#!/usr/bin/env python3
"""tv_link.py - nối bảng TV (sheet "2 Thu vien gia") với file thư viện giá bằng workbook link (Data > Edit Links).

  python3 tv_link.py link   BOQ.xlsx OUT.xlsx --libdir LIBDIR --linkbase "C:\\...\\2. Thu vien gia"
  python3 tv_link.py verify BOQ.xlsx --libdir LIBDIR

Một chiều: ô giá trong TV = công thức trỏ vào ô giá của file thư viện; sửa trong BOQ không ghi ngược file thư viện.
Cột nối: Giá VL / NC thư viện (thô) (P, Q), Giá VL chưa CK, Giá NC ước tính. Cột Giá VL / NC gốc (E, F) giữ nguyên
(= thô x hệ số Loại VL, do hsloai). Chạy SAU hsloai. Thêm cột cuối "Kiểm dòng nguồn" (OK / LỆCH DÒNG / KHÔNG LINK).
Ghi bằng XML, không mở-lưu file BOQ bằng openpyxl.
"""
import sys, os, re, zipfile, argparse, bisect, collections, urllib.parse
import openpyxl
from lxml import etree
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI
from xml.sax.saxutils import escape

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
q = lambda t: '{%s}%s' % (NS, t)

LINK_COLS = {  # cột TV -> tên cột chuẩn ở file nguồn
    'Giá VL thư viện (thô)': 'Giá VL gốc', 'Giá NC thư viện (thô)': 'Giá NC gốc',
    'Giá VL chưa CK': 'Giá VL chưa CK', 'Giá NC ước tính': 'Giá NC ước tính'}
CHECK_COL = 'Kiểm dòng nguồn'


def norm(s): return re.sub(r'\s+', ' ', str(s or '')).strip().lower()


def txt(v):
    if v is None: return ''
    if isinstance(v, float) and v == int(v): v = int(v)
    return norm(v)


# ---- cùng luật nhận diện tiêu đề với tv_loader.py ----
def std_name(top, sub):
    t, s = norm(top), norm(sub)
    if s == 'vật liệu' and ('số liệu gốc' in t or t == ''): return 'Giá VL gốc'
    if s == 'nhân công' and ('số liệu gốc' in t or t == ''): return 'Giá NC gốc'
    if 'nhân công' in t and 'ước' in t: return 'Giá NC ước tính'
    if ('báo giá' in t or 'báo gía' in t or 'b.giá' in t or 'vật tư theo báo giá' in t or 'chưa chiết khấu' in t) and 'nhân công' not in t:
        return 'Giá VL chưa CK'
    if 'mã hiệu' in t and 'mã hiệu' not in s: return 'Mã hiệu'
    if t == 'tên công việc': return 'Tên công việc'
    if t.startswith('unit') or 'đ. vị' in t or t == 'đvt': return 'ĐVT'
    return None


def isnum(v): return isinstance(v, (int, float)) and not isinstance(v, bool)


def src_sheet(ws):
    """-> {'rows': [ {row, ma, ten, vals:{tên chuẩn:(cột, giá trị)}} ], 'cols': {tên chuẩn: số cột}} hoặc None"""
    data = list(ws.iter_rows(min_row=1, values_only=True))
    h = next((i for i, r in enumerate(data[:30]) if any('mã hiệu' in norm(v) for v in r)), None)
    if h is None: return None
    top, sub = list(data[h]), (list(data[h + 1]) if h + 1 < len(data) else [])
    last, colmap = '', {}
    for j in range(len(top)):
        if top[j] not in (None, ''): last = top[j]
        name = std_name(last if top[j] in (None, '') else top[j], sub[j] if j < len(sub) else None)
        if name and name not in colmap.values(): colmap[j] = name
    start = h + 2 if any(norm(v) in ('vật liệu', 'nhân công') for v in sub) else h + 1
    rows = []
    for i in range(start, len(data)):
        r = data[i]
        rec = {colmap[j]: (j + 1, r[j]) for j in colmap if j < len(r)}
        ma, ten = txt(rec.get('Mã hiệu', (0, None))[1]), txt(rec.get('Tên công việc', (0, None))[1])
        if not ma and not ten: continue
        rows.append({'row': i + 1, 'ma': ma, 'ten': ten, 'vals': rec})
    return {'rows': rows, 'cols': {n: j + 1 for j, n in colmap.items()}}


# ---- đọc gói xlsx ----
class Pkg:
    def __init__(self, path):
        z = zipfile.ZipFile(path)
        self.infos = z.infolist()
        self.files = {i.filename: z.read(i.filename) for i in self.infos}
        self.new = {}

    def xml(self, name): return etree.fromstring(self.files[name])


def rels_of(pkg, part):
    d, b = os.path.split(part)
    rp = '%s/_rels/%s.rels' % (d, b)
    if rp not in pkg.files: return {}
    return {r.get('Id'): r for r in etree.fromstring(pkg.files[rp])}


def find_tv(pkg):
    for n in pkg.files:
        if re.match(r'xl/tables/table\d+\.xml$', n):
            t = etree.fromstring(pkg.files[n])
            if t.get('name') == 'TV' or t.get('displayName') == 'TV':
                tn = n
                break
    else:
        raise SystemExit('DỪNG: không có bảng TV.')
    for n in pkg.files:
        m = re.match(r'(xl/worksheets)/_rels/(sheet\d+\.xml)\.rels$', n)
        if m:
            for r in etree.fromstring(pkg.files[n]):
                if r.get('Target', '').endswith(os.path.basename(tn)):
                    return tn, '%s/%s' % (m.group(1), m.group(2))
    raise SystemExit('DỪNG: không tìm thấy sheet chứa bảng TV.')


def shared_strings(pkg):
    if 'xl/sharedStrings.xml' not in pkg.files: return []
    return [''.join(t.itertext()) for t in pkg.xml('xl/sharedStrings.xml').iter(q('si'))]


def cell_val(c, sst):
    t = c.get('t')
    if t == 'inlineStr': return ''.join(c.find(q('is')).itertext())
    v = c.find(q('v'))
    if v is None or v.text is None: return None
    if t == 's': return sst[int(v.text)]
    if t in ('str', 'e'): return v.text
    if t == 'b': return bool(int(v.text))
    f = float(v.text)
    return int(f) if f == int(f) and abs(f) < 1e15 else f


def split_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)$', ref); return CI(m.group(1)), int(m.group(2))


def load_tv(pkg):
    tn, sp = find_tv(pkg)
    tbl = pkg.xml(tn)
    ref = tbl.get('ref')
    a, b = ref.split(':'); c0, r0 = split_ref(a); c1, r1 = split_ref(b)
    names = [tc.get('name') for tc in tbl.find(q('tableColumns'))]
    sst = shared_strings(pkg)
    ws = pkg.xml(sp)
    rows = {}
    for row in ws.find(q('sheetData')):
        for c in row:
            ci, ri = split_ref(c.get('r'))
            if r0 <= ri <= r1 and c0 <= ci <= c1:
                rows.setdefault(ri, {})[names[ci - c0]] = cell_val(c, sst)
    return tn, sp, tbl, ws, names, c0, r0, r1, rows


def lib_files(libdir):
    return {f: os.path.join(libdir, f) for f in os.listdir(libdir) if f.lower().endswith('.xlsx') and not f.startswith('~$')}


def plan(rows, r0, r1, libdir):
    """-> (links, stat): links[row] = {'file','sheet','src_row','cols':{tvcol:(src_col_idx,value)}, 'ten_col'}"""
    files = lib_files(libdir)
    groups = collections.OrderedDict()
    for r in range(r0 + 1, r1 + 1):
        d = rows.get(r, {})
        sh = re.sub(r' \(sheet ẩn\)$', '', str(d.get('Sheet nguồn') or ''))
        groups.setdefault((d.get('File nguồn'), sh), []).append(r)
    links, stat = {}, collections.Counter()
    cache = {}
    for (fn, sh), rlist in groups.items():
        if fn not in files:
            stat['file nguồn không có trong thư mục'] += len(rlist); continue
        if fn not in cache: cache[fn] = openpyxl.load_workbook(files[fn], read_only=True, data_only=True)
        wb = cache[fn]
        if sh not in wb.sheetnames:
            stat['sheet nguồn không có'] += len(rlist); continue
        S = src_sheet(wb[sh])
        if not S:
            stat['không nhận diện được tiêu đề'] += len(rlist); continue
        idx = collections.defaultdict(list)
        for k, s in enumerate(S['rows']): idx[(s['ma'], s['ten'])].append(k)
        p = 0
        for r in rlist:
            d = rows[r]
            key = (txt(d.get('Mã hiệu')), txt(d.get('Tên công việc')))
            lst = idx.get(key, [])
            pos = bisect.bisect_left(lst, p)
            if pos >= len(lst):
                stat['không khớp dòng nguồn'] += 1; continue
            k = lst[pos]; p = k + 1
            s = S['rows'][k]
            cols = {}
            for tvc, nm in LINK_COLS.items():
                tv, sv = d.get(tvc), s['vals'].get(nm)
                if tv is None or sv is None or not isnum(tv) or not isnum(sv[1]):
                    continue
                if abs(float(tv) - float(sv[1])) > 1e-6 * max(1, abs(float(tv))):
                    stat['giá lệch file nguồn (giữ giá tĩnh): ' + tvc] += 1; continue
                cols[tvc] = (sv[0], sv[1])
            tcol = S['cols'].get('Tên công việc')
            mcol = S['cols'].get('Mã hiệu')
            links[r] = {'file': fn, 'sheet': sh, 'src_row': s['row'], 'cols': cols, 'ten_col': tcol, 'ma_col': mcol,
                        'ten': (s['vals'].get('Tên công việc') or (0, ''))[1], 'ma': (s['vals'].get('Mã hiệu') or (0, ''))[1]}
            stat['dòng đã nối'] += 1
    for wb in cache.values(): wb.close()
    return links, stat, cache.keys()


# ---- ghi XML ----
def put(row_el, ref, f=None, v=None, kind='n', style=None):
    ci = split_ref(ref)[0]
    for c in row_el:
        if c.get('r') == ref:
            row_el.remove(c); break
    c = etree.Element(q('c')); c.set('r', ref)
    if style: c.set('s', style)
    if f is not None:
        etree.SubElement(c, q('f')).text = f
    if kind == 'str': c.set('t', 'str')
    if v is not None: etree.SubElement(c, q('v')).text = (repr(float(v)) if kind == 'n' and isinstance(v, float) else str(v))
    pos = len(row_el)
    for i, o in enumerate(row_el):
        if split_ref(o.get('r'))[0] > ci: pos = i; break
    row_el.insert(pos, c)
    return c


def ext_xml(sheets, used, relid='rId1'):
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           '<externalLink xmlns="%s" xmlns:r="%s"><externalBook r:id="%s"><sheetNames>' % (NS, RNS, relid)]
    out += ['<sheetName val="%s"/>' % escape(s, {'"': '&quot;'}) for s in sheets]
    out.append('</sheetNames><sheetDataSet>')
    for si in sorted(used):
        out.append('<sheetData sheetId="%d">' % si)
        for r in sorted(used[si]):
            out.append('<row r="%d">' % r)
            for cidx in sorted(used[si][r]):
                v = used[si][r][cidx]
                ref = '%s%d' % (CL(cidx), r)
                if isnum(v): out.append('<cell r="%s"><v>%r</v></cell>' % (ref, float(v) if isinstance(v, float) else v))
                else: out.append('<cell r="%s" t="str"><v>%s</v></cell>' % (ref, escape(str(v))))
            out.append('</row>')
        out.append('</sheetData>')
    out.append('</sheetDataSet></externalBook></externalLink>')
    return ''.join(out).encode('utf8')


def do_link(a):
    pkg = Pkg(a.boq)
    tn, sp, tbl, ws, names, c0, r0, r1, rows = load_tv(pkg)
    for need in list(LINK_COLS)[:2] + ['Mã hiệu', 'Tên công việc', 'File nguồn', 'Sheet nguồn']:
        if need not in names:
            raise SystemExit('DỪNG: bảng TV thiếu cột "%s". Chạy hsloai (cột P–S) trước khi nối link.' % need)
    if CHECK_COL in names: raise SystemExit('DỪNG: bảng TV đã có cột "%s" (đã nối link). Dùng verify, hoặc Break Link rồi nối lại.' % CHECK_COL)
    if any(re.match(r'xl/externalLinks/externalLink\d+\.xml$', n) and b'tv_link' in pkg.files.get(n.replace('externalLinks/', 'externalLinks/_rels/') + '.rels', b'') for n, b in pkg.files.items()):
        raise SystemExit('DỪNG: file đã có link do tv_link tạo.')
    links, stat, _ = plan(rows, r0, r1, a.libdir)
    allowed = set(LINK_COLS) if a.cols == 'all' else set(list(LINK_COLS)[:2])   # mặc định chỉ nối 2 cột giá thô (E/F đã = thô × hệ số)
    for L in links.values(): L['cols'] = {k: v for k, v in L['cols'].items() if k in allowed}
    check = not a.no_check
    if not links: raise SystemExit('DỪNG: không nối được dòng nào. ' + str(dict(stat)))

    # chỉ mục externalLink hiện có
    wb = pkg.xml('xl/workbook.xml')
    exist = wb.find(q('externalReferences'))
    n_exist = len(exist) if exist is not None else 0
    n_ext = 1 + max([int(m.group(1)) for k in pkg.files for m in [re.match(r'xl/externalLinks/externalLink(\d+)\.xml$', k)] if m] or [0])
    wbrels = pkg.xml('xl/_rels/workbook.xml.rels')

    files = sorted({l['file'] for l in links.values()})
    fidx = {}      # file -> chỉ số [n] dùng trong công thức
    allsheets = {}
    for k, fn in enumerate(files):
        fidx[fn] = n_exist + 1 + k
    srcwb = {fn: openpyxl.load_workbook(os.path.join(a.libdir, fn), read_only=True) for fn in files}
    used = {fn: collections.defaultdict(lambda: collections.defaultdict(dict)) for fn in files}

    sheetdata = ws.find(q('sheetData'))
    rowel = {int(r.get('r')): r for r in sheetdata}
    ncol_new = c0 + len(names)           # cột kiểm
    chk_letter = CL(ncol_new)
    if check:
        hdr = rowel[r0]
        put(hdr, '%s%d' % (chk_letter, r0), v=None)
        c = hdr[-1]; c.set('t', 'inlineStr'); etree.SubElement(etree.SubElement(c, q('is')), q('t')).text = CHECK_COL
    colletter = {n: CL(c0 + i) for i, n in enumerate(names)}
    n_link = 0
    for r in range(r0 + 1, r1 + 1):
        el = rowel.get(r)
        if el is None: continue
        L = links.get(r)
        if not L:
            if not check: continue
            c = put(el, '%s%d' % (chk_letter, r), f=None, v='KHÔNG LINK', kind='str')
            c.attrib.pop('t', None); c.set('t', 'inlineStr'); c.remove(c.find(q('v')))
            etree.SubElement(etree.SubElement(c, q('is')), q('t')).text = 'KHÔNG LINK'
            continue
        n = fidx[L['file']]; sh = L['sheet']
        qs = "'[%d]%s'" % (n, sh.replace("'", "''"))
        si = srcwb[L['file']].sheetnames.index(sh)
        for tvc, (scol, val) in L['cols'].items():
            st = None
            old = [c for c in el if c.get('r') == '%s%d' % (colletter[tvc], r)]
            if old: st = old[0].get('s')
            put(el, '%s%d' % (colletter[tvc], r), f='%s!$%s$%d' % (qs, CL(scol), L['src_row']), v=val, style=st)
            used[L['file']][si][L['src_row']][scol] = val
            n_link += 1
        if check and L['ten_col'] and L['ma_col']:
            sr = L['src_row']
            used[L['file']][si][sr][L['ten_col']] = L['ten'] if L['ten'] is not None else ''
            used[L['file']][si][sr][L['ma_col']] = L['ma'] if L['ma'] is not None else ''
            f = ('IF(EXACT(TRIM(%s!$%s$%d&"")&"|"&TRIM(%s!$%s$%d&""),TRIM($%s%d&"")&"|"&TRIM($%s%d&"")),"OK","LỆCH DÒNG")' % (
                qs, CL(L['ma_col']), sr, qs, CL(L['ten_col']), sr,
                colletter['Mã hiệu'], r, colletter['Tên công việc'], r))
            put(el, '%s%d' % (chk_letter, r), f=f, v='OK', kind='str')
    for o in rowel.values():
        pass
    if check:
        # bảng: thêm cột
        new_ref = '%s%d:%s%d' % (CL(c0), r0, chk_letter, r1)
        tbl.set('ref', new_ref)
        af = tbl.find(q('autoFilter'))
        if af is not None: af.set('ref', new_ref)
        tcs = tbl.find(q('tableColumns'))
        nid = 1 + max(int(t.get('id')) for t in tcs)
        etree.SubElement(tcs, q('tableColumn')).attrib.update({'id': str(nid), 'name': CHECK_COL})
        tcs.set('count', str(len(tcs)))
        dim = ws.find(q('dimension'))
        if dim is not None: dim.set('ref', re.sub(r':[A-Z]+(\d+)$', lambda m: ':%s%s' % (chk_letter, m.group(1)), dim.get('ref')))
    pkg.new[tn] = etree.tostring(tbl, xml_declaration=True, encoding='UTF-8', standalone=True)
    pkg.new[sp] = etree.tostring(ws, xml_declaration=True, encoding='UTF-8', standalone=True)

    # externalLink parts
    ct = pkg.files['[Content_Types].xml'].decode('utf8'); add = ''
    if exist is None:
        exist = etree.Element(q('externalReferences'))
        sh_el = wb.find(q('sheets')); sh_el.addnext(exist)
    for k, fn in enumerate(files):
        part = 'xl/externalLinks/externalLink%d.xml' % (n_ext + k)
        pkg.new[part] = ext_xml(srcwb[fn].sheetnames, used[fn])
        lb = a.linkbase.rstrip('\\/')
        target = ('file://' + lb + '/' + fn) if lb.startswith('/') else ('file:///' + lb + '\\' + fn)
        target = target.replace(' ', '%20')   # Excel bỏ cả link nếu Target còn dấu cách (đã gặp 22/10/2026)
        pkg.new['xl/externalLinks/_rels/externalLink%d.xml.rels' % (n_ext + k)] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="%s">'
            '<Relationship Id="rId1" Type="%s/externalLinkPath" Target="%s" TargetMode="External"/></Relationships>'
            % (PR, RNS, escape(target, {'"': '&quot;'}))).encode('utf8')
        rid = 'rIdExt%d' % (n_ext + k)
        etree.SubElement(wbrels, '{%s}Relationship' % PR, Id=rid, Type=RNS + '/externalLink', Target='externalLinks/externalLink%d.xml' % (n_ext + k))
        etree.SubElement(exist, q('externalReference')).set('{%s}id' % RNS, rid)
        add += '<Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>' % part
    pkg.new['xl/_rels/workbook.xml.rels'] = etree.tostring(wbrels, xml_declaration=True, encoding='UTF-8', standalone=True)
    pkg.new['xl/workbook.xml'] = etree.tostring(wb, xml_declaration=True, encoding='UTF-8', standalone=True)
    pkg.new['[Content_Types].xml'] = ct.replace('</Types>', add + '</Types>').encode('utf8')
    zo = zipfile.ZipFile(a.out, 'w', zipfile.ZIP_DEFLATED)
    for it in pkg.infos: zo.writestr(it, pkg.new.pop(it.filename, pkg.files[it.filename]))
    for k, v in pkg.new.items(): zo.writestr(k, v)
    zo.close()
    tot = r1 - r0
    print('Đã nối %d / %d dòng TV với %d file thư viện (%d ô giá). Cột kiểm: %s.' % (stat['dòng đã nối'], tot, len(files), n_link, ('"%s"' % CHECK_COL) if check else 'không có (--no-check)'))
    for k, v in stat.items():
        if k != 'dòng đã nối': print('  %s: %d' % (k, v))
    print('Link trỏ tới: %s\\<tên file>' % a.linkbase.rstrip('\\/'))


def do_verify(a):
    """So cache TV với file thư viện hiện tại: ô nguồn có còn đúng Tên công việc và giá."""
    pkg = Pkg(a.boq)
    tn, sp, tbl, ws, names, c0, r0, r1, rows = load_tv(pkg)
    sst = shared_strings(pkg)
    ext = {}
    for k in sorted(pkg.files):
        m = re.match(r'xl/externalLinks/externalLink(\d+)\.xml$', k)
        if m:
            rl = etree.fromstring(pkg.files['xl/externalLinks/_rels/externalLink%s.xml.rels' % m.group(1)])
            ext[m.group(1)] = urllib.parse.unquote(os.path.basename(rl[0].get('Target').replace('\\', '/')))
    order = {}
    for i, rid in enumerate(r.get('{%s}id' % RNS) for r in pkg.xml('xl/workbook.xml').find(q('externalReferences'))):
        rels = {r.get('Id'): r.get('Target') for r in pkg.xml('xl/_rels/workbook.xml.rels')}
        order[i + 1] = re.search(r'externalLink(\d+)\.xml', rels[rid]).group(1)
    cl = {n: CL(c0 + i) for i, n in enumerate(names)}
    # Pass 1: gom mọi ô nguồn cần đọc (file, sheet) -> {ref}
    need = collections.defaultdict(set)
    price_tasks, check_tasks = [], []
    for c in ws.find(q('sheetData')).iter(q('c')):
        f = c.find(q('f'))
        if f is None or not f.text: continue
        ci, ri = split_ref(c.get('r'))
        nm = names[ci - c0]
        if nm in LINK_COLS:
            m = re.match(r"'\[(\d+)\](.+)'!\$([A-Z]+)\$(\d+)$", f.text)
            if not m: continue
            fn = ext[order[int(m.group(1))]]; sh = m.group(2).replace("''", "'")
            need[(fn, sh)].add(m.group(3) + m.group(4))
            price_tasks.append((fn, sh, m.group(3) + m.group(4), cell_val(c, sst)))
        elif nm == CHECK_COL and ri != r0:
            ms = re.findall(r"'\[(\d+)\](.+?)'!\$([A-Z]+)\$(\d+)", f.text)
            if len(ms) < 2: continue
            fn = ext[order[int(ms[0][0])]]; sh = ms[0][1].replace("''", "'")
            refs = [c_ + n_ for (_, _, c_, n_) in ms[:2]]
            need[(fn, sh)].update(refs)
            check_tasks.append((ri, fn, sh, refs, ms[0][2] + ms[0][3]))
    # Pass 2: mỗi sheet nguồn chỉ quét MỘT lần (read_only mà truy ô lẻ thì mỗi lần đọc lại từ đầu sheet)
    vals = {}
    for (fn, sh), refs in need.items():
        wb = openpyxl.load_workbook(os.path.join(a.libdir, fn), read_only=True, data_only=True)
        d = {}
        if sh in wb.sheetnames:
            rr = []
            for ref in refs:
                mm = re.match(r'([A-Z]+)(\d+)$', ref); rr.append((int(mm.group(2)), ref))
            lo, hi = min(r for r, _ in rr), max(r for r, _ in rr)
            want = {}
            for r, ref in rr: want.setdefault(r, []).append(ref)
            for ri_, row in enumerate(wb[sh].iter_rows(min_row=lo, max_row=hi), lo):
                if ri_ in want:
                    byref = {cc.coordinate: cc.value for cc in row if hasattr(cc, 'coordinate')}
                    for ref in want[ri_]: d[ref] = byref.get(ref)
        vals[(fn, sh)] = d
        wb.close()
    ok, price_diff = 0, 0
    for fn, sh, ref, v in price_tasks:
        cell = vals[(fn, sh)].get(ref)
        if not (isnum(cell) and isnum(v) and abs(cell - v) <= 1e-6 * max(1, abs(v))): price_diff += 1
        else: ok += 1
    lech, tong = [], 0
    for ri, fn, sh, refs, ref0 in check_tasks:
        cur_ma, cur_ten = (vals[(fn, sh)].get(x) for x in refs)
        tong += 1
        if (txt(cur_ma), txt(cur_ten)) != (txt(rows[ri].get('Mã hiệu')), txt(rows[ri].get('Tên công việc'))):
            lech.append((ri, fn, sh, ref0))
    print('Ô giá nối: %d khớp với thư viện hiện tại, %d khác.' % (ok, price_diff))
    if not check_tasks: print('Lưu ý: file không có cột "%s" nên chỉ so giá, không phát hiện được lệch dòng.' % CHECK_COL)
    print('Kiểm dòng nguồn: %d / %d dòng LỆCH DÒNG (ô nguồn không còn đúng Mã hiệu + Tên công việc, thường do chèn / xóa dòng trong file thư viện).' % (len(lech), tong))
    for ri, fn, sh, ref in lech[:20]: print('  TV dòng %d -> %s [%s] %s' % (ri, fn, sh, ref))
    if lech: print('  => Nối link lại sau khi Break Link, hoặc dùng Power Query (đọc theo tên cột, không lệch dòng).')
    elif price_diff: print('  => Giá thư viện đã đổi so với lần lưu cuối của BOQ: mở BOQ, Update links.')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    sp = ap.add_subparsers(dest='cmd', required=True)
    l = sp.add_parser('link'); l.add_argument('boq'); l.add_argument('out'); l.add_argument('--libdir', required=True); l.add_argument('--linkbase', required=True)
    l.add_argument('--cols', choices=['raw', 'all'], default='raw', help='raw (mặc định): chỉ nối 2 cột giá thô P,Q; all: nối cả Giá VL chưa CK / Giá NC ước tính')
    l.add_argument('--no-check', action='store_true', help='không thêm cột Kiểm dòng nguồn (file nhẹ hơn, mở nhanh hơn; mất cảnh báo lệch dòng)')
    v = sp.add_parser('verify'); v.add_argument('boq'); v.add_argument('--libdir', required=True)
    a = ap.parse_args()
    do_link(a) if a.cmd == 'link' else do_verify(a)