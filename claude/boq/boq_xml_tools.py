#!/usr/bin/env python3
"""boq_xml_tools.py - công cụ ghi trực tiếp XML vào file BOQ template TDI (không dùng openpyxl để ghi).

Lệnh:
  sheets   FILE                         : tên sheet -> file XML (luôn tra lại, Excel hay đổi số)
  cols     FILE SHEET                   : nhận diện cột TDI + cột phụ theo tiêu đề (in JSON)
  helpers  FILE OUT SHEET...            : tạo 5 cột phụ (Nhóm VT, Độ tin cậy, Ghi chú gán mã, Hệ số ĐV, Loại dòng) nếu chưa có
  insert   FILE OUT PLAN.json           : bóc tách - chèn dòng con dưới dòng cụm, tô màu phần TDI
  fillcdt  FILE OUT CFG.json            : đưa đơn giá TDI sang cột giá BOQ CĐT bằng công thức
  recalc   FILE [--sheets 'A|B']          : xóa cache <v>, tính lại bằng LibreOffice trên BẢN SAO, báo lỗi/tổng/có mã không giá
  cleanup  FILE OUT                     : xóa sheet tạm _TDI_JOB / _TDI_TMP_*
  migrate  FILE OUT                     : sau khi nạp lại TV (tv_loader v3): đổi Nhóm VT + SUMIFS viết cứng sang Loại vật liệu (theo mã); liệt kê dòng mơ hồ
  nguongia FILE OUT [SHEET...]          : Bước 6c: cột Nguồn giá (công thức) + nới autoFilter
  cho      FILE OUT [SHEET...]          : Bước 6e: công thức chờ cho ô giá trống + bọc chặn mã trống (Quy tắc 34)
  kiem     FILE [--json]                : tự kiểm các mục Bảng kiểm máy kiểm được, 1 lần tính lại (QT21, 23, 29, 34, 6b)
  dropdown FILE OUT SHEET...            : gắn droplist cho cột Mã hiệu: chỉ liệt kê mã của Nhóm VT trên chính dòng đó
                                          (danh sách lấy từ bảng TV, lưu ở sheet ẩn DS_MaHieu; chạy lại được, chạy sau insert cũng được)

PLAN.json (insert): [[sheet, parent_row, expected_desc_prefix, children[], parent_note, parent_conf], ...]
  child = {g, key ('Mã hiệu'|'Tên công việc'|null), val, ma, name, unit, coef, conf, note}
  parent_row theo số dòng HIỆN TẠI của FILE. Chèn là bước cuối (số dòng đổi sau đó).
CFG.json (fillcdt): {"<sheet>": {"kl":["D","E","F"], "zones":["K","N","Q"], "h":"H","i":"I","j":"J",
                     "unit":"C", "start":8}}  - cột TDI (VL/NC sau hệ số, Loại dòng) tự nhận theo tiêu đề.
  CĐT chỉ 1 cột đơn giá: {"<sheet>": {"kl":["E"], "unit":"D", "start":6, "single":{"f":"F","g":"G"}}}
detect_cols: cột "Nội Dung Công Việc" lấy ô NGOÀI CÙNG BÊN PHẢI (CĐT có thể dùng cùng tiêu đề ở cột B).
Công thức giá (tv_sumifs, bản 14/10/2026): SUMIFS tham chiếu ô Nhóm VT / Mã hiệu của chính dòng, ví dụ
  SUMIFS(TV[Giá NC gốc],TV[Loại vật liệu],$T265,TV[Mã hiệu],$N265). Chỉ viết cứng khi tra theo Tên công việc (mã 0),
  số hạng ghép khác nhóm/mã của dòng, hoặc giá trị có ký tự đại diện ~ * ?.
"""
import posixpath, zipfile, re, json, copy, sys, os, subprocess, tempfile, shutil, collections
from lxml import etree
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
q = lambda t: '{%s}%s' % (NS, t)
GCOL = 'Loại vật liệu'   # tên cột nhóm của bảng TV (trước 15/10/2026: 'Column1'); Pkg tự nhận theo bảng TV trong file
HELPERS = ['Nhóm VT', 'Độ tin cậy', 'Ghi chú gán mã', 'Hệ số ĐV', 'Loại dòng', 'Mã tra cứu', 'Tên tra cứu']   # 2 cột cuối (16/10/2026): tách Mã / Tên từ ô Mã hiệu dạng "Mã | Tên công việc"
SEP = ' | '
COLOR_NAME, COLOR_CODE, FILL_CHILD = 'FF0070C0', 'FFFF0000', 'FFFFF2CC'
CONF_ORDER = ['Cao', 'TB', 'Thấp', 'Chưa khớp']


# ---------------------------------------------------------------- package I/O
class Pkg:
    def __init__(self, path):
        z = zipfile.ZipFile(path)
        self.infos = z.infolist()
        self.files = {i.filename: z.read(i.filename) for i in self.infos}
        self.new = {}
        self.wb = etree.fromstring(self.files['xl/workbook.xml'])
        rels = etree.fromstring(self.files['xl/_rels/workbook.xml.rels'])
        rid = {r.get('Id'): r.get('Target') for r in rels}
        self.sheets = []
        for s in self.wb.find(q('sheets')):
            t = rid[s.get('{%s}id' % RNS)]
            t = t.lstrip('/')
            t = t if t.startswith('xl/') else 'xl/' + t
            self.sheets.append((s.get('name'), t))
        self.path = dict(self.sheets)
        self.ss = []
        if 'xl/sharedStrings.xml' in self.files:
            sst = etree.fromstring(self.files['xl/sharedStrings.xml'])
            self.ss = [''.join(t.text or '' for t in si.iter(q('t'))) for si in sst]
        self._roots = {}
        self._detect_gcol()

    def _detect_gcol(self):
        global GCOL
        for n, d in self.files.items():
            if re.match(r'xl/tables/table\d+\.xml$', n) and re.search(rb'(name|displayName)="TV"', d):
                try: GCOL = etree.fromstring(d).find(q('tableColumns'))[0].get('name')
                except Exception: pass

    def root(self, sheet):
        p = self.path[sheet]
        if p not in self._roots:
            self._roots[p] = etree.fromstring(self.new.get(p, self.files[p]))
        return self._roots[p]

    def touch(self, sheet):
        self.new[self.path[sheet]] = None  # serialise at save

    def val(self, c):
        t = c.get('t'); v = c.find(q('v'))
        if t == 's' and v is not None: return self.ss[int(v.text)]
        if t == 'inlineStr': return ''.join(x.text or '' for x in c.iter(q('t')))
        f = c.find(q('f'))
        if f is not None and f.text: return '=' + f.text
        return v.text if v is not None else None

    def grid(self, sheet, r1, r2):
        out = {}
        for row in self.root(sheet).find(q('sheetData')):
            r = int(row.get('r'))
            if r1 <= r <= r2:
                for c in row:
                    out[c.get('r')] = self.val(c)
        return out

    def save(self, out, drop=()):
        for p, rt in self._roots.items():
            if p in self.new:
                self.new[p] = etree.tostring(rt, xml_declaration=True, encoding='UTF-8', standalone=True)
        cp = self.wb.find(q('calcPr'))
        if cp is None: cp = etree.SubElement(self.wb, q('calcPr'))
        cp.set('fullCalcOnLoad', '1')
        self.new['xl/workbook.xml'] = etree.tostring(self.wb, xml_declaration=True, encoding='UTF-8', standalone=True)
        z = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
        for it in self.infos:
            if it.filename in drop: continue
            z.writestr(it, self.new.get(it.filename) or self.files[it.filename])
        have = {it.filename for it in self.infos}
        for p, data in self.new.items():
            if p not in have and data and p not in drop: z.writestr(p, data)   # part mới (vd. sheet DS_MaHieu)
        z.close()


def colnum(ref): return CI(re.match(r'[A-Z]+', ref).group(0))


def setc(row, col, val=None, formula=None, num=None, style=None, clear=False):
    """Ghi 1 ô vào <row>, giữ thứ tự cột và style sẵn có."""
    ref = CL(col) + row.get('r'); cell = None; pos = len(row)
    for i, c in enumerate(row):
        cc = colnum(c.get('r'))
        if cc == col: cell = c; break
        if cc > col: pos = i; break
    if cell is None:
        if clear: return None
        cell = etree.Element(q('c')); cell.set('r', ref); row.insert(pos, cell)
    s = cell.get('s') if style is None else str(style)
    for ch in list(cell): cell.remove(ch)
    for a in list(cell.attrib):
        if a != 'r': del cell.attrib[a]
    if s: cell.set('s', s)
    if formula is not None:
        etree.SubElement(cell, q('f')).text = formula
    elif num is not None:
        etree.SubElement(cell, q('v')).text = repr(num)
    elif val not in (None, ''):
        cell.set('t', 'inlineStr')
        etree.SubElement(etree.SubElement(cell, q('is')), q('t')).text = str(val)
    return cell


def esc(s): return str(s).replace('~', '~~').replace('*', '~*').replace('?', '~?').replace('"', '""')


WILD = re.compile(r'[~*?]')


def split_code(txt):
    """Giống công thức cột phụ: 'Mã | Tên' -> (mã, tên); '(mã 0) | Tên' / '(không mã) | Tên' -> ('', tên);
    '(mã 0) Tên' (dạng cũ) -> ('', Tên); mã thuần -> (mã, '')."""
    t = str(txt or '')
    if SEP in t:
        c, n = t.split(SEP, 1)
        return ('' if c.startswith('(') else c), n
    if t.startswith('(mã 0)'): return '', t[7:]
    return t, ''


def legacy_code(txt):
    """Ô Mã hiệu dạng mới -> dạng cũ lưu trong Bảng khớp: 'Mã | Tên' -> Mã; '(mã 0) | Tên' -> '(mã 0) Tên'."""
    t = str(txt or '')
    if SEP in t:
        c, n = t.split(SEP, 1)
        return ('(mã 0) ' + n) if c.startswith('(') else c
    return t


def _esc_f(x): return 'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(%s,"~","~~"),"*","~*"),"?","~?")' % x


def hf_code(L, r):
    P = '$%s%d' % (L, r)
    return 'IF(%s="","",IF(ISNUMBER(FIND(" | ",%s)),IF(LEFT(%s,1)="(","",%s),IF(LEFT(%s,6)="(mã 0)","",%s)))' % (
        P, P, P, _esc_f('LEFT(%s,FIND(" | ",%s)-1)' % (P, P)), P, _esc_f(P))


def hf_name(L, r):
    P = '$%s%d' % (L, r)
    return 'IF(%s="","",IF(ISNUMBER(FIND(" | ",%s)),%s,IF(LEFT(%s,6)="(mã 0)",%s,"")))' % (
        P, P, _esc_f('MID(%s,FIND(" | ",%s)+3,500)' % (P, P)), P, _esc_f('MID(%s,8,500)' % P))


def tv_sumifs(lc, g, key, val, r=None, gcol=None, row_g=None, ccol=None, row_code=None, mcol=None, tcol=None):
    """Một số hạng SUMIFS tra bảng TV (bản 16/10/2026).
    Dòng có cột phụ 'Mã tra cứu' / 'Tên tra cứu' (tách từ ô Mã hiệu 'Mã | Tên công việc') -> tra Loại vật liệu + Mã + Tên của chính dòng:
      SUMIFS(TV[Giá VL gốc],TV[Loại vật liệu],$AX11,TV[Mã hiệu],IF($BC11="","<>§",$BC11),TV[Tên công việc],IF($BD11="","<>§",$BD11))
      (mã hoặc tên trống = không lọc theo tiêu chí đó)
    Còn lại giữ cách cũ: nhóm / mã tham chiếu ô, viết cứng (có escape) khi tra theo Tên của số hạng ghép khác nhóm/mã, hoặc giá trị có ~ * ?."""
    sg, sv = str(g if g is not None else ''), str(val if val is not None else '')
    own_g = bool(r and gcol and row_g is not None and sg == str(row_g) and not WILD.search(sg))
    gs = '$%s%d' % (CL(gcol), r) if own_g else '"%s"' % esc(sg)
    if own_g and mcol and tcol and row_code is not None and sv:
        rc, rn = split_code(row_code)
        if (key == 'Mã hiệu' and sv == rc) or (key == 'Tên công việc' and sv == rn):
            M, N = '$%s%d' % (CL(mcol), r), '$%s%d' % (CL(tcol), r)
            return 'SUMIFS(TV[%s],TV[%s],%s,TV[Mã hiệu],IF(%s="","<>§",%s),TV[Tên công việc],IF(%s="","<>§",%s))' % (lc, GCOL, gs, M, M, N, N)
    if r and ccol and key == 'Mã hiệu' and row_code is not None and sv == str(row_code) and sv and not WILD.search(sv):
        vs = '$%s%d' % (CL(ccol), r)
    else: vs = '"%s"' % esc(sv)
    return 'SUMIFS(TV[%s],TV[%s],%s,TV[%s],%s)' % (lc, GCOL, gs, key, vs)


def set_helpers(row, cd, mc, tc, code_text):
    """Ghi (hoặc xóa) 2 công thức cột phụ trên một dòng."""
    r = int(row.get('r'))
    if str(code_text or '').strip():
        setc(row, mc, formula=hf_code(CL(cd), r)); setc(row, tc, formula=hf_name(CL(cd), r))
    else:
        setc(row, mc, ''); setc(row, tc, '')


# ---------------------------------------------------------------- column detection
def detect_cols(pkg, sheet, hdr_rows=(1, 12)):
    g = pkg.grid(sheet, *hdr_rows)
    norm = lambda s: re.sub(r'\s+', ' ', str(s or '')).strip().lower()
    by = collections.defaultdict(list)
    for ref, v in g.items():
        if v is None or str(v).startswith('='): continue
        by[norm(v)].append((int(re.search(r'\d+', ref).group(0)), colnum(ref)))
    def first(name, after=0):
        cands = sorted(c for r, c in by.get(norm(name), []) if c > after)
        return cands[0] if cands else None
    _c = sorted(c for r, c in by.get(norm('Nội Dung Công Việc'), []))
    name = _c[-1] if _c else None   # phần TDI nằm bên phải; CĐT có thể trùng tiêu đề
    if not name: raise SystemExit('Không thấy cột "Nội Dung Công Việc" (phần TDI) trong %s' % sheet)
    res = {'name': name, 'code': name + 1, 'brand': name + 2, 'unit_tdi': name + 3,
           'vl0': name + 4, 'nc0': name + 5, 'chung_loai': name + 6}
    dg = first('Đơn giá', name + 6)          # nhóm đơn giá sau hệ số
    if dg: res.update(vl=dg, nc=dg + 1, dg=dg + 2)
    for h in HELPERS:
        res[h] = first(h, name)
    # dòng tiêu đề chính (để tạo cột phụ)
    rows = [r for r, c in by.get(norm('Nội Dung Công Việc'), []) if c == name]
    res['hdr_row'] = rows[0] if rows else 5
    res['letters'] = {k: CL(v) for k, v in res.items() if isinstance(v, int) and k != 'hdr_row'}
    return res


def cmd_cols(f, sheet):
    print(json.dumps(detect_cols(Pkg(f), sheet)['letters'], ensure_ascii=False, indent=1))


def cmd_helpers(f, out, sheets):
    pkg = Pkg(f)
    for sh in sheets:
        c = detect_cols(pkg, sh)
        miss = [h for h in HELPERS if not c[h]]
        if not miss: print(sh, 'đã có cột phụ', [CL(c[h]) for h in HELPERS]); continue
        # đặt sau cột có dữ liệu cuối cùng của dòng tiêu đề
        sd = pkg.root(sh).find(q('sheetData'))
        rows = {int(r.get('r')): r for r in sd}
        hrow = rows[c['hdr_row']]
        last = max(colnum(x.get('r')) for x in hrow)
        sty = next((x.get('s') for x in hrow if x.get('r') == '%s%d' % (CL(c['Loại dòng']), c['hdr_row'])), None) if c['Loại dòng'] else None
        start = last + 1
        for i, h in enumerate(miss):
            setc(hrow, start + i, h, style=sty)
        pkg.touch(sh)
        print(sh, 'tạo cột phụ', miss, CL(start), '->', CL(start + len(miss) - 1))
    pkg.save(out)


# ---------------------------------------------------------------- insert (bóc tách)
def cmd_insert(f, out, planfile):
    P = json.load(open(planfile))
    pkg = Pkg(f)
    plan = collections.defaultdict(list)
    for sh, rn, exp, ch, note, conf in P: plan[sh].append((rn, len(ch)))
    for sh in plan: plan[sh].sort()
    fmap = lambda sh, n: n + sum(k for p, k in plan.get(sh, []) if n > p)
    CELL = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")
    ROWR = re.compile(r"^(\$?)(\d+):(\$?)(\d+)$")

    def shift_ref(ref, sh):
        if '[' in ref: return ref
        m = ROWR.match(ref)
        if m: return '%s%d:%s%d' % (m.group(1), fmap(sh, int(m.group(2))), m.group(3), fmap(sh, int(m.group(4))))
        return CELL.sub(lambda m: '%s%s%s%d' % (m.group(1), m.group(2), m.group(3), fmap(sh, int(m.group(4)))), ref)

    def shift_formula(fx, cur):
        tok = Tokenizer('=' + fx); outv = []; changed = False
        for t in tok.items:
            v = t.value
            if t.type == 'OPERAND' and t.subtype == 'RANGE':
                if '!' in v:
                    pre, ref = v.rsplit('!', 1); sn = pre.strip("'").replace("''", "'")
                else:
                    pre, ref, sn = None, v, cur
                if sn in plan:
                    nr = shift_ref(ref, sn)
                    if nr != ref: v = (pre + '!' if pre is not None else '') + nr; changed = True
            outv.append(v)
        return ''.join(outv) if changed else fx

    def rt_ok(fx):
        try: return ''.join(t.value for t in Tokenizer('=' + fx).items) == fx
        except Exception: return False
    # styles
    st = etree.fromstring(pkg.files['xl/styles.xml'])
    fonts, fills, xfs = st.find(q('fonts')), st.find(q('fills')), st.find(q('cellXfs'))
    fl = etree.SubElement(fills, q('fill')); pf = etree.SubElement(fl, q('patternFill'), patternType='solid')
    etree.SubElement(pf, q('fgColor'), rgb=FILL_CHILD); etree.SubElement(pf, q('bgColor'), indexed='64')
    fills.set('count', str(len(fills))); FILL = len(fills) - 1
    fcache, xcache = {}, {}

    def font_with(fid, rgb):
        if (fid, rgb) in fcache: return fcache[(fid, rgb)]
        fo = copy.deepcopy(fonts[fid]); c = fo.find(q('color'))
        if c is None:
            c = etree.Element(q('color')); order = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz']
            pos = 0
            for i, ch in enumerate(fo):
                if etree.QName(ch).localname in order: pos = i + 1
            fo.insert(pos, c)
        for a in list(c.attrib): del c.attrib[a]
        c.set('rgb', rgb); fonts.append(fo); fonts.set('count', str(len(fonts)))
        fcache[(fid, rgb)] = len(fonts) - 1; return fcache[(fid, rgb)]

    def xf_child(s, rgb=None):
        if (s, rgb) in xcache: return xcache[(s, rgb)]
        x = copy.deepcopy(xfs[int(s or 0)]); x.set('fillId', str(FILL)); x.set('applyFill', '1')
        if rgb: x.set('fontId', str(font_with(int(x.get('fontId', '0')), rgb))); x.set('applyFont', '1')
        xfs.append(x); xfs.set('count', str(len(xfs))); xcache[(s, rgb)] = len(xfs) - 1; return xcache[(s, rgb)]

    def unshare(sd):
        masters = {}
        for row in sd:
            for c in row:
                fx = c.find(q('f'))
                if fx is not None and fx.get('t') == 'shared' and fx.text: masters[fx.get('si')] = (fx.text, c.get('r'))
        for row in sd:
            for c in row:
                fx = c.find(q('f'))
                if fx is not None and fx.get('t') == 'shared':
                    txt, org = masters[fx.get('si')]
                    if not fx.text: fx.text = Translator('=' + txt, origin=org).translate_formula(c.get('r'))[1:]
                    for a in ('t', 'si', 'ref'):
                        if a in fx.attrib: del fx.attrib[a]

    report = []; TVI = tv_index(pkg)
    for name, p in pkg.sheets:
        root = pkg.root(name); sd = root.find(q('sheetData')); changed = name in plan
        if changed: unshare(sd)
        bad = 0
        for row in sd:
            for c in row:
                fx = c.find(q('f'))
                if fx is not None and fx.text:
                    if not rt_ok(fx.text): bad += 1; continue
                    nf = shift_formula(fx.text, name)
                    if nf != fx.text: fx.text = nf; changed = True
        if bad: print('CẢNH BÁO: %s có %d công thức tokenizer không tái tạo được, không dời' % (name, bad))
        if name in plan:
            C = detect_cols(pkg, name)
            miss = [h for h in HELPERS if not C[h]]
            if miss: raise SystemExit('%s thiếu cột phụ %s - chạy lệnh helpers trước' % (name, miss))
            nm, cd, un, vl0, nc0 = C['name'], C['code'], C['unit_tdi'], C['vl0'], C['nc0']
            hg, hc, hn, hk, hl = (C[h] for h in HELPERS[:5]); mcl, tcl = C['Mã tra cứu'], C['Tên tra cứu']
            KL, VL, NC = CL(hk), CL(vl0), CL(nc0)
            rows = {int(r.get('r')): r for r in sd}
            for sh, rn, exp, ch, note, conf in P:
                if sh != name: continue
                desc = next((pkg.val(c) for c in rows[rn] if c.get('r') == 'B%d' % rn), None)
                if exp and not str(desc or '').strip().startswith(exp):
                    raise SystemExit('Dòng cha lệch %s!%d: cần "%s", thấy "%s"' % (sh, rn, exp, desc))
            for r in sorted(rows, reverse=True):
                nr = fmap(name, r)
                if nr != r:
                    rows[r].set('r', str(nr))
                    for c in rows[r]: c.set('r', re.sub(r'\d+$', str(nr), c.get('r')))
            for sh, rn, exp, ch, note, conf in P:
                if sh != name: continue
                prow = rows[rn]; pr = int(prow.get('r')); idx = list(sd).index(prow)
                first, last = pr + 1, pr + len(ch)
                for j, c in enumerate(ch):
                    nr = pr + 1 + j
                    row = etree.Element(q('row')); row.set('r', str(nr))
                    if prow.get('spans'): row.set('spans', prow.get('spans'))
                    for pc in prow:
                        col = colnum(pc.get('r'))
                        ncell = etree.SubElement(row, q('c')); ncell.set('r', CL(col) + str(nr))
                        if col >= nm:   # chỉ tô phần TDI
                            ncell.set('s', str(xf_child(pc.get('s'), COLOR_NAME if col == nm else COLOR_CODE if col == cd else None)))
                        elif pc.get('s'):
                            ncell.set('s', pc.get('s'))
                    if prow.get('s'): row.set('s', prow.get('s')); row.set('customFormat', '1')
                    setc(row, nm, c['name'])
                    ctext = new_code_text(TVI, c['g'], c.get('ma') or ('(mã 0) ' + c['val'] if c.get('key') == 'Tên công việc' else ''))
                    setc(row, cd, ctext); set_helpers(row, cd, mcl, tcl, ctext)
                    setc(row, un, c.get('unit'))
                    if c.get('key'):
                        for col, lc in ((vl0, 'Giá VL gốc'), (nc0, 'Giá NC gốc')):
                            setc(row, col, formula=tv_sumifs(lc, c['g'], c['key'], c['val'], nr, hg, c['g'], cd, ctext, mcl, tcl))
                    setc(row, hg, c['g']); setc(row, hc, c['conf'])
                    if c.get('note'): setc(row, hn, c['note'])
                    setc(row, hk, num=c.get('coef', 1)); setc(row, hl, 'Con')
                    sd.insert(idx + 1 + j, row)
                setc(prow, cd, ''); set_helpers(prow, cd, mcl, tcl, '')
                for col, L in ((vl0, VL), (nc0, NC)):
                    setc(prow, col, formula='SUMPRODUCT($%s$%d:$%s$%d,%s%d:%s%d)' % (KL, first, KL, last, L, first, L, last))
                confs = [c['conf'] for c in ch]
                pconf = max(confs, key=lambda x: CONF_ORDER.index(x) if x in CONF_ORDER else 3)
                if 'Chưa khớp' in confs and len(set(confs)) > 1:
                    pconf = 'Thấp'; note += '; còn thành phần chưa có giá, đơn giá cụm chưa đủ'
                setc(prow, hg, ' + '.join(dict.fromkeys(c['g'] for c in ch))); setc(prow, hc, pconf)
                setc(prow, hn, note + '. Đơn giá = SUMPRODUCT(Hệ số ĐV, đơn giá) các dòng con bên dưới')
                setc(prow, hk, ''); setc(prow, hl, 'Cụm')
                report.append((name, rn, pr, len(ch)))
            for mc in root.iter(q('mergeCell')): mc.set('ref', shift_ref(mc.get('ref'), name))
            for cf in root.iter(q('conditionalFormatting')):
                cf.set('sqref', ' '.join(shift_ref(x, name) for x in cf.get('sqref').split()))
                for fo in cf.iter(q('formula')): fo.text = shift_formula(fo.text, name)
            for dv in root.iter(q('dataValidation')):   # droplist Mã hiệu: dời vùng áp dụng + ô tham chiếu Nhóm VT
                dv.set('sqref', ' '.join(shift_ref(x, name) for x in dv.get('sqref').split()))
                for fo in list(dv.iter(q('formula1'))) + list(dv.iter(q('formula2'))):
                    if fo.text: fo.text = shift_formula(fo.text, name)
            for tag in ('autoFilter', 'dimension'):
                for e in root.iter(q(tag)): e.set('ref', shift_ref(e.get('ref'), name))
            for e in root.iter(q('selection')):
                if e.get('sqref'): e.set('sqref', ' '.join(shift_ref(x, name) for x in e.get('sqref').split()))
                if e.get('activeCell'): e.set('activeCell', shift_ref(e.get('activeCell'), name))
        if changed: pkg.touch(name)
    for dn in pkg.wb.iter(q('definedName')):
        if dn.text and '!' in dn.text and not dn.text.startswith('{'):
            try:
                if rt_ok(dn.text): dn.text = shift_formula(dn.text, '')
                else: print('CẢNH BÁO: không dời được defined name', dn.get('name'))
            except Exception as e:
                print('CẢNH BÁO: defined name %s: %s' % (dn.get('name'), e))
    pkg.new['xl/styles.xml'] = etree.tostring(st, xml_declaration=True, encoding='UTF-8', standalone=True)
    _drop_calcchain(pkg)
    pkg.save(out, drop=('xl/calcChain.xml',))
    json.dump(report, open(out + '.insert_report.json', 'w'), ensure_ascii=False)
    print('Đã bóc tách %d cụm, %d dòng con. Bản đồ dòng: %s.insert_report.json' % (len(report), sum(r[3] for r in report), out))


def _drop_calcchain(pkg):
    rels = etree.fromstring(pkg.new.get('xl/_rels/workbook.xml.rels') or pkg.files['xl/_rels/workbook.xml.rels'])
    for r in list(rels):
        if 'calcChain' in r.get('Target'): rels.remove(r)
    pkg.new['xl/_rels/workbook.xml.rels'] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)
    ct = (pkg.new.get('[Content_Types].xml') or pkg.files['[Content_Types].xml']).decode('utf8')
    pkg.new['[Content_Types].xml'] = re.sub(r'<Override[^>]*calcChain[^>]*/>', '', ct).encode('utf8')


# ---------------------------------------------------------------- fill CĐT
def cmd_fillcdt(f, out, cfgfile):
    CFG = json.load(open(cfgfile)); pkg = Pkg(f)
    for sh, c in CFG.items():
        C = detect_cols(pkg, sh)
        VL, NC, LO = CL(C['vl']), CL(C['nc']), CL(C['Loại dòng']) if C['Loại dòng'] else None
        Hc, Ic, Jc, U = c.get('h', 'H'), c.get('i', 'I'), c.get('j', 'J'), c.get('unit', 'C')
        sd = pkg.root(sh).find(q('sheetData')); rows = {int(r.get('r')): r for r in sd}
        get = lambda r, col: next((pkg.val(x) for x in rows.get(r, []) if x.get('r') == '%s%d' % (col, r)), None)
        n, over, tot = 0, [], None
        if c.get('single'):   # CĐT chỉ có 1 cột đơn giá + 1 cột thành tiền (vd. MTT: E=KL, F=Đơn giá, G=Thành tiền)
            Fc, Gc, KLc = c['single'].get('f', 'F'), c['single'].get('g', 'G'), c['kl'][0]
            DG = CL(C['dg']); gadd = []
            for r in sorted(rows):
                if r < c.get('start', 8): continue
                if str(get(r, U) or '').strip() in ('', '\xa0'): continue
                if LO and get(r, LO) == 'Con': continue
                v = get(r, Fc)
                if v not in (None, '') and not str(v).startswith('=N('): over.append('%s%d=%s' % (Fc, r, str(v)[:30]))
                setc(rows[r], CI(Fc), formula='N(%s%d)' % (DG, r))   # số (0 khi chưa có giá) để G = E*F không ra #VALUE!
                gc = next((x for x in rows[r] if x.get('r') == '%s%d' % (Gc, r)), None)
                if gc is None or (gc.find(q('f')) is None and gc.find(q('v')) is None and gc.find(q('is')) is None):
                    setc(rows[r], CI(Gc), formula='%s%d*%s%d' % (KLc, r, Fc, r)); gadd.append(r)
                n += 1
            pkg.touch(sh)
            print('%s: %d dòng, ô CĐT bị ghi đè %d %s, thêm công thức %s ở %s' % (sh, n, len(over), over[:15], Gc, gadd))
            continue
        if c.get('amount'):   # CĐT: 3 cột đơn giá (Vật tư, Nhân công, Tổng) + 1 cột Thành tiền (TBA, MFĐ, BOQ 03–08 của TDI25538)
            Ac, KLc = c['amount'], c['kl'][0]; lab = c.get('total_label', 'TỔNG CỘNG TRƯỚC THUẾ').upper(); lc = c.get('label', 'B')
            for r in sorted(rows):
                if r < c.get('start', 6): continue
                if str(get(r, lc) or '').strip().upper().startswith(lab): tot = r; continue
                if str(get(r, U) or '').strip() in ('', '\xa0'): continue
                if LO and get(r, LO) == 'Con': continue
                for col in (Hc, Ic, Jc, Ac):
                    v = get(r, col)
                    if v not in (None, '') and not str(v).startswith('=IF('): over.append('%s%d=%s' % (col, r, str(v)[:30]))
                row = rows[r]
                setc(row, CI(Hc), formula='IF(N(%s%d)=0,"",%s%d)' % (VL, r, VL, r))
                setc(row, CI(Ic), formula='IF(N(%s%d)=0,"",%s%d)' % (NC, r, NC, r))
                setc(row, CI(Jc), formula='IF(COUNT(%s%d:%s%d)=0,"",SUM(%s%d:%s%d))' % (Hc, r, Ic, r, Hc, r, Ic, r))
                setc(row, CI(Ac), formula='IF(OR(%s%d="",N(%s%d)=0),"",%s%d*%s%d)' % (Jc, r, KLc, r, Jc, r, KLc, r))
                n += 1
            note = ''
            if tot:
                cur = get(tot, Ac)
                if cur in (None, ''): setc(rows[tot], CI(Ac), formula='SUBTOTAL(9,%s%d:%s%d)' % (Ac, c.get('start', 6), Ac, tot - 1)); note = ', dòng tổng %d: thêm SUBTOTAL' % tot
                else: note = ', dòng tổng %d giữ nguyên (%s)' % (tot, str(cur)[:30])
            else: note = ', không thấy dòng "%s"' % lab
            pkg.touch(sh)
            print('%s: %d dòng, ô CĐT bị ghi đè %d %s%s' % (sh, n, len(over), over[:15], note))
            continue
        for r in sorted(rows):
            if r < c.get('start', 8): continue
            if str(get(r, CL(C['name'])) or '').startswith('Tổng cộng trước thuế') or str(get(r, CL(C['unit_tdi'])) or '').startswith('Tổng cộng trước thuế'):
                tot = r; continue
            if str(get(r, U) or '').strip() == '': continue
            if LO and get(r, LO) == 'Con': continue
            for col in (Hc, Ic):
                v = get(r, col)
                if v not in (None, '') and not str(v).startswith('=IF(N('): over.append('%s%d=%s' % (col, r, str(v)[:30]))
            row = rows[r]
            setc(row, CI(Hc), formula='IF(N(%s%d)=0,"",%s%d)' % (VL, r, VL, r))
            setc(row, CI(Ic), formula='IF(N(%s%d)=0,"",%s%d)' % (NC, r, NC, r))
            setc(row, CI(Jc), formula='IF(COUNT(%s%d:%s%d)=0,"",SUM(%s%d:%s%d))' % (Hc, r, Ic, r, Hc, r, Ic, r))
            for kl, z in zip(c['kl'], c['zones']):
                a = CI(z); A, B, S = CL(a), CL(a + 1), CL(a + 2)
                setc(row, a, formula='IF(OR(%s%d="",N(%s%d)=0),"",%s%d*%s%d)' % (Hc, r, kl, r, Hc, r, kl, r))
                setc(row, a + 1, formula='IF(OR(%s%d="",N(%s%d)=0),"",%s%d*%s%d)' % (Ic, r, kl, r, Ic, r, kl, r))
                setc(row, a + 2, formula='IF(COUNT(%s%d:%s%d)=0,"",SUM(%s%d:%s%d))' % (A, r, B, r, A, r, B, r))
            n += 1
        if tot:
            for z in c['zones']:
                for k in range(3):
                    L = CL(CI(z) + k); setc(rows[tot], CI(L), formula='SUBTOTAL(9,%s%d:%s%d)' % (L, c.get('start', 8), L, tot - 1))
        pkg.touch(sh)
        print('%s: %d dòng, dòng tổng %s, ô CĐT bị ghi đè %d %s' % (sh, n, tot, len(over), over[:15]))
        if len(c['zones']) != len(c['kl']): print('  LƯU Ý: số khu giá trị khác số cột KL - kiểm tra lại cấu hình')
    pkg.save(out)


# ---------------------------------------------------------------- dropdown Mã hiệu theo Nhóm VT
DS_NAME, ALL_LABEL = 'DS_MaHieu', '(Tất cả)'
WS_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
WS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'


def _cellv(pkg, c):
    v = pkg.val(c)
    if isinstance(v, str) and v.startswith('='):      # công thức: lấy giá trị cache
        x = c.find(q('v')); v = x.text if x is not None else None
    return v


def _tv_table(pkg):
    """(sheet, tablePart, tableRoot) của bảng TV."""
    for sh, p in pkg.sheets:
        rp = p.replace('worksheets/', 'worksheets/_rels/') + '.rels'
        if rp not in pkg.files: continue
        for r in etree.fromstring(pkg.files[rp]):
            if not r.get('Type', '').endswith('/table'): continue
            t = r.get('Target')
            tp = t.lstrip('/') if t.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(p), t))
            tb = etree.fromstring(pkg.files[tp])
            if 'TV' in (tb.get('name'), tb.get('displayName')): return sh, tp, tb
    raise SystemExit('Không thấy bảng TV (sheet "2 Thu vien gia")')


def _tv_iter(pkg):
    """[(Loại vật liệu, mã hiệu, tên công việc)] theo thứ tự bảng TV (mã '' nếu không mã, '0' nếu mã 0)."""
    from openpyxl.utils.cell import range_boundaries
    sh, tp, tb = _tv_table(pkg)
    c1, r1, c2, r2 = range_boundaries(tb.get('ref'))
    names = [tc.get('name') for tc in tb.find(q('tableColumns'))]
    cg, cm, ct = c1, c1 + names.index('Mã hiệu'), c1 + names.index('Tên công việc')
    out = []
    for row in pkg.root(sh).find(q('sheetData')):
        rn = int(row.get('r'))
        if rn <= r1 or rn > r2: continue
        g = m = n = None
        for c in row:
            cc = colnum(c.get('r'))
            if cc == cg: g = _cellv(pkg, c)
            elif cc == cm: m = _cellv(pkg, c)
            elif cc == ct: n = _cellv(pkg, c)
        g, m, n = str(g or '').strip(), str(m or '').strip(), str(n or '')
        if g and (m or n): out.append((g, m, n))
    return out


def entry_text(code, name):
    """Chuỗi hiển thị / lưu trong ô Mã hiệu: 'Mã | Tên công việc'."""
    lab = '(không mã)' if code == '' else '(mã 0)' if code == '0' else code
    return '%s%s%s' % (lab, SEP, name)


def _read_tv(pkg):
    """OrderedDict {Loại vật liệu: ['Mã | Tên', ...]} theo thứ tự TV, bỏ trùng."""
    out = collections.OrderedDict(); seen = collections.defaultdict(set)
    for g, m, n in _tv_iter(pkg):
        if not n: continue
        e = entry_text(m, n)
        if e not in seen[g]: seen[g].add(e); out.setdefault(g, []).append(e)
    return out


class TVIndex:
    def __init__(self, rows):
        self.byc = collections.defaultdict(list); self.byn = collections.defaultdict(list)
        for g, m, n in rows:
            if m not in ('', '0'): self.byc[(g, m)].append(n)
            self.byn[(g, n)].append(m)


def tv_index(pkg): return TVIndex(_tv_iter(pkg))


def new_code_text(TVI, g, legacy):
    """Mã dạng cũ (mã thuần hoặc '(mã 0) Tên') -> ô Mã hiệu dạng mới 'Mã | Tên'. Không tra được thì giữ nguyên."""
    t = str(legacy or '')
    if not t.strip(): return ''
    if SEP in t: return t
    if t.startswith('(mã 0) '):
        n = t[7:]; cs = set(TVI.byn.get((g, n), []))
        if not cs: return t
        if len(cs) == 1:
            c = next(iter(cs)); return entry_text(c, n)
        return entry_text('0', n)
    ns = set(TVI.byc.get((g, t), []))
    if len(ns) == 1: return entry_text(t, next(iter(ns)))
    return t


def _ds_sheet_xml(groups):
    cols = collections.OrderedDict()
    allc = []; seen = set()
    for g, ms in groups.items():
        for m in ms:
            if m not in seen: seen.add(m); allc.append(m)
    cols[ALL_LABEL] = allc                      # cột A = tất cả (dùng khi Nhóm VT trống / ghép cụm / Chưa xác định)
    cols.update(groups)
    root = etree.Element(q('worksheet'), nsmap={None: NS})
    sd = etree.SubElement(root, q('sheetData'))
    n = max(len(v) for v in cols.values()) + 2
    rows = [etree.SubElement(sd, q('row'), r=str(i)) for i in range(1, n + 1)]
    for ci, (g, ms) in enumerate(cols.items(), 1):
        L = CL(ci)
        setc(rows[0], ci, g)
        setc(rows[1], ci, formula='COUNTA(%s3:%s20000)' % (L, L))
        for i, m in enumerate(ms): setc(rows[2 + i], ci, m)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True), list(cols)


def _ensure_ds(pkg, xml):
    if DS_NAME in pkg.path:
        pkg.new[pkg.path[DS_NAME]] = xml; pkg._roots.pop(pkg.path[DS_NAME], None); return
    n = 1
    while 'xl/worksheets/sheet%d.xml' % n in pkg.files or 'xl/worksheets/sheet%d.xml' % n in pkg.new: n += 1
    path = 'xl/worksheets/sheet%d.xml' % n
    RP = 'xl/_rels/workbook.xml.rels'
    rels = etree.fromstring(pkg.new.get(RP) or pkg.files[RP])
    ids = {r.get('Id') for r in rels}; k = 1
    while 'rIdDS%d' % k in ids: k += 1
    rid = 'rIdDS%d' % k
    etree.SubElement(rels, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship',
                     Id=rid, Type=WS_REL, Target='worksheets/sheet%d.xml' % n)
    pkg.new[RP] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)
    shs = pkg.wb.find(q('sheets'))
    sid = max(int(s.get('sheetId')) for s in shs) + 1
    e = etree.SubElement(shs, q('sheet'), name=DS_NAME, sheetId=str(sid), state='hidden'); e.set('{%s}id' % RNS, rid)
    ct = (pkg.new.get('[Content_Types].xml') or pkg.files['[Content_Types].xml']).decode('utf8')
    ct = ct.replace('</Types>', '<Override PartName="/%s" ContentType="%s"/></Types>' % (path, WS_CT))
    pkg.new['[Content_Types].xml'] = ct.encode('utf8')
    pkg.new[path] = xml; pkg.sheets.append((DS_NAME, path)); pkg.path[DS_NAME] = path


def _dv_formula(gcol, r):
    m = 'IFERROR(MATCH($%s%d,%s!$1:$1,0),1)' % (CL(gcol), r, DS_NAME)
    return 'OFFSET(%s!$A$3,0,%s-1,INDEX(%s!$2:$2,%s),1)' % (DS_NAME, m, DS_NAME, m)   # < 255 ký tự (giới hạn của Excel)


def cmd_dropdown(f, out, sheets):
    pkg = Pkg(f)
    groups = _read_tv(pkg)
    xml, order = _ds_sheet_xml(groups)
    _ensure_ds(pkg, xml)
    print('DS_MaHieu: %d nhóm, %d mã hiệu (bỏ trùng) - %s' % (len(groups), sum(len(v) for v in groups.values()), ', '.join('%s=%d' % (g, len(v)) for g, v in groups.items())))
    for sh in sheets:
        C = detect_cols(pkg, sh)
        if not C['Nhóm VT']: raise SystemExit('%s chưa có cột Nhóm VT - chạy lệnh helpers trước' % sh)
        cd, hg = C['code'], C['Nhóm VT']
        root = pkg.root(sh); sd = root.find(q('sheetData'))
        rows = list(sd); first, last = C['hdr_row'] + 1, max(int(r.get('r')) for r in rows)
        # nhóm trong sheet không có trong TV -> dùng danh sách (Tất cả)
        miss = collections.Counter()
        for row in rows:
            for c in row:
                if colnum(c.get('r')) == hg and int(row.get('r')) >= first:
                    g = str(_cellv(pkg, c) or '').strip()
                    if g and g not in groups: miss[g] += 1
        dvs = root.find(q('dataValidations'))
        if dvs is None:
            dvs = etree.Element(q('dataValidations'))
            order_t = ('sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState',
                       'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting')
            pos = -1
            for i, ch in enumerate(root):
                if isinstance(ch.tag, str) and etree.QName(ch).localname in order_t: pos = i
            root.insert(pos + 1, dvs)
        removed = 0
        for dv in list(dvs):                                    # bỏ droplist cũ của skill + droplist khác chỉ nằm trong cột Mã hiệu (tránh chồng vùng)
            f1 = dv.find(q('formula1')); txt = f1.text if f1 is not None and f1.text else ''
            toks = dv.get('sqref', '').split()
            only_code = toks and all(re.fullmatch(r'%s\d+(:%s\d+)?' % (CL(cd), CL(cd)), t) for t in toks)
            if DS_NAME in txt or only_code: dvs.remove(dv); removed += 1
            elif any(colnum(re.match(r'[A-Z]+', t.split(':')[0]).group(0) + '1') <= cd <= colnum(re.match(r'[A-Z]+', t.split(':')[-1]).group(0) + '1') for t in toks):
                print('  CẢNH BÁO: %s có validation khác phủ cột %s (%s) - kiểm tra thủ công, có thể Excel báo lỗi chồng vùng' % (sh, CL(cd), dv.get('sqref')))
        dv = etree.SubElement(dvs, q('dataValidation'), type='list', errorStyle='warning', allowBlank='1',
                              showInputMessage='1', showErrorMessage='1',
                              errorTitle='Mã hiệu ngoài danh sách', error='Mục này không thuộc Loại vật liệu của dòng. Vẫn nhập tiếp nếu đã chắc chắn (giá sẽ ra 0 nếu không có trong TV). Dạng chuẩn: Mã | Tên công việc.',
                              promptTitle='Mã hiệu | Tên công việc', prompt='Chọn trong danh sách (mã và tên) của Loại vật liệu trên dòng này (cột %s).' % CL(hg),
                              sqref='%s%d:%s%d' % (CL(cd), first, CL(cd), last))
        etree.SubElement(dv, q('formula1')).text = _dv_formula(hg, first)
        for old in list(dvs):   # droplist cột Nhóm VT (Loại vật liệu): danh sách = tiêu đề các cột của DS_MaHieu, bỏ (Tất cả)
            f1 = old.find(q('formula1'))
            if f1 is not None and f1.text and f1.text.startswith('%s!$B$1' % DS_NAME): dvs.remove(old)
        dg = etree.SubElement(dvs, q('dataValidation'), type='list', errorStyle='warning', allowBlank='1', showErrorMessage='1',
                              errorTitle='Loại vật liệu ngoài danh sách', error='Tên không có trong thư viện (giá sẽ ra 0). Vẫn nhập tiếp nếu là Chưa xác định / dòng cụm.',
                              sqref='%s%d:%s%d' % (CL(hg), first, CL(hg), last))
        etree.SubElement(dg, q('formula1')).text = '%s!$B$1:$%s$1' % (DS_NAME, CL(len(order)))
        dvs.set('count', str(len(dvs)))
        pkg.touch(sh); _widen(pkg, sh, cd, 58)
        print('%s: droplist cột Mã hiệu %s%d:%s%d theo Nhóm VT cột %s (thay %d validation cũ)%s' % (
            sh, CL(cd), first, CL(cd), last, CL(hg), removed,
            '; nhóm không có trong TV (dùng (Tất cả)): %s' % dict(miss) if miss else ''))
    _drop_calcchain(pkg)
    pkg.save(out, drop=('xl/calcChain.xml',))
    print('Đã ghi', out)


# ---------------------------------------------------------------- migrate: Nhóm VT cũ (theo file) -> Loại vật liệu (theo sheet)
def _read_tv_full(pkg):
    """[(loại VL, nhóm file, mã hiệu, tên công việc)] từ bảng TV."""
    from openpyxl.utils.cell import range_boundaries
    for sh, p in pkg.sheets:
        rp = p.replace('worksheets/', 'worksheets/_rels/') + '.rels'
        if rp not in pkg.files: continue
        for r in etree.fromstring(pkg.files[rp]):
            if not r.get('Type', '').endswith('/table'): continue
            t = r.get('Target'); tp = t.lstrip('/') if t.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(p), t))
            tb = etree.fromstring(pkg.files[tp])
            if 'TV' not in (tb.get('name'), tb.get('displayName')): continue
            c1, r1, c2, r2 = range_boundaries(tb.get('ref'))
            names = [tc.get('name') for tc in tb.find(q('tableColumns'))]
            if 'Nhóm file' not in names: raise SystemExit('Bảng TV chưa có cột "Nhóm file": nạp lại thư viện bằng tv_loader.py v3 (load --replace) trước khi migrate')
            ix = {n: c1 + i for i, n in enumerate(names)}; out = []
            for row in pkg.root(sh).find(q('sheetData')):
                rn = int(row.get('r'))
                if rn <= r1 or rn > r2: continue
                v = {}
                for c in row:
                    cc = colnum(c.get('r'))
                    for n in (names[0], 'Nhóm file', 'Mã hiệu', 'Tên công việc'):
                        if cc == ix[n]: v[n] = str(_cellv(pkg, c) or '').strip()
                if v.get(names[0]): out.append((v[names[0]], v.get('Nhóm file', ''), v.get('Mã hiệu', ''), v.get('Tên công việc', '')))
            return out
    raise SystemExit('Không thấy bảng TV')


def cmd_migrate(f, out, sheets=None):
    """Sau tv_loader v3 (load --replace): đổi Nhóm VT của các dòng BOQ + SUMIFS viết cứng từ nhóm theo file (TCap) sang Loại vật liệu theo sheet (Thang cáp / Phụ kiện T.cáp).
    Quyết theo MÃ: mã chỉ có ở một loại VL của nhóm cũ -> đổi; mã có ở nhiều loại VL -> để nguyên và liệt kê (chọn tay)."""
    pkg = Pkg(f); tv = _read_tv_full(pkg)
    new_names = {t[0] for t in tv}
    OM = collections.defaultdict(set); NEWOF = collections.defaultdict(set)
    for lv, old, ma, ten in tv:
        NEWOF[old].add(lv)
        if ma and ma != '0': OM[(old, 'Mã hiệu', ma)].add(lv)
        if ten: OM[(old, 'Tên công việc', ten)].add(lv)
    def mapg(old, key, val):
        if old in new_names: return old, 'ok'
        c = OM.get((old, key, val)) if key else None
        if c and len(c) == 1: return next(iter(c)), 'đổi'
        if c and len(c) > 1: return old, 'mơ hồ'
        n = NEWOF.get(old)
        if n and len(n) == 1: return next(iter(n)), 'đổi'
        return old, 'không đổi được'
    SUM = re.compile(r'SUMIFS\(TV\[(Giá [^\]]*)\],TV\[(?:Column1|Loại vật liệu)\],"((?:[^"]|"")*)",TV\[(Mã hiệu|Tên công việc)\],("(?:[^"]|"")*"|\$?[A-Z]+\$?\d+)\)')
    stat = collections.Counter(); amb = []
    for sh, p in list(pkg.sheets):
        if sh == DS_NAME: continue
        root = pkg.root(sh); sd = root.find(q('sheetData')); changed = False
        try: C = detect_cols(pkg, sh)
        except SystemExit: C = None
        hg = C['Nhóm VT'] if C else None; cd = C['code'] if C else None
        for row in sd:
            rn = int(row.get('r')); code = gcell = None
            if C and rn > C['hdr_row']:
                for c in row:
                    cc = colnum(c.get('r'))
                    if cc == cd: code = str(_cellv(pkg, c) or '').strip()
                    if cc == hg: gcell = c
            key = val = None
            if code: key, val = ('Tên công việc', code[7:]) if code.startswith('(mã 0) ') else ('Mã hiệu', code)
            for c in row:   # công thức viết cứng + đổi tên cột trong mọi công thức
                fx = c.find(q('f'))
                if fx is None or not fx.text or 'TV[' not in fx.text: continue
                def sub(m):
                    g = m.group(2).replace('""', '"'); k = m.group(3); v = m.group(4)
                    v = v[1:-1].replace('""', '"').replace('~~', '~') if v.startswith('"') else code
                    ng, how = mapg(g, k, v); stat['SUMIFS ' + how] += 1
                    if how == 'mơ hồ': amb.append('%s!%s SUMIFS %s/%s' % (sh, c.get('r'), g, v))
                    return m.group(0).replace('"%s"' % m.group(2), '"%s"' % ng.replace('"', '""'), 1).replace('TV[Column1]', 'TV[%s]' % GCOL)
                nf = SUM.sub(sub, fx.text).replace('TV[Column1]', 'TV[%s]' % GCOL)
                if nf != fx.text: fx.text = nf; changed = True
            if gcell is not None:
                gv = str(_cellv(pkg, gcell) or '').strip()
                if gv and gv not in new_names:
                    parts = [x.strip() for x in gv.split(' + ')]
                    if len(parts) > 1:
                        res = [mapg(x, None, None) for x in parts]; ng = ' + '.join(r[0] for r in res); how = 'đổi' if all(r[1] in ('ok', 'đổi') for r in res) else 'không đổi được'
                    else: ng, how = mapg(gv, key, val)
                    stat['Nhóm VT ' + how] += 1
                    if how == 'mơ hồ': amb.append('%s!%s%d Nhóm VT %s mã %s' % (sh, CL(hg), rn, gv, val))
                    if how == 'đổi': setc(row, hg, ng); changed = True
        if changed: pkg.touch(sh)
    _drop_calcchain(pkg); pkg.save(out, drop=('xl/calcChain.xml',))
    print('migrate:', dict(stat)); print('MƠ HỒ (mã có ở cả sheet chính và phụ kiện, chọn tay): %d' % len(amb)); print('\n'.join(amb[:60]))


# ---------------------------------------------------------------- recalc check
def cmd_recalc(f, sheets=None):
    import openpyxl
    tmp = tempfile.mkdtemp()
    nc = os.path.join(tmp, 'nocache.xlsx')
    zin = zipfile.ZipFile(f); zo = zipfile.ZipFile(nc, 'w', zipfile.ZIP_DEFLATED)
    pat = re.compile(r'(<f>[^<]*</f>|<f [^>]*>[^<]*</f>|<f [^>]*/>)<v>[^<]*</v>')
    for it in zin.infolist():
        d = zin.read(it.filename)
        if it.filename.startswith('xl/worksheets/sheet'): d = pat.sub(r'\1', d.decode('utf8')).encode('utf8')
        zo.writestr(it, d)
    zo.close()
    od = os.path.join(tmp, 'out'); os.makedirs(od)
    subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', od, nc], capture_output=True, timeout=600)
    wbv = openpyxl.load_workbook(os.path.join(od, 'nocache.xlsx'), data_only=True)
    wbf = openpyxl.load_workbook(f, read_only=False)
    pkg = Pkg(f)
    for sh in (sheets or [s for s in wbv.sheetnames if s in wbf.sheetnames]):
        ws, wf = wbv[sh], wbf[sh]
        errs = [(c.coordinate, c.value) for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith('#')]
        msg = '%s: ô lỗi %d %s' % (sh, len(errs), errs[:8])
        try:
            C = detect_cols(pkg, sh)
            no = []
            for r in range(C['hdr_row'] + 3, ws.max_row + 1):
                code = wf.cell(r, C['code']).value
                if code and str(code).strip() and not str(code).startswith('='):
                    a, b = ws.cell(r, C['vl0']).value, ws.cell(r, C['nc0']).value
                    if not (isinstance(a, (int, float)) and a) and not (isinstance(b, (int, float)) and b): no.append(r)
            msg += ' | có mã không giá %d %s' % (len(no), no[:15])
            for r in range(ws.max_row, 0, -1):
                if str(ws.cell(r, C['name']).value or ws.cell(r, C['unit_tdi']).value or '').startswith('Tổng cộng trước thuế'):
                    vals = [(CL(cc), ws.cell(r, cc).value) for cc in range(1, ws.max_column + 1) if isinstance(ws.cell(r, cc).value, (int, float)) and ws.cell(r, cc).value]
                    msg += ' | dòng tổng %d: %s' % (r, vals[:20]); break
        except SystemExit:
            pass
        print(msg)
    shutil.rmtree(tmp, ignore_errors=True)


# ---------------------------------------------------------------- cleanup temp sheets
def cmd_cleanup(f, out):
    pkg = Pkg(f)
    shs = pkg.wb.find(q('sheets'))
    rels = etree.fromstring(pkg.files['xl/_rels/workbook.xml.rels'])
    drop = set(); removed = []
    for s in list(shs):
        n = s.get('name')
        if n == '_TDI_JOB' or n.startswith('_TDI_TMP'):
            rid = s.get('{%s}id' % RNS); shs.remove(s); removed.append(n)
            for r in list(rels):
                if r.get('Id') == rid:
                    t = r.get('Target').lstrip('/'); t = t if t.startswith('xl/') else 'xl/' + t
                    drop.add(t); rels.remove(r)
                    rp = t.replace('worksheets/', 'worksheets/_rels/') + '.rels'
                    drop.add(rp)
    if not removed: print('Không có sheet tạm'); shutil.copy(f, out); return
    # defined names gắn localSheetId cần đánh lại -> bỏ các name của sheet tạm, đánh lại chỉ số
    names = [s.get('name') for s in shs]
    old = [n for n, _ in pkg.sheets]
    dns = pkg.wb.find(q('definedNames'))
    if dns is not None:
        for dn in list(dns):
            lid = dn.get('localSheetId')
            if lid is not None:
                on = old[int(lid)]
                if on in removed: dns.remove(dn)
                else: dn.set('localSheetId', str(names.index(on)))
            elif dn.text and any(r in dn.text for r in removed): dns.remove(dn)
    bv = pkg.wb.find(q('bookViews'))
    if bv is not None:
        for v in bv:
            for a in ('activeTab', 'firstSheet'):
                if v.get(a) and int(v.get(a)) >= len(names): v.set(a, '0')
    pkg.new['xl/_rels/workbook.xml.rels'] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)
    ct = pkg.files['[Content_Types].xml'].decode('utf8')
    for d in drop: ct = re.sub(r'<Override[^>]*PartName="/%s"[^>]*/>' % re.escape(d), '', ct)
    pkg.new['[Content_Types].xml'] = ct.encode('utf8')
    pkg.save(out, drop=drop)
    print('Đã xóa sheet tạm:', removed)


# ---------------------------------------------------------------- cột rộng cho ô Mã hiệu "Mã | Tên"
def _widen(pkg, sh, col, width):
    root = pkg.root(sh); cols = root.find(q('cols'))
    if cols is None:
        cols = etree.Element(q('cols'))
        sdp = [i for i, ch in enumerate(root) if isinstance(ch.tag, str) and etree.QName(ch).localname == 'sheetData'][0]
        root.insert(sdp, cols)
    done = False
    for c in list(cols):
        a, b = int(c.get('min')), int(c.get('max'))
        if a <= col <= b:
            if a == b == col: c.set('width', str(width)); c.set('customWidth', '1'); done = True; break
            parts = []
            if a < col: parts.append((a, col - 1))
            parts.append((col, col))
            if col < b: parts.append((col + 1, b))
            pos = list(cols).index(c); cols.remove(c)
            for k, (x, y) in enumerate(parts):
                n = etree.fromstring(etree.tostring(c)); n.set('min', str(x)); n.set('max', str(y))
                if x == col: n.set('width', str(width)); n.set('customWidth', '1')
                cols.insert(pos + k, n)
            done = True; break
    if not done:
        n = etree.Element(q('col')); n.set('min', str(col)); n.set('max', str(col)); n.set('width', str(width)); n.set('customWidth', '1')
        pos = len(cols)
        for i, c in enumerate(cols):
            if int(c.get('min')) > col: pos = i; break
        cols.insert(pos, n)


# ---------------------------------------------------------------- hsloai: nhân hệ số Loại VL (Summary) vào giá thư viện
TV_RAW = ['Giá VL thư viện (thô)', 'Giá NC thư viện (thô)', 'HS VL (Loại VL)', 'HS NC (Loại VL)']


def _add_ss(pkg, texts):
    RP = 'xl/sharedStrings.xml'
    root = etree.fromstring(pkg.new.get(RP) or pkg.files[RP]); idx = []
    for t in texts:
        si = etree.SubElement(root, q('si')); etree.SubElement(si, q('t')).text = t
        idx.append(len(root) - 1); pkg.ss.append(t)
    root.set('uniqueCount', str(len(root))); root.set('count', str(int(root.get('count', len(root))) + len(texts)))
    pkg.new[RP] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    return idx


def _lookup_hs(r, nm_col, val_col, nm2, val2):
    A = '$A%d' % r; W = 'SUBSTITUTE(%s," ","*")' % A
    rng = lambda c: 'Summary!$%s$11:$%s$40' % (c, c)
    def one(nc, vc, key): return 'INDEX(%s,MATCH(%s,%s,0))' % (rng(vc), key, rng(nc))
    return ('IFERROR(%s,IFERROR(%s,IFERROR(%s,IFERROR(%s,1))))' %
            (one(nm_col, val_col, A), one(nm2, val2, A), one(nm_col, val_col, W), one(nm2, val2, W)))


def cmd_hsloai(f, out):
    """Thêm 4 cột vào bảng TV: P/Q = giá thư viện thô, R/S = hệ số Loại VL (tra Summary B11:G24).
    E/F (Giá VL gốc / Giá NC gốc) thành công thức = thô x hệ số, nên mọi SUMIFS TV[Giá VL gốc]/TV[Giá NC gốc] tự nhân hệ số Loại VL."""
    from openpyxl.utils.cell import range_boundaries
    import copy
    pkg = Pkg(f)
    if 'Summary' not in pkg.path: raise SystemExit('Không thấy sheet Summary')
    sh, tp, tb = _tv_table(pkg)
    names = [tc.get('name') for tc in tb.find(q('tableColumns'))]
    if TV_RAW[0] in names: print('Bảng TV đã có cột hệ số Loại VL - không đổi'); shutil.copy(f, out); return
    c1, r1, c2, r2 = range_boundaries(tb.get('ref'))
    if (names.index('Giá VL gốc'), names.index('Giá NC gốc')) != (4, 5) or c1 != 1:
        raise SystemExit('Bố cục bảng TV khác dự kiến (E = Giá VL gốc, F = Giá NC gốc)')
    nc = len(names); PC = nc + 1                       # cột P (nếu 15 cột)
    L = [CL(PC + i) for i in range(4)]                 # P Q R S
    root = pkg.root(sh); sd = root.find(q('sheetData'))
    ss = _add_ss(pkg, TV_RAW)
    n = 0
    for row in sd:
        rn = int(row.get('r'))
        if rn < r1 or rn > r2: continue
        row.set('spans', '1:%d' % (PC + 3))
        if rn == r1:
            for i, k in enumerate(ss):
                c = setc(row, PC + i, 'x'); c.set('t', 's')
                for ch in list(c): c.remove(ch)
                etree.SubElement(c, q('v')).text = str(k)
            continue
        E = next((c for c in row if colnum(c.get('r')) == 5), None); F = next((c for c in row if colnum(c.get('r')) == 6), None)
        for src, dstc in ((E, PC), (F, PC + 1)):
            if src is not None and src.find(q('v')) is not None:
                d = copy.deepcopy(src); d.set('r', '%s%d' % (CL(dstc), rn))
                pos = len(row)
                for i, c in enumerate(row):
                    if colnum(c.get('r')) > dstc: pos = i; break
                row.insert(pos, d)
        setc(row, PC + 2, formula=_lookup_hs(rn, 'B', 'C', 'E', 'F'))
        setc(row, PC + 3, formula=_lookup_hs(rn, 'B', 'D', 'E', 'G'))
        setc(row, 5, formula='IF(%s%d="","",%s%d*%s%d)' % (L[0], rn, L[0], rn, L[2], rn))
        setc(row, 6, formula='IF(%s%d="","",%s%d*%s%d)' % (L[1], rn, L[1], rn, L[3], rn))
        n += 1
    ref = 'A%d:%s%d' % (r1, CL(PC + 3), r2)
    tb.set('ref', ref); tb.find(q('autoFilter')).set('ref', ref)
    tcs = tb.find(q('tableColumns')); mx = max(int(t.get('id')) for t in tcs)
    for i, nm in enumerate(TV_RAW): etree.SubElement(tcs, q('tableColumn'), id=str(mx + 1 + i), name=nm)
    tcs.set('count', str(len(tcs)))
    pkg.new[tp] = etree.tostring(tb, xml_declaration=True, encoding='UTF-8', standalone=True)
    dm = root.find(q('dimension'))
    if dm is not None: dm.set('ref', 'A1:%s%d' % (CL(PC + 3), r2))
    cols = root.find(q('cols'))
    if cols is not None:
        e = etree.SubElement(cols, q('col')); e.set('min', str(PC)); e.set('max', str(PC + 3)); e.set('width', '16'); e.set('customWidth', '1')
    pkg.touch(sh); _drop_calcchain(pkg)
    pkg.save(out, drop=('xl/calcChain.xml',))
    print('hsloai: %d dòng TV; %s = giá thô VL/NC, %s = hệ số Loại VL/NC; E/F = thô x hệ số. Đã ghi %s' % (n, '/'.join(L[:2]), '/'.join(L[2:]), out))


# ---------------------------------------------------------------- migrate2: ô Mã hiệu -> "Mã | Tên", công thức giá -> tra Loại VL + Mã + Tên
OLDSUM = re.compile(r'SUMIFS\(TV\[(Giá [VN][LC][^\]]*)\],TV\[(?:Column1|Loại vật liệu)\],(\$?[A-Z]+\$?\d+|"(?:[^"]|"")*"),TV\[(Mã hiệu|Tên công việc)\],(\$?[A-Z]+\$?\d+|"(?:[^"]|"")*")\)')


def cmd_migrate2(f, out, sheets=None):
    pkg = Pkg(f); TVI = tv_index(pkg); stat = collections.Counter(); left = []
    for sh, p in list(pkg.sheets):
        if sh == DS_NAME: continue
        try: C = detect_cols(pkg, sh)
        except SystemExit: continue
        if sheets and sh not in sheets: continue
        miss = [h for h in HELPERS if not C[h]]
        if miss: print(sh, 'bỏ qua: thiếu cột phụ', miss, '(chạy helpers trước)'); continue
        cd, hg, mc, tc = C['code'], C['Nhóm VT'], C['Mã tra cứu'], C['Tên tra cứu']
        root = pkg.root(sh); changed = False
        for row in root.find(q('sheetData')):
            rn = int(row.get('r'))
            if rn <= C['hdr_row']: continue
            cc = {colnum(c.get('r')): c for c in row}
            ccell = cc.get(cd)
            code = str(_cellv(pkg, ccell) or '').strip() if ccell is not None else ''
            if not code or code.startswith('='): continue
            g = str(_cellv(pkg, cc[hg]) or '').strip() if hg in cc else ''
            new = new_code_text(TVI, g, code)
            if SEP not in new:      # ô mã ghi tên trần / 'TV: tên' (bản cũ): lấy tên từ công thức tra theo Tên của chính dòng
                nm = set()
                for col in (C['vl0'], C['nc0']):
                    cl_ = cc.get(col); f_ = cl_.find(q('f')) if cl_ is not None else None
                    for m in OLDSUM.finditer(f_.text if f_ is not None and f_.text else ''):
                        if m.group(3) == 'Tên công việc' and m.group(4).startswith('"') and (not m.group(2).startswith('"') or unq_ref(m.group(2)) == g):
                            nm.add(re.sub(r'~([~*?])', r'\1', m.group(4)[1:-1].replace('""', '"')))
                if len(nm) == 1:
                    cand = new_code_text(TVI, g, '(mã 0) ' + next(iter(nm)))
                    if SEP in cand: new = cand
            if new != code:
                setc(row, cd, new); stat['đổi ô mã'] += 1
            else: stat['giữ ô mã'] += 1
            if SEP not in new: left.append('%s!%s%d %s | %s' % (sh, CL(cd), rn, g, code[:60]))
            set_helpers(row, cd, mc, tc, new)
            for col in (C['vl0'], C['nc0']):
                cell = cc.get(col); fx = cell.find(q('f')) if cell is not None else None
                if fx is None or not fx.text or 'TV[' not in fx.text: continue
                def sub(m):
                    gg = unq_ref(m.group(2)); key = m.group(3); v = m.group(4)
                    gg = g if not m.group(2).startswith('"') else gg
                    if v.startswith('"'): v = v[1:-1].replace('""', '"'); v = re.sub(r'~([~*?])', r'\1', v)
                    else: v = code if key == 'Mã hiệu' and not code.startswith('(mã 0)') else (code[7:] if code.startswith('(mã 0) ') else code)
                    return tv_sumifs(m.group(1), gg, key, v, rn, hg, g, cd, new, mc, tc)
                nf = OLDSUM.sub(sub, fx.text)
                if nf != fx.text: fx.text = nf; changed = True; stat['công thức giá'] += 1
        if changed or True: pkg.touch(sh)
        _widen(pkg, sh, cd, 58)
    _drop_calcchain(pkg); pkg.save(out, drop=('xl/calcChain.xml',))
    print('migrate2:', dict(stat)); print('ô mã chưa chuyển được sang dạng "Mã | Tên" (giữ dạng cũ, vẫn tra đúng theo mã): %d' % len(left)); print('\n'.join(left[:40]))



# ---------------------------------------------------------------- Bước 6c / 6e (thêm 21/10/2026: trước đây AI ghi tay bằng lxml)
NGUON = 'Nguồn giá'


def _wb_sheet_index(pkg, sheet):
    return [n for n, _ in pkg.sheets].index(sheet)


def _work_sheets(pkg, sheets):
    """Sheet BOQ đã có đủ cột phụ (Nhóm VT, Loại dòng, Mã tra cứu, Tên tra cứu); sheet khác bị bỏ qua và báo."""
    out = []
    for sh in (sheets or [n for n, _ in pkg.sheets]):
        try: c = detect_cols(pkg, sh)
        except SystemExit: 
            if sheets: print(sh, 'BỎ QUA: không nhận diện được cột TDI')
            continue
        miss = [h for h in ('Nhóm VT', 'Loại dòng', 'Mã tra cứu', 'Tên tra cứu') if not c[h]]
        if miss:
            if sheets: print(sh, 'BỎ QUA: thiếu cột phụ', miss)
            continue
        out.append((sh, c))
    return out


def _cell(row, col):
    for c in row:
        if colnum(c.get('r')) == col: return c
    return None


def _blank(c):
    return c is None or (c.find(q('f')) is None and c.find(q('v')) is None and c.find(q('is')) is None)


def _is_work_row(pkg, row, c, kinds):
    kc = _cell(row, c['Loại dòng'])
    if kc is None or pkg.val(kc) not in kinds: return False
    uc = _cell(row, c['unit_tdi'])
    return uc is not None and not _blank(uc)


def _wait_formula(lc, r, c):
    L = lambda k: CL(c[k])
    return ('IF(OR($%s%d="",AND($%s%d="",$%s%d="")),0,SUMIFS(TV[%s],TV[%s],IF(OR($%s%d="",$%s%d="Chưa xác định"),"<>§",$%s%d),'
            'TV[Mã hiệu],IF($%s%d="","<>§",$%s%d),TV[Tên công việc],IF($%s%d="","<>§",$%s%d)))') % (
        L('code'), r, L('Mã tra cứu'), r, L('Tên tra cứu'), r, lc, GCOL, L('Nhóm VT'), r, L('Nhóm VT'), r, L('Nhóm VT'), r,
        L('Mã tra cứu'), r, L('Mã tra cứu'), r, L('Tên tra cứu'), r, L('Tên tra cứu'), r)


def _std_formula(lc, r, c):
    L = lambda k: CL(c[k])
    return 'SUMIFS(TV[%s],TV[%s],$%s%d,TV[Mã hiệu],IF($%s%d="","<>§",$%s%d),TV[Tên công việc],IF($%s%d="","<>§",$%s%d))' % (
        lc, GCOL, L('Nhóm VT'), r, L('Mã tra cứu'), r, L('Mã tra cứu'), r, L('Tên tra cứu'), r, L('Tên tra cứu'), r)


def cmd_cho(f, out, sheets=None):
    """Bước 6e (Quy tắc 34): công thức chờ cho ô giá gốc VL / NC trống của dòng Dòng / Con; bọc chặn mã trống cho SUMIFS chuẩn;
    công thức cột Mã tra cứu / Tên tra cứu nếu đang trống. Không động vào Cụm, ô số cứng, công thức khác (liệt kê ô bỏ qua)."""
    pkg = Pkg(f)
    for sh, c in _work_sheets(pkg, sheets):
        stat = collections.Counter(); skipped = []
        for row in pkg.root(sh).find(q('sheetData')):
            r = int(row.get('r'))
            if r <= c['hdr_row'] or not _is_work_row(pkg, row, c, ('Dòng', 'Con')): continue
            for key, lc in (('vl0', 'Giá VL gốc'), ('nc0', 'Giá NC gốc')):
                cell = _cell(row, c[key])
                if _blank(cell):
                    cc = _cell(row, c['code'])
                    if cc is not None and not _blank(cc) and str(pkg.val(cc) or '').strip():
                        # dòng ĐÃ có mã mà ô giá cố ý trống (chỉ NC / DC, thiếu giá VL): nối công thức sẽ tự kéo giá TV vào, đổi tổng -> giữ trống
                        stat['trống có mã (giữ)'] += 1; skipped.append(CL(c[key]) + str(r)); continue
                    setc(row, c[key], formula=_wait_formula(lc, r, c)); stat['chờ'] += 1; continue
                fx = cell.find(q('f'))
                if fx is not None and fx.get('t') != 'shared' and fx.text == _std_formula(lc, r, c):
                    setc(row, c[key], formula='IF(OR($%s%d="",AND($%s%d="",$%s%d="")),0,%s)' % (
                        CL(c['code']), r, CL(c['Mã tra cứu']), r, CL(c['Tên tra cứu']), r, fx.text)); stat['bọc'] += 1
                elif fx is not None and fx.text and fx.text.startswith('IF(OR($%s%d=""' % (CL(c['code']), r)): stat['đã có'] += 1
                else: stat['bỏ qua'] += 1; skipped.append(CL(c[key]) + str(r))
            # cột phụ Mã / Tên tra cứu: chỉ ghi khi đang trống hoặc chữ tĩnh, theo đúng dạng công thức của dòng đã gán
            for key, fn in (('Mã tra cứu', hf_code), ('Tên tra cứu', hf_name)):
                cell = _cell(row, c[key])
                if cell is None or cell.find(q('f')) is None:
                    setc(row, c[key], formula=fn(CL(c['code']), r)); stat['cột phụ'] += 1
        pkg.touch(sh)
        print('%s: %s%s' % (sh, dict(stat), (' | bỏ qua: ' + ','.join(skipped[:12])) if skipped else ''))
    pkg.save(out)


def cmd_nguongia(f, out, sheets=None):
    """Bước 6c: cột 'Nguồn giá' (tạo nếu chưa có), công thức Giá thư viện / Chưa có giá cho mọi dòng công việc (Dòng, Con, Cụm),
    giữ nguyên chữ tĩnh 'Giá tham khảo' / 'NCC'; nới autoFilter và _xlnm._FilterDatabase tới cột Nguồn giá."""
    pkg = Pkg(f)
    for sh, c in _work_sheets(pkg, sheets):
        sd = pkg.root(sh).find(q('sheetData')); rows = {int(r.get('r')): r for r in sd}
        hrow = rows[c['hdr_row']]; ng = None
        for x in hrow:
            if str(pkg.val(x) or '').strip() == NGUON and colnum(x.get('r')) > c['name']: ng = colnum(x.get('r'))
        if ng is None:
            ng = max(colnum(x.get('r')) for x in hrow if not _blank(x)) + 1
            sty = next((x.get('s') for x in hrow if colnum(x.get('r')) == c['Tên tra cứu']), None)
            setc(hrow, ng, NGUON, style=sty); created = True
        else: created = False
        stat = collections.Counter()
        for row in sd:
            r = int(row.get('r'))
            if r <= c['hdr_row'] or not _is_work_row(pkg, row, c, ('Dòng', 'Con', 'Cụm')): continue
            old = _cell(row, ng); ov = pkg.val(old) if old is not None else None
            if ov in ('Giá tham khảo', 'NCC'): stat['giữ ' + ov] += 1; continue
            setc(row, ng, formula='IF(N($%s%d)+N($%s%d)>0,"Giá thư viện","Chưa có giá")' % (CL(c['vl0']), r, CL(c['nc0']), r)); stat['ghi'] += 1
        # autoFilter
        af = pkg.root(sh).find(q('autoFilter')); msg = 'không có autoFilter (không tự tạo, báo người dùng)'
        if af is not None:
            m = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)$', af.get('ref'))
            if CI(m.group(3)) < ng:
                new = '%s%s:%s%s' % (m.group(1), m.group(2), CL(ng), m.group(4)); af.set('ref', new)
                idx = str(_wb_sheet_index(pkg, sh)); dn = pkg.wb.find(q('definedNames'))
                for d in (dn if dn is not None else []):
                    if d.get('name') == '_xlnm._FilterDatabase' and d.get('localSheetId') == idx:
                        d.text = re.sub(r"\$[A-Z]+\$(\d+)$", lambda mm: '$%s$%s' % (CL(ng), mm.group(1)), d.text)
                msg = 'autoFilter %s -> %s' % (m.group(0), new)
            else: msg = 'autoFilter %s đã phủ cột Nguồn giá' % af.get('ref')
        pkg.touch(sh)
        print('%s: cột %s%s %s | %s' % (sh, CL(ng), ' (mới tạo)' if created else '', dict(stat), msg))
    pkg.save(out)



# ---------------------------------------------------------------- kiem: các mục Bảng kiểm tự kiểm bằng máy (thêm 21/10/2026)
def _recalc_wb(f):
    """Tính lại bản sao (xóa cache <v>, LibreOffice) và trả workbook data_only; file gốc không bị ghi."""
    import openpyxl
    tmp = tempfile.mkdtemp(); nc = os.path.join(tmp, 'nocache.xlsx')
    zin = zipfile.ZipFile(f); zo = zipfile.ZipFile(nc, 'w', zipfile.ZIP_DEFLATED)
    pat = re.compile(r'(<f>[^<]*</f>|<f [^>]*>[^<]*</f>|<f [^>]*/>)<v>[^<]*</v>')
    for it in zin.infolist():
        d = zin.read(it.filename)
        if it.filename.startswith('xl/worksheets/sheet'): d = pat.sub(r'\1', d.decode('utf8')).encode('utf8')
        zo.writestr(it, d)
    zo.close()
    od = os.path.join(tmp, 'out'); os.makedirs(od)
    subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', od, nc], capture_output=True, timeout=600)
    wb = openpyxl.load_workbook(os.path.join(od, 'nocache.xlsx'), data_only=True)
    shutil.rmtree(tmp, ignore_errors=True)
    return wb


def cmd_kiem(f, sheets=None, as_json=False):
    """Một lần tính lại, in kết quả các mục Bảng kiểm máy kiểm được (QT21, QT23, QT29 cột Nguồn giá + autoFilter, QT34, droplist 6b).
    Mục cần phán đoán (QT2, QT7, QT15...) KHÔNG nằm ở đây. Thoát mã 1 nếu có mục Lệch."""
    pkg = Pkg(f); wbv = _recalc_wb(f); res = []
    ok = lambda sh, mid, passed, detail: res.append({'sheet': sh, 'muc': mid, 'dat': bool(passed), 'chi_tiet': detail})
    dn = pkg.wb.find(q('definedNames')); junk = [d.get('name') for d in (dn if dn is not None else []) if re.match(r'^[A-Za-z]{1,2}\d+$', d.get('name') or '')]
    ok('(toàn file)', 'QT21 tên vùng rác dạng địa chỉ ô', not junk, '%d tên %s' % (len(junk), junk[:5]))
    ok('(toàn file)', 'Bước 6b DS_MaHieu có mặt', 'DS_MaHieu' in [n for n, _ in pkg.sheets], '')
    for sh, c in _work_sheets(pkg, sheets):
        ws = wbv[sh]; root = pkg.root(sh); sd = root.find(q('sheetData'))
        errs = sum(1 for row in ws.iter_rows() for x in row if isinstance(x.value, str) and x.value.startswith('#'))
        ok(sh, 'QT21 ô lỗi = 0', errs == 0, '%d ô lỗi' % errs)
        hrow = next((r for r in sd if int(r.get('r')) == c['hdr_row']), None)
        ng = next((colnum(x.get('r')) for x in hrow if str(pkg.val(x) or '').strip() == NGUON and colnum(x.get('r')) > c['name']), None) if hrow is not None else None
        af = root.find(q('autoFilter')); m = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)$', af.get('ref')) if af is not None else None
        ok(sh, 'QT29 cột Nguồn giá + autoFilter phủ tới cột này', bool(ng and m and CI(m.group(3)) >= ng), 'Nguồn giá=%s autoFilter=%s' % (CL(ng) if ng else None, af.get('ref') if af is not None else None))
        no_ng = no_wait = pure = n = 0
        for row in sd:
            r = int(row.get('r'))
            if r <= c['hdr_row'] or not _is_work_row(pkg, row, c, ('Dòng', 'Con', 'Cụm')): continue
            n += 1
            kind = pkg.val(_cell(row, c['Loại dòng']))
            if ng and _blank(_cell(row, ng)): no_ng += 1
            cc = _cell(row, c['code']); ct = str(pkg.val(cc) or '').strip() if cc is not None else ''
            if ct and not ct.startswith('=') and ' | ' not in ct and not ct.startswith('('): pure += 1
            if kind in ('Dòng', 'Con') and not ct:
                for key in ('vl0', 'nc0'):
                    if _blank(_cell(row, c[key])): no_wait += 1
        ok(sh, 'QT29 mọi dòng công việc có Nguồn giá', no_ng == 0, '%d/%d dòng thiếu' % (no_ng, n))
        ok(sh, 'QT34 ô giá trống của dòng chưa có mã đã nối công thức chờ', no_wait == 0, '%d ô chưa nối' % no_wait)
        ok(sh, 'QT23 ô Mã hiệu dạng mã thuần = 0', pure == 0, '%d ô' % pure)
        raw = pkg.new.get(pkg.path[sh]) or pkg.files[pkg.path[sh]]   # droplist liên sheet nằm trong extLst (x14:dataValidations)
        ok(sh, 'Bước 6b droplist trỏ DS_MaHieu', b'DS_MaHieu' in raw, '')
    if as_json: print(json.dumps(res, ensure_ascii=False, indent=1))
    else:
        for x in res: print('%-6s %-28s %-62s %s' % ('ĐẠT' if x['dat'] else 'LỆCH', x['sheet'][:28], x['muc'], x['chi_tiet']))
        print('Tổng: %d mục, đạt %d, lệch %d' % (len(res), sum(x['dat'] for x in res), sum(not x['dat'] for x in res)))
    return 0 if all(x['dat'] for x in res) else 1


def unq_ref(x): return x[1:-1].replace('""', '"') if x.startswith('"') else x


if __name__ == '__main__':
    a = sys.argv[1:]
    if not a: print(__doc__); sys.exit()
    cmd = a[0]
    if cmd == 'sheets':
        for n, p in Pkg(a[1]).sheets: print('%-40s %s' % (n, p))
    elif cmd == 'cols': cmd_cols(a[1], a[2])
    elif cmd == 'helpers': cmd_helpers(a[1], a[2], a[3:])
    elif cmd == 'insert': cmd_insert(a[1], a[2], a[3])
    elif cmd == 'fillcdt': cmd_fillcdt(a[1], a[2], a[3])
    elif cmd == 'recalc':
        sh = a[a.index('--sheets') + 1].split('|') if '--sheets' in a else None
        cmd_recalc(a[1], sh)
    elif cmd == 'cleanup': cmd_cleanup(a[1], a[2])
    elif cmd == 'dropdown': cmd_dropdown(a[1], a[2], a[3:])
    elif cmd == 'migrate': cmd_migrate(a[1], a[2], a[3:] or None)
    elif cmd == 'migrate2': cmd_migrate2(a[1], a[2], a[3:] or None)
    elif cmd == 'hsloai': cmd_hsloai(a[1], a[2])
    elif cmd == 'cho': cmd_cho(a[1], a[2], a[3:] or None)
    elif cmd == 'kiem': sys.exit(cmd_kiem(a[1], None, '--json' in a))
    elif cmd == 'nguongia': cmd_nguongia(a[1], a[2], a[3:] or None)
    else: print(__doc__)
