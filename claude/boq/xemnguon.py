#!/usr/bin/env python3
# xemnguon.py (Quy tac 39, ban 27/10/2026): o Don gia goc VL / NC cua dong da gan ma = THAM CHIEU O TRUC TIEP
# toi o gia trong sheet thu vien, vi du  ='2 Thu vien gia'!E2911  (KHONG con HYPERLINK).
#   python3 -I xemnguon.py SRC.xlsx DST.xlsx [--dry]
# - Chi ghi XML (lxml), khong dung openpyxl ghi lai BOQ. Giu nguyen <v> (gia tri cache) nen tong tien khong doi.
# - Phan loai tung o gia (dong Loai dong = Dong / Con, o co SUMIFS(TV[Gia VL goc]) hoac TV[Gia NC goc]):
#     * dong da gan ma, khop DUNG 1 dong TV, gia TV = gia hien thi  -> =<sheet TV>!E{dong}  (NC: cot F), giu phan "*he so" neu co
#     * dong chua co ma (cong thuc cho Buoc 6e)                       -> GIU NGUYEN SUMIFS (de chon ma o droplist van ra gia)
#     * khop 0 hoac >1 dong TV, hoac gia lech, hoac cong thuc la      -> GIU NGUYEN va liet ke ly do
# - Chay lai duoc: o da la tham chieu truc tiep thi tinh lai dong TV theo khoa hien tai; o con HYPERLINK cu (ban 22/10) thi go bo roi lam lai.
# - Doi ma o droplist tren o DA tham chieu truc tiep thi gia KHONG tu doi: chay lai script nay (hoac nho Claude).
import sys, zipfile, re
from lxml import etree

M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
ns = {'m': M}
q = lambda t: '{%s}%s' % (M, t)
nrm = lambda s: re.sub(r'\s+', ' ', str(s or '')).strip().casefold()
col_of = lambda ref: re.sub(r'\d', '', ref)


def col_num(c):
    n = 0
    for ch in c: n = n * 26 + ord(ch) - 64
    return n


def match_paren(s, i):
    """s[i] = '(' -> chi so dau ')' tuong ung (bo qua chuoi trong dau nhay kep)"""
    d, ins = 0, False
    for j in range(i, len(s)):
        ch = s[j]
        if ch == '"': ins = not ins
        elif not ins:
            if ch == '(': d += 1
            elif ch == ')':
                d -= 1
                if d == 0: return j
    return -1


args = [a for a in sys.argv[1:] if not a.startswith('--')]
DRY = '--dry' in sys.argv
SRC, DST = args[0], args[1]
z = zipfile.ZipFile(SRC)

# ---------- shared strings, workbook, sheet map ----------
S = []
if 'xl/sharedStrings.xml' in z.namelist():
    S = [''.join(si.itertext()) for si in etree.fromstring(z.read('xl/sharedStrings.xml'))]
wbx = etree.fromstring(z.read('xl/workbook.xml'))
rels = {r.get('Id'): r.get('Target') for r in etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
sheets = {}
for s in wbx.find('m:sheets', ns):
    t = rels[s.get('{%s}id' % R)].lstrip('/')
    sheets[s.get('name')] = t if t.startswith('xl/') else 'xl/' + t


def cval(c):
    if c is None: return None
    t = c.get('t')
    if t == 'inlineStr': return ''.join(c.find('m:is', ns).itertext())
    v = c.find('m:v', ns)
    if v is None: return None
    return S[int(v.text)] if t == 's' else v.text


def fnum(x):
    try: return float(x)
    except (TypeError, ValueError): return None


# ---------- bang TV: tim sheet + cot theo tieu de ----------
tv_sheet = tv_ref = None
for name, path in sheets.items():
    rl = path.replace('worksheets/', 'worksheets/_rels/') + '.rels'
    if rl not in z.namelist(): continue
    for r in etree.fromstring(z.read(rl)):
        if r.get('Type', '').endswith('/table'):
            tp = 'xl/' + r.get('Target').replace('../', '').lstrip('/')
            tx = etree.fromstring(z.read(tp))
            if tx.get('name') == 'TV' or tx.get('displayName') == 'TV':
                tv_sheet, tv_ref, tv_cols = name, tx.get('ref'), [c.get('name') for c in tx.find('m:tableColumns', ns)]
assert tv_sheet, 'Khong thay bang TV trong workbook'
a, b = tv_ref.split(':')
c0, r0, r1 = col_num(col_of(a)), int(re.sub(r'\D', '', a)), int(re.sub(r'\D', '', b))


def letter(n):
    s = ''
    while n: n, k = divmod(n - 1, 26); s = chr(65 + k) + s
    return s


TC = {nm: letter(c0 + i) for i, nm in enumerate(tv_cols)}
for need in ('Loại vật liệu', 'Mã hiệu', 'Tên công việc', 'Giá VL gốc', 'Giá NC gốc'):
    assert need in TC, 'Bang TV thieu cot ' + need
tvx = etree.fromstring(z.read(sheets[tv_sheet]))
TV = []   # (dong, loai, ma, ten, giaVL, giaNC)
for row in tvx.find('m:sheetData', ns):
    n = int(row.get('r'))
    if n <= r0 or n > r1: continue
    g = {col_of(c.get('r')): c for c in row}
    gv = lambda nm: cval(g.get(TC[nm]))
    TV.append((n, nrm(gv('Loại vật liệu')), nrm(gv('Mã hiệu')), nrm(gv('Tên công việc')), fnum(gv('Giá VL gốc')), fnum(gv('Giá NC gốc'))))
SHQ = "'%s'" % tv_sheet.replace("'", "''")

# ---------- xu ly tung sheet BOQ ----------
HDR_NEED = ['Nhóm VT', 'Loại dòng', 'Mã tra cứu', 'Tên tra cứu']
out_parts, report = {}, []
tot = dict(da_ghi=0, giu_cho_ma=0, giu_khac=0, go_hyperlink=0)
for name, path in sheets.items():
    if name == tv_sheet: continue
    sh = etree.fromstring(z.read(path)); sd = sh.find('m:sheetData', ns)
    hdr = {}
    for r in sd:
        if r.get('r') == '4':
            for c in r:
                if cval(c): hdr[cval(c).strip()] = col_of(c.get('r'))
    if not all(k in hdr for k in HDR_NEED): continue
    cAX, cBB, cBC, cBD = (hdr[k] for k in HDR_NEED)
    st = dict(da_ghi=0, giu_cho_ma=0, giu_khac=0, go_hyperlink=0); changed = False
    for r in sd:
        n = int(r.get('r'))
        if n < 6: continue
        cells = {col_of(c.get('r')): c for c in r}
        if cBB not in cells or cval(cells[cBB]) not in ('Dòng', 'Con'): continue
        for col, c in cells.items():
            f = c.find('m:f', ns)
            if f is None or not f.text: continue
            ft = f.text
            kind = 'VL' if 'TV[Giá VL gốc]' in ft else 'NC' if 'TV[Giá NC gốc]' in ft else None
            direct = re.match(r'^%s!\$?([A-Z]+)\$?\d+(\*[^()+\-/,]+)?$' % re.escape(SHQ), ft)
            if kind is None and direct:
                kind = 'VL' if direct.group(1) == TC['Giá VL gốc'] else 'NC' if direct.group(1) == TC['Giá NC gốc'] else None
            if kind is None and 'HYPERLINK("#' in ft and tv_sheet in ft: kind = '?'
            if kind is None: continue
            if f.get('t') in ('shared', 'array'):
                st['giu_khac'] += 1; report.append((name, c.get('r'), 'công thức dùng chung / mảng, không tự đổi')); continue
            body = ft
            # go HYPERLINK cu (ban 22/10): IF(AND($BC,$BD),<G>,IFERROR(HYPERLINK(...,<G>),<G>))
            pre = 'IF(AND($%s%d="",$%s%d=""),' % (cBC, n, cBD, n)
            if body.startswith(pre) and ',IFERROR(HYPERLINK(' in body:
                body = body[len(pre):body.index(',IFERROR(HYPERLINK(')]; st['go_hyperlink'] += 1
            mult = ''
            if direct and body == ft:
                mult = direct.group(2) or ''; kind_x = kind
            else:
                kind_x = 'VL' if 'TV[Giá VL gốc]' in body else 'NC' if 'TV[Giá NC gốc]' in body else None
                if kind_x is None:
                    st['giu_khac'] += 1; report.append((name, c.get('r'), 'không nhận được công thức giá chuẩn')); continue
                inner = body
                if inner.startswith('IF(OR('):          # bọc chặn mã trống (6e)
                    e = match_paren(inner, 5)
                    if e < 0 or inner[e + 1:e + 4] != ',0,' or not inner.endswith(')'):
                        st['giu_khac'] += 1; report.append((name, c.get('r'), 'bọc IF(OR()) lạ')); continue
                    inner = inner[e + 4:-1]
                if not inner.startswith('SUMIFS('):
                    st['giu_khac'] += 1; report.append((name, c.get('r'), 'không phải SUMIFS chuẩn')); continue
                e = match_paren(inner, 6)
                tail = inner[e + 1:]
                if tail and not re.match(r'^\*[^()+\-/,]+$', tail):
                    st['giu_khac'] += 1; report.append((name, c.get('r'), 'có phép tính khác ngoài SUMIFS: ' + tail)); continue
                mult = tail
            # khoa cua chinh dong (gia tri cache cua cot phu)
            ax, bc, bd = (nrm(cval(cells.get(k))) for k in (cAX, cBC, cBD))
            if bc == '' and bd == '':
                if body != ft:    # da go HYPERLINK: ghi lai G
                    f.text = body; changed = True
                st['giu_cho_ma'] += 1; continue
            cand = [t for t in TV if (ax in ('', 'chưa xác định') or t[1] == ax) and (bc == '' or t[2] == bc) and (bd == '' or t[3] == bd)]
            if len(cand) != 1:
                if body != ft: f.text = body; changed = True
                st['giu_khac'] += 1
                report.append((name, c.get('r'), 'khớp %d dòng TV (cần đúng 1): giữ SUMIFS' % len(cand))); continue
            t = cand[0]; price = t[4] if kind_x == 'VL' else t[5]
            shown = fnum(cval(c)); k = 1.0
            if mult:
                mm = re.match(r'^\*(-?[\d.]+)$', mult); k = float(mm.group(1)) if mm else None
            if price is None or shown is None or k is None or abs((price * k) - shown) > 0.5:
                if body != ft: f.text = body; changed = True
                st['giu_khac'] += 1
                report.append((name, c.get('r'), 'giá TV %s ≠ giá đang hiện %s: giữ SUMIFS' % (price, shown))); continue
            f.text = '%s!%s%d%s' % (SHQ, TC['Giá VL gốc' if kind_x == 'VL' else 'Giá NC gốc'], t[0], mult)
            st['da_ghi'] += 1; changed = True
    for k2 in tot: tot[k2] += st[k2]
    print(name, st)
    if changed: out_parts[path] = etree.tostring(sh, xml_declaration=True, encoding='UTF-8', standalone=True)
print('TONG', tot)
for x in report[:60]: print('  GIU', x)
if len(report) > 60: print('  ... và %d dòng nữa' % (len(report) - 60))
if DRY: sys.exit(0)
out = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
for it in z.infolist():
    out.writestr(it, out_parts.get(it.filename, z.read(it.filename)))
out.close()
print('Đã ghi', DST)