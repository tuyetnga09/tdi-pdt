#!/usr/bin/env python3
"""ncc_rasoat.py - Từ thư mục BG NCC + BOQ -> Bảng rà soát khớp giá (xlsx) -> ASSIGN.json + sổ cái cho skill boq-ghi-gia-ncc.
Lệnh:
  extract --root <thư mục BG NCC> --work <thư mục tạm>
  match   --boq <BOQ.xlsx> --work <thư mục tạm> [--sheets "B2_MEP Service" ...]
  build   --work <thư mục tạm> --xlsx <OUT.xlsx> [--boq-name ...] [--unread-list file.txt]
  assign  --review <OUT.xlsx đã duyệt> --out ASSIGN.json [--auto-duyet] [--component VL|NC]
Chạy bằng: python3 ncc_rasoat.py <lệnh> ...  (cần openpyxl, xlrd, rapidfuzz, numpy; soffice nếu có .xls hỏng)
"""
import os, re, sys, json, argparse, unicodedata, collections, warnings, subprocess, shutil, datetime
warnings.filterwarnings('ignore')

# ---------------- chung ----------------
def noacc(s):
    s = unicodedata.normalize('NFD', str(s))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return s.replace('đ', 'd').replace('Đ', 'D')

def nrm(s):
    return re.sub(r'\s+', ' ', unicodedata.normalize('NFC', str(s)).lower()).strip()

STOP = re.compile(r'\b(cung cap|lap dat|va|bao gom|cho|cua|loai|the|and|for|with|c/w|including|hoan chinh|ca|cai|chiec)\b')
def norm(s):
    s = noacc(s).lower().replace('×', 'x').replace('ø', ' d').replace('φ', ' d').replace('mm²', 'mm2').replace('²', '2')
    s = re.sub(r'(\d),(\d)', r'\1.\2', s)
    s = re.sub(r'[^a-z0-9\.\+x/ ]', ' ', s)
    s = re.sub(r'(?<=\d)(?=[a-z])|(?<=[a-z])(?=\d)', ' ', s)
    s = STOP.sub(' ', s)
    return re.sub(r'\s+', ' ', s).strip()

def words(s):
    t = re.sub(r'[^\w ]', ' ', unicodedata.normalize('NFC', str(s)).lower())
    skip = ('cung', 'cấp', 'lắp', 'đặt', 'và', 'bao', 'gồm', 'cho', 'của', 'loại', 'the', 'and', 'for', 'with', 'cái', 'chiếc', 'bộ')
    return {w for w in t.split() if len(w) >= 2 and not re.search(r'\d', w) and w not in skip}

def nums(s):
    t = re.sub(r'(\d),(\d)', r'\1.\2', noacc(s).lower().replace('×', 'x'))
    out = collections.Counter()
    for m in re.findall(r'\d+(?:\.\d+)?', t):
        out[m.rstrip('0').rstrip('.') if '.' in m else (m.lstrip('0') or '0')] += 1
    return out

UN = {'m': 'm', 'md': 'm', 'met': 'm', 'm dai': 'm', 'cai': 'cai', 'chiec': 'cai', 'pcs': 'cai', 'pc': 'cai', 'ea': 'cai', 'unit': 'cai',
      'bo': 'bo', 'set': 'bo', 'sets': 'bo', 'cap': 'cap', 'cuon': 'cuon', 'roll': 'cuon', 'tu': 'tu', 'panel': 'tu', 'kg': 'kg',
      'm2': 'm2', 'm3': 'm3', 'nrs': 'cai', 'no': 'cai', 'nos': 'cai', 'ht': 'ht', 'he thong': 'ht', 'lot': 'lo', 'lo': 'lo', 'goi': 'goi'}
SPECIAL_UNITS = {'lo', 'ht', 'goi', 'bo', 'cum'}   # đơn vị cần chiết tính/kiểm tra
def ucanon(u):
    u = re.sub(r'\s+', ' ', noacc(u).lower().strip().replace('.', ''))
    return UN.get(u, u)

SIGR = re.compile(r'\d+(?:\.\d+)?(?:\s*c?\s*[x×]\s*\d+(?:\.\d+)?c?)+', re.I)
def sigs(s):
    """Chữ ký kích thước: 4Cx25, 3x2.5, 800x200 -> tuple số (để phân biệt 1Cx10 với 3Cx10)."""
    t = re.sub(r'(\d),(\d)', r'\1.\2', noacc(s).lower())
    out = set()
    for mm in SIGR.finditer(t):
        v = tuple(x.rstrip('0').rstrip('.') if '.' in x else x for x in re.findall(r'\d+(?:\.\d+)?', mm.group(0)))
        if len(v) >= 2: out.add(v); out.add(tuple(reversed(v)))
    return out

def vendor_dir(rel):
    p = rel.split('/')
    return '/'.join(p[:-1]) if len(p) > 1 else '(gốc)'

def vendor_name(rel):
    """Tên NCC: thư mục con đầu tiên sau nhóm hệ, bỏ 'Cap nhat ...' và số thứ tự."""
    parts = [x for x in rel.split('/')[:-1] if not re.match(r'(?i)^(cap nhat|cập nhật|bao gia|báo giá)', x)]
    name = parts[-1] if parts else os.path.splitext(rel.split('/')[-1])[0]
    return re.sub(r'^\s*\d+[\.\)]?\s*', '', name).strip()

def date_from_path(rel):
    """Chỉ nhận ngày không mơ hồ. Trả 'YYYY-MM-DD' hoặc ''."""
    s = rel
    m = re.search(r'(?<!\d)(\d{2})[\.\-/_](\d{2})[\.\-/_](20\d{2})(?!\d)', s) or re.search(r'(?<!\d)(\d{2})(\d{2})(20\d{2})(?!\d)', s)
    cand = set()
    if m:
        cand.add((int(m.group(3)), int(m.group(2)), int(m.group(1))))
    for m in re.finditer(r'(?<!\d)(\d{6})(?!\d)', s):
        a, b, c = int(m.group(1)[:2]), int(m.group(1)[2:4]), int(m.group(1)[4:])
        for y, mo, d in ((2000 + a, b, c), (2000 + c, b, a)):     # yymmdd | ddmmyy
            if 2024 <= y <= 2027 and 1 <= mo <= 12 and 1 <= d <= 31:
                cand.add((y, mo, d))
    if len(cand) == 1:
        y, mo, d = next(iter(cand))
        return '%04d-%02d-%02d' % (y, mo, d)
    return ''

# ---------------- extract ----------------
DESC = re.compile(r'(hàng hóa|hang hoa|chi tiết|sản phẩm|đk ngoài|mô tả|description|tên|diễn giải|nội dung|hạng mục|^item|danh mục|thiết bị|product|material|vật tư|specification|quy cách|chủng loại|model|loại|công việc|part)')
UNITR = re.compile(r'^(đvt|đơn vị|unit$|uom|dvt|đơn vị tính|unit \(|đv)|(\bunit\b|đvt|đơn vị tính)')
QTYR = re.compile(r"(khối lượng|số lượng|^sl|qty|quantity|q'ty|k\.lượng|klg|^kl)")
PRICER = re.compile(r'(đơn giá|unit price|unit rate|^price|giá|rate|đg)')
NOTPRICE = re.compile(r'(thành tiền|total|amount|tổng|thuế|sau vat|after vat|incl|tiền)')
CODE = re.compile(r'\b[A-Z]{2,6}(?:-[A-Z0-9]{1,8}){2,}\b')

def convert_xls(path, work):
    d = os.path.join(work, 'conv'); os.makedirs(d, exist_ok=True)
    src = os.path.join(d, 'src_%d.xls' % (abs(hash(path)) % 10**9)); shutil.copy(path, src)
    subprocess.run(['soffice', '--headless', '--convert-to', 'xlsx', '--outdir', d, src], capture_output=True, timeout=180)
    return src[:-4] + '.xlsx'

def rows_of(path, work):
    import openpyxl, xlrd
    ext = path.lower().rsplit('.', 1)[-1]; out = {}
    if ext == 'xls':
        try: xlrd.open_workbook(path, on_demand=True)
        except Exception: path = convert_xls(path, work); ext = 'xlsx'
    if ext in ('xlsx', 'xlsm'):
        wb = openpyxl.load_workbook(path, data_only=True)
        for ws in wb:
            if ws.sheet_state == 'visible':
                out[ws.title] = [list(r) for r in ws.iter_rows(min_row=1, max_row=4000, max_col=40, values_only=True)]
    else:
        wb = xlrd.open_workbook(path, on_demand=True)
        for sh in wb.sheets():
            if sh.visibility == 0:
                out[sh.name] = [[c if c != '' else None for c in sh.row_values(i)[:40]] for i in range(min(sh.nrows, 4000))]
    return out

def tonum(v):
    if isinstance(v, bool): return None
    if isinstance(v, (int, float)): return float(v)
    if isinstance(v, str):
        s = v.strip().replace(' ', '')
        if re.fullmatch(r'[\d\.,]+', s) and re.search(r'\d', s):
            if s.count(',') and s.count('.'):
                s = s.replace(',', '') if s.rfind('.') > s.rfind(',') else s.replace('.', '').replace(',', '.')
            elif s.count('.') > 1: s = s.replace('.', '')
            elif s.count(',') > 1: s = s.replace(',', '')
            elif s.count(','): s = s.replace(',', '.') if len(s.split(',')[-1]) != 3 else s.replace(',', '')
            try: return float(s)
            except Exception: return None
    return None

def find_header(rows):
    best = None
    for i, r in enumerate(rows[:80]):
        cat = {'desc': [], 'unit': [], 'qty': [], 'price': []}
        for j, c in enumerate(r):
            if c is None or isinstance(c, (int, float)): continue
            t = nrm(c)
            if not t or len(t) > 60: continue
            if PRICER.search(t) and not NOTPRICE.search(t): cat['price'].append(j)
            elif QTYR.search(t): cat['qty'].append(j)
            elif UNITR.search(t) and len(t) < 25: cat['unit'].append(j)
            elif DESC.search(t) and not NOTPRICE.search(t): cat['desc'].append(j)
        score = (1 if cat['desc'] else 0) + (1 if cat['unit'] else 0) + (1 if cat['qty'] else 0) + (2 if cat['price'] else 0)
        if cat['price'] and cat['desc'] and score >= 3 and (best is None or score > best[0]):
            best = (score, i, cat)
            if score >= 5: break
    return best

def extract_file(path, rel, work, codes, sheets=None, src='xlsx'):
    items, log = [], []
    if sheets is None:
        try: sheets = rows_of(path, work)
        except Exception as e: return items, [{'file': rel, 'error': str(e)[:100]}]
    for sn, rows in sheets.items():
        for k, r in enumerate(rows):
            for c in r:
                if isinstance(c, str) and '-' in c:
                    for m in CODE.findall(c): codes.append((m, rel, sn, k + 1))
    for sn, rows in sheets.items():
        h = find_header(rows)
        if not h: log.append({'file': rel, 'sheet': sn, 'n': 0, 'note': 'không nhận được dòng tiêu đề'}); continue
        _, hi, cat = h
        pcols = list(cat['price']); nxt = rows[hi + 1] if hi + 1 < len(rows) else []
        for j, c in enumerate(nxt):
            if c is not None and not isinstance(c, (int, float)):
                t = nrm(c)
                if PRICER.search(t) and not NOTPRICE.search(t) and j not in pcols: pcols.append(j)
        dcols = cat['desc'][:3]; ucol = cat['unit'][0] if cat['unit'] else None; qcol = cat['qty'][0] if cat['qty'] else None
        if not dcols:
            best_c, best_l = None, 0
            for j in range(12):
                ls = [len(str(rows[k][j])) for k in range(hi + 1, min(hi + 40, len(rows))) if j < len(rows[k]) and isinstance(rows[k][j], str)]
                if ls and sum(ls) / len(ls) > best_l: best_c, best_l = j, sum(ls) / len(ls)
            if best_c is not None: dcols = [best_c]
        n = 0; heading = ''
        for k in range(hi + 1, len(rows)):
            r = rows[k]
            texts = [str(r[j]).strip() for j in dcols if j < len(r) and r[j] is not None and not isinstance(r[j], (int, float)) and str(r[j]).strip()]
            if not texts:
                c2 = [str(c).strip() for c in r if isinstance(c, str) and len(c.strip()) > 3]
                if c2: texts = [max(c2, key=len)]
            if not texts: continue
            prices = [(j, tonum(r[j])) for j in pcols if j < len(r) and tonum(r[j]) and tonum(r[j]) > 0]
            if not prices:
                if len(texts[0]) < 120: heading = texts[0]
                continue
            unit = str(r[ucol]).strip() if ucol is not None and ucol < len(r) and r[ucol] is not None else ''
            qty = tonum(r[qcol]) if qcol is not None and qcol < len(r) else None
            if not unit and qty is None: continue
            items.append({'file': rel, 'sheet': sn, 'row': k + 1, 'desc': ' | '.join(texts)[:400], 'unit': unit, 'qty': qty,
                          'price': prices[0][1], 'ctx': heading, 'src': src})
            n += 1
        log.append({'file': rel, 'sheet': sn, 'n': n, 'hdr_row': hi + 1})
    return items, log


# ---- PDF / Word ----
PUNIT = r'(?:m2|m²|m3|md|m|cây|cay|cái|cai|bộ|bo|set|sets|cuộn|cuon|roll|kg|tấn|tan|lít|lit|chiếc|chiec|pcs|pc|ea|unit|ống|ong|tấm|tam|thanh|hộp|hop|cặp|cap|lot|lô|gói|goi|ht|đôi|doi|bình|binh|van|tủ|tu|panel|nos|no|nrs|sheet|cục|cuc|trụ|tru|mét|met|pce|lon|thùng|thung|bao|chai|tuýp|kiện|kien)'
PNUM = r'\d{1,3}(?:[.,]\d{3})+(?:[.,]\d+)?|\d+(?:[.,]\d+)?'
def pval(tok, mode):
    t = tok.replace(',', '') if mode == 'us' else tok.replace('.', '').replace(',', '.')
    try: return float(t)
    except Exception: return None
def parse_text_line(line):
    """Dòng văn bản có 'mô tả  ĐVT  SL  ĐG  TT' -> chỉ nhận khi SL x ĐG = TT (sai số 0,3%)."""
    ns = [(mm.group(0), mm.start()) for mm in re.finditer(r'(?<![\w/])(' + PNUM + r')(?![\w])', line)]
    if len(ns) < 3: return None
    for mode in ('us', 'eu'):
        vs = [(pval(t, mode), s) for t, s in ns]
        for i in range(len(vs) - 2):
            for j in range(i + 1, len(vs) - 1):
                for k in range(j + 1, len(vs)):
                    q, p, t = vs[i][0], vs[j][0], vs[k][0]
                    if q and p and t and q > 0 and p >= 100 and t >= 1000 and abs(q * p - t) <= max(2, 0.003 * t):
                        pre = line[:vs[i][1]].strip(); um = re.search(r'(?i)(?:^|\s)(' + PUNIT + r')\s*$', pre)
                        unit = um.group(1) if um else ''; desc = pre[:um.start()].strip() if um else pre
                        desc = re.sub(r'^[#\d\.\)\-]+\s+', '', desc).strip()
                        desc = re.sub(r'(?:\s+\d+(?:[.,]\d+)?)+\s*$', '', desc).replace('Þ', 'Ø').strip()
                        if len(desc) >= 3: return desc, unit, q, p
    return None
def quote_date(text):
    mm = re.search(r'(?:ngày|date|dated)[^\d]{0,14}(\d{1,2})[\./\-](\d{1,2})[\./\-](20\d{2})', text, re.I)
    if not mm: mm = re.search(r'(?i)ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(20\d{2})', text)
    return '%s-%02d-%02d' % (mm.group(3), int(mm.group(2)), int(mm.group(1))) if mm else ''

VATS = (1.0, 1.05, 1.08, 1.10)
GROUPED = re.compile(r'^\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{1,2})?$')
def _nums_in(line):
    return [(m.group(0), m.start(), m.end()) for m in re.finditer(r'(?<![\w/])(' + PNUM + r')(?![\w])', line)]
def _is_money_line(line):
    return len([1 for t, _, _ in _nums_in(line) if GROUPED.match(t)]) >= 2
_LL = ['']
def _triple(line_nums, win_nums):
    """a,b trên dòng neo; c ở dòng neo hoặc 2 dòng sau; a*b = c (hoặc c gồm VAT 5/8/10%). Ưu tiên cặp SL-ĐG liền kề."""
    best = None
    for mode in ('us', 'eu'):
        va = [(pval(t, mode), s, e) for t, s, e in line_nums]
        vw = [(pval(t, mode), pos) for pos, t in win_nums]
        for i in range(len(va)):
            for j in range(i + 1, len(va)):
                x, y = va[i][0], va[j][0]
                if not x or not y or x <= 0 or y <= 0: continue
                for q, p, qi, pi_ in ((x, y, i, j), (y, x, j, i)):
                    if p < 1000 or q > p: continue          # p là đơn giá (lớn hơn SL)
                    if qi == 0 and len(va) >= 4 and re.fullmatch(r'\W*', _LL[0][:va[qi][1]]): continue   # SL trùng STT đầu dòng
                    for c, cpos in vw:
                        if not c or c < 1000 or cpos in (va[qi][1], va[pi_][1]): continue
                        for vi, v in enumerate(VATS):
                            if abs(q * p * v - c) <= max(2, 0.003 * c):
                                sc = (abs(qi - pi_), vi, 0 if mode == 'us' else 1)
                                if best is None or sc < best[0]: best = (sc, (q, p, qi, pi_, c, v))
    return best[1] if best else None
HDR_LINE = re.compile(r'(?i)(\(vn[đd]\)|^\s*suất\b|commodit|sản phẩm/diễn giải|mô tả|description|thành tiền|amount|đơn giá|unit price|\bstt\b|\bđvt\b|tổng|total|thuế|vat|ghi chú|remarks|hiệu lực|thanh toán|payment)')
def text_items(text, rel, src, label):
    items = []; d = quote_date(text)
    for pg, page in enumerate(text.split('\f'), 1):
        lines = page.split('\n'); anchors = []
        i = 0
        while i < len(lines):
            ln = lines[i]; ns = _nums_in(ln)
            if len(ns) >= 2 and not re.match(r'\s*(Tổng|Total|Sub|VAT|Thuế|Thành tiền|Cộng|Giá trị|Tiền|Số tiền)', ln, re.I):
                win = []
                for k in range(i, min(i + 3, len(lines))):
                    for t, s, e in _nums_in(lines[k]): win.append(((k - i) * 1000 + s, t))
                _LL[0] = ln; r = _triple(ns, win)
                if r:
                    q, p, qi, pi_, c, v = r
                    unit = ''
                    um = re.search(r'(?i)(?:^|[\s\d])(' + PUNIT + r')(?=\s|$)', ln)
                    if um: unit = um.group(1)
                    anchors.append((i, q, p, unit, min(ns[qi][1], ns[pi_][1]))); i += 1; continue
            i += 1
        # gom mô tả: dòng giữa 2 neo chia đôi (nửa đầu thuộc neo trước, nửa sau thuộc neo sau)
        for n, (i, q, p, unit, firstpos) in enumerate(anchors):
            prev_i = anchors[n - 1][0] if n else -1
            next_i = anchors[n + 1][0] if n + 1 < len(anchors) else len(lines)
            lo = prev_i + 1 + ((i - prev_i - 1 + 1) // 2 if n else 0)
            hi = i + 1 + ((next_i - i - 1) // 2 if n + 1 < len(anchors) else min(4, next_i - i - 1))
            parts = []; lo = max(lo, i - 7)
            for k in range(lo, i):
                if HDR_LINE.search(lines[k]) and not _is_money_line(lines[k]): lo = k + 1
            for k in range(lo, min(hi, len(lines))):
                t = lines[k]
                if k == i: t = t[:firstpos]
                elif re.match(r'\s*(Tổng|Total|Thành tiền|Ghi chú|\*?Remarks|Thuế|VAT)', t, re.I):
                    if k > i: break
                    continue
                elif _is_money_line(t):
                    t = re.sub(r'(?<![\w/])' + PNUM + r'%?(?![\w])', ' ', t) if False else re.sub(r'(?<![\w/])\d{1,3}(?:[.,]\d{3})+(?:[.,]\d+)?%?(?![\w])|\d+%', ' ', t)
                    if len(re.findall(r'[^\W\d_]{2,}', t)) < 1: continue
                if HDR_LINE.search(t) and k != i and len(parts) == 0 and k < i: continue
                t = re.sub(r'\s{3,}', ' ', t.strip())
                t = re.sub(r'(?:^|\s)(?:STT|\d{1,3})\s*$' if k == i else r'^$', '', t).strip()
                if k == i:
                    t = re.sub(r'^[#\d\.\)\-]+\s+', '', t);
                    t = re.sub(r'(?i)\s*(?:' + PUNIT + r')\s*$', '', t).strip()
                if len(t) >= 2: parts.append(t)
            desc = ' | '.join(parts).replace('Þ', 'Ø')
            desc = re.sub(r'\s+', ' ', desc).strip(' |')[:300]
            if len(desc) >= 3:
                items.append({'file': rel, 'sheet': '%s %d' % (label, pg), 'row': i + 1, 'desc': desc, 'unit': unit, 'qty': q, 'price': p, 'ctx': '', 'src': src, 'date': d})
    if not items: items = price_list_items(text, rel, src, label, d)
    return items

LIST_ROW = re.compile(r'^\s*(\d{1,3})\s+(.{4,}?)\s{2,}(?:\S+\s+)?(' + PUNIT + r')\s+(\d{1,3}(?:[.,]\d{3})+)(?![\d.,])', re.I)
def price_list_items(text, rel, src, label, d):
    """Bảng giá niêm yết (không có SL/thành tiền): STT | tên | ... | ĐVT | đơn giá. SL = None."""
    if re.search(r'(?i)thành tiền|amount', text): return []
    out = []
    for pg, page in enumerate(text.split('\f'), 1):
        for ln, line in enumerate(page.split('\n'), 1):
            m = LIST_ROW.match(line)
            if not m: continue
            desc = re.sub(r'\s{2,}.*$', '', m.group(2)).strip()
            if re.fullmatch(r'\S+/\S+', desc):      # mô tả nằm ở dòng kế bên
                near = [L.strip() for L in (page.split('\n')[ln - 2], page.split('\n')[ln] if ln < len(page.split('\n')) else '') if L.strip() and not _nums_in(L.strip()[:0] or 'x') == 1]
                near = [re.sub(r'\s{2,}.*$', '', L) for L in near if re.search(r'[^\W\d_]{3,}', L) and not LIST_ROW.match(L)]
                if near: desc = near[0]
            price = float(re.sub(r'[.,]', '', m.group(4)))
            if price < 1000: continue
            out.append({'file': rel, 'sheet': '%s %d' % (label, pg), 'row': ln, 'desc': desc[:300], 'unit': m.group(3), 'qty': None, 'price': price,
                        'ctx': 'Bảng giá niêm yết (không có KL)', 'src': src, 'date': d})
    return out if len(out) >= 3 else []
def docx_sheets_and_text(path):
    import docx
    doc = docx.Document(path); sheets = {}
    for i, t in enumerate(doc.tables, 1):
        rows = []
        for r in t.rows:
            seen, cells = set(), []
            for c in r.cells:
                if id(c._tc) in seen: continue      # ô gộp
                seen.add(id(c._tc)); cells.append(c.text.strip().replace('\n', ' ') or None)
            rows.append(cells)
        sheets['Bảng %d' % i] = rows
    text = '\n'.join(p.text for p in doc.paragraphs)
    return sheets, text

SKIP_NAME = re.compile(r'(catalog|datasheet|data sheet|spec sheet|technical data|^td_|/ref\. \d|catalouge|profile|elta fans|kruger)', re.I)
def cmd_extract(a):
    os.makedirs(a.work, exist_ok=True); allit, alllog, codes, skipped = [], [], [], []
    skip = re.compile(a.skip_regex, re.I) if a.skip_regex else SKIP_NAME
    for dp, dn, fn in os.walk(a.root):
        for f in sorted(fn):
            low = f.lower(); p = os.path.join(dp, f); rel = os.path.relpath(p, a.root).replace('\\', '/')
            ext = low.rsplit('.', 1)[-1] if '.' in low else ''
            if ext not in ('xlsx', 'xls', 'xlsm', 'pdf', 'docx', 'doc'):
                skipped.append('Chưa đọc (' + ext + ')|' + rel); continue
            if skip.search('/' + rel):
                skipped.append('Catalogue/datasheet/bảng chọn thiết bị (không có đơn giá, bỏ qua)|' + rel); continue
            try:
                if ext in ('xlsx', 'xls', 'xlsm'):
                    it, lg = extract_file(p, rel, a.work, codes)
                elif ext == 'pdf':
                    txt = subprocess.run(['pdftotext', '-layout', p, '-'], capture_output=True, text=True).stdout
                    if len(txt.strip()) < 200: it, lg = [], [{'file': rel, 'sheet': '', 'n': 0, 'note': 'PDF scan/không có chữ: cần đọc ảnh'}]
                    else:
                        it = text_items(txt, rel, 'pdf', 'PDF trang')
                        lg = [{'file': rel, 'sheet': '', 'n': len(it), 'note': '' if it else 'PDF có chữ nhưng không nhận dạng được dòng SL×ĐG=TT'}]
                else:   # docx / doc
                    pp = p
                    if ext == 'doc':
                        d = os.path.join(a.work, 'conv'); os.makedirs(d, exist_ok=True)
                        subprocess.run(['soffice', '--headless', '--convert-to', 'docx', '--outdir', d, p], capture_output=True, timeout=180)
                        pp = os.path.join(d, os.path.splitext(f)[0] + '.docx')
                    sheets, text = docx_sheets_and_text(pp)
                    it, lg = extract_file(pp, rel, a.work, codes, sheets=sheets, src='docx')
                    if not it:
                        it = text_items(text, rel, 'docx', 'Văn bản'); lg.append({'file': rel, 'sheet': 'Văn bản', 'n': len(it)})
                    for x in it: x.setdefault('date', quote_date(text))
            except Exception as e:
                it, lg = [], [{'file': rel, 'sheet': '', 'n': 0, 'error': str(e)[:100]}]
            allit += it; alllog += lg; print('%5d  %s' % (sum(l.get('n', 0) for l in lg), rel[:110]), flush=True)
    json.dump(allit, open(os.path.join(a.work, 'ncc_items.json'), 'w'), ensure_ascii=False)
    json.dump(alllog, open(os.path.join(a.work, 'ncc_log.json'), 'w'), ensure_ascii=False)
    json.dump(codes, open(os.path.join(a.work, 'ncc_codes.json'), 'w'), ensure_ascii=False)
    open(os.path.join(a.work, 'unread_files.txt'), 'w').write('\n'.join(skipped))
    print('TỔNG dòng báo giá có giá:', len(allit), collections.Counter(x.get('src', 'xlsx') for x in allit), '| chưa đọc/bỏ qua:', len(skipped))

# ---------------- match ----------------
def load_boq(path, sheets):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True); items = []
    num = lambda v: (float(v) if isinstance(v, (int, float)) else None)
    names = sheets or [s for s in wb.sheetnames if s.upper().startswith('B') and '_' in s]
    for sn in names:
        ws = wb[sn]; ctx = []; s1 = s2 = ''
        for i, r in enumerate(ws.iter_rows(min_row=1, max_col=30, values_only=True), 1):
            stt, d, u = r[0], r[1], r[2]
            if d is None or str(d).strip() == '': continue
            d = str(d).strip(); u = str(u).strip() if u not in (None, '') else ''; st = str(stt).strip() if stt is not None else ''
            if i <= 7 and ('STT' in st or 'Mô tả' in d): continue
            if not u:
                if re.fullmatch(r'[IVX]+', st): s1, s2 = d, ''
                elif re.fullmatch(r'[IVX]+\.\d+', st): s2 = d
                ctx = (ctx + [d])[-2:]; continue
            q = num(r[6]);  q = num(r[28]) if q is None else q
            title = ctx[0] if len(ctx) == 2 else (ctx[-1] if ctx else '')
            items.append({'sheet': sn, 'row': i, 'stt': st, 'desc': d, 'unit': u, 'qty': q, 'title': title[:80], 's1': s1, 's2': s2})
    return items

TAGLINE = re.compile(r'^[A-Za-z]{1,6}(?:-[A-Za-z0-9]{1,8}){1,6}(?:\s*,\s*\d+)*$')
def cmd_match(a):
    import numpy as np
    from rapidfuzz import fuzz, process
    W = a.work; boq = load_boq(a.boq, a.sheets); ncc = json.load(open(os.path.join(W, 'ncc_items.json')))
    codes = json.load(open(os.path.join(W, 'ncc_codes.json')))
    print('BOQ', len(boq), 'NCC', len(ncc))
    for b in boq:
        short = len(words(b['desc'])) <= 1; base = (b['title'] + ' ' + b['desc']) if short else b['desc']
        b.update(text=norm(base), nums=nums(base), u=ucanon(b['unit']), w=words(base), sig=sigs(b['desc']))
    for n in ncc:
        short = len(words(n['desc'])) <= 1; base = (n.get('ctx', '') + ' ' + n['desc']) if short else n['desc']
        n.update(text=norm(base), nums=nums(n['desc']), u=ucanon(n['unit']), w=words(base), vendor=vendor_dir(n['file']), sig=sigs(n['desc']))
    M = process.cdist([b['text'] for b in boq], [n['text'] for n in ncc], scorer=fuzz.token_set_ratio, workers=-1, dtype=np.uint8)
    # chỉ mục mã hiệu (tag)
    idx = collections.defaultdict(list); key = lambda s: re.sub(r'\s+', '', noacc(s).upper())
    for j, n in enumerate(ncc):
        for part in n['desc'].split(' | '):
            p = part.strip(); m = re.match(r'([A-Za-z]{1,6}(?:-[A-Za-z0-9]{1,8}){1,6})(?=$|[\s\(\[,;:/])', p)
            if m: idx[key(m.group(1))].append(j)
            r = re.match(r'((?:[A-Za-z]{1,6}-)+[A-Za-z]*?)(\d+)\s*[~\-–]\s*(\d+)(?=$|[\s\(\[,;:/])', p)   # dải: DB-A-L19~21
            if r and 0 < int(r.group(3)) - int(r.group(2)) < 40:
                for k in range(int(r.group(2)), int(r.group(3)) + 1): idx[key(r.group(1) + str(k))].append(j)
    cidx = collections.defaultdict(list)
    for c, f, s, r in codes: cidx[c.upper()].append((f, s, r))
    def cand(n, lvl, sim, method, qok, uok):
        return {'level': lvl, 'sim': round(sim, 1), 'method': method, 'file': n['file'], 'sheet': n['sheet'], 'row': n['row'], 'desc': n['desc'],
                'unit': n['unit'], 'qty': n['qty'], 'price': n['price'], 'qty_eq': qok, 'unit_ok': uok, 'vendor': n['vendor'], 'src': n.get('src', 'xlsx'), 'date': n.get('date', '')}
    out = []
    for i, b in enumerate(boq):
        cs = []
        for j in np.where(M[i] >= 70)[0]:
            n = ncc[j]
            if not b['text'] or not n['text']: continue
            sim = float(M[i][j]); bn, nn = b['nums'], n['nums']
            if b['sig'] and n['sig'] and not (b['sig'] & n['sig']): continue   # khác 1Cx10/3Cx10, 800x200/200x200
            if bn and nn:
                jac = sum((bn & nn).values()) / sum((bn | nn).values()); sub = all(nn.get(k, 0) >= v for k, v in bn.items())
            elif not bn and not nn: jac, sub = 1.0, True
            else: jac, sub = 0.0, not bn
            uok = (b['u'] == n['u']) or not n['u']
            qok = bool(b['qty'] and n['qty'] and abs(b['qty'] - n['qty']) <= 0.01 * max(b['qty'], 1))
            wb_, wn_ = b['w'], n['w']; cover = len(wb_ & wn_) / len(wb_) if wb_ else 1.0
            dice = 2 * len(wb_ & wn_) / (len(wb_) + len(wn_)) if (wb_ or wn_) else 1.0
            lvl = 0
            if sub and bn and sim >= 85 and cover >= .6 and (dice >= .4 or len(wb_) <= 1): lvl = 3
            elif sim >= 98 and cover >= .8: lvl = 3
            elif qok and sim >= 65 and jac >= .6 and cover >= .5: lvl = 3
            elif jac >= .6 and sim >= 70 and cover >= .5: lvl = 2
            elif not bn and not nn and sim >= 88 and cover >= .6 and dice >= .5: lvl = 2
            elif jac > .5 and sim >= 65 and cover >= .5: lvl = 1
            if lvl > 1 and not uok: lvl -= 1
            for kw in ('xlpe', 'mica', 'lszh'):      # vật liệu cách điện khác nhau -> hạ 1 mức
                if lvl and (kw in b['text'].split()) != (kw in n['text'].split()):
                    lvl -= 1
            if lvl: cs.append((lvl, int(qok), sim, j, uok, qok))
        cs.sort(key=lambda x: (-x[0], -x[1], -x[2]))
        res = [cand(ncc[x[3]], x[0], x[2], 'mô tả', x[5], x[4]) for x in cs]
        # mã hiệu (tag): ưu tiên cao nhất khi dòng BOQ chỉ là mã
        desc = b['desc'].strip()
        if TAGLINE.match(desc) and not (res and res[0]['level'] == 3):
            m = re.match(r'([A-Za-z]{1,6}(?:-[A-Za-z0-9]{1,8}){1,6})', desc); js = idx.get(key(m.group(1)), [])
            if js:
                js = sorted(js, key=lambda j: -(ncc[j]['qty'] is not None and b['qty'] is not None and abs(ncc[j]['qty'] - b['qty']) < 1e-6))
                res = [cand(ncc[j], 3, 100, 'mã hiệu', bool(b['qty'] and ncc[j]['qty'] and abs(b['qty'] - ncc[j]['qty']) <= .01 * b['qty']), True) for j in js] + res
            elif not res:
                h = cidx.get(desc.split(',')[0].strip().upper())
                if h:
                    f, s, r = h[0]
                    res = [{'level': 2, 'sim': 100, 'method': 'mã hiệu – chưa có giá', 'file': f, 'sheet': s, 'row': r,
                            'desc': '(mã %s có trong báo giá, chưa đọc được đơn giá)' % desc.split(',')[0].strip(), 'unit': '', 'qty': None, 'price': None,
                            'qty_eq': False, 'unit_ok': True, 'vendor': vendor_dir(f), 'src': 'xlsx', 'date': ''}]
        # giữ ứng viên đa dạng: tốt nhất của mỗi thư mục NCC trước, tối đa 10
        seen, keep = set(), []
        for c in res:
            if c['vendor'] not in seen: seen.add(c['vendor']); keep.append(c)
        keep = keep[:10]
        out.append({**{k: b[k] for k in ('sheet', 'row', 'stt', 'desc', 'unit', 'qty', 's1', 's2')}, 'n_cands': len(res), 'cands': keep})
    json.dump(out, open(os.path.join(W, 'lines.json'), 'w'), ensure_ascii=False)
    print('mức:', collections.Counter((l['cands'][0]['level'] if l['cands'] else 0) for l in out))

# ---------------- build ----------------
LV = {3: 'Cao', 2: 'Trung bình', 1: 'Thấp'}
def cmd_build(a):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    W = a.work; lines = json.load(open(os.path.join(W, 'lines.json'))); log = json.load(open(os.path.join(W, 'ncc_log.json')))
    unread = open(os.path.join(W, 'unread_files.txt')).read().split('\n') if os.path.exists(os.path.join(W, 'unread_files.txt')) else []
    for i, l in enumerate(lines, 1): l['id'] = 'R%04d' % i
    wb = Workbook(); H = PatternFill('solid', fgColor='1F4E78'); HU = PatternFill('solid', fgColor='C55A11'); hf = Font(bold=True, color='FFFFFF')
    G = PatternFill('solid', fgColor='E2EFDA'); Y = PatternFill('solid', fgColor='FFF2CC'); R = PatternFill('solid', fgColor='FCE4D6')
    def mk(name, hdr, widths, user_cols=()):
        ws = wb.create_sheet(name); ws.append(hdr)
        for j, c in enumerate(ws[1], 1):
            c.font = hf; c.fill = HU if j in user_cols else H; c.alignment = Alignment(wrap_text=True, vertical='center')
        for j, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(j)].width = w
        ws.freeze_panes = 'G2'; return ws
    # --- 1. Duyệt giá
    hdr = ['ID', 'Sheet BOQ', 'Dòng BOQ', 'STT', 'Hệ thống', 'Mô tả BOQ', 'ĐVT BOQ', 'KL BOQ', 'Độ tin cậy', 'Cách khớp', 'Đề xuất của máy',
           'Cảnh báo', 'NCC (đề xuất)', 'Mô tả trong báo giá NCC', 'ĐVT NCC', 'KL NCC', 'Đơn giá NCC', 'Số báo giá khớp', 'Số NCC khác nhau',
           'Giá thấp nhất', 'Giá cao nhất', 'Quyết định (Duyệt / Bỏ qua)', 'Chọn báo giá khác (số hạng ở sheet 2. Báo giá thay thế)', 'NCC (sửa nếu cần)',
           'Ngày báo giá (YYYY-MM-DD)', 'Thành phần giá (VL hoặc NC)', 'VAT của giá NCC', 'Ghi chú người duyệt', 'File nguồn', 'Sheet NCC', 'Dòng NCC']
    ws = mk('1. Duyệt giá', hdr, [8, 16, 8, 7, 22, 46, 8, 9, 10, 12, 14, 38, 18, 46, 8, 9, 14, 9, 9, 14, 14, 14, 16, 18, 14, 14, 12, 28, 60, 14, 8], user_cols=(22, 23, 24, 25, 26, 27, 28))
    uc = mk('2. Báo giá thay thế', ['ID', 'Hạng', 'Độ tin cậy', 'Cách khớp', 'NCC', 'Thư mục NCC', 'File', 'Sheet NCC', 'Dòng NCC', 'Mô tả trong báo giá', 'ĐVT', 'KL', 'Đơn giá', 'Ngày (suy từ tên file)', 'Cảnh báo ĐVT/KL'],
            [8, 6, 10, 14, 18, 40, 60, 14, 8, 50, 8, 9, 14, 14, 24])
    nm = collections.Counter(); unmatched = []
    for l in lines:
        cs = l['cands']
        if not cs: unmatched.append(l); continue
        b = cs[0]; warn = []
        ub = ucanon(l['unit'])
        if ub in SPECIAL_UNITS: warn.append('ĐVT BOQ là lô/bộ/gói/HT: cần chiết tính')
        if b['unit'] and ucanon(b['unit']) != ub: warn.append('ĐVT khác: BOQ %s / NCC %s' % (l['unit'], b['unit']))
        if b['price'] is None: warn.append('Chưa đọc được đơn giá')
        if b['qty'] and l['qty'] and not b['qty_eq']: warn.append('KL BOQ %.4g khác KL báo giá %.4g (chỉ để tham khảo)' % (l['qty'], b['qty']))
        if b['method'] == 'mã hiệu' and len(set((c['file'], c['row']) for c in cs)) >= 1 and re.search(r'[~]', b['desc']): warn.append('Báo giá gộp dải mã (KL là của cả dải)')
        prices = [c['price'] for c in cs if c['price'] and c['level'] == b['level']]
        pmin, pmax = (min(prices), max(prices)) if prices else (None, None)
        if pmin and pmax / pmin > 1.1: warn.append('Giá lệch %.0f%% giữa các NCC/đợt' % ((pmax / pmin - 1) * 100))
        if len({c['vendor'] for c in cs if c['level'] == b['level']}) > 1: warn.append('Nhiều NCC: chọn đúng báo giá')
        if b.get('src', 'xlsx') != 'xlsx': warn.append('Nguồn %s (máy trích): đối chiếu file gốc' % b['src'].upper())
        if b['level'] == 1: warn.append('Chỉ tương đồng mô tả')
        if b['level'] == 2 and b['method'] == 'mô tả': warn.append('Lệch một phần quy cách')
        good = b['level'] == 3 and b['price'] is not None and b.get('src', 'xlsx') == 'xlsx' and not any(w.startswith(('ĐVT', 'Chưa', 'Chỉ')) for w in warn)
        prop = 'Nên fill' if good and len({c['vendor'] for c in cs if c['level'] == 3}) <= 1 else ('Cần so sánh NCC' if good else ('Không nên fill' if b['level'] == 1 or b['price'] is None else 'Cần xem kỹ'))
        d = b.get('date') or date_from_path(b['file'])
        ws.append([l['id'], l['sheet'], l['row'], l['stt'], l['s1'], l['desc'], l['unit'], l['qty'], LV[b['level']], b['method'], prop, '; '.join(warn),
                   vendor_name(b['file']), b['desc'], b['unit'], b['qty'], b['price'], l['n_cands'], len({c['vendor'] for c in cs}), pmin, pmax,
                   None, None, vendor_name(b['file']), d, 'VL', 'Không rõ', None, b['file'], b['sheet'], b['row']])
        r = ws.max_row; ws.cell(r, 11).fill = G if prop == 'Nên fill' else (Y if prop.startswith('Cần') else R)
        for col in (17, 20, 21): ws.cell(r, col).number_format = '#,##0'
        for k, c in enumerate(cs, 1):
            w = []
            if c.get('src', 'xlsx') != 'xlsx': w.append('Nguồn ' + c['src'].upper())
            if c['unit'] and ucanon(c['unit']) != ub: w.append('ĐVT khác')
            if c['qty'] and l['qty'] and not c['qty_eq']: w.append('KL khác')
            uc.append([l['id'], k, LV[c['level']], c['method'], vendor_name(c['file']), c['vendor'], c['file'], c['sheet'], c['row'], c['desc'], c['unit'], c['qty'], c['price'], c.get('date') or date_from_path(c['file']), ', '.join(w)])
            uc.cell(uc.max_row, 13).number_format = '#,##0'
        nm[prop] += 1
    ws.auto_filter.ref = ws.dimensions; uc.auto_filter.ref = uc.dimensions
    dv = DataValidation(type='list', formula1='"Duyệt,Bỏ qua"', allow_blank=True); ws.add_data_validation(dv); dv.add('V2:V%d' % (ws.max_row + 500))
    dv2 = DataValidation(type='list', formula1='"VL,NC"', allow_blank=True); ws.add_data_validation(dv2); dv2.add('Z2:Z%d' % (ws.max_row + 500))
    # --- Chưa khớp
    cu = mk('3. BOQ chưa có báo giá', ['ID', 'Sheet BOQ', 'Dòng BOQ', 'STT', 'Hệ thống', 'Mô tả BOQ', 'ĐVT', 'KL BOQ'], [8, 16, 8, 7, 24, 70, 8, 10])
    for l in unmatched: cu.append([l['id'], l['sheet'], l['row'], l['stt'], l['s1'], l['desc'], l['unit'], l['qty']])
    cu.auto_filter.ref = cu.dimensions
    # --- File chưa đọc
    fu = wb.create_sheet('4. File NCC chưa đọc'); fu.append(['Loại', 'File', 'Ghi chú'])
    for c in fu[1]: c.font = hf; c.fill = H
    for x in log:
        if not x.get('n'): fu.append(['Đã mở nhưng không trích được giá', x['file'] + ((' [' + x['sheet'] + ']') if x.get('sheet') else ''), x.get('note') or x.get('error', '')])
    for f in unread:
        t, _, p = f.partition('|')
        fu.append([t if p else 'Không phải Excel (chưa đọc)', p or f, 'xem thủ công hoặc đọc riêng' if not p else ''])
    fu.column_dimensions['A'].width = 28; fu.column_dimensions['B'].width = 100; fu.column_dimensions['C'].width = 40
    # --- Tổng quan (trang đầu)
    t = wb.active; t.title = '0. Hướng dẫn'; t.column_dimensions['A'].width = 34; t.column_dimensions['B'].width = 110
    tot = len(lines); m3 = sum(1 for l in lines if l['cands'] and l['cands'][0]['level'] == 3); m2 = sum(1 for l in lines if l['cands'] and l['cands'][0]['level'] == 2)
    m1 = sum(1 for l in lines if l['cands'] and l['cands'][0]['level'] == 1)
    rows = [('BẢNG RÀ SOÁT KHỚP GIÁ BG NCC ↔ BOQ', ''), ('BOQ', a.boq_name or ''), ('Phạm vi', 'Đọc Excel (xlsx/xls), PDF có chữ và Word (docx/doc). Dòng từ PDF/Word do máy trích (kiểm tra SL×ĐG=TT) nên luôn cần đối chiếu file gốc. Ảnh, email, file nén, PDF scan chưa đọc; xem sheet 4. File NCC chưa đọc.'), ('', ''),
            ('SỐ LIỆU', ''), ('Tổng dòng BOQ (có ĐVT)', tot), ('Khớp Cao', m3), ('Khớp Trung bình', m2), ('Khớp Thấp (tham khảo)', m1), ('Chưa khớp', tot - m3 - m2 - m1),
            ('Đề xuất: Nên fill', nm['Nên fill']), ('Đề xuất: Cần so sánh NCC', nm['Cần so sánh NCC']), ('Đề xuất: Cần xem kỹ', nm['Cần xem kỹ']), ('Đề xuất: Không nên fill', nm['Không nên fill']), ('', ''),
            ('CÁCH DÙNG (3 bước)', ''),
            ('Bước 1 – Xem 1. Duyệt giá', 'Mỗi dòng = 1 đầu mục BOQ + báo giá NCC khớp nhất. Lọc cột "Đề xuất của máy" (xanh = Nên fill, vàng = Cần xem, đỏ = Không nên). Đọc cột "Cảnh báo".'),
            ('Bước 2 – Quyết định', 'Cột cam (V→AB) do người duyệt điền: V = Duyệt / Bỏ qua. Muốn dùng báo giá khác: xem sheet 2. Báo giá thay thế theo ID, nhập số Hạng vào cột W. Sửa NCC, Ngày, VAT nếu cần. Thành phần chỉ chọn VL hoặc NC (giá gộp cung cấp + lắp đặt chọn VL và ghi chú); một giá không bao giờ được ghi vào cả hai. Dòng để trống = không fill.'),
            ('Bước 3 – Fill vào BOQ', 'Lưu file này rồi chạy lệnh "assign" (tạo ASSIGN.json + sổ cái) và chuyển ASSIGN.json và sổ cái cho skill boq-ghi-gia-ncc (kiểm tra → apply --dry → apply → audit). File này KHÔNG sửa BOQ.'),
            ('', ''), ('GIẢI THÍCH CỘT', ''),
            ('Độ tin cậy', 'Cao = trùng tên + thông số (hoặc trùng mã hiệu), Trung bình = lệch một phần quy cách/ĐVT hoặc mã có trong báo giá nhưng chưa đọc được giá, Thấp = chỉ giống mô tả.'),
            ('Số báo giá khớp', 'Số dòng báo giá NCC (mọi file/đợt) khớp đầu mục này. Số NCC khác nhau = số thư mục NCC/đợt khác nhau. Giá thấp nhất/cao nhất so trong cùng mức tin cậy.'),
            ('Đơn giá NCC', 'Giá đọc nguyên văn từ file báo giá; VAT/chiết khấu CHƯA được chuẩn hóa (cột VAT mặc định "Không rõ").'),
            ('KL khác', 'KL trong báo giá thường khác KL BOQ (báo giá theo bản vẽ/đợt khác): chỉ để tham khảo, không dùng để tính.'),
            ('GIẢ ĐỊNH', '"Khớp" dựa trên mô tả + thông số số học + mã hiệu; chưa đối chiếu hãng/model theo spec. Người duyệt chịu trách nhiệm chọn NCC và giá chốt.')]
    for r in rows: t.append(list(r))
    for r in t.iter_rows():
        r[1].alignment = Alignment(wrap_text=True, vertical='top')
        if r[0].value: r[0].font = Font(bold=True, size=12 if (r[1].value in ('', None) and str(r[0].value).isupper()) else 11)
    wb.move_sheet('0. Hướng dẫn', offset=-wb.index(wb['0. Hướng dẫn']))
    wb.save(a.xlsx); print('Đã ghi', a.xlsx, dict(nm), 'chưa khớp', len(unmatched))

# ---------------- assign ----------------
def cmd_assign(a):
    import openpyxl
    wb = openpyxl.load_workbook(a.review, data_only=True); ws = wb['1. Duyệt giá']; uc = wb['2. Báo giá thay thế']
    H = [c.value for c in ws[1]]; col = {h: i for i, h in enumerate(H)}
    def c_(name_start):
        for h, i in col.items():
            if h and h.startswith(name_start): return i
        raise KeyError(name_start)
    iQ, iPick, iNcc, iDate, iComp, iVat, iNote = c_('Quyết định'), c_('Chọn báo giá khác'), c_('NCC (sửa'), c_('Ngày báo giá'), c_('Thành phần'), c_('VAT'), c_('Ghi chú người')
    iProp = c_('Đề xuất'); iBoqDesc = c_('Mô tả BOQ')
    cands = collections.defaultdict(dict)
    for r in uc.iter_rows(min_row=2, values_only=True):
        if r[0]: cands[r[0]][r[1]] = r
    assign, ledger, skipped = [], [], []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[0]: continue
        q = (r[iQ] or '').strip()
        if not q and a.auto_duyet and r[iProp] == 'Nên fill': q = 'Duyệt'
        if q != 'Duyệt': continue
        rk = int(r[iPick]) if r[iPick] not in (None, '') else 1
        cd = cands[r[0]].get(rk)
        if not cd: skipped.append((r[0], 'không có báo giá hạng %s ở 2. Báo giá thay thế' % rk)); continue
        price = cd[12]; ncc = (r[iNcc] or '').strip(); date = (str(r[iDate]).strip()[:10] if r[iDate] else (cd[13] or ''))
        comp = (r[iComp] or a.component or 'VL').strip().upper().replace(' ', '')
        if comp not in ('VL', 'NC'):
            skipped.append((r[0], 'thành phần "%s" không hợp lệ: chỉ VL hoặc NC. Giá gộp (cung cấp + lắp đặt) chọn VL và ghi chú rõ; muốn ghi cả NC thì nhập báo giá NC riêng' % comp)); continue
        if price is None: skipped.append((r[0], 'chưa có đơn giá')); continue
        if not ncc or not date or not cd[6]: skipped.append((r[0], 'thiếu nguồn (NCC/ngày/file) – điền cột NCC, Ngày báo giá')); continue
        lid = 'BG-%04d' % (len(ledger) + 1)
        item = {'sheet': r[1], 'row': r[2], 'desc_prefix': str(r[iBoqDesc])[:30], 'ncc': ncc, 'date': date, 'file': cd[6],
                'status': 'Cần kiểm tra', 'calc': '',
                'note': '%s!dòng %s; đầu mục NCC: %s%s' % (cd[7], cd[8], str(cd[9]).replace('\n', ' ')[:80], ('; ' + r[iNote]) if r[iNote] else '') + '; VAT: ' + str(r[iVat] or 'Không rõ')}
        item[comp.lower()] = {'price': price, 'ledger_id': lid}
        assign.append(item)
        ledger.append([lid, vendor_name(cd[6]), ncc, date, cd[6], cd[7], cd[8], cd[9], cd[10], price, comp, r[iVat] or 'Không rõ', cd[2]])
    json.dump(assign, open(a.out, 'w'), ensure_ascii=False, indent=1)
    from openpyxl import Workbook
    lw = Workbook(); s = lw.active; s.title = 'SoCai'
    s.append(['ID', 'Hệ NCC (thư mục)', 'NCC', 'Ngày BG', 'File nguồn', 'Sheet', 'Dòng', 'Đầu mục nguyên văn', 'ĐVT', 'Đơn giá', 'Loại giá', 'VAT', 'Độ tin cậy'])
    for x in ledger: s.append(x)
    lw.save(os.path.splitext(a.out)[0] + '_SoCai.xlsx')
    print('ASSIGN: %d dòng | bỏ qua: %d' % (len(assign), len(skipped)))
    for x in skipped[:50]: print('  -', x)

if __name__ == '__main__':
    p = argparse.ArgumentParser(); sp = p.add_subparsers(dest='cmd', required=True)
    e = sp.add_parser('extract'); e.add_argument('--root', required=True); e.add_argument('--work', required=True); e.add_argument('--skip-regex', dest='skip_regex', default='')
    m = sp.add_parser('match'); m.add_argument('--boq', required=True); m.add_argument('--work', required=True); m.add_argument('--sheets', nargs='*')
    b = sp.add_parser('build'); b.add_argument('--work', required=True); b.add_argument('--xlsx', required=True); b.add_argument('--boq-name', dest='boq_name', default='')
    s = sp.add_parser('assign'); s.add_argument('--review', required=True); s.add_argument('--out', required=True); s.add_argument('--auto-duyet', action='store_true'); s.add_argument('--component', default='VL')
    a = p.parse_args(); {'extract': cmd_extract, 'match': cmd_match, 'build': cmd_build, 'assign': cmd_assign}[a.cmd](a)
