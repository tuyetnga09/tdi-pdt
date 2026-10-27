# -*- coding: utf-8 -*-
"""
ycbg_gen.py - Gen file Yeu cau bao gia (YCBG) theo he, moi nhom NCC/hang 1 sheet,
tu file "Dau muc chua co trong thu vien gia" (dau ra skill boq-gan-ma-gia-thu-vien)
va file YCBG mau cua TDI.

Lenh:
  python3 ycbg_gen.py SRC.xlsx TEMPLATE.xlsx OUTDIR [tuy chon]

Tuy chon:
  --ma TDI25538            Ma du an (mac dinh doc tu dong 2 sheet 'Chưa có trong TV')
  --du-an "..."            Ten du an (mac dinh doc tu SRC)
  --nguoi-nhan "..."       Noi dung dong Ten nguoi nhan (mac dinh 'Email: ngaptt@v365.vn')
  --dia-chi-giao "..."     Dia chi giao hang (bat buoc neu SRC khong suy ra duoc)
  --han "..."              Ngay yeu cau bao gia (mac dinh giu chu cua mau: 'Sớm nhất có thể')
  --ngay dd/mm/yyyy        Ngay gui yeu cau (mac dinh de trong, nguoi gui dien)
  --hang-khoa "A,B,..."    Bo sung hang hay khoa du an (uu tien gui som)
  --he-map file.json       Ghi de bang map sheet BOQ -> he
  --chi-chua-co-gia        Chi lay dong 'Chưa tìm được giá' (mac dinh lay het dau muc chua co trong TV)
  --dry                    Chi in thong ke, khong ghi file

Khong gui gi ra ngoai: chi tao file nhap de nguoi duyet o Buoc 5.
"""
import sys, os, re, copy, json, argparse, math, datetime
from io import BytesIO
from collections import OrderedDict, defaultdict
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- cau hinh
HE_MAP = [  # (regex ten sheet BOQ, thu muc, ma ngan) - thu tu quan trong
    (r"ĐIỆN NHẸ|DIEN NHE|ELV", "2. HT Dien nhe", "DIENNHE"),
    (r"PCCC|CHỮA CHÁY|CHUA CHAY|BÁO CHÁY|BAO CHAY", "5. PCCC", "PCCC"),
    (r"CẤP THOÁT|CAP THOAT|CTN|PLUMB", "3. CTN", "CTN"),
    (r"ĐHKK|DHKK|HVAC|TG SC|THÔNG GIÓ|THONG GIO|ĐIỀU HÒA|DIEU HOA|ACMV", "4. HVAC", "HVAC"),
    (r"TBA|MFĐ|MFD|MÁY PHÁT|ĐIỆN|DIEN|TU DIEN|TỦ ĐIỆN", "1. HT dien", "DIEN"),
]
HE_ORDER = ["1. HT dien", "2. HT Dien nhe", "3. CTN", "4. HVAC", "5. PCCC"]

HANG_KHOA = ["Grundfos", "Daikin", "York", "Schneider", "Trane", "Carrier", "Ebara",
             "Wilo", "Mitsubishi", "Hitachi", "ABB", "Siemens", "Hochiki", "Notifier"]
# Nhom thiet bi chinh / hang dai ngay: uu tien gui YCBG som
NHOM_UU_TIEN = ["Thiết bị trung thế, MBA, máy phát", "Cáp trung thế", "Busway",
                "Bơm, bể, thiết bị nước", "Điều hòa (dàn nóng, dàn lạnh)", "Tủ điện",
                "PCCC thiết bị", "Quạt thông gió"]

# Ten sheet ngan (<=31 ky tu, khong / \ ? * [ ] :)
SHEET_NGAN = {
    "Thiết bị trung thế, MBA, máy phát": "TB trung thế-MBA-MPĐ",
    "Thiết bị điện nhẹ (CCTV, PA, VDP, ACS, bãi xe)": "TB điện nhẹ",
    "Chống sét, tiếp địa, đèn báo không": "Chống sét-tiếp địa",
    "Thông gió - điều khiển, rèm ngăn cháy": "TG điều khiển-rèm NC",
    "Xây dựng phụ trợ (hố ga, bệ móng)": "XD phụ trợ",
    "Ống gang, phụ kiện gang/inox": "Ống gang-inox",
    "Ống gang, phụ kiện gang": "Ống gang",
    "Công tác NC (thu hồi, hoàn trả)": "Công tác NC",
    "Cáp tín hiệu, mạng, quang": "Cáp tín hiệu-mạng-quang",
    "Van, thiết bị đường ống": "Van-TB đường ống",
    "Phụ kiện thang máng cáp": "PK thang máng cáp",
    "Điều hòa (dàn nóng, dàn lạnh)": "Điều hòa",
    "Thiết bị điện (công tắc, ổ cắm)": "Công tắc-ổ cắm",
}

# Phan nhom theo tu khoa cho dong khong co cot Nhom (sheet 'Có mã – TV thiếu giá VL',
# 'Thành phần bóc tách') va sua nhom sai ro rang. Thu tu quan trong.
KW_NHOM = [  # (regex, nhom, chi ap cho he nao | None)
    (r"^\s*-?\s*(cụm )?bơm\b", "Bơm, bể, thiết bị nước", None),
    (r"dàn lạnh|dàn nóng|điều ho[àa]|vrv|vrf|chiller|ahu|fcu", "Điều hòa (dàn nóng, dàn lạnh)", None),
    (r"\bquạt\b", "Quạt thông gió", None),
    (r"van chặn lửa|\bnrd\b|\bvcd\b|\bobd\b|\bmfd\b|\bfd\b|\bprd\b|van xả áp|van đóng mở.*động cơ", "HVAC van gió", None),
    (r"van điều chỉnh lưu lượng", "HVAC van gió", "4. HVAC"),
    (r"cửa gió|miệng gió|miệng hút|miệng cấp|louver|cửa tăng áp|cửa hút|hộp cho cửa gió", "HVAC cửa gió", None),
    (r"tiêu âm|bảo ôn|cách nhiệt", "HVAC phụ kiện", "4. HVAC"),
    (r"đầu báo|nút ấn|nút nhấn|chuông|còi|đèn chớp|đèn báo cháy|sprinkler|tủ đựng phương tiện|tủ chữa cháy|bình chữa cháy|trung tâm báo|fire alarm|module|họng tiếp nước|lăng|cuộn vòi|van góc", "PCCC thiết bị", None),
    (r"đèn", "Đèn chiếu sáng", None),
    (r"van|y lọc|rọ hút|rọ bơm|búa nước|đồng hồ|công tắc áp|khớp nối|mối nối mềm", "Van, thiết bị đường ống", None),
    (r"công tắc|ổ cắm|hạt\+mặt", "Thiết bị điện (công tắc, ổ cắm)", None),
    (r"vòi|sen\b|lavabo|bồn|\bxí\b|tiểu|chậu|giá treo|gương|bình nóng lạnh|thuyền tắm|vệ sinh|phễu thu|cầu thu|cầu chắn rác|hộp đựng|\bkệ\b|móc treo", "Thiết bị vệ sinh", None),
    (r"\bcáp\b", "Cáp điện", None),
    (r"máy tính|switch|camera|đầu ghi|\bloa\b|ampli|wifi|access", "Thiết bị điện nhẹ (CCTV, PA, VDP, ACS, bãi xe)", None),
]

CHU_Y = [
    "- Đơn giá chào bằng VNĐ, ghi rõ đã/chưa gồm VAT; giá giao đến chân công trình.",
    "- Ghi rõ Hãng sản xuất, Mã hiệu (model), xuất xứ; gửi kèm catalogue/datasheet; hàng có CO, CQ.",
    "- Ghi rõ thời gian giao hàng, thời gian bảo hành và thời hạn hiệu lực của báo giá.",
    "- Nếu chào hàng tương đương hoặc khác quy cách, ghi rõ tại cột Ghi chú.",
    "- Vui lòng giữ nguyên STT và thứ tự dòng của bảng này khi gửi lại báo giá.",
]

# ---------------------------------------------------------------- tien ich
def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()

def he_of(sheet, he_map):
    up = str(sheet).upper()
    for rx, folder, code in he_map:
        if re.search(rx, up):
            return folder, code
    return None, None

def nhom_kw(desc, default=None, he=None):
    d = norm(desc)
    for rx, g, only in KW_NHOM:
        if only and only != he:
            continue
        if re.search(rx, d):
            return g
    return default

def safe_sheet(name, used):
    s = SHEET_NGAN.get(name, name)
    s = re.sub(r"[\\/?*\[\]:]", "-", s)[:31].strip()
    base, i = s, 2
    while s.lower() in used:
        s = (base[:28] + "-" + str(i)); i += 1
    used.add(s.lower())
    return s

def num(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except Exception:
        return None

def find_header(ws, must):
    """tra ve (row, {ten cot: chi so 0-based}) cua dong tieu de chua cac cot must"""
    for r in range(1, 15):
        vals = [norm(c.value) for c in ws[r]]
        if all(any(m in v for v in vals) for m in must):
            return r, {v: i for i, v in enumerate(vals) if v}
    return None, None

def col(hmap, *keys):
    for k in keys:
        for name, i in hmap.items():
            if name.startswith(k):
                return i
    return None

# ---------------------------------------------------------------- doc nguon
def read_source(path, he_map, chi_chua_co_gia=False):
    wb = openpyxl.load_workbook(path, data_only=True)
    items, excluded, warn = [], [], []
    meta = {"ma": None, "du_an": None}
    names = wb.sheetnames

    def sheet_like(key):
        for n in names:
            if norm(n).startswith(norm(key)):
                return wb[n]
        return None

    # 1) Danh muc chinh
    ws = sheet_like("Chưa có trong TV")
    if ws is None:
        raise SystemExit("Không thấy sheet 'Chưa có trong TV' trong file nguồn")
    for r in range(1, 4):
        t = str(ws.cell(r, 1).value or "")
        m = re.search(r"Dự án:\s*(.+?)\s*\((TDI\d+)\)", t)
        if m:
            meta["du_an"], meta["ma"] = m.group(1).strip(), m.group(2)
    hr, h = find_header(ws, ["id", "nhóm", "mô tả", "đvt"])
    cI, cG, cD, cU = col(h, "id"), col(h, "nhóm"), col(h, "mô tả"), col(h, "đvt")
    cK, cP, cS = col(h, "tổng kl", "kl"), col(h, "vị trí"), col(h, "tình trạng")
    cGia, cLink, cHang = col(h, "giá tham khảo"), col(h, "link"), col(h, "hãng")
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        if row[cD] in (None, ""):
            continue
        st = str(row[cS] or "")
        if chi_chua_co_gia and not st.startswith("Chưa tìm"):
            continue
        pos = [p.strip() for p in str(row[cP] or "").split(",") if p.strip()]
        items.append(dict(src="Chưa có trong TV", id=row[cI], nhom=row[cG] or "Khác",
                          desc=str(row[cD]).strip(), unit=row[cU], kl=num(row[cK]), pos=pos,
                          tinh_trang=st, gia_tk=num(row[cGia]) if cGia is not None else None,
                          link=row[cLink] if cLink is not None else None,
                          hang_tk=row[cHang] if cHang is not None else None, quycach=""))

    # 2) Co ma - TV thieu gia VL
    ws = sheet_like("Có mã")
    if ws is not None:
        hr, h = find_header(ws, ["mô tả", "đvt", "vị trí"])
        cD, cU, cK, cM, cP = col(h, "mô tả"), col(h, "đvt"), col(h, "tổng kl", "kl"), col(h, "mã"), col(h, "vị trí")
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            if row[cD] in (None, ""):
                continue
            pos = [p.strip() for p in str(row[cP] or "").split(",") if p.strip()]
            items.append(dict(src="Có mã – TV thiếu giá VL", id=None, nhom=None,
                              desc=str(row[cD]).strip(), unit=row[cU], kl=num(row[cK]), pos=pos,
                              tinh_trang="Có mã TV, TV chỉ có giá NC", gia_tk=None, link=None,
                              hang_tk=None, quycach="", ma_tv=row[cM]))

    # 3) Thanh phan boc tach: con Chua khop
    ws = sheet_like("Thành phần bóc tách")
    cum = defaultdict(list)
    if ws is not None:
        hr, h = find_header(ws, ["sheet", "thành phần", "tình trạng"])
        cSh, cPr, cPD, cCh, cTP = col(h, "sheet"), col(h, "dòng cha"), col(h, "đầu mục cha"), col(h, "dòng con"), col(h, "thành phần")
        cU, cHS, cKLC, cKL, cS = col(h, "đvt"), col(h, "hệ số"), col(h, "kl cha"), col(h, "kl thành phần"), col(h, "tình trạng")
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            if not row[cSh] or not str(row[cS] or "").startswith("Chưa khớp"):
                continue
            kl = num(row[cKL])
            if kl is None and num(row[cHS]) is not None and num(row[cKLC]) is not None:
                kl = num(row[cHS]) * num(row[cKLC])
            cum[(row[cSh], row[cPr], str(row[cPD]).strip())].append(
                dict(name=str(row[cTP]).strip(), unit=row[cU], hs=num(row[cHS]), kl=kl,
                     pos=f"{row[cSh]}!{row[cCh]}"))
    main_by_desc = {}
    orphans = []
    for it in items:
        if it["src"] == "Chưa có trong TV":
            main_by_desc.setdefault(norm(it["desc"]), it)
    for (sh, pr, pdesc), kids in cum.items():
        parent = None
        pd = norm(pdesc).rstrip(" -")
        for k, it in main_by_desc.items():
            if k.startswith(pd[:60]) and any(p.startswith(str(sh) + "!") for p in it["pos"]):
                parent = it; break
        if parent:  # gop thanh phan vao Quy cach cua dong cum
            parts = "; ".join(f"{k['name']} ({(k['hs'] or 0):g} {k['unit'] or ''}/1 {parent['unit'] or 'bộ'})" for k in kids)
            parent["quycach"] = (parent["quycach"] + " " if parent["quycach"] else "") + "Gồm: " + parts
        else:
            for k in kids:
                orphans.append(dict(k, pdesc=pdesc))
    # con khong co dong cum trong danh muc chinh: gop theo ten + DVT;
    # trung ten voi dong da co trong danh muc chinh (vd 'Tiêu âm lót trong' da gop theo m2) thi gop vao dong do
    og = OrderedDict()
    for k in orphans:
        og.setdefault((norm(k["name"]), norm(k["unit"])), []).append(k)
    for (nm, un), ks in og.items():
        kl = sum(k["kl"] or 0 for k in ks) if any(k["kl"] is not None for k in ks) else None
        hit = main_by_desc.get(nm) or next((it for d, it in main_by_desc.items() if d.startswith(nm[:40]) and norm(it["unit"]) == un), None)
        if not hit:  # cung 3 tu dau + cung DVT + KL lech <= 2% -> da co trong danh muc chinh
            w3 = nm.split()[:3]
            for it in main_by_desc.values():
                if norm(it["desc"]).split()[:3] == w3 and norm(it["unit"]) == un and kl and it["kl"] \
                        and abs(it["kl"] - kl) <= 0.02 * max(kl, it["kl"]):
                    excluded.append(dict(ly_do=f"Thành phần bóc tách trùng dòng ID {it['id']} của danh mục chính",
                                         pos=", ".join(k["pos"] for k in ks), desc=ks[0]["name"], unit=ks[0]["unit"],
                                         kl=kl, note=it["desc"][:120]))
                    hit = "dup"; break
            if hit == "dup":
                continue
        if hit:
            if hit["kl"] in (None, 0) and kl:
                hit["kl"] = kl
                warn.append(f"KL '{hit['desc'][:50]}' lấy theo tổng thành phần bóc tách = {kl:g}")
            hit["pos"] += [k["pos"] for k in ks if k["pos"] not in hit["pos"]]
            continue
        cha = list(OrderedDict.fromkeys(k["pdesc"] for k in ks))
        qc = "Thành phần của: " + "; ".join(c[:60] for c in cha[:3]) + (f"; … ({len(cha)} đầu mục)" if len(cha) > 3 else "")
        items.append(dict(src="Thành phần bóc tách", id=None, nhom=None, pdesc=ks[0]["pdesc"],
                          desc=ks[0]["name"], unit=ks[0]["unit"], kl=kl, pos=[k["pos"] for k in ks],
                          tinh_trang="Thành phần cụm chưa khớp", gia_tk=None, link=None, hang_tk=None,
                          quycach=qc))

    # 4) Cau hinh tu dien
    tu_blocks = OrderedDict()
    ws = sheet_like("TĐ – linh kiện") or sheet_like("TĐ - linh kiện")
    if ws is not None:
        hr, h = find_header(ws, ["tủ điện", "linh kiện", "sl"])
        cT, cR, cL, cU, cSL = col(h, "tủ điện"), col(h, "dòng"), col(h, "linh kiện"), col(h, "đvt"), col(h, "sl")
        last_r, cur = None, None
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            if not row[cL]:
                continue
            name = row[cT]
            r = row[cR]
            if not name:
                # khoi dau tien co the khong co ten tu
                if cur is None or (last_r is not None and r is not None and r - last_r > 2):
                    name = cur if cur and cur.startswith("(Tủ") else f"(Tủ chưa có tên – CT TU DIEN từ dòng {r})"
                else:
                    name = cur
            if name != cur:
                cur = name
            tu_blocks.setdefault(cur, []).append(dict(name=str(row[cL]).strip(), unit=row[cU], sl=num(row[cSL]), r=r))
            last_r = r
        for k in tu_blocks:
            if k.startswith("(Tủ chưa có tên"):
                warn.append(f"Khối linh kiện tủ không có tên trong sheet TĐ – linh kiện: {k} – kiểm tra tên tủ trước khi gửi")

    # 5) Can chiet tinh -> khong dua vao YCBG
    ws = sheet_like("Cần chiết tính")
    if ws is not None:
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0]:
                excluded.append(dict(ly_do="Cần chiết tính (lô/gói/XD phụ trợ) – không xin giá NCC",
                                     pos=f"{row[0]}!{row[1]}", desc=row[2], unit=row[3], kl=row[4], note=row[5]))

    # gan he
    for it in items:
        hes = []
        for p in it["pos"]:
            f, c = he_of(p.rsplit("!", 1)[0], he_map)
            if f and f not in hes:
                hes.append(f)
        if not hes:
            warn.append(f"Không xác định được hệ: {it['desc'][:60]} ({', '.join(it['pos'])})")
            it["he"] = None
            continue
        it["he"] = hes[0]
        if it["nhom"] is None:
            it["nhom"] = nhom_kw(it["desc"], None, it["he"]) or nhom_kw(it.get("pdesc", ""), "Khác", it["he"])
        if len(hes) > 1:
            warn.append(f"Đầu mục dùng chung nhiều hệ {hes} – đặt vào {hes[0]}: {it['desc'][:60]}")
        # sua nhom sai ro rang (vd bom bi xep vao Tu dien)
        g2 = nhom_kw(it["desc"])
        if it["src"] == "Chưa có trong TV" and g2 == "Bơm, bể, thiết bị nước" and it["nhom"] != g2:
            warn.append(f"Chuyển nhóm '{it['nhom']}' → '{g2}': {it['desc'][:60]}")
            it["nhom"] = g2
        # tach nhom Tu dien: dong tu / linh kien
        if it["nhom"] == "Tủ điện & linh kiện tủ":
            if re.match(r"^\s*(tủ|vỏ tủ|ph box|hộp)", norm(it["desc"])) and not any(p.startswith("CT TU DIEN") for p in it["pos"]):
                it["nhom"] = "Tủ điện"
            else:
                it["nhom"] = "Thiết bị đóng cắt, linh kiện tủ"
    return items, tu_blocks, excluded, warn, meta


def merge_dups(rows):
    out = OrderedDict()
    for it in rows:
        k = (norm(it["desc"]), norm(it["unit"]), norm(it["quycach"]))
        if k in out:
            o = out[k]
            o["kl"] = (o["kl"] or 0) + (it["kl"] or 0) if (o["kl"] is not None or it["kl"] is not None) else None
            o["pos"] += [p for p in it["pos"] if p not in o["pos"]]
            o["src"] = o["src"] if it["src"] in o["src"] else o["src"] + " + " + it["src"]
            o.setdefault("ids", []).append(it["id"])
        else:
            it = dict(it); it["ids"] = [it["id"]]
            out[k] = it
    return list(out.values())

# ---------------------------------------------------------------- ghi YCBG
THIN = Side(style="thin"); HAIR = Side(style="hair")

def est_height(texts_widths, base=15.75, line=15.0):
    lines = 1
    for t, w in texts_widths:
        if t:
            n = sum(max(1, math.ceil(len(seg) / max(w * 1.15, 1))) for seg in str(t).split("\n"))
            lines = max(lines, n)
    return max(base, lines * line)

def copy_style(src, dst):
    dst._style = copy.copy(src._style)

class Book:
    def __init__(self, template):
        self.wb = openpyxl.load_workbook(template)
        self.tpl = self.wb.worksheets[0]
        self.img = self.tpl._images[0] if self.tpl._images else None
        self.img_bytes = self.img._data() if self.img else None
        # vi tri cac moc trong mau
        self.r_head = self.r_chuy = self.r_nhan = None
        for r in range(1, self.tpl.max_row + 1):
            a = norm(self.tpl.cell(r, 1).value)
            if a == "stt": self.r_head = r
            elif a.startswith("chú ý"): self.r_chuy = r
            elif a.startswith("thông tin người nhận"): self.r_nhan = r
        assert self.r_head and self.r_chuy and self.r_nhan, "Mẫu YCBG không đúng cấu trúc (STT / Chú ý / Thông tin người nhận)"
        self.n_tpl = self.r_chuy - self.r_head - 1  # so dong hang trong mau
        self.used = set()

    def find_label(self, ws, text, cols=range(1, 9)):
        for r in range(1, ws.max_row + 1):
            for c in cols:
                if norm(ws.cell(r, c).value).startswith(norm(text)):
                    return r, c
        return None, None

    def add_sheet(self, title, rows, info, block_rows=False):
        """rows: list dict(stt, desc, quycach, unit, kl, block=bool)"""
        ws = self.wb.copy_worksheet(self.tpl)
        ws.title = safe_sheet(title, self.used)
        if self.img_bytes:
            im = XLImage(BytesIO(self.img_bytes))
            im.anchor = copy.deepcopy(self.img.anchor)
            im.width, im.height = self.img.width, self.img.height
            ws.add_image(im)
        h0 = self.r_head + 1
        n = len(rows)
        tpl_rows = self.n_tpl
        # them dong hang
        if n > tpl_rows:
            add = n - tpl_rows
            ws.insert_rows(self.r_chuy - 1, add)  # chen truoc dong hang cuoi (giu vien duoi thin)
            for r in range(self.r_chuy - 1, self.r_chuy - 1 + add):
                for c in range(1, 9):
                    copy_style(ws.cell(h0, c), ws.cell(r, c))
        shift = max(0, n - tpl_rows)
        # chieu cao dong: dat lai theo mau cho phan duoi
        for r in range(self.r_chuy, self.tpl.max_row + 1):
            ws.row_dimensions[r + shift].height = self.tpl.row_dimensions[r].height
        widths = {c: (ws.column_dimensions[get_column_letter(c)].width or 10) for c in range(1, 9)}
        blk_fill = PatternFill("solid", fgColor="DDEBF7")
        for i, it in enumerate(rows):
            r = h0 + i
            if it.get("block"):
                ws.cell(r, 1).value = None
                ws.cell(r, 2).value = it["desc"]
                for c in range(1, 9):
                    ws.cell(r, c).fill = blk_fill
                    ws.cell(r, c).font = Font(name="Times New Roman", size=12, bold=True)
                ws.cell(r, 6).value = it.get("unit")
                ws.cell(r, 7).value = it.get("kl")
                ws.cell(r, 3).value = it.get("quycach") or None
            else:
                ws.cell(r, 1).value = it["stt"]
                ws.cell(r, 2).value = it["desc"]
                ws.cell(r, 3).value = it.get("quycach") or None
                ws.cell(r, 6).value = it.get("unit")
                ws.cell(r, 7).value = it.get("kl")
            ws.row_dimensions[r].height = est_height([(it["desc"], widths[2]), (it.get("quycach"), widths[3])])
        for r in range(h0 + n, h0 + max(n, tpl_rows)):
            ws.row_dimensions[r].height = self.tpl.row_dimensions[h0].height
        # header: So, Ngay, Vv
        rS, cS = self.find_label(ws, "Số", cols=[6])
        if rS:
            ws.cell(rS, 7).value = info["so"]
            ws.cell(rS, 7).font = Font(name="Times New Roman", size=11, bold=True)
        rN, cN = self.find_label(ws, "Ngày gửi yêu cầu", cols=[6])
        if rN and info.get("ngay"):
            ws.cell(rN, 7).value = info["ngay"]
        rV, _ = self.find_label(ws, "Vv", cols=[1])
        if rV:
            b = ws.cell(rV, 2)
            b.value = f"{str(b.value).rstrip('.')}\nDự án: {info['du_an']} ({info['ma']}) – Hạng mục: {info['hang_muc']}."
            ws.merge_cells(start_row=rV, start_column=2, end_row=rV, end_column=8)
            b.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[rV].height = 31.5
            ws.cell(rV, 1).alignment = Alignment(vertical="top")
        # chu y
        rC = self.r_chuy + shift
        free = self.r_nhan - self.r_chuy - 2  # so dong trong giua Chu y va Thong tin nguoi nhan (de 1 dong trong)
        if len(CHU_Y) > free:
            ws.insert_rows(rC + 1, len(CHU_Y) - free)
            shift2 = len(CHU_Y) - free
            for r in range(self.r_nhan, self.tpl.max_row + 1):
                ws.row_dimensions[r + shift + shift2].height = self.tpl.row_dimensions[r].height
        else:
            shift2 = 0
        for k, t in enumerate(CHU_Y):
            c = ws.cell(rC + 1 + k, 1)
            c.value = t
            c.font = Font(name="Times New Roman", size=11, italic=True)
            ws.row_dimensions[rC + 1 + k].height = 15.75
        # thong tin nguoi nhan
        def setv(label, val):
            r, c = self.find_label(ws, label, cols=[2])
            if r and val:
                ws.cell(r, 3).value = ": " + val
        setv("Tên người nhận", info.get("nguoi_nhan"))
        setv("Địa chỉ giao hàng", info.get("dia_chi_giao"))
        if info.get("han"):
            setv("Ngày yêu cầu báo giá", info["han"])
        last = self.tpl.max_row + shift + shift2
        ws.print_area = f"A1:H{last}"
        ws.print_title_rows = f"{self.r_head}:{self.r_head}"
        ws.page_setup.orientation = "landscape" if False else ws.page_setup.orientation
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        if info.get("uu_tien"):
            ws.sheet_properties.tabColor = "C00000"
        return ws

    def save(self, path):
        self.wb.remove(self.tpl)
        for n in list(self.wb.defined_names.keys()):
            if re.fullmatch(r"_+a1", n):
                del self.wb.defined_names[n]
        self.wb.active = 0
        self.wb.save(path)


def unique_path(p):
    if not os.path.exists(p):
        return p
    b, e = os.path.splitext(p)
    i = 2
    while os.path.exists(f"{b}_v{i}{e}"):
        i += 1
    return f"{b}_v{i}{e}"


def write_tracking(path, sheets, details, excluded, warn, meta):
    wb = openpyxl.Workbook()
    hdr_font = Font(bold=True, color="FFFFFF"); hdr_fill = PatternFill("solid", fgColor="1F4E78")
    bd = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    def table(ws, headers, data, widths):
        ws.append(headers)
        for c in range(1, len(headers) + 1):
            x = ws.cell(1, c); x.font = hdr_font; x.fill = hdr_fill
            x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in data:
            ws.append(row)
        for r in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=len(headers)):
            for x in r:
                x.border = bd
                if x.row > 1:
                    x.alignment = Alignment(vertical="top", wrap_text=True)
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    ws = wb.active; ws.title = "Theo dõi YCBG"
    table(ws, ["Ưu tiên", "Hệ", "File YCBG", "Sheet", "Số YCBG", "Nhóm vật tư", "Số dòng",
               "Giá trị tham khảo (VNĐ)", "Lý do ưu tiên", "NCC gửi", "Ngày gửi", "Hạn phản hồi",
               "Trạng thái", "Ghi chú"],
          [[s["uu_tien"], s["he"], s["file"], s["sheet"], s["so"], s["nhom"], s["n"], s["gt"] or None,
            s["ly_do"], None, None, None, "Chưa gửi", None] for s in sheets],
          [8, 14, 30, 22, 26, 30, 8, 16, 30, 22, 12, 12, 14, 30])
    dv = DataValidation(type="list", formula1='"Chưa gửi,Đã gửi,Đã nhận BG,Không phản hồi,Hủy"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"M2:M{ws.max_row}")
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 8).number_format = "#,##0"
        if ws.cell(r, 1).value == "Cao":
            ws.cell(r, 1).fill = PatternFill("solid", fgColor="F8CBAD")

    ws = wb.create_sheet("Chi tiết dòng")
    table(ws, ["Số YCBG", "STT", "Hệ", "Sheet", "Diễn giải (BOQ CĐT)", "Quy cách", "ĐVT", "KL",
               "Vị trí trong BOQ", "ID nguồn", "Nguồn danh sách", "Tình trạng giá", "Giá tham khảo VL",
               "Hãng/model tham khảo", "Link", "Đơn giá NCC", "NCC", "Ngày BG"],
          details, [24, 6, 12, 18, 50, 30, 8, 10, 32, 10, 22, 26, 14, 20, 20, 14, 18, 12])
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 13).number_format = "#,##0"; ws.cell(r, 16).number_format = "#,##0"

    ws = wb.create_sheet("Không đưa vào YCBG")
    table(ws, ["Lý do", "Vị trí", "Mô tả", "ĐVT", "KL", "Ghi chú"],
          [[e["ly_do"], e["pos"], e["desc"], e["unit"], e["kl"], e.get("note")] for e in excluded],
          [36, 26, 60, 8, 10, 50])

    ws = wb.create_sheet("Cảnh báo")
    table(ws, ["Cảnh báo cần người rà soát"], [[w] for w in warn], [120])
    wb.save(path)


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src"); ap.add_argument("template"); ap.add_argument("outdir")
    ap.add_argument("--ma"); ap.add_argument("--du-an")
    ap.add_argument("--nguoi-nhan", default="Email: ngaptt@v365.vn")
    ap.add_argument("--dia-chi-giao"); ap.add_argument("--han"); ap.add_argument("--ngay")
    ap.add_argument("--hang-khoa", default=""); ap.add_argument("--he-map")
    ap.add_argument("--chi-chua-co-gia", action="store_true"); ap.add_argument("--dry", action="store_true")
    a = ap.parse_args()

    he_map = HE_MAP
    if a.he_map:
        he_map = [tuple(x) for x in json.load(open(a.he_map, encoding="utf-8"))]
    hang_khoa = HANG_KHOA + [h.strip() for h in a.hang_khoa.split(",") if h.strip()]

    items, tu_blocks, excluded, warn, meta = read_source(a.src, he_map, a.chi_chua_co_gia)
    ma = a.ma or meta["ma"] or "DUAN"
    du_an = a.du_an or meta["du_an"] or ""
    if not a.dia_chi_giao:
        warn.append("Chưa có địa chỉ giao hàng – dòng 'Địa chỉ giao hàng' để trống")

    for it in items:
        if it["kl"] in (None, 0):
            warn.append(f"Thiếu KL hoặc KL = 0: {it['desc'][:60]} ({', '.join(it['pos'])})")

    groups = defaultdict(lambda: defaultdict(list))
    for it in items:
        if it["he"]:
            groups[it["he"]][it["nhom"]].append(it)

    sheets, details, stats = [], [], []
    he_code = {f: c for _, f, c in he_map}
    tu_he = he_of("CT TU DIEN", he_map)[0]
    for he in sorted(groups, key=lambda x: HE_ORDER.index(x) if x in HE_ORDER else 99):
        fname = f"YCBG_{ma}_{he.split('. ', 1)[-1]}.xlsx"
        fdir = os.path.join(a.outdir, he)
        fpath = unique_path(os.path.join(fdir, fname))
        book = None if a.dry else Book(a.template)
        nhoms = groups[he]

        def prio(g):
            rows = nhoms[g]
            hk = sorted({h for it in rows for h in hang_khoa if re.search(r"\b" + re.escape(h) + r"\b", it["desc"], re.I)})
            return hk
        order = sorted(nhoms, key=lambda g: (0 if (prio(g) or g in NHOM_UU_TIEN) else 1, g))
        for k, g in enumerate(order, 1):
            rows = merge_dups(nhoms[g])
            so = f"YCBG-{ma}-{he_code.get(he, 'HE')}-{k:02d}"
            hk = prio(g)
            uu = "Cao" if (hk or g in NHOM_UU_TIEN) else "Thường"
            ly = ("Hãng khóa: " + ", ".join(hk)) if hk else ("Thiết bị chính / hàng đặt dài ngày" if g in NHOM_UU_TIEN else "")
            gt = sum((it["kl"] or 0) * (it["gia_tk"] or 0) for it in rows)
            out_rows = [dict(stt=i, desc=it["desc"], quycach=it["quycach"], unit=it["unit"], kl=it["kl"]) for i, it in enumerate(rows, 1)]
            title = g
            if g == "Tủ điện" and tu_blocks and he == tu_he:
                for it in out_rows:
                    if not norm(it["desc"]).startswith("tủ điện"):
                        continue
                    it["quycach"] = (it["quycach"] + " " if it["quycach"] else "") + "Cấu hình chi tiết: xem sheet 'Tủ điện - cấu hình'."
            info = dict(so=so, ngay=a.ngay, du_an=du_an, ma=ma, hang_muc=g, nguoi_nhan=a.nguoi_nhan,
                        dia_chi_giao=a.dia_chi_giao, han=a.han, uu_tien=(uu == "Cao"))
            sname = title
            if book:
                ws = book.add_sheet(title, out_rows, info)
                sname = ws.title
            sheets.append(dict(uu_tien=uu, he=he, file=os.path.basename(fpath), sheet=sname, so=so, nhom=g,
                               n=len(rows), gt=gt, ly_do=ly))
            for i, it in enumerate(rows, 1):
                details.append([so, i, he, sname, it["desc"], it["quycach"] or None, it["unit"], it["kl"],
                                ", ".join(it["pos"]), ", ".join(str(x) for x in it["ids"] if x is not None) or None,
                                it["src"], it["tinh_trang"], it["gia_tk"], it["hang_tk"], it["link"], None, None, None])
            # phu luc cau hinh tu dien
            if g == "Tủ điện" and tu_blocks and he == tu_he:
                brow, stt = [], 0
                for tu, lk in tu_blocks.items():
                    brow.append(dict(block=True, desc=tu, quycach="Cấu hình 01 tủ", unit="Tủ", kl=None))
                    for x in lk:
                        stt += 1
                        brow.append(dict(stt=stt, desc=x["name"], quycach=None, unit=x["unit"], kl=x["sl"]))
                if book:
                    info2 = dict(info, hang_muc="Tủ điện – cấu hình chi tiết (phụ lục của " + so + ")")
                    ws2 = book.add_sheet("Tủ điện - cấu hình", brow, info2)
                    ws2.sheet_properties.tabColor = "C00000"
                sheets.append(dict(uu_tien=uu, he=he, file=os.path.basename(fpath), sheet="Tủ điện - cấu hình",
                                   so=so + " (phụ lục)", nhom="Cấu hình tủ điện", n=stt, gt=None,
                                   ly_do="Phụ lục cấu hình linh kiện từng tủ (SL/1 tủ)"))
        if book:
            os.makedirs(fdir, exist_ok=True)
            book.save(fpath)
        stats.append((he, fpath, len(order)))

    tpath = unique_path(os.path.join(a.outdir, f"YCBG_{ma}_Theo doi.xlsx"))
    if not a.dry:
        write_tracking(tpath, sheets, details, excluded, warn, meta)

    print(json.dumps(dict(ma=ma, du_an=du_an, files=[dict(he=h, file=p, sheets=n) for h, p, n in stats],
                          tracking=tpath, so_dong=len(details), so_sheet=len(sheets),
                          khong_dua_vao=len(excluded), canh_bao=len(warn),
                          theo_he={h: sum(1 for d in details if d[2] == h) for h, _, _ in stats}),
                     ensure_ascii=False, indent=1))
    for w in warn[:40]:
        print("CẢNH BÁO:", w)
    if len(warn) > 40:
        print(f"... còn {len(warn) - 40} cảnh báo trong sheet 'Cảnh báo' của file theo dõi")


if __name__ == "__main__":
    main()
