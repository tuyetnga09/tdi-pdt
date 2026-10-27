#!/usr/bin/env python3
"""Chuyển BOQ của Chủ đầu tư (CĐT) sang template chào giá TDI.

Ba lệnh con:
  inspect  <file.xlsx> [--rows N]            Xem nhanh các sheet để xác định cột/dòng
  extract  --config cfg.json --out rows.json  Đọc BOQ CĐT, phân loại dòng, gợi ý Chủng loại
  review   --rows rows.json [--only none]     In bảng gọn để Claude soát Chủng loại
  build    --config cfg.json --rows rows.json [--overrides ov.json] --out out.xlsx

Layout đầu ra (Z = số khu khối lượng của CĐT, ví dụ 3 khu -> giống hệt sheet B2_SERVICE):
  Trái (CĐT):  A STT | B Hạng mục | C Thông số | D Mã hiệu | E Nhãn hiệu | F Xuất xứ | G Đơn vị
               | Z cột KL theo khu | 1 cột Tổng KL
  Phải (TDI):  Nội dung | Mã hiệu TDI | Nhãn hiệu XX | Đơn vị | ĐG gốc VL | ĐG gốc NC | Chủng loại
               | In ấn | Tổng KL tạm tính | 18 cột diễn giải KL | KL | ĐG VL | ĐG NC | ĐG tổng
               | 3 cột thành tiền x mỗi khu | 3 cột Total | Ghi chú
"""
import argparse
import copy
import json
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import column_index_from_string as cidx
from openpyxl.utils import get_column_letter as L

HERE = Path(__file__).resolve().parent

# ----------------------------------------------------------------------------
# Chủng loại: luật gợi ý
# Viết luật bằng tiếng Việt có dấu bình thường; khi so khớp, cả luật lẫn văn bản đều
# được bỏ DẤU THANH (sắc huyền hỏi ngã nặng) nhưng GIỮ â ă ê ô ơ ư đ, để "cáp" ≠ "cấp".
# Thứ tự quan trọng: luật cụ thể đứng trước, luật chung đứng sau.
# ----------------------------------------------------------------------------
_TONES = {"\u0300", "\u0301", "\u0303", "\u0309", "\u0323"}


def detone(s):
    s = unicodedata.normalize("NFD", str(s or "").lower())
    s = "".join(ch for ch in s if ch not in _TONES)
    return unicodedata.normalize("NFC", s)


RULES = [
    # Bộ mã theo bảng "Tra cứu từ khóa" (Tong_hop_Tu_khoa_Chung_loai, 10/2026).
    # Thứ tự quan trọng: luật đầu tiên khớp (trên CHỦ NGỮ của mô tả trước, rồi cả mô tả) thắng.
    # --- chi phí, công tác ---
    ("CTK", r"chi phí|\bphí\b(?! gia công)|lệ phí"),
    ("CPK", r"thí nghiệm|thử nghiệm|xét nghiệm|^\s*kiểm định|hiệu chỉnh rơ ?le"),
    ("CTK", r"nghiệm thu|vận hành thử|chạy thử|thử áp|thử kín|thử tải|testing|commissioning|đào tạo|hồ sơ hoàn công|thủ tục|bảo trì|bảo dưỡng|^\s*vận chuyển|nhân công|giàn giáo|xe nâng|sleeve|ống lồng|băng (?:cảnh )?báo|mốc (?:sứ )?báo|sứ (?:đánh dấu|báo)|đào đắp|cát (?:lót|đen|vàng|san)|^\s*(?:bệ|bê tông|xây|trát|gia cố|hoàn thiện|đào|đắp|lấp|khoan|đục|cắt)\b"),
    # --- thông tắc sàn: mã riêng TS (thêm 27/10/2026, nhóm thiết bị × HSTB; trước đây ra TBVS / mã ống).
    #     Đứng ngay sau luật chi phí / thử nghiệm, trước mọi luật khác (RMU có "fco", VTP, TBVS, INOX, UPVC…)
    #     để "Thông tắc sàn FCO / inox / uPVC D110…" không bị nhận nhầm.
    ("TS", r"thông\s+tắc\s+sàn|floor\s*clean.?out"),
    # --- trung thế, máy biến áp ---
    ("MBA", r"^(?!\s*tủ\b).*?(?:máy biến áp|trạm biến áp)|transformer"),
    ("RMU", r"\brmu\b|ring main unit|switchgear|tủ trung thế|mvsg|recloser|\blbs\b|cầu dao phụ tải|máy cắt (?:chân không|trung thế)|\bvcb\b|\bfco\b|cầu chì tự rơi|chống sét van|cảnh báo sự cố cáp"),
    ("CAPLON", r"đầu cáp|cáp trung thế|cáp[^,;]*\b(?:22|24|35)\s?kv|\b(?:22|24|35)\s?kv[^,;]*cáp"),
    ("CAP", r"\b(?:cu|al)\s*/|xlpe|cxv|cvv|lszh|mm2|mm²"),
    ("CSTĐ", r"tiếp địa|tiếp đất|nối đất|thu sét|thoát sét|chống sét(?! lan truyền)|thu lôi|kim thu|earth|grounding|lightning|đồng trần|hàn hóa nhiệt|hàn hoá nhiệt|mối hàn|bộ ghép(?! kênh)|cọc[^,;]*tiếp|băng đồng (?:tiếp|trần|\d)|đếm sét|giảm điện trở"),
    ("VTP", r"vật tư phụ|vật tư đấu nối|phụ kiện kết nối đầu cáp|consumable|đầu cốt|băng keo|giá (?:đỡ|treo|đặt|kê)|ty ren|ty treo|thanh ren|thép hình|\bđai\b(?! ốc)|\bkẹp\b|\bcùm\b|bu ?lông|tắc kê"),
    ("BW", r"busway|busduct|bus duct|thanh dẫn điện|thanh cái|\bfeeder\b|workmanship|hanger|elbow|flanged end|tr connection|plug-?in|\bph box|tap.?off|gia công co|end cap"),
    ("TMC", r"máng cáp|thang cáp|trunking|cable tray|cable ladder|khay cáp|cầu đỡ cáp"),
    ("MFD", r"máy phát|genset|generator|tủ cấp dầu|ống xả.*giảm thanh|\bmfđ\b|nhiên liệu|cấp dầu|bồn dầu"),
    ("UPS", r"\bups\b|ắc quy|ác quy|acquy|battery|bộ nguồn|power supply|bộ lưu điện"),
    # --- PCCC (đứng trước tủ điện / aptomat: "Bình chữa cháy ACB/MFZL-8") ---
    ("TBCC", r"tủ (?:đựng )?(?:vòi|chữa cháy)|hộp (?:đựng )?(?:dụng cụ )?chữa cháy|bình chữa cháy|bình bột|bình co2|\blăng\b|vòi chữa cháy|cuộn (?:ống|vòi)|họng (?:nước|tiếp nước|chữa cháy)|trụ (?:tiếp nước|cứu hỏa|cứu hoả|chữa cháy)|siamese|hose|extinguisher|chăn chiên|đầu phun|sprinkler|drencher"),
    ("FM200", r"khí chữa cháy|chữa cháy (?:bằng )?khí|fm.?200|novec|inergen|aerosol|bình xịt|nút ấn (?:tạm dừng )?xả khí|(?:đèn|còi)[^,;]*xả khí|bình chứa khí"),
    ("FOAM", r"\bfoam\b|bọt (?:chữa cháy|cô đặc)"),
    # --- thiết bị mạng (đứng trước Cáp ĐN: "Patch panel Cat6") ---
    ("CAP DN", r"dây nhảy|patch cord"),
    ("TBDN", r"\bpdu\b|patch|\bodf\b|\bsfp\b|modul(?:e)? quang|mô[\s-]?đun (?:quang|tăng âm)|(?<!công tắc )\bswitch\b|router|firewall|server|máy chủ|\brack\b|tủ mạng|tăng âm|amply|amplifier|khuếch đại|chuông cửa|video|intercom|converter|chuyển đổi quang|\bpoe\b"),
    ("CAP", r"^\s*(?:cáp|dây)[^,;]{0,40}báo cháy"),
    ("CAP DN", r"cat\.?\s?[56]|cáp mạng|cáp quang|cáp đồng trục|dây tín hiệu|cáp tín hiệu|cáp điện thoại|\butp\b|\bftp\b|coaxial|\brg-?(?:6|11)\b|dây loa|cáp loa|hdmi|rs-?485"),
    # --- tủ điện (trước module báo cháy, cửa gió, công tắc: "Tủ điện chiếu sáng ổ cắm") ---
    ("TĐ", r"^\s*(?:e?msb|smdb|mdb|db|dp|ldb|pdb|edb|cdb|acp|mcc|lp|pp(?!-?r\b))[\s\-_.]|tủ điện|tủ phân phối|tủ điều khiển|tủ nút nhấn|tủ cấp nguồn|tủ động lực|tủ chiếu sáng|tủ ats|\bats\b|switchboard|distribution board|tủ bù|tụ bù|tủ hạ thế|\bspd\b|chống sét lan truyền|cắt sét|^\s*tủ\b(?! (?:đựng |)vòi| rack| mạng| trung tâm)"),
    ("KSRV", r"đầu đọc thẻ|thẻ (?:từ|cảm ứng|nhân viên|ra vào|rfid|mifare|proximity)|\brfid\b|\bcard\b"),
    ("TBBC", r"báo cháy|đầu (?:báo|dò)|\bchuông\b|\bcòi\b|nút (?:ấn|nhấn) (?:báo cháy|khẩn|điều khiển)|fire alarm|detector|\bloop\b|module|mô[\s-]?đun|modul|điện trở cuối|tủ trung tâm|bảng hiển thị phụ|đèn (?:báo|chớp)(?! không)|cảm biến (?:khói|nhiệt)(?! độ)"),
    ("TBDN", r"camera|cctv|\bnvr\b|\bhdd\b|đầu ghi|màn hình|\btv\b|tivi|\bbms\b|scada|\bloa\b|speaker|micro|âm thanh|bộ phát|ipbx|tổng đài|matv|splitter|bộ chia tín hiệu|máy trạm|wifi|access point|bộ chuyển mạch|phần mềm|máy in|bàn phím|điện thoại|kiểm soát (?:ra|vào)|access control|khóa (?:điện )?từ|khoá (?:điện )?từ|barrier|exit button|nút (?:nhấn |bấm )?(?:mở cửa|exit)"),
    # --- aptomat ---
    ("ACB", r"\bacb\b|máy cắt không khí|air circuit breaker"),
    ("MCCB", r"\bmccb\b"),
    ("MCB", r"\bmcb\b|\brcbo\b|\brccb\b|\belcb\b|\brcd\b|isolator|cầu dao cách ly"),
    ("CB", r"aptomat|áp tô mát|attomat|cầu dao tự động|circuit breaker|\bcb\b"),
    # --- điều hòa, thông gió, vệ sinh (trước van: "Lavabo ... + van góc", "Van gió VCD") ---
    ("TBDH", r"\bidu\b|\bcdu\b|\bfcu\b|\bahu\b|\bpau\b|\berv\b|\bhru\b|chiller|cooling tower|tháp giải nhiệt|điều hòa|điều hoà|\bvrv\b|\bvrf\b|dàn nóng|dàn lạnh|split|cassette|refnet|bộ chia gas"),
    ("QUAT", r"quạt|\bfan\b|\b(?:paf|eaf|spf|sef|cef|bpf|tef|kef|saf|ef|jf)[\s\-]"),
    ("VCG", r"miệng gió|cửa gió|cửa (?:cấp|hồi|thải|hút|lấy|tăng áp|xả)|van gió|louver|\blover\b|grille|diffuser|lưới chắn côn|\b(?:sag|rag|seg|reg|eag|oag|fag|lg|sad|rad|ead|is|vcd|obd|fd|fsd|nrd)\b|\bmfd\b|damper|van chặn lửa|van xả áp|\bprd\b|van đóng mở|linear|hộp gió|plenum"),
    ("TBVS", r"bồn cầu|\bb\.? ?cầu\b|xí bệt|\bwc\b|lavabo|chậu rửa|bồn rửa|bồn tiểu|tiểu nam|vòi sen|sen tắm|gương|vòi xịt|bồn tắm|treo khăn|sấy tay|giấy vệ sinh|lô (?:để )?giấy|khay xà|kệ kính|kệ xà phòng|thiết bị vệ sinh|phễu thu|phễu thoát|thu sàn|thoát sàn|floor drain|roof drain|cầu chắn rác|nhà bếp|vòi rửa|\bbếp\b"),
    ("VAN", r"^\s*van\b|\bvan (?:bướm|cổng|cầu|bi|một chiều|1 chiều|góc|giảm áp|xả|an toàn|phao|điện từ|chặn|khóa|khoá|báo động|tràn|cân bằng|ổn áp|hai chiều|2 chiều)|valve|y lọc|strainer|đồng hồ|khớp nối mềm|khớp (?:giãn nở|chống rung)|rọ bơm|flow switch|công tắc (?:áp|dòng|lưu lượng|mực|chênh|giám sát)|bình tích áp|bình giãn nở|áp kế|nhiệt kế|gauge|trạm van|chống va|chống xoáy|kính quan sát|cảm biến (?:áp|mực|nhiệt độ|lưu lượng)|que cảm biến"),
    ("CTOC", r"công tắc(?! (?:áp|dòng|mực|chênh|giám sát|hành trình|lưu lượng|kiểm soát))|ổ cắm|socket|hạt (?:công tắc|ổ)"),
    # --- chiếu sáng ---
    ("EXIT", r"sơ đồ|biển (?:báo|chỉ dẫn)|tiêu lệnh|nội quy|bảng (?:chỉ dẫn|hướng dẫn)"),
    ("DEN", r"(?<!thép )(?<!van )(?<!màu )(?<!sơn )\bđèn\b|\bexit\b|thoát hiểm|mắt ếch|\blight|luminaire|highbay|downlight|\bled\b|cảm biến|sensor|aviation|obstruction"),
    # --- năng lượng, nước nóng ---
    ("LLMT", r"năng lượng mặt trời|pin mặt trời|tấm pin|solar"),
    ("MNN", r"máy nước nóng|bình nóng lạnh|bình nước nóng|heat pump|bơm nhiệt|nước nóng \d+ ?l\b"),
    # --- ống luồn, hộp điện ---
    ("OLD", r"ống luồn|luồn dây|conduit|\bemt\b|\bimc\b|\brsc\b|ruột gà|ống dẫn(?! nước)|ống g\.?i d\d|ống hdpe d\d+/\d+|hộp (?:nối|đấu|chia|kỹ thuật đấu|điện)|\bbox\b(?! gió)|đế âm|hộp âm"),
    # --- ống, ống gió ---
    ("BO", r"^\s*(?:bảo ôn|cách nhiệt|cách âm|tiêu âm|insulation)"),
    ("OD", r"ống đồng|copper|\bø\s*\d+[.,]\d+\s*mm"),
    ("UPVC", r"nước ngưng|condensate"),
    ("OG", r"ống gió|g\.?i dày|\bduct\b|ductwork|tôn mạ kẽm|\bei\s?\d+"),
    ("CCL", r"chống cháy lan|bịt kín|fire ?stop|chèn chống cháy|vữa chống cháy"),
    ("BO", r"bảo ôn|cách nhiệt|cách âm|tiêu âm|insulation"),
    ("Vencap", r"vent ?cap|chụp thông hơi|nắp thông hơi"),
    ("INOX", r"ống (?:thép )?(?:inox|không gỉ|không rỉ)|stainless steel pipe|ống sus|inox[^,;]*\bdn\s?\d"),
    ("PPR", r"\bppr\b|pp-r|nước lạnh|nước nóng|nước cấp sinh hoạt|hàn nhiệt"),
    ("HDPE", r"hdpe|âm đất.*thải|thải.*âm đất"),
    ("UPVC", r"upvc|u-pvc|\bpvc\b|trục đứng|ống xả|thoát nước thải trong nhà|\bpp\b"),
    ("OTH", r"ống thép|thép tráng kẽm|thép mạ kẽm|mạ kẽm|thép đen|\bsch\s?\d+|bs\s?1387|âm đất|steel pipe|\bg\.?i\b|ống trong nhà"),
    # --- bơm, lọc ---
    ("BOM", r"bơm|pump|jockey|lọc nước|bình lọc|than hoạt tính|lọc áp lực|softener|làm mềm|khử trùng|\buv\b"),
    # --- công tác xây dựng, bể, hố ga ---
    ("CTK", r"\bbể\b|\bbệ\b|hố ga|hố thu|bể phốt|cống (?:thoát|hộp|tròn|bản|bê tông|d\d|ø)|song chắn|nắp (?:hố|ga|bể|đan)|tấm đan|bồn chứa|bồn nước|\btéc\b|\btank\b|gas bổ sung|nạp gas|công tác khác|\blấp\b|bê tông|\bmóng\b|\bxây\b|\btrát\b|ốp lát|\bgạch\b|chống thấm|\bsơn\b|chống rỉ|khoan|\bđục\b|cắt (?:bê tông|tường|sàn|đường)|cửa sập|cửa (?:kiểm tra|thăm)|access door|access panel"),
    # --- cáp chung (cuối cùng) ---
    ("CAP", r"mm2|mm²|\bcu/|cxv|cvv|\bcv\b|xlpe|lszh|\bcu\b.*pvc|cáp điều khiển|\bawg\b|cáp điện|dây điện|dây dẫn|\bcáp\b|cable"),
]
_RULES = [(code, re.compile(detone(p), re.IGNORECASE)) for code, p in RULES]

# Mô tả chỉ là kích thước / chung chung -> phải dựa vào tiêu đề cha
GENERIC = re.compile(detone(
    r"^\s*(dn|d|ø|phi|size)?\s*[\d.,/x×\s\-]+(mm)?\s*$|^\s*phụ kiện( cho ống)?|^\s*phòng\b|^\s*(tầng|phần|khối|tháp|tower|block|hầm)\b"))

# Hệ thống "chi phối": mọi thiết bị nằm dưới tiêu đề này lấy luôn mã hệ thống
DOMINANT = [(code, re.compile(detone(p))) for code, p in [
    ("FM200", r"bình xịt|aerosol|chữa cháy khí|khí chữa cháy|fm.?200"), ("FOAM", r"\bfoam\b"),
    ("MFD", r"máy phát"),
    ("TBBC", r"hệ thống báo cháy|fire alarm system"),
]]
NON_DOMINATED = {"VTP", "CTK", "CPK", "CAP", "CAPLON", "CAP DN", "OLD", "TMC", "CSTĐ", "TBCC", "TBDN",
                 "OTH", "INOX", "PPR", "HDPE", "UPVC", "BO", "CCL", "TS"}
ELEC_SYSTEM = re.compile(detone(r"điện|\belv\b|\blv\b|chống sét|camera|cctv|âm thanh|thông tin|truyền hình|kiểm soát ra vào|báo cháy"))
CONDUIT_HEAD = re.compile(detone(r"ống luồn|luồn dây|conduit|ruột gà"))
PIPE_WORD = re.compile(detone(r"\bống\b|\bpvc\b|hdpe|\bg\.?i\b|conduit"))
NOT_CONDUIT = re.compile(detone(r"ống đồng|ống gió|nước|thoát|cấp thoát"))
LOWV_SYSTEM = re.compile(detone(r"điện nhẹ|\belv\b|camera|cctv|âm thanh|thông tin|truyền hình|mạng|kiểm soát ra vào|điện thoại|\bdata\b"))
AMPS = re.compile(r"(\d+(?:[.,]\d+)?)\s*a\b")
CABLE_SIZE = re.compile(r"(\d+(?:[.,]\d+)?)\s*mm(?:2|²)")


PIPE_CODES = {"OLD", "OD", "OG", "INOX", "PPR", "HDPE", "UPVC", "OTH"}
_PREFIX = re.compile(detone(r"^\s*(cung cấp|vận chuyển|lắp đặt|thi công|supply|install)[\s,/&và-]*((cung cấp|vận chuyển|lắp đặt|thi công|and|và|,|&|/)\s*)*"))
_CUT = re.compile(detone(r"\s+(bao gồm|gồm|c/w|including|kèm theo|kèm|có|with|tại)\b|[(:;]"))


def subject(text):
    """'Cung cấp và lắp đặt van cổng có công tắc giám sát' -> 'van cổng'."""
    t = _PREFIX.sub("", detone(text))
    m = _CUT.search(t)
    return t[:m.start()] if m and m.start() > 2 else t


def _match(text, only=None):
    t = detone(text)
    for code, rx in _RULES:
        if only and code not in only:
            continue
        if rx.search(t):
            return code
    return None


def _match_ctx(text, only=None):
    return _match(subject(text), only) or _match(text, only)


def _cap_size(code, *texts, system=""):
    """Tinh chỉnh mã sau khi khớp luật:
    CAP: tiết diện pha >= 70mm2 -> CAPLON; nằm trong hệ Điện nhẹ -> CAP DN.
    CB (aptomat chưa ghi loại): <= 63A -> MCB, 64-630A -> MCCB, >= 800A -> ACB, không rõ dòng -> MCCB."""
    if code == "CB":
        for t in texts:
            m = AMPS.search(detone(t))
            if m:
                amp = float(m.group(1).replace(",", "."))
                return "MCB" if amp <= 63 else ("MCCB" if amp < 800 else "ACB")
        return "MCCB"
    if code != "CAP":
        return code
    for t in texts:
        m = CABLE_SIZE.search(detone(t))
        if m and float(m.group(1).replace(",", ".")) >= 70:
            return "CAPLON"
    if system and LOWV_SYSTEM.search(detone(system)):
        return "CAP DN"
    return code


def classify(desc, context, system="", ranked=None):
    """Trả (code, source). source: 'desc' | 'system' | 'context' | 'none'.
    context: tiêu đề phía trên (gần nhất trước); system: tiêu đề cấp A/I;
    ranked: [(level, text)] để biết tiêu đề nào là cấp con (>=2) hay ghi chú (9)."""
    d = detone(desc)
    # "Phụ kiện …" luôn là VTP (bảng Tra cứu); co/cút/tê/côn… ghi rõ thì theo vật liệu ống
    if PHU_KIEN_HEAD.match(d) and not VTP_EXCEPT.search(d):
        return "VTP", "desc"
    generic = (not d) or bool(GENERIC.search(d))
    desc_code = None
    if not generic:
        sub_code, full_code = _match(subject(d)), _match(d)
        desc_code = sub_code or full_code
        if sub_code in PIPE_CODES and full_code and full_code not in PIPE_CODES:
            desc_code = full_code  # "Đường ống xả … giảm thanh" -> MFĐ, không phải ống
    near_subj = subject(context[0]) if context else ""
    own = near_subj if generic else d
    # 1) Ống trong hệ Điện / Điện nhẹ / Báo cháy -> ống luồn dây (OLD)
    in_conduit = any(CONDUIT_HEAD.search(detone(c)) for c in (context or [])[:2])
    if (ELEC_SYSTEM.search(detone(system)) or in_conduit) and PIPE_WORD.search(own) and not NOT_CONDUIT.search(own) \
            and (desc_code is None or desc_code in PIPE_CODES):
        return "OLD", "system"
    # 2) Hệ thống chi phối (bình xịt, foam, kiểm soát ra vào, máy phát, báo cháy)
    if desc_code not in NON_DOMINATED:
        for lvl, txt in (ranked or []):
            if lvl is not None and lvl < 2:
                continue
            for code, rx in DOMINANT:
                if rx.search(subject(txt)) or rx.search(detone(txt)[:60]):
                    return code, "system"
    if desc_code:
        return _cap_size(desc_code, desc, system=system), "desc"
    # 3) Mô tả chỉ là kích thước dưới một tiêu đề "Ống ..." -> chỉ tìm vật liệu ống
    if context and re.match(detone(r"\s*(đường )?ống\b|\s*pipe"), near_subj) and not re.search(detone(r"\bvan\b|văn chặn"), near_subj):
        for ctx in context:
            code = _match_ctx(ctx, PIPE_CODES)
            if code:
                return code, "context"
        return None, "none"
    # 4) Dựa vào tiêu đề cha, gần nhất trước
    for ctx in context:
        code = _match_ctx(ctx)
        if code:
            return _cap_size(code, desc, ctx, system=system), "context"
    if d:
        code = _match(d)
        if code:
            return _cap_size(code, desc, system=system), "desc"
        # "Phụ kiện …" / "Vật tư …" không có mã ống phía trên -> VTP (bước extract vẫn cho kế thừa mã ống)
        if PHU_KIEN_HEAD.match(d):
            return "VTP", "fallback"
    return None, "none"


PHU_KIEN_HEAD = re.compile(detone(r"\s*(phụ kiện|vật tư)"))
VTP_EXCEPT = re.compile(detone(r"đầu cáp|chống sét|thu sét|thoát sét|tiếp địa"))


ROMAN = re.compile(r"^(?=[IVXLC])M*(C[MD]|D?C{0,3})(X[CL]|L?X{0,3})(I[XV]|V?I{0,3})\.?$")


def norm(s):
    return unicodedata.normalize("NFC", str(s)).strip() if s is not None else ""


# Nhóm hệ số đơn giá vật liệu (quyết định công thức cột ĐG VL)
THIET_BI = {
    "RMU", "MBA", "TĐ", "MFD", "UPS", "MCB", "MCCB", "ACB", "CTOC", "DEN", "BW",
    "TBBC", "TBDN", "KSRV", "VAN", "BOM", "TBCC", "FM200", "FOAM",
    "QUAT", "VCG", "TBDH", "TBVS", "TS", "MNN", "LLMT",
}
LINEAR_UNITS = {"m", "m2", "m3", "md", "m²", "m³", "mét", "met"}


def price_group(code, unit):
    u = norm(unit).lower().replace(".", "")
    if u in LINEAR_UNITS:
        return "hao_hut"
    if code in THIET_BI:
        return "thiet_bi"
    return "vat_lieu"


# ----------------------------------------------------------------------------
# inspect
# ----------------------------------------------------------------------------
def cmd_inspect(a):
    wb = openpyxl.load_workbook(a.file, data_only=True)
    for ws in wb.worksheets:
        print(f"\n===== Sheet '{ws.title}'  dims={ws.dimensions}  max_row={ws.max_row}  state={ws.sheet_state}")
        shown = 0
        for r in range(1, ws.max_row + 1):
            vals = [(c.coordinate, c.value) for c in ws[r] if c.value not in (None, "")]
            if not vals:
                continue
            print("  ", [(k, str(v)[:40]) for k, v in vals[:18]])
            shown += 1
            if shown >= a.rows:
                break
        mr = [str(m) for m in ws.merged_cells.ranges if m.min_row <= 20][:25]
        if mr:
            print("   merged (header):", mr)


# ----------------------------------------------------------------------------
# extract
# ----------------------------------------------------------------------------
def _cell(ws, r, col):
    if not col:
        return None
    v = ws.cell(r, cidx(col)).value
    if isinstance(v, str):
        v = v.strip()
        return v or None
    return v


def detect_kind(stt, unit, qtys, prev_letter):
    s = norm(stt)
    has_unit = bool(norm(unit))
    has_qty = any(isinstance(q, (int, float)) and q != 0 for q in qtys)
    if s:
        letter = s.rstrip(".")
        if re.fullmatch(r"[A-H]", letter) and not has_unit:
            if letter not in "CD" or prev_letter == chr(ord(letter) - 1):
                return "section", 0
        if ROMAN.match(s) and not has_unit:
            return "section", 1
        if re.fullmatch(r"[IVXLC]+\.\d+\.?", s) and not has_unit:
            return "section", 2
        if re.fullmatch(r"[IVXLC]+\.\d+\.\d+\.?", s) and not has_unit:
            return "section", 3
    if has_unit or has_qty:
        return "item", None
    return "note", None


def cmd_extract(a):
    cfg = json.loads(Path(a.config).read_text(encoding="utf-8"))
    wb = openpyxl.load_workbook(cfg["input"], data_only=True)
    wbf = openpyxl.load_workbook(cfg["input"])  # để biết ô nào là công thức thiếu giá trị cache
    out = {"sheets": []}
    warn_formula = 0
    for sc in cfg["sheets"]:
        ws = wb[sc["source_sheet"]]
        wsf = wbf[sc["source_sheet"]]
        cols = sc["columns"]
        zones = sc["zones"]
        start = sc["data_start_row"]
        end = sc.get("data_end_row") or ws.max_row
        rows = []
        context = []  # [(level_rank, text)]
        prev_letter = None
        for r in range(start, end + 1):
            stt = _cell(ws, r, cols.get("stt"))
            desc = _cell(ws, r, cols.get("desc"))
            unit = _cell(ws, r, cols.get("unit"))
            qtys = []
            for z in zones:
                v = _cell(ws, r, z["col"])
                if v is None and isinstance(wsf.cell(r, cidx(z["col"])).value, str) and str(wsf.cell(r, cidx(z["col"])).value).startswith("="):
                    warn_formula += 1
                if isinstance(v, str):
                    try:
                        v = float(v.replace(",", ""))
                    except ValueError:
                        v = None
                qtys.append(v)
            other = [_cell(ws, r, cols.get(k)) for k in ("spec", "code", "brand", "origin")]
            if all(x in (None, "") for x in [stt, desc, unit, *other]) and not any(qtys):
                continue
            kind, level = detect_kind(stt, unit, qtys, prev_letter)
            if kind == "section" and level == 0:
                prev_letter = norm(stt).rstrip(".")
            row = {
                "src_row": r, "stt": stt, "desc": desc,
                "spec": other[0], "code_cdt": other[1], "brand": other[2], "origin": other[3],
                "unit": unit, "qty": qtys, "kind": kind, "level": level,
                "bold": bool(ws.cell(r, cidx(cols["desc"])).font.b) if cols.get("desc") else False,
            }
            # ngữ cảnh: tiêu đề/ghi chú phía trên, gần nhất trước
            if kind == "section":
                rank = level
                context = [c for c in context if c[0] < rank] + [(rank, norm(desc))]
            elif kind == "note":
                context = [c for c in context if c[0] < 9] + [(9, norm(desc))]
                context = [c for c in context if c[0] < 9] + [c for c in context if c[0] == 9][-2:]
            row["context"] = [t for _, t in reversed(context)][:5]
            row["system"] = " | ".join(t for lv, t in context if lv is not None and lv <= 1)
            row["ranked"] = [[lv, t] for lv, t in reversed(context)]
            if kind == "item":
                code, src = classify(desc, row["context"], row["system"], list(reversed(context)))
                if code is None and row["spec"]:
                    code, src = classify(row["spec"], [], row["system"])
                row["cl"], row["cl_source"] = code, src
            else:
                row["cl"], row["cl_source"] = None, "-"
            # "Phụ kiện cho ống…" lấy mã của ống ngay phía trên (cùng tiêu đề cha)
            if kind == "item" and re.match(detone(r"\s*phụ kiện"), detone(desc)):
                prev = next((p for p in reversed(rows) if p["kind"] == "item"), None)
                if prev and prev.get("cl") and prev["context"][:1] == row["context"][:1] \
                        and row.get("cl_source") != "desc":
                    row["cl"], row["cl_source"] = prev["cl"], "inherit"
            rows.append(row)
        out["sheets"].append({"output_sheet": sc["output_sheet"], "rows": rows})
        n_item = sum(1 for x in rows if x["kind"] == "item")
        n_none = sum(1 for x in rows if x["kind"] == "item" and not x["cl"])
        n_ctx = sum(1 for x in rows if x["cl_source"] == "context")
        print(f"[{sc['output_sheet']}] {len(rows)} dòng | {n_item} hạng mục | chưa phân loại: {n_none} | gán theo tiêu đề cha: {n_ctx}")
    if warn_formula:
        print(f"CẢNH BÁO: {warn_formula} ô KL là công thức nhưng không có giá trị cache -> mở file bằng Excel, Save lại rồi chạy lại.")
    Path(a.out).write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print("Đã ghi", a.out)


# ----------------------------------------------------------------------------
# review
# ----------------------------------------------------------------------------
def group_key(x):
    """Khóa nhóm = tiêu đề gần nhất < tiêu đề cha (tránh gộp nhầm khi tiêu đề chung chung như 'Tháp B')."""
    ctx = x.get("context") or []
    return " < ".join(t.strip()[:80] for t in ctx[:2])


def review_groups(rows, limit=400):
    """Gom các hạng mục gán theo tiêu đề cha / chưa có mã theo tiêu đề gần nhất -> soát 1 lần cho cả nhóm."""
    groups = {}
    for x in rows:
        if x["kind"] != "item" or x["cl_source"] not in ("context", "none", "inherit", "system"):
            continue
        head = group_key(x)
        g = groups.setdefault(head, {"n": 0, "codes": {}, "rows": [], "samples": []})
        g["n"] += 1
        g["codes"][x["cl"] or "??"] = g["codes"].get(x["cl"] or "??", 0) + 1
        g["rows"].append(x["src_row"])
        if len(g["samples"]) < 2:
            g["samples"].append(" ".join(str(x["desc"] or "").split())[:35])
    items = sorted(groups.items(), key=lambda kv: (-("??" in kv[1]["codes"]), -kv[1]["n"]))
    return items[:limit], len(items)


def cmd_review(a):
    data = json.loads(Path(a.rows).read_text(encoding="utf-8"))
    if a.only == "group":
        for sh in data["sheets"]:
            items, total = review_groups(sh["rows"])
            print(f"\n===== {sh['output_sheet']}: {total} nhóm tiêu đề (hiện {len(items)})")
            for head, g in items:
                codes = ", ".join(f"{k}×{v}" for k, v in sorted(g["codes"].items(), key=lambda kv: -kv[1]))
                print(f"{g['n']:>5} dòng | {codes:<22} | ctx:{head}  vd: {' / '.join(g['samples'])}  [dòng {g['rows'][0]}…{g['rows'][-1]}]")
        return
    for sh in data["sheets"]:
        print(f"\n===== {sh['output_sheet']}")
        for x in sh["rows"]:
            if a.only == "none" and not (x["kind"] == "item" and not x["cl"]):
                continue
            if a.only == "context" and x["cl_source"] != "context":
                continue
            if a.only == "items" and x["kind"] != "item":
                continue
            tag = x["cl"] or ("??" if x["kind"] == "item" else "")
            lvl = f"S{x['level']}" if x["kind"] == "section" else x["kind"][0].upper()
            ctx = (" < " + x["context"][0][:35]) if x["kind"] == "item" and x["context"] else ""
            print(f"{x['src_row']:>5} {lvl:<3} {str(x['stt'] or ''):<6} {tag:<6} {str(x['unit'] or ''):<6} {str(x['desc'] or '')[:60]}{ctx}")


# ----------------------------------------------------------------------------
# build
# ----------------------------------------------------------------------------
DIENGIAI = [("Trục chính", ["Trục 1", "Trục 2"]), ("Tầng Hầm", ["Hầm 1", "Hầm 2", "Hầm 3"]),
            ("Khối đế phía trên", ["Tầng 1", "Tầng 2", "Tầng 3"]), ("Hành lang điển hình", ["HL1", "HL2", "HL3"]),
            ("Căn hộ điển hình", ["CH-01", "CH-02", "CH-03", "CH-04", "CH-05"]), ("Khu vực khác", ["Khác 1", "Khác 2"])]
STYLE_TEMPLATE = HERE.parent / "assets" / "style_template.xlsx"
# Dòng mẫu style trong sheet _TPL (xem make_style_template.py)
TPL_ROW = {"sec0": 13, "sec1": 14, "sec2": 15, "sec3": 16, "note": 17, "note_bold": 18, "item": 19, "total": 20}


def layout(nz):
    c = {"stt": 1, "desc": 2, "spec": 3, "code": 4, "brand": 5, "origin": 6, "unit": 7}
    c["zones"] = [8 + i for i in range(nz)]
    c["total"] = 8 + nz
    s = c["total"] + 1
    c.update({"r_desc": s, "r_code": s + 1, "r_brand": s + 2, "r_unit": s + 3, "p_vl": s + 4, "p_nc": s + 5,
              "cl": s + 6, "inan": s + 7, "t_kl": s + 8, "dg0": s + 9, "dg1": s + 26, "kl": s + 27,
              "an": s + 28, "ao": s + 29, "ap": s + 30})
    c["amt"] = [(s + 31 + 3 * i, s + 32 + 3 * i, s + 33 + 3 * i) for i in range(nz)] if nz > 1 else []
    t0 = s + 31 + 3 * len(c["amt"])
    c["amt_total"] = (t0, t0 + 1, t0 + 2)
    c["note"] = t0 + 3
    c["hs_chung"] = t0 + 2
    c["lam_tron"] = t0 + 3
    return c


def col_map(c):
    """Cột đầu ra -> cột tương ứng trong sheet mẫu (mẫu có 3 khu: A..BC)."""
    t = layout(3)
    m = {}
    for k in ("stt", "desc", "spec", "code", "brand", "origin", "unit", "total", "r_desc", "r_code", "r_brand",
              "r_unit", "p_vl", "p_nc", "cl", "inan", "t_kl", "kl", "an", "ao", "ap", "note"):
        m[c[k]] = t[k]
    for i, zc in enumerate(c["zones"]):
        m[zc] = t["zones"][i % 3]
    for off in range(18):
        m[c["dg0"] + off] = t["dg0"] + off
    for i, grp in enumerate(c["amt"]):
        for j, col in enumerate(grp):
            m[col] = t["amt"][i % 3][j]
    for j, col in enumerate(c["amt_total"]):
        m[col] = t["amt_total"][j]
    return m


def style_row(ws, tpl, r, tpl_r, cmap, last_col):
    for col in range(1, last_col + 1):
        ws.cell(r, col)._style = copy.copy(tpl.cell(tpl_r, cmap[col])._style)


def put(ws, r, c, v):
    ws.cell(r, c).value = v


def hdr(ws, r1, c1, r2, c2, text):
    if (r1, c1) != (r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    ws.cell(r1, c1).value = text


def write_header(ws, tpl, cfg, sc, c, cmap):
    last = c["note"]
    # Dòng 1-12: chép style từ mẫu (riêng dòng 7-8 vùng đơn giá/thành tiền của mẫu là bảng phụ, bỏ qua)
    skip = set(range(c["an"], c["amt_total"][1] + 1))  # giữ ô "HS chung"
    for r in range(1, 13):
        for col in range(1, last + 1):
            if r in (7, 8) and col in skip:
                continue
            ws.cell(r, col)._style = copy.copy(tpl.cell(r, cmap[col])._style)
        ws.row_dimensions[r].height = tpl.row_dimensions[r].height
    company = cfg.get("company", {})
    for i, k in enumerate(["name", "name_en", "address", "contact"], 1):
        put(ws, i, last, company.get(k) or tpl.cell(i, 55).value)
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=last)
    put(ws, 5, 1, "CHI TIẾT GIÁ")
    put(ws, 6, 1, "=Summary!A32")
    put(ws, 7, 1, "Hạng mục: " + sc.get("title", sc["output_sheet"]))
    put(ws, 6, c["hs_chung"], "=Summary!G32")
    put(ws, 6, c["lam_tron"], "=Summary!H32")
    put(ws, 7, c["hs_chung"], "=Summary!G33")
    put(ws, 7, c["lam_tron"], "=Summary!H33")
    hao = cfg.get("coefficients", {}).get("hao_hut", 1.02)
    for col, lab, val in [(c["r_unit"], "HS hao hụt", hao), (c["p_vl"], "HSVL", "=Summary!C8"),
                          (c["p_nc"], "HSNC", "=Summary!E8"), (c["cl"], "HSTB", "=Summary!D8"),
                          (c["t_kl"], "HSKL", 1), (c["hs_chung"], "HS chung", "=Summary!F8"),
                          (c["lam_tron"], "LÀM TRÒN", "=Summary!G8")]:
        put(ws, 8, col, lab)
        put(ws, 9, col, val)
    ws.cell(9, c["lam_tron"]).number_format = "General"
    # Số / Ngày báo giá: tối thiểu cỡ chữ 10 cho dễ đọc
    for r in (6, 7):
        for col in (c["hs_chung"], c["lam_tron"]):
            f = copy.copy(ws.cell(r, col).font)
            if (f.sz or 11) < 10:
                f.sz = 11
                ws.cell(r, col).font = f
    ws.cell(7, c["lam_tron"]).number_format = "dd/mm/yyyy"
    for key, text in [("stt", "STT"), ("desc", "Hạng mục công việc"), ("spec", "Thông số kỹ thuật"),
                      ("code", "Mã hiệu"), ("brand", "Nhãn hiệu"), ("origin", "Xuất xứ"), ("unit", "Đơn vị")]:
        hdr(ws, 10, c[key], 12, c[key], text)
    hdr(ws, 10, c["zones"][0], 10, c["total"], "Quantity/ Khối lượng")
    for i, z in enumerate(sc["zones"]):
        hdr(ws, 11, c["zones"][i], 12, c["zones"][i], z["name"])
    hdr(ws, 11, c["total"], 12, c["total"], "Tổng khối lượng")
    for key, text in [("r_desc", "Nội Dung Công Việc"), ("r_code", "Quy cách\nMã Hiệu"),
                      ("r_brand", "Nhãn Hiệu\nXuất Xứ"), ("r_unit", "Đơn Vị"), ("t_kl", "Tổng KL tạm tính"),
                      ("kl", "KL"), ("note", "Ghi Chú")]:
        hdr(ws, 10, c[key], 12, c[key], text)
    hdr(ws, 10, c["p_vl"], 10, c["p_nc"], "Đơn giá gốc")
    hdr(ws, 11, c["p_vl"], 12, c["p_vl"], "Vật liệu")
    hdr(ws, 11, c["p_nc"], 12, c["p_nc"], "Nhân Công")
    hdr(ws, 10, c["cl"], 10, c["inan"], "Bộ lọc")
    hdr(ws, 11, c["cl"], 12, c["cl"], "Chủng loại")
    hdr(ws, 11, c["inan"], 12, c["inan"], "In ấn")
    hdr(ws, 10, c["dg0"], 10, c["dg1"], "Diễn giải chi tiết KL")
    col = c["dg0"]
    for grp, subs in DIENGIAI:
        hdr(ws, 11, col, 11, col + len(subs) - 1, grp)
        for j, sub in enumerate(subs):
            put(ws, 12, col + j, sub)
        col += len(subs)
    hdr(ws, 10, c["kl"], 12, c["kl"], "KL")
    hdr(ws, 10, c["an"], 10, c["ap"], "Đơn giá")
    for key, text in [("an", "Vật Liệu"), ("ao", "Nhân công"), ("ap", "ĐG Tổng")]:
        hdr(ws, 11, c[key], 12, c[key], text)
    groups = [(z["name"], c["amt"][i]) for i, z in enumerate(sc["zones"])] if c["amt"] else []
    groups.append(("Total", c["amt_total"]))
    for name, (a1, a2, a3) in groups:
        hdr(ws, 10, a1, 10, a3, name)
        hdr(ws, 11, a1, 11, a2, "Amount/ thành tiền")
        hdr(ws, 11, a3, 12, a3, "Total Amount/ Tổng thành tiền")
        put(ws, 12, a1, "Supply/ Cung cấp")
        put(ws, 12, a2, "Installation/ lắp đặt")


# ----------------------------------------------------------------------------
# Dễ đọc: độ rộng cột theo nội dung, xuống dòng, chiều cao dòng tự tính
# (openpyxl không tự co giãn dòng; Excel cũng không tự giãn dòng có ô merge)
# ----------------------------------------------------------------------------
LINE_H = 15.0  # pt cho Times New Roman 11


def _lines(text, width, bold=False):
    if text is None or text == "":
        return 1
    cpl = max(1.0, (width - 1.5) * (1.15 if bold else 1.28))  # số ký tự/dòng ước lượng
    n = 0
    for part in str(text).split("\n"):
        words, cur, k = part.split(" "), 0, 1
        for w in words:
            lw = len(w)
            if cur and cur + 1 + lw > cpl:
                k += 1
                cur = lw
            else:
                cur += (1 if cur else 0) + lw
            while cur > cpl:  # từ quá dài
                k += 1
                cur -= cpl
        n += k
    return n


def _align(cell, wrap=True, horizontal=None, vertical="center", indent=None):
    a = copy.copy(cell.alignment)
    cell.alignment = Alignment(horizontal=horizontal if horizontal is not None else a.horizontal,
                               vertical=vertical, wrap_text=wrap,
                               indent=a.indent if indent is None else indent, text_rotation=a.text_rotation)


def make_readable(ws, c, rows, first, last, tr):
    """Đặt độ rộng cột, căn chỉnh, xuống dòng và chiều cao dòng để không mất chữ."""
    items = rows

    def maxlen(key, cap=200):
        vals = [len(str(x.get(key) or "")) for x in items]
        return min(max(vals or [0]), cap)

    def qty_len(i):
        m = 0
        for x in items:
            q = x["qty"][i] if i < len(x["qty"]) else None
            if isinstance(q, (int, float)):
                m = max(m, len(f"{q:,.2f}"))
        return m

    W = {}
    W[c["stt"]] = min(max(5.5, maxlen("stt") * 1.1 + 2), 9)
    W[c["desc"]] = 46
    W[c["r_desc"]] = 46
    W[c["spec"]] = 0 if maxlen("spec") == 0 else min(max(16, maxlen("spec") * 0.5), 32)
    for k, key, lo, hi in (("code", "code_cdt", 12, 26), ("brand", "brand", 12, 22), ("origin", "origin", 12, 20)):
        n = maxlen(key)
        W[c[k]] = 0 if n == 0 else min(max(lo, n * 0.8 + 2), hi)
    W[c["unit"]] = min(max(8, maxlen("unit") + 3), 12)
    W[c["r_unit"]] = W[c["unit"]]
    for i, zc in enumerate(c["zones"]):
        W[zc] = max(13, qty_len(i) + 4)
    W[c["total"]] = max(14, max([qty_len(i) for i in range(len(c["zones"]))] or [0]) + 5)
    W[c["r_code"]], W[c["r_brand"]] = 18, 18
    W[c["p_vl"]], W[c["p_nc"]] = 15, 15
    W[c["cl"]], W[c["inan"]] = 9.5, 8
    W[c["t_kl"]], W[c["kl"]] = 13, 13
    W[c["an"]], W[c["ao"]], W[c["ap"]] = 15, 15, 16
    for grp in c["amt"]:
        for col in grp:
            W[col] = 17
    for col in c["amt_total"]:
        W[col] = 19
    W[c["note"]] = 24
    for col, w in W.items():
        dim = ws.column_dimensions[L(col)]
        if w == 0:  # cột CĐT không có dữ liệu: thu hẹp nhưng vẫn đủ đọc tiêu đề
            w = 11
        dim.width = round(w, 1)

    text_cols = {c[k]: k for k in ("stt", "desc", "spec", "code", "brand", "origin", "unit", "r_desc", "r_code",
                                   "r_brand", "r_unit", "cl", "inan", "note")}
    center_cols = {c["stt"], c["unit"], c["r_unit"], c["cl"], c["inan"]}
    num_cols = set(c["zones"]) | {c["total"], c["p_vl"], c["p_nc"], c["t_kl"], c["kl"], c["an"], c["ao"], c["ap"]}
    for grp in c["amt"] + [c["amt_total"]]:
        num_cols |= set(grp)

    # Tiêu đề bảng (dòng 10-12): căn giữa, xuống dòng
    for r in (10, 11, 12):
        for col in range(1, c["note"] + 1):
            _align(ws.cell(r, col), True, "center", "center")
    ws.row_dimensions[10].height = 33
    ws.row_dimensions[11].height = 33
    ws.row_dimensions[12].height = 33

    # Dữ liệu
    for x in rows + [{"_r": tr, "desc": "Tổng cộng trước thuế", "kind": "total"}]:
        r = x["_r"]
        bold = x["kind"] in ("section", "total") or x.get("bold", False)
        for col in range(1, c["note"] + 1):
            cell = ws.cell(r, col)
            if col in text_cols:
                if col in center_cols:
                    _align(cell, True, "center")
                elif col in (c["desc"], c["r_desc"], c["spec"], c["note"]):
                    _align(cell, True, "left", indent=1 if x["kind"] != "note" else 0)
                else:
                    _align(cell, True, "center")
            elif col in num_cols:
                _align(cell, False, "right", indent=1)
        n = 1
        for key, ck in (("desc", "desc"), ("spec", "spec"), ("code_cdt", "code"), ("brand", "brand"),
                        ("origin", "origin"), ("unit", "unit"), ("stt", "stt")):
            if x.get(key):
                n = max(n, _lines(x.get(key), ws.column_dimensions[L(c[ck])].width, bold))
        if x.get("note"):
            n = max(n, _lines(x["note"], ws.column_dimensions[L(c["note"])].width))
        ws.row_dimensions[r].height = min(409, max(18.0, n * LINE_H + 4))


def section_end(rows, i, last_excel_row):
    lv = rows[i]["level"]
    for j in range(i + 1, len(rows)):
        if rows[j]["kind"] == "section" and rows[j]["level"] <= lv:
            return rows[j]["_r"] - 1
    return last_excel_row


def build_sheet(wb, tpl, cfg, sc, rows):
    nz = len(sc["zones"])
    c = layout(nz)
    cmap = col_map(c)
    last_col = c["note"]
    ws = wb.create_sheet(sc["output_sheet"][:31])
    write_header(ws, tpl, cfg, sc, c, cmap)
    first = 13
    for i, x in enumerate(rows):
        x["_r"] = first + i
    last = first + len(rows) - 1
    Ls = {k: L(v) for k, v in c.items() if isinstance(v, int)}
    zL = [L(z) for z in c["zones"]]
    tot = Ls["total"]
    P9, Q9, R9, O9, T9 = (f"${Ls['p_vl']}$9", f"${Ls['p_nc']}$9", f"${Ls['cl']}$9", f"${Ls['r_unit']}$9", f"${Ls['t_kl']}$9")
    BB9, BC9 = f"${Ls['hs_chung']}$9", f"${Ls['lam_tron']}$9"
    for i, x in enumerate(rows):
        r = x["_r"]
        if x["kind"] == "section":
            key = f"sec{min(max(x.get('level') or 1, 0), 3)}"
        elif x["kind"] == "note":
            key = "note_bold" if x.get("bold") else "note"
        else:
            key = "item"
        style_row(ws, tpl, r, TPL_ROW[key], cmap, last_col)
        # ---- trái: dữ liệu CĐT giữ nguyên ----
        put(ws, r, c["stt"], x["stt"])
        put(ws, r, c["desc"], x["desc"])
        for k, key in (("spec", "spec"), ("code", "code_cdt"), ("brand", "brand"), ("origin", "origin")):
            put(ws, r, c[k], x.get(key))
        put(ws, r, c["unit"], x["unit"])
        for zi, zc in enumerate(c["zones"]):
            put(ws, r, zc, x["qty"][zi])
        # ---- phải: template TDI ----
        put(ws, r, c["r_desc"], f"=+B{r}")
        if x["kind"] == "section":
            end = section_end(rows, i, last)
            # [FIX audit #3] tiêu đề không có dòng con (end <= r): không ghi SUBTOTAL, tránh vòng tham chiếu vào dòng tổng
            if end > r:
                for grp in c["amt"] + [c["amt_total"]]:
                    for col in grp:
                        put(ws, r, col, f"=+SUBTOTAL(9,{L(col)}{r + 1}:{L(col)}{end})")
            continue
        if x["kind"] == "note":
            # [FIX #VALUE 26/10/2026] dòng ghi chú / tiêu đề nhận nhầm: không ghi công thức giá và thành tiền (ô KL CĐT có thể là chữ / ký tự trắng)
            put(ws, r, c["note"], x.get("note"))
            continue
        put(ws, r, c["total"], "=+" + "+".join(f"{z}{r}" for z in zL))
        put(ws, r, c["r_code"], x.get("code_tdi"))
        put(ws, r, c["r_brand"], x.get("brand_tdi"))
        put(ws, r, c["r_unit"], f"=+G{r}")
        put(ws, r, c["p_vl"], x.get("p_vl"))
        put(ws, r, c["p_nc"], x.get("p_nc"))
        put(ws, r, c["cl"], x.get("cl"))
        put(ws, r, c["inan"], x.get("inan"))
        put(ws, r, c["t_kl"], f"=+{tot}{r}")
        O, T = f"{Ls['r_unit']}{r}", f"{Ls['t_kl']}{r}"
        put(ws, r, c["kl"], f'=IF(OR({O}="m",{O}="m2",{O}="m3"),{T}*{T9},{T})')
        P, Q = f"{Ls['p_vl']}{r}", f"{Ls['p_nc']}{r}"
        g = price_group(x.get("cl"), x.get("unit"))
        if g == "hao_hut":
            an = f"=ROUNDDOWN({P}*{P9}*{BB9}*{O9},{BC9})"
        elif g == "thiet_bi":
            an = f"=ROUNDDOWN({P}*{R9}*{BB9},{BC9})"
        else:
            an = f"=ROUNDDOWN({P}*{P9}*{BB9},{BC9})"
        AN, AO, AP = f"{Ls['an']}{r}", f"{Ls['ao']}{r}", f"{Ls['ap']}{r}"
        put(ws, r, c["an"], an)
        put(ws, r, c["ao"], f"=ROUNDDOWN({Q}*{Q9}*{BB9},{BC9})")
        put(ws, r, c["ap"], f"={AN}+{AO}")
        for zi, (a1, a2, a3) in enumerate(c["amt"]):
            q = f"{zL[zi]}{r}"
            put(ws, r, a1, f"=+{AN}*N({q})")
            put(ws, r, a2, f"=+{AO}*N({q})")
            put(ws, r, a3, f"=+{AP}*N({q})")
        t1, t2, t3 = c["amt_total"]
        put(ws, r, t1, f"={AN}*N(${tot}{r})")
        put(ws, r, t2, f"={AO}*N(${tot}{r})")
        put(ws, r, t3, f"={AP}*N(${tot}{r})")
        put(ws, r, c["note"], x.get("note"))
    # Dòng tổng cộng
    tr = last + 1
    style_row(ws, tpl, tr, TPL_ROW["total"], cmap, last_col)
    put(ws, tr, c["desc"], "Tổng cộng trước thuế")
    put(ws, tr, c["r_desc"], f"=+B{tr}")
    for grp in c["amt"] + [c["amt_total"]]:
        for col in grp:
            put(ws, tr, col, f"=SUBTOTAL(9,{L(col)}{first}:{L(col)}{last})")
    # Độ rộng cột theo mẫu, rồi chỉnh lại cho vừa nội dung
    for col in range(1, last_col + 1):
        src = tpl.column_dimensions[L(cmap[col])]
        ws.column_dimensions[L(col)].width = src.width if src.width else 13
    make_readable(ws, c, rows, first, last, tr)
    for col in range(c["dg0"], c["dg1"] + 1):
        ws.column_dimensions[L(col)].outlineLevel = 1
        ws.column_dimensions[L(col)].hidden = True
    ws.freeze_panes = "C13"
    ws.auto_filter.ref = f"A12:{L(last_col)}{last}"
    ws.sheet_view.zoomScale = 85
    ws.sheet_properties.tabColor = tpl.sheet_properties.tabColor
    return {"sheet": ws.title, "total_row": tr, "vl": L(c["amt_total"][0]), "nc": L(c["amt_total"][1]),
            "tong": L(c["amt_total"][2]), "title": sc.get("summary_name") or sc.get("title", sc["output_sheet"])}


def build_summary(sm, cfg, infos):
    """Điền sheet Summary (chép từ mẫu): hệ số, thông tin dự án, danh sách hệ thống."""
    co = cfg.get("coefficients", {})
    for col, key, dflt in [(2, "HSKL", 1), (3, "HSVL", 1.25), (4, "HSTB", 1.1), (5, "HSNC", 1.15),
                           (6, "HS_chung", 1), (7, "lam_tron", -2)]:
        sm.cell(8, col).value = co.get(key, dflt)
    if cfg.get("project"):
        sm["A32"] = "Dự án: " + cfg["project"]
    if cfg.get("hang_muc"):
        sm["A33"] = "Hạng mục: " + cfg["hang_muc"]
    sm["H32"] = cfg.get("so_bao_gia") or sm["H32"].value
    if cfg.get("ngay"):
        sm["H33"] = cfg["ngay"]
    # Bảng diện tích của dự án mẫu -> để trống cho dự án mới, chống chia cho 0
    for r in range(25, 31):
        sm.cell(r, 5).value = None
        sm.cell(r, 6).value = None
    sm["H25"] = "=SUM(F25:F30)"
    sm["H26"] = "=IFERROR(G57/H25,0)"
    sm["H27"] = "=IFERROR(SUBTOTAL(9,E39:E50)/H25,0)"
    sm["H28"] = "=IFERROR(SUBTOTAL(9,F39:F50)/H25,0)"
    # Danh sách hệ thống: dòng 39..50 (style lấy theo dòng 39 của mẫu)
    if len(infos) > 12:
        print("CẢNH BÁO: Summary mẫu chỉ có chỗ cho 12 hệ thống; các sheet sau không được tổng hợp.")
    row_style = {col: copy.copy(sm.cell(39, col)._style) for col in range(1, 9)}
    blank_style = {col: copy.copy(sm.cell(40, col)._style) for col in range(1, 9)}
    for r in range(39, 51):
        for col in range(1, 9):
            sm.cell(r, col).value = None
            sm.cell(r, col)._style = copy.copy(blank_style[col])
    for i, inf in enumerate(infos[:12]):
        r = 39 + i
        for col in range(1, 9):
            sm.cell(r, col)._style = copy.copy(row_style[col])
        sq = "'" + inf["sheet"].replace("'", "''") + "'"
        sm.cell(r, 1).value = chr(ord("A") + i)
        sm.cell(r, 2).value = inf["title"]
        sm.cell(r, 3).value = "Hệ"
        sm.cell(r, 4).value = 1
        sm.cell(r, 5).value = f"=+{sq}!{inf['vl']}{inf['total_row']}"
        sm.cell(r, 6).value = f"=+{sq}!{inf['nc']}{inf['total_row']}"
        sm.cell(r, 7).value = f"=+{sq}!{inf['tong']}{inf['total_row']}"
    sm["G38"] = "=SUM(G39:G50)"
    for i, inf in enumerate(infos[:12]):
        cell = sm.cell(39 + i, 2)
        _align(cell, True, "left")
        sm.row_dimensions[39 + i].height = max(20.0, _lines(inf["title"], sm.column_dimensions["B"].width or 45, True) * LINE_H + 5)
    for col in "EFG":
        idx = cidx(col)
        dim = next((d for d in sm.column_dimensions.values() if d.min and d.max and d.min <= idx <= d.max), None)
        if dim is None:
            dim = sm.column_dimensions[col]
        if (dim.width or 0) < 18:
            dim.width = 18


def apply_overrides(data, ov):
    for sh in data["sheets"]:
        m = ov.get(sh["output_sheet"], {})
        # "ctx:<khóa nhóm như review --only group in ra>": "MÃ" -> áp cho cả nhóm (trừ dòng đã có override riêng)
        heads = {k[4:].strip(): v for k, v in m.items() if k.startswith("ctx:")}
        if heads:
            for x in sh["rows"]:
                h = group_key(x)
                if x["kind"] == "item" and h in heads and str(x["src_row"]) not in m:
                    x["cl"], x["cl_source"] = heads[h], "claude"
        for x in sh["rows"]:
            o = m.get(str(x["src_row"]))
            if o is None:
                continue
            if isinstance(o, str):
                x["cl"], x["cl_source"] = o, "claude"
            else:
                for k, v in o.items():
                    x[k] = v
                if "cl" in o:
                    x["cl_source"] = "claude"
                if o.get("kind") == "item" and "cl" not in o and not x.get("cl"):
                    x["cl"], x["cl_source"] = classify(x["desc"], x.get("context", []), x.get("system", ""))
                if o.get("kind") in ("section", "note") and "level" not in o:
                    x["level"] = x.get("level") if o.get("kind") == "section" else None


# ============================================================================
# CHẾ ĐỘ IN-PLACE (mặc định): giữ nguyên file + sheet CĐT, dựng phần TDI ngay bên phải
# ============================================================================
import colorsys  # noqa: E402
import zipfile  # noqa: E402
from openpyxl.styles import Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.styles.colors import Color  # noqa: E402


class ThemeRGB:
    """Đổi màu theme của file mẫu TDI sang RGB, để màu TDI không đổi khi dán vào workbook CĐT (theme khác)."""

    def __init__(self, wb):
        xml = (wb.loaded_theme or b"").decode("utf-8", "ignore")
        names = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
                 "hlink", "folHlink"]
        found = {}
        for n in names:
            m = re.search(rf"<a:{n}>.*?(?:srgbClr val=\"([0-9A-Fa-f]{{6}})\"|lastClr=\"([0-9A-Fa-f]{{6}})\")", xml, re.S)
            found[n] = (m.group(1) or m.group(2)) if m else "000000"
        self.pal = [found["lt1"], found["dk1"], found["lt2"], found["dk2"]] + \
                   [found[f"accent{i}"] for i in range(1, 7)] + [found["hlink"], found["folHlink"]]
        self.cache = {}

    def color(self, c):
        if c is None or getattr(c, "type", None) != "theme":
            return copy.copy(c) if c is not None else None
        key = (c.theme, round(c.tint or 0, 4))
        if key not in self.cache:
            base = self.pal[c.theme] if c.theme < len(self.pal) else "000000"
            r, g, b = [int(base[i:i + 2], 16) / 255 for i in (0, 2, 4)]
            h, l, sat = colorsys.rgb_to_hls(r, g, b)
            t = c.tint or 0
            l = l * (1 + t) if t < 0 else l * (1 - t) + t
            r, g, b = colorsys.hls_to_rgb(h, l, sat)
            self.cache[key] = "FF%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))
        return Color(rgb=self.cache[key])

    def apply(self, dst, src):
        """Chép style ô src (workbook mẫu) sang ô dst (workbook khác), màu theme -> RGB."""
        f = copy.copy(src.font)
        dst.font = Font(name=f.name, sz=f.sz, b=f.b, i=f.i, u=f.u, strike=f.strike, vertAlign=f.vertAlign,
                        color=self.color(f.color))
        fl = src.fill
        if getattr(fl, "fill_type", None):
            dst.fill = PatternFill(fill_type=fl.fill_type, fgColor=self.color(fl.fgColor), bgColor=self.color(fl.bgColor))
        else:
            dst.fill = PatternFill(fill_type=None)
        bd = src.border
        side = lambda sd: Side(style=sd.style, color=self.color(sd.color)) if sd is not None and sd.style else Side()  # noqa: E731
        dst.border = Border(left=side(bd.left), right=side(bd.right), top=side(bd.top), bottom=side(bd.bottom))
        dst.alignment = copy.copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy.copy(src.protection)


def last_used_col(ws):
    m = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value not in (None, ""):
                m = max(m, cell.column)
    for mr in ws.merged_cells.ranges:
        m = max(m, mr.max_col)
    return m


def layout_inplace(sc, ws):
    """Cột CĐT giữ nguyên vị trí; cột TDI bắt đầu ngay sau cột cuối cùng CĐT đang dùng."""
    nz = len(sc["zones"])
    base = layout(nz)
    start = cidx(sc["tdi_start_col"]) if sc.get("tdi_start_col") else last_used_col(ws) + 1
    shift = start - base["r_desc"]
    c = {}
    for k, v in base.items():
        if k in ("stt", "desc", "spec", "code", "brand", "origin", "unit", "zones", "total"):
            continue
        if isinstance(v, int):
            c[k] = v + shift
        elif k == "amt":
            c[k] = [tuple(x + shift for x in g) for g in v]
        elif k == "amt_total":
            c[k] = tuple(x + shift for x in v)
    cols = sc["columns"]
    for k in ("stt", "desc", "spec", "code", "brand", "origin", "unit"):
        c[k] = cidx(cols[k]) if cols.get(k) else None
    c["zones"] = [cidx(z["col"]) for z in sc["zones"]]
    c["total"] = cidx(sc["total_col"]) if sc.get("total_col") else None
    c["_shift"] = shift
    return c


def col_map_right(c):
    """Cột TDI đầu ra -> cột tương ứng trong sheet mẫu _TPL (mẫu 3 khu, phần TDI bắt đầu ở L)."""
    t = layout(3)
    m = {}
    for k in ("r_desc", "r_code", "r_brand", "r_unit", "p_vl", "p_nc", "cl", "inan", "t_kl", "kl", "an", "ao", "ap", "note"):
        m[c[k]] = t[k]
    for off in range(18):
        m[c["dg0"] + off] = t["dg0"] + off
    for i, grp in enumerate(c["amt"]):
        for j, col in enumerate(grp):
            m[col] = t["amt"][i % 3][j]
    for j, col in enumerate(c["amt_total"]):
        m[col] = t["amt_total"][j]
    return m


def build_inplace_sheet(wb, tpl, th, cfg, sc, rows, sum_name):
    ws = wb[sc["source_sheet"]]
    c = layout_inplace(sc, ws)
    cmap = col_map_right(c)
    right = list(range(c["r_desc"], c["note"] + 1))
    Lc = lambda k: L(c[k])  # noqa: E731
    zL = [L(z) for z in c["zones"]]
    desc_L = L(c["desc"])
    unit_L = L(c["unit"]) if c.get("unit") else None
    S = "'" + sum_name.replace("'", "''") + "'"
    # ---- vị trí tiêu đề TDI: căn theo dòng tiêu đề bảng của CĐT
    h2 = (sc.get("header_rows") or [None, sc["data_start_row"] - 1])[-1]
    flat = h2 < 3
    tier = {1: h2 if flat else h2 - 2, 2: h2 if flat else h2 - 1, 3: h2}
    coef_ok = (not flat) and tier[1] - 2 >= 1
    lab_r, val_r = (tier[1] - 2, tier[1] - 1) if coef_ok else (None, None)
    if coef_ok:
        ref = lambda k: f"${Lc(k)}${val_r}"  # noqa: E731
        P9, Q9, R9, O9, T9, BB9, BC9 = (ref("p_vl"), ref("p_nc"), ref("cl"), ref("r_unit"), ref("t_kl"),
                                        ref("hs_chung"), ref("lam_tron"))
    else:  # không còn dòng trống phía trên -> tham chiếu thẳng hệ số ở Summary
        P9, Q9, R9, O9, T9, BB9, BC9 = (f"{S}!$C$8", f"{S}!$E$8", f"{S}!$D$8", f"{S}!$H$8", f"{S}!$B$8",
                                        f"{S}!$F$8", f"{S}!$G$8")

    def sty(r, col, tpl_r):
        th.apply(ws.cell(r, col), tpl.cell(tpl_r, cmap[col]))

    # ---- hệ số
    if coef_ok:
        for k, lab, val in [("r_unit", "HS hao hụt", f"={S}!H8"), ("p_vl", "HSVL", f"={S}!C8"), ("p_nc", "HSNC", f"={S}!E8"),
                            ("cl", "HSTB", f"={S}!D8"), ("t_kl", "HSKL", f"={S}!B8"), ("hs_chung", "HS chung", f"={S}!F8"),
                            ("lam_tron", "LÀM TRÒN", f"={S}!G8")]:
            sty(lab_r, c[k], 8)
            sty(val_r, c[k], 9)
            ws.cell(lab_r, c[k]).value = lab
            ws.cell(val_r, c[k]).value = val
        ws.cell(val_r, c["lam_tron"]).number_format = "General"
    # ---- tiêu đề bảng 3 tầng
    for t in (1, 2, 3):
        for col in right:
            sty(tier[t], col, 9 + t)

    def hdr(t1, c1, t2, c2, text):
        r1, r2 = tier[t1], tier[t2]
        if flat and t2 < 3:
            return  # chỉ còn 1 dòng: bỏ nhãn nhóm, giữ nhãn chi tiết
        if (r1, c1) != (r2, c2):
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        ws.cell(r1, c1).value = text

    for k, text in [("r_desc", "Nội Dung Công Việc"), ("r_code", "Quy cách\nMã Hiệu"), ("r_brand", "Nhãn Hiệu\nXuất Xứ"),
                    ("r_unit", "Đơn Vị"), ("t_kl", "Tổng KL tạm tính"), ("kl", "KL"), ("note", "Ghi Chú")]:
        hdr(1, c[k], 3, c[k], text)
    hdr(1, c["p_vl"], 1, c["p_nc"], "Đơn giá gốc")
    hdr(2, c["p_vl"], 3, c["p_vl"], "Vật liệu")
    hdr(2, c["p_nc"], 3, c["p_nc"], "Nhân Công")
    hdr(1, c["cl"], 1, c["inan"], "Bộ lọc")
    hdr(2, c["cl"], 3, c["cl"], "Chủng loại")
    hdr(2, c["inan"], 3, c["inan"], "In ấn")
    hdr(1, c["dg0"], 1, c["dg1"], "Diễn giải chi tiết KL")
    col = c["dg0"]
    for grp, subs in DIENGIAI:
        hdr(2, col, 2, col + len(subs) - 1, grp)
        for j, sub in enumerate(subs):
            ws.cell(tier[3], col + j).value = sub
        col += len(subs)
    hdr(1, c["an"], 1, c["ap"], "Đơn giá")
    for k, text in [("an", "Vật Liệu"), ("ao", "Nhân công"), ("ap", "ĐG Tổng")]:
        hdr(2, c[k], 3, c[k], text)
    groups = [(z["name"], c["amt"][i]) for i, z in enumerate(sc["zones"])] if c["amt"] else []
    groups.append(("Total", c["amt_total"]))
    for name, (a1, a2, a3) in groups:
        hdr(1, a1, 1, a3, name)
        hdr(2, a1, 2, a2, "Amount/ thành tiền")
        hdr(2, a3, 3, a3, "Total Amount/ Tổng thành tiền")
        ws.cell(tier[3], a1).value = ("Cung cấp - " + name) if flat else "Supply/ Cung cấp"
        ws.cell(tier[3], a2).value = ("Lắp đặt - " + name) if flat else "Installation/ lắp đặt"
    for t in (1, 2, 3):
        for col in right:
            a = copy.copy(ws.cell(tier[t], col).alignment)
            ws.cell(tier[t], col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True,
                                                        indent=a.indent)
    # ---- dữ liệu
    first, last = rows[0]["src_row"], rows[-1]["src_row"]
    for x in rows:
        x["_r"] = x["src_row"]
    tot_col = L(c["total"]) if c.get("total") else Lc("t_kl")
    text_center = {c["r_code"], c["r_brand"], c["r_unit"], c["cl"], c["inan"]}
    num_cols = {c["p_vl"], c["p_nc"], c["t_kl"], c["kl"], c["an"], c["ao"], c["ap"], *c["amt_total"]}
    for g in c["amt"]:
        num_cols |= set(g)
    style_cache = {}
    for i, x in enumerate(rows):
        r = x["_r"]
        if x["kind"] == "section":
            key = f"sec{min(max(x.get('level') or 1, 0), 3)}"
        elif x["kind"] == "note":
            key = "note_bold" if x.get("bold") else "note"
        else:
            key = "item"
        # style phía TDI: tính 1 lần cho mỗi (loại dòng, cột) rồi chép chỉ số style -> nhanh với BOQ lớn
        for col in right:
            ck = (key, col)
            cell = ws.cell(r, col)
            if ck in style_cache:
                cell._style = copy.copy(style_cache[ck])
                continue
            sty(r, col, TPL_ROW[key])
            if col in (c["r_desc"], c["note"]):
                _align(cell, True, "left", indent=1 if x["kind"] != "note" else 0)
            elif col in text_center:
                _align(cell, True, "center")
            elif col in num_cols:
                _align(cell, False, "right", indent=1)
            style_cache[ck] = copy.copy(cell._style)
        put(ws, r, c["r_desc"], f"=+{desc_L}{r}")
        # giữ kiểu chữ (đậm / nghiêng / gạch chân) của mô tả CĐT cho cột Nội dung bên TDI
        sf = ws.cell(r, c["desc"]).font
        if sf.b or sf.i or sf.u:
            df = ws.cell(r, c["r_desc"]).font
            ws.cell(r, c["r_desc"]).font = Font(name=df.name, sz=df.sz, color=copy.copy(df.color),
                                                b=bool(sf.b) or bool(df.b), i=bool(sf.i), u=sf.u)
        if x["kind"] == "section":
            end = section_end(rows, i, last)
            # [FIX audit #3] tiêu đề không có dòng con: không ghi SUBTOTAL (tránh vòng tham chiếu)
            if end > r:
                for grp in c["amt"] + [c["amt_total"]]:
                    for col in grp:
                        put(ws, r, col, f"=+SUBTOTAL(9,{L(col)}{r + 1}:{L(col)}{end})")
        elif x["kind"] == "note":
            # [FIX #VALUE 26/10/2026] dòng ghi chú / tiêu đề nhận nhầm: không ghi công thức giá và thành tiền
            put(ws, r, c["note"], x.get("note"))
        else:
            put(ws, r, c["r_code"], x.get("code_tdi"))
            put(ws, r, c["r_brand"], x.get("brand_tdi"))
            put(ws, r, c["r_unit"], f"=+{unit_L}{r}" if unit_L else None)
            put(ws, r, c["p_vl"], x.get("p_vl"))
            put(ws, r, c["p_nc"], x.get("p_nc"))
            put(ws, r, c["cl"], x.get("cl"))
            put(ws, r, c["inan"], x.get("inan"))
            put(ws, r, c["t_kl"], f"=+{L(c['total'])}{r}" if c.get("total") else "=+" + "+".join(f"{z}{r}" for z in zL))
            O, T = f"{Lc('r_unit')}{r}", f"{Lc('t_kl')}{r}"
            put(ws, r, c["kl"], f'=IF(OR({O}="m",{O}="m2",{O}="m3"),{T}*{T9},{T})')
            P, Q = f"{Lc('p_vl')}{r}", f"{Lc('p_nc')}{r}"
            g = price_group(x.get("cl"), x.get("unit"))
            if g == "hao_hut":
                an = f"=ROUNDDOWN({P}*{P9}*{BB9}*{O9},{BC9})"
            elif g == "thiet_bi":
                an = f"=ROUNDDOWN({P}*{R9}*{BB9},{BC9})"
            else:
                an = f"=ROUNDDOWN({P}*{P9}*{BB9},{BC9})"
            AN, AO, AP = f"{Lc('an')}{r}", f"{Lc('ao')}{r}", f"{Lc('ap')}{r}"
            put(ws, r, c["an"], an)
            put(ws, r, c["ao"], f"=ROUNDDOWN({Q}*{Q9}*{BB9},{BC9})")
            put(ws, r, c["ap"], f"={AN}+{AO}")
            for zi, (a1, a2, a3) in enumerate(c["amt"]):
                q = f"{zL[zi]}{r}"
                put(ws, r, a1, f"=+{AN}*N({q})")
                put(ws, r, a2, f"=+{AO}*N({q})")
                put(ws, r, a3, f"=+{AP}*N({q})")
            t1, t2, t3 = c["amt_total"]
            put(ws, r, t1, f"={AN}*N(${tot_col}{r})")
            put(ws, r, t2, f"={AO}*N(${tot_col}{r})")
            put(ws, r, t3, f"={AP}*N(${tot_col}{r})")
            put(ws, r, c["note"], x.get("note"))
    # ---- dòng tổng phía TDI (chỉ ghi vào các cột TDI, không đụng dòng tổng của CĐT)
    tr = last + 1
    for col in right:
        sty(tr, col, TPL_ROW["total"])
    put(ws, tr, c["r_desc"], "Tổng cộng trước thuế (TDI)")
    for grp in c["amt"] + [c["amt_total"]]:
        for col in grp:
            put(ws, tr, col, f"=SUBTOTAL(9,{L(col)}{first}:{L(col)}{last})")
    # ---- độ rộng cột TDI: cột Nội dung rộng bằng cột mô tả CĐT để xuống dòng giống hệt
    dw = ws.column_dimensions[desc_L].width or 46
    widths = {c["r_desc"]: dw, c["r_code"]: 18, c["r_brand"]: 18, c["r_unit"]: 9, c["p_vl"]: 15, c["p_nc"]: 15,
              c["cl"]: 9.5, c["inan"]: 8, c["t_kl"]: 13, c["kl"]: 13, c["an"]: 15, c["ao"]: 15, c["ap"]: 16,
              c["note"]: 24}
    for g in c["amt"]:
        for col in g:
            widths[col] = 17
    for col in c["amt_total"]:
        widths[col] = 19
    for col, w in widths.items():
        ws.column_dimensions[L(col)].width = w
    for col in range(c["dg0"], c["dg1"] + 1):
        ws.column_dimensions[L(col)].width = 9
        ws.column_dimensions[L(col)].outlineLevel = 1
        ws.column_dimensions[L(col)].hidden = True
    if flat:
        print(f"   LƯU Ý [{ws.title}]: tiêu đề bảng CĐT ở dòng {h2} (< 3) nên tiêu đề TDI chỉ có 1 dòng, hệ số lấy thẳng từ {sum_name}.")
    return {"sheet": ws.title, "total_row": tr, "vl": L(c["amt_total"][0]), "nc": L(c["amt_total"][1]),
            "tong": L(c["amt_total"][2]), "title": sc.get("summary_name") or sc.get("title", ws.title),
            "tdi_cols": f"{L(c['r_desc'])}:{L(c['note'])}"}


def copy_sheet_across(src, dst, th):
    """Chép sheet Summary mẫu sang workbook CĐT: giá trị, style (màu -> RGB), merge, độ rộng, chiều cao, logo."""
    for row in src.iter_rows():
        for cell in row:
            if cell.__class__.__name__ == "MergedCell":
                continue
            d = dst.cell(cell.row, cell.column)
            d.value = cell.value
            if cell.has_style:
                th.apply(d, cell)
    for mr in src.merged_cells.ranges:
        dst.merge_cells(str(mr))
    for k, dim in src.column_dimensions.items():
        if dim.width:
            for ci in range((dim.min or cidx(k)), (dim.max or cidx(k)) + 1):
                if ci > 30:
                    break
                dst.column_dimensions[L(ci)].width = dim.width
    for r, dim in src.row_dimensions.items():
        if dim.height:
            dst.row_dimensions[r].height = dim.height
    try:
        from io import BytesIO
        from openpyxl.drawing.image import Image as XLImage
        for img in getattr(src, "_images", []):
            ni = XLImage(BytesIO(img._data()))
            ni.width, ni.height = img.width, img.height
            ni.anchor = img.anchor
            dst.add_image(ni)
    except Exception as e:  # logo không chép được thì bỏ qua
        print("   (không chép được logo Summary:", e, ")")
    dst.sheet_view.showGridLines = src.sheet_view.showGridLines


def build_inplace(cfg, data, out_path):
    src_path = cfg["input"]
    keep_vba = str(src_path).lower().endswith(".xlsm")
    wb = openpyxl.load_workbook(src_path, rich_text=True, keep_vba=keep_vba)
    try:
        with zipfile.ZipFile(src_path) as z:
            if any(n.startswith("xl/drawings/") and n.endswith(".xml") for n in z.namelist()):
                print("LƯU Ý: file CĐT có hình vẽ / shape / biểu đồ. openpyxl giữ được ảnh và biểu đồ cơ bản nhưng "
                      "có thể làm mất shape, textbox. Hãy mở file kết quả kiểm tra.")
    except zipfile.BadZipFile:
        pass
    twb = openpyxl.load_workbook(cfg.get("style_template") or str(STYLE_TEMPLATE))
    th = ThemeRGB(twb)
    tpl = twb["_TPL"]
    sum_name = "Summary"
    k = 1
    while sum_name in wb.sheetnames:
        k += 1
        sum_name = f"Summary TDI{'' if k == 2 else ' ' + str(k)}"
    by_name = {s["output_sheet"]: s for s in cfg["sheets"]}
    infos = []
    for sh in data["sheets"]:
        for x in sh["rows"]:
            if x["kind"] == "section" and x.get("level") is None:
                x["level"] = 1
        infos.append(build_inplace_sheet(wb, tpl, th, cfg, by_name[sh["output_sheet"]], sh["rows"], sum_name))
    sm = wb.create_sheet(sum_name)
    copy_sheet_across(twb["_SUM"], sm, th)
    build_summary(sm, cfg, infos)
    th.apply(sm["H7"], sm["G7"])
    th.apply(sm["H8"], sm["G8"])
    sm["H7"] = "HS hao hụt"
    sm["H8"] = cfg.get("coefficients", {}).get("hao_hut", 1.02)
    wb.calculation.fullCalcOnLoad = True  # Excel tự tính lại khi mở (không cần lưu qua LibreOffice)
    wb.save(out_path)
    for inf in infos:
        print(f"[{inf['sheet']}] phần TDI ở cột {inf['tdi_cols']}, dòng tổng TDI: {inf['total_row']}")
    return infos



def cmd_build(a):
    cfg = json.loads(Path(a.config).read_text(encoding="utf-8"))
    data = json.loads(Path(a.rows).read_text(encoding="utf-8"))
    if a.overrides:
        apply_overrides(data, json.loads(Path(a.overrides).read_text(encoding="utf-8")))
    if cfg.get("mode", "inplace") == "inplace":
        build_inplace(cfg, data, a.out)
        by_name = {s["output_sheet"]: s for s in cfg["sheets"]}
        for sh in data["sheets"]:
            nz = len(by_name[sh["output_sheet"]]["zones"])
            sums = [round(sum((x["qty"][i] or 0) for x in sh["rows"] if x["kind"] != "section" and isinstance(x["qty"][i], (int, float))), 3) for i in range(nz)]
            miss = [x["src_row"] for x in sh["rows"] if x["kind"] == "item" and not x.get("cl")]
            print(f"   tổng KL theo khu: {sums} | chưa có Chủng loại: {len(miss)} {miss[:20]}")
        print("Đã ghi", a.out)
        return
    template = cfg.get("style_template") or str(STYLE_TEMPLATE)
    wb = openpyxl.load_workbook(template)  # giữ nguyên theme màu của file mẫu TDI
    tpl, sm = wb["_TPL"], wb["_SUM"]
    infos = []
    by_name = {s["output_sheet"]: s for s in cfg["sheets"]}
    for sh in data["sheets"]:
        for x in sh["rows"]:
            if x["kind"] == "section" and x.get("level") is None:
                x["level"] = 1
        infos.append(build_sheet(wb, tpl, cfg, by_name[sh["output_sheet"]], sh["rows"]))
    build_summary(sm, cfg, infos)
    sm.title = "Summary"
    wb.remove(tpl)
    wb.move_sheet("Summary", offset=-wb.index(sm))
    wb.active = 0
    wb.save(a.out)
    # Đối soát khối lượng
    for sh in data["sheets"]:
        nz = len(by_name[sh["output_sheet"]]["zones"])
        sums = [round(sum((x["qty"][i] or 0) for x in sh["rows"] if x["kind"] != "section" and isinstance(x["qty"][i], (int, float))), 3) for i in range(nz)]
        items = [x for x in sh["rows"] if x["kind"] == "item"]
        miss = [x["src_row"] for x in items if not x.get("cl")]
        print(f"[{sh['output_sheet']}] tổng KL theo khu (để đối chiếu với file CĐT): {sums}")
        print(f"   hạng mục: {len(items)} | chưa có Chủng loại: {len(miss)} {miss[:20]}")
    print("Đã ghi", a.out)


# ============================================================================
# auto: tự đoán cấu trúc BOQ CĐT -> config.json nháp (Claude chỉ cần xem lại)
# ============================================================================
HDR_KEYS = {
    "stt": r"^(stt|tt|no\.?|item|số tt|s\.tt)$",
    "desc": r"mô tả|hạng mục|nội dung|description|tên (vật tư|thiết bị|công việc|hàng)|diễn giải|work item",
    "spec": r"thông số|quy cách|specification|đặc tính|spec",
    "code": r"mã hiệu|model|mã số|ký hiệu|\bcode\b|part no",
    "brand": r"nhãn hiệu|hãng|brand|nhà sản xuất|manufacturer",
    "origin": r"xuất xứ|origin|nước sản xuất|country",
    "unit": r"^(đvt|đơn vị( tính)?|unit|đv)$",
    "qty": r"khối lượng|\bkl\b|quantity|q'?ty|số lượng|^sl$",
    "price": r"đơn giá|thành tiền|unit rate|rate|amount|giá|price|total amount",
}
_HK = {k: re.compile(detone(v)) for k, v in HDR_KEYS.items()}
SKIP_SHEET = re.compile(detone(r"bìa|cover|chỉ dẫn|chi dan|hướng dẫn|tổng hợp|summary|thuyết minh|ghi chú|note|đơn giá|dulieu|data"))


def _txt(ws, r, c):
    v = ws.cell(r, c).value
    return detone(" ".join(str(v).split())) if v not in (None, "") else ""


def _merged_text(ws, r, c, merges):
    for mr in merges:
        if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
            return _txt(ws, mr.min_row, mr.min_col)
    return _txt(ws, r, c)


def auto_sheet(ws):
    merges = list(ws.merged_cells.ranges)
    maxc = min(ws.max_column, 40)
    best = (0, None)
    for r in range(1, min(ws.max_row, 40) + 1):
        hits = set()
        for c in range(1, maxc + 1):
            t = _merged_text(ws, r, c, merges)
            for k, rx in _HK.items():
                if t and rx.search(t):
                    hits.add(k)
        score = len(hits & {"stt", "desc", "unit", "qty"}) * 2 + len(hits)
        if score > best[0]:
            best = (score, r)
    if best[0] < 5:
        return None
    h1 = best[1]
    # tiêu đề 2 tầng: dòng dưới có chữ ở cột thuộc nhóm "khối lượng" hoặc ô merge dọc 2 dòng
    h2 = h1
    nxt = [(_txt(ws, h1 + 1, c), c) for c in range(1, maxc + 1)]
    vertical = any(mr.min_row == h1 and mr.max_row >= h1 + 1 for mr in merges)
    if vertical or (sum(1 for t, _ in nxt if t and not re.match(r"^[\d\s.,-]+$", t)) >= 2 and
                    not any(isinstance(ws.cell(h1 + 1, c).value, (int, float)) for c in range(1, maxc + 1))):
        h2 = h1 + 1
    # dòng đánh số cột (1, 2, 3 …) ngay dưới tiêu đề -> bỏ qua
    data_start = h2 + 1
    nums = [ws.cell(data_start, c).value for c in range(1, maxc + 1)]
    if sum(1 for v in nums if isinstance(v, (int, float))) >= 4 and \
            all(isinstance(v, (int, float)) or v in (None, "") for v in nums):
        data_start += 1
    cols, zones, total = {}, [], None
    for c in range(1, maxc + 1):
        top = _merged_text(ws, h1, c, merges)
        low = _merged_text(ws, h2, c, merges) if h2 != h1 else top
        both = (top + " | " + low).strip()
        L_ = L(c)
        if _HK["price"].search(both) and not _HK["qty"].search(top if top != low else both):
            continue
        for k in ("stt", "desc", "spec", "code", "brand", "origin", "unit"):
            if k not in cols and (_HK[k].search(low) or _HK[k].search(top)):
                cols[k] = L_
                break
        else:
            if _HK["qty"].search(top) or _HK["qty"].search(low):
                if re.search(detone(r"tổng|total|cộng|sum"), low if h2 != h1 else both):
                    total = L_
                else:
                    zones.append({"name": " ".join(str(ws.cell(h2, c).value or ws.cell(h1, c).value or f"KL {L_}").split()), "col": L_})
            elif h2 != h1 and zones and not low == "" and not re.search(detone(r"tổng|total"), low) and \
                    any(mr.min_row == h1 and mr.min_col <= c <= mr.max_col and _HK["qty"].search(_txt(ws, h1, mr.min_col)) for mr in merges):
                zones.append({"name": " ".join(str(ws.cell(h2, c).value).split()), "col": L_})
            elif h2 != h1 and re.search(detone(r"tổng"), low) and any(mr.min_row == h1 and mr.min_col <= c <= mr.max_col and _HK["qty"].search(_txt(ws, h1, mr.min_col)) for mr in merges):
                total = L_
    if not zones and "unit" in cols:
        # tiêu đề chỉ ghi tên khu (Basement, Block A…): lấy các cột chủ yếu là số nằm bên phải cột đơn vị
        for c in range(cidx(cols["unit"]) + 1, maxc + 1):
            head = _merged_text(ws, h2, c, merges) or _merged_text(ws, h1, c, merges)
            if not head or _HK["price"].search(head):
                if zones or total:
                    break
                continue
            vals = [ws.cell(rr, c).value for rr in range(data_start, min(data_start + 300, ws.max_row + 1))]
            filled = [v for v in vals if v not in (None, "")]
            if not filled or sum(isinstance(v, (int, float)) for v in filled) / len(filled) < 0.7:
                break
            if L(c) == total or re.search(detone(r"tổng|total|cộng|sum"), head):
                total = L(c)
                break
            else:
                zones.append({"name": " ".join(str(ws.cell(h2, c).value or ws.cell(h1, c).value).split()), "col": L(c)})
    if not zones and total:  # chỉ có 1 cột KL tổng
        zones, total = [{"name": "Khối lượng", "col": total}], None
    if "desc" not in cols or not zones:
        return None
    # dòng cuối dữ liệu: trước dòng "Tổng cộng" cuối bảng
    dc = cidx(cols["desc"])
    end = ws.max_row
    while end > data_start and all(ws.cell(end, c).value in (None, "") for c in range(1, maxc + 1)):
        end -= 1
    tot_rx = re.compile(detone(r"^(tổng cộng|tổng giá trị|tổng hợp|cộng\s*(:|$)|total\b|grand total)"))
    for r in range(end, max(data_start, end - 40), -1):   # dòng "Tổng cộng" gần cuối nhất
        if tot_rx.match(_txt(ws, r, dc)) or any(tot_rx.match(_txt(ws, r, c)) for c in range(1, dc)):
            end = r - 1
            while end > data_start and all(ws.cell(end, c).value in (None, "") for c in range(1, maxc + 1)):
                end -= 1
            break
    return {"header_rows": [h1, h2], "data_start_row": data_start, "data_end_row": end,
            "columns": {k: cols.get(k) for k in ("stt", "desc", "spec", "code", "brand", "origin", "unit")},
            "zones": zones, "total_col": total}


def cmd_auto(a):
    wb = openpyxl.load_workbook(a.file, data_only=True)
    project = ""
    sheets = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        for r in range(1, min(ws.max_row, 12) + 1):
            for c in range(1, min(ws.max_column, 8) + 1):
                v = ws.cell(r, c).value
                if not project and isinstance(v, str) and re.search(detone(r"dự án|project|công trình"), detone(v)):
                    project = re.sub(r"^\s*(dự án|project|công trình)\s*[:：-]?\s*", "", v.strip(), flags=re.I)
        if SKIP_SHEET.search(detone(ws.title)):
            print(f"- bỏ qua sheet '{ws.title}' (tên giống bìa / chỉ dẫn / tổng hợp)")
            continue
        g = auto_sheet(ws)
        if not g:
            print(f"- bỏ qua sheet '{ws.title}' (không thấy bảng khối lượng)")
            continue
        n = g["data_end_row"] - g["data_start_row"] + 1
        if n < 5:
            print(f"- bỏ qua sheet '{ws.title}' (chỉ {n} dòng)")
            continue
        g.update({"source_sheet": ws.title, "output_sheet": ws.title, "summary_name": ws.title.upper()[:30],
                  "title": ws.title})
        sheets.append(g)
        print(f"+ sheet '{ws.title}': tiêu đề dòng {g['header_rows']}, dữ liệu {g['data_start_row']}–{g['data_end_row']}, "
              f"cột {g['columns']}, khu {[(z['name'], z['col']) for z in g['zones']]}, tổng KL {g['total_col']}")
    cfg = {"input": str(Path(a.file).resolve()), "mode": "inplace", "project": project, "hang_muc": "Cung cấp & lắp đặt hệ thống cơ điện",
           "so_bao_gia": "", "ngay": "",
           "coefficients": {"HSKL": 1, "HSVL": 1.25, "HSTB": 1.1, "HSNC": 1.15, "HS_chung": 1, "lam_tron": -2, "hao_hut": 1.02},
           "sheets": [{k: g[k] for k in ("source_sheet", "output_sheet", "summary_name", "title", "header_rows", "data_start_row",
                                         "data_end_row", "total_col", "columns", "zones")} for g in sheets]}
    Path(a.out).write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")
    print("Đã ghi", a.out, "-> xem lại rồi sửa nếu cần")


# ============================================================================
# check: gộp mọi bước kiểm tra sau build
# ============================================================================
def _col_sig(c):
    return None if c is None else (c.type, c.rgb if c.type == "rgb" else c.theme if c.type == "theme" else c.indexed,
                                   round(c.tint or 0, 4))


def _cell_sig(x):
    f, fl, b, al = x.font, x.fill, x.border, x.alignment
    return (repr(x.value), f.name, f.sz, f.b, f.i, f.u, f.strike, _col_sig(f.color),
            getattr(fl, "fill_type", None), _col_sig(getattr(fl, "fgColor", None)) if getattr(fl, "fill_type", None) else None,
            tuple((sd.style, _col_sig(sd.color)) for sd in (b.left, b.right, b.top, b.bottom)),
            al.horizontal, al.vertical, al.wrap_text, al.indent, al.text_rotation, x.number_format)


_REF_RX = re.compile(r"(?<![A-Za-z0-9_.\[])(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?")


def find_cycles(wb, skip_sheets=("2 Thu vien gia", "DS_MaHieu")):
    """Tìm vòng tham chiếu (Tarjan lặp, không đệ quy) trên mọi công thức của workbook (kể cả chéo sheet).
    Trả về danh sách vòng (mỗi vòng là list 'Sheet!A1'); gồm cả ô tự trỏ chính nó."""
    import bisect
    F = {}                                   # (sheet,row,col) -> công thức
    colrows = {}                             # (sheet,col) -> list dòng có công thức (đã sort)
    for ws in wb.worksheets:
        if ws.title in skip_sheets:
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    F[(ws.title, c.row, c.column)] = v
                    colrows.setdefault((ws.title, c.column), []).append(c.row)
    for lst in colrows.values():
        lst.sort()
    names = {w.title for w in wb.worksheets}
    G = {}
    for (sh, r, c), f in F.items():
        tg = []
        for m in _REF_RX.finditer(re.sub(r'"[^"]*"', '""', f)):
            tsh = (m.group(1) or m.group(2) or sh).replace("''", "'")
            if tsh not in names or tsh in skip_sheets:
                continue
            c1, r1 = cidx(m.group(3)), int(m.group(4))
            c2, r2 = (cidx(m.group(5)), int(m.group(6))) if m.group(5) else (c1, r1)
            for cc in range(min(c1, c2), max(c1, c2) + 1):
                lst = colrows.get((tsh, cc))
                if not lst:
                    continue
                lo, hi = bisect.bisect_left(lst, min(r1, r2)), bisect.bisect_right(lst, max(r1, r2))
                tg.extend((tsh, rr, cc) for rr in lst[lo:hi])
        G[(sh, r, c)] = tg
    idx, low, onst, st, out, n = {}, {}, set(), [], [], 0
    for root in G:
        if root in idx:
            continue
        work = [(root, 0)]
        while work:
            v, i = work.pop()
            if i == 0:
                idx[v] = low[v] = n
                n += 1
                st.append(v)
                onst.add(v)
            ch = G.get(v, [])
            recursed = False
            while i < len(ch):
                w = ch[i]
                i += 1
                if w not in G:
                    continue
                if w not in idx:
                    work.append((v, i))
                    work.append((w, 0))
                    recursed = True
                    break
                if w in onst:
                    low[v] = min(low[v], idx[w])
            if recursed:
                continue
            if low[v] == idx[v]:
                comp = []
                while True:
                    w = st.pop()
                    onst.discard(w)
                    comp.append(w)
                    if w == v:
                        break
                if len(comp) > 1 or v in G.get(v, []):
                    out.append([f"{a}!{L(c)}{r}" for a, r, c in comp])
            if work:
                u = work[-1][0]
                low[u] = min(low[u], low[v])
    return out


def cmd_checkcycle(a):
    wb = openpyxl.load_workbook(a.file, rich_text=True)
    cyc = find_cycles(wb)
    print(f"Vòng tham chiếu: {len(cyc)}")
    for cy in cyc[:50]:
        print("  ", cy)
    sys.exit(1 if cyc else 0)


def cmd_check(a):
    import random
    import shutil
    import subprocess
    import tempfile
    cfg = json.loads(Path(a.config).read_text(encoding="utf-8"))
    ok = True
    # [FIX audit #1/#2] mở mỗi file một lần (A = đầu vào, B = đầu ra, dùng chung cho mọi bước); chỉ copy bản sao khi chạy recalc (full)
    quick = getattr(a, "quick", False)
    need_fid = cfg.get("mode", "inplace") == "inplace" and not getattr(a, "skip_fidelity", False)
    tmpd = cp = None
    if not quick:
        tmpd = tempfile.mkdtemp()
        cp = str(Path(tmpd) / "check.xlsx")
        shutil.copy(a.out, cp)
    A = openpyxl.load_workbook(cfg["input"], rich_text=True) if (quick or need_fid) and need_fid else None
    B = openpyxl.load_workbook(a.out, rich_text=True)
    if quick:
        # Kiểm tra nhanh (không tính lại): quét tĩnh mọi công thức phía TDI + Summary
        wbq = B
        names = set(wbq.sheetnames)
        fn_rx = re.compile(r"([A-Z][A-Z0-9.]*)\(")
        allowed = {"ROUNDDOWN", "ROUND", "IF", "OR", "AND", "SUBTOTAL", "SUM", "IFERROR", "TODAY", "N"}
        for wsi in (A.worksheets if A is not None else openpyxl.load_workbook(cfg["input"], read_only=True).worksheets):
            for row in wsi.iter_rows(values_only=True):
                for v in row:
                    if isinstance(v, str) and v.startswith("="):
                        allowed.update(fn_rx.findall(v))
        ref_rx = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!")
        nf, badf = 0, []
        for wsq in wbq.worksheets:
            for row in wsq.iter_rows():
                for cell in row:
                    v = cell.value
                    if isinstance(v, str) and v.startswith("="):
                        nf += 1
                        if "#REF" in v or "[" in v or any(f not in allowed for f in fn_rx.findall(re.sub(r'"[^"]*"', "", v))):
                            badf.append((wsq.title, cell.coordinate, v))
                            continue
                        for m in ref_rx.finditer(v):
                            nm = (m.group(1) or m.group(2)).replace("''", "'")
                            if nm not in names:
                                badf.append((wsq.title, cell.coordinate, v))
                                break
        print(f"1) Công thức (kiểm tra tĩnh, không tính lại): {nf} công thức | tham chiếu / tên hàm hỏng: {len(badf)}")
        for b_ in badf[:5]:
            print("   ", b_)
        ok &= not badf
    else:
        pass
    # 1) công thức: recalc trên BẢN SAO (không đụng file giao cho người dùng)
    if not getattr(a, "quick", False):
        rows = sum(max(0, (s.get("data_end_row") or 0) - s["data_start_row"]) for s in cfg["sheets"])
        tmo = str(max(180, rows // 50))
        r = subprocess.run([sys.executable, "/mnt/skills/public/xlsx/scripts/recalc.py", cp, tmo], capture_output=True, text=True)
        try:
            res = json.loads(r.stdout[r.stdout.index("{"):])
            print(f"1) Công thức: {res.get('status')} | lỗi: {res.get('total_errors')} | số công thức: {res.get('total_formulas')}")
            ok &= res.get("status") == "success" and not res.get("total_errors")
        except Exception:
            print("1) Không chạy được recalc:", (r.stderr or r.stdout)[-300:])
            ok = False
    # 2) tổng KL theo khu: cộng lại từ file CĐT gốc, so với bản tính trong file kết quả
    src = openpyxl.load_workbook(cfg["input"], data_only=True) if not quick else None   # [FIX audit #2] chế độ nhanh không in/dùng mục 2 -> bỏ
    for sc in (cfg["sheets"] if src is not None else []):
        ws = src[sc["source_sheet"]]
        end = sc.get("data_end_row") or ws.max_row
        sums = []
        for z in sc["zones"]:
            t = 0.0
            for rr in range(sc["data_start_row"], end + 1):
                v = ws.cell(rr, cidx(z["col"])).value
                if isinstance(v, (int, float)):
                    t += v
            sums.append(round(t, 3))
        print(f"2) [{sc['source_sheet']}] tổng KL theo khu trong file CĐT: {sums}")
    # 4) vòng tham chiếu (Excel báo Circular Reference, LibreOffice thì không)
    cyc = find_cycles(B)
    print(f"4) Vòng tham chiếu: {len(cyc)}")
    for cy in cyc[:5]:
        print("   vòng:", cy)
    ok &= not cyc
    # 3) in-place: phía CĐT không đổi (so mẫu ngẫu nhiên giá trị + định dạng). Bỏ qua ở chế độ nhanh (lần chạy lại)
    if cfg.get("mode", "inplace") == "inplace" and not getattr(a, "skip_fidelity", False):
        bad = 0
        tot = 0
        for sc in cfg["sheets"]:
            wa, wb_ = A[sc["source_sheet"]], B[sc["source_sheet"]]
            lastc = last_used_col(wa)
            cells = [(rr, cc) for rr in range(1, wa.max_row + 1) for cc in range(1, lastc + 1)]
            random.seed(1)
            for rr, cc in random.sample(cells, min(3000, len(cells))):
                x, y = wa.cell(rr, cc), wb_.cell(rr, cc)
                tot += 1
                if _cell_sig(x) != _cell_sig(y):
                    bad += 1
                    if bad <= 3:
                        print("   khác ở", x.coordinate, _cell_sig(x), "->", _cell_sig(y))
            if set(map(str, wa.merged_cells.ranges)) - set(map(str, wb_.merged_cells.ranges)):
                bad += 1
                print("   mất merge ở", sc["source_sheet"])
        print(f"3) Phía CĐT giữ nguyên: {tot - bad}/{tot} ô mẫu" + ("" if bad == 0 else "  <-- CÓ THAY ĐỔI"))
        ok &= bad == 0
    if tmpd:
        shutil.rmtree(tmpd, ignore_errors=True)
    print("KẾT LUẬN:", "ĐẠT - giao file được" if ok else "CHƯA ĐẠT - xem các dòng trên")



# ============================================================================
# run: MỘT LỆNH cho cả file (auto -> extract -> build -> check), in gọn kết quả cần soát
# ============================================================================
# ============================================================================
def doubt_groups(rows, limit=20):
    """Chỉ những nhóm thật sự cần người soát: có dòng chưa mã (??) hoặc mã lẫn lộn trong cùng nhóm."""
    items, _ = review_groups(rows, limit=10_000)
    out = [(h, g) for h, g in items if "??" in g["codes"] or len(g["codes"]) > 1]
    return out[:limit], len(out), len(items)


def cmd_run(a):
    import contextlib
    import io
    import time
    T = time.time()
    work = Path(a.workdir)
    work.mkdir(parents=True, exist_ok=True)
    cfg_p, rows_p = work / "config.json", work / "rows.json"
    quiet = io.StringIO()
    # 1) cấu hình: dùng config có sẵn (lần chạy lại) hoặc tự nhận diện
    if a.config:
        cfg_p = Path(a.config)
    else:
        with contextlib.redirect_stdout(quiet):
            cmd_auto(argparse.Namespace(file=a.file, out=str(cfg_p)))
    cfg = json.loads(cfg_p.read_text(encoding="utf-8"))
    if a.mode:
        cfg["mode"] = a.mode
    for k in ("project", "so_bao_gia", "ngay", "hang_muc"):
        if getattr(a, k, None):
            cfg[k] = getattr(a, k)
    cfg_p.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")
    if not cfg["sheets"]:
        print("KHÔNG tìm thấy sheet BOQ nào. Xem bằng: inspect", a.file)
        print(quiet.getvalue())
        return
    t1 = time.time()
    # 2) phân loại (bỏ qua nếu đã có rows.json cùng config và chỉ đang áp sửa)
    if not (a.reuse_rows and rows_p.exists()):
        with contextlib.redirect_stdout(quiet):
            cmd_extract(argparse.Namespace(config=str(cfg_p), out=str(rows_p)))
    data = json.loads(rows_p.read_text(encoding="utf-8"))
    t2 = time.time()
    # 3) dựng
    out = a.out or str(Path("/mnt/user-data/outputs") / (Path(cfg["input"]).stem + "_TDI.xlsx"))
    build_log = io.StringIO()
    with contextlib.redirect_stdout(build_log):
        cmd_build(argparse.Namespace(config=str(cfg_p), rows=str(rows_p), overrides=a.overrides, out=out))
    t3 = time.time()
    # 4) kiểm tra
    chk = io.StringIO()
    if a.check != "none":
        with contextlib.redirect_stdout(chk):
            cmd_check(argparse.Namespace(config=str(cfg_p), out=out, quick=a.check == "quick",
                                         skip_fidelity=a.reuse_rows and a.check == "quick"))
    t4 = time.time()
    # ---- báo cáo ngắn
    ov = json.loads(Path(a.overrides).read_text(encoding="utf-8")) if a.overrides else {}
    if ov:
        apply_overrides(data, ov)
    print(f"FILE: {out}")
    print(f"Dự án: {cfg.get('project') or '(chưa rõ)'} | chế độ: {cfg.get('mode', 'inplace')} | {len(cfg['sheets'])} sheet")
    for line in quiet.getvalue().splitlines():
        if line.startswith("- bỏ qua"):
            print("  " + line)
    tdi_cols = {m.group(1): (m.group(2), m.group(3)) for m in re.finditer(r"\[(.+?)\] phần TDI ở cột (\S+), dòng tổng TDI: (\d+)", build_log.getvalue())}
    for sc, sh in zip(cfg["sheets"], data["sheets"]):
        rows = sh["rows"]
        items = [x for x in rows if x["kind"] == "item"]
        none = [x["src_row"] for x in items if not x.get("cl")]
        dg, ndg, ng = doubt_groups(rows, a.max_groups)
        loc = tdi_cols.get(sc["source_sheet"])
        print(f"\n[{sc['source_sheet']}] dòng {sc['data_start_row']}–{sc.get('data_end_row')} | tiêu đề {sc.get('header_rows')} | "
              f"khu {[z['col'] for z in sc['zones']]} tổng {sc.get('total_col')} | {len(items)} hạng mục | chưa mã {len(none)}"
              + (f" | TDI {loc[0]}, tổng TDI dòng {loc[1]}" if loc else ""))
        if ndg:
            print(f"  Cần soát {ndg}/{ng} nhóm (có ?? hoặc mã lẫn lộn):")
            for h, g in dg:
                codes = ", ".join(f"{k}×{v}" for k, v in sorted(g["codes"].items(), key=lambda kv: -kv[1]))
                print(f"   {g['n']:>4} | {codes:<20} | ctx:{h}  vd: {' / '.join(g['samples'])}")
        else:
            print(f"  Không có nhóm nào cần soát ({ng} nhóm đều thuần một mã).")
    res = chk.getvalue()
    for line in res.splitlines():
        if line.startswith(("1)", "3)", "4)", "KẾT LUẬN", "   khác", "   vòng")):
            print(line)
    if a.check == "none":
        print("(bỏ qua kiểm tra)")
    print(f"\nThời gian: nhận diện {t1 - T:.1f}s | phân loại {t2 - t1:.1f}s | dựng {t3 - t2:.1f}s | kiểm tra {t4 - t3:.1f}s | tổng {t4 - T:.1f}s")
    print(f"Để áp sửa Chủng loại: ghi overrides.json rồi chạy lại: run --config {cfg_p} --reuse-rows --overrides <overrides.json> --out {out}")



def main():
    p = argparse.ArgumentParser()
    sp = p.add_subparsers(dest="cmd", required=True)
    s = sp.add_parser("inspect"); s.add_argument("file"); s.add_argument("--rows", type=int, default=25)
    s = sp.add_parser("run"); s.add_argument("file", nargs="?"); s.add_argument("--config"); s.add_argument("--overrides")
    s.add_argument("--out"); s.add_argument("--workdir", default="/home/claude/boq"); s.add_argument("--mode", choices=["inplace", "rebuild"])
    s.add_argument("--project"); s.add_argument("--so_bao_gia"); s.add_argument("--ngay"); s.add_argument("--hang_muc")
    s.add_argument("--max-groups", type=int, default=25)
    s.add_argument("--reuse-rows", action="store_true")
    s.add_argument("--check", choices=["quick", "full", "none"], default="quick",
                   help="quick: soát tĩnh công thức + so định dạng CĐT (mặc định); full: thêm tính lại bằng LibreOffice")
    s = sp.add_parser("auto"); s.add_argument("file"); s.add_argument("--out", default="config.json")
    s = sp.add_parser("check"); s.add_argument("--config", required=True); s.add_argument("--out", required=True)
    s.add_argument("--quick", action="store_true"); s.add_argument("--skip-fidelity", dest="skip_fidelity", action="store_true")
    s = sp.add_parser("extract"); s.add_argument("--config", required=True); s.add_argument("--out", required=True)
    s = sp.add_parser("checkcycle"); s.add_argument("file")
    s = sp.add_parser("review"); s.add_argument("--rows", required=True)
    s.add_argument("--only", choices=["all", "none", "context", "items", "group"], default="all")
    s = sp.add_parser("build"); s.add_argument("--config", required=True); s.add_argument("--rows", required=True)
    s.add_argument("--overrides"); s.add_argument("--out", required=True)
    a = p.parse_args()
    {"inspect": cmd_inspect, "run": cmd_run, "auto": cmd_auto, "check": cmd_check, "checkcycle": cmd_checkcycle, "extract": cmd_extract, "review": cmd_review, "build": cmd_build}[a.cmd](a)


if __name__ == "__main__":
    main()