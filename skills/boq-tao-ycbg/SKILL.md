---
name: "boq-tao-ycbg"
description: "Gen file Đề nghị báo giá (YCBG) theo mẫu TDI, mỗi hệ một file, từ file tổng hợp đầu mục chưa có trong thư viện (đầu ra Bước 6d của boq-gan-ma-gia-thu-vien, đã được boq-gia-tham-khao-web bổ sung giá tham khảo nếu đã chạy)."
---

# Tạo YCBG theo hệ từ file tổng hợp đầu mục chưa có giá

Chạy **sau `boq-gan-ma-gia-thu-vien`** (đã xuất file tổng hợp Bước 6d) và **sau `boq-gia-tham-khao-web`** (ghi giá tham khảo web vào file tổng hợp; chưa chạy thì YCBG vẫn tạo được, chỉ thiếu giá tham khảo nội bộ ở sheet `Chi tiết dòng`), **trước Bước 5** (người rà soát và gửi NCC), và trước `boq-lay-gia-ncc` (lấy báo giá NCC về BOQ).

Đầu vào chính: file `<tên BOQ>_Dau_muc_chua_co_trong_thu_vien_gia_tham_khao.xlsx` trong thư mục `Claude outputs` cạnh file BOQ TDI. Không đọc lại BOQ để tự lọc dòng: file tổng hợp đã áp quy tắc loại dòng của skill gán giá. YCBG lập cho **mọi đầu mục chưa có giá trong thư viện** (mọi dòng sheet `Chưa có trong TV`), kể cả đầu mục đã có giá tham khảo web, vì giá web chưa phải giá NCC.

Đầu ra mỗi lần chạy:
- `YCBG_<mã dự án>_<Hệ>.xlsx`: mỗi hệ một file, một sheet, copy từ mẫu YCBG trắng của TDI (giữ logo, font, khổ in). Các nhóm vật tư trong hệ ngăn nhau bằng dòng tiêu đề nền vàng đánh số I, II, III.
- `YCBG_<mã dự án>_TongHop.xlsx`: sheet `Theo dõi YCBG` (ưu tiên, NCC, ngày gửi, hạn phản hồi, trạng thái), `Chi tiết dòng` (Ref BOQ từng dòng, cột để điền đơn giá NCC ở Bước 6), `Không đưa vào YCBG`, `Cảnh báo`.

## Quy tắc bắt buộc

1. **Không gửi gì ra ngoài.** Chỉ tạo file nháp. Gửi NCC là việc của người ở Bước 5. Ô `Ngày gửi yêu cầu` để trống, người gửi tự điền.
2. **Không sửa file BOQ và file tổng hợp.** Chỉ đọc.
3. **Diễn giải chép nguyên văn đầu mục BOQ** (kể cả phần tiếng Trung), chỉ bỏ khoảng trắng thừa. `boq-lay-gia-ncc` khớp giá theo đầu mục nguyên văn.
4. **Không đưa dữ liệu nội bộ vào file gửi NCC**: giá tham khảo web, hãng/model tham chiếu, link nguồn, lý do chưa khớp. Các thông tin này chỉ nằm trong `YCBG_..._TongHop.xlsx`.
5. **Không đoán quy cách, mã hiệu, hãng.** Cột Quy cách chỉ có nội dung khi đã có căn cứ (dòng thành phần cụm ghi "Thành phần của: …"). Muốn điền quy cách từ spec thì phải có trang hoặc mục nguồn, ghi vào Ghi chú.
6. **Mã dự án** của file tổng hợp phải khớp mã người dùng đang làm và mã trong cấu hình. Không khớp thì dừng và hỏi.
7. **Mỗi dòng YCBG có Ref BOQ** dạng `sheet!dòng` (cột ẩn sau bảng, và cột `Ref BOQ` trong tổng hợp). Dòng gộp từ nhiều vị trí ghi đủ các ref.
8. **Không ghi đè file cũ**: nếu đã có file cùng tên, script tạo `_dot2`, `_dot3`.
9. **Dòng cần chiết tính** (lô, gói, đơn vị đặc biệt) không đưa vào YCBG, nằm ở sheet `Không đưa vào YCBG`.

## Công cụ

- Script `ycbg_tools.py` và cấu hình mẫu `YCBG_CONFIG_mau.json` nằm trong Project (`claude/boq/`), bản dùng trên máy ở `_Claude_tools` cạnh `boq_xml_tools.py`. Chưa có thì lấy bằng `project_read` rồi ghi vào workspace.
- Cần Python 3 và openpyxl. Chạy trong workspace cloud: stage file tổng hợp và mẫu YCBG lên, chạy, rồi commit kết quả về thư mục đích. Nếu máy người dùng có openpyxl thì chạy bằng `device_bash` tại chỗ.
- Mẫu YCBG: file `YCBG*.xlsx` còn trắng (tiêu đề bảng ở cột A là `STT`, có `Chú ý`, `THÔNG TIN NGƯỜI NHẬN HÀNG`, `THÔNG TIN XUẤT HÓA ĐƠN`) trong thư mục `1.1 YCBG` của dự án. Script tự dò vị trí các khối, không cần cố định dòng.

| Lệnh | Việc |
|---|---|
| `python3 ycbg_tools.py read SRC.xlsx ITEMS.json --ma <mã>` | Đọc file tổng hợp thành ITEMS.json (sheet `Chưa có trong TV`, `Có mã – TV thiếu giá VL`, `Thành phần bóc tách`, `Cần chiết tính`). Tên sheet BOQ có dấu phẩy được đọc đúng. |
| `python3 ycbg_tools.py plan ITEMS.json CONFIG.json` | In bảng hệ × nhóm, ví dụ đầu mục, và các dòng chưa xếp được hệ, để duyệt ở điểm dừng 1. |
| `python3 ycbg_tools.py build MAU.xlsx ITEMS.json CONFIG.json OUTDIR [--cot-gia]` | Gen file theo hệ và file tổng hợp. `--cot-gia` thêm cột Đơn giá VL, Đơn giá NC, Thành tiền (công thức), Thời gian giao hàng cho NCC điền, trang in chuyển sang khổ ngang. Script dừng nếu cấu hình còn dấu `<…>` hoặc thiếu địa chỉ giao hàng, người nhận. |
| `python3 ycbg_tools.py check OUTDIR ITEMS.json CONFIG.json` | Kiểm: mọi Ref BOQ có trong tổng hợp, tổng KL vào bằng ra, mỗi file còn logo, không dòng nào thiếu ĐVT (trừ dòng nguồn đã thiếu), phần Chú ý không đè lên khối người nhận, không còn dấu `<…>`. Phải ra `KẾT QUẢ: ĐẠT`. |

## Các bước

### Bước 0. Xác định đầu vào
- File tổng hợp Bước 6d của đúng dự án (mã dự án trong dòng 2 sheet `Chưa có trong TV`). Không thấy file này thì dừng và báo người dùng chạy `boq-gan-ma-gia-thu-vien` trước.
- Mở sheet `Chưa có trong TV`: nếu mọi dòng đều trống cột giá tham khảo và `Tình trạng giá` = `Chưa tra giá web` thì `boq-gia-tham-khao-web` chưa chạy. Báo người dùng và hỏi chạy skill đó trước hay tạo YCBG ngay (YCBG không phụ thuộc giá tham khảo vì giá web không đưa vào file gửi NCC, Quy tắc 4).
- Mẫu YCBG trắng, `YCBG_CONFIG.json` (tìm ở `1.1 YCBG` của dự án, rồi `_Claude_tools`). Chưa có cấu hình thì tạo từ Phụ lục, rồi hỏi người dùng các trường còn trống: **tên người nhận hàng, địa chỉ giao hàng, hạn báo giá**. Không in dấu `<…>` ra file.
- Thư mục đích mặc định: `<dự án>\2. HSDT\1.1 YCBG\YCBG_<yyyy-mm-dd>\`.

### Bước 1. Đọc
Chạy `read`. Báo số dòng đọc theo từng sheet, số dòng không đưa vào YCBG, và các cảnh báo (thiếu KL, thiếu ĐVT, có sheet `TĐ – linh kiện`).

### Bước 2. Xếp hệ và **điểm dừng 1**
- Hệ lấy theo `quy_tac_he` trong cấu hình, xét từ trên xuống, quy tắc khớp đầu tiên thắng. Quy tắc có `nhom` (nhóm trong file tổng hợp) và/hoặc `mo_ta` (từ khóa trong đầu mục). Quy tắc theo mô tả đặt trước quy tắc theo nhóm vì cột Nhóm không phải lúc nào cũng đúng hệ (ví dụ đèn báo FIRE nằm trong nhóm PCCC, louver nằm trong nhóm PCCC nhưng thuộc HVAC).
- Chạy `plan` và đưa người dùng: số dòng và nhóm mỗi hệ, vài đầu mục tiêu biểu, các dòng rơi vào hệ `Khác`, số dòng sau khi gộp trùng. Muốn chỉ định hệ cho một dòng thì điền `he` cho dòng đó trong ITEMS.json, hoặc thêm quy tắc.
- Hỏi thêm: có thêm cột giá (`--cot-gia`) không, có người nhận riêng cho từng hệ không.
- **Chờ người dùng xác nhận rồi mới build.**

### Bước 3. Tạo file
Chạy `build`, rồi `check`. Hệ có từ hai nhóm trở lên thì có dòng tiêu đề nền vàng cho từng nhóm. Số YCBG: `<mã>-YCBG-<HỆ>-<NN>`, NN tăng khi chạy lại.

### Bước 4. Kiểm tra
1. `check` ra `ĐẠT`.
2. Chuyển 1–2 file sang PDF (`soffice --headless --convert-to pdf`), xuất ảnh trang (`pdftoppm -png -r 60`) và xem: logo, tiêu đề bảng lặp lại ở trang sau, phần Chú ý và khối người nhận không đè nhau, chữ dài xuống dòng.
3. Mở `Theo dõi YCBG` và `Cảnh báo` trong file tổng hợp.

### Bước 5. Giao kết quả
Ghi file vào thư mục đích trên máy người dùng và báo: danh sách file và số dòng từng hệ, tổng số đầu mục, số dòng đã gộp, các dòng thiếu KL hoặc ĐVT, các dòng ở hệ `Khác`. Nêu rõ các việc người dùng làm ở Bước 5:
- Duyệt các dòng ưu tiên cao (cột `Ưu tiên` trong `Theo dõi YCBG`: hãng khóa dự án, thiết bị chính, hàng đặt dài ngày).
- Điền địa chỉ giao hàng còn thiếu nếu có, rồi tự gửi NCC. Điền NCC, ngày gửi, hạn phản hồi vào `Theo dõi YCBG`.
- Khi NCC trả file, đặt vào thư mục `BG NCC` rồi chạy `boq-lay-gia-ncc`.

## Giới hạn đã biết
- Sheet `TĐ – linh kiện` (cấu hình từng tủ) **chưa** được xuất thành phụ lục. Script chỉ cảnh báo. Linh kiện tủ gửi NCC theo bản vẽ hoặc sơ đồ nguyên lý riêng.
- Cột `Ưu tiên` xếp theo từ khóa (`uu_tien`, `hang_khoa`), chỉ để gợi ý thứ tự rà soát.
- Cột Quy cách, Mã hiệu, Hãng để trống, NCC tự điền.

## Phụ lục. Cấu hình mẫu `YCBG_CONFIG.json`
Các trường có dấu `<…>` bắt buộc thay trước khi chạy (script từ chối chạy nếu còn).
```json
{
  "du_an": "<Tên dự án>", "ma_du_an": "<TDIxxxxx>",
  "ngay_gui": "", "han_bao_gia": "Sớm nhất có thể",
  "dia_chi_giao_hang": "<Địa chỉ giao hàng>",
  "nguoi_nhan": {"_mac_dinh": "Lê Minh Tuấn - 0942.468.130 - Email: tuanlm@tdimep.vn"},
  "hoa_don": {"cong_ty": "CÔNG TY CỔ PHẦN TẬP ĐOÀN TDI",
              "dia_chi": "A16-NV6, ô số 10, KĐT mới đường Lê Trọng Tấn, An Khánh, Hoài Đức, HN",
              "mst": "0307717044",
              "dien_thoai": "(024) 7308.6632 - info@tdigroup.vn - www.tdigroup.vn"},
  "thu_tu_he": ["Điện", "Điện nhẹ", "Cấp thoát nước", "HVAC", "PCCC", "Xây dựng"],
  "he_mac_dinh": "Khác",
  "nhom_mac_dinh": "Hạng mục khác",
  "doi_ten_nhom": {"Khác (ME)": "Vật tư, thiết bị khác"},
  "ten_he": {"Điện": "hệ Điện", "HVAC": "hệ Thông gió - Điều hòa (HVAC)", "PCCC": "hệ PCCC", "Xây dựng": "hạng mục Xây dựng"},
  "vv": "Đề nghị báo giá {he} - Dự án {du_an}, quy cách như bảng dưới đây.",
  "quy_tac_he": [
    {"he": "Xây dựng", "nhom": "^xd –|^xây dựng"},
    {"he": "Điện", "mo_ta": "tủ điện|ring main|\\brmu\\b|máy biến áp|trạm biến áp|đầu cáp trung thế"},
    {"he": "PCCC", "mo_ta": "chữa cháy|báo cháy|đầu phun|sprinkler|trụ tiếp nước|họng|\\bfire\\b|pccc|van xả khí"},
    {"he": "HVAC", "mo_ta": "quạt|louver|cửa gió|miệng gió|van gió|cổ bạt|chụp thời tiết|ống gió|\\bprd\\b|\\bvcd|\\bnrd\\b|bảo ôn|điều hòa"},
    {"he": "HVAC", "mo_ta": "van xả áp|\\bmfd\\b|\\bmd\\b|chặn lửa|ống gió"},
    {"he": "Điện nhẹ", "mo_ta": "cáp quang|cat6|camera|mạng lan|\\bloa\\b|access"},
    {"he": "Điện", "mo_ta": "\\bmcb\\b|\\bmccb\\b|cầu dao|recloser|\\blbs\\b|sứ|cáp ngầm|kios|máy phát|giằng neo|hộp kỹ thuật đấu dây|điện trở sấy|bình dầu"},
    {"he": "Cấp thoát nước", "mo_ta": "công tắc áp suất|công tắc dòng chảy|phao|bình tích áp|cút hướng dòng|composite|ventcap|ống thông hơi"},
    {"he": "Điện", "nhom": "tủ điện|trung thế|chống sét|dụng cụ an toàn điện|chiếu sáng|thang / máng cáp"},
    {"he": "Điện nhẹ", "nhom": "điện nhẹ"},
    {"he": "PCCC", "nhom": "^pccc"},
    {"he": "HVAC", "nhom": "^hvac"},
    {"he": "Cấp thoát nước", "nhom": "bơm|^van /"}
  ],
  "uu_tien": ["tủ điện", "trung thế", "máy biến áp", "\\brmu\\b", "bơm", "quạt", "điều hòa", "busway", "máy phát"],
  "hang_khoa": ["Grundfos", "Daikin", "York", "Schneider", "Trane", "Carrier", "Ebara", "Wilo", "Mitsubishi", "Hitachi", "ABB", "Siemens", "Hochiki", "Notifier"],
  "chu_y": ["- Đơn giá chào bằng VNĐ, ghi rõ đã hay chưa gồm VAT; giá giao đến chân công trình.",
            "- Ghi rõ hãng sản xuất, mã hiệu (model), xuất xứ; gửi kèm catalogue; hàng có CO, CQ.",
            "- Ghi rõ thời gian giao hàng, thời gian bảo hành và hiệu lực của báo giá.",
            "- Nếu chào hàng tương đương hoặc khác quy cách, ghi rõ tại cột Ghi chú.",
            "- Vui lòng giữ nguyên STT và thứ tự dòng của bảng này khi gửi lại báo giá."]
}
```