---
name: "boq-lay-gia-ncc"
description: "Quét thư mục BG NCC (Excel, PDF, Word), khớp với BOQ, xuất Bảng rà soát để người duyệt trước, rồi xuất ASSIGN.json cho bước fill giá vào BOQ."
---

# Rà soát khớp BOQ ↔ báo giá NCC (trước khi fill giá)

Dùng khi người dùng muốn: thống kê đầu mục BOQ nào đã có báo giá trong BG NCC, kiểm tra/duyệt kết quả khớp, rồi mới fill giá. Đứng GIỮA bước nhận báo giá NCC và skill `boq-ghi-gia-ncc` (skill đó ghi giá đã duyệt vào BOQ qua `apply`/`audit`, thay cả ô giá tham khảo web bằng giá NCC). Skill này KHÔNG sửa BOQ, không sửa file OneDrive, không gửi email.

## Nguồn báo giá đọc được
- **Excel** (xlsx/xlsm/xls; .xls hỏng tự chuyển bằng soffice): tự nhận dòng tiêu đề và cột mô tả/ĐVT/KL/đơn giá.
- **PDF có chữ** (pdftotext -layout): nhận dòng có hai số bất kỳ nhân ra số thứ ba (SL × ĐG = TT, sai số 0,3%; TT có thể đã gồm VAT 5/8/10%; TT có thể nằm ở 2 dòng kế; SL và ĐG theo thứ tự nào cũng được; không nhận SL trùng STT đầu dòng; bỏ dòng Tổng/VAT/Thuế). Mô tả nhiều dòng được gom từ các dòng quanh dòng giá (có thể lẫn chữ của dòng bên cạnh, người duyệt đối chiếu file gốc). Bảng giá niêm yết không có SL/thành tiền (vd VNF) đọc theo STT|tên|ĐVT|đơn giá, SL để trống. Số có định dạng sai (vd 2.450.00) bị bỏ qua, không đoán.
- **Word** (docx; .doc chuyển bằng soffice): đọc bảng (nhận tiêu đề như Excel); không có bảng thì đọc từng dòng văn bản theo quy tắc SL × ĐG = TT. Chưa có file Word thật để kiểm chứng ngoài file mẫu: lần đầu gặp Word thật phải xem kỹ kết quả.
- **Chưa đọc**: PDF scan, ảnh (jpg/png), email (.msg), file nén (rar/zip) — liệt kê ở sheet 4 để xem tay hoặc đọc riêng.
- Tự bỏ: catalogue, datasheet, bảng chọn thiết bị (tên chứa catalog, datasheet, data sheet, technical data, Ref. n, catalouge, profile, elta fans, kruger); đổi bằng `--skip-regex`.

## Bước 0 – Lấy công cụ `ncc_rasoat.py` (làm TRƯỚC mọi bước khác)
Script tự chứa `claude/boq/ncc_rasoat.py` nằm ở Project "Tối ưu hóa quy trình Phòng đấu thầu". Phiên làm việc có thể gắn Project khác (vd. project của từng gói thầu) nên KHÔNG mặc định có sẵn. Tìm theo thứ tự, dừng ở nơi đầu tiên thấy:
1. `project_read` đường dẫn `claude/boq/ncc_rasoat.py` (nếu phiên gắn đúng Project công cụ); thử thêm `project_search` "ncc_rasoat".
2. Thư mục dự án đã kết nối trên máy người dùng (tìm file tên `ncc_rasoat*.py`, vd thư mục `PDT_BOQ` hoặc thư mục công cụ) → `device_stage_files`.
3. Kho OneDrive/SharePoint của Phòng đấu thầu (tìm `ncc_rasoat`) qua connector Microsoft 365.
4. Nếu cả ba nơi đều không có: DỪNG, báo người dùng, chọn một: (a) người dùng đính kèm file script; (b) chạy lại skill trong phiên gắn Project "Tối ưu hóa quy trình Phòng đấu thầu"; (c) chỉ khi người dùng đồng ý rõ mới tự viết lại công cụ từ mô tả skill (nêu rõ kết quả có thể khác bản gốc và phải đối chiếu mẫu). Không tự viết lại âm thầm.
Sau khi có script, chép vào thư mục làm việc và ghi lại để không phải tìm lại; gợi ý người dùng lưu một bản vào Project của gói thầu đang làm.

## Công cụ
Cần python3, openpyxl, xlrd, python-docx, rapidfuzz, numpy (`pip install --break-system-packages`), poppler (`pdftotext`), soffice nếu có .xls/.doc. Chạy bằng `python3 -I`.

| Lệnh | Việc |
|---|---|
| `extract --root <BG NCC> --work <tmp> [--skip-regex ...]` | Quét Excel + PDF + Word, ghi dòng báo giá có giá (mỗi dòng có `src` = xlsx/pdf/docx), mã hiệu, danh sách file chưa đọc |
| `match --boq <BOQ.xlsx> --work <tmp> [--sheets ...]` | Khớp từng dòng BOQ: mô tả + thông số số + chữ ký kích thước (1Cx10 ≠ 3Cx10, 800x200 ≠ 200x200) + vật liệu cách điện (XLPE/MICA/LSZH) + mã hiệu tủ/quạt (kể cả dải DB-A-L19~21); giữ tối đa 10 báo giá/dòng |
| `build --work <tmp> --xlsx <Ra-soat.xlsx> --boq-name <tên BOQ>` | Dựng Bảng rà soát |
| `assign --review <Ra-soat.xlsx đã duyệt> --out ASSIGN.json [--auto-duyet] [--component VL\|NC]` | Đọc quyết định → ASSIGN.json + `_SoCai.xlsx` |

GIẢ ĐỊNH của `match`: BOQ có cột A=STT, B=Mô tả, C=ĐVT, G=KL (hoặc AC); sheet BOQ tên dạng `B2_...`. Khác bố cục thì truyền `--sheets` hoặc hỏi người dùng.

## Quy trình
0. Lấy script (Bước 0 ở trên).
1. Hỏi một lần: đúng BOQ nào (thư mục có nhiều phiên bản: vd 251111, 260226hs, 260319hs…; gợi ý bản mới nhất); thư mục BG NCC nào (thường `2. HSDT/1.2. BG NCC`; BOQ nằm ở `2. HSDT`).
2. Stage BOQ và file BG NCC (xlsx/xls, PDF báo giá, Word). Cẩn thận đường dẫn Unicode NFC/NFD. Thư mục BG NCC lớn (hàng trăm file): `device_list_dir` đệ quy có thể vượt giới hạn token, kết quả được lưu ra file, đọc bằng python/JSON. Loại catalogue/datasheet trước khi stage (PDF catalogue nặng, không có giá). Stage theo lô ≤ 50 file; nếu `device_bash` có mặt, ưu tiên chạy `extract` ngay trên máy người dùng để khỏi stage cả thư mục.
3. Chạy `extract` → `match` → `build`.
4. Giao Bảng rà soát và DỪNG chờ duyệt. Báo ngắn: số dòng BOQ; Cao/Trung bình/Thấp/Chưa khớp; số dòng "Nên fill"; số file chưa đọc; giả định.
5. Người dùng duyệt (cột cam), gửi lại. Chạy `assign`; báo số dòng vào ASSIGN và danh sách dòng bị bỏ qua kèm lý do (thiếu NCC/ngày/file, chưa có đơn giá).
6. Chuyển `ASSIGN.json` và sổ cái `*_SoCai.xlsx` cho skill `boq-ghi-gia-ncc` (kiểm tra → `apply --dry` → xác nhận → `apply` → `audit`). Skill này KHÔNG tự ghi giá và không tự chạy `boq-ghi-gia-ncc` khi người dùng chưa đồng ý.

## Các sheet của Bảng rà soát
- `0. Hướng dẫn`: số liệu và hướng dẫn 3 bước.
- `1. Duyệt giá`: 1 dòng = 1 đầu mục BOQ + báo giá khớp nhất; cột máy (Độ tin cậy, Cách khớp, Đề xuất, Cảnh báo, Số báo giá khớp, Số NCC, giá thấp/cao); cột người duyệt (cam): Quyết định (Duyệt/Bỏ qua), Chọn báo giá khác, NCC, Ngày báo giá, Thành phần (VL/NC/VL+NC), VAT, Ghi chú.
- `2. Báo giá thay thế`: các báo giá ứng viên theo ID (hạng, NCC, file, sheet/trang, dòng, mô tả nguyên văn, ĐVT, KL, đơn giá).
- `3. BOQ chưa có báo giá`: nguồn lập YCBG bổ sung (gồm cả hạng mục DC/NSC không cần hỏi giá).
- `4. File NCC chưa đọc`: file mở được nhưng không trích được giá, catalogue bỏ qua, loại file chưa đọc.

## Quy tắc bất biến
- Dòng để trống Quyết định = không fill. `--auto-duyet` chỉ khi người dùng yêu cầu rõ, và chỉ cho dòng "Nên fill".
- Dòng từ PDF/Word do máy trích: luôn cảnh báo "Nguồn PDF/DOCX", KHÔNG bao giờ xếp "Nên fill"; người duyệt đối chiếu file gốc.
- Mỗi giá phải có nguồn đủ: NCC, file, sheet/trang + dòng, ngày báo giá, ID sổ cái. Thiếu ngày/NCC thì KHÔNG đưa vào ASSIGN (ngày chỉ tự suy khi không mơ hồ).
- Đơn giá là giá nguyên văn: VAT/chiết khấu/hết hạn chưa chuẩn hóa; trạng thái luôn "Cần kiểm tra". KL trong báo giá chỉ để tham khảo.
- Dòng ĐVT lô/bộ/gói/HT hoặc ĐVT khác phải chiết tính (`calc`) trước khi fill; không tự quy đổi.
- Khớp theo mô tả/mã hiệu, chưa đối chiếu hãng/model theo spec/Vendor list; người duyệt chịu trách nhiệm chọn NCC và giá chốt.
- Bảng chọn quạt (Nicotra/Elta/Kruger) khớp theo mã nhưng chưa có đơn giá: đánh "mã hiệu – chưa có giá", không đưa vào ASSIGN.
- Nêu rõ BOQ phiên bản/ngày nào; không sửa/di chuyển/xóa file OneDrive; không gửi email.

## Lỗi dễ gặp
- Không tìm thấy `ncc_rasoat.py`: xem Bước 0; không tự bịa kết quả, không tự viết lại công cụ khi người dùng chưa đồng ý.
- Dòng BOQ chỉ là mã tủ/quạt (DB-A-L27, MSB-A1) khớp mô tả sẽ hỏng: script khớp theo mã hiệu (KL trong dải "DB-A-L19~21" là của cả dải).
- Cáp/ống gió sai quy cách: script đã chặn/hạ mức nhưng vẫn cần người duyệt xem lại.
- Một đầu mục có nhiều NCC/đợt cập nhật: xem Số NCC, giá thấp/cao và sheet 2; ưu tiên đợt mới nhất còn hiệu lực.
- PDF không ra dòng nào: thường không phải báo giá (datasheet, giấy kiểm định, công văn) — mở xem trước khi kết luận lỗi. PDF scan/ảnh và bảng giá quá lạ: cần đọc tay hoặc đọc ảnh trang.
- Bảng giá niêm yết (không KL) và PDF có VAT trong thành tiền: giá lấy theo đơn giá nguyên văn, chưa chuẩn hóa VAT.