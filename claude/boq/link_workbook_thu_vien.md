# Link sống thư viện giá bằng workbook link (`tv_link.py`)

_Soạn 22/10/2026 theo yêu cầu phòng (cách thủ công cũ là link từng ô giá). Dùng với skill `covert-boq-cdt-into-boq-tdi`, mục 3b. Script: `claude/boq/tv_link.py`._

## 1. Hành vi

- Ô giá trong bảng `TV` (sheet `2 Thu vien gia`) là **công thức trỏ vào ô giá của file thư viện** (`='[1]Ong HDPE'!$E$9`), đúng như link tay: các cột `Giá VL thư viện (thô)`, `Giá NC thư viện (thô)`, `Giá NC ước tính`.
- `Giá VL gốc` / `Giá NC gốc` (E, F) vẫn là thô × hệ số Loại VL (Quy tắc 24), nên đổi giá thư viện thì E, F và mọi công thức `SUMIFS(TV[…])` đổi theo.
- **Một chiều**: sửa giá trong BOQ không ghi ngược vào file thư viện. Sửa giá ở file nguồn rồi **lưu**, mở BOQ → Excel hỏi cập nhật link → **Update Values** (hoặc Data → Edit Links → Update Values). File thư viện **không cần mở**.
- Cột cuối mới `Kiểm dòng nguồn`: `OK` nếu ô nguồn vẫn đúng **Mã hiệu + Tên công việc** của dòng TV; `LỆCH DÒNG` nếu thư viện bị chèn / xóa dòng làm link trỏ sai ô; `KHÔNG LINK` nếu dòng không nối được (giá giữ tĩnh).

## 2. Thứ tự làm (Claude làm trong skill chuyển template)

1. Nạp `TV` bằng `tv_loader.py load` rồi `check` = 0/0/0 (như thường).
2. `boq_xml_tools.py hsloai FILE OUT` (bản 15/10/2026 chiều trở lên) để có cột O–R; `check` lại.
3. `python3 tv_link.py link FILE OUT --libdir "<thư mục thư viện ME đã stage>" --linkbase "<đường dẫn Windows thật của thư mục thư viện trên máy người dùng>"`.
   - `--linkbase` là nơi file thư viện nằm **trên máy người dùng** (ví dụ `C:\Users\Admin\OneDrive - TDI Group\…\2. Thu vien gia`), không phải đường dẫn tạm của cầu nối. Không biết chắc thì hỏi người dùng.
   - `--libdir` phải là đúng bộ file đã dùng để nạp `TV` (script đối chiếu từng dòng theo Mã + Tên và giá; dòng không khớp được giữ giá tĩnh và được liệt kê).
4. Giao file; dặn người dùng bấm **Enable Content / Update Values** khi mở.

Chạy `check` của `tv_loader.py` **trước** khi nối link (check đọc giá trị, không đọc công thức). Sau khi nối link dùng `tv_link.py verify`.

## 3. Việc người dùng cần biết

- Máy mở BOQ phải thấy được thư mục thư viện ở **đúng đường dẫn đã nối** (OneDrive đồng bộ, ổ đĩa giống nhau). Khác đường dẫn: Data → Edit Links → **Change Source** cho từng file (18 file), hoặc nhờ Claude nối lại với `--linkbase` mới.
- Chưa cập nhật link thì BOQ vẫn hiện giá lần cập nhật cuối (Excel lưu giá trị cache trong file).
- Thư viện bị **chèn / xóa dòng** thì link bám theo vị trí ô nên lệch: cột `Kiểm dòng nguồn` báo `LỆCH DÒNG`. Gặp thì không dùng tiếp: Break Link, nạp lại `TV` và nối lại; hoặc dùng cách Power Query (`power_query_tv_song.md`, đọc theo tên cột nên không lệch dòng).
- `verify` (Claude chạy): `python3 tv_link.py verify FILE --libdir "<thư mục thư viện>"` đếm ô giá còn khớp thư viện hiện tại và liệt kê dòng `LỆCH DÒNG` (so Mã + Tên tại ô đã nối).
- **Chốt giá trước khi nộp thầu**: Data → Edit Links → chọn từng link → **Break Link** (chuyển công thức thành giá trị, gỡ link). Làm ở bản `_chot`. Không chốt thì ai sửa file thư viện cũng làm đổi giá BOQ đã duyệt khi cập nhật link.

## 4. Với `boq-gan-ma-gia-thu-vien` khi TV đã nối link

- Không chạy `tv_loader.py load … --replace` và không chạy lại `hsloai` (phá link). Thư viện đổi thì nhờ người dùng Update Values.
- `recalc` bằng LibreOffice trên bản sao dùng cache của link nên vẫn ra giá (đã thử với file thử; chưa thử trên file thật của phòng).
- Ghi file bằng XML phải giữ nguyên `xl/externalLinks/*`, `<externalReferences>` và cột `Kiểm dòng nguồn`.

## 5. Đã kiểm / chưa kiểm

- Đã thử trên thư viện giả lập (tiêu đề 2 tầng, dòng nhóm, dòng trống đầu, sheet ẩn): nối đúng 15/15 ô giá, openpyxl và LibreOffice mở được file, LibreOffice nhận đúng 2 liên kết ngoài, `verify` bắt được việc chèn dòng.
- **Chưa thử trên Excel**: lần đầu mở file phải xem Excel có báo lỗi gói file hay không và bấm Update Values thử; chưa thử trên bộ TDI25538 thật.

## Chạy nhanh hơn (bản 22/10)

- `link` mặc định chỉ nối 2 cột giá thô (Giá VL / Giá NC thư viện, cột O–P). Giá gốc E/F đã bằng thô × hệ số nên vẫn cập nhật theo thư viện. Số link giảm khoảng một nửa. Muốn nối cả `Giá NC ước tính` thì thêm `--cols all`.
- `--no-check`: không thêm cột `Kiểm dòng nguồn`, file nhẹ hơn nhưng mất cảnh báo lệch dòng khi thư viện bị chèn/xóa dòng (`verify` khi đó chỉ so giá).
- `verify` đã sửa lỗi chậm: trước đây đọc từng ô nguồn riêng lẻ (mỗi lần quét lại sheet từ đầu, 6.887 dòng mất hơn 11 phút), nay mỗi sheet nguồn quét một lần.
- Khi cập nhật link trong Excel, nên đặt Formulas → Calculation Options → Manual, Update Values xong rồi nhấn F9 một lần.
