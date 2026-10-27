# Kiểm tra tính lại, vòng tham chiếu và lỗi đã gặp

## Kiểm tra tính lại (`--check full` hoặc kiểm tra thủ công)

**File do `build` vừa ghi**: không có giá trị cache. `--check full` dùng được.
- Nếu không có `/mnt/skills/public/xlsx/scripts/recalc.py` (Cowork): dùng `soffice --headless --convert-to xlsx` trên bản sao rồi đọc bằng `openpyxl(data_only=True)`.

**File đã được người dùng mở và lưu bằng Excel**: LibreOffice dùng lại giá trị cache cũ, nên kết quả đọc ra có thể sai. Phải xóa cache trên bản sao trước:
```python
re.sub(r'(<f>[^<]*</f>|<f [^>]*>[^<]*</f>|<f [^>]*/>)<v>[^<]*</v>', r'\1', sheet_xml)
```
Lệnh `recalc` của `boq_xml_tools.py` làm sẵn việc này.

Không bao giờ recalc trực tiếp file giao cho người dùng (LibreOffice có thể làm xê dịch định dạng CĐT). File đã bật cờ để Excel tự tính khi mở.

**LibreOffice không báo vòng tham chiếu** (SUBTOTAL lồng nhau vẫn ra số đúng), nên `recalc` sạch không có nghĩa là Excel sẽ không hiện "Circular Reference". Luôn chạy kiểm vòng ở mục 4.5.

## Vòng tham chiếu (chi tiết lịch sử)
   - **Lỗi đã gặp (dự án số 2, 15/10/2026)**: sheet có dòng tiêu đề cuối cùng không có dòng con (ví dụ "B – HẠNG MỤC NHÀ THẦU BỔ SUNG" nằm sát dòng tổng). `build` ghi `SUBTOTAL(9, <dòng kế tiếp>:<dòng kế tiếp>)` cho dòng tiêu đề; dòng kế tiếp chính là dòng "Tổng cộng trước thuế (TDI)", mà dòng tổng lại cộng cả dòng tiêu đề. Hai ô trỏ vòng vào nhau và kéo theo các ô Summary lấy số từ dòng tổng. Gặp ở 7/8 sheet MEP (3 ô mỗi sheet).
   - **Cách kiểm**: dựng đồ thị phụ thuộc của mọi công thức (kể cả tham chiếu chéo sheet, bỏ qua sheet `2 Thu vien gia` và `DS_MaHieu`) rồi tìm thành phần liên thông mạnh (Tarjan). Mẫu (đọc file bằng openpyxl, không ghi lại file):
   - **Cách sửa**: ô tiêu đề không có dòng con thì **bỏ công thức** (để trống, giữ style), không để `SUBTOTAL` trỏ tới dòng tổng. Dòng tổng vẫn là `SUBTOTAL(9, đầu:cuối)`: SUBTOTAL tự bỏ qua SUBTOTAL lồng nên số không đổi. Sửa bằng patch XML trên chính file (thay `<c r="AU33" s=".."><f>…</f><v/></c>` bằng `<c r="AU33" s=".."/>`), **không** lưu lại bằng openpyxl. Sau đó chạy lại kiểm vòng và `recalc`: 0 vòng, tổng các sheet không đổi.
   - **Sửa tận gốc trong script**: ở `boq_tdi.py` (hai chỗ ghi công thức dòng tiêu đề `section`), chỉ ghi `SUBTOTAL` khi `section_end(...) > r` và vùng cộng nằm hoàn toàn trong `[r+1, last]`; không ghi nếu tiêu đề không có dòng con. Chưa sửa script thì làm bước kiểm và sửa ở trên sau mỗi lần `build`.

## Lỗi `boq_tdi.py` đã gặp
- **Lỗi `boq_tdi.py` đã gặp (dự án MTT 20260403) và cách xử lý:**
  - **File CĐT có tên in `_xlnm.Print_Titles`/`Print_Area` = `#N/A`/`#REF!`** làm openpyxl không mở được: bỏ đúng các tên in hỏng đó trên bản sao (không xóa tên khác), chạy trên bản sao; báo người dùng.
  - **Cột Nhãn hiệu nhận nhầm** cột "TỔNG HẠNG MỤC" (H): kiểm `columns.brand` trong config; CĐT ghi hãng trong cột "THÔNG SỐ" thì đặt `brand = null`.
  - **Dòng tiêu đề đánh số "1.1", "1.3"** bị nhận là dòng ghi chú và nhận công thức như hạng mục → `#VALUE!` khi ô KL CĐT là ký tự trắng (`\xa0`). Sửa `kind` = `section`, `level` = 1 trong `rows.json` rồi chạy lại `--reuse-rows`.
  - **Tiêu đề CĐT trùng "NỘI DUNG CÔNG VIỆC"** (cột B) với tiêu đề phần TDI: script dùng chung `boq_xml_tools.py` bản mới lấy ô ngoài cùng bên phải; bản cũ nhận nhầm cột B.
  - Báo "CHƯA ĐẠT" chỉ vì `#REF!` có sẵn trong sheet CĐT (Total Summary, CFA…) hoặc lệch số thực dạng 42.228 → 42.22799999 là không phải lỗi; kiểm toàn bộ ô CĐT (không chỉ 3.000 ô mẫu) với dung sai số thực trước khi kết luận.
- **Lỗi vòng tham chiếu dòng tiêu đề cuối (dự án số 2, 15/10/2026)**: `build` ghi `SUBTOTAL` cho dòng tiêu đề không có dòng con, trỏ thẳng vào dòng "Tổng cộng trước thuế (TDI)" → Excel báo Circular Reference ở 7/8 sheet MEP và các ô Summary lấy số từ đó. Xử lý ở mục 4.5.
- **Thông tắc sàn ra TBVS (TGC-CT04, 27/10/2026)**: người dùng yêu cầu đầu mục "Thông tắc sàn" có Chủng loại riêng `TS`. Đã thêm luật `TS` trong `boq_tdi.py` (ngay sau các luật chi phí / thử nghiệm, trước luật RMU có `fco`, VTP, TBVS và các mã ống), thêm `TS` vào `THIET_BI` (giữ × HSTB như khi còn là TBVS) và `NON_DOMINATED`, sinh lại JS add-in bằng `build_excel_core.py`. Bảng mã: `chung_loai.md`.
