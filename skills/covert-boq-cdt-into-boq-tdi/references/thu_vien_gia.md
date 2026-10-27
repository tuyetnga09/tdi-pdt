# Thư viện giá ME/XD — vị trí, nạp `tv_loader.py`, kiểm tra

_Tách từ SKILL.md (không đổi nội dung nghiệp vụ). Đọc khi tới bước nạp thư viện (sau `run` đạt)._

### Vị trí thư viện giá: thư mục dự án trước, thư mục chung sau (chốt 16/10/2026)

Mỗi dự án có thư viện riêng, nằm **cùng cấp với file BOQ CĐT**:

```
<Thư mục dự án>\                    ví dụ ...\PDT_BOQ\BOQ BKY CN8_Khối lượng 工程量\
├── <BOQ CĐT>.xlsx, <BOQ CĐT>_TDI.xlsx
├── 2. Thu vien gia\                thư viện ME: một THƯ MỤC, nhiều file .xlsx theo nhóm
├── 2.Thu vien gia XD.xlsx          thư viện XD: một FILE, chứa sheet TVGIA
└── 03. HSDT\ (HSMT, spec…)
```

**Dò tên**: chuẩn hóa (bỏ dấu, chữ thường, bỏ số thứ tự đầu, gộp dấu chấm / gạch / gạch dưới / khoảng trắng). Tên chứa `thu vien gia`: là **thư mục** → ứng viên thư viện ME; là **file `.xlsx`** có thêm `xd` → ứng viên thư viện XD. Không nhầm hai loại: file XD không bao giờ được nạp vào bảng `TV`, thư mục ME không bao giờ thành `TVGIA`.

**Thư viện ME, thứ tự ưu tiên**
1. Đường dẫn người dùng chỉ trong chat.
2. Thư mục ứng viên ME trong thư mục dự án. Có từ hai ứng viên → hỏi một lần.
3. Không có → thư mục chung `2.Thu vien gia` của người dùng (dò theo skill `tdi-duong-dan-khong-co-dinh`, không ghi cứng đường dẫn), lọc theo mã dự án (cách cũ).

**Lọc file khi lấy từ thư mục dự án** (thư mục đã thuộc riêng dự án):
- Có file mang tiền tố mã dự án (ví dụ `TDI25538 - 01.BG-Cap dien_.xlsx`) → chỉ lấy các file đó, như cũ.
- **Không file nào mang tiền tố mã** → lấy mọi `.xlsx` trong thư mục (bỏ `~$…`, thư mục con, file `_backup`). `tv_loader.py` lọc theo `--code`: đọc đoạn lọc trong script; nếu `--code ""` (chuỗi rỗng) nhận mọi file thì dùng `--code ""` cho cả `load` và `check`. Không được thì stage bản sao vào thư mục cloud với tên `<mã> - <tên gốc>` rồi chạy như thường, ghi việc này vào "Bước đã bỏ qua / làm khác skill" (cột `File nguồn` sẽ mang tiền tố thêm vào). Không sửa script trên máy người dùng.
- File mang **mã dự án khác** (thư viện dự án khác chép nhầm vào), hoặc lẫn file có mã và không mã → dừng, liệt kê, hỏi.
- Vẫn kiểm như cũ: thiếu nhóm so với 18 file chuẩn, file `cloudOnly`, `mtimes.json`.

**Thư viện XD, thứ tự ưu tiên**
1. File người dùng chỉ trong chat.
2. File ứng viên XD trong thư mục dự án (ví dụ `2.Thu vien gia XD.xlsx`). Có từ hai ứng viên → hỏi một lần.
3. Không có → `01. File nội bộ TDI.xlsx` bản mới nhất của phòng (hỏi đường dẫn nếu chưa biết).

**Kiểm bố cục file XD trước khi dùng** (stage về cloud, đọc bằng openpyxl `read_only=True, data_only=True`, không ghi lại file):
- Có sheet tên đúng **`TVGIA`** (`boq_xd_tools.py build` chỉ chép sheet tên này).
- Bố cục khớp skill: **cột B** tên công tác, **cột G** giá VL 1 ĐVT, **cột J** giá NC 1 ĐVT, dữ liệu trong A–K. Cách kiểm: in hàng tiêu đề và 30 dòng đầu có tên ở B; đa số các dòng đó có số ở G hoặc J; ô công thức có giá trị đã tính (file đã lưu bằng Excel, không phải `None` hàng loạt).
- Sai tên sheet hoặc lệch cột → **dừng**, báo danh sách sheet và tiêu đề cột đang có, nhờ người dùng sửa file nguồn (đổi tên sheet, chép đúng bố cục). Không tự đổi tên, không dựng lại `TVGIA` bằng openpyxl, không tự đổi cột VLOOKUP.
- File BOQ **đã có `TVGIA`**: so với file nguồn vừa chọn (số dòng có tên ở B, ΣG, ΣJ, ghi chú nguồn ở L1). Lệch → báo, hỏi giữ ảnh chụp cũ hay thay; không tự thay.
- Ghi tên file và ngày sửa (mtime) để báo cáo.

**Script và Bảng khớp không đổi chỗ**: luôn lấy ở `2.Thu vien gia\_Claude_tools\` của thư mục chung, kể cả khi thư viện lấy theo dự án. Không tìm `_Claude_tools` trong thư mục dự án. Cần quyền cả hai nơi: `device_request_folder_access` thư mục dự án (đủ cho BOQ, thư viện ME / XD, HSMT) và thư mục chung `2.Thu vien gia`.

**Báo cáo** (điểm dừng 1 của skill gán mã, tóm tắt giao file của skill chuyển template): thư viện ME lấy ở đâu (dự án / thư mục chung / người dùng chỉ), cách lọc (theo mã / cả thư mục), số file; thư viện XD lấy ở đâu, tên file, ngày sửa, kết quả kiểm bố cục.

### 3. Nạp thư viện giá vào file (Bước 2 quy trình TDI) — tự động

Làm **tự động** ngay sau khi `run` đạt, **trước** khi giao file, cho mọi file có sheet MEP (file chỉ có sheet XD thì dùng `TVGIA`, bỏ qua mục này). Sheet thư viện là ảnh chụp giá tại thời điểm nạp (giá trị, không liên kết ngược), để giá trong BOQ không đổi khi thư viện gốc cập nhật.

**Thư mục thư viện ME**: dò theo mục "Vị trí thư viện giá" (mục Quy trình): người dùng chỉ → thư mục `2. Thu vien gia` trong thư mục dự án → thư mục chung `2.Thu vien gia` của người dùng (dò theo skill `tdi-duong-dan-khong-co-dinh`). Các bước dưới đây áp cho thư mục đã chọn.

1. **Liệt kê thư mục** bằng `device_list_dir`. Chưa được cấp quyền thì `device_request_folder_access` cho đúng thư mục này.
2. **Lọc theo mã dự án**: chỉ lấy file `.xlsx` có tên bắt đầu bằng mã dự án (ví dụ `TDI25538 - 01.BG-Cap dien_.xlsx`).
   - **Thư mục thư viện của dự án mà không file nào mang mã** → lấy cả thư mục theo mục "Vị trí thư viện giá" (`--code ""` hoặc bản sao có tiền tố), không dừng.
   - **Thư mục chung không có file nào khớp mã** → dừng, báo người dùng danh sách mã đang có trong thư mục và hỏi dùng bộ nào. Không tự lấy bộ của dự án khác.
   - Thiếu nhóm so với 18 file chuẩn → báo, hỏi có tiếp tục không.
   - File OneDrive chưa tải về (`cloudOnly`) → báo người dùng trước khi stage.
   - **Không truy cập được máy người dùng** (không có cầu nối): xin người dùng đính kèm các file thư viện của dự án; không có thì giao file chưa có thư viện và nói rõ.
3. **Stage các file đã lọc**. Ghi `mtimes.json` = `{tên file: mtimeMs}` lấy từ kết quả stage (ngày cập nhật thật trên máy người dùng).
4. **Lấy `tv_loader.py` bản 15/10/2026** (stage từ `_Claude_tools`, hoặc `project_read` `claude/boq/tv_loader.py` ở Project). **Không dùng bản cũ** còn cột `Column1` (nhận ra qua `GROUP='Column1'` hoặc không có lệnh in "Loại vật liệu"). Rồi chạy:
   ```bash
   python3 /tmp/tv_loader.py load "<thư mục thư viện ME đã stage>" "<file>_TDI.xlsx" "<file>_TDI.tmp.xlsx" --code <mã dự án> --mtimes mtimes.json
   mv "<file>_TDI.tmp.xlsx" "<file>_TDI.xlsx"
   ```
   Kết quả: sheet `2 Thu vien gia` ở cuối file, bảng Excel tên **`TV`**. Các cột:
   - **`Loại vật liệu`** (cột đầu, trước 15/10/2026 tên `Column1`; chính là Nhóm VT, công thức dùng `TV[Loại vật liệu]`)
   - `Mã hiệu`, `Tên công việc`, `ĐVT`
   - `Giá VL gốc`, `Giá NC gốc`, `Giá VL chưa CK`, `Giá NC ước tính`
   - `Description`, `Ghi chú TV`, `Nhãn Hiệu Xuất Xứ`
   - `Sheet nguồn`, `File nguồn`, `Ngày cập nhật` (dùng làm nguồn giá "Thư viện")
   - `Nhóm file` (nhóm cũ theo file: TCap, OLDay, PVC…; chỉ để đối chiếu và chuyển đổi BOQ cũ, không dùng để gán)
5. **`Loại vật liệu` ghi rõ theo từng sheet con**, không còn dùng chung một mã cho cả file. Tên theo bảng "Loại VL" của Summary TDI. Ví dụ file `Thang cap` có sheet "Thang cap" → `Thang cáp`, sheet "PK thang cap" → `Phụ kiện T.cáp`.

   | Nhóm file | Sheet nguồn → Loại vật liệu |
   |---|---|
   | Cap dien | Cáp điện |
   | OLDay | Ống luồn dây |
   | TCap | Thang cap → Thang cáp; PK thang cap → Phụ kiện T.cáp |
   | MCap | Mang cap → Máng cáp; PK mang cap → Phụ kiện M.cáp |
   | OThep | Ong thep → Ống thép; PK thep → PK ống thép |
   | Van Khoa | Van thep → Van Khóa Thép |
   | CTKhac | Công tác khác |
   | ODong | Ống đồng |
   | TB Dien | Thiết bị Điện |
   | TB Ve Sinh | Thiết bị Vệ Sinh |
   | TB HVAC | TB dieu hoa → Điều Hoà; TB thong gio → Quạt thông gió; Van gio → Van gió |
   | PVC | Ong PVC → Ống PVC; PK PVC → PK Ống PVC |
   | PPR | Ong PPR → Ống PPR; PK PPR → PK Ống PPR |
   | HDPE | Ong HDPE → Ống HDPE; PK HDPE → PK HDPE |
   | OGio | Ong gio → Ống gió; PK ong gio → PK Ống gió; Cua gio (sheet ẩn) → Cửa gió |
   | Bao on | Bảo ôn |
   | PCCC | TB PCCC → Thiết bị chữa cháy |
   | TB Bao chay | TB PCCC → Thiết bị báo cháy |

   - Tổng cộng **28 tên hợp lệ**: Cáp điện, Ống luồn dây, Thang cáp, Phụ kiện T.cáp, Máng cáp, Phụ kiện M.cáp, Ống thép, PK ống thép, Van Khóa Thép, Công tác khác, Ống đồng, Thiết bị Điện, Thiết bị Vệ Sinh, Điều Hoà, PK Ống PVC, Ống PVC, Ống PPR, PK Ống PPR, Ống HDPE, PK HDPE, Ống gió, PK Ống gió, Cửa gió, Van gió, Bảo ôn, Quạt thông gió, Thiết bị chữa cháy, Thiết bị báo cháy.
   - Bảng ánh xạ nằm trong `tv_loader.py` (`LOAI_VL`). **Sheet lạ không có trong bảng** thì loader dùng nhóm file làm Loại vật liệu và in cảnh báo: báo người dùng, không tự đặt tên mới.
   - Tên Loại vật liệu phải khớp bảng "Loại VL – Hệ số VL/NC" trong Summary TDI (28 dòng); lệch thì báo.
6. **Thư viện chuẩn phải ĐỦ và ĐÚNG - quy tắc bắt buộc của loader** (bản v1 và Power Query từng thiếu 685/6.887 dòng của bộ TDI25538):
   - **Đọc mọi sheet, kể cả sheet ẩn.** Sheet ẩn ghi `"<tên> (sheet ẩn)"` ở cột Sheet nguồn. Đã gặp: sheet ẩn "Cua gio" của file `10.BG-OGio` (591 dòng cửa gió 1/2 lớp, louver nan Z, khuếch tán, tiêu âm FCU).
   - **Giữ dòng không có mã hiệu** nếu có Tên công việc và ĐVT (hoặc giá). Đã gặp: 73 dòng CTKhac (cắt đục tường, chèn trát, giá đỡ thang máng/ống gió/ống nước, thử áp, đầu cốt). Ghi chú TV: "Không có mã hiệu - tra theo tên".
   - **Giữ dòng thiếu ĐVT** nếu có đủ Mã và Tên và có giá. Đã gặp: 21 dòng `WE-D…-HDPE` (cút hàn ghép HDPE). Ghi chú TV: "Thiếu ĐVT trong file nguồn".
   - **Chỉ bỏ** dòng nhóm/tiêu đề (tên nhóm nằm ở cột Mã, giá là hệ số 1 / 1,165…) và dòng mẫu trống (mã 0/trống và tên trống). Loader in số dòng bỏ theo từng sheet.
   - **Giữ nguyên giá kể cả 0**; ô giá dạng chữ ghi vào Ghi chú TV. Không làm tròn, không đổi đơn vị.
   - **Tự đối chiếu**: đếm lại độc lập số dòng, ΣVL gốc, ΣNC gốc từng (file, sheet) nguồn so với đã nạp. Lệch bất kỳ sheet nào thì **dừng, không ghi file**; tìm nguyên nhân, không được bỏ qua.
   - **Báo khóa trùng** Loại vật liệu+Mã và Loại vật liệu+Tên (bộ TDI25538 có 51 mã trùng ở Máng cáp 350x100…, 124 tên trùng). Skill gán mã không dùng các khóa này cho SUMIFS.
   - **Tên vùng trùng "TV"** (không phân biệt hoa thường): loader xóa tên rác ẩn (`{…}`, `#REF!`) và báo; tên thật thì dừng. Đã gặp: tên ẩn `tv` làm Excel tự đổi bảng thành `TV_1`, mọi công thức `SUMIFS(TV[...])` lỗi.
   - **Hai file cùng chứa sheet "TB PCCC"** (file PCCC và file TB Bao chay): loader gán Loại vật liệu khác nhau (`Thiết bị chữa cháy` / `Thiết bị báo cháy`). Nếu nội dung hai sheet trùng nhau thì báo người dùng.
7. **File đã có bảng `TV`**: loader dừng. Hỏi người dùng có thay không; thay thì chạy lại với `--replace` (giữ tên sheet/bảng, thay toàn bộ nội dung). `--replace` tự đổi `TV[Column1]` thành `TV[Loại vật liệu]` trong công thức của mọi sheet. Nếu file cũ còn nhóm cũ ở cột Nhóm VT (`TCap`, `OLDay`…) thì chạy tiếp `boq_xml_tools.py migrate` (skill `boq-gan-ma-gia-thu-vien`, Bước 0b). Sau khi thay, `recalc` lại và so tổng: giá đã gán chỉ được đổi nếu file nguồn đổi.
8. **Kiểm tra bắt buộc trước khi giao** (lệnh độc lập, đọc chính file đầu ra):
   ```bash
   python3 /tmp/tv_loader.py check "<thư mục thư viện ME đã stage>" "<file>_TDI.xlsx" --code <mã dự án>
   ```
   Phải ra `THIẾU 0 | THỪA 0 | LỆCH GIÁ 0`. Kiểm thêm `xl/tables/*.xml` có `displayName="TV"` (không phải `TV_1`) và cột đầu của bảng tên đúng `Loại vật liệu`.

Đã kiểm chứng với bộ TDI25538 (15/10/2026): 18 file, 28 sheet (1 sheet ẩn), **6.887 dòng**, 28 giá trị `Loại vật liệu`; `check` đạt 0/0/0. Bảng `TV` cũ 6.282 dòng (v1/Power Query) thiếu 685 dòng và thừa 80 dòng mẫu trống.

## Script nạp thư viện `tv_loader.py`

Skill không còn nhúng mã nguồn loader (bản nhúng cũ dùng cột `Column1`, sẽ gây sai). Lấy bản 15/10/2026 từ một trong hai nơi:
1. `_Claude_tools\tv_loader.py` trong thư mục thư viện (dò vị trí theo skill `tdi-duong-dan-khong-co-dinh`, không ghi cứng đường dẫn; stage về `/tmp/tv_loader.py`).
2. Project của phòng đấu thầu: `project_read` đường dẫn `claude/boq/tv_loader.py`, ghi ra `/tmp/tv_loader.py`.

Không lấy được cả hai → xin người dùng đính kèm; **không tự viết lại**. Cần `openpyxl`, `lxml`. **Không dùng bản v1 hoặc bản còn cột `Column1`** (bỏ sheet ẩn, bỏ dòng không mã/thiếu ĐVT, chung một nhóm cho cả file).

Lệnh: `load LIBDIR BOQ.xlsx OUT.xlsx --code <mã> [--mtimes m.json] [--replace]` và `check LIBDIR BOQ.xlsx --code <mã>`.