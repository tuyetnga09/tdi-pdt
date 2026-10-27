---
name: "covert-boq-cdt-into-boq-tdi"
description: "Chuyển BOQ (bảng khối lượng mời thầu) của Chủ đầu tư / Tổng thầu sang template chào giá TDI, giữ nguyên sheet CĐT bên trái, dựng phần TDI bên phải, nạp thư viện giá và gắn bộ lọc cho mọi cột."
---

# Chuyển BOQ CĐT sang template TDI (hệ Cơ điện + hệ Xây dựng)

Ứng với **Bước 1 và Bước 2** của quy trình lập giá TDI: chuyển BOQ sang template, rồi nạp thư viện giá vào cùng file. Skill nối tiếp là `boq-gan-ma-gia-thu-vien` (Bước 3–4, 6).

**Hai hệ.** Sheet Cơ điện (MEP) làm theo các mục dưới đây với `boq_tdi.py`. Sheet **Xây dựng (XD)** làm theo `references/xd.md` với `boq_xd_tools.py` (đọc file đó khi file có sheet XD). File có cả hai thì chạy cả hai, mỗi sheet một lần.

## QUY TẮC VỊ TRÍ FILE VÀ LINK (chốt 26/10/2026, sau lỗi "Không thể đọc sổ làm việc được liên kết")

**Lỗi đã gặp (26/10/2026, dự án TDI26582):** file `_TDI.xlsx` có link thư viện tương đối `2.Thu vien gia/<file>`, nhưng người dùng mở bản nằm ở thư mục con `Claude outputs\` (bản tự đồng bộ từ `/mnt/user-data/outputs/` lên SharePoint). Excel ghép đường dẫn tương đối vào `Claude outputs\2.Thu vien gia\…` (không tồn tại) nên báo "Không thể đọc sổ làm việc được liên kết". Cùng lúc, bản ở thư mục dự án lại là bản CŨ (chưa có thư viện, chưa có link) vì lần ghi đầu của `run` lên máy người dùng không được thay bằng bản cuối. **Nguyên nhân gốc: link tương đối chỉ đúng khi file nằm ĐÚNG thư mục dùng để tính `--linkrel`; file bị nhân bản sang thư mục khác thì link vỡ.**

Quy tắc để không gặp lại:

1. **Chốt thư mục đích trước khi nối link.** Thư mục đích = thư mục chứa file BOQ CĐT gốc trên máy người dùng (cùng cấp thư mục thư viện, thường `...\3.1 HS thực hiện chào giá\`). `--linkrel` luôn tính từ thư mục đích này, **không** tính từ `/mnt/user-data/outputs/` hay thư mục tạm.
2. **Không để `boq_tdi.py run` ghi mặc định vào `/mnt/user-data/outputs/`.** Mọi lần chạy `run` / `tv_loader` / `hsloai` / `tv_link` / `tdi_filter` đều dùng `--out` hoặc thư mục làm việc tạm (`/home/claude/w/…`, scratchpad). Chỉ **bản cuối đã hoàn tất** (đã nạp thư viện, đã nối link, đã lọc, đã kiểm) mới được chép vào `/mnt/user-data/outputs/`. Lý do: file ở `outputs/` được tự đồng bộ/giao cho người dùng và có thể thành bản thứ hai ở `Claude outputs\` trên SharePoint; bản trung gian ở đó sẽ bị người dùng mở nhầm.
3. **Chỉ một bản duy nhất.** Sau khi ghi vào thư mục đích bằng `device_commit_files`, báo rõ cho người dùng: "mở đúng file ở `<thư mục đích>`, không mở bản trong `Claude outputs`". Nếu người dùng đã có bản trong `Claude outputs`, bảo họ đóng và bỏ bản đó.
4. **Kiểm sau khi commit, trên máy người dùng** (bằng `device_bash`, không tin báo `written`): (a) kích thước và số phần tử zip của file ở thư mục đích khớp file cuối trên cloud (nếu file đích nhỏ hơn rõ rệt / không có `xl/externalLinks/*` thì commit chưa ghi bản cuối: commit lại bằng `fileUuid` của `SendUserFile` mới nhất và `expectedMtimeMs` lấy từ `device_list_dir` ngay trước đó); (b) file có đủ số link (`xl/externalLinks/_rels/*` = số file thư viện, ví dụ 18); (c) từ thư mục chứa file, đường dẫn tương đối của từng `Target` (giải mã `%20`) tồn tại: `test -f "<thư mục đích>/<linkrel>/<tên file>"` cho cả 18 file. Thiếu thì dừng, không báo xong.
5. **Không còn bản `_TDI.xlsx` cũ ở thư mục đích trước khi giao.** Trước khi commit, `device_list_dir` thư mục đích: nếu đã có `<tên>_TDI.xlsx`, so mtime với lần ghi trước của phiên này; ghi đè bằng bản cuối, không để bản cũ khác nội dung.
6. **Nói rõ trong tóm tắt giao file** ba điều: (a) link là tương đối, file phải nằm cùng cấp thư mục `<linkrel>`; (b) không chép/di chuyển file BOQ riêng ra thư mục khác (chép cả thư mục dự án); (c) nếu mở file thấy báo không đọc được liên kết thì kiểm file đang mở có nằm đúng thư mục không, rồi Data → Workbook Links → Change Source trỏ về `2.Thu vien gia` (làm một lần).
7. **Khi người dùng buộc file nằm ở thư mục khác thư mục chuẩn** (ví dụ giao vào `Claude outputs`): dùng `--linkrel "../2.Thu vien gia"` cho đúng vị trí đó, hoặc nói rõ link sẽ vỡ nếu mở ở chỗ khác. Không tự quyết; hỏi người dùng một lần nếu thư mục đích không rõ.

**Không ghi cứng đường dẫn máy.** Không bao giờ đưa tên user Windows hoặc ổ đĩa (`C:\Users\Admin\...`, `C:\Users\lexxu\...`) vào file giao cho người dùng hay vào lệnh mặc định. Đường dẫn trên máy lấy lúc chạy từ `get_device_info` / `resolvedPath` của `device_list_dir`, hoặc hỏi người dùng. Mọi liên kết ngoài trong file Excel phải là **đường dẫn tương đối so với thư mục chứa file BOQ** (xem mục 3b), để file chạy được trên máy khác, tài khoản khác, OneDrive / SharePoint sync ở ổ khác.

**Thư viện giá được nạp tự động.** Với file có sheet MEP, ngay sau khi chuyển template đạt, **tự nạp thư viện giá vào cùng file** (sheet `2 Thu vien gia`, bảng `TV`, cột đầu `Loại vật liệu`) và tự kiểm `check` = 0/0/0, **không chờ người dùng yêu cầu thêm**. Với sheet XD, nạp `TVGIA` khi `build`. Chi tiết (vị trí thư viện, lọc file, loader, quy tắc đủ/đúng): **`references/thu_vien_gia.md`**. Chỉ không nạp khi người dùng nói rõ "không nạp thư viện".

**Bộ lọc cho mọi cột (thêm 22/10/2026).** Mọi sheet BOQ sau khi chuyển phải có bộ lọc (AutoFilter) phủ **tất cả các cột**, từ cột đầu của CĐT đến cột cuối của phần TDI (kể cả 5 cột phụ gán mã nếu đã có), để người lập giá lọc theo Chủng loại, Mã hiệu, Nhóm VT, độ tin cậy, dòng trống giá… Làm tự động sau khi nạp thư viện, bằng `tdi_filter.py` (mục 3c), **không chờ người dùng yêu cầu thêm**. Chỉ không gắn khi người dùng nói rõ "không cần bộ lọc". **Dòng chứa nút lọc phải nằm trong vùng cố định (freeze pane)** (thêm 23/10/2026): nếu freeze dừng ở dòng trên dòng lọc thì khi cuộn xuống nút lọc trôi mất và người dùng báo "không thấy bộ lọc" dù ref đã phủ đủ cột. `tdi_filter.py apply` tự nâng freeze tới dòng lọc; `verify` báo LỖI nếu chưa đạt.

**Link sống thư viện giá (thêm 22/10/2026, sửa 26/10/2026: đường dẫn tương đối + quy tắc vị trí file ở trên).** Bảng `TV` nạp bằng `tv_loader.py` là ảnh chụp giá. Khi người dùng muốn giá trong BOQ đổi theo file thư viện (thư mục → BOQ, một chiều, BOQ không ghi ngược vào thư viện), làm theo mục 3b: sau khi nạp và `check` đạt, nối ô giá của `TV` vào file thư viện bằng workbook link (Data → Edit Links) với `tv_link.py link ... --linkrel "<đường dẫn tương đối>"`. Bản `tv_link.py` 26/10/2026 tự ghi link **tương đối**, không cần bước sửa riêng.

**Cột Đơn vị của phần TDI để trống (chốt 23/10/2026).** Skill này **không chép ĐVT của CĐT** sang phần TDI: ô Đơn vị TDI (cột "Đơn Vị" bên phải, `r_unit`) giữ tiêu đề nhưng **để trống**, không còn công thức `=+G`/`=+H`. ĐVT của phần TDI chỉ đến từ **thư viện giá** (`TV[ĐVT]`), do skill `boq-gan-ma-gia-thu-vien` ghi cùng lúc với Mã hiệu và Đơn giá gốc (Quy tắc 38 của skill đó); dòng chưa gán mã hoặc thư viện không có ĐVT thì ô Đơn vị TDI **tiếp tục để trống**, không dự phòng bằng ĐVT CĐT. Công thức cột `KL` (hao hụt m / m2 / m3) đọc **ĐVT của CĐT** (cột Đơn vị bên trái) vì `Tổng KL tạm tính` là khối lượng theo ĐVT CĐT; công thức đơn giá VL vẫn chọn nhánh hao hụt / thiết bị / vật liệu như cũ. **Không tự điền, không đoán ĐVT cho ô Đơn vị TDI ở skill này.** Phạm vi: MEP (`boq_tdi.py`, cả hai chế độ và add-in `tdi_b` / `tdi_write`); hệ XD (`references/xd.md`, `boq_xd_tools.py`) chưa đổi, vẫn tham chiếu ĐVT CĐT ở c3 và cột Đơn vị của khối Vật liệu / Nhân công, chờ quyết định riêng.

**Lưu ý (26/10/2026):** `boq_tdi.py` bản đang cài vẫn còn ghi `=+<ĐVT CĐT>` vào ô Đơn vị TDI (dòng `put(ws, r, c["r_unit"], ...)`) và công thức `KL` trỏ ô Đơn vị TDI. Nếu script chưa được sửa, sau `run` phải vá bằng XML trên file: xóa nội dung cột Đơn vị TDI cho các dòng hạng mục (giữ tiêu đề và style) và đổi `<cột Đơn vị TDI><r>="m"` trong công thức `KL` thành `<cột Đơn vị CĐT><r>="m"` (ví dụ `BL12="m"` thành `R12="m"`; kiểm bằng cách đọc lại vài ô). Ghi vào tóm tắt là đã vá. Không mở-lưu bằng openpyxl.

Hai chế độ (MEP), chọn bằng `"mode"` trong config:

- **`inplace` (mặc định).**
  - Đầu ra là chính file CĐT, lưu tên `<tên gốc>_TDI.xlsx`.
  - Bên trái giữ nguyên tuyệt đối: giá trị, công thức, font, màu, đậm/nghiêng, rich text, merge, độ rộng, chiều cao dòng, freeze, các sheet khác. **Bộ lọc là ngoại lệ có kiểm soát**: Excel chỉ cho một AutoFilter mỗi sheet nên bộ lọc được mở rộng sang phải để phủ cả phần TDI (mục 3c). Điều kiện lọc đang bật của CĐT (nếu có) được giữ nguyên. **Freeze cũng là ngoại lệ có kiểm soát**: chỉ được nâng ySplit lên bằng dòng lọc khi freeze của CĐT đang dừng ở dòng trên đó; không bao giờ giảm, không tạo freeze mới.
  - Bên phải là phần TDI: bắt đầu ngay sau cột cuối của CĐT, cùng dòng với CĐT.
  - Thêm sheet `Summary` (hoặc `Summary TDI` nếu tên đã có).
- **`rebuild`.** Dựng sheet TDI mới theo khuôn riêng. Chỉ dùng khi người dùng muốn file TDI tách riêng. Mode này `boq_tdi.py` đã tự đặt AutoFilter từ dòng 12; vẫn chạy `tdi_filter.py verify` để chắc phủ hết cột và dòng lọc nằm trong vùng cố định.

Phân công:

- **Script `scripts/boq_tdi.py`**: nhận diện cấu trúc, công thức, định dạng theo `assets/style_template.xlsx`, Summary, kiểm tra (gồm kiểm vòng tham chiếu).
- **`tv_loader.py` (bản 15/10/2026, cột `Loại vật liệu`)**: nạp thư viện giá, tự đối chiếu và dừng nếu thiếu/lệch. Lệnh `check` kiểm lại bảng `TV`. Xem `references/thu_vien_gia.md`.
- **`tdi_filter.py` (`2.Thu vien gia\_Claude_tools` hoặc Project `claude/boq/tdi_filter.py`, bản 23/10/2026 có xử lý freeze)**: gắn / mở rộng bộ lọc cho mọi cột và đưa dòng lọc vào vùng cố định bằng patch XML, không mở-lưu bằng openpyxl nên không phá link thư viện, bảng `TV`, định dạng.
- **`tv_link.py` (`_Claude_tools` hoặc Project `claude/boq/tv_link.py`, bản 26/10/2026)**: nối link thư viện bằng workbook link, **luôn ghi đường dẫn tương đối** (`--linkrel`), có lệnh `relink` sửa file đã nối tuyệt đối / sai thư mục và `verify` kiểm. Chỉ ghi XML. Bản cũ hơn (còn `--linkbase` ghi tuyệt đối) không dùng.
- **Claude**: xác nhận cấu trúc; soát mã **Chủng loại** (quyết định hệ số giá: thiết bị × HSTB, vật liệu × HSVL, hàng theo mét thêm hao hụt, nên gán sai là giá sai).

Tài liệu (chỉ đọc khi cần):

- `references/chung_loai.md`: bảng mã Chủng loại theo bảng "Tra cứu từ khóa" TDI (khi có nhóm chưa có mã hoặc cần soát).
- `references/layout_inplace.md`: layout inplace + hệ số theo Chủng loại (khi đối chiếu cột/ô file inplace).
- `references/template_layout.md`: layout **rebuild** (cột L–BC). **Không dùng để đối chiếu file inplace.**
- `references/thu_vien_gia.md`: thư viện ME/XD, loader.
- `claude/boq/link_workbook_thu_vien.md` + `claude/boq/tv_link.py` (hoặc `2.Thu vien gia\_Claude_tools\`): link sống thư viện bằng workbook link (cách làm, kiểm, chốt giá). Phương án dự phòng không lệch dòng: `claude/boq/power_query_tv_song.md` (Power Query, người dùng tự dán; trong query dùng đường dẫn tương đối hoặc tham số, không ghi cứng `C:\Users\...`).
- `claude/boq/tdi_filter.py`: gắn bộ lọc cho mọi cột và đưa dòng lọc vào vùng cố định (mục 3c). Đọc rồi chạy, không tự viết lại.
- `references/xd.md`: hệ Xây dựng.
- `references/loi_da_gap.md`: recalc, lịch sử lỗi đã gặp.
- `references/excel_addin.md`: chạy trong add-in Claude for Excel (không có Python). Khi xong phải xóa sheet ẩn `_TDI_JOB`, `_TDI_TMP_*`.

## Mục tiêu: 2–3 phút cho cả file

- **Một lệnh `run` làm hết phần chuyển template, rồi nạp thư viện giá và gắn bộ lọc trong cùng lượt.** Không gọi lẻ `inspect` / `auto` / `extract` / `build` / `check`, trừ khi `run` báo lỗi hoặc nhận diện sai.
- **Chế độ nhanh (mặc định)**: chỉ xử lý các **nhóm cần soát** mà `run` in ra (có `??` hoặc mã lẫn lộn). Nhóm thuần một mã thì tin kết quả tự động. Nhóm không chắc thì để trống và báo. Soát kỹ khi người dùng yêu cầu.
- Chỉ đọc `chung_loai.md` khi có nhóm chưa có mã cần quyết định. Báo cáo cuối gọn.
- **Gom lời gọi tool qua cầu nối máy người dùng** (mỗi lời gọi là một lượt chờ): xin quyền **một lần** cho mọi thư mục cần (`device_request_folder_access` với `paths` = thư mục dự án + thư mục chung `2.Thu vien gia`); `device_list_dir` thư mục dự án **một lần `recursive=true`** để dò thư viện ME/XD (và để kiểm đường dẫn BOQ người dùng đưa có tồn tại không: nếu không, liệt kê thư mục cha và dùng file `.xlsx` cùng tên; đường dẫn có dấu tiếng Việt báo "does not exist" thì dò bằng `device_bash` với `$HOME/mnt/<tên thư mục>` rồi stage bằng đường dẫn `~/mnt/...`); **stage một lần** mọi file cần (BOQ + file thư viện đã lọc + `tv_loader.py` + `tv_link.py` + `boq_xml_tools.py` + `tdi_filter.py` nếu có + `boq_xd_tools.py` nếu có XD) rồi ghi `mtimes.json` từ kết quả stage. Đường dẫn có chữ có dấu hoặc khoảng trắng lạ có thể bị từ chối khi xin quyền từ xa: khi đó nhờ người dùng bấm "Add folder" trong app.

## Quy trình (Cowork / Desktop / claude.ai)

`S=<thư mục skill>/scripts/boq_tdi.py`.

**Chuẩn bị đầu vào**:
- File `.xls`: chuyển sang `.xlsx` trước (`soffice --headless --convert-to xlsx`).
- File `.xlsm`: vẫn giữ macro.
- File trên máy người dùng (Cowork có thư mục kết nối): stage vào `/mnt/user-data/uploads/…` rồi **chép sang thư mục làm việc tạm** (`/home/claude/w/`) và chạy ở đó; không chạy ghi vào `outputs/`.
- **Mã dự án** (ví dụ `TDI25538`): lấy từ người dùng, hoặc từ tên file/thư mục BOQ, hoặc từ tiền tố file thư viện. Chỉ hỏi khi không suy ra được hoặc có nhiều ứng viên. Mã này dùng để chọn bộ file thư viện giá, nên cần chắc **trước khi chạy `run`**. Thư mục dự án chỉ có thư viện mang mã dự án khác (ví dụ dự án TDI26582 dùng bộ TDI25538) thì hỏi người dùng một lần có dùng bộ đó không.
- **Thư mục dự án** = thư mục chứa file BOQ CĐT; dò thư viện theo `references/thu_vien_gia.md`. **Đây cũng là thư mục đích của file `_TDI.xlsx`** (xem Quy tắc vị trí file và link).

### 1. Một lệnh cho cả file

```bash
mkdir -p /home/claude/w && cd /home/claude/w
python3 $S run "<file CĐT đã chép vào /home/claude/w>.xlsx" --project "<nếu có>" --so_bao_gia "<…>" --out "/home/claude/w/<tên gốc>_TDI.xlsx"
```

`run` làm lần lượt:
1. `auto`: tìm sheet BOQ, nhận diện tiêu đề 1–2 tầng, dòng đầu–cuối, các cột, khu KL, tổng KL.
2. `extract`: phân loại dòng, gán Chủng loại.
3. In nhóm cần soát.
4. `build`: ghi file `_TDI.xlsx` (**vào `--out`, thư mục tạm**, không vào `outputs/`).
5. `check` nhanh: soát tĩnh công thức, **kiểm vòng tham chiếu (dòng `4)`)**, so 3.000 ô CĐT / sheet với file gốc, in `KẾT LUẬN: ĐẠT / CHƯA ĐẠT`.

**Sau `run`, SOÁT config trước khi tin kết quả** (lỗi đã gặp 26/10/2026 với TGC-CT04: `run` nhận diện sai `total_col`, khu KL và dòng cuối, ra 12 khu thay vì 6, tổng ở cột AB thay vì Y):
- `zones` phải đúng các cột khối lượng thực tế (không lẫn cột "Khối lượng nhà thầu" / "Đơn giá" / "Thành tiền" bên phải); `total_col` đúng cột tổng KL tham chiếu CĐT; cột `spec`, `desc`, `code`, `unit` đúng (in vài dòng đối chiếu); `data_end_row` là dòng cuối dữ liệu **trước** các dòng "NHÀ THẦU BỔ SUNG" / "TỔNG CỘNG" / "HƯỚNG DẪN" của CĐT, không phải cuối vùng lọc CĐT.
- Sai thì sửa `config.json` rồi chạy lại `run --config` (mục 2).

**Ngay sau `run` đạt (hoặc đạt có giải thích), nạp thư viện giá (`references/thu_vien_gia.md`, mục 3) trên file `_TDI.xlsx` tạm, rồi vá cột Đơn vị TDI (nếu script chưa sửa), nối link (mục 3b, nếu người dùng muốn link sống) và gắn bộ lọc (mục 3c)**, sau đó mới kiểm tra (mục 4) và giao (mục 5). Không dừng hỏi "có nạp thư viện không", "có gắn bộ lọc không".

File trung gian: `/home/claude/boq/` (`config.json`, `rows.json`) và `/home/claude/w/`.

| Tùy chọn | Tác dụng |
|---|---|
| `--mode rebuild` | Sheet TDI riêng |
| `--out <file>` | Đường dẫn file đầu ra (luôn đặt về thư mục tạm) |
| `--hang_muc`, `--ngay` | Thông tin cho Summary |
| `--max-groups N` | Số nhóm cần soát in ra mỗi sheet (mặc định 25) |
| `--check full` | Tính lại bằng LibreOffice trên bản sao (xem `references/loi_da_gap.md`) |
| `--check none` | Bỏ kiểm tra |

### 2. Chỉ khi cần: sửa rồi chạy lại

**Nhận diện sai** (khu, dòng cuối, cột, sheet):
- Sửa `/home/claude/boq/config.json` (`header_rows`, `data_start_row`, `data_end_row`, `total_col`, `columns`, `zones`, `tdi_start_col`).
- Chạy lại `run --config /home/claude/boq/config.json`, **không** thêm `--reuse-rows`.
- Đặt `tdi_start_col` khi CĐT có cột ghi chú hoặc cột ẩn ở xa bên phải. Không để phần TDI đè lên cột Ghi chú/Remark của CĐT.

**Sửa mã Chủng loại**:
- Ghi vào `/home/claude/boq/overrides.json`:
  ```json
  {"<tên sheet>": {"ctx:<khóa nhóm, chép nguyên văn từ dòng ctx:>": "HDPE", "1234": "VAN"}}
  ```
  (khóa số là số dòng nguồn; giá trị rỗng `""` để trống mã).
- Chạy lại:
  ```bash
  python3 $S run --config /home/claude/boq/config.json --reuse-rows --overrides /home/claude/boq/overrides.json --out "<cùng file đầu ra tạm>"
  ```
- Chạy lại `run` ghi đè file `_TDI.xlsx` và **làm mất sheet `2 Thu vien gia` cùng bộ lọc và link**: sau đó phải nạp lại thư viện và làm lại các bước sau.

Gợi ý soát:
- **Ống**: đúng vật liệu và đúng hệ; ống trong hệ Điện là OLD.
- **Thiết bị dưới hệ máy phát, báo cháy, chữa cháy khí / bình xịt, foam**: lấy mã của hệ (MFD, TBBC, FM200, FOAM).
- **Cáp ≥ 70mm² hoặc trung thế**: CAPLON. Cáp trong hệ Điện nhẹ: CAP DN.
- **"Phụ kiện …"**: VTP. Co/cút/tê ghi rõ vật liệu: theo mã ống.
- **Thông tắc sàn**: TS (nhóm thiết bị, × HSTB; thêm 27/10/2026). Không sửa về TBVS hay mã ống.
- Chỉ dùng mã trong `references/chung_loai.md` (bộ mã theo bảng "Tra cứu từ khóa" TDI).
- Sửa theo số dòng ưu tiên hơn sửa theo nhóm.

### 3. Nạp thư viện giá (tự động, sau `run` đạt)

Làm theo `references/thu_vien_gia.md` (mục 3 trong file đó), **trước** khi giao file, cho mọi file có sheet MEP. Kết quả bắt buộc: bảng `TV`, cột đầu `Loại vật liệu`, `tv_loader.py check` = `THIẾU 0 | THỪA 0 | LỆCH GIÁ 0`.

### 3c. Bộ lọc cho tất cả các cột (tự động, thêm 22/10/2026; freeze thêm 23/10/2026)

**Mục tiêu**: mỗi sheet BOQ có một AutoFilter phủ **toàn bộ cột** từ cột đầu CĐT (STT / Mô tả…) đến cột cuối phần TDI (Ghi Chú, và 5 cột phụ gán mã `Nhóm VT | Độ tin cậy | Ghi chú gán mã | Hệ số ĐV | Loại dòng` nếu đã có), để lọc được theo mọi cột: Chủng loại, Mã hiệu, Nhãn hiệu, Đơn vị, Giá gốc VL/NC trống, Nhóm VT, Độ tin cậy… **Và nút lọc phải luôn nhìn thấy khi cuộn** (dòng lọc nằm trong vùng cố định).

**Cách làm** (script `tdi_filter.py` ở `_Claude_tools` hoặc Project `claude/boq/`; chạy trên file `_TDI.xlsx`, ghi bằng XML nên **an toàn cả khi file đã nạp thư viện / đã nối link**; không mở-lưu bằng openpyxl):

```bash
F=/home/claude/w/tools/tdi_filter.py
python3 $F apply "<file>_TDI.xlsx" --config /home/claude/boq/config.json        # mọi sheet MEP trong config
python3 $F verify "<file>_TDI.xlsx"                                              # in ref từng sheet + kiểm freeze
```

**Lưu ý (26/10/2026):** `apply --config` chỉ phủ tới cột cuối có chữ ở dòng tiêu đề. Nếu ref in ra dừng ở cột `Total` (trước `Ghi Chú`) thì chạy lại từng sheet với `--sheet "<tên>" --header-row <dòng> --first-col A --last-col <cột Ghi Chú TDI> --last-row <data_end_row>` để phủ đủ tới Ghi Chú. `verify` phải in ref kết thúc ở cột Ghi Chú.

Quy tắc ref (script tự tính, in ra để kiểm):
- **Dòng lọc** = dòng tiêu đề **thấp nhất** của bảng (`header_rows[-1]` trong config), vì tiêu đề CĐT/TDI có 1–3 tầng gộp ô; nút lọc đặt ở dòng chi tiết cuối, không đặt ở tầng nhóm.
- **Cột đầu** = cột đầu của bộ lọc CĐT nếu CĐT đã có, nếu không là cột đầu có chữ ở dòng tiêu đề. **Cột cuối** = cột cuối có chữ ở dòng tiêu đề.
- **Dòng cuối** = `data_end_row`. **Không** gồm dòng "Tổng cộng trước thuế (TDI)" ngay dưới, để dòng tổng luôn hiện khi lọc.
- **Vùng cố định (freeze)**: nếu sheet đang freeze với `ySplit` nhỏ hơn dòng lọc (ví dụ tiêu đề dòng 4–6 nhưng freeze dừng ở dòng 5) thì `apply` nâng `ySplit` lên đúng dòng lọc và đẩy `topLeftCell` xuống dưới dòng đó nếu cần. Không bao giờ giảm `ySplit`, không tạo freeze mới khi sheet chưa freeze (chỉ in CẢNH BÁO để người dùng quyết). `--no-freeze` để bỏ qua. `verify` báo `LỖI` (exit 1) khi freeze thấp hơn dòng lọc.
- Bộ lọc CĐT có sẵn: ref được mở rộng sang phải, điều kiện lọc đang bật giữ nguyên (`--keep-criteria` nếu muốn giữ cả khi đổi dòng). Không tự xóa lọc của CĐT.
- Cột diễn giải KL đang ẩn (nhóm cột outline): vẫn nằm trong vùng lọc, không đổi trạng thái ẩn.
- Sheet có Table (ListObject) chồng vùng: script **bỏ qua và báo**; không ép (`--force-tables` chỉ khi người dùng đồng ý).
- Công thức tổng dùng `SUBTOTAL(9,…)` nên khi lọc, tổng tự tính theo dòng đang hiện; dòng tiêu đề nhóm (hạng mục) nằm giữa các dòng chi tiết sẽ bị ẩn khi lọc theo cột TDI, đó là hành vi bình thường của Excel (gợi ý người dùng lọc thêm cột `Loại dòng` để giữ tiêu đề).

**Chạy lại khi**:
- `run` ghi đè file `_TDI.xlsx` (mục 2): chạy lại `apply`.
- Sau `boq-gan-ma-gia-thu-vien` thêm 5 cột phụ gán mã bên phải Ghi Chú: chạy lại `apply` để ref phủ tới cột `Loại dòng` (skill đó phải gọi lại bước này sau khi tạo cột phụ).
- **Sheet XD**: chạy `apply --sheet "<tên sheet XD>" --header-row <dòng tiêu đề khối phân tích XD> --last-row <dòng cuối dữ liệu>` (xem dòng tiêu đề và dòng cuối ở `references/xd.md` / `config` của `boq_xd_tools.py`). Không tìm ra thì hỏi người dùng, không đoán.

**Không làm**: không tạo thêm bộ lọc thứ hai trên cùng sheet (Excel không cho), không dùng Table thay AutoFilter (merge ô tiêu đề và SUBTOTAL của template bị vỡ), không sửa tên đầu mục / KL của CĐT.

**Khi `tdi_filter.py` không có** (đã gặp 22/10/2026): không dừng, tự patch XML trên bản sao, không mở-lưu bằng openpyxl. Trong `xl/worksheets/<sheetN>.xml` chèn `<autoFilter ref="<cột đầu><header_rows[-1]>:<cột cuối><data_end_row>"/>` ngay sau `</sheetData>` (trước `<mergeCells>`); trong `xl/workbook.xml` thêm đầu `<definedNames>`: `<definedName name="_xlnm._FilterDatabase" localSheetId="<chỉ số sheet từ 0>" hidden="1">'<tên sheet>'!$A$<từ>:$AW$<đến></definedName>` (đổi vùng cho đúng). Sheet đã có `<autoFilter>` thì sửa ref, không thêm cái thứ hai. **Đồng thời kiểm `<pane … ySplit="N" state="frozen"/>` trong `<sheetViews>`: nếu N nhỏ hơn dòng lọc thì sửa `ySplit` thành dòng lọc** (và `topLeftCell` xuống dưới dòng đó nếu cần). Kiểm bằng openpyxl `ws.auto_filter.ref`.

Trạng thái kiểm: script đã thử trên file giả lập và file thật (23/10 và 26/10/2026), mở được bằng openpyxl và LibreOffice. **Chưa thử trên Excel của phòng**, nên lần đầu giao phải nhắc người dùng mở file kiểm có hiện nút lọc ở mọi cột (kể cả khi cuộn xuống) và Excel không báo lỗi sửa file.

### 3b. Link sống thư viện giá bằng workbook link (tùy chọn, chốt 22/10/2026; đường dẫn tương đối chốt 26/10/2026)

**Khi nào**: người dùng nói "link", "liên kết sống", "workbook link", "giá thư viện đổi thì BOQ đổi theo", "cập nhật theo thư mục thư viện". Không nói thì chỉ nhắc một dòng trong tóm tắt giao file rằng có thể bật link sống. Chỉ cho file MEP (XD dùng `TVGIA`, không áp).

**Nguyên tắc**: một chiều, **file thư viện → ô giá trong bảng `TV`** (giống cách phòng link tay từng ô giá). Sửa giá trong BOQ không ghi ngược vào file thư viện. Cập nhật khi người dùng mở BOQ và chọn Update Values; file thư viện không cần mở.

**Link luôn là đường dẫn TƯƠNG ĐỐI (bắt buộc), và chỉ đúng khi file BOQ nằm ĐÚNG thư mục dùng để tính `--linkrel`.** Hai lỗi đã gặp:
- (26/10/2026, lần 1) Bản `tv_link.py` cũ (`--linkbase`) ghi `Target="file:///C:\Users\<tên user>\..."` tuyệt đối. Máy khác mở file thì Excel báo "CẢNH BÁO BẢO MẬT: cập nhật tự động của các nối kết đã bị vô hiệu hóa" và "KHÔNG THỂ LÀM MỚI".
- (26/10/2026, lần 2, TGC-CT04) Link tương đối đúng nhưng người dùng mở bản nằm ở `Claude outputs\` (thư mục con): Excel báo "Không thể đọc sổ làm việc được liên kết", đường dẫn thử `…/Claude outputs/2.Thu vien gia/…`. Xem **Quy tắc vị trí file và link** ở đầu skill: chỉ giao một bản, ở thư mục đích, và kiểm lại trên máy người dùng.

Hệ quả cho cách tổ chức thư mục:
- Thư viện ME nằm **trong thư mục dự án** (cùng cấp file BOQ, ví dụ `2.Thu vien gia\`) → `--linkrel "2.Thu vien gia"` (tính từ thư mục đích chứa file BOQ). Đây là cách chuẩn.
- Thư viện nằm ở thư mục anh em / cha (ví dụ `..\2.Thu vien gia`) → `--linkrel "../2.Thu vien gia"`.
- Thư viện chỉ có ở thư mục chung ngoài cây thư mục dự án (OneDrive cá nhân, khác ổ): không dùng link tương đối được. **Đề nghị chép bộ file thư viện đã lọc vào `2.Thu vien gia` của thư mục dự án** rồi nối. Nếu người dùng không muốn chép thì không nối link sống, giữ snapshot `TV`, nói rõ lý do. Không ghi cứng `C:\Users\...`.

**Cách làm của Claude** (script `tv_link.py` bản 26/10/2026 và hướng dẫn đầy đủ `link_workbook_thu_vien.md` ở `_Claude_tools` của thư mục thư viện hoặc Project `claude/boq/`; không tự viết lại; **kiểm `python3 tv_link.py link -h` có `--linkrel`; không có thì đó là bản cũ: dừng, lấy bản mới**):
1. Làm Bước 3 như thường: `tv_loader.py load`, `check` = 0/0/0.
2. Chạy `hsloai FILE OUT` (bản `boq_xml_tools.py` 15/10/2026 chiều trở lên, stage từ `_Claude_tools`; cần đủ 2 tham số) để `TV` có cột O–R, rồi `check` lại. Link nối vào cột thô O, P nên bắt buộc có O–R; thiếu thì `tv_link.py` dừng.
3. `python3 tv_link.py link FILE OUT --libdir "<thư mục thư viện ME đã stage, đúng bộ file đã nạp>" --linkrel "<đường dẫn TƯƠNG ĐỐI từ thư mục đích chứa file BOQ tới thư mục thư viện>"` (ví dụ `--linkrel "2.Thu vien gia"`). Không biết đường dẫn tương đối thì đưa hai đường dẫn thật `--boqdir "<thư mục đích>" --linkbase "<thư mục thư viện>"` để script tự tính (cùng ổ đĩa; khác ổ thì script dừng, đề nghị chép thư viện vào cây thư mục dự án). Đường dẫn thật lấy từ `resolvedPath` của `device_list_dir` / `device_stage_files` hoặc hỏi người dùng, **không phải đường dẫn tạm của cầu nối** (`$HOME/mnt/…`) và không phải `/mnt/user-data/outputs/`, không giả định tên user. **Không dùng `--absolute`**. Script nối theo Mã hiệu + Tên công việc + giá, ghi cột cuối `Kiểm dòng nguồn`, liệt kê dòng không nối được (giữ giá tĩnh), tự mã hóa khoảng trắng thành `%20`.
4. **Kiểm đường dẫn link trong file vừa nối** (bắt buộc): `unzip -p OUT 'xl/externalLinks/_rels/*' | grep -o 'Target="[^"]*"'`. Mọi link do Claude tạo phải dạng `<linkrel>/<tên file>.xlsx` (có `%20`), **không có `C:\`, `Users`, `file:///`** (trừ link cũ của CĐT mà Claude không tạo), và tên file giải mã `%20` có tồn tại trong thư mục thư viện. Thấy đường dẫn tuyệt đối thì dừng, không giao.
5. Ghi file bằng XML qua `tv_link.py` (không mở-lưu bằng openpyxl). Chạy `tv_link.py verify FILE --libdir "<thư mục thư viện>"` rồi giao. Không chạy `tv_loader.py check` sau khi đã nối link (check đọc giá trị, không đọc công thức); chạy `check` trước khi nối. **Sau khi commit lên máy người dùng, kiểm lại ở đó theo Quy tắc vị trí file (mục 4 của quy tắc).**
6. Trong tóm tắt giao file nói ngắn: số dòng đã nối / tổng, số dòng không nối; link là **tương đối**, **file phải nằm đúng `<thư mục đích>` cùng cấp `<linkrel>`**, **giữ nguyên cấu trúc thư mục** khi chép / gửi sang máy khác (chép cả thư mục dự án, không chỉ file BOQ); **không mở bản trong `Claude outputs`**; khi mở file bấm Enable Content / Update Values; nếu vẫn báo không làm mới được thì Data → Edit Links (Workbook Links) → Change Source trỏ về thư mục thư viện của máy đó (làm một lần, Excel lưu lại tương đối); thư viện bị **chèn / xóa dòng** thì link lệch và cột `Kiểm dòng nguồn` báo `LỆCH DÒNG` (khi đó nối lại hoặc dùng phương án Power Query `claude/boq/power_query_tv_song.md`); trước khi nộp thầu phải **chốt giá** (Data → Edit Links → Break Link ở bản `_chot`), nếu không giá BOQ đã duyệt đổi theo file thư viện.
7. Nêu rõ: chưa kiểm trên Excel của phòng (mới thử trên thư viện giả lập, LibreOffice và kiểm chuỗi Target), lần đầu mở file phải kiểm có báo lỗi gói file không và có Update Values được không.

**Sửa file đã nối link tuyệt đối / sai thư mục từ trước**: chạy trên bản sao `python3 tv_link.py relink FILE OUT --linkrel "<đường dẫn tương đối tới thư mục thư viện tính từ vị trí file>"` (chỉ đổi đường dẫn link, không đụng ô nào), giao bản `_v2` ở thư mục đích, không ghi đè bản người dùng đang mở. Link cũ của CĐT (Kone / Nissan…) cũng bị đổi nếu cùng tên thư mục: chỉ chạy khi người dùng đồng ý.

**Bốn điểm đã gặp khi nối link (22/10/2026), làm trước / ngay sau `hsloai` và `tv_link.py`:**
1. **File `_TDI.xlsx` do `boq_tdi.py run` + `tv_loader.py` ghi dùng `inlineStr`, không có `xl/sharedStrings.xml`.** `boq_xml_tools.py hsloai` sẽ lỗi `KeyError: 'xl/sharedStrings.xml'`. Cách xử lý: trên bản sao làm việc, thêm `xl/sharedStrings.xml` rỗng (`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>`), thêm `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` vào `[Content_Types].xml` và một `<Relationship Id="rIdSS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` vào `xl/_rels/workbook.xml.rels`, rồi mới chạy `hsloai` và `tv_loader.py check`.
2. **Excel âm thầm bỏ toàn bộ công thức link nếu `Target` trong `xl/externalLinks/_rels/externalLinkN.xml.rels` còn dấu cách chưa mã hóa.** `tv_link.py` (bản 26/10/2026) đã tự mã hóa `%20`, nên không cần sửa tay nữa; chỉ cần kiểm lại sau khi chạy (bước 4). Dấu hiệu lỗi nếu còn: người dùng mở file, bấm Update, lưu lại, và sheet `2 Thu vien gia` mất cột O–S cùng công thức; bảng `TV` co về A:N; số link còn lại chỉ là link cũ. Gặp thì không lưu đè bản người dùng; chạy `tv_link.py relink`, giao bản `_v2`, nhờ chụp bảng "Removed Records / Repaired Records" nếu còn lỗi. Chưa kiểm trên Excel của phòng.
3. File CĐT cũ thường mang theo hàng chục link ngoài hỏng (đã gặp 35 link Kone/Nissan/Vinamilk…, không công thức nào dùng; chỉ nằm trong tên vùng `#REF!`). Chúng làm hộp thoại Update báo lỗi. Nhắc người dùng; chỉ xóa khi họ đồng ý.
4. Thứ tự đúng: `run` (ra thư mục tạm) → soát config → `tv_loader.py load` + `check` → (sharedStrings rỗng nếu thiếu) `hsloai` + `check` → vá cột Đơn vị TDI (nếu cần) → `tv_link.py link --linkrel …` + kiểm Target + `verify` → bộ lọc 3c → `checkcycle` → chép bản cuối vào `outputs/` → commit vào thư mục đích → kiểm lại trên máy người dùng.

**Khi file đã nối link và quay lại skill này / `boq-gan-ma-gia-thu-vien`**:
- Nhận biết: file có `xl/externalLinks/*` do `tv_link.py` tạo, bảng `TV` có cột `Kiểm dòng nguồn`.
- **Không chạy `tv_loader.py load … --replace`** và **không chạy lại `hsloai`** (phá link). Thư viện đổi thì nhờ người dùng Update Values trong Excel.
- Ghi file bằng XML phải giữ nguyên `xl/externalLinks/*`, `<externalReferences>` và cột `Kiểm dòng nguồn`. `tdi_filter.py` ghi bằng XML nên dùng được trên file đã nối link.
- Chạy lại `run` ghi đè `_TDI.xlsx` và làm mất sheet thư viện cùng link và bộ lọc: phải làm lại mục 3, 3b và 3c.

### 4. Kiểm tra trước khi giao

1. **`KẾT LUẬN: ĐẠT`** (đã gồm `4) Vòng tham chiếu: 0`). Nếu in `CHƯA ĐẠT`: (a) vài ô số CĐT khác dạng ghi (ví dụ `167.30999999999997` → `167.31`) hoặc ô chuỗi rỗng `''` thành ô trống: lệch biểu diễn, không phải sai giá trị; (b) vòng tham chiếu nằm trong **ô của CĐT** (ví dụ `A1314` có `MAX($A$12:$A1317)` trỏ vào chính nó): không sửa phần CĐT, ghi vào báo cáo và đưa vào danh sách câu hỏi cho CĐT. Mọi trường hợp khác thì sửa. Ghi rõ lý do trong báo cáo, không coi là lỗi nếu thuộc (a) hoặc (b).
2. **Có sheet `2 Thu vien gia` và bảng `TV`** (trừ khi người dùng bảo không nạp), **cột đầu tên `Loại vật liệu`**, và **`tv_loader.py check` ra THIẾU 0 | THỪA 0 | LỆCH GIÁ 0**.
3. **Không còn sheet tạm.**
   ```bash
   python3 -c "import openpyxl,sys;print([s for s in openpyxl.load_workbook(sys.argv[1],read_only=True).sheetnames if s.startswith('_TDI')])" "<file>"
   ```
   Nếu còn (thường do file đã qua add-in): xóa bằng lệnh `cleanup` của `boq_xml_tools.py` (Phụ lục C skill `boq-gan-ma-gia-thu-vien`), hoặc nhờ add-in xóa. **Không mở-lưu lại bằng openpyxl.**
4. **Hợp đồng cột với skill gán mã** (xem `references/layout_inplace.md`): ô "Nội Dung Công Việc" có ở dòng tiêu đề TDI; bên phải cột "Ghi Chú" của phần TDI không có dữ liệu khác, để chừa chỗ cho 5 cột phụ.
5. **Không có vòng tham chiếu (circular reference) do TDI tạo.** `boq_tdi.py run` đã tự kiểm (dòng `4)`); `build` không còn ghi `SUBTOTAL` cho tiêu đề không có dòng con. Với file đã qua add-in / `boq_xd_tools.py` / sửa tay, chạy:
   ```bash
   python3 $S checkcycle "<file>_TDI.xlsx"     # in số vòng; exit code 1 nếu có
   ```
   Nếu còn vòng ở ô TDI: ô tiêu đề không có dòng con thì **bỏ công thức** (để trống, giữ style) bằng patch XML trên chính file, **không** lưu lại bằng openpyxl; sau đó chạy lại `checkcycle` và `recalc`. LibreOffice `recalc` không báo vòng nên không thay được bước này. Lịch sử lỗi: `references/loi_da_gap.md`.
6. **Bộ lọc phủ mọi cột và nút lọc luôn hiện**: `python3 $F verify "<file>_TDI.xlsx"` in ref từng sheet; ref phải bắt đầu ở cột đầu CĐT, kết thúc ở cột cuối TDI (Ghi Chú, hoặc cột `Loại dòng` nếu đã gán mã), dòng đầu là dòng tiêu đề thấp nhất, dòng cuối là `data_end_row` (không gồm dòng Tổng cộng TDI). Không có cảnh báo `CÓ TABLE`. **Không có dòng `LỖI: freeze chỉ tới dòng …`**; có dòng `CẢNH BÁO: chưa freeze` thì nêu trong tóm tắt giao file.
7. **Cột Đơn vị TDI để trống**: mọi dòng hạng mục của từng sheet MEP có ô Đơn vị TDI rỗng (không công thức `=+G`/`=+H`/`=+R`), tiêu đề "Đơn Vị" vẫn còn; công thức `KL` trỏ cột ĐVT của CĐT (không trỏ ô Đơn vị TDI). Đếm và ghi vào tóm tắt.
8. **Không ghi cứng đường dẫn máy**: nếu file có `xl/externalLinks/*` do Claude tạo thì mọi `Target` là đường dẫn tương đối (không có `C:\`, `Users`, `file:///`); nêu kết quả kiểm trong tóm tắt.
9. **File lành**: mọi phần XML của file cuối parse được (không lỗi `XMLSyntaxError`); `recalc` trên bản sao **không có link** (hoặc LibreOffice) không báo ô lỗi do TDI gây ra (lỗi `#NAME?` ở ô `XLOOKUP` của CĐT là do LibreOffice, bỏ qua).
10. **Vị trí và link trên máy người dùng** (bắt buộc, mục 4 của Quy tắc vị trí file): sau commit, số phần tử zip / kích thước / số link của file ở thư mục đích khớp bản cuối; mỗi file thư viện tồn tại tại `<thư mục đích>/<linkrel>/<tên>`; không còn bản `_TDI.xlsx` cũ khác nội dung ở thư mục đích; không còn bản trung gian ở `Claude outputs`.

### 5. Giao file

**Tóm tắt kèm file** (3–6 dòng):
- Số sheet, số hạng mục, vị trí phần TDI của từng sheet.
- Các dòng còn trống Chủng loại (nêu số dòng, để người dùng xác nhận).
- **Nguồn thư viện**: ME lấy ở đâu (thư mục dự án / thư mục chung / người dùng chỉ), lọc theo mã hay cả thư mục, số file; XD lấy từ file nào, ngày sửa, kết quả kiểm bố cục `TVGIA`.
- Thư viện đã nạp: mã dự án, số `Loại vật liệu` (trên 28), số dòng (trong đó từ sheet ẩn / không mã / thiếu ĐVT), kết quả `check`, ngày cập nhật mới nhất của các file nguồn, khóa trùng, sheet lạ chưa có trong bảng ánh xạ.
- Đơn giá gốc VL/NC và **cột Đơn vị TDI** đang trống, chờ skill `boq-gan-ma-gia-thu-vien` (ĐVT lấy từ thư viện).
- **Bộ lọc**: mỗi sheet đã có bộ lọc phủ cột nào đến cột nào (ref), dòng lọc là dòng mấy, freeze đã tới dòng mấy (và có phải do `apply` nâng lên không); nếu CĐT có bộ lọc sẵn thì nói đã mở rộng; sheet nào bị bỏ qua và lý do.
- Link sống thư viện (mục 3b): bảng `TV` đã nối workbook link hay vẫn là snapshot; số dòng đã nối / không nối; số dòng `LỆCH DÒNG` từ `verify`; link là đường dẫn tương đối `<linkrel>` nào; **file phải nằm ở `<thư mục đích>`**; số link cũ hỏng của CĐT còn lại.
- Cảnh báo shape/textbox nếu `build` báo.
- Kết quả kiểm vòng tham chiếu (số vòng tìm thấy / đã sửa / thuộc CĐT).

**Cách giao theo môi trường**:
- **Cowork**: chép **bản cuối** vào `/mnt/user-data/outputs/` (chỉ bản cuối, không bản trung gian), `SendUserFile`, rồi nếu có thư mục kết nối: `device_commit_files` vào **cùng thư mục với file CĐT gốc** (thư mục đích), tên `<tên gốc>_TDI.xlsx`; không ghi đè file gốc. Báo tên thư mục và tên file. Làm bước kiểm trên máy (mục 4 của Quy tắc vị trí file) trước khi báo xong.
  - **Chỉ một file làm việc** `<tên gốc>_TDI.xlsx` ở thư mục đích. Không tạo thêm bản `_da-gan-ma`, `_ban-truoc…` trên máy người dùng; bản trung gian chỉ ở cloud (thư mục tạm). Ngoại lệ duy nhất: người dùng đã mở và lưu đè bản cũ làm mất link (mục 3b) thì giao `_v2`.
  - Trước khi ghi: `device_list_dir` thư mục đích; nếu có file khóa `~$<tên>_TDI.xlsx` thì xin người dùng đóng file; nếu file `_TDI.xlsx` đã tồn tại thì ghi `expectedMtimeMs` lấy từ `device_list_dir` ngay trước đó và không dùng `force`. Quy trình đầy đủ ở mục "Một file duy nhất: giao thức ghi" của skill `boq-gan-ma-gia-thu-vien`.
  - **Nhắc người dùng**: mở file ở thư mục dự án, không mở bản trong `Claude outputs` (nếu có thì đóng và bỏ).
- **claude.ai**: `present_files`.

**Điểm dừng (Bước 1 quy trình TDI)**:
- Gửi danh sách thông tin thiếu cần hỏi CĐT: STT, đầu mục, vấn đề, câu hỏi đề xuất.
- Các vấn đề cần nêu: dòng thiếu KL/đơn vị, mô tả mơ hồ, mâu thuẫn với bản vẽ/spec, lỗi công thức trong file CĐT (vòng tham chiếu…).
- **Chờ người dùng xác nhận** rồi mới sang gán mã.

## Lưu ý

- **CĐT gộp nhiều hệ trong một sheet**: giữ một sheet. **CĐT tách mỗi hệ một sheet**: mỗi sheet một mục trong `sheets`. Summary nhận tối đa 12 hệ.
- **Mô tả 2 cột Anh/Việt**: cột Việt làm `desc`, cột Anh làm `spec`.
- **"Phụ kiện cho ống…"**: lấy mã ống ngay trên. **"Vật tư phụ"** dạng lot: luôn là VTP.
- **Cột Nội dung phía TDI** là công thức `=B…`: giữ được định dạng cả ô, không giữ rich text từng đoạn.
- **Không sửa tên đầu mục và KL CĐT**: đây là khóa để map giá ngược về BOQ gốc ở Bước 6.
- **BOQ trên ~20.000 dòng**: dùng bản Python thay cho add-in.
- **Add-in Claude for Excel** (không có Python): `tdi_write.js` đang đặt `autoFilter` cho `A12:<cột Ghi Chú>` ở mode rebuild; với inplace trong add-in, đặt `sheet.autoFilter.apply` cho vùng từ cột đầu CĐT đến cột Ghi Chú TDI, dòng tiêu đề thấp nhất đến `data_end_row`, **và đặt freeze rows tới dòng tiêu đề thấp nhất đó nếu freeze hiện tại thấp hơn**, rồi nhắc người dùng vùng đã đặt.
- **TDI đổi mẫu màu**: `python3 scripts/make_style_template.py "<file mẫu>.xlsx" --sheet "<sheet>" --out assets/style_template.xlsx`
- **Sửa luật Chủng loại** trong `boq_tdi.py`: chạy `python3 scripts/build_excel_core.py` để sinh lại JS cho add-in.
- **Lỗi `boq_tdi.py` đã gặp** (tên in `#N/A`, cột Nhãn hiệu nhận nhầm, tiêu đề "1.1" bị nhận là ghi chú, tiêu đề trùng "NỘI DUNG CÔNG VIỆC", "CHƯA ĐẠT" giả do `#REF!` có sẵn, nhận diện sai `zones` / `total_col` / `data_end_row` khi CĐT có cột Khối lượng nhà thầu và Đơn giá bên phải…): xem `references/loi_da_gap.md`.