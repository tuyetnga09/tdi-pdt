---
name: "boq-ghi-gia-ncc"
description: "Ghi giá NCC đã duyệt (ASSIGN.json từ boq-lay-gia-ncc) vào BOQ TDI, thay cả ô giá tham khảo web; ghi nguồn, audit, đẩy sang cột CĐT khi được đồng ý."
---

# Ghi giá NCC đã duyệt vào BOQ TDI

Skill này CHỈ ghi giá. Việc đọc báo giá trong BG NCC, dựng sổ cái, khớp đầu mục và để người dùng duyệt là của `boq-lay-gia-ncc`; skill này nhận kết quả đã duyệt của skill đó (`ASSIGN.json` + sổ cái `*_SoCai.xlsx`) và ghi vào BOQ. Không đọc báo giá, không khớp lại, không gửi gì ra ngoài.

Vị trí trong quy trình: `covert-boq-cdt-into-boq-tdi` → `boq-gan-ma-gia-thu-vien` → `boq-gia-tham-khao-web` → `boq-tao-ycbg` → (người gửi YCBG, Bước 5) → `boq-lay-gia-ncc` (quét BG NCC, rà soát, duyệt, ra ASSIGN.json) → **`boq-ghi-gia-ncc`** (skill này) → gửi/nộp. Ứng với Bước 6 quy trình TDI. Dùng được cho mọi dự án và mọi sheet BOQ TDI (nhận cột theo tiêu đề). Thiếu `ASSIGN.json` đã duyệt: dừng, hướng người dùng chạy `boq-lay-gia-ncc` trước, không tự dựng sổ cái hay tự khớp.

## Tên file

- `<tên>_TDI.xlsx`: BOQ làm việc của dự án (kết quả của gán mã, giá tham khảo web). Skill này KHÔNG sửa file này; nó là bản "trước" để `audit`.
- `<tên>_NCC.xlsx`: file kết quả của skill này (BOQ TDI đã ghi giá NCC; nếu người dùng đồng ý đẩy sang cột CĐT thì cột CĐT cũng nằm trong file này). Lần chạy bổ sung khi có thêm báo giá: lấy `<tên>_NCC.xlsx` làm đầu vào, ghi ra `<tên>_NCC_<yyyymmdd>.xlsx`, `audit` so hai file đó.
- `BAO_CAO_GHI_GIA_NCC_<ngày>.xlsx`: báo cáo, đặt trong thư mục `_Claude` (hoặc Claude outputs) cạnh thư mục báo giá. Không cần bản sao `_ban-truoc-NCC`.
- Không ghi vào thư mục thư viện giá, không đụng file trong thư mục BG NCC.

## Nguyên tắc bất biến (sai ở đây là sai giá dự thầu)

1. **Chỉ có 3 trường hợp được ghi vào ô Đơn giá gốc VL hoặc NC:**
   - **Ô TRỐNG** hoàn toàn: ghi số.
   - **Ô có CÔNG THỨC cho kết quả 0 hoặc rỗng** (thư viện có mã nhưng chưa có giá): giữ công thức, bọc `=IF(N((cũ))=0, giá NCC, (cũ))`. Khi thư viện có giá lớn hơn 0 thì giá thư viện luôn thắng.
   - **Ô VL đang là giá tham khảo web** do `boq-gia-tham-khao-web` ghi: nhận ra khi đủ cả 3 dấu hiệu (ô số cứng lớn hơn 0; cột `Nguồn giá` = `Giá tham khảo`; ghi chú gán mã có `GIÁ THAM KHẢO WEB`). Giá NCC luôn thay giá tham khảo web (kể cả khi NCC cao hơn hoặc thấp hơn): ghi giá NCC, nền xanh chuyển sang cam, `Nguồn giá` = `NCC`, ghi chú đổi thành `GIÁ WEB CŨ (đã thay bằng giá NCC): ...` giữ nguyên số liệu cũ để truy vết. Thiếu một trong 3 dấu hiệu thì coi là giá người dùng, không ghi, báo lý do.
   Ô đã có giá lớn hơn 0 khác (công thức thư viện, công thức nhập tay như `250000*1.15`, số nhập tay) TUYỆT ĐỐI không ghi, kể cả khi NCC rẻ hơn hay mới hơn. Số 0 gõ tay: bỏ qua, trừ khi người dùng xác nhận (`--ghi-so-0`). Ô chứa chữ, hoặc công thức chưa có kết quả tính: bỏ qua.
2. **Dòng đã có giá một thành phần** (ví dụ có NC, thiếu VL): không tự điền nốt. Chỉ điền thành phần còn thiếu khi ASSIGN.json có giá của đúng thành phần đó (độ tin cậy Cao hoặc Trung bình). Thành phần đã có giữ nguyên tuyệt đối; không suy từ thành phần kia hay từ dòng khác.
3. Dòng đã có giá thư viện/nhập tay cho CÙNG thành phần mà cũng có báo giá NCC: giữ giá hiện hữu, đưa báo giá NCC vào mục `ThamKhao` của báo cáo, không ghi vào BOQ.
4. **Không sửa dữ liệu CĐT** (STT, tên đầu mục, ĐVT, KL), không sửa thư viện giá, không tạo đầu mục mới.
5. **Không bịa giá.** Mỗi giá ghi vào BOQ truy được về một dòng trong sổ cái và file báo giá gốc. Không suy giá từ dòng tương tự, không nội suy, không quy đổi tiền tệ khi chưa có tỷ giá người dùng chốt.
5b. **Bắt buộc ghi nguồn từng giá.** Mỗi mục trong ASSIGN.json phải có `ncc`, `file`, `date` (YYYY-MM-DD hoặc dd/mm/yyyy), `ledger_id` của từng thành phần, và `note` mở đầu bằng vị trí trong file báo giá (`<sheet NCC>!dòng <n>; STT YCBG <x>; đầu mục NCC: ...`, PDF: `trang <n>!dòng <n>`, email: `email <ngày>`). Script `apply` từ chối ghi mục thiếu một trong các trường này và in lý do.
6. **Dòng lô / lot / bộ / HT / gói / tấn / 10m / 100m**: không gán giá trực tiếp. Phải có diễn giải chiết tính (`calc`); không có thì đưa vào `CanChietTinh`.
7. **Không gửi gì ra ngoài** (YCBG, email, tin nhắn).
8. **Không lưu đè `<tên>_TDI.xlsx`** hay bản CĐT phát hành. Nhắc người dùng đóng bản đang mở rồi mở lại, không lưu bản cũ.
9. Mã dự án của BOQ, thư viện, sổ cái hoặc báo giá không khớp dự án đang làm: dừng và hỏi.
10. Chọn NCC, hệ số, giá chốt là việc của người rà soát và Trưởng phòng.

## Công cụ

- `claude/boq/boq_ncc_tools.py` trong Project "Tối ưu hóa quy trình Phòng đấu thầu" (tự chứa, cần python3, openpyxl, lxml). Lấy bằng `project_read`, chép ra thư mục tạm. Ghi file BOQ CHỈ qua script này, không dùng openpyxl ghi lại BOQ. Khi BOQ nằm trong thư mục kết nối từ máy người dùng: ưu tiên chạy bằng `device_bash` nếu có; nếu không thì stage đúng file cần dùng, xử lý, rồi `device_commit_files` về cạnh file gốc (commit với `expectedMtimeMs`).
- `boq_xml_tools.py` (`fillcdt`, `recalc`) và `CFG.json` của Bước 8 skill `boq-gan-ma-gia-thu-vien`, dùng ở Bước D và kiểm tra tính lại.

| Lệnh | Dùng khi |
|---|---|
| `sheets BOQ.xlsx` | Phân loại mọi sheet: sheet nào là BOQ TDI, cột đã nhận, có cột phụ gán mã và cột `Nguồn giá` chưa, thống kê trạng thái ô giá (kể cả `VL_web_ref` = đang là giá tham khảo web) |
| `todo BOQ.xlsx [SHEET...] --out todo.json` | Dòng có thể nhận giá NCC: ô trống, công thức ra 0, hoặc đang là giá tham khảo web; kèm KL, ĐVT, cờ lô đặc biệt, cờ một thành phần |
| `ledger-check SO_CAI.xlsx` | Kiểm tra sổ cái: thiếu trường, VAT không rõ, hết hạn, trích từ ảnh/email, trùng NCC và hàng |
| `apply BOQ.xlsx OUT.xlsx ASSIGN.json [--dry]` | Ghi giá NCC, 4 cột nguồn, cập nhật `Nguồn giá`. Luôn `--dry` trước |
| `audit BEFORE.xlsx AFTER.xlsx` | Bằng chứng: mọi ô thay đổi chỉ thuộc 3 trường hợp được phép, 4 cột nguồn NCC, và `Nguồn giá`/ghi chú gán mã của đúng dòng đã ghi. FAIL thì không giao |

Thêm `--map MAP.json` khi một sheet có bố cục lạ (mẫu trong phần đầu script). Sheet chưa phải template TDI: chuyển template bằng `covert-boq-cdt-into-boq-tdi` trước; chỉ dùng `--map` khi người dùng yêu cầu và nói rõ cột giá nào được ghi.

`apply` tự: tạo 4 cột `Nguồn giá NCC | Ngày BG NCC | Trạng thái giá NCC | Chiết tính NCC` bên phải cột có dữ liệu xa nhất nếu chưa có; tô ô giá lấy từ NCC màu cam nhạt; cập nhật `Nguồn giá` (`NCC`; `Giá thư viện + NCC` hoặc `Giá tham khảo + NCC` khi dòng chỉ được thay một phần); bỏ qua dòng Cụm (ghi vào dòng con), dòng lệch `desc_prefix`, ĐVT đặc biệt thiếu `calc`, mục thiếu nguồn, và từng ô đã có giá. Mọi mục bị bỏ qua được in kèm lý do.

Định dạng `ASSIGN.json` (do `boq-lay-gia-ncc assign` tạo): mỗi mục `{"sheet","row","desc_prefix","vl":{"price","ledger_id"},"nc":{...},"ncc","date","file","status","calc","note"}`; `vl`/`nc` tùy chọn. Dòng đã có một thành phần thì chỉ ghi thành phần còn thiếu.

## Quy trình

### Bước A. Điều kiện và hỏi một lần

1. Xác định BOQ đang làm (`<tên>_TDI.xlsx`; hỏi bản nào nếu thư mục có nhiều phiên bản), mã dự án, `ASSIGN.json` đã duyệt, sổ cái `*_SoCai.xlsx`. Thiếu `ASSIGN.json` hoặc sổ cái: dừng (xem phần đầu).
2. Chạy `sheets`. Báo sheet BOQ nhận được, sheet bị bỏ qua và lý do. Nếu báo cột phụ gán mã hoặc cột `Nguồn giá` "CHƯA CÓ": skill `boq-gan-ma-gia-thu-vien` chưa chạy đủ cho BOQ này, dừng và báo người dùng.
3. Chạy `ledger-check` trên sổ cái, báo kết quả (VAT không rõ, hết hiệu lực, ảnh/email, trùng).
4. Chạy `todo`, báo theo từng sheet: số dòng nhận giá (VL, NC, cả hai), số dòng đang là giá tham khảo web sẽ bị thay, số dòng một thành phần, số dòng ĐVT đặc biệt, số dòng công thức chưa có kết quả tính (nếu có: nhờ người dùng mở file bằng Excel, tính lại, lưu, rồi chạy lại).
5. Hỏi một lần nếu chưa có trong hội thoại hoặc Project: (a) giá thư viện tính chưa VAT hay đã VAT; (b) giá NCC trong ASSIGN đã chuẩn hóa theo cơ sở đó chưa (`boq-lay-gia-ncc` giữ giá nguyên văn, trạng thái "Cần kiểm tra"); (c) giá NCC nhập vào cột gốc (hệ số HSVL/HSTB/HSNC áp sau) hay đã là giá chào; (d) chiết khấu chỉ tính khi NCC xác nhận bằng văn bản hay tính cả chiết khấu thương thảo. Cơ sở VAT chưa khớp thì dừng, nhờ người duyệt chuẩn hóa trong Bảng rà soát và chạy lại `assign`; skill này không tự quy đổi VAT.

### Bước B. Kiểm tra ASSIGN.json trước khi ghi

Duyệt từng mục, loại khỏi ASSIGN và đưa vào `RaSoat` (kèm lý do) nếu:
- thiếu trường nguồn theo nguyên tắc 5b;
- `ledger_id` không có trong sổ cái, hoặc giá khác giá trong sổ cái mà không có `calc` giải thích;
- trạng thái "Hết hạn", hoặc báo giá hết hiệu lực trong sổ cái;
- hãng/model NCC chào khác yêu cầu spec hoặc danh sách hãng khóa dự án (Grundfos, Daikin, York, Schneider…);
- ĐVT đặc biệt (lô, bộ, HT, tấn, 10m, 100m…) mà thiếu `calc`, hoặc quy đổi ĐVT không có đúng một hệ số rõ ràng;
- dòng là Cụm (phải ghi vào dòng con);
- báo giá chỉ có VL hoặc chỉ có NC mà ASSIGN ghi cả hai (chỉ ghi thành phần có trong báo giá).
Dòng đã có giá thư viện/nhập tay cho cùng thành phần: chuyển sang `ThamKhao`, không đưa vào ASSIGN. Dòng đang là giá tham khảo web và có báo giá NCC: đưa vào ghi (trường hợp 3).

Điểm dừng: báo số mục sẽ ghi, số mục bị loại và lý do, các mặc định đã áp; chờ xác nhận nếu người dùng chưa cho chạy liền.

### Bước C. Ghi vào BOQ và kiểm chứng

1. `apply <tên>_TDI.xlsx <tên>_NCC.xlsx ASSIGN.json --dry`; đọc danh sách mục bị bỏ qua và lý do; xử lý nếu có lý do bất thường. Chờ OK nếu người dùng chưa cho chạy liền.
2. `apply` thật ra `<tên>_NCC.xlsx`.
3. **`audit <tên>_TDI.xlsx <tên>_NCC.xlsx`** phải PASS. FAIL thì dừng, không giao, tìm nguyên nhân.
4. Mở lại vài dòng: 4 cột nguồn hiện đúng NCC, tên file, ngày, ID sổ cái; cột `Nguồn giá` = `NCC`; ô giá nền cam; dòng thay giá web có ghi chú `GIÁ WEB CŨ`. Script không ghi đủ thì dừng, báo người dùng, không giao file.
5. Tính lại bằng LibreOffice trên bản sao (`recalc`, xóa cache rồi tính lại): không phát sinh ô lỗi; tổng các dòng giá thư viện/nhập tay trước và sau bằng nhau; chỉ các dòng mới có giá hoặc được thay giá web mới đổi, báo chênh lệch tổng của nhóm dòng thay giá web.
6. Kiểm tra AutoFilter vẫn phủ cột `Nguồn giá` (thiếu thì nới theo Bước 6c skill `boq-gan-ma-gia-thu-vien`).
7. Không giao file chưa qua `audit`.

### Bước D. Đẩy giá sang cột CĐT (chỉ khi người dùng đồng ý)

Quy trình TDI yêu cầu "BOQ CĐT hoàn chỉnh có giá" nên bước này là phần của Bước 6, nhưng ghi vào cột CĐT là việc có hậu quả, nên hỏi trước. Sau khi người dùng đồng ý: chạy `fillcdt` (Bước 8 skill `boq-gan-ma-gia-thu-vien`, cùng `CFG.json`, tìm trong thư mục `_Claude`; không có thì hỏi người dùng hoặc dựng theo mô tả Bước 8) trên `<tên>_NCC.xlsx` (ghi ra file tạm rồi thay `<tên>_NCC.xlsx`; `audit` đã chạy ở Bước C, không chạy lại sau `fillcdt`). Báo mọi ô CĐT bị ghi đè (nếu có) và số dòng, tổng thành tiền CĐT theo từng nguồn giá (Giá thư viện / Giá tham khảo / NCC). Nhắc: khi chốt hồ sơ thì copy → Paste Values.

### Bước E. Báo cáo và bàn giao

File `BAO_CAO_GHI_GIA_NCC_<ngày>.xlsx` (openpyxl) gồm các sheet:
- `DaGhi`: sheet, dòng, tên đầu mục BOQ, thành phần ghi (VL/NC), giá mới, giá bị thay (giá tham khảo web cũ, nếu có), NCC, ngày, file, vị trí trong file, ID sổ cái, độ tin cậy, trạng thái.
- `BoQua`: mục ASSIGN bị loại hoặc bị `apply` bỏ qua, kèm lý do.
- `ThamKhao`: dòng đã có giá thư viện/nhập tay nhưng cũng có báo giá NCC cho cùng thành phần (không ghi).
- `CanChietTinh`: ĐVT đặc biệt, kèm đề xuất cách chiết tính.
- `BangNguonGia`: mọi dòng công việc: sheet, dòng, STT, tên đầu mục, ĐVT, KL, đơn giá VL, đơn giá NC, `Nguồn giá` (Giá thư viện / Giá tham khảo / NCC / Chưa có giá), NCC và ngày nếu có. Đây là "bảng nguồn giá từng dòng" của Bước 6.
- `ConThieuGia`: chạy lại `todo` trên file sau khi ghi; dòng vẫn trống nhóm theo NCC/hãng và theo hệ, kèm trạng thái YCBG nếu có bảng theo dõi (đã gửi chưa, hạn phản hồi, NCC đã trả lời chưa). Dòng còn đang dùng giá tham khảo web (chưa có NCC thay) ghi riêng một cột. Hãng khóa dự án xin báo giá sớm.
- `RaSoat`: dòng cần người kiểm tra, sắp theo độ tin cậy Thấp/Cần kiểm tra, rồi KL lớn, rồi giá trị lớn (KL × giá).

Tóm tắt chat (ngắn): số mục ASSIGN; số dòng đã ghi (trong đó số ô bọc IF, số ô thay giá tham khảo web, số dòng chỉ điền thành phần còn thiếu); số mục bị loại; số dòng còn thiếu giá; số dòng còn dùng giá tham khảo web; kết quả `audit`; đã/chưa đẩy sang cột CĐT; các mặc định đã áp; việc người dùng cần làm tiếp.

**Nhắc người dùng**: đóng và mở lại file BOQ, không lưu bản cũ; Trưởng phòng là người kiểm tra cuối và quyết định hệ số chốt dự toán; kiểm tra lại dòng khối lượng lớn, giá trị lớn, ĐVT đặc biệt trước khi nộp. Giá NCC đã duyệt có thể đề xuất đưa vào thư viện giá, nhưng chỉ khi người dùng yêu cầu và do người phụ trách thư viện cập nhật; skill này không tự sửa thư viện.

## Điểm dừng chờ xác nhận

1. Sau Bước A: xác nhận các mặc định (VAT, cột gốc, chiết khấu), và nếu BOQ chưa gán mã/chưa có cột `Nguồn giá` thì dừng.
2. Sau Bước B và sau `apply --dry`: báo số dòng sắp ghi và mục bị loại/bỏ qua, chờ OK nếu người dùng chưa cho chạy liền.
3. Trước Bước D: chờ đồng ý đẩy sang cột CĐT.
4. Kết thúc: dừng. Không gửi YCBG, không gửi email.

## Lỗi dễ gặp

- Giá báo theo m nhưng BOQ tính 10m hoặc 100m: sai gấp 10 hoặc 100 lần nếu gán trực tiếp; bắt buộc `calc`.
- NCC báo "đã gồm VAT" nhưng giá hiện hữu chưa VAT: lệch 8 đến 10% nếu ASSIGN chưa chuẩn hóa (Bước A.5).
- Báo giá thiết bị chỉ gồm cung cấp, chưa gồm lắp đặt và vận chuyển: không điền vào cột NC.
- Giá web bị thay mà người dùng muốn giữ giá web: chỉ ghi khi mục đó có trong ASSIGN đã duyệt; muốn giữ thì bỏ mục khỏi ASSIGN.
- Ô giá tham khảo web bị người dùng sửa tay hoặc mất ghi chú/`Nguồn giá`: script coi là giá người dùng và bỏ qua; báo để người dùng quyết định.
- Công thức dùng chung (shared formula) của Excel: script tự bung đúng; không tự sửa công thức bằng tay.
- File đang mở bị khóa hoặc lưu đè: ghi ra file mới, nhắc đóng và mở lại.
- Giá ghi vào BOQ mà cột nguồn trống hoặc chỉ có tên NCC: coi là lỗi, không giao.