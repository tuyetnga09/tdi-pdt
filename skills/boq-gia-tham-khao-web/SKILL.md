---
name: "boq-gia-tham-khao-web"
description: "Tra giá tham khảo trên web (song song theo lô) cho đầu mục Chưa khớp của BOQ TDI, ghi vào cột Đơn giá gốc VL tô xanh, Nguồn giá = Giá tham khảo, cập nhật file tổng hợp. Chạy sau boq-gan-ma-gia-thu-vien, trước boq-tao-ycbg."
---

# Giá tham khảo từ web cho đầu mục chưa có trong thư viện

Tách từ Bước 6c và 6d của `boq-gan-ma-gia-thu-vien` (chốt 19/10/2026) để bước gán mã chạy nhanh. Vị trí trong quy trình TDI:

`covert-boq-cdt-into-boq-tdi` → `boq-gan-ma-gia-thu-vien` (xuất file tổng hợp) → **skill này** → `boq-tao-ycbg` → Bước 5 (người rà soát, gửi YCBG) → `boq-lay-gia-ncc` + gen giá NCC vào BOQ → trưởng phòng chốt.

## Việc skill này làm và không làm

- **Làm**: tra web cho từng đầu mục Chưa khớp, ghi giá tham khảo vào BOQ TDI (cột Đơn giá gốc VL) và vào file tổng hợp, đưa giá sang cột giá CĐT cho các dòng đó.
- **Không làm**: gán mã, gán giá thư viện, bóc tách, sửa tên đầu mục hay KL của CĐT, soạn YCBG, gửi gì ra ngoài. Không lưu giá web vào Bảng khớp (giá web đổi theo thời gian).

## Quy tắc bắt buộc

1. **Không bịa giá.** Chỉ ghi giá đọc được trên trang web, đúng đầu mục (đúng loại, cỡ, thông số chính) và đúng ĐVT (hoặc quy đổi 1:1 rõ ràng). Không lấy giá cỡ / model khác, không trung bình, không tự ước, không trạng thái "Ước tính". Không thấy thì để trống và ghi đã tra gì.
2. **Truy vấn web chỉ chứa tên vật tư và thông số** (kèm hãng nếu BOQ ghi hãng). **Không** đưa tên dự án, tên CĐT, mã dự án, khối lượng vào truy vấn hay vào bất kỳ chỗ nào ngoài file nội bộ.
3. **Chỉ tra các đầu mục Chưa khớp** có trong sheet `Chưa có trong TV` của file tổng hợp. Không tra: Cần chiết tính; dòng cha cụm (giá = SUMPRODUCT con); thu hồi / hoàn trả chỉ NC; hạng mục DC chỉ NC; PS; dòng "CĐT yêu cầu không báo giá" / "đã tính ở …"; dòng có đầu mục thư viện nhưng giá = 0 (Quy tắc 11 của skill gán mã); mã trùng chờ chọn tay. Các nhóm này đã nằm ngoài sheet `Chưa có trong TV`, nên chỉ cần đọc sheet đó, không tự lọc lại BOQ.
4. **Dòng vẫn là Chưa khớp**: không gán mã, không đổi độ tin cậy, vẫn vào YCBG. Giá tham khảo chỉ để định hướng, thay bằng báo giá NCC khi có; hệ số và giá chốt là việc của trưởng phòng.
5. **Không sửa dữ liệu CĐT** (STT, tên, thông số, ĐVT, KL, dòng tổng). Không động vào ô giá người dùng tự nhập (ô số cứng không có ghi chú `GIÁ THAM KHẢO WEB`).
6. **Một file làm việc duy nhất** `<tên BOQ>_TDI.xlsx`, ghi theo "Một file duy nhất: giao thức ghi" của `boq-gan-ma-gia-thu-vien` (file trung gian ở cloud, bản cuối commit một lần với `expectedMtimeMs`; không dùng openpyxl ghi lại BOQ, ghi bằng lxml giữ style). File tổng hợp là file riêng, cập nhật tại chỗ; nếu mtime khác lần ghi trước (người dùng đã sửa tay) thì không ghi đè, giao `…_lan-N.xlsx` và nói file nào đúng.
7. **Mã dự án** của file tổng hợp (dòng 2 sheet `Chưa có trong TV`) phải khớp file BOQ TDI và dự án đang làm. Không khớp thì dừng và hỏi.
8. **Không gửi gì ra ngoài.** Skill này chỉ ghi file nội bộ.

## Công cụ

- Script ở `_Claude_tools` (hoặc Project `claude/boq/`), lấy theo mục "Công cụ" của `boq-gan-ma-gia-thu-vien`: `boq_xml_tools.py` (`cols`, `fillcdt`, `recalc`). Chưa có lệnh ghi giá tham khảo: ghi bằng lxml theo Bước 3 (ngoại lệ có chủ đích của "chỉ ghi qua script").
- `WebSearch` ở chế độ `standard`; chỉ chuyển `extended` cho đầu mục khó (thiết bị đặc thù, hãng ít gặp) khi `standard` không ra giá.
- Subagent (công cụ Agent) để tra song song. Subagent chỉ tra web và trả JSON, **không ghi file**; chỉ phiên chính ghi file.
- Lưu `REFPRICE.json` ở cloud: `[{"key","sheet","row","query","rank","url","site","price","unit","vat","date","hang_model","khoang_gia","ghi_chu"}]`. `key` = khóa gom (tên chuẩn hóa + thông số chính + ĐVT) giống file tổng hợp.

## Bước 0. Điều kiện

1. Có file BOQ TDI (đã qua `boq-gan-ma-gia-thu-vien`, có cột `Nguồn giá`) và file `<tên BOQ>_Dau_muc_chua_co_trong_thu_vien_gia_tham_khao.xlsx` trong `Claude outputs`. Thiếu một trong hai thì dừng, báo người dùng chạy `boq-gan-ma-gia-thu-vien` trước.
2. Kiểm mã dự án (Quy tắc 7), file BOQ không bị khóa (`~$` cạnh file), `cols` nhận đúng cột TDI từng sheet.
3. Đọc sheet `Chưa có trong TV` thành danh sách đầu mục (ID, Nhóm, Mô tả, ĐVT, Tổng KL, Vị trí trong BOQ, lý do chưa khớp). Đầu mục đã có giá ở cột H (lần chạy trước) thì **bỏ qua**, trừ khi người dùng bảo tra lại; đầu mục `Chưa tìm được giá` chỉ tra lại khi người dùng yêu cầu.
4. **Cache** (nếu có): `REFPRICE_CACHE.json` ở `_Claude_tools` hoặc `claude/boq/refprice_cache.json` trong Project. Đầu mục cùng `key` đã tra trong vòng 30 ngày thì dùng lại, ghi **ngày tra gốc** và "(cache)" vào ghi chú; quá 30 ngày thì tra lại. Không có cache thì bỏ qua bước này.

## Bước 1. Lập kế hoạch tra (báo ngắn, không chờ duyệt)

- Số đầu mục duy nhất cần tra, số dòng BOQ ảnh hưởng, số đầu mục lấy từ cache, số lô.
- **Thứ tự ưu tiên**: hãng khóa dự án (Grundfos, Daikin, York, Schneider…) và thiết bị chính trước, rồi theo `Tổng KL` giảm dần, rồi còn lại.
- **Chia lô** 15–20 đầu mục, giữ nguyên thứ tự ưu tiên. Mỗi lô giao một subagent; chạy nhiều subagent cùng lúc (tối đa khoảng 5 lô song song để tránh quá tải), lô sau chạy khi lô trước xong.
- **Giới hạn mỗi đầu mục**: tối đa 3 truy vấn và 2 lần mở trang. Hết giới hạn mà chưa thấy giá thì ghi "không thấy giá", **không bỏ đầu mục nào** vì nhiều. Nếu người dùng đặt giới hạn thời gian mà chưa tra hết thì báo còn bao nhiêu đầu mục, và ghi rõ đầu mục chưa tra.

## Bước 2. Tra web song song (subagent)

Mẫu prompt cho mỗi subagent (điền danh sách đầu mục của lô):

> Bạn tra giá tham khảo trên web cho các đầu mục vật tư sau (mỗi đầu mục gồm key, tên, thông số, ĐVT, hãng nếu có). Chỉ dùng WebSearch / WebFetch. Không ghi file. Với mỗi đầu mục: (1) truy vấn = tên vật tư tiếng Việt + thông số chính + "giá" (thêm hãng nếu có); không thêm tên dự án, chủ đầu tư, khối lượng. (2) Lấy **kết quả đầu tiên** trong danh sách tìm kiếm có ghi giá cụ thể cho đúng đầu mục (đúng loại, đúng cỡ / thông số chính) và đúng ĐVT hoặc quy đổi 1:1 rõ ràng; kết quả đầu không thỏa thì lấy kết quả kế tiếp đầu tiên thỏa và ghi thứ hạng `rank`. Snippet chưa đủ thì mở trang. (3) Không lấy giá của cỡ / model khác, không lấy trung bình, không tự ước. (4) Giá ghi nguyên số trên trang (VNĐ / 1 ĐVT); trang ghi khoảng giá thì lấy số đầu và ghi cả khoảng; giá ngoại tệ thì không quy đổi, để `price` rỗng và ghi giá gốc trong `ghi_chu`; VAT ghi `đã VAT` / `chưa VAT` / `không rõ`. (5) Tối đa 3 truy vấn và 2 lần mở trang mỗi đầu mục; hết mà không thấy giá thì `price` rỗng và `ghi_chu` = "Đã tra web: <truy vấn>, không thấy giá". Trả về **chỉ một mảng JSON**, mỗi phần tử: `{"key","query","rank","url","site","price","unit","vat","hang_model","khoang_gia","ghi_chu"}`. Không thêm lời giải thích.

Phiên chính nhận JSON từng lô, **kiểm lại trước khi nhận**: mỗi `key` trong lô có đúng một phần tử; mỗi `price` có `url` và `rank`; `unit` khớp ĐVT BOQ; truy vấn không chứa tên dự án / CĐT / KL; giá không có dấu hiệu lệch cỡ (ví dụ giá cho DN100 đầu mục DN50). Phần tử nào không đạt thì trả lại tra riêng hoặc để trống. Ghép tất cả vào `REFPRICE.json` (thêm `sheet`, `row` từ cột Vị trí; một đầu mục nhiều dòng thì ghi mỗi dòng một bản), gắn `date` = ngày tra dd/mm/yyyy.

## Bước 3. Ghi vào BOQ TDI

Cho mọi sheet đã xử lý (MEP và XD), mọi dòng có giá trong `REFPRICE.json`:

- **Ô Đơn giá gốc VL** của dòng (MEP: cột `Giá VL gốc` theo `cols`; XD: cột Đơn giá VL của khối Công tác VL): ghi **số cứng**. Ô NC để trống, ghi chú "NC chưa có, cần báo giá / chiết tính".
- **Tô nền ô đó xanh lá nhạt `E2EFDA`**: thêm `fill` mới và một `xf` mới chép từ `xf` cũ của ô (giữ font, viền, định dạng số), chỉ đổi `fillId`. Chỉ tô ô giá VL, không tô phần CĐT; dòng con nền `FFF2CC` chỉ đổi ô giá VL sang xanh.
- Mã hiệu để trống, Nhóm VT giữ nguyên, **Độ tin cậy giữ `Chưa khớp`**.
- **Ghi chú gán mã** thêm: `GIÁ THAM KHẢO WEB: <giá> đ/<ĐVT>, <tên trang / domain>, <URL>, kết quả #<n>, tra ngày dd/mm/yyyy, <VAT>; cần báo giá NCC thay thế`.
- **Cột `Nguồn giá`** của dòng = `Giá tham khảo`. Dòng cha cụm: `Giá tham khảo` nếu có ít nhất một con giá tham khảo. Dòng không thấy giá: giữ `Chưa có giá`. AutoFilter đã phủ cột `Nguồn giá` từ `boq-gan-ma-gia-thu-vien`; kiểm lại, thiếu thì nới theo mục "Bộ lọc" ở Bước 6c của skill đó.
- Đơn giá sau hệ số, thành tiền tự tính từ ô VL như dòng thường; giá tham khảo **không nhân hệ số Loại VL** (không có trong `TV`), vẫn qua HSVL / HS chung.
- **Chạy lại**: dòng có ghi chú `GIÁ THAM KHẢO WEB` mà nay đã khớp thư viện thì bỏ số cứng, ghi công thức thư viện qua `write`, trả style cũ, `Nguồn giá` = `Giá thư viện`. Ô số cứng **không** có ghi chú này vẫn là giá người dùng, không động vào.

## Bước 4. Đưa giá sang cột giá CĐT

Chạy lại `fillcdt` (Bước 8 của `boq-gan-ma-gia-thu-vien`, cùng `CFG.json`: cột theo tiêu đề từng sheet, bỏ qua dòng `Con` và dòng không đơn vị) để các dòng `Giá tham khảo` có giá ở cột giá CĐT. Báo mọi ô CĐT bị ghi đè (nếu có) và số dòng, tổng thành tiền CĐT đang dùng giá tham khảo. Nhắc khi chốt hồ sơ thì copy → Paste Values.

## Bước 5. Cập nhật file tổng hợp

Mở file tổng hợp (bản `recalc` của file TDI ở Bước 6 làm căn cứ đối chiếu), ghi vào sheet `Chưa có trong TV` theo **tiêu đề cột**, không theo chữ cái:

- `Giá tham khảo VL (VNĐ/ĐVT)`: số cứng, nền `E2EFDA`. `Thành tiền tham khảo` giữ công thức `=IF(OR(H4="",E4=""),"",E4*H4)`.
- `Khoảng giá`, `Hãng / model tham chiếu`, `VAT` (`đã VAT` / `chưa VAT` / `không rõ`), `Nguồn` (tên trang), `Link nguồn` (URL), `Ghi chú giá` (điều kiện của giá; dòng không thấy giá: `Đã tra web: <truy vấn>, không thấy giá`), `Thứ hạng KQ` (`#n`), `Ngày tra`, `Truy vấn`.
- `Tình trạng giá` chỉ một trong: `Giá niêm yết` (trang hãng / NCC công bố bảng giá), `Giá bán lẻ/đại lý` (đại lý, cửa hàng, sàn), `Chưa tìm được giá – cần báo giá NCC`. Dòng tủ điện (cụm) giữ `Xem sheet "TĐ – tổng hợp"`; sheet `TĐ – linh kiện`: giá lấy `INDEX/MATCH` theo ID như skill gán mã mô tả.
- Sheet `Tổng hợp`: cập nhật A2 phần "Tra giá web ngày <ngày>"; các ô `COUNTIFS` / `SUMIFS` sẵn có tự tính lại (`Có giá tham khảo`, `Chưa có giá = Số đầu mục − Có giá`, `Thành tiền tham khảo`). Không có cột "Ước tính".
- Không sửa cột A–G và sheet phụ khác. Chạy `recalc` trên bản sao: 0 ô lỗi. Commit về `Claude outputs` với `expectedMtimeMs` (Quy tắc 6).

## Bước 6. Kiểm tra trước khi báo xong

1. `recalc` file TDI: 0 ô lỗi, tham chiếu vòng = 0; **tổng CĐT = tổng KL × đơn giá TDI** từng sheet.
2. Mỗi ô xanh `E2EFDA` có ghi chú `GIÁ THAM KHẢO WEB` kèm URL; số dòng `Nguồn giá` = `Giá tham khảo` bằng số ô xanh; **không ô giá thư viện nào bị tô xanh**; mọi dòng công việc có `Nguồn giá`; không dòng Chưa khớp nào bị gán mã vì có giá web.
3. File tổng hợp: số đầu mục có giá = số đầu mục có giá trong `REFPRICE.json`, từng giá khớp số đã ghi ở file TDI; mọi dòng có giá đều có `Link nguồn`, `Thứ hạng KQ`, `Ngày tra`; **0 dòng `Ước tính`**; Σ `Số đầu mục` ở `Tổng hợp` không đổi so với trước khi chạy.
4. Mọi đầu mục trong danh sách tra đều có kết quả (có giá hoặc ghi "không thấy giá"); đầu mục chưa tra (nếu hết giới hạn thời gian) được liệt kê rõ.
5. Dữ liệu CĐT, mã hiệu, độ tin cậy không đổi so với trước khi chạy (đếm lại số ô độ tin cậy `Chưa khớp`).
6. Truy vấn trong `REFPRICE.json` và trong file tổng hợp không chứa tên dự án, CĐT, KL.

## Báo cáo cho người dùng (ngắn)

- Số đầu mục đã tra (dùng cache / tra mới), số có giá, số không thấy giá, số dòng BOQ đã ghi giá; thời gian chạy và số lô.
- Bảng ngắn (đầu mục, giá, ĐVT, trang nguồn, VAT) cho các đầu mục KL lớn hoặc giá trị lớn để người rà soát kiểm; nêu đầu mục có `không rõ VAT`.
- Số ô CĐT bị ghi đè khi chạy lại `fillcdt`; tổng thành tiền CĐT đang dùng giá tham khảo.
- File đã ghi: file TDI (tên, kích thước, giờ ghi) và file tổng hợp (tên, thư mục).
- **Nhắc**: lọc cột `Nguồn giá` = `Giá tham khảo` (ô xanh lá nhạt) để rà; giá tham khảo chưa gồm hệ số, nhiều trang không ghi VAT; khi báo giá NCC về, bước gen giá NCC phải ghi đè ô giá tham khảo, bỏ nền xanh và đổi `Nguồn giá` thành `NCC` (kèm tên NCC + ngày báo giá).
- **Bước tiếp theo**: `boq-tao-ycbg` (lập YCBG cho các đầu mục chưa có giá trong thư viện, kể cả đầu mục đã có giá tham khảo); không tự chạy khi người dùng chưa xác nhận.

## Bảng kiểm tuân thủ (ngắn, luôn có ở cuối báo cáo)

| Mục | Cách kiểm |
|---|---|
| QT1 Không bịa giá | Mọi giá có URL, thứ hạng, ngày tra; không "Ước tính", không giá cỡ khác |
| QT2 Truy vấn sạch | 0 truy vấn chứa tên dự án / CĐT / mã dự án / KL |
| QT3 Phạm vi tra | Số đầu mục tra = số đầu mục sheet `Chưa có trong TV` (trừ đầu mục đã có giá hoặc cache) |
| QT4 Giữ Chưa khớp | Độ tin cậy và mã hiệu không đổi |
| QT5 Không sửa CĐT | Tên, KL, ĐVT CĐT không đổi (đối chiếu trước / sau) |
| QT6 Một file | Đúng một file TDI, file tổng hợp ghi đúng tên, kiểm mtime sau commit |
| QT7 Mã dự án | File tổng hợp khớp file TDI |
| Bước 3 | Ô xanh, ghi chú, `Nguồn giá`, bộ lọc |
| Bước 4 | `fillcdt` chạy lại, tổng CĐT = tổng KL × đơn giá TDI |
| Bước 5 | Cột giá, link, thứ hạng, ngày tra đủ; `Tổng hợp` tính lại |
| Bước 6 | Các mục kiểm 1–6 ở trên |