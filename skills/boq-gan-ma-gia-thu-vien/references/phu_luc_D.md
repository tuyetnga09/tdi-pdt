<!-- Tách nguyên văn từ SKILL.md ngày 21/10/2026. Gồm Phụ lục D (D1–D3). -->

# Phụ lục D. Bảng kiểm tuân thủ (bắt buộc mỗi lần chạy, Quy tắc 30)

Mỗi mục ghi: trạng thái (`Đã áp` / `Không phát sinh` / `Lệch` / `Chưa làm`) + bằng chứng. Cột "Cách kiểm" là bằng chứng tối thiểu. Thêm quy tắc mới vào skill thì thêm dòng vào bảng này cùng lúc.

## D1. Quy tắc bắt buộc

| Mục | Cách kiểm (bằng chứng) |
|---|---|
| QT1 Bảng khớp là cuối cùng | `diff`: MẤT = 0, ĐỔI = 0 (trừ "Đề xuất đổi"); số dòng lấy từ bảng |
| QT2 Không tự thêm loại trừ giá | Số dòng có cờ A/DC/NSC/PS, mỗi dòng có trích HSMT; dòng chưa xác nhận vẫn có giá |
| QT3 Thang/máng nhiều độ dày | Lọc nhóm TCap/MCap: 0 dòng Chưa khớp vì độ dày |
| QT4 uPVC PN12 → PN16 | Lọc "PN12": mọi dòng gán PN16, Thấp, có ghi chú |
| QT5 Khớp theo tên A5 | Lọc công tắc / ổ cắm / ổ sàn / chống nước: 0 dòng Chưa khớp |
| QT6 Cáp trung thế | Lọc 3.6kV / 12,7/24kV / 22kV / 35kV: gán CXV cùng lõi tiết diện, Thấp, cờ bắt buộc báo giá |
| QT7 Ống hệ điện là ống luồn | Lọc hệ Điện / ĐN / BC / ELV: 0 dòng mang mã `P-D…-PVC/PPR/HDPE` |
| QT8 Ghi quy tắc đã áp | Mọi dòng Claude gán có Ghi chú nêu quy tắc; dòng Chưa khớp nêu đã tra gì |
| QT9 Phụ kiện đi kèm Chưa khớp | Số dòng ghi "gộp YCBG với …" |
| QT10 Không làm tắt | `todo` còn 0 dòng trên mọi sheet |
| QT11 Đầu mục giá 0 vẫn gán mã | Danh sách "có mã không giá" hợp lệ, có ghi chú cần báo giá |
| QT12 Tra cả bảng TV | Dòng Chưa khớp đều nêu nhóm / sheet ẩn / dòng không mã đã tra |
| QT13 Thu hồi / lắp lại / hoàn trả | Lọc từ khóa: 0 dòng có giá VL vật tư mới; lắp lại chỉ NC |
| QT14 HDPE gân xoắn | Lọc "HDPE" hệ điện: tra A10, lệch ghi rõ |
| QT15 A6 chỉ là gợi ý, nhất quán họ | Rà họ vật tư: 0 cặp cùng họ cùng sheet mà một gán được một Chưa khớp không giải thích |
| QT16 Cụm liệt kê thành phần vẫn bóc tách | Lọc "bao gồm / gồm / kèm": mọi cụm đã có dòng con |
| QT17 Rà lần 2 A12 | Đã chạy A12 cho mọi sheet; số dòng gán thêm ở vòng 2 |
| QT18 Dòng gộp phải bóc tách | Lọc "+" / "kèm" / danh sách QT18: 0 dòng gộp còn một mã |
| QT19 Khóa trùng không SUMIFS | Cảnh báo khóa trùng của `write` = đã xử lý hết |
| QT20 Phụ kiện ống gió theo diện tích | Mọi dòng Thấp, có công thức, nằm trong danh sách cần duyệt |
| QT21 Tên vùng rác, vòng, ô lỗi | `recalc`: ô lỗi 0, vòng 0, tên vùng rác 0 |
| QT22 Không bỏ quy tắc A | So số mục A1–A12 với bản skill trước: không giảm |
| QT23 Ô `Mã \| Tên` | Lọc cột Mã hiệu: 0 ô mã thuần / dạng cũ |
| QT24 Hệ số Loại VL | `TV` có cột P–S; thử đổi hệ số trên bản sao chỉ đổi đúng Loại VL |
| QT25 Bước 8 mặc định chạy | Danh sách sheet đã fill CĐT; tổng CĐT = KL × đơn giá TDI |
| QT26 Bóc tách kể cả thiếu script | Số cụm / dòng con đã chèn; nếu thiếu `insert` đã chèn XML |
| QT27 Ghép chỉ khi duyệt | Lọc `[Ghép]` / `terms`: mọi dòng có "GHÉP – đã được duyệt ngày …" |
| QT28 Mục bỏ qua / làm khác skill | Có trong báo cáo (hoặc "Không có") |
| QT29 Không tra web ở skill này | Không có truy vấn web; không ô xanh `E2EFDA` mới; dòng Chưa khớp để trống giá, `Nguồn giá` = `Chưa có giá`; autoFilter phủ cột `Nguồn giá`; ô `Giá tham khảo` có từ trước được giữ nguyên |
| QT30 Bảng kiểm này | 0 mục `Chưa làm` |
| QT31 Thư viện theo thư mục dự án | Nguồn ME và XD đã dò theo thứ tự ưu tiên (đường dẫn, cách lọc file, số file); file XD qua kiểm bố cục `TVGIA` B / G / J; script lấy ở `_Claude_tools` chung |
| QT32 File tổng hợp đầu mục chưa có trong thư viện | File ghi đúng tên vào `Claude outputs`; Σ đầu mục = số đầu mục Chưa khớp duy nhất; cột giá tham khảo / link / thứ hạng / ngày tra để trống, `Tình trạng giá` = `Chưa tra giá web`; 0 dòng "Ước tính"; `recalc` file tổng hợp 0 lỗi |
| QT34 Công thức chờ dòng chưa có giá | Bước 7 mục 14: số ô nối / số ô trong phạm vi, ô bỏ qua có lý do; tổng trước = sau; thử chọn / xóa mã trên bản sao ra đúng; `diff` 0 / 0 |
| QT33 Bóc tách XD theo TVGIA + cột Tên công việc | Sheet XD có cột N + droplist; số cụm / dòng con; Loại dòng `Con` không cộng tổng (SUMIFS); **tổng XD trước = sau** (so từng ô); Σ con ≠ cha đã liệt kê và giữ giá cha; dòng giá viết riêng đã ghi chú; `TVGIA` đúng định dạng nguồn; `<cols>` không chồng |

## D2. Bước quy trình

| Mục | Cách kiểm |
|---|---|
| Đọc hết skill + references | Đã xem SKILL.md đến dòng cuối và mọi file `references/` |
| Bước 0 (0a–0f) | Hệ từng sheet; nguồn thư viện ME / XD đã báo (QT31); `cols` OK; `tv_loader check` = 0/0/0; cặp cột giá; A9; hệ số hàng 4; không có `~$`, đã ghi mtime |
| Bước 1 | Cột nhận theo tiêu đề cho từng sheet |
| Bước 2 | Đã đọc phân mục, Ghi chú CĐT, spec / bản vẽ khi cần; dòng hỏi vật liệu đã liệt kê |
| Bước 3 | Mọi dòng có Nhóm VT; không ép CTKhac |
| Bước 4 (4a, 4b) | `apply` đã chạy; `write` đã chạy; thang độ tin cậy đủ |
| Bước 5 | `insert` (hoặc XML) đã chạy sau điểm dừng 2 |
| Bước 6 | Công thức giá tham chiếu ô của dòng; dòng Thấp có giá |
| Bước 6b | `DS_MaHieu` ẩn có mặt; droplist mọi sheet |
| Bước 6c | Cột `Nguồn giá` và bộ lọc như QT29 |
| Bước 6e | Như QT34 |
| Bước 6d | Như QT32 (file tổng hợp đã ghi, sheet phụ đủ dòng so với nguồn) |
| Bước 7 | Đủ mục 1–13 |
| Bước 8 | Như QT25; ô CĐT bị ghi đè đã báo |
| Bước 9 / 9b | Chỉ khi người dùng báo đã duyệt; chưa duyệt ghi `Không phát sinh (chờ duyệt)` |
| Giao thức một file | Đúng một `<tên>_TDI.xlsx` (hoặc `_ra-soat-lan-N` theo ngoại lệ) cộng file tổng hợp Bước 6d trong `Claude outputs`; kiểm mtime sau commit |
| Điểm dừng 1–4 | Đã dừng đúng chỗ, hoặc ghi rõ người dùng cho chạy liền mạch |

## D3. Bảng tra và hệ XD

| Mục | Cách kiểm |
|---|---|
| A1–A12 | Mỗi bảng: số dòng đã áp, hoặc `Không phát sinh` kèm từ khóa đã lọc (ví dụ A11: lọc "kV" ra 0) |
| Phần XD (khi có sheet XD) | Nguồn `TVGIA` (file dự án / file nội bộ) và kết quả kiểm bố cục, so `TVGIA` trong file với nguồn; `tv`, `todo`, `apply`, `write`, `recalc` đã chạy; tên trùng dùng tham chiếu trực tiếp; số cứng thầu phụ / tỷ lệ % giữ nguyên; bóc tách theo TVGIA (Quy tắc 33): xem QT33 |
| C3 | Rà các khớp sai C3: 0 dòng phạm |
| Không có sheet XD | Ghi `Không phát sinh (file không có sheet XD)` |
