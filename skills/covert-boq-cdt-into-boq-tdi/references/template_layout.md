# Layout template TDI (bên trái CĐT, bên phải TDI)

Script `boq_tdi.py build` tạo đúng layout này. Tài liệu này dùng để:

- Kiểm tra file đầu ra.
- Làm tay khi không chạy được Python, ví dụ trong add-in Claude for Excel: ghi trực tiếp vào workbook theo đúng các cột và công thức dưới đây.

Ví dụ bên dưới dùng **Z = 3 khu** (Basement, Block A, Block B), giống hệt sheet `B2_SERVICE` mẫu. Nếu CĐT có số khu khác thì mọi cột từ "Tổng KL" trở sang phải dịch tương ứng (mỗi khu thêm 1 cột KL bên trái và 3 cột thành tiền bên phải).

## Khối đầu sheet (dòng 1–9)

| Ô | Nội dung |
|---|---|
| BC1:BC4 | Tên công ty TDI, tên tiếng Anh, địa chỉ, Tel/email (căn phải) |
| A5 (merge A5:BC5) | `CHI TIẾT GIÁ` |
| A6 | `="Dự án: "&Summary!$C$10` |
| A7 | `Hạng mục: <tên hệ thống>` |
| BB6 / BC6 | `Số:` / `=Summary!$C$11` |
| BB7 / BC7 | `Ngày:` / `=Summary!$C$12` |
| Dòng 8 | Nhãn: O8 `HS hao hụt`, P8 `HSVL`, Q8 `HSNC`, R8 `HSTB`, T8 `HSKL`, BB8 `HS chung`, BC8 `LÀM TRÒN` |
| Dòng 9 | O9 = 1.02 (nhập tay), P9 `=Summary!$C$8`, Q9 `=Summary!$E$8`, R9 `=Summary!$D$8`, T9 `=Summary!$B$8`, BB9 `=Summary!$F$8`, BC9 `=Summary!$G$8` |

## Tiêu đề bảng (dòng 10–12) và cột

**Bên trái (dữ liệu CĐT, chép nguyên giá trị, nền xanh nhạt):**

| Cột | Tiêu đề |
|---|---|
| A | STT |
| B | Hạng mục công việc |
| C | Thông số kỹ thuật |
| D | Mã hiệu |
| E | Nhãn hiệu |
| F | Xuất xứ |
| G | Đơn vị |
| H, I, J | Quantity/ Khối lượng, chia theo từng khu |
| K | Tổng khối lượng |

**Bên phải (template TDI, nền cam nhạt, chữ tiêu đề đỏ):**

| Cột | Tiêu đề | Công thức dòng hạng mục (dòng r) |
|---|---|---|
| L | Nội Dung Công Việc | `=+Br` |
| M | Quy cách Mã hiệu (TDI) | nhập tay / VLOOKUP thư viện |
| N | Nhãn hiệu Xuất xứ (TDI) | nhập tay |
| O | Đơn Vị | `=+Gr` |
| P | Đơn giá gốc: Vật liệu | nhập tay (chữ xanh) |
| Q | Đơn giá gốc: Nhân công | nhập tay (chữ xanh) |
| R | Bộ lọc: Chủng loại | mã, xem `chung_loai.md` |
| S | Bộ lọc: In ấn | để trống (file mẫu dùng `XOA` để đánh dấu dòng không in) |
| T | Tổng KL tạm tính | `=+Kr` |
| U–AL | Diễn giải chi tiết KL (18 cột: Trục chính 2, Tầng hầm 3, Khối đế 3, Hành lang 3, Căn hộ 5, Khác 2) | để trống, cột được group và ẩn |
| AM | KL | `=IF(OR(Or="m",Or="m2",Or="m3"),Tr*$T$9,Tr)` |
| AN | Đơn giá: Vật liệu | xem bảng công thức bên dưới |
| AO | Đơn giá: Nhân công | `=ROUNDDOWN(Qr*$Q$9*$BB$9,$BC$9)` |
| AP | Đơn giá: ĐG Tổng | `=ANr+AOr` |
| AQ, AR, AS | Khu 1: Cung cấp / Lắp đặt / Tổng | `=+ANr*Hr`, `=+AOr*Hr`, `=+APr*Hr` |
| AT, AU, AV | Khu 2 | tương tự với `Ir` |
| AW, AX, AY | Khu 3 | tương tự với `Jr` |
| AZ, BA, BB | Total | `=ANr*$Kr`, `=AOr*$Kr`, `=APr*$Kr` |
| BC | Ghi Chú | |

Cột K bên trái có công thức `=+Hr+Ir+Jr`.

**Công thức đơn giá vật liệu (AN) theo nhóm:**

| Điều kiện | Công thức |
|---|---|
| Đơn vị m / m2 / m3 (hao hụt) | `=ROUNDDOWN(Pr*$P$9*$BB$9*$O$9,$BC$9)` |
| Mã thuộc nhóm thiết bị (TB) | `=ROUNDDOWN(Pr*$R$9*$BB$9,$BC$9)` |
| Còn lại (VL) | `=ROUNDDOWN(Pr*$P$9*$BB$9,$BC$9)` |

## Loại dòng

| Loại | Nhận biết | Xử lý |
|---|---|---|
| Section cấp 0 | STT là chữ in hoa A, B… (không có đơn vị) | nền vàng, chữ đậm; các cột thành tiền dùng `=SUBTOTAL(9, dòng kế tiếp : dòng cuối section)` |
| Section cấp 1 | STT là số La Mã I, II… | như cấp 0 |
| Section cấp 2 / 3 | STT dạng I.1 / I.1.2 | như cấp 0 |
| Ghi chú (note) | không đơn vị, không KL, ví dụ "Cung cấp và lắp đặt…" | đủ công thức như hạng mục (kết quả = 0) |
| Hạng mục (item) | có đơn vị hoặc có KL | đủ công thức như bảng trên |

Vì SUBTOTAL lồng nhau tự bỏ qua nhau, dòng cuối **TỔNG CỘNG** chỉ cần `=SUBTOTAL(9, toàn bộ vùng dữ liệu)` cho từng cột thành tiền.

## Sheet Summary (chép nguyên từ mẫu)

| Vị trí | Nội dung |
|---|---|
| B7:G8 | Hệ số HSKL, HSVL, HSTB (ô D), HSNC (ô E), HS chung, LÀM TRÒN |
| B9:G24 | Hệ số theo loại vật liệu (giữ nguyên mẫu) |
| E25:H30 | Bảng diện tích, đơn giá/m². Để trống cho dự án mới; công thức có IFERROR |
| A32 / A33 | `Dự án: …` / `Hạng mục: …` |
| G32:H33 | Số / Ngày |
| Dòng 38 | CHI PHÍ TRỰC TIẾP, `G38 = SUM(G39:G50)` |
| Dòng 39–50 | Mỗi sheet chi tiết một dòng: A/B/C…, tên, `Hệ`, 1. Cột E/F/G tham chiếu cột Total (cung cấp / lắp đặt / tổng) ở dòng "Tổng cộng trước thuế" của sheet |
| Dòng 51–59 | Chi phí gián tiếp (% trên G38), tổng chưa VAT, VAT 10%, tổng có VAT (giữ nguyên mẫu) |

## Định dạng

Mọi màu, font, viền và định dạng số lấy từ `assets/style_template.xlsx`:

- Sheet `_TPL`: dòng 1–12 là đầu sheet. Các dòng mẫu: 13 mục A, 14 mục I, 15 mục I.1, 16 mục I.1.1, 17 ghi chú, 18 ghi chú đậm, 19 hạng mục, 20 tổng cộng.
- Sheet `_SUM`: Summary.

Khi làm tay trong Excel, chép định dạng (Format Painter) từ file chào giá mẫu của TDI theo đúng loại dòng.

Freeze panes tại C13; AutoFilter từ dòng 12; nhóm cột diễn giải KL được ẩn.

**Dễ đọc (hàm `make_readable`):**

- Độ rộng cột:
  - STT 5,5–9
  - Hạng mục / Nội dung 46
  - Thông số, Mã hiệu, Nhãn hiệu, Xuất xứ: theo độ dài nội dung (12–32)
  - KL ≥ 13
  - Đơn giá 15–16
  - Thành tiền theo khu 17, Total 19
  - Ghi chú 24
- Mọi cột chữ bật wrap text, căn giữa theo chiều dọc.
- Chiều cao dòng = số dòng chữ ước tính × 15pt + 4.
- Tiêu đề bảng (dòng 10–12) cao 33pt mỗi dòng.
