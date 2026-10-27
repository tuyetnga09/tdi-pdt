# Bảng tra gán mã + spec TDI25538

Trích từ skill boq-gan-ma-gia-thu-vien (Phụ lục A, B). Kiểm lại với thư viện mỗi dự án.

# Phụ lục A. Bảng tra (kiểm lại với thư viện mỗi dự án)

## A1. Quy đổi kích thước

**PPR** — mã `P-D{OD}-PN{n}-PPR`, có PN10/16/20/25.

| DN | 15 | 20 | 25 | 32 | 40 | 50 | 65 | 80 | 100 | 125 | 150 |
|---|---|---|---|---|---|---|---|---|---|---|---|
| OD | 20 | 25 | 32 | 40 | 50 | 63 | 75 | 90 | 110 | 140 | 160 |

**HDPE** — mã `P-D{OD}-PN{n}-HDPE-PE100`.

| DN | 50 | 65 | 80 | 100 | 125 | 150 | 200 | 250 | 300 |
|---|---|---|---|---|---|---|---|---|---|
| OD | 63 | 75 | 90 | 110 | 140 | 160 | 225 | 280 | 315 |

**PVC** — mã `P-D{OD}-PN{n}-PVC`.

| DN | 15 | 20 | 25 | 32 | 40 | 50 | 65 | 80 | 100 | 125 | 150 | 200 | 250 | 300 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| OD | 21 | 27 | 34 | 42 | 48 | 60 | 75 | 90 | 110 | 140 | 160 | 200 | 250 | 315 |

**Inox** — mã `P-DN{dn}x{dày}-SLT`. Chọn độ dày SCH10S gần nhất, không nhỏ hơn:

| DN | 15–20 | 25–50 | 65–100 | 125–150 | 200 | 250 |
|---|---|---|---|---|---|---|
| Dày (mm) | 2.11 | 2.77 | 3.05 | 3.4 | 3.76 | 4.19 |

**Thép đen / mạ kẽm**: một DN có nhiều độ dày → Thấp, ghi ứng viên, chọn theo spec nếu có.

**Ống đồng** — mã `P-D{OD}x{dày}Cu` hoặc `P-D{OD}x{dày}-Cu` (tra cả hai dạng, dấu phẩy thập phân).
- BOQ ghi Ø làm tròn; quy đổi sang OD thư viện:

| Ø BOQ | 6.4 | 9.5 | 12.7 | 15.9 | 19.1 | 22.2 | 25.4 | 28.6 | 31.8 | 34.9 | 41.3 |
|---|---|---|---|---|---|---|---|---|---|---|---|
| OD thư viện | 6,35 | 9,52 | 12,7 | 15,88 | 19,05 | 22,22 | 25,4 | 28,58 | 31,75 | 34,93 | 41,28 |

- Ø38.1 và Ø44.5 không có trong thư viện.
- **Chọn độ dày** thư viện gần nhất, không nhỏ hơn spec. 0.81 mm ≈ ống 0.8 mm thương mại (TB).
- **Thư viện không đạt spec** → lấy loại dày nhất có, Thấp, ghi "thấp hơn spec".

**Bảo ôn ống** — mã `P-D{ID}x{dày}mm-BO`.
- Độ dày có: 10/13/19/25/38 mm.
- ID có: 6, 10, 13, 16, 19, 22, 25, 28, 32, 35, 38, 42, 48, 51, 54, 60, 67, 73, 76, 79, 89, 102, 105, 108, 114, 127, 130, 133, 140.
- **Chọn ID gần OD nhất**, lệch tối đa ~1 mm:

| Ống | OD → ID bảo ôn |
|---|---|
| Đồng | 6.35→6, 9.52→10, 12.7→13, 15.88→16, 19.05→19, 22.22→22, 25.4→25, 28.58→28, 31.75→32, 34.93→35, 41.28→42 |
| uPVC | 21→22, 27→28, 34→35, 42→42, 48→48, 60→60, 75→76, 90→89, 110→114, 140→140 |

- **D≥160**: không có bảo ôn ống; báo giá, hoặc dùng bảo ôn tấm `S-{t}mm-BO` nếu người dùng đồng ý.
- **Độ dày** lấy mức thư viện gần nhất nằm trong khoảng spec (spec ≥15 → 19; 20–25 → 25).

## A2. Cáp điện

**Mã theo cấu tạo:**

| Cấu tạo | Mã | Độ tin cậy |
|---|---|---|
| Cu/PVC | `CV-{n}x{s}` | |
| Cu/PVC/PVC | `CVV-{n}x{s}` | |
| Cu/XLPE/PVC | `CXV-{n}x{s}` | |
| Cu/XLPE/DSTA | `DSTA-{n}x{s}` | |
| Cu/MICA/XLPE/FR-PVC | `FRN/CXV-{n}x{s}` | Thấp |

- Mica luôn là chống cháy (FRN), không gán FSN.
- **Bỏ "mm2" trước khi tách tiết diện**, để không nhầm số "2".
- **Ký hiệu số sợi × lõi:**

| BOQ ghi | Cách hiểu |
|---|---|
| `1x2Cx1.5` | `CVV-2x1.5` |
| `4x2Cx1.0` | 4 sợi `CVV-2x1` |
| `3x1Cx2.5 + (E)1.5` | 3 sợi `CV-1x2.5` + 1 sợi `CV-1x1.5` |

  Dòng ghép sợi → bóc tách, hệ số = số sợi. Không gán mã nhiều lõi gộp.
- **Tiết diện không chuẩn** (2.0 mm²) → Chưa khớp, hỏi CĐT.
- **Cáp tín hiệu, âm thanh, CAT6, quang, RS485** → Chưa khớp.

## A3. Van (Van Khoa)

**Kiểu nối**: DN≤32 nối ren `-TI-`; DN≥40 mặt bích `-F-`. DN40–100 có cả hai bản: chọn mặt bích, TB, ghi "giả định nối mặt bích".

**Mã theo họ van:**

| Họ | Mã |
|---|---|
| Van cổng | `GV` |
| Van 1 chiều | `CV` |
| Van bi tay gạt (đến DN100) | `BV-TI` |
| Van bướm | `BV-F` |
| Van bướm có công tắc giám sát | `BVLS-F` |
| Van phao | `BFV` |
| Van an toàn | `PRV` |
| Van giảm áp | `RV` |
| Van báo động | `AV` |
| Búa nước | `WHA` |
| Y lọc | `YF` |
| Khớp nối mềm | `FJ` |
| Van đáy | `FV` |
| Đồng hồ nước lạnh / nóng | `CWM` / `HWM` |
| Van cân bằng | `VCB` |
| Van tuần hoàn | `VTH` |

**DN ngoài dải thư viện** → Chưa khớp, không lấy DN gần nhất.

## A4. Từ đồng nghĩa
| Tên trong BOQ | Đầu mục thư viện | Tin cậy |
|---|---|---|
| Khớp nối mềm DN | `FJ-F/TI-DN` | Cao |
| Bộ chống nước va, Van búa nước | `WHA-F-DN` | TB/Cao |
| Van báo động | `AV-F-DN` | Cao |
| Van phao | `BFV-F-DN` | Cao |
| Van xả áp | `RV-F-DN` (tạm) | Thấp |
| Bộ lọc đường ống | `YF-F-DN` (tạm) | Thấp |
| Rọ hút | `FV-F-DN` (tạm) | Thấp |
| Van cổng có công tắc giám sát | `BVLS-F-DN` | Thấp |
| Đồng hồ nước | `CWM-…` | Thấp |
| Thông tắc sàn DN | `CO-D{OD}-PVC` | Thấp |
| Ống G.I (hệ điện) | `OLD-THEP-C-DN…` | TB |
| Ống PVC / PVC conduit (điện, báo cháy) | `OLD-PVC-C-D…-N` (nổi; âm dùng `-C`) | TB |
| Ống mềm PVC | `OLD-PVC-M-D…` | Cao |
| Ống nước nóng DN | `P-D{OD}-PN20-PPR` | TB |
| Thép mạ kẽm dày 0.58–1.38 mm (m2, ống gió) | `P-{t}-STEEL` | TB |
| Ống gió chống cháy EI30/45/60/120 | `EI30…EI120` (không lấy `PK-EI…`) | TB |
| Ống gió mềm có/không bảo ôn | `FAD-I-D…` / `FAD-D…` | Cao |
| Ø dưới "ống dẫn chất làm lạnh có lớp cách nhiệt" (ODBO) | Bóc tách: `P-D…Cu` + `P-D…-BO` | theo độ dày |
| DN dưới "ống thoát nước ngưng có lớp cách nhiệt" (ONN) | Bóc tách: `P-D{OD}-PN10-PVC` + `P-D…-BO` | TB |
| Vỏ tủ chữa cháy | "Hộp chữa cháy 600x400x200mm" / "1200x650x200mm" (mã 0) | Thấp |
| Van góc giảm áp DN50/65 | "Van góc DN50/DN65" (PCCC) | Thấp |
| Cuộn vòi DN50 / DN65 20 m | "Cuộn ống mềm chữa cháy (trong nhà) - DN50mm" / "(ngoài trời) DN65mm" | TB/Thấp |
| Lăng phun D13 | "Lăng chữa cháy (trong nhà) -DN50x13mm" | Cao |
| Đế đầu báo | "Đế cho đầu báo khói, nhiệt (4-6\")" (mã 0) | Cao |

## A5. Tra theo tên (mã 0 hoặc mã trùng, tên duy nhất trong nhóm)

**Cách ghi:**
- **Ô Mã hiệu**: `(mã 0) <Tên công việc>`.
- **Công thức giá**: `SUMIFS(TV[Giá NC gốc],TV[Column1],"<nhóm>",TV[Tên công việc],"<tên>")`.
- **Ghi chú**: "VL cần báo giá NCC" khi thư viện không có giá VL.

**Các khớp đã dùng (tên BOQ → tên thư viện):**

| Tên BOQ | Tên thư viện | Tin cậy |
|---|---|---|
| Công tắc 1/2/3 nút, 1/2 chiều | "Công tắc N chiều X" | vỏ kim loại: Thấp |
| Ổ cắm | "Ổ cắm 3 chấu … (âm tường / âm sàn / chống nước)" | TB |
| Nút nhấn khẩn cấp, Điểm báo cháy bằng tay | "Nút nhấn khẩn cấp (địa chỉ)" | |
| Đầu báo khói / nhiệt / kết hợp | đầu báo tương ứng | loại không địa chỉ: Thấp |
| Còi đèn, module giám sát / cách ly / điều khiển | đầu mục cùng tên | |
| Thoát sàn | "Thoát sàn loại 1" | |
| Cầu thu mưa | "Cầu thu rác" | Thấp |
| Đèn EXIT, đèn sự cố, downlight, ốp trần | đầu mục cùng tên | |
| LED tuýp | theo máng huỳnh quang | Thấp |
| Sen tắm, bồn tiểu, phụ kiện WC | đầu mục cùng tên | |
| Máy nước nóng | "Bình nóng lạnh" | Thấp |
| Bồn cầu | "Bồn cầu bệt (xí bệt) loai 1" + vòi xịt | |
| Chậu rửa mặt | "Chậu rửa mặt (Lavabo) treo tường loại 1" + "Vòi chậu rửa mặt loại 1" | |
| Chậu rửa hố đơn/đôi | "Chậu rữa bát loại 1/2" + vòi | |

## A6. Không có trong thư viện TDI25538 (Chưa khớp → cần báo giá)

| Nhóm | Hạng mục |
|---|---|
| Van, đo lường | Van cầu, van điện từ, công tắc dòng chảy, lưu lượng kế, công tắc / đồng hồ áp lực |
| Thoát nước | Đĩa chống xoáy, chụp thoát mùi, ống PP thoát nước giảm âm |
| HVAC | Miệng gió, van gió VCD/MFD/NRD/PRD/OBD, ống gió EI90, ống khói, ống đồng Ø38.1/Ø44.5, bảo ôn D≥160, hộp che ống đồng, refnet / bộ chia gas VRV |
| PCCC | Họng khô, lăng D19, vòi và hộp vòi rửa sàn |
| Điện | Ống luồn thép mềm, băng thép chống sét, cáp tín hiệu / âm thanh, máng cáp cỡ thiếu (100x100) |
| Báo cháy | Hộp module, module nhận/xuất, rơ le |
| Thiết bị | Điện thoại, tay vịn người khuyết tật, RMU, máy biến áp, pô máy phát, thiết bị mạng / CCTV / ACS, bơm, tủ điều khiển bơm, máy phát |
| Quạt | Thư viện có đầu mục nhưng giá = 0 |

**Dòng lô trong phân mục điều hòa**: ghi hướng chiết tính.
- **Nạp gas**: tính kg theo chiều dài ống lỏng, tham chiếu `R410` (CTKhac, kg).
- **Phụ kiện ống đồng / ống nước ngưng**: đếm theo bản vẽ.
- **Refnet**: xin báo giá hãng.

## A7. Khớp sai đã bắt (dùng khi rà chéo)

| Dòng BOQ | Bị gán nhầm thành |
|---|---|
| Tủ chữa cháy | Lăng |
| Họng khô | Trụ chữa cháy |
| Điểm báo cháy bằng tay | Tủ báo cháy trung tâm |
| Băng thép | Ống thép đen |
| Cáp âm thanh | CXV |
| Máng 100x100 | 100x50 |
| Máy nước nóng | Van an toàn |
| Van bi DN250 | DN15 |
| Van cầu | Van cổng |
| EI90 | EI30 |
| Cáp 4x2Cx1.0 | CV-1x4 |
| Cáp 3x1Cx1.5+E | CVV-3x2.5+1x1.5 |
| Dòng cụm | một mã đơn lẻ |

# Phụ lục B. Spec đã kiểm chứng — dự án TDI25538 (tham khảo; dự án mới vẫn phải đọc spec riêng)

**Cấp thoát nước**

| Hạng mục | Vật liệu theo spec |
|---|---|
| Cấp nước DN<100 | PPR PN10 hàn nhiệt |
| Cấp nước DN≥100 | HDPE PE100 PN12.5 (người dùng đã chốt) |
| Nước nóng | PPR PN20 |
| Thoát trong nhà | **PP 3 lớp giảm âm EN1451** (spec 2.03). Thư viện chỉ có PP-R → Chưa khớp, YCBG |
| Thoát âm đất (kể cả bếp) | HDPE ISO 4427 |
| Phân mục ghi "uPVC" | uPVC PN12; thư viện thiếu thì PN16, Thấp |
| Inox | SS304 SCH10S; mặt bích ≥PN16 |

**Ống môi chất lạnh** (spec 1.02 §2.4): đồng BS EN 1057; ống <50 mm loại cứng class Y. Độ dày tối thiểu:

| Ø (mm) | 6.35 / 9.52 / 12.7 | 15.88 | 19.05 | 22.2 | 28.58 | 34.92 | 41.28 |
|---|---|---|---|---|---|---|---|
| Dày (mm) | 0.81 | 0.99 | 1.07 | 1.14 | 1.27 | 1.40 | 1.52 |

**Ống nước ngưng** (spec 1.02 §2.5): uPVC PN10 (BS 3505), bẫy nước ≥50 mm, bọc cao su ô kín.

**Bảo ôn** (spec 1.06):

| Ống | Đường kính | Độ dày bảo ôn |
|---|---|---|
| Môi chất | D6.4–15.9 | ≥15 mm |
| Môi chất | D19.1–41.3 | 20–25 mm |
| Môi chất | >D41.3 | 25–32 mm |
| Nước ngưng | ≤D48 | 15–19 mm |
| Nước ngưng | >D48 | 20–25 mm |

**Bộ bơm tăng áp** (spec 2.09 §2.5 + sơ đồ cấp nước):
- **Thành phần**: bơm ly tâm trục đứng đa cấp, bể áp lực, tủ điều khiển (mỗi bơm một biến tần), van, khớp nối mềm, chống va, đồng hồ, công tắc áp lực, van an toàn.
- **Mỗi nhánh**: van cổng – khớp mềm – bơm – khớp mềm – van 1 chiều – van cổng.
- **Bypass**: 1 nhánh gồm van 1 chiều + van cổng.
- **Lệch bản vẽ**: bản vẽ MEP-HR-CD-CT7-P trang 4 ghi Q=18 m3/h, H=23 m, trong khi BOQ ghi Q=27–39 m3/h, H=67–141 m. Đã báo.

