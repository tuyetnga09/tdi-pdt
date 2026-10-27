<!-- Tách nguyên văn từ SKILL.md ngày 21/10/2026. Gồm Phụ lục A1–A12. -->

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

**PVC** (ống nước, chỉ dùng cho hệ cấp thoát nước) — mã `P-D{OD}-PN{n}-PVC`. Phân mục ghi PN12 mà thư viện không có → PN16, Thấp.

| DN | 15 | 20 | 25 | 32 | 40 | 50 | 65 | 80 | 100 | 125 | 150 | 200 | 250 | 300 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| OD | 21 | 27 | 34 | 42 | 48 | 60 | 75 | 90 | 110 | 140 | 160 | 200 | 250 | 315 |

**Inox** — mã `P-DN{dn}x{dày}-SLT`. Chọn độ dày SCH10S gần nhất, không nhỏ hơn:

| DN | 15–20 | 25–50 | 65–100 | 125–150 | 200 | 250 |
|---|---|---|---|---|---|---|
| Dày (mm) | 2.11 | 2.77 | 3.05 | 3.4 | 3.76 | 4.19 |

**Thép đen / mạ kẽm**: một DN có nhiều độ dày → Thấp, ghi ứng viên, chọn theo spec nếu có.

**Ống đồng** — mã `P-D{OD}x{dày}Cu` hoặc `P-D{OD}x{dày}-Cu` (tra cả hai dạng, dấu phẩy thập phân).

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

**Mã theo cấu tạo (chỉ cáp hạ thế 0.6/1kV):**

| Cấu tạo | Mã | Độ tin cậy |
|---|---|---|
| Cu/PVC | `CV-{n}x{s}` | |
| Cu/PVC/PVC | `CVV-{n}x{s}` | |
| Cu/XLPE/PVC | `CXV-{n}x{s}` | |
| Cu/XLPE/DSTA | `DSTA-{n}x{s}` | |
| Cu/MICA/XLPE/FR-PVC | `FRN/CXV-{n}x{s}` | Thấp |

- **Cáp trung thế** (3.6kV trở lên, 12,7/24kV, 22kV…): gán tạm `CXV-{n}x{s}` hạ thế cùng số lõi, tiết diện, độ tin cậy Thấp, cờ bắt buộc báo giá (Quy tắc bắt buộc 6, A11).
- Mica luôn là chống cháy (FRN), không gán FSN.
- **Bỏ "mm2" trước khi tách tiết diện**, để không nhầm số "2".
- **Ký hiệu số sợi × lõi:**

| BOQ ghi | Cách hiểu |
|---|---|
| `1x2Cx1.5` | `CVV-2x1.5` |
| `4x2Cx1.0` | 4 sợi `CVV-2x1` |
| `3x1Cx2.5 + (E)1.5` | 3 sợi `CV-1x2.5` + 1 sợi `CV-1x1.5` |

  Dòng ghép sợi → bóc tách, hệ số = số sợi. Không gán mã nhiều lõi gộp.
- **Cáp ≥70mm²** vẫn gán mã và giá bình thường (phạm vi NSC theo mục A9 nếu HSMT ghi).
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
| Ống PVC / PVC conduit (điện, điện nhẹ, báo cháy) | `OLD-PVC-C-D…-N` / `-C` theo A10 | TB |
| Ống mềm PVC | `OLD-PVC-M-D…` | Cao |
| Ống HDPE D{ID}/{OD} luồn cáp (hệ điện, cáp ngầm) | `OLD-HDPE-DN{OD}-{ID}` theo bảng A10 | TB; lệch OD ≤5mm: Thấp |
| Ống nước nóng DN | `P-D{OD}-PN20-PPR` | TB |
| Thép mạ kẽm dày 0.58–1.38 mm (m2, ống gió) | `P-{t}-STEEL` | TB |
| Ống gió chữ nhật BOQ tính m, cỡ có trong thư viện (bước 100) và đúng độ dày | `P-{W}x{H}x{t}-RD` | Cao |
| Ống gió chữ nhật BOQ tính m, cỡ không có (250x150, 1350x400…) | `P-{t}-STEEL` (m2), Hệ số ĐV = chu vi 2(W+H)/1000 | TB; độ dày giả định: Thấp |
| Ống gió chống cháy EIxx BOQ tính m | `EIxx` (m2), Hệ số ĐV = chu vi 2(W+H)/1000 | TB |
| Ống gió + "bọc tiêu âm trong" (rockwool) | Bóc tách: ống gió + tiêu âm (Chưa khớp, hệ số = chu vi) | Thấp |
| Miệng gió 1 lớp (kèm OBD), cỡ ≤ 1000x1000 | `VF-1L-{W}x{H}-RD` (OGio, sheet ẩn "Cua gio") | Thấp (OBD, lưới, G4 chưa tính) |
| Miệng gió 2 lớp | `VF-2L-{W}x{H}-RD` | TB |
| Miệng gió sọt trứng | `VF-1L-…` (thư viện không có sọt trứng) | Thấp |
| Louver chắn mưa / đầu ống gió chắn mưa, cỡ ≤ 1000x1000 | `LV-Z-{W}x{H}-RD` (louver 1 lớp nan Z) | Thấp |
| Van gió VCD / NRD / FD / MFD (FDM) / PRD | "Van điều chỉnh tay gạt VCD, kích thước…" / "Van gió 1 chiều NRD, kích thước…." / "Van gió dập lửa thường FD, kích thước…." / "Van gió dập lửa điều khiển bằng motor điện FD, kích thước…." / "Van gió điều áp PRD, kích thước…." (TB HVAC, tra theo tên, giá = 0 → QT11) | TB; MPRD có động cơ: Thấp |
| Giá đỡ thang máng / ống gió / ống nước, cắt đục - chèn trát ống âm, thử áp - thử kín, đầu cốt | Đầu mục CTKhac không mã, tra theo tên đúng dải kích thước | TB |
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

**Cách ghi:** `write` với `keycol` = `Tên công việc`; ô Mã hiệu tự ghi `Mã | Tên` (mã 0: `(mã 0) | <Tên công việc>`). Ghi chú "VL cần báo giá NCC" khi thư viện không có giá VL.

**Các khớp bắt buộc (tên BOQ → tên thư viện):**

| Tên BOQ | Tên thư viện | Tin cậy |
|---|---|---|
| Công tắc đơn/đôi/ba, 1/2 chiều | "Công tắc N chiều X" | TB; vỏ kim loại: Thấp |
| Ổ cắm đơn/đôi | "Ổ cắm 3 chấu … (âm tường)" | TB; vỏ kim loại: Thấp |
| Ổ cắm sàn (kể cả chống chịu thời tiết) | "Ổ cắm 3 chấu … (âm sàn)" | TB |
| Ổ cắm chống chịu thời tiết / chống nước | "Ổ cắm 3 chấu … (chống nước, âm tường)" | TB |
| Ổ cắm nguồn cho thiết bị / tủ lạnh | "Ổ cắm 3 chấu đơn (âm tường)" | Thấp |
| Ổ cắm điện thoại / mạng | đầu mục cùng tên | TB |
| Nút nhấn khẩn cấp, Điểm báo cháy bằng tay | "Nút nhấn khẩn cấp (địa chỉ)" | TB |
| Đầu báo khói / nhiệt / kết hợp | đầu báo tương ứng | loại không địa chỉ: Thấp |
| Còi đèn, module giám sát / cách ly / điều khiển | đầu mục cùng tên | |
| Thoát sàn | "Thoát sàn loại 1" | |
| Cầu thu mưa | "Cầu thu rác" | Thấp |
| Đèn EXIT, đèn sự cố, downlight, ốp trần | đầu mục cùng tên; khác công suất: Thấp | |
| LED tuýp | theo máng huỳnh quang | Thấp |
| Sen tắm, bồn tiểu, phụ kiện WC | đầu mục cùng tên | |
| Máy nước nóng | "Bình nóng lạnh" | Thấp |
| Bồn cầu | "Bồn cầu bệt (xí bệt) loai 1" + vòi xịt | |
| Chậu rửa mặt | "Chậu rửa mặt (Lavabo) treo tường loại 1" + "Vòi chậu rửa mặt loại 1" | |
| Chậu rửa hố đơn/đôi | "Chậu rữa bát loại 1/2" + vòi | |

## A6. Không có trong thư viện TDI25538 (Chưa khớp → cần báo giá)

Danh sách này chỉ là gợi ý đã kiểm chứng ở TDI25538, không thay cho việc tra `TV` (Quy tắc 15).

| Nhóm | Hạng mục |
|---|---|
| Van, đo lường | Van cầu, van điện từ, công tắc dòng chảy, lưu lượng kế, công tắc / đồng hồ áp lực |
| Thoát nước | Đĩa chống xoáy, chụp thoát mùi, ống PP thoát nước giảm âm |
| HVAC | Miệng gió/louver cỡ > 1000x1000, louver tính theo m², vát góc chắn mưa, OBD rời, tiêu âm ATT cho quạt (thư viện chỉ có tiêu âm FCU cao 300 `TA-…`), ống gió EI90, ống khói, ống đồng Ø38.1/Ø44.5, bảo ôn D≥160, hộp che ống đồng, refnet; bộ chia gas VRV (tên trùng 2 dòng → chọn tay) |
| PCCC | Họng khô, lăng D19, vòi và hộp vòi rửa sàn |
| Điện | Busway và phụ kiện busway, tủ điện căn hộ, ống luồn thép mềm, băng thép chống sét, cáp tín hiệu / âm thanh, máng cáp cỡ thiếu (ví dụ 1100x500, 500x250) |
| Điện (trung thế) | Tủ RMU / LBS / VCB, relay bảo vệ 50/51, tủ RTU SCADA, chống sét van, hộp đầu cáp trung thế (T-plug, trong nhà): Chưa khớp, bóc tách tủ theo Quy tắc 16. Cáp trung thế: gán tạm hạ thế (Quy tắc 6, A11) |
| Báo cháy | Hộp module, module nhận/xuất, rơ le |
| Thiết bị | Điện thoại, tay vịn người khuyết tật, RMU, máy biến áp, pô máy phát, thiết bị mạng / CCTV / ACS / phần mềm quản lý, bơm, tủ điều khiển bơm, máy phát |
| Điện (linh kiện tủ) | Sheet linh kiện tủ điện (ví dụ `CT TU DIEN`, 516 dòng): Chưa khớp, thư viện không có; gom YCBG theo hãng / tủ |
| HVAC (lưu ý) | Miệng gió `VF-1L-WxH-RD` có đến cỡ ~1200 rộng: tra trước khi kết luận "cỡ > 1000 không có" |
| Bơm | Bơm chìm / bơm tăng áp: thư viện chỉ có theo đầu mục cụ thể, tra theo tên trước; ghép "2 bơm + tủ" phải tách (Quy tắc 18) |
| Quạt, van gió | Thư viện CÓ đầu mục nhưng giá = 0 → **gán mã theo Quy tắc bắt buộc 11**, vẫn đưa vào YCBG (không còn là Chưa khớp) |

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
| Cáp trung thế 24kV | `DSTA-…` (cáp có giáp) hoặc mã hạ thế mà không ghi cờ "giá tạm, bắt buộc báo giá": gán tạm đúng là `CXV` cùng tiết diện, Thấp, có cờ |
| Ống HDPE D100/130, D195/150 | Chưa khớp "không có cỡ" trong khi thư viện có `OLD-HDPE-DN130-100`, `OLD-HDPE-DN190-150` |
| Thu hồi / lắp đặt lại cáp tận dụng | Giá VL của cáp mới (phải là chỉ NC) |
| Tủ RMU 4 khoang (CĐT liệt kê 2 LBS + 2 VCB + relay) | Giữ một dòng Chưa khớp, không bóc tách |
| Ống PVC D20/D25 hệ điện thoại/dữ liệu, CCTV, điện nhẹ | Ống nước `P-D90-PN3-PVC` / `P-D21-PN25-PVC` (sai hệ và sai cỡ) |
| Máng 100x100 | 100x50 |
| Máy nước nóng | Van an toàn |
| Van bi DN250 | DN15 |
| Van cầu | Van cổng |
| EI90 | EI30 |
| Cáp 4x2Cx1.0 | CV-1x4 |
| Cáp 3x1Cx1.5+E | CVV-3x2.5+1x1.5 |
| Cút ren, côn ren, rắc co ren | Cút / côn PVC hàn dán (hệ ren không phải PVC) |
| Tê gang FBF, phụ kiện gang | Tê PVC |
| Tê ống gió 1300–1700x0,95 | `T-1300…1700x0,95-RD` qua SUMIFS (khóa trùng, bị cộng dồn): dùng ước tính diện tích, Thấp (Quy tắc 19) |
| Dòng cụm | một mã đơn lẻ |
| Dòng gộp (dàn lạnh + điều khiển, đèn trụ + cần + bóng, xí bệt + vòi xịt, ống gió + tiêu âm…) khi thiếu script `insert` (TDI25538, 15/10/2026) | Một dòng `[Ghép]` cộng nhiều SUMIFS thay vì bóc tách thành dòng con (sai Quy tắc 16, 18; xem Quy tắc 26, 27) |
| Ô Mã hiệu chỉ ghi mã `DN50mm` (PCCC, trùng cho Van góc và Cuộn vòi) rồi lưu Bảng khớp | Mất tên, giá thành giá của cả hai đầu mục: luôn lưu và truyền ô `Mã \| Tên` (Quy tắc 23) |

Lỗi ngược (để trống sai) đã gặp khi chạy lại (vòng 2 dự án số 2 còn bắt thêm ~585 dòng có trong thư viện nhưng để Chưa khớp: ống gió, cút/côn/tê, miệng gió, van, mặt bích, đèn, hộp nối; xem A12): thang/máng cáp, ống uPVC PN12, công tắc / ổ cắm vỏ kim loại, ổ cắm sàn, cáp ≥70mm² bị bỏ giá cắt vì tự suy diễn "NSC"; ống HDPE luồn cáp D100/130, D195/150 ghi "không có cỡ" trong khi D65/50 cùng sheet gán được.

## A8. Thang cáp, máng cáp (TCap, MCap)

- **Mã** `{Rộng}x{Cao}x{Dày}-{LD|TR}-{MK|MKN|STD}`. Thang cáp = `LD`, máng cáp = `TR`.
- **Lớp phủ**: trong nhà / không ghi → `MK` (mạ kẽm điện phân); ghi "ngoài trời", "mạ kẽm nhúng (nóng)" → `MKN`; ghi "sơn tĩnh điện" → `STD`.
- **Độ dày**: theo spec nếu có; không có → **1,0 mm**, ghi ứng viên dày hơn trong ghi chú.
- **Độ tin cậy**: BOQ ghi đủ kích thước và lớp phủ → TB; thiếu lớp phủ hoặc độ dày giả định → Thấp.
- **Kích thước không có trong thư viện** (ví dụ 1100x500) → Chưa khớp, không lấy cỡ gần nhất. Phụ kiện thư viện chỉ có Cút (E-) và Tê (T-): côn thu, giảm, thập (ngã 4) → Chưa khớp; "Ngã 2" hiểu là cút (Thấp), "Ngã 3" là tê.
- **Khay cáp** lấy theo máng cáp (TR, có nắp) của thư viện, Thấp. **Trunking** = máng cáp TR, TB.
- **Phụ kiện** (cút, tê, nắp): theo nhóm `PK thang cap` / `PK mang cap` cùng quy tắc.

## A9. Phạm vi cung cấp (A / DC / NSC / PS)

**Nguồn**: ma trận trách nhiệm và phạm vi công việc trong HSMT (TDI25538: `251027-TGC-CT7-P04-THI CONG MEPF-PROJECT MASTER RESPONSIBILITY MATRIX.pdf`, thư mục `1. HSMT/2. Cap nhat SOW 251030`; bản trước ở `1. HSMT 251023/V2 Requirement/MEP- Scope of work`). Dùng bản mới nhất.

| Ký hiệu | Ý nghĩa | Cách tính giá |
|---|---|---|
| `[A]` | Nhà thầu MEP cung cấp và lắp đặt | VL + NC theo thư viện (mặc định) |
| `[DC]` | CĐT hoặc NCC do CĐT chỉ định cấp vật tư/thiết bị, nhà thầu MEP nhận, lắp đặt, đấu nối | Sau khi người dùng xác nhận: **chỉ NC** (+ vật tư phụ nếu có), VL để trống. Chưa xác nhận: giữ VL + NC, gắn cờ |
| `[NSC]` | Thầu phụ chỉ định của CĐT | Giữ giá thư viện làm tham khảo, gắn cờ; cách đưa vào BOQ (giá tạm, phí quản lý) do người dùng quyết |
| `[PS]` | Khoản tạm tính CĐT đưa sẵn | Không gán giá thư viện, ghi theo số CĐT |

**Quy tắc**:
- **Mỗi nhãn DC / NSC / PS phải có trích dẫn**: tên tài liệu + số mục + câu gốc. Không tìm được căn cứ trong HSMT thì là `[A]`.
- **Ghi chú dòng**: `[DC] <tài liệu> mục <n>: <tóm tắt>`; thêm "đã xác nhận chỉ NC" hoặc "chờ xác nhận".
- **Hỏi một lần** ở điểm dừng 1, kèm bảng hạng mục DC/NSC/PS và căn cứ. Áp theo câu trả lời cho cả dự án; ghi quyết định vào Bảng khớp ở Bước 9.
- **Ví dụ đã kiểm chứng (TDI25538)**: ma trận trách nhiệm mục 73 "Sanitary fixtures, Mixers, Kitchen Sink & Bath Accessories … shall be supplied by MH or a supplier nominated by MH"; nhà thầu MEP làm "quantification, coordination, receiving, handling, storage, installation, connection, testing, and commissioning" → thiết bị vệ sinh, vòi, phụ kiện WC là `[DC]`.
- Các nhãn khác từng xuất hiện khi chạy thử (công tắc/ổ cắm, đèn, cáp ≥70mm², RMU, busway) **chưa tìm được căn cứ trong ma trận trách nhiệm**: phải tìm lại trong HSMT trước khi gắn nhãn.

## A10. Ống luồn dây (OLDay)

- **Áp cho**: mọi dòng ống dưới hệ Điện, Điện nhẹ (điện thoại, dữ liệu, CCTV, ACS, âm thanh, bãi xe), Báo cháy, kể cả khi Chủng loại ghi `uPVC`.
- **Mã**: PVC cứng `OLD-PVC-C-D{D}-N` (đi nổi) / `-C` (đi âm); PVC mềm `OLD-PVC-M-D{D}`; thép (G.I) `OLD-THEP-C-DN{D}`. D trong BOQ (D20, D25…) là cỡ ống luồn, không quy đổi sang OD ống nước.
- **Nổi hay âm**: BOQ/tiêu đề ghi "âm tường/âm sàn" → `-C`; ghi "nổi", "trên trần", "hầm", "phòng kỹ thuật" hoặc không ghi → `-N` (mặc định đã duyệt ở TDI25538). Độ tin cậy TB, ghi giả định.
- **Ống G.I mềm / ống luồn thép mềm**: Chưa khớp (thư viện không có).
- **HDPE gân xoắn luồn cáp** (`OLD-HDPE-DN{OD}-{ID}`, BG Santo 2019). BOQ ghi `D{ID}/{OD}` hoặc `D{OD}/{ID}`:

| ID / OD (mm) | Mã |
|---|---|
| 25/32 | `OLD-HDPE-DN32-25` |
| 30/40 | `OLD-HDPE-DN40-30` |
| 40/50 | `OLD-HDPE-DN50-40` |
| 50/65 | `OLD-HDPE-DN65-50` |
| 65/85 | `OLD-HDPE-DN85-65` |
| 80/105 | `OLD-HDPE-DN105-80` |
| 100/130 | `OLD-HDPE-DN130-100` |
| 125/160 | `OLD-HDPE-DN160-125` |
| 150/190 | `OLD-HDPE-DN190-150` |

  Khớp đúng ID/OD: độ tin cậy TB. Lệch OD ≤5mm (ví dụ D195/150 → `DN190-150`): Thấp, ghi độ lệch. Lệch lớn hơn: Chưa khớp, ghi cỡ gần nhất.

## A11. Hệ trung thế: cáp, tủ RMU, công tác tận dụng cáp

Đã kiểm chứng ở TDI25538 (sheet BOQ 01. TBA), kiểm lại thư viện: không có cáp trung thế, tủ RMU/LBS/VCB, relay bảo vệ, hộp đầu cáp trung thế, đơn giá thu hồi cáp.

| Dòng BOQ | Cách xử lý | Độ tin cậy |
|---|---|---|
| Cáp trung thế `Cu/XLPE/PVC-12,7/24kV (3x70)` | `CXV-3x70`, cờ "giá hạ thế, bắt buộc báo giá" | Thấp |
| Cáp ngầm trung thế `…/DATA/PVC-W (1x300)` | `CXV-1x300` (thư viện không có `DSTA-1x300`), cờ như trên | Thấp |
| Thu hồi cáp ngầm trung thế (tận dụng) | Chưa khớp, chỉ NC, chiết tính | Chưa khớp |
| Lắp đặt lại cáp ngầm thu hồi (tận dụng) | Chỉ NC tham chiếu `CXV-1x300` (`"vt": []`), VL = 0 | Thấp |
| Hoàn trả cáp thu hồi cho điện lực | Chưa khớp, chỉ NC/vận chuyển, chiết tính | Chưa khớp |
| Tủ RMU n khoang (n LBS + m VCB + relay) | Bóc tách (Quy tắc 16): khoang LBS ×n, khoang VCB ×m, relay bảo vệ ×m; mỗi con Chưa khớp, YCBG hãng khóa. Tủ RTU, UPS, cảnh báo cáp, chống sét van đã có dòng riêng thì không tính lại | Thấp (cha), con Chưa khớp |
| Hộp đầu cáp trung thế (T-plug, trong nhà) | Chưa khớp, YCBG | Chưa khớp |

Phụ kiện đi kèm các dòng này gộp YCBG với hạng mục chính (Quy tắc 9).

## A12. Bộ quy tắc rà soát họ vật tư (rà lần 2, TDI25538, 15/10/2026)

Chạy cho mọi sheet trước khi để Chưa khớp (Quy tắc 17). Đây là gợi ý đã kiểm chứng, không thay cho tra `TV` (Quy tắc 15).

| Họ | Cách khớp | Độ tin cậy |
|---|---|---|
| Ống gió chữ nhật | `P-WxHxt-RD`; độ dày t theo cạnh dài (0,58 / 0,75 / 0,95 / 1,15); tối đa 3000x1200 | Cao nếu đúng cỡ |
| Phụ kiện ống gió | `E-` (cút), `RS-` (côn), `T-` (tê, theo bề rộng); khóa trùng hoặc cỡ ngoài bảng: ước tính theo m² bằng `P-t-STEEL` / `PK-t-STEEL` (Quy tắc 19, 20) | TB / Thấp |
| Ống gió tròn | `P-DxT-SD`; van chặn lửa `FAD-D` / `FAD-I-D` | Cao |
| Miệng gió, louver | `VF-1L-WxH-RD` (đến ~1200 rộng); `LV-Z` | Cao / TB |
| Van gió, quạt, điều hòa | Có đầu mục nhưng giá = 0: vẫn gán mã (Quy tắc 11), cờ báo giá | Cao, không giá |
| Ống / phụ kiện PVC, PPR, HDPE (hệ nước) | Theo vật liệu ở tiêu đề phụ; phụ kiện ren / gang không khớp sang PVC | Theo A1, A7 |
| Van | `BV-TI`, `BVLS-F`, `RV-F`, `CWM` theo loại và DN | Cao / TB |
| Mặt bích | `F-DNxx-ST/MK`, ĐVT Cái; "cặp bích" → `coef` 2 | TB |
| Khác | Đèn, hộp nối, cáp, tiếp địa, PCCC phụ: xem A2, A4, A5 trước | Theo bảng |

Còn Chưa khớp sau vòng 2: `CT TU DIEN` (516 dòng), RMU / LBS / VCB và relay, thiết bị hãng khóa không có trong thư viện.
