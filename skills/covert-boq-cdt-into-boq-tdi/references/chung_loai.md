# Mã Chủng loại (cột "Bộ lọc / Chủng loại" bên TDI)

Bộ mã chuẩn theo bảng **"Tra cứu từ khóa"** của TDI (file `Tong_hop_Tu_khoa_Chung_loai`, 10/2026). Luật trong `boq_tdi.py` đã được kiểm thử trên 2.940 dòng đã gán tay ở sheet "Chi tiết" (3 BOQ: Dự án số 2, BKY Ninh Bình, Unichems Hải Phòng).

Cột **Nhóm** quyết định công thức đơn giá vật liệu:

- **TB**: thiết bị, nhân với HSTB.
- **VL**: vật liệu thường, nhân với HSVL.
- Đơn vị m / m2 / m3 thì luôn dùng công thức **hao hụt** (HSVL × hệ số hao hụt), bất kể mã.

Script tự chọn công thức. Bảng này để Claude soát và sửa mã. **Chỉ dùng các mã trong bảng này.** Không dùng mã cũ (CAPLO, CAPDN, CSTD, MFĐ, OGI, OGM, MG, VG, ĐHKK, SOL, ODBO, uPVC, ĐP, CAM, RACK, TBAT, VDP, CAMB, ISO, DBK, BON, LOC, PTS, BNL, BEP, ONN, PPL, CUA, CP) hay mã chi tiết riêng (OGEI, OLD GI, OTD, PKT, MK, DAU CAP, GĐ, HG…).

## Điện động lực

| Mã | Nghĩa / từ khóa | Nhóm |
|---|---|---|
| CAP | Cáp điện, dây điện, dây dẫn, cáp điều khiển (tiết diện < 70mm²), cáp báo cháy | VL |
| CAPLON | Cáp trung thế; cáp hạ thế tiết diện pha **≥ 70mm²**; đầu cáp, hộp đầu cáp | VL |
| TĐ | Tủ điện, tủ phân phối, tủ điều khiển, tủ nút nhấn, LP/DB/MSB…, tụ bù, ATS, SPD | TB |
| MCB / MCCB / ACB | Aptomat. Ghi rõ loại thì theo loại. "Aptomat/CB" chưa rõ loại: ≤ 63A là MCB, 64–630A là MCCB, ≥ 800A là ACB, không ghi dòng là MCCB. RCBO, RCCB, ELCB, isolator là MCB | TB |
| RMU | Tủ trung thế, RMU, recloser, LBS, FCO, chống sét van, bộ cảnh báo sự cố cáp | TB |
| MBA | Máy biến áp, trạm biến áp | TB |
| MFD | Máy phát điện **và mọi thứ thuộc hệ cấp dầu/xả khói máy phát** | TB |
| UPS | UPS, ắc quy, bộ nguồn | TB |
| BW | Busway, thanh cái, thanh dẫn điện, plug-in, tap-off, phí gia công co/đầu nối busway | TB |
| TMC | Máng cáp, thang cáp, khay cáp, nắp/cút/tê máng | VL |
| OLD | Ống luồn dây, ruột gà (mọi ống trong hệ Điện, Điện nhẹ, Báo cháy); hộp nối/chia dây, hộp điện, box, đế âm | VL |
| CTOC | Công tắc, ổ cắm | TB |
| DEN | Đèn (LED, highbay, downlight, đường, sân vườn, cao áp, exit, sự cố, đèn báo không), cảm biến chiếu sáng | TB |
| CSTĐ | Chống sét, thu lôi, tiếp địa, cọc/băng đồng, hàn hóa nhiệt, mối hàn, bộ ghép | VL |

## Điện nhẹ

| Mã | Nghĩa / từ khóa | Nhóm |
|---|---|---|
| CAP DN | Cáp quang, cáp mạng, Cat6, UTP, cáp tín hiệu, cáp loa, HDMI, dây nhảy. Cáp điện thường nằm trong hệ Điện nhẹ cũng là CAP DN | VL |
| TBDN | Thiết bị điện nhẹ: loa, âm thanh, điện thoại, camera, CCTV, NVR, rack, patch panel, ODF, PDU, switch, router, firewall, wifi, màn hình, TV, BMS, SCADA, kiểm soát vào ra, máy chủ, phần mềm, máy in, video door phone | TB |
| TBBC | Báo cháy: tủ trung tâm, module, đầu báo khói/nhiệt, nút nhấn, chuông, còi, đèn báo cháy, điện trở cuối | TB |
| KSRV | Thẻ, card, đầu đọc thẻ | TB |

## Cấp thoát nước

| Mã | Nghĩa / từ khóa | Nhóm |
|---|---|---|
| VAN | Van, Y lọc, đồng hồ nước/áp, khớp nối mềm, công tắc áp/dòng chảy/mực, cảm biến áp/mực/nhiệt độ, bình tích áp, bình giãn nở | TB |
| BOM | Máy bơm, cụm bơm, thiết bị lọc nước | TB |
| PPR | Ống PPR và phụ kiện PPR | VL |
| UPVC | Ống uPVC và phụ kiện uPVC; ống nước ngưng; ống PP | VL |
| HDPE | Ống HDPE và phụ kiện HDPE | VL |
| OTH | Ống thép, thép mạ kẽm và phụ kiện thép (co, cút, tê, côn, mặt bích…) | VL |
| INOX | Ống inox | VL |
| OD | Ống đồng | VL |
| Vencap | Ventcap, chụp thông hơi | VL |
| TBVS | Thiết bị vệ sinh và phụ kiện (bồn cầu, lavabo, gương, lô giấy, khay xà phòng, kệ kính), phễu thu sàn, thiết bị bếp | TB |
| TS | Thông tắc sàn (floor clean-out), kể cả nắp / ống thông tắc sàn, mọi vật liệu (uPVC, inox…). Thêm 27/10/2026; trước đây gán TBVS | TB |
| MNN | Máy nước nóng, bình nóng lạnh, bơm nhiệt | TB |
| LLMT | Năng lượng mặt trời | TB |

Co / cút / măng sông / nối ren / côn thu / tê **theo vật liệu ống**: ghi rõ vật liệu thì lấy theo vật liệu, chỉ ghi kích thước thì lấy mã ống ở tiêu đề phía trên.

## Điều hòa, thông gió

| Mã | Nghĩa / từ khóa | Nhóm |
|---|---|---|
| TBDH | Dàn lạnh, dàn nóng, FCU, AHU, VRF, chiller, refnet | TB |
| QUAT | Quạt | TB |
| VCG | Cửa gió, miệng gió, louver, **van gió**, damper, van chặn lửa, van xả áp | TB |
| OG | Ống gió (tôn, mềm, EI) | VL |
| BO | Bảo ôn, cách nhiệt, tiêu âm (khi là dòng riêng) | VL |

## Phòng cháy chữa cháy

| Mã | Nghĩa / từ khóa | Nhóm |
|---|---|---|
| TBCC | Họng, hộp, tủ, trụ chữa cháy, cuộn vòi, lăng, **đầu phun / sprinkler**, bình chữa cháy | TB |
| FM200 | Chữa cháy khí, FM200, bình xịt aerosol, nút ấn xả khí | TB |
| FOAM | Chữa cháy bọt | TB |
| CCL | Chống cháy lan, bịt kín | VL |

## Chung

| Mã | Nghĩa / từ khóa | Nhóm |
|---|---|---|
| VTP | Vật tư phụ, **mọi dòng "Phụ kiện …"**, giá đỡ, giá treo, ty ren, đai, kẹp, đầu cốt | VL |
| CTK | Chi phí, phí; thử áp, vận hành, nghiệm thu; vận chuyển, giàn giáo, xe nâng; sơn, khoan, đục, sleeve; đào đắp, bê tông, móng, gạch, trát, chống thấm; bể, hố ga, bể phốt, cống, nắp, song chắn; băng báo, mốc báo cáp | VL |
| CPK | Thí nghiệm, thử nghiệm, kiểm định | VL |
| EXIT | Sơ đồ, biển báo chỉ dẫn, tiêu lệnh, nội quy | VL |

## Nguyên tắc gán khi mơ hồ

1. **Mô tả chỉ là kích thước** (DN150, Ø 9.5mm) hoặc chỉ là tên phụ kiện ống (Cút DN50) thì xét tiêu đề cha gần nhất. Đọc **chủ ngữ** của tiêu đề, bỏ phần "Cung cấp và lắp đặt…" và phần sau "bao gồm…".
2. **Ống thì xem hệ thống.** Ống trong hệ Điện / ELV / Báo cháy là OLD. Ống âm đất thuộc hệ thoát thải là HDPE, thuộc PCCC là OTH.
3. **Hệ thống "chi phối" thắng mã thiết bị lẻ.** Dưới hệ chữa cháy khí / bình xịt: FM200. Dưới foam: FOAM. Dưới hệ máy phát: MFD. Dưới Hệ thống báo cháy: TBBC (kể cả bộ nguồn, ắc quy). Ngoại lệ (giữ mã riêng): VTP, CTK, CPK, cáp, ống luồn, máng cáp, chống sét, TBCC, TBDN, ống, bảo ôn, CCL, TS.
4. **"Phụ kiện …" luôn là VTP**, kể cả trong hệ ống hay máng cáp. Ngoại lệ: phụ kiện đầu cáp, phụ kiện chống sét / tiếp địa.
5. **Có chữ "thông tắc sàn" (hoặc floor clean-out) thì là TS**, kể cả khi ghi vật liệu (uPVC, inox), mã FCO hay nằm dưới tiêu đề thoát nước / thiết bị vệ sinh. Chỉ thua dòng chi phí, thử áp / thử kín (CTK, CPK) và dòng "Phụ kiện …" (VTP). "Thông tắc" ống đứng, tường… không có chữ "sàn" thì vẫn theo luật cũ.
6. Cáp: tiết diện pha lấy từ chính dòng (vd `4Cx240mm2`, `1x4x(1Cx185)`). Dòng tiêu đề "Cáp, dây điện (Cu/PVC)" không có tiết diện thì từng dòng con tự xét.
7. Nếu vẫn không chắc, để trống và liệt kê cho người dùng. Gán sai còn tệ hơn để trống, vì mã quyết định hệ số giá.
