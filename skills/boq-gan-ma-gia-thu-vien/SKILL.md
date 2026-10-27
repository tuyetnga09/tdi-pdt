---
name: "boq-gan-ma-gia-thu-vien"
description: "Gán mã hiệu và đơn giá thư viện TDI vào BOQ hệ Cơ điện (bảng TV) và hệ Xây dựng (TVGIA, có bóc tách theo TVGIA kèm cột Tên công việc), file tổng hợp đầu mục chưa có trong thư viện, Bảng kiểm tuân thủ. Không tra giá web (dùng boq-gia-tham-khao-web)."
---

# Gán Mã hiệu và Đơn giá từ Thư viện giá vào BOQ (hệ Cơ điện + hệ Xây dựng)

Skill chạy sau `covert-boq-cdt-into-boq-tdi` (BOQ đã có phần TDI bên phải và sheet `2 Thu vien gia`, bảng `TV`). Ứng với **Bước 3–4 và Bước 6** của quy trình lập giá TDI. Mỗi dòng công việc nhận đúng 4 thông tin: **Nhóm VT, Mã hiệu, Đơn giá gốc VL, Đơn giá gốc NC**, cộng 7 cột phụ để người duyệt lọc và tra giá (xem Công cụ) và cột **`Nguồn giá`** (`Giá thư viện` / `Chưa có giá`, có bộ lọc, Bước 6c; giá trị `Giá tham khảo` do skill `boq-gia-tham-khao-web` ghi sau), và ô **Đơn giá gốc VL / NC của dòng đã gán mã là tham chiếu ô trực tiếp tới ô giá trong sheet thư viện**, ví dụ `='2 Thu vien gia'!E2911` (không dùng `HYPERLINK`, không thêm cột, Quy tắc 39). Đầu mục không có trong thư viện **để trống giá**, skill này **không tra web** (Quy tắc 29), và được tổng hợp ra một file Excel riêng (Bước 6d, Quy tắc 32) làm đầu vào cho `boq-gia-tham-khao-web` rồi `boq-tao-ycbg`. Không thêm cột giá khác nếu người dùng không yêu cầu.

**Hai hệ trong một skill.** Mỗi sheet BOQ thuộc một hệ, nhận diện ở Bước 0:
- **Cơ điện (MEP)**: thư viện bảng `TV` (sheet `2 Thu vien gia`), nguồn là **thư mục** thư viện ME của dự án; làm theo Bước 0–9 và Phụ lục A.
- **Xây dựng (XD)**: thư viện sheet `TVGIA` trong file BOQ, nguồn là **một file** thư viện XD của dự án; không có Mã hiệu / Nhóm VT; làm theo **Phần XD** và Phụ lục C (`references/phan_xd.md`).
- Nguồn thư viện của từng hệ dò theo mục "Vị trí thư viện giá" (Quy tắc 31).
- File có cả hai hệ thì xử lý song song, dùng chung điểm dừng và một báo cáo chia hai phần.

Bố cục: nguyên tắc → quy tắc → công cụ → Bước 0–9 (MEP, gồm 6c cột `Nguồn giá` và 6d file tổng hợp đầu mục chưa có trong thư viện) và 9b (làm giàu quy tắc) → báo cáo. Phần còn lại nằm ở `references/` (xem "Bản đồ references"): `phan_xd.md` (Phần XD + Phụ lục C, tra `TVGIA`), `phu_luc_A.md` (bảng tra quy đổi đầy đủ A1–A12, được làm giàu sau mỗi BOQ), `phu_luc_D.md` (Bảng kiểm tuân thủ). Phụ lục B (spec TDI25538), bảng tra XD chi tiết C1/C2 và hệ XD chi tiết lưu trong Project: `claude/SKILL_boq-gan-ma-gia-thu-vien_cap-nhat-vong-2.md`.

## Bắt buộc: chạy đủ mọi quy tắc (đọc trước khi làm bất cứ việc gì)

Chốt 16/10/2026: người dùng yêu cầu **mọi quy tắc trong skill đều phải được chạy**, không bỏ, không làm tắt.
1. **Đọc hết SKILL.md này, rồi đọc hết đúng các file `references/` theo bảng "Bản đồ references" bên dưới** (chốt 21/10/2026 thay cho "đọc hết cả Phụ lục A, C, D ngay từ đầu", để giảm token): mỗi file đã kích hoạt thì đọc **đầy đủ**, không đọc lướt, file bị cắt khi xem thì xem tiếp theo `view_range` đến hết. Quy tắc "mọi quy tắc phải được chạy" ở đầu mục này **không đổi**; chỉ đổi thời điểm nạp.
2. **Lập Bảng kiểm tuân thủ** theo Phụ lục D ngay sau Bước 0 (`KIEM_QUY_TAC.json` ở cloud) và cập nhật sau mỗi bước. Mỗi mục có đúng một trạng thái: `Đã áp` (số dòng / sheet, ví dụ dòng), `Không phát sinh` (đã kiểm bằng cách nào: lọc từ khóa nào, trên sheet nào, ra 0 dòng), `Lệch` (lý do, việc cần làm tiếp), `Chưa làm`.
3. **Không được báo xong khi còn mục `Chưa làm`.** `Lệch` chỉ hợp lệ khi có lý do thật (thiếu công cụ, chờ người dùng xác nhận) và phải ghi vào mục "Bước đã bỏ qua / làm khác skill" (Quy tắc 28). "Không phát sinh" phải có bằng chứng kiểm, không ghi theo cảm nhận.
4. Mỗi điểm dừng báo kèm tình trạng Bảng kiểm đến thời điểm đó (số mục Đã áp / Không phát sinh / Lệch / Chưa làm); báo cáo cuối in đủ bảng (Quy tắc 30).

### Bản đồ references (đọc đầy đủ khi điều kiện kích hoạt)

| File | Nội dung (nguyên văn phần đã tách khỏi SKILL.md) | Khi nào đọc |
|---|---|---|
| `references/phu_luc_D.md` | Phụ lục D: Bảng kiểm tuân thủ D1–D3 | **Mọi lần chạy**, ngay sau Bước 0 (lập `KIEM_QUY_TAC.json`) |
| `references/phu_luc_A.md` | Phụ lục A1–A12: bảng tra quy đổi, từ đồng nghĩa, khớp sai, thang / máng, A9, A10, A11, A12 | Khi file có **sheet MEP**: đọc trước Bước 2 / 4b. File chỉ có sheet XD: ghi A1–A12 = `Không phát sinh (file không có sheet MEP)` |
| `references/phan_xd.md` | Phần XD (hệ Xây dựng, `TVGIA`) + Phụ lục C (C3 khớp sai XD) | Khi Bước 0 nhận diện có **sheet XD**. File không có sheet XD: ghi `Không phát sinh (file không có sheet XD)` như D3 |
| `references/quy_tac_nhom_vt.md` | Quy tắc xếp Nhóm VT | Bước 3 (như cũ) |
| `references/power_query_thu_vien.md` | Dựng sheet thư viện bằng Power Query | Chỉ khi chạy bằng add-in Excel (như cũ) |

Mọi chỗ trong SKILL.md nhắc "Phần XD", "Phụ lục A / C / D", "A1–A12", "C3", "D1–D3" là chỉ vào các file trên.

## Vì sao phải dùng Bảng khớp đã duyệt

Để Claude suy luận lại từ đầu thì kết quả lệch giữa các lần chạy (đã gặp: chạy lại TDI25538 mất mã/giá 136 dòng, đổi 300 dòng, do tự siết quy tắc hoặc tự thêm quy tắc loại giá). Vì vậy: (1) kết quả đã duyệt lưu trong `BANG_KHOP_DA_DUYET.xlsx` và áp lại bằng script; (2) Claude chỉ suy luận cho dòng chưa có trong bảng, ghi qua script; (3) sau khi người dùng duyệt, xuất kết quả vào bảng (Bước 9).


## Nguyên tắc (sai ở đây là sai giá dự thầu)

- **Không bịa mã, không bịa giá.** Mã và giá chỉ lấy từ bảng `TV`; không thấy thì để trống và ghi lý do. **Skill này không tra giá web** (chốt 19/10/2026: tách sang skill `boq-gia-tham-khao-web`, chạy sau skill này): đầu mục không có trong thư viện để trống giá, cột `Nguồn giá` = `Chưa có giá`. Không bao giờ tự ước giá. Mỗi giá có nguồn (tên file + ngày cập nhật, cột `File nguồn`, `Ngày cập nhật`). Có mã thì phải có giá hoặc có ghi chú vì sao không (ngoại lệ: Quy tắc 11).
- **Một file làm việc duy nhất** `<tên BOQ>_TDI.xlsx`; không tạo bản thứ hai trên máy người dùng, không rải file trung gian vào thư mục dự án (giao thức ghi ở mục "Một file duy nhất"). **Ngoại lệ**: người dùng yêu cầu rõ "gửi file mới", hoặc file TDI bị khóa mà không đóng được → giao **file mới** hậu tố `_ra-soat-lan-N`, **không ghi đè file cũ**, nói rõ file nào đúng. **Ngoại lệ có chủ đích**: file tổng hợp đầu mục chưa có trong thư viện (Bước 6d, Quy tắc 32) là file đầu ra riêng cạnh BOQ, không phải bản thứ hai của file làm việc.
- **Thư viện phải đủ và đúng trước khi gán**: `tv_loader.py check` = THIẾU 0 / THỪA 0 / LỆCH GIÁ 0 (Bước 0b). Bảng nạp bằng bản cũ từng thiếu 685 dòng (sheet ẩn "Cua gio", dòng CTKhac không mã, HDPE thiếu ĐVT). Mã dự án của file thư viện không khớp dự án → dừng và hỏi (file trong thư mục thư viện của chính dự án không mang mã thì hợp lệ, xem Quy tắc 31).
- **Không sửa dữ liệu CĐT** (STT, tên, thông số, đơn vị, KL, dòng tổng). Bóc tách chỉ chèn dòng con, không tô màu phần CĐT; cột giá CĐT chỉ ghi ở Bước 8, **mặc định chạy ở mọi BOQ** (Quy tắc 25).
- **Đọc hiểu trước khi khớp** (Bước 2): phân mục, Chủng loại, **Ghi chú CĐT**, spec, bản vẽ, ma trận trách nhiệm HSMT. Dòng lô, lot, gói, HT, 100m, tấn → "Cần chiết tính". Dòng cụm có thành phần rõ → bóc tách (Bước 5).
- **Không gửi gì ra ngoài** (YCBG chỉ soạn nháp). Skill này không truy vấn web (việc đó do `boq-gia-tham-khao-web` làm, và truy vấn ở đó không chứa tên dự án, CĐT, mã dự án, khối lượng). **Không khẳng định điều chưa kiểm chứng**: tổng / thành tiền báo người dùng lấy từ `recalc`. Hệ số và giá chốt là việc của trưởng phòng; quy tắc nghiệp vụ chưa chốt thì hỏi.

## Quy tắc bắt buộc (không được tự nới hoặc tự siết)

1. **Dòng đã có trong Bảng khớp thì dùng nguyên kết quả của bảng.** Không tự đổi mã, hạ độ tin cậy hay xóa giá. Thấy bảng sai → ghi vào "Đề xuất đổi so với Bảng khớp", chờ duyệt (bảng cũng có thể sai, từng gặp 6 dòng ống luồn điện nhẹ gán nhầm ống nước), nên khi rà chéo vẫn đọc cả dòng lấy từ bảng.
2. **Không tự thêm quy tắc loại trừ giá.** Phạm vi A / DC / NSC / PS chỉ áp theo A9: phải có trích dẫn HSMT và người dùng xác nhận; chưa xác nhận thì **giữ mã và giá thư viện**, chỉ gắn cờ trong Ghi chú.
3. **Thang/máng cáp** nhiều độ dày không được để Chưa khớp vì lý do đó (A8).
4. **Ống uPVC** ghi PN12 mà thư viện không có: lấy PN16, Thấp, ghi rõ.
5. **Đầu mục khớp theo tên ở A5** (công tắc vỏ kim loại, ổ cắm sàn / chống nước…) gán theo bảng đó, không để Chưa khớp.
6. **Cáp trung thế** (3.6kV trở lên, 12,7/24kV, 22kV, 35kV): thư viện chỉ có hạ thế 0.6/1kV. **Chốt 15/10/2026: gán tạm cáp hạ thế cùng số lõi và tiết diện** (`CXV-{n}x{s}`; cáp 1 lõi giáp DATA lấy `CXV-1x{s}`), **Thấp**, Ghi chú bắt buộc "Cáp trung thế, thư viện chỉ có hạ thế, giá thấp hơn thực tế nhiều lần, BẮT BUỘC BÁO GIÁ NCC (A11)", vào YCBG. Dự án khác: hỏi lại một lần ở điểm dừng 1.
7. **Ống dưới hệ Điện / Điện nhẹ / Báo cháy / ELV luôn là ống luồn dây** (`OLDay`), không bao giờ gán mã ống nước `P-D…-PVC/PPR/HDPE` (A10).
8. **Mỗi dòng Claude tự gán ghi quy tắc đã áp** trong Ghi chú (ví dụ "A1 PVC DN50→D60, PN12→PN16"); dòng Chưa khớp ghi đã tra nhóm, từ khóa, từ đồng nghĩa nào.
9. **Phụ kiện đi kèm hạng mục Chưa khớp** (giá treo, hộp cắm busway): cùng nhóm YCBG với hạng mục chính, ghi "gộp YCBG với …".
10. **Không làm tắt**: xử lý hết mọi dòng `todo` của mọi sheet; không đủ thời gian thì báo còn bao nhiêu dòng.
11. **Thư viện có đầu mục nhưng giá = 0 (hoặc trống) vẫn gán mã** (chốt 14/10/2026): gán Nhóm + khóa như thường, Ghi chú "THƯ VIỆN CÓ ĐẦU MỤC NHƯNG GIÁ VL = NC = 0 → cần báo giá", vào YCBG. Áp cho van gió VCD/NRD/FD/MFD/PRD, quạt (`TB HVAC`), thiết bị chỉ có giá NC… Ở Bước 7 là "có mã không giá" hợp lệ, liệt kê riêng.
12. **Trước khi để Chưa khớp tra cả bảng `TV`**: theo mã, tên, `Description`, kể cả dòng "(sheet ẩn)" và "Không có mã hiệu - tra theo tên"; Ghi chú nêu đã tra nhóm / sheet nào.
13. **Thu hồi / tận dụng / lắp đặt lại / hoàn trả / di dời / tháo dỡ** chỉ có nhân công, **không lấy giá VL vật tư mới**. Chốt 15/10/2026: "lắp đặt lại" gán NC tham chiếu vật tư cùng cỡ (`"vt": []`, `"nt"`), Thấp; "thu hồi", "hoàn trả" Chưa khớp, vào danh sách chiết tính NC (A11).
14. **Ống HDPE gân xoắn luồn cáp** ghi `D{ID}/{OD}` hoặc `D{OD}/{ID}`: tra bảng A10 sang `OLD-HDPE-DN{OD}-{ID}`; lệch ≤5mm lấy cỡ gần nhất, Thấp, ghi độ lệch. Không ghi "không có cỡ" khi chưa đối chiếu.
15. **A6 và quy tắc ghi sẵn chỉ là gợi ý, không thay cho tra `TV`**: "thư viện không có" chỉ ghi sau khi tra đúng khóa. Cuối Bước 7 rà **nhất quán theo họ vật tư**: dòng Chưa khớp cùng họ và cùng sheet với dòng gán được (HDPE D65/50 gán được mà D100/130 Chưa khớp) phải tra lại, giải thích hoặc sửa.
16. **Dòng cụm CĐT đã liệt kê thành phần** vẫn bóc tách ở Bước 5 **dù thành phần không có trong thư viện** (tủ RMU: khoang LBS, VCB, relay); thành phần không có giá là dòng con Chưa khớp, vào YCBG. Chỉ chèn dòng sau điểm dừng 2 (hoặc sau khi người dùng đã cho tự duyệt bóc tách, xem mục "Điểm dừng chờ xác nhận").
17. **Rà lần 2 theo họ vật tư (A12) trước khi để Chưa khớp**, trên **mọi** sheet: ống gió, cút / côn / tê, ống gió tròn, van chặn lửa, miệng gió, quạt, van, mặt bích, đèn, hộp nối, ống nước PVC / PPR / HDPE. Lần đầu từng bỏ sót ~585 dòng có trong thư viện.
18. **Dòng gộp nhiều đầu mục phải bóc tách ở Bước 5**, mỗi thành phần một mã (thành phần không có trong thư viện là con Chưa khớp, Quy tắc 16). Đã gặp: dàn lạnh + điều khiển có dây; cửa gió + hộp gió; hộp gió + bảo ôn; ống gió tròn + bảo ôn; ống gió / phụ kiện + tiêu âm lót trong; van giảm áp + đồng hồ áp; đồng hồ áp + van bi + ống cong; bồn tắm + sen; đèn cao áp / sân vườn (trụ + cần + đèn); đèn còi + nút nhấn; 2 bơm + tủ; bơm chìm + phụ kiện; rèm ngăn cháy; tủ chữa cháy kèm phương tiện. "Cặp bích" = 2 mặt bích → `coef` 2.
19. **Khóa trùng trong thư viện không gán bằng SUMIFS** (tổng bị cộng dồn), ví dụ tê ống gió `T-1300…1700x0,95-RD`: dùng ước tính diện tích `PK-t-STEEL` hoặc tra theo tên duy nhất, Thấp, ghi "khóa trùng". Không bỏ qua cảnh báo của `write`. (Khóa trùng chỉ về mã nhưng khác tên đã được phân biệt bởi ô `Mã | Tên`, Quy tắc 23.)
20. **Phụ kiện ống gió ước tính theo diện tích luôn là Thấp**, Ghi chú nêu công thức (kích thước, m², đơn giá `P-t-STEEL`/`PK-t-STEEL`, hệ số) và vào danh sách "giá bất thường / cần duyệt" (Bước 6).
21. **Trước khi giao file**: xóa tên vùng rác dạng địa chỉ ô (a129, a130…, làm `recalc` báo Err:539 / #N/A; từng gặp 96 tên); **tham chiếu vòng = 0**, **ô lỗi = 0**; file TDI cũ bị khóa thì giao file mới (ngoại lệ ở Nguyên tắc).
22. **Quy tắc quy đổi ở Phụ lục A không được bỏ hoặc gộp bớt khi cập nhật skill; chỉ được thêm hoặc sửa có lý do.** Sau mỗi BOQ đã duyệt, bổ sung quy tắc mới học được theo Bước 9b.
23. **Ô Mã hiệu ghi `Mã | Tên công việc`** (chốt 15/10/2026, mọi BOQ): mã 0 ghi `(mã 0) | Tên`, đầu mục không mã ghi `(không mã) | Tên`. Hai cột phụ `Mã tra cứu`, `Tên tra cứu` tự tách từ ô này; giá tra theo **Loại VL + Mã + Tên** của chính dòng (mã trùng nhiều tên không còn cộng dồn). Chỉ ghi qua script; ASSIGN.json vẫn truyền mã thuần hoặc `keycol` = `Tên công việc`, script đổi sang dạng mới. File cũ (mã thuần, `(mã 0) Tên`, tên trần) chuyển bằng `migrate2`.
24. **Hệ số Loại VL (Summary B11:G24) nhân vào giá thư viện** (chốt 15/10/2026, mọi BOQ): bảng `TV` có 4 cột P–S (`Giá VL / NC thư viện (thô)`, `HS VL / NC (Loại VL)`); `Giá VL gốc` / `Giá NC gốc` (E, F) = thô × hệ số của đúng Loại vật liệu, nên sửa hệ số NC "Cáp điện" thì mọi giá NC nhóm Cáp điện đổi theo. Hệ số **nhân thêm** vào HSVL / HSNC / HS chung sẵn có. **Sửa giá thư viện ở cột thô (P, Q), không sửa E, F.** Thiếu tên Loại VL trong Summary → hệ số 1 (tên lệch khoảng trắng vẫn khớp). Cài bằng `hsloai`.
25. **Bước 8 (giá TDI → cột giá BOQ CĐT) mặc định chạy ở mọi BOQ** (chốt 15/10/2026, người dùng hỏi "sao không gen ngược giá vào BOQ CĐT"); không chờ yêu cầu. Chỉ ghi ô giá CĐT, báo mọi ô CĐT bị ghi đè để người dùng quyết.
26. **Bước 5 (bóc tách) là bắt buộc, kể cả khi thiếu script `insert`** (chốt 16/10/2026: lần chạy 15/10 thiếu `boq_xml_tools.py` nên Claude ghép nhiều mã vào một dòng thay vì bóc tách, sai Quy tắc 16, 18). Thiếu script không phải lý do "không bóc tách được": Claude tự chèn dòng con ở mức XML theo "Chèn dòng khi thiếu script" (Bước 5), dời **mọi** tham chiếu và kiểm tra lại bằng `recalc` (0 ô lỗi, tổng thành tiền chỉ đổi do giá con mới, giá cha = SUMPRODUCT `Tổng KL tạm tính` con × giá con, Quy tắc 37).
27. **Ghép nhiều vật tư vào một dòng (`terms`) chỉ khi người dùng đồng ý cho từng dòng** (chốt 16/10/2026). Ghi chú dòng ghi "GHÉP – đã được duyệt ngày …, lý do: …". Không tự ghép vì thiếu công cụ, vì muốn nhanh hay vì "khi không bóc tách được"; dòng gộp chưa duyệt ghép thì phải bóc tách (Quy tắc 18) hoặc để `Cụm` không giá và hỏi.
28. **Báo cáo cuối luôn có mục "Bước đã bỏ qua / làm khác skill"** (chốt 16/10/2026): liệt kê mọi chỗ lệch quy trình (thiếu script, thiếu Bảng khớp, bỏ điểm dừng, chưa bóc tách, ghép dòng, chưa áp A9…), lý do và việc cần làm tiếp. Không gộp các lệch này vào ghi chú chung; không có lệch thì ghi "Không có".
29. **Không tra web ở skill này** (chốt 19/10/2026, thay quy tắc 29 ngày 16/10/2026; áp cho mọi BOQ, cả MEP và XD): đầu mục không có trong thư viện sau khi đã tra `TV` / `TVGIA` đủ (Quy tắc 12, 15, 17) giữ độ tin cậy `Chưa khớp`, **để trống giá VL và NC** (không ghi số; ô giá chỉ mang công thức chờ của Bước 6e, ra 0 khi chưa có mã, Quy tắc 34), không gán mã, `Nguồn giá` = `Chưa có giá`, vào danh sách báo giá / YCBG. Giá tham khảo web do skill `boq-gia-tham-khao-web` ghi **sau** skill này và trước `boq-tao-ycbg`. Chạy lại skill này gặp ô `Giá tham khảo` có ghi chú `GIÁ THAM KHẢO WEB` thì giữ nguyên, trừ khi dòng nay đã khớp thư viện: cách xử lý ở Bước 6c.
30. **Mọi quy tắc phải được chạy và chứng minh đã chạy** (chốt 16/10/2026): Quy tắc 1–39, Bước 0–9 (gồm 6b, 6c (cột `Nguồn giá`), 6d, 6e (công thức chờ), 9b khi đã duyệt), Phụ lục A1–A12 và Phần XD / Phụ lục C (khi file có sheet XD) đều có dòng trong Bảng kiểm tuân thủ (Phụ lục D) với trạng thái và bằng chứng. Chạy liền mạch, chạy lại, chạy một phần sheet hay thiếu script đều **không** miễn bảng kiểm. Còn mục `Chưa làm` thì không báo xong, không commit bản cuối.
31. **Thư viện lấy theo thư mục dự án trước** (chốt 16/10/2026, mọi BOQ): ME là thư mục ứng viên `thu vien gia` cùng cấp file BOQ, XD là file ứng viên `thu vien gia … XD` cùng cấp file BOQ; không có mới về thư mục chung / `01. File nội bộ TDI.xlsx`. File XD phải qua kiểm bố cục (`TVGIA`, cột B / G / J) trước khi dùng. Script và Bảng khớp vẫn lấy ở `_Claude_tools` thư mục chung. Chi tiết ở mục "Vị trí thư viện giá".
32. **Xuất file tổng hợp đầu mục chưa có trong thư viện** (chốt 16/10/2026, sửa 19/10/2026; mọi BOQ, cả MEP và XD): sau `recalc` (Bước 7 mục 1), mỗi lần chạy xuất **một file Excel riêng** `<tên BOQ>_Dau_muc_chua_co_trong_thu_vien_gia_tham_khao.xlsx` vào thư mục con `Claude outputs` cạnh file TDI; nội dung, sheet phụ và cách ghi ở Bước 6d. **Lần xuất của skill này không có giá**: các cột giá tham khảo, hãng / model, khoảng giá, VAT, nguồn, link, thứ hạng, ngày tra, truy vấn **để trống** nhưng **giữ nguyên tiêu đề, vị trí và tên cột** (vì `boq-gia-tham-khao-web` và `ycbg_tools.py read` đọc theo tiêu đề); `Tình trạng giá` = `Chưa tra giá web`. Đây là file đầu ra riêng, ngoại lệ có chủ đích của Nguyên tắc "một file"; không sửa file TDI.
33. **Hệ XD: bóc tách theo TVGIA và cột `Tên công việc (TVGIA)`** (chốt 16/10/2026, người dùng yêu cầu sau khi thấy sheet XD không có dòng con như thư viện): sau khi gán và được duyệt, mọi sheet XD có **cột N "TÊN CÔNG VIỆC (TVGIA)" ngay sau Nội dung (TDI)** có droplist, và **mọi dòng BOQ gán vào đầu mục TVGIA có con được bóc tách đủ các cấp** của thư viện. Dòng con (Loại dòng = `Con`) **chỉ để xem, không cộng vào tổng**; **tổng XD trước / sau bóc tách phải bằng nhau**; tổng các con khác giá cha thì **giữ giá cha**, ghi chênh vào Ghi chú và báo người dùng. Cách làm ở Phần XD ("Bóc tách XD theo TVGIA"). Chuỗi chạy lại có bước ghi giá thì `xd_restruct.py` luôn chạy **cuối cùng**.
34. **Dòng chưa có giá vẫn phải nối công thức chờ để chọn mã từ droplist là giá tự lên** (chốt 20/10/2026, người dùng gặp ở 270726_HH4: chọn mã ở dòng Chưa khớp mà giá không đổi; áp cho mọi BOQ MEP từ nay). Mọi ô Đơn giá gốc VL / NC đang trống của dòng công việc được ghi **công thức chờ** có chặn mã trống (chưa có mã ra **0**, có mã tra `TV` như công thức chuẩn), và công thức SUMIFS chuẩn của dòng đã gán được bọc chặn mã trống. Không ghi số, không đổi Độ tin cậy, không gán mã. Cách làm và kiểm ở Bước 6e, Bước 7 mục 14.
35. **Tô xám nhạt ô Đơn giá VL (đã nhân hệ số) khi ĐVT của CĐT là m, md, m2, m3, kg** (chốt 22/10/2026, người dùng yêu cầu; áp cho mọi BOQ, cả MEP và XD). Điều kiện xét trên **ĐVT của CĐT** ở dòng đó (cột Đơn vị phần CĐT, không phải ĐVT TDI của dòng con), so sánh không phân biệt hoa thường, bỏ khoảng trắng thừa, chấp nhận dạng viết khác: `m2` = `m²` = `M2`, `m3` = `m³`, `md` = `m dài` = `mdai`, `kg` = `Kg` = `KG`. Dòng thỏa thì tô **nền xám nhạt `D9D9D9`** đúng **một ô: Đơn giá VL sau hệ số của dòng đó** (cột `ROUNDDOWN(giá gốc × HSVL × HS chung)`, nhận theo tiêu đề "Đơn giá" / VL, không theo chữ cái; Quy tắc 24 và mục 0f). Không tô ô Đơn giá NC, Tổng, Thành tiền, không tô phần CĐT (Nguyên tắc: không tô màu phần CĐT). Dòng `Con`, dòng tiêu đề, dòng tổng, dòng `Cụm` có ĐVT CĐT ngoài 5 đơn vị trên (ví dụ bộ, HT, cái, lô, tấn, 10m, 100m) **không tô xám**; ĐVT CĐT trống thì không tô. **Dòng thiết bị tô xanh dương theo Quy tắc 36, không tô xám** (thứ tự ưu tiên ở Quy tắc 36). Không đè lên màu xanh `E2EFDA` của giá tham khảo (cột khác, không xung đột). Thời điểm và công cụ: theo Quy tắc 36 (tô lần 1 trước điểm dừng 2, tô lại sau Bước 5 / 6e). Ô tô chỉ là dấu nhận biết: **không đổi giá trị, công thức, hệ số**. **ĐVT CĐT** ở đây và ở Quy tắc 36 luôn là cột Đơn vị của phần **CĐT** (cột H ở BOQ du toan so 2), **không** phải cột Đơn vị TDI, vì cột Đơn vị TDI từ nay lấy theo thư viện (Quy tắc 38) và có thể khác ĐVT CĐT. Kiểm ở Bước 7 mục 15.
36. **Tô màu ô Đơn giá VL (đã nhân hệ số) theo loại dòng: xanh dương cho thiết bị, xám cho ĐVT m / md / m2 / m3 / kg, trắng cho phần còn lại** (chốt 22/10/2026, người dùng yêu cầu; áp cho mọi BOQ MEP, đã thử ở DA3139_HH4; bổ sung vòng 2 ngày 22/10/2026 sau BOQ du toan so 2). **Đầu mục thiết bị** = dòng công việc mà công thức Đơn giá VL sau hệ số dùng **HCTB** (HSTB, `Summary TDI` dòng 8, cột D, ví dụ 1,1) thay cho HSVL (ví dụ 1,25), tức Chủng loại thuộc nhóm thiết bị (mục 0f: "HCTB với thiết bị theo Chủng loại"). **Nhận diện theo công thức của chính dòng** (công thức có tham chiếu ô HSTB: `$AQ$3` ở B2.1 MEPF, `$D$8` của `Summary TDI` ở các sheet còn lại, `$U$3` ở BOQ du toan so 2; vị trí ô lấy theo tiêu đề / hàng hệ số của sheet), **không đoán theo tên đầu mục**. Ví dụ Chủng loại thiết bị ở HH4: VAN, TĐ, TBDN, VG, MG, QUAT, BW, CTOC, KSRV, TBBC, DEN, ĐHKK, TBAT, CAM, MBA, PTS, LOC, CAMB, BOM, BON, TBVS. Chủng loại vật tư dùng HSVL nên không phải thiết bị (CAP, uPVC, OGI, TMC, CAPLO, OLD, CAPDN, PPR, HDPE, BO, OTH, VTP và dòng không có Chủng loại). Sheet XD không có HCTB thì quy tắc xanh dương không phát sinh, ghi `Không phát sinh` kèm cách kiểm. **Thứ tự ưu tiên cho từng dòng công việc** (có ĐVT, không phải dòng `Con`, tiêu đề, tổng): (1) thiết bị → nền **xanh dương `9BC2E6`**; (2) không phải thiết bị và ĐVT CĐT thuộc m, md, m2, m3, kg → nền **xám nhạt `D9D9D9`** (Quy tắc 35); (3) **còn lại → nền trắng `FFFFFF`**, kể cả khi ô đang mang nền khác của template (ở HH4 cả cột đơn giá VL mang nền xanh nhạt `DEEBF7` sẵn, nên dòng còn lại phải đặt về trắng; không dùng `9BC2E6` trùng `DEEBF7` gây nhầm). Chỉ tô **đúng một ô** là Đơn giá VL sau hệ số của dòng; dòng tiêu đề / nhóm / tổng giữ nguyên nền template; không tô ô NC, Tổng, Thành tiền, phần CĐT. Tô bằng cách thêm fill và xf vào `styles.xml` (nhân bản xf cũ của ô, chỉ đổi `fillId`, giữ số định dạng, font, viền), không dùng openpyxl ghi lại BOQ. **Không đổi giá trị, công thức, hệ số.** **Dòng `Cụm`** có công thức dùng HCTB tính là thiết bị (xanh dương); dòng `Con` không có ô Đơn giá sau hệ số nên không tô. **Lệch giữa Chủng loại và công thức** (đã gặp ở BOQ du toan so 2: 4 dòng Chủng loại MNN, LLMT dùng HCTB): **theo công thức**, tô xanh dương, đồng thời **liệt kê các dòng đó** (Chủng loại, dòng, tên) ở mục "Cần người dùng xác nhận" của báo cáo để người duyệt quyết giữ hay đổi hệ số / Chủng loại; không tự đổi công thức. **Thời điểm** (chốt vòng 2, 22/10/2026; lý do: ở điểm dừng 2 người dùng mở file vẫn thấy cả cột nền xanh nhạt của template và hỏi vì sao "xanh hết"): tô **lần 1 ngay sau Bước 6 và trước điểm dừng 2** (số dòng chưa đổi), để file giao ở điểm dừng 2 đã đúng màu; tô **lại lần 2** sau Bước 5 / 6e / `fillcdt` khi có dòng mới (chạy lại được, kết quả giống từng byte vì chỉ đặt `fillId`). **Công cụ**: chưa là lệnh của `boq_xml_tools.py`; dùng script mẫu `colour.py` trong Project (`claude/boq/colour.py`, chạy `python3 colour.py FILE OUT`, sửa hằng `S` thành tên từng sheet MEP hoặc lặp qua mọi sheet MEP đã xử lý; nhận thiết bị theo `$U$3` trong công thức, đổi theo ô HSTB của sheet), thiếu thì tự viết bằng lxml theo đúng quy tắc này và ghi vào "Bước đã bỏ qua / làm khác skill". Kiểm ở Bước 7 mục 16.
37. **MEP: dòng `Con` có `Tổng KL tạm tính` = `Hệ số ĐV`, và giá dòng cha `SUMPRODUCT` theo cột `Tổng KL tạm tính`** (chốt 22/10/2026, người dùng yêu cầu; áp cho mọi BOQ MEP). (a) Sau bóc tách (Bước 5), ở **mọi dòng `Loại dòng` = `Con`** của sheet MEP, ô **`Tổng KL tạm tính`** của phần TDI (cột N+8, nhận theo tiêu đề "Tổng KL" của khối TDI, không theo chữ cái) ghi **số tĩnh** bằng giá trị ô `Hệ số ĐV` của chính dòng đó (ví dụ `Hệ số ĐV` = 3.535 thì `Tổng KL tạm tính` = 3.535). Ô `Hệ số ĐV` là công thức thì lấy giá trị đã tính. **Không nhân KL cha**, không làm tròn, không ghi công thức. (b) Ở **mọi dòng `Cụm`**, công thức giá gốc VL / NC đổi vùng hệ số từ cột `Hệ số ĐV` sang cột `Tổng KL tạm tính` của các dòng con, giữ nguyên vùng giá và kiểu tham chiếu: ví dụ `=SUMPRODUCT($BA$188:$BA$189,S188:S189)` (BA = Hệ số ĐV, S = giá gốc VL) thành `=SUMPRODUCT($<Tổng KL tạm tính>$188:$<Tổng KL tạm tính>$189,S188:S189)`. Chỉ đổi đúng vùng trỏ cột `Hệ số ĐV`; phần khác của công thức (bọc IF, nhân hệ số) giữ nguyên. Làm (a) trước (b): ô `Tổng KL tạm tính` của con trống thì SUMPRODUCT coi là 0 và giá cha bị hụt, nên dòng con nào chưa có số thì không đổi công thức cha đó, liệt kê. Vì hai cột bằng nhau, **giá cha, tổng thành tiền TDI, CĐT, Summary trước = sau**. Từ đây `Tổng KL tạm tính` là ô điều khiển giá cha: người duyệt sửa định mức ở ô này thì giá cha đổi theo; cột `Hệ số ĐV` giữ nguyên làm giá trị gốc lúc bóc tách (sửa `Hệ số ĐV` thì không đổi giá, nhắc người dùng sửa ở `Tổng KL tạm tính`). Không ghi vào `Tổng KL tạm tính` của dòng `Dòng` / `Cụm` / tiêu đề / tổng (giữ công thức template `=+<Tổng KL CĐT>`), không động vào phần CĐT, không thêm thành tiền cho dòng con. Ô `Tổng KL tạm tính` của dòng con đang trống hoặc là công thức template trỏ Tổng KL CĐT của chính dòng thì ghi đè; ô đang là số khác `Hệ số ĐV` (người dùng tự nhập) thì **không ghi đè** (giá cha sẽ theo số đó), liệt kê để người dùng quyết. Ghi bằng lxml giữ style (định dạng số theo ô `Hệ số ĐV` của dòng đó). **Công cụ** (vòng 2, 22/10/2026): `insert` chưa làm bước này; dùng script mẫu `qt37.py` trong Project (`claude/boq/qt37.py`, chạy `python3 qt37.py FILE OUT` ngay sau `insert`, trước `dropdown` / `cho`, vì 6e không động vào dòng `Cụm`; sửa hằng `S` theo tên sheet hoặc lặp qua mọi sheet MEP; script đã kiểm ở BOQ du toan so 2: 8 dòng con, 8 công thức cha đổi, tổng trước = sau), thiếu thì tự viết bằng lxml và ghi vào "Bước đã bỏ qua / làm khác skill". Chạy lại mỗi lần `insert`. Hệ XD không áp (XD có cơ chế riêng ở Quy tắc 33). Kiểm ở Bước 7 mục 17.
38. **MEP: cột `Đơn vị` của phần TDI lấy ĐVT từ thư viện (`TV[ĐVT]`) theo khóa của chính dòng; đổi mã thì ĐVT đổi theo** (chốt 22/10/2026, người dùng yêu cầu sau khi thấy cột Đơn vị TDI chỉ chép ĐVT của CĐT `=+H`; áp cho mọi BOQ MEP, cùng cơ chế với Đơn giá gốc VL / NC). (a) **Công thức**: ở mọi dòng `Loại dòng` = `Dòng` hoặc `Con`, ô Đơn vị TDI (cột R ở BOQ du toan so 2, nhận theo tiêu đề "Đơn vị" của khối TDI, không theo chữ cái) ghi `=IF(OR($P{n}="",AND($BC{n}="",$BD{n}="")),<DỰ PHÒNG>,IFERROR(IF(INDEX(TV[ĐVT],<MATCH>)="",<DỰ PHÒNG>,INDEX(TV[ĐVT],<MATCH>)),<DỰ PHÒNG>))`, trong đó `<MATCH>` = `MATCH(1,INDEX(((($AX{n}="")+($AX{n}="Chưa xác định")+(TV[Loại vật liệu]=$AX{n}))>0)*((($BC{n}="")+(TV[Mã hiệu]=$BC{n}))>0)*((($BD{n}="")+(TV[Tên công việc]=$BD{n}))>0),0),0)` (cùng khóa Loại VL + Mã tra cứu + Tên tra cứu với công thức giá của Quy tắc 23, ví dụ P = Mã hiệu, AX = Nhóm VT, BC / BD = Mã / Tên tra cứu; vị trí cột lấy theo tiêu đề). Dòng chưa có mã (ô Mã hiệu trống) hoặc không tra ra ĐVT thì hiện `<DỰ PHÒNG>`; chọn mã khác ở droplist thì ô Đơn vị đổi theo mã mới, không cần chạy lại script. (b) **Giá trị dự phòng**: dòng `Dòng` = `+<ô ĐVT CĐT của dòng>` (đúng như template cũ, nên dòng chưa gán mã vẫn hiện ĐVT CĐT); dòng `Con` = chữ ĐVT tĩnh `insert` đã ghi trước đó (chuỗi trong dấu nháy kép, thoát `"` thành `""`). Dòng `Cụm`, tiêu đề, tổng giữ nguyên. Ô Đơn vị đã là công thức tra `TV[ĐVT]` (chạy lại) thì giữ dự phòng cũ, chỉ ghi lại công thức. (c) **Hao hụt**: công thức cột KL (AP ở BOQ du toan so 2) `=IF(OR(R="m",R="m2",R="m3"),W*$W$3,W)` của dòng `Dòng` đổi tham chiếu `R{n}` sang **ĐVT CĐT** (`H{n}`), vì `Tổng KL tạm tính` của dòng `Dòng` đang là KL theo ĐVT CĐT; nếu không đổi thì ĐVT thư viện khác (ví dụ `md` ↔ `m`) làm hao hụt bật / tắt sai và tổng đổi. Dòng `Con` giữ tham chiếu ĐVT TDI (KL con tính theo `Hệ số ĐV` theo ĐVT thư viện). (d) **ĐVT thư viện khác ĐVT CĐT** (so không phân biệt hoa thường, bỏ khoảng trắng thừa; ví dụ CĐT `cặp bích` ↔ thư viện `Cái`, CĐT `cái` ↔ thư viện `Bộ`, CĐT `cái` ↔ thư viện `m`): đơn giá thư viện tính **theo ĐVT thư viện** nhưng nhân với KL CĐT, nên với dòng lệch thật là sai ĐVT thì thành tiền sai. Claude **không tự đổi mã, KL hay `Hệ số ĐV`**; **liệt kê toàn bộ dòng lệch** (số dòng, ĐVT CĐT, ĐVT thư viện, mã) ở mục "Cần người dùng xác nhận" của báo cáo, phân loại: (i) lệch quy đổi đã biết (cặp bích → Cái hệ số 2, Quy tắc 18 và Phụ lục A), (ii) lệch cần xem lại mã (ví dụ cái ↔ m, có thể gán nhầm họ vật tư), (iii) lệch cách gọi (Bộ ↔ cái). Dòng nhóm (ii) ghi cảnh báo "ĐVT thư viện <x> khác ĐVT CĐT <y>" vào Ghi chú gán mã; **không tự đổi mã và không tự đổi Độ tin cậy** của dòng lấy từ Bảng khớp hoặc đã duyệt (Quy tắc 1), chỉ đề xuất ở "Đề xuất đổi so với Bảng khớp". (e) **Không đổi giá trị tiền**: Đơn giá gốc, hao hụt, thành tiền, tổng từng sheet, Summary, CĐT **trước = sau** khi đặt công thức (chỉ chữ trong cột Đơn vị đổi). Quy tắc 35 và 36 vẫn đọc ĐVT **CĐT** (cột H), không đọc cột Đơn vị TDI. Không đổi style ô (giữ nền, font, viền). (f) **Thứ tự**: chạy **sau `cho` (Bước 6e)** và **trước `fillcdt`**; chạy lại sau mỗi lần `insert` hoặc ghi lại file có thêm dòng (dòng `Con` mới). **Công cụ**: chưa là lệnh của `boq_xml_tools.py`; dùng script mẫu `donvi.py` trong Project (`claude/boq/donvi.py`, chạy `python3 donvi.py FILE OUT [SHEET…]`, mặc định mọi sheet có cột `Loại dòng` và `Mã tra cứu`; đã kiểm ở BOQ du toan so 2: 589 ô Đơn vị đặt công thức, 0 ô lỗi, giá trị tiền và tổng trước = sau, 52 / 581 dòng `Dòng` có ĐVT thư viện khác ĐVT CĐT), thiếu thì tự viết bằng lxml theo đúng quy tắc này và ghi vào "Bước đã bỏ qua / làm khác skill". (g) **Skill `covert-boq-cdt-into-boq-tdi`** vẫn ghi `=+H` khi dựng template (đó là dự phòng hợp lệ); Quy tắc 38 chạy ở skill này nên không cần sửa skill đó. (h) **Hệ XD**: áp dụng chỉ khi sheet `TVGIA` có cột ĐVT của đầu mục (khóa theo tên công việc như Phần XD); không có cột ĐVT thì ghi `Không phát sinh` kèm cách kiểm. Phụ lục D chưa có dòng QT38: thêm khóa `QT38` vào `KIEM_QUY_TAC.json` theo mẫu QT37. Kiểm ở Bước 7 mục 18.
39. **Ô Đơn giá gốc VL và Đơn giá gốc NC của dòng đã gán mã là THAM CHIẾU Ô TRỰC TIẾP tới ô giá trong sheet thư viện, ví dụ `='2 Thu vien gia'!E2911`; KHÔNG dùng `HYPERLINK`, KHÔNG thêm cột riêng** (chốt 22–23/10/2026 dưới dạng ô link bấm để nhảy; **sửa 27/10/2026** theo yêu cầu người dùng: ô giá phải là công thức tham chiếu ô như trên, không phải link để bấm; áp cho mọi BOQ, cả MEP và XD). (a) **Cách làm (MEP)**: ở dòng `Loại dòng` = `Dòng` hoặc `Con` **đã gán mã**, thay công thức tra `SUMIFS(TV[Giá VL gốc]…)` bằng `='<sheet TV>'!E{n}` (VL, cột `Giá VL gốc`) và `='<sheet TV>'!F{n}` (NC, cột `Giá NC gốc`), với `{n}` là số dòng của dòng `TV` khớp khóa **Loại VL + Mã tra cứu + Tên tra cứu** của chính dòng (cùng khóa Quy tắc 23, 38). E / F là giá thư viện đã nhân hệ số Loại VL (Quy tắc 24) nên số hiển thị không đổi; phần `*hệ số` đứng sau `SUMIFS` (nếu có) giữ nguyên: `='2 Thu vien gia'!E2911*2`. Sửa giá thư viện vẫn ở cột thô P / Q. Tên sheet thư viện lấy theo tên thật, luôn đặt trong dấu nháy đơn. Bấm F2 vào ô thấy đúng công thức tham chiếu, Ctrl+[ nhảy tới ô nguồn. Không thêm cột, không chèn cột, không dời gì. (b) **Cách xác định `{n}`**: script `xemnguon.py` tìm bảng `TV` theo tên bảng, lấy khóa của dòng từ **giá trị đã tính** của ba ô Nhóm VT, Mã tra cứu, Tên tra cứu (Nhóm VT trống hoặc `Chưa xác định` thì không lọc Loại VL; so không phân biệt hoa thường, bỏ khoảng trắng thừa), đòi **đúng 1 dòng TV khớp** và kiểm `giá TV × hệ số` bằng **giá đang hiện** (lệch quá 0,5 đồng thì giữ nguyên công thức cũ và báo). Cần `recalc` trước (ô phải có giá trị đã tính). (c) **Dòng không đổi sang tham chiếu** (giữ nguyên, có liệt kê lý do): dòng **chưa có mã** (công thức chờ Bước 6e, giữ `SUMIFS` để người duyệt chọn mã ở droplist vẫn ra giá); khớp 0 hoặc nhiều hơn 1 dòng TV (khóa trùng, `SUMIFS` đang cộng dồn); giá TV lệch giá đang hiện; công thức lạ hoặc có phép tính khác ngoài `*hệ số`; công thức dạng chung (shared / array); dòng `Cụm`, tiêu đề, tổng; giá tham khảo web (`E2EFDA`), giá NCC, số cứng của người dùng. (d) **Hệ quả bắt buộc báo người dùng**: ô đã là tham chiếu trực tiếp thì **đổi mã ở droplist, giá KHÔNG tự đổi** (khác công thức chờ `SUMIFS`). Đổi mã xong thì chạy lại `xemnguon.py` hoặc nhờ Claude chạy lại (script tính lại dòng TV theo khóa mới). Dòng chưa có mã vẫn chọn mã là ra giá. (e) **Giữ vị trí bảng `TV`**: tham chiếu bám theo vị trí ô. Chèn / xóa dòng trong Excel thì Excel tự dời đúng; nhưng **nạp lại `TV` (`tv_loader load --replace`), Refresh Power Query hoặc sắp xếp lại bảng `TV`** làm đổi thứ tự dòng thì tham chiếu trỏ sai ô hoặc `#REF!`: sau các thao tác đó **phải chạy lại `xemnguon.py`** (nó tính lại `{n}` theo khóa; ô đã là tham chiếu trực tiếp cũng được tính lại), rồi `recalc` và đối chiếu giá trước / sau. Không đổi tên hay xóa sheet thư viện. (f) **Chạy lại / chuyển từ bản cũ**: ô còn dạng `HYPERLINK` (bản 22/10/2026) được gỡ lớp bọc rồi đổi sang tham chiếu trực tiếp trong cùng lần chạy; ô đã là tham chiếu trực tiếp được tính lại theo khóa hiện tại. (g) **Định dạng**: giữ nguyên style, màu nền theo Quy tắc 35, 36 và định dạng số của ô giá; không đổi font. Giữ nguyên `<v>` đã tính nên **không cần `fullCalcOnLoad`** cho riêng quy tắc này (workbook đã có cờ thì để nguyên). (h) **Không đổi giá trị tiền**: Đơn giá gốc, thành tiền, tổng từng sheet, Summary, CĐT **trước = sau**. (i) **Thứ tự**: chạy **sau `cho` (Bước 6e), `donvi` và `colour`**, **trước `fillcdt`**; chạy lại khi ghi lại file có dòng mới, khi công thức giá bị ghi đè (ví dụ `apply` / `migrate2`), khi người duyệt đã đổi mã ở droplist, hoặc khi `TV` bị nạp lại. (j) **Công cụ**: script `xemnguon.py` bản 27/10/2026 ở `2.Thu vien gia\_Claude_tools\` (thiếu hoặc là bản cũ còn `HYPERLINK` thì lấy bản trong Project `claude/boq/xemnguon.py`): `python3 -I xemnguon.py SRC.xlsx DST.xlsx [--dry]` (`--dry` chỉ báo số ô và lý do giữ lại, không ghi); sửa trực tiếp OOXML bằng lxml, không dùng openpyxl ghi lại BOQ; tự nhận sheet BOQ theo tiêu đề dòng 4 (Nhóm VT, Loại dòng, Mã tra cứu, Tên tra cứu) và cột giá theo nội dung công thức. Khi ghi vào thư mục người dùng, nếu file đang mở trong Excel thì OneDrive sinh bản `-1`: nhờ người dùng đóng file rồi stage lại trước khi ghi. Phụ lục D thêm khóa `QT39` vào `KIEM_QUY_TAC.json` theo mẫu QT37. Kiểm ở Bước 7 mục 19. (k) **Hệ XD** (chưa có script): dòng đã có số trong cột `Dòng TVGIA` ghi `=TVGIA!G{r}` (VL) và `=TVGIA!J{r}` (NC) với `{r}` là giá trị cột đó (không hệ số); dòng chưa có số giữ công thức cũ; tự viết bằng lxml rồi ghi vào "Bước đã bỏ qua / làm khác skill".

## Điểm dừng chờ xác nhận

Mỗi điểm dừng báo kèm tình trạng Bảng kiểm tuân thủ (Phụ lục D) đến thời điểm đó.

Người dùng có thể cho phép chạy liền mạch trong prompt; khi đó chỉ dừng ở điểm 4 và khi gặp điều kiện dừng bắt buộc (mã thư viện lệch, vật liệu không suy ra được). Yêu cầu chung kiểu "chạy skill rồi làm tiếp X" **không** phải là cho phép chạy liền mạch: vẫn dừng ở điểm 2. Chạy liền mạch vẫn phải bóc tách đủ (Quy tắc 26) và báo kế hoạch bóc tách trong báo cáo.

**Tự duyệt bóc tách** (chốt 22/10/2026, vòng 2: người dùng nói "không cần phải duyệt bóc tách, tự động bóc tách giúp tôi" ngay ở điểm dừng 2). Câu nói rõ kiểu "không cần duyệt bóc tách", "tự động bóc tách", "bóc tách luôn" là cho phép **riêng cho điểm 2** (không phải mọi điểm dừng): Claude **bỏ chờ** và chạy ngay Bước 5 theo kế hoạch tự lập, vẫn giữ nguyên các điểm dừng khác. Khi tự duyệt: (a) lập kế hoạch theo đúng Bước 5 và Quy tắc 16, 18, 26 (thành phần không có trong thư viện là con Chưa khớp, hàng đồng bộ giữ một dòng kèm lý do, **không ghép dòng**, Quy tắc 27); (b) khi có nhiều ứng viên ngang điểm (ví dụ bồn tắm loại 1 / loại 2 cùng giá) chọn một mã, độ tin cậy **Thấp**, ghi giả định vào Ghi chú; (c) in **kế hoạch và giả định** trong báo cáo cuối, mục "Cần người dùng xác nhận"; (d) ghi "bỏ điểm dừng 2 theo yêu cầu người dùng ngày …" vào "Bước đã bỏ qua / làm khác skill" (Quy tắc 28); (e) vẫn chạy đủ Quy tắc 37, dropdown, `nguongia`, `cho`, Quy tắc 38 (`donvi.py`), tô màu lần 2, `fillcdt`, `recalc`, `kiem`, Bước 6d và đối chiếu tổng. Người dùng nói tự duyệt ở prompt đầu thì áp ngay từ điểm 2 đầu tiên.

1. **Sau Bước 0**: báo thư viện đã nạp (nhóm, số dòng, tên file, ngày cập nhật, mã dự án) và **nguồn thư viện từng hệ** (thư mục dự án / thư mục chung / người dùng chỉ; mục "Vị trí thư viện giá"). Hỏi cặp cột giá, cách xử lý dòng Thấp và cách tính hạng mục DC/NSC (mục A9) nếu chưa chốt.
2. **Sau Bước 6 (đã tô màu lần 1, Quy tắc 36), trước bóc tách**: báo thống kê độ tin cậy, số dòng lấy từ Bảng khớp / Claude tự gán, các dòng cần xác nhận vật liệu và **bảng kế hoạch bóc tách** (dòng cha → các con, hệ số, căn cứ; cụm giữ nguyên vì hàng đồng bộ kèm lý do; dòng đề nghị ghép theo Quy tắc 27). Chờ duyệt, vì bóc tách làm đổi số dòng (trừ khi người dùng đã cho tự duyệt, xem trên). Thiếu script `insert` thì nói rõ ở đây là sẽ tự chèn dòng (Quy tắc 26), không chuyển sang ghép. Báo luôn **số đầu mục Chưa khớp sẽ vào file tổng hợp** ở Bước 6d (số đầu mục duy nhất, số dòng); skill này không tra web, bước đó chạy ở `boq-gia-tham-khao-web` sau khi người dùng duyệt.
3. **Sau Bước 7**: chạy Bước 8 (fill giá sang CĐT) không cần hỏi lại (Quy tắc 25); báo ô CĐT bị ghi đè, chờ người dùng quyết phần đó.
4. **Kết thúc**: dừng trước Bước 5 của quy trình TDI. Chuẩn bị danh sách cho người rà soát (xem phần Báo cáo), gồm file tổng hợp đầu mục chưa có trong thư viện (Bước 6d). Gợi ý bước tiếp theo: `boq-gia-tham-khao-web`, rồi `boq-tao-ycbg`; không tự chạy khi người dùng chưa xác nhận. Bước 9 (lưu Bảng khớp) chỉ làm khi người dùng báo đã duyệt.

## Công cụ và cách ghi file

Cowork (Python + cầu nối máy người dùng) khớp cả sheet, áp Bảng khớp, kiểm tra LibreOffice, đọc spec / bản vẽ PDF; **không dùng openpyxl ghi lại file BOQ** (mất hình, add-in, cấu hình in). Add-in Claude for Excel (Office.js) ghi vào file đang mở, không có Python, mỗi script ~90 giây nên làm từng phân mục 150–300 dòng.

**Script và Bảng khớp** nằm ở thư mục `_Claude_tools` bên trong thư mục thư viện chung `2.Thu vien gia` (đường dẫn dò lúc chạy theo skill `tdi-duong-dan-khong-co-dinh`, không ghi cứng tên user hay ổ đĩa): `boq_xml_tools.py`, `boq_memory.py`, `tv_loader.py` (v2: đọc cả sheet ẩn, giữ dòng không mã / thiếu ĐVT, lệnh `check`, `--replace`), `xemnguon.py` (Quy tắc 39, ô giá = tham chiếu ô trực tiếp tới ô giá trong sheet thư viện, bản 27/10/2026; dự phòng Project `claude/boq/xemnguon.py`), `BANG_KHOP_DA_DUYET.xlsx` (sheet `Khop` kết quả từng dòng, sheet `Cum` cấu hình bóc tách); hệ XD: `boq_xd_tools.py`, `BANG_KHOP_XD.xlsx`. Bản cũ lưu kèm đuôi `_backup`. **Bản 15/10/2026 (chiều)** của 3 script MEP thêm: ô `Mã | Tên`, `migrate2`, `hsloai`, kiểu `amount` của `fillcdt`, `check` so cột thô; script cũ hơn không có các lệnh này. **Script mẫu vòng 2 (22/10/2026)** nằm ở Project, không ở `_Claude_tools`: `claude/boq/qt37.py` (Quy tắc 37), `claude/boq/colour.py` (Quy tắc 35, 36) và `claude/boq/donvi.py` (Quy tắc 38); đề nghị chép vào `boq_xml_tools.py` thành ba lệnh `qt37`, `mau` và `donvi` khi trưởng phòng đồng ý, đến lúc đó dùng bản mẫu.

**Thiếu script** (không lấy được theo cách dưới): vẫn làm đủ mọi bước bằng ghi XML trực tiếp (lxml, giữ style, không dùng openpyxl ghi lại BOQ), riêng bóc tách theo "Chèn dòng khi thiếu script" ở Bước 5; ghi vào mục "Bước đã bỏ qua / làm khác skill" (Quy tắc 28).

**Cách lấy**: máy phòng có thư mục trên → `device_list_dir` (chưa có quyền thì `device_request_folder_access` thư mục `2.Thu vien gia` của thư mục chung, không phải thư mục thư viện của dự án), stage 4 file, chép script vào `/tmp/boqtools/`. Máy khác → script từ Project "Tối ưu hóa quy trình Phòng đấu thầu" (`claude/boq/*.py`), **xin người dùng đính kèm `BANG_KHOP_DA_DUYET.xlsx`**; không có bảng thì báo "chạy không có Bảng khớp, kết quả có thể lệch lần trước".

`X=/tmp/boqtools/boq_xml_tools.py`, `M=…/boq_memory.py`, `T=…/tv_loader.py`. `FILE` và `OUT` trong các lệnh là **file trung gian ở cloud**, chỉ bản cuối mới ghi về file của người dùng theo giao thức bên dưới.

| Lệnh | Dùng khi |
|---|---|
| `python3 $X cols FILE SHEET` | Nhận diện cột TDI theo tiêu đề |
| `python3 $X helpers FILE OUT SHEET…` | Tạo cột phụ còn thiếu trong 7 cột (`Nhóm VT, Độ tin cậy, Ghi chú gán mã, Hệ số ĐV, Loại dòng, Mã tra cứu, Tên tra cứu`) |
| `python3 $X migrate2 FILE OUT [SHEET…]` | File cũ: ô Mã hiệu → `Mã \| Tên`, công thức giá → tra Loại VL + Mã + Tên (giá không đổi, đã kiểm 0 lệch) |
| `python3 $X hsloai FILE OUT` | Thêm cột P–S vào `TV`, E/F = thô × hệ số Loại VL (Quy tắc 24); chạy lại sau `tv_loader load --replace` |
| `python3 $M apply FILE OUT BANG_KHOP.xlsx --project <mã>` | Áp Bảng khớp. Ra `OUT`, `OUT.todo.json` (dòng cần Claude gán), `OUT.plan.json` (kế hoạch bóc tách từ bảng) |
| `python3 $M write FILE OUT ASSIGN.json` | Ghi các dòng Claude tự gán: kiểm tra khóa có trong TV, cảnh báo khóa trùng, sinh công thức chuẩn |
| `python3 $X insert FILE OUT PLAN.json` | Bóc tách: chèn dòng con, dời mọi tham chiếu, tô màu phần TDI |
| `python3 qt37.py FILE OUT` (script mẫu Project) | Quy tắc 37: ghi `Tổng KL tạm tính` dòng con = `Hệ số ĐV`, đổi `SUMPRODUCT` dòng cha sang cột `Tổng KL tạm tính`; chạy ngay sau `insert` |
| `python3 colour.py FILE OUT` (script mẫu Project) | Quy tắc 35, 36: tô ô Đơn giá VL sau hệ số xanh dương / xám / trắng; chạy lần 1 sau Bước 6 (trước điểm dừng 2) và lần 2 sau `insert` / `cho` / `fillcdt` |
| `python3 donvi.py FILE OUT [SHEET…]` (script mẫu Project) | Quy tắc 38: cột Đơn vị TDI = `TV[ĐVT]` theo khóa của dòng (đổi mã là đổi ĐVT), dự phòng ĐVT CĐT / chữ cũ; đổi `R{n}` → ĐVT CĐT trong công thức hao hụt của dòng `Dòng`; chạy sau `cho`, trước `fillcdt` |
| `python3 -I xemnguon.py SRC OUT [--dry]` (`_Claude_tools`, dự phòng Project) | Quy tắc 39: đổi công thức Đơn giá gốc VL / NC của dòng đã gán mã thành tham chiếu ô trực tiếp `='2 Thu vien gia'!E{n}` / `F{n}` (MEP, bảng `TV`; không `HYPERLINK`, không thêm cột), giữ `SUMIFS` ở dòng chưa có mã; tự gỡ `HYPERLINK` bản cũ; hệ XD tự viết theo 39(k); chạy sau `donvi.py` và `colour.py`, trước `fillcdt`, chạy lại sau mỗi lần `insert`, đổi mã ở droplist hoặc nạp lại `TV` |
| `python3 $X fillcdt FILE OUT CFG.json` | Bước 8 (mặc định chạy): đưa giá TDI sang cột giá CĐT |
| `python3 $X recalc FILE --sheets 'A\|B'` | Xóa cache trên bản sao, tính lại bằng LibreOffice; báo ô lỗi, dòng có mã không giá, dòng tổng |
| `python3 $M diff FILE BANG_KHOP.xlsx --project <mã>` | So file với Bảng khớp: dòng MẤT mã/giá, ĐỔI, THÊM |
| `python3 $M export FILE BANG_KHOP.xlsx --project <mã> [--fix FIX.json]` | Bước 9: lưu kết quả đã duyệt vào Bảng khớp (thay dòng cũ cùng dự án) |
| `python3 $X dropdown FILE OUT SHEET…` | Bước 6b: droplist `Mã \| Tên` theo Loại VL của dòng, nới rộng cột Mã hiệu |
| `python3 $X nguongia FILE OUT [SHEET…]` | Bước 6c: tạo cột `Nguồn giá` nếu chưa có, ghi công thức `Giá thư viện` / `Chưa có giá` cho mọi dòng công việc (giữ chữ tĩnh `Giá tham khảo` / `NCC`), nới autoFilter và `_xlnm._FilterDatabase` tới cột này. Chạy lại được (kết quả giống từng byte) |
| `python3 $X cho FILE OUT [SHEET…]` | Bước 6e (Quy tắc 34): công thức chờ cho ô giá gốc VL / NC trống của dòng chưa có mã, bọc chặn mã trống cho SUMIFS chuẩn, ghi công thức `Mã tra cứu` / `Tên tra cứu` nếu trống. Ô trống của dòng **đã có mã** (chỉ NC / DC, thiếu giá VL) **giữ trống** và được liệt kê, vì nối công thức sẽ kéo giá `TV` vào và đổi tổng. Ô công thức khác bị bỏ qua và liệt kê |
| `python3 $X kiem FILE [--json]` | Tự kiểm một lần tính lại: QT21 (ô lỗi, tên vùng rác), QT23, QT29 (`Nguồn giá` + autoFilter), QT34 (ô chờ), droplist 6b. Mục cần phán đoán (QT2, QT7, QT15…) vẫn do Claude kiểm. Thoát mã 1 nếu có mục Lệch |
| `python3 $X cleanup FILE OUT` | Xóa sheet tạm `_TDI_JOB` / `_TDI_TMP_*` |
| `python3 $T check LIBDIR FILE --code <mã>` | Đối chiếu bảng `TV` trong FILE với thư viện nguồn từng dòng: THIẾU / THỪA / LỆCH GIÁ, so cột giá thô P/Q nếu có (`T=/tmp/boqtools/tv_loader.py`) |
| `python3 $T load LIBDIR FILE OUT --code <mã> --mtimes m.json --replace` | Nạp lại toàn bộ bảng `TV` (giữ tên sheet/bảng); dùng khi `check` báo thiếu/lệch; sau đó chạy `hsloai` |

**Chạy `recalc` / `kiem` khi file có liên kết ngoài** (vòng 2, 22/10/2026): `recalc` của LibreOffice báo lỗi khi workbook có `externalLinks` (thư viện nối trực tiếp bằng `tv_link`). Làm trên **bản phẳng** (xóa phần `externalLinks`, công thức có `[n]`) chỉ để tính và kiểm; file giao người dùng giữ nguyên liên kết. Không chạy `tv_loader load --replace` trên file đã nối liên kết (xem skill chuyển template).

**`ASSIGN.json`** (mỗi dòng Claude tự gán):
```json
[{"sheet":"B2_MEP Service","row":120,"g":"PVC","keycol":"Mã hiệu","keyval":"P-D60-PN16-PVC","ma":"P-D60-PN16-PVC",
  "conf":"Thấp","note":"A1 PVC DN50→D60; PN12 không có, lấy PN16","loai":"Dòng"},
 {"sheet":"B2_MEP Service","row":130,"g":"Chưa xác định","conf":"Chưa khớp","note":"Đã tra Van Khoa, PCCC theo 'van cầu'; thư viện không có"}]
```
Ghép nhiều vật tư vào một dòng (**chỉ khi người dùng đã duyệt ghép dòng đó**, Quy tắc 27): dùng `"terms": [[nhóm, "Mã hiệu"|"Tên công việc", giá trị], …]`. Mã 0: `keycol` = `Tên công việc`, script tự đổi sang ô `Mã | Tên` (mã 0: `(mã 0) | Tên`). Dòng chỉ tính NC (hạng mục DC đã được xác nhận, hoặc công tác "lắp đặt lại" theo Quy tắc 13): truyền `"vt": []` và `"nt": [[nhóm, khóa, giá trị]]`.

Quy tắc khi ghi:
- **Chỉ ghi qua script** (`apply`, `write`, `insert`, `fillcdt`, `dropdown`, `hsloai`, `migrate2`): giữ style, thứ tự cột, escape ký tự đại diện, không động vào giá người dùng tự nhập. Thứ tự chuỗi ghi: `helpers` → `hsloai` → `migrate2` (file cũ) → `apply`/`write` → **tô màu lần 1 (Quy tắc 35, 36)** → *(điểm dừng 2, hoặc tự duyệt)* → `insert` → **`Tổng KL tạm tính` dòng con (Quy tắc 37)** → `dropdown` → **cột `Nguồn giá` + bộ lọc (Bước 6c)** → **công thức chờ (Bước 6e)** → **cột Đơn vị TDI từ thư viện (Quy tắc 38, `donvi.py`)** → **tham chiếu ô trực tiếp trong ô Đơn giá gốc VL / NC của dòng đã gán mã (Quy tắc 39, `xemnguon.py`)** → `fillcdt` (chèn dòng trước, droplist, rồi CĐT) → **tô màu lần 2 (Quy tắc 35, 36)**. Bước 6c và 6e có lệnh `nguongia` và `cho` (bản 21/10/2026 của `boq_xml_tools.py`); script cũ hơn không có hai lệnh này thì ghi bằng lxml theo đúng Bước 6c / 6e và ghi vào "Bước đã bỏ qua / làm khác skill". Bước Quy tắc 37, Quy tắc 38 và bước tô màu (Quy tắc 35, 36) chưa có lệnh trong `boq_xml_tools.py`: dùng script mẫu `qt37.py` / `colour.py` / `donvi.py` trong Project (xem bảng lệnh), thiếu thì làm bằng lxml theo đúng quy tắc và ghi vào "Bước đã bỏ qua / làm khác skill". Bước 6d (file tổng hợp) vẫn dựng bằng openpyxl. Giá tham khảo web không ghi ở skill này (skill `boq-gia-tham-khao-web` ghi sau). File tổng hợp (Bước 6d) dựng sau `recalc` ở Bước 7 bằng openpyxl vào một file **mới** (không phải file BOQ, nên không trái quy tắc không dùng openpyxl ghi lại BOQ).
- **Không viết cứng chuỗi Nhóm/Mã trong công thức** khi tham chiếu ô được (xem "Công thức giá" ở Bước 6). Không tự gõ công thức tay; script sinh đúng dạng.
- **SUMIFS cộng dồn khi khóa trùng.** `write` cảnh báo khóa trùng; gặp thì đổi sang tra theo tên, hoặc ghi "mã trùng, cần chọn tay" và không gán giá.
- **Tên vùng (defined names) rác dạng địa chỉ ô** (a129, a130…, thường ẩn) làm `recalc` lỗi: xóa trước khi tính lại (Quy tắc 21). Khóa trùng: Quy tắc 19.
- **Vùng xám ngoài khung** là `view="pageBreakPreview"`, không phải tô màu.
- **Đọc file bằng openpyxl `read_only`** phải gọi `ws.reset_dimensions()`, nếu không các cột phụ mới thêm sẽ bị cắt.
- **`detect_cols` / `read_sheet` (bản 14/10/2026)**: cột "Nội Dung Công Việc" lấy ô ngoài cùng bên phải (CĐT có thể trùng tiêu đề ở cột B); cột Đơn vị CĐT nhận theo tiêu đề "ĐƠN VỊ"/"ĐVT"/"UNIT" (mặc định C), cột "THÔNG SỐ" đưa vào `todo` là `spec`. Luôn chạy `cols` cho từng sheet và kiểm `name` là cột TDI trước khi ghi.

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

### Một file duy nhất: giao thức ghi

**Vì sao.** Đã gặp: Claude ghi bản đã gán mã, người dùng vẫn mở và lưu một bản Excel cũ cùng tên → mất toàn bộ phần vừa gán. Tạo file thứ hai buộc người dùng đoán bản nào đúng nên **cấm**; ghi đúng một file, một lần, có kiểm tra xung đột.

**Quy ước.**
- File sống duy nhất: `<tên BOQ>_TDI.xlsx` trong thư mục dự án (`FILE_USER`). Mọi bản trung gian chỉ nằm ở cloud. **Không** tạo `_da-gan-ma`, `_v2`, `_ban-truoc-boc-tach`… trên máy người dùng; sao lưu = bản stage ban đầu ở cloud và lịch sử phiên bản OneDrive.
- Danh sách Chưa khớp / báo giá, YCBG nháp là file khác, đặt vào thư mục con `Claude outputs` hoặc gửi qua chat, không đặt tên giống BOQ. File tổng hợp Bước 6d cũng đặt ở `Claude outputs`, tên `<tên BOQ>_Dau_muc_chua_co_trong_thu_vien_gia_tham_khao.xlsx`.
- **Ngoại lệ (đã gặp 15/10/2026)**: người dùng yêu cầu file mới, hoặc file TDI bị khóa (`~$`) và không đóng được → dựng từ bản stage, đặt tên `<tên BOQ>_TDI_ra-soat-lan-N.xlsx`, commit cạnh file cũ, **để nguyên file cũ**, báo "file đúng là …, file cũ … để nguyên, đóng file cũ không lưu".

**Trước khi làm (Bước 0).**
1. `device_list_dir` thư mục dự án. Có `~$<tên>_TDI.xlsx` → **dừng, xin đóng file**; không đóng được mà cần kết quả ngay → áp ngoại lệ trên.
2. Ghi `mtimeMs` và kích thước `FILE_USER`, stage về cloud làm bản gốc; mọi bước trung gian dựng từ bản này.
3. Dặn người dùng một lần: trong lúc Claude chạy, đóng file và đừng lưu bản cũ đang mở.

**Khi ghi.** Không commit từng bước; chỉ commit tại điểm dừng hoặc khi xong. Ngay trước commit `device_list_dir` lại: có `~$` → xin đóng; `mtimeMs` / kích thước khác lần stage → người dùng đã lưu thêm, **không ghi đè**: stage lại bản mới, so (số ô độ tin cậy, số sheet, `DS_MaHieu`, ô khác biệt); bản mới chỉ có sửa tay → áp lại bước của mình lên bản mới nhất bằng script; bản mới là bản cũ chưa gán → báo người dùng rồi ghi bản đúng lên chính file đó. Commit **vào đúng `FILE_USER`** với `expectedMtimeMs` của lần stage mới nhất; `rejected` thì không `force`, stage lại. Sau commit `device_list_dir` kiểm kích thước và mtime, báo tên file và thư mục.

**Sau khi ghi.**
- Nhắc: nếu Excel đang mở file này thì **đóng không lưu rồi mở lại**. Lỡ bị lưu đè: khôi phục bằng Lịch sử phiên bản OneDrive hoặc báo Claude áp lại.
- **Mỗi lần stage lại để sửa tiếp, quét vòng tham chiếu và đếm ô độ tin cậy** (covert-boq-cdt-into-boq-tdi mục 4.5): bản của người dùng có thể là bản cũ bị lưu đè, đã gặp 21 vòng quay lại. Vá lại cùng đợt thay đổi mới.
- Gửi BOQ qua `SendUserFile` chỉ khi người dùng yêu cầu; bản đó chỉ để xem.

**Kiểm tra**: LibreOffice dùng lại giá trị cache Excel, nên luôn dùng `recalc`, không bao giờ recalc file giao cho người dùng.

## Bước 0. Điều kiện

**Nhận diện hệ từng sheet trước tiên** (XD.1); sheet XD theo Phần XD, các mục 0a–0f áp cho sheet MEP. Báo ở điểm dừng 1 danh sách sheet theo hệ. **Kiểm file** theo giao thức một file: không có `~$`, ghi `mtimeMs`, stage bản gốc.

**0a. Template TDI**: `cols` chạy được (có cột Nội dung, Mã hiệu, Đơn giá gốc VL/NC). Chưa có → dừng, dùng skill `covert-boq-cdt-into-boq-tdi`. Chưa có đủ 7 cột phụ → `helpers` cho mọi sheet BOQ.

**0b. Thư viện giá.** File có sheet `2 Thu vien gia`, bảng `TV` (cột `Loại vật liệu` = Nhóm VT (trước 15/10/2026 là `Column1`), `Mã hiệu`, `Tên công việc`, `ĐVT`, `Giá VL gốc`, `Giá NC gốc`, `Giá VL chưa CK`, `Giá NC ước tính`, …, `File nguồn`, `Ngày cập nhật`, **P–S: `Giá VL thư viện (thô)`, `Giá NC thư viện (thô)`, `HS VL (Loại VL)`, `HS NC (Loại VL)`**).
- Chưa có: Cowork nạp bằng `tv_loader.py` theo mục 3 của `covert-boq-cdt-into-boq-tdi`, nguồn dò theo mục "Vị trí thư viện giá" (thư mục dự án trước; lọc theo mã, hoặc cả thư mục khi file không mang mã); chỉ có add-in dùng `references/power_query_thu_vien.md`. Báo lại số nhóm, số dòng, mã dự án, ngày cập nhật mới nhất.
- **Kiểm bắt buộc kể cả khi đã có `TV`**: dò nguồn theo mục "Vị trí thư viện giá", stage thư mục thư viện ME đó, `python3 $T check "<thư mục>" FILE --code <mã>` (`--code ""` hoặc bản sao có tiền tố nếu file không mang mã, đúng như lúc nạp). Nguồn lúc kiểm khác nguồn lúc nạp (ví dụ trước nạp từ thư mục chung, nay dự án đã có thư mục riêng) → báo, hỏi dùng nguồn nào trước khi `--replace`. THIẾU / THỪA / LỆCH GIÁ khác 0 → `load … --replace`, `check` lại đến 0/0/0, báo số dòng bổ sung; lệch giá do file nguồn đã cập nhật → hỏi giữ ảnh chụp cũ hay lấy giá mới.
- **Chưa có cột P–S** (file cũ, hoặc vừa `load --replace`): chạy `hsloai` (Quy tắc 24); `check` so cột thô P/Q. Sau đó `migrate2` nếu ô Mã hiệu còn dạng cũ, rồi `dropdown`.
- Mở `xl/tables/*.xml` kiểm `displayName="TV"`; ra `TV_1` (tên vùng `tv` trùng) thì xóa tên vùng rác và đặt lại `TV` (loader v2 tự làm). Tổng hợp khóa trùng do loader in ra, không dùng chúng cho SUMIFS.

**0c. Rà thư viện.** Mã trùng (Nhóm + Mã, ví dụ PCCC `DN50mm` cho Van góc và Cuộn vòi): ô `Mã | Tên` phân biệt được thì giá tra đúng; không phân biệt được thì "mã trùng trong thư viện, cần chọn tay", Thấp, không giá. Mã = 0 / trống (TB Ve Sinh, TB Dien, TB Bao Chay, TB HVAC, PCCC), thường chỉ có giá NC: tra theo tên (A5). Tên công việc trùng trong nhóm: không tra theo tên.

**0d. Cặp cột giá.** Mặc định phương án A `Giá VL gốc` / `Giá NC gốc`, **có điền giá dòng Thấp** kèm "Giá tạm, độ tin cậy Thấp cần duyệt". Chưa xác nhận cho dự án thì hỏi một lần (phương án B: `Giá VL chưa CK` / `Giá NC ước tính`).

**0e. Phạm vi cung cấp (A9).** Thư mục HSMT / HSDT của dự án (ví dụ `03. HSDT`) đang `cloudOnly` → nhờ người dùng chọn "Always keep on this device" rồi mới đọc. Tìm trong HSMT ma trận trách nhiệm / phạm vi (tên file có `RESPONSIBILITY MATRIX`, `Scope of work`, `SOW`, `Exclusion`); lập danh sách DC / NSC / PS kèm số mục rồi hỏi cách tính một lần.

**0f. Hệ số.** `Summary TDI` dòng 8 (HSKL, HSVL, HCTB, HSNC, HS chung, LÀM TRÒN, hao hụt) kéo về **hàng 4** mỗi sheet. Đơn giá VL sau hệ số = `ROUNDDOWN(giá gốc × HSVL (HCTB với thiết bị theo Chủng loại) × HS chung, làm tròn)`, NC tương tự với HSNC. **Giá gốc đã gồm hệ số Loại VL** (Summary dòng 11–24, Quy tắc 24): đổi hệ số ở đó thì giá thư viện của đúng Loại VL đổi theo, rồi mới nhân HSVL/HSNC. Không tự sửa hệ số, chỉ giải thích; hệ số chốt là việc trưởng phòng. **Dòng dùng HCTB chính là đầu mục thiết bị** (Quy tắc 36).

## Bước 1. Nhận diện BOQ và cột

- Xác định cột **theo tiêu đề**, không theo chữ cái (STT, Mô tả, Thông số, Đơn vị, KL theo khu, **Ghi chú CĐT**, Chủng loại). Sheet có số khu khác thì cột dịch theo (B4: TDI bắt đầu ở X, đơn giá sau hệ số ở AZ/BA).
- **Dòng công việc** = có Đơn vị; làm **mọi sheet BOQ**.
- **Chủng loại** do skill chuyển đổi gán, chỉ dùng làm gợi ý họ vật tư (`ODBO`, `ONN`, `CAPLO`, `VTP`, `CP`). Chủng loại `uPVC` ở "Ống PVC D.." hệ điện là gán sai của bước chuyển đổi; quyết theo hệ (Quy tắc 7).

## Bước 2. Đọc hiểu (bắt buộc, cho các dòng trong `todo`)

- **Cây phân mục**: `todo.json` có `heading`, `cl`, `remark`; bổ sung phân mục La Mã / mã con (I.8, III.1.4) và dòng mô tả "Cung cấp và lắp đặt … bao gồm …" (cho biết hệ).
- **Lọc Ghi chú CĐT** có "bao gồm", "gồm", "kèm", "trọn bộ", "chỉ", "không bao gồm", "đã bao gồm", và cả cột Thông số / mô tả nhiều dòng. Liệt kê thành phần (bơm tăng áp "bao gồm bơm; van; tủ điều khiển", đầu báo "kèm đế", module "kèm hộp", RMU "2 LBS + 2 VCB kèm relay") → bóc tách. "Chỉ bao gồm…" / "Không bao gồm…" → không bóc tách, không gộp. Dòng KL = 0 "đã bao gồm theo cụm" → không định giá riêng, dùng làm căn cứ thành phần và cỡ.
- **Xác định hệ**: ống dưới Điện / Điện nhẹ / Báo cháy là ống luồn dây (A10); "Ống G.I" hệ điện là ống luồn thép.
- **Vật liệu cho dòng chỉ ghi DN/Ø**: đọc tiêu đề phụ → Chủng loại → Ghi chú CĐT → **spec HSMT** (`pdftotext -layout`, tìm PP, PPR, HDPE, uPVC, copper, SCH, PN, thickness, insulation, booster); vẫn chưa rõ thì **hỏi**, không đoán.
- **Bản vẽ** khi cần cấu hình cụm: `pdftotext -layout` tìm trang có ký hiệu, `pdftotext -bbox` lấy tọa độ, `pdftoppm -f N -l N -r 180 -x -y -W -H` cắt vùng xem sơ đồ; lệch BOQ thì báo.
- **Dòng cụm / bộ** nhận qua: "cụm", "bao gồm", "kèm", "trọn bộ", dấu "+", ĐVT cụm / bộ / HT, tiêu đề "có lớp cách nhiệt", mô tả gạch đầu dòng. **Phạm vi cung cấp** theo Bước 0e.

## Bước 3. Nhóm VT

- Đọc `references/quy_tac_nhom_vt.md`; chỉ dùng tên nhóm có trong thư viện. Ưu tiên nhóm có mã khớp thật (miệng / van gió, ống thép PCCC, van PCCC, đồng hồ nước); không nhóm nào khớp thì nhóm theo quy tắc tạm, Chưa khớp, ghi "Nhóm theo quy tắc tạm".
- **Không ép vào CTKhac**: công tác xây dựng (đào, đắp, hố ga, bể, gối đỡ) và thiết bị không có nhóm (bơm, quạt, miệng gió, thiết bị mạng…) ghi `Chưa xác định`. Dòng cụm nối nhóm bằng " + ".


## Bước 4. Gán Mã hiệu

**4a. Áp Bảng khớp trước (bắt buộc).**
```bash
python3 $M apply FILE A.xlsx BANG_KHOP_DA_DUYET.xlsx --project <mã dự án>
```
- Khớp theo khóa (mô tả CĐT, ĐVT, tiêu đề phụ, Chủng loại), nới dần khi kết quả duy nhất. Ưu tiên dòng cùng dự án.
- Script tự kiểm tra mã có trong thư viện của dự án này. Mã không còn thì đưa dòng vào `todo` kèm gợi ý.
- Dòng cụm có cấu hình trong sheet `Cum` được đưa vào `A.xlsx.plan.json`.
- **Kết quả từ Bảng khớp là cuối cùng** (Quy tắc bắt buộc 1).

**4b. Claude gán các dòng trong `A.xlsx.todo.json`.**
- Chỉ tìm trong cùng Nhóm VT, so khớp **thông số** (DN/D/Ø, PN, độ dày, tiết diện, lõi, công suất, kích thước).
- Thứ tự áp quy tắc (Phụ lục A): (1) xác định hệ, ống luồn hay ống nước (A10); (2) phân mục (ống); (3) cáp, kể cả trung thế (A2, A11); (4) van; (5) thang / máng cáp; (6) từ đồng nghĩa; (7) tra theo tên, mã 0; (8) không có trong thư viện, sau khi tra `TV` đúng khóa (Quy tắc 12, 15); (9) xác định cụm; (10) phạm vi cung cấp (A9); (11) thu hồi / tận dụng / lắp lại (Quy tắc 13, A11); (12) rà họ vật tư A12 (Quy tắc 17) và tách dòng gộp (Quy tắc 18).
- **Dòng tương tự dòng trong Bảng khớp** (cùng họ, khác cỡ): làm theo cách của Bảng khớp. Mở sheet `Khop` lọc theo từ khóa để xem mẫu.
- Ghi bằng `write` (ASSIGN.json), mỗi dòng có quy tắc đã áp trong ghi chú.

| Độ tin cậy | Khi nào |
|---|---|
| Cao | Khớp loại và mọi thông số chính |
| TB | Khớp loại, BOQ thiếu một thông số, phải giả định (ghi giả định) |
| Thấp | Gần đúng, ứng viên ngang điểm, vật tư thay thế, ĐVT lệch, thư viện không đạt spec, phụ kiện ống gió ước tính theo diện tích hoặc khóa trùng (Quy tắc 19, 20) |
| Chưa khớp | Không có đầu mục tương ứng (ghi lý do) → danh sách cần báo giá; giá để trống (giá tham khảo web do `boq-gia-tham-khao-web` ghi sau), độ tin cậy là Chưa khớp |
| Cần chiết tính | Lô, lot, gói, HT, đơn vị đặc biệt |

- **Không dùng cụm "chiết tính" trong ghi chú của dòng đã có mã và giá.** Viết "phụ kiện tính riêng" thay vì "phụ kiện cần chiết tính".
- **ĐVT lệch 1:1** (bình/bộ): điền giá, độ tin cậy Thấp. Hệ số khác 1 ghi `coef` (script nhân vào công thức).

## Bước 5. Bóc tách cụm thành dòng con (theo mẫu TDI)

**Khi nào**: sau Bước 6 và sau điểm dừng 2 (bước cuối của phần gán vì đổi số dòng; nếu người dùng đã cho tự duyệt bóc tách thì chạy ngay, không chờ). **Kế hoạch** = `plan.json` từ `apply` (cụm đã duyệt) + cụm mới Claude lập, gộp một file, chạy `insert` một lần. **Thành phần cụm mới lấy theo thứ tự**: mô tả và Ghi chú CĐT → dòng KL = 0 "đã bao gồm theo cụm" → spec → sơ đồ bản vẽ → cụm cùng loại trong sheet `Cum`; mỗi thành phần ghi căn cứ.
- **Cụm CĐT đã liệt kê thành phần vẫn bóc tách dù thành phần không có trong thư viện** (Quy tắc 16; RMU 4 khoang: 2 LBS, 2 VCB, relay): thành phần không có giá là con Chưa khớp.
- **Dòng gộp nhiều đầu mục (Quy tắc 18)**: tách con có mã riêng ngay cả khi `Cum` chưa có cấu hình, căn cứ mô tả "A + B", "kèm", "bao gồm". "Cặp bích" → con mặt bích `coef` 2.
- **Hàng đồng bộ không tách** khi thành phần là phụ kiện liền của thiết bị, NCC chào cả bộ (máy biến áp kèm quạt, cảm biến; máy phát kèm tủ đầu cực; điều hòa cục bộ dàn lạnh + dàn nóng + remote; ống gió mềm có bảo ôn; louver kèm lưới chắn côn trùng; bình FM200 đầy đủ phụ kiện; đèn exit kèm bộ lưu điện): ghi lý do giữ nguyên trong kế hoạch ở điểm dừng 2.
- **Không tự thêm vật tư không căn cứ.** Cụm không có cấu hình (van giảm áp, bình lọc, split / multi, tủ báo cháy): Loại dòng `Cụm`, không giá, "cần bóc tách, chưa có cấu hình", hỏi người dùng. **Không tính hai lần** thành phần đã có dòng riêng (bình áp lực, tủ RTU, UPS). Thành phần cần đo bóc (ống gom, phụ kiện): một con không giá, "đo bóc theo bản vẽ".

**`PLAN.json`**: `[[sheet, parent_row, đoạn đầu mô tả cột B, [con…], ghi chú cha, độ tin cậy cha], …]`; mỗi con `g, key, val, ma, name, unit, coef, conf, note` (`coef` = số lượng thành phần cho 1 đơn vị cha); con Chưa khớp: `g`=`Chưa xác định`, `key`/`val`/`ma`=null, `conf`=`Chưa khớp`.

`insert` làm: dòng con chỉ ghi phần TDI (tên, mã dạng `Mã | Tên`, ĐVT, giá SUMIFS tham chiếu ô Nhóm VT + cột phụ Mã / Tên tra cứu của chính dòng con, Nhóm, Độ tin cậy, Ghi chú, Hệ số ĐV, `Con`; phần CĐT trống, không KL); **màu chỉ ở phần TDI** (nền `FFF2CC`, tên `0070C0`, mã `FF0000`); dòng cha: mã trống, giá = `SUMPRODUCT(Hệ số ĐV con, giá con)`, độ tin cậy = mức thấp nhất của con (có con Chưa khớp → Thấp), Loại dòng `Cụm`; dời mọi tham chiếu (công thức mọi sheet, merge, conditional formatting, autoFilter, defined names, vùng droplist), xóa calcChain. Ngay sau `insert` (Quy tắc 37): ghi `Tổng KL tạm tính` của mọi dòng `Con` = giá trị `Hệ số ĐV` của chính dòng (số tĩnh), rồi đổi `SUMPRODUCT` của dòng cha sang vùng cột `Tổng KL tạm tính` của các con, ví dụ `=SUMPRODUCT($BA$188:$BA$189,S188:S189)` → `=SUMPRODUCT($<Tổng KL tạm tính>$188:$<Tổng KL tạm tính>$189,S188:S189)` (script mẫu `qt37.py`). Dòng con do `insert` tạo **chưa có ô Đơn giá sau hệ số, Chủng loại, `Nguồn giá`**: Đơn giá sau hệ số không cần (giá cha lấy giá gốc con × `Tổng KL tạm tính`), `Nguồn giá` do `nguongia` ghi sau.

Ví dụ đã duyệt (sheet `Cum`): cáp ghép sợi (hệ số = số sợi); ODBO (ống đồng + bảo ôn); ONN (uPVC PN10 + bảo ôn); tủ chữa cháy (hộp + van góc + 2 cuộn vòi + lăng); đầu báo + đế; module + hộp; thiết bị vệ sinh + vòi / tay vịn; cụm bơm tăng áp 3 bơm (bơm ×3, tủ biến tần ×1, van cổng DN80 ×7, van 1 chiều DN80 ×4, khớp nối mềm DN80 ×6, đồng hồ áp, công tắc áp, ống gom đo bóc); tủ RMU 630-24kV 4 khoang (LBS ×2, VCB ×2, relay 50/51 ×2, mỗi con Chưa khớp). Vòng 2 (BOQ du toan so 2, 22/10/2026, tự duyệt): van giảm áp thủy lực DNxx + đồng hồ áp → `RV-F-DNxx` (nhóm Van Khóa Thép, TB) + đồng hồ đo áp (Chưa khớp); thuyền tắm + hương sen → `Bồn tắm nằm loại 1` + `Sen, vòi cho bồn tắm nằm loại 1` (nhóm Thiết bị Vệ Sinh, Thấp, khóa theo tên mã 0); bơm đồng bộ nguyên cụm giữ một dòng.


**Chèn dòng khi thiếu script** (Quy tắc 26, đã làm thử ở TDI25538 ngày 16/10/2026: 129 cụm, 269 dòng con, 0 lỗi):
1. **Gỡ shared formula** ở mọi sheet bị ảnh hưởng: đổi từng ô con thành công thức tường minh (dịch công thức gốc theo vị trí, ví dụ `openpyxl.formula.translate.Translator`) trước khi dời, vì công thức dùng chung bị lệch khi chèn dòng (STT `=A7+1` sẽ trỏ vào dòng con trống).
2. **Hàm dời dòng** cho từng sheet: dòng r → r + tổng số dòng con chèn sau các dòng cha có số dòng < r. Áp cho mọi tham chiếu tới sheet đó trong **mọi sheet** (có tên sheet hoặc không), bỏ qua chuỗi trong ngoặc kép và tham chiếu file ngoài `[n]`; vùng có dòng cha nằm trong khoảng thì tự giãn (SUBTOTAL phân mục bao luôn dòng con).
3. Dời cả: defined names (gồm print area, autofilter ẩn), mergeCells, autoFilter, dimension, data validation (sqref + công thức), conditional formatting, row breaks, drawing/comment anchor nếu sheet có.
4. Đánh số lại dòng và ô từ dưới lên, chèn `<row>` con ngay sau dòng cha; phần CĐT của dòng con chỉ copy style, không giá trị; phần TDI: tên, `Mã | Tên`, ĐVT, giá SUMIFS theo ô của chính dòng con, đơn giá sau hệ số dịch từ dòng cha (không có thành tiền), 7 cột phụ, Loại dòng `Con`, **`Tổng KL tạm tính` = số tĩnh bằng `Hệ số ĐV`** (Quy tắc 37); style mới nền `FFF2CC`, tên `0070C0`, mã `FF0000` (thêm fill/font/xf vào styles.xml).
5. Dòng cha: xóa mã, giá VL/NC = `SUMPRODUCT(Tổng KL tạm tính con, giá con)` (Quy tắc 37), độ tin cậy thấp nhất của con (có con Chưa khớp → Thấp), Loại dòng `Cụm`.
6. Thứ tự: ghi mã/giá → chèn dòng → droplist (lấy danh sách dòng sau chèn) → hệ số Loại VL / DS_MaHieu → đóng gói → `recalc`. Cập nhật mọi danh sách vị trí giao kèm (YCBG, danh sách Chưa khớp) theo số dòng mới.

## Bước 6. Điền đơn giá

- **Ghi công thức** SUMIFS theo Loại VL + Mã + Tên của chính dòng (mã 0 / không mã: chỉ lọc theo Tên). `apply`/`write`/`insert` làm việc này. Số cứng chỉ khi được yêu cầu.
- **Công thức giá tham chiếu ô trên chính dòng** (Quy tắc 23), để người duyệt nhìn dòng là biết giá tra theo gì, sửa ô Mã hiệu thì giá tự đổi:
  - `=SUMIFS(TV[Giá NC gốc],TV[Loại vật liệu],$<Nhóm VT>n,TV[Mã hiệu],IF($<Mã tra cứu>n="","<>§",$<Mã tra cứu>n),TV[Tên công việc],IF($<Tên tra cứu>n="","<>§",$<Tên tra cứu>n))`: mã hoặc tên trống = không lọc theo tiêu chí đó. Cột lấy theo `cols` (Nhóm VT là cột phụ, **không phải Chủng loại**).
  - **Vẫn viết cứng** (có escape) khi số hạng ghép khác nhóm hoặc khác mã / tên của dòng, hoặc giá trị có `~ * ?`. Dòng có hệ số: `(SUMIFS(…))*coef`. `insert` tự dời tham chiếu.
  - Giá đọc từ `TV[Giá VL gốc]` / `TV[Giá NC gốc]` đã nhân hệ số Loại VL (Quy tắc 24). `export`/`diff` đọc được mọi dạng công thức; Bảng khớp lưu ô `Mã | Tên` (mã thuần cũ vẫn đọc được), `apply` tự đổi sang dạng mới.
  - **Nhắc người dùng**: sửa tay ô Mã hiệu / Nhóm VT là đổi giá; mã sai ra 0, `recalc` báo ở "có mã không giá". Nên chọn từ droplist.
- **Điền cho Cao, TB, Thấp.** Chỉ chặn giá thư viện: Chưa khớp, Cần chiết tính, cụm chưa bóc tách, mã trùng chưa chọn tay. Dòng Chưa khớp để trống giá (không ghi số; Bước 6e nối công thức chờ để người duyệt chọn mã là giá tự lên); giá tham khảo web do `boq-gia-tham-khao-web` ghi sau, không phải ở skill này. Ghi chú khác (kể cả cờ phạm vi chưa xác nhận) không chặn giá.
- **Hạng mục DC đã được xác nhận "chỉ tính NC"**, và công tác "lắp đặt lại" (Quy tắc 13): VL để trống, NC theo thư viện; ghi nguồn và trích mục HSMT (A9) hoặc lý do.
- **Ô giá thư viện trống**: để trống, ghi "Thiếu giá VL" / "Thiếu giá NC".
- **Tô màu ô Đơn giá VL sau hệ số** theo loại dòng (Quy tắc 35, 36): thiết bị (dòng dùng HCTB) xanh dương `9BC2E6`; ĐVT CĐT là m, md, m2, m3, kg (không phải thiết bị) xám `D9D9D9`; dòng công việc còn lại trắng `FFFFFF`. **Tô lần 1 ngay sau khi ghi xong công thức giá của sheet và trước điểm dừng 2** (để file ở điểm dừng 2 không còn nền xanh nhạt của template ở cột này), rồi tô lại lần 2 sau `insert` / `cho` / `fillcdt` (Quy tắc 36, script mẫu `colour.py`).
- **Giá bất thường** (ví dụ Búa nước DN150 VL > 22 triệu) và **phụ kiện ống gió ước tính theo diện tích** (Quy tắc 20): đưa vào danh sách kiểm tra, không sửa.

## Bước 6b. Droplist cho cột Mã hiệu (bắt buộc, sau khi gán mã và bóc tách xong)

Người duyệt đổi mã bằng cách chọn trong danh sách **`Mã | Tên công việc`** của đúng Loại VL trên dòng đó (Quy tắc 23).
```bash
python3 $X dropdown FILE OUT 'Sheet1' 'Sheet2' …
```
- Chạy cho **mọi sheet đã xử lý** (kể cả `CT TU DIEN`, `CHI PHI KHAC`), sau `apply`/`write`/`insert` (chạy lại được; `insert` tự dời vùng droplist). Tạo sheet **ẩn `DS_MaHieu`** (cột A = `(Tất cả)`, mỗi Loại VL một cột, mỗi mục là `Mã | Tên`), droplist cột Mã hiệu, **đổi Nhóm VT thì danh sách đổi theo**; Loại VL trống / `Chưa xác định` / ghép nhiều nhóm → `(Tất cả)`. Cột Mã hiệu tự nới rộng (58). Droplist cột Nhóm VT lấy từ `DS_MaHieu`.
- Kiểu cảnh báo, **không chặn**: vẫn nhập được mục ngoài danh sách; mã sai ra giá 0. **Không tự ghi data validation bằng openpyxl.** Sau `tv_loader load --replace` chạy lại `hsloai` rồi `dropdown`. Không xóa `DS_MaHieu`. Có validation khác chồng vùng cột Mã hiệu: báo người dùng, không tự xóa.


## Bước 6c. Cột `Nguồn giá` và bộ lọc (không tra web ở skill này)

Chốt 19/10/2026: tra giá web đã tách sang skill `boq-gia-tham-khao-web` (chạy sau skill này, trước `boq-tao-ycbg`) để skill này chạy nhanh và giá web không lẫn vào kết quả gán mã. Skill này **không** tra web, không ghi giá tham khảo, không tô xanh. Dòng Chưa khớp để trống giá VL và NC, ghi chú đã tra gì.

Chạy sau `insert` và `dropdown` (số dòng đã cố định), trước `fillcdt`, cho **mọi sheet đã xử lý** (MEP và XD). Chạy lại mỗi lần ghi lại file để cột `Nguồn giá` luôn đúng.

**Cột `Nguồn giá` và bộ lọc**
- Cột phụ thứ 8, tiêu đề `Nguồn giá`, đặt ngay sau cột phụ cuối (`Tên tra cứu`; sheet XD: sau cột TDI cuối), cùng hàng và style tiêu đề cột phụ. Đã có thì nhận theo tiêu đề và ghi lại.
- Giá trị (chữ tĩnh) cho **mọi dòng công việc**: `Giá thư viện` (giá VL hoặc NC lấy từ `TV` / `TVGIA`), `Chưa có giá` (còn lại). Giá trị `Giá tham khảo` do `boq-gia-tham-khao-web` ghi: dòng đã là `Giá tham khảo` (ô giá VL số cứng có ghi chú `GIÁ THAM KHẢO WEB`) thì **giữ nguyên**, kể cả nền xanh `E2EFDA`. Dòng cha cụm: `Giá tham khảo` nếu có ít nhất một con giá tham khảo, không thì theo con.
- **Bộ lọc**: Excel chỉ cho một autoFilter mỗi sheet. Sheet đã có autoFilter → nới `ref` sang phải đến cột `Nguồn giá` (giữ `filterColumn` đang có) và sửa defined name `_xlnm._FilterDatabase` của sheet theo. Chưa có → tạo autoFilter từ hàng tiêu đề TDI, cột đầu đến cột `Nguồn giá`, đến dòng dữ liệu cuối, kèm `_xlnm._FilterDatabase` (ẩn). Hàng tiêu đề autoFilter cũ khác hàng tiêu đề cột `Nguồn giá` → không tự đổi, báo người dùng. **Lưu ý** (vòng 2): `insert` đã dời `ref` của autoFilter và `_xlnm._FilterDatabase` theo số dòng mới nên sau bóc tách chỉ cần `nguongia` nới tới cột `Nguồn giá`; script tạo autoFilter mới (ví dụ `fix1.py`) **không được cứng số dòng cuối** (đã gặp `O6:BE630` sau khi chèn 8 dòng).
- **Chạy lại**: dòng có ghi chú `GIÁ THAM KHẢO WEB` mà nay đã khớp thư viện → bỏ số cứng, ghi công thức thư viện qua `write`, trả style cũ (bỏ xanh), `Nguồn giá` = `Giá thư viện`. Ô giá tham khảo không bị coi là "giá người dùng tự nhập"; ô số cứng **không** có ghi chú này thì vẫn là giá người dùng, không động vào.

## Bước 6e. Công thức chờ cho dòng chưa có giá (Quy tắc 34)

Chốt 20/10/2026. Lý do: Bước 6 chỉ ghi SUMIFS cho dòng Cao / TB / Thấp, dòng Chưa khớp để ô giá trống, nên người duyệt chọn `Mã | Tên` từ droplist (Bước 6b) mà giá không đổi. Không thể chép nguyên công thức chuẩn sang dòng chưa có mã: tiêu chí "mã / tên trống = không lọc" làm SUMIFS **cộng giá cả nhóm**.

Chạy sau 6c (số dòng đã cố định), trước `fillcdt`, cho **mọi sheet MEP đã xử lý**. Chạy lại mỗi lần ghi lại file (sau `apply` / `write` / `insert` / `migrate2` / `boq-gia-tham-khao-web` chạy lại). Ghi bằng lxml giữ style (có lệnh script thì dùng script).

**Phạm vi ô.** Dòng công việc (`Loại dòng` = `Dòng` hoặc `Con`, có ĐVT) có ô Đơn giá gốc VL và / hoặc NC **trống** (không giá trị, không công thức). Không động vào: dòng `Cụm` (giữ `SUMPRODUCT`), dòng tiêu đề / tổng, ô **số cứng** (giá người dùng, giá tham khảo web `E2EFDA`, giá NCC), ô công thức không phải dạng chuẩn của chính dòng (viết cứng, ghép, nhân hệ số, công thức tay). Ô bị bỏ qua thì liệt kê.

**Công thức chờ** (ô VL; ô NC đổi `Giá VL gốc` → `Giá NC gốc`; cột theo `cols`):
`=IF(OR($<Mã hiệu>n="",AND($<Mã tra cứu>n="",$<Tên tra cứu>n="")),0,SUMIFS(TV[Giá VL gốc],TV[Loại vật liệu],IF(OR($<Nhóm VT>n="",$<Nhóm VT>n="Chưa xác định"),"<>§",$<Nhóm VT>n),TV[Mã hiệu],IF($<Mã tra cứu>n="","<>§",$<Mã tra cứu>n),TV[Tên công việc],IF($<Tên tra cứu>n="","<>§",$<Tên tra cứu>n)))`
- Ra **0** khi chưa có mã (không ra `""`: ô đơn giá sau hệ số / thành tiền TDI và CĐT nhân với chuỗi rỗng sẽ ra `#VALUE!`). Tổng sheet trước và sau bước này phải bằng nhau.
- Nhóm VT trống / `Chưa xác định` (droplist đang là `(Tất cả)`) thì không lọc Loại VL, để chọn mã ở nhóm nào cũng ra giá. Khóa `Mã + Tên` trùng ở hai Loại VL (loader in khóa trùng) thì giá bị cộng: Bước 7 mục 14 kiểm.
- Không nhân hệ số (dòng `Con` đã có `Tổng KL tạm tính` nhân ở `SUMPRODUCT` của dòng cha, Quy tắc 37).

**Bọc chặn mã trống cho dòng đã gán.** Ô giá đang là SUMIFS chuẩn của chính dòng (Bước 6) → bọc `=IF(OR($<Mã hiệu>n="",AND($<Mã tra cứu>n="",$<Tên tra cứu>n="")),0,<công thức cũ>)`, giữ nguyên phần SUMIFS cũ. Giá trị không đổi (dòng đã có mã). Sau bọc chạy `diff`: MẤT 0, ĐỔI 0; nếu `diff` / `export` không đọc được dạng bọc thì bỏ bọc ở dòng đã gán, chỉ giữ công thức chờ ở dòng chưa có giá, ghi `Lệch` kèm lý do.

**Cột phụ `Mã tra cứu` / `Tên tra cứu`** của mọi dòng trong phạm vi phải là **công thức tách từ ô Mã hiệu của chính dòng** (chép đúng dạng công thức của dòng đã gán cùng sheet, dời số dòng; không tự nghĩ dạng mới). Ô đang trống hoặc chữ tĩnh → ghi công thức đó, nếu không chọn mã xong giá vẫn 0.

**Cột `Nguồn giá`** của dòng trong phạm vi: `=IF(N(<Giá VL gốc>n)+N(<Giá NC gốc>n)>0,"Giá thư viện","Chưa có giá")`. Dòng đã là `Giá tham khảo` / `NCC` giữ nguyên chữ tĩnh.

**Cột giá CĐT (Bước 8)**: `fillcdt` chạy sau bước này nên ô giá CĐT của dòng trong phạm vi tham chiếu ô TDI như mọi dòng. Dòng nào `fillcdt` bỏ qua vì chưa có giá thì ghi tham chiếu đó bằng lxml, chỉ vào ô CĐT đang trống, báo số ô.

**Không đổi**: Độ tin cậy, Ghi chú gán mã, mã (vẫn trống), màu. Dòng vẫn vào file tổng hợp (6d) và YCBG như Chưa khớp. Người duyệt chọn mã xong thì tự sửa Độ tin cậy / Ghi chú; khi báo duyệt, Bước 9 lưu dòng đó vào Bảng khớp như dòng thường.

**Giới hạn đã biết**: ô đã bị `boq-gia-tham-khao-web` ghi số cứng thì không còn công thức chờ (chọn mã không đổi giá) đến khi chạy lại skill này: 6c "Chạy lại" đổi về công thức thư viện khi dòng đã có mã. `boq-ghi-gia-ncc` gặp công thức chờ ra 0 sẽ bọc `IF(N(cũ)=0, giá NCC, cũ)` nên vẫn tương thích. Tên có `* ? ~` do người duyệt chọn có thể ra 0 (SUMIFS coi là ký tự đại diện): Bước 7 báo ở "có mã không giá".

## Bước 6d. File tổng hợp đầu mục chưa có trong thư viện (Quy tắc 32)

Chạy **sau `recalc` ở Bước 7 mục 1** (số dòng đã cố định), cho **mọi sheet đã xử lý** (MEP và XD), mặc định ở mọi BOQ, không chờ yêu cầu. Chạy lại mỗi lần ghi lại file TDI (số dòng đổi). File này là đầu vào của `boq-gia-tham-khao-web` (ghi giá tham khảo vào các cột giá) rồi `boq-tao-ycbg`. Kết quả kiểm ở Bước 7 mục 12.

**Nguồn dữ liệu** (chỉ đọc, không sửa file TDI): bản `recalc` của file TDI ở cloud, `PLAN.json` (dòng con bóc tách), danh sách Cần chiết tính và "có mã không giá" của Bước 7. **Không tra web** và **không điền giá** ở bước này: các cột giá để trống cho `boq-gia-tham-khao-web` ghi.

**Đầu mục đưa vào sheet `Chưa có trong TV`**: mọi đầu mục duy nhất (khóa gom: tên + thông số chính + ĐVT; `boq-gia-tham-khao-web` dùng đúng khóa này) có Độ tin cậy `Chưa khớp`, gồm cả dòng con Chưa khớp sau bóc tách và linh kiện tủ điện Chưa khớp. Dòng tủ điện (cụm, giá = tổng linh kiện) không có giá riêng: Tình trạng giá ghi `Xem sheet "TĐ – tổng hợp"`. **Không đưa**: dòng có giá thư viện; Cần chiết tính (sheet riêng); có mã nhưng giá = 0 (sheet riêng); dòng cha cụm; thu hồi / hoàn trả chỉ NC; "CĐT yêu cầu không báo giá" / "đã tính ở …".

**Sheet `Chưa có trong TV`**: hàng 1 tiêu đề file, hàng 2 dự án + ngày tra, **hàng 3 tiêu đề cột**, dữ liệu từ hàng 4, autoFilter hàng 3 đến cột cuối, freeze `D4`, sắp theo Nhóm rồi Mô tả. Cột:

| Cột | Nội dung |
|---|---|
| A `ID` | Số thứ tự đầu mục theo lần xuất hiện đầu tiên trong BOQ; `Thành phần bóc tách` và `TĐ – linh kiện` tham chiếu theo ID này |
| B `Nhóm` | Nhóm hàng (Busway, Bơm / bể / thiết bị nước, Tủ điện & linh kiện tủ, Thiết bị điện nhẹ, Van / thiết bị đường ống, HVAC, PCCC…; XD: hạng mục lớn của BOQ) |
| C `Mô tả đầu mục (theo BOQ CĐT)` | Nguyên văn BOQ CĐT, không sửa |
| D `ĐVT`, E `Tổng KL` | ĐVT CĐT; tổng KL các dòng cùng đầu mục (số tĩnh, lấy từ KL CĐT; dòng con bóc tách cộng `Hệ số ĐV` × KL dòng cha) |
| F `Vị trí trong BOQ` | Liệt kê hết `sheet!dòng` (số dòng sau bóc tách, khớp file TDI hiện hành) |
| G `Lý do chưa khớp / đã tra` | Chép từ Ghi chú gán mã: đã tra nhóm, từ khóa, từ đồng nghĩa nào |
| H `Giá tham khảo VL (VNĐ/ĐVT)` | **Để trống ở skill này**; `boq-gia-tham-khao-web` ghi số cứng, nền `E2EFDA`; không thấy giá thì trống |
| I `Thành tiền tham khảo (VNĐ)` | Công thức `=IF(OR(H4="",E4=""),"",E4*H4)` |
| J `Khoảng giá` | Khoảng giá trên trang (H lấy số đầu) |
| K `Hãng / model tham chiếu` | Sản phẩm cụ thể trên trang đã lấy giá |
| L `Tình trạng giá` | Skill này ghi `Chưa tra giá web`. Sau `boq-gia-tham-khao-web` chỉ một trong: `Giá niêm yết` (trang hãng / NCC công bố bảng giá), `Giá bán lẻ/đại lý` (đại lý, cửa hàng, sàn), `Chưa tìm được giá – cần báo giá NCC`, `Xem sheet "TĐ – tổng hợp"` |
| M `VAT` | `đã VAT` / `chưa VAT` / `không rõ` |
| N `Nguồn`, O `Link nguồn` | Tên trang; URL kết quả đã lấy giá |
| P `Ghi chú giá` | Điều kiện của giá (chưa gồm phụ kiện, bản 220V / 380V…); dòng không thấy giá ghi `Đã tra web: <truy vấn>, không thấy giá`; để trống ở skill này, `boq-gia-tham-khao-web` ghi |
| Q `Thứ hạng KQ`, R `Ngày tra`, S `Truy vấn` | `#n`, dd/mm/yyyy, truy vấn đã dùng (không chứa tên dự án, CĐT, KL); để trống ở skill này, `boq-gia-tham-khao-web` ghi |

**Quy tắc giá và trạng thái**: skill này không điền giá, giá chỉ do `boq-gia-tham-khao-web` ghi; ĐVT của trang = ĐVT BOQ (hoặc quy đổi 1:1 rõ ràng). **Cấm** trạng thái `Ước tính từ mặt hàng tương tự`, giá trung bình, giá cỡ / model khác (`boq-gia-tham-khao-web`, Bước 2). File cũ dựng trước 16/10/2026 còn dòng "Ước tính" thì bỏ giá đó, tra lại bằng `boq-gia-tham-khao-web`, không có thì `Chưa tìm được giá`. Dòng đã khớp thư viện sau lần chạy lại thì rời khỏi file.

**Sheet `Tổng hợp`**: A1 tiêu đề, A2 nguồn (`BOQ <tên> – <mã dự án> · Thư viện: <số Loại VL>, <số dòng>, cập nhật <ngày> · Tra giá web: chưa tra`), khối "Cách đọc" (sau khi `boq-gia-tham-khao-web` chạy: giá tham khảo chỉ để định hướng, không thay báo giá NCC; nhiều trang không ghi rõ VAT; giá không nhân hệ số Loại VL; tủ điện tính theo tổng linh kiện có giá), rồi bảng theo Nhóm: `Nhóm | Số đầu mục | Có giá tham khảo | Chưa có giá | Thành tiền tham khảo (VNĐ)` bằng `COUNTIFS` / `SUMIFS` trỏ sheet `Chưa có trong TV` (vùng đến dòng cuối thực tế), `Chưa có giá = Số đầu mục − Có giá`, dòng Tổng cộng. Không có cột "Ước tính".

**Sheet phụ** (chỉ tạo khi có dữ liệu; tiêu đề ở hàng 1, autoFilter, freeze hàng 1):
- `TĐ – linh kiện` + `TĐ – tổng hợp` (file có sheet linh kiện tủ điện, ví dụ `CT TU DIEN`): linh kiện gồm `Tủ điện | Dòng | Linh kiện | ĐVT | SL | ID | Giá tham khảo VL | Thành tiền`, giá lấy bằng `INDEX/MATCH` theo ID từ sheet `Chưa có trong TV`; tổng hợp theo tủ: tổng linh kiện có giá (`SUMIFS`), số linh kiện chưa có giá (`COUNTIFS`), ghi chú "Chưa gồm vỏ / thanh cái nếu chưa có giá, vật tư phụ và nhân công".
- `Có mã – TV thiếu giá VL`: dòng có mã thư viện nhưng giá VL = 0 (Quy tắc 11): `Mô tả | ĐVT | Tổng KL | Mã / đầu mục TV đã gán | Vị trí | Ghi chú`. Không tra web (nhóm này nằm ngoài phạm vi `boq-gia-tham-khao-web`).
- `Thành phần bóc tách`: mỗi dòng con của cụm: `Sheet | Dòng cha | Đầu mục cha | Dòng con | Thành phần | ĐVT | Hệ số / 1 ĐV cha | KL cha | KL thành phần (=Hệ số × KL cha) | Tình trạng (Có mã TV (Cao / TB / Thấp) hoặc Chưa khớp – cần báo giá) | Ghi chú`.
- `Cần chiết tính`: `Sheet | Dòng | Mô tả | ĐVT | KL | Ghi chú` (hướng chiết tính). Không tra web.

**Ghi file**: dựng bằng openpyxl ở cloud (file mới); tổng, đếm, thành tiền đều là công thức; số `#,##0`; tiêu đề đậm, nền xám nhạt, wrap text, độ rộng cột đọc được. Chạy `recalc` trên bản sao: 0 ô lỗi. Commit bằng `device_commit_files` vào thư mục con `Claude outputs` cạnh file TDI (tạo thư mục nếu chưa có), tên `<tên BOQ>_Dau_muc_chua_co_trong_thu_vien_gia_tham_khao.xlsx`; không có thư mục kết nối thì gửi qua `SendUserFile`. **Chạy lại**: ghi đè lần trước với `expectedMtimeMs`; người dùng đã sửa tay (mtime khác lần ghi trước, ví dụ đã điền báo giá NCC hoặc `boq-gia-tham-khao-web` đã ghi giá tham khảo) thì **không ghi đè**, giao `…_lan-N.xlsx` cạnh file cũ và nói file nào đúng. Không dùng hậu tố `_v2`, `_da-gan`.

**Lưu ý**: file là tài liệu nội bộ (có tên dự án, vị trí BOQ), không gửi NCC; YCBG soạn riêng, không kèm tên dự án (Nguyên tắc "không gửi gì ra ngoài"). Giá tham khảo chỉ để định hướng, thay bằng báo giá NCC khi có; hệ số và giá chốt là việc của trưởng phòng.

## Bước 7. Kiểm tra trước khi báo xong

1. **`recalc`**: 0 ô lỗi; **tham chiếu vòng = 0**; không còn tên vùng rác dạng địa chỉ ô (Quy tắc 21); **có mã không giá = 0** hoặc từng dòng có lý do (giá thư viện = 0 theo Quy tắc 11 hợp lệ, liệt kê riêng); **tổng CĐT = tổng TDI** từng sheet (từ bản tính lại); `tv_loader.py check` vẫn 0/0/0.
2. **`diff` với Bảng khớp**: **MẤT = 0 và ĐỔI = 0** (trừ dòng đã đưa vào "Đề xuất đổi").
3. **`todo` đã xử lý hết**: mỗi dòng có Nhóm VT và Độ tin cậy; Chưa khớp có ghi chú đã tra gì.
4. **Rà chéo theo họ vật tư** (cả dòng từ Bảng khớp lẫn Claude gán): đọc mẫu đầu, giữa, cuối mỗi đầu mục được gán; **không dòng hệ điện / điện nhẹ nào mang mã ống nước**; cỡ mã trùng cỡ BOQ (D20 không ra D90). Lỗi đã gặp: A7.
5. **Rà ngược Chưa khớp**: tìm lại theo từ khóa, từ đồng nghĩa, kể cả mã 0, dòng không mã, sheet ẩn; lọc lại Ghi chú CĐT để không sót cụm. **Chạy A12 cho mọi sheet (Quy tắc 17)**, nhất quán theo họ vật tư (Quy tắc 15), dòng gộp đã tách (Quy tắc 18). Ghi chú "không có cỡ trong thư viện" phải đối chiếu được với `TV`.
6. **Không dòng cụm nào mang mã đơn lẻ có giá**; màu chỉ phần TDI; mọi cụm CĐT đã liệt kê thành phần đều đã bóc tách (Quy tắc 16). **Không còn dòng `[Ghép]` / `terms` nào chưa được người dùng duyệt** (Quy tắc 27): lọc cột Mã hiệu và Ghi chú để kiểm.
7. **File gốc còn nguyên** hình, add-in, định dạng; không còn sheet tạm (`cleanup`). **Droplist đã gắn** cho mọi sheet MEP sau lần ghi cuối, `DS_MaHieu` có mặt (ẩn), sheet BOQ có `data_validations` trỏ `DS_MaHieu`.
8. **Hệ số Loại VL và CĐT**: bảng `TV` có P–S, tất cả hệ số Loại VL = 1 thì giá mọi dòng **bằng giá trước `hsloai`** (đã kiểm 0 lệch); thử đổi một hệ số trên bản sao thì chỉ dòng đúng Loại VL đổi. Bước 8 đã chạy ở mọi sheet có cột giá CĐT, **tổng thành tiền CĐT = tổng KL × đơn giá TDI** từng sheet.
10. **Cột `Nguồn giá` (Bước 6c)**: mọi dòng công việc có `Nguồn giá` (`Giá thư viện` / `Chưa có giá`; dòng `Giá tham khảo` do lần chạy trước của `boq-gia-tham-khao-web` được giữ nguyên); autoFilter mỗi sheet phủ tới cột `Nguồn giá`; skill này không ghi thêm ô xanh `E2EFDA` nào, không có dòng Chưa khớp nào bị gán mã hay điền giá.
11. **Bảng kiểm tuân thủ (Phụ lục D, Quy tắc 30)**: đi lần lượt từng mục, cập nhật trạng thái và bằng chứng; **0 mục `Chưa làm`**; mục `Lệch` đã có trong "Bước đã bỏ qua / làm khác skill". Kiểm lại bằng số liệu từ `recalc` / `diff` / lọc cột, không ghi theo trí nhớ.
9. **File trên máy người dùng đúng bản**: sau commit đọc lại (kích thước / mtime), đếm ô độ tin cậy, `DS_MaHieu` có mặt, **vòng tham chiếu = 0**, không còn file `_da-gan-ma`, `_v2` hay bản trùng (trừ file `_ra-soat-lan-N` theo ngoại lệ; file tổng hợp Bước 6d trong `Claude outputs` không phải bản trùng).
13. **Hệ XD sau bóc tách (Quy tắc 33)**: tổng từng sheet XD và Summary XD bằng trước bóc tách; mọi ô cũ giữ nguyên giá trị; 0 ô lỗi; dòng tổng dùng SUMIFS loại `Con`; droplist cột N phủ đúng dòng gán; `<cols>` mọi sheet không chồng nhau (cả sheet ME); sheet `TVGIA` đúng định dạng thư viện nguồn.
12. **File tổng hợp (Bước 6d, Quy tắc 32)**: đã ghi đúng tên vào `Claude outputs` (đọc lại kích thước / mtime); `recalc` file tổng hợp 0 ô lỗi; Σ `Số đầu mục` ở sheet `Tổng hợp` = số đầu mục Chưa khớp duy nhất đã gom; các cột giá tham khảo, link, thứ hạng, ngày tra để trống và `Tình trạng giá` = `Chưa tra giá web` (không có dòng nào đã điền giá); **0 dòng `Ước tính`**; Σ `Tổng KL` đối chiếu được với KL CĐT; `Vị trí trong BOQ` khớp số dòng file TDI hiện hành; sheet phụ có đủ dòng so với nguồn (bóc tách, tủ điện, có mã giá 0, cần chiết tính).
14. **Công thức chờ (Bước 6e, Quy tắc 34)**: đếm ô trong phạm vi = ô đã nối (100%, ô bị bỏ qua có danh sách và lý do); **giá trị mọi ô và tổng từng sheet, Summary, CĐT trước = sau** 6e; 0 ô lỗi, 0 vòng; `diff` MẤT 0 / ĐỔI 0. **Thử chọn trên bản sao** (không ghi vào file giao): mỗi sheet lấy ít nhất 3 dòng chờ (một dòng `Dòng`, một dòng `Con` nếu có, một dòng Nhóm VT `Chưa xác định` nếu có), gán một `Mã | Tên` có giá trong `DS_MaHieu` → `recalc`: giá = giá `TV`, thành tiền TDI / CĐT, `Nguồn giá`, giá dòng cha cụm đổi đúng; xóa mã → giá về 0, tổng về như cũ. Dòng Nhóm `Chưa xác định`: kiểm khóa `Mã + Tên` đã chọn chỉ có một dòng trong `TV`.
15. **Tô xám ĐVT m / md / m2 / m3 / kg (Quy tắc 35)**: lọc ĐVT CĐT từng sheet, đếm số dòng thỏa (không phải thiết bị) và số ô Đơn giá VL sau hệ số đã tô `D9D9D9` (phải bằng nhau, 0 ô tô nhầm dòng ĐVT khác, dòng `Con`, ô NC / Tổng / phần CĐT); giá trị và công thức các ô tô trước = sau; báo số dòng tô theo từng ĐVT trong báo cáo.
16. **Tô màu theo loại dòng (Quy tắc 36)**: từng sheet MEP đếm dòng công việc theo 3 nhóm và đối chiếu số ô đã tô: thiết bị (công thức dùng HCTB) = ô `9BC2E6`; không phải thiết bị và ĐVT CĐT thuộc m / md / m2 / m3 / kg = ô `D9D9D9`; còn lại = ô `FFFFFF`; tổng ba nhóm = số dòng công việc (không phải dòng `Con`, tiêu đề, tổng). Không dòng nào thuộc hai nhóm; không ô NC / Tổng / phần CĐT bị đổi nền; giá trị và công thức mọi ô trước = sau; báo số dòng theo từng nhóm (và theo Chủng loại thiết bị) trong báo cáo. Sheet không có HCTB: nhóm thiết bị = 0, ghi `Không phát sinh` kèm cách kiểm. **Đối chiếu Chủng loại với công thức** (vòng 2): liệt kê dòng xanh dương mà Chủng loại không thuộc nhóm thiết bị và dòng thuộc nhóm thiết bị mà công thức không dùng HCTB; mỗi dòng lệch có tên Chủng loại và số dòng trong báo cáo (đã gặp 4 dòng MNN / LLMT). Kiểm **sau lần tô thứ hai**, và xác nhận file giao ở điểm dừng 2 đã qua lần tô thứ nhất.
17. **`Tổng KL tạm tính` dòng con và SUMPRODUCT dòng cha (Quy tắc 37)**: từng sheet MEP đếm dòng `Con` và số ô `Tổng KL tạm tính` đã ghi (phải bằng nhau, trừ ô người dùng tự nhập đã liệt kê); mọi ô là số tĩnh (không công thức) và **bằng đúng** `Hệ số ĐV` cùng dòng (0 lệch); mọi dòng `Cụm` có `SUMPRODUCT` trỏ vùng cột `Tổng KL tạm tính` của đúng các dòng con (**0 công thức cha còn trỏ cột `Hệ số ĐV`**, vùng bắt đầu ở con đầu và kết thúc ở con cuối); giá gốc VL / NC từng dòng cha trước = sau; 0 ô ghi vào `Tổng KL tạm tính` của dòng `Dòng` / `Cụm` / tiêu đề / tổng / phần CĐT; tổng thành tiền TDI, CĐT, Summary từng sheet trước = sau; 0 ô lỗi. **Khi bóc tách làm xuất hiện giá mới** (cụm trước đó chưa có giá, nay con có mã), tổng thay đổi là đúng: tính chênh lệch kỳ vọng (Σ hệ số × KL cha × giá con có giá) và đối chiếu với chênh lệch thực, ví dụ BOQ du toan so 2: +96 triệu = van RV-F ×8 + bồn tắm + sen. Thử trên bản sao: sửa `Tổng KL tạm tính` một dòng con → giá cha đổi đúng. Bảng kiểm ghi mục `QT37 Tổng KL tạm tính dòng con + SUMPRODUCT cha` (Phụ lục D chưa có dòng này thì vẫn thêm vào `KIEM_QUY_TAC.json`).
18. **Cột Đơn vị TDI từ thư viện (Quy tắc 38)**: từng sheet MEP đếm dòng `Dòng` / `Con` và số ô Đơn vị đã đặt công thức tra `TV[ĐVT]` (phải bằng nhau, ô bỏ qua có danh sách và lý do); sau `recalc` 0 ô lỗi, ô Đơn vị không rỗng; **tổng hao hụt, thành tiền từng sheet, Summary, CĐT trước = sau** (so với bản chưa đặt công thức); công thức hao hụt của dòng `Dòng` trỏ ĐVT CĐT, không còn trỏ ô Đơn vị TDI; **thử đổi một mã ở droplist trên bản sao** (không ghi vào file giao) sang mã có ĐVT khác thì ô Đơn vị đổi theo; Quy tắc 35 / 36 vẫn đếm đúng theo ĐVT CĐT; báo số dòng ĐVT thư viện khác ĐVT CĐT và danh sách theo 3 nhóm (i / ii / iii). Bảng kiểm ghi mục `QT38 Đơn vị TDI từ thư viện`.
19. **Tham chiếu ô trực tiếp trong ô Đơn giá gốc VL / NC (Quy tắc 39)**: từng sheet đếm dòng `Dòng` / `Con` có giá là công thức tra thư viện, số ô đã đổi sang tham chiếu `='<sheet TV>'!E{n}` / `F{n}` và số ô giữ lại (dòng chưa có mã giữ `SUMIFS`; khóa khớp 0 hoặc nhiều dòng; giá lệch; công thức lạ: có danh sách và lý do); **0 ô còn `HYPERLINK`** (tìm trong XML từng sheet); không có cột thêm; sau `recalc` 0 ô lỗi và **giá trị mọi ô giá, thành tiền, tổng từng sheet, Summary, CĐT trước = sau**; **mẫu thử ít nhất 5 dòng mỗi sheet (ưu tiên dòng có `*hệ số`, dòng `Con`, mã 0 / không mã)**: dòng `TV` được trỏ tới có cùng Loại VL + Mã + Tên với dòng BOQ, giá ở ô được trỏ (E hoặc F, cột đã nhân hệ số Loại VL) × hệ số đi kèm bằng đúng số hiển thị; dòng chưa có mã vẫn là `SUMIFS` và chọn mã ở droplist vẫn ra giá; dòng `Cụm`, giá tham khảo, giá NCC, số cứng không bị đổi. Bảng kiểm ghi mục `QT39 Tham chiếu ô trực tiếp trong ô giá VL / NC`.


## Bước 8. Đưa giá TDI sang cột giá BOQ CĐT (mặc định chạy, Quy tắc 25)

Chạy sau Bước 7, cho **mọi sheet có cột giá CĐT**; nguồn là đơn giá TDI sau hệ số (nhóm cột "Đơn giá", tự nhận theo tiêu đề), nên hệ số Loại VL và HS chung đều đã tính vào. Xem hàng tiêu đề từng sheet để lập `CFG.json` (cột theo tiêu đề, không theo chữ cái):

| Kiểu CĐT | Cấu hình |
|---|---|
| 3 cột đơn giá (Vật tư, Nhân công, Tổng) + 1 cột Thành tiền (TBA, MFĐ, BOQ 03–08 của TDI25538) | `{"<sheet>": {"kl": ["I"], "unit": "H", "h": "J", "i": "K", "j": "L", "amount": "M", "start": 6}}` (MFĐ lệch một cột: `kl` J, `unit` I, `h` K, `i` L, `j` M, `amount` N) |
| Nhiều khu, mỗi khu Cung cấp / Lắp đặt / Tổng | `{"kl": ["D","E","F"], "zones": ["K","N","Q"], "h": "H", "i": "I", "j": "J", "unit": "C", "start": 8}` (`zones` = cột đầu của mỗi nhóm giá trị) |
| 1 cột Đơn giá + 1 Thành tiền (MTT, CHI PHI KHAC) | `{"kl": ["D"], "unit": "C", "start": 6, "single": {"f": "E", "g": "F"}}` |

- **Kiểu `amount`**: `h = IF(N(VL)=0,"",VL)`, `i` tương tự NC, `j = tổng`, `amount = IF(OR(j="",N(KL)=0),"",KL×j)`; dòng "TỔNG CỘNG TRƯỚC THUẾ" (cột B) thêm `SUBTOTAL(9, …)` nếu ô trống, **giữ nguyên nếu đã có công thức** (TBA có `=SUM(M7:M52)`).
- Bỏ qua dòng `Con` và dòng không đơn vị. Sheet không có cột giá CĐT (`CT TU DIEN`: chỉ tên, đơn vị, số lượng): không chạy, ghi trong báo cáo.
- Sau khi chạy **báo mọi ô CĐT bị ghi đè** (TDI25538: 10 ô `=I192*100`, `=I221*5`… ở cột Vật tư của BOQ 08 PCCC, công thức cũ của CĐT; người dùng quyết giữ hay thay); số khu khác số cột KL thì chỉ điền khu khớp. `recalc`: tổng CĐT = tổng KL × đơn giá TDI từng sheet. Nhắc khi chốt hồ sơ thì copy → Paste Values. Dòng `Giá tham khảo` (nếu đã có từ lần chạy `boq-gia-tham-khao-web` trước) cũng được đưa sang CĐT; ở lần chạy đầu chưa có giá tham khảo, nên `boq-gia-tham-khao-web` sẽ chạy lại `fillcdt` cho các dòng đó (dùng lại `CFG.json` của Bước 8).


## Bước 9. Lưu vào Bảng khớp (khi người dùng báo đã duyệt)

1. Stage lại file đã duyệt (người dùng có thể đã sửa tay) và `BANG_KHOP_DA_DUYET.xlsx` mới nhất.
2. Chạy:
   ```bash
   python3 $M export FILE BANG_KHOP_DA_DUYET.xlsx --project <mã dự án> [--fix FIX.json]
   ```
   `FIX.json` dùng khi người dùng sửa bằng lời, chưa sửa trong file: `{"<sheet>!<dòng>": {"g":…, "code":…, "vt":[[…]], "nt":[[…]], "conf":…, "note":…}}`.
3. Commit bảng về `2.Thu vien gia\_Claude_tools\` (dùng `expectedMtimeMs` của lần stage). Không có thư mục đó trên máy đang kết nối thì gửi file qua chat và nhắc người dùng chép đè vào thư mục. Báo số dòng và số cụm đã lưu. **Không lưu giá tham khảo web vào Bảng khớp** (giá web thay đổi theo thời gian): dòng đó lưu là Chưa khớp, lần chạy sau tra lại ở Bước 6c. Bước này chỉ đọc file BOQ, không ghi vào file BOQ của người dùng.

## Bước 9b. Làm giàu quy tắc quy đổi (sau Bước 9, khi người dùng báo đã duyệt)

Bảng khớp lưu từng dòng; Phụ lục A lưu **quy tắc** để áp cho dòng chưa gặp. Sau khi duyệt, rút từ BOQ vừa chạy:
1. **Quy đổi mới** (cỡ BOQ → mã thư viện, DN → OD, độ dày, lớp phủ, cách tính hệ số ĐV) → thêm vào A1 / A2 / A3 / A8 / A10 đúng bảng.
2. **Từ đồng nghĩa mới** (tên BOQ → đầu mục thư viện, kèm độ tin cậy) → A4 hoặc A5.
3. **Khớp sai đã bắt** khi rà chéo → A7. **Hạng mục thư viện không có** → A6 (ghi tên dự án).
4. **Họ vật tư mới bị bỏ sót ở vòng 1** → A12. **Phạm vi cung cấp, hệ trung thế** → A9, A11.
5. **Quyết định người dùng chốt** (kèm ngày) → Quy tắc bắt buộc hoặc mục A liên quan.

Cách làm: liệt kê đề xuất (mục, nội dung thêm, nguồn dòng BOQ), **chờ người dùng duyệt**, rồi cập nhật skill bằng `propose_skills` (đọc SKILL.md hiện hành đủ, gửi bản đầy đủ). Không xóa hay gộp dòng cũ; mâu thuẫn với quy tắc cũ thì nêu rõ để người dùng quyết. Ghi tên dự án và ngày cạnh mỗi mục mới. Dung lượng skill tăng thì báo người dùng trước khi cắt bất kỳ phần nào.

## Báo cáo cho người dùng (ngắn)

File có cả hai hệ: thêm phần XD ("Báo cáo XD").

- **Bảng kiểm tuân thủ** (Quy tắc 30, luôn có, đặt đầu báo cáo): bảng đủ mọi mục Phụ lục D, cột Mục / Trạng thái / Bằng chứng (số dòng, sheet, ví dụ dòng hoặc cách kiểm ra 0). Dòng tổng: số mục Đã áp / Không phát sinh / Lệch / Chưa làm (phải = 0).
- **Bước đã bỏ qua / làm khác skill** (Quy tắc 28, luôn có, ngay sau Bảng kiểm): từng mục kèm lý do và việc cần làm tiếp; không có thì ghi "Không có". Gồm cả "bỏ điểm dừng 2 theo yêu cầu người dùng" khi đã tự duyệt bóc tách, và việc dùng script mẫu `qt37.py` / `colour.py` / `donvi.py` thay lệnh chưa có.
- **Phạm vi**: sheet đã xử lý, số dòng công việc, số dòng có giá; nguồn giá (tên file thư viện + ngày cập nhật).
- **File đã ghi**: đúng một file làm việc `<tên>_TDI.xlsx` trong thư mục dự án (tên, kích thước, giờ ghi), cộng file tổng hợp Bước 6d trong `Claude outputs`. Nếu giao file mới (ngoại lệ): ghi tên file mới, nói file cũ để nguyên và nhắc đóng file cũ không lưu.
- **Nguồn kết quả** (Bảng khớp / Claude gán / chưa xử lý) và **độ tin cậy** (Cao / TB / Thấp / Chưa khớp / Cần chiết tính; số cụm và dòng con; mã trùng cần chọn tay); kết quả `diff` (MẤT / ĐỔI / THÊM); droplist (sheet, số nhóm / mã, nhóm không có trong TV).
- **Đề xuất đổi so với Bảng khớp** (nếu có) kèm lý do. **Phạm vi cung cấp**: số dòng A / DC / NSC / PS kèm trích mục HSMT; dòng đã áp "chỉ NC" và dòng chờ xác nhận.
- **Cần người dùng xác nhận**: dòng cần chốt vật liệu (kèm trích spec); cờ phạm vi; thư viện không đạt spec; cụm chưa có cấu hình; BOQ lệch bản vẽ; giá bất thường; ô CĐT bị ghi đè; **cáp trung thế gán tạm giá hạ thế (bắt buộc báo giá)**; **phụ kiện ống gió ước tính theo diện tích (Thấp, kèm công thức)**; dòng gộp đã tách (kèm **giả định khi tự duyệt bóc tách**, ví dụ chọn bồn tắm loại 1); **dòng tô xanh dương theo công thức HCTB nhưng Chủng loại không thuộc nhóm thiết bị** (Quy tắc 36); **dòng có ĐVT thư viện khác ĐVT CĐT, chia 3 nhóm (i quy đổi đã biết / ii cần xem lại mã / iii cách gọi)** (Quy tắc 38); **số ô giá VL / NC đã đổi sang tham chiếu ô trực tiếp, dòng giữ `SUMIFS` và lý do; nhắc: đổi mã ở droplist trên ô đã tham chiếu thì giá không tự đổi, cần chạy lại `xemnguon.py`** (Quy tắc 39).
- **Bước tiếp theo**: chạy `boq-gia-tham-khao-web` để tra giá tham khảo web cho các đầu mục Chưa khớp (báo số đầu mục duy nhất và số dòng sẽ tra); skill này không tra web. Sau đó chạy `boq-tao-ycbg`.
- **File tổng hợp đầu mục chưa có trong thư viện** (Bước 6d, Quy tắc 32): tên file và thư mục; số đầu mục duy nhất (chia theo Nhóm; các cột giá tham khảo để trống, chờ `boq-gia-tham-khao-web`), số tủ điện và linh kiện, số dòng ở từng sheet phụ; nhắc dòng thuộc hãng khóa dự án (Grundfos, Daikin, York, Schneider…) nên xin báo giá NCC sớm.
- **Chuẩn bị cho Bước 5 (người rà soát)**: dòng cần kiểm tra trước (Thấp → KL lớn → giá trị lớn); danh sách Chưa khớp nhóm theo NCC / hãng (hãng khóa dự án như Grundfos, Daikin, York, Schneider xin báo giá sớm); YCBG nháp và bảng theo dõi trống, không gửi.
- **CĐT và hệ số**: sheet đã fill giá CĐT (số dòng, tổng CĐT = tổng TDI), ô CĐT bị ghi đè, sheet không có cột giá CĐT; bảng `TV` đã có cột hệ số Loại VL (hệ số hiện tại).
- **Nhắc**: Excel đang mở file này thì đóng không lưu rồi mở lại (lỡ bị đè: Lịch sử phiên bản OneDrive hoặc báo Claude); cột Mã hiệu có droplist `Mã | Tên` theo Nhóm VT (dòng Chưa khớp đã có công thức chờ: chọn mã là giá tự lên, nhớ sửa Độ tin cậy / Ghi chú và kiểm Nhóm VT đúng), **cột Đơn vị TDI tự đổi theo mã đã chọn (lấy từ thư viện); ĐVT thư viện có thể khác ĐVT CĐT nên xem danh sách lệch ĐVT**, định mức dòng con sửa ở cột `Tổng KL tạm tính` (giá cha `SUMPRODUCT` theo cột này, sửa `Hệ số ĐV` không đổi giá), sửa giá thư viện ở cột thô P/Q của `TV`, đổi bảng `TV` thì chạy lại `hsloai` rồi `dropdown`; giá thư viện là giá cơ sở, trưởng phòng chốt hệ số; sau khi duyệt báo để lưu Bảng khớp (Bước 9) và làm giàu quy tắc (Bước 9b).