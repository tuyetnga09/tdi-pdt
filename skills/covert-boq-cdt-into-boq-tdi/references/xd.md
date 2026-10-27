# Hệ Xây dựng (XD) — chi tiết

_Tách từ SKILL.md (không đổi nội dung)._

BOQ phần xây dựng (cọc, đất, bê tông, ván khuôn, cốt thép, kết cấu thép, xây trát, hoàn thiện, hạ tầng…) **không đi qua `boq_tdi.py`**: không có Mã hiệu, Chủng loại MEP hay hệ số HSVL/HSTB. Dùng `boq_xd_tools.py` (cùng script gán giá XD của skill `boq-gan-ma-gia-thu-vien`). Layout lấy theo form XD nội bộ TDI (file `01. File nội bộ TDI.xlsx`, dự án BKY).

**Nhận diện hệ của từng sheet** trước khi chạy:
- Sheet XD: hạng mục là công tác xây dựng như trên, thường có tên nhà/bể/hạ tầng (NHÀ CN8-1, BỂ…, HẠ TẦNG). Phần ống thoát nước mái, cống, rãnh nằm trong BOQ xây dựng vẫn là XD.
- Sheet MEP: điện, nước, HVAC, PCCC, báo cháy, ELV.
- **File có cả hai**: chạy `boq_tdi.py run` cho sheet MEP trước (sửa `config.json`, bỏ sheet XD khỏi `sheets` rồi chạy lại `run --config`), nạp thư viện MEP (mục 3), sau đó chạy `build` XD trên chính file `_TDI.xlsx` vừa ra, chỉ với sheet XD. Không sheet nào được chuyển hai lần.
- Không chắc sheet thuộc hệ nào → hỏi một lần.

**Chế độ**: giữ nguyên sheet CĐT bên trái (giống `inplace` MEP), dựng **khối XD** bên phải, cách cột cuối của CĐT một cột trống. Không ghi đè ô CĐT nào.

**Lấy script**: `boq_xd_tools.py` ở `2.Thu vien gia\_Claude_tools\` (lấy như mục "Công cụ" của skill gán mã). Không có → xin người dùng đính kèm; không tự viết lại. `XD=/tmp/boqtools/boq_xd_tools.py`.

### Các bước XD

1. **Nhận diện**:
   ```bash
   python3 $XD auto "<file>.xlsx" /home/claude/boq/xdcfg.json "<mã dự án>"
   ```
   Mỗi sheet: hàng tiêu đề (1 hoặc 2 tầng, nhận cả tiêu đề song ngữ Việt/Trung/Anh), dòng dữ liệu đầu–cuối, cột `stt`, `desc`, `unit`, `kl`, `note`, cột bắt đầu khối XD (`block_start`).
   - Sửa `xdcfg.json` khi nhận sai. **Bỏ khỏi `sheets`** mọi sheet không phải BOQ XD (sheet MEP, tổng hợp, chỉ dẫn).
   - **Nhiều cột khối lượng** (theo nhà/khu): `kl` phải là cột tổng KL. CĐT không có cột tổng → hỏi người dùng.
   - **Tiêu đề sát đầu sheet**: khối XD cần 2 hàng phía trên hàng tiêu đề cuối (hàng tiêu đề trên + hàng hệ số). Không đủ thì báo người dùng, không chèn dòng vào sheet CĐT.
   - Đặt `block_start` xa hơn nếu CĐT có cột ẩn hoặc ghi chú ở xa bên phải.
2. **Dựng khối XD và nạp `TVGIA`**:
   ```bash
   python3 $XD build "<file>.xlsx" "/mnt/user-data/outputs/<tên gốc>_TDI.xlsx" /home/claude/boq/xdcfg.json "<file thư viện XD đã stage>"
   ```
   - **File thư viện XD dò theo mục "Vị trí thư viện giá"**: file `2.Thu vien gia XD.xlsx` (ứng viên `thu vien gia … XD`) trong thư mục dự án trước; không có mới dùng `01. File nội bộ TDI.xlsx` bản mới nhất của phòng (hỏi đường dẫn nếu chưa biết). Stage từ máy như thư viện MEP và **kiểm bố cục trước khi `build`** (sheet `TVGIA`, cột B / G / J); không đạt thì dừng, không chạy `build`. Script chép sheet `TVGIA` cột A–K, **giữ công thức nội bộ sheet và giá trị đã tính**, bỏ định dạng, ghi nguồn + ngày ở L1. File đã có `TVGIA` thì giữ nguyên, báo người dùng.
   - Thêm sheet **`Summary XD`**: hệ số chung VL (C6) / NC (C7), mặc định 1, **trưởng phòng chốt** (BKY dùng 1.12); bảng thành tiền từng sheet và tổng.
   - Ghi bằng XML, giữ nguyên phần CĐT; xóa calcChain, bật tính lại khi mở.
3. **Kiểm tra**:
   - `python3 $XD todo OUT todo.json <các sheet XD>`: số dòng công việc mỗi sheet bằng số dòng có ĐVT của CĐT; trạng thái "chưa gán" (trừ dòng có ghi chú loại trừ).
   - `python3 $XD recalc OUT rep.json <các sheet XD>`: 0 ô lỗi; tổng = 0 (chưa có giá).
   - **Kiểm vòng tham chiếu** như mục 4.5 (dòng tiêu đề nhóm không có dòng con không được cộng vào dòng tổng).
   - So ô CĐT với file gốc (giá trị, đậm, nền, viền): không khác ô nào.
   - Còn sheet `_TDI*` thì xóa như mục 4 phía trên.
4. **Giao file và điểm dừng**: như mục 5 phía trên. Tóm tắt thêm cho XD: sheet XD, vị trí khối XD, số hạng mục / tiêu đề, `TVGIA` đã nạp (file nguồn: thư mục dự án hay file nội bộ, ngày sửa, kết quả kiểm bố cục), hệ số chung đang để 1 chờ chốt. Bước sau: skill `boq-gan-ma-gia-thu-vien` (Phần XD).

### Layout khối XD (bắt đầu tại `block_start`, gọi là c0)

Hàng tiêu đề phụ = hàng tiêu đề cuối của CĐT (`data_start_row − 1`); hàng tiêu đề trên = hàng ngay trên; hàng hệ số = trên nữa.

| Cột | Tiêu đề trên / phụ | Dòng hạng mục r |
|---|---|---|
| c0 | – / STT | `=IF(ISBLANK(STT CĐT),"",STT CĐT)` |
| c1 | NỘI DUNG (TDI) | tham chiếu Nội dung CĐT (ô này là dấu hiệu nhận khối XD, không đổi tên) |
| c2 | CHỦNG LOẠI VẬT TƯ | nhập tay (chữ xanh): mác, cấp thép, quy cách theo spec |
| c3 | ĐƠN VỊ | tham chiếu ĐVT CĐT |
| c4–c8 | SL, Dài, Rộng, Cao, Hệ số | diễn giải KL, nhập tay |
| c9 | Khối lượng TDI | nhập tay (KL TDI tự đo bóc) |
| c10 | KHỐI LƯỢNG | tham chiếu KL CĐT |
| c11 / c12 / c13 | ĐƠN GIÁ: VẬT LIỆU / NHÂN CÔNG / TỔNG | `IF(N(KL)=0,0,ROUND(TT phân tích VL × HS chung VL / KL,0))`, tương tự NC; tổng |
| c14 | TỔNG (thành tiền) | `ROUND(N(KL)×c13,0)` |
| c15 | GHI CHÚ (CĐT) | tham chiếu ghi chú CĐT (dòng có "không báo giá", "đã tính ở…" được skill gán mã bỏ qua) |
| c16 | NHÃN | trống; hàng hệ số ghi "HS chung" |
| c17–c21 | VẬT LIỆU: Công tác / Đơn vị / Khối lượng / Đơn giá / Thành tiền | Công tác, Đơn giá **để trống** cho skill gán mã; Đơn vị = c3, KL = c10, Thành tiền = `N(KL)×N(Đơn giá)` |
| c22–c26 | NHÂN CÔNG: (như VL) | như VL |

- **Hệ số chung**: ô hàng hệ số tại c17 (VL) và c22 (NC) = `'Summary XD'!$C$6` / `$C$7`.
- **Dòng tiêu đề** (STT La Mã, chữ cái, `*`, I.1…): nền vàng, c14 = `SUBTOTAL(9, …)` tới hết mục. Tiêu đề không có dòng con thì để trống ô tổng, không trỏ vào dòng tổng cộng.
- **Dòng ghi chú** (có mô tả, không ĐVT, không KL): chỉ STT, Nội dung.
- **Dòng "TỔNG CỘNG (TDI)"** ở `data_end_row + 2`: `SUBTOTAL(9)` cột c14.
- Ô tham chiếu được ghi kèm giá trị tạm (cache) để `todo` đọc được ngay mà không cần mở Excel.
- Khác form BKY: BKY là sheet TDI riêng (A–AB, hệ số S1/X1, `SUMMARY 預算表`); khối XD là cùng bộ cột đó dịch sang phải. Skill gán mã nhận cả hai theo tiêu đề.

Đã chạy thử: BOQ CĐT dựng lại từ 2 sheet BKY (CN8-1, CN8-6) → `build` → gán bằng Bảng khớp XD: 287/383 dòng gán tự động, giá VL/NC từng dòng trùng 100% với file BKY; phần CĐT không đổi ô nào.
