# Layout inplace và hệ số theo Chủng loại

_Tách từ SKILL.md. Đọc khi cần đối chiếu cột/ô của file inplace._

## Layout inplace (đối chiếu file thật, ví dụ TDI25538 sheet B2/B3)

**Phần CĐT** (ví dụ A–T): STT, Mô tả, Đơn vị, KL theo khu, Tổng KL, Đơn giá Cung cấp/Lắp đặt/Tổng, Giá trị theo khu, Ghi chú/Remark. Giữ nguyên, skill không ghi vào.

**Hệ số**: nằm ngay trên tiêu đề bảng. TDI25538: nhãn ở hàng 3, giá trị ở hàng 4.

| Hệ số | Ô (TDI25538) | Công thức |
|---|---|---|
| Hao hụt | X4 | `='Summary TDI'!H8` |
| HSVL | Y4 | `='Summary TDI'!C8` |
| HSNC | Z4 | `='Summary TDI'!E8` |
| HSTB | AA4 | `='Summary TDI'!D8` |
| HSKL | AC4 | `='Summary TDI'!B8` |
| HS chung | BK4 | `='Summary TDI'!F8` |
| Làm tròn | BL4 | `='Summary TDI'!G8` |

**Phần TDI**: tiêu đề ở hàng 5–7. Cột tính từ ô "Nội Dung Công Việc" (gọi là N):

| Vị trí | Cột | Ví dụ B2 |
|---|---|---|
| N | Nội dung (`=B`) | U |
| N+1 | Quy cách – Mã hiệu | V |
| N+2 | Nhãn hiệu | W |
| N+3 | Đơn vị | X |
| N+4 | Giá gốc VL | Y |
| N+5 | Giá gốc NC | Z |
| N+6 | Chủng loại | AA |
| N+7 | In ấn | AB |
| N+8 | Tổng KL | AC |
| — | Diễn giải KL (18 cột) | AD–AU |
| — | KL × HSKL | AV |
| — | **Đơn giá sau hệ số VL / NC / Tổng** | AW / AX / AY |
| — | Thành tiền theo khu và Total | AZ–BK |
| — | Ghi Chú | BL |

**Cột phụ của skill gán mã**: bên phải cột Ghi Chú, theo thứ tự `Nhóm VT | Độ tin cậy | Ghi chú gán mã | Hệ số ĐV | Loại dòng` (B2: BM–BQ). Skill gán mã tạo các cột này nếu chưa có và luôn tìm chúng theo tiêu đề. Cột `Nhóm VT` nhận giá trị từ cột `Loại vật liệu` của bảng `TV`.

**Sheet có số khu khác thì mọi cột dịch theo.** B4 (CĐT có thêm cột) bắt đầu TDI ở X; đơn giá sau hệ số ở AZ/BA; cột phụ ở BP–BT. Luôn xác định cột theo tiêu đề.

**Dòng cuối** "Tổng cộng trước thuế (TDI)": SUBTOTAL các cột thành tiền. Dòng tiêu đề cuối cùng (ví dụ "B – HẠNG MỤC NHÀ THẦU BỔ SUNG") nếu không có dòng con thì ô thành tiền để trống, không trỏ vào dòng tổng (xem mục 4.5).

**Hệ số theo Chủng loại**:
- Đơn giá VL dùng `$Y$4` (HSVL) hoặc `$AA$4` (HSTB). Script chọn theo nhóm TB/VL của mã Chủng loại (`chung_loai.md`).
- Bảng "Loại VL – Hệ số VL/NC" trong Summary TDI (dòng 11–24) **hiện không được công thức nào tham chiếu**. Nếu người dùng hỏi:
  - Nói rõ điều này.
  - Đề xuất bảng quy đổi Chủng loại → Loại VL (tên Loại VL trùng với cột `Loại vật liệu` của bảng `TV`).
  - Chỉ nối công thức khi người dùng (trưởng phòng) duyệt.
