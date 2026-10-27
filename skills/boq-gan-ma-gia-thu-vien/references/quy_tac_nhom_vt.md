# Quy tắc xếp Nhóm VT

18 nhóm tương ứng 18 file thư viện. Tên nhóm phải ghi **đúng như cột `Nhóm VT` của sheet thư viện** (không dấu, như tên file).

Các dòng có dấu (?) là suy luận từ tên file, **chưa được phòng xác nhận**. Nếu kết quả xếp nhóm gây tranh cãi, nêu ra để người dùng chốt, rồi cập nhật lại file này.

| Nhóm VT | Bao gồm | Phân biệt |
|---|---|---|
| Cap dien | Cáp điện lực, cáp điều khiển, dây điện (CV, CVV, CXV, Cu/XLPE…) | |
| TCap | Thang cáp và phụ kiện thang (?) | |
| MCap | Máng cáp, trunking và phụ kiện máng (?) | |
| OLDay | Ống luồn dây điện: PVC luồn dây, ống thép luồn dây (EMT/IMC), ống ruột gà, và phụ kiện (?) | Ống nằm dưới tiêu đề hệ **Điện / Điện nhẹ / Báo cháy** thì vào OLDay, kể cả khi là ống PVC |
| PVC | Ống và phụ kiện PVC / uPVC cấp thoát nước | |
| PPR | Ống và phụ kiện PPR | |
| HDPE | Ống và phụ kiện HDPE, kể cả ống gân / ngăn sóng 2 lớp | |
| ODong | Ống đồng gas điều hòa và phụ kiện | |
| Bao On | Bảo ôn ống đồng, ống gió, ống nước lạnh | Dòng ghi gộp "ống đồng kèm bảo ôn": nhóm chính là ODong, Ghi chú nêu Bao On |
| OGio | Ống gió và phụ kiện ống gió | Miệng gió, van gió: **chưa chốt** (xem dưới) |
| OThep | Ống thép mạ kẽm, ống thép đen và phụ kiện | Ống thép hệ PCCC: **chưa chốt** |
| Van Khoa | Van, khóa các loại; lọc Y; đồng hồ nước (?) | Van hệ PCCC: **chưa chốt** |
| PCCC | Đầu phun sprinkler, tủ/hộp chữa cháy, lăng vòi, bình chữa cháy, vật tư chuyên dụng chữa cháy | |
| CTKhac | Công tác khác: giá đỡ, sơn, thử áp, vệ sinh đường ống… (?) | Chỉ dùng khi chắc chắn có đầu mục tương ứng trong thư viện CTKhac |
| TB Dien | Tủ điện, thiết bị đóng cắt (MCB, MCCB, ACB), đèn, ổ cắm, công tắc | |
| TB Ve Sinh | Bồn cầu, lavabo, tiểu nam, vòi, sen, phụ kiện vệ sinh | |
| TB HVAC | Điều hòa, dàn nóng / lạnh, FCU, AHU, chiller, quạt | |
| TB Bao Chay | Tủ trung tâm báo cháy, đầu báo, nút ấn, chuông / đèn báo, module | |

## Các điểm chưa chốt (quy tắc tạm)

Đến khi phòng chốt, dùng quy tắc tạm dưới đây. Các dòng thuộc diện này để Độ tin cậy **TB**, Ghi chú "theo quy tắc tạm", để người làm giá dễ lọc ra.

| Trường hợp | Quy tắc tạm | Phương án khác |
|---|---|---|
| Miệng gió, van gió, van chặn lửa | OGio | TB HVAC |
| Ống thép hệ PCCC | OThep (xếp theo vật liệu) | PCCC |
| Van hệ PCCC (van cổng, van báo động, van 1 chiều) | Van Khoa | PCCC |
| Đồng hồ nước | Van Khoa | CTKhac |

Nếu sheet thư viện cho thấy nhóm kia có đúng mã đó (ví dụ trong PCCC có "Van cổng DN100 PCCC"), thì ưu tiên nhóm có mã khớp.

## Không thuộc nhóm nào → `Chưa xác định`

Đào đất, lấp đất, vận chuyển đất, hố ga BTCT / xây, nắp hố ga, bể tự hoại, bể tách mỡ, gối đỡ bê tông, chi phí đấu nối do CĐT, chi phí xin phép. Đây là danh sách cần chào giá theo cách khác (định mức xây dựng, thầu phụ). Đừng ép vào CTKhac.
