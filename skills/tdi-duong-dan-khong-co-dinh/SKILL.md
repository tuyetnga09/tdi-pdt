---
name: "tdi-duong-dan-khong-co-dinh"
description: "Dùng cùng các skill BOQ TDI (covert-boq-cdt-into-boq-tdi, boq-gan-ma-gia-thu-vien, boq-*): quy tắc dò đường dẫn thư viện giá và link Excel, không ghi cứng thư mục hay tên user."
---

# Đường dẫn không cố định cho các skill BOQ TDI (chốt 26/10/2026)

Skill này **ưu tiên hơn** mọi đường dẫn cứng còn sót trong `SKILL.md` và `references/` của các skill BOQ TDI. Các skill đó chạy trên nhiều máy, nhiều tài khoản, nhiều dự án, nên không có thư mục nào là mặc định.

## 1. Đường dẫn cũ chỉ là ví dụ, không phải mặc định

Các chuỗi sau trong skill khác (`C:\Users\Admin\OneDrive - TDI Group\2.Thu vien gia`, `...\2.Thu vien gia\_Claude_tools\`, `C:\Users\Admin\OneDrive - TDI Group\Desktop\PĐT Test\...`, và bất kỳ `C:\Users\<tên>\...` nào) là **ví dụ của một máy**. Không dùng chúng làm bước dự phòng "thư mục chung", không ghi vào file giao cho người dùng, không đưa vào câu lệnh. Chỗ nào skill khác viết "thư mục chung `2.Thu vien gia`" hoặc "`_Claude_tools` của thư mục chung", làm theo mục 2 dưới đây.

## 2. Cách dò thư mục (theo thứ tự, dừng ở bước đầu tiên tìm được)

1. Đường dẫn người dùng chỉ trong chat.
2. Thư mục dự án (chứa file BOQ CĐT): `device_list_dir` với `recursive=true`, tìm thư mục có tên chuẩn hóa chứa `thu vien gia` (bỏ dấu, chữ thường, bỏ số thứ tự đầu, gộp dấu chấm / gạch / khoảng trắng), và thư mục `_Claude_tools` nếu cần script. Có từ hai ứng viên thì hỏi một lần.
3. Thư mục gốc đồng bộ của người dùng: `get_device_info.homeDirectories` (ví dụ `TDI Group`, `OneDrive - TDI Group`, `OneDrive`), rồi `device_list_dir` từng cấp tìm `2.Thu vien gia` / `_Claude_tools`. Tên user và ổ đĩa luôn lấy từ `get_device_info` / `resolvedPath`, không giả định.
4. Không tìm ra: hỏi người dùng đường dẫn, hoặc xin họ đính kèm file. Không dùng đường dẫn của máy khác.

Đường dẫn có chữ có dấu hoặc khoảng trắng lạ có thể bị từ chối khi xin quyền từ xa: khi đó nhờ người dùng bấm "Add folder" trong app Claude desktop.

## 3. Link Excel luôn là đường dẫn tương đối

- Mọi liên kết ngoài (`xl/externalLinks/_rels/*.rels`) trong file giao phải là đường dẫn **tương đối so với thư mục chứa file BOQ**, ví dụ `2.%20Thu%20vien%20gia/<file>.xlsx`. Không có `C:\`, `Users`, `file:///`.
- Nối link bằng `python3 tv_link.py link FILE OUT --libdir "<thư mục thư viện đã stage>" --linkrel "<thư mục thư viện tính từ thư mục chứa file BOQ>"` (bản `tv_link.py` 26/10/2026 ghi link tương đối, không cần bước sửa riêng). `--linkrel` tính theo dự án: `2. Thu vien gia` nếu nằm cùng cấp, `../2. Thu vien gia` nếu ở thư mục cha / anh em. Không biết đường dẫn tương đối thì đưa `--boqdir` và `--linkbase` (đường dẫn thật lấy từ `resolvedPath`) để script tự tính. **Không dùng `--absolute`.** Chạy `python3 tv_link.py link -h` thấy không có `--linkrel` là bản cũ: dừng, lấy bản mới ở `_Claude_tools`. Thư viện nằm ngoài cây thư mục dự án (khác ổ): script dừng; đề nghị chép vào thư mục dự án, không chép thì không nối link sống, giữ snapshot `TV`.
- File đã nối bằng đường dẫn tuyệt đối: `python3 tv_link.py relink FILE OUT --linkrel "<đường dẫn tương đối>"` trên bản sao, giao bản `_v2`.
- Kiểm trước khi giao: `unzip -p FILE 'xl/externalLinks/_rels/*' | grep -c -E 'C:|Users|file:'` phải ra 0 với link do Claude tạo.
- Báo người dùng: khi gửi file sang máy khác phải giữ nguyên cấu trúc thư mục dự án; Excel vẫn báo không làm mới được thì Data → Edit Links → Change Source một lần.

## 4. Script và Bảng khớp (`_Claude_tools`)

Không giả định vị trí. Dò theo mục 2 (ưu tiên trong thư mục dự án, sau đó thư mục gốc đồng bộ của người dùng). Không có thì dùng bản ở Project `claude/boq/*.py` và xin người dùng đính kèm `BANG_KHOP_DA_DUYET.xlsx`.

## 5. Power Query (dự phòng)

Đường dẫn thư viện đặt trong một ô cấu hình (bảng `CauHinh`, ô B1) để máy khác chỉ sửa ô đó. Không ghi cứng `Folder.Files("C:\Users\...")` trong query. Không dùng `CELL("filename")` để tự dò, vì file ở thư mục OneDrive / SharePoint đồng bộ có thể trả về địa chỉ web.