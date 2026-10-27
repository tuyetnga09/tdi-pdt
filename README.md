# tdi-boq

Plugin gom các skill lập giá BOQ dự thầu của Phòng Đấu thầu TDI.

## Các skill (theo thứ tự chạy)

| Skill | Việc làm |
|---|---|
| `covert-boq-cdt-into-boq-tdi` | Chuyển BOQ CĐT sang template TDI (Bước 1–2) |
| `boq-gan-ma-gia-thu-vien` | Gán mã hiệu và đơn giá thư viện (Bước 3–4) |
| `boq-gia-tham-khao-web` | Tra giá tham khảo cho đầu mục Chưa khớp |
| `boq-tao-ycbg` | Tạo file Đề nghị báo giá nháp |
| `boq-lay-gia-ncc` | Quét báo giá NCC, xuất Bảng rà soát và ASSIGN.json |
| `boq-ghi-gia-ncc` | Ghi giá NCC đã duyệt vào BOQ (Bước 6) |
| `tdi-duong-dan-khong-co-dinh` | Quy tắc dò đường dẫn, không ghi cứng thư mục/tên user |

## Quy tắc sửa đổi (làm việc trên Git)

1. Không commit thẳng vào `main`. Tạo nhánh: `git checkout -b sua/<ten-skill>-<noi-dung>`.
2. Mỗi lần sửa SKILL.md, ghi rõ ngày và lý do ở mục "lỗi đã gặp" của skill đó nếu có.
3. Mở Pull Request, người quản trị duyệt rồi mới merge.
4. Sau khi merge, tăng `version` trong `.claude-plugin/plugin.json` (0.1.0 -> 0.1.1 cho sửa nhỏ, 0.2.0 cho thêm tính năng).
5. Không đưa vào repo: đường dẫn máy cá nhân, thông tin đăng nhập, file BOQ/báo giá thật của dự án.

## Đẩy lên Git lần đầu

```bash
cd tdi-boq
git init -b main
git add .
git commit -m "Khởi tạo plugin tdi-boq v0.1.0"
git remote add origin <URL-repo-private>
git push -u origin main
```

Nên dùng repo **private** vì skill có thông tin nội bộ TDI (mẫu YCBG, thông tin liên hệ).

## Công cụ đi kèm (`claude/boq/`)

Thư mục `claude/boq/` chứa 14 file công cụ lấy từ Project "Tối ưu hóa quy trình Phòng đấu thầu": `tv_loader.py`, `tv_link.py`, `boq_xml_tools.py`, `boq_ncc_tools.py`, `boq_memory.py`, `ncc_rasoat.py`, `tdi_filter.py`, `xemnguon.py`, `ycbg_tools.py`, `ycbg_gen.py`, `YCBG_CONFIG_mau.json` và 3 file hướng dẫn `.md` (link workbook thư viện, Power Query, tham chiếu gán mã TDI).

Lưu ý:

- `tdi_filter.py` mặc định đọc `/home/claude/boq/config.json`; cần chỉnh đường dẫn hoặc truyền tham số khi chạy ở nơi khác.
- `xemnguon.py` là bản 27/10/2026 (Quy tắc 39: tham chiếu ô trực tiếp, có `--dry`).
- Các file này được chép lại từ Project, nên chạy thử một lần trên dự án nhỏ trước khi dùng thật.

## Còn thiếu (chưa có trong Project)

Skill có nhắc tới 7 file sau nhưng chúng không nằm trong Project nên chưa vào plugin. Nếu bạn có, hãy bổ sung vào `claude/boq/`: `config.json`, `overrides.json`, `qt37.py`, `colour.py`, `donvi.py`, `refprice_cache.json`, `xdcfg.json`.

## Lịch sử

- 0.2.3 (27/10/2026): `boq-gan-ma-gia-thu-vien` và `boq_xml_tools.py`: dòng con bóc tách (đầu mục con) tự ghi `XOA` vào cột `In ấn`.
- 0.2.2 (27/10/2026): thay `xemnguon.py` bằng bản 27/10/2026.
- 0.2.1 (27/10/2026): cập nhật skill `boq-gan-ma-gia-thu-vien` (Quy tắc 39 mới: ô giá là tham chiếu ô trực tiếp, không dùng `HYPERLINK`).
- 0.2.0: thêm 14 file công cụ `claude/boq/`.
- 0.1.0: 7 skill ban đầu.
