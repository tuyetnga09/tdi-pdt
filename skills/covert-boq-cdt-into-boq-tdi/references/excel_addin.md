## Chạy trong add-in Claude for Excel (tách từ SKILL.md)

**Đang chạy trong add-in Claude for Excel** (không có Python):
- Đọc `references/excel_addin.md` và chỉ dùng 2 script A/B trong đó.
- Khi xong, **xóa sheet ẩn `_TDI_JOB` và `_TDI_TMP_*`** bằng một script cuối, rồi xác nhận lại danh sách sheet:
  ```js
  for (const n of ["_TDI_JOB"]) context.workbook.worksheets.getItemOrNullObject(n).delete();
  // và mọi sheet tên bắt đầu "_TDI_TMP": load worksheets/items/name, lọc, delete()
  ```
- Thư viện giá: hướng dẫn người dùng dán Power Query theo skill `boq-gan-ma-gia-thu-vien` (`references/power_query_thu_vien.md`). **Lưu ý: bản Power Query bỏ dòng không có mã hiệu nên THIẾU dòng** (bộ TDI25538 thiếu 73 dòng CTKhac) và **chưa có cột `Loại vật liệu` đúng tên mới**. Nạp bằng add-in thì sau đó phải chạy lại `tv_loader.py load --replace` + `check` ở Cowork trước khi gán giá.

---

# Chạy trong add-in Claude for Excel (Office.js)

Mặc định là **chế độ in-place**: workbook CĐT đang mở được giữ nguyên hoàn toàn (giá trị, font, màu, merge, độ rộng, chiều cao dòng, công thức, sheet khác…). Phần TDI được ghi **ngay bên phải** cột cuối cùng của từng sheet BOQ, cùng dòng với CĐT. Thêm một sheet `Summary` (hoặc `Summary TDI` nếu tên đã có).

Trong add-in không có Python. Mỗi script chạy tối đa **90 giây** và trong một môi trường bị khóa:

- **Không có** `eval`, `new Function`, `import`, `require`, `fetch` để đọc file skill.
- Mỗi script **không nhớ gì** từ script trước.

Vì vậy skill có sẵn hai thư viện JavaScript thuần, đã nén, sinh tự động từ bản Python và đã kiểm tra cho kết quả giống hệt. **Cả file, dù nhiều sheet, chỉ cần 2 script:**

| Script | Dán file nào | Làm gì |
|---|---|---|
| **A: Phân tích** | `scripts/excel/tdi_a.min.js` | **Tự tìm mọi sheet BOQ** (bỏ bìa, chỉ dẫn, tổng hợp). Tự nhận diện dòng tiêu đề, cột, khu, tổng KL, dòng cuối của từng sheet. Phân loại, gán Chủng loại, chọn cột bắt đầu phần TDI |
| **B: Ghi** | `scripts/excel/tdi_b.min.js` | Áp chỗ sửa Chủng loại, ghi phần TDI cho mọi sheet, dòng tổng, Summary gộp mọi hệ, dọn sheet tạm |

Mỗi script tự chia khúc, tự dừng trước 90 giây và lưu chỗ dừng. Khi chưa xong chỉ cần **chạy lại y nguyên**. (`tdi_analyze.js` / `tdi_write.js` là bản đọc được của cùng code, để tra cứu; không dán.)

## Mục tiêu thời gian: 2–3 phút cho cả file

- **Không** tự đọc thử sheet, **không** tự dựng cấu hình: script A tự nhận diện.
- **Chế độ nhanh (mặc định):**
  - Chỉ xử lý `nhom_can_soat`: nhóm có dòng chưa mã (`??`) hoặc mã lẫn lộn trong cùng nhóm. Thường chỉ vài nhóm cho cả file.
  - Nhóm thuần một mã thì tin kết quả tự động (luật khớp khoảng 97%).
  - Nhóm nào không chắc thì **để trống**, rồi báo người dùng cho người làm giá tự lọc cột Chủng loại.
- Chỉ soát kỹ toàn bộ khi người dùng yêu cầu ("soát kỹ Chủng loại").
- Không giải thích dài giữa các lượt chạy. Lượt nào trả `xong: false` thì chạy lại ngay.

## Quy tắc bắt buộc

1. **Chỉ dùng đúng 2 mẫu script dưới đây.** Dán NGUYÊN nội dung file `.min.js` (đọc bằng công cụ đọc file của skill), rồi thêm đúng lời gọi như mẫu. Không sửa thư viện, không tự viết script phân tích / ghi / định dạng, không nạp code bằng chuỗi. Script kiểm tra phụ (nếu thật cần) chỉ được đụng dưới ~1.000 ô.
2. **Trước lượt chạy đầu tiên, nhắc người dùng:** không gõ dở trong ô nào và không mở hộp thoại trong Excel khi Claude đang chạy. Excel ở chế độ sửa ô sẽ chặn mọi lệnh của add-in.
3. **Không đổi chế độ tính** (`calculationMode`).
4. Script không ghi hay định dạng bất kỳ ô nào thuộc phần CĐT, không đổi freeze pane hay bộ lọc của CĐT.
5. **Khi một script quá giờ:** đợi vài giây rồi chạy lại y nguyên; script tự tiếp tục.
6. Nếu công cụ đã cấp sẵn `context`, thay `await Excel.run(async (context) => …)` bằng `return await TDI_A.runAll(context, {…});` (tương tự với B).

## Script A: Phân tích cả file

```js
<dán nguyên tdi_a.min.js>
await Excel.run(async (context) => TDI_A.runAll(context, {}));
```

Tham số tùy chọn (thêm vào `{}`):

- `sheets: ["BOQ MEP", "BOQ PCCC"]`: chỉ làm các sheet này. Dùng khi người dùng chỉ định, hoặc khi tự tìm sai sheet.
- `budget: 55000`: mặc định, không tăng.

Kết quả:

- `xong: false`: chạy lại y nguyên.
- `xong: true`: trả `du_an`, `bo_qua` (sheet đã bỏ qua) và với từng sheet: `tieu_de_dong`, `du_lieu` (dòng đầu–cuối), `khu` (tên@cột), `tong_kl`, `tdi_tu_cot`, `hang_muc`, `chua_ma`, `tong_nhom`, `nhom_can_soat`.

**Liếc nhanh `khu`, `du_lieu`, `tdi_tu_cot` của từng sheet.** Nếu thấy sai (ví dụ nhầm cột đơn giá thành cột khu):

1. Xóa sheet ẩn `_TDI_JOB` và các sheet `_TDI_TMP_*`.
2. Chạy lại A với `sheets` đúng.
3. Nếu vẫn sai, báo người dùng cấu trúc lạ và chuyển sang bản Python.

## Script B: Ghi phần TDI cho cả file

```js
<dán nguyên tdi_b.min.js>
await Excel.run(async (context) => TDI_B.runAll(context, {
  OV: { "BOQ PCCC": { "ctx:<khóa nhóm, chép nguyên văn>": "VAN", "123": "BOM" } },
  info: { project: "", hang_muc: "Cung cấp & lắp đặt hệ thống cơ điện", so_bao_gia: "", ngay: "",
          coefficients: { HSKL: 1, HSVL: 1.25, HSTB: 1.1, HSNC: 1.15, HS_chung: 1, lam_tron: -2, hao_hut: 1.02 } } }));
```

Cách điền:

- **`OV`:** theo từng sheet. `"ctx:<khóa nhóm>"` sửa cả nhóm; `"<số dòng CĐT>"` sửa riêng một dòng. Không có gì cần sửa thì để `OV: {}`.
- **`info.project`:** để trống thì lấy tên dự án tự đọc được ở bước A.
- **Giữa các lượt chạy lại:** giữ nguyên `OV` và `info`.

Kết quả:

- `xong: false`: chạy lại y nguyên.
- `xong: true`: với từng sheet, trả `phan_tdi` (vùng cột TDI), `dong_tong_tdi`, `summary`.

## Muốn làm lại

- **Đang dở:** xóa các sheet ẩn `_TDI_JOB`, `_TDI_TMP_*` và các cột TDI đã ghi, rồi chạy lại A.
- **Đã xong:** xóa các cột TDI, và ô `Z1` của Summary nếu muốn làm lại cả Summary. Nếu chưa xóa, script A sẽ báo lỗi để tránh ghi chồng.

## Chế độ dựng lại (sheet mới, không giữ định dạng CĐT)

Rất ít khi cần trong Excel; khuyên người dùng làm chế độ này bằng Python trên claude.ai / Desktop.

## Lỗi thường gặp

| Lỗi | Cách xử lý |
|---|---|
| `eval is not a function` | Script đã nạp code động. Dán nguyên nội dung file JS vào script thay vì nạp bằng chuỗi |
| `Script execution timed out` | Đợi vài giây rồi **chạy lại y nguyên** script A hoặc B; cả hai tự tiếp tục. Nếu là script tự viết (không phải A/B), dừng lại và chuyển sang dùng A/B |
| "Không tìm thấy sheet BOQ nào" | Tiêu đề bảng không có các chữ quen thuộc (STT / mô tả / đơn vị / khối lượng). Hỏi người dùng tên sheet + dòng tiêu đề, hoặc dùng bản Python (`inspect`) |
| Báo lỗi biểu thức chính quy (`Invalid regular expression`) | Excel trên máy quá cũ, không hỗ trợ lookbehind. Cập nhật Office, hoặc chạy skill trên claude.ai / Desktop |
| Quá giờ dù dữ liệu nhỏ | Thường do người dùng đang gõ dở trong một ô, hoặc đang mở hộp thoại trong Excel. Nhắc thoát ô (Enter / Esc), đóng hộp thoại, rồi chạy lại |
| Phần TDI đè lên dữ liệu CĐT | Cột tự chọn sai do CĐT có cột ẩn / ghi chú xa. Xóa phần đã ghi và các sheet `_TDI_*`; dùng bản Python với `tdi_start_col` |

## Bảng màu (hex) của mẫu TDI

| Vị trí | Màu |
|---|---|
| Tiêu đề bảng | nền `#9DC3E6`, chữ `#C00000` đậm |
| Đơn giá gốc, Tổng KL tạm tính | nền `#FFFF00` |
| Bộ lọc / Chủng loại | nền `#92D050` |
| Diễn giải KL | nền `#FBE5D6` |
| Dòng mục A / I | nền `#BDD7EE` |
| Dòng mục I.1 trở xuống | nền `#D9D9D9` |
| ĐG vật liệu ở dòng hạng mục | nền `#DEEBF7` |
| Chữ công thức phía TDI | `#002060` |

Font: Times New Roman 11.

## BOQ rất lớn

| Số dòng CĐT | Add-in Excel | claude.ai / Desktop (Python) |
|---|---|---|
| ≤ 5.000 (cả file) | Tốt: thường 1 lượt A + 1 lượt B | ~15 giây |
| 5.000 – 20.000 | Làm được: A, B mỗi script vài lượt (chạy lại y nguyên); mỗi lượt dán lại thư viện nên chậm hơn. Gợi ý bản Python | ~30–60 giây |
| > 20.000 | **Khuyên chạy bằng Python** rồi mở file bằng Excel | vài phút |

## Bảo trì

Sửa luật trong `scripts/boq_tdi.py` xong thì chạy `python3 scripts/build_excel_core.py` để sinh lại các file JS (cần `node` và gói `terser` để nén). Giao diện nằm ở `scripts/excel/*_template.js`.
