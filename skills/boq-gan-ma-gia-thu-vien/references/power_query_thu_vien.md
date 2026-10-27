# Dựng sheet Thư viện giá bằng Power Query

Người dùng tự làm các bước này trong Excel. Add-in Claude không tạo được Power Query, nên Claude chỉ hướng dẫn và đưa code.

Thư mục thư viện chứa 18 file dạng `TDIxxx - 01.BG-Cap dien_.xlsx` … `TDIxxx - 18.BG TB Bao Chay.xlsx`. Tiền tố `TDIxxx` đổi theo mã dự án. Code dưới đây không phụ thuộc vào tiền tố đó.

## Các bước dán code

1. Mở file BOQ (query phải nằm **trong chính file BOQ** thì kết quả mới đổ vào đó).
2. Nếu đã có query `2 Thu vien gia`: vào **Data → Queries & Connections**, chuột phải vào query → **Edit**. Nếu chưa có: vào **Data → Get Data → From Other Sources → Blank Query**.
3. Vào **Home → Advanced Editor**, xóa hết code cũ, dán code ở mục dưới. Đường dẫn không nằm trong query: tạo sheet `Cau hinh`, ô A1 ghi `DuongDanThuVien`, ô B1 dán đường dẫn thư mục thư viện của máy này, chọn A1:B1 → Ctrl+T (không tick My table has headers) → đặt tên bảng `CauHinh`.
4. Bấm **Done → Close & Load**. (Tùy chọn) Bấm vào một ô của bảng → **Table Design → Table Name** → đặt `TV` cho công thức gọn.

Đường dẫn là thư mục thư viện ME **của dự án** (thư mục `2. Thu vien gia` nằm cùng cấp file BOQ); dự án chưa có thư mục riêng mới dùng thư mục chung `2.Thu vien gia` của người dùng (khi đó cần lọc thêm theo mã dự án). Không trỏ vào file `2.Thu vien gia XD.xlsx`: đó là thư viện XD, nạp qua sheet `TVGIA`, không qua Power Query.

Lấy đường dẫn: mở thư mục **chứa trực tiếp 18 file** trong File Explorer, click vào thanh địa chỉ, Ctrl+C. Dán vào ô B1 của bảng `CauHinh` (không cần dấu ngoặc kép, có hay không có `\` ở cuối đều được). Máy khác chỉ cần sửa ô B1 rồi Data → Refresh All.

## Code

```
let
    // Đường dẫn lấy từ ô cấu hình (bảng CauHinh, ô B1), mỗi máy tự điền; không ghi cứng đường dẫn vào query.
    DuongDan = Excel.CurrentWorkbook(){[Name="CauHinh"]}[Content]{0}[Column2],
    Source = Folder.Files(if Text.EndsWith(DuongDan, "\") then DuongDan else DuongDan & "\"),
    LocFile = Table.SelectRows(Source, each Text.Lower([Extension]) = ".xlsx"
                and not Text.StartsWith([Name], "~$")),
    ThemNhom = Table.AddColumn(LocFile, "Nhóm VT", each
                Text.Trim(Text.AfterDelimiter(
                    Text.BeforeDelimiter([Name], ".", {0, RelativePosition.FromEnd}),
                    ".BG"), {"-", " ", "_"}), type text),
    GiuCot = Table.SelectColumns(ThemNhom, {"Nhóm VT", "Content"}),
    MoFile = Table.AddColumn(GiuCot, "Sheets", each Excel.Workbook([Content])),
    BoContent = Table.RemoveColumns(MoFile, {"Content"}),
    BungSheet = Table.ExpandTableColumn(BoContent, "Sheets",
                {"Name", "Data", "Kind", "Hidden"}, {"Sheet", "Data", "Kind", "Hidden"}),
    ChiSheet = Table.SelectRows(BungSheet, each [Kind] = "Sheet" and [Hidden] = false),

    LamSach = (x) => Text.Lower(Text.Combine(List.Select(
                Text.SplitAny(Text.From(x), " #(lf)#(cr)#(tab)"), each _ <> ""), " ")),
    LaMaHieu = (c) => c <> null and Text.Contains(LamSach(c), "mã hiệu"),

    DoiTenChuan = (ten as text) as text =>
        let t = LamSach(ten) in
        if Text.Contains(t, "mã hiệu") then "Mã hiệu"
        else if Text.Contains(t, "tên công việc") then "Tên công việc"
        else if Text.Contains(t, "description") then "Description"
        else if Text.Contains(t, "unit") or Text.Contains(t, "đvt") or Text.Contains(t, "đơn vị")
             or Text.Contains(Text.Remove(t, " "), "đ.vị") then "ĐVT"
        else if Text.Contains(t, "báo giá vt") then "Giá VL chưa CK"
        else if Text.Contains(t, "nhân công ước") then "Giá NC ước tính"
        else if t = "vật liệu" then "Giá VL gốc"
        else if t = "nhân công" then "Giá NC gốc"
        else if Text.Contains(t, "ghi chú") then "Ghi chú TV"
        else ten,

    LayBang = (t as table, nhom as text, sheet as text) as nullable table =>
        let
            Rows = Table.ToRows(t),
            ViTri = List.PositionOf(List.Transform(List.FirstN(Rows, 30),
                        (r) => List.AnyTrue(List.Transform(r, LaMaHieu))), true),
            KetQua = if ViTri < 0 then null else
                let
                    Dong1 = Rows{ViTri},
                    Dong2 = if ViTri + 1 < List.Count(Rows) then Rows{ViTri + 1} else {},
                    ViTriMa = List.PositionOf(List.Transform(Dong1, LaMaHieu), true),
                    LaTieuDePhu = List.Count(Dong2) > 0 and
                        (Dong2{ViTriMa} = null or Text.Trim(Text.From(Dong2{ViTriMa})) = ""),
                    TieuDe0 = List.Transform(List.Positions(Dong1), (i) =>
                        let
                            phu = if LaTieuDePhu then Dong2{i} else null,
                            chinh = Dong1{i},
                            ten = if phu <> null and Text.Trim(Text.From(phu)) <> "" then Text.From(phu)
                                  else if chinh <> null and Text.Trim(Text.From(chinh)) <> "" then Text.From(chinh)
                                  else "Cot" & Text.From(i + 1)
                        in DoiTenChuan(ten)),
                    TieuDe = List.Transform(List.Positions(TieuDe0), (i) =>
                        if List.Contains(List.FirstN(TieuDe0, i), TieuDe0{i})
                        then TieuDe0{i} & "_" & Text.From(i + 1) else TieuDe0{i}),
                    DuLieu = Table.Skip(t, ViTri + (if LaTieuDePhu then 2 else 1)),
                    DatTen = Table.RenameColumns(DuLieu, List.Zip({Table.ColumnNames(DuLieu), TieuDe})),
                    ThemCot = Table.AddColumn(Table.AddColumn(DatTen, "Nhóm VT", each nhom),
                                "Sheet nguồn", each sheet)
                in ThemCot
        in KetQua,

    XuLy = Table.AddColumn(ChiSheet, "Bang", each LayBang([Data], [Nhóm VT], [Sheet])),
    CoBang = Table.SelectRows(XuLy, each [Bang] <> null),
    GopLai = Table.Combine(CoBang[Bang]),
    LocDong = Table.SelectRows(GopLai, each
                [Mã hiệu] <> null and Text.Trim(Text.From([Mã hiệu])) <> ""
                and not LaMaHieu([Mã hiệu])
                and Record.FieldOrDefault(_, "ĐVT") <> null
                and Text.Trim(Text.From(Record.FieldOrDefault(_, "ĐVT"))) <> ""),
    CotGia = List.Select({"Giá VL gốc", "Giá NC gốc", "Giá VL chưa CK", "Giá NC ước tính"},
                each List.Contains(Table.ColumnNames(LocDong), _)),
    DoiSo = Table.TransformColumns(LocDong, List.Transform(CotGia, (c) =>
                {c, each let n = try Number.From(_) otherwise null in if n = 0 then null else n,
                 type nullable number})),
    Cot = Table.ColumnNames(DoiSo),
    Dau = List.Select({"Nhóm VT", "Mã hiệu", "Tên công việc", "ĐVT",
                "Giá VL gốc", "Giá NC gốc", "Giá VL chưa CK", "Giá NC ước tính"},
                each List.Contains(Cot, _)),
    SapXep = Table.ReorderColumns(DoiSo, Dau & List.RemoveItems(Cot, Dau & {"Sheet nguồn"}) & {"Sheet nguồn"})
in
    SapXep
```

## Code làm gì (để giải thích khi người dùng hỏi)

| Bước | Lý do |
|---|---|
| Lấy `Nhóm VT` từ tên file: phần sau `.BG`, bỏ đuôi, cắt `-`, khoảng trắng, `_` | Tên file có hai kiểu viết (`BG-PPR`, `BG Van Khoa`), và tên nhóm cũng chứa dấu chấm (`06.BG-PPR`), nên phải cắt ở dấu chấm **cuối cùng** |
| Bỏ file `~$…` | File tạm Excel sinh ra khi có người đang mở file thư viện |
| Dò dòng tiêu đề chứa "mã hiệu" trong 30 dòng đầu | Phía trên bảng có tiêu đề "BẢNG GIÁ…", ô chiết khấu, ngày lấy giá |
| Ghép tiêu đề 2 tầng | "Số liệu gốc" ở dòng trên, "Vật liệu / Nhân công" ở dòng dưới |
| Đổi tên cột về tên chuẩn | 18 file đặt tên cột hơi khác nhau, mà Table.Combine cần tên giống nhau |
| Bỏ dòng không có ĐVT | Loại các dòng nhóm con, ví dụ "Cáp điện CV-1x (Cu/PVC)" kèm hệ số 1.00 / 0.85 |
| Giá "-" hoặc 0 → trống | Để không lấy nhầm giá 0 |

Hạn chế: các **hệ số của dòng nhóm con** và **chiết khấu ở đầu file** (ví dụ 0.840 / 0.75) không được mang theo. Nếu phòng tính giá theo các hệ số này, cần bổ sung code: điền xuống (fill down) hệ số của dòng nhóm con, và đọc ô nằm dưới chữ "Chiết khấu".

## Xử lý sự cố

| Lỗi / hiện tượng | Cách xử lý |
|---|---|
| `DataSource.NotFound … couldn't find the folder` | Sai đường dẫn. Lấy lại từ thanh địa chỉ của File Explorer. Thư mục thường lồng 2 cấp cùng tên. |
| `The process cannot access the file … being used by another process` | Đóng file thư viện đang mở. Chờ OneDrive đồng bộ xong (dấu ✓). Vẫn lỗi thì End task các tiến trình Excel chạy ngầm, hoặc tạm dừng OneDrive 2 giờ. |
| Chạy rất chậm, hoặc không đọc được file trên OneDrive | Chuột phải vào thư mục → **Always keep on this device** |
| Thiếu một nhóm trong `Nhóm VT` | File đó không có ô tiêu đề chứa "mã hiệu" (có thể là "Mã VT", "Mã số", hoặc gõ bằng bảng mã khác). Sửa từ khóa trong `LaMaHieu`, hoặc sửa tiêu đề trong file thư viện. |
| Có cột `Cot7`, `Số liệu gốc`, `…_8` | Tiêu đề của file đó khác mẫu. Thêm luật đổi tên trong `DoiTenChuan`. |
| Cột giá của một nhóm trống hết | Tiêu đề cột giá của file đó không khớp các tên đã biết. Xem file gốc rồi thêm luật đổi tên. |

Khi Refresh để cập nhật giá, nhớ đóng hết các file thư viện trước.