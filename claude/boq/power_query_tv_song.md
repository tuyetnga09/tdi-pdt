# Link sống thư viện giá vào bảng `TV` bằng Power Query (một chiều: thư mục → BOQ)

_Soạn 22/10/2026. Dùng với skill `covert-boq-cdt-into-boq-tdi` (Bước 2). Chưa chạy thử trên Excel của phòng: lần đầu dùng phải kiểm theo mục 4._

## 1. Hành vi

- Sửa giá trong file thư viện (thư mục `2. Thu vien gia` của dự án), **lưu file**, rồi ở BOQ bấm **Data → Refresh All**: bảng `TV` (sheet `2 Thu vien gia`) đổi theo, mọi công thức `SUMIFS(TV[…])` đổi theo.
- Sửa giá trong sheet `2 Thu vien gia` của BOQ: **file thư viện không đổi** (Power Query chỉ đọc). Lần Refresh sau, giá sửa tay bị ghi đè bằng giá từ thư mục.
- Không tức thì tuyệt đối: phải Refresh (có thể đặt tự Refresh khi mở file hoặc mỗi N phút ở Query Properties). File thư viện phải được lưu; file OneDrive phải ở trạng thái "Always keep on this device".
- Bảng `TV` do query tạo có đủ 19 cột như bảng nạp bằng `tv_loader.py` + `hsloai`: A `Loại vật liệu` … O `Nhóm file`, P `Giá VL thư viện (thô)`, Q `Giá NC thư viện (thô)`, R `HS VL (Loại VL)`, S `HS NC (Loại VL)`. `Giá VL gốc` / `Giá NC gốc` (E, F) = thô × hệ số Loại VL (Quy tắc 24). Cột `Ngày cập nhật` lấy từ ngày sửa file thật.

## 2. Thời điểm chuyển sang link sống (quan trọng)

Làm **ngay sau khi skill chuyển template giao file `_TDI.xlsx`, TRƯỚC khi chạy `boq-gan-ma-gia-thu-vien`**, lúc chưa có công thức nào chứa `TV[`. Lý do: bước 2 dưới đây xóa bảng `TV` snapshot; nếu BOQ đã có công thức `SUMIFS(TV[…])` thì chúng thành `#REF!`.

Kiểm nhanh: Ctrl+F, tìm `TV[` trong Formulas, toàn workbook. Không có kết quả thì chuyển được.

Skill chuyển template vẫn nạp snapshot như cũ để file luôn chạy được; link sống thay thế snapshot đó.

## 3. Cách làm (một lần cho mỗi dự án, khoảng 5 phút)

1. Đóng các file thư viện đang mở. Mở file `_TDI.xlsx`.
2. **Xóa sheet `2 Thu vien gia`** (bảng `TV` snapshot). Việc này giải phóng tên bảng `TV`.
3. **Data → Get Data → From Other Sources → Blank Query**. Trong cửa sổ Query Editor, ô **Name** bên phải: đặt tên query là **`TV`**.
4. **Home → Advanced Editor**, xóa hết, dán code ở mục 6. Sửa 2 dòng đầu:
   - `DuongDan`: mở thư mục `2. Thu vien gia` của dự án trong File Explorer, bấm vào thanh địa chỉ, Ctrl+C, dán vào giữa hai dấu ngoặc kép. Không thêm `\` ở cuối.
   - `MaDuAn`: để `""` nếu tên file thư viện không mang mã dự án; nếu có (ví dụ `TDI25538`) thì điền mã để chỉ lấy file mang mã đó.
5. **Done → Close & Load To… → Table → New worksheet → OK**.
6. Đổi tên sheet mới thành **`2 Thu vien gia`**. Bấm vào một ô của bảng → **Table Design → Table Name**: phải là **`TV`** (không phải `TV_1`; nếu là `TV_1` thì còn tên ẩn `tv` cũ, xóa tên rác ở Formulas → Name Manager rồi đặt lại).
7. Chuột phải vào bảng → **Table → External Data Properties** (hoặc Data → Queries & Connections → chuột phải query → Properties): bật **Preserve column sort/filter/layout**, bật **Refresh data when opening the file** (khuyến nghị), bỏ **Enable background refresh**.
8. Lưu file. Gửi lại cho Claude để kiểm (mục 4) rồi chạy `boq-gan-ma-gia-thu-vien`.

## 4. Kiểm lần đầu (bắt buộc)

Bảng nạp bằng bản cũ của Power Query từng thiếu 685 dòng, nên không tin số dòng bằng mắt.

- Nhờ Claude chạy `tv_loader.py check "<thư mục thư viện ME>" "<file>_TDI.xlsx" --code <mã>` trên file đã Refresh: phải ra `THIẾU 0 | THỪA 0 | LỆCH GIÁ 0`. Với bộ TDI25538 chuẩn là 6.887 dòng, 28 giá trị `Loại vật liệu`.
- Lọc cột `Loại vật liệu` tìm giá trị bắt đầu bằng `(Chưa ánh xạ)`: đó là sheet lạ chưa có trong bảng ánh xạ, báo Claude để bổ sung, không tự đặt tên mới.
- Thiếu hoặc thừa dòng so với `check`: gửi Claude file và kết quả `check`, không dùng tiếp.

## 5. Khi dùng tiếp với `boq-gan-ma-gia-thu-vien`

- Bảng `TV` đã có cột O–R: **không chạy `hsloai`**.
- **Không chạy `tv_loader.py load … --replace`** trên bảng `TV` này (sẽ phá bảng của query). Thư viện đổi thì bấm Refresh.
- `check` báo LỆCH GIÁ do thư viện vừa cập nhật: đó là giá mới đúng ý link sống, Refresh rồi `check` lại.
- Sửa giá thư viện ở **file nguồn** (cột Số liệu gốc), không sửa trong sheet `2 Thu vien gia`.
- **Trước khi nộp thầu: chốt giá.** Refresh lần cuối, lưu bản `_TDI_chot_<ngày>.xlsx`, ở bản này chọn toàn bảng `TV` → Copy → Paste Values, rồi xóa query (Queries & Connections → Delete). Nếu không chốt, ai sửa file thư viện sau đó cũng làm đổi giá BOQ đã duyệt khi Refresh.

## 6. Code Power Query (dán vào Advanced Editor)

```
let
    // ===== SỬA 2 DÒNG NÀY CHO TỪNG DỰ ÁN =====
    DuongDan = "C:\Users\Admin\OneDrive - TDI Group\...\2. Thu vien gia",
    MaDuAn   = "",
    // =========================================

    Chuan = (x as any) as text =>
        if x = null then "" else
        Text.Lower(Text.Combine(List.Select(
            Text.SplitAny(Text.From(x, "en-US"), " #(lf)#(cr)#(tab)#(00A0)"), each _ <> ""), " ")),
    Khoa = (x as any) as text => Text.Remove(Chuan(x), {" ", "_", "-", "."}),
    ToText = (v as any) as nullable text =>
        if v = null then null
        else if v is number then
            (if v = Number.Round(v, 0) then Text.From(Number.Round(v, 0), "en-US") else Text.From(v, "en-US"))
        else let s = Text.Trim(Text.From(v, "en-US")) in if s = "" then null else s,
    Gia = (v as any) as nullable number => if v is number then v else null,
    GiaChu = (v as any, ten as text) as nullable text =>
        if v is text and Text.Trim(v) <> "" and Text.Trim(v) <> "-"
        then ten & " dạng chữ: " & Text.Trim(v) else null,

    // Hệ số Loại VL = 1 (xem mục "Bật hệ số Loại VL" nếu cần khác 1)
    HeSo = (loai as text, cot as text) as number => 1,

    // ----- Ánh xạ Loại vật liệu (28 tên chuẩn của Summary TDI) -----
    MapNhom = {
        {"capdien", "Cáp điện"}, {"olday", "Ống luồn dây"}, {"vankhoa", "Van Khóa Thép"},
        {"ctkhac", "Công tác khác"}, {"odong", "Ống đồng"}, {"tbdien", "Thiết bị Điện"},
        {"tbvesinh", "Thiết bị Vệ Sinh"}, {"baoon", "Bảo ôn"},
        {"pccc", "Thiết bị chữa cháy"}, {"tbbaochay", "Thiết bị báo cháy"}
    },
    MapSheet = {
        {"thangcap", "Thang cáp"}, {"pkthangcap", "Phụ kiện T.cáp"},
        {"mangcap", "Máng cáp"}, {"pkmangcap", "Phụ kiện M.cáp"},
        {"ongthep", "Ống thép"}, {"pkthep", "PK ống thép"}, {"vanthep", "Van Khóa Thép"},
        {"tbdieuhoa", "Điều Hoà"}, {"tbthonggio", "Quạt thông gió"}, {"vangio", "Van gió"},
        {"ongpvc", "Ống PVC"}, {"pkpvc", "PK Ống PVC"},
        {"ongppr", "Ống PPR"}, {"pkppr", "PK Ống PPR"},
        {"onghdpe", "Ống HDPE"}, {"pkhdpe", "PK HDPE"},
        {"onggio", "Ống gió"}, {"pkonggio", "PK Ống gió"}, {"cuagio", "Cửa gió"}
    },
    LoaiVL = (nhom as text, sheet as text) as text =>
        let
            a = List.First(List.Select(MapNhom, each _{0} = Khoa(nhom)), null),
            b = List.First(List.Select(MapSheet, each _{0} = Khoa(sheet)), null)
        in
            if a <> null then a{1} else if b <> null then b{1}
            else "(Chưa ánh xạ) " & nhom & " / " & sheet,

    // ----- Nhận diện tiêu đề cột (cùng luật với tv_loader.py) -----
    LaMa = (c as any) as logical => c <> null and Text.Contains(Chuan(c), "mã hiệu"),
    TenChuan = (t as text, s as text) as nullable text =>
        let vlOk = Text.Contains(t, "số liệu gốc") or t = "" in
        if s = "vật liệu" and vlOk then "Giá VL gốc"
        else if s = "nhân công" and vlOk then "Giá NC gốc"
        else if Text.Contains(t, "nhân công") and Text.Contains(t, "ước") then "Giá NC ước tính"
        else if Text.Contains(t, "mã hiệu") and not Text.Contains(s, "mã hiệu") then "Mã hiệu"
        else if t = "tên công việc" then "Tên công việc"
        else if Text.StartsWith(t, "unit") or Text.Contains(t, "đ. vị") or t = "đvt" then "ĐVT"
        else if t = "description" then "Description"
        else if t = "ghi chú" then "Ghi chú TV"
        else if Text.Contains(t, "nhãn hiệu") and Text.Contains(t, "xuất xứ") then "Nhãn Hiệu Xuất Xứ"
        else null,

    CotNguon = {"Mã hiệu", "Tên công việc", "ĐVT", "Giá VL gốc", "Giá NC gốc",
                "Giá NC ước tính", "Description", "Ghi chú TV", "Nhãn Hiệu Xuất Xứ"},
    CotRa = {"Loại vật liệu", "Mã hiệu", "Tên công việc", "ĐVT", "Giá VL gốc", "Giá NC gốc",
             "Giá NC ước tính", "Description", "Ghi chú TV", "Nhãn Hiệu Xuất Xứ",
             "Sheet nguồn", "File nguồn", "Ngày cập nhật", "Nhóm file",
             "Giá VL thư viện (thô)", "Giá NC thư viện (thô)", "HS VL (Loại VL)", "HS NC (Loại VL)"},

    LayBang = (t as table, nhom as text, sheet as text, hidden as logical, fn as text, ngay as text) as nullable table =>
        let
            Rows = Table.ToRows(t),
            H    = List.PositionOf(List.Transform(List.FirstN(Rows, 30),
                        (r) => List.AnyTrue(List.Transform(r, LaMa))), true),
            Loai = LoaiVL(nhom, sheet),
            HSV  = HeSo(Loai, "HS VL"),
            HSN  = HeSo(Loai, "HS NC"),
            Nhan = sheet & (if hidden then " (sheet ẩn)" else ""),

            XuLyDong = (r as record) as nullable record =>
                let
                    ma  = ToText(r[#"Mã hiệu"]),
                    ten = ToText(r[#"Tên công việc"]),
                    dvt = ToText(r[#"ĐVT"]),
                    vl  = Gia(r[#"Giá VL gốc"]),
                    nc  = Gia(r[#"Giá NC gốc"]),
                    ut  = Gia(r[#"Giá NC ước tính"]),
                    coMa  = ma <> null and ma <> "0",
                    coGia = List.AnyTrue(List.Transform({vl, nc, ut}, each _ <> null and _ <> 0)),
                    giu   = (coMa and dvt <> null) or (coMa and ten <> null and coGia) or (ten <> null and dvt <> null),
                    tu = List.Select({
                            ToText(r[#"Ghi chú TV"]),
                            if not coMa then "Không có mã hiệu - tra theo tên" else null,
                            if dvt = null then "Thiếu ĐVT trong file nguồn" else null,
                            GiaChu(r[#"Giá VL gốc"], "Giá VL gốc"),
                            GiaChu(r[#"Giá NC gốc"], "Giá NC gốc")}, each _ <> null),
                    ghi = if List.Count(tu) = 0 then null else Text.Combine(tu, "; ")
                in
                    if not giu then null else
                    [#"Loại vật liệu" = Loai, #"Mã hiệu" = ma, #"Tên công việc" = ten, #"ĐVT" = dvt,
                     #"Giá VL gốc" = (if vl = null then null else vl * HSV),
                     #"Giá NC gốc" = (if nc = null then null else nc * HSN),
                     #"Giá NC ước tính" = ut,
                     #"Description" = ToText(r[Description]), #"Ghi chú TV" = ghi,
                     #"Nhãn Hiệu Xuất Xứ" = ToText(r[#"Nhãn Hiệu Xuất Xứ"]),
                     #"Sheet nguồn" = Nhan, #"File nguồn" = fn, #"Ngày cập nhật" = ngay, #"Nhóm file" = nhom,
                     #"Giá VL thư viện (thô)" = vl, #"Giá NC thư viện (thô)" = nc,
                     #"HS VL (Loại VL)" = HSV, #"HS NC (Loại VL)" = HSN],

            KetQua =
                if H < 0 then null else
                let
                    Top     = Rows{H},
                    Sub     = if H + 1 < List.Count(Rows) then Rows{H + 1} else {},
                    TopFill = List.Skip(List.Accumulate(Top, {""}, (st, x) =>
                                  st & {if Chuan(x) = "" then List.Last(st) else Chuan(x)}), 1),
                    Ten0    = List.Transform(List.Positions(Top), (j) =>
                                  TenChuan(TopFill{j}, if j < List.Count(Sub) then Chuan(Sub{j}) else "")),
                    Ten     = List.Transform(List.Positions(Ten0), (j) =>
                                  if Ten0{j} <> null and not List.Contains(List.FirstN(Ten0, j), Ten0{j})
                                  then Ten0{j} else null),
                    Idx     = List.Select(List.Positions(Ten), (j) => Ten{j} <> null),
                    TenKeep = List.Transform(Idx, (j) => Ten{j}),
                    BatDau  = H + (if List.AnyTrue(List.Transform(Sub,
                                  (v) => List.Contains({"vật liệu", "nhân công"}, Chuan(v)))) then 2 else 1),
                    Dong    = List.Transform(List.Skip(Rows, BatDau),
                                  (r) => List.Transform(Idx, (j) => if j < List.Count(r) then r{j} else null)),
                    Bang    = Table.FromRows(Dong, TenKeep),
                    Du      = Table.SelectColumns(Bang, CotNguon, MissingField.UseNull),
                    Recs    = List.Select(List.Transform(Table.ToRecords(Du), XuLyDong), each _ <> null)
                in
                    Table.FromRecords(Recs, CotRa)
        in
            KetQua,

    NhomFile = (ten as text) as text =>
        let
            b = Text.BeforeDelimiter(ten, ".", {0, RelativePosition.FromEnd}),
            p = Text.PositionOf(b, ".BG", Occurrence.First, Comparer.OrdinalIgnoreCase),
            s = if p < 0 then b else Text.Middle(b, p + 3)
        in
            Text.Trim(s, {"-", " ", "_"}),

    // ----- Đọc thư mục -----
    Source  = Folder.Files(DuongDan),
    LocFile = Table.SelectRows(Source, each
                  Text.Lower([Extension]) = ".xlsx"
                  and not Text.StartsWith([Name], "~$")
                  and not Text.Contains(Text.Lower([Name]), "_backup")
                  and Text.TrimEnd(Text.Lower([Folder Path]), "\") = Text.TrimEnd(Text.Lower(DuongDan), "\")
                  and (MaDuAn = "" or Text.StartsWith(Text.Upper([Name]), Text.Upper(MaDuAn)))),
    ThemNhom = Table.AddColumn(LocFile, "NhomFile", each NhomFile([Name]), type text),
    ThemNgay = Table.AddColumn(ThemNhom, "NgayCN", each Date.ToText(Date.From([Date modified]), "dd/MM/yyyy"), type text),
    MoFile   = Table.AddColumn(ThemNgay, "Sheets", each Excel.Workbook([Content])),
    Giu      = Table.SelectColumns(MoFile, {"Name", "NhomFile", "NgayCN", "Sheets"}),
    Bung     = Table.ExpandTableColumn(Giu, "Sheets", {"Name", "Data", "Kind", "Hidden"}, {"Sheet", "Data", "Kind", "Hidden"}),
    ChiSheet = Table.SelectRows(Bung, each [Kind] = "Sheet"),
    XuLy     = Table.AddColumn(ChiSheet, "Bang", each
                   LayBang([Data], [NhomFile], [Sheet], [Hidden] = true, [Name], [NgayCN])),
    CoBang   = Table.SelectRows(XuLy, each [Bang] <> null),
    GopLai   = Table.Combine(CoBang[Bang]),
    Kieu     = Table.TransformColumnTypes(GopLai, {
                   {"Giá VL gốc", type nullable number}, {"Giá NC gốc", type nullable number},
                   {"Giá NC ước tính", type nullable number},
                   {"Giá VL thư viện (thô)", type nullable number}, {"Giá NC thư viện (thô)", type nullable number},
                   {"HS VL (Loại VL)", type number}, {"HS NC (Loại VL)", type number}})
in
    Kieu

```

## 7. Bật hệ số Loại VL (tùy chọn)

Mặc định hệ số = 1, nên `Giá VL gốc` / `Giá NC gốc` = giá thư viện thô. Đúng khi bảng "Loại VL – Hệ số VL/NC" trong Summary TDI đang là 1 hết. Nếu đang dùng hệ số khác 1:

1. Tạo một bảng Excel tên **`HS_LoaiVL`** (Insert → Table, Table Design → Table Name) có đúng 3 cột `Loại vật liệu`, `HS VL`, `HS NC`, 28 dòng, tên Loại VL khớp bảng Summary (có thể dùng công thức trỏ về ô Summary để hệ số đổi theo Summary).
2. Trong Advanced Editor, thay đúng dòng `HeSo = (loai as text, cot as text) as number => 1,` bằng đoạn sau:

```
    HSBang = Table.Buffer(try Excel.CurrentWorkbook(){[Name = "HS_LoaiVL"]}[Content]
                 otherwise #table({"Loại vật liệu", "HS VL", "HS NC"}, {})),
    HeSo = (loai as text, cot as text) as number =>
        let
            r = Table.SelectRows(HSBang, each Khoa([#"Loại vật liệu"]) = Khoa(loai)),
            v = if Table.RowCount(r) = 0 then null
                else try Number.From(Table.Column(r, cot){0}) otherwise null
        in
            if v = null then 1 else v,
```

3. Nếu Excel hỏi quyền riêng tư dữ liệu (Formula.Firewall), chọn **Ignore Privacy Levels for this file**.
4. Hệ số trong `HS_LoaiVL` chỉ có hiệu lực sau mỗi lần Refresh (không đổi tức thì khi sửa Summary).

Chưa xác nhận bố cục khối B11:G24 của Summary TDI nên chưa nối tự động; khi cần, nhờ Claude dựng bảng `HS_LoaiVL` từ Summary.

## 8. Hạn chế đã biết

- Chiết khấu ở đầu file nguồn (ví dụ 0,85176 ở file HDPE) và hệ số dòng nhóm con **không** được query đọc; BOQ lấy theo cột Số liệu gốc như `tv_loader.py`. Cột "Báo giá VT chưa chiết khấu" lấy giá trị đã tính trong file nguồn, nên phải lưu file nguồn bằng Excel.
- Đường dẫn trong query là đường dẫn máy người dùng. Máy khác mở file BOQ muốn Refresh phải có cùng đường dẫn hoặc sửa `DuongDan`; không Refresh được thì bảng `TV` vẫn giữ giá lần Refresh cuối.
- Code chưa chạy thử trên dữ liệu thật của phòng (đã kiểm cú pháp bằng parser Power Query).
