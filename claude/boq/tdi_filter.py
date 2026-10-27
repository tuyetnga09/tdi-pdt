#!/usr/bin/env python3
"""tdi_filter.py - gắn bộ lọc (AutoFilter) cho mọi cột của sheet BOQ TDI (inplace): CĐT + TDI + cột phụ gán mã.

Ghi bằng XML (zipfile), KHÔNG mở-lưu bằng openpyxl => giữ nguyên externalLinks, bảng TV, định dạng, công thức.
Chạy được ở mọi thời điểm: ngay sau `boq_tdi.py run`, sau khi nạp thư viện, sau khi gán mã (chạy lại để mở rộng ref).

Dùng:
  python3 tdi_filter.py apply FILE [--out OUT] [--config /home/claude/boq/config.json]
                        [--sheet TÊN] [--header-row N] [--first-row N] [--last-row N] [--first-col A] [--last-col BQ]
                        [--keep-criteria] [--no-freeze] [--dry]
  python3 tdi_filter.py verify FILE      # in ref bộ lọc của từng sheet + kiểm vùng lọc + kiểm vùng cố định (freeze)

Quy tắc ref (khi không truyền tay):
  - dòng tiêu đề lọc  = dòng tiêu đề cuối của bảng (config: header_rows[-1]); là dòng THẤP NHẤT của tiêu đề nhiều tầng.
  - cột đầu           = cột đầu của bộ lọc CĐT nếu có, nếu không là cột đầu có chữ ở dòng tiêu đề.
  - cột cuối          = cột cuối có chữ ở dòng tiêu đề (đã gồm Ghi Chú TDI và 5 cột phụ gán mã nếu có).
  - dòng cuối         = data_end_row của config (KHÔNG gồm dòng "Tổng cộng trước thuế (TDI)" ngay dưới).
Mỗi sheet chỉ có 1 AutoFilter (giới hạn của Excel); ref mới phủ từ cột đầu CĐT tới cột cuối TDI.

Vùng cố định (freeze pane):
  Nút lọc nằm ở dòng tiêu đề lọc (dòng THẤP NHẤT của tiêu đề nhiều tầng). Nếu freeze chỉ giữ tới dòng trên đó
  (vd tiêu đề dòng 4-6 nhưng freeze ySplit=5) thì khi cuộn xuống dòng lọc trôi mất => người dùng "không thấy bộ lọc".
  `apply` tự nâng ySplit của pane đang freeze lên bằng dòng tiêu đề lọc (không bao giờ giảm, không tạo freeze mới
  khi sheet chưa có freeze, chỉ cảnh báo). Dùng --no-freeze để bỏ qua.
"""
import argparse
import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(s):
    n = 0
    for ch in s.upper():
        n = n * 26 + ord(ch) - 64
    return n


def split_ref(ref):
    m = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?", ref.strip())
    if not m:
        raise ValueError(f"ref lạ: {ref}")
    c1, r1, c2, r2 = m.groups()
    return col_index(c1), int(r1), col_index(c2 or c1), int(r2 or r1)


def read_sheet_map(z):
    """tên sheet -> (đường dẫn xml, localSheetId)"""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2t = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        tag = m.group(0)
        i = re.search(r'\bId="([^"]+)"', tag)
        t = re.search(r'\bTarget="([^"]+)"', tag)
        if i and t:
            tg = t.group(1)
            rid2t[i.group(1)] = tg.lstrip("/") if tg.startswith("/") else "xl/" + tg
    out = {}
    for idx, m in enumerate(re.finditer(r"<sheet\b[^>]*>", wb)):
        tag = m.group(0)
        nm = re.search(r'\bname="([^"]*)"', tag)
        rid = re.search(r'\br:id="([^"]+)"', tag) or re.search(r'\bid="([^"]+)"', tag)
        if nm and rid:
            name = (nm.group(1).replace("&amp;", "&").replace("&apos;", "'").replace("&quot;", '"')
                    .replace("&lt;", "<").replace("&gt;", ">"))
            out[name] = (rid2t[rid.group(1)], idx)
    return out


def shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    x = z.read("xl/sharedStrings.xml").decode("utf-8")
    res = []
    for si in re.finditer(r"<si\b[^>]*>(.*?)</si>", x, re.S):
        res.append("".join(re.findall(r"<t\b[^>]*>(.*?)</t>", si.group(1), re.S)))
    return res


def header_cols(xml, row, sst):
    """các cột có chữ ở dòng `row`"""
    m = re.search(rf'<row\b[^>]*\br="{row}"[^>]*?(?:/>|>(.*?)</row>)', xml, re.S)
    cols = []
    if not m or not m.group(1):
        return cols
    for c in re.finditer(r'<c\b([^>]*?)(?:/>|>(.*?)</c>)', m.group(1), re.S):
        attrs, body = c.group(1), c.group(2) or ""
        ref = re.search(r'\br="([A-Z]+)\d+"', attrs)
        if not ref:
            continue
        t = re.search(r'\bt="(\w+)"', attrs)
        v = re.search(r"<v>(.*?)</v>", body, re.S)
        is_ = re.search(r"<is>.*?<t\b[^>]*>(.*?)</t>", body, re.S)
        txt = None
        if t and t.group(1) == "s" and v:
            i = int(v.group(1))
            txt = sst[i] if i < len(sst) else None
        elif is_:
            txt = is_.group(1)
        elif v and not (t and t.group(1) == "e"):
            txt = v.group(1)
        if txt not in (None, ""):
            cols.append(col_index(ref.group(1)))
    return cols


AF_RE = re.compile(r"<autoFilter\b[^>]*?(?:/>|>.*?</autoFilter>)", re.S)
AFTER_SHEETDATA = re.compile(r"\s*<(sheetCalcPr|sheetProtection|protectedRanges|scenarios)\b[^>]*?(?:/>|>.*?</\1>)", re.S)
PANE_RE = re.compile(r"<pane\b[^>]*?/?>")


def get_freeze(xml):
    """(xSplit, ySplit, frozen?) của pane đầu tiên; ySplit=0 nếu không có"""
    m = PANE_RE.search(xml)
    if not m:
        return 0, 0, False
    tag = m.group(0)
    xs = re.search(r'\bxSplit="([\d.]+)"', tag)
    ys = re.search(r'\bySplit="([\d.]+)"', tag)
    frozen = bool(re.search(r'\bstate="frozen(?:Split)?"', tag))
    return (int(float(xs.group(1))) if xs else 0), (int(float(ys.group(1))) if ys else 0), frozen


def ensure_freeze(xml, hdr):
    """Nâng ySplit của pane đang freeze lên >= dòng tiêu đề lọc `hdr`. Trả (xml_mới, thông báo hoặc None)."""
    m = PANE_RE.search(xml)
    if not m:
        return xml, f"CẢNH BÁO: sheet chưa freeze; nút lọc ở dòng {hdr} sẽ trôi khi cuộn (đặt freeze tay nếu cần)"
    tag = m.group(0)
    xs, ys, frozen = get_freeze(xml)
    if not frozen:
        return xml, f"CẢNH BÁO: pane không ở trạng thái freeze; nút lọc dòng {hdr} có thể trôi khi cuộn"
    if ys >= hdr:
        return xml, None
    new = re.sub(r'\bySplit="[^"]*"', f'ySplit="{hdr}"', tag) if re.search(r'\bySplit=', tag) \
        else tag.replace("<pane ", f'<pane ySplit="{hdr}" ', 1)
    # topLeftCell của vùng cuộn phải nằm dưới dòng freeze
    tl = re.search(r'\btopLeftCell="([A-Z]+)(\d+)"', new)
    if tl and int(tl.group(2)) <= hdr:
        new = new.replace(tl.group(0), f'topLeftCell="{tl.group(1)}{hdr + 1}"', 1)
    # pane đang chỉ có xSplit (activePane=topRight) -> sau khi thêm ySplit phải là bottomRight
    if xs and not ys and 'activePane="topRight"' in new:
        new = new.replace('activePane="topRight"', 'activePane="bottomRight"', 1)
    return xml[:m.start()] + new + xml[m.end():], f"freeze ySplit {ys} -> {hdr}"


def set_autofilter(xml, ref, keep_criteria):
    old = AF_RE.search(xml)
    if old:
        tag = old.group(0)
        children = ""
        if keep_criteria:
            mm = re.search(r">(.*)</autoFilter>", tag, re.S)
            children = mm.group(1) if mm else ""
        new = f'<autoFilter ref="{ref}">{children}</autoFilter>' if children else f'<autoFilter ref="{ref}"/>'
        return xml[:old.start()] + new + xml[old.end():], old.group(0)
    i = xml.find("</sheetData>")
    if i < 0:
        raise RuntimeError("sheet không có sheetData")
    pos = i + len("</sheetData>")
    while True:
        m = AFTER_SHEETDATA.match(xml, pos)
        if not m:
            break
        pos = m.end()
    return xml[:pos] + f'<autoFilter ref="{ref}"/>' + xml[pos:], None


def set_filter_name(wbxml, sheet_name, local_id, ref):
    c1, r1, c2, r2 = split_ref(ref)
    q = "'" + sheet_name.replace("'", "''") + "'"
    q = q.replace("&", "&amp;")
    rng = f"{q}!${col_letter(c1)}${r1}:${col_letter(c2)}${r2}"
    dn = f'<definedName name="_xlnm._FilterDatabase" localSheetId="{local_id}" hidden="1">{rng}</definedName>'
    pat = re.compile(rf'<definedName\b[^>]*name="_xlnm\._FilterDatabase"[^>]*localSheetId="{local_id}"[^>]*>.*?</definedName>'
                     rf'|<definedName\b[^>]*localSheetId="{local_id}"[^>]*name="_xlnm\._FilterDatabase"[^>]*>.*?</definedName>', re.S)
    if pat.search(wbxml):
        return pat.sub(dn, wbxml, count=1)
    if "</definedNames>" in wbxml:
        return wbxml.replace("</definedNames>", dn + "</definedNames>", 1)
    if re.search(r"<definedNames\s*/>", wbxml):
        return re.sub(r"<definedNames\s*/>", f"<definedNames>{dn}</definedNames>", wbxml, count=1)
    m = re.search(r"</sheets>(\s*<functionGroups\b.*?/>|\s*<externalReferences>.*?</externalReferences>)*", wbxml, re.S)
    return wbxml[:m.end()] + f"<definedNames>{dn}</definedNames>" + wbxml[m.end():]


def plan_for_sheet(z, name, path, cfg_sheet, a):
    xml = z.read(path).decode("utf-8")
    if "<tableParts" in xml:
        # AutoFilter của sheet không được chồng lên Table (ListObject)
        tparts = re.findall(r'<tablePart\b', xml)
        if tparts and not a.force_tables:
            return None, f"BỎ QUA: sheet có {len(tparts)} Table (ListObject); AutoFilter sheet sẽ chồng lên Table. Dùng --force-tables nếu chắc chắn."
    sst = shared_strings(z)
    old = AF_RE.search(xml)
    old_ref = None
    if old:
        m = re.search(r'\bref="([^"]+)"', old.group(0))
        old_ref = m.group(1) if m else None
    hdr = a.header_row or (cfg_sheet.get("header_rows") or [None, None])[-1] if cfg_sheet else a.header_row
    if not hdr and old_ref:
        hdr = split_ref(old_ref)[1]
    if not hdr:
        return None, "THIẾU dòng tiêu đề: truyền --header-row N hoặc --config"
    last_row = a.last_row or (cfg_sheet or {}).get("data_end_row")
    if not last_row:
        return None, "THIẾU dòng cuối: truyền --last-row N hoặc --config (data_end_row)"
    cols = header_cols(xml, hdr, sst)
    if not cols:
        return None, f"Dòng {hdr} không có chữ nào, kiểm lại --header-row"
    c_first = col_index(a.first_col) if a.first_col else (split_ref(old_ref)[0] if old_ref else min(cols))
    c_last = col_index(a.last_col) if a.last_col else max(cols + ([split_ref(old_ref)[2]] if old_ref else []))
    first_row = a.first_row or hdr
    ref = f"{col_letter(c_first)}{first_row}:{col_letter(c_last)}{last_row}"
    return {"ref": ref, "old_ref": old_ref, "xml": xml, "header_row": hdr}, None


def cmd_apply(a):
    src = Path(a.file)
    out = Path(a.out) if a.out else src
    cfg = {}
    cp = Path(a.config)
    if cp.exists():
        cfg = json.loads(cp.read_text(encoding="utf-8"))
    cfg_by = {s.get("output_sheet") or s.get("source_sheet"): s for s in cfg.get("sheets", [])}
    with zipfile.ZipFile(src) as z:
        smap = read_sheet_map(z)
        targets = [a.sheet] if a.sheet else [n for n in cfg_by if n in smap]
        if not targets:
            sys.exit("Không xác định được sheet: truyền --sheet hoặc --config")
        files = {n: z.read(n) for n in z.namelist()}
        infos = {i.filename: i for i in z.infolist()}
        wbxml = z.read("xl/workbook.xml").decode("utf-8")
        report = []
        for name in targets:
            if name not in smap:
                report.append((name, "KHÔNG THẤY sheet"))
                continue
            path, lid = smap[name]
            plan, err = plan_for_sheet(z, name, path, cfg_by.get(name), a)
            if err:
                report.append((name, err))
                continue
            new_xml, _ = set_autofilter(plan["xml"], plan["ref"], a.keep_criteria)
            fz = ""
            if not a.no_freeze:
                new_xml, fmsg = ensure_freeze(new_xml, plan["header_row"])
                fz = f"; {fmsg}" if fmsg else ""
            files[path] = new_xml.encode("utf-8")
            wbxml = set_filter_name(wbxml, name, lid, plan["ref"])
            report.append((name, f"ref {plan['old_ref'] or '(chưa có)'} -> {plan['ref']} (tiêu đề dòng {plan['header_row']}){fz}"))
        files["xl/workbook.xml"] = wbxml.encode("utf-8")
        # ghi bằng tên mới nếu out khác src để không đọc-ghi cùng file
        tmp = out.with_suffix(out.suffix + ".tmp")
        if a.dry:
            for n, m in report:
                print(f"[{n}] {m}")
            return
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            order = [i.filename for i in z.infolist()]
            for n in order:
                zi = zipfile.ZipInfo(n, date_time=infos[n].date_time)
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.external_attr = infos[n].external_attr
                zo.writestr(zi, files[n])
    shutil.move(str(tmp), str(out))
    for n, m in report:
        print(f"[{n}] {m}")
    print(f"=> {out}")


def cmd_verify(a):
    bad = 0
    with zipfile.ZipFile(a.file) as z:
        smap = read_sheet_map(z)
        wbxml = z.read("xl/workbook.xml").decode("utf-8")
        for name, (path, lid) in smap.items():
            xml = z.read(path).decode("utf-8")
            m = AF_RE.search(xml)
            if not m:
                continue
            r = re.search(r'\bref="([^"]+)"', m.group(0)).group(1)
            has_name = bool(re.search(rf'name="_xlnm\._FilterDatabase"[^>]*localSheetId="{lid}"|localSheetId="{lid}"[^>]*name="_xlnm\._FilterDatabase"', wbxml))
            c1, r1, c2, r2 = split_ref(r)
            tp = "<tableParts" in xml
            note = []
            if not has_name:
                note.append("thiếu defined name _FilterDatabase (Excel tự tạo khi mở)")
            if tp:
                note.append("CÓ TABLE: AutoFilter có thể chồng Table")
                bad += 1
            _, ys, frozen = get_freeze(xml)
            if not frozen:
                note.append(f"CẢNH BÁO: chưa freeze, nút lọc dòng {r1} trôi khi cuộn")
            elif ys < r1:
                note.append(f"LỖI: freeze chỉ tới dòng {ys} < dòng lọc {r1} => nút lọc biến mất khi cuộn (chạy `apply`)")
                bad += 1
            print(f"[{name}] {r} ({col_letter(c1)}..{col_letter(c2)}, {c2 - c1 + 1} cột, dòng {r1}-{r2}) {'; '.join(note)}")
    sys.exit(1 if bad else 0)


def main():
    p = argparse.ArgumentParser()
    sp = p.add_subparsers(dest="cmd", required=True)
    a = sp.add_parser("apply")
    a.add_argument("file")
    a.add_argument("--out")
    a.add_argument("--config", default="/home/claude/boq/config.json")
    a.add_argument("--sheet")
    a.add_argument("--header-row", type=int)
    a.add_argument("--first-row", type=int)
    a.add_argument("--last-row", type=int)
    a.add_argument("--first-col")
    a.add_argument("--last-col")
    a.add_argument("--keep-criteria", action="store_true", help="giữ điều kiện lọc đang bật của CĐT")
    a.add_argument("--force-tables", action="store_true")
    a.add_argument("--no-freeze", action="store_true", help="không chỉnh vùng cố định (freeze pane)")
    a.add_argument("--dry", action="store_true")
    v = sp.add_parser("verify")
    v.add_argument("file")
    args = p.parse_args()
    {"apply": cmd_apply, "verify": cmd_verify}[args.cmd](args)


if __name__ == "__main__":
    main()
