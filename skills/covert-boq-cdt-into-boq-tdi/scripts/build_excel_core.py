#!/usr/bin/env python3
"""Sinh scripts/excel/tdi_analyze.js và tdi_write.js từ luật trong boq_tdi.py (để bản Excel và bản Python luôn khớp).

Chạy lại mỗi khi sửa RULES / GENERIC / DOMINANT... trong boq_tdi.py:
    python3 scripts/build_excel_core.py
"""
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import boq_tdi as B  # noqa: E402

# \b của Python hiểu chữ có dấu; \b của JavaScript thì không -> thay bằng ranh giới Unicode
WB = r"(?:(?<=[\p{L}\p{N}_])(?![\p{L}\p{N}_])|(?<![\p{L}\p{N}_])(?=[\p{L}\p{N}_]))"


def js(p):
    out, i = [], 0
    while i < len(p):
        if p[i] == "\\" and i + 1 < len(p):
            out.append(WB if p[i + 1] == "b" else p[i:i + 2])
            i += 2
        else:
            out.append(p[i])
            i += 1
    return "".join(out)


def main():
    # Giữ nguyên cú pháp Python (\b); phần JS tự đổi \b sang ranh giới Unicode khi khởi tạo
    js = lambda p: p  # noqa: E731
    rules = [[code, js(rx.pattern)] for code, rx in B._RULES]
    dom = [[code, js(rx.pattern)] for code, rx in B.DOMINANT]
    data = {
        "RULES": rules,
        "GENERIC": js(B.GENERIC.pattern),
        "DOMINANT": dom,
        "NON_DOMINATED": sorted(B.NON_DOMINATED),
        "PIPE_CODES": sorted(B.PIPE_CODES),
        "ELEC_SYSTEM": js(B.ELEC_SYSTEM.pattern),
        "PIPE_WORD": js(B.PIPE_WORD.pattern),
        "NOT_CONDUIT": js(B.NOT_CONDUIT.pattern),
        "PREFIX": js(B._PREFIX.pattern),
        "CUT": js(B._CUT.pattern),
        "PIPE_HEAD": js(B.detone(r"\s*(đường )?ống\b|\s*pipe")),
        "VAN_HEAD": js(B.detone(r"\bvan\b|văn chặn")),
        "PHU_KIEN": js(B.detone(r"\s*phụ kiện")),
        "PHU_KIEN_HEAD": js(B.PHU_KIEN_HEAD.pattern),
        "VTP_EXCEPT": js(B.VTP_EXCEPT.pattern),
        "LOWV_SYSTEM": js(B.LOWV_SYSTEM.pattern),
        "CONDUIT_HEAD": js(B.CONDUIT_HEAD.pattern),
        "HDR_KEYS": {k: B.detone(v) for k, v in B.HDR_KEYS.items()},
        "SKIP_SHEET": B.SKIP_SHEET.pattern,
        "TOT_HEAD": B.detone(r"tổng|total|cộng|sum"),
        "TOT_ROW": B.detone(r"^(tổng cộng|tổng giá trị|tổng hợp|cộng\s*(:|$)|total\b|grand total)"),
        "PROJ": B.detone(r"dự án|project|công trình"),
        "THIET_BI": sorted(B.THIET_BI),
        "LINEAR_UNITS": sorted(B.LINEAR_UNITS),
    }
    dw = {"THIET_BI": data.pop("THIET_BI"), "LINEAR_UNITS": data.pop("LINEAR_UNITS")}
    for name, key, payload in (("analyze", "/*__DATA__*/null", data), ("write", "/*__DATA_W__*/null", dw)):
        tpl = (HERE / "excel" / f"{name}_template.js").read_text(encoding="utf-8")
        body = tpl.replace(key, json.dumps(payload, ensure_ascii=False))
        out = HERE / "excel" / f"tdi_{name}.js"
        out.write_text(body, encoding="utf-8")
        print("Đã ghi", out.name, f"({len(body):,} ký tự)")
        if name == "write":
            # tdi_b.js: chỉ phần GHI CHẾ ĐỘ IN-PLACE (bỏ toàn bộ phần dựng sheet mới) -> nhẹ hơn để dán vào script
            b = body.replace("PHẦN 2: GHI & ĐỊNH DẠNG", "PHẦN B: GHI IN-PLACE").replace("const TDI_W = ", "const TDI_B = ")
            for start, end in (("  // Ghi dòng 1-12", "  // ----------------------------- Office.js: định dạng vùng dữ liệu"),
                               ("  // Gọi cho từng lô (r1..r2)", "  // ----------------------------- Office.js: Summary"),
                               ("  // Văn bản bắt đầu bằng", "  // ======================= CHẾ ĐỘ IN-PLACE"),
                               ("  const planInplace = ", "  // ----------------------------- Office.js: đầu sheet")):
                i, j = b.index(start), b.index(end)
                b = b[:i] + b[j:]
            k = b.index("  return { runAll, run, applyOV, layoutInplace")
            k2 = b.index("};", k)
            b = b[:k] + "  return { runAll, run, layoutInplace, headerPlan, buildRight, rightBlock, rightTotal, formatRight, writeHeaderRight, writeSummary, layout, col, fromTmp, applyOV, priceGroup, PAL };" + b[k2 + 2:]
            b = b.replace("// ============================== HẾT PHẦN 2 ==", "// ============================== HẾT PHẦN B ==")
            (HERE / "excel" / "tdi_b.js").write_text(b, encoding="utf-8")
    # Bản nén để dán vào add-in (Claude chỉ phải viết ra ít ký tự hơn)
    import subprocess
    for src, dst, title in (("tdi_analyze.js", "tdi_a.min.js", "TDI_A: phân tích BOQ CĐT"), ("tdi_b.js", "tdi_b.min.js", "TDI_B: ghi phần TDI in-place")):
        r = subprocess.run(["node", str(HERE / "excel" / "minify.js"), str(HERE / "excel" / src), str(HERE / "excel" / dst)],
                           capture_output=True, text=True)
        if r.returncode != 0:
            print("  (không nén được", dst, ":", (r.stderr or "").strip(), ")")
            continue
        code = (HERE / "excel" / dst).read_text(encoding="utf-8")
        head = f"/*! {title} | skill covert-boq-cdt-into-boq-tdi | JS thuần, không eval. Dán NGUYÊN file này. */\n"
        (HERE / "excel" / dst).write_text(head + code, encoding="utf-8")
        print("Đã ghi", dst, f"({len(head + code):,} ký tự)")
    old_batch = HERE / "excel" / "tdi_batch.js"
    if old_batch.exists():
        old_batch.unlink()


if __name__ == "__main__":
    main()
