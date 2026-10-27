// ============================================================================
// BOQ CĐT -> TDI | PHẦN 2: GHI & ĐỊNH DẠNG (dùng trong add-in Claude for Excel)
// JavaScript thuần: KHÔNG eval / new Function / import / require.
// Dán NGUYÊN khối này vào đầu mỗi script ghi dữ liệu / định dạng / Summary.
// ============================================================================
const TDI_W = (() => {
  const D = /*__DATA_W__*/null;
  const THIET_BI = new Set(D.THIET_BI), LINEAR = new Set(D.LINEAR_UNITS);
  const norm = (v) => (v === null || v === undefined ? "" : String(v).normalize("NFC").trim());
  // ---- đọc sheet CĐT theo khối: cols = {stt:"A", desc:"B", spec:"C"|null, code, brand, origin, unit:"G", zones:["H","I","J"]}
  const cIdx = (L) => { let n = 0; for (const ch of String(L).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
  const cLet = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  function srcSpan(cols) {
    const all = ["stt", "desc", "spec", "code", "brand", "origin", "unit"].map((k) => cols[k]).filter(Boolean).concat(cols.zones || []);
    const idx = all.map(cIdx); return [Math.min(...idx), Math.max(...idx)];
  }
  // Xếp lệnh đọc khối r1..r2 (values + chữ đậm cột mô tả). Gọi await context.sync() rồi đưa kết quả vào srcRows().
  function loadSource(sheet, cols, r1, r2) {
    const [a, b] = srcSpan(cols);
    const rg = sheet.getRange(`${cLet(a)}${r1}:${cLet(b)}${r2}`); rg.load("values");
    const bold = cols.desc ? sheet.getRange(`${cols.desc}${r1}:${cols.desc}${r2}`).getCellProperties({ format: { font: { bold: true, italic: true } } }) : null;
    return { rg, bold, r1, a };
  }
  function srcRows(h, cols) {
    const v = h.rg.values, out = [];
    const g = (row, L) => (L ? row[cIdx(L) - h.a] : null);
    for (let i = 0; i < v.length; i++) {
      const row = v[i];
      out.push({ r: h.r1 + i, stt: g(row, cols.stt), desc: g(row, cols.desc), spec: g(row, cols.spec), code: g(row, cols.code),
                 brand: g(row, cols.brand), origin: g(row, cols.origin), unit: g(row, cols.unit),
                 qty: (cols.zones || []).map((z) => g(row, z)),
                 bold: !!(h.bold && h.bold.value[i] && h.bold.value[i][0].format.font.bold),
                 italic: !!(h.bold && h.bold.value[i] && h.bold.value[i][0].format.font.italic) });
    }
    return out;
  }
  const cellVal = (v) => (v === null || v === undefined || (typeof v === "string" && v.trim() === "") ? null : (typeof v === "string" ? v.trim() : v));
  const toQty = (v) => { if (typeof v === "number") return v; const t = cellVal(v); if (t === null) return null; const n = Number(String(t).replace(/,/g, "")); return Number.isFinite(n) ? n : null; };
  // Ghép kết quả phân tích (đọc từ _TDI_TMP) với giá trị gốc của sheet CĐT
  function attach(tmpRows, src) {
    const by = new Map(src.map((s) => [s.r, s]));
    return tmpRows.map((t) => { const s = by.get(t.src_row) || {};
      return Object.assign({}, t, { stt: cellVal(s.stt), desc: cellVal(s.desc), spec: cellVal(s.spec), code_cdt: cellVal(s.code),
        brand: cellVal(s.brand), origin: cellVal(s.origin), unit: cellVal(s.unit), qty: (s.qty || []).map(toQty),
        src_bold: !!s.bold, src_italic: !!s.italic }); });
  }

  const priceGroup = (code, unit) => {
    const u = norm(unit).toLowerCase().replace(/\./g, "");
    if (LINEAR.has(u)) return "hao_hut";
    return THIET_BI.has(code) ? "thiet_bi" : "vat_lieu";
  };

  // ----------------------------- layout & công thức -----------------------------
  function col(n) { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function layout(nz) {
    const c = { stt: 1, desc: 2, spec: 3, code: 4, brand: 5, origin: 6, unit: 7 };
    c.zones = Array.from({ length: nz }, (_, i) => 8 + i);
    c.total = 8 + nz;
    const s = c.total + 1;
    Object.assign(c, { r_desc: s, r_code: s + 1, r_brand: s + 2, r_unit: s + 3, p_vl: s + 4, p_nc: s + 5, cl: s + 6,
      inan: s + 7, t_kl: s + 8, dg0: s + 9, dg1: s + 26, kl: s + 27, an: s + 28, ao: s + 29, ap: s + 30 });
    c.amt = nz > 1 ? Array.from({ length: nz }, (_, i) => [s + 31 + 3 * i, s + 32 + 3 * i, s + 33 + 3 * i]) : [];
    const t0 = s + 31 + 3 * c.amt.length;
    c.amt_total = [t0, t0 + 1, t0 + 2];
    c.note = t0 + 3; c.hs_chung = t0 + 2; c.lam_tron = t0 + 3;
    return c;
  }
  // Văn bản bắt đầu bằng = + - @ hoặc trông như số/ngày -> thêm ' để Excel giữ nguyên là chữ
  const asText = (v) => {
    if (v === null || v === undefined) return "";
    if (typeof v !== "string") return v;
    if (/^[=+\-@]/.test(v) || (/^[\d.,\/\-\s]+$/.test(v) && /\d/.test(v))) return "'" + v;
    return v;
  };
  // Một dòng đầu ra: mảng công thức/giá trị cho cột A .. Ghi chú (dùng cho range.formulas)
  function buildRow(x, c) {
    const r = x.out_row, out = new Array(c.note).fill("");
    const put = (k, v) => { out[k - 1] = v === null || v === undefined ? "" : v; };
    const L = (k) => col(k);
    put(c.stt, asText(x.stt)); put(c.desc, asText(x.desc));
    put(c.spec, asText(x.spec)); put(c.code, asText(x.code_cdt)); put(c.brand, asText(x.brand)); put(c.origin, asText(x.origin));
    put(c.unit, asText(x.unit));
    c.zones.forEach((zc, i) => put(zc, typeof x.qty[i] === "number" ? x.qty[i] : ""));
    put(c.r_desc, `=+B${r}`);
    if (x.kind === "section") {
      for (const g of c.amt.concat([c.amt_total])) for (const k of g) put(k, `=+SUBTOTAL(9,${L(k)}${r + 1}:${L(k)}${x.sec_end})`);
      return out;
    }
    const tot = L(c.total), z = c.zones.map(L);
    const P9 = `$${L(c.p_vl)}$9`, Q9 = `$${L(c.p_nc)}$9`, R9 = `$${L(c.cl)}$9`, O9 = `$${L(c.r_unit)}$9`, T9 = `$${L(c.t_kl)}$9`;
    const BB9 = `$${L(c.hs_chung)}$9`, BC9 = `$${L(c.lam_tron)}$9`;
    put(c.total, "=+" + z.map((a) => `${a}${r}`).join("+"));
    put(c.r_unit, `=+G${r}`); put(c.cl, x.cl || ""); put(c.t_kl, `=+${tot}${r}`);
    const O = `${L(c.r_unit)}${r}`, T = `${L(c.t_kl)}${r}`, P = `${L(c.p_vl)}${r}`, Q = `${L(c.p_nc)}${r}`;
    put(c.kl, `=IF(OR(${O}="m",${O}="m2",${O}="m3"),${T}*${T9},${T})`);
    const g = priceGroup(x.cl, x.unit);
    put(c.an, g === "hao_hut" ? `=ROUNDDOWN(${P}*${P9}*${BB9}*${O9},${BC9})`
      : g === "thiet_bi" ? `=ROUNDDOWN(${P}*${R9}*${BB9},${BC9})` : `=ROUNDDOWN(${P}*${P9}*${BB9},${BC9})`);
    const AN = `${L(c.an)}${r}`, AO = `${L(c.ao)}${r}`, AP = `${L(c.ap)}${r}`;
    put(c.ao, `=ROUNDDOWN(${Q}*${Q9}*${BB9},${BC9})`);
    put(c.ap, `=${AN}+${AO}`);
    c.amt.forEach(([a1, a2, a3], i) => { put(a1, `=+${AN}*${z[i]}${r}`); put(a2, `=+${AO}*${z[i]}${r}`); put(a3, `=+${AP}*${z[i]}${r}`); });
    const [t1, t2, t3] = c.amt_total;
    put(t1, `=${AN}*$${tot}${r}`); put(t2, `=${AO}*$${tot}${r}`); put(t3, `=${AP}*$${tot}${r}`);
    return out;
  }
  function buildTotalRow(c, p) {
    const out = new Array(c.note).fill(""), r = p.total_row;
    out[c.desc - 1] = "Tổng cộng trước thuế"; out[c.r_desc - 1] = `=+B${r}`;
    for (const g of c.amt.concat([c.amt_total])) for (const k of g) out[k - 1] = `=SUBTOTAL(9,${col(k)}${p.first}:${col(k)}${p.last})`;
    return out;
  }


  // ======================= CHẾ ĐỘ IN-PLACE: giữ nguyên sheet CĐT, ghi TDI ở bên phải =======================
  // cols: cột của CĐT (giống CFG.cols), total: cột tổng KL của CĐT (hoặc null), start: cột bắt đầu phần TDI
  function layoutInplace(cols, nz, start, total) {
    const base = layout(nz), shift = cIdx(start) - base.r_desc, c = {};
    for (const [k, v] of Object.entries(base)) {
      if (["stt", "desc", "spec", "code", "brand", "origin", "unit", "zones", "total"].includes(k)) continue;
      if (typeof v === "number") c[k] = v + shift;
      else if (k === "amt") c[k] = v.map((g) => g.map((x) => x + shift));
      else if (k === "amt_total") c[k] = v.map((x) => x + shift);
    }
    for (const k of ["stt", "desc", "spec", "code", "brand", "origin", "unit"]) c[k] = cols[k] ? cIdx(cols[k]) : null;
    c.zones = (cols.zones || []).map(cIdx); c.total = total ? cIdx(total) : null;
    return c;
  }
  // Vị trí tiêu đề TDI theo dòng tiêu đề cuối của bảng CĐT (h2): 3 tầng h2-2..h2, hệ số ở 2 dòng trên đó
  function headerPlan(c, h2, sumName) {
    const flat = h2 < 3, tier = { 1: flat ? h2 : h2 - 2, 2: flat ? h2 : h2 - 1, 3: h2 };
    const coefOk = !flat && tier[1] - 2 >= 1, S = "'" + sumName.replace(/'/g, "''") + "'";
    const val = coefOk ? tier[1] - 1 : null;
    const ref = (k) => `$${col(c[k])}$${val}`;
    const refs = coefOk ? { P9: ref("p_vl"), Q9: ref("p_nc"), R9: ref("cl"), O9: ref("r_unit"), T9: ref("t_kl"), BB9: ref("hs_chung"), BC9: ref("lam_tron") }
      : { P9: `${S}!$C$8`, Q9: `${S}!$E$8`, R9: `${S}!$D$8`, O9: `${S}!$H$8`, T9: `${S}!$B$8`, BB9: `${S}!$F$8`, BC9: `${S}!$G$8` };
    return { flat, tier, coefOk, lab: coefOk ? tier[1] - 2 : null, val, refs, S };
  }
  // Một dòng phía TDI (cột Nội dung .. Ghi chú), giống hệt công thức bản Python in-place
  function buildRight(x, c, hp) {
    const r = x.out_row, n = c.note - c.r_desc + 1, out = new Array(n).fill("");
    const put = (k, v) => { out[k - c.r_desc] = v === null || v === undefined ? "" : v; };
    const L = col, R = hp.refs;
    put(c.r_desc, `=+${L(c.desc)}${r}`);
    if (x.kind === "section") {
      for (const g of c.amt.concat([c.amt_total])) for (const k of g) put(k, `=+SUBTOTAL(9,${L(k)}${r + 1}:${L(k)}${x.sec_end})`);
      return out;
    }
    const z = c.zones.map(L), totL = c.total ? L(c.total) : L(c.t_kl);
    put(c.r_unit, c.unit ? `=+${L(c.unit)}${r}` : ""); put(c.cl, x.cl || "");
    put(c.t_kl, c.total ? `=+${L(c.total)}${r}` : "=+" + z.map((a) => `${a}${r}`).join("+"));
    const O = `${L(c.r_unit)}${r}`, T = `${L(c.t_kl)}${r}`, P = `${L(c.p_vl)}${r}`, Q = `${L(c.p_nc)}${r}`;
    put(c.kl, `=IF(OR(${O}="m",${O}="m2",${O}="m3"),${T}*${R.T9},${T})`);
    const g = priceGroup(x.cl, x.unit);
    put(c.an, g === "hao_hut" ? `=ROUNDDOWN(${P}*${R.P9}*${R.BB9}*${R.O9},${R.BC9})`
      : g === "thiet_bi" ? `=ROUNDDOWN(${P}*${R.R9}*${R.BB9},${R.BC9})` : `=ROUNDDOWN(${P}*${R.P9}*${R.BB9},${R.BC9})`);
    const AN = `${L(c.an)}${r}`, AO = `${L(c.ao)}${r}`, AP = `${L(c.ap)}${r}`;
    put(c.ao, `=ROUNDDOWN(${Q}*${R.Q9}*${R.BB9},${R.BC9})`); put(c.ap, `=${AN}+${AO}`);
    c.amt.forEach(([a1, a2, a3], i) => { put(a1, `=+${AN}*${z[i]}${r}`); put(a2, `=+${AO}*${z[i]}${r}`); put(a3, `=+${AP}*${z[i]}${r}`); });
    const [t1, t2, t3] = c.amt_total;
    put(t1, `=${AN}*$${totL}${r}`); put(t2, `=${AO}*$${totL}${r}`); put(t3, `=${AP}*$${totL}${r}`);
    return out;
  }
  // Mảng 2 chiều liền mạch cho dòng r1..r2 (dòng trống của CĐT -> để trống phía TDI)
  function rightBlock(rows, c, hp, r1, r2) {
    const by = new Map(rows.map((x) => [x.out_row, x])), n = c.note - c.r_desc + 1, out = [];
    for (let r = r1; r <= r2; r++) out.push(by.has(r) ? buildRight(by.get(r), c, hp) : new Array(n).fill(""));
    return out;
  }
  function rightTotal(c, first, last, tr) {
    const out = new Array(c.note - c.r_desc + 1).fill("");
    out[0] = "Tổng cộng trước thuế (TDI)";
    for (const g of c.amt.concat([c.amt_total])) for (const k of g) out[k - c.r_desc] = `=SUBTOTAL(9,${col(k)}${first}:${col(k)}${last})`;
    return out;
  }
  // Định dạng phía TDI cho dòng r1..r2 (không đụng cột CĐT)
  function formatRight(sheet, c, rows, r1, r2, isTotal) {
    const a = col(c.r_desc), z = col(c.note), n = r2 - r1 + 1;
    const all = sheet.getRange(`${a}${r1}:${z}${r2}`);
    all.format.font.name = "Times New Roman"; all.format.font.size = 11; all.format.font.color = PAL.navy;
    all.format.verticalAlignment = "Center";
    ["EdgeLeft", "EdgeRight", "InsideVertical"].forEach((e) => { all.format.borders.getItem(e).style = "Continuous"; });
    ["EdgeTop", "EdgeBottom", "InsideHorizontal"].forEach((e) => { all.format.borders.getItem(e).style = "Dot"; });
    const colFmt = (k, f) => { sheet.getRange(`${col(k)}${r1}:${col(k)}${r2}`).numberFormat = Array.from({ length: n }, () => [f]); };
    [c.t_kl, c.kl].forEach((k) => colFmt(k, QTY_FMT));
    [c.p_vl, c.p_nc, c.an, c.ao, c.ap].concat(...c.amt, c.amt_total).forEach((k) => colFmt(k, MONEY_FMT));
    [c.r_desc, c.note].forEach((k) => { const g = sheet.getRange(`${col(k)}${r1}:${col(k)}${r2}`); g.format.wrapText = true; g.format.horizontalAlignment = "Left"; });
    [c.r_code, c.r_brand, c.r_unit, c.cl, c.inan].forEach((k) => { const g = sheet.getRange(`${col(k)}${r1}:${col(k)}${r2}`); g.format.wrapText = true; g.format.horizontalAlignment = "Center"; });
    sheet.getRange(`${col(c.p_vl)}${r1}:${col(c.p_nc)}${r2}`).format.horizontalAlignment = "Right";
    sheet.getRange(`${col(c.t_kl)}${r1}:${col(c.amt_total[2])}${r2}`).format.horizontalAlignment = "Right";
    sheet.getRange(`${col(c.cl)}${r1}:${col(c.cl)}${r2}`).format.font.size = 10;
    // gộp các dòng liền nhau cùng kiểu -> ít lệnh định dạng hơn (nhanh hơn nhiều với BOQ lớn)
    const runs = (pred) => { const out = []; let s = null, p = null;
      for (const x of rows) { if (x.out_row < r1 || x.out_row > r2 || !pred(x)) continue;
        if (s !== null && x.out_row === p + 1) p = x.out_row; else { if (s !== null) out.push([s, p]); s = p = x.out_row; } }
      if (s !== null) out.push([s, p]); return out; };
    for (const [s, e] of runs((x) => x.kind === "section" && x.level <= 1)) { const g = sheet.getRange(`${a}${s}:${z}${e}`); g.format.fill.color = PAL.sec1; g.format.font.bold = true; }
    for (const [s, e] of runs((x) => x.kind === "section" && x.level > 1)) { const g = sheet.getRange(`${a}${s}:${z}${e}`); g.format.fill.color = PAL.sec2; g.format.font.bold = true; }
    for (const [s, e] of runs((x) => x.kind === "item")) sheet.getRange(`${col(c.an)}${s}:${col(c.an)}${e}`).format.fill.color = PAL.an;
    for (const [s, e] of runs((x) => x.src_bold || x.bold)) sheet.getRange(`${a}${s}:${a}${e}`).format.font.bold = true;
    for (const [s, e] of runs((x) => x.src_italic)) sheet.getRange(`${a}${s}:${a}${e}`).format.font.italic = true;
    if (isTotal) { const g = sheet.getRange(`${a}${r2}:${z}${r2}`); g.format.fill.color = PAL.sec1; g.format.font.bold = true; g.format.font.color = PAL.red; }
  }
  // Tiêu đề + hệ số phía TDI (chỉ các cột TDI). zones: tên khu; descWidthPt: độ rộng cột mô tả CĐT (điểm)
  function writeHeaderRight(sheet, c, zones, hp, descWidthPt) {
    const T = hp.tier, a = col(c.r_desc), z = col(c.note);
    if (hp.coefOk) {
      [[c.r_unit, "HS hao hụt", `=${hp.S}!H8`], [c.p_vl, "HSVL", `=${hp.S}!C8`], [c.p_nc, "HSNC", `=${hp.S}!E8`], [c.cl, "HSTB", `=${hp.S}!D8`],
       [c.t_kl, "HSKL", `=${hp.S}!B8`], [c.hs_chung, "HS chung", `=${hp.S}!F8`], [c.lam_tron, "LÀM TRÒN", `=${hp.S}!G8`]].forEach(([k, lab, v]) => {
        const l = sheet.getRange(`${col(k)}${hp.lab}`); l.values = [[lab]]; l.format.font.bold = true; l.format.font.size = 10; l.format.font.color = PAL.blue;
        l.format.font.name = "Times New Roman"; l.format.horizontalAlignment = "Center";
        const q = sheet.getRange(`${col(k)}${hp.val}`); q.formulas = [[v]]; q.format.font.size = 10; q.format.font.color = PAL.navy;
        q.format.font.name = "Times New Roman"; q.format.horizontalAlignment = "Center"; });
    }
    const top = T[1], bot = T[3];
    const H = sheet.getRange(`${a}${top}:${z}${bot}`);
    H.format.fill.color = PAL.head; H.format.font.name = "Times New Roman"; H.format.font.size = 11; H.format.font.bold = true;
    H.format.font.color = PAL.red; H.format.wrapText = true; H.format.horizontalAlignment = "Center"; H.format.verticalAlignment = "Center";
    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"].forEach((e) => { H.format.borders.getItem(e).style = "Continuous"; });
    const m = (t1, c1, t2, c2, text) => {
      if (hp.flat && t2 < 3) return;
      const rg = sheet.getRange(`${col(c1)}${T[t1]}:${col(c2)}${T[t2]}`);
      if (T[t1] !== T[t2] || c1 !== c2) rg.merge(false);
      rg.getCell(0, 0).values = [[text]];
    };
    [["r_desc", "Nội Dung Công Việc"], ["r_code", "Quy cách\nMã Hiệu"], ["r_brand", "Nhãn Hiệu\nXuất Xứ"], ["r_unit", "Đơn Vị"],
     ["t_kl", "Tổng KL tạm tính"], ["kl", "KL"], ["note", "Ghi Chú"]].forEach(([k, t]) => m(1, c[k], 3, c[k], t));
    m(1, c.p_vl, 1, c.p_nc, "Đơn giá gốc"); m(2, c.p_vl, 3, c.p_vl, "Vật liệu"); m(2, c.p_nc, 3, c.p_nc, "Nhân Công");
    sheet.getRange(`${col(c.p_vl)}${top}:${col(c.p_nc)}${bot}`).format.fill.color = PAL.yellow;
    sheet.getRange(`${col(c.t_kl)}${top}:${col(c.t_kl)}${bot}`).format.fill.color = PAL.yellow;
    m(1, c.cl, 1, c.inan, "Bộ lọc"); m(2, c.cl, 3, c.cl, "Chủng loại"); m(2, c.inan, 3, c.inan, "In ấn");
    sheet.getRange(`${col(c.cl)}${top}:${col(c.inan)}${bot}`).format.fill.color = PAL.green;
    m(1, c.dg0, 1, c.dg1, "Diễn giải chi tiết KL");
    let k = c.dg0;
    for (const [g, subs] of DIENGIAI) { m(2, k, 2, k + subs.length - 1, g); subs.forEach((s2, j) => { sheet.getRange(`${col(k + j)}${bot}`).values = [[s2]]; }); k += subs.length; }
    if (!hp.flat) sheet.getRange(`${col(c.dg0)}${top}:${col(c.dg1)}${T[2]}`).format.fill.color = PAL.dg;
    sheet.getRange(`${col(c.dg0)}${bot}:${col(c.dg1)}${bot}`).format.fill.color = PAL.yellow;
    m(1, c.an, 1, c.ap, "Đơn giá"); [["an", "Vật Liệu"], ["ao", "Nhân công"], ["ap", "ĐG Tổng"]].forEach(([kk, t]) => m(2, c[kk], 3, c[kk], t));
    const groups = c.amt.map((g, i) => [zones[i], g]).concat([["Total", c.amt_total]]);
    for (const [name, [a1, a2, a3]] of groups) {
      m(1, a1, 1, a3, name); m(2, a1, 2, a2, "Amount/ thành tiền"); m(2, a3, 3, a3, "Total Amount/ Tổng thành tiền");
      sheet.getRange(`${col(a1)}${bot}`).values = [[hp.flat ? "Cung cấp - " + name : "Supply/ Cung cấp"]];
      sheet.getRange(`${col(a2)}${bot}`).values = [[hp.flat ? "Lắp đặt - " + name : "Installation/ lắp đặt"]];
    }
    // độ rộng: cột Nội dung = cột mô tả CĐT để xuống dòng giống hệt; ẩn cột diễn giải
    const w = (kk, ch) => { sheet.getRange(`${col(kk)}:${col(kk)}`).format.columnWidth = Math.round((ch * 7 + 5) * 0.75); };
    sheet.getRange(`${a}:${a}`).format.columnWidth = descWidthPt || Math.round((46 * 7 + 5) * 0.75);
    w(c.r_code, 18); w(c.r_brand, 18); w(c.r_unit, 9); w(c.p_vl, 15); w(c.p_nc, 15); w(c.cl, 9.5); w(c.inan, 8); w(c.t_kl, 13); w(c.kl, 13);
    w(c.an, 15); w(c.ao, 15); w(c.ap, 16); c.amt.forEach((g) => g.forEach((kk) => w(kk, 17))); c.amt_total.forEach((kk) => w(kk, 19)); w(c.note, 24);
    sheet.getRange(`${col(c.dg0)}:${col(c.dg1)}`).columnHidden = true;
  }
  const planInplace = (rows) => {   // dùng khi đọc lại từ _TDI_TMP: out_row = dòng CĐT
    rows.forEach((x) => { x.out_row = x.src_row; });
    return { first: rows[0].src_row, last: rows[rows.length - 1].src_row, total_row: rows[rows.length - 1].src_row + 1 };
  };

  // ----------------------------- Office.js: đầu sheet -----------------------------
  const PAL = { head: "#9DC3E6", yellow: "#FFFF00", green: "#92D050", dg: "#FBE5D6", sec1: "#BDD7EE", sec2: "#D9D9D9",
                an: "#DEEBF7", navy: "#002060", red: "#C00000", blue: "#0000FF", orange: "#DE5A00" };
  const DIENGIAI = [["Trục chính", ["Trục 1", "Trục 2"]], ["Tầng Hầm", ["Hầm 1", "Hầm 2", "Hầm 3"]],
    ["Khối đế phía trên", ["Tầng 1", "Tầng 2", "Tầng 3"]], ["Hành lang điển hình", ["HL1", "HL2", "HL3"]],
    ["Căn hộ điển hình", ["CH-01", "CH-02", "CH-03", "CH-04", "CH-05"]], ["Khu vực khác", ["Khác 1", "Khác 2"]]];
  const QTY_FMT = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-';
  const MONEY_FMT = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-';

  // Ghi dòng 1-12 (giá trị + merge + màu). sheet: Excel.Worksheet; zones: ["Basement", ...]
  function writeHeader(sheet, c, zones, title, opt) {
    opt = opt || {};
    const last = col(c.note);
    const set = (a, v) => { sheet.getRange(a).values = [[v]]; };
    const m = (a, v) => { const rg = sheet.getRange(a); rg.merge(false); if (v !== undefined) rg.getCell(0, 0).values = [[v]]; return rg; };
    const all = sheet.getRange(`A1:${last}12`);
    all.format.font.name = "Times New Roman"; all.format.font.size = 11;
    [["CÔNG TY CỔ PHẦN TẬP ĐOÀN TDI", 16, PAL.orange], ["TDI GROUP JOINT STOCK COMPANY", 12, PAL.orange],
     ["A16-NV6, ô số 10, KĐT mới đường Lê Trọng Tấn, An Khánh, Hoài Đức, HN", 9, PAL.navy],
     ["Tel: (024) 7308.6632 - info@tdigroup.vn - www.tdigroup.vn", 9, PAL.navy]].forEach(([t, sz, cl], i) => {
      const rg = sheet.getRange(`${last}${i + 1}`); rg.values = [[t]]; rg.format.font.size = sz; rg.format.font.bold = true;
      rg.format.font.color = cl; rg.format.horizontalAlignment = "Right"; });
    const t5 = m(`A5:${last}5`, "CHI TIẾT GIÁ");
    t5.format.font.size = 20; t5.format.font.bold = true; t5.format.font.color = PAL.navy; t5.format.horizontalAlignment = "Center";
    t5.format.borders.getItem("EdgeTop").style = "Double";
    set("A6", "=Summary!A32"); set("A7", "Hạng mục: " + title);
    sheet.getRange("A6").format.font.color = PAL.navy; sheet.getRange("A7").format.font.color = "#FF0000";
    sheet.getRange("A6:A7").format.font.bold = true; sheet.getRange("A6:A7").format.font.size = 12;
    set(`${col(c.hs_chung)}6`, "=Summary!G32"); set(`${col(c.lam_tron)}6`, "=Summary!H32");
    set(`${col(c.hs_chung)}7`, "=Summary!G33"); set(`${col(c.lam_tron)}7`, "=Summary!H33");
    sheet.getRange(`${col(c.lam_tron)}7`).numberFormat = [["dd/mm/yyyy"]];
    sheet.getRange(`${col(c.hs_chung)}6:${col(c.lam_tron)}7`).format.font.bold = true;
    [[c.r_unit, "HS hao hụt", opt.hao_hut || 1.02], [c.p_vl, "HSVL", "=Summary!C8"], [c.p_nc, "HSNC", "=Summary!E8"],
     [c.cl, "HSTB", "=Summary!D8"], [c.t_kl, "HSKL", 1], [c.hs_chung, "HS chung", "=Summary!F8"],
     [c.lam_tron, "LÀM TRÒN", "=Summary!G8"]].forEach(([k, lab, v]) => {
      const a = sheet.getRange(`${col(k)}8`); a.values = [[lab]]; a.format.font.bold = true; a.format.font.size = 10;
      a.format.font.color = PAL.blue; a.format.horizontalAlignment = "Center";
      const b = sheet.getRange(`${col(k)}9`); b.formulas = [[v]]; b.format.font.size = 10; b.format.font.color = PAL.navy;
      b.format.horizontalAlignment = "Center"; });
    // Tiêu đề bảng 10-12
    const H = sheet.getRange(`A10:${last}12`);
    H.format.fill.color = PAL.head; H.format.font.bold = true; H.format.wrapText = true;
    H.format.horizontalAlignment = "Center"; H.format.verticalAlignment = "Center";
    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"].forEach((e) => {
      H.format.borders.getItem(e).style = "Continuous"; });
    sheet.getRange(`${col(c.r_desc)}10:${last}12`).format.font.color = PAL.red;
    const v3 = (k, t) => m(`${col(k)}10:${col(k)}12`, t);
    [["stt", "STT"], ["desc", "Hạng mục công việc"], ["spec", "Thông số kỹ thuật"], ["code", "Mã hiệu"], ["brand", "Nhãn hiệu"],
     ["origin", "Xuất xứ"], ["unit", "Đơn vị"], ["r_desc", "Nội Dung Công Việc"], ["r_code", "Quy cách\nMã Hiệu"],
     ["r_brand", "Nhãn Hiệu\nXuất Xứ"], ["r_unit", "Đơn Vị"], ["t_kl", "Tổng KL tạm tính"], ["kl", "KL"], ["note", "Ghi Chú"]]
      .forEach(([k, t]) => v3(c[k], t));
    m(`${col(c.zones[0])}10:${col(c.total)}10`, "Quantity/ Khối lượng");
    zones.forEach((zn, i) => m(`${col(c.zones[i])}11:${col(c.zones[i])}12`, zn));
    m(`${col(c.total)}11:${col(c.total)}12`, "Tổng khối lượng");
    m(`${col(c.p_vl)}10:${col(c.p_nc)}10`, "Đơn giá gốc");
    m(`${col(c.p_vl)}11:${col(c.p_vl)}12`, "Vật liệu"); m(`${col(c.p_nc)}11:${col(c.p_nc)}12`, "Nhân Công");
    sheet.getRange(`${col(c.p_vl)}10:${col(c.p_nc)}12`).format.fill.color = PAL.yellow;
    sheet.getRange(`${col(c.t_kl)}10:${col(c.t_kl)}12`).format.fill.color = PAL.yellow;
    m(`${col(c.cl)}10:${col(c.inan)}10`, "Bộ lọc");
    m(`${col(c.cl)}11:${col(c.cl)}12`, "Chủng loại"); m(`${col(c.inan)}11:${col(c.inan)}12`, "In ấn");
    sheet.getRange(`${col(c.cl)}10:${col(c.inan)}12`).format.fill.color = PAL.green;
    m(`${col(c.dg0)}10:${col(c.dg1)}10`, "Diễn giải chi tiết KL");
    let k = c.dg0;
    for (const [g, subs] of DIENGIAI) {
      m(`${col(k)}11:${col(k + subs.length - 1)}11`, g);
      subs.forEach((s, j) => { sheet.getRange(`${col(k + j)}12`).values = [[s]]; });
      k += subs.length;
    }
    sheet.getRange(`${col(c.dg0)}10:${col(c.dg1)}11`).format.fill.color = PAL.dg;
    sheet.getRange(`${col(c.dg0)}12:${col(c.dg1)}12`).format.fill.color = PAL.yellow;
    m(`${col(c.an)}10:${col(c.ap)}10`, "Đơn giá");
    [["an", "Vật Liệu"], ["ao", "Nhân công"], ["ap", "ĐG Tổng"]].forEach(([kk, t]) => m(`${col(c[kk])}11:${col(c[kk])}12`, t));
    const groups = c.amt.map((g, i) => [zones[i], g]).concat([["Total", c.amt_total]]);
    for (const [name, [a1, a2, a3]] of groups) {
      m(`${col(a1)}10:${col(a3)}10`, name);
      m(`${col(a1)}11:${col(a2)}11`, "Amount/ thành tiền");
      m(`${col(a3)}11:${col(a3)}12`, "Total Amount/ Tổng thành tiền");
      sheet.getRange(`${col(a1)}12`).values = [["Supply/ Cung cấp"]];
      sheet.getRange(`${col(a2)}12`).values = [["Installation/ lắp đặt"]];
    }
    sheet.getRange("10:12").format.rowHeight = 33;
  }

  // ----------------------------- Office.js: định dạng vùng dữ liệu -----------------------------
  // Gọi cho từng lô (r1..r2), sau khi đã ghi dữ liệu. rows: các dòng của lô (đã có out_row, kind, level, bold)
  function formatBatch(sheet, c, rows, r1, r2, isTotalBatch) {
    const last = col(c.note), n = r2 - r1 + 1;
    const all = sheet.getRange(`A${r1}:${last}${r2}`);
    all.format.font.name = "Times New Roman"; all.format.font.size = 11; all.format.verticalAlignment = "Center";
    ["EdgeLeft", "EdgeRight", "InsideVertical"].forEach((e) => { all.format.borders.getItem(e).style = "Continuous"; });
    ["EdgeTop", "EdgeBottom", "InsideHorizontal"].forEach((e) => { all.format.borders.getItem(e).style = "Dot"; });
    sheet.getRange(`${col(c.r_desc)}${r1}:${last}${r2}`).format.font.color = PAL.navy;
    const colFmt = (k, f) => { sheet.getRange(`${col(k)}${r1}:${col(k)}${r2}`).numberFormat = Array.from({ length: n }, () => [f]); };
    c.zones.concat([c.total, c.t_kl, c.kl]).forEach((k) => colFmt(k, QTY_FMT));
    [c.p_vl, c.p_nc, c.an, c.ao, c.ap].concat(...c.amt, c.amt_total).forEach((k) => colFmt(k, MONEY_FMT));
    const wrapLeft = [c.desc, c.spec, c.r_desc, c.note], wrapCenter = [c.stt, c.code, c.brand, c.origin, c.unit, c.r_code, c.r_brand, c.r_unit, c.cl, c.inan];
    wrapLeft.forEach((k) => { const g = sheet.getRange(`${col(k)}${r1}:${col(k)}${r2}`); g.format.wrapText = true; g.format.horizontalAlignment = "Left"; });
    wrapCenter.forEach((k) => { const g = sheet.getRange(`${col(k)}${r1}:${col(k)}${r2}`); g.format.wrapText = true; g.format.horizontalAlignment = "Center"; });
    sheet.getRange(`${col(c.zones[0])}${r1}:${col(c.total)}${r2}`).format.horizontalAlignment = "Right";
    sheet.getRange(`${col(c.p_vl)}${r1}:${col(c.p_nc)}${r2}`).format.horizontalAlignment = "Right";
    sheet.getRange(`${col(c.t_kl)}${r1}:${last}${r2}`).format.horizontalAlignment = "Right";
    sheet.getRange(`${col(c.note)}${r1}:${col(c.note)}${r2}`).format.horizontalAlignment = "Left";
    sheet.getRange(`${col(c.cl)}${r1}:${col(c.cl)}${r2}`).format.font.size = 10;
    for (const x of rows) {
      const rr = x.out_row;
      if (x.kind === "section") {
        const g = sheet.getRange(`A${rr}:${last}${rr}`);
        g.format.fill.color = x.level <= 1 ? PAL.sec1 : PAL.sec2; g.format.font.bold = true;
      } else {
        if (x.bold) sheet.getRange(`A${rr}:${col(c.unit)}${rr}`).format.font.bold = true;
        if (x.kind === "item") sheet.getRange(`${col(c.an)}${rr}`).format.fill.color = PAL.an;
      }
    }
    if (isTotalBatch) {
      const tr = r2, g = sheet.getRange(`A${tr}:${last}${tr}`);
      g.format.fill.color = PAL.sec1; g.format.font.bold = true; g.format.font.color = PAL.red;
    }
    all.format.autofitRows();
  }

  // Độ rộng cột (đơn vị ký tự như Excel), ẩn cột diễn giải, cố định khung, lọc
  function finishSheet(sheet, c, p, rows) {
    const w = (k, chars) => { sheet.getRange(`${col(k)}:${col(k)}`).format.columnWidth = Math.round((chars * 7 + 5) * 0.75); };
    const maxLen = (key) => Math.max(0, ...rows.map((x) => String(x[key] === null || x[key] === undefined ? "" : x[key]).length));
    w(c.stt, Math.min(Math.max(5.5, maxLen("stt") * 1.1 + 2), 9)); w(c.desc, 46); w(c.r_desc, 46);
    w(c.spec, maxLen("spec") ? Math.min(Math.max(16, maxLen("spec") * 0.5), 32) : 11);
    [["code_cdt", c.code, 12, 26], ["brand", c.brand, 12, 22], ["origin", c.origin, 12, 20]].forEach(([key, k, lo, hi]) => {
      const n = maxLen(key); w(k, n ? Math.min(Math.max(lo, n * 0.8 + 2), hi) : 11); });
    const uw = Math.min(Math.max(8, maxLen("unit") + 3), 12); w(c.unit, uw); w(c.r_unit, uw);
    c.zones.forEach((k) => w(k, 14)); w(c.total, 15);
    w(c.r_code, 18); w(c.r_brand, 18); w(c.p_vl, 15); w(c.p_nc, 15); w(c.cl, 9.5); w(c.inan, 8); w(c.t_kl, 13); w(c.kl, 13);
    w(c.an, 15); w(c.ao, 15); w(c.ap, 16);
    c.amt.forEach((g) => g.forEach((k) => w(k, 17))); c.amt_total.forEach((k) => w(k, 19)); w(c.note, 24);
    sheet.getRange(`${col(c.dg0)}:${col(c.dg1)}`).columnHidden = true;
    sheet.freezePanes.freezeAt(sheet.getRange("A1:B12"));
    sheet.autoFilter.apply(sheet.getRange(`A12:${col(c.note)}${p.last}`));
  }

  // ----------------------------- Office.js: Summary -----------------------------
  // infos: [{sheet, name, total_row, vl, nc, tong}] ; cfg: {project, hang_muc, so_bao_gia, ngay, coefficients}
  function writeSummary(sm, infos, cfg) {
    cfg = cfg || {}; const co = cfg.coefficients || {};
    const S = (a, v) => { sm.getRange(a).formulas = [[v]]; };
    const all = sm.getRange("A1:H62"); all.format.font.name = "Times New Roman"; all.format.font.size = 11;
    S("H1", "CÔNG TY CỔ PHẦN TẬP ĐOÀN TDI"); S("H2", "TDI GROUP JOINT STOCK COMPANY");
    sm.getRange("H1:H2").format.font.bold = true; sm.getRange("H1:H2").format.font.color = PAL.orange;
    sm.getRange("H1:H4").format.horizontalAlignment = "Right";
    const t = sm.getRange("A6:H6"); t.merge(false); t.getCell(0, 0).values = [["BẢNG TỔNG HỢP DỰ TOÁN"]];
    t.format.font.size = 24; t.format.font.bold = true; t.format.font.color = PAL.navy; t.format.horizontalAlignment = "Center";
    [["B", "HSKL", co.HSKL ?? 1], ["C", "HSVL", co.HSVL ?? 1.25], ["D", "HSTB", co.HSTB ?? 1.1], ["E", "HSNC", co.HSNC ?? 1.15],
     ["F", "HS chung", co.HS_chung ?? 1], ["G", "LÀM TRÒN", co.lam_tron ?? -2]].forEach(([cc, lab, v]) => { S(cc + "7", lab); S(cc + "8", v); });
    S("H7", "HS hao hụt"); S("H8", co.hao_hut ?? 1.02);
    sm.getRange("B7:H7").format.font.bold = true; sm.getRange("B7:H7").format.font.color = PAL.blue;
    sm.getRange("B8:H8").format.fill.color = "#FFFF00"; sm.getRange("B7:H8").format.horizontalAlignment = "Center";
    S("A32", "Dự án: " + (cfg.project || "")); S("A33", "Hạng mục: " + (cfg.hang_muc || "Cung cấp & lắp đặt hệ thống cơ điện"));
    S("G32", "Số:"); S("H32", cfg.so_bao_gia || ""); S("G33", "Ngày:"); S("H33", cfg.ngay || "=TODAY()");
    sm.getRange("H33").numberFormat = [["dd/mm/yyyy"]];
    sm.getRange("A32:H33").format.font.bold = true; sm.getRange("A32:H33").format.font.color = PAL.navy;
    const head = [["A34:A35", "STT"], ["B34:B35", "Hạng Mục Công Việc"], ["C34:C35", "Đơn Vị"], ["D34:D35", "KL"],
                  ["E34:F34", "Đơn Giá"], ["G34:G35", "Thành Tiền"], ["H34:H35", "Ghi Chú"]];
    head.forEach(([a, v]) => { const g = sm.getRange(a); g.merge(false); g.getCell(0, 0).values = [[v]]; });
    S("E35", "Vật Tư"); S("F35", "Nhân Công");
    const H = sm.getRange("A34:H35"); H.format.fill.color = PAL.head; H.format.font.bold = true; H.format.font.color = PAL.red;
    H.format.horizontalAlignment = "Center"; H.format.verticalAlignment = "Center"; H.format.wrapText = true;
    S("B38", "CHI PHÍ TRỰC TIẾP"); S("G38", "=SUM(G39:G50)");
    infos.slice(0, 12).forEach((inf, i) => {
      const r = 39 + i, q = "'" + inf.sheet.replace(/'/g, "''") + "'";
      sm.getRange(`A${r}:G${r}`).formulas = [[String.fromCharCode(65 + i), inf.name, "Hệ", 1,
        `=+${q}!${inf.vl}${inf.total_row}`, `=+${q}!${inf.nc}${inf.total_row}`, `=+${q}!${inf.tong}${inf.total_row}`]];
      sm.getRange(`A${r}:H${r}`).format.fill.color = PAL.sec2; sm.getRange(`A${r}:H${r}`).format.font.bold = true;
    });
    S("B51", "CHI PHÍ GIÁN TIẾP"); S("G51", "=SUBTOTAL(9,G52:G56)");
    [["Chi Phí Văn Phòng, Kho Bãi,…", 0.01], ["Chi phí Quản Lý Gián Tiếp", 0.035],
     ["Chi phí tiện ích trong thi công (Điện, Nước, Vận thang, Vận chuyển rác thải ra trong phạm vi công trường.)", 0.02],
     ["Bảo Hộ, Dụng Cụ An Toàn, Bảo hiểm cho người lao động", 0.01],
     ["Vận chuyển, bảo quản và bảo vệ thiết bị do chủ đầu tư cấp (bồi thường nếu mất mát, hỏng hóc)", 0.035]].forEach(([b, d], i) => {
      const r = 52 + i; sm.getRange(`B${r}:D${r}`).values = [[b, "Lot", d]];
      if (i < 4) S(`G${r}`, `=$G$38*D${r}`); });
    S("H56", "Nếu chủ đầu tư cấp thiết bị");
    sm.getRange("D52:D56").numberFormat = Array.from({ length: 5 }, () => ["0.00%"]);
    S("B57", "TỔNG (Chưa bao gồm VAT)"); S("G57", "=G38+G51");
    S("B58", "Thuế VAT 10%"); S("G58", "=G57*0.1");
    S("B59", "TỔNG (Đã bao gồm VAT)"); S("G59", "=SUM(G57:G58)");
    [38, 51, 57, 59].forEach((r) => { const g = sm.getRange(`A${r}:H${r}`); g.format.font.bold = true; g.format.fill.color = PAL.dg; });
    S("G62", "CÔNG TY CP TẬP ĐOÀN TDI"); sm.getRange("G62").format.font.bold = true;
    const body = sm.getRange("A34:H59");
    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical"].forEach((e) => { body.format.borders.getItem(e).style = "Continuous"; });
    body.format.borders.getItem("InsideHorizontal").style = "Dot";
    sm.getRange("E38:G59").numberFormat = Array.from({ length: 22 }, () => Array(3).fill(MONEY_FMT));
    sm.getRange("D52:D56").numberFormat = Array.from({ length: 5 }, () => ["0.00%"]);
    sm.getRange("B38:B59").format.wrapText = true;
    [["A", 6], ["B", 46], ["C", 9], ["D", 9], ["E", 19], ["F", 19], ["G", 20], ["H", 22]].forEach(([cc, ch]) => {
      sm.getRange(`${cc}:${cc}`).format.columnWidth = Math.round((ch * 7 + 5) * 0.75); });
    sm.getRange("A34:H59").format.autofitRows();
  }


  // ================= GHI MỘT SHEET (in-place): không đọc lại sheet CĐT, mọi thứ cần đã nằm trong sheet tạm =================
  // CFG: { src, cols, header_row, total, zone_names, name, OV, info:{...}, tdi_start? }
  async function run(context, CFG, deadline, reserve) {
    deadline = deadline || Date.now() + (CFG.budget || 45000);
    reserve = reserve === undefined ? 6000 : reserve;
    const T0 = Date.now(), sec = () => Math.round((Date.now() - T0) / 1000), left = () => deadline - Date.now();
    let SIZE = 500, lastDur = 0;
    const wb = context.workbook, ws = wb.worksheets.getItem(CFG.src);
    const tmp = wb.worksheets.getItemOrNullObject("_TDI_TMP_" + String(CFG.src).slice(0, 20)); await context.sync();
    if (tmp.isNullObject) return { loi: "Chưa có kết quả phân tích: chạy script phân tích (TDI_A) trước" };
    const meta = tmp.getRange("P1:T1"); meta.load("values"); await context.sync();
    const n = meta.values[0][2], info = meta.values[0][3] ? JSON.parse(meta.values[0][3]) : null;
    if (!info) return { loi: "Phân tích chưa hoàn tất: chạy lại script phân tích cho tới khi xong: true" };
    const prog = meta.values[0][4] ? JSON.parse(meta.values[0][4]) : { header: false, from: 0 };
    const c = layoutInplace(CFG.cols, CFG.cols.zones.length, CFG.tdi_start || info.tdi_start, CFG.total || null);
    const hp = headerPlan(c, CFG.header_row, info.summary);
    const save = () => { tmp.getRange("T1").values = [[JSON.stringify(prog)]]; };
    if (!prog.header) {
      const sm = wb.worksheets.getItemOrNullObject(info.summary); await context.sync();
      if (sm.isNullObject) wb.worksheets.add(info.summary);
      const d = ws.getRange(`${CFG.cols.desc}:${CFG.cols.desc}`); d.load("format/columnWidth"); await context.sync();
      writeHeaderRight(ws, c, CFG.zone_names || CFG.cols.zones, hp, d.format.columnWidth);
      prog.header = true; save(); await context.sync();
    }
    while (prog.from < n && left() > lastDur * 1.5 + 2000) {
      const t1 = Date.now();
      const t = tmp.getRange(`A${prog.from + 2}:M${Math.min(prog.from + SIZE, n) + 1}`); t.load("values"); await context.sync();
      const rows = applyOV(fromTmp(t.values), CFG.OV);
      if (!rows.length) break;
      const r1 = rows[0].out_row, r2 = rows[rows.length - 1].out_row;
      context.application.suspendApiCalculationUntilNextSync();
      ws.getRange(`${col(c.r_desc)}${r1}:${col(c.note)}${r2}`).formulas = rightBlock(rows, c, hp, r1, r2);
      formatRight(ws, c, rows, r1, r2, false);
      prog.from += rows.length; save();
      await context.sync();
      lastDur = Date.now() - t1;
      if (lastDur > 12000) SIZE = Math.max(25, Math.floor(SIZE / 2));
    }
    if (prog.from < n || left() < reserve) return { xong: false, da_ghi: prog.from, tren: n, giay: sec() };
    const tr = info.total_row, tdi = `${col(c.r_desc)}:${col(c.note)}`;
    ws.getRange(`${col(c.r_desc)}${tr}:${col(c.note)}${tr}`).formulas = [rightTotal(c, info.first, info.last, tr)];
    formatRight(ws, c, [], tr, tr, true);
    const sm = wb.worksheets.getItem(info.summary), z = sm.getRange("Z1"); z.load("values"); await context.sync();
    let infos = []; try { infos = JSON.parse(z.values[0][0] || "[]") || []; } catch (e) { infos = []; }
    infos = infos.filter((x) => x.sheet !== CFG.src).concat([{ sheet: CFG.src, name: CFG.name || CFG.src, total_row: tr,
      vl: col(c.amt_total[0]), nc: col(c.amt_total[1]), tong: col(c.amt_total[2]), tdi }]);
    writeSummary(sm, infos, CFG.info || {});
    sm.getRange("Z1").values = [[JSON.stringify(infos)]]; sm.getRange("Z:Z").columnHidden = true;
    tmp.delete();
    await context.sync();
    return { xong: true, phan_tdi: tdi, dong_tong_tdi: tr, summary: info.summary, giay: sec() };
  }

  // ================= CẢ FILE: ghi mọi sheet đã phân tích (gọi lại y nguyên cho tới khi xong) =================
  // OPT: { OV?: { "<sheet>": { "<dòng>": "MÃ", "ctx:<khóa nhóm>": "MÃ" } }, info?: { project, hang_muc, so_bao_gia, ngay, coefficients }, budget? }
  async function runAll(context, OPT) {
    OPT = OPT || {};
    const BUD = OPT.budget || 60000, T0 = Date.now(), deadline = T0 + BUD, sec = () => Math.round((Date.now() - T0) / 1000);
    const minLeft = Math.min(10000, BUD * 0.15), reserve = Math.min(6000, BUD * 0.1);
    const j = context.workbook.worksheets.getItemOrNullObject("_TDI_JOB"); await context.sync();
    if (j.isNullObject) return { loi: "Chưa có kết quả phân tích: chạy script A (TDI_A.runAll) trước" };
    const r0 = j.getRange("A1"); r0.load("values"); await context.sync();
    const data = JSON.parse(r0.values[0][0] || "null");
    if (!data) return { loi: "Thiếu dữ liệu phân tích: chạy lại script A" };
    if (data.sheets.some((s) => !s.a)) return { loi: "Phân tích chưa xong: chạy lại script A cho tới khi xong: true" };
    const info = Object.assign({ project: data.project || "" }, OPT.info || {});
    for (const s of data.sheets) {
      if (s.a !== 1 || s.b) continue;
      if (deadline - Date.now() < minLeft) break;
      const r = await run(context, Object.assign({}, s.cfg, { OV: (OPT.OV || {})[s.src] || {}, info }), deadline, reserve);
      if (r.loi) { s.b = 2; s.loi_b = r.loi; }
      else if (r.xong) { s.b = 1; s.kq_b = { phan_tdi: r.phan_tdi, dong_tong_tdi: r.dong_tong_tdi, summary: r.summary }; }
      j.getRange("A1").values = [[JSON.stringify(data)]]; await context.sync();
      if (!r.xong && !r.loi) break;
    }
    const pending = data.sheets.filter((s) => s.a === 1 && !s.b);
    if (pending.length) return { buoc: "ghi", xong: false, con_lai: pending.map((s) => s.src), giay: sec(), viec_tiep: "Chạy lại y nguyên script này" };
    j.delete(); await context.sync();
    return { buoc: "ghi", xong: true, giay: sec(), sheets: data.sheets.map((s) => ({ sheet: s.src, ...(s.kq_b || {}), loi: s.loi || s.loi_b })) };
  }

  // Bảng tạm _TDI_TMP: lưu kết quả phân tích để các script sau dùng lại (mỗi script không nhớ gì)
  const TMP_HEAD = ["src_row", "kind", "level", "cl", "cl_source", "bold", "out_row", "sec_end", "desc", "context0", "context1", "unit", "italic"];
  // chữ bắt đầu bằng = + - @ sẽ bị Excel hiểu là công thức -> thêm dấu '
  const safe = (v) => { const t = String(v === null || v === undefined ? "" : v).slice(0, 120); return /^[=+\-@]/.test(t) ? "'" + t : t; };
  const toTmp = (rows) => rows.map((x) => [x.src_row, x.kind, x.level === null ? "" : x.level, x.cl || "", x.cl_source || "",
    x.bold ? 1 : 0, x.out_row || "", x.sec_end || "", safe(x.desc), safe((x.context || [])[0])]);
  const fromTmp = (vals) => vals.filter((v) => v[0] !== "").map((v) => ({ src_row: v[0], kind: v[1], level: v[2] === "" ? null : v[2],
    cl: v[3] || null, cl_source: v[4], bold: v[5] === 1, out_row: v[6], sec_end: v[7] || undefined,
    context: [v[9], v[10]].filter((t) => t !== "" && t !== null && t !== undefined),
    unit: v[11] === "" ? null : v[11], src_bold: v[5] === 1, src_italic: v[12] === 1 }));
  // Áp chỗ sửa Chủng loại ngay lúc ghi: { "<dòng CĐT>": "MÃ", "ctx:<khóa nhóm>": "MÃ" } (khóa nhóm như bước phân tích trả về)
  function applyOV(rows, OV) {
    const heads = {}; for (const k of Object.keys(OV || {})) if (k.startsWith("ctx:")) heads[k.slice(4).trim()] = OV[k];
    for (const x of rows) {   // giống bản Python: sửa theo số dòng áp cho mọi dòng, sửa theo nhóm chỉ áp cho hạng mục
      const own = OV[String(x.src_row)];
      if (own) { x.cl = own; continue; }
      if (x.kind !== "item") continue;
      const h = heads[(x.context || []).slice(0, 2).map((t) => String(t).trim().slice(0, 80)).join(" < ")];
      if (h) x.cl = h;
    }
    return rows;
  }

  return { runAll, run, applyOV, layoutInplace, headerPlan, buildRight, rightBlock, rightTotal, formatRight, writeHeaderRight, planInplace,
           loadSource, srcRows, attach, layout, col, buildRow, buildTotalRow, writeHeader, formatBatch, finishSheet, writeSummary, fromTmp, priceGroup, PAL };
})();
// ============================== HẾT PHẦN 2 ==============================
