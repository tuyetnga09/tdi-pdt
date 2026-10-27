// ============================================================================
// BOQ CĐT -> TDI | PHẦN 1: PHÂN TÍCH (dùng trong add-in Claude for Excel)
// JavaScript thuần: KHÔNG eval / new Function / import / require.
// Dán NGUYÊN khối này vào đầu script phân tích. Sinh tự động bởi scripts/build_excel_core.py.
// ============================================================================
const TDI_A = (() => {
  const D = {"RULES": [["CTK", "chi phi|\\bphi\\b(?! gia công)|lê phi"], ["CPK", "thi nghiêm|thư nghiêm|xet nghiêm|^\\s*kiêm đinh|hiêu chinh rơ ?le"], ["CTK", "nghiêm thu|vân hanh thư|chay thư|thư ap|thư kin|thư tai|testing|commissioning|đao tao|hô sơ hoan công|thu tuc|bao tri|bao dương|^\\s*vân chuyên|nhân công|gian giao|xe nâng|sleeve|ông lông|băng (?:canh )?bao|môc (?:sư )?bao|sư (?:đanh dâu|bao)|đao đăp|cat (?:lot|đen|vang|san)|^\\s*(?:bê|bê tông|xây|trat|gia cô|hoan thiên|đao|đăp|lâp|khoan|đuc|căt)\\b"], ["TS", "thông\\s+tăc\\s+san|floor\\s*clean.?out"], ["MBA", "^(?!\\s*tu\\b).*?(?:may biên ap|tram biên ap)|transformer"], ["RMU", "\\brmu\\b|ring main unit|switchgear|tu trung thê|mvsg|recloser|\\blbs\\b|câu dao phu tai|may căt (?:chân không|trung thê)|\\bvcb\\b|\\bfco\\b|câu chi tư rơi|chông set van|canh bao sư cô cap"], ["CAPLON", "đâu cap|cap trung thê|cap[^,;]*\\b(?:22|24|35)\\s?kv|\\b(?:22|24|35)\\s?kv[^,;]*cap"], ["CAP", "\\b(?:cu|al)\\s*/|xlpe|cxv|cvv|lszh|mm2|mm²"], ["CSTĐ", "tiêp đia|tiêp đât|nôi đât|thu set|thoat set|chông set(?! lan truyên)|thu lôi|kim thu|earth|grounding|lightning|đông trân|han hoa nhiêt|han hoa nhiêt|môi han|bô ghep(?! kênh)|coc[^,;]*tiêp|băng đông (?:tiêp|trân|\\d)|đêm set|giam điên trơ"], ["VTP", "vât tư phu|vât tư đâu nôi|phu kiên kêt nôi đâu cap|consumable|đâu côt|băng keo|gia (?:đơ|treo|đăt|kê)|ty ren|ty treo|thanh ren|thep hinh|\\bđai\\b(?! ôc)|\\bkep\\b|\\bcum\\b|bu ?lông|tăc kê"], ["BW", "busway|busduct|bus duct|thanh dân điên|thanh cai|\\bfeeder\\b|workmanship|hanger|elbow|flanged end|tr connection|plug-?in|\\bph box|tap.?off|gia công co|end cap"], ["TMC", "mang cap|thang cap|trunking|cable tray|cable ladder|khay cap|câu đơ cap"], ["MFD", "may phat|genset|generator|tu câp dâu|ông xa.*giam thanh|\\bmfđ\\b|nhiên liêu|câp dâu|bôn dâu"], ["UPS", "\\bups\\b|ăc quy|ac quy|acquy|battery|bô nguôn|power supply|bô lưu điên"], ["TBCC", "tu (?:đưng )?(?:voi|chưa chay)|hôp (?:đưng )?(?:dung cu )?chưa chay|binh chưa chay|binh bôt|binh co2|\\blăng\\b|voi chưa chay|cuôn (?:ông|voi)|hong (?:nươc|tiêp nươc|chưa chay)|tru (?:tiêp nươc|cưu hoa|cưu hoa|chưa chay)|siamese|hose|extinguisher|chăn chiên|đâu phun|sprinkler|drencher"], ["FM200", "khi chưa chay|chưa chay (?:băng )?khi|fm.?200|novec|inergen|aerosol|binh xit|nut ân (?:tam dưng )?xa khi|(?:đen|coi)[^,;]*xa khi|binh chưa khi"], ["FOAM", "\\bfoam\\b|bot (?:chưa chay|cô đăc)"], ["CAP DN", "dây nhay|patch cord"], ["TBDN", "\\bpdu\\b|patch|\\bodf\\b|\\bsfp\\b|modul(?:e)? quang|mô[\\s-]?đun (?:quang|tăng âm)|(?<!công tăc )\\bswitch\\b|router|firewall|server|may chu|\\brack\\b|tu mang|tăng âm|amply|amplifier|khuêch đai|chuông cưa|video|intercom|converter|chuyên đôi quang|\\bpoe\\b"], ["CAP", "^\\s*(?:cap|dây)[^,;]{0,40}bao chay"], ["CAP DN", "cat\\.?\\s?[56]|cap mang|cap quang|cap đông truc|dây tin hiêu|cap tin hiêu|cap điên thoai|\\butp\\b|\\bftp\\b|coaxial|\\brg-?(?:6|11)\\b|dây loa|cap loa|hdmi|rs-?485"], ["TĐ", "^\\s*(?:e?msb|smdb|mdb|db|dp|ldb|pdb|edb|cdb|acp|mcc|lp|pp(?!-?r\\b))[\\s\\-_.]|tu điên|tu phân phôi|tu điêu khiên|tu nut nhân|tu câp nguôn|tu đông lưc|tu chiêu sang|tu ats|\\bats\\b|switchboard|distribution board|tu bu|tu bu|tu ha thê|\\bspd\\b|chông set lan truyên|căt set|^\\s*tu\\b(?! (?:đưng |)voi| rack| mang| trung tâm)"], ["KSRV", "đâu đoc the|the (?:tư|cam ưng|nhân viên|ra vao|rfid|mifare|proximity)|\\brfid\\b|\\bcard\\b"], ["TBBC", "bao chay|đâu (?:bao|do)|\\bchuông\\b|\\bcoi\\b|nut (?:ân|nhân) (?:bao chay|khân|điêu khiên)|fire alarm|detector|\\bloop\\b|module|mô[\\s-]?đun|modul|điên trơ cuôi|tu trung tâm|bang hiên thi phu|đen (?:bao|chơp)(?! không)|cam biên (?:khoi|nhiêt)(?! đô)"], ["TBDN", "camera|cctv|\\bnvr\\b|\\bhdd\\b|đâu ghi|man hinh|\\btv\\b|tivi|\\bbms\\b|scada|\\bloa\\b|speaker|micro|âm thanh|bô phat|ipbx|tông đai|matv|splitter|bô chia tin hiêu|may tram|wifi|access point|bô chuyên mach|phân mêm|may in|ban phim|điên thoai|kiêm soat (?:ra|vao)|access control|khoa (?:điên )?tư|khoa (?:điên )?tư|barrier|exit button|nut (?:nhân |bâm )?(?:mơ cưa|exit)"], ["ACB", "\\bacb\\b|may căt không khi|air circuit breaker"], ["MCCB", "\\bmccb\\b"], ["MCB", "\\bmcb\\b|\\brcbo\\b|\\brccb\\b|\\belcb\\b|\\brcd\\b|isolator|câu dao cach ly"], ["CB", "aptomat|ap tô mat|attomat|câu dao tư đông|circuit breaker|\\bcb\\b"], ["TBDH", "\\bidu\\b|\\bcdu\\b|\\bfcu\\b|\\bahu\\b|\\bpau\\b|\\berv\\b|\\bhru\\b|chiller|cooling tower|thap giai nhiêt|điêu hoa|điêu hoa|\\bvrv\\b|\\bvrf\\b|dan nong|dan lanh|split|cassette|refnet|bô chia gas"], ["QUAT", "quat|\\bfan\\b|\\b(?:paf|eaf|spf|sef|cef|bpf|tef|kef|saf|ef|jf)[\\s\\-]"], ["VCG", "miêng gio|cưa gio|cưa (?:câp|hôi|thai|hut|lây|tăng ap|xa)|van gio|louver|\\blover\\b|grille|diffuser|lươi chăn côn|\\b(?:sag|rag|seg|reg|eag|oag|fag|lg|sad|rad|ead|is|vcd|obd|fd|fsd|nrd)\\b|\\bmfd\\b|damper|van chăn lưa|van xa ap|\\bprd\\b|van đong mơ|linear|hôp gio|plenum"], ["TBVS", "bôn câu|\\bb\\.? ?câu\\b|xi bêt|\\bwc\\b|lavabo|châu rưa|bôn rưa|bôn tiêu|tiêu nam|voi sen|sen tăm|gương|voi xit|bôn tăm|treo khăn|sây tay|giây vê sinh|lô (?:đê )?giây|khay xa|kê kinh|kê xa phong|thiêt bi vê sinh|phêu thu|phêu thoat|thu san|thoat san|floor drain|roof drain|câu chăn rac|nha bêp|voi rưa|\\bbêp\\b"], ["VAN", "^\\s*van\\b|\\bvan (?:bươm|công|câu|bi|môt chiêu|1 chiêu|goc|giam ap|xa|an toan|phao|điên tư|chăn|khoa|khoa|bao đông|tran|cân băng|ôn ap|hai chiêu|2 chiêu)|valve|y loc|strainer|đông hô|khơp nôi mêm|khơp (?:gian nơ|chông rung)|ro bơm|flow switch|công tăc (?:ap|dong|lưu lương|mưc|chênh|giam sat)|binh tich ap|binh gian nơ|ap kê|nhiêt kê|gauge|tram van|chông va|chông xoay|kinh quan sat|cam biên (?:ap|mưc|nhiêt đô|lưu lương)|que cam biên"], ["CTOC", "công tăc(?! (?:ap|dong|mưc|chênh|giam sat|hanh trinh|lưu lương|kiêm soat))|ô căm|socket|hat (?:công tăc|ô)"], ["EXIT", "sơ đô|biên (?:bao|chi dân)|tiêu lênh|nôi quy|bang (?:chi dân|hương dân)"], ["DEN", "(?<!thep )(?<!van )(?<!mau )(?<!sơn )\\bđen\\b|\\bexit\\b|thoat hiêm|măt êch|\\blight|luminaire|highbay|downlight|\\bled\\b|cam biên|sensor|aviation|obstruction"], ["LLMT", "năng lương măt trơi|pin măt trơi|tâm pin|solar"], ["MNN", "may nươc nong|binh nong lanh|binh nươc nong|heat pump|bơm nhiêt|nươc nong \\d+ ?l\\b"], ["OLD", "ông luôn|luôn dây|conduit|\\bemt\\b|\\bimc\\b|\\brsc\\b|ruôt ga|ông dân(?! nươc)|ông g\\.?i d\\d|ông hdpe d\\d+/\\d+|hôp (?:nôi|đâu|chia|ky thuât đâu|điên)|\\bbox\\b(?! gio)|đê âm|hôp âm"], ["BO", "^\\s*(?:bao ôn|cach nhiêt|cach âm|tiêu âm|insulation)"], ["OD", "ông đông|copper|\\bø\\s*\\d+[.,]\\d+\\s*mm"], ["UPVC", "nươc ngưng|condensate"], ["OG", "ông gio|g\\.?i day|\\bduct\\b|ductwork|tôn ma kem|\\bei\\s?\\d+"], ["CCL", "chông chay lan|bit kin|fire ?stop|chen chông chay|vưa chông chay"], ["BO", "bao ôn|cach nhiêt|cach âm|tiêu âm|insulation"], ["Vencap", "vent ?cap|chup thông hơi|năp thông hơi"], ["INOX", "ông (?:thep )?(?:inox|không gi|không ri)|stainless steel pipe|ông sus|inox[^,;]*\\bdn\\s?\\d"], ["PPR", "\\bppr\\b|pp-r|nươc lanh|nươc nong|nươc câp sinh hoat|han nhiêt"], ["HDPE", "hdpe|âm đât.*thai|thai.*âm đât"], ["UPVC", "upvc|u-pvc|\\bpvc\\b|truc đưng|ông xa|thoat nươc thai trong nha|\\bpp\\b"], ["OTH", "ông thep|thep trang kem|thep ma kem|ma kem|thep đen|\\bsch\\s?\\d+|bs\\s?1387|âm đât|steel pipe|\\bg\\.?i\\b|ông trong nha"], ["BOM", "bơm|pump|jockey|loc nươc|binh loc|than hoat tinh|loc ap lưc|softener|lam mêm|khư trung|\\buv\\b"], ["CTK", "\\bbê\\b|\\bbê\\b|hô ga|hô thu|bê phôt|công (?:thoat|hôp|tron|ban|bê tông|d\\d|ø)|song chăn|năp (?:hô|ga|bê|đan)|tâm đan|bôn chưa|bôn nươc|\\btec\\b|\\btank\\b|gas bô sung|nap gas|công tac khac|\\blâp\\b|bê tông|\\bmong\\b|\\bxây\\b|\\btrat\\b|ôp lat|\\bgach\\b|chông thâm|\\bsơn\\b|chông ri|khoan|\\bđuc\\b|căt (?:bê tông|tương|san|đương)|cưa sâp|cưa (?:kiêm tra|thăm)|access door|access panel"], ["CAP", "mm2|mm²|\\bcu/|cxv|cvv|\\bcv\\b|xlpe|lszh|\\bcu\\b.*pvc|cap điêu khiên|\\bawg\\b|cap điên|dây điên|dây dân|\\bcap\\b|cable"]], "GENERIC": "^\\s*(dn|d|ø|phi|size)?\\s*[\\d.,/x×\\s\\-]+(mm)?\\s*$|^\\s*phu kiên( cho ông)?|^\\s*phong\\b|^\\s*(tâng|phân|khôi|thap|tower|block|hâm)\\b", "DOMINANT": [["FM200", "binh xit|aerosol|chưa chay khi|khi chưa chay|fm.?200"], ["FOAM", "\\bfoam\\b"], ["MFD", "may phat"], ["TBBC", "hê thông bao chay|fire alarm system"]], "NON_DOMINATED": ["BO", "CAP", "CAP DN", "CAPLON", "CCL", "CPK", "CSTĐ", "CTK", "HDPE", "INOX", "OLD", "OTH", "PPR", "TBCC", "TBDN", "TMC", "TS", "UPVC", "VTP"], "PIPE_CODES": ["HDPE", "INOX", "OD", "OG", "OLD", "OTH", "PPR", "UPVC"], "ELEC_SYSTEM": "điên|\\belv\\b|\\blv\\b|chông set|camera|cctv|âm thanh|thông tin|truyên hinh|kiêm soat ra vao|bao chay", "PIPE_WORD": "\\bông\\b|\\bpvc\\b|hdpe|\\bg\\.?i\\b|conduit", "NOT_CONDUIT": "ông đông|ông gio|nươc|thoat|câp thoat", "PREFIX": "^\\s*(cung câp|vân chuyên|lăp đăt|thi công|supply|install)[\\s,/&va-]*((cung câp|vân chuyên|lăp đăt|thi công|and|va|,|&|/)\\s*)*", "CUT": "\\s+(bao gôm|gôm|c/w|including|kem theo|kem|co|with|tai)\\b|[(:;]", "PIPE_HEAD": "\\s*(đương )?ông\\b|\\s*pipe", "VAN_HEAD": "\\bvan\\b|văn chăn", "PHU_KIEN": "\\s*phu kiên", "PHU_KIEN_HEAD": "\\s*(phu kiên|vât tư)", "VTP_EXCEPT": "đâu cap|chông set|thu set|thoat set|tiêp đia", "LOWV_SYSTEM": "điên nhe|\\belv\\b|camera|cctv|âm thanh|thông tin|truyên hinh|mang|kiêm soat ra vao|điên thoai|\\bdata\\b", "CONDUIT_HEAD": "ông luôn|luôn dây|conduit|ruôt ga", "HDR_KEYS": {"stt": "^(stt|tt|no\\.?|item|sô tt|s\\.tt)$", "desc": "mô ta|hang muc|nôi dung|description|tên (vât tư|thiêt bi|công viêc|hang)|diên giai|work item", "spec": "thông sô|quy cach|specification|đăc tinh|spec", "code": "ma hiêu|model|ma sô|ky hiêu|\\bcode\\b|part no", "brand": "nhan hiêu|hang|brand|nha san xuât|manufacturer", "origin": "xuât xư|origin|nươc san xuât|country", "unit": "^(đvt|đơn vi( tinh)?|unit|đv)$", "qty": "khôi lương|\\bkl\\b|quantity|q'?ty|sô lương|^sl$", "price": "đơn gia|thanh tiên|unit rate|rate|amount|gia|price|total amount"}, "SKIP_SHEET": "bia|cover|chi dân|chi dan|hương dân|tông hơp|summary|thuyêt minh|ghi chu|note|đơn gia|dulieu|data", "TOT_HEAD": "tông|total|công|sum", "TOT_ROW": "^(tông công|tông gia tri|tông hơp|công\\s*(:|$)|total\\b|grand total)", "PROJ": "dư an|project|công trinh"};
  const TONES = /[\u0300\u0301\u0303\u0309\u0323]/g;
  const detone = (s) => (s === null || s === undefined ? "" : String(s)).toLowerCase().normalize("NFD").replace(TONES, "").normalize("NFC");
  const WB = "(?:(?<=[\\p{L}\\p{N}_])(?![\\p{L}\\p{N}_])|(?<![\\p{L}\\p{N}_])(?=[\\p{L}\\p{N}_]))";
  // \b của JavaScript không hiểu chữ có dấu -> đổi sang ranh giới Unicode (xử lý chuỗi, không phải eval)
  const conv = (p) => { let o = ""; for (let i = 0; i < p.length; i++) { if (p[i] === "\\" && i + 1 < p.length) { o += p[i + 1] === "b" ? WB : p.slice(i, i + 2); i++; } else o += p[i]; } return o; };
  const rx = (p, f) => new RegExp(conv(p), f || "u");
  const RULES = D.RULES.map(([c, p]) => [c, rx(p, "iu")]);
  const GENERIC = rx(D.GENERIC), ELEC = rx(D.ELEC_SYSTEM), PIPE_WORD = rx(D.PIPE_WORD), NOT_CONDUIT = rx(D.NOT_CONDUIT);
  const PREFIX = rx(D.PREFIX), CUT = rx(D.CUT);
  const PIPE_HEAD = rx("^(?:" + D.PIPE_HEAD + ")"), VAN_HEAD = rx(D.VAN_HEAD), PHU_KIEN = rx("^(?:" + D.PHU_KIEN + ")");
  const DOMINANT = D.DOMINANT.map(([c, p]) => [c, rx(p)]);
  const NON_DOM = new Set(D.NON_DOMINATED), PIPE_CODES = new Set(D.PIPE_CODES);
  const CABLE = /(\d+(?:[.,]\d+)?)\s*mm(?:2|²)/u;
  const AMPS = /(\d+(?:[.,]\d+)?)\s*a(?![\p{L}\p{N}_])/u;
  const PHU_KIEN_HEAD = rx("^(?:" + D.PHU_KIEN_HEAD + ")"), VTP_EXCEPT = rx(D.VTP_EXCEPT), LOWV = rx(D.LOWV_SYSTEM), CONDUIT = rx(D.CONDUIT_HEAD);
  const ROMAN = /^(?=[IVXLC])M*(C[MD]|D?C{0,3})(X[CL]|L?X{0,3})(I[XV]|V?I{0,3})\.?$/;
  const norm = (v) => (v === null || v === undefined ? "" : String(v).normalize("NFC").trim());
  const empty = (v) => v === null || v === undefined || (typeof v === "string" && v.trim() === "");
  const cellVal = (v) => (empty(v) ? null : (typeof v === "string" ? v.trim() : v));
  const toQty = (v) => {
    if (typeof v === "number") return v;
    if (empty(v)) return null;
    const n = Number(String(v).replace(/,/g, ""));
    return Number.isFinite(n) ? n : null;
  };

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

  function subject(text) {
    const t = detone(text).replace(PREFIX, "");
    const m = CUT.exec(t);
    return m && m.index > 2 ? t.slice(0, m.index) : t;
  }
  function match(text, only) {
    const t = detone(text);
    for (const [c, r] of RULES) {
      if (only && only.size && !only.has(c)) continue;
      if (r.test(t)) return c;
    }
    return null;
  }
  const matchCtx = (text, only) => match(subject(text), only) || match(text, only);
  // Tinh chỉnh mã sau khi khớp luật (giống _cap_size bản Python):
  // CAP: tiết diện >= 70mm2 -> CAPLON; hệ Điện nhẹ -> CAP DN. CB: <=63A MCB, <800A MCCB, >=800A ACB, không rõ -> MCCB.
  function refine(code, system, ...texts) {
    if (code === "CB") {
      for (const t of texts) {
        const m = AMPS.exec(detone(t));
        if (m) { const a = parseFloat(m[1].replace(",", ".")); return a <= 63 ? "MCB" : (a < 800 ? "MCCB" : "ACB"); }
      }
      return "MCCB";
    }
    if (code !== "CAP") return code;
    for (const t of texts) {
      const m = CABLE.exec(detone(t));
      if (m && parseFloat(m[1].replace(",", ".")) >= 70) return "CAPLON";
    }
    if (system && LOWV.test(detone(system))) return "CAP DN";
    return code;
  }

  // Trả [mã, nguồn]. nguồn: desc | system | context | none
  function classify(desc, context, system, ranked) {
    context = context || [];
    const d = detone(desc);
    // "Phụ kiện …" luôn là VTP (bảng Tra cứu)
    if (PHU_KIEN_HEAD.test(d) && !VTP_EXCEPT.test(d)) return ["VTP", "desc"];
    const generic = !d || GENERIC.test(d);
    let descCode = null;
    if (!generic) {
      const sub = match(subject(d)), full = match(d);
      descCode = sub || full;
      if (sub && PIPE_CODES.has(sub) && full && !PIPE_CODES.has(full)) descCode = full;
    }
    const nearSubj = context.length ? subject(context[0]) : "";
    const own = generic ? nearSubj : d;
    const inConduit = context.slice(0, 2).some((c) => CONDUIT.test(detone(c)));
    if ((ELEC.test(detone(system)) || inConduit) && PIPE_WORD.test(own) && !NOT_CONDUIT.test(own) && (descCode === null || PIPE_CODES.has(descCode)))
      return ["OLD", "system"];
    if (!NON_DOM.has(descCode)) {
      for (const [lvl, txt] of ranked || []) {
        if (lvl !== null && lvl < 2) continue;
        for (const [c, r] of DOMINANT) if (r.test(subject(txt)) || r.test(detone(txt).slice(0, 60))) return [c, "system"];
      }
    }
    if (descCode) return [refine(descCode, system, desc), "desc"];
    if (context.length && PIPE_HEAD.test(nearSubj) && !VAN_HEAD.test(nearSubj)) {
      for (const ctx of context) { const c = matchCtx(ctx, PIPE_CODES); if (c) return [c, "context"]; }
      return [null, "none"];
    }
    for (const ctx of context) { const c = matchCtx(ctx); if (c) return [refine(c, system, desc, ctx), "context"]; }
    if (d) {
      const c = match(d); if (c) return [refine(c, system, desc), "desc"];
      if (PHU_KIEN_HEAD.test(d)) return ["VTP", "fallback"];
    }
    return [null, "none"];
  }

  function detectKind(stt, unit, qtys, prevLetter) {
    const s = norm(stt), hasUnit = !!norm(unit);
    const hasQty = qtys.some((q) => typeof q === "number" && q !== 0);
    if (s) {
      const letter = s.replace(/\.+$/, "");
      if (/^[A-H]$/.test(letter) && !hasUnit) {
        if (!"CD".includes(letter) || prevLetter === String.fromCharCode(letter.charCodeAt(0) - 1)) return ["section", 0];
      }
      if (ROMAN.test(s) && !hasUnit) return ["section", 1];
      if (/^[IVXLC]+\.\d+\.?$/.test(s) && !hasUnit) return ["section", 2];
      if (/^[IVXLC]+\.\d+\.\d+\.?$/.test(s) && !hasUnit) return ["section", 3];
    }
    if (hasUnit || hasQty) return ["item", null];
    return ["note", null];
  }

  // src: [{r, stt, desc, spec, code, brand, origin, unit, qty:[...], bold}] theo thứ tự dòng của sheet CĐT
  // state: trạng thái nối giữa các khúc (tiêu đề cha, mục chữ cái, hạng mục trước) -> phân tích từng khúc vẫn ra kết quả y hệt
  function analyze(src, state) {
    state = state || {};
    const rows = [];
    let context = state.context || [], prevLetter = state.prevLetter || null, prevItem = state.prevItem || null;
    for (const s of src) {
      const stt = cellVal(s.stt), desc = cellVal(s.desc), unit = cellVal(s.unit);
      const other = [s.spec, s.code, s.brand, s.origin].map(cellVal);
      const qty = (s.qty || []).map(toQty);
      if ([stt, desc, unit, ...other].every((v) => v === null) && !qty.some((q) => q)) continue;
      const [kind, level] = detectKind(stt, unit, qty, prevLetter);
      if (kind === "section" && level === 0) prevLetter = norm(stt).replace(/\.+$/, "");
      const row = { src_row: s.r, stt, desc, spec: other[0], code_cdt: other[1], brand: other[2], origin: other[3],
                    unit, qty, kind, level, bold: !!s.bold, italic: !!s.italic };
      if (kind === "section") context = context.filter((c) => c[0] < level).concat([[level, norm(desc)]]);
      else if (kind === "note") {
        context = context.filter((c) => c[0] < 9).concat([[9, norm(desc)]]);
        context = context.filter((c) => c[0] < 9).concat(context.filter((c) => c[0] === 9).slice(-2));
      }
      const rev = context.slice().reverse();
      row.context = rev.slice(0, 5).map((c) => c[1]);
      row.system = context.filter((c) => c[0] !== null && c[0] <= 1).map((c) => c[1]).join(" | ");
      if (kind === "item") {
        let [code, srcTag] = classify(desc, row.context, row.system, rev);
        if (code === null && row.spec) [code, srcTag] = classify(row.spec, [], row.system, null);
        row.cl = code; row.cl_source = srcTag;
      } else { row.cl = null; row.cl_source = "-"; }
      if (kind === "item" && PHU_KIEN.test(detone(desc))) {
        const prev = prevItem;
        const same = (prev && (prev.ctx0 || null) === (row.context[0] || null));
        if (prev && prev.cl && same && row.cl_source !== "desc") { row.cl = prev.cl; row.cl_source = "inherit"; }
      }
      if (kind === "item") prevItem = { cl: row.cl, ctx0: row.context[0] || null };
      rows.push(row);
    }
    state.context = context; state.prevLetter = prevLetter; state.prevItem = prevItem;
    return rows;
  }

  // ---- soát theo nhóm (BOQ lớn): khóa = tiêu đề gần nhất < tiêu đề cha, giống bản Python
  const groupKey = (x) => (x.context || []).slice(0, 2).map((t) => String(t).trim().slice(0, 80)).join(" < ");
  function review(rows, limit) {
    const G = new Map();
    for (const x of rows) {
      if (x.kind !== "item" || !["context", "none", "inherit", "system"].includes(x.cl_source)) continue;
      const k = groupKey(x); if (!G.has(k)) G.set(k, { n: 0, codes: {}, first: x.src_row, last: x.src_row, vd: [] });
      const g = G.get(k); g.n++; g.last = x.src_row; const c = x.cl || "??"; g.codes[c] = (g.codes[c] || 0) + 1;
      if (g.vd.length < 2) g.vd.push(String(x.desc || "").replace(/\s+/g, " ").slice(0, 35));
    }
    const list = [...G.entries()].sort((p, q) => (("??" in q[1].codes) - ("??" in p[1].codes)) || q[1].n - p[1].n);
    return { tong_nhom: list.length, nhom: list.slice(0, limit || 150).map(([k, g]) => ({ ctx: k, n: g.n, ma: g.codes, vd: g.vd, dong: [g.first, g.last] })) };
  }
  // OV: { "1234": "MÃ" (theo dòng CĐT), "ctx:<khóa nhóm>": "MÃ" (cả nhóm) }
  function applyOverrides(rows, OV) {
    const heads = {}; for (const k of Object.keys(OV || {})) if (k.startsWith("ctx:")) heads[k.slice(4).trim()] = OV[k];
    for (const x of rows) {   // giống bản Python: sửa theo số dòng áp cho mọi dòng, sửa theo nhóm chỉ áp cho hạng mục
      const own = OV[String(x.src_row)];
      if (own) { x.cl = own; x.cl_source = "claude"; continue; }
      if (x.kind !== "item") continue;
      const h = heads[groupKey(x)];
      if (h) { x.cl = h; x.cl_source = "claude"; }
    }
    return rows;
  }
  // Gán số dòng đầu ra (bắt đầu từ dòng 13) và dòng kết thúc của mỗi section
  // first = 13 (dựng sheet mới) hoặc "inplace" (giữ nguyên dòng CĐT)
  function plan(rows, first) {
    const inplace = first === "inplace";
    first = inplace ? rows[0].src_row : (first || 13);
    rows.forEach((x, i) => { x.out_row = inplace ? x.src_row : first + i; if (x.kind === "section" && (x.level === null || x.level === undefined)) x.level = 1; });
    const last = inplace ? rows[rows.length - 1].src_row : first + rows.length - 1;
    rows.forEach((x, i) => {
      if (x.kind !== "section") return;
      let end = last;
      for (let j = i + 1; j < rows.length; j++) if (rows[j].kind === "section" && rows[j].level <= x.level) { end = rows[j].out_row - 1; break; }
      x.sec_end = Math.max(end, x.out_row + 1);
    });
    return { first, last, total_row: last + 1 };
  }

  // Bước hoàn tất (không dùng regex, rất nhanh): vals = các dòng A:K của _TDI_TMP (không gồm tiêu đề)
  // -> áp OV, tính out_row / sec_end, trả cột D:E (cl, nguồn) và G:H (out_row, sec_end) để ghi lại, kèm danh sách nhóm cần soát
  function finalize(vals, OV, first) {
    const rows = vals.filter((v) => v[0] !== "" && v[0] !== null).map((v) => ({ src_row: v[0], kind: v[1], level: v[2] === "" ? null : v[2],
      cl: v[3] || null, cl_source: v[4], bold: v[5] === 1, desc: v[8], context: [v[9], v[10]].filter((t) => t !== "" && t !== null && t !== undefined) }));
    applyOverrides(rows, OV || {});
    const p = plan(rows, first === undefined ? "inplace" : first);
    return { p, rows, DE: rows.map((x) => [x.cl || "", x.cl_source || ""]), GH: rows.map((x) => [x.out_row, x.sec_end || ""]) };
  }

  // ================= TỰ NHẬN DIỆN CẤU TRÚC SHEET (giống lệnh auto của bản Python) =================
  const HK = Object.fromEntries(Object.entries(D.HDR_KEYS).map(([k, p]) => [k, rx(p)]));
  const SKIP = rx(D.SKIP_SHEET), TOT_HEAD = rx(D.TOT_HEAD), TOT_ROW = rx(D.TOT_ROW), PROJ = rx(D.PROJ, "iu");
  const sp = (v) => String(v).split(/\s+/).filter(Boolean).join(" ");
  const tx = (v) => (v === null || v === undefined || v === "" ? "" : detone(sp(v)));
  const isNum = (v) => typeof v === "number";
  // head: giá trị dòng 1..n (mảng 2 chiều từ A1), merges: [[r1,c1,r2,c2]], probe(r,c), tail: {r0, vals}, lastRow
  function detectSheet(head, merges, probe, tail, lastRow) {
    const maxc = Math.min(40, Math.max(0, ...head.map((r) => r.length)));
    const cell = (r, c) => (head[r - 1] && head[r - 1][c - 1] !== undefined ? head[r - 1][c - 1] : (probe ? probe(r, c) : ""));
    const mt = (r, c) => { for (const m of merges) if (m[0] <= r && r <= m[2] && m[1] <= c && c <= m[3]) return tx(cell(m[0], m[1])); return tx(cell(r, c)); };
    let best = [0, null];
    for (let r = 1; r <= Math.min(head.length, 40); r++) {
      const hits = new Set();
      for (let c = 1; c <= maxc; c++) { const t = mt(r, c); if (!t) continue; for (const k in HK) if (HK[k].test(t)) hits.add(k); }
      const sc = ["stt", "desc", "unit", "qty"].filter((k) => hits.has(k)).length * 2 + hits.size;
      if (sc > best[0]) best = [sc, r];
    }
    if (best[0] < 5) return null;
    const h1 = best[1]; let h2 = h1;
    const nxt = Array.from({ length: maxc }, (_, i) => tx(cell(h1 + 1, i + 1)));
    const vertical = merges.some((m) => m[0] === h1 && m[2] >= h1 + 1);
    if (vertical || (nxt.filter((t) => t && !/^[\d\s.,-]+$/.test(t)).length >= 2 &&
        !Array.from({ length: maxc }, (_, i) => cell(h1 + 1, i + 1)).some(isNum))) h2 = h1 + 1;
    let start = h2 + 1;
    const nums = Array.from({ length: maxc }, (_, i) => cell(start, i + 1));
    if (nums.filter(isNum).length >= 4 && nums.every((v) => isNum(v) || v === null || v === "")) start++;
    const cols = {}, zones = []; let total = null;
    const qtyMerge = (c) => merges.some((m) => m[0] === h1 && m[1] <= c && c <= m[3] && HK.qty.test(tx(cell(h1, m[1]))));
    const nm = (r, c) => sp(cell(r, c) === "" || cell(r, c) === null ? "" : cell(r, c));
    for (let c = 1; c <= maxc; c++) {
      const top = mt(h1, c), low = h2 !== h1 ? mt(h2, c) : top, both = (top + " | " + low).trim(), L = colL(c);
      if (HK.price.test(both) && !HK.qty.test(top !== low ? top : both)) continue;
      let f = null;
      for (const k of ["stt", "desc", "spec", "code", "brand", "origin", "unit"]) if (!(k in cols) && (HK[k].test(low) || HK[k].test(top))) { f = k; break; }
      if (f) { cols[f] = L; continue; }
      if (HK.qty.test(top) || HK.qty.test(low)) {
        if (TOT_HEAD.test(h2 !== h1 ? low : both)) total = L;
        else zones.push({ name: nm(h2, c) || nm(h1, c) || "KL " + L, col: L });
      } else if (h2 !== h1 && zones.length && low !== "" && !/tông|total/.test(low) && qtyMerge(c)) zones.push({ name: nm(h2, c), col: L });
      else if (h2 !== h1 && /tông/.test(low) && qtyMerge(c)) total = L;
    }
    if (!zones.length && cols.unit && !probe) return null;   // cần đọc thêm số liệu để dò cột KL
    if (!zones.length && cols.unit && probe) {
      for (let c = cIdx(cols.unit) + 1; c <= maxc; c++) {
        const hd = mt(h2, c) || mt(h1, c);
        if (!hd || HK.price.test(hd)) { if (zones.length || total) break; continue; }
        const filled = []; for (let r = start; r < start + 300 && r <= lastRow; r++) { const v = probe(r, c); if (v !== null && v !== "" && v !== undefined) filled.push(v); }
        if (!filled.length || filled.filter(isNum).length / filled.length < 0.7) break;
        if (colL(c) === total || TOT_HEAD.test(hd)) { total = colL(c); break; }
        zones.push({ name: nm(h2, c) || nm(h1, c), col: colL(c) });
      }
    }
    if (!zones.length && total) { zones.push({ name: "Khối lượng", col: total }); total = null; }
    if (!cols.desc || !zones.length) return null;
    const dc = cIdx(cols.desc), tv = (r, c) => { const row = tail.vals[r - tail.r0]; return row ? row[c - 1] : ""; };
    const emptyRow = (r) => { for (let c = 1; c <= maxc; c++) { const v = tv(r, c); if (v !== null && v !== "" && v !== undefined) return false; } return true; };
    let end = lastRow;
    while (end > start && emptyRow(end)) end--;
    for (let r = end; r > Math.max(start, end - 40); r--) {
      let hit = TOT_ROW.test(tx(tv(r, dc)));
      for (let c = 1; c < dc && !hit; c++) hit = TOT_ROW.test(tx(tv(r, c)));
      if (hit) { end = r - 1; while (end > start && emptyRow(end)) end--; break; }
    }
    for (const k of ["stt", "desc", "spec", "code", "brand", "origin", "unit"]) if (!(k in cols)) cols[k] = null;
    cols.zones = zones.map((z) => z.col);
    return { header_rows: [h1, h2], start, end, header_row: h2, cols, total, zone_names: zones.map((z) => z.name) };
  }
  const parseAreas = (addr) => String(addr || "").split(",").map((a) => a.trim().replace(/^.*!/, "").replace(/\$/g, "")).filter(Boolean).map((a) => {
    const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a); return m ? [+m[2], cIdx(m[1]), +(m[4] || m[2]), cIdx(m[3] || m[1])] : null; }).filter(Boolean);

  // ================= PHÂN TÍCH MỘT SHEET (gọi lại cho tới khi xong) =================
  const tmpName = (src) => "_TDI_TMP_" + String(src).slice(0, 20);
  const colL = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  async function run(context, CFG, OV, deadline, reserve) {
    deadline = deadline || Date.now() + (CFG.budget || 40000);
    reserve = reserve === undefined ? 12000 : reserve;   // thời gian chừa cho bước hoàn tất
    const T0 = Date.now(), sec = () => Math.round((Date.now() - T0) / 1000), left = () => deadline - Date.now();
    let SIZE = 250, lastDur = 0;
    const wb = context.workbook, ws = wb.worksheets.getItem(CFG.src), name = tmpName(CFG.src);
    const s1 = wb.worksheets.getItemOrNullObject("Summary"), s2 = wb.worksheets.getItemOrNullObject("Summary TDI");
    let tmp = wb.worksheets.getItemOrNullObject(name);
    await context.sync();
    const z = [s1, s2].map((s) => { if (s.isNullObject) return null; const r = s.getRange("Z1"); r.load("values"); return r; });
    await context.sync();
    const zl = z.map((r) => { try { return r ? JSON.parse(r.values[0][0] || "null") : null; } catch (e) { return null; } });
    const sumName = s1.isNullObject || zl[0] ? "Summary" : "Summary TDI";
    const doneList = zl[0] || zl[1] || [];
    if (tmp.isNullObject && doneList.some((x) => x.sheet === CFG.src))
      return { loi: `Sheet "${CFG.src}" đã có phần TDI (cột ${doneList.find((x) => x.sheet === CFG.src).tdi}). Muốn làm lại thì xóa các cột đó và dòng của sheet này trong ${sumName} trước.` };
    let state = {}, next = CFG.start, n = 0, info = null;
    if (tmp.isNullObject) {
      tmp = wb.worksheets.add(name); tmp.getRange("A1:M1").values = [TMP_HEAD]; tmp.visibility = "Hidden";
    } else {
      const meta = tmp.getRange("P1:S1"); meta.load("values"); await context.sync();
      if (meta.values[0][0]) { state = JSON.parse(meta.values[0][0]); next = meta.values[0][1]; n = meta.values[0][2]; }
      if (meta.values[0][3]) info = JSON.parse(meta.values[0][3]);
    }
    while (next <= CFG.end && left() > lastDur * 1.5 + 2000) {
      const t1 = Date.now(), r2 = Math.min(next + SIZE - 1, CFG.end);
      const h = loadSource(ws, CFG.cols, next, r2); await context.sync();
      const rows = analyze(srcRows(h, CFG.cols), state);
      if (rows.length) { const data = toTmp(rows); tmp.getRange(`A${n + 2}:M${n + 1 + data.length}`).values = data; n += data.length; }
      next = r2 + 1;
      tmp.getRange("P1:R1").values = [[JSON.stringify(state), next, n]];
      await context.sync();
      lastDur = Date.now() - t1;
      if (lastDur > 10000) SIZE = Math.max(25, Math.floor(SIZE / 2));
    }
    if (next <= CFG.end) return { xong: false, da_toi_dong_cdt: next - 1, tren: CFG.end, giay: sec() };
    if (n > 30000) return { loi: "Sheet trên 30.000 dòng: nên chạy skill bằng Python trên claude.ai / Desktop" };
    if (left() < reserve) return { xong: false, giay: sec() };
    const parts = [];
    for (let r = 2; r <= n + 1; r += 5000) { const g = tmp.getRange(`A${r}:K${Math.min(r + 4999, n + 1)}`); g.load("values"); parts.push(g); }
    let used = null;
    if (!(info && info.tdi_start) && !CFG.tdi_start) { used = ws.getUsedRange(true); used.load("columnIndex,columnCount"); }
    await context.sync();
    const fin = finalize(parts.flatMap((g) => g.values), OV || {}, "inplace");
    for (let i = 0; i < n; i += 5000) {
      const k = Math.min(i + 5000, n);
      tmp.getRange(`D${i + 2}:E${k + 1}`).values = fin.DE.slice(i, k);
      tmp.getRange(`G${i + 2}:H${k + 1}`).values = fin.GH.slice(i, k);
    }
    const tdi_start = CFG.tdi_start || (info && info.tdi_start) || colL(used.columnIndex + used.columnCount + 1);
    info = { tdi_start, summary: sumName, first: fin.p.first, last: fin.p.last, total_row: fin.p.total_row };
    tmp.getRange("S1:T1").values = [[JSON.stringify(info), ""]];
    await context.sync();
    const items = fin.rows.filter((x) => x.kind === "item");
    return { xong: true, tdi_start, summary: sumName, dong: n, hang_muc: items.length,
      chua_ma: items.filter((x) => !x.cl).length, dong_tong_tdi: fin.p.total_row, giay: sec(), ...review(fin.rows, 150) };
  }

  // ================= CẢ FILE: tự nhận diện mọi sheet BOQ rồi phân tích hết (gọi lại y nguyên cho tới khi xong) =================
  // OPT: { sheets?: ["tên", ...] (mặc định: tự tìm), OV?: { "<sheet>": {...} }, budget? }
  const JOB = "_TDI_JOB";
  async function readJob(context) {
    const j = context.workbook.worksheets.getItemOrNullObject(JOB); await context.sync();
    if (j.isNullObject) return null;
    const r = j.getRange("A1"); r.load("values"); await context.sync();
    try { return { ws: j, data: JSON.parse(r.values[0][0] || "null") }; } catch (e) { return { ws: j, data: null }; }
  }
  const saveJob = (job) => { job.ws.getRange("A1").values = [[JSON.stringify(job.data)]]; };
  async function detectAll(context, only) {
    const wb = context.workbook, list = wb.worksheets; list.load("items/name,items/visibility"); await context.sync();
    const cand = list.items.filter((w) => w.visibility === "Visible" && !/^_TDI|^Summary( TDI)?$/.test(w.name) && (!only || only.includes(w.name)) && (only || !SKIP.test(detone(w.name))));
    const skipped = list.items.filter((w) => !cand.includes(w) && !/^_TDI/.test(w.name)).map((w) => w.name);
    const H = cand.map((w) => { const u = w.getUsedRange(true); u.load("rowIndex,rowCount,columnIndex,columnCount"); return { w, u }; });
    await context.sync();
    for (const h of H) {
      h.lastRow = h.u.rowIndex + h.u.rowCount; h.maxc = Math.min(40, h.u.columnIndex + h.u.columnCount);
      h.head = h.w.getRange(`A1:${colL(h.maxc)}${Math.min(40, h.lastRow)}`); h.head.load("values");
      h.mg = h.head.getMergedAreasOrNullObject(); h.mg.load("address");
      h.tail = h.w.getRange(`A${Math.max(1, h.lastRow - 60)}:${colL(h.maxc)}${h.lastRow}`); h.tail.load("values");
    }
    await context.sync();
    let project = "";
    const firstPass = H.map((h) => {
      for (const row of h.head.values.slice(0, 12)) for (const v of row.slice(0, 8))
        if (!project && typeof v === "string" && PROJ.test(detone(v))) project = v.trim().replace(/^\s*(dự án|project|công trình)\s*[:：-]?\s*/i, "");
      return detectSheet(h.head.values, h.mg.isNullObject ? [] : parseAreas(h.mg.address), null, { r0: Math.max(1, h.lastRow - 60), vals: h.tail.values }, h.lastRow);
    });
    // lần 2: đọc thêm 300 dòng dữ liệu cho sheet cần dò cột KL theo số liệu
    const probes = H.map((h, i) => { const d = firstPass[i]; if (!d) { const g = h.w.getRange(`A1:${colL(h.maxc)}${Math.min(h.lastRow, 345)}`); g.load("values"); return g; } return null; });
    if (probes.some(Boolean)) await context.sync();
    const out = [];
    H.forEach((h, i) => {
      let d = firstPass[i];
      if (!d && probes[i]) { const pv = probes[i].values; d = detectSheet(h.head.values, h.mg.isNullObject ? [] : parseAreas(h.mg.address), (r, c) => (pv[r - 1] ? pv[r - 1][c - 1] : ""), { r0: Math.max(1, h.lastRow - 60), vals: h.tail.values }, h.lastRow); }
      if (!d || d.end - d.start + 1 < 5) { skipped.push(h.w.name); return; }
      out.push({ src: h.w.name, cfg: { src: h.w.name, name: h.w.name.toUpperCase().slice(0, 30), ...d }, a: 0, b: 0 });
    });
    return { sheets: out, skipped, project };
  }
  async function runAll(context, OPT) {
    OPT = OPT || {};
    const BUD = OPT.budget || 55000, T0 = Date.now(), deadline = T0 + BUD, sec = () => Math.round((Date.now() - T0) / 1000);
    const minLeft = Math.min(8000, BUD * 0.15), reserve = Math.min(12000, BUD * 0.2);
    let job = await readJob(context);
    if (!job || !job.data) {
      const d = await detectAll(context, OPT.sheets);
      if (!d.sheets.length) return { loi: "Không tìm thấy sheet BOQ nào (cần có cột STT / mô tả / đơn vị / khối lượng)", bo_qua: d.skipped };
      const ws = job && job.ws ? job.ws : context.workbook.worksheets.add(JOB); ws.visibility = "Hidden";
      job = { ws, data: d }; saveJob(job); await context.sync();
    }
    const S = job.data.sheets;
    for (const s of S) {
      if (s.a) continue;
      if (deadline - Date.now() < minLeft) break;
      const r = await run(context, s.cfg, (OPT.OV || {})[s.src] || {}, deadline, reserve);
      if (r.loi) { s.a = 2; s.loi = r.loi; }
      else if (r.xong) {
        s.a = 1;
        s.kq = { tdi_start: r.tdi_start, summary: r.summary, hang_muc: r.hang_muc, chua_ma: r.chua_ma, tong_nhom: r.tong_nhom,
                 nhom_can_soat: r.nhom.filter((g) => g.ma["??"] || Object.keys(g.ma).length > 1).slice(0, 25) };
        s.cfg.tdi_start = r.tdi_start;
      }
      saveJob(job); await context.sync();
      if (!r.xong && !r.loi) break;
    }
    const pending = S.filter((s) => !s.a).length;
    if (pending) return { buoc: "phân tích", xong: false, con_lai: pending + " sheet", giay: sec(), viec_tiep: "Chạy lại y nguyên script này" };
    return { buoc: "phân tích", xong: true, giay: sec(), du_an: job.data.project, bo_qua: job.data.skipped,
      sheets: S.map((s) => s.loi ? { sheet: s.src, loi: s.loi } : ({ sheet: s.src, tieu_de_dong: s.cfg.header_rows, du_lieu: [s.cfg.start, s.cfg.end],
        khu: s.cfg.zone_names.map((n, i) => n + "@" + s.cfg.cols.zones[i]), tong_kl: s.cfg.total, tdi_tu_cot: s.kq.tdi_start,
        hang_muc: s.kq.hang_muc, chua_ma: s.kq.chua_ma, tong_nhom: s.kq.tong_nhom, nhom_can_soat: s.kq.nhom_can_soat })) };
  }

  // Bảng tạm _TDI_TMP: lưu kết quả phân tích để các script sau dùng lại (mỗi script không nhớ gì)
  const TMP_HEAD = ["src_row", "kind", "level", "cl", "cl_source", "bold", "out_row", "sec_end", "desc", "context0", "context1", "unit", "italic"];
  // chữ bắt đầu bằng = + - @ sẽ bị Excel hiểu là công thức -> thêm dấu '
  const safe = (v) => { const t = String(v === null || v === undefined ? "" : v).slice(0, 120); return /^[=+\-@]/.test(t) ? "'" + t : t; };
  const toTmp = (rows) => rows.map((x) => [x.src_row, x.kind, x.level === null ? "" : x.level, x.cl || "", x.cl_source || "",
    x.bold ? 1 : 0, x.out_row || "", x.sec_end || "", safe(x.desc), safe((x.context || [])[0]), safe((x.context || [])[1]), safe(x.unit), x.italic ? 1 : 0]);
  const fromTmp = (vals) => vals.filter((v) => v[0] !== "").map((v) => ({ src_row: v[0], kind: v[1], level: v[2] === "" ? null : v[2],
    cl: v[3] || null, cl_source: v[4], bold: v[5] === 1, out_row: v[6], sec_end: v[7] || undefined }));

  return { runAll, run, detectSheet, parseAreas, detone, classify, detectKind, analyze, finalize, plan, groupKey, review, applyOverrides, TMP_HEAD, toTmp, fromTmp, loadSource, srcRows };
})();
// ============================== HẾT PHẦN 1 ==============================
