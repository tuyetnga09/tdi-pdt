// Nén file JS cho add-in Excel (bỏ chú thích, rút gọn tên biến cục bộ; giữ nguyên tên TDI_A / TDI_B)
// Dùng: node minify.js <vào.js> <ra.js>   (cần gói terser: npm install terser)
const fs = require('fs');
let terser;
try { terser = require('terser'); } catch (e) { try { terser = require(process.env.TERSER_PATH || '/home/claude/jstools/node_modules/terser'); } catch (e2) { console.error('Thiếu terser'); process.exit(2); } }
(async () => {
  const src = fs.readFileSync(process.argv[2], 'utf8');
  const r = await terser.minify(src, { ecma: 2020, compress: { passes: 2 }, mangle: true, format: { comments: /^!/ } });
  fs.writeFileSync(process.argv[3], r.code);
})();
