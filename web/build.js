#!/usr/bin/env node
/* dist/artifact.html 생성 — web/의 여러 파일을 하나로 인라인.
 *
 * 두 가지 산출물:
 *   1) web/index.html + web/app.js + web/styles.css + web/vendor/exceljs.min.js
 *      → 그대로 GitHub Pages / 로컬(더블클릭)에서 동작 (외부 CDN 불필요)
 *   2) dist/artifact.html
 *      → 모든 CSS/JS를 인라인한 단일 파일. claude.ai Artifact 규격(문서 골격 없이
 *        본문만)에 맞춰 <title>/<style>/본문/<script> 만 포함.
 *
 * 사용:  node web/build.js
 */
const fs = require("fs");
const path = require("path");

const WEB = __dirname;
const ROOT = path.dirname(WEB);
const read = (p) => fs.readFileSync(path.join(WEB, p), "utf8");

const css = read("styles.css");
const exceljs = read("vendor/exceljs.min.js");
const app = read("app.js");
const html = read("index.html");

// index.html 본문(.wrap ... </div>)만 추출
const bodyMatch = html.match(/<div class="wrap">[\s\S]*?<\/div>\s*<\/div>/);
// 위 정규식은 마지막 닫는 div까지 못 잡으므로, body 태그 사이를 안전하게 추출
const between = html.slice(html.indexOf("<body>") + "<body>".length, html.indexOf("</body>"));
// <script src=...> 라인 제거 (인라인으로 대체)
const bodyContent = between.replace(/<script\s+src=["'][^"']*["']><\/script>\s*/g, "").trim();

const out = `<title>발주양식 통합 서비스</title>
<style>
${css}
</style>
${bodyContent}
<script>
${exceljs}
</script>
<script>
${app}
</script>
`;

const distDir = path.join(ROOT, "dist");
if (!fs.existsSync(distDir)) fs.mkdirSync(distDir, { recursive: true });
const outPath = path.join(distDir, "artifact.html");
fs.writeFileSync(outPath, out, "utf8");
console.log(`생성됨: ${outPath}  (${(out.length / 1024).toFixed(0)} KB)`);
