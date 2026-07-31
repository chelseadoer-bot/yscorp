/* 발주양식 통합 서비스 — 브라우저 변환 엔진 + UI
 * 모든 처리는 브라우저 안에서만 이루어집니다. (파일이 서버로 전송되지 않음)
 * 엑셀 입출력: ExcelJS (전역 ExcelJS)
 */
(function () {
  "use strict";

  // =====================================================================
  // 표준 출력 양식 정의 (기준: 통합본)
  // =====================================================================
  const STANDARD_COLUMNS = [
    { key: "no",        header: "NO" },
    { key: "orderer",   header: "주문인명" },
    { key: "recipient", header: "수령인명" },
    { key: "phone",     header: "수령인핸드폰번호" },
    { key: "phone2",    header: "수령인핸드폰번호" },
    { key: "postcode",  header: "우편번호" },
    { key: "address",   header: "주소" },
    { key: "message",   header: "배송메세지" },
    { key: "product",   header: "상품정보" },
    { key: "qty",       header: "주문수량" },
    { key: "courier",   header: "택배사" },
    { key: "invoice",   header: "송장번호" },
  ];
  const FIELD_KEYS = STANDARD_COLUMNS.map((c) => c.key);
  const HEADERS = STANDARD_COLUMNS.map((c) => c.header);
  const WIDTHS = [4.25, 12.625, 13, 15, 13, 8.375, 23.25, 6.5, 42.25, 8, 13, 14.375];
  const CENTER_COLS = new Set([1, 2, 3, 10]);       // NO, 주문인명, 수령인명, 주문수량
  const HEADER_FILL_COLS = new Set([11, 12]);       // 택배사, 송장번호 (연한 주황)
  const TEXT_COLS = new Set([4, 5, 6, 12]);         // 전화/우편/송장 → 텍스트 서식
  const FONT_NAME = "맑은 고딕";

  // =====================================================================
  // 값 → 문자열 변환 (ExcelJS 셀 값의 다양한 형태 처리)
  // =====================================================================
  function cellToText(v) {
    if (v === null || v === undefined) return "";
    if (v instanceof Date) {
      const y = v.getFullYear(), m = v.getMonth() + 1, d = v.getDate();
      return `${y}${String(m).padStart(2, "0")}${String(d).padStart(2, "0")}`;
    }
    if (typeof v === "object") {
      if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join("");
      if ("text" in v) return String(v.text);
      if ("result" in v) return cellToText(v.result);
      if ("formula" in v) return "";
      if ("error" in v) return "";
      return String(v);
    }
    return String(v);
  }

  // =====================================================================
  // 정리 유틸
  // =====================================================================
  function cleanText(v) {
    if (v === null || v === undefined) return "";
    let t = String(v).trim();
    if (t.toLowerCase() === "nan") return "";
    return t.replace(/\s+/g, " ");
  }

  function normalizePhone(v) {
    let t = cleanText(v);
    if (!t) return "";
    return t.replace(/~/g, "-").replace(/\s/g, "");
  }

  function cleanPostcode(v) {
    let t = cleanText(v);
    if (!t) return "";
    let digits = t.replace(/\D/g, "");
    if (digits.length === 4) digits = "0" + digits;      // 엑셀이 날린 앞자리 0 복원
    if (digits.length === 5) return digits.slice(0, 3) + "-" + digits.slice(3);
    if (digits.length === 6) return digits.slice(0, 3) + "-" + digits.slice(3);
    return t.replace(/-{2,}/g, "-");
  }

  const QTY_TAIL = /\s*[/·\-]?\s*(\d+)\s*개\s*$/;
  function splitProductQty(v, defaultQty) {
    if (defaultQty === undefined) defaultQty = 1;
    let t = cleanText(v);
    if (!t) return { product: "", qty: defaultQty };
    const m = t.match(QTY_TAIL);
    if (m) {
      const qty = parseInt(m[1], 10);
      let product = t.slice(0, m.index).replace(/[\s/·\-▶]+$/, "").trim();
      return { product: product, qty: qty };
    }
    return { product: t, qty: defaultQty };
  }

  const ADDR_POST = /^\s*\[\s*(\d{3})\s*-?\s*(\d{2})\s*\]\s*/;
  function extractPostcodeFromAddress(v) {
    let t = cleanText(v);
    if (!t) return { postcode: "", address: "" };
    const m = t.match(ADDR_POST);
    if (m) return { postcode: m[1] + "-" + m[2], address: t.slice(m[0].length).trim() };
    return { postcode: "", address: t };
  }

  const ADDR_TAIL_NAMEPHONE = /\s{2,}[가-힣A-Za-z]{2,5}\s+0\d[\d~\-\s]{7,}$/;
  const ADDR_TAIL_PHONE = /\s{2,}0\d[\d~\-\s]{7,}$/;
  function cleanAddress(v, stripTrailingContact) {
    if (v === null || v === undefined) return "";
    let t = String(v).trim();
    if (t.toLowerCase() === "nan") return "";
    if (stripTrailingContact) {
      t = t.replace(ADDR_TAIL_NAMEPHONE, "").replace(ADDR_TAIL_PHONE, "");
    }
    t = t.replace(/~/g, "-");
    return t.replace(/\s+/g, " ").trim();
  }

  function parseIntSafe(v, def) {
    if (def === undefined) def = 1;
    const t = cleanText(v);
    if (!t) return def;
    const m = t.match(/\d+/);
    return m ? parseInt(m[0], 10) : def;
  }

  function toYYMMDD(v) {
    const t = cleanText(v).replace(/[-.]/g, "");
    const m = t.match(/(?:20)?(\d{2})(\d{2})(\d{2})/);
    return m ? m[1] + m[2] + m[3] : null;
  }

  function normalizeHeader(v) {
    if (v === null || v === undefined) return "";
    return String(v).trim().replace(/\s+/g, "");
  }

  // =====================================================================
  // 한 행 접근자
  // =====================================================================
  function makeRow(values, headerMap) {
    return {
      h: function (header) {
        const idx = headerMap[normalizeHeader(header)];
        return idx ? values[idx - 1] : undefined;
      },
      pos: function (idx) {
        return idx >= 1 && idx <= values.length ? values[idx - 1] : undefined;
      },
    };
  }

  // =====================================================================
  // 업체별 매핑
  // =====================================================================
  function mapUnier(row) {
    const name = cleanText(row.h("수신인"));
    const phone = normalizePhone(row.h("연락처1"));
    let product = cleanText(row.h("상품명"));
    const attr = cleanText(row.h("속성1 속성2"));
    if (attr) product = (product + " " + attr).trim();
    return {
      orderer: name, recipient: name, phone: phone, phone2: phone,
      postcode: cleanPostcode(row.h("우편번호")),
      address: cleanAddress(row.h("주소"), true),
      message: cleanText(row.h("배송메모")),
      product: product, qty: parseIntSafe(row.h("실수량")),
      courier: cleanText(row.h("택배사명")), invoice: cleanText(row.h("송장번호")),
    };
  }

  function mapBlueberry(row) {
    const name = cleanText(row.h("수령인명"));
    const phone = normalizePhone(row.h("수령인연락처"));
    const pq = splitProductQty(row.h("주문상품명") || row.pos(4));
    return {
      orderer: name, recipient: name, phone: phone, phone2: phone,
      postcode: "", address: cleanAddress(row.h("주소")),
      message: cleanText(row.h("배송시 요청사항")),
      product: pq.product, qty: pq.qty, courier: "", invoice: "",
    };
  }

  function mapChikjeup(row) {
    const name = cleanText(row.h("수령인"));
    const phone = normalizePhone(row.h("수령인연락처"));
    const pq = splitProductQty(row.h("주문상품명"));
    return {
      orderer: name, recipient: name, phone: phone, phone2: phone,
      postcode: "", address: cleanAddress(row.h("주소")),
      message: cleanText(row.h("비고")),
      product: pq.product, qty: pq.qty, courier: "", invoice: "",
    };
  }

  function mapFashiongeo(row) {
    const orderer = cleanText(row.h("주문자명"));
    const recipient = cleanText(row.h("수취인 명"));
    const phone = normalizePhone(row.h("수취인 휴대폰번호") || row.h("수취인 전화번호"));
    const pa = extractPostcodeFromAddress(row.h("배송지"));
    return {
      orderer: orderer || recipient, recipient: recipient,
      phone: phone, phone2: phone,
      postcode: pa.postcode, address: pa.address,
      message: cleanText(row.h("배송메세지")),
      product: cleanText(row.h("품목명")), qty: parseIntSafe(row.h("수량")),
      courier: cleanText(row.h("택배사")), invoice: cleanText(row.h("송장번호")),
    };
  }

  function unierName(headerMap) {
    for (const norm in headerMap) {
      if (norm.indexOf("식별번호") !== -1) {
        const m = norm.match(/식별번호[-\s]*([가-힣A-Za-z0-9]+)/);
        if (m) return m[1];
      }
    }
    return null;
  }
  function fashiongeoName(sheetName) {
    const m = (sheetName || "").match(/\d{6,8}_\s*([^_]+?)\s*발주/);
    if (m) return m[1].replace(/\(주\)|주식회사/g, "").trim();
    return null;
  }

  const VENDORS = [
    { key: "unier", label: "유니어", defaultName: "유니어",
      signature: ["수신인", "상품명", "실수량"], mapRow: mapUnier,
      nameFromHeaders: unierName, nameConfident: true },
    { key: "fashiongeo", label: "패션지오", defaultName: "패션지오",
      signature: ["품목명", "수취인 명", "배송지"], mapRow: mapFashiongeo,
      nameFromSheet: fashiongeoName, dateHeader: "주문일자", nameConfident: true },
    { key: "chikjeup", label: "칡즙", defaultName: "칡즙",
      signature: ["수령인", "주문상품명", "비고"], mapRow: mapChikjeup,
      nameConfident: false },
    { key: "blueberry", label: "블루베리퓨레", defaultName: "블루베리퓨레",
      signature: ["수령인명", "수령인연락처", "배송시 요청사항"], mapRow: mapBlueberry,
      nameConfident: false },
  ].map((v) => Object.assign(v, { signatureNorm: v.signature.map(normalizeHeader) }));

  function detectVendor(present) {
    let best = null, bestScore = 0;
    for (const v of VENDORS) {
      if (v.signatureNorm.every((s) => present.has(s)) && v.signatureNorm.length > bestScore) {
        best = v; bestScore = v.signatureNorm.length;
      }
    }
    return best;
  }

  // =====================================================================
  // 워크북 → 표준 행
  // =====================================================================
  function sheetToGrid(ws) {
    const R = ws.rowCount, C = Math.max(ws.columnCount, 1);
    const grid = [];
    for (let r = 1; r <= R; r++) {
      const row = ws.getRow(r);
      const arr = [];
      for (let c = 1; c <= C; c++) arr.push(cellToText(row.getCell(c).value));
      grid.push(arr);
    }
    return grid;
  }

  function findHeaderRow(grid, maxScan) {
    maxScan = Math.min(maxScan || 5, grid.length);
    for (let r = 0; r < maxScan; r++) {
      const present = new Set();
      const headerMap = {};
      grid[r].forEach((cell, i) => {
        const norm = normalizeHeader(cell);
        if (norm) { present.add(norm); if (!(norm in headerMap)) headerMap[norm] = i + 1; }
      });
      const vendor = detectVendor(present);
      if (vendor) return { vendor: vendor, headerRow: r, headerMap: headerMap };
    }
    return null;
  }

  function convertSheet(ws) {
    const grid = sheetToGrid(ws);
    const found = findHeaderRow(grid, 5);
    if (!found) return null;
    const { vendor, headerRow, headerMap } = found;
    const rows = [];
    let orderDate = null;
    for (let r = headerRow + 1; r < grid.length; r++) {
      const values = grid[r];
      if (!values.some((v) => cleanText(v))) continue;
      const row = makeRow(values, headerMap);
      const std = vendor.mapRow(row);
      if (!cleanText(std.recipient) && !cleanText(std.address)) continue;
      std.no = rows.length + 1;
      rows.push(std);
      if (orderDate === null && vendor.dateHeader) orderDate = toYYMMDD(row.h(vendor.dateHeader));
    }
    let vendorName = vendor.defaultName;
    if (vendor.nameFromHeaders) vendorName = vendor.nameFromHeaders(headerMap) || vendorName;
    if (vendor.nameFromSheet) vendorName = vendor.nameFromSheet(ws.name) || vendorName;
    return {
      vendor: vendor, vendorName: vendorName, nameConfident: vendor.nameConfident,
      date: orderDate, rows: rows, sheetName: ws.name,
    };
  }

  async function convertArrayBuffer(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    // 데이터가 있는 시트를 순회하며 인식되는 첫 시트를 사용
    let firstNonEmpty = null;
    for (const ws of wb.worksheets) {
      if (ws.rowCount > 1 && !firstNonEmpty) firstNonEmpty = ws;
      const result = convertSheet(ws);
      if (result) return result;
    }
    const err = new Error("업체 양식을 인식하지 못했습니다.");
    err.headers = firstNonEmpty ? sheetToGrid(firstNonEmpty)[0] : [];
    throw err;
  }

  // =====================================================================
  // 표준 행 → 엑셀 버퍼 (서식 포함)
  // =====================================================================
  async function buildOutputBuffer(rows, sheetTitle) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet((sheetTitle || "발주").slice(0, 31));
    ws.columns = STANDARD_COLUMNS.map((c, i) => ({ key: c.key + i, width: WIDTHS[i] }));
    ws.addRow(HEADERS);
    rows.forEach((r) => ws.addRow(FIELD_KEYS.map((k) => (r[k] === undefined || r[k] === null ? "" : r[k]))));

    const border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    ws.eachRow({ includeEmpty: true }, (row, rn) => {
      if (rn === 1) row.height = 30.75;
      for (let cn = 1; cn <= HEADERS.length; cn++) {
        const cell = row.getCell(cn);
        cell.font = { name: FONT_NAME, size: 10 };
        cell.border = border;
        const center = CENTER_COLS.has(cn);
        cell.alignment = {
          vertical: "middle",
          horizontal: rn === 1 ? (cn === 4 || cn === 5 ? "left" : "center") : (center ? "center" : "left"),
          wrapText: rn === 1,
        };
        if (rn === 1 && HEADER_FILL_COLS.has(cn)) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCE4D6" } };
        }
        if (rn > 1 && TEXT_COLS.has(cn)) cell.numFmt = "@";
      }
    });
    return await wb.xlsx.writeBuffer();
  }

  function outputFilename(date, vendorName) {
    return `${date}_${vendorName}_발주 수정본.xlsx`;
  }

  function todayYYMMDD() {
    const d = new Date();
    return String(d.getFullYear()).slice(2) + String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0");
  }

  // 공개 API (테스트/빌드에서 사용)
  window.POEngine = {
    convertArrayBuffer, buildOutputBuffer, outputFilename, todayYYMMDD,
    STANDARD_COLUMNS, HEADERS, VENDORS,
    _clean: { cleanPostcode, splitProductQty, cleanAddress, extractPostcodeFromAddress, normalizePhone },
  };

  // =====================================================================
  // UI
  // =====================================================================
  if (typeof document === "undefined") return;

  const state = [];   // 처리된 파일들
  let seq = 0;

  const els = {};
  function $(id) { return document.getElementById(id); }

  function download(buffer, filename) {
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1500);
  }

  function esc(s) {
    return String(s === undefined || s === null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  async function handleFiles(fileList) {
    const files = Array.from(fileList).filter((f) => /\.(xlsx|xlsm|xls)$/i.test(f.name));
    if (!files.length) return;
    for (const file of files) {
      const item = { id: ++seq, sourceName: file.name, status: "loading" };
      state.push(item);
      render();
      try {
        const buf = await file.arrayBuffer();
        const result = await window.POEngine.convertArrayBuffer(buf);
        Object.assign(item, {
          status: "ok",
          vendorKey: result.vendor.key,
          formatLabel: result.vendor.label,
          vendorName: result.vendorName,
          nameConfident: result.nameConfident,
          date: result.date || window.POEngine.todayYYMMDD(),
          rows: result.rows,
          sheetName: result.sheetName,
        });
      } catch (e) {
        Object.assign(item, { status: "error", error: e.message, headers: e.headers || [] });
      }
      render();
    }
  }

  async function downloadItem(item) {
    const title = `${item.date}_${item.vendorName}`;
    const buf = await window.POEngine.buildOutputBuffer(item.rows, title);
    download(buf, window.POEngine.outputFilename(item.date, item.vendorName));
  }

  async function downloadAll() {
    const oks = state.filter((i) => i.status === "ok");
    for (const item of oks) { await downloadItem(item); await new Promise((r) => setTimeout(r, 350)); }
  }

  function removeItem(id) {
    const i = state.findIndex((x) => x.id === id);
    if (i !== -1) { state.splice(i, 1); render(); }
  }

  function previewTable(item) {
    const cols = window.POEngine.STANDARD_COLUMNS;
    const shown = item.rows.slice(0, 6);
    let head = cols.map((c) => `<th>${esc(c.header)}</th>`).join("");
    let body = shown.map((r) =>
      "<tr>" + cols.map((c) => `<td>${esc(r[c.key])}</td>`).join("") + "</tr>"
    ).join("");
    const more = item.rows.length > shown.length
      ? `<div class="preview-more">+ ${item.rows.length - shown.length}건 더</div>` : "";
    return `<div class="preview-wrap"><table class="preview"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>${more}`;
  }

  function cardHTML(item) {
    if (item.status === "loading") {
      return `<div class="card loading"><div class="card-head"><span class="fname">${esc(item.sourceName)}</span><span class="chip chip-muted">읽는 중…</span></div></div>`;
    }
    if (item.status === "error") {
      const hs = (item.headers || []).filter(Boolean).slice(0, 12).map(esc).join(" · ");
      return `<div class="card error">
        <div class="card-head">
          <span class="fname">${esc(item.sourceName)}</span>
          <span class="chip chip-error">인식 실패</span>
          <button class="x" data-remove="${item.id}" aria-label="제거">✕</button>
        </div>
        <p class="err-msg">${esc(item.error)} 지원 양식(유니어·패션지오·칡즙·블루베리)의 헤더가 있는지 확인해 주세요.</p>
        ${hs ? `<p class="err-headers"><span>원본 헤더</span>${hs}</p>` : ""}
      </div>`;
    }
    const nameChip = item.nameConfident
      ? `<span class="chip chip-ok">업체명 자동인식</span>`
      : `<span class="chip chip-warn">업체명 확인 필요</span>`;
    return `<div class="card">
      <div class="card-head">
        <span class="fname">${esc(item.sourceName)}</span>
        <span class="chip chip-ok">${esc(item.formatLabel)} 양식</span>
        ${nameChip}
        <span class="count"><b>${item.rows.length}</b>건</span>
        <button class="x" data-remove="${item.id}" aria-label="제거">✕</button>
      </div>
      <div class="fields">
        <label class="field"><span>날짜</span>
          <input type="text" value="${esc(item.date)}" data-field="date" data-id="${item.id}" inputmode="numeric" maxlength="8"></label>
        <label class="field"><span>업체명</span>
          <input type="text" value="${esc(item.vendorName)}" data-field="vendorName" data-id="${item.id}" class="${item.nameConfident ? "" : "needs"}"></label>
        <div class="field grow"><span>출력 파일명</span>
          <div class="outname" id="outname-${item.id}">${esc(window.POEngine.outputFilename(item.date, item.vendorName))}</div></div>
        <button class="btn btn-primary" data-download="${item.id}">다운로드</button>
      </div>
      ${previewTable(item)}
    </div>`;
  }

  function render() {
    const list = $("results");
    const oks = state.filter((i) => i.status === "ok").length;
    els.bulk.hidden = state.filter((i) => i.status === "ok").length < 2;
    els.bulkCount.textContent = oks;
    if (!state.length) {
      list.innerHTML = "";
      els.empty.hidden = false;
      return;
    }
    els.empty.hidden = true;
    list.innerHTML = state.map(cardHTML).join("");

    list.querySelectorAll("[data-download]").forEach((b) =>
      b.addEventListener("click", () => {
        const item = state.find((x) => x.id === +b.dataset.download);
        if (item) downloadItem(item);
      }));
    list.querySelectorAll("[data-remove]").forEach((b) =>
      b.addEventListener("click", () => removeItem(+b.dataset.remove)));
    list.querySelectorAll("input[data-field]").forEach((inp) =>
      inp.addEventListener("input", () => {
        const item = state.find((x) => x.id === +inp.dataset.id);
        if (!item) return;
        item[inp.dataset.field] = inp.value;
        const out = $("outname-" + item.id);
        if (out) out.textContent = window.POEngine.outputFilename(item.date, item.vendorName);
      }));
  }

  function init() {
    els.drop = $("drop");
    els.input = $("file");
    els.empty = $("empty");
    els.bulk = $("bulk");
    els.bulkCount = $("bulk-count");

    els.drop.addEventListener("click", () => els.input.click());
    els.drop.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); els.input.click(); } });
    els.input.addEventListener("change", (e) => { handleFiles(e.target.files); els.input.value = ""; });

    ["dragenter", "dragover"].forEach((ev) =>
      els.drop.addEventListener(ev, (e) => { e.preventDefault(); els.drop.classList.add("over"); }));
    ["dragleave", "drop"].forEach((ev) =>
      els.drop.addEventListener(ev, (e) => { e.preventDefault(); if (ev === "dragleave" && els.drop.contains(e.relatedTarget)) return; els.drop.classList.remove("over"); }));
    els.drop.addEventListener("drop", (e) => { if (e.dataTransfer) handleFiles(e.dataTransfer.files); });

    $("bulk-btn").addEventListener("click", downloadAll);
    render();
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
