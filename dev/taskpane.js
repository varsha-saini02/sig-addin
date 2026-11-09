"use strict";

/* =============================================================
   Word Signature Add-in - DEV build, Step 3d (+ dark mode)
   NCCS CADCAM
   Scan any record book -> choose fields -> draw -> sign.
   Step 3d: a next cell is used only if it is at least MIN_CELL_W_PT
            wide (narrow spacer cells fall back to label end);
            next cell never holds another label; signed check matches
            the role in the SIG tag; switching role clears the drawing.
   ============================================================= */

/* ---------- CONFIG: all wording and rules in one place ---------- */
const CONFIG = {
  BUILD: "step3d",       // shown at the end of the Developer box first line
  MAX_LINE: 60,          // longer lines are never labels or headings
  MAX_PLAIN_SIG: 25,     // a role-less "Signature" line must be this short to be guessed
  LOOKBACK: 8,           // non-empty lines to look back for a role word
  SIG_RE: /\bsign(ature|ed)?\b/,
  ROLE_WORDS: {
    student:  ["student", "learner", "candidate", "trainee", "apprentice"],
    assessor: ["assessor", "trainer", "teacher", "verifier", "supervisor"]
  },
  SECTION_RE: /^(?:(?:assessment\s+)?task\s*(\d+)|unit (?:assessment )?summary)/,
  DEFAULT_SECTION: "Receipt of Unit Guide",
  REASSESS_ON_RE: /\bre-?\s?assess/,
  REASSESS_ON_EXCEPT_RE: /^date for/,
  REASSESS_OFF_RE: /^assessment task outcome|^(?:assessment\s+)?task\s*\d+|acknowledg|unit assessment summary/,
  // Our own "email | date" text. Removed before reading a line, so a
  // signed label still looks the same to the scan.
  STAMP_RE: /\s*[^\s@|]+@[^\s@|]+\s*\|\s*\d{1,2}\/\d{1,2}\/\d{2,4}/g,
  SIG_TAG: "SIG|",
  SIG_MAX_W_PT: 120,     // inserted signature size limits (points)
  SIG_MAX_H_PT: 30,
  MIN_CELL_W_PT: 80,     // a next cell narrower than this is a spacer, not a signing space
  SIG_INK: "#1a1a1a",    // ink colour (canvas is always white)
  EMAIL_KEY: "nccsSig.email",
  THEME_KEY: "nccsSig.theme"
};

const EMAIL_OK = /^[^\s@|]+@[^\s@|]+\.[^\s@|]+$/;

const ROLE_RE = {};
for (const r in CONFIG.ROLE_WORDS) {
  ROLE_RE[r] = new RegExp("\\b(?:" + CONFIG.ROLE_WORDS[r].join("|") + ")s?\\b");
}

/* ---------- state ---------- */
const $ = id => document.getElementById(id);
let role = "student";
let scanState = null;          // { fields: [plain objects], paraCount }
let checked = new Map();       // field key -> ticked?
let hasSignature = false;
let busy = false;
let canvas, ctx;
let devText = "", devErr = "";

/* =============================================================
   THEME (light / dark)
   Auto = Word's theme if Word reports one, else the system setting.
   The Theme button overrides: Auto -> Light -> Dark -> Auto.
   ============================================================= */
function getThemePref() {
  try { return localStorage.getItem(CONFIG.THEME_KEY) || "auto"; } catch (e) { return "auto"; }
}

function setThemePref(p) {
  try { localStorage.setItem(CONFIG.THEME_KEY, p); } catch (e) { /* storage blocked */ }
}

function brightness(hex) {
  const m = /^#?([0-9a-f]{6})$/i.exec(String(hex || "").trim());
  if (!m) return null;
  const n = parseInt(m[1], 16);
  return (0.299 * ((n >> 16) & 255) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255;
}

function officeThemeIsDark() {
  try {
    const t = Office.context && Office.context.officeTheme;
    if (t && t.bodyBackgroundColor) {
      const b = brightness(t.bodyBackgroundColor);
      if (b !== null) return b < 0.5;
    }
  } catch (e) { /* not supported here */ }
  return null;
}

function systemIsDark() {
  try { return window.matchMedia("(prefers-color-scheme: dark)").matches; } catch (e) { return false; }
}

function applyTheme() {
  const pref = getThemePref();
  let dark;
  if (pref === "dark") dark = true;
  else if (pref === "light") dark = false;
  else {
    const o = officeThemeIsDark();
    dark = (o !== null) ? o : systemIsDark();
  }
  document.documentElement.setAttribute("data-theme", dark ? "dark" : "light");
  const btn = $("themeBtn");
  if (btn) btn.textContent = "Theme: " + pref.charAt(0).toUpperCase() + pref.slice(1);
}

function initTheme() {
  applyTheme();
  $("themeBtn").addEventListener("click", () => {
    const order = ["auto", "light", "dark"];
    const next = order[(order.indexOf(getThemePref()) + 1) % order.length];
    setThemePref(next);
    applyTheme();
  });
  // Follow system changes while in Auto
  try {
    const mq = window.matchMedia("(prefers-color-scheme: dark)");
    if (mq.addEventListener) mq.addEventListener("change", applyTheme);
    else if (mq.addListener) mq.addListener(applyTheme);
  } catch (e) { /* ignore */ }
  // Follow Office theme changes where Office supports the event
  try {
    if (Office.EventType && Office.EventType.OfficeThemeChanged) {
      Office.context.document.addHandlerAsync(Office.EventType.OfficeThemeChanged, applyTheme, () => {});
    }
  } catch (e) { /* not supported here */ }
}

/* ---------- text helpers ---------- */
function norm(t) {
  return (t || "")
    .replace(/[\u0000-\u001f\u200b\u200e\u200f\ufeff\ufffc]/g, " ")
    .replace(/[\u00a0\u2000-\u200a\u202f\u205f\u3000]/g, " ")
    .replace(/[\u2010-\u2015\u2212]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function lineOf(text) {
  return norm(text).replace(CONFIG.STAMP_RE, " ").replace(/\s+/g, " ").trim();
}

function cleanDisplay(text) {
  return (text || "")
    .replace(/[\u0000-\u001f\u200b\u200e\u200f\ufeff\ufffc]/g, " ")
    .replace(/[\u00a0\u2000-\u200a\u202f\u205f\u3000]/g, " ")
    .replace(CONFIG.STAMP_RE, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function roleOf(line) {
  const s = ROLE_RE.student.test(line);
  const a = ROLE_RE.assessor.test(line);
  if (s && a) return "both";
  return s ? "student" : (a ? "assessor" : null);
}

function lookBackRole(lines, i) {
  let seen = 0;
  for (let j = i - 1; j >= 0 && seen < CONFIG.LOOKBACK; j--) {
    const t = lines[j];
    if (!t) continue;
    seen++;
    if (t.length > CONFIG.MAX_LINE) continue;
    const r = roleOf(t);
    if (r === "student" || r === "assessor") return { role: r, j: j };
  }
  return null;
}

function isSigPic(p) {
  return (p.altTextDescription || "").indexOf(CONFIG.SIG_TAG) === 0;
}

// Role stored in our picture tag: "SIG|role|email|time"
function sigRoleOf(p) {
  return (p.altTextDescription || "").split("|")[1] || "";
}

// Our signature for this exact role
function isOurSigFor(p, r) {
  return isSigPic(p) && sigRoleOf(p) === r;
}

// Text that is (or contains) a label: a signature word or a role word.
// Such a cell belongs to another field and is never a signing target.
function looksLikeLabel(line) {
  return CONFIG.SIG_RE.test(line) || roleOf(line) !== null;
}

function roleName(r) { return (r || role) === "student" ? "Student" : "Assessor"; }
function pad(s, n) { s = String(s); return s.length >= n ? s : s + " ".repeat(n - s.length); }
function clip(s, n) { s = String(s); return s.length > n ? s.slice(0, n - 3) + "..." : s; }

function docName() {
  try {
    const u = Office.context.document.url || "";
    return decodeURIComponent(u.split("/").pop().split("?")[0]) || "(unknown)";
  } catch (e) { return "(unknown)"; }
}

function userError(msg) { const e = new Error(msg); e.userMessage = msg; return e; }

/* =============================================================
   scanFields(context) -> { fields (with _para/_cell proxies), paraCount }
   ============================================================= */
async function scanFields(context) {
  const paras = context.document.body.paragraphs;   // includes table cells
  paras.load("items/text");
  await context.sync();

  const items = paras.items;
  const lines = items.map(p => lineOf(p.text));
  const fields = [];
  let section = CONFIG.DEFAULT_SECTION;
  let reassess = false;

  /* Pass 1: linear walk - sections, re-assessment state, fields */
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!line || line.length > CONFIG.MAX_LINE) continue;

    const sm = line.match(CONFIG.SECTION_RE);
    if (sm) section = sm[1] ? "Assessment Task " + sm[1] : "Unit Summary";

    if (CONFIG.REASSESS_ON_RE.test(line) && !CONFIG.REASSESS_ON_EXCEPT_RE.test(line)) {
      reassess = true;
    } else if (CONFIG.REASSESS_OFF_RE.test(line)) {
      reassess = false;
    }

    if (!CONFIG.SIG_RE.test(line)) continue;

    let r = roleOf(line), guessed = false, guessFrom = "";
    if (r === "both") continue;
    if (!r) {
      if (line.length > CONFIG.MAX_PLAIN_SIG) continue;
      const lb = lookBackRole(lines, i);
      if (!lb) continue;
      r = lb.role; guessed = true; guessFrom = cleanDisplay(items[lb.j].text);
    }
    fields.push({
      key: i + "|" + r + "|" + line,
      index: i, label: line, text: cleanDisplay(items[i].text) || line,
      role: r, section: section, isReassessment: reassess,
      guessed: guessed, guessFrom: guessFrom
    });
  }

  if (!fields.length) return { fields: fields, paraCount: lines.length };

  /* Pass 2: is the label in a table cell? */
  const cells = fields.map(f => {
    const c = items[f.index].parentTableCellOrNullObject;
    c.load("rowIndex,cellIndex");
    return c;
  });
  await context.sync();

  /* Pass 3: next cell in the same row */
  const nexts = cells.map(c => {
    if (c.isNullObject) return null;
    const n = c.getNextOrNullObject();
    n.load("rowIndex,cellIndex,value,columnWidth");
    return n;
  });
  await context.sync();

  /* Pass 4: pictures at label and next cell */
  const labelPics = [], cellPics = [], nextCell = [];
  fields.forEach((f, k) => {
    const c = cells[k], n = nexts[k];
    nextCell[k] = (n && !n.isNullObject && n.rowIndex === c.rowIndex && n.cellIndex > c.cellIndex) ? n : null;
    labelPics[k] = items[f.index].inlinePictures;
    labelPics[k].load("items/altTextDescription");
    if (nextCell[k]) {
      cellPics[k] = nextCell[k].body.inlinePictures;
      cellPics[k].load("items/altTextDescription");
    }
  });
  await context.sync();

  /* Decide target (step3d):
     next cell ONLY if it is wide enough (MIN_CELL_W_PT), holds no label
     text, no other-role signature, and is either empty or holds our
     signature for THIS role. Otherwise sign at the end of the label
     paragraph. If Word does not report a width, the width test is
     skipped (shown as "-" in the Developer box). */
  fields.forEach((f, k) => {
    let pics = labelPics[k].items;
    f.target = "label end";
    f._para = items[f.index];
    f._cell = null;
    if (nextCell[k]) {
      const cp = cellPics[k].items;
      const w = nextCell[k].columnWidth;
      f.nextW = w ? Math.round(w) : 0;
      const wideEnough = !f.nextW || f.nextW >= CONFIG.MIN_CELL_W_PT;
      const cellText = lineOf(nextCell[k].value);
      const isLabel = looksLikeLabel(cellText);
      const otherSig = cp.some(p => isSigPic(p) && sigRoleOf(p) !== f.role);
      const ourSig = cp.some(p => isOurSigFor(p, f.role));
      if (wideEnough && !isLabel && !otherSig && (cellText === "" || ourSig)) {
        f.target = "next cell"; f._cell = nextCell[k]; pics = cp;
      }
    }
    const sig = pics.find(p => isOurSigFor(p, f.role));
    f.signed = !!sig;
    f.signedBy = sig ? (sig.altTextDescription.split("|")[2] || "") : "";
    f.otherPics = pics.filter(p => !isSigPic(p)).length;
    f.ticked = !f.isReassessment && !f.guessed && !f.signed;
  });

  return { fields: fields, paraCount: lines.length };
}

function plainFields(fields) {
  return fields.map(f => {
    const o = {};
    for (const k in f) if (k.charAt(0) !== "_") o[k] = f[k];
    return o;
  });
}

/* ---------- scan + list ---------- */
async function runScan(quiet) {
  const t0 = Date.now();
  await Word.run(async context => {
    const r = await scanFields(context);
    const fields = plainFields(r.fields);
    const prev = checked;
    checked = new Map();
    fields.forEach(f => {
      checked.set(f.key, f.signed ? false : (prev.has(f.key) ? prev.get(f.key) : f.ticked));
    });
    scanState = { fields: fields, paraCount: r.paraCount };
  });
  devReport(Date.now() - t0);
  $("scanBtn").textContent = "Refresh list";
  renderList();
  updateSignBtn();
  if (!quiet) clearStatus();
}

function roleFields() {
  return scanState ? scanState.fields.filter(f => f.role === role) : [];
}

function selectedFields() {
  return roleFields().filter(f => !f.signed && checked.get(f.key));
}

function renderList() {
  const box = $("fieldList");
  box.innerHTML = "";
  if (!scanState) return;

  const fields = roleFields();
  if (!fields.length) {
    const p = document.createElement("p");
    p.className = "empty";
    p.textContent = "No " + roleName() + " signature fields found in this document. " +
      "Check you have the right document open, then click Refresh list.";
    box.appendChild(p);
    return;
  }

  const nSigned = fields.filter(f => f.signed).length;
  const summary = document.createElement("p");
  summary.className = "list-summary";
  summary.textContent = fields.length + " " + roleName() + " field" + (fields.length > 1 ? "s" : "") +
    " found" + (nSigned ? ", " + nSigned + " already signed" : "") + ". Ticked fields will be signed.";
  box.appendChild(summary);

  const groups = new Map();
  fields.forEach(f => {
    if (!groups.has(f.section)) groups.set(f.section, []);
    groups.get(f.section).push(f);
  });

  groups.forEach((list, section) => {
    const h = document.createElement("div");
    h.className = "section-title";
    h.textContent = section;
    box.appendChild(h);

    list.forEach(f => {
      const row = document.createElement("label");
      row.className = "field" + (f.signed ? " done" : "");

      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = !!checked.get(f.key);
      cb.disabled = f.signed;
      cb.addEventListener("change", () => { checked.set(f.key, cb.checked); updateSignBtn(); });
      row.appendChild(cb);

      const txt = document.createElement("span");
      txt.className = "txt";
      txt.appendChild(document.createTextNode(clip(f.text, 60)));
      if (f.isReassessment) txt.appendChild(badge("Re-assessment", "re"));
      if (f.guessed && !f.signed) txt.appendChild(badge("Check: is this yours?", "guess"));
      if (f.signed) txt.appendChild(badge("Signed", "done"));
      if (f.guessed) txt.appendChild(sub("Near: " + clip(f.guessFrom, 50)));
      if (f.signed && f.signedBy) txt.appendChild(sub("Signed by " + f.signedBy));
      row.appendChild(txt);

      const go = document.createElement("button");
      go.className = "link-btn";
      go.textContent = "Show";
      go.title = "Show this spot in the document";
      go.addEventListener("click", e => { e.preventDefault(); guarded(() => goTo(f)); });
      row.appendChild(go);

      box.appendChild(row);
    });
  });
}

function badge(text, cls) {
  const b = document.createElement("span");
  b.className = "badge " + cls;
  b.textContent = text;
  return b;
}

function sub(text) {
  const s = document.createElement("span");
  s.className = "sub";
  s.textContent = text;
  return s;
}

async function goTo(f) {
  await Word.run(async context => {
    const paras = context.document.body.paragraphs;
    paras.load("items/text");
    await context.sync();
    const p = paras.items[f.index];
    if (!p || lineOf(p.text) !== f.label) {
      throw userError("That spot has moved. Click Refresh list.");
    }
    p.select();
    await context.sync();
  });
}

/* ---------- signature image ---------- */
function getSignatureImage() {
  const w = canvas.width, h = canvas.height;
  if (!w || !h) return null;
  const data = ctx.getImageData(0, 0, w, h).data;
  let minX = w, minY = h, maxX = -1, maxY = -1;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      if (data[(y * w + x) * 4 + 3] > 0) {
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
        if (y < minY) minY = y;
        if (y > maxY) maxY = y;
      }
    }
  }
  if (maxX < 0) return null;

  const padPx = 6;
  minX = Math.max(0, minX - padPx); minY = Math.max(0, minY - padPx);
  maxX = Math.min(w - 1, maxX + padPx); maxY = Math.min(h - 1, maxY + padPx);
  const cw = maxX - minX + 1, ch = maxY - minY + 1;

  const out = document.createElement("canvas");
  out.width = cw; out.height = ch;
  const o = out.getContext("2d");
  // White background: Word's dark mode does not recolour pictures, so a
  // see-through signature would be almost invisible on a dark page view.
  o.fillStyle = "#ffffff";
  o.fillRect(0, 0, cw, ch);
  o.drawImage(canvas, minX, minY, cw, ch, 0, 0, cw, ch);

  const ratio = cw / ch;
  let hPt = CONFIG.SIG_MAX_H_PT, wPt = hPt * ratio;
  if (wPt > CONFIG.SIG_MAX_W_PT) { wPt = CONFIG.SIG_MAX_W_PT; hPt = wPt / ratio; }

  return {
    b64: out.toDataURL("image/png").split(",")[1],   // no data: prefix
    wPt: Math.round(wPt * 10) / 10,
    hPt: Math.round(hPt * 10) / 10
  };
}

/* ---------- sign ---------- */
async function doSign() {
  const email = $("emailInput").value.trim();
  if (!EMAIL_OK.test(email)) { showStatus("Please enter a valid email address.", "error"); return; }
  if (!hasSignature) { showStatus("Please draw your signature first.", "error"); return; }

  const chosen = selectedFields();
  if (!chosen.length) { showStatus("Tick at least one field to sign.", "error"); return; }

  const img = getSignatureImage();
  if (!img) { showStatus("Please draw your signature first.", "error"); return; }

  const now = new Date();
  const dateStr = now.toLocaleDateString("en-AU", { day: "2-digit", month: "2-digit", year: "numeric" });
  const stamp = "  " + email + " | " + dateStr;
  const tag = CONFIG.SIG_TAG + role + "|" + email + "|" + now.toISOString();
  let done = 0, skipped = 0;

  showStatus("Signing " + chosen.length + " field" + (chosen.length > 1 ? "s" : "") + "...", "info");

  await Word.run(async context => {
    const fresh = await scanFields(context);
    const byKey = new Map(fresh.fields.map(f => [f.key, f]));
    const targets = [];

    for (const c of chosen) {
      const f = byKey.get(c.key);
      if (!f) {
        throw userError("The document has changed since the list was made. " +
          "Click Refresh list, check your ticks, then sign again. Nothing was signed.");
      }
      if (f.signed) { skipped++; continue; }
      targets.push(f);
    }

    for (const f of targets) {
      let pic;
      if (f._cell) {
        pic = f._cell.body.insertInlinePictureFromBase64(img.b64, "End");
      } else {
        f._para.insertText(" ", "End");
        pic = f._para.insertInlinePictureFromBase64(img.b64, "End");
      }
      pic.width = img.wPt;
      pic.height = img.hPt;
      pic.altTextDescription = tag;
      pic.altTextTitle = "SIG";
      pic.insertText(stamp, "After");
    }
    await context.sync();
    done = targets.length;
  });

  try { localStorage.setItem(CONFIG.EMAIL_KEY, email); } catch (e) { /* storage blocked */ }
  await runScan(true);

  let msg = "Signed " + done + " " + roleName().toLowerCase() + " field" + (done !== 1 ? "s" : "") + ".";
  if (skipped) msg += " " + skipped + " already signed, skipped.";
  showStatus(msg, "success");
}

/* ---------- sign button state ---------- */
function updateSignBtn() {
  const n = selectedFields().length;
  const emailOk = EMAIL_OK.test($("emailInput").value.trim());
  const btn = $("signBtn");
  btn.className = "btn btn-insert" + (role === "assessor" ? " assessor" : "");
  btn.textContent = n ? "Sign " + n + " field" + (n > 1 ? "s" : "") + " as " + roleName() : "Sign selected fields";
  btn.disabled = busy || !n || !hasSignature || !emailOk;

  let hint = "";
  if (!emailOk) hint = "Enter your email.";
  else if (!hasSignature) hint = "Draw your signature.";
  else if (!n) hint = "Tick at least one field.";
  $("signHint").textContent = hint;
}

/* ---------- canvas (always white paper, dark ink) ---------- */
function setupPen() {
  ctx.strokeStyle = CONFIG.SIG_INK;
  ctx.fillStyle = CONFIG.SIG_INK;
  ctx.lineWidth = 3;
  ctx.lineCap = "round";
  ctx.lineJoin = "round";
}

function sizeCanvas() {
  const rect = canvas.getBoundingClientRect();
  if (!rect.width || !rect.height) return;
  const dpr = window.devicePixelRatio || 1;
  let saved = null;
  if (hasSignature && canvas.width && canvas.height) {
    saved = document.createElement("canvas");
    saved.width = canvas.width; saved.height = canvas.height;
    saved.getContext("2d").drawImage(canvas, 0, 0);
  }
  canvas.width = Math.round(rect.width * dpr);
  canvas.height = Math.round(rect.height * dpr);
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  setupPen();
  if (saved) ctx.drawImage(saved, 0, 0, saved.width / dpr, saved.height / dpr);
}

function initCanvas() {
  canvas = $("sigCanvas");
  ctx = canvas.getContext("2d");
  sizeCanvas();

  let drawing = false, lastX = 0, lastY = 0;
  const pos = e => {
    const r = canvas.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  };
  const markDrawn = () => {
    if (!hasSignature) {
      hasSignature = true;
      canvas.classList.add("active");
      updateSignBtn();
    }
  };

  canvas.addEventListener("pointerdown", e => {
    e.preventDefault();
    drawing = true;
    try { canvas.setPointerCapture(e.pointerId); } catch (err) { /* ignore */ }
    const p = pos(e);
    lastX = p.x; lastY = p.y;
    ctx.beginPath();
    ctx.arc(p.x, p.y, ctx.lineWidth / 2, 0, Math.PI * 2);
    ctx.fill();
    markDrawn();
  });
  canvas.addEventListener("pointermove", e => {
    if (!drawing) return;
    e.preventDefault();
    const p = pos(e);
    ctx.beginPath();
    ctx.moveTo(lastX, lastY);
    ctx.lineTo(p.x, p.y);
    ctx.stroke();
    lastX = p.x; lastY = p.y;
  });
  ["pointerup", "pointercancel"].forEach(t =>
    canvas.addEventListener(t, () => { drawing = false; }));

  let timer = null;
  window.addEventListener("resize", () => {
    clearTimeout(timer);
    timer = setTimeout(sizeCanvas, 150);
  });
}

function clearCanvas() {
  ctx.save();
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.restore();
  canvas.classList.remove("active");
  hasSignature = false;
  updateSignBtn();
  clearStatus();
}

/* ---------- toggle + email ---------- */
function showRoleBadge() {
  const badgeEl = $("modeBadge");
  badgeEl.textContent = roleName();
  badgeEl.className = "mode-badge " + role;
}

function initToggle() {
  const t = $("modeToggle");
  t.checked = false;             // always open on Student
  role = "student";
  showRoleBadge();
  t.addEventListener("change", function () {
    role = this.checked ? "assessor" : "student";
    showRoleBadge();
    clearCanvas();               // each role draws its own signature
    renderList();
    updateSignBtn();
    clearStatus();
  });
}

function initEmail() {
  const inp = $("emailInput");
  try { inp.value = localStorage.getItem(CONFIG.EMAIL_KEY) || ""; } catch (e) { /* storage blocked */ }
  inp.addEventListener("input", () => {
    const v = inp.value.trim();
    if (EMAIL_OK.test(v)) {
      try { localStorage.setItem(CONFIG.EMAIL_KEY, v); } catch (e) { /* storage blocked */ }
    }
    updateSignBtn();
  });
}

/* ---------- status, busy, errors ---------- */
function showStatus(msg, type) {
  const el = $("statusMsg");
  el.textContent = msg;
  el.className = "status " + type;
}

function clearStatus() { $("statusMsg").className = "status"; }

function setBusy(b) {
  busy = b;
  $("scanBtn").disabled = b;
  $("clearBtn").disabled = b;
  updateSignBtn();
}

async function guarded(fn) {
  if (busy) return;
  setBusy(true);
  try {
    await fn();
  } catch (e) {
    showStatus(e && e.userMessage ? e.userMessage :
      "Something went wrong: " + ((e && e.message) || e) +
      ". If the document says Viewing at the top, switch to Editing and try again.", "error");
    devError(e);
  } finally {
    setBusy(false);
  }
}

/* ---------- developer box ---------- */
function summarise(fields) {
  const s = { student: 0, assessor: 0, reStudent: 0, reAssessor: 0, guessed: 0, signed: 0 };
  for (const f of fields) {
    s[f.role]++;
    if (f.isReassessment) s[f.role === "student" ? "reStudent" : "reAssessor"]++;
    if (f.guessed) s.guessed++;
    if (f.signed) s.signed++;
  }
  return s;
}

function devReport(ms) {
  const f = scanState.fields;
  const s = summarise(f);
  const L = [];
  L.push("SCAN  " + new Date().toLocaleString() + "  " + docName() + "  [" + CONFIG.BUILD + "]");
  L.push("Theme " + document.documentElement.getAttribute("data-theme") +
         " (pref " + getThemePref() + ", office " + officeThemeIsDark() + ", system " + systemIsDark() + ")");
  L.push("Paragraphs " + scanState.paraCount + "   Fields " + f.length + "   (" + ms + " ms)");
  L.push("Student " + s.student + " (reassess " + s.reStudent + ")   Assessor " + s.assessor +
         " (reassess " + s.reAssessor + ")   guessed " + s.guessed + "   signed " + s.signed);
  L.push("");
  L.push("flags: R=reassess G=guessed S=signed P=other picture   width = next cell width in pt");
  f.forEach((x, n) => {
    const flags = (x.isReassessment ? "R" : "-") + (x.guessed ? "G" : "-") +
                  (x.signed ? "S" : "-") + (x.otherPics ? "P" : "-");
    L.push(pad(n + 1, 3) + (checked.get(x.key) ? "[x] " : "[ ] ") + "p" + pad(x.index, 5) +
           pad(x.section, 22) + pad(x.role.toUpperCase(), 9) + pad(x.target, 10) +
           pad(x.nextW ? x.nextW + "pt" : "-", 7) +
           flags + '  "' + clip(x.label, 40) + '"');
  });
  devText = L.join("\n");
  $("devOut").value = devText + devErr;
}

function devError(e) {
  devErr += "\n\nERROR " + new Date().toLocaleTimeString() + ": " + (e.code || "") + " " + (e.message || e);
  if (e.debugInfo) devErr += "\n" + JSON.stringify(e.debugInfo, null, 1);
  $("devOut").value = devText + devErr;
}

async function copyDev() {
  const ta = $("devOut");
  try { await navigator.clipboard.writeText(ta.value); }
  catch (e) { ta.focus(); ta.select(); try { document.execCommand("copy"); } catch (e2) { /* ignore */ } }
}

/* ---------- start ---------- */
Office.onReady(info => {
  initTheme();
  if (info.host !== Office.HostType.Word) {
    showStatus("This add-in only runs in Word.", "error");
    return;
  }
  initToggle();
  initEmail();
  initCanvas();
  $("clearBtn").addEventListener("click", clearCanvas);
  $("scanBtn").addEventListener("click", () => guarded(() => runScan(false)));
  $("signBtn").addEventListener("click", () => guarded(doSign));
  $("copyDevBtn").addEventListener("click", copyDev);
  updateSignBtn();
  guarded(() => runScan(true));     // list appears as soon as the panel opens
});
