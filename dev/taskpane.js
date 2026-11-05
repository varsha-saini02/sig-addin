"use strict";

/* =============================================================
   Word Signature Add-in - DEV build, Step 3c (+ dark mode)
   NCCS CADCAM
   Scan any record book -> choose fields -> draw -> sign.
   Step 3c: next-cell target never uses a cell holding another label;
            signed check matches the role in the SIG tag;
            switching role clears the drawing.
   ============================================================= */

/* ---------- CONFIG: all wording and rules in one place ---------- */
const CONFIG = {
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
    return decodeURIComponent(u.split("/").pop().split(
