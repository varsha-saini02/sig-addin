"use strict";

/* =============================================================
   Word Signature Add-in - DEV build, Step 2
   - Scan test: prints detected fields as plain text
   - Alt text test: insert tagged picture, reopen doc, read back
   Nothing here signs the document.
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
  // Anchored at the start for BOTH alternatives, so "see the Unit Assessment
  // Summary page" in a short sentence does not switch the section.
  SECTION_RE: /^(?:(?:assessment\s+)?task\s*(\d+)|unit (?:assessment )?summary)/,
  DEFAULT_SECTION: "Receipt of Unit Guide",
  // \b stops "are assessed" from matching "re assess"
  REASSESS_ON_RE: /\bre-?\s?assess/,
  REASSESS_ON_EXCEPT_RE: /^date for/,
  REASSESS_OFF_RE: /^assessment task outcome|^(?:assessment\s+)?task\s*\d+|acknowledg|unit assessment summary/,
  SIG_TAG: "SIG|"
};

const EXPECTED = {
  MEM09009: { student: 8, assessor: 7, reStudent: 3, reAssessor: 3, guessed: 1 },
  MEM11011: { student: 4, assessor: 3, reStudent: 0, reAssessor: 0, guessed: 1 },
  MEM07005: { student: 9, assessor: 8, reStudent: 2, reAssessor: 2, guessed: 1 }
};

const ROLE_RE = {};
for (const role in CONFIG.ROLE_WORDS) {
  ROLE_RE[role] = new RegExp("\\b(?:" + CONFIG.ROLE_WORDS[role].join("|") + ")s?\\b");
}

/* ---------- text helpers ---------- */
function norm(t) {
  return (t || "")
    .replace(/[\u0000-\u001f\u200b\u200e\u200f\ufeff]/g, " ") // control chars, zero-width
    .replace(/[\u00a0\u2000-\u200a\u202f\u205f\u3000]/g, " ")  // non-breaking and odd spaces
    .replace(/[\u2010-\u2015\u2212]/g, "-")                     // non-breaking/fancy hyphens
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
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
    if (t.length > CONFIG.MAX_LINE) continue;   // ignore long declarations
    const r = roleOf(t);
    if (r === "student" || r === "assessor") return { role: r, from: t };
  }
  return null;
}

function isSigPic(p) {
  return (p.altTextDescription || "").indexOf(CONFIG.SIG_TAG) === 0;
}

/* ---------- output helpers ---------- */
let OUT = [];
function out(s) { OUT.push(s === undefined ? "" : String(s)); }
function flush() { document.getElementById("out").value = OUT.join("\n"); }
function pad(s, n) { s = String(s); return s.length >= n ? s : s + " ".repeat(n - s.length); }
function clip(s, n) { s = String(s); return s.length > n ? s.slice(0, n - 3) + "..." : s; }
function tl(i, what, line) { return "p" + pad(i, 5) + pad(what, 30) + '"' + clip(line, 48) + '"'; }
function status(s) { document.getElementById("status").textContent = s; }

function docName() {
  try {
    const u = Office.context.document.url || "";
    return decodeURIComponent(u.split("/").pop().split("?")[0]) || "(unknown)";
  } catch (e) { return "(unknown)"; }
}

function reportError(e) {
  out("");
  out("ERROR: " + (e.code || "") + " " + (e.message || e));
  if (e.debugInfo) out(JSON.stringify(e.debugInfo, null, 1));
  out("(If inserting: is the document in Editing mode, not Viewing?)");
}

/* =============================================================
   scanFields(context, trace) -> { fields, paraCount }
   ============================================================= */
async function scanFields(context, trace) {
  const paras = context.document.body.paragraphs;   // includes table cells
  paras.load("items/text");
  await context.sync();

  const items = paras.items;
  const lines = items.map(p => norm(p.text));
  const fields = [];
  let section = CONFIG.DEFAULT_SECTION;
  let reassess = false;

  /* Pass 1: linear walk - sections, re-assessment state, fields */
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!line || line.length > CONFIG.MAX_LINE) continue;

    const sm = line.match(CONFIG.SECTION_RE);
    if (sm) {
      const s = sm[1] ? "Assessment Task " + sm[1] : "Unit Summary";
      if (s !== section) { section = s; trace.push(tl(i, "SECTION -> " + s, line)); }
    }

    // ON is checked BEFORE OFF, and OFF only runs if ON did not match
    if (CONFIG.REASSESS_ON_RE.test(line) && !CONFIG.REASSESS_ON_EXCEPT_RE.test(line)) {
      if (!reassess) trace.push(tl(i, "REASSESS ON", line));
      reassess = true;
    } else if (CONFIG.REASSESS_OFF_RE.test(line)) {
      if (reassess) trace.push(tl(i, "REASSESS OFF", line));
      reassess = false;
    }

    if (!CONFIG.SIG_RE.test(line)) continue;

    let role = roleOf(line), guessed = false, guessFrom = "";
    if (role === "both") { trace.push(tl(i, "SKIP (both roles)", line)); continue; }
    if (!role) {
      if (line.length > CONFIG.MAX_PLAIN_SIG) { trace.push(tl(i, "SKIP (no role, too long)", line)); continue; }
      const lb = lookBackRole(lines, i);
      if (!lb) { trace.push(tl(i, "SKIP (no role nearby)", line)); continue; }
      role = lb.role; guessed = true; guessFrom = lb.from;
    }
    fields.push({ index: i, label: line, role, section,
                  isReassessment: reassess, guessed, guessFrom });
  }

  if (!fields.length) return { fields, paraCount: lines.length };

  /* Pass 2: is the label in a table cell? */
  const cells = fields.map(f => {
    const c = items[f.index].parentTableCellOrNullObject;
    c.load("rowIndex,cellIndex");
    return c;
  });
  await context.sync();

  /* Pass 3: the next cell (must be in the same row) */
  const nexts = cells.map(c => {
    if (c.isNullObject) return null;
    const n = c.getNextOrNullObject();
    n.load("rowIndex,cellIndex,value");
    return n;
  });
  await context.sync();

  /* Pass 4: pictures at label and next cell (for already-signed check) */
  const labelPics = [], cellPics = [], nextCell = [];
  fields.forEach((f, k) => {
    const c = cells[k], n = nexts[k];
    f.inTable = !c.isNullObject;
    nextCell[k] = (n && !n.isNullObject && n.rowIndex === c.rowIndex && n.cellIndex > c.cellIndex) ? n : null;
    labelPics[k] = items[f.index].inlinePictures;
    labelPics[k].load("items/altTextDescription");
    if (nextCell[k]) {
      cellPics[k] = nextCell[k].body.inlinePictures;
      cellPics[k].load("items/altTextDescription");
    }
  });
  await context.sync();

  /* Decide sign target. Next cell counts as the target if it is empty OR
     already holds an add-in signature (otherwise signed cells would fall
     back to the label and duplicate checks would miss them). */
  fields.forEach((f, k) => {
    let pics = labelPics[k].items;
    f.target = "label end";
    if (nextCell[k]) {
      const cp = cellPics[k].items;
      const empty = norm(nextCell[k].value) === "";
      if (empty || cp.some(isSigPic)) { f.target = "next cell"; pics = cp; }
    }
    f.signed = pics.some(isSigPic);
    f.otherPics = pics.filter(p => !isSigPic(p)).length;
    f.ticked = !f.isReassessment && !f.guessed && !f.signed;
  });

  return { fields, paraCount: lines.length };
}

/* ---------- summary / compare ---------- */
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

/* ---------- button: Scan ---------- */
async function runScan() {
  OUT = [];
  const showTrace = document.getElementById("trace").checked;
  const book = document.getElementById("expected").value;
  const trace = [];
  const t0 = Date.now();

  out("SCAN  " + new Date().toLocaleString());
  out("Document: " + docName());
  out("WordApi 1.3 supported: " + Office.context.requirements.isSetSupported("WordApi", "1.3"));

  await Word.run(async context => {
    const r = await scanFields(context, trace);
    const fields = r.fields;
    out("Paragraphs: " + r.paraCount + "   Fields: " + fields.length + "   (" + (Date.now() - t0) + " ms)");

    if (showTrace) {
      out("");
      out("--- TRACE ---");
      trace.forEach(t => out(t));
    }

    out("");
    out("--- FIELDS ---  [x]=ticked by default   flags: R=reassess G=guessed S=signed P=other picture");
    fields.forEach((f, n) => {
      const flags = (f.isReassessment ? "R" : "-") + (f.guessed ? "G" : "-") +
                    (f.signed ? "S" : "-") + (f.otherPics ? "P" : "-");
      out(pad(n + 1, 3) + (f.ticked ? "[x] " : "[ ] ") + "p" + pad(f.index, 5) +
          pad(f.section, 22) + pad(f.role.toUpperCase(), 9) + pad(f.target, 10) +
          flags + '  "' + clip(f.label, 40) + '"');
      if (f.guessed) out("                role guessed from: \"" + clip(f.guessFrom, 50) + "\"");
    });

    const s = summarise(fields);
    out("");
    out("--- SUMMARY ---");
    out("Student " + s.student + " (reassess " + s.reStudent + ")   Assessor " + s.assessor +
        " (reassess " + s.reAssessor + ")   guessed " + s.guessed + "   already signed " + s.signed);

    if (book && EXPECTED[book]) {
      const e = EXPECTED[book];
      let ok = true;
      out("");
      out("--- COMPARE with " + book + " ---");
      for (const key in e) {
        const pass = s[key] === e[key];
        if (!pass) ok = false;
        out(pad(key, 12) + "got " + pad(s[key], 4) + "expected " + pad(e[key], 4) + (pass ? "OK" : "MISMATCH"));
      }
      out("RESULT: " + (ok ? "PASS" : "FAIL"));
    }
  });
}

/* ---------- button: Insert test picture ---------- */
function makeTestPng() {
  const c = document.createElement("canvas");
  c.width = 200; c.height = 60;
  const g = c.getContext("2d");
  g.fillStyle = "#ffffff"; g.fillRect(0, 0, 200, 60);
  g.strokeStyle = "#1a3d8f"; g.lineWidth = 2; g.strokeRect(1, 1, 198, 58);
  g.fillStyle = "#1a3d8f"; g.font = "20px Arial"; g.fillText("TEST SIG", 50, 37);
  return c.toDataURL("image/png").split(",")[1];   // no data: prefix
}

async function insertTest() {
  OUT = [];
  const tag = CONFIG.SIG_TAG + "student|test|" + new Date().toISOString();
  out("INSERT TEST  " + new Date().toLocaleString());
  out("Document: " + docName());
  await Word.run(async context => {
    const pic = context.document.getSelection().insertInlinePictureFromBase64(makeTestPng(), "End");
    pic.altTextDescription = tag;
    pic.altTextTitle = "SIG";          // backup, in case description is dropped
    await context.sync();
    pic.load("altTextDescription,altTextTitle");
    await context.sync();
    out("Wrote:     desc=\"" + tag + "\" title=\"SIG\"");
    out("Read now:  desc=\"" + pic.altTextDescription + "\" title=\"" + pic.altTextTitle + "\"");
  });
  out("");
  out("Next: wait for 'Saved', close the document tab, reopen it from SharePoint,");
  out("open the DEV panel and click 'Read back pictures'.");
}

/* ---------- button: Read back ---------- */
async function readBack() {
  OUT = [];
  out("READ BACK  " + new Date().toLocaleString());
  out("Document: " + docName());
  await Word.run(async context => {
    const pics = context.document.body.inlinePictures;
    pics.load("items/altTextDescription,items/altTextTitle,items/width,items/height");
    await context.sync();
    let sig = 0, titleOnly = 0;
    out("Inline pictures in body: " + pics.items.length);
    pics.items.forEach((p, i) => {
      const d = p.altTextDescription || "", t = p.altTextTitle || "";
      if (isSigPic(p)) sig++;
      else if (t === "SIG") titleOnly++;
      out(pad(i + 1, 4) + (isSigPic(p) ? "SIG " : "    ") + 'desc="' + clip(d, 55) +
          '" title="' + t + '" ' + Math.round(p.width) + "x" + Math.round(p.height));
    });
    out("");
    if (sig) out("PASS: " + sig + " picture(s) kept the SIG description");
    else if (titleOnly) out("PARTIAL: description lost but title kept (" + titleOnly + ") - we can tag via title");
    else out("FAIL: no tagged pictures found");
  });
}

/* ---------- copy ---------- */
async function copyOut() {
  const ta = document.getElementById("out");
  try { await navigator.clipboard.writeText(ta.value); status("Copied"); }
  catch (e) {
    ta.focus(); ta.select();
    try { document.execCommand("copy"); status("Copied"); }
    catch (e2) { status("Press Ctrl+C"); }
  }
}

/* ---------- wiring ---------- */
function setBusy(b) {
  document.querySelectorAll("button").forEach(x => { x.disabled = b; });
  status(b ? "Working..." : "");
}

function wire(id, fn) {
  document.getElementById(id).onclick = async () => {
    setBusy(true);
    try { await fn(); } catch (e) { reportError(e); }
    finally { flush(); setBusy(false); }
  };
}

Office.onReady(info => {
  if (info.host !== Office.HostType.Word) {
    document.getElementById("out").value = "This add-in only runs in Word.";
    return;
  }
  const m = docName().match(/MEM\d{5}/i);
  if (m && EXPECTED[m[0].toUpperCase()]) document.getElementById("expected").value = m[0].toUpperCase();
  wire("btnScan", runScan);
  wire("btnInsertTest", insertTest);
  wire("btnReadBack", readBack);
  document.getElementById("btnCopy").onclick = copyOut;
});
