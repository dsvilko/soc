"use strict";
// Offline harness: exercises the REAL <script> from ../index.html against
// ../test.xlsx (no browser needed).
//
// Run once:  npm install   (inside this directory)
// Run tests: npm test       (or: node test.js)
//
// NOTE: xlsx@0.18.5 is the closest npm release to the CDN build the app
// loads (xlsx-0.20.3); here it is only used to parse the fixture sheets.
// The style/color logic under test reads xl/*.xml directly via JSZip,
// exactly like the app does.

const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const XLSX = require("xlsx");

// ---------------------------------------------------------------- stubs
const registry = [];
function mkEl(tag) {
  const el = {
    tag: tag || "div",
    children: [],
    dataset: {},
    style: {},
    value: "",
    textContent: "",
    title: "",
    // NB: the app assigns el.className directly, so keep it synced with
    // the classList set like a real DOM element does.
    get className() { return [...this._cls].join(" "); },
    set className(v) { this._cls = new Set(String(v || "").split(/\s+/).filter(Boolean)); },
    placeholder: "",
    type: "",
    _cls: new Set(),
    _listeners: {},
    _html: "",
    _scrolled: false,
    _focused: false,
    get innerHTML() { return this._html; },
    set innerHTML(v) { this._html = v; this.children = []; },
    classList: null, // wired below
    setAttribute() {},
    appendChild(c) { this.children.push(c); return c; },
    append(...cs) { this.children.push(...cs); },
    addEventListener(t, fn) { this._listeners[t] = fn; },
    select() {},
    remove() {},
    focus() { this._focused = true; },
    scrollIntoView() { this._scrolled = true; },
    querySelectorAll(sel) {
      const out = [];
      const walk = (n) => {
        if (sel === "input[data-col]" && n.tag === "input" && n.dataset.col !== undefined) out.push(n);
        if (sel === "input[data-from]" && n.tag === "input" && n.dataset.from !== undefined) out.push(n);
        if (sel === "input[data-to]" && n.tag === "input" && n.dataset.to !== undefined) out.push(n);
        if (sel === "input" && n.tag === "input") out.push(n);
        if (sel === ".replace-row" && n._cls.has("replace-row")) out.push(n);
        if (sel === "button" && n.tag === "button") out.push(n);
        (n.children || []).forEach(walk);
      };
      walk(this);
      return out;
    },
    querySelector(sel) {
      const all = this.querySelectorAll(sel);
      return all.length ? all[0] : null;
    },
  };
  el.classList = {
    add: (...cs) => cs.forEach((c) => el._cls.add(c)),
    remove: (...cs) => cs.forEach((c) => el._cls.delete(c)),
    toggle: (c) => (el._cls.has(c) ? (el._cls.delete(c), false) : (el._cls.add(c), true)),
    contains: (c) => el._cls.has(c),
  };
  registry.push(el);
  return el;
}

const __els = {};
function getEl(id) {
  if (!__els[id]) __els[id] = mkEl("div");
  return __els[id];
}
// NOTE: assign via globalThis, never `const window`/`const document` in this
// module scope — the xlsx UMD build does `typeof window`, which throws under
// the temporal dead zone if this scope later declares `const window`.
globalThis.window = { scrollTo() {} };
// NB: Node ≥21 ships a getter-only global `navigator`, so plain assignment
// throws — define it as an own property instead.
Object.defineProperty(globalThis, "navigator", {
  value: { clipboard: { writeText: async (t) => { captured.clipboard = t; } } },
  configurable: true,
});
globalThis.document = {
  getElementById: getEl,
  createElement: (t) => mkEl(t),
  body: mkEl("body"),
  execCommand() { return true; },
  querySelectorAll(sel) {
    if (sel === ".student-name.current") {
      return registry.filter((e) => e._cls.has("student-name") && e._cls.has("current"));
    }
    return [];
  },
};
const captured = { clipboard: null };

// ------------------------------------------------------- load app script
const ROOT = path.join(__dirname, "..");
const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
const script = html.split("<script>")[1].split("</script>")[0];
eval(script + `
;globalThis.__api = { findDataSheet, makeQuestionData, makeAttempts, computeGroup,
  buildGroups, candidatePairs, parseCorrections, applyCorrections,
  resolveCorrectionTexts, correctionMatches, renderMissing,
  continueWithManualAnswers, proceedWithQuestions, finish, copyReport,
  applyTextReplacements, isMultipleChoice, displayPair,
  collectReplacements, addReplacementRow,
  __setPendingMissing: (v) => { pendingMissing = v; } };
`);
const A = globalThis.__api;

// ------------------------------------------------------------- assertions
let failures = 0;
function assert(cond, msg) {
  if (cond) console.log("  ok  " + msg);
  else { failures++; console.log("  FAIL " + msg); }
}

(async () => {
  console.log("fixture: test.xlsx");
  const buf = fs.readFileSync(path.join(ROOT, "test.xlsx"));
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  const wb = XLSX.read(buf, { type: "buffer", cellStyles: true, bookFiles: true, cellNF: true });

  const { ws, rows, headerRow } = A.findDataSheet(wb);
  assert(headerRow === 5, "header row is index 5 (Student Name row 6)");
  assert(rows[headerRow].slice(4).filter(Boolean).length === 13, "13 question columns (E..Q)");

  // --- question parsing: test.xlsx column K (mm) has zero green cells
  const parsed = await A.makeQuestionData(wb, ab, ws, rows, headerRow);
  assert(parsed.questions.length === 12, "12 questions parsed from colors (got " + parsed.questions.length + ")");
  assert(parsed.missing.length === 1 && parsed.missing[0].col === 10,
    "exactly 1 missing question (col K, mm question)");

  // --- manual-answer fallback for the missing column
  const manual = parsed.missing.map((m) => ({ col: m.col, header: m.header, type: m.type, correct: "17" }));
  const questions = parsed.questions.concat(manual).sort((a, b) => a.col - b.col);
  assert(questions.length === 13, "manual answer merged back (13 questions)");

  // --- corrections: "2. B" overrides Q2 file answer "metar"
  assert(questions[1].correct === "metar", "Q2 file correct is 'metar' before correction");
  A.applyCorrections(questions, A.parseCorrections("2. B"));
  const attempts = A.makeAttempts(rows, headerRow, questions, parsed.style);
  assert(attempts.length === 21, "21 attempts (got " + attempts.length + ")");
  A.resolveCorrectionTexts(questions, attempts);
  assert(questions[1].correct === "centimetar",
    "single-letter correction resolved to full text (got " + JSON.stringify(questions[1].correct) + ")");
  const g0 = A.computeGroup(attempts, [0], questions);
  assert(g0.answers[1].correct === false, "'A. metar' re-marked wrong after '2. B' correction");

  // --- full UI flow: missing box -> manual continue -> duplicates modal -> report
  getEl("zaglavlje").value = "Kviz provjera znanja.";
  getEl("opisi").value = "88 Odlicno.\n0 Lose.";
  getEl("pitanja").value = "Pitanja na koja nije dan ispravan odgovor #:";
  getEl("ispravci").value = "2. B";
  // mimic generate()'s missing branch (file already parsed above)
  A.__setPendingMissing({ questions: parsed.questions, missing: parsed.missing, style: parsed.style, rows, headerRow });
  A.renderMissing(parsed.missing);
  assert(getEl("missingWrap").style.display === "block", "missing box shown");
  assert(getEl("missingWrap")._scrolled === true, "auto-scrolled to missing box");
  const inputs = getEl("missingList").querySelectorAll("input[data-col]");
  assert(inputs.length === 1, "one manual input rendered");
  // empty submit is blocked
  inputs[0].value = "";
  getEl("error").style.display = "none";
  A.continueWithManualAnswers();
  assert(getEl("error").style.display === "block", "empty manual answer blocked with error");
  assert(inputs[0]._focused === true, "empty manual input focused");
  // valid submit continues into the duplicate-merge modal (ask-level pair exists)
  inputs[0].value = "17";
  A.addReplacementRow("predmenta", "predmeta");
  A.continueWithManualAnswers();
  assert(getEl("missingWrap").style.display === "none", "missing box hidden after valid input");
  assert(getEl("matchModal").style.display === "block", "duplicate-merge modal shown for ask-level pair");

  // --- report: hidden blocks, names with scores, copy content rules
  A.finish();
  const reportBoxes = registry.filter((e) => e._cls.has("student-copy"));
  const nameLabels = registry.filter((e) => e._cls.has("student-name"));
  assert(reportBoxes.length === nameLabels.length && reportBoxes.length > 15,
    "report blocks rendered (" + reportBoxes.length + ")");
  assert(getEl("reportContent")._cls.has("hideBlocks"), "blocks hidden by default");
  assert(getEl("toggleBlocks").textContent === "Prikaži izvještaje", "toggle button reset to show-state");
  assert(/\(\d+%\)$/.test(nameLabels[0].textContent), "score appended to name (" + JSON.stringify(nameLabels[0].textContent) + ")");

  // click the first name (= same as clicking its block): capture clipboard
  captured.clipboard = null;
  await nameLabels[0]._listeners.click();
  const text = captured.clipboard;
  assert(typeof text === "string" && text.length > 0, "name click copies report text");
  assert(text.startsWith("Kviz provjera znanja."), "copied text starts with zaglavlje (no student name)");
  assert(!text.split("\n")[0].match(/^[A-ZŠĐŽĆČ ]+ \(\d+%\)$/), "student name is not part of copied text");
  assert(/REZULTAT: \d+% \(\d+% je prosjek razreda\)/.test(text), "REZULTAT line carries class average");
  assert(/\(odgovoreno \| ispravan odgovor\)/.test(text), "question-list prefix has parentheses");
  assert(nameLabels[0]._cls.has("current"), "clicked name marked current");
  assert(reportBoxes[0]._cls.has("used"), "clicked block marked used");
  // clicking a second name moves the highlight
  await nameLabels[1]._listeners.click();
  assert(!nameLabels[0]._cls.has("current") && nameLabels[1]._cls.has("current"),
    "current highlight moves to last clicked name");
  assert(nameLabels[0]._cls.has("used"), "previously clicked name stays used");
  // every report: bad-answer lines show plain text, no "X. " choice prefixes
  let totalBad = 0, leaked = null;
  for (const lbl of nameLabels) {
    await lbl._listeners.click();
    for (const line of captured.clipboard.split("\n").filter((l) => /^\d+\. /.test(l))) {
      totalBad++;
      if (/\([A-Za-z0-9]\. /.test(line) && leaked === null) leaked = line;
    }
  }
  assert(totalBad > 0, "bad-answer lines exist (" + totalBad + ")");
  assert(leaked === null, "no choice prefix in bad lines" + (leaked ? ": " + leaked : ""));

  // --- corrections parser edge cases
  assert(JSON.stringify(A.parseCorrections("5. A\n2: B")) === JSON.stringify({ 2: "B", 5: "A" }),
    "parseCorrections handles . and : separators");
  let threw = 0;
  try { A.parseCorrections("abc"); } catch (e) { threw++; }
  try { A.applyCorrections([{ correct: "x" }], { 9: "A" }); } catch (e) { threw++; }
  assert(threw === 2, "bad correction line / number throws");

  // --- text replacements (Ispravci teksta): whole-word, case-sensitive
  const allHtml = reportBoxes.map((e) => e._html).join("\n");
  assert(allHtml.includes("predmeta"), "replacement applied to question header in report");
  assert(!allHtml.includes("predmenta"), "original typo gone from report");
  assert(A.applyTextReplacements("metr i metri", [{ from: "metr", to: "metar" }]) === "metar i metri",
    "whole-word only (metri untouched)");
  assert(A.applyTextReplacements("Metr metr", [{ from: "metr", to: "X" }]) === "Metr X",
    "case-sensitive replacement");
  assert(A.applyTextReplacements("a+b (a+b)", [{ from: "a+b", to: "c" }]) === "c (c)",
    "special chars escaped, all occurrences replaced");
  assert(A.applyTextReplacements("predmenta, predmenta!", [{ from: "predmenta", to: "predmeta" }]) === "predmeta, predmeta!",
    "punctuation counts as word boundary");
  assert(A.applyTextReplacements("metr", [{ from: "", to: "X" }, { from: "metr", to: "" }]) === "",
    "empty search skipped, empty replacement deletes");
  assert(A.applyTextReplacements("a b", [{ from: "a", to: "b" }, { from: "b", to: "c" }]) === "c c",
    "pairs applied in order");
  assert(A.isMultipleChoice("A. metar") === true, "MC answer detected by choice prefix");
  assert(A.isMultipleChoice("17") === false, "free numeric answer is not MC");
  assert(A.isMultipleChoice("-") === false, "dash is not MC");
  const dp1 = A.displayPair("A. metr", "B. metr", [{ from: "metr", to: "metar" }]);
  assert(dp1.ans === "metar" && dp1.ok === "metar", "MC student + correct answers replaced");
  const dp2 = A.displayPair("17", "17", [{ from: "17", to: "18" }]);
  assert(dp2.ans === "17" && dp2.ok === "18", "free student answer untouched, correct replaced");
  // dynamic rows: collect skips empty search
  getEl("zamjeneList").innerHTML = "";
  A.addReplacementRow("predmenta", "predmeta");
  A.addReplacementRow("", "ignored");
  A.addReplacementRow("metr", "metar");
  assert(JSON.stringify(A.collectReplacements()) === JSON.stringify([{ from: "predmenta", to: "predmeta" }, { from: "metr", to: "metar" }]),
    "collectReplacements skips empty search");

  console.log(failures === 0 ? "\nALL TESTS PASSED" : "\n" + failures + " TEST(S) FAILED");
  process.exit(failures === 0 ? 0 : 1);
})().catch((e) => { console.error("FATAL", e); process.exit(1); });
