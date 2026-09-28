# AGENTS.md

Single-file static app: `index.html` only (plus fixtures `test.xlsx`, `t2.xlsx`). No build, no CI; the app itself needs no `npm` deps.

## Run
- No build step. Open `index.html` directly or serve: `python3 -m http.server` then open `http://localhost:8000/index.html`.
- Requires internet on load for CDN: SheetJS `xlsx-0.20.3` + `JSZip@3.10.1`. Data processing itself is fully local in browser.
- Offline logic tests live in `test/` (own `package.json`: `jszip@3.10.1` + `xlsx@0.18.5`, the closest npm release to the CDN build, used only to parse fixtures). Install once: `npm install` in `test/` (`node_modules` stays there, no reinstalls needed). Run: `npm test` in `test/` — exercises the real `<script>` from `index.html` against `test.xlsx` (35 asserts: colors, missing fallback, corrections, report text, copy/used/current UI states). Uses `globalThis` (not `const`) for `window`/`navigator` stubs — `const` would break the xlsx UMD `typeof window` guard via TDZ, and Node ≥21 needs `defineProperty` for `navigator`.
- Verify manually: load a real Socrative `.xlsx` export, click "Generiraj Socrative izvještaj", check per-student reports + duplicate-merge modal.
- Test fixture `test.xlsx` is a real Socrative export. Verified 2026-09-10: column K (mm question) has zero green cells, which exercises the no-green-column fallback (manual answer entry via `#missingWrap`), not the old hard error.
- Correct answer = first non-empty cell in column with green fill. Columns with no green cell do NOT abort: they are collected into `missing` (`makeQuestionData` returns `{questions, missing, style}`), rendered as question text + text input in `#missingWrap` below the form (auto-scrolled into view via `renderMissing`), and merged back sorted by column in `continueWithManualAnswers` (same `A. `-prefix strip as file answers). Empty manual input blocks with an error and focuses the field.

## Structure
- Everything (CSS + UI + logic in one `<script>` with `"use strict"`, vanilla JS) lives in `index.html`. Keep it that way — do not split or add a bundler.

## Critical logic — do not regress
- Correct/wrong comes from cell fills, not text: green `DAF0DE` = correct, red `F9CCCC` = incorrect (`isGreen`/`isRed`, suffix match to tolerate alpha prefix).
- `styleMapFromXlsx()` must read `xl/styles.xml` (`<fills>` → `fillColors`, `<cellXfs>` → `fillId`) + worksheet `s="…"` attrs via JSZip. Do not rely only on SheetJS `cell.s` — it varies by version/browser.
- Keep the anchored regexes for `<fills>…</fills>`, `<fill>…</fill>`, `<xf …>` — a broad regex previously merged multiple `<fill>` blocks.
- Sheet detection: first sheet with `Student Name` in col 0 (`findDataSheet`). Questions start at col 4; strip `(type)` prefix and `[...]` suffix. Correct answer = first non-empty cell in column with green fill. Rows `Class Scoring` / `Report Generated:` are skipped; attempts with 0 answers are skipped.
- Name matching: `cleanName` strips diacritics/digits, uppercases. `similarity()` = max of whole-string, token, containment scores. `likelySame`: `>=0.94` auto-merge, `>=0.86` ask user. Never use row order to match. Multi-attempt rule: correct only if never `wrong` and at least once `correct`; unanswered → wrong with `-`.
- Groups answering `<= questions.length - 6` columns are filtered out (`createReport`). Sort is last-name + first-name with `localeCompare(..., "hr")`.
- `parseDescriptions` formats: `88 text`, `88A text` / `A88 text` (per-type), `(type) label`. Equal thresholds = rotating variants via `randomSeed` + counter.
- Corrections (`#ispravci`, empty by default): `parseCorrections` lines like `5. A` (number = 1-indexed position in final sorted `questions`, value = letter or full text). `applyCorrections` in `proceedWithQuestions` overwrites `q.correct` and sets `q.correction`; `makeAttempts` then recomputes correct/wrong via `correctionMatches` (single alnum char = leading `X.` choice letter, else prefix-stripped case-insensitive full-text match) instead of colors. `resolveCorrectionTexts` then replaces a single-letter `q.correct` with the full stripped text of the first attempt answer carrying that letter. Report lines always show plain text: both student (`ans`) and correct (`q.correct`) get the `X. `-prefix stripped at display in `createReport`. Bad line numbers / bad question numbers throw (surfaced via existing error box).
