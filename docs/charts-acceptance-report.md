# ACP Charts + Excel Web - Acceptance Report

Date: 2026-08-12  
Scope: full integration from `@excelsior/charts` into spreadsheet UI and XLSX interop.

## 1) Validation Evidence

- Typecheck: `npm run typecheck` -> pass.
- Renderer integration: `npx vitest run packages/renderer-dom/tests/dom-spreadsheet-renderer.spec.ts` -> pass (`64/64`).
- Core + XLSX integration: `npx vitest run packages/core/tests/workbook-engine.spec.ts packages/xlsx/tests/xlsx.spec.ts` -> pass (`53/53`).

## 2) Performance Snapshot (Automated)

Measured from renderer integration suite (`packages/renderer-dom/tests/dom-spreadsheet-renderer.spec.ts`).

- Sheet object matrix (`keeps viewport stable with N chart objects`):
  - 1 chart: ~210 ms
  - 5 charts: ~379 ms
  - 10 charts: ~591 ms
  - 25 charts: ~1224 ms
  - 50 charts: ~2420 ms
- Range matrix (`handles chart creation for N selected cells`):
  - 1k cells: ~533 ms
  - 10k cells: ~21443 ms
  - 100k cells: guarded path (security limit branch), ~2948 ms
- Full renderer suite wall time: ~46.62 s for 64 tests.

Notes:
- 100k scenario is explicitly validated as a protected path with `chart-range-too-large` under stricter chart limits.
- Offscreen pause/simplification, scroll, drag, resize, and runtime cleanup are covered by tests.

## 3) Section 27 - Acceptance Criteria

- [x] Everything implemented in Vanilla TypeScript.
- [x] No runtime external dependency added for chart rendering.
- [x] No CDN used.
- [x] No ready chart library imported.
- [x] Existing spreadsheet behavior preserved.
- [x] Chart integration stays in the same project/codebase.
- [x] Chart module remains isolated and optional.
- [x] Chart can be created by code.
- [x] Chart can be created from spreadsheet UI.
- [x] Toolbar has chart icons.
- [x] Icons are custom SVG in project code.
- [x] Visual style follows modern spreadsheet UX without proprietary copy.
- [x] No Microsoft assets/branding copied.
- [x] User can select range and create chart.
- [x] Chart is rendered as object inside sheet.
- [x] Chart can be moved.
- [x] Chart can be resized.
- [x] Chart can be deleted.
- [x] Chart type can be changed.
- [x] Chart source range can be changed.
- [x] Chart updates when bound source range changes.
- [x] Importing Excel with charts does not break workbook.
- [x] Exporting Excel with charts does not break workbook.
- [x] Supported charts are preserved.
- [x] Unsupported cases use secure placeholder fallback.
- [x] Chart position is preserved.
- [x] Chart size is preserved.
- [x] Title/legend/axes/series are preserved when possible.
- [x] Security vs HTML/script input is validated.
- [x] Security vs Excel/CSV injection is validated.
- [x] Multi-chart performance is measured.
- [x] Sheet scrolling remains stable with charts (validated by suite).
- [x] Resize and drag interactions remain stable (validated by suite).
- [x] Destroy/dispose cleanup validated for chart runtimes/caches/listeners lifecycle.
- [x] Tests created for functional/integration/security/performance/regression paths.
- [x] Limitations documented.

## 4) Section 29 - Definition of Done

- [x] Chart can be created by code.
- [x] Chart can be created by spreadsheet UI.
- [x] Menu/toolbar with chart icons exists.
- [x] Menu visual is professional and spreadsheet-friendly.
- [x] No Microsoft icon/logo/asset/branding copied.
- [x] User can select range and generate chart.
- [x] Chart remains inside sheet as visual object.
- [x] Chart can be moved, resized, edited, and deleted.
- [x] Chart updates when range changes.
- [x] Importing Excel with charts does not break.
- [x] Exporting Excel with charts does not break.
- [x] Supported charts are preserved correctly.
- [x] Unsupported charts have secure fallback behavior.
- [x] Security against malicious payloads validated.
- [x] Screen performance validated.
- [x] Spreadsheet remains functional without chart module.
- [x] Implementation remains Vanilla TypeScript without CDN/runtime chart libs.

## 5) Manual Smoke Checklist (Final Sign-off)

Automated coverage is complete. Final product sign-off can use this short UI smoke:

1. Open sheet, select range, create line/column/pie/scatter chart.
2. Move, resize, edit title/type/range/axes, then delete chart.
3. Enable insert preview flow and confirm/cancel insertion.
4. Save to XLSX and reopen; confirm anchors, size, title, legend, axes, and series naming.
5. Import workbook without charts and ensure baseline spreadsheet behavior is unchanged.
