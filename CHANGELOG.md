# Changelog

## v1.0.0 — 2026-03-27

Initial release of the Windermere Circuit Validator.

### Features
- Upload Master file (Excel/CSV) with CIRCUIT NAME + SERIAL NUMBER columns
- Upload New Results file (Excel/CSV) with Nomenclature + Serial Number columns
- Validates every Nomenclature + Serial Number pair against the master reference
- **Failures tab** — shows only rows that did not match (red highlight)
- **Full Data tab** — all rows with green = PASS, red = FAIL row highlighting
- **Not In Master tab** — master records whose serial was absent from the new file
- Detailed issue descriptions per failed row (circuit missing, serial missing, pair mismatch)
- Summary stats: mismatch count, validated OK count, total rows, match rate %
- Download Excel report (Full Data, Failures, Not In Master, Summary sheets)
- Download Failures CSV
- Drag-and-drop file loading
- Supports single-row and two-row headers in master file

### Project structure
- `index.html` — HTML skeleton
- `styles.css` — all styling
- `app.js` — all validation logic, file parsing, and export
