# Windermere Circuit Validator

A browser-based tool for validating new test results against a master circuit reference. No server required — runs entirely in the browser.

## What it does

Compares two files and checks whether every **Nomenclature + Serial Number** pair in your new results file exists in the master reference. Flags any rows where the circuit name, serial number, or combination don't match.

## How to use

1. Open `index.html` in any modern browser (or use the [GitHub Pages link](https://nehez.github.io/relay-data-checker/))
2. **Step 1** — Load your Master / Compare file (Excel or CSV)
   - Must contain `CIRCUIT NAME` and `SERIAL NUMBER` columns
   - Supports single-row or two-row headers
3. **Step 2** — Load your New Results file (Excel or CSV)
   - Must contain `Nomenclature` and `Serial Number` columns
4. Click **Run Validation**

## Results tabs

| Tab | What it shows |
|-----|---------------|
| **Failures** | Only the rows that did not match the master (red) |
| **Full Data** | All rows — green = passed, red = failed |
| **Not In Master** | Master records whose serial number was not found in the new file at all |

## Downloads

- **Download Excel Report** — Full report with separate sheets: Full Data, Failures, Not In Master, Summary
- **Download Failures CSV** — Just the failed rows as a CSV

## File structure

```
index.html    — HTML structure
styles.css    — All styling
app.js        — All validation logic and file parsing
```

## Enabling GitHub Pages

To host this tool directly from GitHub:

1. Go to **Settings → Pages** in this repository
2. Under **Source**, select `main` branch and `/ (root)`
3. Click **Save** — the tool will be live at `https://nehez.github.io/relay-data-checker/`
