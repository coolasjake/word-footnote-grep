# Footnote Grep

A Microsoft Word task pane add-in that searches and replaces text in **footnotes** and **endnotes** using JavaScript regular expressions (grep-style).

## Requirements

- Microsoft Word 2021+, Microsoft 365, or Word on the web
- WordApi 1.5 (footnotes/endnotes API)
- Node.js 18+

## Quick start

```bash
cd Projects/word-footnote-grep
npm install
npm run certs          # trust localhost dev certificate (once)
npm run dev-server     # serves https://localhost:3000
npm start              # sideload add-in into Word (separate terminal)
```

In Word: **Home → Footnote Grep** opens the task pane.

## Usage

| Field | Example | Description |
|-------|---------|-------------|
| Pattern | `\bSmith\b` | JavaScript regex |
| Replacement | `$1 et al.` | Supports `$1`, `$2` capture groups |
| Global (g) | ✓ | Replace all matches per note |
| Ignore case (i) | | Case-insensitive matching |
| Multiline (m) | | `^` / `$` match line boundaries |

1. **Preview matches** — scans footnotes/endnotes and shows before/after
2. **Replace all** — applies replacement to every matching note
3. **Replace selected** — applies only to checked rows in the results table

## Scope

- **Footnotes** — body footnotes only
- **Endnotes** — endnotes only
- **Both** — scan and replace in both collections

## Example patterns

```
Find years:           \b(19|20)\d{2}\b
Fix spacing:          \s{2,}
Remove bracket refs:    \[(\d+)\]     →  ($1)
Normalize initials:     ([A-Z])\.     →  $1.
```

## Build for production

```bash
npm run build
npm run validate
```

Update `manifest.xml` URLs from `localhost:3000` to your hosted CDN before distributing.

## Project structure

```
manifest.xml          Office add-in manifest
src/taskpane/         Task pane UI and grep logic
webpack.config.js     Bundler config
```

## Limitations

- Regex runs in JavaScript, not Word's native wildcard engine
- Each note is processed independently (no cross-note patterns)
- Requires WordApi 1.5; older Word desktop builds are unsupported
