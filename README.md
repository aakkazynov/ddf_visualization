# DDF Funders Network

An interactive network of funders that pooled or co-funded grant programs with the Doris Duke Foundation (DDF). It has filters for program, program area, year, and **where the money goes**: US only, international, or both, plus individual recipient countries.

Built with Claude Code from a single prompt ([ddf-network-prompt.md](ddf-network-prompt.md)) for DPI 851M, Assignment 3.

## Files

| File | What it is |
|---|---|
| `index.html` | The visualization (D3.js v7, loaded from a CDN) |
| `how-claude-built-this.html` | The prompt and a summary of how the page was built |
| `data.js` / `data.json` | The dataset (the page loads `data.js`, so it also works when opened directly from disk) |
| `DDF_Pooled_Fund_Programs.xlsx` | Workbook with tabs for Programs, Funders, Program–Funder links, Countries, Excluded items and Leads |
| `research/*.json` | Raw research output (programs, international programs, funder endowments) |
| `build_data.py` | Rebuilds `data.json`, `data.js` and the workbook from `research/` (needs `openpyxl`) |

## View it

Double-click `index.html`. You need an internet connection for D3 and the world map. Or serve the folder:

```powershell
python -m http.server 8000
```

Then open http://localhost:8000.

## Change the data

Edit the JSON files in `research/`, then run:

```powershell
python build_data.py
```

## Publish to GitHub Pages (when you're ready)

1. Create an empty repository on GitHub, for example `ddf-funders-network`.
2. In this folder, run:

```powershell
git remote add origin https://github.com/<your-username>/ddf-funders-network.git
git push -u origin main
```

3. On GitHub, go to **Settings → Pages → Deploy from a branch → main / (root)**. The site appears at `https://<your-username>.github.io/ddf-funders-network/`.

Vercel also works: import the repository, with no build step.
