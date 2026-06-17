# Markdown-to-CV 

Markdown-based CV with one-command DOCX and PDF generation.  
Edit `cv.md` → run build → get a formatted Word document and PDF.

---

## Quickstart

```bash
# 1. Clone and install
git clone https://github.com/rohan12patil/markdown-cv
cd markdown-cv
npm install

# 2. Edit your CV
# Open cv.md in VS Code and make changes

# 3. Build
npm run build           # → dist/<Name>_CV_MonYYYY.docx
npm run build:pdf       # → dist/<Name>_CV_MonYYYY.docx + .pdf
```

---

## File structure

```
markdown-cv/
├── cv.md                  ← Edit this — your CV source of truth
├── package.json
├── scripts/
│   └── build.js           ← Build script (md → docx + pdf)
├── .github/
│   └── workflows/
│       └── build.yml      ← CI: auto-builds on push to main
├── dist/                  ← Generated output (git-ignored)
│   ├── <Name>_CV_MonYYYY.docx
│   ├── <Name>_CV_MonYYYY.html  ← Intermediate file used for PDF generation
│   └── <Name>_CV_MonYYYY.pdf
└── .vscode/
    ├── tasks.json         ← VS Code build tasks (Ctrl+Shift+B)
    └── extensions.json    ← Recommended extensions
```

The output filename is derived automatically from the `# Name` heading in `cv.md` and stamped with the current month and year (e.g. `RohanPatil_CV_Jun2025.docx`).

---

## Editing cv.md

The CV uses standard Markdown. The build script parses specific patterns:

### Header fields
```markdown
# Your Name
**Role:** Your Title
**Email:** you@email.com
**Phone:** +44 7700000000
**Location:** City, Country
**LinkedIn:** https://linkedin.com/in/yourhandle
**GitHub:** https://github.com/yourhandle
```

### Professional Summary
Bullet-point list rendered as formatted bullets in both DOCX and PDF:
```markdown
## Professional Summary

- First summary point
- Second summary point
```

### Skills table
Standard Markdown table — add/remove rows freely:
```markdown
| Category    | Tools / Technologies         |
|-------------|------------------------------|
| Automation  | Playwright, Selenium, Appium |
| Languages   | TypeScript, Python, Java     |
| AI Testing  | Evals, Databricks            |
| Performance | K6                           |
| Methodology | Agile, Waterfall             |
```

### Employment entries
```markdown
### Job Title
**Company Name** · Location
_Start Date – End Date_

- Bullet point one
- Bullet point two
```

### Projects
```markdown
### Project Name — Short Description
**Stack:** TypeScript, Playwright
**URL:** github.com/yourhandle/project-name

- What it does
- Impact or metrics
```

### Education
```markdown
**Degree Name**
University Name, Location
_Start Date – End Date_
```

---

## Building in VS Code

Press `Ctrl+Shift+B` (or `Cmd+Shift+B` on Mac) to open the task picker:

- **Build CV (DOCX)** — generates `dist/<Name>_CV_MonYYYY.docx`
- **Build CV (DOCX + PDF)** — generates both
- **Watch CV** — auto-rebuilds every time you save `cv.md`

Or use the terminal:
```bash
npm run build
npm run build:pdf
npm run watch
```

---

## PDF generation

PDF is generated using a headless browser — **no LibreOffice required**.

The build script automatically detects an installed browser on your machine:

| Platform | Browsers checked (in order) |
|----------|------------------------------|
| Windows  | Microsoft Edge, Google Chrome |
| macOS    | Google Chrome, Microsoft Edge, Chromium |
| Linux    | google-chrome, chromium-browser, chromium |

Just run:
```bash
npm run build:pdf
```

If no supported browser is found, install [Google Chrome](https://www.google.com/chrome/) or [Microsoft Edge](https://www.microsoft.com/edge) and retry.

---

## GitHub Actions

The repository includes `.github/workflows/build.yml` which automatically builds the CV on every push to `main` (when `cv.md`, `build.js`, or `package.json` change) and on manual trigger.

Both the DOCX and PDF are uploaded as a downloadable artifact named **CV**, retained for 60 days. Find them under **Actions → your run → Artifacts**.

You can also trigger a build manually via **Actions → Build CV → Run workflow**.

---

