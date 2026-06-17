# CV — Rohan Patil

Markdown-based CV with one-command DOCX/PDF generation.  
Edit `cv.md` → run build → get a formatted Word document.

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
npm run build           # → dist/RohanPatil_CV.docx
npm run build:pdf       # → dist/RohanPatil_CV.docx + .pdf (needs LibreOffice)
```

---

## File structure

```
cv/
├── cv.md                  ← Edit this — your CV source of truth
├── package.json
├── scripts/
│   └── build.js           ← Build script (md → docx/pdf)
├── dist/                  ← Generated output (git-ignored)
│   ├── RohanPatil_CV.docx
│   └── RohanPatil_CV.pdf
└── .vscode/
    ├── tasks.json         ← VS Code build tasks (Ctrl+Shift+B)
    └── extensions.json    ← Recommended extensions
```

---

## Editing cv.md

The CV uses standard Markdown. The build script parses specific patterns:

### Header fields
```markdown
# Your Name
**Role:** Your Title
**Email:** you@email.com
**Phone:** 07700000000
**Location:** City, Country
**LinkedIn:** linkedin.com/in/yourhandle
**GitHub:** github.com/yourhandle
```

### Skills table
Standard Markdown table — add/remove rows freely:
```markdown
| Category   | Tools / Technologies         |
|------------|------------------------------|
| Automation | Playwright, Selenium, Appium |
| Languages  | TypeScript, Python, Java     |
```

### Employment entries
```markdown
### Job Title
**Company Name** · Location
_Start Date – End Date_

- Bullet point one
- Bullet point two
```

### Open source projects
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
_Start Date – End Date_
```

---

## Building in VS Code

Press `Ctrl+Shift+B` (or `Cmd+Shift+B` on Mac) to open the task picker:

- **Build CV (DOCX)** — generates `dist/RohanPatil_CV.docx`
- **Build CV (DOCX + PDF)** — generates both (requires LibreOffice)
- **Watch CV** — auto-rebuilds every time you save `cv.md`

Or use the terminal:
```bash
npm run build
npm run build:pdf
npm run watch
```

---

## PDF generation

PDF export requires LibreOffice:

```bash
# Ubuntu / Debian
sudo apt install libreoffice

# macOS (via Homebrew)
brew install --cask libreoffice
```

Then run `npm run build:pdf`.

---

## GitHub Actions (optional)

To auto-generate and attach the DOCX on every push, create `.github/workflows/build.yml`:

```yaml
name: Build CV
on: [push]
jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 20
      - run: npm install
      - run: npm run build
      - uses: actions/upload-artifact@v4
        with:
          name: CV
          path: dist/RohanPatil_CV.docx
```

This uploads the generated `.docx` as a downloadable artifact on every commit.

---

## Recommended VS Code extensions

Install when prompted, or manually:

- **Markdown All in One** — shortcuts, TOC, formatting
- **Markdown Preview Enhanced** — live side-by-side preview
- **markdownlint** — catches formatting issues
