#!/usr/bin/env node
/**
 * build.js — converts cv.md → RohanPatil_CV.docx (and optionally .pdf)
 * Usage:
 *   node scripts/build.js          → generates dist/RohanPatil_CV.docx
 *   node scripts/build.js --pdf    → also generates dist/RohanPatil_CV.pdf
 */

const fs   = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, BorderStyle, WidthType,
  TabStopType, TabStopPosition
} = require('docx');

// ─── Config ──────────────────────────────────────────────────────────────
const INPUT  = path.join(__dirname, '..', 'cv.md');
const OUTDIR = path.join(__dirname, '..', 'dist');
const OUTPUT = path.join(OUTDIR, 'CV.docx');

// ─── Colours ─────────────────────────────────────────────────────────────
const NAVY   = "1E2761";
const ACCENT = "3A86FF";
const MUTED  = "5A6585";
const WHITE  = "FFFFFF";

// ─── Border helpers ───────────────────────────────────────────────────────
const noBorder  = { style: BorderStyle.NONE, size: 0, color: WHITE };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ─── Paragraph factories ──────────────────────────────────────────────────
function sectionDivider() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 1 } },
    spacing: { before: 0, after: 160 },
    children: []
  });
}

function sectionHeading(text) {
  return new Paragraph({
    spacing: { before: 280, after: 60 },
    children: [new TextRun({ text, bold: true, size: 26, color: NAVY, font: "Cambria" })]
  });
}

function jobTitle(title, dateStr) {
  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: { before: 160, after: 40 },
    children: [
      new TextRun({ text: title, bold: true, size: 22, color: NAVY, font: "Calibri" }),
      new TextRun({ text: "\t" + dateStr, size: 20, color: ACCENT, font: "Calibri" })
    ]
  });
}

function jobCompany(text) {
  return new Paragraph({
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text, italics: true, size: 20, color: MUTED, font: "Calibri" })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 20, font: "Calibri" })]
  });
}

function projectTitle(title, dateOrUrl) {
  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: { before: 120, after: 40 },
    children: [
      new TextRun({ text: title, bold: true, size: 22, color: NAVY, font: "Calibri" }),
      new TextRun({ text: "\t" + dateOrUrl, size: 19, color: ACCENT, font: "Calibri" })
    ]
  });
}

function projectStack(text) {
  return new Paragraph({
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text, italics: true, size: 19, color: MUTED, font: "Calibri" })]
  });
}

function spacer() {
  return new Paragraph({ spacing: { before: 160, after: 0 }, children: [] });
}

function smallSpacer() {
  return new Paragraph({ spacing: { before: 120, after: 0 }, children: [] });
}

// ─── Parse cv.md ─────────────────────────────────────────────────────────
function parseMd(mdText) {
  const lines = mdText.split('\n');
  const data  = { header: {}, summary: '', skills: [], jobs: [], projects: [], education: '', languages: '' };

  let section = null;
  let currentJob = null;
  let currentProject = null;
  let summaryLines = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    // Top-level header fields (Name is h1, rest are bold key: value)
    if (line.startsWith('# ')) { data.header.name = line.slice(2).trim(); continue; }
    const fieldMatch = line.match(/^\*\*(.+?):\*\*\s*(.+)$/);
    if (fieldMatch && section === null) {
      data.header[fieldMatch[1].toLowerCase()] = fieldMatch[2].trim();
      continue;
    }

    // Section headings (h2)
    if (line.startsWith('## ')) {
      const s = line.slice(3).trim().toLowerCase();
      if (s.includes('summary'))   { section = 'summary';    continue; }
      if (s.includes('skill'))     { section = 'skills';     continue; }
      if (s.includes('employment')){ section = 'employment'; continue; }
      if (s.includes('project'))    { section = 'projects';  continue; }
      if (s.includes('education')) { section = 'education';  continue; }
      if (s.includes('language'))  { section = 'languages';  continue; }
      continue;
    }

    // Skip dividers
    if (line === '---') continue;

    // --- Summary ---
    if (section === 'summary' && line) { summaryLines.push(line); continue; }

    // --- Skills table rows ---
    if (section === 'skills') {
      const row = line.replace(/^\||\|$/g, '').split('|').map(c => c.trim());
      if (row.length >= 2 && row[0] && !row[0].startsWith('-') && !row[0].toLowerCase().includes('category')) {
        data.skills.push({ label: row[0], items: row[1].split(',').map(s => s.trim()) });
      }
      continue;
    }

    // --- Employment ---
    if (section === 'employment') {
      if (line.startsWith('### ')) {
        if (currentJob) data.jobs.push(currentJob);
        currentJob = { title: line.slice(4).trim(), company: '', date: '', bullets: [] };
        continue;
      }
      if (currentJob) {
        const compMatch = line.match(/^\*\*(.+?)\*\*\s*·?\s*(.+)?$/);
        if (compMatch && !currentJob.company) {
          currentJob.company = compMatch[1].trim() + (compMatch[2] ? ' · ' + compMatch[2].trim() : '');
          continue;
        }
        const dateMatch = line.match(/^_(.+)_$/);
        if (dateMatch) { currentJob.date = dateMatch[1].trim(); continue; }
        if (line.startsWith('- ')) { currentJob.bullets.push(line.slice(2).trim()); continue; }
      }
      continue;
    }

    // --- Projects ---
    if (section === 'projects') {
      if (line.startsWith('### ')) {
        if (currentProject) data.projects.push(currentProject);
        currentProject = { title: '', url: '', stack: '', bullets: [] };
        const titleLine = line.slice(4).trim();
        const urlSep = titleLine.lastIndexOf(' — ');
        currentProject.title = urlSep > -1 ? titleLine : titleLine;
        continue;
      }
      if (currentProject) {
        const stackMatch = line.match(/^\*\*Stack:\*\*\s*(.+)$/);
        if (stackMatch) { currentProject.stack = stackMatch[1].trim(); continue; }
        const urlMatch = line.match(/^\*\*URL:\*\*\s*(.+)$/);
        if (urlMatch) { currentProject.url = urlMatch[1].trim(); continue; }
        if (line.startsWith('- ')) { currentProject.bullets.push(line.slice(2).trim()); continue; }
      }
      continue;
    }

    // --- Education ---
    if (section === 'education') {
      if (line.startsWith('**')) {
        const m = line.match(/^\*\*(.+?)\*\*/);
        if (m) { data.education = m[1].trim(); }
      }
      const dateMatch = line.match(/^_(.+)_$/);
      if (dateMatch) { data.educationDate = dateMatch[1].trim(); }
      if (!line.startsWith('**') && !line.startsWith('_') && line) {
        data.educationInstitution = line.trim();
      }
      continue;
    }

    // --- Languages ---
    if (section === 'languages' && line) { data.languages = line; continue; }
  }

  // Flush open job/project
  if (currentJob) data.jobs.push(currentJob);
  if (currentProject) data.projects.push(currentProject);
  data.summary = summaryLines;
  return data;
}

// ─── Build skills table rows ──────────────────────────────────────────────
function skillRow(label, items) {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: 1800, type: WidthType.DXA },
        borders: noBorders,
        margins: { top: 60, bottom: 60, left: 0, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 19, color: NAVY, font: "Calibri" })] })]
      }),
      new TableCell({
        width: { size: 7560, type: WidthType.DXA },
        borders: noBorders,
        margins: { top: 60, bottom: 60, left: 0, right: 0 },
        children: [new Paragraph({ children: [new TextRun({ text: items.join('  ·  '), size: 19, color: "2C2C2C", font: "Calibri" })] })]
      })
    ]
  });
}

// ─── HTML builder (used for PDF generation) ──────────────────────────────
function buildHtml(cv) {
  const li = items => items.map(b => `<li>${b}</li>`).join('\n');
  const summaryBullets = cv.summary
    .filter(l => l.startsWith('- '))
    .map(l => l.slice(2).trim());

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<style>
  @page { margin: 18mm 16mm; }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Calibri, Arial, sans-serif; font-size: 10pt; color: #2C2C2C; }
  h1  { font-family: Cambria, Georgia, serif; color: #1E2761; font-size: 24pt; margin-bottom: 4px; }
  .role { color: #5A6585; font-size: 12pt; margin-bottom: 5px; }
  .contact { font-size: 9.5pt; margin-bottom: 10px; }
  .contact .link { color: #3A86FF; }
  .contact .sep  { color: #5A6585; margin: 0 5px; }
  .divider { border: none; border-top: 2px solid #1E2761; margin: 8px 0 14px; }
  h2 { font-family: Cambria, Georgia, serif; color: #1E2761; font-size: 12pt; font-weight: bold;
       border-bottom: 1px solid #1E2761; padding-bottom: 2px; margin: 14px 0 6px; }
  .job-block { margin-bottom: 10px; }
  .job-intro { break-inside: avoid; page-break-inside: avoid; break-after: avoid; page-break-after: avoid; }
  .job-row { display: flex; justify-content: space-between; align-items: baseline; margin: 8px 0 1px; }
  .job-title { font-weight: bold; color: #1E2761; font-size: 10.5pt; }
  .job-date  { color: #3A86FF; font-size: 9.5pt; white-space: nowrap; margin-left: 8px; }
  .job-company { color: #5A6585; font-style: italic; font-size: 9.5pt; margin-bottom: 4px; }
  ul { margin: 4px 0 8px 18px; }
  li { margin: 2px 0; font-size: 9.5pt; line-height: 1.4; }
  table { width: 100%; border-collapse: collapse; margin: 2px 0 8px; }
  td { padding: 2px 6px 2px 0; font-size: 9.5pt; vertical-align: top; }
  td.label { font-weight: bold; color: #1E2761; white-space: nowrap; width: 110px; }
  .proj-url   { color: #3A86FF; font-size: 9pt; }
  .proj-stack { color: #5A6585; font-style: italic; font-size: 9pt; margin-bottom: 3px; }
  p { font-size: 9.5pt; margin: 4px 0; }
</style>
</head>
<body>
  <h1>${cv.header.name || ''}</h1>
  <div class="role">${cv.header.role || ''}</div>
  <div class="contact">
    <span class="link">${cv.header.email || ''}</span>
    <span class="sep">·</span>${cv.header.phone || ''}
    <span class="sep">·</span>${cv.header.location || ''}
    <span class="sep">·</span><span class="link">${cv.header.linkedin || ''}</span>
  </div>
  <hr class="divider">

  <h2>Professional Summary</h2>
  <ul>${li(summaryBullets)}</ul>

  <h2>Skills</h2>
  <table>${cv.skills.map(s => `
    <tr><td class="label">${s.label}</td><td>${s.items.join('  ·  ')}</td></tr>`).join('')}
  </table>

  <h2>Employment History</h2>
  ${cv.jobs.map(j => `
    <div class="job-block">
      <div class="job-intro">
        <div class="job-row">
          <span class="job-title">${j.title}</span>
          <span class="job-date">${j.date}</span>
        </div>
        <div class="job-company">${j.company}</div>
      </div>
      <ul>${li(j.bullets)}</ul>
    </div>`).join('')}

  <h2>Projects</h2>
  ${cv.projects.map(p => `
    <div class="job-block">
      <div class="job-intro">
        <div class="job-row">
          <span class="job-title">${p.title}</span>
          <span class="proj-url">${p.url}</span>
        </div>
        <div class="proj-stack">${p.stack}</div>
      </div>
      <ul>${li(p.bullets)}</ul>
    </div>`).join('')}

  <h2>Education</h2>
  <div class="job-row">
    <span class="job-title">${cv.education}</span>
    <span class="job-date">${cv.educationDate || ''}</span>
  </div>
  ${cv.educationInstitution ? `<div class="job-company">${cv.educationInstitution}</div>` : ''}

</body>
</html>`;
}

// ─── Find installed headless browser ─────────────────────────────────────
function findBrowser() {
  const candidates = {
    win32: [
      'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
      'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
      'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
      'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
    ],
    darwin: [
      '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
      '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
      '/Applications/Chromium.app/Contents/MacOS/Chromium',
    ],
    linux: [
      '/usr/bin/google-chrome',
      '/usr/bin/chromium-browser',
      '/usr/bin/chromium',
    ],
  };
  const platform = process.platform === 'win32' ? 'win32'
                 : process.platform === 'darwin' ? 'darwin'
                 : 'linux';
  return (candidates[platform] || []).find(p => fs.existsSync(p)) || null;
}

// ─── Main build ───────────────────────────────────────────────────────────
async function build() {
  if (!fs.existsSync(INPUT)) { console.error(`cv.md not found at ${INPUT}`); process.exit(1); }
  const mdText = fs.readFileSync(INPUT, 'utf8');
  const cv = parseMd(mdText);

  const children = [];

  // Header
  children.push(new Paragraph({
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text: cv.header.name || 'Name', bold: true, size: 52, color: NAVY, font: "Cambria" })]
  }));
  children.push(new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: cv.header.role || '', size: 26, color: MUTED, font: "Calibri" })]
  }));
  children.push(new Paragraph({
    spacing: { before: 0, after: 0 },
    children: [
      new TextRun({ text: cv.header.email || '', size: 20, color: ACCENT, font: "Calibri" }),
      new TextRun({ text: `   ·   ${cv.header.phone || ''}   ·   ${cv.header.location || ''}   ·   `, size: 20, color: MUTED, font: "Calibri" }),
      new TextRun({ text: cv.header.linkedin || '', size: 20, color: ACCENT, font: "Calibri" })
    ]
  }));
  children.push(new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 1 } },
    spacing: { before: 120, after: 160 },
    children: []
  }));

  // Summary
  children.push(sectionHeading('Professional Summary'));
  children.push(sectionDivider());
  cv.summary.forEach(line => {
    if (line.startsWith('- ')) {
      children.push(bullet(line.slice(2).trim()));
    } else if (line) {
      children.push(new Paragraph({
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: line, size: 20, font: "Calibri", color: "1A1A2E" })]
      }));
    }
  });
  children.push(spacer());

  // Skills
  children.push(sectionHeading('Skills'));
  children.push(sectionDivider());
  if (cv.skills.length > 0) {
    children.push(new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [1800, 7560],
      borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideH: noBorder, insideV: noBorder },
      rows: cv.skills.map(s => skillRow(s.label, s.items))
    }));
  }
  children.push(spacer());

  // Employment
  children.push(sectionHeading('Employment History'));
  children.push(sectionDivider());
  cv.jobs.forEach(job => {
    children.push(jobTitle(job.title, job.date));
    children.push(jobCompany(job.company));
    job.bullets.forEach(b => children.push(bullet(b)));
  });
  children.push(spacer());

  // Projects
  if (cv.projects.length > 0) {
    children.push(sectionHeading('Projects'));
    children.push(sectionDivider());
    cv.projects.forEach((proj, i) => {
      children.push(projectTitle(proj.title, proj.url));
      children.push(projectStack(proj.stack));
      proj.bullets.forEach(b => children.push(bullet(b)));
      if (i < cv.projects.length - 1) children.push(smallSpacer());
    });
    children.push(spacer());
  }

  // Education
  children.push(sectionHeading('Education'));
  children.push(sectionDivider());
  children.push(new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: { before: 60, after: 20 },
    children: [
      new TextRun({ text: cv.education, bold: true, size: 22, color: NAVY, font: "Calibri" }),
      new TextRun({ text: '\t' + (cv.educationDate || ''), size: 20, color: ACCENT, font: "Calibri" })
    ]
  }));
  if (cv.educationInstitution) {
    children.push(new Paragraph({
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: cv.educationInstitution, italics: true, size: 20, color: MUTED, font: "Calibri" })]
    }));
  }
  children.push(spacer());


  // Assemble document
  const doc = new Document({
    numbering: {
      config: [{
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 540, hanging: 300 } } } }]
      }]
    },
    styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
        }
      },
      children
    }]
  });

  if (!fs.existsSync(OUTDIR)) fs.mkdirSync(OUTDIR, { recursive: true });
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(OUTPUT, buf);
  console.log(`✅ DOCX → ${OUTPUT}`);

  // Optional PDF via headless Chrome/Edge
  if (process.argv.includes('--pdf')) {
    const htmlFile = path.join(OUTDIR, 'CV.html');
    const pdfFile  = path.join(OUTDIR, 'CV.pdf');
    fs.writeFileSync(htmlFile, buildHtml(cv), 'utf8');

    const browserPath = findBrowser();
    if (!browserPath) {
      console.warn('⚠️  PDF skipped — Chrome or Edge not found. Install one and retry.');
    } else {
      try {
        const puppeteer = require('puppeteer-core');
        const fileUrl = `file:///${htmlFile.replace(/\\/g, '/')}`;
        const browser = await puppeteer.launch({
          executablePath: browserPath,
          args: ['--no-sandbox', '--disable-gpu']
        });
        const page = await browser.newPage();
        await page.goto(fileUrl, { waitUntil: 'networkidle0' });
        await page.pdf({
          path: pdfFile,
          format: 'A4',
          margin: { top: '18mm', right: '16mm', bottom: '18mm', left: '16mm' },
          displayHeaderFooter: false,
          printBackground: true
        });
        await browser.close();
        console.log(`✅ PDF  → ${pdfFile}`);
      } catch (e) {
        console.warn('⚠️  PDF conversion failed:', e.message);
      }
    }
  }
}

build().catch(e => { console.error(e); process.exit(1); });
