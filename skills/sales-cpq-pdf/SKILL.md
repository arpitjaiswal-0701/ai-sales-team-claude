---
name: sales-cpq-pdf
description: "End-to-end workflow for generating a customer-facing PDF quote from an Adobe CPQ Word template. Covers two sub-skills: (1) build-quote — populate a .docx CPQ template with deal-specific pricing rows via XML manipulation; (2) docx-to-pdf — convert the output .docx to PDF on Windows using Word COM automation. Use this any time a deal requires a PDF quote from the standard Adobe CPQ template."
risk: low
source: custom
date_added: "2026-05-08"
---

# Sales CPQ → PDF Export

## Overview

Converts an Adobe CPQ Word template into a customer-ready PDF with deal-specific pricing. Two sub-skills — **build** and **export** — can run independently or in sequence.

```
CPQ Template (.docx) + Pricing Config
        ↓  [sub-skill: build-quote]
   Modified Quote (.docx)
        ↓  [sub-skill: docx-to-pdf]
   Final Quote (.pdf)
```

---

## Sub-Skill 1: `build-quote` — Populate the DOCX Template

### When to use
Any time pricing changes or a new deal quote is needed from the Adobe CPQ template.

### How it works
The Adobe CPQ `.docx` is a ZIP archive of XML. The script:
1. Copies the source `.docx` → `.tmp.zip` (PowerShell 5.1 `Expand-Archive` rejects `.docx` extension)
2. Extracts to a temp directory
3. Surgically modifies `word/document.xml`:
   - Replaces the existing single-line data row with multi-year pricing rows
   - Updates the "Adobe Products and Services" subtotal cell
   - Inserts an Add-On Services section (if applicable) before the Summary of Fees section
   - Updates the "Total Quote Fees" cell
4. Repacks as `.tmp.zip`, renames to `.docx`

### Config format

When building a new deal script, pass pricing as a config object:

```js
const config = {
  src:   'C:/path/to/CPQ-Template.docx',          // source template (never modified)
  out:   'C:/path/to/deals/<account>/Quote.docx',  // output file
  licenseRows: [                                    // one entry per year/tier
    ['0010', 'LEARNING MANAGER: YEAR 1 (POST-POC, MO 7-12)', '300.00',    'Each USER Per Year', '6.00',  '45.00',   '6,750.00'],
    ['0020', 'LEARNING MANAGER: YEAR 2',                      '5,000.00',  'Each USER Per Year', '12.00', '25.00', '125,000.00'],
    ['0030', 'LEARNING MANAGER: YEAR 3',                      '16,000.00', 'Each USER Per Year', '12.00', '12.00', '192,000.00'],
    ['0040', 'LEARNING MANAGER: YEAR 4',                      '40,000.00', 'Each USER Per Year', '12.00', '8.00',  '320,000.00'],
  ],
  licenseSubtotal: '643,750.00',
  addOnRows: [                                      // optional; set to [] if no add-ons
    ['0050', 'LMS ADMIN - LEARNING MANAGER ADMINISTRATION (YEAR 1)', '1.00', 'Service Per Year', '12.00', '65,000.00', '65,000.00'],
    ['0060', 'LEARNER SUPPORT (ACTIVATES AT 6,000+ LEARNERS)',        'TBC',  'User Per Year',    'TBC',   'TBC',       'TBC'],
  ],
  addOnSubtotal: '65,000.00',                       // sum of add-on rows (TBC rows = $0)
  grandTotal:    '708,750.00',
  addOnFootnote: '* Learner Support pricing confirmed when learner count reaches 6,000+. Not included in Total Quote Fees.',
};
```

### Column widths (Adobe CPQ standard — do not change)

| Col | Content | Width (twips) |
|-----|---------|--------------|
| 0 | Line Number | 2402 |
| 1 | Product Description | 1999 |
| 2 | Quantity | 880 |
| 3 | Unit of Measure / Metric | 1557 |
| 4 | Term (Months) | 813 |
| 5 | Unit Price | 866 |
| 6 | Total Fees | 2206 |

### XML markers used for replacement

| Target | Marker |
|--------|--------|
| Existing single-line data row | `w14:paraId` of the row (inspect `word/document.xml` for each new template) |
| "Adobe Products and Services" subtotal | First `<w:t>NNN,NNN.NN</w:t>` matching the original CPQ amount |
| Add-On section insertion point | `<w:tbl><w:tblPr><w:tblStyle w:val="SectionTable0"` |
| "Total Quote Fees" total | Second `<w:t>NNN,NNN.NN</w:t>` matching the original CPQ amount |

**Finding the data row paraId for a new template:**
```bash
node -e "
const fs = require('fs');
const xml = fs.readFileSync('word/document.xml', 'utf8');
const m = xml.match(/w14:paraId=\"([A-F0-9]+)\"/g);
console.log(m.slice(0,20));
"
```
Look for the paraId that belongs to the single existing line item row in Table 3.

### Reference implementation
`C:\Users\arjaiswa\.claude\skills\pptx-official\kc-workspace\build-quote.js`
(Kimberly-Clark deal — 4-year Option B ramp)

### Scaffold for new deals
`scripts/build-quote-template.js` in this skill directory — copy and fill in config.

---

## Sub-Skill 2: `docx-to-pdf` — Convert DOCX → PDF (Windows)

### When to use
After `build-quote` produces the output `.docx`, or any time a `.docx` needs PDF conversion without opening Word manually.

### Script
`scripts/docx-to-pdf.ps1` in this skill directory.

### Invocation
```powershell
# Basic — PDF lands next to the .docx with same name
powershell -ExecutionPolicy Bypass -File "C:\Users\arjaiswa\.claude\skills\sales-cpq-pdf\scripts\docx-to-pdf.ps1" -InputDocx "C:\path\to\Quote.docx"

# Explicit output path
powershell -ExecutionPolicy Bypass -File "C:\Users\arjaiswa\.claude\skills\sales-cpq-pdf\scripts\docx-to-pdf.ps1" -InputDocx "C:\path\to\Quote.docx" -OutputPdf "C:\path\to\Quote.pdf"
```

Or from Node.js within `build-quote.js`:
```js
const { execSync } = require('child_process');
const SCRIPT = 'C:/Users/arjaiswa/.claude/skills/sales-cpq-pdf/scripts/docx-to-pdf.ps1';

execSync(
  `powershell -ExecutionPolicy Bypass -File "${SCRIPT}" -InputDocx "${OUT.replace(/\//g,'\\')}"`,
  { shell: 'cmd.exe', stdio: 'inherit' }
);
```

### How it works
Uses Word COM automation — opens the `.docx` in a hidden Word instance, calls `SaveAs2` with format `wdFormatPDF` (17), then closes cleanly. Requires Word to be installed (guaranteed on this machine).

### Output
PDF is placed at `<same-directory-as-docx>/<same-name>.pdf` unless `-OutputPdf` is specified.

---

## Full Pipeline

When both steps are needed end-to-end:

1. Copy and fill `scripts/build-quote-template.js` with deal-specific config
2. Run: `node build-quote-<account>.js`
3. Run: `powershell -ExecutionPolicy Bypass -File ".../docx-to-pdf.ps1" -InputDocx "<out-path>"`
4. Verify: open PDF and confirm all tables, branding, and totals are correct

Or call `docx-to-pdf.ps1` from within the Node script itself (append after `fs.renameSync(outZip, OUT)`).

---

## Pricing Calculation Reference

For a 4-year ALM ramp deal with Year 1 = 6-month post-POC period:

```
durationFactor = [0.5, 1, 1, 1]   // Y1 is months 7–12 only

Year 1 total = learners[0] × rate[0] × 0.5
Year 2 total = learners[1] × rate[1]
Year 3 total = learners[2] × rate[2]
Year 4 total = learners[3] × rate[3]

License subtotal = sum(Y1..Y4)
Add-On total     = sum of fixed add-on fees (TBC rows = $0)
Grand total      = License subtotal + Add-On total
```

Current KC Option B tiers (2026-05-08):
| Tier | Users | $/user/yr |
|------|-------|-----------|
| Y1   | 300   | $45       |
| Y2   | 5,000 | $25       |
| Y3   | 16,000| $12       |
| Y4   | 40,000| $8        |

---

## Troubleshooting

| Error | Cause | Fix |
|-------|-------|-----|
| `ENOENT` on SRC | Template path wrong or file moved | Verify path; check `deals/` folder |
| `EBUSY` on OUT | Output `.docx` open in Word | Close Word or change output filename (append `-v2`, etc.) |
| `NotSupportedArchiveFileExtension` | PowerShell `Expand-Archive` rejecting `.docx` | Script already handles this via `.tmp.zip` copy — verify `fs.copyFileSync(SRC, srcZip)` is present |
| Data row paraId not found | New CPQ template version with different paraId | Run the paraId grep above and update the pattern |
| Subtotal/total value not found | Original CPQ template had a non-standard amount | Update the search token to match the template's actual value |
| PDF blank or missing header | Word failed silently | Open the `.docx` in Word manually to confirm it's valid before re-running PDF step |

---

## When to Use
Invoke when a deal requires a PDF quote from the Adobe CPQ Word template. Covers any account using the standard Adobe `LineItemsTable` / `SectionTable0` CPQ format.

## Limitations
- PDF conversion requires Microsoft Word installed (Windows only — this machine qualifies)
- Does not support non-CPQ Word templates — use `docx-official` for general DOCX work
- The `w14:paraId` marker is template-version-specific; re-identify for each new CPQ template download
