// build-quote-template.js — Scaffold for generating an Adobe CPQ quote DOCX
// and converting it to PDF for a new deal.
//
// SETUP INSTRUCTIONS:
//   1. Copy this file to the deal folder:
//      deals/<account>-<year>/build-quote.js
//   2. Fill in all TODO sections below
//   3. Run: node build-quote.js
//   4. PDF will be created alongside the DOCX automatically
//
// HOW TO FIND THE DATA ROW paraId:
//   After extracting the source template, run:
//     node -e "const fs=require('fs'); const x=fs.readFileSync('word/document.xml','utf8');
//              x.match(/w14:paraId=\"[A-F0-9]+\"/g).slice(0,20).forEach(m=>console.log(m))"
//   Identify the paraId belonging to the single existing line item row in Table 3.

const fs   = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// ── TODO: fill these in ───────────────────────────────────────────────────────
const SRC = 'C:/Users/arjaiswa/Desktop/PATH-TO-CPQ-TEMPLATE.docx';  // source template
const OUT = 'C:/Users/arjaiswa/Desktop/claude-workspace/deals/ACCOUNT-YEAR/Quote.docx';
const TMP = path.join(process.env.TEMP || 'C:/Temp', 'quote-build');

// The single existing data row paraId in the source template's Table 3
// Find with the grep command above after first extraction.
const DATA_ROW_PARAID = '00000000';  // TODO: replace with actual paraId

// Original CPQ amount in the source template (as it appears in document.xml)
// This is what the subtotal and total cells currently show.
const ORIGINAL_AMOUNT = '0,000.00';  // TODO: e.g. '117,000.00'

const config = {
  licenseRows: [
    // [lineNo, description, qty, uom, term(mo), unitPrice, total]
    // TODO: fill with actual pricing rows
    ['0010', 'LEARNING MANAGER: YEAR 1 (POST-POC, MO 7-12)', '300.00',    'Each USER Per Year', '6.00',  '45.00',   '6,750.00'],
    ['0020', 'LEARNING MANAGER: YEAR 2',                      '5,000.00',  'Each USER Per Year', '12.00', '25.00', '125,000.00'],
    ['0030', 'LEARNING MANAGER: YEAR 3',                      '16,000.00', 'Each USER Per Year', '12.00', '12.00', '192,000.00'],
    ['0040', 'LEARNING MANAGER: YEAR 4',                      '40,000.00', 'Each USER Per Year', '12.00', '8.00',  '320,000.00'],
  ],
  licenseSubtotal: '643,750.00',  // TODO: sum of license rows
  addOnRows: [
    // Set to [] if no add-ons
    ['0050', 'LMS ADMIN - LEARNING MANAGER ADMINISTRATION (YEAR 1)', '1.00', 'Service Per Year', '12.00', '65,000.00', '65,000.00'],
    ['0060', 'LEARNER SUPPORT (ACTIVATES AT 6,000+ LEARNERS)',        'TBC',  'User Per Year',    'TBC',   'TBC',       'TBC'],
  ],
  addOnSubtotal:  '65,000.00',   // TODO: sum of confirmed add-on fees (TBC = $0)
  grandTotal:     '708,750.00',  // TODO: licenseSubtotal + addOnSubtotal
  addOnFootnote:  '* Learner Support pricing confirmed when learner count reaches 6,000+. Not included in Total Quote Fees.',
  includeAddOns:  true,          // set false if addOnRows is []
};

// ── Column widths (Adobe CPQ standard — do not modify) ───────────────────────
const COLS = [2402, 1999, 880, 1557, 813, 866, 2206];
const LABEL_W = COLS.slice(0, 6).reduce((s, w) => s + w, 0);

// ── XML helpers ──────────────────────────────────────────────────────────────
function dataRow(cells) {
  const tcs = cells.map((v, i) =>
    `<w:tc><w:tcPr><w:tcW w:w="${COLS[i]}" w:type="dxa"/></w:tcPr>` +
    `<w:p><w:pPr><w:keepLines/></w:pPr><w:r><w:t xml:space="preserve">${v}</w:t></w:r></w:p></w:tc>`
  ).join('');
  return `<w:tr><w:trPr><w:cantSplit/></w:trPr>${tcs}</w:tr>`;
}

function subtotalDataRow(label, value) {
  return `<w:tr><w:trPr><w:cantSplit/></w:trPr>` +
    `<w:tc><w:tcPr><w:tcW w:w="${LABEL_W}" w:type="dxa"/><w:gridSpan w:val="6"/>` +
    `<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="6A737B"/></w:tcBorders></w:tcPr>` +
    `<w:p><w:pPr><w:keepLines/><w:jc w:val="right"/></w:pPr>` +
    `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${label}</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="${COLS[6]}" w:type="dxa"/>` +
    `<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="6A737B"/></w:tcBorders></w:tcPr>` +
    `<w:p><w:pPr><w:keepLines/></w:pPr>` +
    `<w:r><w:rPr><w:b/></w:rPr><w:t>${value}</w:t></w:r></w:p></w:tc></w:tr>`;
}

function lineItemsTable(rows) {
  const headers = ['Line Number', 'Product Description', 'Quantity', 'Unit of Measure/ Metric', 'Term\n(Months)', 'Unit Price', 'Total Fees'];
  const hCells = headers.map((h, i) => {
    const inner = h.includes('\n')
      ? h.split('\n').map(p => `<w:p><w:pPr><w:keepNext/><w:keepLines/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t xml:space="preserve">${p} </w:t></w:r></w:p>`).join('')
      : `<w:p><w:pPr><w:keepNext/><w:keepLines/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>${h}</w:t></w:r></w:p>`;
    return `<w:tc><w:tcPr><w:tcW w:w="${COLS[i]}" w:type="dxa"/>` +
      `<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>` +
      `<w:shd w:val="clear" w:color="auto" w:fill="6A737B"/></w:tcPr>${inner}</w:tc>`;
  }).join('');
  const gridCols = COLS.map(w => `<w:gridCol w:w="${w}"/>`).join('');
  return `<w:tbl><w:tblPr><w:tblStyle w:val="LineItemsTable"/><w:tblW w:w="0" w:type="auto"/>` +
    `<w:tblInd w:w="112" w:type="dxa"/>` +
    `<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr>` +
    `<w:tblGrid>${gridCols}</w:tblGrid>` +
    `<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>${hCells}</w:tr>` +
    rows.map(r => dataRow(r)).join('') + `</w:tbl>`;
}

function sectionHeaderTable(title) {
  return `<w:tbl><w:tblPr><w:tblStyle w:val="SectionTable"/><w:tblW w:w="0" w:type="auto"/>` +
    `<w:tblLook w:val="0000"/></w:tblPr><w:tblGrid><w:gridCol w:w="10748"/></w:tblGrid>` +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="12000" w:type="dxa"/></w:tcPr>` +
    `<w:p><w:pPr><w:keepNext/><w:keepLines/></w:pPr>` +
    `<w:r><w:rPr><w:color w:val="6A737B"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>` +
    `<w:t>${title}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
}

function subsectionHeaderTable(title) {
  return `<w:tbl><w:tblPr><w:tblStyle w:val="LineItemsTable"/><w:tblW w:w="0" w:type="auto"/>` +
    `<w:tblInd w:w="107" w:type="dxa"/><w:tblLook w:val="0000"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="10728"/></w:tblGrid>` +
    `<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>` +
    `<w:tc><w:tcPr><w:tcW w:w="10728" w:type="dxa"/>` +
    `<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>` +
    `<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/></w:tcPr>` +
    `<w:p><w:pPr><w:keepNext/><w:keepLines/></w:pPr>` +
    `<w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>${title}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
}

function subtotalTable(label, value) {
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>` +
    `<w:tblCellMar><w:top w:w="43" w:type="dxa"/><w:left w:w="72" w:type="dxa"/>` +
    `<w:bottom w:w="29" w:type="dxa"/><w:right w:w="72" w:type="dxa"/></w:tblCellMar>` +
    `<w:tblLook w:val="0000"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="5313"/><w:gridCol w:w="4254"/><w:gridCol w:w="1268"/></w:tblGrid>` +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/>` +
    `<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders></w:tcPr>` +
    `<w:p><w:pPr><w:keepLines/></w:pPr></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="4704" w:type="dxa"/>` +
    `<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="6A737B"/>` +
    `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="6A737B"/>` +
    `<w:right w:val="single" w:sz="4" w:space="0" w:color="6A737B"/></w:tcBorders>` +
    `<w:shd w:val="clear" w:color="auto" w:fill="6A737B"/></w:tcPr>` +
    `<w:p><w:pPr><w:keepLines/></w:pPr>` +
    `<w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>` +
    `<w:t xml:space="preserve">${label}</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="1296" w:type="dxa"/>` +
    `<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="6A737B"/>` +
    `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="6A737B"/>` +
    `<w:right w:val="single" w:sz="4" w:space="0" w:color="6A737B"/></w:tcBorders>` +
    `<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/></w:tcPr>` +
    `<w:p><w:pPr><w:keepLines/><w:jc w:val="center"/></w:pPr>` +
    `<w:r><w:rPr><w:b/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>` +
    `<w:t>${value}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
}

const sep = `<w:p><w:pPr><w:keepLines/></w:pPr></w:p>`;

// ── Build ────────────────────────────────────────────────────────────────────
console.log('\nBuilding CPQ quote...');

if (fs.existsSync(TMP)) execSync(`rmdir /s /q "${TMP.replace(/\//g,'\\')}"`, { shell: 'cmd.exe' });
fs.mkdirSync(TMP, { recursive: true });

const srcZip = SRC.replace(/\.docx$/, '.tmp.zip');
fs.copyFileSync(SRC, srcZip);
try {
  execSync(`powershell -Command "Expand-Archive -LiteralPath '${srcZip}' -DestinationPath '${TMP}' -Force"`, { shell: 'cmd.exe' });
} finally {
  fs.unlinkSync(srcZip);
}

const docPath = path.join(TMP, 'word', 'document.xml');
let xml = fs.readFileSync(docPath, 'utf8');
const changes = [];

// 1. Replace data row
const t3Pattern = new RegExp(`<w:tr[^>]*w14:paraId="${DATA_ROW_PARAID}"[\\s\\S]*?<\\/w:tr>`);
const t3Match = xml.match(t3Pattern);
if (t3Match) {
  const newRows = config.licenseRows.map(r => dataRow(r)).join('') +
    subtotalDataRow('License Total', config.licenseSubtotal);
  xml = xml.replace(t3Match[0], newRows);
  changes.push(`Table 3: data row replaced (${config.licenseRows.length} years + subtotal)`);
} else {
  console.error(`  ✗ Data row paraId "${DATA_ROW_PARAID}" not found — update DATA_ROW_PARAID`);
  process.exit(1);
}

// 2. Update subtotal (first occurrence of original amount)
const t4Token = `<w:t>${ORIGINAL_AMOUNT}</w:t>`;
const t4Idx = xml.indexOf(t4Token);
if (t4Idx !== -1) {
  xml = xml.slice(0, t4Idx) + `<w:t>${config.licenseSubtotal}</w:t>` + xml.slice(t4Idx + t4Token.length);
  changes.push(`Table 4: subtotal → $${config.licenseSubtotal}`);
} else {
  console.error(`  ✗ Original amount "${ORIGINAL_AMOUNT}" not found for Table 4 subtotal`);
  process.exit(1);
}

// 3. Insert Add-On section (if applicable)
if (config.includeAddOns && config.addOnRows.length > 0) {
  const summaryMarker = '<w:tbl><w:tblPr><w:tblStyle w:val="SectionTable0"';
  const summaryIdx = xml.indexOf(summaryMarker);
  if (summaryIdx !== -1) {
    const addOnSection =
      sep +
      sectionHeaderTable('Add-On Services Pricing Detail:') +
      sep +
      subsectionHeaderTable('Adobe Add-On Services') +
      sep +
      lineItemsTable(config.addOnRows) +
      sep +
      subtotalTable('Adobe Add-On Services:', config.addOnSubtotal) +
      sep +
      `<w:p><w:pPr><w:keepLines/></w:pPr>` +
      `<w:r><w:rPr><w:color w:val="6A737B"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:i/></w:rPr>` +
      `<w:t xml:space="preserve">${config.addOnFootnote}</w:t></w:r></w:p>` +
      sep;
    xml = xml.slice(0, summaryIdx) + addOnSection + xml.slice(summaryIdx);
    changes.push(`Add-On Services section inserted ($${config.addOnSubtotal})`);
  } else {
    console.error('  ✗ SectionTable0 marker not found — could not insert Add-On section');
    process.exit(1);
  }
}

// 4. Update grand total (next occurrence of original amount)
const t6Token = `<w:t>${ORIGINAL_AMOUNT}</w:t>`;
const t6Idx = xml.indexOf(t6Token);
if (t6Idx !== -1) {
  xml = xml.slice(0, t6Idx) + `<w:t>${config.grandTotal}</w:t>` + xml.slice(t6Idx + t6Token.length);
  changes.push(`Table 6: Total Quote Fees → $${config.grandTotal}`);
} else {
  console.error(`  ✗ Original amount "${ORIGINAL_AMOUNT}" not found for Table 6 total`);
  process.exit(1);
}

// ── Write + repack ────────────────────────────────────────────────────────────
fs.writeFileSync(docPath, xml, 'utf8');

const outZip = OUT.replace(/\.docx$/, '.tmp.zip');
if (fs.existsSync(outZip)) fs.unlinkSync(outZip);
if (fs.existsSync(OUT))    fs.unlinkSync(OUT);

const outZipWin = outZip.replace(/\//g, '\\');
const tmpWin    = TMP.replace(/\//g, '\\');
execSync(`powershell -Command "Push-Location '${tmpWin}'; Compress-Archive -Path * -DestinationPath '${outZipWin}' -Force; Pop-Location"`, { shell: 'cmd.exe' });
fs.renameSync(outZip, OUT);

console.log('\n  DOCX built.');
console.log('  Output:', OUT);
console.log('  Size:  ', Math.round(fs.statSync(OUT).size / 1024), 'KB');
console.log('\n  Changes:');
changes.forEach(c => console.log('    ✓', c));

// ── Convert to PDF ────────────────────────────────────────────────────────────
const PDF_SCRIPT = 'C:/Users/arjaiswa/.claude/skills/sales-cpq-pdf/scripts/docx-to-pdf.ps1';
const outDocx = OUT.replace(/\//g, '\\');

console.log('\n  Converting to PDF...');
try {
  execSync(
    `powershell -ExecutionPolicy Bypass -File "${PDF_SCRIPT}" -InputDocx "${outDocx}"`,
    { shell: 'cmd.exe', stdio: 'inherit' }
  );
} catch (e) {
  console.error('  ✗ PDF conversion failed. Open the .docx in Word and use File > Save As > PDF as fallback.');
}
