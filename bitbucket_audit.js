// bitbucket_audit.js

'use strict';

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');
const puppeteer = require('puppeteer');
require('dotenv').config();
const { Document, Packer, Paragraph, ImageRun } = require("docx");
const officegen = require('officegen');



// ------------------- CONFIG -------------------

const INPUT_DIR = path.join(__dirname, 'input_files');
const OUTPUT_DIR = path.join(__dirname, 'output_files');

const INPUT_XLSX = path.join(INPUT_DIR, 'SP_Decision_Sheet_Dummy.xlsx');
const REVOKED_CSV = path.join(OUTPUT_DIR, 'revoked_rows.csv');
const FORMATTED_CSV = path.join(OUTPUT_DIR, 'formatted_revoked_rows.csv');

const ACCESS_OUTPUT_CSV = path.join(OUTPUT_DIR, 'access_check_results.csv');
const NO_ACCESS_OUTPUT_CSV = path.join(OUTPUT_DIR, 'no_access_check_results.csv');
const SCREENSHOTS_DIR = path.join(OUTPUT_DIR, 'screenshots');

// Bitbucket API config
const URL = process.env.BB_URL;
const USERNAME = process.env.BB_USERNAME;
const KEYNAME = process.env.BB_KEYNAME;

if (!URL || !USERNAME || !KEYNAME) {
  console.error("Missing Bitbucket environment variables. Add them to .env");
  process.exit(1);
}


// ------------------- UTILS -------------------

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

// Trim helper (like xargs in bash)
function trim(value) {
  return String(value || '').trim();
}

// Format timestamp: equivalent to `date "+%Y-%m-%d %H:%M:%S"`
function formatTimestamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const yyyy = d.getFullYear();
  const mm = pad(d.getMonth() + 1);
  const dd = pad(d.getDate());
  const hh = pad(d.getHours());
  const mi = pad(d.getMinutes());
  const ss = pad(d.getSeconds());
  return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
}

// Safe timestamp for filenames: equivalent to `date "+%Y-%m-%d_%H-%M-%S"`
function formatSafeTimestamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const yyyy = d.getFullYear();
  const mm = pad(d.getMonth() + 1);
  const dd = pad(d.getDate());
  const hh = pad(d.getHours());
  const mi = pad(d.getMinutes());
  const ss = pad(d.getSeconds());
  return `${yyyy}-${mm}-${dd}_${hh}-${mi}-${ss}`;
}

// Simple CSV row builder (no quoting, to stay close to your shell output)
function csvRow(fields) {
  return fields.map((f) => String(f ?? '').replace(/\r?\n/g, ' ')).join(',');
}

// ------------------- STEP 1 + 2: XLSX -> revoked_rows.csv -------------------

function extractRevokedRowsFromXlsx() {
  console.log('Step 1 & 2: Reading XLSX and extracting revoked rows...');

  if (!fs.existsSync(INPUT_XLSX)) {
    throw new Error(`Input XLSX not found at: ${INPUT_XLSX}`);
  }

  const workbook = xlsx.readFile(INPUT_XLSX);
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];

  // Convert sheet to JSON, first row as headers
  const rows = xlsx.utils.sheet_to_json(worksheet, { defval: '' });

  if (rows.length === 0) {
    console.warn('No data found in XLSX file.');
  }

  // Header keys in order
  const header = rows.length > 0 ? Object.keys(rows[0]) : [];

  // Filter rows where Decision == "Revoked"
  const revokedRows = rows.filter(
    (r) => trim(r['Decision']) === 'Revoked'
  );

  // Write revoked_rows.csv (header + revoked rows)
  ensureDir(OUTPUT_DIR);
  const lines = [];
  if (header.length > 0) {
    lines.push(csvRow(header));
    for (const row of revokedRows) {
      const fields = header.map((h) => row[h] ?? '');
      lines.push(csvRow(fields));
    }
  }

  fs.writeFileSync(REVOKED_CSV, lines.join('\n'), 'utf8');

  console.log('Extraction completed.');
  console.log(`Filtered rows saved to: ${REVOKED_CSV}`);

  return revokedRows;
}

// ------------------- STEP 3: revoked_rows.csv -> formatted_revoked_rows.csv -------------------

function transformRevokedRowsToFormatted(revokedRows) {
  console.log('Step 3: Transforming revoked rows to formatted_revoked_rows.csv...');

  // Header as per your shell script
  const header = ['User SSO', 'Account ID', 'Project Key', 'Access Permission', 'Decision'];
  const lines = [csvRow(header)];

  for (const row of revokedRows) {
    const user_sso = trim(row['User SSO']);
    const account_id = trim(row['Account ID']);
    const entitlement_desc = trim(row['Entitlement Description']);
    const decision = trim(row['Decision']);

    // Example entitlement_desc: "P : PB-Admin"
    let project_key = '';
    let access_permission = '';

    if (entitlement_desc) {
      // Split on ':'
      const colonParts = entitlement_desc.split(':');
      if (colonParts[1]) {
        const rightSide = colonParts[1].trim(); // e.g. "PB-Admin"
        const dashParts = rightSide.split('-');
        project_key = trim(dashParts[0]);       // e.g. "PB"
        access_permission = trim(dashParts[1]); // e.g. "Admin"
      }
    }

    lines.push(csvRow([user_sso, account_id, project_key, access_permission, decision]));
  }

  fs.writeFileSync(FORMATTED_CSV, lines.join('\n'), 'utf8');

  console.log('Transformation complete.');
  console.log(`Output saved to ${FORMATTED_CSV}`);
}


// ------------------- STEP 4: Bitbucket access check + screenshots -------------------

async function performAccessCheck() {
  console.log('Step 4: Performing Bitbucket access check...');

  const HTML_DIR = path.join(OUTPUT_DIR, 'html');
  const PNG_DIR = path.join(OUTPUT_DIR, 'png');

  ensureDir(OUTPUT_DIR);
  ensureDir(HTML_DIR);
  ensureDir(PNG_DIR);

  // Prepare CSV headers
  fs.writeFileSync(
    ACCESS_OUTPUT_CSV,
    csvRow([
      'Username',
      'Account ID',
      'Project Key',
      'Access Permission',
      'Access Status',
      'Timestamp',
      'Screenshot File',
    ]) + '\n'
  );

  fs.writeFileSync(
    NO_ACCESS_OUTPUT_CSV,
    csvRow([
      'Username',
      'Account ID',
      'Project Key',
      'Access Permission',
      'Access Status',
      'Timestamp',
    ]) + '\n'
  );

  if (!fs.existsSync(FORMATTED_CSV)) {
    console.warn(`Formatted CSV not found at ${FORMATTED_CSV}`);
    return;
  }

  const content = fs.readFileSync(FORMATTED_CSV, 'utf8');
  const rows = content.split(/\r?\n/).filter((l) => l.trim());

  if (rows.length <= 1) return;

  let browser = null;
  let page = null;

  async function getBrowserPage() {
    if (!browser) {
      browser = await puppeteer.launch({ headless: true });
      page = await browser.newPage();
    }
    return page;
  }

  for (let i = 1; i < rows.length; i++) {
    const [user_sso, account_id, project_key, access_permission] =
      rows[i].split(',').map(trim);

    const api_url = `http://${URL}/rest/api/1.0/projects/${project_key}/permissions/users?filter=${encodeURIComponent(
      user_sso
    )}`;

    let responseData;
    try {
      const resp = await axios.get(api_url, {
        auth: { username: USERNAME, password: KEYNAME },
      });
      responseData = resp.data;
    } catch {
      responseData = { values: [] };
    }

    const has_access = Array.isArray(responseData.values)
      ? responseData.values.length > 0
      : false;

    const ts = formatTimestamp();
    const safe_ts = formatSafeTimestamp();

    if (has_access) {
      // HTML + PNG file paths
      const htmlFile = path.join(
        HTML_DIR,
        `${user_sso}_${project_key}_${safe_ts}.html`
      );
      const pngFile = path.join(
        PNG_DIR,
        `${user_sso}_${project_key}_${safe_ts}.png`
      );

      // Generate HTML
      const html = `
<html>
  <body style="font-family: monospace; padding: 20px;">
    <h2>Bitbucket Access Check</h2>
    <b>User:</b> ${user_sso}<br>
    <b>Project:</b> ${project_key}<br>
    <b>Timestamp:</b> ${ts}<br><br>

    <h3>REST API URL</h3>
    <p>${api_url}</p>

    <h3>API Response</h3>
    <pre>${JSON.stringify(responseData, null, 2)}</pre>
  </body>
</html>`;

      fs.writeFileSync(htmlFile, html);

      // Take screenshot without excessive blank area
      const pageIns = await getBrowserPage();
      await pageIns.goto('file://' + htmlFile, { waitUntil: 'networkidle0' });

      const requiredHeight = await pageIns.evaluate(() => {
        return document.body.scrollHeight;
      });

      await pageIns.setViewport({ width: 1280, height: requiredHeight });

      await pageIns.screenshot({ path: pngFile });

      // Log result
      fs.appendFileSync(
        ACCESS_OUTPUT_CSV,
        csvRow([
          user_sso,
          account_id,
          project_key,
          access_permission,
          'HAS_ACCESS',
          ts,
          pngFile,
        ]) + '\n'
      );
    } else {
      // NO ACCESS
      fs.appendFileSync(
        NO_ACCESS_OUTPUT_CSV,
        csvRow([
          user_sso,
          account_id,
          project_key,
          access_permission,
          'NO_ACCESS',
          ts,
        ]) + '\n'
      );
    }
  }

  if (browser) await browser.close();

  console.log('✔ Access verification complete');
  console.log(`✔ HTML stored under: ${HTML_DIR}`);
  console.log(`✔ PNG stored under: ${PNG_DIR}`);
}

// ------------------- Generate DOCX with screenshots -------------------

async function generateDocWithScreenshots() {
  console.log("Generating DOCX using officegen...");

  const PNG_DIR = path.join(OUTPUT_DIR, "png");
  const DOC_DIR = path.join(OUTPUT_DIR, "doc");
  ensureDir(DOC_DIR);

  const pngFiles = fs.readdirSync(PNG_DIR).filter(f => f.endsWith(".png"));
  if (pngFiles.length === 0) {
    console.log("No PNG files found — skipping DOC creation.");
    return;
  }

  // Create DOCX document
  const docx = officegen('docx');

  docx.on('error', (err) => {
    console.error('officegen error:', err);
  });

  // ---- Title page ----
  let p = docx.createP({ align: 'center' });
  p.addText('Bitbucket Access Evidence Collection Report', { bold: true, font_size: 28 });

  docx.createP(); // blank line

  p = docx.createP({ align: 'center' });
  p.addText(`Generated: ${new Date().toLocaleString()}`, { font_size: 14 });

  // Page break before screenshots
  let pb = docx.createP();
  pb.addText('', { pageBreakBefore: true });

  // ---- Screenshots ----
  pngFiles.forEach((file, idx) => {
    const filePath = path.join(PNG_DIR, file);

    // Heading above image
    // let headingP = docx.createP();
    // headingP.addText(`Screenshot: ${file}`, { bold: true, font_size: 20 });

    // Centered inline image
    let imgP = docx.createP({ align: 'center' });
    imgP.addImage(filePath); // officegen embeds as inline picture, not attachment

    // Page break *after* each image except the last
    if (idx < pngFiles.length - 1) {
      const br = docx.createP();
      br.addText('', { pageBreakBefore: true });
    }
  });

  const outputPath = path.join(DOC_DIR, 'Bitbucket_Access_Report.docx');
  ensureDir(DOC_DIR);
  const out = fs.createWriteStream(outputPath);

  return new Promise((resolve, reject) => {
    out.on('error', (err) => {
      console.error('Write stream error:', err);
      reject(err);
    });

    out.on('close', () => {
      console.log(`✔ DOCX generated → ${outputPath}`);
      resolve();
    });

    docx.generate(out);
  });
}



// ------------------- MAIN -------------------

async function main() {
  ensureDir(INPUT_DIR);
  ensureDir(OUTPUT_DIR);

  // Step 1 & 2: XLSX -> revoked_rows.csv
  const revokedRows = extractRevokedRowsFromXlsx();

  // Step 3: revoked_rows.csv -> formatted_revoked_rows.csv
  transformRevokedRowsToFormatted(revokedRows);

  // Step 4: access check via Bitbucket API + screenshots
  await performAccessCheck();

  // Generate DOC with screenshots
  await generateDocWithScreenshots();
}

main().catch((err) => {
  console.error('Error running Bitbucket audit script:', err);
  process.exit(1);
});
