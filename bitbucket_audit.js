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

const HAS_ACCESS_PNG_DIR = path.join(OUTPUT_DIR, "png/has_access");
const NO_ACCESS_PNG_DIR = path.join(OUTPUT_DIR, "png/no_access");

const DOC_DIR = path.join(OUTPUT_DIR, "doc");

// Create all required directories
ensureDir(OUTPUT_DIR);
ensureDir(HAS_ACCESS_PNG_DIR);
ensureDir(NO_ACCESS_PNG_DIR);
ensureDir(DOC_DIR);

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

  ensureDir(HTML_DIR);
  ensureDir(HAS_ACCESS_PNG_DIR);
  ensureDir(NO_ACCESS_PNG_DIR);

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
      'Screenshot File'
    ]) + '\n'
  );

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

    const htmlFile = path.join(
      HTML_DIR,
      `${user_sso}_${project_key}_${safe_ts}.html`
    );

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

    const pageIns = await getBrowserPage();
    await pageIns.goto('file://' + htmlFile, { waitUntil: 'networkidle0' });

    const requiredHeight = await pageIns.evaluate(() => document.body.scrollHeight);
    await pageIns.setViewport({ width: 1280, height: requiredHeight });

    // SELECT PNG DIRECTORY BASED ON ACCESS
    const pngFile = has_access
      ? path.join(HAS_ACCESS_PNG_DIR, `${user_sso}_${project_key}_${safe_ts}.png`)
      : path.join(NO_ACCESS_PNG_DIR, `${user_sso}_${project_key}_${safe_ts}.png`);

    await pageIns.screenshot({ path: pngFile });

    // LOG CSV
    if (has_access) {
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
      fs.appendFileSync(
        NO_ACCESS_OUTPUT_CSV,
        csvRow([
          user_sso,
          account_id,
          project_key,
          access_permission,
          'NO_ACCESS',
          ts,
          pngFile
        ]) + '\n'
      );
    }
  }

  if (browser) await browser.close();

  console.log('✔ Access verification complete');
}

// ------------------- Generate DOCX with screenshots -------------------

async function generateDoc(hasAccessDir, outputFileName) {
  console.log(`Generating DOCX → ${outputFileName}`);

  const images = fs.readdirSync(hasAccessDir).filter(f => f.endsWith(".png"));
  if (images.length === 0) {
    console.log(`No screenshots found in ${hasAccessDir}`);
    return;
  }

  const docx = officegen("docx");

  docx.on("error", console.error);

  let p = docx.createP({ align: "center" });
  p.addText(outputFileName.replace(".docx", ""), { bold: true, font_size: 28 });

  docx.createP();
  docx.createP({ align: "center" }).addText(
    `Generated: ${new Date().toLocaleString()}`,
    { font_size: 14 }
  );

  docx.createP().addText("", { pageBreakBefore: true });

  images.forEach((file, idx) => {
    const filePath = path.join(hasAccessDir, file);

    let imgP = docx.createP({ align: "center" });
    imgP.addImage(filePath);

    if (idx < images.length - 1) {
      docx.createP().addText("", { pageBreakBefore: true });
    }
  });

  const outPath = path.join(DOC_DIR, outputFileName);
  const out = fs.createWriteStream(outPath);

  return new Promise((resolve, reject) => {
    out.on("error", reject);
    out.on("close", () => {
      console.log(`✔ DOCX generated → ${outPath}`);
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
    await generateDoc(HAS_ACCESS_PNG_DIR, "Bitbucket_Has_Access_Report.docx");
    await generateDoc(NO_ACCESS_PNG_DIR, "Bitbucket_No_Access_Report.docx");

}

main().catch((err) => {
  console.error('Error running Bitbucket audit script:', err);
  process.exit(1);
});
