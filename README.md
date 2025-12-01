# **Bitbucket Access Audit Automation**

Automated XLSX parsing → permission evaluation → screenshot capture → DOCX report generation

This tool automates Bitbucket access audits by:

* Reading and filtering an XLSX decision sheet
* Extracting all users whose access was **Revoked**
* Checking their permissions using the Bitbucket REST API
* Generating HTML evidence
* Capturing screenshots using **Puppeteer + Chromium**
* Generating a professional **DOCX report** with inline screenshots
* Producing CSV outputs for both access and no-access users

---

# 📁 **Project Structure**

```
project/
│
├── input_files/
│     └── SP_Decision_Sheet_Dummy.xlsx
│
├── output_files/
│     ├── revoked_rows.csv
│     ├── formatted_revoked_rows.csv
│     ├── access_check_results.csv
│     ├── no_access_check_results.csv
│     ├── html/
│     ├── png/
│     └── doc/
│          └── Bitbucket_Access_Report.docx
│
├── bitbucket_audit.js
├── .env
└── README.md
```

---

# 🚀 **Features**

### ✔ XLSX → CSV extraction

Parses input sheet and extracts only `Decision = "Revoked"` rows.

### ✔ Access verification

Uses Bitbucket REST API:

```
GET /rest/api/1.0/projects/{projectKey}/permissions/users?filter={username}
```

### ✔ Screenshot generation

Full-page screenshot of access evidence using Puppeteer + Chromium.

### ✔ DOCX report

Generates a **clean audit report** with:

* Title page
* One screenshot per page
* Inline images (NOT attachments)

### ✔ Output CSV reports

* `access_check_results.csv`
* `no_access_check_results.csv`

---

# 🔧 **Prerequisites (Ubuntu)**

Install all required dependencies in one command:

```bash
sudo apt update && sudo apt install -y \
  nodejs npm curl chromium-browser chromium-browser-l10n \
  libatk1.0-0 libatk-bridge2.0-0 libx11-xcb1 libxcomposite1 \
  libxcursor1 libxdamage1 libxi6 libxtst6 libnss3 libxrandr2 \
  libasound2 libpangocairo-1.0-0 libgtk-3-0 libgbm1 libpango-1.0-0 \
  libcairo2 libxss1 fonts-liberation xdg-utils ca-certificates \
  build-essential graphicsmagick git libreoffice
```

---

# 🧩 **Puppeteer Required Environment Variables**

Add to `.bashrc`:

```bash
export PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium-browser
export PUPPETEER_SKIP_DOWNLOAD=true
```

Reload:

```bash
source ~/.bashrc
```

---

# 🔐 **Environment Variables (.env)**

Create a `.env` file:

```
BB_URL=localhost:7990
BB_USERNAME=admin
BB_KEYNAME=your-bitbucket-api-token
```

---

# 📦 **Install Node Dependencies**

Inside project root:

```bash
npm install xlsx axios puppeteer dotenv officegen
```

---

# ▶️ **Run the Script**

```bash
node bitbucket_audit.js
```

---

# 📝 **Outputs Generated**

### 🔹 CSV files

| File                        | Description                  |
| --------------------------- | ---------------------------- |
| revoked_rows.csv            | Raw filtered revoked records |
| formatted_revoked_rows.csv  | Clean formatted fields       |
| access_check_results.csv    | Users who still have access  |
| no_access_check_results.csv | Users with no access         |

---

### 🔹 Evidence Files

| Directory             | Contents                     |
| --------------------- | ---------------------------- |
| `/output_files/html/` | Rendered HTML evidence       |
| `/output_files/png/`  | Screenshots of API responses |
| `/output_files/doc/`  | Final audit report (`.docx`) |

---

# 📘 **Final DOCX Report**

The generated report includes:

* Title page
* Automatically generated timestamp
* One screenshot per page
* Screenshots fit full document width
* Inline VML images compatible with MS Word

---

# 🧪 **Testing the Script**

To validate:

```bash
node bitbucket_audit.js
ls output_files/png
ls output_files/doc
```

Ensure:

* PNGs are created
* DOCX is generated
* No images appear as attachments

---
