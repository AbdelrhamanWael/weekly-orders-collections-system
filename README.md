# 📊 Weekly Orders & Collections System

A full-stack web application that **links weekly order data with collection data** from multiple e-commerce platforms, and generates detailed Excel reports automatically.

![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-2.x-black?logo=flask&logoColor=white)
![SQLite](https://img.shields.io/badge/SQLite-Database-lightblue?logo=sqlite)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

---

## ✨ Features

- 📁 **Multi-platform file upload** — supports CSV, XLSX, XLS from Amazon, Noon, Trendyol, Ilasouq, Tabby, SMSA
- ⚙️ **Automated data processing** — links orders with collections and calculates profits
- 📊 **Interactive charts** — visual analytics powered by Chart.js
- 📈 **Excel report generation** — styled weekly reports with one click
- 🔄 **Weekly reset** — clears previous week's data while keeping all old reports
- 🖥️ **Simple launcher** — double-click `run.bat` to start everything

---

## 🚀 Quick Start

### For End Users (Windows)

```
Double-click:  run.bat
```

The browser will open automatically at **http://127.0.0.1:5000**

### For Developers

```bash
# Install dependencies (once)
pip install -r requirements.txt

# Initialize the database (once)
python init_db.py

# Run the application
python app.py
```

---

## 📅 Weekly Workflow

```
Every new week:

1️⃣  Run run.bat
      ↓ Browser opens automatically

2️⃣  Dashboard → Click "🔄 New Week"
      ↓ Clears previous week's data

3️⃣  Upload Files → Upload this week's order & collection files
      (CSV or XLSX or XLS from all platforms)

4️⃣  Process Data → Click "🚀 Run Full Processing"
      ↓ Processes, links, and calculates everything

5️⃣  Reports → Download your Excel report ✅
```

> **Note:** Old reports are never deleted — they are always available on the Reports page.

---

## 🗂️ Project Structure

```
weekly-orders-collections-system/
│
├── 📄 app.py                    # Flask backend — main entry point
├── 📄 process_data.py           # Orders file processing
├── 📄 process_collections.py    # Collections file processing
├── 📄 generate_report.py        # Excel report generation
├── 📄 init_db.py                # Database initialization (run once)
│
├── 📁 templates/
│   └── index.html               # Frontend UI (HTML/CSS/JS)
│
├── 📁 samples/                  # ← Place your order & collection files here
├── 📁 reports/                  # ← Generated Excel reports saved here
│
├── 🗄️ finance_system.db        # SQLite database
│
├── 📁 database/                 # Database management modules
│   ├── db_manager.py
│   └── models.py
│
├── 📁 processors/               # Processing helper modules
│   ├── calculator.py
│   └── file_transformer.py
│
├── 📁 utils/                    # Utility modules
│   └── exporters.py
│
├── 📄 run.bat                   # One-click launcher (Windows)
├── 📄 requirements.txt          # Python dependencies
└── 📄 .gitignore
```

---

## 🏪 Supported Platforms

| Platform | File Type  | Detection Method             |
| -------- | ---------- | ---------------------------- |
| Amazon   | CSV / XLSX | Filename contains `amazon`   |
| Noon     | XLSX       | Filename contains `noon`     |
| Trendyol | XLSX       | Filename contains `trendyol` |
| Ilasouq  | XLSX       | Filename contains `ilasouq`  |
| Tabby    | XLSX       | Filename contains `tabby`    |
| SMSA     | XLSX       | Filename contains `smsa`     |

---

## 📊 Application Pages

| Page            | Function                                      |
| --------------- | --------------------------------------------- |
| 🏠 Dashboard    | KPI cards + platform table + new week button  |
| 📁 Upload Files | Drag & Drop upload for orders and collections |
| ⚙️ Process Data | Run full processing pipeline with live log    |
| 📊 Analytics    | Interactive charts (Chart.js)                 |
| 📈 Reports      | Download weekly Excel reports                 |

---

## 🔌 API Endpoints

| Endpoint               | Method | Description                        |
| ---------------------- | ------ | ---------------------------------- |
| `/`                    | GET    | Main dashboard page                |
| `/api/stats`           | GET    | Database statistics                |
| `/api/charts`          | GET    | Chart data                         |
| `/api/files`           | GET    | List uploaded files                |
| `/api/reports`         | GET    | List generated reports             |
| `/upload`              | POST   | Upload files                       |
| `/delete-file`         | POST   | Delete a file                      |
| `/process`             | POST   | Run full processing pipeline       |
| `/new-week`            | POST   | Start new week (clears DB + files) |
| `/download/<filename>` | GET    | Download a report                  |
| `/reset-db`            | POST   | Reset database only                |

---

## ⚙️ Requirements

- **Python** 3.8+
- **Windows** (for `run.bat` launcher)
- Internet connection on first run (to load Chart.js from CDN)

### Python Libraries

```
flask
pandas
openpyxl
xlsxwriter
xlrd
python-dateutil
pytz
```

Install all at once:

```bash
pip install -r requirements.txt
```

---

## 🛠️ Developer Guide

### Adding a New Platform

1. In `process_data.py` — add filename detection condition
2. In `process_collections.py` — add a collection processing function for the new platform
3. In `generate_report.py` — add a badge color for the platform in the table

### Customizing the Report

- `generate_report.py` controls everything: columns, colors, sheets
- Main function: `generate_weekly_report()`

---

## 📦 Tech Stack

| Layer    | Technology            |
| -------- | --------------------- |
| Backend  | Python + Flask        |
| Database | SQLite                |
| Frontend | HTML + CSS + JS       |
| Charts   | Chart.js (CDN)        |
| Reports  | openpyxl / xlsxwriter |
| Launcher | Windows Batch (.bat)  |

---

_Developed with ❤️ using Python + Flask + Chart.js_
