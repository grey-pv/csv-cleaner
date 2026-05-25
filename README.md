<div align="center">

# 🧹 CSV Cleaner Tool

**Clean messy CSV files automatically and get an Excel report of everything that was fixed.**

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Status](https://img.shields.io/badge/Status-Active-brightgreen)]()
[![PRs Welcome](https://img.shields.io/badge/PRs-Welcome-orange)](CONTRIBUTING.md)
[![pandas](https://img.shields.io/badge/Powered%20by-pandas-purple)](https://pandas.pydata.org/)

</div>

---

> **Messy CSVs are everywhere. Cleaning them by hand is a waste of time.**
> This tool fixes encoding issues, removes duplicates and empty rows, standardizes dates, trims whitespace, and exports a full Excel report of everything it did.

---

## 🎯 The Problem It Solves

Anyone who works with data knows the pain:

- CSV exported from a system with broken encoding (special characters show as `?` or `â€™`)
- 500-row file with 80 duplicate rows nobody noticed
- Date columns with three different formats in the same file
- Column names with random spaces and capitalization
- Empty rows scattered everywhere

This tool fixes all of that automatically — and tells you exactly what it changed.

**Who this is for:**
- 📊 Data analysts cleaning raw exports before analysis
- 🧾 Finance teams dealing with system-generated CSVs
- 🏢 Operations teams processing supplier or government data files
- 🐍 Python developers who want a clean automation script to study or extend

---

## ✨ Features

- ✅ **Auto-detects encoding** — UTF-8, Latin-1, CP1252, ISO-8859-1 and more
- ✅ **Removes empty rows** — fully blank rows are dropped automatically
- ✅ **Removes duplicate rows** — exact duplicates eliminated
- ✅ **Trims whitespace** — leading/trailing spaces removed from every cell
- ✅ **Standardizes dates** — DD/MM/YYYY, MM/DD/YYYY, and more → YYYY-MM-DD
- ✅ **Standardizes column names** — spaces, caps, special chars → clean snake_case
- ✅ **Excel report exported** — summary sheet + columns renamed sheet
- ✅ Handles **multiple files at once** — folder or explicit file list
- ✅ Zero configuration — works out of the box

---

## 🚀 Quick Start

### Requirements

- Python 3.10 or higher
- pip

### Installation

```bash
git clone https://github.com/grey-pv/csv-cleaner.git
cd csv-cleaner
pip install -r requirements.txt
```

### Usage

```bash
# Clean a single file
python clean_csv.py data.csv

# Clean multiple files
python clean_csv.py jan.csv feb.csv mar.csv

# Clean all CSVs in a folder
python clean_csv.py ./exports

# Custom output folder
python clean_csv.py data.csv --output-dir ./results
```

---

## 📸 Terminal Output

```
🧹 CSV Cleaner Tool
────────────────────────────────────────
  Found 1 CSV file(s) to clean...

  📄 messy_customers.csv
  ────────────────────────────────────
  Encoding : latin-1
  Rows     : 11 original
  ✓ Empty rows removed    : 2
  ✓ Duplicates removed    : 2
  ✓ Whitespace cells fixed: 5
  ✓ Dates standardized    : 6
  ✓ Column names fixed    : 6
  Final rows              : 7

  ✅ Cleaned : messy_customers_cleaned.csv
  📊 Report  : messy_customers_report.xlsx

────────────────────────────────────────
✅ All done! 1 file(s) cleaned.
   Duplicates removed : 2
   Empty rows removed : 2
   Whitespace fixed   : 5
   Dates standardized : 6
   Results saved to   : output/
────────────────────────────────────────
```

**Output per file:**
- `filename_cleaned.csv` — the clean version, UTF-8 encoded
- `filename_report.xlsx` — styled Excel report with summary and details

---

## ⚙️ CLI Options

| Option | Short | Description |
|---|---|---|
| `--output-dir FOLDER` | `-d` | Output folder (default: `output/`) |
| `--help` | `-h` | Show help and exit |

---

## 🗂️ Project Structure

```
csv-cleaner/
├── clean_csv.py             ← Main script — run this
├── generate_sample_csv.py   ← Creates a messy CSV for testing
├── requirements.txt         ← 2 dependencies only
├── README.md
├── samples/                 ← Place your input CSVs here
└── output/                  ← Cleaned files and reports land here
```

---

## 🧠 How It Works

```
Your messy CSV file
        │
        ▼
  Encoding detection    ← Tries UTF-8, Latin-1, CP1252... until it reads
        │
        ▼
  Column name cleanup   ← Strips spaces, lowercases, removes special chars
        │
        ▼
  Empty row removal     ← Drops rows where every cell is blank
        │
        ▼
  Duplicate removal     ← Drops exact duplicate rows
        │
        ▼
  Whitespace trimming   ← Strips leading/trailing spaces from every cell
        │
        ▼
  Date standardization  ← Converts date columns to YYYY-MM-DD format
        │
        ▼
  Export cleaned CSV    ← UTF-8 encoded, ready for analysis
        │
        ▼
  Export Excel report   ← Summary + columns renamed, fully formatted
```

---

## 🧪 Try It With a Sample File

```bash
python generate_sample_csv.py
python clean_csv.py samples/messy_customers.csv
```

The sample file has encoding issues, duplicate rows, empty rows, whitespace problems, mixed date formats, and messy column names — all planted intentionally so you can see every feature in action.

---

## 📦 Dependencies

| Package | Version | Purpose |
|---|---|---|
| [pandas](https://pandas.pydata.org/) | ≥ 2.0.0 | Reading, cleaning, and structuring data |
| [openpyxl](https://openpyxl.readthedocs.io/) | ≥ 3.1.0 | Excel report creation and formatting |

---

## 🔮 Roadmap

- [ ] **`--no-report` flag** — skip the Excel report for faster batch processing
- [ ] **Column type detection** — auto-cast numeric columns stored as text
- [ ] **Custom date format flag** — `--date-format "%d/%m/%Y"`
- [ ] **Merge cleaned files** — combine output into one master CSV
- [ ] **Streamlit UI** — drag & drop browser interface
- [ ] **JSON and TSV support**

---

## ❓ FAQ

**Does it modify the original file?**
No. It only reads the source and writes new files to the output folder.

**What date formats does it detect?**
DD/MM/YYYY, MM/DD/YYYY, DD-MM-YYYY, MM-DD-YYYY, DD/MM/YY, and more. All are converted to YYYY-MM-DD.

**What if my file has a completely different encoding?**
Open an issue and attach the file — adding new encodings is straightforward.

**Can it handle files with thousands of rows?**
Yes. pandas handles large files efficiently. Memory usage depends on file size.

**Is this free to use commercially?**
Yes. MIT license.

---

## 🤝 Contributing

PRs are welcome. If the cleaner misses something in your file, open an issue and attach a sample — that's the fastest way to improve it.

---

## 📄 License

MIT — free to use, modify, and distribute.

---

## 👤 Author

Built by **Paulo Victor (grey)** — shipping Python automation tools in public.

- Twitter/X: [@grey]
- LinkedIn: [linkedin.com/in/paulo-victor-f-gonçalves-6b0b26318]
- GitHub: [@grey-pv](https://github.com/grey-pv)

---

<div align="center">

**Part of the [Python Automation Toolkit](https://github.com/grey-pv) — open source tools that eliminate repetitive manual work.**

⭐ Star the repo if this saved you time.

</div>
