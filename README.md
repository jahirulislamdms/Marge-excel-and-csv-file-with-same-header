# 📊 Marge Excel and CSV File With Same Header

A powerful Python script that merges multiple Excel and CSV files based on matching headers.

Files with identical headers are merged into the same output CSV.  
Files with different headers generate separate merged output files.

---

## 🚀 Features

- Supports:
  - `.csv`
  - `.txt`
  - `.xlsx`
  - `.xls`
  - `.xlsm`
- Automatically groups files by identical headers
- Skips malformed rows (wrong column count or parsing errors)
- Copies problematic files to an `error_files/` folder
- Automatic encoding detection (optional `chardet`)
- Handles very large CSV field sizes
- Outputs clean UTF-8 CSV files
- GUI file picker if no command-line arguments provided
- Command-line usage supported

---

## 🛠 Requirements

```bash
pip install openpyxl
pip install chardet   # optional but recommended
```

---

## ▶️ Usage

### Option 1 — Command Line

```bash
python merge_excel_csv_by_header.py file1.csv file2.xlsx file3.csv
```

### Option 2 — File Picker (GUI)

```bash
python merge_excel_csv_by_header.py
```

A file selection window will open.

---

## 📂 Output Behavior

- Output files are saved in the **common parent directory** of selected files.
- Naming format:

```
merged_<header_name>_<hash>.csv
```

- Files containing malformed rows are copied to:

```
error_files/
```

---

## ⚙️ How It Works

1. Reads the first row as header.
2. Normalizes header values.
3. Groups files that share identical headers.
4. Writes merged data into UTF-8 CSV format.
5. Skips malformed rows.
6. Copies original problematic files into `error_files/`.

---

## 📄 Summary

After processing, the script prints:

- Total files processed
- Files with errors
- Number of rows added per merged file

---

## 📄 License

Free to use and modify.
