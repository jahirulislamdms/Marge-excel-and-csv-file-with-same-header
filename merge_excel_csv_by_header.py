#!/usr/bin/env python3
"""
merge_excel_csv.py

Merge multiple Excel (.xlsx/.xls) and CSV (.csv/.txt) files by header.

Behavior:
- First row of each file is treated as header.
- Files that share the same header are merged into a single CSV.
- Files with different headers produce separate merged CSV outputs.
- If any line/row in a file is malformed (wrong column count or CSV parse error),
  that row is skipped and the original file is copied into `error_files/`.
- All outputs are saved as UTF-8 CSVs into the common parent folder of inputs.

Notes:
- For reading Excel files this script uses openpyxl.
- For better encoding detection it optionally uses chardet.
- It raises CSV field size limit to avoid "field larger than field limit" errors.

Requirements:
    pip install openpyxl
    pip install chardet   # optional but recommended

Usage:
    python app.py file1.csv file2.xlsx ...
    # or
    python app.py            # opens a file selection dialog (tkinter required)
"""
from pathlib import Path
import argparse
import csv
import sys
import os
import shutil
import hashlib
from typing import List, Tuple, Iterable

# Optional dependencies
try:
    import openpyxl
except Exception:
    openpyxl = None

try:
    import chardet
except Exception:
    chardet = None

# Increase CSV field size limit to avoid field-too-large errors
try:
    csv.field_size_limit(sys.maxsize)
except OverflowError:
    # Some platforms can't accept sys.maxsize
    try:
        csv.field_size_limit(10 ** 9)
    except Exception:
        pass
except Exception:
    pass


def detect_encoding(path: Path, sample_size: int = 131072) -> str:
    """
    Try to detect file encoding. If chardet is available, use it on a sample.
    Otherwise try a list of common encodings and return the first that works.
    Fallback: 'utf-8'.
    """
    if chardet:
        try:
            with path.open('rb') as f:
                raw = f.read(sample_size)
            result = chardet.detect(raw)
            enc = result.get('encoding') or 'utf-8'
            return enc
        except Exception:
            return 'utf-8'

    # fallback: try common encodings
    for enc in ('utf-8-sig', 'utf-8', 'cp1252', 'latin1', 'iso-8859-1'):
        try:
            with path.open('r', encoding=enc, errors='strict') as f:
                f.read(2048)
            return enc
        except Exception:
            continue
    return 'utf-8'


def sanitize_header(header: List[str]) -> Tuple[str, int]:
    norm = [str(h).strip() for h in header]
    key = "|".join(norm)
    return key, len(norm)


def sanitize_filename(s: str, max_len: int = 200) -> str:
    s2 = s.lower()
    out = []
    for ch in s2:
        if ch.isalnum() or ch in " _-":
            out.append(ch)
        else:
            out.append("_")
    name = "".join(out).replace(" ", "_")
    if len(name) > max_len:
        name = name[:max_len]
    while "__" in name:
        name = name.replace("__", "_")
    name = name.strip("_")
    return name or "merged"


def read_csv_rows(path: Path) -> Tuple[List[str], Iterable[List[str]], bool]:
    """
    Read header and a generator of rows from CSV file.
    Returns: (header_list, rows_generator, had_error)
    - had_error is True if any row was skipped or a parse/decoding error occurred.
    """
    had_error = False
    # detect encoding and try strict first then replace
    enc = detect_encoding(path)

    def _open_reader(encoding: str, errors: str):
        f = open(path, "r", encoding=encoding, errors=errors, newline="")
        reader = csv.reader(f)
        return f, reader

    # Attempt strategies: (detected encoding, strict) then (detected encoding, replace)
    attempts = [(enc, "strict"), (enc, "replace")]
    for encoding_try, errors_try in attempts:
        try:
            f, reader = _open_reader(encoding_try, errors_try)
            try:
                header = next(reader)
            except StopIteration:
                f.close()
                return [], iter([]), had_error
            header_clean = [h.strip() for h in header]
            expected_cols = len(header_clean)

            def gen():
                nonlocal had_error
                line_no = 1
                try:
                    for row in reader:
                        line_no += 1
                        if len(row) == 0 or all((c is None or str(c).strip() == "") for c in row):
                            continue
                        if len(row) != expected_cols:
                            had_error = True
                            print(
                                f"Skipping malformed row {line_no} in {path.name}: expected {expected_cols} columns, found {len(row)}"
                            )
                            continue
                        yield [c for c in row]
                except csv.Error as e:
                    had_error = True
                    print(f"CSV parsing error in {path.name} at line {line_no}: {e}")
                except Exception as e:
                    had_error = True
                    print(f"Unexpected error reading {path.name} at line {line_no}: {e}")
                finally:
                    try:
                        f.close()
                    except Exception:
                        pass

            return header_clean, gen(), had_error
        except UnicodeDecodeError:
            # try next attempt
            continue
        except csv.Error as e:
            # CSV module error while opening/reading
            had_error = True
            print(f"csv.Error while opening {path.name}: {e}")
            # try next attempt (replace)
            continue
        except Exception as e:
            had_error = True
            print(f"Error reading CSV {path.name}: {e}")
            return [], iter([]), had_error

    return [], iter([]), True


def read_excel_rows(path: Path) -> Tuple[List[str], Iterable[List[str]], bool]:
    """
    Read header and rows from an Excel file using openpyxl.
    Returns: (header_list, rows_generator, had_error)
    """
    if openpyxl is None:
        raise RuntimeError("openpyxl is required to read Excel files. Install with: pip install openpyxl")
    had_error = False
    try:
        wb = openpyxl.load_workbook(filename=str(path), read_only=True, data_only=True)
        ws = wb.worksheets[0]
        rows = ws.iter_rows(values_only=True)
        try:
            header_row = next(rows)
        except StopIteration:
            wb.close()
            return [], iter([]), had_error
        header = [str(c).strip() if c is not None else "" for c in header_row]
        expected_cols = len(header)

        def gen():
            nonlocal had_error
            line_no = 1
            try:
                for row in rows:
                    line_no += 1
                    row_list = ["" if c is None else str(c) for c in row]
                    if len(row_list) != expected_cols:
                        had_error = True
                        print(
                            f"Skipping malformed row {line_no} in {path.name}: expected {expected_cols} columns, found {len(row_list)}"
                        )
                        continue
                    yield row_list
            except Exception as e:
                had_error = True
                print(f"Error while iterating rows in {path.name} at row {line_no}: {e}")
            finally:
                try:
                    wb.close()
                except Exception:
                    pass

        return header, gen(), had_error
    except Exception as e:
        print(f"Error reading Excel {path.name}: {e}")
        return [], iter([]), True


def write_row_to_csv(path: Path, header: List[str], rows: Iterable[List[str]], append: bool = True) -> int:
    """
    Write header (if file is created) and rows into path.
    Returns number of rows written (not counting header).
    """
    mode = "a" if append and path.exists() else "w"
    path.parent.mkdir(parents=True, exist_ok=True)
    written = 0
    with path.open(mode, newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if mode == "w":
            writer.writerow(header)
        for row in rows:
            writer.writerow(row)
            written += 1
    return written


def process_files(files: List[Path]):
    if not files:
        print("No files provided.")
        return

    parents = [str(p.parent) for p in files]
    common_parent = Path(os.path.commonpath(parents))
    output_dir = common_parent
    print(f"Saving merged outputs into: {output_dir}")

    error_dir = output_dir / "error_files"
    error_dir.mkdir(parents=True, exist_ok=True)

    stats = {"files_processed": 0, "files_with_errors": [], "merged_files": {}}

    for path in files:
        stats["files_processed"] += 1
        print(f"\nProcessing: {path}")
        if not path.exists():
            print(f"  Skipping missing file: {path}")
            continue

        ext = path.suffix.lower()
        if ext in (".csv", ".txt"):
            header, rows_iter, had_error = read_csv_rows(path)
        elif ext in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"):
            header, rows_iter, had_error = read_excel_rows(path)
        else:
            print(f"  Unsupported file type {ext}, skipping.")
            continue

        if not header:
            print(f"  Empty or unreadable header in {path.name}, skipping file.")
            if had_error:
                dest = error_dir / path.name
                if dest.exists():
                    i = 1
                    while True:
                        newname = dest.with_name(f"{dest.stem}_{i}{dest.suffix}")
                        if not newname.exists():
                            dest = newname
                            break
                        i += 1
                try:
                    shutil.copy2(path, dest)
                    stats["files_with_errors"].append(path)
                    print(f"  Copied original {path.name} to {dest}")
                except Exception as e:
                    print(f"  Failed to copy file with errors: {e}")
            continue

        header_key, expected_cols = sanitize_header(header)
        safe_name = sanitize_filename(header_key)
        hhash = hashlib.sha1(header_key.encode("utf-8")).hexdigest()[:8]
        out_name = f"merged_{safe_name}_{hhash}.csv"
        out_path = output_dir / out_name
        first_time = not out_path.exists()

        try:
            rows_written = write_row_to_csv(out_path, header, rows_iter, append=not first_time)
        except Exception as e:
            print(f"  Failed to write rows from {path.name} into {out_path.name}: {e}")
            had_error = True
            rows_written = 0

        print(f"  Merged into {out_path.name} (rows added: {rows_written})")
        stats["merged_files"].setdefault(out_path.name, 0)
        stats["merged_files"][out_path.name] += rows_written

        if had_error:
            dest = error_dir / path.name
            if dest.exists():
                i = 1
                while True:
                    newname = dest.with_name(f"{dest.stem}_{i}{dest.suffix}")
                    if not newname.exists():
                        dest = newname
                        break
                    i += 1
            try:
                shutil.copy2(path, dest)
                stats["files_with_errors"].append(path)
                print(f"  Copied original {path.name} to {dest}")
            except Exception as e:
                print(f"  Failed to copy file with errors: {e}")

    # Summary
    print("\n=== Summary ===")
    print(f"Files processed: {stats['files_processed']}")
    if stats["files_with_errors"]:
        print(f"Files with errors (copied to {error_dir}):")
        for p in stats["files_with_errors"]:
            print(f"  - {p.name}")
    else:
        print("No files had malformed rows.")

    print("Merged files produced in:")
    for name, rows in stats["merged_files"].items():
        print(f"  - {name} (rows added: {rows})")


def choose_files_via_dialog() -> List[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        print("tkinter is not available on this system. Please supply file paths as command-line arguments.")
        return []
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="Select CSV/Excel files to merge",
        filetypes=[("CSV/Excel files", "*.csv *.txt *.xlsx *.xls *.xlsm *.xltx *.xltm"), ("All files", "*.*")],
    )
    root.update()
    return [Path(p) for p in file_paths]


def main():
    parser = argparse.ArgumentParser(description="Merge multiple CSV/Excel files that share the same header into CSV outputs.")
    parser.add_argument("files", nargs="*", help="Files to merge (CSV/Excel). If omitted, a file chooser dialog opens).")
    args = parser.parse_args()

    if args.files:
        files = [Path(f) for f in args.files]
    else:
        files = choose_files_via_dialog()
        if not files:
            print("No files selected. Exiting.")
            sys.exit(0)

    process_files(files)


if __name__ == "__main__":
    main()
