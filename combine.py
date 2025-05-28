import sys
import subprocess

# === DEPENDENCY CHECK AND AUTO-INSTALL ===
REQUIRED_MODULES = ['pandas', 'openpyxl']
for module in REQUIRED_MODULES:
    try:
        __import__(module)
    except ImportError:
        print(f"📦 Missing dependency: {module} — attempting to install it...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", module])
        except subprocess.CalledProcessError:
            print(f"❌ Failed to install {module}. Please install it manually.")
            sys.exit(1)

# === MAIN IMPORTS ===
import argparse
import pandas as pd
from pathlib import Path
import re

# === CONFIGURATION ===
DEFAULT_INPUT_FOLDER = "relatorios"
DEFAULT_OUTPUT_FILE = "combined_report.csv"
FILE_EXTENSION = "*.xlsx"

# === MESSAGES ===
MSG_NO_FILES_FOUND = "⚠️  No files with extension '{ext}' found in folder '{folder}'."
MSG_READING_FILE = "Reading: {filename}"
MSG_FILE_SAVED = "✅ Combined file saved to: {output_path}"
MSG_SKIPPED_EMPTY = "⚠️  Skipped empty file: {filename}"

# === MOJIBAKE FIX ===
MOJIBAKE_PATTERN = re.compile(r"[ÃÂ¢§]")

def fix_encoding_issues(df):
    def maybe_fix(val):
        if isinstance(val, str) and MOJIBAKE_PATTERN.search(val):
            try:
                return val.encode("latin1", errors="strict").decode("utf-8", errors="strict")
            except Exception:
                return val
        return val

    for col in df.columns:
        df[col] = df[col].apply(maybe_fix)

    return df

# === REMOVE NEWLINES FROM TEXT FIELDS ===
def normalize_text_fields(df):
    for col in df.columns:
        df[col] = df[col].str.replace(r'\s*[\r\n]+\s*', ' ', regex=True).str.strip()
    return df

# === FIX TIME FORMATTING FROM "0 days HH:MM:SS" ===
def clean_time_columns(df):
    for col in df.columns:
        df[col] = df[col].str.extract(r'(\d{1,2}:\d{2}:\d{2})', expand=False).fillna(df[col])
    return df

# === MAIN FUNCTION ===
def combine_spreadsheets(folder, output_file):
    path = Path(folder)
    xlsx_files = list(path.glob(FILE_EXTENSION))

    if not xlsx_files:
        print(MSG_NO_FILES_FOUND.format(ext=FILE_EXTENSION, folder=folder))
        return

    dataframes = []
    for file in xlsx_files:
        print(MSG_READING_FILE.format(filename=file.name))
        
        # ✅ Import all as string and clean
        df = pd.read_excel(file).astype(str)
        df.replace(['nan', 'NaT', 'NaN', 'None'], '', inplace=True)

        if df.empty:
            print(MSG_SKIPPED_EMPTY.format(filename=file.name))
            continue

        df = fix_encoding_issues(df)
        df = normalize_text_fields(df)

        # === CUSTOM ===
        if len(df) >= 3:
            df = df.iloc[:-3]

        df = clean_time_columns(df)

        dataframes.append(df)

    combined = pd.concat(dataframes, ignore_index=True)

    # 💥 REMOVE DUPLICATES
    combined.drop_duplicates(inplace=True)

    combined.to_csv(output_file, index=False, encoding='utf-8-sig')
    print(MSG_FILE_SAVED.format(output_path=output_file))

# === CLI ===
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Combine all spreadsheet files from a folder into a single CSV file."
    )
    
    parser.add_argument(
        "--folder",
        type=str,
        default=DEFAULT_INPUT_FOLDER,
        help=f"Folder containing {FILE_EXTENSION} files (default: {DEFAULT_INPUT_FOLDER})"
    )
    
    parser.add_argument(
        "--output",
        type=str,
        default=DEFAULT_OUTPUT_FILE,
        help=f"Output CSV file name (default: {DEFAULT_OUTPUT_FILE})"
    )

    args = parser.parse_args()
    combine_spreadsheets(args.folder, args.output)
