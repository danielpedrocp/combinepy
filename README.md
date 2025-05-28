# 🍰 combinepy

*A piece of cake CLI for terminal Excel file merging*

> Combine your `.xlsx` files into a single, clean, deduplicated CSV — in one line. Simple, just like data prep should be.


## 🚀 What is it?

**combinepy** is a simple command-line utility to:

- 📥 Load Excel files from a folder
- 🔄 Clean formatting issues (mojibake, `NaN`, `NaT`, time garbage like "0 days")
- 🧼 Normalize text fields and remove linebreak noise
- 🗑️ Remove duplicated records (100% match)
- 📤 Export everything as a UTF-8 encoded CSV, ready for analysis



## 🛠️ Usage

```bash
python combine.py --folder relatorios --output combined.csv
```

### Optional arguments:
- `--folder`: directory containing `.xlsx` files (default: `relatorios`)
- `--output`: name of the CSV to be created (default: `combined_report.csv`)


## 📦 Dependencies

- `pandas`
- `openpyxl`

They are **automatically installed** if missing when the script runs.


## ✅ Features

- ✅ Handles garbled encoding (`Ã§`, `Ã£`, etc.)
- ✅ Treats all cells as strings — no accidental datetimes, `NaN`, or `NaT`
- ✅ Normalizes line breaks (`\n`, `\r`, etc.) into single-line text
- ✅ Removes trailing metadata/footer rows (last 3 lines per file)
- ✅ Removes exact duplicates from merged files
- ✅ Compatible with Windows terminals and Excel CSV import


## ⚠️ Limitations

- ⚠️ Assumes all `.xlsx` files share the **same column structure**
- ⚠️ Processes all `.xlsx` files in the folder — no filtering or date range
- ⚠️ Still a single-script tool (but built to be extended)


## 🧪 Example use cases

- Merging daily crime reports
- Joining audit exports
- Cleaning messy `.xlsx` logs
- Prepping structured data for BI dashboards or SQL pipelines


## ✨ Future ideas

- `--log` mode: generate a log of the process
- `--dry-run` mode: preview actions without writing output
- `--subset` option: deduplicate based on specific columns
- `--stats` command: generate summary insights