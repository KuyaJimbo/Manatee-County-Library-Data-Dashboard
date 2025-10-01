# Library Dashboard Data Refresh Script Documentation

## Overview

This repository contains a Python automation script that consolidates raw monthly Excel reports from library branches into a single **master dataset** (`MasterDataset.xlsx`). The processed dataset is then used to power **Power BI dashboards**, enabling streamlined analysis of library system performance.

The script is necessary because the raw Excel files stored in the **Statistics folder** are **not formatted for a simple ETL process**. Each workbook uses custom layouts, merged cells, and scattered metrics, making them difficult to process directly in Power BI. This script extracts, organizes, and cleans the data into consistent, analysis-ready tables.

## Features
* **Automated ETL**: Reads multiple fiscal-year folders (`October YYYY - September YYYY`) and extracts metrics across all branches.
* **Dynamic Branch Detection**: Automatically identifies branch-level workbooks ending in Branch.xlsx.
* **Data Consolidation**: Creates a **single standardized master dataset** with multiple worksheets:
   * General Statistics
   * Programming
   * Digital Information
   * Interlibrary Loan (ILL)
   * Tech Statistics (and Part 2)
   * Computer & Study Room Usage
   * Branch Legend
* **Error Handling**: Skips problematic rows safely and reports issues without crashing.

* **Good Practices**:

   * Modular function design (clear separation of data extraction, cleaning, and processing).
   * Configurable mappings (cell references and worksheet headers defined at the top of the script).
   * Dynamic fiscal year handling and month-to-date string formatting.
   * Uses openpyxl with data_only=True to extract actual values instead of formulas.

## File Structure

```
├── Library Data Dashboard/
│     ├── Refresh Dashboard Dataset.py
│     ├── Create New Branch.py
│     ├── Internal_Library_Dashboard.pbix
│     ├── Public_Library_Dashboard.pbix
│     ├── MasterDataset.xlsx
│     └── ...
├── October 2023 - September 2024/
│     ├── <Branch Name> Branch.xlsx
│     ├── ILL.xlsx
│     ├── Digital Information.xlsx
│     ├── Tech Statistics.xlsx
│     ├── Summary Usage Report.xlsx
│     └── ...
├── October 2024 - September 2025/
│     └── ...
...
```

## Key Functions

####  Utility Functions

* `detect_branch_files(folder_path)` → Finds all branch Excel files in a fiscal year folder.
* `get_month_number_from_name(month_name)` → Converts month names to numeric values.
* `get_date_string(year, month_name)` → Formats a standard date string (`YYYY-MM-01 00:00:00`).
* `clean_data_row(row)` → Replaces `None` with `0` and validates numeric columns.
* `safe_append_row(...)` → Appends rows safely with error handling.

#### Dataset Creation

* `create_master_dataset()` → Initializes a new workbook with all worksheets and headers.
* `populate_legend_worksheets()` → Builds the Branch Legend tab.

#### Data Extraction

* `extract_general_statistics(...)` → Pulls branch-level statistics like patrons, hours, references.
* `extract_programming_data(...)` → Extracts program counts and attendance by category/age group.
* `extract_digital_info(...)` → Extracts system-wide circulation and digital usage data.
* `extract_ill_data(...)` → Extracts interlibrary loan borrowed/supplied counts.
* `extract_tech_statistics(...)` / extract_tech_statistics_pt2(...) → Extracts checkouts, check-ins, and PAC/Reserve usage.
* `extract_computer_study_room_usage(...)` → Extracts study/computer room usage metrics.

#### Processing Functions

* `process_library_file(...)` → Handles branch-level Excel files.
* `process_digital_info_file(...)` → Handles Digital Information workbooks.
* `process_tech_stats_file(...)` → Handles Tech Statistics workbooks.
* `process_library_usage(...)` → Handles study room usage data.

#### Execution

* `main()` → Orchestrates the entire ETL pipeline:
   1. Deletes old MasterDataset.xlsx (if exists).
   2. Iterates through all fiscal-year folders.
   3. Processes branch/system workbooks.
   4. Compiles all results into a single output file.

### Step 1: Export Data from LibCal

1. **Access LibCal**
   - Log into your LibCal account
   - Navigate to: **Events Tab > Manatee Library Events > Event Explorer**

2. **Configure Export Filters**
   - **Date Range**: From 2022-01-01 To 2025-09-23 (or current date)
   - **Show Registration Responses**: No
   - Click **"Submit"** to apply filters

3. **Export Data**
   - Click **"Export Data"**
   - Save file with naming convention: `lc_events_[timestamp].csv`

## Requirements

* Python 3.9+
* `openpyxl` library
```
pip install openpyxl
```

## Usage
1. Place this script (Refresh Dashboard Dataset.py) in the Dashboard folder, one level above all fiscal-year folders.
2. Run:
```
python Refresh\ Dashboard\ Dataset.py
```
3. The script generates `MasterDataset.xlsx` in the same folder.
4. Import `MasterDataset.xlsx` into **Power BI** for reporting.

## Why This Matters for Power BI

Power BI struggles with messy Excel structures (merged cells, multiple tables per sheet, and non-tabular metrics). This script acts as the ETL layer:
* Converts **branch-level** reports into normalized rows.
* Ensures **consistent column headers** across fiscal years.
* Produces a **single master dataset** ready for ingestion into dashboards.

## Good Practices in the Code

* `Separation of Concerns`: Extraction, cleaning, and loading are handled by dedicated functions.
* `Reusability`: Configurable mappings at the top mean minimal code changes if Excel formats evolve.
* `Data Safety`: Built-in exception handling (safe_append_row) prevents incomplete files from breaking the process.
* `Scalability`: Supports multiple fiscal year folders automatically.