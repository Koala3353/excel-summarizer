# Employee and Job Order Excel Summary Scripts

## Overview

This repository contains two Python scripts:

- **script-employee.py**  
  Summarizes employee salary data per week (or date range) from multiple Excel files.

- **script-job.py**  
  Summarizes job order expenses per week (or date range) from multiple Excel files.

Both scripts are designed to process a batch of Excel files and generate a summary Excel file with columns for each week/date and a total column.

---

## Folder Structure

Before running the scripts, **create a folder named `data`** in the same directory as the scripts.  
Place all your Excel files (`.xlsx` or `.xlsm`) to be processed inside this `data` folder.

```
your_project/
│
├── script-employee.py
├── script-job.py
├── data/
│   ├── 14 APR. 2-8, 2025.xlsm
│   ├── 15 APR. 9-15, 2025.xlsm
│   └── ... (other Excel files)
```

---

## How the Scripts Work

### Common Features

- **Reads all Excel files** in the `data` folder.
- **Extracts the week label** from cell **C1** of the relevant sheet (`Input` for employees, `Bossing` for job orders).  
  If C1 is empty or missing, the script falls back to parsing the filename (ignoring any leading number).
- **Aggregates data** per week/date and per entity (employee or job order).
- **Outputs a summary Excel file** with columns for each week/date and a total column.
- **Logs all actions** to a log file in the parent directory.

---

### script-employee.py

**Purpose:**  
Aggregates employee salary data per week/date from the "Input" sheet of each Excel file.

**How it works:**
- For each file, it tries to read the week label from cell C1 of the "Input" sheet.
- If C1 is missing, it parses the filename (e.g., `14 APR. 2-8, 2025.xlsm` → `APR 2-8`).
- It locates the salary column and sums salaries per employee for each week.
- Outputs `EmployeeSummary.xlsx` with columns:  
  `Employee Name | <Week1> | <Week2> | ... | Total`

**Log file:**  
`employee_script_log.txt`

---

### script-job.py

**Purpose:**  
Aggregates job order expenses per week/date from the "Bossing" sheet of each Excel file.

**How it works:**
- For each file, it tries to read the week label from cell C1 of the "Bossing" sheet.
- If C1 is missing, it parses the filename (e.g., `14 APR. 2-8, 2025.xlsm` → `APR 2-8`).
- It sums expenses per job order for each week.
- Outputs `Summary.xlsx` with columns:  
  `Job Order | <Week1> | <Week2> | ... | Total`

**Log file:**  
`job_script_log.txt`

---

## Requirements

- Python 3.7+
- Packages: `pandas`, `openpyxl`

Install requirements with:
```
pip install pandas openpyxl
```

---

## Usage

1. **Place your Excel files in the `data` folder.**
2. **Run the scripts:**
   ```
   python script-employee.py
   python script-job.py
   ```
3. **Check the output Excel files** (`EmployeeSummary.xlsx`, `Summary.xlsx`) and log files in the parent directory.

---

## Notes

- The scripts are robust to missing or malformed files and will log any issues encountered.
- Week/date columns are dynamically created based on the C1 cell or filename of each file.
- The scripts create a `debug` folder with raw extracted sheets for troubleshooting.

---
