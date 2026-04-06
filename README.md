# Budget Query Automation

Automates the monthly PeopleSoft budget report process for SBSC. Combines `OU_BUD_ORG` and `OU_BUD_SOURCE` query exports into a single formatted Excel file with the correct naming convention, sorting, and formatting.

## What It Does

- Reads both PeopleSoft query exports (`.xlsx`)
- Skips the PeopleSoft title row and reads the real headers automatically
- Excludes `OUFND` from `OU_BUD_SOURCE` to prevent duplicate rows
- Applies the required custom sort to each query before combining
- Outputs a single formatted Excel file with:
  - `Org Budget Inquiry` label in A1
  - Retrieved date in B1 (bold, yellow)
  - Bold headers with auto-filter on row 2
  - Accounting-style USD formatting on dollar columns
  - Parent rows highlighted in light green
  - Auto-fit column widths
  - File named `P## Month Year - Department.xlsx`

## Requirements

- Python 3.x — [python.org](https://python.org/downloads)
- pandas and openpyxl

```
py -m pip install pandas openpyxl
```

## Usage

```
py budget_query_automation.py
```

1. File picker opens — select `OU_BUD_ORG` export
2. File picker opens — select `OU_BUD_SOURCE` export
3. Enter the department name (ex: `Biology`)
4. Excel file is saved in the same directory as the script

## Output Filename Format

```
P## Month Year - Department.xlsx
```

The period and month are automatically calculated based on the previous month.
