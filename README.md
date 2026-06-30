# EXCEL AI AGENT — Profit / Loss Reporting

An AI-agent–driven reporting tool that reads a Merrill brokerage Excel workbook
and produces a formatted **Profit / Loss PDF report** for stock holdings.

## What it does

For each holding (rows flagged with `f` in column A of the `Merrill` sheet) the
agent captures the ticker, share count, unit price, current price, total cost
and current value, then computes **Profit / Loss** and **% change**. The report
is rendered as a styled landscape PDF with a losing-vs-winning summary section.

## Reports

| Folder      | Script                     | Purpose                                  |
|-------------|----------------------------|------------------------------------------|
| `Report_1/` | `generate_report.py`       | Full holdings P/L report with summary    |
| `Report_2/` | `generate_indiv_report.py` | Individual holdings report               |

Generated PDFs are written to the `output/` directory.

## Usage

```bash
cd Report_1
python generate_report.py        # default: ordered ascending by Profit / Loss
python generate_report.py OBT    # OBT = Order By Ticker (alphabetical)
```

## Requirements

- Python 3
- [`openpyxl`](https://pypi.org/project/openpyxl/) — read the Excel workbook
- [`reportlab`](https://pypi.org/project/reportlab/) — generate the PDF

```bash
pip install openpyxl reportlab
```

## Configuration

The source workbook path and sheet are defined at the top of each script:

```python
EXCEL_PATH = r"C:\Users\johna\Documents\MerrylAccount.xlsx"
SHEET_NAME = "Merrill"
```

The Excel file is always opened **read-only**.

## Notes

- The workbook (`*.xlsx`) and generated PDFs are git-ignored — they contain
  personal account data and are regenerated artifacts.
