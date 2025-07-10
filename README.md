# Prepayment Automation Tool

A Python script to automate the generation of accounting entries for prepaid item amortization. This tool helps accountants process monthly prepayment schedules and generate the required journal entries efficiently.

## Overview

Companies use prepayment schedules to track the amortization of prepaid items. At each month-end, accountants need to record journal entries to recognize expenses for that month. This tool automates that process by reading prepayment data from Excel files and generating standardized accounting entries.

### Example Output

For "Webhosting" in May 2024:

```
Date       | Description                                  | Reference | Account | Amount
31/5/2024  | Prepayment amortisation for Webhosting     | 46248     | EXP001  | 833.33
31/5/2024  | Prepayment amortisation for Webhosting     | 46248     | PRE001  | -833.33
```

## Installation

1. **Clone or download the project files**
2. **Create Python virtual environment**:

   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install required dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## Configuration

The tool uses the following configuration in `config.py`:

```python
EXPENSE_ACCOUNT = "EXP001"        # Account code for expense entries
PREPAYMENT_ACCOUNT = "PRE001"     # Account code for prepayment entries
SHEET_NAME = "Schedule"           # Excel sheet name containing prepayment data
REPAYMENT_PERIOD = 12            # Amortization period in months
INPUT_FILE = "input/sample.xlsx"  # Input Excel file path
```

You can modify these values to match your organization's chart of accounts and business rules.

## Usage

### Interactive Mode (Default)

```bash
python main.py
```

The tool will prompt you for:

- Month/year (in MM/YY format, e.g., 05/24)
- Export format (excel or csv)

### Command Line Mode

```bash
# Generate entries for May 2024 in Excel format
python main.py --month 05/24 --format excel

# Generate entries for December 2023 in CSV format
python main.py --month 12/23 --format csv
```

### Input File Structure

Place your prepayment data in an Excel file in the `input/` directory. The expected structure is:

- **File location**: Define file path as `INPUT_FILE` in `config.py`
- **Sheet name**: Define file path as `SHEET_NAME` in `config.py`
- **Headers** (row 3):
  - `Items`: Description of prepaid items
  - `Invoice number`: Reference number for the invoice
  - `Invoice amount`: Total prepaid amount
  - Date columns: One column per month (e.g., "2024-05-01" for May 2024)

## Project Structure

```
backbone/
├── main.py              # Main entry point and CLI interface
├── processing.py        # Core business logic for processing prepayments
├── helpers.py          # Utility functions (date parsing, file I/O)
├── config.py           # Configuration settings
├── requirements.txt    # Python dependencies
├── input/             # Input Excel files
│   └── sample.xlsx
└── output/            # Generated output files
    ├── result_05-2024.csv
    └── result_05-2024.xlsx
```

## Output Files

Generated files are saved in the `output/` directory with the naming convention:

- `result_MM-YYYY.xlsx` for Excel format
- `result_MM-YYYY.csv` for CSV format

Each file contains columns:

- **Date**: Last day of the target month (DD/M/YYYY format)
- **Description**: Descriptive text for the journal entry
- **Reference**: Invoice number from the source data
- **Account**: Account code (EXP001 or PRE001)
- **Amount**: Calculated monthly amortization amount

## Technical Requirements

- Python 3.7+
- pandas 2.3.1+
- openpyxl 3.1.5+
- numpy 2.3.1+

## License

This project was developed as part of an internship assessment for Backbone and is intended for educational and evaluation purposes.

## Contact

For questions or issues, please contact the development team.
