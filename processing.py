
import datetime
from helpers import read_excel, get_last_day_of_month
import pandas as pd
from config import EXPENSE_ACCOUNT, PREPAYMENT_ACCOUNT, REPAYMENT_PERIOD

def process_prepayments(file_path: str, sheet_name: str, target_month: datetime.date) -> pd.DataFrame:
    """Process prepayments from an Excel file and return a DataFrame with entries."""
    # Read the Excel file with formulas
    df_formulas = read_excel(file_path, sheet_name, header_row=2)

    # Find the target month column
    target_column = None
    for col in df_formulas.columns:
        try:
            # Try to parse the column name as a date
            col_date = pd.to_datetime(col).date()

            # Check if year and month match the target
            if col_date.year == target_month.year and col_date.month == target_month.month:
                target_column = col
                break
        except (ValueError, TypeError):
            # Skip columns that can't be parsed as dates
            continue
    
    if target_column is None:
        print(f"No column found for target month: {target_month}")
        return pd.DataFrame()

    # Filter out rows where Items is NaN
    df_clean = df_formulas.dropna(subset=["Items"]).copy()
    
    if df_clean.empty:
        print("No valid data found after filtering NaN")
        return pd.DataFrame()
    
    # Create prepayment entries
    df_prepayment = df_clean.copy()
    df_prepayment["Amount"] = (df_clean["Invoice amount"] / REPAYMENT_PERIOD).round(2)
    df_prepayment["Account"] = PREPAYMENT_ACCOUNT
    df_prepayment["Description"] = "Prepayment for " + df_clean["Items"].astype(str)
    df_prepayment["Reference"] = df_clean["Invoice number"].astype(int).astype(str)
    df_prepayment["Date"] = get_last_day_of_month(target_month)
    
    # Create expense entries
    df_expense = df_clean.copy()
    df_expense["Amount"] = (-(df_clean["Invoice amount"] / REPAYMENT_PERIOD)).round(2)
    df_expense["Account"] = EXPENSE_ACCOUNT
    df_expense["Description"] = "Expense for " + df_clean["Items"].astype(str)
    df_expense["Reference"] = df_clean["Invoice number"].astype(int).astype(str)
    df_expense["Date"] = get_last_day_of_month(target_month)
    
    # Intertwine DataFrames - group prepayment and expense for each item together
    df_combined_list = []
    for idx in df_clean.index:
        # Add prepayment entry for this item
        df_combined_list.append(df_prepayment.loc[idx][["Date", "Description", "Reference", "Account", "Amount"]])
        # Add expense entry for this item
        df_combined_list.append(df_expense.loc[idx][["Date", "Description", "Reference", "Account", "Amount"]])
    
    # Create final DataFrame
    df_combined = pd.DataFrame(df_combined_list)
    
    return df_combined