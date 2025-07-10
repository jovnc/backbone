from processing import process_prepayments
from config import SHEET_NAME, INPUT_FILE
import os
import argparse
from helpers import parse_month_year, save_dataframe, get_user_input

def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Process prepayments and generate invoices')
    parser.add_argument('--month', type=str, help='Month/year in MM/YY format (e.g., 05/24)')
    parser.add_argument('--format', type=str, choices=['excel', 'csv'], help='Export format (excel or csv)')
    parser.add_argument('--interactive', action='store_true', default=True, help='Run in interactive mode (default)')
    
    args = parser.parse_args()
    
    # If command line arguments are provided, use them; otherwise ask user
    if args.month and args.format:
        try:
            target_date = parse_month_year(args.month)
            export_format = args.format
            print(f"Using command line arguments: {args.month} -> {target_date.strftime('%B %Y')}, format: {export_format}")
        except ValueError as e:
            print(f"Error in command line arguments: {e}")
            return
    else:
        # Interactive mode
        target_date, export_format = get_user_input()
    
    print(f"\nProcessing prepayments for {target_date.strftime('%B %Y')}...")
    
    # Process prepayments from the specified Excel file and sheet
    entries_df = process_prepayments(INPUT_FILE, SHEET_NAME, target_date)

    # Save the DataFrame if it exists
    if not entries_df.empty:
        output_file = save_dataframe(entries_df, export_format, target_date)
        
        # Check if file exists and notify about overwrite
        if os.path.exists(output_file):
            print(f"File created: {output_file}")
        
        print(f"Successfully saved {len(entries_df)} entries to {output_file}")
    else:
        print("No entries to save.")


if __name__ == "__main__":
    main()
