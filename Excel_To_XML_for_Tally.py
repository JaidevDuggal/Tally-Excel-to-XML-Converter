import pandas as pd
import sys
from pathlib import Path
import re
from xml.sax.saxutils import escape  # <-- Zaroori fix for '&' symbol errors
import sys
# New Import for Colors
from colorama import Fore, Style, init

# Ye line zaroori hai Windows ke liye:
init(autoreset=True) 
# autoreset=True ka matlab hai ki har print statement ke baad color automatic hat jayega,
# taaki agli line normal black/white mein print ho.

# --- Configuration: Define Standard Column Names ---
EXPECTED_COLUMNS_MAP = {
    's.no.': 'sno',
    'date': 'date',
    'drledger': 'drledger',
    'crledger': 'crledger',
    'amount': 'amount',
    'narration': 'narration',
    'vouchernumber': 'vouchernumber'
}

# Define which columns are essential
ESSENTIAL_COLUMNS = ['sno', 'date', 'drledger', 'crledger', 'amount', 'vouchernumber']
# Define which columns must have data
ROW_ESSENTIAL_COLUMNS = ['sno', 'date', 'drledger', 'crledger', 'amount']

# --- Helper Functions ---

def get_user_paths():
    excel_path_str = input("Please enter the full path to your Excel file: ").strip().strip('"')
    excel_path = Path(excel_path_str)
    
    # Check if file exists
    if not excel_path.is_file():
        print(Fore.RED + Style.BRIGHT + f"\nERROR: Excel file not found at the specified path.")
        sys.exit(1)
    
    # Check extension
    elif not excel_path_str.lower().endswith(".xlsx"):
        print(Fore.RED + Style.BRIGHT + "Error: Only .xlsx files are allowed.")
        sys.exit(1)

    output_folder_str = input("Please enter the full path for the output XML folder: ").strip().strip('"')
    output_folder = Path(output_folder_str)

    if not output_folder.is_dir():
        print(Fore.RED + Style.BRIGHT + f"\nERROR: Output folder not found at the specified path.")
        sys.exit(1)
        
    output_filename = output_folder / "Tally_Import.xml"
    return excel_path, output_filename

def standardize_headers(df):
    original_headers = df.columns
    cleaned_headers = [re.sub(r'[\s\._]+', '', str(h).lower()) for h in original_headers]
    final_headers = [EXPECTED_COLUMNS_MAP.get(h, h) for h in cleaned_headers]
    df.columns = final_headers
    return df

def check_required_columns(df_columns):
    missing_cols = [col for col in ESSENTIAL_COLUMNS if col not in df_columns]
    if missing_cols:
        print(Fore.RED + Style.BRIGHT + "\n--- ERROR: Missing Required Columns ---")
        print(f"{Fore.RED}{Style.BRIGHT}Your Excel file is missing the following required columns:{Style.RESET_ALL} {', '.join(missing_cols)}")
        sys.exit(1)

def create_voucher_xml(data):
    # Safety Fix: Handle special characters like '&' in names/narration
    safe_narration = escape(str(data['narration']))
    safe_drledger = escape(str(data['drledger']))
    safe_crledger = escape(str(data['crledger']))
    
    return f"""
        <TALLYMESSAGE>
          <VOUCHER ACTION="Create" VCHTYPE="Journal"> 
            <DATE>{data['date']}</DATE> 
            <VOUCHERTYPENAME>Journal</VOUCHERTYPENAME>
            <VOUCHERNUMBER>{data['vouchernumber']}</VOUCHERNUMBER>
            <NARRATION>{safe_narration}</NARRATION> 
            
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>{safe_drledger}</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-{data['amount']:.2f}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>{safe_crledger}</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>{data['amount']:.2f}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>
        </TALLYMESSAGE>
    """.strip()

def generate_full_xml(voucher_xml_list):
    all_vouchers_str = "\n".join(voucher_xml_list)
    return f"""
<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Import Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
      </REQUESTDESC>
      <REQUESTDATA>
        {all_vouchers_str}
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>
    """.strip()

# --- Main Execution ---

def main():
    try:
        print(f"""
        {Fore.CYAN}{Style.BRIGHT}====================================================================================
        {Fore.CYAN}{Style.BRIGHT}                     🚀 Tally Voucher Import Tool (v1.0)
        {Fore.WHITE}{Style.DIM}                              By: Jaidev Duggal
        {Fore.YELLOW}------------------------------------------------------------------------------------
        {Style.RESET_ALL}{Fore.CYAN}💡 This tool securely converts your customized Excel Sheet data 
           into a Tally-compatible XML file for direct voucher import.

        {Fore.GREEN}{Style.BRIGHT}✅ RELIABILITY GUARANTEED:{Style.RESET_ALL} The tool strictly validates every row for Date, 
           Ledger, and Amount before creating the XML.

        {Fore.MAGENTA}{Style.BRIGHT}🔗 Connect with the Developer:
        - LinkedIn: {Fore.BLUE}https://www.linkedin.com/in/jaidev-duggal{Fore.MAGENTA}{Style.BRIGHT}
        - GitHub:   {Fore.BLUE}https://github.com/JaidevDuggal
        {Fore.YELLOW}{Style.DIM}------------------------------------------------------------------------------------{Style.RESET_ALL}
        
        {Fore.RED}{Style.BRIGHT}[IMPORTANT WARNING]: Please read the 'Instructions' sheet in the Excel Template before running this tool.
        
        {Fore.CYAN}{Style.BRIGHT}===================================================================================={Style.RESET_ALL}
        """)

        excel_path, output_filename = get_user_paths()

        print(f"\nReading Excel file: {excel_path}...")
        print("Note: Reading the first sheet of the Excel file only.")
        
        try:
            # Pandas handles headers and empty rows automatically (Much Smarter!)
            df = pd.read_excel(excel_path)
        except PermissionError:
            print(f"{Fore.RED}{Style.BRIGHT}\n--- ERROR: File Locked ---")
            print("The Excel file is currently open. Please close it and try again.")
            sys.exit(1)
        except Exception as e:
            print(f"{Fore.RED}{Style.BRIGHT}\nERROR: Could not read the Excel file. Details: {Style.RESET_ALL}{e}")
            sys.exit(1)

        df = standardize_headers(df)
        check_required_columns(df.columns)
        print("Excel headers are valid.")

        all_vouchers_xml = []
        print("Validating rows and converting data...")

        for index, row in df.iterrows():
            # Get S.No safely
            s_no = row.get('sno', f"Excel Row {index + 2}") 
            
            # --- Guardrail 2: Critical Data Validation ---
            for col in ROW_ESSENTIAL_COLUMNS:
                if pd.isna(row[col]) or str(row[col]).strip() == "":
                    print(f"{Fore.RED}{Style.BRIGHT}\n--- ERROR: Missing Data in Row ---")
                    print(f"The row with S. No. '{s_no}' has a blank value in the '{col}' column.")
                    print("This is a required field. Process stopped.")
                    sys.exit(1)

            # --- Guardrail 3: Date Format Validation ---
            try:
                # dayfirst=True handles 04/07/2025 correctly as 4th July
                date_obj = pd.to_datetime(row['date'], dayfirst=True)
                tally_date = date_obj.strftime('%Y%m%d')
            except Exception:
                print(f"{Fore.RED}{Style.BRIGHT}\n--- ERROR: Invalid Date Format ---")
                print(f"Row S. No. '{s_no}': Invalid date '{row.get('date')}'.")
                print("Please correct the date and run again.")
                sys.exit(1)

            # --- Amount Validation ---
            try:
                amount = float(row['amount'])
            except ValueError:
                print(f"{Fore.RED}{Style.BRIGHT}\n--- ERROR: Invalid Amount ---")
                print(f"Row S. No. '{s_no}': Invalid amount '{row.get('amount')}'.")
                sys.exit(1)

            # --- Prepare Data ---
            narration = str(row['narration']) if 'narration' in row and pd.notna(row['narration']) else ""
            voucher_number = str(row['vouchernumber']) if pd.notna(row['vouchernumber']) else ""

            voucher_data = {
                'date': tally_date,
                'vouchernumber': voucher_number,
                'narration': narration,
                'drledger': str(row['drledger']).strip(),
                'crledger': str(row['crledger']).strip(),
                'amount': amount
            }
            
            all_vouchers_xml.append(create_voucher_xml(voucher_data))

        if not all_vouchers_xml:
            print(f"{Fore.RED}{Style.BRIGHT}\nERROR: No data rows found in the Excel file.")
            sys.exit(1)
            
        print(f"{Fore.GREEN}Validation complete. {len(all_vouchers_xml)} vouchers processed.")
        
        final_xml_content = generate_full_xml(all_vouchers_xml)
        
        with open(output_filename, "w", encoding="utf-8") as f:
            f.write(final_xml_content)
            
        print(f"{Fore.GREEN}{Style.BRIGHT}\n--- SUCCESS! ---")
        print(f"XML file created at: {output_filename}")
        print(f"{Fore.GREEN}{Style.BRIGHT}\n\n------------------------------------------------------------------")
        print(f"{Fore.GREEN}{Style.BRIGHT}------------------------------------------------------------------")
        print(f"{Fore.GREEN}{Style.BRIGHT}👍 If helpful, please connect with the developer!")

    except Exception as e:
        print(f"{Fore.RED}{Style.BRIGHT}\n--- An Unexpected Error Occurred ---")
        print(f"Error: {e}")
        
    finally:
        input("\nPress Enter to exit.")

if __name__ == "__main__":
    main()

#pyinstaller --noconfirm --clean "Excel_To_XML_for_Tally.spec"
#pyinstaller --noconfirm --onedir --console --name "Tally_Excel_to_XML" --clean --exclude-module tkinter --exclude-module matplotlib --exclude-module scipy --exclude-module pygame --exclude-module IPython --exclude-module notebook "Excel_To_XML_for_Tally.py"