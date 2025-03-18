import os
import sys
import time
import logging
import pandas as pd
import pickle
import importlib.util
from datetime import datetime
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from colorama import Fore, Back, Style
import colorama

# Initialize colorama
colorama.init(autoreset=True)

# Custom logging formatter with colors
class ColoredFormatter(logging.Formatter):
    COLORS = {
        'DEBUG': Fore.CYAN,
        'INFO': Fore.GREEN,
        'WARNING': Fore.YELLOW,
        'ERROR': Fore.RED,
        'CRITICAL': Fore.RED + Back.WHITE
    }
    
    def format(self, record):
        # Add timestamp with color
        timestamp = datetime.now().strftime('%H:%M:%S')
        log_color = self.COLORS.get(record.levelname, Fore.WHITE)
        
        # Format the message with colors
        if record.levelname == 'INFO':
            # For INFO messages, we'll add some variety
            if "Processing" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.BLUE}➤ "
            elif "Successfully" in record.msg or "Found" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.GREEN}✓ "
            elif "Comparing" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.CYAN}🔄 "
            elif "Updated" in record.msg or "Marked" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.GREEN}📝 "
            elif "Authenticating" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.YELLOW}🔑 "
            elif "Matched" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.GREEN}🔍 "
            else:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {log_color}ℹ️ "
        elif record.levelname == 'WARNING':
            prefix = f"{Fore.MAGENTA}[{timestamp}] {log_color}⚠️ "
        elif record.levelname == 'ERROR':
            prefix = f"{Fore.MAGENTA}[{timestamp}] {log_color}❌ "
        elif record.levelname == 'CRITICAL':
            prefix = f"{Fore.MAGENTA}[{timestamp}] {log_color}🔥 "
        else:
            prefix = f"{Fore.MAGENTA}[{timestamp}] {log_color}"
        
        return f"{prefix}{record.msg}"

# Configure logging with pretty colors
console_handler = logging.StreamHandler()
console_handler.setFormatter(ColoredFormatter())

file_handler = logging.FileHandler("staff_comparator.log")
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

logger = logging.getLogger("StaffComparator")
logger.setLevel(logging.INFO)
logger.addHandler(console_handler)
logger.addHandler(file_handler)

# Check for config file existence
CONFIG_FILE = 'config.py'
if not os.path.exists(CONFIG_FILE):
    logger.error(f"Configuration file not found: {CONFIG_FILE}")
    sys.exit(1)

# Import configuration
try:
    import config
    logger.info(f"Configuration loaded successfully.")
except Exception as e:
    logger.error(f"Error loading configuration: {e}")
    logger.error(f"Please check your config.py file for syntax errors.")
    sys.exit(1)

# Set up paths - assuming comparator.py is in the /core/ directory
CORE_DIR = os.path.dirname(os.path.abspath(__file__))  # /core/
ROOT_DIR = os.path.dirname(CORE_DIR)  # Main directory

# Set up directories
PROCESSED_DIR = os.path.join(CORE_DIR, "processed")

# Set up Google Sheets configuration
# Use a specific token file for this script to avoid scope conflicts
TOKEN_FILE = "token_comparator.pickle"
# Scopes needed for reading AND writing to sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Get OUTPUT_DIR from config if available, otherwise use default
if hasattr(config, 'OUTPUT_DIR') and config.OUTPUT_DIR:
    OUTPUT_DIR = config.OUTPUT_DIR
else:
    OUTPUT_DIR = os.path.join(os.getcwd(), "EMR_Reports")

def authenticate():
    """Authenticate with Google using OAuth with write permissions"""
    creds = None
    logger.info("Authenticating with Google Sheets API (with write permission)...")
    
    # Check if we have stored credentials
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)
        logger.info("Found existing credentials")
    
    # If credentials don't exist or are invalid, get new ones
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            logger.info("Refreshing expired credentials")
            creds.refresh(Request())
        else:
            logger.info("Need to obtain new credentials with write permissions")
            print(f"\n{Fore.YELLOW}You need to authorize this application to access and modify your Google Sheets.")
            print(f"{Fore.CYAN}A browser window will open. Please log in and grant permission.\n")
            
            try:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
                logger.info("Successfully obtained new credentials with write permissions")
            except FileNotFoundError:
                logger.error("credentials.json file not found!")
                print(f"\n{Fore.RED}Error: credentials.json file not found!")
                print(f"{Fore.YELLOW}Please download OAuth credentials from Google Cloud Console:")
                print(f"{Fore.WHITE}1. Go to https://console.cloud.google.com/")
                print(f"{Fore.WHITE}2. Create a project or select an existing one")
                print(f"{Fore.WHITE}3. Go to APIs & Services > Credentials")
                print(f"{Fore.WHITE}4. Create OAuth client ID credentials (type: Desktop app)")
                print(f"{Fore.WHITE}5. Download the credentials.json file and place it in the same directory as this script\n")
                return None
        
        # Save credentials for future use
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)
        logger.info("Saved credentials for future use")
    
    return creds

def load_need_staff():
    """Load Need Staff data from processed directory"""
    need_staff_file = os.path.join(PROCESSED_DIR, "NeedStaff.csv")
    if not os.path.exists(need_staff_file):
        logger.error(f"Need Staff file not found: {need_staff_file}")
        return None
    
    try:
        need_staff_df = pd.read_csv(need_staff_file)
        logger.info(f"Loaded {len(need_staff_df)} need staff entries from {need_staff_file}")
        return need_staff_df
    except Exception as e:
        logger.error(f"Error loading Need Staff data: {e}")
        return None

def load_active_cases():
    """Load all Active Cases data from processed files"""
    patient_files = []
    for filename in os.listdir(PROCESSED_DIR):
        if filename.endswith("_Processed.csv"):
            patient_files.append(os.path.join(PROCESSED_DIR, filename))
    
    if not patient_files:
        logger.error(f"No processed patient files found in {PROCESSED_DIR}")
        return None
    
    try:
        all_patients = []
        for file_path in patient_files:
            file_name = os.path.basename(file_path)
            try:
                df = pd.read_csv(file_path)
                all_patients.append(df)
                logger.info(f"Loaded {len(df)} patients from {file_name}")
            except Exception as e:
                logger.error(f"Error loading patient data from {file_name}: {e}")
        
        # Combine all patient dataframes
        if all_patients:
            combined_df = pd.concat(all_patients, ignore_index=True)
            logger.info(f"Combined {len(combined_df)} total active patients")
            return combined_df
        else:
            logger.error("No patient data was successfully loaded")
            return None
    except Exception as e:
        logger.error(f"Error loading patient data: {e}")
        return None

def update_google_sheet(service, matches):
    """Update the Google Sheet with 'Already Staffed' for matching IDs"""
    if not matches:
        logger.info("No matches to update in Google Sheet")
        return 0
    
    # Get spreadsheet info from config
    if not hasattr(config, 'GOOGLE_SHEETS') or 'need_staff' not in config.GOOGLE_SHEETS:
        logger.error("Need Staff Google Sheet configuration not found in config.py")
        return 0
    
    sheet_config = config.GOOGLE_SHEETS['need_staff']
    spreadsheet_id = sheet_config['url'].split('/')[-1]
    worksheet_name = sheet_config['worksheet']
    
    logger.info(f"Getting data from Google Sheet: {worksheet_name}")
    
    try:
        # Get the current data to find row indices
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{worksheet_name}'!A:I"
        ).execute()
        
        values = result.get('values', [])
        if not values:
            logger.warning("No data found in Google Sheet")
            return 0
        
        # Convert matches to strings for comparison
        str_matches = set(str(m) for m in matches)
        
        # Find the ID column index (should be column A)
        id_col_idx = 0  # Assuming ID is in column A (index 0)
        
        # Find matching rows and prepare updates
        batch_updates = []
        update_count = 0
        
        for i, row in enumerate(values[1:], start=1):  # Skip header row
            if len(row) <= id_col_idx:
                continue  # Row doesn't have enough columns
                
            try:
                sheet_id = str(row[id_col_idx]).strip()
                if sheet_id in str_matches:
                    # Prepare to write "Already Staffed" in column I
                    row_idx = i + 1  # 1-based row index for API
                    
                    batch_updates.append({
                        'range': f"'{worksheet_name}'!I{row_idx}",
                        'values': [["Already Staffed"]]
                    })
                    
                    logger.info(f"Marking ID {sheet_id} as 'Already Staffed' at row {row_idx}")
                    update_count += 1
            except Exception as e:
                logger.error(f"Error processing row {i}: {e}")
        
        # Update in smaller batches to avoid issues
        batch_size = 10
        for i in range(0, len(batch_updates), batch_size):
            batch_chunk = batch_updates[i:i+batch_size]
            try:
                service.spreadsheets().values().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={
                        'valueInputOption': 'RAW',
                        'data': batch_chunk
                    }
                ).execute()
                logger.info(f"Updated batch of {len(batch_chunk)} rows")
            except HttpError as error:
                logger.error(f"Error updating batch: {error}")
                # Try individual updates for this batch
                for update in batch_chunk:
                    try:
                        range_name = update['range']
                        values = update['values']
                        service.spreadsheets().values().update(
                            spreadsheetId=spreadsheet_id,
                            range=range_name,
                            valueInputOption='RAW',
                            body={'values': values}
                        ).execute()
                        logger.info(f"Updated {range_name} individually")
                    except Exception as e:
                        logger.error(f"Failed to update {range_name}: {e}")
        
        logger.info(f"Successfully updated {update_count} rows in Google Sheet")
        return update_count
        
    except HttpError as error:
        logger.error(f"Error updating Google Sheet: {error}")
        print(f"\n{Fore.RED}Error: {error}")
        if "insufficient authentication scopes" in str(error):
            print(f"\n{Fore.YELLOW}Authentication Error: The token doesn't have write permissions.")
            print(f"{Fore.CYAN}Try deleting the token_comparator.pickle file and run the script again.")
        return 0
    except Exception as e:
        logger.error(f"Unexpected error updating Google Sheet: {e}")
        return 0

def compare_staff_data():
    """Compare Need Staff with Active Cases and mark matches in Google Sheet"""
    # Load Need Staff data
    need_staff_df = load_need_staff()
    if need_staff_df is None:
        return False
    
    # Load Active Cases data
    active_cases_df = load_active_cases()
    if active_cases_df is None:
        return False
    
    # Compare the data
    logger.info("Comparing Need Staff with Active Cases")
    
    # Extract patient IDs from both dataframes
    need_staff_ids = set(need_staff_df['ID'].astype(str))
    active_cases_ids = set(active_cases_df['Patient_Id'].astype(str))
    
    # Find matches
    matches = need_staff_ids.intersection(active_cases_ids)
    
    if matches:
        logger.info(f"Found {len(matches)} matching patient IDs")
        
        # Authenticate with Google Sheets
        creds = authenticate()
        if not creds:
            logger.error("Failed to authenticate with Google Sheets")
            return False
        
        # Build the Google Sheets service
        service = build('sheets', 'v4', credentials=creds)
        
        # Update Google Sheet with matches
        updated_count = update_google_sheet(service, matches)
        if updated_count > 0:
            logger.info(f"Successfully updated {updated_count} entries in Google Sheet")
            return True
        else:
            logger.warning("No entries were updated in Google Sheet")
            return False
    else:
        logger.info("No matching IDs found between Need Staff and Active Cases")
        return True

def main():
    """Main function to compare staff data and update Google Sheet"""
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
    print(f"{Fore.YELLOW}{Style.BRIGHT}{'STAFF COMPARATOR':^80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}\n")
    
    logger.info("Starting staff comparison process")
    
    # Track timing
    total_start_time = time.time()
    
    # Check if the processed directory exists
    if not os.path.exists(PROCESSED_DIR):
        logger.error(f"Processed directory not found: {PROCESSED_DIR}")
        print(f"{Fore.RED}Error: Processed directory not found: {PROCESSED_DIR}")
        print(f"{Fore.YELLOW}Please run the data_processor.py script first.")
        return
    
    # Perform comparison and update
    success = compare_staff_data()
    
    # Print summary
    total_duration = time.time() - total_start_time
    minutes = int(total_duration // 60)
    seconds = int(total_duration % 60)
    
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
    if success:
        print(f"{Fore.GREEN}{Style.BRIGHT}{'STAFF COMPARISON COMPLETED SUCCESSFULLY':^80}")
    else:
        print(f"{Fore.RED}{Style.BRIGHT}{'STAFF COMPARISON COMPLETED WITH ERRORS':^80}")
    print(f"{Fore.GREEN}{Style.BRIGHT}{f'Total time: {minutes} minutes {seconds} seconds':^80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}\n")
    
    logger.info(f"Staff comparison completed in {minutes}m {seconds}s")

if __name__ == "__main__":
    main()