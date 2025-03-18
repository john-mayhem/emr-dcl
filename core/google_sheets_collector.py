import os
import sys
import time
import pandas as pd
import pickle
import logging
import importlib.util
from datetime import datetime
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from colorama import Fore, Back, Style
import colorama

# Initialize colorama
colorama.init(autoreset=True)

# Custom logging formatter with colors - matching EMR collector style
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
            elif "Successfully" in record.msg or "Retrieved" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.GREEN}✓ "
            elif "Waiting" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.CYAN}⏳ "
            elif "Downloaded" in record.msg or "Saved" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.GREEN}📄 "
            elif "Authenticating" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.YELLOW}🔑 "
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

file_handler = logging.FileHandler("google_sheets_collection.log")
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

logger = logging.getLogger("GoogleSheetsCollector")
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

# Set up configuration variables from config.py
TOKEN_FILE = "token.pickle"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Get OUTPUT_DIR from config if available, otherwise use default
if hasattr(config, 'OUTPUT_DIR') and config.OUTPUT_DIR:
    OUTPUT_DIR = config.OUTPUT_DIR
else:
    OUTPUT_DIR = os.path.join(os.getcwd(), "EMR_Reports")

# Create output directory if it doesn't exist
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)
    logger.info(f"Created output directory: {OUTPUT_DIR}")

# Use Google Sheets configuration from config.py
if hasattr(config, 'GOOGLE_SHEETS') and config.GOOGLE_SHEETS:
    SPREADSHEETS = []
    for key, sheet_info in config.GOOGLE_SHEETS.items():
        # Extract the spreadsheet ID from the URL
        sheet_id = sheet_info['url'].split('/')[-1]
        
        SPREADSHEETS.append({
            "name": key.capitalize(),
            "id": sheet_id,
            "sheet": sheet_info['worksheet'],
            "range": f"A1:{'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[sheet_info['columns']-1]}",  # Calculate range based on columns
            "output": sheet_info['output_file']
        })
    logger.info(f"Loaded {len(SPREADSHEETS)} sheet configurations from config.py")
else:
    logger.error("No Google Sheets configuration found in config.py")
    sys.exit(1)

def authenticate():
    """Authenticate with Google using OAuth"""
    creds = None
    logger.info("Authenticating with Google Sheets API...")
    
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
            logger.info("Need to obtain new credentials")
            print(f"\n{Fore.YELLOW}You need to authorize this application to access your Google Sheets.")
            print(f"{Fore.CYAN}A browser window will open. Please log in and grant permission.\n")
            
            try:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
                logger.info("Successfully obtained new credentials")
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

def get_sheet_data(service, sheet):
    """Get data from a Google Sheet"""
    try:
        logger.info(f"Processing {sheet['name']} sheet from {sheet['id']}")
        
        # Construct the range string
        range_name = f"'{sheet['sheet']}'!{sheet['range']}"
        
        # Get the sheet data
        logger.info(f"Fetching data from range: {range_name}")
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet['id'],
            range=range_name
        ).execute()
        
        # Get the values from the result
        values = result.get('values', [])
        
        if not values:
            logger.warning(f"No data found in {sheet['name']}")
            return None
            
        logger.info(f"Retrieved {len(values)} rows from {sheet['name']}")
        return values
    except Exception as e:
        logger.error(f"Error getting data from {sheet['name']}: {e}")
        return None

def save_as_csv(data, filename):
    """Save data as a CSV file"""
    if not data:
        logger.warning(f"No data to save for {filename}")
        return False
        
    try:
        # Create DataFrame
        df = pd.DataFrame(data)
        
        # Use first row as header
        headers = df.iloc[0]
        df = df[1:]  # Remove header row
        df.columns = headers
        
        # Save as CSV
        output_path = os.path.join(OUTPUT_DIR, filename)
        df.to_csv(output_path, index=False)
        logger.info(f"Saved data to {output_path}")
        return True
    except Exception as e:
        logger.error(f"Error saving data as CSV: {e}")
        return False

def main():
    """Main function to download Google Sheets data"""
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
    print(f"{Fore.YELLOW}{Style.BRIGHT}{'GOOGLE SHEETS DATA COLLECTION':^80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}\n")
    
    logger.info("Starting Google Sheets data collection process")
    
    # Authenticate with Google
    creds = authenticate()
    if not creds:
        return
    
    # Build the service
    service = build('sheets', 'v4', credentials=creds)
    logger.info("Google Sheets API service initialized")
    
    # Track success and overall timing
    success_count = 0
    total_start_time = time.time()
    
    # Process each sheet
    for i, sheet in enumerate(SPREADSHEETS):
        sheet_num = i + 1
        
        print(f"\n{Fore.CYAN}{'─'*80}")
        print(f"{Fore.YELLOW}[{sheet_num}/{len(SPREADSHEETS)}] Processing: {Fore.CYAN}{sheet['name']}")
        print(f"{Fore.CYAN}{'─'*80}{Style.RESET_ALL}")
        
        start_time = time.time()
        
        # Get sheet data
        data = get_sheet_data(service, sheet)
        
        # Save as CSV
        if data and save_as_csv(data, sheet['output']):
            end_time = time.time()
            duration = end_time - start_time
            
            success_count += 1
            print(f"{Fore.GREEN}✓ {Style.BRIGHT}Sheet completed in {duration:.2f} seconds ({success_count}/{len(SPREADSHEETS)} done)")
        else:
            print(f"{Fore.RED}✗ {Style.BRIGHT}Failed to process {sheet['name']} ({success_count}/{len(SPREADSHEETS)} done)")
    
    # Print summary
    total_duration = time.time() - total_start_time
    minutes = int(total_duration // 60)
    seconds = int(total_duration % 60)
    
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
    print(f"{Fore.YELLOW}{Style.BRIGHT}{f'DATA COLLECTION COMPLETED: {success_count}/{len(SPREADSHEETS)} SHEETS':^80}")
    print(f"{Fore.GREEN}{Style.BRIGHT}{f'Total time: {minutes} minutes {seconds} seconds':^80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}\n")
    
    logger.info(f"Google Sheets data collection completed: {success_count}/{len(SPREADSHEETS)} sheets processed")

if __name__ == "__main__":
    main()