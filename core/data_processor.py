import os
import sys
import time
import logging
import pandas as pd
import glob
import re
import shutil
from datetime import datetime
import importlib.util

# Check for colorama availability
try:
    import colorama
    from colorama import Fore, Back, Style
    colorama.init(autoreset=True)
    HAS_COLOR = True
except ImportError:
    HAS_COLOR = False
    # Create dummy color classes if colorama is not available
    class DummyFore:
        def __getattr__(self, name):
            return ""
    class DummyBack:
        def __getattr__(self, name):
            return ""
    class DummyStyle:
        def __getattr__(self, name):
            return ""
    Fore = DummyFore()
    Back = DummyBack()
    Style = DummyStyle()

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
            elif "Successfully" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.GREEN}✓ "
            elif "Combining" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.CYAN}🔄 "
            elif "Saved" in record.msg or "Created" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.GREEN}📄 "
            elif "Calculating" in record.msg or "Mapping" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.YELLOW}🧮 "
            elif "Found" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.GREEN}🔍 "
            elif "Copied" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.CYAN}📋 "
            elif "Extracted" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.YELLOW}📊 "
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

file_handler = logging.FileHandler("data_processing.log")
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

logger = logging.getLogger("DataProcessor")
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

# Set up paths - assuming data_processor.py is in the /core/ directory
CORE_DIR = os.path.dirname(os.path.abspath(__file__))  # /core/
ROOT_DIR = os.path.dirname(CORE_DIR)  # Main directory

# First try to read from config
if hasattr(config, 'OUTPUT_DIR') and config.OUTPUT_DIR:
    DATA_DIR = config.OUTPUT_DIR
else:
    # Based on the directory structure shown, data files are in /core/EMR_Reports/
    DATA_DIR = os.path.join(CORE_DIR, "EMR_Reports")
    
    # If that doesn't exist, fall back to the root directory
    if not os.path.exists(DATA_DIR):
        DATA_DIR = os.path.join(ROOT_DIR, "EMR_Reports")

logger.info(f"Using data directory: {DATA_DIR}")

# Create processed directory within the core folder
PROCESSED_DIR = os.path.join(CORE_DIR, "processed")
if not os.path.exists(PROCESSED_DIR):
    os.makedirs(PROCESSED_DIR)
    logger.info(f"Created processed directory: {PROCESSED_DIR}")

def process_need_staff():
    """Copy NeedStaff.csv as-is to the processed directory"""
    try:
        input_file = os.path.join(DATA_DIR, "NeedStaff.csv")
        output_file = os.path.join(PROCESSED_DIR, "NeedStaff.csv")
        
        logger.info(f"Copying NeedStaff.csv")
        
        if not os.path.exists(input_file):
            logger.error(f"Input file not found: {input_file}")
            return False
        
        # Simply copy the file
        shutil.copy2(input_file, output_file)
        logger.info(f"Successfully copied NeedStaff.csv to {output_file}")
        return True
    except Exception as e:
        logger.error(f"Error copying NeedStaff.csv: {e}")
        return False

def copy_active_therapists():
    """Copy ActiveTherapists.csv to the processed directory"""
    try:
        input_file = os.path.join(DATA_DIR, "ActiveTherapists.csv")
        output_file = os.path.join(PROCESSED_DIR, "ActiveTherapists.csv")
        
        logger.info(f"Copying ActiveTherapists.csv")
        
        if not os.path.exists(input_file):
            logger.error(f"Input file not found: {input_file}")
            return False
        
        shutil.copy2(input_file, output_file)
        logger.info(f"Successfully copied ActiveTherapists.csv to {output_file}")
        return True
    except Exception as e:
        logger.error(f"Error copying ActiveTherapists.csv: {e}")
        return False

def clean_address(address, city, state, zip_code):
    """
    Clean up address field by removing duplicated city/state/zip information
    and correcting cases where the address contains different location info than the columns
    """
    if not isinstance(address, str):
        return address
    
    # First check if address contains a different zipcode than the zipcode column
    # This likely means the address is more accurate (from treating address)
    if isinstance(zip_code, str):
        zip_matches = re.findall(r'\b\d{5}\b', address)
        if zip_matches:
            embedded_zip = zip_matches[-1]  # Take the last zipcode in the address
            if embedded_zip != zip_code:
                # If different zipcodes, use zip from address and update zip_code variable
                zip_code = embedded_zip
    
    # Extract different components from the address
    # This regex looks for "City, State Zip" pattern at the end of the address
    location_match = re.search(r'(.+?)(?:,\s*([A-Z]{2})\s*(\d{5}))?$', address)
    
    if location_match:
        clean_address = location_match.group(1).strip()
        
        # If embedded state/zip were extracted, they take precedence
        if location_match.group(2) and location_match.group(3):
            embedded_state = location_match.group(2)
            embedded_zip = location_match.group(3)
            
            # If the address has precise location info, truncate at the city
            city_match = re.search(r'(.+?)(?:,\s*([^,]+))(?:,\s*[A-Z]{2}\s*\d{5})$', address)
            if city_match:
                clean_address = city_match.group(1).strip()
    else:
        clean_address = address
    
    # Remove any trailing commas or spaces
    clean_address = re.sub(r'[,\s]+$', '', clean_address)
    
    # Handle special case where only the zipcode is at the end
    clean_address = re.sub(r',?\s+\d{5}$', '', clean_address)
    
    # Remove "Apt X, City" patterns at the end
    clean_address = re.sub(r'(Apt\s+[\w\d]+),\s*[^,]+$', r'\1', clean_address)
    
    # Handle "Street, City, State ZIP" pattern
    street_city_pattern = re.match(r'(.+),\s*([^,]+),\s*[A-Z]{2}\s*\d{5}', clean_address)
    if street_city_pattern:
        clean_address = street_city_pattern.group(1)
    
    return clean_address

def extract_year_from_age(age_str):
    """Extract just the years from an age string like '98y 10m 17d'"""
    if not isinstance(age_str, str):
        return age_str
    
    # Extract the years portion
    match = re.search(r'(\d+)y', age_str)
    if match:
        return f"{match.group(1)}y"
    
    return age_str

def extract_treating_address(comment):
    """Extract the treating address from the comment field"""
    if not isinstance(comment, str):
        return None
    
    # Look for patterns like "Treating add:" or similar
    treating_patterns = [
        r"Treating add(?:ress)?:\s*(.*?)(?:\n|$)",
        r"Treating address(?:es)?:\s*(.*?)(?:\n|$)",
        r"Treating Add(?:ress)?:\s*(.*?)(?:\n|$)",
        r"Treatment add(?:ress)?:\s*(.*?)(?:\n|$)",
        r"Treatment Address(?:es)?:\s*(.*?)(?:\n|$)"
    ]
    
    for pattern in treating_patterns:
        match = re.search(pattern, comment, re.IGNORECASE)
        if match:
            address = match.group(1).strip()
            return address
    
    return None

def process_excel_file(file_path):
    """Process a single Excel file to extract the required columns and handle treating addresses"""
    try:
        logger.info(f"Processing Excel file: {os.path.basename(file_path)}")
        
        # Extract office and discipline from the filename
        file_basename = os.path.basename(file_path)
        
        # IMPROVED REGEX: Handle the Rehab on Wheels case correctly
        # This pattern looks for "_OT_" or "_PT_" at the end of the filename before "With Cases"
        match = re.search(r'(.+?)_(OT|PT)_With Cases', file_basename)
        if match:
            office = match.group(1)
            discipline = match.group(2)
        else:
            office = "Unknown"
            discipline = "Unknown"
            logger.warning(f"Could not extract office and discipline from filename: {file_basename}")
        
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Check if the dataframe is empty
        if df.empty:
            logger.warning(f"Excel file is empty: {file_basename}")
            return None
        
        # Log the column names to help with debugging
        logger.info(f"Excel file columns: {df.columns.tolist()}")
        
        # Try to map columns by name instead of position
        # Define expected column names and their new names
        column_mapping = {
            'Patient Id': 'Patient_Id',
            'First Name': 'First_Name',
            'Last Name': 'Last_Name',
            'Age': 'Age',
            'Address': 'Address',
            'City': 'City',
            'State': 'State',
            'Zip': 'Zip',
            'Status': 'Status'
        }
        
        # Check if columns exist and create a mapping
        available_columns = {}
        for orig_col, new_col in column_mapping.items():
            if orig_col in df.columns:
                available_columns[orig_col] = new_col
            else:
                logger.warning(f"Column '{orig_col}' not found in {file_basename}")
        
        # If critical columns are missing, try positional mapping as fallback
        if 'Status' not in available_columns:
            logger.warning(f"Status column not found by name, trying to find it by position")
            
            # Log the first few rows to see what data looks like
            logger.info(f"First row data: {df.iloc[0].values}")
            
            # Check if we have enough columns for Status (which should be at index 21)
            if len(df.columns) > 21:
                # Check if column 21 has values like 'active', 'pending', etc.
                status_values = df.iloc[:, 21].astype(str).str.lower()
                has_status_keywords = status_values.str.contains('active|home care|pending').any()
                
                if has_status_keywords:
                    logger.info(f"Found Status column at position 21")
                    # Add Status to available columns with positional mapping
                    df['Status'] = df.iloc[:, 21]
                    available_columns['Status'] = 'Status'
        
        # If we still don't have Status column, try to find it by content
        if 'Status' not in available_columns:
            logger.warning(f"Still couldn't find Status column, searching all columns for status values")
            
            for col_idx, col_name in enumerate(df.columns):
                col_values = df.iloc[:, col_idx].astype(str).str.lower()
                has_status_keywords = col_values.str.contains('active|home care|pending').any()
                
                if has_status_keywords:
                    logger.info(f"Found probable Status column in '{col_name}' at position {col_idx}")
                    df['Status'] = df.iloc[:, col_idx]
                    available_columns['Status'] = 'Status'
                    break
        
        # Create a subset with only the columns we found and mapped
        if available_columns:
            selected_columns = list(available_columns.keys())
            selected_df = df[selected_columns].copy()
            selected_df.columns = [available_columns[col] for col in selected_columns]
            
            # Check if we have the minimal required columns
            required_columns = ['Patient_Id', 'Address', 'Status']
            missing_columns = [col for col in required_columns if col not in selected_df.columns]
            
            if missing_columns:
                logger.error(f"Missing required columns: {missing_columns}")
                return None
        else:
            logger.error(f"Could not find any required columns in {file_basename}")
            return None
        
        # First, handle Name if we have First_Name and Last_Name but not Name
        if 'First_Name' in selected_df.columns and 'Last_Name' in selected_df.columns and 'Name' not in selected_df.columns:
            selected_df['Name'] = selected_df['First_Name'] + ' ' + selected_df['Last_Name']
            
            # Drop the individual name columns if they exist
            if 'First_Name' in selected_df.columns:
                selected_df = selected_df.drop('First_Name', axis=1)
            if 'Last_Name' in selected_df.columns:
                selected_df = selected_df.drop('Last_Name', axis=1)
        
        # Ensure we have a 'Name' column
        if 'Name' not in selected_df.columns:
            selected_df['Name'] = '[Name Not Available]'
        
        # Debug: Log unique status values
        unique_statuses = selected_df['Status'].unique()
        logger.info(f"Found unique status values: {unique_statuses}")
        
        # Convert Status to string and lowercase for case-insensitive comparison
        selected_df['Status'] = selected_df['Status'].astype(str).str.lower()
        
        # Filter out statuses we don't want - use lowercase since we converted above
        keep_statuses = ['active', 'home care', 'pending']
        filtered_df = selected_df[selected_df['Status'].str.contains('|'.join(keep_statuses))]
        logger.info(f"Filtered from {len(selected_df)} to {len(filtered_df)} records based on status")
        
        # If we filtered everything out, log a warning
        if filtered_df.empty and not selected_df.empty:
            logger.warning(f"All records filtered out. Check the status values in your Excel file.")
        
        # Clean up the Age column to just extract the years
        if 'Age' in filtered_df.columns:
            filtered_df['Age'] = filtered_df['Age'].apply(extract_year_from_age)
        else:
            # Add a placeholder Age column if missing
            filtered_df['Age'] = 'N/A'
        
        # Make a copy of the original address, city, state, and zip before cleaning
        if 'Address' in filtered_df.columns:
            filtered_df['Original_Address'] = filtered_df['Address']
        else:
            filtered_df['Address'] = 'N/A'
            filtered_df['Original_Address'] = 'N/A'
            
        if 'City' in filtered_df.columns:
            filtered_df['Original_City'] = filtered_df['City']
        else:
            filtered_df['City'] = 'N/A'
            filtered_df['Original_City'] = 'N/A'
            
        if 'State' in filtered_df.columns:
            filtered_df['Original_State'] = filtered_df['State']
        else:
            filtered_df['State'] = 'N/A'
            filtered_df['Original_State'] = 'N/A'
            
        if 'Zip' in filtered_df.columns:
            filtered_df['Original_Zip'] = filtered_df['Zip']
        else:
            filtered_df['Zip'] = 'N/A'
            filtered_df['Original_Zip'] = 'N/A'
        
        # Clean up addresses only if we have valid data
        address_cols_valid = all(filtered_df[col].iloc[0] != 'N/A' for col in ['Address', 'City', 'State', 'Zip'])
        
        if address_cols_valid:
            # Clean up addresses
            filtered_df['Address'] = filtered_df.apply(
                lambda row: clean_address(row['Address'], row['City'], row['State'], row['Zip']), 
                axis=1
            )
            
            # After address cleaning, update the city/zip if needed based on embedded info
            for idx, row in filtered_df.iterrows():
                address = row['Original_Address']
                if isinstance(address, str):
                    zip_matches = re.findall(r'\b\d{5}\b', address)
                    if zip_matches:
                        embedded_zip = zip_matches[-1]  # Take the last zipcode in the address
                        if embedded_zip != row['Zip']:
                            # If the zip in the address doesn't match the cell, use the one from address
                            filtered_df.at[idx, 'Zip'] = embedded_zip
                            
                            # Try to identify city/state as well
                            city_state_match = re.search(r'([^,]+),\s*([A-Z]{2})\s*\d{5}', address)
                            if city_state_match:
                                embedded_city = city_state_match.group(1).strip()
                                embedded_state = city_state_match.group(2)
                                filtered_df.at[idx, 'City'] = embedded_city
                                filtered_df.at[idx, 'State'] = embedded_state
        
        # Drop the original columns
        filtered_df = filtered_df.drop(['Original_Address', 'Original_City', 'Original_State', 'Original_Zip'], axis=1)
        
        # Create output filename
        output_filename = f"{office}_{discipline}_Processed.csv"
        output_path = os.path.join(PROCESSED_DIR, output_filename)
        
        # Save processed data
        filtered_df.to_csv(output_path, index=False)
        logger.info(f"Successfully processed and saved: {output_filename}")
        
        return filtered_df
        
    except Exception as e:
        logger.error(f"Error processing Excel file {file_path}: {e}")
        import traceback
        logger.error(traceback.format_exc())  # Add detailed traceback
        return None

def main():
    """Main function to process all data files"""
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
    print(f"{Fore.YELLOW}{Style.BRIGHT}{'EMR DATA PROCESSING':^80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}\n")
    
    logger.info("Starting data processing")
    
    # Track timing and success count
    total_start_time = time.time()
    success_count = 0
    total_files = 0
    
    # 1. Process ActiveTherapists.csv
    if copy_active_therapists():
        success_count += 1
        total_files += 1
    
    # 2. Process NeedStaff.csv
    if process_need_staff():
        success_count += 1
        total_files += 1
    
    # 3. Process Excel files
    excel_files = glob.glob(os.path.join(DATA_DIR, "*.xlsx"))
    total_files += len(excel_files)
    
    # Create a list to store dataframes for potential combining later
    processed_dfs = []
    
    for i, file_path in enumerate(excel_files):
        print(f"\n{Fore.CYAN}{'─'*80}")
        print(f"{Fore.YELLOW}[{i+1}/{len(excel_files)}] Processing: {Fore.CYAN}{os.path.basename(file_path)}")
        print(f"{Fore.CYAN}{'─'*80}{Style.RESET_ALL}")
        
        start_time = time.time()
        processed_df = process_excel_file(file_path)
        end_time = time.time()
        
        if processed_df is not None:
            duration = end_time - start_time
            processed_dfs.append(processed_df)
            success_count += 1
            print(f"{Fore.GREEN}✓ {Style.BRIGHT}File processed in {duration:.2f} seconds ({success_count}/{total_files} done)")
        else:
            print(f"{Fore.RED}✗ {Style.BRIGHT}Failed to process file ({success_count}/{total_files} done)")
    
    # Note: As per requirements, we're not combining Excel files
    # Each processed file is saved separately with its own name
    
    # Print summary
    total_duration = time.time() - total_start_time
    minutes = int(total_duration // 60)
    seconds = int(total_duration % 60)
    
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
    print(f"{Fore.YELLOW}{Style.BRIGHT}{f'DATA PROCESSING COMPLETED: {success_count}/{total_files} FILES':^80}")
    print(f"{Fore.GREEN}{Style.BRIGHT}{f'Total time: {minutes} minutes {seconds} seconds':^80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}\n")
    
    logger.info(f"Data processing completed: {success_count}/{total_files} files processed")

if __name__ == "__main__":
    main()