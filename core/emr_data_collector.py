import os
import sys
import time
import logging
import shutil
import importlib.util
from datetime import datetime

# Function to check dependencies at startup
def check_dependencies():
    """Check and install required dependencies if missing"""
    required_packages = ['selenium', 'colorama']
    missing_packages = []
    
    print("Checking required dependencies...")
    
    for package in required_packages:
        if importlib.util.find_spec(package) is None:
            missing_packages.append(package)
    
    if missing_packages:
        print(f"Missing required packages: {', '.join(missing_packages)}")
        install = input("Would you like to install them now? (y/n): ")
        
        if install.lower() == 'y':
            try:
                import subprocess
                for package in missing_packages:
                    print(f"Installing {package}...")
                    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                print("All dependencies installed successfully!")
            except Exception as e:
                print(f"Error installing packages: {e}")
                print("Please install the required packages manually:")
                print(f"pip install {' '.join(missing_packages)}")
                sys.exit(1)
        else:
            print("Please install the required packages before running the script:")
            print(f"pip install {' '.join(missing_packages)}")
            sys.exit(1)
    else:
        print("All required dependencies are already installed.")

# Check dependencies first
check_dependencies()

# Now import the dependencies
import colorama
from colorama import Fore, Back, Style
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Initialize colorama
colorama.init(autoreset=True)

# Check for config file existence
CONFIG_FILE = 'config.py'
if not os.path.exists(CONFIG_FILE):
    print(f"{Fore.YELLOW}Configuration file not found. Creating default config.py file...")
    with open(CONFIG_FILE, 'w') as f:
        f.write('''# EMR Data Collection Configuration File

# User Credentials
USERNAME = "your_username"
PASSWORD = "your_password"
AUTO_LOGIN = False  # Set to True to use the credentials above for automatic login

# Browser Settings
HEADLESS_MODE = False  # Set to True to run browser in background without UI

# Offices and Disciplines to Process
OFFICES = [
    "Americare Certified Special Services Inc",
    "Four Seasons Home Health Care.",
    "Girling Health Care",
    "Personal Touch",
    "Rehab on Wheels OT PT PLLC",
    "Shining Star Home Care"
]

DISCIPLINES = ["PT", "OT"]

# Filter type to use for reports
FILTER_TYPE = "With Cases"

# Time delays for operations (in seconds)
DELAY_BETWEEN_UI_OPERATIONS = 0.3  # Reduced for faster UI interactions
DELAY_AFTER_PAGE_LOAD = 2.0
DELAY_BETWEEN_REPORTS = 2.0
DELAY_AFTER_LOGIN = 5.0
DELAY_BEFORE_EXPORT = 2.0

# Output Directory
# Leave empty to use default "EMR_Reports" in script directory
OUTPUT_DIR = ""
''')
    print(f"{Fore.GREEN}Default configuration file created at {os.path.abspath(CONFIG_FILE)}")
    print(f"{Fore.YELLOW}Please update the configuration with your settings before running the script again.")
    sys.exit(0)

# Import configuration
try:
    import config
    print(f"{Fore.GREEN}Configuration loaded successfully.")
except Exception as e:
    print(f"{Fore.RED}Error loading configuration: {e}")
    print(f"{Fore.YELLOW}Please check your config.py file for syntax errors.")
    sys.exit(1)

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
            elif "Waiting" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.CYAN}⏳ "
            elif "Downloaded" in record.msg or "Renamed" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.GREEN}📄 "
            elif "Button clicked" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.YELLOW}🖱️ "
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
        
        # Style the message content
        message = record.msg
        if "Processing report for Office" in message:
            parts = message.split("Office: ")
            if len(parts) > 1:
                office_part = parts[1].split(",")[0]
                message = f"Processing report for Office: {Fore.CYAN}{Style.BRIGHT}{office_part}{Style.RESET_ALL}{log_color}, Discipline: {parts[1].split('Discipline: ')[1].split(',')[0]}, Filter: {parts[1].split('Filter: ')[1]}"
        
        return f"{prefix}{message}"

# Configure logging with pretty colors
console_handler = logging.StreamHandler()
console_handler.setFormatter(ColoredFormatter())

file_handler = logging.FileHandler("emr_data_collection.log")
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

logger = logging.getLogger("EMRDataCollector")
logger.setLevel(logging.INFO)
logger.addHandler(console_handler)
logger.addHandler(file_handler)

# Get configuration values
USERNAME = config.USERNAME
PASSWORD = config.PASSWORD
AUTO_LOGIN = config.AUTO_LOGIN
HEADLESS_MODE = config.HEADLESS_MODE
OFFICES = config.OFFICES
DISCIPLINES = config.DISCIPLINES
FILTER_TYPE = config.FILTER_TYPE
DELAY_BETWEEN_UI_OPERATIONS = config.DELAY_BETWEEN_UI_OPERATIONS
DELAY_AFTER_PAGE_LOAD = config.DELAY_AFTER_PAGE_LOAD
DELAY_BETWEEN_REPORTS = config.DELAY_BETWEEN_REPORTS
DELAY_AFTER_LOGIN = config.DELAY_AFTER_LOGIN
DELAY_BEFORE_EXPORT = config.DELAY_BEFORE_EXPORT

# Configuration for output directory
if hasattr(config, 'OUTPUT_DIR') and config.OUTPUT_DIR:
    OUTPUT_DIR = config.OUTPUT_DIR
else:
    OUTPUT_DIR = os.path.join(os.getcwd(), "EMR_Reports")

# Create output directory if it doesn't exist
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)
    logger.info(f"Created output directory: {OUTPUT_DIR}")

def setup_driver():
    """Setup and return configured Chrome WebDriver"""
    chrome_options = Options()
    
    # Set download directory and disable security features
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": OUTPUT_DIR,  # Set download directory directly to output folder
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,  # Disable SafeBrowsing to avoid virus scan warnings
        "safebrowsing.disable_download_protection": True,  # Disable download protection
        "profile.default_content_setting_values.automatic_downloads": 1,  # Allow multiple downloads
        "browser.download.manager.showWhenStarting": False,
        "browser.download.manager.focusWhenStarting": False,
        "browser.helperApps.neverAsk.saveToDisk": "application/vnd.ms-excel;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    })
    
    # Add additional Chrome arguments to disable security features
    chrome_options.add_argument("--disable-web-security")
    chrome_options.add_argument("--allow-running-insecure-content")
    chrome_options.add_argument("--disable-features=IsolateOrigins,site-per-process")
    
    # Add headless mode if configured
    if HEADLESS_MODE:
        chrome_options.add_argument("--headless=new")
        logger.info("Running in headless mode (no browser UI)")
    
    # Create a new Chrome WebDriver
    service = Service()
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.maximize_window()
    return driver

def perform_login(driver):
    """Attempt to log in using the provided credentials"""
    if not AUTO_LOGIN:
        logger.info("Automatic login disabled. Please log in manually and press Enter...")
        input()
        return True
    
    try:
        logger.info(f"Attempting to log in as {USERNAME}...")
        
        # Find username field and enter username
        username_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "userNameOrEmailAddress"))
        )
        username_field.clear()
        username_field.send_keys(USERNAME)
        
        # Find password field and enter password
        password_field = driver.find_element(By.NAME, "password")
        password_field.clear()
        password_field.send_keys(PASSWORD)
        
        # Click login button
        login_button = driver.find_element(By.XPATH, "//button[@type='submit' and contains(text(), 'Log in')]")
        login_button.click()
        
        # Wait for login to complete - look for an element that appears after successful login
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".card-body"))
        )
        
        logger.info("Login successful")
        return True
    except Exception as e:
        logger.error(f"Login failed: {e}")
        logger.info("Please log in manually and press Enter...")
        input()
        return True

def set_dropdown_value(driver, dropdown_name, target_value):
    """
    Set dropdown to target value using fast JavaScript execution.
    This function uses a direct approach to minimize waiting time.
    """
    try:
        # One-step approach to find and open dropdown
        js_open = f"""
        // Find label and associated dropdown
        const label = Array.from(document.querySelectorAll('label')).find(el => el.textContent.trim() === '{dropdown_name}');
        if (!label) return "Label not found";
        
        // Find and click the dropdown trigger
        const parentDiv = label.parentElement;
        const trigger = parentDiv.querySelector('.p-dropdown-trigger');
        if (trigger) {{
            trigger.click();
            return "Opened dropdown";
        }}
        return "Trigger not found";
        """
        
        result = driver.execute_script(js_open)
        logger.info(f"Open {dropdown_name} dropdown: {result}")
        time.sleep(DELAY_BETWEEN_UI_OPERATIONS)
        
        # Select the target option
        js_select = f"""
        // Find the dropdown panel
        const panel = document.querySelector('.p-dropdown-panel');
        if (!panel) return "No panel found";
        
        // Find the option
        const options = Array.from(panel.querySelectorAll('li'));
        const target = options.find(li => li.textContent.trim() === '{target_value}');
        
        if (target) {{
            target.click();
            return "Selected option";
        }}
        
        return "Option not found: " + options.map(o => o.textContent.trim()).join(", ");
        """
        
        result = driver.execute_script(js_select)
        logger.info(f"Select {target_value} in {dropdown_name}: {result}")
        time.sleep(DELAY_BETWEEN_UI_OPERATIONS)
        
        return "Selected option" in result
    except Exception as e:
        logger.error(f"Error setting {dropdown_name} to {target_value}: {e}")
        return False

def click_button_by_text(driver, button_text):
    """Click a button using its text content via JavaScript for reliability"""
    # Try multiple approaches to click the button
    js_approaches = [
        # Approach 1: Find button by text content
        f"""
        const buttons = Array.from(document.querySelectorAll('button'));
        const targetButton = buttons.find(btn => btn.textContent.trim().includes('{button_text}'));
        if (targetButton) {{
            targetButton.click();
            return "Button clicked via standard approach";
        }}
        return "Button not found via standard approach";
        """,
        
        # Approach 2: Try to find by class/attributes that might identify the export button
        f"""
        const exportBtn = document.querySelector('button.btn-outline-success');
        if (exportBtn) {{
            exportBtn.click();
            return "Button clicked via class selector";
        }}
        return "Button not found via class selector";
        """,
        
        # Approach 3: Try a more direct event dispatch
        f"""
        const buttons = Array.from(document.querySelectorAll('button'));
        const targetButton = buttons.find(btn => btn.textContent.trim().includes('{button_text}'));
        if (targetButton) {{
            // Try to trigger events that may help with download dialogs
            const clickEvent = new MouseEvent('click', {{
                bubbles: true,
                cancelable: true,
                view: window
            }});
            
            targetButton.dispatchEvent(clickEvent);
            return "Button clicked via event dispatch";
        }}
        return "Button not found for event dispatch";
        """
    ]
    
    # Try each approach
    for i, js in enumerate(js_approaches):
        try:
            result = driver.execute_script(js)
            logger.info(f"Click button approach {i+1} result: {result}")
            if "clicked" in result:
                time.sleep(DELAY_BETWEEN_UI_OPERATIONS)
                return True
        except Exception as e:
            logger.error(f"Error in button click approach {i+1}: {e}")
    
    # If all JavaScript approaches failed, try the standard Selenium approach
    try:
        button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, f"//button[contains(text(), '{button_text}')]"))
        )
        button.click()
        logger.info("Button clicked via Selenium")
        time.sleep(DELAY_BETWEEN_UI_OPERATIONS)
        return True
    except Exception as e:
        logger.error(f"Error clicking button via Selenium: {e}")
        return False

def verify_and_rename_downloaded_file(office, discipline, filter_type):
    """Verify that the file was downloaded and rename it with descriptive info"""
    try:
        # Wait for the file to appear in the download directory
        max_wait_time = 30  # Increased timeout for virus scans
        start_time = time.time()
        
        # Look for the most recent Excel file in the output directory
        while time.time() - start_time < max_wait_time:
            excel_files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith('.xlsx') or
                          (f.endswith('.xlsx.crdownload') or f.endswith('.xlsx.tmp'))]
            
            # First check for complete downloads
            complete_files = [f for f in excel_files if f.endswith('.xlsx')]
            if complete_files:
                # Sort by modification time, newest first
                complete_files.sort(key=lambda x: os.path.getmtime(os.path.join(OUTPUT_DIR, x)), reverse=True)
                most_recent_file = complete_files[0]
                
                # Handle file rename
                if "ActivePatientsReport" in most_recent_file:
                    # Clean office name for filename
                    office_clean = office.replace(" ", "_").replace(".", "").replace(",", "")
                    
                    # Create the new filename
                    new_filename = f"{office_clean}_{discipline}_{filter_type}_ActivePatients.xlsx"
                    old_path = os.path.join(OUTPUT_DIR, most_recent_file)
                    new_path = os.path.join(OUTPUT_DIR, new_filename)
                    
                    # Make sure the file is complete and not being scanned
                    try:
                        # Attempt to open and read the file to ensure it's not locked
                        with open(old_path, 'rb') as f:
                            # Just read a small portion to check if file is accessible
                            data = f.read(1024)
                            
                        # Check if target already exists and remove if needed
                        if os.path.exists(new_path):
                            os.remove(new_path)
                        
                        # Rename the file
                        os.rename(old_path, new_path)
                        logger.info(f"Renamed {most_recent_file} to {new_filename}")
                        return True
                    except (PermissionError, IOError) as e:
                        # File might still be locked by virus scanner or download process
                        logger.info(f"File not yet accessible, waiting: {e}")
                        time.sleep(1)
                        continue
            
            # Check if downloads are in progress
            if any(f.endswith('.xlsx.crdownload') or f.endswith('.xlsx.tmp') for f in excel_files):
                logger.info("Download still in progress, waiting...")
                time.sleep(1)
                continue
                
            # If no files found yet, brief wait
            time.sleep(0.5)
        
        logger.error("Timed out waiting for downloaded file")
        return False
    except Exception as e:
        logger.error(f"Error verifying/renaming file: {e}")
        return False

def process_report(driver, office, discipline, filter_type="With Cases"):
    """Process a single report for the given office, discipline, and filter"""
    try:
        logger.info(f"Processing report for Office: {office}, Discipline: {discipline}, Filter: {filter_type}")
        
        # Load the reports page
        driver.get("https://emr.appv2.hellonote.com/app/main/reports/activePatients")
        time.sleep(DELAY_AFTER_PAGE_LOAD)
        
        # Try a more direct JavaScript approach to set all values at once
        # This can be much faster than individual dropdown interactions
        js_set_all_values = f"""
        // Function to set a dropdown value
        function setDropdownValue(labelText, targetValue) {{
            // Find the label and associated dropdown
            const label = Array.from(document.querySelectorAll('label'))
                .find(el => el.textContent.trim() === labelText);
            if (!label) return `Label "${{labelText}}" not found`;
            
            // Find parent div and dropdown component
            const parentDiv = label.parentElement;
            const dropdownTrigger = parentDiv.querySelector('.p-dropdown-trigger');
            
            if (dropdownTrigger) {{
                // Click to open dropdown
                dropdownTrigger.click();
                
                // Small delay to let dropdown open
                setTimeout(() => {{
                    // Find dropdown panel and options
                    const panel = document.querySelector('.p-dropdown-panel');
                    if (!panel) return `Panel for "${{labelText}}" not found`;
                    
                    const options = Array.from(panel.querySelectorAll('li'));
                    const target = options.find(li => li.textContent.trim() === targetValue);
                    
                    if (target) {{
                        // Click the target option
                        target.click();
                        return `Set ${{labelText}} to ${{targetValue}}`;
                    }} else {{
                        return `Option "${{targetValue}}" not found for ${{labelText}}`;
                    }}
                }}, 100);
            }}
        }}
        
        // Set the values with small delays between actions
        setDropdownValue("Office", "{office}");
        
        setTimeout(() => {{
            setDropdownValue("Filter", "{filter_type}");
            
            setTimeout(() => {{
                setDropdownValue("Discipline", "{discipline}");
            }}, 300);
        }}, 300);
        
        return "Setting values...";
        """
        
        # Try the fast JavaScript approach first
        driver.execute_script(js_set_all_values)
        logger.info("Attempted to set all dropdown values with JavaScript")
        time.sleep(1.5)  # Give time for the dropdowns to update
        
        # As a backup, try individual dropdown setting if needed
        # Check if the office value is correctly set
        office_set_correctly = driver.execute_script(f"""
            const officeLabel = document.querySelector('p-dropdown[name="modal_organizationUnitId"] .p-dropdown-label');
            return officeLabel && officeLabel.textContent.trim() === '{office}';
        """)
        
        if not office_set_correctly:
            logger.info("JavaScript approach didn't set office correctly, using fallback approach")
            # Set Office dropdown
            if not set_dropdown_value(driver, "Office", office):
                logger.error(f"Failed to select office: {office}")
                return False
            
            # Set Filter dropdown
            if not set_dropdown_value(driver, "Filter", filter_type):
                logger.error(f"Failed to select filter: {filter_type}")
                return False
            
            # Set Discipline dropdown
            if not set_dropdown_value(driver, "Discipline", discipline):
                logger.error(f"Failed to select discipline: {discipline}")
                return False
        
        # Add a short delay to ensure UI is ready
        time.sleep(DELAY_BEFORE_EXPORT)
        
        # Clear any existing downloads for this combination
        office_clean = office.replace(" ", "_").replace(".", "").replace(",", "")
        existing_file = os.path.join(OUTPUT_DIR, f"{office_clean}_{discipline}_{filter_type}_ActivePatients.xlsx")
        if os.path.exists(existing_file):
            try:
                os.remove(existing_file)
                logger.info(f"Removed existing file: {existing_file}")
            except Exception as e:
                logger.warning(f"Could not remove existing file: {e}")
        
        # Add specific code to help with the browser security warning
        try:
            # This can sometimes clear the browser's security popup or warnings
            driver.execute_script("""
            // Try to dismiss any active security dialogs
            document.body.dispatchEvent(new KeyboardEvent('keydown', {key: 'Escape', bubbles: true}));
            document.body.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', bubbles: true}));
            """)
        except Exception:
            pass
        
        # Export to Excel
        if not click_button_by_text(driver, "Export to excel"):
            logger.error("Failed to click Export to excel button")
            return False
        
        # Verify and rename the downloaded file
        if not verify_and_rename_downloaded_file(office, discipline, filter_type):
            logger.error("Failed to verify or rename downloaded file")
            return False
        
        logger.info(f"Successfully processed report for {office} - {discipline} - {filter_type}")
        return True
        
    except Exception as e:
        logger.error(f"Error processing report for {office} - {discipline}: {e}")
        return False

def main():
    """Main function to run the data collection process"""
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
    print(f"{Fore.YELLOW}{Style.BRIGHT}{'EMR DATA COLLECTION SCRIPT':^80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}\n")
    
    logger.info("Starting EMR data collection process")
    
    driver = None
    try:
        offices_count = len(OFFICES)
        disciplines_count = len(DISCIPLINES)
        total_reports = offices_count * disciplines_count
        
        print(f"{Fore.CYAN}▶ Will process {Fore.YELLOW}{total_reports}{Fore.CYAN} reports:")
        print(f"{Fore.CYAN}  ├─ {Fore.WHITE}{offices_count} offices: {Fore.YELLOW}{', '.join(OFFICES)}")
        print(f"{Fore.CYAN}  └─ {Fore.WHITE}{disciplines_count} disciplines: {Fore.YELLOW}{', '.join(DISCIPLINES)}\n")
        
        driver = setup_driver()
        
        # Open the EMR login page
        driver.get("https://emr.appv2.hellonote.com/")
        logger.info("Opened EMR login page")
        
        # Handle login
        if not perform_login(driver):
            logger.error("Login process failed")
            return
        
        # Add significant delay after login to let the browser fully initialize
        logger.info(f"Waiting {DELAY_AFTER_LOGIN} seconds after login to stabilize...")
        time.sleep(DELAY_AFTER_LOGIN)
            
        # After login, add a certificate exception if needed
        try:
            # This JavaScript can help bypass certificate warnings in some cases
            driver.execute_script("""
            try {
                // For Chrome's "Your connection is not private" warnings
                const classNames = ['secondary-button', 'ssl-opt-in', 'advanced-button'];
                for (const className of classNames) {
                    const button = document.querySelector(`.${className}`);
                    if (button) button.click();
                }
                
                // For "Proceed to site" type links
                const proceedLinks = document.querySelectorAll('a');
                for (const link of proceedLinks) {
                    if (link.textContent.includes('Proceed') || 
                        link.textContent.includes('Continue') || 
                        link.textContent.includes('Advanced')) {
                        link.click();
                    }
                }
            } catch (e) {
                // Ignore errors if elements not found
            }
            """)
        except Exception:
            # Ignore errors, this is just a precaution
            pass
        
        # Process each office and discipline
        completed = 0
        total_start_time = time.time()
        
        for office_idx, office in enumerate(OFFICES):
            for discipline_idx, discipline in enumerate(DISCIPLINES):
                report_num = office_idx * len(DISCIPLINES) + discipline_idx + 1
                
                print(f"\n{Fore.CYAN}{'─'*80}")
                print(f"{Fore.YELLOW}[{report_num}/{total_reports}] Processing: {Fore.CYAN}{office} {Fore.MAGENTA}- {Fore.GREEN}{discipline}")
                print(f"{Fore.CYAN}{'─'*80}{Style.RESET_ALL}")
                
                start_time = time.time()
                success = process_report(driver, office, discipline, FILTER_TYPE)
                end_time = time.time()
                
                if success:
                    completed += 1
                    duration = end_time - start_time
                    print(f"{Fore.GREEN}✓ {Style.BRIGHT}Report completed in {duration:.2f} seconds ({completed}/{total_reports} done)")
                else:
                    print(f"{Fore.RED}✗ {Style.BRIGHT}Failed to process report ({completed}/{total_reports} done)")
                    logger.warning(f"Failed to process {office} - {discipline}, moving to next combination")
                
                time.sleep(DELAY_BETWEEN_REPORTS)
        
        total_duration = time.time() - total_start_time
        minutes = int(total_duration // 60)
        seconds = int(total_duration % 60)
        
        print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
        print(f"{Fore.YELLOW}{Style.BRIGHT}{f'DATA COLLECTION COMPLETED: {completed}/{total_reports} REPORTS':^80}")
        print(f"{Fore.GREEN}{Style.BRIGHT}{f'Total time: {minutes} minutes {seconds} seconds':^80}")
        print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}\n")
        
        logger.info(f"Data collection completed: {completed}/{total_reports} reports processed")
        
    except Exception as e:
        logger.error(f"An error occurred during data collection: {e}")
    finally:
        if driver:
            driver.quit()
            logger.info("WebDriver closed")

if __name__ == "__main__":
    main()