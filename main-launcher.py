import os
import sys
import time
import subprocess
import importlib.util
import shutil

# First, check if colorama is available for nice output
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

# Set up directory paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CORE_DIR = os.path.join(SCRIPT_DIR, "core")

# Create core directory if it doesn't exist
if not os.path.exists(CORE_DIR):
    os.makedirs(CORE_DIR)
    print(f"Created core directory at {CORE_DIR}")

# Configuration - scripts to run in order
SCRIPTS = [

]

def print_header():
    """Print a colorful header for the launcher"""
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
    print(f"{Fore.YELLOW}{Style.BRIGHT}{'DATA COLLECTION LAUNCHER':^80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}\n")

def check_script_exists(script_path):
    """Check if a script file exists"""
    return os.path.exists(script_path)

def check_dependencies():
    """Check for required Python dependencies"""
    required_packages = ['colorama']
    missing_packages = []
    
    print("Checking basic dependencies...")
    
    for package in required_packages:
        if importlib.util.find_spec(package) is None:
            missing_packages.append(package)
    
    if missing_packages:
        print(f"Missing required packages: {', '.join(missing_packages)}")
        install = input("Would you like to install them now? (y/n): ")
        
        if install.lower() == 'y':
            try:
                for package in missing_packages:
                    print(f"Installing {package}...")
                    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                print("All dependencies installed successfully!")
                
                # Re-import colorama if it was installed
                if 'colorama' in missing_packages:
                    global colorama, Fore, Back, Style, HAS_COLOR
                    import colorama
                    from colorama import Fore, Back, Style
                    colorama.init(autoreset=True)
                    HAS_COLOR = True
                    
            except Exception as e:
                print(f"Error installing packages: {e}")
                print("Please install the required packages manually:")
                print(f"pip install {' '.join(missing_packages)}")
                return False
        else:
            print("Please install the required packages before running the script:")
            print(f"pip install {' '.join(missing_packages)}")
            return False
    
    return True

def create_config_if_needed():
    """Create the config.py file if it doesn't exist"""
    config_path = os.path.join(CORE_DIR, 'config.py')
    if not os.path.exists(config_path):
        print(f"{Fore.YELLOW}Configuration file not found. Creating default config.py file...")
        with open(config_path, 'w') as f:
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

# Google Sheets Configuration
# URLs for the Google Sheets to access
GOOGLE_SHEETS = {
    # Therapists sheet
    "therapists": {
        "url": "https://docs.google.com/spreadsheets/d/1H0YTVkNLjowHmVZkNrnKs-CxQNvkQBoaoCc6kRmhr6U",
        "worksheet": "ActiveTherapists",
        "output_file": "ActiveTherapists.csv",
        "columns": 3  # Number of columns to keep (first N columns)
    },
    # Need Staff sheet
    "need_staff": {
        "url": "https://docs.google.com/spreadsheets/d/1YecC12XUwp9wY9bRq26Y1YAgMGxBGfT-aZDKQ--ovlQ",
        "worksheet": "Need Staff-Feb 2024",
        "output_file": "NeedStaff.csv",
        "columns": 5  # Number of columns to keep (first N columns)
    }
}
''')
        print(f"{Fore.GREEN}Default configuration file created at {os.path.abspath(config_path)}")
        print(f"{Fore.YELLOW}Please update the configuration with your settings before running again.")
        update_config = input("Would you like to edit the configuration now? (y/n): ")
        if update_config.lower() == 'y':
            try:
                if sys.platform == 'win32':
                    os.system(f'notepad.exe "{config_path}"')
                elif sys.platform == 'darwin':  # macOS
                    os.system(f'open -t "{config_path}"')
                else:  # Linux or other Unix-like systems
                    editors = ['nano', 'vim', 'vi', 'gedit']
                    for editor in editors:
                        if os.system(f'which {editor} > /dev/null 2>&1') == 0:
                            os.system(f'{editor} "{config_path}"')
                            break
            except Exception as e:
                print(f"Error opening editor: {e}")
                print(f"Please edit {config_path} manually with a text editor.")
        return False
    return True

def run_script(script):
    """Run a Python script and return success status"""
    script_path = os.path.join(CORE_DIR, script["file"])
    
    if not os.path.exists(script_path):
        print(f"{Fore.RED}Error: Script '{script_path}' not found.")
        return False
    
    print(f"\n{Fore.CYAN}{'─'*80}")
    print(f"{Fore.YELLOW}Running: {Fore.CYAN}{script['name']}")
    print(f"{Fore.YELLOW}Description: {Fore.WHITE}{script['description']}")
    print(f"{Fore.CYAN}{'─'*80}{Style.RESET_ALL}")
    
    try:
        print(f"{Fore.YELLOW}Starting {script['name']}...")
        # Set PYTHONPATH to include core directory
        env = os.environ.copy()
        if 'PYTHONPATH' in env:
            env['PYTHONPATH'] = f"{CORE_DIR}{os.pathsep}{env['PYTHONPATH']}"
        else:
            env['PYTHONPATH'] = CORE_DIR
            
        # Run the script
        result = subprocess.run([sys.executable, script_path], 
                                env=env, 
                                cwd=CORE_DIR,
                                check=True)
        
        if result.returncode == 0:
            print(f"{Fore.GREEN}✓ {Style.BRIGHT}{script['name']} completed successfully.")
            return True
        else:
            print(f"{Fore.RED}✗ {Style.BRIGHT}{script['name']} failed with return code {result.returncode}.")
            return False
    except subprocess.CalledProcessError as e:
        print(f"{Fore.RED}✗ {Style.BRIGHT}{script['name']} failed with return code {e.returncode}.")
        return False
    except Exception as e:
        print(f"{Fore.RED}✗ {Style.BRIGHT}Error running {script['name']}: {e}")
        return False

def copy_kml_directory():
    """Copy the KML directory from core to parent directory"""
    source_kml_dir = os.path.join(CORE_DIR, "kml")
    target_kml_dir = os.path.join(SCRIPT_DIR, "KML")
    
    if not os.path.exists(source_kml_dir):
        print(f"{Fore.YELLOW}⚠️ KML directory not found in core. Skipping copy operation.")
        return False
    
    # Create target directory if it doesn't exist
    if not os.path.exists(target_kml_dir):
        os.makedirs(target_kml_dir)
        print(f"{Fore.CYAN}Created target KML directory: {target_kml_dir}")
    
    # Copy all files from source to target
    try:
        file_count = 0
        for filename in os.listdir(source_kml_dir):
            if filename.endswith(".kml"):
                source_file = os.path.join(source_kml_dir, filename)
                target_file = os.path.join(target_kml_dir, filename)
                shutil.copy2(source_file, target_file)
                file_count += 1
                print(f"{Fore.GREEN}Copied: {filename}")
        
        if file_count > 0:
            print(f"{Fore.GREEN}✓ {Style.BRIGHT}Successfully copied {file_count} KML files to {target_kml_dir}")
            return True
        else:
            print(f"{Fore.YELLOW}⚠️ No KML files found to copy.")
            return False
    except Exception as e:
        print(f"{Fore.RED}✗ {Style.BRIGHT}Error copying KML files: {e}")
        return False

def check_for_updates():
    """Check for updates using the updater module"""
    updater_path = os.path.join(SCRIPT_DIR, "updater.py")
    
    if not os.path.exists(updater_path):
        print(f"{Fore.YELLOW}Updater script not found. Skipping update check.")
        return False
    
    print(f"\n{Fore.CYAN}{'─'*80}")
    print(f"{Fore.YELLOW}Checking for updates...")
    print(f"{Fore.CYAN}{'─'*80}{Style.RESET_ALL}")
    
    try:
        # Run the updater with direct subprocess call instead of capturing output
        # This allows the updater to handle its own input/output correctly
        result = subprocess.call([sys.executable, updater_path])
        
        # If updater returns 0, no update needed or user declined
        # If it returns 1, an update was performed and we should restart
        return result == 1
    except Exception as e:
        print(f"{Fore.RED}Error checking for updates: {e}")
        return False

def main():
    """Main function to run the launcher"""
    print_header()
    
    # Check for updates first
    should_restart = check_for_updates()
    if should_restart:
        print(f"{Fore.GREEN}Update completed. Restarting application...")
        os.execv(sys.executable, ['python'] + sys.argv)
        return
    
    # Check dependencies
    if not check_dependencies():
        sys.exit(1)
    
    # Create config file if needed
    if not create_config_if_needed():
        return
    
    # Run all scripts in sequence
    total_scripts = len(SCRIPTS)
    successful = 0
    
    total_start_time = time.time()
    
    for i, script in enumerate(SCRIPTS):
        print(f"\n{Fore.YELLOW}[{i+1}/{total_scripts}] Running {script['name']}...")
        if run_script(script):
            successful += 1
    
    # Copy KML files to parent directory
    print(f"\n{Fore.CYAN}{'─'*80}")
    print(f"{Fore.YELLOW}Final Step: Copying KML files to parent directory")
    print(f"{Fore.CYAN}{'─'*80}{Style.RESET_ALL}")
    
    copy_kml_directory()
    
    total_duration = time.time() - total_start_time
    minutes = int(total_duration // 60)
    seconds = int(total_duration % 60)
    
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
    print(f"{Fore.YELLOW}{Style.BRIGHT}{f'ALL TASKS COMPLETED: {successful}/{total_scripts} SCRIPTS':^80}")
    print(f"{Fore.GREEN}{Style.BRIGHT}{f'Total time: {minutes} minutes {seconds} seconds':^80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}")

if __name__ == "__main__":
    main()
