import os
import sys
import time
import urllib.request
import json
import shutil
import logging
from datetime import datetime

# Check if colorama is available for nice output
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

# Custom logging formatter with colors - matching the style in other modules
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
            if "Checking" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.BLUE}🔍 "
            elif "Downloaded" in record.msg or "Saved" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.GREEN}📥 "
            elif "Backed up" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.CYAN}📦 "
            elif "Restored" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.GREEN}📤 "
            elif "Created" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.GREEN}✓ "
            elif "Version" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.YELLOW}🏷️ "
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

# Configure logging with pretty colors and file output
console_handler = logging.StreamHandler()
console_handler.setFormatter(ColoredFormatter())

file_handler = logging.FileHandler(os.path.join(CORE_DIR, "updater.log"))
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

logger = logging.getLogger("Updater")
logger.setLevel(logging.INFO)
logger.addHandler(console_handler)
logger.addHandler(file_handler)

# Version information
CURRENT_VERSION = "1.0.5"  # Update this when you release a new version
REPO_USER = "john-mayhem"
REPO_NAME = "emr-dcl"
REPO_BRANCH = "main"

# File list to check and update
FILES_TO_UPDATE = [
    {"path": "core/comparator.py", "local": os.path.join(CORE_DIR, "comparator.py")},
    {"path": "core/data_processor.py", "local": os.path.join(CORE_DIR, "data_processor.py")},
    {"path": "core/emr_data_collector.py", "local": os.path.join(CORE_DIR, "emr_data_collector.py")},
    {"path": "core/google_sheets_collector.py", "local": os.path.join(CORE_DIR, "google_sheets_collector.py")},
    {"path": "core/kml_generator.py", "local": os.path.join(CORE_DIR, "kml_generator.py")},
    {"path": "MapScript.bat", "local": os.path.join(SCRIPT_DIR, "MapScript.bat")},
    {"path": "main-launcher.py", "local": os.path.join(SCRIPT_DIR, "main-launcher.py")},
    {"path": "updater.py", "local": os.path.join(SCRIPT_DIR, "updater.py")}
]

def get_raw_file_url(file_path):
    """Get the raw URL for a file in the GitHub repository with aggressive cache busting"""
    # Add random number to completely prevent caching
    random_param = str(time.time()) + str(os.urandom(4).hex())
    return f"https://raw.githubusercontent.com/{REPO_USER}/{REPO_NAME}/{REPO_BRANCH}/{file_path}?nocache={random_param}"

def get_remote_version():
    """Get the version from the remote version.txt file with aggressive cache prevention"""
    try:
        logger.info(f"Checking remote version from {REPO_USER}/{REPO_NAME}")
        version_url = get_raw_file_url("version.txt")
        logger.info(f"Fetching version from: {version_url}")
        
        # Create a completely fresh request with aggressive cache prevention
        req = urllib.request.Request(
            version_url,
            headers={
                'User-Agent': f'EMR-Data-Mapper-Updater-{os.urandom(4).hex()}',
                'Cache-Control': 'no-cache, no-store, must-revalidate, max-age=0',
                'Pragma': 'no-cache',
                'Expires': '-1'
            }
        )
        
        # Create a custom opener that ignores caches
        opener = urllib.request.build_opener(urllib.request.HTTPHandler())
        urllib.request.install_opener(opener)
        
        with urllib.request.urlopen(req, timeout=10) as response:
            logger.info(f"Got response from GitHub (status: {response.status})")
            content = response.read().decode('utf-8').strip()
            
            # Clear any connection pool that might be caching
            opener.close()
            
            # Log the content
            logger.info(f"Version file content: '{content}'")
            
            # Simple validation check
            import re
            if re.match(r'^\d+\.\d+\.\d+$', content):
                logger.info(f"Found remote version: {content}")
                return content
            else:
                logger.warning(f"Invalid version format: '{content}'")
                return None
            
    except Exception as e:
        logger.error(f"Error checking remote version: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return None

def check_for_updates():
    """Check if a newer version is available"""
    logger.info("Starting update check")
    
    try:
        logger.info(f"Current version: {CURRENT_VERSION}")
        remote_version = get_remote_version()
        
        if not remote_version:
            logger.warning("Could not determine remote version")
            return False
        
        logger.info(f"Remote version: {remote_version}")
        
        # Compare versions (simple string comparison for now)
        if remote_version > CURRENT_VERSION:
            logger.info(f"Update available: v{remote_version}")
            return remote_version
        else:
            logger.info("You have the latest version")
            return False
    except Exception as e:
        logger.error(f"Error checking for updates: {e}")
        return False

def download_file(file_path, local_path):
    """Download a single file from the GitHub repository with cache prevention"""
    try:
        file_url = get_raw_file_url(file_path)
        logger.info(f"Downloading {file_path} from {file_url}")
        
        req = urllib.request.Request(
            file_url,
            headers={
                'User-Agent': f'EMR-Data-Mapper-Downloader-{os.urandom(4).hex()}',
                'Cache-Control': 'no-cache, no-store, must-revalidate, max-age=0',
                'Pragma': 'no-cache',
                'Expires': '-1'
            }
        )
        
        # Create a custom opener that ignores caches
        opener = urllib.request.build_opener(urllib.request.HTTPHandler())
        urllib.request.install_opener(opener)
        
        # Ensure the directory exists
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        
        # Download the file
        with urllib.request.urlopen(req, timeout=10) as response:
            logger.info(f"Received response (status: {response.status})")
            content = response.read()
            logger.info(f"Downloaded {len(content)} bytes")
            
            with open(local_path, 'wb') as f:
                f.write(content)
            
            logger.info(f"Saved to {local_path}")
            return True
    except Exception as e:
        logger.error(f"Error downloading {file_path}: {e}")
        return False

def download_update(version):
    """Download and install the latest update without creating backups"""
    logger.info(f"Starting update to version {version}")
    
    try:
        # Download each file
        success_count = 0
        total_files = len(FILES_TO_UPDATE)
        
        for i, file_info in enumerate(FILES_TO_UPDATE):
            file_path = file_info["path"]
            local_path = file_info["local"]
            
            logger.info(f"[{i+1}/{total_files}] Updating {file_path}")
            
            if download_file(file_path, local_path):
                success_count += 1
                logger.info(f"Successfully updated {file_path}")
            else:
                logger.error(f"Failed to update {file_path}")
        
        if success_count > 0:
            logger.info(f"Update completed: {success_count}/{total_files} files updated")
            return True
        else:
            logger.warning(f"Update failed: No files were updated")
            return False
            
    except Exception as e:
        logger.error(f"Error during update: {e}")
        return False


def run_updater():
    """Main function to run the updater without user interaction"""
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
    print(f"{Fore.YELLOW}{Style.BRIGHT}{'EMR DATA MAPPER UPDATER':^80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}\n")
    
    logger.info("Starting updater")
    
    # Check for updates
    update_version = check_for_updates()
    if update_version:
        logger.info(f"Update available: {update_version}")
        print(f"{Fore.GREEN}Update available! Automatically updating from v{CURRENT_VERSION} to v{update_version}...")
        
        if download_update(update_version):
            logger.info("Update completed successfully. Application will restart.")
            print(f"{Fore.GREEN}Update completed successfully! The application will now restart.")
            # Return exit code 1 to signal update was performed
            sys.exit(1)
        else:
            logger.warning("Update failed.")
            print(f"{Fore.RED}Update failed. Continuing with current version.")
    else:
        logger.info("No updates available")
    
    # Return exit code 0 to signal no update needed
    sys.exit(0)

if __name__ == "__main__":
    run_updater()
