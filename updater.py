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
CURRENT_VERSION = "1.0.1"  # Update this when you release a new version
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
    """Get the raw URL for a file in the GitHub repository with cache busting"""
    # Add cache-busting timestamp parameter to prevent caching
    timestamp = int(time.time())
    return f"https://raw.githubusercontent.com/{REPO_USER}/{REPO_NAME}/{REPO_BRANCH}/{file_path}?t={timestamp}"

def get_remote_version():
    """Get the version from the remote version.txt file"""
    try:
        logger.info(f"Checking remote version from {REPO_USER}/{REPO_NAME}")
        version_url = get_raw_file_url("version.txt")
        logger.info(f"Fetching version from: {version_url}")
        
        req = urllib.request.Request(
            version_url,
            headers={
                'User-Agent': 'EMR-Data-Mapper-Updater',
                'Cache-Control': 'no-cache, no-store, must-revalidate',
                'Pragma': 'no-cache',
                'Expires': '0'
            }
        )
        
        with urllib.request.urlopen(req, timeout=10) as response:
            logger.info(f"Got response from GitHub (status: {response.status})")
            content = response.read().decode('utf-8').strip()
            
            # Log the content
            logger.info(f"Version file content: {content}")
            
            # Simple validation check
            import re
            if re.match(r'^\d+\.\d+\.\d+$', content):
                logger.info(f"Found remote version: {content}")
                return content
            else:
                logger.warning(f"Invalid version format: {content}")
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

def create_backup():
    """Create a backup of the current installation"""
    try:
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        backup_dir = os.path.join(SCRIPT_DIR, f"backup-{timestamp}")
        
        logger.info(f"Creating backup directory: {backup_dir}")
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
        
        # Create core directory in backup if needed
        backup_core_dir = os.path.join(backup_dir, "core")
        if not os.path.exists(backup_core_dir):
            os.makedirs(backup_core_dir)
            logger.info(f"Created core directory in backup: {backup_core_dir}")
        
        # Copy each file to backup
        backup_count = 0
        for file_info in FILES_TO_UPDATE:
            local_path = file_info["local"]
            if os.path.exists(local_path):
                # Determine the backup path
                if "core/" in file_info["path"]:
                    backup_path = os.path.join(backup_core_dir, os.path.basename(local_path))
                else:
                    backup_path = os.path.join(backup_dir, os.path.basename(local_path))
                
                # Copy the file
                shutil.copy2(local_path, backup_path)
                logger.info(f"Backed up: {os.path.basename(local_path)}")
                backup_count += 1
        
        logger.info(f"Backup completed: {backup_count} files backed up")
        return backup_dir
    except Exception as e:
        logger.error(f"Error creating backup: {e}")
        return None

def download_file(file_path, local_path):
    """Download a single file from the GitHub repository"""
    try:
        file_url = get_raw_file_url(file_path)
        logger.info(f"Downloading {file_path} from {file_url}")
        
        req = urllib.request.Request(
            file_url,
            headers={
                'User-Agent': 'EMR-Data-Mapper-Updater'
            }
        )
        
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
    """Download and install the latest update"""
    logger.info(f"Starting update to version {version}")
    
    try:
        # Create a backup of current files
        logger.info("Creating backup before update")
        backup_dir = create_backup()
        if not backup_dir:
            logger.error("Failed to create backup. Update aborted.")
            return False
        
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
        
        if success_count == total_files:
            logger.info(f"Update completed successfully: {success_count}/{total_files} files updated")
            return True
        else:
            logger.warning(f"Partial update: {success_count}/{total_files} files updated")
            
            if success_count < total_files / 2:
                logger.warning("Less than half of files updated. Restoring from backup...")
                restore_from_backup(backup_dir)
                return False
            
            return True
            
    except Exception as e:
        logger.error(f"Error during update: {e}")
        return False

def restore_from_backup(backup_dir):
    """Restore files from backup after failed update"""
    try:
        logger.info(f"Restoring from backup: {backup_dir}")
        
        # Copy each file from backup
        restored_count = 0
        
        for file_info in FILES_TO_UPDATE:
            local_path = file_info["local"]
            
            # Determine the backup path
            if "core/" in file_info["path"]:
                backup_path = os.path.join(backup_dir, "core", os.path.basename(local_path))
            else:
                backup_path = os.path.join(backup_dir, os.path.basename(local_path))
            
            if os.path.exists(backup_path):
                # Ensure the target directory exists
                os.makedirs(os.path.dirname(local_path), exist_ok=True)
                
                # Copy the file back
                shutil.copy2(backup_path, local_path)
                restored_count += 1
                logger.info(f"Restored: {os.path.basename(local_path)}")
        
        logger.info(f"Restoration complete: {restored_count} files restored")
        return True
    except Exception as e:
        logger.error(f"Error restoring from backup: {e}")
        return False

def run_updater():
    """Main function to run the updater"""
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
    print(f"{Fore.YELLOW}{Style.BRIGHT}{'EMR DATA MAPPER UPDATER':^80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}\n")
    
    logger.info("Starting updater")
    
    # Check for updates
    update_version = check_for_updates()
    if update_version:
        logger.info(f"Update available: {update_version}")
        do_update = input(f"{Fore.YELLOW}Would you like to update to version {update_version}? (y/n): ")
        logger.info(f"User response to update prompt: {do_update}")
        
        if do_update.lower() == 'y':
            if download_update(update_version):
                logger.info("Update completed successfully. Application will restart.")
                # Return exit code 1 to signal update was performed
                sys.exit(1)
            else:
                logger.warning("Update failed. Continuing with current version.")
    else:
        logger.info("No updates available")
    
    # Return exit code 0 to signal no update needed
    sys.exit(0)

if __name__ == "__main__":
    run_updater()
