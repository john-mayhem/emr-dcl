import os
import sys
import time
import subprocess
import importlib.util
import shutil
import urllib.request
import json
import zipfile
import tempfile

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

# Version information
CURRENT_VERSION = "1.0.0"  # Update this when you release a new version
GITHUB_REPO = "https://api.github.com/repos/YOUR_USERNAME/YOUR_REPO"  # Replace with your repo details
GITHUB_RELEASE_URL = f"{GITHUB_REPO}/releases/latest"
GITHUB_ZIPBALL_URL = f"{GITHUB_REPO}/zipball/main"  # Or master, depending on your branch name

def check_for_updates():
    """Check if a newer version is available"""
    print(f"{Fore.CYAN}Checking for updates...")
    
    try:
        # Setup request with appropriate headers
        req = urllib.request.Request(
            GITHUB_RELEASE_URL,
            headers={
                'User-Agent': 'EMR-Data-Mapper-Updater'
            }
        )
        
        # Get the latest release info
        with urllib.request.urlopen(req, timeout=10) as response:
            data = json.loads(response.read().decode('utf-8'))
            latest_version = data.get('tag_name', '').replace('v', '')
            
            if not latest_version:
                print(f"{Fore.YELLOW}Could not determine latest version.")
                return False
            
            print(f"{Fore.GREEN}Current version: {CURRENT_VERSION}")
            print(f"{Fore.GREEN}Latest version: {latest_version}")
            
            # Compare versions (simple string comparison for now, could be improved)
            if latest_version > CURRENT_VERSION:
                print(f"{Fore.YELLOW}Update available: v{latest_version}")
                return latest_version
            else:
                print(f"{Fore.GREEN}You have the latest version.")
                return False
    except Exception as e:
        print(f"{Fore.RED}Error checking for updates: {e}")
        return False

def download_update(version):
    """Download and install the latest update"""
    print(f"{Fore.CYAN}Downloading update v{version}...")
    
    try:
        # Create a backup of current files
        backup_dir = create_backup()
        if not backup_dir:
            print(f"{Fore.RED}Failed to create backup. Update aborted.")
            return False
        
        # Setup request with appropriate headers
        req = urllib.request.Request(
            GITHUB_ZIPBALL_URL,
            headers={
                'User-Agent': 'EMR-Data-Mapper-Updater'
            }
        )
        
        # Download the zipball
        with tempfile.NamedTemporaryFile(suffix='.zip', delete=False) as temp_file:
            temp_zip_path = temp_file.name
            
            with urllib.request.urlopen(req, timeout=30) as response:
                # Download with progress reporting
                file_size = int(response.headers.get('Content-Length', 0))
                downloaded = 0
                chunk_size = 8192
                
                print(f"{Fore.CYAN}Downloading {file_size/1024/1024:.2f} MB...")
                
                while True:
                    chunk = response.read(chunk_size)
                    if not chunk:
                        break
                    
                    temp_file.write(chunk)
                    downloaded += len(chunk)
                    
                    # Print progress
                    if file_size > 0:
                        percent = downloaded * 100 / file_size
                        sys.stdout.write(f"\r{Fore.CYAN}Progress: {percent:.1f}%")
                        sys.stdout.flush()
            
            print(f"\n{Fore.GREEN}Download complete.")
        
        # Extract and install the update
        if install_update(temp_zip_path, backup_dir):
            os.unlink(temp_zip_path)  # Remove the temporary zip file
            return True
        else:
            print(f"{Fore.RED}Update installation failed. Restoring from backup...")
            restore_from_backup(backup_dir)
            os.unlink(temp_zip_path)
            return False
            
    except Exception as e:
        print(f"{Fore.RED}Error downloading update: {e}")
        return False

def create_backup():
    """Create a backup of the current installation"""
    try:
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        backup_dir = os.path.join(SCRIPT_DIR, f"backup-{timestamp}")
        
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
        
        print(f"{Fore.CYAN}Creating backup in {backup_dir}")
        
        # Copy core directory
        if os.path.exists(CORE_DIR):
            core_backup = os.path.join(backup_dir, "core")
            shutil.copytree(CORE_DIR, core_backup)
        
        # Copy main script files
        for file in os.listdir(SCRIPT_DIR):
            if file.endswith('.py') or file.endswith('.bat'):
                source = os.path.join(SCRIPT_DIR, file)
                target = os.path.join(backup_dir, file)
                shutil.copy2(source, target)
        
        print(f"{Fore.GREEN}Backup created successfully.")
        return backup_dir
    except Exception as e:
        print(f"{Fore.RED}Error creating backup: {e}")
        return None

def find_root_dir_in_zip(zip_ref):
    """Find the root directory in the GitHub zipball"""
    for name in zip_ref.namelist():
        if name.endswith('/'):
            parts = name.split('/')
            if len(parts) == 2:  # First level directory
                return parts[0]
    return None

def install_update(zip_path, backup_dir):
    """Extract and install files from the update zip"""
    try:
        print(f"{Fore.CYAN}Installing update...")
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # GitHub zipballs have a root directory with the repo name and commit hash
            # We need to extract from that directory
            root_dir = find_root_dir_in_zip(zip_ref)
            
            if not root_dir:
                print(f"{Fore.RED}Could not find root directory in zip file.")
                return False
            
            # Extract to a temporary directory
            temp_extract_dir = tempfile.mkdtemp()
            zip_ref.extractall(temp_extract_dir)
            
            # Copy files from the extracted temp directory to the installation directory
            extracted_root = os.path.join(temp_extract_dir, root_dir)
            
            # Copy/update core directory
            extracted_core = os.path.join(extracted_root, "core")
            if os.path.exists(extracted_core):
                if os.path.exists(CORE_DIR):
                    shutil.rmtree(CORE_DIR)
                shutil.copytree(extracted_core, CORE_DIR)
            
            # Copy/update main script files
            for file in os.listdir(extracted_root):
                if file.endswith('.py') or file.endswith('.bat'):
                    source = os.path.join(extracted_root, file)
                    target = os.path.join(SCRIPT_DIR, file)
                    shutil.copy2(source, target)
            
            # Clean up temporary directory
            shutil.rmtree(temp_extract_dir)
            
            print(f"{Fore.GREEN}Update installed successfully.")
            return True
    except Exception as e:
        print(f"{Fore.RED}Error installing update: {e}")
        return False

def restore_from_backup(backup_dir):
    """Restore files from backup after failed update"""
    try:
        print(f"{Fore.CYAN}Restoring from backup...")
        
        # Restore core directory
        backup_core = os.path.join(backup_dir, "core")
        if os.path.exists(backup_core):
            if os.path.exists(CORE_DIR):
                shutil.rmtree(CORE_DIR)
            shutil.copytree(backup_core, CORE_DIR)
        
        # Restore main script files
        for file in os.listdir(backup_dir):
            if file.endswith('.py') or file.endswith('.bat'):
                source = os.path.join(backup_dir, file)
                target = os.path.join(SCRIPT_DIR, file)
                shutil.copy2(source, target)
        
        print(f"{Fore.GREEN}Restoration completed successfully.")
        return True
    except Exception as e:
        print(f"{Fore.RED}Error restoring from backup: {e}")
        return False

def run_updater():
    """Main function to run the updater"""
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
    print(f"{Fore.YELLOW}{Style.BRIGHT}{'EMR DATA MAPPER UPDATER':^80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}\n")
    
    # Check for updates
    update_version = check_for_updates()
    if update_version:
        do_update = input(f"{Fore.YELLOW}Would you like to update to version {update_version}? (y/n): ")
        if do_update.lower() == 'y':
            if download_update(update_version):
                print(f"{Fore.GREEN}Update complete! The program will now restart.")
                return True
            else:
                print(f"{Fore.RED}Update failed. Continuing with current version.")
    
    return False

if __name__ == "__main__":
    run_updater()