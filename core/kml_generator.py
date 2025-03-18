import os
import sys
import time
import logging
import pandas as pd
import xml.etree.ElementTree as ET
import colorsys
import re
import glob
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
            elif "Creating" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.CYAN}📌 "
            elif "Saved" in record.msg or "Created" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.GREEN}📄 "
            elif "Found" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.GREEN}🔍 "
            elif "Generating" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.YELLOW}⚙️ "
            elif "Added" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.CYAN}➕ "
            elif "Skipping" in record.msg:
                prefix = f"{Fore.MAGENTA}[{timestamp}] {Fore.YELLOW}⏭️ "
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

file_handler = logging.FileHandler("kml_generation.log")
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

logger = logging.getLogger("KMLGenerator")
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

# Set up paths - assuming kml_generator.py is in the /core/ directory
CORE_DIR = os.path.dirname(os.path.abspath(__file__))  # /core/
ROOT_DIR = os.path.dirname(CORE_DIR)  # Main directory

# Set up directories
PROCESSED_DIR = os.path.join(CORE_DIR, "processed")
KML_DIR = os.path.join(CORE_DIR, "kml")

# Create KML directory if it doesn't exist
if not os.path.exists(KML_DIR):
    os.makedirs(KML_DIR)
    logger.info(f"Created KML directory: {KML_DIR}")

# Configuration options for KML generation
POLYGON_TRANSPARENCY = '7f'  # '00' (fully transparent) to 'ff' (solid)
BORDER_WIDTH = '1'  # Border width in pixels
OT_COLOR = '7f0000ff'  # Blue with 50% transparency
PT_COLOR = '7fff0000'  # Red with 50% transparency

# Reference to the ZIP code KML file
ZIP_KML_FILE = os.path.join(CORE_DIR, "tl_2024_us_zcta520.kml")
if not os.path.exists(ZIP_KML_FILE):
    logger.error(f"ZIP Code KML file not found: {ZIP_KML_FILE}")
    logger.error("Please place the ZIP Code KML file in the core directory.")

def get_unique_color(index):
    """Generate a unique color based on index"""
    hue = (index * 0.618033988749895) % 1
    rgb = colorsys.hsv_to_rgb(hue, 0.7, 0.95)
    return f'{POLYGON_TRANSPARENCY}{int(rgb[2]*255):02x}{int(rgb[1]*255):02x}{int(rgb[0]*255):02x}'

def extract_zipcodes_from_area(area):
    """Extract valid ZIP codes from an area string using a more reliable approach"""
    if not isinstance(area, str) or area == 'Anywhere':
        return []
    
    # Split by comma and filter for valid 5-digit ZIP codes
    zipcodes = [zip.strip() for zip in area.split(',') if zip.strip().isdigit() and len(zip.strip()) == 5]
    return zipcodes

def create_therapist_coverage_kml(discipline, therapist_df, zip_kml_file, color):
    """Create KML file for all therapists of a specific discipline with therapist listings"""
    logger.info(f"Creating {discipline} therapist coverage KML")
    
    # Check if the input dataframe is valid
    if therapist_df is None or therapist_df.empty:
        logger.error("No therapist data provided")
        return None
    
    # Create mapping of ZIP codes to therapists
    zipcode_therapists = {}
    
    # Filter therapists by discipline (PT* or OT*)
    pattern = f"^{discipline}" if discipline in ["PT", "OT"] else discipline
    
    # Process each therapist
    for _, row in therapist_df.iterrows():
        therapist_discipline = str(row.get('Discipline', ''))
        
        # Skip therapists of other disciplines
        if not re.match(pattern, therapist_discipline):
            continue
            
        # Check if "Area" field exists and has zipcode information
        if isinstance(row.get('ZIP'), str) and row['ZIP'] != 'Anywhere':
            zipcodes = extract_zipcodes_from_area(row['ZIP'])
            
            for zipcode in zipcodes:
                if zipcode not in zipcode_therapists:
                    zipcode_therapists[zipcode] = []
                
                # Add therapist information
                therapist_info = f"{row['Discipline']} {row['Name']}"
                # Add language if available
                if 'Language' in row and pd.notna(row['Language']):
                    therapist_info += f" ({row['Language']})"
                
                zipcode_therapists[zipcode].append(therapist_info)
        elif 'ZIP' in row and row['ZIP'] == 'Anywhere':
            logger.info(f"Skipping therapist with 'Anywhere' coverage: {row['Name']}")
        else:
            logger.warning(f"Therapist missing ZIP data: {row['Name']}")
    
    logger.info(f"Found {len(zipcode_therapists)} ZIP codes for {discipline} therapists")
    
    # If no ZIP codes found, return
    if not zipcode_therapists:
        logger.warning(f"No ZIP codes found for {discipline} therapists")
        return None
    
    # Parse the ZIP code KML file
    try:
        tree = ET.parse(zip_kml_file)
        root = tree.getroot()
        logger.info("Successfully parsed ZIP code KML file")
    except Exception as e:
        logger.error(f"Error parsing ZIP code KML file: {e}")
        return None
    
    # Create new KML structure
    kml = ET.Element('kml', {'xmlns': 'http://www.opengis.net/kml/2.2'})
    doc = ET.SubElement(kml, 'Document')
    
    # Add name
    name_elem = ET.SubElement(doc, 'name')
    name_elem.text = f"All {discipline} Therapist Coverage"
    
    # Add styles
    style = ET.SubElement(doc, 'Style', {'id': 'coverageStyle'})
    
    # Line style (border)
    line_style = ET.SubElement(style, 'LineStyle')
    line_width = ET.SubElement(line_style, 'width')
    line_width.text = BORDER_WIDTH
    line_color = ET.SubElement(line_style, 'color')
    line_color.text = color
    
    # Polygon style (fill)
    poly_style = ET.SubElement(style, 'PolyStyle')
    poly_color = ET.SubElement(poly_style, 'color')
    poly_color.text = color
    fill = ET.SubElement(poly_style, 'fill')
    fill.text = '1'
    outline = ET.SubElement(poly_style, 'outline')
    outline.text = '1'
    
    # Process each ZIP code placemark from the source KML
    placemark_count = 0
    
    # Find placemarks with specified ZIP codes
    for original_placemark in root.findall('.//{http://www.opengis.net/kml/2.2}Placemark'):
        zipcode_elem = original_placemark.find('.//{http://www.opengis.net/kml/2.2}Data[@name="ZCTA5CE20"]/{http://www.opengis.net/kml/2.2}value')
        
        if zipcode_elem is not None and zipcode_elem.text in zipcode_therapists:
            zipcode = zipcode_elem.text
            placemark = ET.SubElement(doc, 'Placemark')
            
            # Add name
            name = ET.SubElement(placemark, 'name')
            name.text = f"ZIP {zipcode}"
            
            # Add description with therapist list
            therapist_list = sorted(zipcode_therapists[zipcode])
            description_text = f'<description><![CDATA[Area serviced by:<br><br>{("<br>").join(therapist_list)}]]></description>'
            placemark.append(ET.XML(description_text))
            
            # Add style reference
            style_url = ET.SubElement(placemark, 'styleUrl')
            style_url.text = '#coverageStyle'
            
            # Copy polygon data
            coords_elem = original_placemark.find('.//{http://www.opengis.net/kml/2.2}coordinates')
            if coords_elem is not None:
                polygon = ET.SubElement(placemark, 'Polygon')
                outer = ET.SubElement(polygon, 'outerBoundaryIs')
                ring = ET.SubElement(outer, 'LinearRing')
                tessellate = ET.SubElement(ring, 'tessellate')
                tessellate.text = '1'
                coords = ET.SubElement(ring, 'coordinates')
                coords.text = coords_elem.text.strip()
                
                placemark_count += 1
    
    logger.info(f"Added {placemark_count} placemarks to the {discipline} coverage KML")
    
    return kml if placemark_count > 0 else None

def create_patient_pins_kml(patient_dfs):
    """Create a KML file with patient pins colored by organization"""
    logger.info(f"Creating patient pins KML")
    
    # Create new KML structure
    kml = ET.Element('kml', {'xmlns': 'http://www.opengis.net/kml/2.2'})
    doc = ET.SubElement(kml, 'Document')
    
    # Add name with current date
    name_elem = ET.SubElement(doc, 'name')
    name_elem.text = f"Active_Cases_{datetime.now().strftime('%m/%d/%Y')}"
    
    # Define colors for different organizations
    org_styles = {
        "Rehab_on_Wheels": "icon-1502-558B2F",  # Green
        "Shining_Star": "icon-1502-FFEA00",  # Yellow
        "Four_Seasons": "icon-1502-C2185B",  # Pink
        "Girling_Health": "icon-1502-3949AB",  # Blue
        "Personal_Touch": "icon-1502-880E4F",  # Purple
        "Americare": "icon-1502-E65100"  # Orange
    }
    
    # Define colors
    colors = {
        "icon-1502-558B2F": "ff2f8b55",  # Green
        "icon-1502-FFEA00": "ff00eaff",  # Yellow
        "icon-1502-C2185B": "ff5b18c2",  # Pink
        "icon-1502-3949AB": "ffab4939",  # Blue
        "icon-1502-880E4F": "ff4f0e88",  # Purple
        "icon-1502-E65100": "ff0051e6"   # Orange
    }
    
    # Create all the styles
    for style_id, color_code in colors.items():
        style = ET.SubElement(doc, 'Style', {'id': style_id})
        icon_style = ET.SubElement(style, 'IconStyle')
        color = ET.SubElement(icon_style, 'color')
        color.text = color_code
        scale = ET.SubElement(icon_style, 'scale')
        scale.text = '1'
        icon = ET.SubElement(icon_style, 'Icon')
        href = ET.SubElement(icon, 'href')
        href.text = 'https://www.gstatic.com/mapspro/images/stock/503-wht-blank_maps.png'
        hotspot = ET.SubElement(icon_style, 'hotSpot')
        hotspot.set('x', '32')
        hotspot.set('xunits', 'pixels')
        hotspot.set('y', '64')
        hotspot.set('yunits', 'insetPixels')
    
    # Track which patient IDs we've already processed
    processed_patients = {}
    
    # First pass - collect all patients and their disciplines
    for df_name, df in patient_dfs:
        # Extract organization and discipline from filename
        match = re.match(r'(.+?)_(OT|PT)_Processed\.csv', df_name)
        if match:
            org = match.group(1)
            disc = match.group(2)
            
            # Skip if no data for this office/discipline
            if df.empty:
                continue
            
            # Process each patient to collect disciplines
            for _, row in df.iterrows():
                patient_id = str(row['Patient_Id'])
                
                if patient_id not in processed_patients:
                    processed_patients[patient_id] = {
                        'row': row,
                        'disciplines': set([disc]),
                        'org': org
                    }
                else:
                    # If patient already exists, add this discipline
                    processed_patients[patient_id]['disciplines'].add(disc)
    
    # Counter for pins
    pin_count = 0
    
    # Second pass - create placemarks with combined disciplines where needed
    for patient_id, patient_data in processed_patients.items():
        row = patient_data['row']
        org = patient_data['org']
        
        if pd.notna(row['Zip']) and pd.notna(row['Address']):
            # Create a placemark
            placemark = ET.SubElement(doc, 'Placemark')
            
            # Add name - Patient ID instead of patient name
            name_elem = ET.SubElement(placemark, 'name')
            name_elem.text = patient_id
            
            # Format the full address string for Google Maps to geocode
            full_address = f"{row['Address']}, {row['City']}, {row['State']}, {row['Zip']}"
            
            # Add address tag
            address_elem = ET.SubElement(placemark, 'address')
            address_elem.text = full_address
            
            # Add simple description - just the Patient ID
            description_elem = ET.SubElement(placemark, 'description')
            description_elem.text = patient_id
            
            # Add style reference based on organization
            org_key = org.replace(" ", "_").replace(".", "").replace(",", "")
            style_id = org_styles.get(org_key, org_styles["Americare"])  # Default if not found
            style_url = ET.SubElement(placemark, 'styleUrl')
            style_url.text = f'#{style_id}'
            
            # Add ExtendedData with the required fields in the correct order
            extended_data = ET.SubElement(placemark, 'ExtendedData')
            
            # Name field
            data_elem = ET.SubElement(extended_data, 'Data', {'name': 'Name'})
            value_elem = ET.SubElement(data_elem, 'value')
            value_elem.text = row['Name']
            
            # address field (lowercase as in your example)
            data_elem = ET.SubElement(extended_data, 'Data', {'name': 'address'})
            value_elem = ET.SubElement(data_elem, 'value')
            value_elem.text = full_address
            
            # Discipline field - Combine disciplines if multiple
            data_elem = ET.SubElement(extended_data, 'Data', {'name': 'Discipline'})
            value_elem = ET.SubElement(data_elem, 'value')
            
            disciplines = patient_data['disciplines']
            if 'PT' in disciplines and 'OT' in disciplines:
                value_elem.text = "PT and OT"
            else:
                value_elem.text = next(iter(disciplines))  # Just use the single discipline
            
            # Location field
            data_elem = ET.SubElement(extended_data, 'Data', {'name': 'Location'})
            value_elem = ET.SubElement(data_elem, 'value')
            value_elem.text = org.replace("_", " ")
            
            pin_count += 1
    
    logger.info(f"Added {pin_count} patient pins to the KML (after combining duplicate patients)")
    
    return kml if pin_count > 0 else None

def create_need_staff_pins_kml(need_staff_df):
    """Create a KML file with need staff pins"""
    logger.info(f"Creating need staff pins KML")
    
    # Check if the input dataframe is valid
    if need_staff_df is None or need_staff_df.empty:
        logger.error("No need staff data provided")
        return None
    
    # Create new KML structure
    kml = ET.Element('kml', {'xmlns': 'http://www.opengis.net/kml/2.2'})
    doc = ET.SubElement(kml, 'Document')
    
    # Add name with current date
    name_elem = ET.SubElement(doc, 'name')
    name_elem.text = f"Need_Staff_{datetime.now().strftime('%m/%d/%Y')}"
    
    # Create style for the pins - Yellow
    style = ET.SubElement(doc, 'Style', {'id': 'icon-1502-FFEA00'})
    icon_style = ET.SubElement(style, 'IconStyle')
    color = ET.SubElement(icon_style, 'color')
    color.text = "ff00eaff"  # Yellow (in ABGR format)
    scale = ET.SubElement(icon_style, 'scale')
    scale.text = '1'
    icon = ET.SubElement(icon_style, 'Icon')
    href = ET.SubElement(icon, 'href')
    href.text = 'https://www.gstatic.com/mapspro/images/stock/503-wht-blank_maps.png'
    hotspot = ET.SubElement(icon_style, 'hotSpot')
    hotspot.set('x', '32')
    hotspot.set('xunits', 'pixels')
    hotspot.set('y', '64')
    hotspot.set('yunits', 'insetPixels')
    
    # Counter for pins
    pin_count = 0
    
    # Process each need staff entry
    for _, row in need_staff_df.iterrows():
        if pd.notna(row['zip']) and pd.notna(row['Address']):
            zipcode = str(row['zip']).strip()
            
            # Create a placemark
            placemark = ET.SubElement(doc, 'Placemark')
            
            # Add name with the ID
            name_elem = ET.SubElement(placemark, 'name')
            name_elem.text = str(row['ID'])
            
            # Format the full address string for Google Maps to geocode
            full_address = f"{row['Address']}, {zipcode}"
            
            # Add address tag for automatic geocoding
            address_elem = ET.SubElement(placemark, 'address')
            address_elem.text = full_address
            
            # Add simple description - just the ID
            description_elem = ET.SubElement(placemark, 'description')
            description_elem.text = str(row['ID'])
            
            # Add style reference
            style_url = ET.SubElement(placemark, 'styleUrl')
            style_url.text = '#icon-1502-FFEA00'
            
            # Add ExtendedData with the required fields
            extended_data = ET.SubElement(placemark, 'ExtendedData')
            
            # ID field
            data_elem = ET.SubElement(extended_data, 'Data', {'name': 'ID'})
            value_elem = ET.SubElement(data_elem, 'value')
            value_elem.text = str(row["ID"])
            
            # Ref Date field
            data_elem = ET.SubElement(extended_data, 'Data', {'name': 'Ref Date'})
            value_elem = ET.SubElement(data_elem, 'value')
            value_elem.text = str(row["Ref Date"])
            
            # Address field (address)
            data_elem = ET.SubElement(extended_data, 'Data', {'name': 'Address'})
            value_elem = ET.SubElement(data_elem, 'value')
            value_elem.text = str(row["Address"])
            
            # zip field
            data_elem = ET.SubElement(extended_data, 'Data', {'name': 'zip'})
            value_elem = ET.SubElement(data_elem, 'value')
            value_elem.text = zipcode
            
            # Discpline field (with the original spelling)
            data_elem = ET.SubElement(extended_data, 'Data', {'name': 'Discpline'})
            value_elem = ET.SubElement(data_elem, 'value')
            value_elem.text = str(row["Discpline"])
            
            pin_count += 1
    
    logger.info(f"Added {pin_count} need staff pins to the KML")
    
    return kml if pin_count > 0 else None

def main():
    """Main function to generate KML files"""
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
    print(f"{Fore.YELLOW}{Style.BRIGHT}{'KML GENERATION':^80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}\n")
    
    logger.info("Starting KML generation")
    
    # Check if the ZIP Code KML file exists
    if not os.path.exists(ZIP_KML_FILE):
        logger.error(f"ZIP Code KML file not found: {ZIP_KML_FILE}")
        logger.error("Please place the ZIP Code KML file in the core directory.")
        print(f"{Fore.RED}Error: ZIP Code KML file not found: {ZIP_KML_FILE}")
        print(f"{Fore.YELLOW}Please place the ZIP Code KML file in the core directory.")
        return
    
    # Check if the processed directory exists
    if not os.path.exists(PROCESSED_DIR):
        logger.error(f"Processed directory not found: {PROCESSED_DIR}")
        logger.error("Please run the data_processor.py script first.")
        print(f"{Fore.RED}Error: Processed directory not found: {PROCESSED_DIR}")
        print(f"{Fore.YELLOW}Please run the data_processor.py script first.")
        return
    
    # Track timing and success count
    total_start_time = time.time()
    success_count = 0
    
    # 1. Load active therapists data
    therapist_file = os.path.join(PROCESSED_DIR, "ActiveTherapists.csv")
    if os.path.exists(therapist_file):
        try:
            therapist_df = pd.read_csv(therapist_file)
            logger.info(f"Loaded {len(therapist_df)} active therapists from {therapist_file}")
            
            # Check DataFrame columns for correct names
            if 'Discipline' not in therapist_df.columns or 'ZIP' not in therapist_df.columns:
                logger.info(f"Looking for alternative column names in therapist data, found: {therapist_df.columns.tolist()}")
                # Try to identify and rename columns if they have different names
                if 'Role' in therapist_df.columns:
                    therapist_df = therapist_df.rename(columns={'Role': 'Discipline'})
                    logger.info("Renamed 'Role' column to 'Discipline'")
                
                if 'Area' in therapist_df.columns:
                    therapist_df = therapist_df.rename(columns={'Area': 'ZIP'})
                    logger.info("Renamed 'Area' column to 'ZIP'")
            
            # 2. Generate OT coverage KML
            ot_kml = create_therapist_coverage_kml("OT", therapist_df, ZIP_KML_FILE, OT_COLOR)
            
            if ot_kml is not None:
                ot_output_file = os.path.join(KML_DIR, "OT_Therapist_Coverage.kml")
                ET.ElementTree(ot_kml).write(ot_output_file, encoding='UTF-8', xml_declaration=True)
                logger.info(f"Successfully created OT coverage KML: {ot_output_file}")
                success_count += 1
                print(f"{Fore.GREEN}✓ {Style.BRIGHT}Generated OT therapist coverage KML")
            else:
                logger.warning("Failed to create OT coverage KML")
                print(f"{Fore.RED}✗ {Style.BRIGHT}Failed to create OT therapist coverage KML")
            
            # 3. Generate PT coverage KML
            pt_kml = create_therapist_coverage_kml("PT", therapist_df, ZIP_KML_FILE, PT_COLOR)
            
            if pt_kml is not None:
                pt_output_file = os.path.join(KML_DIR, "PT_Therapist_Coverage.kml")
                ET.ElementTree(pt_kml).write(pt_output_file, encoding='UTF-8', xml_declaration=True)
                logger.info(f"Successfully created PT coverage KML: {pt_output_file}")
                success_count += 1
                print(f"{Fore.GREEN}✓ {Style.BRIGHT}Generated PT therapist coverage KML")
            else:
                logger.warning("Failed to create PT coverage KML")
                print(f"{Fore.RED}✗ {Style.BRIGHT}Failed to create PT therapist coverage KML")
                
        except Exception as e:
            logger.error(f"Error processing active therapists: {e}")
            print(f"{Fore.RED}Error processing active therapists: {e}")
        else:
            logger.error(f"Active therapists file not found: {therapist_file}")
            print(f"{Fore.RED}Active therapists file not found: {therapist_file}")
    
    # 4. Load patient data from processed files
    patient_files = glob.glob(os.path.join(PROCESSED_DIR, "*_Processed.csv"))
    
    if patient_files:
        try:
            # Load all patient dataframes
            patient_dfs = []
            
            for file_path in patient_files:
                file_name = os.path.basename(file_path)
                try:
                    df = pd.read_csv(file_path)
                    patient_dfs.append((file_name, df))
                    logger.info(f"Loaded {len(df)} patients from {file_name}")
                except Exception as e:
                    logger.error(f"Error loading patient data from {file_name}: {e}")
            
            # Generate patient pins KML
            patient_kml = create_patient_pins_kml(patient_dfs)
            
            if patient_kml is not None:
                patient_output_file = os.path.join(KML_DIR, "Active_Cases.kml")
                ET.ElementTree(patient_kml).write(patient_output_file, encoding='UTF-8', xml_declaration=True)
                logger.info(f"Successfully created patient pins KML: {patient_output_file}")
                success_count += 1
                print(f"{Fore.GREEN}✓ {Style.BRIGHT}Generated patient pins KML")
            else:
                logger.warning("Failed to create patient pins KML")
                print(f"{Fore.RED}✗ {Style.BRIGHT}Failed to create patient pins KML")
                
        except Exception as e:
            logger.error(f"Error processing patient data: {e}")
            print(f"{Fore.RED}Error processing patient data: {e}")
    else:
        logger.error(f"No processed patient files found in {PROCESSED_DIR}")
        print(f"{Fore.RED}No processed patient files found")
    
    # 5. Load need staff data
    need_staff_file = os.path.join(PROCESSED_DIR, "NeedStaff.csv")
    if os.path.exists(need_staff_file):
        try:
            need_staff_df = pd.read_csv(need_staff_file)
            logger.info(f"Loaded {len(need_staff_df)} need staff entries from {need_staff_file}")
            
            # Generate need staff pins KML
            need_staff_kml = create_need_staff_pins_kml(need_staff_df)
            
            if need_staff_kml is not None:
                need_staff_output_file = os.path.join(KML_DIR, "Need_Staff.kml")
                ET.ElementTree(need_staff_kml).write(need_staff_output_file, encoding='UTF-8', xml_declaration=True)
                logger.info(f"Successfully created need staff pins KML: {need_staff_output_file}")
                success_count += 1
                print(f"{Fore.GREEN}✓ {Style.BRIGHT}Generated need staff pins KML")
            else:
                logger.warning("Failed to create need staff pins KML")
                print(f"{Fore.RED}✗ {Style.BRIGHT}Failed to create need staff pins KML")
                
        except Exception as e:
            logger.error(f"Error processing need staff data: {e}")
            print(f"{Fore.RED}Error processing need staff data: {e}")
    else:
        logger.error(f"Need staff file not found: {need_staff_file}")
        print(f"{Fore.RED}Need staff file not found: {need_staff_file}")
    
    # Print summary
    total_duration = time.time() - total_start_time
    minutes = int(total_duration // 60)
    seconds = int(total_duration % 60)
    
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'='*80}")
    print(f"{Fore.YELLOW}{Style.BRIGHT}{f'KML GENERATION COMPLETED: {success_count}/4 FILES':^80}")
    print(f"{Fore.GREEN}{Style.BRIGHT}{f'Total time: {minutes} minutes {seconds} seconds':^80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'='*80}{Style.RESET_ALL}\n")
    
    logger.info(f"KML generation completed: {success_count}/4 files generated")
    
    # Provide instruction for using the KML files
    if success_count > 0:
        print(f"{Fore.CYAN}KML files are located in: {KML_DIR}")
        print(f"{Fore.YELLOW}You can import these files into Google Earth or Google Maps to visualize the coverage areas.")
        print(f"{Fore.YELLOW}Instructions for Google Maps:")
        print(f"{Fore.WHITE}1. Go to https://www.google.com/maps/")
        print(f"{Fore.WHITE}2. Click on the menu (hamburger icon) in the top left")
        print(f"{Fore.WHITE}3. Select 'Your Places' > 'Maps' > 'Create Map'")
        print(f"{Fore.WHITE}4. Click 'Import' and select the KML file you want to view")

if __name__ == "__main__":
    main()