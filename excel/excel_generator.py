import os
import sys
import logging
import pandas as pd
import xlwt
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional, Union, Tuple
import csv
import time
import re

# Add project root to path for imports
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

try:
    from ..utils.file_manager import FileManager
except ImportError:
    print("Warning: Could not import FileManager, using fallback implementation")
    class FileManager:
        """Fallback FileManager implementation"""
        def __init__(self):
            # Get project root directory (parent of excel directory)
            excel_dir = Path(__file__).parent  # excel directory
            self.project_root = excel_dir.parent    # main project directory
        def get_csv_directory(self):
            """Get CSV directory with fallback"""
            today = datetime.now()
            date_folder = today.strftime("%d%b").lower()  # Format: 22jul
            return self.project_root / "EHC_Data" / date_folder
        def get_excel_directory(self):
            """Get Excel directory with fallback"""
            today = datetime.now()
            date_folder = today.strftime("%d%b").lower()
            return self.project_root / "EHC_Data_Merge" / date_folder
        def count_csv_files(self, directory=None):
            """Count CSV files with fallback"""
            if directory is None:
                directory = self.get_csv_directory()
            try:
                return len(list(Path(directory).glob("*.csv")))
            except:
                return 0

# --- Define constants matching JavaScript ---
REQUIRED_COLUMNS = [
    'Hostname',
    'IP Address',
    'MAC Address',
    'WLAN (SSID)',
    'AP MAC',
    'Data Rate (up)',
    'Data Rate (down)'
]

# Column name variations to handle different CSV formats (Python version of JS COLUMN_MAPPINGS)
COLUMN_MAPPINGS = {
    'Hostname': ['hostname', 'host name', 'host_name', 'device name', 'device_name'],
    'IP Address': ['ip address', 'ip_address', 'ipaddress', 'ip addr', 'ip_addr'],
    'MAC Address': ['mac address', 'mac_address', 'macaddress', 'mac addr', 'mac_addr'],
    'WLAN (SSID)': ['wlan (ssid)', 'wlan ssid', 'wlan_ssid', 'ssid', 'network name', 'network_name'],
    'AP MAC': ['ap mac', 'ap_mac', 'apmac', 'access point mac', 'access_point_mac'],
    'Data Rate (up)': ['data rate (up)', 'data_rate_up', 'upload rate', 'upload_rate', 'up rate', 'up_rate'],
    'Data Rate (down)': ['data rate (down)', 'data_rate_down', 'download rate', 'download_rate', 'down rate', 'down_rate']
}

def normalize_headers(headers: List[str]) -> Dict[str, str]:
    """
    Normalize CSV headers to match standard column names.
    Similar to JavaScript normalizeHeaders function.
    Returns a dictionary mapping original header names to standard names.
    """
    header_mapping = {}
    for header in headers:
        normalized_header = header.lower().strip()
        # Find matching column
        for standard_column, variations in COLUMN_MAPPINGS.items():
             # Check exact match (case insensitive) or variation match
            if normalized_header == standard_column.lower() or normalized_header in variations:
                header_mapping[header] = standard_column
                break # Found a match, move to next header
    return header_mapping


class ExcelGenerator:
    def __init__(self, file_manager: FileManager = None):
        """Initialize Excel Generator with file management"""
        self.file_manager = file_manager or FileManager()
        self.logger = logging.getLogger(__name__)
        # Setup logging if not already configured
        if not self.logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
            self.logger.setLevel(logging.INFO)

        # Target headers for Excel file (in exact order for VBS compatibility)
        # These are the final headers in the Excel file
        self.excel_headers = [
            'Hostname',
            'IP_Address',
            'MAC_Address',
            'Package',
            'AP_MAC',
            'Upload',
            'Download'
        ]

        # Mapping from standard CSV column names (after normalization) to Excel headers
        # This replaces the old csv_to_excel_mapping
        self.standard_to_excel_mapping = {
            'Hostname': 'Hostname',
            'IP Address': 'IP_Address',
            'MAC Address': 'MAC_Address',
            'WLAN (SSID)': 'Package',
            'AP MAC': 'AP_MAC',
            'Data Rate (up)': 'Upload',
            'Data Rate (down)': 'Download'
        }

    def parse_csv_file(self, csv_file: Path) -> List[Dict[str, str]]:
        """
        Parse a single CSV file, normalize headers, and map to standard Excel format.
        Similar logic to JavaScript parseCSVFile chunk processing.
        """
        try:
            self.logger.info(f"Parsing CSV file: {csv_file}")
            processed_rows = []

            # Use pandas for easier handling of various CSV formats (like JS PapaParse)
            # Specify encoding and handle potential issues
            df = pd.read_csv(csv_file, encoding='utf-8-sig', on_bad_lines='skip')
            if df.empty:
                 self.logger.warning(f"CSV file {csv_file.name} is empty or could not be read.")
                 return []

            # Normalize headers from the DataFrame
            raw_headers = df.columns.tolist()
            header_mapping = normalize_headers(raw_headers) # Use the new function

            if not header_mapping:
                self.logger.warning(f"No recognized columns found in {csv_file.name}. Expected columns related to: {list(COLUMN_MAPPINGS.keys())}")
                return []

            self.logger.debug(f"Header mapping for {csv_file.name}: {header_mapping}")

            # Rename DataFrame columns to standard names based on mapping
            rename_dict = {orig: std for orig, std in header_mapping.items()}
            df.rename(columns=rename_dict, inplace=True)

            # Keep only the standard columns we are interested in
            standard_columns = list(self.standard_to_excel_mapping.keys())
            existing_standard_columns = [col for col in standard_columns if col in df.columns]
            df_filtered = df[existing_standard_columns].copy() # Work with a subset

            # Convert DataFrame rows to dictionaries
            # Iterate through rows and map to Excel format
            for _, row in df_filtered.iterrows():
                 processed_row = {excel_header: "" for excel_header in self.excel_headers} # Initialize with empty strings

                 # Map values from standard columns to Excel headers
                 for standard_col, excel_col in self.standard_to_excel_mapping.items():
                      if standard_col in row and pd.notna(row[standard_col]): # Check if column exists and value is not NaN
                           # Convert value to string and strip
                           value = str(row[standard_col]).strip()
                           processed_row[excel_col] = value
                      # If column doesn't exist or value is NaN, it remains an empty string due to initialization

                 # Check if row has any non-empty values (basic check, can be omitted if all rows are needed)
                 has_data = any(value != "" for value in processed_row.values())
                 if has_data: # Only append rows with data, matching JS logic slightly
                      processed_rows.append(processed_row)

            self.logger.info(f"Successfully parsed {len(processed_rows)} rows from {csv_file.name}")
            return processed_rows

        except Exception as e:
            self.logger.error(f"Failed to parse CSV file {csv_file}: {e}", exc_info=True) # Log full traceback
            return []

    # ... (rest of the class methods remain largely the same,
    # but ensure they use the updated parsing logic) ...

    def discover_csv_files(self) -> List[Path]:
        """Discover all CSV files in the CSV directory"""
        try:
            csv_dir = self.file_manager.get_csv_directory()
            csv_files = []
            if csv_dir.exists():
                csv_files = list(csv_dir.glob("*.csv"))
                self.logger.info(f"Found {len(csv_files)} CSV files in {csv_dir}")
            else:
                self.logger.warning(f"CSV directory does not exist: {csv_dir}")
            return csv_files
        except Exception as e:
            self.logger.error(f"Failed to discover CSV files: {e}")
            return []

    def combine_csv_data(self, all_data: List[List[Dict[str, str]]]) -> List[Dict[str, str]]:
        """Combine data from all CSV files (simple concatenation)"""
        try:
            combined_data = []
            for file_data in all_data:
                combined_data.extend(file_data)
            self.logger.info(f"Combined data from all files: {len(combined_data)} total rows")
            return combined_data
        except Exception as e:
            self.logger.error(f"Failed to combine CSV data: {e}")
            return []

    def generate_filename(self, date_folder: str = None) -> str:
        """Generate Excel filename with date from folder or current date"""
        try:
            if date_folder:
                # Extract date from folder name (e.g., "21jul" -> "21072025")
                match = re.match(r'(\d{1,2})([a-z]{3})', date_folder.lower())
                if match:
                    day = match.group(1).zfill(2)  # Ensure 2 digits
                    month_abbr = match.group(2)
                    # Convert month abbreviation to number
                    month_map = {
                        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
                        'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
                        'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
                    }
                    month = month_map.get(month_abbr, '07')  # Default to July
                    year = datetime.now().strftime("%Y")  # Current year
                    filename = f"EHC_Upload_Mac_{day}{month}{year}.xls" # Keep .xls for VBS
                    self.logger.info(f"Generated filename from folder '{date_folder}': {filename}")
                    return filename
            # Fallback to current date
            today = datetime.now().strftime("%d%m%Y")  # 24072025 format
            filename = f"EHC_Upload_Mac_{today}.xls" # Keep .xls for VBS
            return filename
        except Exception as e:
            self.logger.error(f"Failed to generate filename: {e}")
            # Return a default name that still indicates .xls format
            return "EHC_Upload_Mac_default.xls"

    def create_excel_file(self, data: List[Dict[str, str]]) -> Optional[Path]:
        """Create Excel file in old format (.xls) for VBS compatibility"""
        try:
            if not data:
                self.logger.error("No data to export to Excel")
                return None

            # Generate filename with today's date or derived from folder
            # The date_folder needs to be determined. If using discover_csv_files,
            # you might derive it from the csv_dir path.
            # For simplicity here, using current date or default logic in generate_filename
            excel_filename = self.generate_filename() # Modify if you have date_folder logic

            excel_dir = self.file_manager.get_excel_directory()
            # Ensure the output directory exists
            excel_dir.mkdir(parents=True, exist_ok=True)
            excel_path = excel_dir / excel_filename
            self.logger.info(f"Creating Excel file: {excel_path}")

            # Create workbook using xlwt (OLD EXCEL FORMAT .xls)
            workbook = xlwt.Workbook()
            worksheet = workbook.add_sheet('EHC_Data') # Same as the root folder name

            # Define styles for better readability
            header_style = xlwt.XFStyle()
            header_font = xlwt.Font()
            header_font.bold = True
            header_style.font = header_font

            # Write headers in exact order (matching self.excel_headers)
            for col, header in enumerate(self.excel_headers): # Use self.excel_headers
                worksheet.write(0, col, header, header_style)

            # Write data rows (NO FILTERING OR DUPLICATE REMOVAL)
            for row_idx, record in enumerate(data, start=1):
                for col_idx, header in enumerate(self.excel_headers): # Use self.excel_headers
                    value = record.get(header, '')
                    # Convert to string to avoid type issues, preserve empty strings
                    worksheet.write(row_idx, col_idx, str(value) if value else '')
                # Log progress for large datasets
                if row_idx % 1000 == 0:
                    self.logger.info(f"Written {row_idx} rows to Excel...")

            # Set column widths for better readability (matching JS)
            # xlwt column width unit is 1/256th of the width of '0' character
            column_widths_chars = [
                25,  # Hostname
                16,  # IP_Address
                20,  # MAC_Address
                30,  # Package
                20,  # AP_MAC
                18,  # Upload
                18   # Download
            ]
            # Convert character widths to xlwt units (approx 256 units per character)
            column_widths_xlwt = [width * 256 for width in column_widths_chars]

            for col, width in enumerate(column_widths_xlwt):
                worksheet.col(col).width = width

            # Save the workbook
            workbook.save(str(excel_path))

            # Validate file creation
            if excel_path.exists() and excel_path.stat().st_size > 0:
                file_size = excel_path.stat().st_size / 1024  # Size in KB
                self.logger.info(f"Excel file created successfully: {excel_path}")
                self.logger.info(f"File size: {file_size:.1f} KB, Rows: {len(data)}")
                return excel_path
            else:
                self.logger.error("Excel file creation failed - file is empty or doesn't exist")
                return None
        except Exception as e:
            self.logger.error(f"Failed to create Excel file: {e}", exc_info=True) # Log full traceback
            return None

    # ... (other methods like estimate_excel_size, validate_csv_files remain) ...

    def execute_excel_generation(self) -> Dict[str, Any]:
        """Execute complete Excel generation process"""
        try:
            start_time = time.time()
            self.logger.info("=== STARTING EXCEL GENERATION PROCESS ===")

            # Step 1: Discover CSV files
            csv_files = self.discover_csv_files()
            if not csv_files:
                return {"success": False, "error": "No CSV files found"}

            # Step 2: Validate CSV files (optional, can integrate checks into parsing)
            # valid_files, validation_errors = self.validate_csv_files(csv_files)
            # For simplicity, proceed with discovered files, relying on try/except in parse_csv_file
            valid_files = csv_files
            validation_errors = [] # Initialize, populate if using validate_csv_files

            if not valid_files:
                return {"success": False, "error": "No valid CSV files found"} # , "validation_errors": validation_errors

            self.logger.info(f"Processing {len(valid_files)} CSV files")

            # Step 3: Parse all CSV files using the improved method
            all_data = []
            for csv_file in valid_files:
                file_data = self.parse_csv_file(csv_file) # Use the fixed parse_csv_file
                if file_data: # Only append if data was successfully parsed
                    all_data.append(file_data)
                # Optional: Add small delay or yield if processing is UI-blocking in any context

            if not any(all_data): # Check if all sublists are empty
                return {"success": False, "error": "No data parsed from CSV files"}

            # Step 4: Combine all data (NO FILTERING OR DUPLICATE REMOVAL)
            combined_data = self.combine_csv_data(all_data)
            if not combined_data:
                return {"success": False, "error": "No combined data available"}

            # Step 5: Estimate file size
            size_estimate = self.estimate_excel_size(len(combined_data))
            self.logger.info(f"Estimated Excel file size: {size_estimate['estimated_mb']} MB")
            if size_estimate.get("warning"):
                self.logger.warning(size_estimate["warning"])

            # Step 6: Create Excel file
            excel_path = self.create_excel_file(combined_data)
            if not excel_path:
                return {"success": False, "error": "Failed to create Excel file"}

            # Step 7: Return success result
            processing_time = time.time() - start_time
            result = {
                "success": True,
                "excel_file": str(excel_path),
                "total_rows": len(combined_data),
                "csv_files_processed": len(valid_files),
                "processing_time": round(processing_time, 2),
                "file_size_estimate": size_estimate
            }
            # if validation_errors: # Uncomment if using validate_csv_files
            #     result["validation_errors"] = validation_errors

            self.logger.info(f"=== EXCEL GENERATION COMPLETED SUCCESSFULLY ===")
            self.logger.info(f"Total rows processed: {len(combined_data)}")
            self.logger.info(f"Processing time: {processing_time:.2f} seconds")
            self.logger.info(f"Output file: {excel_path}")
            return result
        except Exception as e:
            self.logger.error(f"Excel generation process failed: {e}", exc_info=True) # Log full traceback
            return {"success": False, "error": f"Process failed: {str(e)}"}

    # ... (rest of the class methods like get_processing_stats, generate_from_csv_directory, estimate_excel_size, validate_csv_files) ...
    # Ensure generate_from_csv_directory also uses the new normalize_headers and mapping logic if needed,
    # or refactor common parts.

    def estimate_excel_size(self, row_count: int) -> Dict[str, Any]:
        """Estimate Excel file size and provide warnings"""
        try:
            # Rough estimation: ~100 bytes per row for 7 columns
            estimated_bytes = row_count * 100
            estimated_mb = estimated_bytes / (1024 * 1024)
            result = {
                "estimated_mb": round(estimated_mb, 1),
                "warning": None
            }
            if estimated_mb > 100:
                result["warning"] = "Large file size expected (>100MB). Processing may take time."
            elif estimated_mb > 500:
                result["warning"] = "Very large file size expected (>500MB). Ensure sufficient disk space."
            return result
        except Exception as e:
            self.logger.error(f"Failed to estimate Excel size: {e}")
            return {"estimated_mb": 0, "warning": "Size estimation failed"}

    def validate_csv_files(self, csv_files: List[Path]) -> Tuple[List[Path], List[str]]:
        """Validate CSV files and return valid files and errors"""
        try:
            valid_files = []
            errors = []
            max_file_size = 500 * 1024 * 1024  # 500MB per file
            for csv_file in csv_files:
                # Check if file exists
                if not csv_file.exists():
                    errors.append(f"{csv_file.name} does not exist")
                    continue
                # Check file size
                if csv_file.stat().st_size > max_file_size:
                    errors.append(f"{csv_file.name} is too large (>500MB)")
                    continue
                # Check if file is readable (basic check)
                try:
                    # Attempt to read the first line using the same method as parse_csv_file
                     # Use pandas head or read with nrows=1 might be better
                     df_test = pd.read_csv(csv_file, nrows=1, encoding='utf-8-sig', on_bad_lines='skip')
                     # If it gets here without error, it's likely readable
                     valid_files.append(csv_file)
                except Exception as e:
                    errors.append(f"{csv_file.name} is not readable: {str(e)}")
            return valid_files, errors
        except Exception as e:
            self.logger.error(f"Failed to validate CSV files: {e}")
            return [], [f"Validation failed: {str(e)}"]


def main():
    """Test the Excel generator independently"""
    try:
        print("🧪 Testing Excel Generator...")
        # Initialize Excel generator
        generator = ExcelGenerator()

        # Option 1: Use the automated process based on FileManager paths
        result = generator.execute_excel_generation() # This should now work better

        if result.get("success"):
            print(f"✅ Excel generation successful!")
            print(f"   📝 Excel file: {result.get('excel_file')}")
            print(f"   📊 Records processed: {result.get('total_rows')}")
            print(f"   📄 CSV files processed: {result.get('csv_files_processed')}")
            print(f"   ⏱️  Processing time: {result.get('processing_time')} seconds")
        else:
            print(f"❌ Excel generation failed: {result.get('error')}")
            # if result.get("validation_errors"):
            #     print(f"   Validation errors: {result.get('validation_errors')}")

        # Option 2: Use generate_from_csv_directory if you have a specific path
        # today_folder = datetime.now().strftime("%d%b").lower() # e.g., "21jul"
        # specific_csv_dir = "/path/to/your/csv/files" # Replace with actual path
        # result2 = generator.generate_from_csv_directory(specific_csv_dir, today_folder)
        # if result2.get("success"):
        #     print(f"✅ Excel generation (specific dir) successful!")
        #     print(f"   📝 Excel file: {result2.get('excel_file')}")
        #     print(f"   📊 Records processed: {result2.get('records_processed')}")
        #     print(f"   📄 CSV files processed: {result2.get('csv_files_processed')}")
        # else:
        #     print(f"❌ Excel generation (specific dir) failed: {result2.get('error')}")

    except Exception as e:
        print(f"❌ Test failed: {e}")
        import traceback
        traceback.print_exc() # Print full stack trace for debugging

if __name__ == "__main__":
    # Ensure basic logging configuration if running standalone
    if not logging.getLogger().hasHandlers():
         logging.basicConfig(level=logging.INFO)
    main()
