#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VDXConvert - Batch converter for Visio files to VDX format

This script processes Visio files (VSD, VSDX, VSDM, VDW) from the input folder,
converts them to VDX format, and saves the results in the output folder.
Original files are moved to the archive folder upon successful conversion.

Author: Sam Lyndon
Version: 1.0.0
License: MIT
"""

import os
import sys
import time
import shutil
import logging
import argparse
import platform
import subprocess
import csv
from datetime import datetime
from xml.etree import ElementTree as ET
from pathlib import Path
import traceback

# Try to import optional dependencies
try:
    import colorama
    from colorama import Fore, Style
    colorama.init(autoreset=True)
    COLOR_SUPPORT = True
except ImportError:
    COLOR_SUPPORT = False

try:
    from tqdm import tqdm
    PROGRESS_BAR_SUPPORT = True
except ImportError:
    PROGRESS_BAR_SUPPORT = False

# Try to import vsdx module
try:
    import vsdx
    VSDX_SUPPORT = True
except ImportError:
    VSDX_SUPPORT = False

# Constants
APP_NAME = "VDXConvert"
VERSION = "1.0.0"
SUPPORTED_EXTENSIONS = ['.vsd', '.vsdx', '.vsdm', '.vdw']
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(SCRIPT_DIR, "input")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output")
ARCHIVE_DIR = os.path.join(SCRIPT_DIR, "archive")
LOGS_DIR = os.path.join(SCRIPT_DIR, "logs")

# Configure logger
def setup_logging(verbose=False):
    """Configure the logging system"""
    os.makedirs(LOGS_DIR, exist_ok=True)
    
    log_level = logging.DEBUG if verbose else logging.INFO
    log_file = os.path.join(LOGS_DIR, f"{APP_NAME}.log")
    
    # Configure root logger
    logger = logging.getLogger()
    logger.setLevel(log_level)
    
    # Clear existing handlers
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # File handler
    file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
    file_handler.setLevel(log_level)
    file_format = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(file_format)
    logger.addHandler(file_handler)
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    
    if COLOR_SUPPORT:
        # Custom formatter with colors
        class ColoredFormatter(logging.Formatter):
            formats = {
                logging.DEBUG: Fore.CYAN + '%(message)s' + Style.RESET_ALL,
                logging.INFO: '%(message)s',
                logging.WARNING: Fore.YELLOW + '%(message)s' + Style.RESET_ALL,
                logging.ERROR: Fore.RED + '%(message)s' + Style.RESET_ALL,
                logging.CRITICAL: Fore.RED + Style.BRIGHT + '%(message)s' + Style.RESET_ALL
            }
            
            def format(self, record):
                log_fmt = self.formats.get(record.levelno)
                formatter = logging.Formatter(log_fmt)
                return formatter.format(record)
        
        console_handler.setFormatter(ColoredFormatter())
    else:
        console_format = logging.Formatter('%(message)s')
        console_handler.setFormatter(console_format)
    
    logger.addHandler(console_handler)
    
    return logger

# Check dependencies
def check_dependencies():
    """Check if all required dependencies are installed"""
    missing_deps = []
    
    # Check for vsdx module
    if not VSDX_SUPPORT:
        missing_deps.append("vsdx")
        logging.warning("vsdx module not found. .vsdx and .vsdm processing will be limited.")
    
    # Check for unoconv/LibreOffice
    has_libreoffice = False
    try:
        if platform.system() == 'Windows':
            # Check if LibreOffice is installed on Windows
            import winreg
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\LibreOffice") as key:
                    logging.debug("LibreOffice found in registry")
                    has_libreoffice = True
            except WindowsError:
                missing_deps.append("LibreOffice")
                logging.warning("LibreOffice not found. .vsd and .vdw processing will be limited.")
        else:
            # Check for LibreOffice on macOS/Linux
            try:
                result = subprocess.run(
                    ["which", "soffice"], 
                    stdout=subprocess.PIPE, 
                    stderr=subprocess.PIPE, 
                    text=True, 
                    check=True
                )
                if result.stdout:
                    logging.debug(f"LibreOffice found at: {result.stdout.strip()}")
                    has_libreoffice = True
            except (subprocess.SubprocessError, FileNotFoundError):
                missing_deps.append("LibreOffice")
                logging.warning("LibreOffice not found. .vsd and .vdw processing will be limited.")
    except Exception as e:
        logging.error(f"Error checking for LibreOffice: {str(e)}")
    
    # Check for unoconv
    has_unoconv = False
    try:
        result = subprocess.run(
            ["which" if platform.system() != "Windows" else "where", "unoconv"], 
            stdout=subprocess.PIPE, 
            stderr=subprocess.PIPE, 
            text=True
        )
        if result.returncode == 0:
            logging.debug(f"unoconv found at: {result.stdout.strip()}")
            has_unoconv = True
        else:
            missing_deps.append("unoconv")
            logging.warning("unoconv not found. .vsd and .vdw processing will be limited.")
    except (subprocess.SubprocessError, FileNotFoundError):
        missing_deps.append("unoconv")
        logging.warning("unoconv not found. .vsd and .vdw processing will be limited.")
    
    # Return overall LibreOffice/unoconv support status
    vsd_support = has_libreoffice or has_unoconv
    
    return VSDX_SUPPORT, vsd_support, missing_deps

# File handling functions
def get_visio_files(directory):
    """Get a list of Visio files in the given directory"""
    files = []
    for filename in os.listdir(directory):
        filepath = os.path.join(directory, filename)
        if os.path.isfile(filepath):
            _, ext = os.path.splitext(filename.lower())
            if ext in SUPPORTED_EXTENSIONS:
                files.append(filepath)
    return files

def get_unique_filename(filepath):
    """Generate a unique filename if file already exists"""
    if not os.path.exists(filepath):
        return filepath
    
    directory, filename = os.path.split(filepath)
    name, ext = os.path.splitext(filename)
    
    counter = 1
    while True:
        new_filepath = os.path.join(directory, f"{name}_{counter}{ext}")
        if not os.path.exists(new_filepath):
            return new_filepath
        counter += 1

# Conversion functions
def convert_vsdx_to_vdx(input_file, output_file):
    """Convert .vsdx or .vsdm file to .vdx format using vsdx library"""
    if not VSDX_SUPPORT:
        raise ImportError("vsdx library is required for processing .vsdx files")
    
    try:
        drawing = vsdx.VisioFile(input_file)
        
        # Create VDX XML structure
        vdx_root = ET.Element("VisioDocument", 
                              xmlns="http://schemas.microsoft.com/visio/2003/core")
        
        # Extract document properties
        doc_props = ET.SubElement(vdx_root, "DocumentProperties")
        title = ET.SubElement(doc_props, "Title")
        title.text = os.path.basename(input_file)
        creator = ET.SubElement(doc_props, "Creator")
        creator.text = "VDXConvert"
        
        # Extract pages
        pages = ET.SubElement(vdx_root, "Pages")
        
        # Process each page in the Visio document
        for idx, page in enumerate(drawing.pages, 1):
            page_elem = ET.SubElement(pages, "Page", ID=str(idx))
            page_elem.set("Name", page.name)
            
            # Add page properties
            page_props = ET.SubElement(page_elem, "PageProperties")
            width = ET.SubElement(page_props, "PageWidth")
            width.text = str(page.width)
            height = ET.SubElement(page_props, "PageHeight")
            height.text = str(page.height)
            
            # Add shapes
            shapes = ET.SubElement(page_elem, "Shapes")
            for shape_id, shape in enumerate(page.shapes, 1):
                shape_elem = ET.SubElement(shapes, "Shape", ID=str(shape_id))
                shape_elem.set("Name", shape.name if shape.name else f"Shape_{shape_id}")
                
                # Add shape properties
                shape_props = ET.SubElement(shape_elem, "ShapeProperties")
                
                # Add position and size
                if hasattr(shape, 'x') and hasattr(shape, 'y'):
                    pos_x = ET.SubElement(shape_props, "PosX")
                    pos_x.text = str(shape.x)
                    pos_y = ET.SubElement(shape_props, "PosY")
                    pos_y.text = str(shape.y)
                
                if hasattr(shape, 'width') and hasattr(shape, 'height'):
                    width = ET.SubElement(shape_props, "Width")
                    width.text = str(shape.width)
                    height = ET.SubElement(shape_props, "Height")
                    height.text = str(shape.height)
        
        # Create XML tree and write to file
        tree = ET.ElementTree(vdx_root)
        tree.write(output_file, encoding="utf-8", xml_declaration=True)
        
        return True
    except Exception as e:
        logging.error(f"Error converting {input_file} to VDX: {str(e)}")
        logging.debug(traceback.format_exc())
        return False

def convert_vsd_to_vdx(input_file, output_file):
    """Convert .vsd or .vdw file to .vdx format using unoconv/LibreOffice"""
    try:
        # Create a temporary directory for conversion
        temp_dir = os.path.join(LOGS_DIR, "temp")
        os.makedirs(temp_dir, exist_ok=True)
        
        # Get base filename without extension
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        
        # Try unoconv first (better conversion quality)
        try:
            logging.debug(f"Attempting conversion with unoconv: {input_file}")
            subprocess.run(
                ["unoconv", "-f", "vdx", "-o", temp_dir, input_file],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                check=True
            )
            
            # Check if the VDX file was created
            temp_vdx = os.path.join(temp_dir, f"{base_name}.vdx")
            if os.path.exists(temp_vdx):
                shutil.copy2(temp_vdx, output_file)
                os.remove(temp_vdx)
                return True
        except (subprocess.SubprocessError, FileNotFoundError) as e:
            logging.debug(f"unoconv conversion failed: {str(e)}")
        
        # Try direct LibreOffice conversion as fallback
        try:
            soffice_cmd = "soffice"
            if platform.system() == "Windows":
                # Try to find LibreOffice executable on Windows
                program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
                libreoffice_path = os.path.join(program_files, "LibreOffice", "program", "soffice.exe")
                if os.path.exists(libreoffice_path):
                    soffice_cmd = f'"{libreoffice_path}"'
            
            logging.debug(f"Attempting conversion with LibreOffice: {input_file}")
            convert_cmd = [
                soffice_cmd,
                "--headless",
                "--convert-to", "vdx",
                "--outdir", temp_dir,
                input_file
            ]
            
            subprocess.run(
                convert_cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                check=True
            )
            
            # Check if the VDX file was created
            temp_vdx = os.path.join(temp_dir, f"{base_name}.vdx")
            if os.path.exists(temp_vdx):
                shutil.copy2(temp_vdx, output_file)
                os.remove(temp_vdx)
                return True
        except (subprocess.SubprocessError, FileNotFoundError) as e:
            logging.debug(f"LibreOffice conversion failed: {str(e)}")
            logging.error(f"Failed to convert {input_file} to VDX format.")
            return False
        
        logging.error(f"All conversion methods failed for {input_file}")
        return False
    except Exception as e:
        logging.error(f"Error during VSD to VDX conversion: {str(e)}")
        logging.debug(traceback.format_exc())
        return False
    finally:
        # Clean up temp directory
        temp_dir = os.path.join(LOGS_DIR, "temp")
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

def process_file(input_file, vsdx_support, vsd_support):
    """Process a single Visio file and convert it to VDX format"""
    filename = os.path.basename(input_file)
    base_name, ext = os.path.splitext(filename)
    ext = ext.lower()
    
    start_time = time.time()
    
    # Create output filename
    output_file = os.path.join(OUTPUT_DIR, f"{base_name}.vdx")
    output_file = get_unique_filename(output_file)
    
    # Create archive filename
    archive_file = os.path.join(ARCHIVE_DIR, filename)
    archive_file = get_unique_filename(archive_file)
    
    logging.info(f"Processing: {filename}")
    
    try:
        success = False
        
        # Process based on file extension
        if ext in ['.vsdx', '.vsdm']:
            if vsdx_support:
                success = convert_vsdx_to_vdx(input_file, output_file)
            else:
                logging.error(f"Cannot process {ext} files: vsdx library not available")
        elif ext in ['.vsd', '.vdw']:
            if vsd_support:
                success = convert_vsd_to_vdx(input_file, output_file)
            else:
                logging.error(f"Cannot process {ext} files: LibreOffice/unoconv not available")
        else:
            logging.error(f"Unsupported file extension: {ext}")
        
        end_time = time.time()
        processing_time = end_time - start_time
        
        # Move original file to archive on success
        if success:
            shutil.move(input_file, archive_file)
            logging.info(f"✅ Conversion successful: {os.path.basename(output_file)} ({processing_time:.2f}s)")
            return {
                "filename": filename,
                "output": os.path.basename(output_file),
                "archive": os.path.basename(archive_file),
                "success": True,
                "time": processing_time,
                "error": None
            }
        else:
            logging.error(f"❌ Conversion failed: {filename}")
            return {
                "filename": filename,
                "output": None,
                "archive": None,
                "success": False,
                "time": processing_time,
                "error": "Conversion process failed"
            }
    except Exception as e:
        end_time = time.time()
        processing_time = end_time - start_time
        logging.error(f"❌ Error processing {filename}: {str(e)}")
        logging.debug(traceback.format_exc())
        return {
            "filename": filename,
            "output": None,
            "archive": None,
            "success": False,
            "time": processing_time,
            "error": str(e)
        }

def save_csv_report(results, save_path):
    """Save analysis report to CSV file"""
    fieldnames = ["filename", "output", "archive", "success", "time", "error"]
    
    with open(save_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for result in results:
            writer.writerow(result)
    
    logging.info(f"CSV report saved to: {save_path}")

def print_summary(results):
    """Print summary of conversion results"""
    total = len(results)
    successful = sum(1 for r in results if r['success'])
    failed = total - successful
    total_time = sum(r['time'] for r in results)
    
    # Print divider
    print("\n" + "="*60)
    
    # Print header
    if COLOR_SUPPORT:
        print(f"\n{Fore.CYAN}{Style.BRIGHT}VDXConvert Summary{Style.RESET_ALL}")
    else:
        print("\nVDXConvert Summary")
    
    print(f"\nTotal files processed: {total}")
    
    if COLOR_SUPPORT:
        print(f"Successful conversions: {Fore.GREEN}{successful}{Style.RESET_ALL}")
        print(f"Failed conversions: {Fore.RED if failed > 0 else ''}{failed}{Style.RESET_ALL}")
    else:
        print(f"Successful conversions: {successful}")
        print(f"Failed conversions: {failed}")
    
    print(f"Total processing time: {total_time:.2f}s")
    
    # Print failed files if any
    if failed > 0:
        print("\nFailed conversions:")
        for result in results:
            if not result['success']:
                if COLOR_SUPPORT:
                    print(f"  {Fore.RED}✖ {result['filename']}: {result['error']}{Style.RESET_ALL}")
                else:
                    print(f"  ✖ {result['filename']}: {result['error']}")
    
    # Print divider
    print("\n" + "="*60 + "\n")

def main():
    """Main function"""
    # Parse command line arguments
    parser = argparse.ArgumentParser(description=f"{APP_NAME} - Batch converter for Visio files to VDX format")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable verbose logging")
    parser.add_argument("--no-report", action="store_true", help="Don't save CSV report")
    args = parser.parse_args()
    
    # Setup logging
    logger = setup_logging(args.verbose)
    
    # Print welcome message
    if COLOR_SUPPORT:
        print(f"\n{Fore.CYAN}{Style.BRIGHT}{APP_NAME} v{VERSION}{Style.RESET_ALL}")
    else:
        print(f"\n{APP_NAME} v{VERSION}")
    print("Batch converter for Visio files to VDX format")
    print("Author: Sam Lyndon\n")
    
    # Ensure directories exist
    for directory in [INPUT_DIR, OUTPUT_DIR, ARCHIVE_DIR, LOGS_DIR]:
        os.makedirs(directory, exist_ok=True)
    
    # Check dependencies
    logging.info("Checking dependencies...")
    vsdx_support, vsd_support, missing_deps = check_dependencies()
    
    if missing_deps:
        if COLOR_SUPPORT:
            print(f"\n{Fore.YELLOW}Warning: Some dependencies are missing:{Style.RESET_ALL}")
        else:
            print("\nWarning: Some dependencies are missing:")
        
        for dep in missing_deps:
            print(f"  - {dep}")
        
        print("\nSome file types may not be processed correctly.")
        print("See README.md for installation instructions.\n")
    
    # Get Visio files from input directory
    logging.info("Scanning input directory...")
    visio_files = get_visio_files(INPUT_DIR)
    
    if not visio_files:
        logging.warning("No Visio files found in input directory.")
        print("\nNo Visio files found in the input directory.")
        print(f"Please place Visio files (VSD, VSDX, VSDM, VDW) in the input folder: {INPUT_DIR}")
        return
    
    # Process files
    logging.info(f"Found {len(visio_files)} Visio files to process")
    print(f"\nFound {len(visio_files)} Visio files to process.")
    
    results = []
    
    if PROGRESS_BAR_SUPPORT:
        for input_file in tqdm(visio_files, desc="Converting", unit="file"):
            result = process_file(input_file, vsdx_support, vsd_support)
            results.append(result)
    else:
        for input_file in visio_files:
            result = process_file(input_file, vsdx_support, vsd_support)
            results.append(result)
    
    # Print summary
    print_summary(results)
    
    # Save CSV report
    if not args.no_report:
        save_csv = True
        try:
            save_response = input("Save detailed CSV report? [Y/n]: ").strip().lower()
            save_csv = save_response != 'n'
        except:
            save_csv = True
        
        if save_csv:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_path = os.path.join(LOGS_DIR, f"conversion_report_{timestamp}.csv")
            save_csv_report(results, report_path)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        sys.exit(1)
    except Exception as e:
        logging.critical(f"Unhandled exception: {str(e)}")
        logging.debug(traceback.format_exc())
        print(f"\nAn unexpected error occurred: {str(e)}")
        sys.exit(1)