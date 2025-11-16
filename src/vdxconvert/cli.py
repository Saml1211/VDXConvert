#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Command-line interface for VDXConvert.

Handles argument parsing, main conversion flow, and user interaction.
"""

import argparse
import logging
import sys
import time
from pathlib import Path
from typing import List, Dict, Any, Optional

from .config import (
    APP_NAME,
    VERSION,
    AUTHOR,
    INPUT_DIR,
    OUTPUT_DIR,
    ARCHIVE_DIR,
    LOGS_DIR,
    OUTPUT_EXTENSION,
    PROMPT_SAVE_CSV
)
from .converters import VSDXConverter, VSDConverter
from .exceptions import VDXConvertError, ValidationError, ConversionError
from .logger import setup_logging, COLOR_SUPPORT
from .utils import (
    get_visio_files,
    get_unique_filename,
    safe_move_file,
    save_csv_report,
    generate_report_filename,
    ensure_directories_exist
)
from .validators import validate_input_file

# Optional progress bar support
try:
    from tqdm import tqdm
    PROGRESS_BAR_SUPPORT = True
except ImportError:
    PROGRESS_BAR_SUPPORT = False

# Color support
if COLOR_SUPPORT:
    from colorama import Fore, Style

logger = logging.getLogger(__name__)


class VDXConvertCLI:
    """Command-line interface for VDXConvert."""

    def __init__(self) -> None:
        """Initialize the CLI."""
        self.converters = [
            VSDXConverter(),
            VSDConverter()
        ]

    def print_banner(self) -> None:
        """Print application banner."""
        if COLOR_SUPPORT:
            print(f"\n{Fore.CYAN}{Style.BRIGHT}{APP_NAME} v{VERSION}{Style.RESET_ALL}")
        else:
            print(f"\n{APP_NAME} v{VERSION}")
        print("Batch converter for Visio files to VDX format")
        print(f"Author: {AUTHOR}\n")

    def check_dependencies(self) -> None:
        """Check and report on available dependencies."""
        logger.info("Checking dependencies...")

        missing_deps = []
        for converter in self.converters:
            missing = converter.get_missing_dependencies()
            missing_deps.extend(missing)

        if missing_deps:
            if COLOR_SUPPORT:
                print(f"\n{Fore.YELLOW}Warning: Some dependencies are missing:{Style.RESET_ALL}")
            else:
                print("\nWarning: Some dependencies are missing:")

            for dep in set(missing_deps):
                print(f"  - {dep}")

            print("\nSome file types may not be processed correctly.")
            print("See README.md for installation instructions.\n")

    def get_converter_for_file(self, filepath: Path) -> Optional[object]:
        """
        Get the appropriate converter for a file.

        Args:
            filepath: Path to the file

        Returns:
            Converter instance or None if no converter available
        """
        for converter in self.converters:
            if converter.can_convert(filepath) and converter.check_dependencies():
                return converter
        return None

    def process_file(self, input_file: Path) -> Dict[str, Any]:
        """
        Process a single Visio file.

        Args:
            input_file: Path to input file

        Returns:
            Dictionary containing processing results
        """
        filename = input_file.name
        start_time = time.time()

        # Create output and archive paths
        output_file = OUTPUT_DIR / f"{input_file.stem}{OUTPUT_EXTENSION}"
        output_file = get_unique_filename(output_file)

        archive_file = ARCHIVE_DIR / filename
        archive_file = get_unique_filename(archive_file)

        logger.info(f"Processing: {filename}")

        try:
            # Validate input file
            validate_input_file(input_file)

            # Get appropriate converter
            converter = self.get_converter_for_file(input_file)
            if not converter:
                raise ConversionError(
                    f"No converter available for {input_file.suffix}",
                    "Required dependencies may be missing"
                )

            # Perform conversion
            success = converter.convert(input_file, output_file)

            end_time = time.time()
            processing_time = end_time - start_time

            if success:
                # Move original to archive
                final_archive = safe_move_file(input_file, archive_file)
                logger.info(f"✅ Conversion successful: {output_file.name} ({processing_time:.2f}s)")

                return {
                    "filename": filename,
                    "output": output_file.name,
                    "archive": final_archive.name,
                    "success": True,
                    "time": processing_time,
                    "error": None
                }
            else:
                raise ConversionError("Conversion returned False")

        except VDXConvertError as e:
            end_time = time.time()
            processing_time = end_time - start_time
            error_msg = str(e)
            logger.error(f"❌ {error_msg}")
            if e.details:
                logger.debug(f"Details: {e.details}")

            return {
                "filename": filename,
                "output": None,
                "archive": None,
                "success": False,
                "time": processing_time,
                "error": error_msg
            }

        except Exception as e:
            end_time = time.time()
            processing_time = end_time - start_time
            error_msg = f"Unexpected error: {str(e)}"
            logger.error(f"❌ Error processing {filename}: {error_msg}")
            logger.debug(f"Exception details:", exc_info=True)

            return {
                "filename": filename,
                "output": None,
                "archive": None,
                "success": False,
                "time": processing_time,
                "error": error_msg
            }

    def print_summary(self, results: List[Dict[str, Any]]) -> None:
        """
        Print summary of conversion results.

        Args:
            results: List of result dictionaries
        """
        total = len(results)
        successful = sum(1 for r in results if r['success'])
        failed = total - successful
        total_time = sum(r['time'] for r in results)

        # Print divider
        print("\n" + "=" * 60)

        # Print header
        if COLOR_SUPPORT:
            print(f"\n{Fore.CYAN}{Style.BRIGHT}{APP_NAME} Summary{Style.RESET_ALL}")
        else:
            print(f"\n{APP_NAME} Summary")

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
        print("\n" + "=" * 60 + "\n")

    def run(self, args: argparse.Namespace) -> int:
        """
        Run the main conversion process.

        Args:
            args: Parsed command-line arguments

        Returns:
            Exit code (0 for success, 1 for error)
        """
        # Setup logging
        setup_logging(args.verbose)

        # Print banner
        self.print_banner()

        # Ensure directories exist
        ensure_directories_exist([INPUT_DIR, OUTPUT_DIR, ARCHIVE_DIR, LOGS_DIR])

        # Check dependencies
        self.check_dependencies()

        # Get Visio files
        logger.info("Scanning input directory...")
        visio_files = get_visio_files(INPUT_DIR)

        if not visio_files:
            logger.warning("No Visio files found in input directory.")
            print("\nNo Visio files found in the input directory.")
            print(f"Please place Visio files (VSD, VSDX, VSDM, VDW) in: {INPUT_DIR}")
            return 0

        # Process files
        logger.info(f"Found {len(visio_files)} Visio files to process")
        print(f"\nFound {len(visio_files)} Visio files to process.")

        results = []

        if PROGRESS_BAR_SUPPORT:
            for input_file in tqdm(visio_files, desc="Converting", unit="file"):
                result = self.process_file(input_file)
                results.append(result)
        else:
            for input_file in visio_files:
                result = self.process_file(input_file)
                results.append(result)

        # Print summary
        self.print_summary(results)

        # Save CSV report
        if not args.no_report:
            save_csv = True
            try:
                save_response = input(PROMPT_SAVE_CSV).strip().lower()
                save_csv = save_response not in ('n', 'no')
            except (EOFError, KeyboardInterrupt):
                save_csv = True
                print()  # New line after interrupt

            if save_csv:
                report_filename = generate_report_filename()
                report_path = LOGS_DIR / report_filename
                save_csv_report(results, report_path)
                logger.info(f"CSV report saved to: {report_path}")

        return 0


def parse_arguments() -> argparse.Namespace:
    """
    Parse command-line arguments.

    Returns:
        Parsed arguments namespace
    """
    parser = argparse.ArgumentParser(
        description=f"{APP_NAME} - Batch converter for Visio files to VDX format"
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose logging"
    )
    parser.add_argument(
        "--no-report",
        action="store_true",
        help="Don't save CSV report"
    )
    return parser.parse_args()


def main() -> int:
    """
    Main entry point for the CLI.

    Returns:
        Exit code
    """
    try:
        args = parse_arguments()
        cli = VDXConvertCLI()
        return cli.run(args)

    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        return 1

    except Exception as e:
        logging.critical(f"Unhandled exception: {str(e)}", exc_info=True)
        print(f"\nAn unexpected error occurred: {str(e)}")
        return 1
