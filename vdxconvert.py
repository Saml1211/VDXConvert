#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VDXConvert - Batch converter for Visio files to VDX format

BACKWARD COMPATIBILITY WRAPPER
==============================
This file provides backward compatibility for the legacy vdxconvert.py script.
The actual implementation has been refactored into a modular package structure
located in src/vdxconvert/.

For new code, please use:
    from vdxconvert.cli import main
    # or
    python -m vdxconvert [options]

This wrapper will be maintained for backward compatibility but may be deprecated
in future versions.

Author: Sam Lyndon
Version: 1.0.0
License: MIT
"""

import sys
from pathlib import Path

# Add src directory to path to import the new package
src_path = Path(__file__).parent / "src"
if src_path.exists():
    sys.path.insert(0, str(src_path))

try:
    # Import the new modular implementation
    from vdxconvert.cli import main

    # Run the CLI
    if __name__ == "__main__":
        sys.exit(main())

except ImportError as e:
    # Fallback message if new package is not available
    print("Error: Unable to import vdxconvert package.")
    print(f"Details: {e}")
    print("\nPlease ensure the package is properly installed:")
    print("  pip install -e .")
    print("\nOr run directly from the package:")
    print("  python -m vdxconvert")
    sys.exit(1)
