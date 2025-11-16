#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Entry point for running VDXConvert as a module.

Usage:
    python -m vdxconvert [options]
"""

import sys
from .cli import main

if __name__ == "__main__":
    sys.exit(main())
