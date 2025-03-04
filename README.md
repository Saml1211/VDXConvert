# VDXConvert

A lightweight, cross-platform tool for batch converting Visio files to VDX format.

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![License](https://img.shields.io/badge/license-MIT-green)

VDXConvert is a Python utility that automates the conversion of Visio files (VSD, VSDX, VSDM, VDW) to VDX format. VDX (Visio XML Drawing) is an XML-based format that offers better interoperability with other software and version control systems.

## Features

- Batch processing of multiple Visio files
- Support for various Visio formats:
  - `.vsdx` - Visio Drawing (XML)
  - `.vsdm` - Visio Drawing with Macros
  - `.vsd` - Visio Drawing (Binary)
  - `.vdw` - Visio Web Drawing
- Automatic organization with input, output, and archive folders
- Detailed logging and error handling
- Comprehensive analysis reports
- Cross-platform compatibility (Windows and macOS)
- File versioning to prevent overwriting existing files

## Directory Structure

```
VDXConvert/
├── vdxconvert.py          # Main conversion script
├── README.md              # Documentation
├── requirements.txt       # Python dependencies
├── input/                 # Source Visio files
├── output/                # Converted VDX files
├── archive/               # Successfully processed source files
└── logs/                  # Log files and reports
```

## Requirements

- Python 3.6 or higher
- Required Python packages (installed via `pip`):
  - `vsdx`: For processing `.vsdx` and `.vsdm` files
  - Additional packages listed in `requirements.txt`
- LibreOffice (for processing `.vsd` and `.vdw` files)
- unoconv (optional, improves conversion quality)

## Installation

### Windows

1. **Install Python:**
   - Download and install Python from [python.org](https://www.python.org/downloads/windows/)
   - Make sure to check "Add Python to PATH" during installation

2. **Set up a Virtual Environment (Recommended):**
   ```cmd
   cd path\to\VDXConvert
   python -m venv venv
   venv\Scripts\activate
   ```
   - This creates an isolated environment for the project
   - Your command prompt should now show `(venv)` at the beginning

3. **Install Python Dependencies:**
   ```cmd
   pip install -r requirements.txt
   ```

4. **Install LibreOffice:**
   - Download and install from [libreoffice.org](https://www.libreoffice.org/download/download/)
   - The default installation path should be detected automatically

5. **Install unoconv (optional but recommended):**
   - Using pip: `pip install unoconv`
   - Or download from [GitHub](https://github.com/unoconv/unoconv)

### macOS

1. **Install Python:**
   ```bash
   # Using Homebrew
   brew install python
   
   # Or download from python.org
   ```

2. **Set up a Virtual Environment (Recommended):**
   ```bash
   cd path/to/VDXConvert
   python3 -m venv venv
   source venv/bin/activate
   ```
   - This creates an isolated environment for the project
   - Your terminal should now show `(venv)` at the beginning

3. **Install Python Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Install LibreOffice:**
   ```bash
   # Using Homebrew
   brew install --cask libreoffice
   
   # Or download from libreoffice.org
   ```

5. **Install unoconv (optional but recommended):**
   ```bash
   # Using Homebrew
   brew install unoconv
   
   # Or using pip
   pip3 install unoconv
   ```

## Usage

1. **Place Visio files in the input folder:**
   - Copy or move your Visio files (`.vsd`, `.vsdx`, `.vsdm`, `.vdw`) to the `input/` directory

2. **Run the conversion script:**

   **Windows:**
   ```cmd
   cd path\to\VDXConvert
   
   # If using virtual environment (recommended)
   venv\Scripts\activate
   python vdxconvert.py
   
   # Or without virtual environment
   python vdxconvert.py
   ```

   **macOS:**
   ```bash
   cd path/to/VDXConvert
   
   # If using virtual environment (recommended)
   source venv/bin/activate
   python vdxconvert.py
   
   # Or without virtual environment
   python3 vdxconvert.py
   ```

3. **Check the results:**
   - Converted VDX files will be in the `output/` directory
   - Original files will be moved to the `archive/` directory upon successful conversion
   - The script will display a summary of the conversion process
   - Detailed logs are stored in the `logs/` directory

### Command Line Options

```
usage: vdxconvert.py [-h] [-v] [--no-report]

VDXConvert - Batch converter for Visio files to VDX format

options:
  -h, --help      show this help message and exit
  -v, --verbose   Enable verbose logging
  --no-report     Don't save CSV report
```

## Troubleshooting

### Common Issues

1. **Missing Dependencies:**
   - If you see warnings about missing dependencies, follow the installation instructions above
   - For `.vsdx` files, ensure the `vsdx` Python package is installed
   - For `.vsd` files, ensure LibreOffice is installed and in your PATH

2. **Conversion Failures:**
   - Check the logs in the `logs/` directory for detailed error messages
   - Ensure the Visio files are not corrupted or password-protected
   - Try opening and re-saving the file in Visio if possible

3. **LibreOffice Not Found:**
   - Windows: Ensure LibreOffice is installed in the default location or add it to your PATH
   - macOS: If installed via Homebrew, it should be detected automatically

4. **Virtual Environment Issues:**
   - If you're having issues with dependencies, ensure your virtual environment is activated
   - Windows: The command prompt should show `(venv)` at the beginning when activated
   - macOS: The terminal should show `(venv)` at the beginning when activated
   - If needed, you can recreate the virtual environment:
     ```
     # Delete the old environment
     rm -rf venv  # macOS/Linux
     rmdir /s /q venv  # Windows
     
     # Create a new environment
     python -m venv venv  # Windows
     python3 -m venv venv  # macOS/Linux
     
     # Activate and reinstall dependencies
     ```

### Debug Mode

Run the script with the `-v` flag to enable verbose logging:

```bash
python vdxconvert.py -v
```

This will provide more detailed information about the conversion process and any errors.

## License

MIT License

Copyright (c) 2025 Sam Lyndon

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.