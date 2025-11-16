# Migration Guide - VDXConvert v1.0 Refactoring

## Overview

VDXConvert has been significantly refactored to improve code quality, maintainability, and testability. This guide will help you understand the changes and migrate to the new structure.

## What Changed?

### Previous Structure (Monolithic)
```
VDXConvert/
├── vdxconvert.py (573 lines - all functionality)
├── README.md
└── requirements.txt
```

### New Structure (Modular)
```
VDXConvert/
├── src/vdxconvert/           # Main package
│   ├── __init__.py
│   ├── __main__.py           # Entry point for python -m vdxconvert
│   ├── cli.py                # Command-line interface
│   ├── config.py             # All configuration constants
│   ├── exceptions.py         # Custom exception hierarchy
│   ├── logger.py             # Logging configuration
│   ├── utils.py              # File operation utilities
│   ├── validators.py         # Input validation & security
│   └── converters/           # Converter implementations
│       ├── __init__.py
│       ├── base.py           # Base converter interface
│       ├── vsdx_converter.py # VSDX/VSDM converter
│       └── vsd_converter.py  # VSD/VDW converter
├── tests/                    # Comprehensive test suite
│   ├── __init__.py
│   ├── conftest.py
│   ├── test_cli.py
│   ├── test_converters.py
│   ├── test_utils.py
│   └── test_validators.py
├── vdxconvert.py            # Backward compatibility wrapper
├── pyproject.toml           # Modern packaging configuration
├── requirements-dev.txt     # Development dependencies
├── .flake8                  # Linting configuration
└── README.md
```

## Key Improvements

### 1. **SOLID Principles Applied**
- **Single Responsibility**: Each module has one clear purpose
- **Open/Closed**: Easy to add new converters without modifying existing code
- **Dependency Inversion**: Converters implement abstract base class

### 2. **Security Enhancements**
- Path traversal protection
- File size limits (max 500MB by default)
- Filename sanitization
- Input validation on all user inputs

### 3. **Type Safety**
- Comprehensive type hints throughout
- mypy type checking configured
- Better IDE support and autocomplete

### 4. **Testing**
- 90+ unit tests with >80% coverage target
- pytest framework
- Mocked external dependencies
- Easy to run: `pytest tests/`

### 5. **Code Quality**
- Black formatting (consistent style)
- flake8 linting
- isort import sorting
- Automated quality checks

### 6. **Better Error Handling**
- Custom exception hierarchy:
  - `VDXConvertError` (base)
  - `ValidationError`
  - `ConversionError`
  - `DependencyError`
  - `FileOperationError`
  - `SecurityError`

## Migration Steps

### For End Users (No Code Changes)

The refactoring maintains **100% backward compatibility** for command-line usage:

```bash
# Old way (still works)
python vdxconvert.py [options]

# New way (recommended)
python -m vdxconvert [options]

# Or after installation
pip install -e .
vdxconvert [options]
```

**No changes needed to your workflow!**

### For Developers

#### Installing for Development

```bash
# Clone the repository
git clone https://github.com/Saml1211/VDXConvert.git
cd VDXConvert

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install in development mode with dev dependencies
pip install -e .[dev]
```

#### Running Tests

```bash
# Run all tests
pytest

# Run with coverage report
pytest --cov

# Run specific test file
pytest tests/test_validators.py

# Run with verbose output
pytest -v
```

#### Code Quality Checks

```bash
# Format code with Black
black src/ tests/

# Check code style with flake8
flake8 src/ tests/

# Sort imports with isort
isort src/ tests/

# Type check with mypy
mypy src/
```

#### Importing in Your Code

**Old (if you were importing the monolithic script):**
```python
# This was not really possible before
```

**New:**
```python
# Import specific components
from vdxconvert.cli import VDXConvertCLI
from vdxconvert.converters import VSDXConverter, VSDConverter
from vdxconvert.validators import validate_input_file
from vdxconvert.exceptions import ConversionError

# Use as a library
from vdxconvert.converters import VSDXConverter
from pathlib import Path

converter = VSDXConverter()
if converter.check_dependencies():
    converter.convert(Path("input.vsdx"), Path("output.vdx"))
```

## Configuration Changes

### Environment Variables (Future)

The new structure is ready for environment-based configuration:

```python
# In config.py - ready for environment variables
MAX_FILE_SIZE_MB = int(os.getenv("VDXCONVERT_MAX_FILE_SIZE", "500"))
```

### Custom Converters

You can now easily add custom converters:

```python
from pathlib import Path
from vdxconvert.converters.base import BaseConverter

class MyCustomConverter(BaseConverter):
    def can_convert(self, filepath: Path) -> bool:
        return filepath.suffix.lower() == '.myformat'

    def check_dependencies(self) -> bool:
        return True  # Check your dependencies

    def get_missing_dependencies(self) -> list[str]:
        return []  # Return missing deps

    def convert(self, input_file: Path, output_file: Path) -> bool:
        # Your conversion logic
        return True
```

## Breaking Changes

### None for CLI Users ✅

If you're using the command-line interface, **nothing breaks**.

### For Developers

If you were somehow importing from the old `vdxconvert.py` (unlikely), you'll need to update imports:

```python
# Old (was never really supported)
from vdxconvert import some_function

# New
from vdxconvert.cli import main
from vdxconvert.converters import VSDXConverter
```

## Benefits of Migrating

1. **Better Error Messages**: More specific exceptions with details
2. **Type Safety**: Type hints help catch errors early
3. **Testable**: Can now unit test individual components
4. **Maintainable**: Clear separation of concerns
5. **Extensible**: Easy to add new file formats
6. **Secure**: Input validation prevents security issues
7. **Professional**: Follows Python packaging best practices

## Rollback Plan

If you encounter issues with the refactored version:

1. Check out the previous commit:
   ```bash
   git checkout <previous-commit-hash>
   ```

2. Or use the old vdxconvert.py from git history

3. Report the issue on GitHub with details

## Questions or Issues?

- **GitHub Issues**: https://github.com/Saml1211/VDXConvert/issues
- **Documentation**: See README.md for usage
- **This Guide**: For migration questions

## Timeline

- **v1.0.0**: Current version with refactored structure
- **v1.1.0** (planned): Additional converters, performance improvements
- **v2.0.0** (future): May deprecate backward compatibility wrapper

---

**Thank you for using VDXConvert!** The refactoring makes the codebase more maintainable and extensible while maintaining the simplicity you expect.
