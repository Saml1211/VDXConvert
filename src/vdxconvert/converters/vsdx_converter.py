#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VSDX/VSDM file converter for VDXConvert.

Converts .vsdx and .vsdm files to VDX format using the vsdx library.
"""

import logging
import traceback
from pathlib import Path
from typing import List
from xml.etree import ElementTree as ET

from ..config import VSDX_EXTENSIONS, VDX_NAMESPACE
from ..exceptions import ConversionError, DependencyError
from .base import BaseConverter

# Try to import vsdx module
try:
    import vsdx
    VSDX_AVAILABLE = True
except ImportError:
    VSDX_AVAILABLE = False

logger = logging.getLogger(__name__)


class VSDXConverter(BaseConverter):
    """Converter for .vsdx and .vsdm files."""

    def can_convert(self, filepath: Path) -> bool:
        """
        Check if this converter can handle the given file.

        Args:
            filepath: Path to the file to check

        Returns:
            True if file extension is .vsdx or .vsdm
        """
        return filepath.suffix.lower() in VSDX_EXTENSIONS

    def check_dependencies(self) -> bool:
        """
        Check if vsdx library is available.

        Returns:
            True if vsdx library is installed
        """
        return VSDX_AVAILABLE

    def get_missing_dependencies(self) -> List[str]:
        """
        Get list of missing dependencies.

        Returns:
            List containing 'vsdx' if not available, empty list otherwise
        """
        if not VSDX_AVAILABLE:
            return ['vsdx']
        return []

    def convert(self, input_file: Path, output_file: Path) -> bool:
        """
        Convert .vsdx or .vsdm file to .vdx format.

        Args:
            input_file: Path to input VSDX/VSDM file
            output_file: Path where VDX output should be saved

        Returns:
            True if conversion was successful

        Raises:
            DependencyError: If vsdx library is not available
            ConversionError: If conversion fails
        """
        if not VSDX_AVAILABLE:
            raise DependencyError(
                "vsdx library not available",
                "Install with: pip install vsdx"
            )

        try:
            logger.debug(f"Opening VSDX file: {input_file}")
            drawing = vsdx.VisioFile(str(input_file))

            # Create VDX XML structure
            vdx_root = ET.Element("VisioDocument", xmlns=VDX_NAMESPACE)

            # Extract document properties
            doc_props = ET.SubElement(vdx_root, "DocumentProperties")
            title = ET.SubElement(doc_props, "Title")
            title.text = input_file.name
            creator = ET.SubElement(doc_props, "Creator")
            creator.text = "VDXConvert"

            # Extract pages
            pages_elem = ET.SubElement(vdx_root, "Pages")

            # Process each page in the Visio document
            for idx, page in enumerate(drawing.pages, 1):
                page_elem = ET.SubElement(pages_elem, "Page", ID=str(idx))
                page_elem.set("Name", page.name)

                # Add page properties
                page_props = ET.SubElement(page_elem, "PageProperties")

                # Add page dimensions
                if hasattr(page, 'width'):
                    width = ET.SubElement(page_props, "PageWidth")
                    width.text = str(page.width)
                if hasattr(page, 'height'):
                    height = ET.SubElement(page_props, "PageHeight")
                    height.text = str(page.height)

                # Add shapes
                shapes_elem = ET.SubElement(page_elem, "Shapes")
                for shape_id, shape in enumerate(page.shapes, 1):
                    shape_elem = ET.SubElement(shapes_elem, "Shape", ID=str(shape_id))
                    shape_name = shape.name if hasattr(shape, 'name') and shape.name else f"Shape_{shape_id}"
                    shape_elem.set("Name", shape_name)

                    # Add shape properties
                    shape_props = ET.SubElement(shape_elem, "ShapeProperties")

                    # Add position and size if available
                    if hasattr(shape, 'x') and hasattr(shape, 'y'):
                        pos_x = ET.SubElement(shape_props, "PosX")
                        pos_x.text = str(shape.x)
                        pos_y = ET.SubElement(shape_props, "PosY")
                        pos_y.text = str(shape.y)

                    if hasattr(shape, 'width') and hasattr(shape, 'height'):
                        width_elem = ET.SubElement(shape_props, "Width")
                        width_elem.text = str(shape.width)
                        height_elem = ET.SubElement(shape_props, "Height")
                        height_elem.text = str(shape.height)

            # Create XML tree and write to file
            logger.debug(f"Writing VDX file: {output_file}")
            tree = ET.ElementTree(vdx_root)
            tree.write(str(output_file), encoding="utf-8", xml_declaration=True)

            logger.debug(f"Successfully converted {input_file.name} to VDX")
            return True

        except Exception as e:
            logger.error(f"Error converting {input_file} to VDX: {str(e)}")
            logger.debug(traceback.format_exc())
            raise ConversionError(
                f"Failed to convert {input_file.name}",
                str(e)
            )
