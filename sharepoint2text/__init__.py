"""
sharepoint-to-text: Text extraction library for SharePoint file formats.

A Python library for extracting plain text content from files typically found
in SharePoint repositories. Supports both modern Office Open XML formats and
legacy binary formats, plus PDF documents.
"""

from sharepoint2text.router import get_extractor

__version__ = "0.1.0"
__all__ = ["get_extractor", "__version__"]
