"""
XLS content extractor using pandas and olefile libraries.
"""

import io
import logging
import typing
from dataclasses import dataclass, field
from typing import Any, Dict, List

import olefile
import pandas as pd

from sharepoint2text.extractors.abstract_extractor import (
    ExtractionInterface,
    FileMetadataInterface,
)

logger = logging.getLogger(__name__)


@dataclass
class MicrosoftXlsMetadata(FileMetadataInterface):
    title: str = ""
    author: str = ""
    subject: str = ""
    company: str = ""
    last_saved_by: str = ""
    created: str = ""
    modified: str = ""


@dataclass
class MicrosoftXlsSheet:
    name: str = ""
    data: List[Dict[str, Any]] = field(default_factory=list)
    text: str = ""


@dataclass
class MicrosoftXlsContent(ExtractionInterface):
    metadata: MicrosoftXlsMetadata = field(default_factory=MicrosoftXlsMetadata)
    sheets: List[MicrosoftXlsSheet] = field(default_factory=list)
    full_text: str = ""

    def iterator(self) -> typing.Iterator[str]:
        for sheet in self.sheets:
            yield sheet.text

    def get_full_text(self) -> str:
        return self.full_text

    def get_metadata(self) -> FileMetadataInterface:
        """Returns the metadata of the extracted file."""
        return self.metadata


def _read_content(file_like: io.BytesIO) -> List[MicrosoftXlsSheet]:
    logger.debug("Reading content")
    xls = pd.read_excel(file_like, engine="calamine", sheet_name=None)

    sheets = []
    for sheet_name, df in xls.items():
        logger.debug(f"Reading sheet: [{sheet_name}]")
        data = df.to_dict(orient="records")
        text = df.to_string(index=False)
        sheets.append(
            MicrosoftXlsSheet(
                name=str(sheet_name),
                data=data,
                text=text,
            )
        )
    return sheets


def _read_metadata(file_like: io.BytesIO) -> MicrosoftXlsMetadata:
    ole = olefile.OleFileIO(file_like)
    meta = ole.get_metadata()

    result = MicrosoftXlsMetadata(
        title=meta.title.decode("utf-8") if meta.title else "",
        author=meta.author.decode("utf-8") if meta.author else "",
        subject=meta.subject.decode("utf-8") if meta.subject else "",
        company=meta.company.decode("utf-8") if meta.company else "",
        last_saved_by=meta.last_saved_by.decode("utf-8") if meta.last_saved_by else "",
        created=meta.create_time.isoformat() if meta.create_time else "",
        modified=meta.last_saved_time.isoformat() if meta.last_saved_time else "",
    )
    ole.close()
    return result


def read_xls(file_like: io.BytesIO, path: str | None = None) -> MicrosoftXlsContent:
    """
    Extract all relevant content from an XLS file.

    Args:
        file_like: A BytesIO object containing the XLS file data.
        path: Optional file path to populate file metadata fields.

    Returns:
        MicrosoftXlsContent dataclass with all extracted content.
    """
    file_like.seek(0)
    sheets = _read_content(file_like=file_like)
    file_like.seek(0)
    metadata = _read_metadata(file_like=file_like)
    metadata.populate_from_path(path)

    full_text = "\n\n".join(sheet.text for sheet in sheets)

    return MicrosoftXlsContent(
        metadata=metadata,
        sheets=sheets,
        full_text=full_text,
    )
