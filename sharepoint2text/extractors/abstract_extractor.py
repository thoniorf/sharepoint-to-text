import typing
from abc import abstractmethod
from dataclasses import dataclass
from pathlib import Path
from typing import Protocol


@dataclass
class FileMetadataInterface:
    filename: str | None = None
    file_extension: str | None = None
    file_path: str | None = None
    folder_path: str | None = None

    def populate_from_path(self, path: str | Path | None) -> None:
        """Populate file metadata fields from a path."""
        if path is None:
            return
        p = Path(path)
        self.filename = p.name
        self.file_extension = p.suffix
        self.file_path = str(p.resolve()) if p.exists() else str(p)
        self.folder_path = (
            str(p.parent.resolve()) if p.parent.exists() else str(p.parent)
        )


class ExtractionInterface(Protocol):
    @abstractmethod
    def iterator(self) -> typing.Iterator[str]:
        """
        Returns an iterator over the extracted text i.e., the main text body of a file.
        Additional text areas may be missing if they are not part of the main text body of the file.
        This greatly depends on the underlying data source.
        A PDF returns text per pages, PowerPoint files return slides as units.
        Excel files return sheets.
        Content of footnotes, headers or alike is not part of this iterator's return values.
        The legacy and modern Word documents have no per-page representation in the files, they return only a single unit which is the full text.
        """
        ...

    @abstractmethod
    def get_full_text(self) -> str:
        """Full text of the slide deck as one single block of text"""
        ...

    @abstractmethod
    def get_metadata(self) -> FileMetadataInterface:
        """Returns the metadata of the extracted file"""
        ...
