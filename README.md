# sharepoint-to-text

A **pure Python** library for extracting text, metadata, and structured elements from Microsoft Office files—both modern (`.docx`, `.xlsx`, `.pptx`) and legacy (`.doc`, `.xls`, `.ppt`) formats—plus PDF and plain text.

## Why This Library?

### Pure Python, No External Dependencies

Unlike popular alternatives that shell out to **LibreOffice** or **Apache Tika** (requiring Java), `sharepoint-to-text` is a **native Python implementation** with no system-level dependencies:

| Approach | Requirements | Cross-platform | Container-friendly |
|----------|-------------|----------------|-------------------|
| **sharepoint-to-text** | `pip install` only | Yes | Yes (minimal image) |
| LibreOffice-based | LibreOffice install, X11/headless setup | Complex | Large images (~1GB+) |
| Apache Tika | Java runtime, Tika server | Complex | Heavy (~500MB+) |
| subprocess-based | Shell access, security concerns | No | Risky |

This library parses Office binary formats (OLE2) and XML-based formats (OOXML) directly in Python, making it ideal for:

- **RAG pipelines** and LLM document ingestion
- **Serverless functions** (AWS Lambda, Google Cloud Functions)
- **Containerized deployments** with minimal footprint
- **Secure environments** where shell execution is restricted
- **Cross-platform** applications (Windows, macOS, Linux)

### Enterprise SharePoint Reality

Enterprise SharePoints contain decades of accumulated documents. While modern `.docx`, `.xlsx`, and `.pptx` files are well-supported, legacy `.doc`, `.xls`, and `.ppt` files remain common. This library provides a **unified interface** for all formats—no conditional logic needed.

## Supported Formats

| Format            | Extension | Description                      |
|-------------------|-----------|----------------------------------|
| Modern Word       | `.docx`   | Word 2007+ documents             |
| Legacy Word       | `.doc`    | Word 97-2003 documents           |
| Modern Excel      | `.xlsx`   | Excel 2007+ spreadsheets         |
| Legacy Excel      | `.xls`    | Excel 97-2003 spreadsheets       |
| Modern PowerPoint | `.pptx`   | PowerPoint 2007+ presentations   |
| Legacy PowerPoint | `.ppt`    | PowerPoint 97-2003 presentations |
| PDF               | `.pdf`    | PDF documents                    |
| Plain Text        | `.txt`    | Plain text files                 |
| CSV               | `.csv`    | Comma-separated values           |
| TSV               | `.tsv`    | Tab-separated values             |
| JSON              | `.json`   | JSON files                       |

## Installation

```bash
pip install sharepoint-to-text
```

Or install from source:

```bash
git clone https://github.com/Horsmann/sharepoint-to-text.git
cd sharepoint-to-text
pip install -e .
```

## Quick Start

### The Unified Interface

All extractors return objects implementing a common interface:

```python
import sharepoint2text

# Works identically for ANY supported format
result = sharepoint2text.read_file("document.docx")  # or .doc, .pdf, .pptx, etc.

# Three methods available on ALL content types:
text = result.get_full_text()       # Complete text as a single string
metadata = result.get_metadata()    # File metadata (author, dates, etc.)

# Iterate over logical units (varies by format - see below)
for unit in result.iterator():
    print(unit)
```

### Understanding `iterator()` Output by Format

Different file formats have different natural structural units:

| Format | `iterator()` yields | Notes |
|--------|-------------------|-------|
| `.docx`, `.doc` | 1 item (full text) | Word documents have no page structure in the file format |
| `.xlsx`, `.xls` | 1 item per **sheet** | Each yield contains sheet content |
| `.pptx`, `.ppt` | 1 item per **slide** | Each yield contains slide text |
| `.pdf` | 1 item per **page** | Each yield contains page text |
| `.txt`, `.csv`, `.json`, `.tsv` | 1 item (full content) | Single unit |

**Note on Word documents:** The `.doc` and `.docx` file formats do not store page boundaries—pages are a rendering artifact determined by fonts, margins, and printer settings. The library returns the full document as a single text unit.

### Basic Usage Examples

```python
import sharepoint2text

# Extract from any file - format auto-detected
result = sharepoint2text.read_file("quarterly_report.docx")
print(result.get_full_text())

# Check format support before processing
if sharepoint2text.is_supported_file("document.xyz"):
    result = sharepoint2text.read_file("document.xyz")

# Access metadata
result = sharepoint2text.read_file("presentation.pptx")
meta = result.get_metadata()
print(f"Author: {meta.author}, Modified: {meta.modified}")
print(meta.to_dict())  # Convert to dictionary
```

### Working with Structured Content

```python
import sharepoint2text

# Excel: iterate over sheets
result = sharepoint2text.read_file("budget.xlsx")
for sheet in result.sheets:
    print(f"Sheet: {sheet.name}")
    print(f"Rows: {len(sheet.data)}")  # List of row dictionaries
    print(sheet.text)                   # Text representation

# PowerPoint: iterate over slides
result = sharepoint2text.read_file("deck.pptx")
for slide in result.slides:
    print(f"Slide {slide.slide_number}: {slide.title}")
    print(slide.content_placeholders)  # Body text
    print(slide.images)                # Image metadata

# PDF: iterate over pages
result = sharepoint2text.read_file("report.pdf")
for page_num, page in result.pages.items():
    print(f"Page {page_num}: {page.text[:100]}...")
    print(f"Images: {len(page.images)}")
```

### Using Format-Specific Extractors with BytesIO

For API responses or in-memory data:

```python
import sharepoint2text
import io

# Direct extractor usage with BytesIO
with open("document.docx", "rb") as f:
    result = sharepoint2text.read_docx(io.BytesIO(f.read()), path="document.docx")

# Get extractor dynamically based on filename
def extract_from_api(filename: str, content: bytes):
    extractor = sharepoint2text.get_extractor(filename)
    return extractor(io.BytesIO(content), path=filename)

result = extract_from_api("report.pdf", pdf_bytes)
```

## API Reference

### Main Functions

```python
import sharepoint2text

# Read any supported file (recommended entry point)
result = sharepoint2text.read_file(path: str | Path) -> ContentType

# Check if a file extension is supported
supported = sharepoint2text.is_supported_file(path: str) -> bool

# Get extractor function for a file type
extractor = sharepoint2text.get_extractor(path: str) -> Callable[[io.BytesIO, str | None], ContentType]
```

### Format-Specific Extractors

All accept `io.BytesIO` and optional `path` for metadata population:

```python
sharepoint2text.read_docx(file: io.BytesIO, path: str | None = None) -> DocxContent
sharepoint2text.read_doc(file: io.BytesIO, path: str | None = None) -> DocContent
sharepoint2text.read_xlsx(file: io.BytesIO, path: str | None = None) -> XlsxContent
sharepoint2text.read_xls(file: io.BytesIO, path: str | None = None) -> XlsContent
sharepoint2text.read_pptx(file: io.BytesIO, path: str | None = None) -> PptxContent
sharepoint2text.read_ppt(file: io.BytesIO, path: str | None = None) -> PptContent
sharepoint2text.read_pdf(file: io.BytesIO, path: str | None = None) -> PdfContent
sharepoint2text.read_plain_text(file: io.BytesIO, path: str | None = None) -> PlainTextContent
```

### Return Types

All content types implement the common interface:

```python
class ExtractionInterface(Protocol):
    def iterator() -> Iterator[str]           # Iterate over logical units
    def get_full_text() -> str                # Complete text as string
    def get_metadata() -> FileMetadataInterface  # Metadata with to_dict()
```

#### DocxContent (.docx)

```python
result.metadata       # DocxMetadata (title, author, created, modified, ...)
result.paragraphs     # List[DocxParagraph] (text, style, runs with formatting)
result.tables         # List[List[List[str]]] (cell data)
result.images         # List[DocxImage] (filename, content_type, data, size_bytes)
result.headers        # List[DocxHeaderFooter]
result.footers        # List[DocxHeaderFooter]
result.hyperlinks     # List[DocxHyperlink] (text, url)
result.footnotes      # List[DocxNote] (id, text)
result.endnotes       # List[DocxNote]
result.comments       # List[DocxComment] (author, date, text)
result.sections       # List[DocxSection] (page dimensions, margins)
result.full_text      # str (pre-computed full text)
```

#### DocContent (.doc)

```python
result.metadata         # DocMetadata (title, author, num_pages, num_words, num_chars, ...)
result.main_text        # str (main document body)
result.footnotes        # str (concatenated footnotes)
result.headers_footers  # str (concatenated headers/footers)
result.annotations      # str (concatenated annotations)
```

#### XlsxContent / XlsContent (.xlsx, .xls)

```python
result.metadata   # XlsxMetadata / XlsMetadata (title, creator, created, modified, ...)
result.sheets     # List[XlsxSheet / XlsSheet]

# Each sheet:
sheet.name   # str (sheet name)
sheet.data   # List[Dict[str, Any]] (rows as dictionaries)
sheet.text   # str (text representation)
```

#### PptxContent (.pptx)

```python
result.metadata   # PptxMetadata (title, author, created, modified, ...)
result.slides     # List[PPTXSlide]

# Each slide:
slide.slide_number          # int (1-indexed)
slide.title                 # str
slide.footer                # str
slide.content_placeholders  # List[str] (body content)
slide.other_textboxes       # List[str] (free-form text)
slide.images                # List[PPTXImage] (filename, content_type, size_bytes, blob)
slide.text                  # str (pre-computed combined text)
```

#### PptContent (.ppt)

```python
result.metadata   # PptMetadata (title, author, num_slides, created, modified, ...)
result.slides     # List[PptSlideContent]
result.all_text   # List[str] (flat list of all text)

# Each slide:
slide.slide_number   # int (1-indexed)
slide.title          # str | None
slide.body_text      # List[str]
slide.other_text     # List[str]
slide.notes          # List[str] (speaker notes)
slide.text_combined  # str (property: title + body + other)
slide.all_text       # List[PptTextBlock] (with text_type info)
```

#### PdfContent (.pdf)

```python
result.metadata    # PdfMetadata (total_pages)
result.pages       # Dict[int, PdfPage] (1-indexed)

# Each page:
page.text    # str
page.images  # List[PdfImage] (index, name, width, height, data, format)
```

#### PlainTextContent (.txt, .csv, .json, .tsv)

```python
result.content   # str (full file content)
result.metadata  # FileMetadataInterface (filename, file_extension, file_path, folder_path)
```

## Examples

### Bulk Processing

```python
import sharepoint2text
from pathlib import Path

def extract_all_documents(folder: Path) -> dict[str, str]:
    """Extract text from all supported files in a folder."""
    results = {}

    for file_path in folder.rglob("*"):
        if sharepoint2text.is_supported_file(str(file_path)):
            try:
                result = sharepoint2text.read_file(file_path)
                results[str(file_path)] = result.get_full_text()
            except Exception as e:
                print(f"Failed to extract {file_path}: {e}")

    return results
```

### Extract Images

```python
import sharepoint2text
import io

# From PDF
result = sharepoint2text.read_file("document.pdf")
for page_num, page in result.pages.items():
    for img in page.images:
        with open(f"page{page_num}_{img.name}.{img.format}", "wb") as out:
            out.write(img.data)

# From PowerPoint
result = sharepoint2text.read_file("slides.pptx")
for slide in result.slides:
    for img in slide.images:
        with open(img.filename, "wb") as out:
            out.write(img.blob)

# From Word
result = sharepoint2text.read_file("document.docx")
for img in result.images:
    if img.data:
        with open(img.filename, "wb") as out:
            out.write(img.data.getvalue())
```

### RAG Pipeline Integration

```python
import sharepoint2text

def prepare_for_rag(file_path: str) -> list[dict]:
    """Prepare document chunks for RAG ingestion."""
    result = sharepoint2text.read_file(file_path)
    meta = result.get_metadata()

    chunks = []
    for i, unit in enumerate(result.iterator()):
        if unit.strip():  # Skip empty units
            chunks.append({
                "text": unit,
                "metadata": {
                    "source": file_path,
                    "chunk_index": i,
                    "author": getattr(meta, "author", None),
                    "title": getattr(meta, "title", None),
                }
            })
    return chunks
```

## Requirements

- Python >= 3.10
- olefile >= 0.47
- openpyxl >= 3.1.5
- pandas >= 2.3.3
- pypdf >= 6.5.0
- python-docx >= 1.2.0
- python-pptx >= 1.0.2
- python-calamine >= 0.6.1

## License

Apache 2.0 - see [LICENSE](LICENSE) for details.
