class ExtractionFileFormatNotSupportedError(Exception):
    """Raised when the file format for extraction is not supported."""

    def __init__(self, file_path: str, message: str = None, *, cause: Exception = None):
        self.file_path = file_path
        if message is None:
            message = f"Extraction file format not supported: {file_path}"
        # Use exception chaining if cause is provided
        super().__init__(message)
        self.__cause__ = cause  # Optional chaining for debugging
