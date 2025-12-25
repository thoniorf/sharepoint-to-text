import io
import logging

logger = logging.getLogger(__name__)


def read_plain_text(file_like: io.BytesIO) -> dict:
    logger.debug("Reading plain text file")
    file_like.seek(0)

    content = file_like.read()

    if isinstance(content, bytes):
        text = content.decode("utf-8", errors="ignore")
    else:
        text = content

    return {"content": text}
