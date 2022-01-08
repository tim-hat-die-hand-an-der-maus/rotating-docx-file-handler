from logging.handlers import RotatingFileHandler
from typing import Optional

import docx

from . import HandlerDocument


class RotatingDocxFileHandler(RotatingFileHandler):
    def __init__(self, filename: str, mode='a', maxBytes=0, backupCount=0, encoding=None, delay=False, errors=None):
        if not filename.endswith(".docx"):
            filename = filename + ".docx"

        super().__init__(filename, mode=mode, maxBytes=maxBytes, backupCount=backupCount, encoding=encoding,
                         delay=delay)
        self.stream: Optional[HandlerDocument] = None

    def _open(self):
        self.stream = HandlerDocument(self.baseFilename)

        self.stream.save(self.baseFilename)

        return self.stream

    def close(self):
        super().close()

        self.stream.save(self.baseFilename)
