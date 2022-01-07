from logging.handlers import RotatingFileHandler

import docx


class RotatingDocxFileHandler(RotatingFileHandler):
    def __init__(self, filename: str, mode='a', maxBytes=0, backupCount=0, encoding=None, delay=False, errors=None):
        if not filename.endswith(".docx"):
            filename = filename + ".docx"

        super().__init__(filename, mode=mode, maxBytes=maxBytes, backupCount=backupCount, encoding=encoding,
                         delay=delay)

    def _open(self):
        # noinspection PyUnresolvedReferences
        try:
            self.document = docx.Document(self.baseFilename)
        except docx.opc.exceptions.PackageNotFoundError:
            self.document = docx.Document()

        self.document.save(self.baseFilename)

    def close(self):
        super().close()
