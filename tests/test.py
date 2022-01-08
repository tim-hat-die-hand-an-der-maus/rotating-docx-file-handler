import logging
import os
import unittest

from src.rotatingdocxfilehandler import handler


def createLoggerWithDocxHandler(name: str, *, level: int = logging.DEBUG):
    logger = logging.Logger(name)
    ch = handler.RotatingDocxFileHandler("test")

    formatting = "[{}] %(asctime)s\t%(levelname)s\t%(module)s.%(funcName)s#%(lineno)d | %(message)s".format(name)
    formatter = logging.Formatter(formatting)
    ch.setFormatter(formatter)

    logger.addHandler(ch)
    logger.setLevel(level)

    return logger


class RotatingDocxFileHandlerTest(unittest.TestCase):
    def test_non_existing_file_without_file_ending(self):
        filename = "./test"
        if os.path.exists(filename):
            os.remove(filename)

        logger = createLoggerWithDocxHandler(filename, level=logging.DEBUG)
        logger.error("error")
        logger.warning("warning")
        logger.info("info")
        logger.debug("debug")

        # TODO: actually read contents from file
        # change formatter for this to `%(message)s`
        self.assertEqual(True, True)
