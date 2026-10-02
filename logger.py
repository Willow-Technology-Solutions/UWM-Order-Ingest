from logging.handlers import RotatingFileHandler
import logging
import os
import sys
from pathlib import Path


def _getBaseDirectory():
    """Return the directory that contains the bundled executable or script."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent


BASE_DIR = os.fspath(_getBaseDirectory())
LOGS_DIR = os.path.join(BASE_DIR, "logs")


def setupLogger():
    """Set up and return a configured logger with both stdout and file handlers."""
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

        stdoutHandler = logging.StreamHandler(sys.stdout)
        stdoutHandler.setLevel(logging.DEBUG)
        stdoutHandler.setFormatter(formatter)
        logger.addHandler(stdoutHandler)

        os.makedirs(LOGS_DIR, exist_ok=True)
        log_file = os.path.join(LOGS_DIR, "log.txt")
        logFileHandler = RotatingFileHandler(
            log_file,
            maxBytes=10000000,  # 10mb
            backupCount=3,
        )
        logFileHandler.setLevel(logging.DEBUG)
        logFileHandler.setFormatter(formatter)
        logger.addHandler(logFileHandler)

    return logger
