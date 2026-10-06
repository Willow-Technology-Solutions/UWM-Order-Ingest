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


def _cleanEnvPath(value: str) -> str:
    text = value.strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in {'"', "'"}:
        text = text[1:-1].strip()
    return text


def resolveLogDirectory() -> str:
    """Return LOGGING_LOCATION when set, otherwise the local logs directory."""
    logging_location = _cleanEnvPath(os.getenv("LOGGING_LOCATION", ""))
    if logging_location:
        return logging_location
    return os.path.join(os.fspath(_getBaseDirectory()), "logs")


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

        log_directory = resolveLogDirectory()
        os.makedirs(log_directory, exist_ok=True)
        log_file = os.path.join(log_directory, "log.txt")
        logFileHandler = RotatingFileHandler(
            log_file,
            maxBytes=10000000,  # 10mb
            backupCount=3,
        )
        logFileHandler.setLevel(logging.DEBUG)
        logFileHandler.setFormatter(formatter)
        logger.addHandler(logFileHandler)
        logger.info("Writing logs to %s", log_file)

    return logger
