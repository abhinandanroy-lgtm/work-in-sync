import logging
import os
from datetime import datetime
from app.config import LOG_DIR

def get_logger(name="wisbot"):
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)

    file_path = os.path.join(LOG_DIR, f"run_{datetime.now().strftime('%Y%m%d')}.log")
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    fh = logging.FileHandler(file_path, encoding="utf-8")
    fh.setFormatter(formatter)

    sh = logging.StreamHandler()
    sh.setFormatter(formatter)

    logger.addHandler(fh)
    logger.addHandler(sh)
    return logger