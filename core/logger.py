import logging
from pathlib import Path
from datetime import datetime


def setup_logger():
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)

    filename = log_dir / f"run_{datetime.now():%Y%m%d_%H%M%S}.log"

    logger = logging.getLogger("nc_auto")
    logger.setLevel(logging.DEBUG)

    fh = logging.FileHandler(filename, encoding="utf-8")
    fh.setLevel(logging.DEBUG)

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    fh.setFormatter(fmt)
    ch.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(ch)

    return logger


log = setup_logger()
