import logging
from datetime import datetime

from core.paths import logs_dir
from core.runtime_mode import is_engine_mode


def setup_logger():
    log_dir = logs_dir()

    filename = log_dir / f"run_{datetime.now():%Y%m%d_%H%M%S}.log"

    logger = logging.getLogger("nc_auto")
    logger.setLevel(logging.DEBUG)

    fh = logging.FileHandler(filename, encoding="utf-8")
    fh.setLevel(logging.DEBUG)

    # 文件始终留全量 DEBUG(开发细节);控制台/stdout 在引擎模式下只冒 WARNING+,
    # 避免 INFO 噪音经子进程管道污染 finance 用户面板。CLI 直跑保持 INFO 可见。
    ch = logging.StreamHandler()
    ch.setLevel(logging.WARNING if is_engine_mode() else logging.INFO)

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    fh.setFormatter(fmt)
    ch.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(ch)

    return logger


log = setup_logger()
