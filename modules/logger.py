"""执行日志模块：同时输出到控制台和 logs/ 文件。"""

from __future__ import annotations

import logging
import os
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

BEIJING_TZ = timezone(timedelta(hours=8))


class BeijingFormatter(logging.Formatter):
    def formatTime(self, record, datefmt=None):
        dt = datetime.fromtimestamp(record.created, BEIJING_TZ)
        if datefmt:
            return dt.strftime(datefmt)
        return dt.isoformat(timespec="seconds")


def get_logger(name: str = "weijiagong") -> logging.Logger:
    """返回单例 logger；日志会进入控制台与 logs/ 目录。"""
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    logger.propagate = False

    formatter = BeijingFormatter("[%(asctime)s] %(levelname)s %(message)s")

    stream = logging.StreamHandler(sys.stdout)
    stream.setFormatter(formatter)
    logger.addHandler(stream)

    log_dir = Path("logs")
    log_dir.mkdir(parents=True, exist_ok=True)
    run_id = os.getenv("GITHUB_RUN_ID", "local")
    stamp = datetime.now(BEIJING_TZ).strftime("%Y%m%d_%H%M%S")
    file_handler = logging.FileHandler(
        log_dir / f"run_{stamp}_{run_id}.log",
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger


def log(message: str) -> None:
    """兼容旧调用。"""
    get_logger().info(message)
