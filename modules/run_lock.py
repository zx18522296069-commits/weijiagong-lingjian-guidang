"""运行锁，防止 GitHub Actions 重复执行。"""

from pathlib import Path

LOCK_FILE = Path(".update_parts.lock")


class RunLock:
    def acquire(self):
        if LOCK_FILE.exists():
            raise RuntimeError("已有任务正在执行，停止本次运行")
        LOCK_FILE.write_text("running", encoding="utf-8")

    def release(self):
        if LOCK_FILE.exists():
            LOCK_FILE.unlink()

    def __enter__(self):
        self.acquire()
        return self

    def __exit__(self, exc_type, exc, tb):
        self.release()
