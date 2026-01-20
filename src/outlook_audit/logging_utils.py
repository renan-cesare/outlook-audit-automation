from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


@dataclass
class Logger:
    log_file: Path

    def info(self, msg: str) -> None:
        self._write("[INFO] " + msg)

    def ok(self, msg: str) -> None:
        self._write("[OK] " + msg)

    def warn(self, msg: str) -> None:
        self._write("[WARN] " + msg)

    def error(self, msg: str) -> None:
        self._write("[ERROR] " + msg)

    def _write(self, msg: str) -> None:
        line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} {msg}"
        print(line)
        self.log_file.parent.mkdir(parents=True, exist_ok=True)
        if self.log_file.exists():
            old = self.log_file.read_text(encoding="utf-8")
            self.log_file.write_text(old + line + "\n", encoding="utf-8")
        else:
            self.log_file.write_text(line + "\n", encoding="utf-8")


def make_logger() -> Logger:
    ts = datetime.now().strftime("%Y%m%d")
    return Logger(log_file=Path("logs") / f"exec_{ts}.log")
