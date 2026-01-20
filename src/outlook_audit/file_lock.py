import os
import psutil


def file_is_open_by_any_process(path: str) -> bool:
    abs_path = os.path.abspath(path)
    for proc in psutil.process_iter(["open_files"]):
        try:
            files = proc.info.get("open_files") or []
            for f in files:
                if getattr(f, "path", "") == abs_path:
                    return True
        except Exception:
            continue
    return False


def assert_files_closed(paths: list[str]) -> None:
    for p in paths:
        if p and os.path.exists(p) and file_is_open_by_any_process(p):
            raise RuntimeError(f'O arquivo está aberto e precisa ser fechado: "{p}"')
