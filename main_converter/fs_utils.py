from __future__ import annotations

import os
import shutil
import stat
import time
from pathlib import Path
from typing import Optional


def _clear_readonly_and_retry(func, path, exc_info) -> None:
    exc = exc_info[1]
    if not isinstance(exc, PermissionError):
        raise exc
    os.chmod(path, stat.S_IWRITE)
    func(path)


def remove_tree_robust(path: Path, retries: int = 3, delay_sec: float = 0.2) -> None:
    last_error: Optional[Exception] = None
    for attempt in range(retries):
        try:
            shutil.rmtree(path, onerror=_clear_readonly_and_retry)
            return
        except FileNotFoundError:
            return
        except PermissionError as exc:
            last_error = exc
            if attempt + 1 == retries:
                break
            time.sleep(delay_sec)
    if last_error is not None:
        raise last_error
