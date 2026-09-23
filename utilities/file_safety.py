import os
import time
import json
import hashlib
import tempfile
import portalocker
from pathlib import Path
from contextlib import contextmanager

# A dedicated, non-browsable directory for all lock/version sidecar files.
# Adjust this base path to wherever your app root actually is.
LOCKS_ROOT = Path(__file__).resolve().parent.parent / ".locks"
LOCKS_ROOT.mkdir(exist_ok=True)


def _lock_key(target_path: Path) -> str:
    """Stable, collision-safe key derived from the file's absolute path."""
    abs_str = str(Path(target_path).resolve())
    return hashlib.sha256(abs_str.encode("utf-8")).hexdigest()


def _lock_path(target_path: Path) -> Path:
    return LOCKS_ROOT / f"{_lock_key(target_path)}.lock"


def _editing_lock_path(target_path: Path) -> Path:
    return LOCKS_ROOT / f"{_lock_key(target_path)}.editing.json"


@contextmanager
def exclusive_write_lock(target_path: Path, timeout=10):
    lock_file = _lock_path(target_path)
    lock_file.touch(exist_ok=True)
    try:
        with portalocker.Lock(str(lock_file), timeout=timeout, mode="w") as fh:
            yield fh
    finally:
        # Optional: Clean up the physical lock file after release
        try:
            lock_file.unlink(missing_ok=True)
        except OSError:
            pass


def atomic_write_docx(target_path: Path, save_fn):
    target_path = Path(target_path)
    tmp_fd, tmp_name = tempfile.mkstemp(
        suffix=".tmp", prefix=target_path.stem + "_", dir=target_path.parent
    )
    os.close(tmp_fd)
    tmp_path = Path(tmp_name)
    try:
        save_fn(tmp_path)
        os.replace(tmp_path, target_path)
    finally:
        if tmp_path.exists():
            tmp_path.unlink(missing_ok=True)


def get_version_stamp(target_path: Path) -> str:
    stat = Path(target_path).stat()
    return f"{stat.st_mtime_ns}-{stat.st_size}"


def acquire_editing_lock(target_path: Path, user_id: str, ttl_seconds=300):
    lock_path = _editing_lock_path(target_path)
    now = time.time()

    if lock_path.exists():
        try:
            holder = json.loads(lock_path.read_text())
        except (json.JSONDecodeError, OSError):
            holder = None

        if holder and now - holder.get("timestamp", 0) < ttl_seconds and holder.get("user_id") != user_id:
            return False, holder

    lock_path.write_text(json.dumps({"user_id": user_id, "timestamp": now}))
    return True, None


def release_editing_lock(target_path: Path, user_id: str):
    lock_path = _editing_lock_path(target_path)
    if lock_path.exists():
        try:
            holder = json.loads(lock_path.read_text())
            if holder.get("user_id") == user_id:
                lock_path.unlink(missing_ok=True)
        except (json.JSONDecodeError, OSError):
            lock_path.unlink(missing_ok=True)


def refresh_editing_lock(target_path: Path, user_id: str):
    lock_path = _editing_lock_path(target_path)
    lock_path.write_text(json.dumps({"user_id": user_id, "timestamp": time.time()}))