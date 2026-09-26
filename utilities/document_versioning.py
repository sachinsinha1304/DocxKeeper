"""
document_versioning.py

Adds version history + audit tracking on top of the existing filesystem-based
document store (repos/subfolders/*.docx). Metadata is stored in a MySQL table,
while physical snapshots of old versions are kept in a local .versions/ directory.

Layout on disk:
  /documents
    /repo_name
      /subfolder
        report.docx                 <- current/live file, unchanged location
        .versions/
          report.docx/
            3f9a1b2c-....docx       <- snapshot of an old version
            7d2e4f11-....docx
"""

import difflib
import hashlib
import json
import shutil
import uuid
from datetime import datetime, timezone
from pathlib import Path
from db.db import get_db_connection

from .docxDisplayHelper import readDocxContent  # your existing module; adjust import path

VERSIONS_DIRNAME = ".versions"


# ---------------------------------------------------------------------------
# Paths (Filesystem for snapshots only)
# ---------------------------------------------------------------------------

def _versions_dir(file_path: Path) -> Path:
    """-versions/<filename>/ sits alongside the live file for binary snapshots."""
    file_path = Path(file_path)
    return file_path.parent / VERSIONS_DIRNAME / file_path.name


# ---------------------------------------------------------------------------
# Metadata read/write (MySQL Implementation)
# ---------------------------------------------------------------------------

def load_metadata(file_path: Path) -> list:
    path_str = str(Path(file_path).resolve())
    conn = get_db_connection()
    try:
        with conn.cursor(dictionary=True) as cursor:
            query = """
                SELECT id, version_number, file_name, is_live, file_hash, 
                       modified_by, modified_at, change_type, change_summary, diff_json
                FROM document_versions
                WHERE file_path = %s
                ORDER BY version_number ASC
            """
            cursor.execute(query, (path_str,))
            rows = cursor.fetchall()
            
            versions = []
            for row in rows:
                # Reconstruct the dictionary format expected by the app
                diff_data = row["diff_json"]
                if isinstance(diff_data, str):
                    diff_data = json.loads(diff_data)
                
                versions.append({
                    "id": row["id"],
                    "version_number": row["version_number"],
                    "file_name": row["file_name"],
                    "is_live": bool(row["is_live"]),
                    "file_hash": row["file_hash"],
                    "modified_by": row["modified_by"],
                    "modified_at": row["modified_at"],
                    "change_type": row["change_type"],
                    "change_summary": row["change_summary"],
                    "diff": diff_data
                })
            return versions
    finally:
        conn.close()


def save_metadata(file_path: Path, versions: list) -> None:
    path_str = str(Path(file_path).resolve())
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            # Sync entire version list for this file using an UPSERT pattern
            for v in versions:
                diff_str = json.dumps(v.get("diff")) if v.get("diff") else None
                sql = """
                    INSERT INTO document_versions 
                    (id, file_path, version_number, file_name, is_live, file_hash, 
                     modified_by, modified_at, change_type, change_summary, diff_json)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON DUPLICATE KEY UPDATE 
                    is_live = VALUES(is_live),
                    change_summary = VALUES(change_summary)
                """
                cursor.execute(sql, (
                    v["id"],
                    path_str,
                    v["version_number"],
                    v["file_name"],
                    1 if v["is_live"] else 0,
                    v["file_hash"],
                    v["modified_by"],
                    v["modified_at"],
                    v["change_type"],
                    v["change_summary"],
                    diff_str
                ))
        conn.commit()
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Hashing / text extraction helpers
# ---------------------------------------------------------------------------

def _file_hash(file_path: Path) -> str:
    h = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _extract_plain_text(file_path: Path) -> str:
    """Flattens readDocxContent()'s block list into plain text lines,
    used only for computing a human-readable diff — not for storage."""
    blocks = readDocxContent(file_path.parent, file_path.name)
    lines = []
    for block in blocks:
        btype = block.get("type")
        if btype in ("paragraph", "heading"):
            lines.append(block.get("text", ""))
        elif btype == "table":
            for row in block.get("rows", []):
                lines.append(" | ".join(row))
        elif btype == "image":
            lines.append("[image]")
    return "\n".join(lines)


def _diff_summary(old_text: str, new_text: str, max_lines: int = 40) -> dict:
    """Line-level diff between two text extractions. Returns added/removed
    line counts plus a filtered preview containing ONLY added/removed lines."""
    old_lines = old_text.splitlines()
    new_lines = new_text.splitlines()
    diff = list(difflib.unified_diff(old_lines, new_lines, lineterm=""))

    added = sum(1 for l in diff if l.startswith("+") and not l.startswith("+++"))
    removed = sum(1 for l in diff if l.startswith("-") and not l.startswith("---"))

    # Keep ONLY lines that were explicitly added (+) or removed (-)
    filtered_preview = [
        l for l in diff 
        if (l.startswith("+") and not l.startswith("+++")) or 
           (l.startswith("-") and not l.startswith("---"))
    ]

    return {
        "lines_added": added,
        "lines_removed": removed,
        "preview": filtered_preview[:max_lines],
        "truncated": len(filtered_preview) > max_lines,
    }
# ---------------------------------------------------------------------------
# Core operations
# ---------------------------------------------------------------------------

def save_new_version(
    file_path,
    new_file_bytes: bytes,
    modified_by: str,
    change_summary: str = None,
) -> dict:
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"{file_path} does not exist")

    versions = load_metadata(file_path)
    new_hash = hashlib.sha256(new_file_bytes).hexdigest()

    if not versions:
        versions.append(_write_version_record(
            file_path, file_path.read_bytes(), modified_by, "Initial version", "created",
            next_number=1,
        ))
        if new_hash == versions[-1]["file_hash"]:
            save_metadata(file_path, versions)
            return versions[-1]

    current_hash = versions[-1]["file_hash"]
    if new_hash == current_hash:
        return versions[-1]

    old_text = _extract_plain_text(file_path)  # live file, pre-overwrite
    record = _write_version_record(
        file_path, new_file_bytes, modified_by, change_summary, "replaced", old_text,
        next_number=versions[-1]["version_number"] + 1,
    )
    for v in versions:
        v["is_live"] = False
    versions.append(record)
    save_metadata(file_path, versions)
    return record


def _write_version_record(file_path, new_bytes, modified_by, change_summary, change_type,
                          old_text="", next_number=1):
    file_path = Path(file_path)
    vdir = _versions_dir(file_path)
    vdir.mkdir(parents=True, exist_ok=True)

    with open(file_path, "wb") as f:
        f.write(new_bytes)

    snapshot_id = str(uuid.uuid4())
    snapshot_path = vdir / f"{snapshot_id}{file_path.suffix}"
    shutil.copy2(file_path, snapshot_path)

    new_text = _extract_plain_text(file_path)
    diff = _diff_summary(old_text, new_text) if old_text else None

    return {
        "id": snapshot_id,
        "version_number": next_number,
        "file_name": snapshot_path.name,
        "is_live": True,
        "file_hash": hashlib.sha256(new_bytes).hexdigest(),
        "modified_by": modified_by,
        "modified_at": datetime.now(timezone.utc).isoformat(),
        "change_type": change_type,
        "change_summary": change_summary,
        "diff": diff,
    }


def list_versions(file_path) -> list:
    file_path = Path(file_path)
    versions = load_metadata(file_path)
    
    if not versions and file_path.exists():
        try:
            record = _write_version_record(
                file_path=file_path,
                new_bytes=file_path.read_bytes(),
                modified_by="System",
                change_summary="Auto-seeded initial version",
                change_type="created",
                old_text="",
                next_number=1
            )
            versions = [record]
            save_metadata(file_path, versions)
        except Exception:
            pass
            
    return versions


def restore_version(file_path, version_id: str, restored_by: str) -> dict:
    file_path = Path(file_path)
    versions = load_metadata(file_path)
    vdir = _versions_dir(file_path)

    target = next((v for v in versions if v["id"] == version_id), None)
    if target is None:
        raise ValueError(f"No version {version_id} found for {file_path}")

    if target.get("is_live"):
        raise ValueError("That version is already the live version")

    snapshot_path = vdir / target["file_name"]
    if not snapshot_path.exists():
        raise FileNotFoundError(f"Snapshot file missing: {snapshot_path}")

    with open(snapshot_path, "rb") as f:
        restored_bytes = f.read()

    return save_new_version(
        file_path,
        restored_bytes,
        modified_by=restored_by,
        change_summary=f"Restored from version {target['version_number']}",
    )


def get_audit_trail(file_path) -> list:
    versions = list_versions(file_path)
    trail = []
    for v in reversed(versions):
        diff = v.get("diff") or {}
        trail.append({
            "version": v["version_number"],
            "by": v["modified_by"],
            "at": v["modified_at"],
            "type": v["change_type"],
            "summary": v.get("change_summary"),
            "lines_added": diff.get("lines_added"),
            "lines_removed": diff.get("lines_removed"),
        })
    return trail