"""
document_versioning.py

Adds version history + audit tracking on top of the existing filesystem-based
document store (repos/subfolders/*.docx). Designed to drop in next to your
existing utility files (repo_utils.py, docx_reader.py, etc.) with no DB
required — metadata is kept in a JSON sidecar file per document.

Layout on disk:

  /documents
    /repo_name
      /subfolder
        report.docx                <- current/live file, unchanged location
        .versions/
          report.docx/
            versions.json          <- ordered metadata for this file
            3f9a1b2c-....docx      <- snapshot of an old version
            7d2e4f11-....docx

Swap-in note: every function here takes/returns plain dicts, so if you later
move metadata into a real DB, only save_metadata()/load_metadata() need to
change — nothing else in this file or its callers.
"""

import difflib
import hashlib
import json
import shutil
import uuid
from datetime import datetime, timezone
from pathlib import Path

from .docxDisplayHelper import readDocxContent  # your existing module; adjust import path

VERSIONS_DIRNAME = ".versions"


# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

def _versions_dir(file_path: Path) -> Path:
    """.versions/<filename>/ sits alongside the live file, hidden by your
    existing HIDDEN_SUFFIXES-style filtering if you also exclude dotfolders."""
    file_path = Path(file_path)
    return file_path.parent / VERSIONS_DIRNAME / file_path.name


def _metadata_path(file_path: Path) -> Path:
    meta_path = _versions_dir(file_path) / "versions.json"
    print(f"DEBUG: Looking for metadata at -> {meta_path.resolve()}")
    return meta_path


# ---------------------------------------------------------------------------
# Metadata read/write
# ---------------------------------------------------------------------------

def load_metadata(file_path: Path) -> list:
    meta_path = _metadata_path(file_path)
    if not meta_path.exists():
        return []
    with open(meta_path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_metadata(file_path: Path, versions: list) -> None:
    meta_path = _metadata_path(file_path)
    meta_path.parent.mkdir(parents=True, exist_ok=True)
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(versions, f, indent=2, ensure_ascii=False)


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
    line counts plus a capped unified diff for display in a UI."""
    old_lines = old_text.splitlines()
    new_lines = new_text.splitlines()
    diff = list(difflib.unified_diff(old_lines, new_lines, lineterm=""))

    added = sum(1 for l in diff if l.startswith("+") and not l.startswith("+++"))
    removed = sum(1 for l in diff if l.startswith("-") and not l.startswith("---"))

    return {
        "lines_added": added,
        "lines_removed": removed,
        "preview": diff[:max_lines],
        "truncated": len(diff) > max_lines,
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
    """
    Call this whenever a user uploads a replacement for an existing .docx.

    1. Hashes the incoming bytes; if identical to the current live file, does
       nothing (no-op save) and returns the existing latest version's record.
    2. Otherwise: snapshots the CURRENT live file into .versions/ before
       overwriting it, computes a text diff for the "what changed" record,
       writes the new bytes to the live path, and appends a new version entry.

    Returns the metadata dict for the newly created version.
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"{file_path} does not exist")

    versions = load_metadata(file_path)
    new_hash = hashlib.sha256(new_file_bytes).hexdigest()

    # first-ever save: seed history with the file's current state as v1,
    # then fall through to record the incoming bytes as v2 if it differs.
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
        # nothing actually changed — don't spam version history
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
    """Writes new_bytes to the live file, snapshots the result into
    .versions/, and returns the metadata record for this version."""
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
    """Returns version history, newest last. Auto-seeds history if 
       the file exists but has no metadata ledger yet (legacy files)."""
    file_path = Path(file_path)
    versions = load_metadata(file_path)
    print(file_path)
    print(versions)
    
    if not versions and file_path.exists():
        try:
            # Auto-seed v1 for files created before versioning was integrated
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
            pass # Fail gracefully if read/write fails
            
    return versions


def restore_version(file_path, version_id: str, restored_by: str) -> dict:
    """
    Restores an old snapshot as the new live content. Does NOT delete any
    history — creates a fresh version entry with change_type='restored', so
    the timeline stays honest (you can see a restore happened and when).
    """
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
    """Flattened who/when/what list, newest first — convenient for a UI table."""
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