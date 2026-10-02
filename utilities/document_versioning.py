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
import os
import uuid
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.table import _Cell
from docx.text.paragraph import Paragraph

from db.db import get_db_connection
from .docxDisplayHelper import readDocxContent  # your existing module; adjust import path

VERSIONS_DIRNAME = ".versions"


# ---------------------------------------------------------------------------
# Paths (Filesystem for snapshots only)
# ---------------------------------------------------------------------------

def _versions_dir(file_path: Path) -> Path:
    """.versions/<filename>/ sits alongside the live file for binary snapshots."""
    file_path = Path(file_path)
    return file_path.parent / VERSIONS_DIRNAME / file_path.name


def _atomic_write(path: Path, data: bytes) -> None:
    """Write to a temp file in the same folder, fsync, then atomically swap it in,
    so a crash mid-write can never leave a truncated/corrupt .docx."""
    path = Path(path)
    tmp = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        with open(tmp, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    finally:
        if tmp.exists():
            tmp.unlink()


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
                    "diff": diff_data,
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
                    diff_str,
                ))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Content extraction (for human-readable diffs)
#
# Produces an ordered list of (line, format_signature) tuples covering:
#   headings, paragraphs, list items, table rows/cells, images (by content hash,
#   so replacing a chart/diagram shows up), page breaks and Word comments.
# The signature captures run formatting (font/size/colour/highlight/bold/...),
# alignment and cell shading, so formatting-only edits are detected too.
# ---------------------------------------------------------------------------

def _run_fmt(run) -> tuple:
    f = run.font
    try:
        rgb = str(f.color.rgb) if f.color.rgb is not None else None
    except Exception:
        rgb = None
    try:
        hl = getattr(f.highlight_color, "name", None)
    except Exception:
        hl = None
    shd = None
    rPr = run._r.rPr
    if rPr is not None:
        el = rPr.find(qn("w:shd"))
        if el is not None:
            shd = el.get(qn("w:fill"))
    return (
        run.bold or None, run.italic or None, bool(run.underline) or None,
        f.strike or None, f.subscript or None, f.superscript or None,
        f.name, f.size.pt if f.size else None, rgb, hl, shd,
    )


def _para_sig(p) -> str:
    """Formatting signature that is independent of how text happens to be split into runs."""
    spans = []                                   # [fmt, char_count]
    for r in p.runs:
        if not r.text:
            continue
        fmt = _run_fmt(r)
        if spans and spans[-1][0] == fmt:
            spans[-1][1] += len(r.text)
        else:
            spans.append([fmt, len(r.text)])
    style = getattr(p.style, "name", "") or ""
    raw = repr((style, str(p.alignment), spans))
    return hashlib.md5(raw.encode()).hexdigest()[:12]


def _cell_fill(tc) -> str:
    tcPr = tc.tcPr
    if tcPr is None:
        return ""
    shd = tcPr.find(qn("w:shd"))
    return (shd.get(qn("w:fill")) or "") if shd is not None else ""


def _walk(parent_elm, parent, doc, out, prefix=""):
    for child in parent_elm.iterchildren():
        if child.tag == qn("w:p"):
            p = Paragraph(child, parent)
            style = (getattr(p.style, "name", "") or "")
            text = p.text.strip()
            if text:
                m = style.lower()
                if m.startswith("heading"):
                    digits = "".join(c for c in m if c.isdigit()) or "1"
                    text = "#" * int(digits) + " " + text
                elif m.startswith("list"):
                    text = "• " + text
                out.append((prefix + text, _para_sig(p)))
            for blip in p._p.iter(qn("a:blip")):
                try:
                    blob = doc.part.rels[blip.get(qn("r:embed"))].target_part.blob
                    out.append((f"{prefix}[image {hashlib.sha1(blob).hexdigest()[:8]}]", ""))
                except Exception:
                    out.append((f"{prefix}[image]", ""))
            if p._p.xpath('.//w:br[@w:type="page"]'):
                out.append((f"{prefix}[page break]", ""))
        elif child.tag == qn("w:tbl"):
            for tr in child.iterchildren(qn("w:tr")):
                fills = []
                for tc in tr.iterchildren(qn("w:tc")):
                    fills.append(_cell_fill(tc))
                    _walk(tc, _Cell(tc, parent), doc, out, prefix="| ")
                out.append(("|--", "|".join(fills)))     # row boundary + cell shading (callout colour)


def _extract_structured(file_path: Path) -> list:
    doc = Document(file_path)
    out = []
    _walk(doc.element.body, doc, doc, out)
    try:                                              # read word/comments.xml directly (any python-docx version)
        root = etree.fromstring(doc.part.part_related_by(RT.COMMENTS).blob)
        items = sorted(root.findall(qn("w:comment")), key=lambda c: int(c.get(qn("w:id")) or 0))
        for c in items:
            text = " ".join("".join(t.text or "" for t in p.iter(qn("w:t"))) for p in c.iter(qn("w:p"))).strip()
            out.append((f"[comment] {c.get(qn('w:author')) or 'Unknown'}: {text}", ""))
    except Exception:
        pass                                          # no comments part => nothing to add
    return out


def _safe_extract(file_path: Path):
    """Never let a diff failure block a save."""
    try:
        return _extract_structured(Path(file_path))
    except Exception:
        return None


def _extract_plain_text(file_path: Path) -> str:
    """Kept for backwards compatibility."""
    return "\n".join(line for line, _ in (_safe_extract(file_path) or []))


def _diff_summary(old: list, new: list, max_lines: int = 40) -> dict:
    """Line-level diff plus formatting-only changes.
    preview contains '+' (added), '-' (removed) and '~' (formatting changed) lines."""
    old_lines = [l for l, _ in old]
    new_lines = [l for l, _ in new]

    diff = list(difflib.unified_diff(old_lines, new_lines, lineterm=""))[2:]   # drop ---/+++ headers
    added_l = [l for l in diff if l.startswith("+")]
    removed_l = [l for l in diff if l.startswith("-")]

    fmt_changed = []
    sm = difflib.SequenceMatcher(a=old_lines, b=new_lines, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                if old[i1 + k][1] != new[j1 + k][1]:
                    fmt_changed.append("~" + new_lines[j1 + k][:120])

    preview = [l for l in diff if l.startswith(("+", "-"))] + fmt_changed
    return {
        "lines_added": len(added_l),
        "lines_removed": len(removed_l),
        "format_changes": len(fmt_changed),
        "preview": preview[:max_lines],
        "truncated": len(preview) > max_lines,
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
    old_bytes = file_path.read_bytes()

    if not versions:
        versions.append(_write_version_record(
            file_path, old_bytes, modified_by, "Initial version", "created",
            next_number=1, write_live=False,        # live file already holds these bytes
        ))
        if new_hash == versions[-1]["file_hash"]:
            save_metadata(file_path, versions)
            return versions[-1]

    if new_hash == versions[-1]["file_hash"]:
        return versions[-1]

    old_struct = _safe_extract(file_path)           # live file, pre-overwrite
    record = _write_version_record(
        file_path, new_file_bytes, modified_by, change_summary, "replaced",
        old_struct=old_struct, next_number=versions[-1]["version_number"] + 1,
    )
    for v in versions:
        v["is_live"] = False
    versions.append(record)

    try:
        save_metadata(file_path, versions)
    except Exception:
        # DB write failed: put the previous live file back and drop the orphan snapshot
        _atomic_write(file_path, old_bytes)
        try:
            (_versions_dir(file_path) / record["file_name"]).unlink()
        except OSError:
            pass
        raise
    return record


def _write_version_record(file_path, new_bytes, modified_by, change_summary, change_type,
                          old_struct=None, next_number=1, write_live=True):
    file_path = Path(file_path)
    vdir = _versions_dir(file_path)
    vdir.mkdir(parents=True, exist_ok=True)

    snapshot_id = str(uuid.uuid4())
    snapshot_path = vdir / f"{snapshot_id}{file_path.suffix}"
    _atomic_write(snapshot_path, new_bytes)         # snapshot first, so a failure never touches the live file
    if write_live:
        _atomic_write(file_path, new_bytes)

    diff = None
    if old_struct is not None:
        new_struct = _safe_extract(snapshot_path)
        if new_struct is not None:
            diff = _diff_summary(old_struct, new_struct)

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
                next_number=1,
                write_live=False,
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

    return save_new_version(
        file_path,
        snapshot_path.read_bytes(),
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
            "format_changes": diff.get("format_changes"),
        })
    return trail


def get_version_path(target_path, version_id):
    target_path = Path(target_path)
    try:
        version_id = str(uuid.UUID(str(version_id)))     # blocks path traversal like "../../x"
    except ValueError:
        raise ValueError("Invalid version id")
    return _versions_dir(target_path) / f"{version_id}{target_path.suffix}"


def read_version_content(version_path):
    """
    Reads a version snapshot using the same DOCX reader
    used for normal documents.
    """
    version_path = Path(version_path)

    if not version_path.exists():
        raise FileNotFoundError(f"Version file not found: {version_path}")

    if version_path.suffix.lower() != ".docx":
        raise ValueError(f"Version is not a DOCX file: {version_path}")

    return readDocxContent(version_path.parent, version_path.name)