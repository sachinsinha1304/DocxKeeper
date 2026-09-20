import os
from pathlib import Path
import base64
import base64
import zipfile
from pathlib import Path
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
import json


path = r"./documents/"

def listAllTeamFolder():
    target_path = Path(path)
    folder_list = [folder_name for folder_name in os.listdir(target_path)]
    if folder_list is None or folder_list == []:
        return [], True
    return folder_list, False

def listAllSubfolderInRepo(repo_name):
    target_path = Path(path + repo_name).resolve()
    # target_path = (path / re/po_name)

    # Security check: ensure path stays inside path
    if not str(target_path).startswith(str(target_path.resolve())):
        raise FileNotFoundError(f"{target_path} not exist")

    if not target_path.exists():
        raise FileNotFoundError(f"{target_path} not exist")

    if target_path.is_file():
        return {"parent": target_path.parent, "name":target_path.name}, True

    contents = []
    for item in target_path.iterdir():
        contents.append({
            "name": item.name,
            "is_file": item.is_file(),
            "route_path": f"{repo_name}/{item.name}" if repo_name else item.name,
        })

    return contents, False


def readDocxContent(file_path, file_name):
    target_path = Path(file_path) / file_name
    content = []

    if target_path.suffix.lower() != ".docx":
        return content

    if not zipfile.is_zipfile(target_path):
        with open(target_path, "rb") as f:
            header = f.read(8)
        return [{"type": "error", "text": f"Not a valid docx (zip) file. Header bytes: {header!r}"}]

    try:
        doc = Document(target_path)

        image_rels = {
            rel_id: rel.target_part
            for rel_id, rel in doc.part.rels.items()
            if "image" in rel.reltype
        }

        def make_image_item(rel_id):
            image_part = image_rels[rel_id]
            mime = image_part.content_type
            b64_data = base64.b64encode(image_part.blob).decode("utf-8")
            return {"type": "image", "mime": mime, "src": f"data:{mime};base64,{b64_data}"}

        def iter_block_items(parent):
            parent_elm = parent.element.body
            for child in parent_elm.iterchildren():
                if child.tag == qn("w:p"):
                    yield Paragraph(child, parent)
                elif child.tag == qn("w:tbl"):
                    yield Table(child, parent)

        def paragraph_style_type(paragraph):
            style_name = (paragraph.style.name or "").lower()
            if "heading" in style_name:
                level = "".join(filter(str.isdigit, style_name)) or "1"
                return "heading", int(level)
            return "paragraph", None

        def process_paragraph(paragraph):
            """Walk the paragraph's XML in document order, emitting text and
            images in the exact sequence they appear (not images-first)."""
            items = []
            text_buffer = ""
            block_type, level = paragraph_style_type(paragraph)

            def flush_text():
                nonlocal text_buffer
                if text_buffer.strip():
                    if block_type == "heading":
                        items.append({"type": "heading", "level": level, "text": text_buffer.strip()})
                    else:
                        items.append({"type": "paragraph", "text": text_buffer.strip()})
                text_buffer = ""

            # iterate every descendant of the paragraph in document order
            for el in paragraph._element.iter():
                tag = el.tag
                if tag == qn("w:t"):
                    text_buffer += el.text or ""
                elif tag == qn("w:tab"):
                    text_buffer += "\t"
                elif tag == qn("w:br") or tag == qn("w:cr"):
                    text_buffer += "\n"
                elif tag == qn("a:blip"):
                    rel_id = el.get(qn("r:embed"))
                    if rel_id and rel_id in image_rels:
                        flush_text()          # emit text seen so far
                        items.append(make_image_item(rel_id))  # then the image

            flush_text()  # emit any trailing text after the last image
            return items

        for block in iter_block_items(doc):
            if isinstance(block, Paragraph):
                content.extend(process_paragraph(block))

            elif isinstance(block, Table):
                rows = [[cell.text.strip() for cell in row.cells] for row in block.rows]
                if rows:
                    content.append({"type": "table", "rows": rows})

    except Exception as e:
        content = [{"type": "error", "text": f"Error reading document: {str(e)}"}]

    return content
