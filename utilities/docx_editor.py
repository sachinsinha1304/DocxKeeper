"""
DOCX <-> HTML conversion for the web editor.

Supports: headings, paragraphs, alignment, indent, nested bullet/numbered lists
(numbering restarts per list), fonts / sizes / colours / highlight, bold / italic /
underline / strike / sub / superscript, hyperlinks, tables (colspan, rowspan, header
cells, cell shading + borders => callout boxes), code blocks, page breaks, images
(charts + draw.io diagrams, with sizing), and Word comments (python-docx >= 1.2).
"""
import base64
import html
import json
import re
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

from bs4 import BeautifulSoup, NavigableString, Tag, Comment as HtmlComment
from lxml import etree
from docx import Document
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .file_safety import exclusive_write_lock, atomic_write_docx, get_version_stamp  # noqa: F401
from .document_versioning import save_new_version

_SAFE_URL = ("http://", "https://", "mailto:", "ftp://")
_CODE_CSS = ("background:#0f172a;color:#e2e8f0;padding:12px;border-radius:6px;"
             "font-family:Consolas,'Courier New',monospace;font-size:.9rem;white-space:pre-wrap")
_HIGHLIGHT = {
    "YELLOW": "#ffff00", "BRIGHT_GREEN": "#00ff00", "TURQUOISE": "#00ffff", "PINK": "#ff00ff",
    "BLUE": "#0000ff", "RED": "#ff0000", "DARK_BLUE": "#000080", "TEAL": "#008080",
    "GREEN": "#008000", "VIOLET": "#800080", "DARK_RED": "#800000", "DARK_YELLOW": "#808000",
    "GRAY_50": "#808080", "GRAY_25": "#c0c0c0", "BLACK": "#000000", "WHITE": "#ffffff",
}


# =====================================================================
# DOCX -> HTML (loading into the editor)
# =====================================================================

class _Reader:
    def __init__(self, doc):
        self.doc = doc
        self.images = {rid: rel.target_part for rid, rel in doc.part.rels.items() if "image" in rel.reltype}
        self.open_comments = []

    # ---- blocks (body or table cell) ----
    def blocks(self, parent_elm, parent):
        out, stack = [], []

        def close_lists(keep=0):
            while len(stack) > keep:
                out.append(f"</li></{stack.pop()}>\n")

        for child in parent_elm.iterchildren():
            if child.tag == qn("w:p"):
                para = Paragraph(child, parent)
                kind = self._list_kind(para)
                if kind is None:
                    close_lists()
                    out.append(self.paragraph(para) + "\n")
                    continue
                tag, level = kind
                close_lists(level + 1)
                if len(stack) == level + 1:
                    if stack[level] == tag:
                        out.append("</li><li>")
                    else:
                        out.append(f"</li></{stack[level]}><{tag}><li>")
                        stack[level] = tag
                else:
                    while len(stack) < level + 1:
                        out.append(f"<{tag}><li>")
                        stack.append(tag)
                out.append(self.inline(para))
            elif child.tag == qn("w:tbl"):
                close_lists()
                out.append(self.table(Table(child, parent)) + "\n")
        close_lists()
        return "".join(out)

    def _list_kind(self, para):
        name = (getattr(para.style, "name", "") or "").lower()
        m = re.match(r"list (bullet|number)(?: (\d+))?$", name)
        if m:
            return ("ol" if m.group(1) == "number" else "ul", min(int(m.group(2) or 1) - 1, 5))
        pPr = para._p.pPr
        if pPr is not None and pPr.numPr is not None and not name.startswith(("heading", "title")):
            try:
                level = pPr.numPr.ilvl.val if pPr.numPr.ilvl is not None else 0
            except Exception:
                level = 0
            return ("ol" if "number" in name else "ul", min(int(level), 5))
        return None

    def paragraph(self, para):
        pPr = para._p.pPr
        shd = pPr.find(qn("w:shd")) if pPr is not None else None
        if shd is not None and (shd.get(qn("w:fill")) or "").upper() == "F1F5F9":   # marker written by _code()
            return f'<pre style="{_CODE_CSS}">{html.escape(para.text)}</pre>'
        name = (getattr(para.style, "name", "") or "")
        tag = "p"
        m = re.match(r"heading (\d)", name.lower())
        if m:
            tag = f"h{min(max(int(m.group(1)), 1), 6)}"
        elif name.lower() == "title":
            tag = "h1"
        inner = self.inline(para) or "<br>"
        align = {WD_ALIGN_PARAGRAPH.CENTER: "center", WD_ALIGN_PARAGRAPH.RIGHT: "right",
                 WD_ALIGN_PARAGRAPH.JUSTIFY: "justify"}.get(para.alignment)
        style = f' style="text-align:{align}"' if align else ""
        return f"<{tag}{style}>{inner}</{tag}>"

    # ---- inline content ----
    def inline(self, para):
        out = []
        # re-open comment ranges that started in an earlier paragraph
        out.extend(self._span(cid) for cid in self.open_comments)
        for child in para._p.iterchildren():
            if child.tag == qn("w:r"):
                out.append(self.run(Run(child, para)))
            elif child.tag == qn("w:hyperlink"):
                inner = "".join(self.run(Run(r, para)) for r in child.iterchildren(qn("w:r")))
                rid = child.get(qn("r:id"))
                url = self.doc.part.rels[rid].target_ref if rid and rid in self.doc.part.rels else None
                if url and url.lower().startswith(_SAFE_URL):
                    out.append(f'<a href="{html.escape(url)}">{inner}</a>')
                else:
                    out.append(inner)
            elif child.tag == qn("w:commentRangeStart"):
                cid = f"dc{child.get(qn('w:id'))}"
                self.open_comments.append(cid)
                out.append(self._span(cid))
            elif child.tag == qn("w:commentRangeEnd"):
                cid = f"dc{child.get(qn('w:id'))}"
                if cid in self.open_comments:
                    self.open_comments.remove(cid)
                    out.append("</span>")
        out.append("</span>" * len(self.open_comments))
        return "".join(out)

    @staticmethod
    def _span(cid):
        return f'<span class="comment-mark" data-cid="{cid}">'

    def run(self, run):
        parts = []
        for blip in run._r.iter(qn("a:blip")):
            part = self.images.get(blip.get(qn("r:embed")))
            if part is None:
                continue
            attrs = ""
            holders = blip.xpath("ancestor::wp:inline | ancestor::wp:anchor")
            if holders:
                ext = holders[0].find(qn("wp:extent"))
                if ext is not None and ext.get("cx"):
                    attrs += f' style="width:{round(int(ext.get("cx")) / 9525)}px;max-width:100%"'
                dp = holders[0].find(qn("wp:docPr"))
                if dp is not None and dp.get("descr"):
                    attrs += f' alt="{html.escape(dp.get("descr"))}"'
            if part.content_type == "image/png" and b"mxfile" in part.blob:  # draw.io PNG => re-editable
                attrs += ' data-diagram="1" title="Double-click to edit diagram"'
            b64 = base64.b64encode(part.blob).decode()
            parts.append(f'<img src="data:{part.content_type};base64,{b64}"{attrs}>')

        text = run.text or ""
        if text:
            t = html.escape(text, quote=False).replace("\n", "<br>").replace("\t", "&emsp;")
            f = run.font
            css = []
            if f.name:
                css.append(f"font-family:'{html.escape(f.name)}'")
            if f.size:
                css.append(f"font-size:{f.size.pt:g}pt")
            try:
                rgb = f.color.rgb
            except Exception:
                rgb = None
            if rgb is not None:
                css.append(f"color:#{rgb}")
            bg = self._run_bg(run)
            if bg:
                css.append(f"background-color:{bg}")
            if css:
                t = f'<span style="{";".join(css)}">{t}</span>'
            if run.bold:
                t = f"<strong>{t}</strong>"
            if run.italic:
                t = f"<em>{t}</em>"
            if run.underline:
                t = f"<u>{t}</u>"
            if f.strike:
                t = f"<s>{t}</s>"
            if f.subscript:
                t = f"<sub>{t}</sub>"
            if f.superscript:
                t = f"<sup>{t}</sup>"
            parts.append(t)
        return "".join(parts)

    @staticmethod
    def _run_bg(run):
        try:
            hc = run.font.highlight_color
        except Exception:
            hc = None
        if hc is not None:
            c = _HIGHLIGHT.get(getattr(hc, "name", str(hc)))
            if c:
                return c
        rPr = run._r.rPr
        if rPr is not None:
            shd = rPr.find(qn("w:shd"))
            if shd is not None:
                fill = shd.get(qn("w:fill"))
                if fill and fill.lower() not in ("auto", "ffffff"):
                    return "#" + fill
        return None

    # ---- tables ----
    def table(self, tbl):
        grid = getattr(tbl.style, "name", "") == "Table Grid"
        rows = []
        for tr in tbl._tbl.tr_lst:
            cells = []
            for tc in tr.tc_lst:
                try:
                    span = tc.grid_span
                except Exception:
                    span = 1
                cs = f' colspan="{span}"' if span and span > 1 else ""
                css = self._cell_css(tc)
                style = f' style="{css}"' if css else ""
                cells.append(f"<td{cs}{style}>{self.blocks(tc, _Cell(tc, tbl))}</td>")
            rows.append(f"<tr>{''.join(cells)}</tr>")
        border = ' border="1"' if grid else ""
        return f'<table{border} style="border-collapse:collapse;width:100%">{"".join(rows)}</table>'

    @staticmethod
    def _cell_css(tc):
        css = []
        tcPr = tc.tcPr
        if tcPr is None:
            return ""
        shd = tcPr.find(qn("w:shd"))
        if shd is not None:
            fill = shd.get(qn("w:fill"))
            if fill and fill.lower() not in ("auto", "ffffff"):
                css.append(f"background:#{fill}")
        borders = tcPr.find(qn("w:tcBorders"))
        if borders is not None:
            for side in ("top", "left", "bottom", "right"):
                el = borders.find(qn(f"w:{side}"))
                if el is None or el.get(qn("w:val")) in (None, "nil", "none"):
                    continue
                px = max(1, round(int(el.get(qn("w:sz")) or 4) / 6))
                col = el.get(qn("w:color"))
                col = f"#{col}" if col and col.lower() != "auto" else "#94a3b8"
                css.append(f"border-{side}:{px}px solid {col}")
        return ";".join(css)


def _read_comments(doc):
    """Reads word/comments.xml directly, so it works on any python-docx version."""
    try:
        part = doc.part.part_related_by(RT.COMMENTS)
        root = etree.fromstring(part.blob)
    except Exception:
        return []
    now = datetime.now(timezone.utc).isoformat()
    out = []
    for c in root.findall(qn("w:comment")):
        text = "\n".join("".join(t.text or "" for t in p.iter(qn("w:t"))) for p in c.iter(qn("w:p")))
        date = c.get(qn("w:date"))
        resolved = text.startswith("[Resolved] ")        # written by _apply_comments for resolved comments
        out.append({
            "id": f"dc{c.get(qn('w:id'))}", "text": text[11:] if resolved else text,
            "author": c.get(qn("w:author")) or "Unknown",
            "time": date or now, "resolved": resolved, "quote": "",
        })
    return out


def get_docx_editor_data(file_path, file_name):
    """Returns (html, version, comments). Pass `comments` to the template."""
    target_path = Path(file_path) / file_name
    doc = Document(target_path)
    reader = _Reader(doc)
    comments = _read_comments(doc)
    # Comments are also embedded in the HTML itself, so the editor gets them even if the
    # route doesn't pass `comments` to the template (and returns them on save the same way).
    html_result = _embed_comments(reader.blocks(doc.element.body, doc), comments)
    return html_result, get_version_stamp(target_path), comments


def _embed_comments(html_str, comments):
    if not comments:
        return html_str
    payload = json.dumps(comments).replace("<", "\\u003c")
    return f'{html_str}\n<script type="application/json" id="embedded-comments">{payload}</script>'


def get_docx_as_html(file_path, file_name):
    """Backwards-compatible: returns (html, version)."""
    html_result, version, _ = get_docx_editor_data(file_path, file_name)
    return html_result, version


# =====================================================================
# HTML -> DOCX (saving)
# =====================================================================

_BLOCKS = {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "li", "table", "pre", "hr", "blockquote"}
_HEADINGS = {"h1", "h2", "h3", "h4", "h5", "h6"}
_SKIP = {"script", "style", "head", "meta", "link", "title"}
_ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
          "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}
_FONT_SIZE_MAP = {"1": 8, "2": 10, "3": 12, "4": 14, "5": 18, "6": 24, "7": 36}
_GENERIC_FONTS = {"serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "inherit"}
_NAMED = {"black": "000000", "white": "FFFFFF", "red": "FF0000", "blue": "0000FF", "green": "008000",
          "yellow": "FFFF00", "orange": "FFA500", "gray": "808080", "grey": "808080"}
_XML_BAD = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\u200b\ufeff]")
# schema-order helpers (Word is strict about child order)
_AFTER_PBDR = ("w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
               "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
               "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
               "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
               "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")
_AFTER_SHD = _AFTER_PBDR[1:]
_RPR_AFTER_SHD = ("w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
                  "w:specVanish", "w:oMath")
_TCPR_AFTER_SHD = ("w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark")


class _Ctx:
    def __init__(self, doc):
        self.doc = doc
        self.comment_runs = {}
        self.warnings = []
        self.indent = 0.0
        s = doc.sections[0]
        self.max_in = (s.page_width - s.left_margin - s.right_margin) / 914400.0


# ---- small parsers ----
def _parse_style(attr):
    out = {}
    for part in (attr or "").split(";"):
        if ":" in part:
            k, v = part.split(":", 1)
            out[k.strip().lower()] = v.strip()
    return out


def _parse_color(v):
    if not v:
        return None
    v = v.strip().lower()
    m = re.match(r"rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?\s*\)", v)
    if m:
        if m.group(4) is not None and float(m.group(4)) == 0:
            return None
        return "%02X%02X%02X" % tuple(min(255, int(m.group(i))) for i in (1, 2, 3))
    m = re.match(r"#([0-9a-f]{6})$", v)
    if m:
        return m.group(1).upper()
    m = re.match(r"#([0-9a-f])([0-9a-f])([0-9a-f])$", v)
    if m:
        return "".join(c * 2 for c in m.groups()).upper()
    return _NAMED.get(v)


def _find_color(spec):
    m = re.search(r"(#[0-9a-fA-F]{3,6}\b|rgba?\([^)]*\))", spec or "")
    return _parse_color(m.group(1)) if m else None


def _size_pt(v):
    m = re.match(r"^([\d.]+)\s*(pt|px)$", (v or "").strip().lower())
    if not m:
        return None
    val = float(m.group(1)) * (0.75 if m.group(2) == "px" else 1)
    return round(val * 2) / 2 or None


def _len_in(v, px_default=False):
    m = re.match(r"^(-?[\d.]+)\s*(px|pt|in|cm|mm)?$", (v or "").strip().lower())
    if not m:
        return 0.0
    n, unit = float(m.group(1)), m.group(2) or ("px" if px_default else None)
    if unit is None:
        return 0.0
    n = n / {"px": 96, "pt": 72, "in": 1, "cm": 2.54, "mm": 25.4}[unit]
    return max(n, 0.0)


def _margin_left_in(css):
    if "margin-left" in css:
        return _len_in(css["margin-left"])
    parts = css.get("margin", "").split()
    if len(parts) == 4:
        return _len_in(parts[3])
    if len(parts) in (2, 3):
        return _len_in(parts[1])
    return _len_in(parts[0]) if parts else 0.0


def _inherit(node, base):
    """Return a copy of the inline-style state updated by this node's tag + style attr."""
    st = dict(base)
    n = node.name
    if n in ("strong", "b"):
        st["bold"] = True
    elif n in ("em", "i"):
        st["italic"] = True
    elif n == "u":
        st["underline"] = True
    elif n in ("s", "strike", "del"):
        st["strike"] = True
    elif n == "sub":
        st["sub"], st["sup"] = True, False
    elif n == "sup":
        st["sup"], st["sub"] = True, False
    elif n in ("code", "kbd", "samp", "tt"):
        st["font"] = "Consolas"
    elif n == "mark":
        st["bg"] = "FFFF00"
    elif n == "a" and (node.get("href") or "").lower().startswith(_SAFE_URL):
        st["link"] = node["href"]
    elif n == "font":
        if node.get("face"):
            st["font"] = node["face"].split(",")[0].strip().strip("'\"")
        c = _parse_color(node.get("color"))
        if c:
            st["color"] = c
        if node.get("size") in _FONT_SIZE_MAP:
            st["size"] = _FONT_SIZE_MAP[node["size"]]

    css = _parse_style(node.get("style"))
    fw = css.get("font-weight")
    if fw:
        st["bold"] = fw.lower() in ("bold", "bolder") or (fw.isdigit() and int(fw) >= 600)
    fs = css.get("font-style")
    if fs:
        st["italic"] = fs.lower() in ("italic", "oblique")
    deco = f'{css.get("text-decoration", "")} {css.get("text-decoration-line", "")}'
    if "underline" in deco:
        st["underline"] = True
    if "line-through" in deco:
        st["strike"] = True
    va = css.get("vertical-align")
    if va == "super":
        st["sup"], st["sub"] = True, False
    elif va == "sub":
        st["sub"], st["sup"] = True, False
    ff = css.get("font-family")
    if ff:
        fam = ff.split(",")[0].strip().strip("'\"")
        if fam and fam.lower() not in _GENERIC_FONTS:
            st["font"] = fam
    size = _size_pt(css.get("font-size"))
    if size:
        st["size"] = size
    c = _parse_color(css.get("color"))
    if c:
        st["color"] = c
    if "background-color" in css and n not in ("td", "th", "tr", "table"):
        bg = _parse_color(css["background-color"])
        if bg:
            st["bg"] = bg
        else:
            st.pop("bg", None)          # "transparent" = highlight removed
    if node.get("data-cid") and "comment-mark" in (node.get("class") or []):
        st["cid"] = node["data-cid"]
    return st


# ---- low-level docx helpers ----
def _clean(text):
    return _XML_BAD.sub("", text)


def _shade_run(run, fill):
    rPr = run._r.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
    rPr.insert_element_before(shd, *_RPR_AFTER_SHD)


def _shade_para(p, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
    p._p.get_or_add_pPr().insert_element_before(shd, *_AFTER_SHD)


def _format_run(paragraph, run, st):
    run.bold = st.get("bold")
    run.italic = st.get("italic")
    run.underline = st.get("underline")
    if st.get("strike"):
        run.font.strike = True
    if st.get("sup"):
        run.font.superscript = True
    elif st.get("sub"):
        run.font.subscript = True
    if st.get("font"):
        run.font.name = st["font"]
        rFonts = run._r.get_or_add_rPr().rFonts
        if rFonts is not None:
            rFonts.set(qn("w:eastAsia"), st["font"]); rFonts.set(qn("w:cs"), st["font"])
    if st.get("size"):
        run.font.size = Pt(st["size"])
    if st.get("color"):
        run.font.color.rgb = RGBColor.from_string(st["color"])
    if st.get("bg"):
        _shade_run(run, st["bg"])
    if st.get("link"):
        if not st.get("color"):
            run.font.color.rgb = RGBColor.from_string("0563C1")
        run.underline = True
        rid = paragraph.part.relate_to(st["link"], RT.HYPERLINK, is_external=True)
        hl = OxmlElement("w:hyperlink")
        hl.set(qn("r:id"), rid)
        run._r.addprevious(hl)
        hl.append(run._r)


def _add_para(target, style=None):
    return target.add_paragraph(style=style)


def _add_image(p, node, ctx):
    src = node.get("src") or ""
    m = re.match(r"data:(image/[\w.+-]+);base64,(.*)$", src, re.S)
    if not m or "svg" in m.group(1).lower():
        p.add_run("[image not embedded]")
        return
    try:
        data = base64.b64decode(m.group(2))
        css = _parse_style(node.get("style"))
        w_in = _len_in(css.get("width") or node.get("width"), px_default=True)
        shape = p.add_run().add_picture(BytesIO(data), width=Inches(w_in) if w_in else None)
        max_emu = Inches(ctx.max_in)
        if shape.width > max_emu:                      # never exceed the text area
            shape.height = int(shape.height * max_emu / shape.width)
            shape.width = max_emu
        if node.get("alt"):
            shape._inline.docPr.set("descr", node["alt"])
    except Exception:
        p.add_run("[image could not be embedded]")


# ---- inline + block walkers ----
def _inline(p, node, st, ctx):
    if isinstance(node, HtmlComment):
        return
    if isinstance(node, NavigableString):
        text = _clean(re.sub(r"[ \t\r\n]+", " ", str(node).replace("\u00a0", " ")))
        if text:
            run = p.add_run(text)
            _format_run(p, run, st)
            if st.get("cid"):
                ctx.comment_runs.setdefault(st["cid"], []).append(run)
        return
    if node.name in _SKIP:
        return
    if node.name == "br":
        sib = node.next_sibling                         # trailing <br> in a block is just a placeholder
        while isinstance(sib, NavigableString) and not str(sib).strip():
            sib = sib.next_sibling
        if sib is not None:
            p.add_run().add_break()
        return
    if node.name == "img":
        _add_image(p, node, ctx)
        return
    st2 = _inherit(node, st)
    for child in node.children:
        _inline(p, child, st2, ctx)


def _finish(p, node, ctx):
    css = _parse_style(node.get("style")) if node is not None else {}
    align = css.get("text-align") or (node.get("align") if node is not None else None)
    if align in _ALIGN:
        p.alignment = _ALIGN[align]
    elif p._p.xpath(".//w:drawing") and not p.text.strip():
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER         # lone image/figure => centred
    ind = ctx.indent + _margin_left_in(css)
    if ind > 0:
        p.paragraph_format.left_indent = Inches(ind)


def _blocks(container, target, ctx, base):
    buf = []

    def flush():
        nonlocal buf
        nodes, buf = buf, []
        if not any(isinstance(n, Tag) or str(n).strip() for n in nodes if not isinstance(n, HtmlComment)):
            return
        p = _add_para(target)
        for n in nodes:
            _inline(p, n, base, ctx)
        _finish(p, None, ctx)

    for node in container.children:
        if isinstance(node, HtmlComment):
            continue
        if isinstance(node, Tag) and node.name in _SKIP:
            continue
        if isinstance(node, NavigableString) or node.name not in _BLOCKS:
            buf.append(node)
            continue
        flush()
        _block(node, target, ctx, base)
    flush()


def _block(node, target, ctx, base):
    name = node.name
    st = _inherit(node, base)
    css = _parse_style(node.get("style"))
    if name in _HEADINGS:
        p = _add_para(target, f"Heading {name[1]}")
        for c in node.children:
            _inline(p, c, st, ctx)
        _finish(p, node, ctx)
    elif name in ("p", "li") or (name == "div" and not any(isinstance(c, Tag) and c.name in _BLOCKS for c in node.children)):
        p = _add_para(target)
        for c in node.children:
            _inline(p, c, st, ctx)
        _finish(p, node, ctx)
    elif name == "div":
        ctx.indent += _margin_left_in(css)
        _blocks(node, target, ctx, st)
        ctx.indent -= _margin_left_in(css)
    elif name == "blockquote":
        extra = 0.5 + _margin_left_in(css) if css else 0.5
        ctx.indent += extra
        _blocks(node, target, ctx, st)
        ctx.indent -= extra
    elif name in ("ul", "ol"):
        _list(node, target, ctx, st, 0)
    elif name == "table":
        _table(node, target, ctx, st)
    elif name == "pre":
        _code(node, target)
    elif name == "hr":
        _rule(node, target)


def _new_num_id(doc, style_name):
    """New numbering instance so each <ol> restarts at 1 (python-docx otherwise continues)."""
    try:
        base_id = doc.styles[style_name].element.pPr.numPr.numId.val
        numbering = doc.part.numbering_part.numbering_definitions._numbering
        abstract = next(n.abstractNumId.val for n in numbering.num_lst if n.numId == base_id)
        new = numbering.add_num(abstract)
        new.add_lvlOverride(ilvl=0).add_startOverride(1)
        return new.numId
    except Exception:
        return None


def _list(node, target, ctx, base, level):
    kind = "List Number" if node.name == "ol" else "List Bullet"
    style = kind if level == 0 else f"{kind} {min(level, 2) + 1}"
    num_id = _new_num_id(ctx.doc, kind) if (node.name == "ol" and level == 0) else None

    for li in node.find_all("li", recursive=False):
        st = _inherit(li, base)
        buf, made = [], False

        def emit(nodes):
            p = _add_para(target, style)
            if num_id is not None:
                try:
                    numPr = p._p.get_or_add_pPr().get_or_add_numPr()
                    numPr.get_or_add_numId().val = num_id
                    numPr.get_or_add_ilvl().val = 0
                except Exception:
                    pass
            for n in nodes:
                _inline(p, n, st, ctx)

        for ch in li.children:
            if isinstance(ch, Tag) and ch.name in ("ul", "ol"):
                if buf:
                    emit(buf); buf, made = [], True
                _list(ch, target, ctx, st, level + 1)
                made = True
            elif isinstance(ch, Tag) and ch.name in ("p", "div"):
                if buf:
                    emit(buf); buf = []
                emit(list(ch.children)); made = True
            else:
                buf.append(ch)
        if buf and any(isinstance(n, Tag) or str(n).strip() for n in buf):
            emit(buf)
        elif not made:
            emit([])


def _code(node, target):
    for br in node.find_all("br"):
        br.replace_with("\n")
    lines = node.get_text().replace("\u00a0", " ").replace("\t", "    ").rstrip("\n").split("\n")
    p = _add_para(target)
    _shade_para(p, "F1F5F9")
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run()
    for i, line in enumerate(lines):
        if i:
            run.add_break()
        run.add_text(_clean(line))
    run.font.name = "Consolas"
    run.font.size = Pt(9.5)
    rFonts = run._r.get_or_add_rPr().rFonts
    if rFonts is not None:
        rFonts.set(qn("w:eastAsia"), "Consolas")


def _rule(node, target):
    p = _add_para(target)
    if "pagebreak" in (node.get("class") or []) or "page-break" in (node.get("style") or ""):
        p.add_run().add_break(WD_BREAK.PAGE)
        return
    pBdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    for k, v in (("val", "single"), ("sz", "6"), ("space", "1"), ("color", "94A3B8")):
        b.set(qn(f"w:{k}"), v)
    pBdr.append(b)
    p._p.get_or_add_pPr().insert_element_before(pBdr, *_AFTER_PBDR)


# ---- tables ----
def _rows(table):
    rows = []
    for ch in table.children:
        if not isinstance(ch, Tag):
            continue
        if ch.name == "tr":
            rows.append(ch)
        elif ch.name in ("thead", "tbody", "tfoot"):
            rows.extend(r for r in ch.children if isinstance(r, Tag) and r.name == "tr")
    return rows


def _int(v):
    try:
        return max(1, int(v))
    except (TypeError, ValueError):
        return 1


def _style_cell(cell, css):
    tcPr = cell._tc.get_or_add_tcPr()
    sides = {}
    if "border" in css:
        sides = {s: css["border"] for s in ("top", "left", "bottom", "right")}
    for s in ("top", "left", "bottom", "right"):
        if f"border-{s}" in css:
            sides[s] = css[f"border-{s}"]
    if sides:
        borders = OxmlElement("w:tcBorders")
        for s in ("top", "left", "bottom", "right"):
            if s not in sides:
                continue
            spec = sides[s].lower()
            m = re.search(r"([\d.]+)px", spec)
            px = float(m.group(1)) if m else 1.0
            el = OxmlElement(f"w:{s}")
            if "none" in spec or px == 0:
                el.set(qn("w:val"), "nil")
            else:
                el.set(qn("w:val"), "dashed" if "dashed" in spec else "single")
                el.set(qn("w:sz"), str(max(2, min(96, round(px * 6)))))
                el.set(qn("w:space"), "0")
                el.set(qn("w:color"), _find_color(spec) or "auto")
            borders.append(el)
        tcPr.insert_element_before(borders, "w:shd", *_TCPR_AFTER_SHD)
    fill = _find_color(css.get("background-color") or css.get("background"))
    if fill:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
        tcPr.insert_element_before(shd, *_TCPR_AFTER_SHD)


def _table(node, target, ctx, base):
    rows = _rows(node)
    if not rows:
        return
    occupied, placed = set(), []
    for r, tr in enumerate(rows):
        c = 0
        for td in tr.find_all(["td", "th"], recursive=False):
            while (r, c) in occupied:
                c += 1
            cs, rs = _int(td.get("colspan")), min(_int(td.get("rowspan")), len(rows) - r)
            for dr in range(rs):
                for dc in range(cs):
                    occupied.add((r + dr, c + dc))
            placed.append((r, c, rs, cs, td))
            c += cs
    if not placed:
        return
    n_cols = max(c + cs for _, c, _, cs, _ in placed)
    table = target.add_table(len(rows), n_cols)
    classes = node.get("class") or []
    if node.get("border") not in (None, "", "0") and "callout" not in " ".join(classes):
        table.style = "Table Grid"

    col_in = ctx.max_in / n_cols
    saved_max = ctx.max_in
    for r, c, rs, cs, td in placed:
        cell = table.cell(r, c)
        if rs > 1 or cs > 1:
            cell = cell.merge(table.cell(r + rs - 1, c + cs - 1))
        css = _parse_style(td.get("style"))
        _style_cell(cell, css)
        first_p = cell._tc.p_lst[0]
        st = _inherit(td, base)
        if td.name == "th":
            st["bold"] = True
        ctx.max_in = max(col_in * cs - 0.25, 0.5)
        _blocks(td, cell, ctx, st)
        ctx.max_in = saved_max
        if len(cell._tc.p_lst) > 1 and first_p.getparent() is cell._tc:
            cell._tc.remove(first_p)                    # drop the placeholder paragraph
    if target is not None and not isinstance(target, _Cell):
        pass


# ---- comments ----
def _apply_comments(doc, ctx, comments):
    if not comments or not ctx.comment_runs:
        return
    if not hasattr(doc, "add_comment"):
        try:
            _apply_comments_legacy(doc, ctx, comments)
        except Exception as e:
            ctx.warnings.append(f"Comments were not saved: {e}")
        return
    for c in comments:
        runs = ctx.comment_runs.get(c.get("id"))
        if not runs:
            continue                                    # the commented text was deleted
        author = (c.get("author") or "Unknown")[:100]
        initials = "".join(w[0] for w in author.split())[:3].upper() or "U"
        text = _clean(c.get("text") or "")
        if c.get("resolved"):
            text = "[Resolved] " + text
        try:
            cm = doc.add_comment(runs=runs[0] if len(runs) == 1 else [runs[0], runs[-1]],
                                 text=text, author=author, initials=initials)
        except Exception as e:
            ctx.warnings.append(f"Comment {c.get('id')} skipped: {e}")
            continue
        try:                                            # keep the original timestamp
            cm._comment_elm.set(qn("w:date"), (c.get("time") or "")[:19] + "Z")
        except Exception:
            pass


def _apply_comments_legacy(doc, ctx, comments):
    """Writes real Word comments (word/comments.xml) without python-docx 1.2's API."""
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    root = etree.Element(qn("w:comments"), nsmap={"w": W})
    next_id = 0
    for c in comments:
        runs = ctx.comment_runs.get(c.get("id"))
        if not runs:
            continue
        author = (c.get("author") or "Unknown")[:100]
        el = etree.SubElement(root, qn("w:comment"))
        el.set(qn("w:id"), str(next_id))
        el.set(qn("w:author"), author)
        el.set(qn("w:initials"), "".join(w[0] for w in author.split())[:3].upper() or "U")
        el.set(qn("w:date"), ((c.get("time") or "")[:19] or datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")) + "Z")
        text = _clean(c.get("text") or "")
        if c.get("resolved"):
            text = "[Resolved] " + text
        for line in text.split("\n"):
            p = etree.SubElement(el, qn("w:p"))
            r = etree.SubElement(p, qn("w:r"))
            t = etree.SubElement(r, qn("w:t"))
            t.text = line
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        def anchor(run):                                 # runs wrapped in a hyperlink: anchor on the hyperlink
            parent = run._r.getparent()
            return parent if parent is not None and parent.tag == qn("w:hyperlink") else run._r

        start, end = OxmlElement("w:commentRangeStart"), OxmlElement("w:commentRangeEnd")
        start.set(qn("w:id"), str(next_id)); end.set(qn("w:id"), str(next_id))
        ref = OxmlElement("w:r")
        cref = OxmlElement("w:commentReference"); cref.set(qn("w:id"), str(next_id))
        ref.append(cref)
        anchor(runs[0]).addprevious(start)
        last = anchor(runs[-1])
        last.addnext(end)
        end.addnext(ref)
        next_id += 1
    if next_id == 0:
        return
    part = Part(PackURI("/word/comments.xml"),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
                etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
                doc.part.package)
    doc.part.relate_to(part, RT.COMMENTS)


def _coerce_comments(comments):
    if isinstance(comments, str):
        try:
            comments = json.loads(comments or "[]")
        except ValueError:
            return []
    return comments if isinstance(comments, list) else []


def _build(html_content, comments=None):
    doc = Document()
    ctx = _Ctx(doc)
    soup = BeautifulSoup(html_content or "", "html.parser")
    embedded = []
    tag = soup.find("script", id="embedded-comments")   # sent by the editor with the document HTML
    if tag is not None:
        embedded = _coerce_comments(tag.get_text())
        tag.decompose()
    _blocks(soup, doc, ctx, {})
    _apply_comments(doc, ctx, _coerce_comments(comments) or embedded)
    return doc, ctx.warnings


def _build_document_from_html(html_content, comments=None):
    """Pure function: HTML -> in-memory Document. No disk access."""
    return _build(html_content, comments)[0]


def save_html_as_docx(file_path, file_name, html_content, expected_version,
                      modified_by="System", change_summary="Updated via web editor", comments=None):
    """
    Rebuilds the docx from the edited HTML and saves it via the versioning module.
    `comments` may be the raw JSON string from the form or an already-parsed list.
    Returns {"success": bool, "reason": str|None, "warnings": [str]}.
    """
    target_path = Path(file_path) / file_name

    with exclusive_write_lock(target_path):
        current_version = get_version_stamp(target_path)
        if expected_version is not None and str(current_version) != str(expected_version):
            return {"success": False, "reason": "This document was changed by someone else since you opened it.",
                    "warnings": []}

        try:
            doc, warnings = _build(html_content, comments)
            buffer = BytesIO()
            doc.save(buffer)
            save_new_version(
                file_path=target_path,
                new_file_bytes=buffer.getvalue(),
                modified_by=modified_by,
                change_summary=change_summary,
            )
        except Exception as e:
            return {"success": False, "reason": str(e), "warnings": []}

    return {"success": True, "reason": None, "warnings": warnings}