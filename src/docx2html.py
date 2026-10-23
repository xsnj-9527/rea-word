# -*- coding: utf-8 -*-
"""docx -> HTML converter.

Reads a .docx (which is a zip of XML) and produces semantic HTML.
Design goal: faithful *reading* fidelity for prose documents, not a
pixel-perfect clone of Word's layout engine. Unsupported constructs
(tracked changes, floating objects, text boxes) degrade gracefully and
are reported in `warnings` rather than silently dropped.
"""

import os
import re
import zipfile
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"

# Heading styles are not standardised; match the common English + Chinese names.
HEADING_PATTERNS = [
    (re.compile(r"^(heading|标题)\s*([1-6])$", re.I), lambda m: int(m.group(2))),
    (re.compile(r"^title$|^标题$", re.I), lambda m: 1),
]

# Style ids that Word emits for headings when the display name is localised.
HEADING_STYLE_IDS = {
    "Heading1": 1, "Heading2": 2, "Heading3": 3,
    "Heading4": 4, "Heading5": 5, "Heading6": 6,
    "Title": 1,
}


def _heading_level(style_name, style_id):
    if style_id and style_id in HEADING_STYLE_IDS:
        return HEADING_STYLE_IDS[style_id]
    if style_name:
        for pattern, level in HEADING_PATTERNS:
            m = pattern.match(style_name.strip())
            if m:
                return level(m)
    return None


def _read_styles(zf):
    """Map styleId -> (display name, is_bold, is_italic, size_half_pt, color)."""
    styles = {}
    try:
        data = zf.read("word/styles.xml")
    except KeyError:
        return styles
    root = ET.fromstring(data)
    for st in root.iter(W + "style"):
        sid = st.get(W + "styleId")
        if not sid:
            continue
        name_el = st.find(W + "name")
        name = name_el.get(W + "val") if name_el is not None else sid
        rpr = st.find(W + "rPr")
        bold = italic = False
        size = None
        color = None
        if rpr is not None:
            b = rpr.find(W + "b")
            bold = b is not None and b.get(W + "val") not in ("0", "false")
            i = rpr.find(W + "i")
            italic = i is not None and i.get(W + "val") not in ("0", "false")
            sz = rpr.find(W + "sz")
            if sz is not None and sz.get(W + "val"):
                try:
                    size = int(sz.get(W + "val"))
                except ValueError:
                    pass
            c = rpr.find(W + "color")
            if c is not None:
                color = c.get(W + "val")
        styles[sid] = {
            "name": name, "bold": bold, "italic": italic,
            "size": size, "color": color,
        }
    return styles


def _run_text(run):
    """Concatenate the text-bearing children of a w:r, preserving breaks/tabs."""
    out = []
    for child in run:
        tag = child.tag
        if tag == W + "t":
            out.append(child.text or "")
        elif tag == W + "tab":
            out.append("\t")
        elif tag in (W + "br", W + "cr"):
            out.append("\n")
        elif tag == W + "noBreakHyphen":
            out.append("-")
    return "".join(out)


# --- embedded images -------------------------------------------------------
# A .docx keeps media in word/media/ and refers to it by relationship id.
# We inline the bytes as data: URIs so the viewer needs no file access.

_MIME_BY_EXT = {
    ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
    ".gif": "image/gif", ".bmp": "image/bmp", ".webp": "image/webp",
    ".tif": "image/tiff", ".tiff": "image/tiff", ".svg": "image/svg+xml",
    ".emf": "image/emf", ".wmf": "image/wmf",
}

# EMU (English Metric Units) -> CSS pixels at 96 dpi: 914400 EMU = 1 inch.
_EMU_PER_PX = 9525

# Luminance below this counts as "dark ink" and gets flipped in dark mode.
# 0.35 sits above pure black but below mid-grey, so ordinary dark text flips
# while coloured accents keep their hue.
_DARK_INK_THRESHOLD = 0.35


def _is_dark_color(hex6):
    """True when a #rrggbb colour is too dark to read on a dark background.

    Rec.709 luminance alone misjudges saturated colours: pure red scores 0.21
    yet is perfectly readable on black. So a colour only counts as "dark ink"
    when it is both dim AND close to grey — i.e. the author wrote black-ish
    text rather than a deliberate accent colour.
    """
    h = hex6.strip().lstrip("#")
    if len(h) != 6:
        return False
    try:
        r = int(h[0:2], 16) / 255.0
        g = int(h[2:4], 16) / 255.0
        b = int(h[4:6], 16) / 255.0
    except ValueError:
        return False

    lum = 0.2126 * r + 0.7152 * g + 0.0722 * b
    if lum >= _DARK_INK_THRESHOLD:
        return False
    # Saturation check: a vivid colour keeps its own contrast against dark.
    mx, mn = max(r, g, b), min(r, g, b)
    sat = 0.0 if mx == 0 else (mx - mn) / mx
    return sat < 0.45


def _read_rels(zf, part="word/_rels/document.xml.rels"):
    """Map relationship id -> (target path inside zip, external flag)."""
    rels = {}
    try:
        data = zf.read(part)
    except KeyError:
        return rels
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return rels
    for rel in root:
        rid = rel.get("Id")
        target = rel.get("Target")
        if not rid or not target:
            continue
        external = rel.get("TargetMode") == "External"
        if not external and not target.startswith("word/"):
            target = "word/" + target.lstrip("/")
        rels[rid] = (target, external)
    return rels


def _media_uri(zf, path, warnings):
    """Return a data: URI for one media part, or None when unreadable."""
    if path not in zf.namelist():
        warnings.append("图片缺失：%s" % path)
        return None
    try:
        raw = zf.read(path)
    except KeyError:
        warnings.append("图片读取失败：%s" % path)
        return None
    ext = os.path.splitext(path)[1].lower()
    mime = _MIME_BY_EXT.get(ext)
    if mime is None:
        warnings.append("不支持的图片格式：%s" % ext)
        return None
    # EMF/WMF are vector metafiles browsers cannot render; say so instead of
    # emitting a broken <img>.
    if mime in ("image/emf", "image/wmf"):
        warnings.append("文档含 %s 矢量图，浏览器无法显示" % ext.lstrip(".").upper())
        return None
    import base64
    return "data:%s;base64,%s" % (mime, base64.b64encode(raw).decode("ascii"))


def _emit_drawing(node, rels, zf, warnings):
    """Render one w:drawing (inline or anchored) as an <img> when possible."""
    # Dimensions live in wp:extent (cx/cy) for both inline and anchored images.
    extent = node.find(".//" + WP + "extent")
    width_px = height_px = None
    if extent is not None:
        try:
            width_px = int(extent.get("cx", "0")) / _EMU_PER_PX
            height_px = int(extent.get("cy", "0")) / _EMU_PER_PX
        except ValueError:
            pass

    # The image relationship is on a:blip inside a:graphic/pic:blipFill.
    blip = node.find(".//" + A + "blip")
    uri = None
    if blip is not None:
        rid = blip.get(R + "embed") or blip.get(R + "link")
        if rid and rid in rels:
            target, external = rels[rid]
            if external:
                # Remote images are not fetched: the viewer stays offline.
                warnings.append("文档含网络图片，出于离线考虑未加载")
            else:
                uri = _media_uri(zf, target, warnings)
        elif rid:
            warnings.append("图片关系 %s 未找到" % rid)

    if uri is None:
        return ""  # already reported in warnings

    # Cap the rendered size at the paper width so wide images do not overflow;
    # the reader can still zoom the page.
    css = ["max-width:100%", "height:auto"]
    if width_px:
        css.append("width:%.0fpx" % width_px)
        # Only pin the height when an explicit width exists, otherwise the
        # aspect ratio could break for images Word sized proportionally.
        if height_px:
            css.append("height:auto")
    css.append("border-radius:2px")
    return ('<img class="doc-img" src="%s" style="%s" loading="lazy">'
            % (uri, ";".join(css)))


def _run_props(run, styles, warnings):
    """Return (bold, italic, underline, strike, sup, sub, size, color, font)."""
    rpr = run.find(W + "rPr")
    bold = italic = underline = strike = sup = sub = False
    size = None
    color = None
    font = None
    if rpr is None:
        return dict(bold=bold, italic=italic, underline=underline, strike=strike,
                    sup=sup, sub=sub, size=size, color=color, font=font)

    # Style reference first; direct formatting below overrides it.
    rstyle = rpr.find(W + "rStyle")
    if rstyle is not None:
        sid = rstyle.get(W + "val")
        info = styles.get(sid)
        if info:
            bold = info["bold"]
            italic = info["italic"]
            size = info["size"]
            color = info["color"]

    def flag(tag, cur):
        el = rpr.find(W + tag)
        if el is None:
            return cur
        return el.get(W + "val") not in ("0", "false")

    bold = flag("b", bold)
    italic = flag("i", italic)
    strike = flag("strike", strike)
    underline = rpr.find(W + "u") is not None
    vert = rpr.find(W + "vertAlign")
    if vert is not None:
        v = vert.get(W + "val")
        sup = v == "superscript"
        sub = v == "subscript"

    sz = rpr.find(W + "sz")
    if sz is not None and sz.get(W + "val"):
        try:
            size = int(sz.get(W + "val"))
        except ValueError:
            pass
    c = rpr.find(W + "color")
    if c is not None:
        val = c.get(W + "val")
        if val and val.lower() != "auto":
            color = val
    f = rpr.find(W + "rFonts")
    if f is not None:
        font = f.get(W + "ascii") or f.get(W + "eastAsia")

    return dict(bold=bold, italic=italic, underline=underline, strike=strike,
                sup=sup, sub=sub, size=size, color=color, font=font)


def _escape(text):
    return (text.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;"))


def _emit_run(run, styles, warnings, rels=None, zf=None):
    # Drawings/pictures are emitted first: they may coexist with text.
    images = []
    if rels is not None and zf is not None:
        for drawing in run.findall(W + "drawing"):
            images.append(_emit_drawing(drawing, rels, zf, warnings))
        for pict in run.findall(W + "pict"):
            # Legacy VML pictures: look for the same blip relationship.
            images.append(_emit_drawing(pict, rels, zf, warnings))

    text = _run_text(run)
    if not text:
        # Runs can be pure formatting marks or images; keep images only.
        return "".join(images)
    p = _run_props(run, styles, warnings)
    css = []
    classes = []
    if p["bold"]:
        css.append("font-weight:700")
    if p["italic"]:
        css.append("font-style:italic")
    if p["underline"]:
        css.append("text-decoration:underline")
    if p["strike"]:
        css.append("text-decoration:line-through")
    if p["size"]:
        # half-points -> pt
        css.append("font-size:%.1fpt" % (p["size"] / 2.0))
    if p["color"]:
        color = p["color"].lstrip("#")
        if _is_dark_color(color):
            # Author picked a near-black colour, which is invisible on the
            # dark page. Tag it so the stylesheet can flip it per theme.
            classes.append("dark-ink")
        else:
            css.append("color:#%s" % color)
    if p["font"]:
        css.append("font-family:'%s'" % p["font"].replace("'", ""))

    # Preserve literal tabs/newlines from w:tab / w:br.
    html = _escape(text).replace("\t", "&#9;").replace("\n", "<br>")
    if p["sup"]:
        html = "<sup>%s</sup>" % html
    elif p["sub"]:
        html = "<sub>%s</sub>" % html
    cls = ' class="%s"' % " ".join(classes) if classes else ""
    if css or classes:
        style_attr = ' style="%s"' % ";".join(css) if css else ""
        return "<span%s%s>%s</span>" % (cls, style_attr, html)
    return html


def _emit_paragraph(para, styles, warnings, rels=None, zf=None):
    ppr = para.find(W + "pPr")
    style_name = style_id = None
    align = None
    indent_left = indent_first = None
    numbering = None
    if ppr is not None:
        ps = ppr.find(W + "pStyle")
        if ps is not None:
            style_id = ps.get(W + "val")
            info = styles.get(style_id)
            style_name = info["name"] if info else style_id
        jc = ppr.find(W + "jc")
        if jc is not None:
            align = {"both": "justify", "center": "center",
                     "right": "right", "left": "left"}.get(jc.get(W + "val"))
        ind = ppr.find(W + "ind")
        if ind is not None:
            for attr, key in ((W + "left", "left"), (W + "firstLine", "first")):
                v = ind.get(attr)
                if v:
                    try:
                        val = int(v) / 20.0
                        if key == "left":
                            indent_left = val
                        else:
                            indent_first = val
                    except ValueError:
                        pass
        numpr = ppr.find(W + "numPr")
        if numpr is not None:
            ilvl = numpr.find(W + "ilvl")
            numbering = int(ilvl.get(W + "val")) if ilvl is not None and ilvl.get(W + "val") else 0

    inner = "".join(_emit_run(r, styles, warnings, rels, zf) for r in para.findall(W + "r"))
    if not inner.strip():
        # A paragraph holding only an image still needs its image rendered.
        if "<img" in inner:
            return '<p class="fig">%s</p>' % inner
        return '<p class="empty"></p>'

    level = _heading_level(style_name, style_id)
    css = []
    if align:
        css.append("text-align:%s" % align)
    if indent_left:
        css.append("margin-left:%.1fpt" % indent_left)
    if indent_first:
        css.append("text-indent:%.1fpt" % indent_first)
    style_attr = ' style="%s"' % ";".join(css) if css else ""

    if level:
        return "<h%d%s>%s</h%d>" % (level, style_attr, inner, level)
    if numbering is not None:
        return '<p class="li"%s>%s</p>' % (style_attr, inner)
    return "<p%s>%s</p>" % (style_attr, inner)


def _emit_table(tbl, styles, warnings, rels=None, zf=None):
    rows = []
    for tr in tbl.findall(W + "tr"):
        cells = []
        for tc in tr.findall(W + "tc"):
            tcpr = tc.find(W + "tcPr")
            span = 1
            if tcpr is not None:
                gs = tcpr.find(W + "gridSpan")
                if gs is not None and gs.get(W + "val"):
                    try:
                        span = int(gs.get(W + "val"))
                    except ValueError:
                        pass
            parts = []
            for child in tc:
                if child.tag == W + "p":
                    parts.append(_emit_paragraph(child, styles, warnings, rels, zf))
                elif child.tag == W + "tbl":
                    parts.append(_emit_table(child, styles, warnings, rels, zf))
            attr = ' colspan="%d"' % span if span > 1 else ""
            cells.append("<td%s>%s</td>" % (attr, "".join(parts)))
        rows.append("<tr>%s</tr>" % "".join(cells))
    if not rows:
        return ""
    return '<table class="doc-table">%s</table>' % "".join(rows)


def convert(path):
    """Return (html_body, warnings, meta)."""
    warnings = []
    meta = {}
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        if "word/document.xml" not in names:
            raise ValueError("not a Word document (word/document.xml missing)")

        styles = _read_styles(zf)

        # Core properties are optional; absence is not an error.
        try:
            core = ET.fromstring(zf.read("docProps/core.xml"))
            for tag, key in (("title", "title"), ("creator", "author")):
                for el in core.iter():
                    if el.tag.endswith("}" + tag) and el.text:
                        meta[key] = el.text
                        break
        except (KeyError, ET.ParseError):
            pass

        # Note unsupported constructs so the UI can say what was skipped.
        doc_xml = zf.read("word/document.xml")
        if b"<w:ins " in doc_xml or b"<w:del " in doc_xml:
            warnings.append("文档含修订标记（tracked changes），已按接受后状态显示")
        if b"<w:txbxContent" in doc_xml:
            warnings.append("文档含文本框，文本框内容可能未完整显示")

        rels = _read_rels(zf)

        root = ET.fromstring(doc_xml)
        body = root.find(W + "body")
        if body is None:
            raise ValueError("document body not found")

        chunks = []
        for child in body:
            if child.tag == W + "p":
                chunks.append(_emit_paragraph(child, styles, warnings, rels, zf))
            elif child.tag == W + "tbl":
                chunks.append(_emit_table(child, styles, warnings, rels, zf))
            elif child.tag == W + "sectPr":
                pass  # page setup; not rendered in reading mode
        html = "\n".join(c for c in chunks if c)

    return html, warnings, meta
