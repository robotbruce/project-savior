#!/usr/bin/env python3
"""Render a deck spec (JSON) into a themed 16:9 .pptx for project handover / teaching.

Run with:
    uv run --with python-pptx python build_deck.py deck.json -o out.pptx

The spec format is documented in ../references/deck-spec.md. Every slide may carry
"notes" (speaker notes). Inline markup inside any text: **bold**, `code`.

After building, the script prints a fit report: any text box whose content is
estimated to overflow even at the minimum font size is listed so you can split
the slide or shorten the text.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

W, H = 13.333, 7.5  # inches, 16:9

DEFAULT_THEME = {
    "primary": "#1E3A5F",    # deep navy: titles, header fills
    "accent": "#F28C28",     # orange: accent bars, highlights
    "text": "#222222",
    "muted": "#6B7280",
    "light": "#F3F5F8",      # card / zebra background
    "border": "#D5DBE3",
    "code_bg": "#1F2430",
    "code_fg": "#E6E6E6",
    "font": "Microsoft JhengHei",   # 微軟正黑體 — present on Windows & Office for Mac
    "mono": "Consolas",
}

WARNINGS: list[str] = []
CUR_SLIDE = [0]


# ---------------------------------------------------------------- helpers
def rgb(hexstr: str) -> RGBColor:
    h = hexstr.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def is_wide(ch: str) -> bool:
    o = ord(ch)
    return (0x2E80 <= o <= 0x9FFF or 0xAC00 <= o <= 0xD7AF or 0xF900 <= o <= 0xFAFF
            or 0xFF00 <= o <= 0xFFEF or 0x3000 <= o <= 0x303F)


def text_width_em(s: str, mono: bool = False) -> float:
    if mono:
        return sum(1.0 if is_wide(c) else 0.6 for c in s)
    return sum(1.0 if is_wide(c) else (0.32 if c in " il.,:;|!'()[]" else
                                       0.62 if c.isupper() or c in "mwMW_" else 0.56) for c in s)


def estimate_height(paras: list[tuple[str, int]], width_in: float, size_pt: float,
                    spacing: float = 1.25, mono: bool = False, para_gap_pt: float = 0) -> float:
    """Return estimated text height in inches for [(text, indent_level)]."""
    total_pt = 0.0
    for text, level in paras:
        avail_pt = max((width_in - 0.25 * level - 0.2) * 72, 20)
        lines = 0
        for seg in (text or " ").split("\n"):
            w = text_width_em(strip_markup(seg), mono) * size_pt
            lines += max(1, int(-(-w // avail_pt)))
        total_pt += lines * size_pt * spacing + para_gap_pt
    return total_pt / 72 + 0.1


def longest_word_em(paras, mono=False) -> float:
    """Width of the longest unbreakable latin token (PowerPoint never splits these mid-word)."""
    words = [w for t, _ in paras for w in re.split(r"[\s\u200b\u2E80-\u9FFF\uFF00-\uFFEF\u3000-\u303F]+",
                                                   strip_markup(t or "")) if w]
    return max((text_width_em(w, mono) for w in words), default=0)


def soft_breaks(text: str) -> str:
    """Let long identifiers wrap at _ . / and camelCase humps instead of shrinking the font.

    Inserts zero-width spaces (invisible, but a legal line-break point in PowerPoint and
    LibreOffice) into tokens longer than 12 characters.
    """
    def fix(m):
        w = m.group(0)
        w = re.sub(r"([_./])(?=[A-Za-z0-9])", "\\1\u200b", w)
        return re.sub(r"(?<=[a-z0-9])(?=[A-Z][a-z])", "\u200b", w)
    return re.sub(r"[A-Za-z0-9_./:\-]{13,}", fix, text or "")


def fit_size(paras, width_in, height_in, max_pt, min_pt, label, **kw) -> float:
    size = max_pt
    word_em = longest_word_em(paras, kw.get("mono", False))
    def too_big(sz):
        return (estimate_height(paras, width_in, sz, **kw) > height_in
                or word_em * sz / 72 > width_in - 0.22)
    while size > min_pt and too_big(size):
        size -= 1
    if estimate_height(paras, width_in, size, **kw) > height_in:
        WARNINGS.append(f"slide {CUR_SLIDE[0]}: '{label}' likely overflows even at {min_pt}pt "
                        f"— split the slide or shorten text")
    return size


def strip_markup(s: str) -> str:
    return re.sub(r"\*\*(.+?)\*\*|`(.+?)`", lambda m: m.group(1) or m.group(2), s)


def set_run_font(run, th, size=None, bold=None, color=None, mono=False, italic=None):
    f = run.font
    name = th["mono"] if mono else th["font"]
    f.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", th["font"] if tag == "a:ea" else name)
    if size:
        f.size = Pt(size)
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    if color:
        f.color.rgb = rgb(color)


def add_rich_text(para, text: str, th, size, color, bold=False, mono=False):
    """Add text to a paragraph honoring **bold** and `code` markup."""
    pos = 0
    for m in re.finditer(r"\*\*(.+?)\*\*|`(.+?)`", text):
        if m.start() > pos:
            r = para.add_run(); r.text = text[pos:m.start()]
            set_run_font(r, th, size, bold, color, mono)
        r = para.add_run()
        if m.group(1) is not None:
            r.text = m.group(1); set_run_font(r, th, size, True, color, mono)
        else:
            r.text = m.group(2); set_run_font(r, th, size * 0.92, bold, th["accent"] if not mono else color, True)
        pos = m.end()
    if pos < len(text) or not text:
        r = para.add_run(); r.text = text[pos:]
        set_run_font(r, th, size, bold, color, mono)


def textbox(slide, x, y, w, h, th, *, anchor=MSO_ANCHOR.TOP, margin=0.05):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", Inches(margin))
    return tb, tf


def put_text(slide, x, y, w, h, text, th, size, color=None, bold=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, mono=False, fit_min=None, label="text"):
    if fit_min:
        size = fit_size([(text, 0)], w, h, size, fit_min, label, mono=mono)
    tb, tf = textbox(slide, x, y, w, h, th, anchor=anchor)
    for i, line in enumerate(str(text).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        add_rich_text(p, line, th, size, color or th["text"], bold, mono)
    return tb


def rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.shadow.inherit = False
    if fill:
        s.fill.solid(); s.fill.fore_color.rgb = rgb(fill)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = rgb(line); s.line.width = Pt(1)
    else:
        s.line.fill.background()
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    return s


def shape_text(s, text, th, size, color, bold=False, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE):
    tf = s.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ("left", "right"):
        setattr(tf, f"margin_{side}", Inches(0.08))
    for side in ("top", "bottom"):
        setattr(tf, f"margin_{side}", Inches(0.04))
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        add_rich_text(p, line, th, size, color, bold)


def norm_items(items) -> list[tuple[str, int]]:
    out = []
    for it in items or []:
        if isinstance(it, str):
            lvl = 0
            while it.startswith("  "):
                lvl += 1; it = it[2:]
            if it.startswith("- "):
                it = it[2:]
            out.append((it, lvl))
        elif isinstance(it, dict):
            out.append((it.get("text", ""), int(it.get("level", 0))))
            for sub in it.get("sub", []) or []:
                out.append((sub, int(it.get("level", 0)) + 1))
    return out


def bullet_list(slide, x, y, w, h, items, th, max_pt=22, min_pt=12, label="bullets", color=None):
    paras = norm_items(items)
    if not paras:
        return
    size = fit_size(paras, w, h, max_pt, min_pt, label, spacing=1.3, para_gap_pt=6)
    tb, tf = textbox(slide, x, y, w, h, th)
    for i, (text, lvl) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(6 if lvl == 0 else 2)
        p.line_spacing = 1.1
        sz = size if lvl == 0 else max(size - 3, min_pt - 1)
        indent = 0.28 + 0.3 * lvl
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(int(Inches(indent))))
        pPr.set("indent", str(-int(Inches(0.26))))
        bu_color = th["accent"] if lvl == 0 else th["muted"]
        for tag in ("a:buClr", "a:buFont", "a:buChar"):
            el = pPr.find(qn(tag))
            if el is not None:
                pPr.remove(el)
        clr = pPr.makeelement(qn("a:buClr"), {})
        srgb = clr.makeelement(qn("a:srgbClr"), {"val": bu_color.lstrip("#")})
        clr.append(srgb); pPr.append(clr)
        pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"}))
        pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "●" if lvl == 0 else "–"}))
        add_rich_text(p, text, th, sz, color or (th["text"] if lvl == 0 else "#444444"))


# ---------------------------------------------------------------- chrome
class Deck:
    def __init__(self, spec: dict):
        self.spec = spec
        self.th = {**DEFAULT_THEME, **(spec.get("theme") or {})}
        self.prs = Presentation()
        self.prs.slide_width = Inches(W)
        self.prs.slide_height = Inches(H)
        self.blank = self.prs.slide_layouts[6]
        self.base = Path(spec.get("_base", "."))
        self.footer = spec.get("footer") or spec.get("title", "")

    def new(self, bg=None):
        s = self.prs.slides.add_slide(self.blank)
        CUR_SLIDE[0] = len(self.prs.slides)
        if bg:
            rect(s, 0, 0, W, H, bg)
        return s

    def chrome(self, s, title, subtitle=None):
        th = self.th
        rect(s, 0.55, 0.5, 0.09, 0.62, th["accent"])
        put_text(s, 0.78, 0.38, W - 1.6, 0.85, title, th, 28, th["primary"], True,
                 anchor=MSO_ANCHOR.MIDDLE, fit_min=20, label="title")
        top = 1.35
        if subtitle:
            put_text(s, 0.8, 1.17, W - 1.6, 0.45, subtitle, th, 15, th["muted"], fit_min=11, label="subtitle")
            top = 1.75
        # footer
        line = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(0.55), Inches(H - 0.5),
                                      Inches(W - 0.55), Inches(H - 0.5))
        line.line.color.rgb = rgb(th["border"]); line.line.width = Pt(0.75)
        put_text(s, 0.55, H - 0.45, 8, 0.32, self.footer, th, 10, th["muted"])
        put_text(s, W - 1.55, H - 0.45, 1.0, 0.32, str(CUR_SLIDE[0]), th, 10, th["muted"], align=PP_ALIGN.RIGHT)
        return top  # y where content may start

    def body_box(self, top):
        return 0.8, top + 0.05, W - 1.6, H - 0.7 - top - 0.15  # x, y, w, h

    def callout(self, s, text, y, h=0.75):
        th = self.th
        box = rect(s, 0.8, y, W - 1.6, h, "#FFF4E8", None, MSO_SHAPE.ROUNDED_RECTANGLE, 0.15)
        rect(s, 0.8, y, 0.08, h, th["accent"])
        size = fit_size([(text, 0)], W - 2.0, h - 0.1, 16, 11, "callout")
        shape_text(box, text, th, size, th["text"], align=PP_ALIGN.LEFT)
        box.text_frame.margin_left = Inches(0.3)

    def notes(self, s, text):
        if text:
            s.notes_slide.notes_text_frame.text = text if isinstance(text, str) else "\n".join(text)

    # ------------------------------------------------------------ types
    def t_title(self, d):
        th = self.th
        s = self.new(th["primary"])
        rect(s, 0, H - 0.18, W, 0.18, th["accent"])
        rect(s, 0.9, 2.35, 0.12, 1.6, th["accent"])
        put_text(s, 1.2, 2.2, W - 2.4, 1.3, d.get("title", self.spec.get("title", "")), th, 42, "#FFFFFF", True,
                 anchor=MSO_ANCHOR.BOTTOM, fit_min=28, label="title")
        if d.get("subtitle"):
            put_text(s, 1.2, 3.5, W - 2.4, 0.8, d["subtitle"], th, 20, "#D9E2EC", fit_min=14, label="subtitle")
        meta = " ｜ ".join(x for x in [d.get("author"), d.get("date")] if x)
        if meta:
            put_text(s, 1.2, 5.6, W - 2.4, 0.5, meta, th, 14, "#C3CFDC")
        return s

    def t_section(self, d):
        th = self.th
        s = self.new(th["primary"])
        if d.get("number") is not None:
            put_text(s, 0.9, 1.9, 3, 1.4, f"{int(d['number']):02d}" if str(d["number"]).isdigit() else str(d["number"]),
                     th, 72, th["accent"], True)
        put_text(s, 0.9, 3.3, W - 1.8, 1.1, d["title"], th, 38, "#FFFFFF", True, fit_min=26, label="section title")
        if d.get("subtitle"):
            put_text(s, 0.9, 4.4, W - 1.8, 1.2, d["subtitle"], th, 18, "#C3CFDC", fit_min=13, label="section subtitle")
        return s

    def t_agenda(self, d):
        th = self.th
        s = self.new()
        top = self.chrome(s, d.get("title", "大綱"), d.get("subtitle"))
        items = d["items"]
        n = len(items)
        cols = 1 if n <= 6 else 2
        per_col = -(-n // cols)
        col_w = (W - 1.6) / cols
        row_h = min(0.85, (H - top - 0.9) / per_col)
        for i, it in enumerate(items):
            c, r = divmod(i, per_col)
            x = 0.8 + c * col_w
            y = top + 0.15 + r * row_h
            circ = rect(s, x, y + (row_h - 0.55) / 2, 0.55, 0.55, th["primary"], shape=MSO_SHAPE.OVAL)
            shape_text(circ, str(i + 1), th, 16, "#FFFFFF", True)
            label = it if isinstance(it, str) else it.get("text", "")
            desc = None if isinstance(it, str) else it.get("desc")
            text = label + (f"\n{desc}" if desc else "")
            tb, tf = textbox(s, x + 0.75, y, col_w - 1.0, row_h, th, anchor=MSO_ANCHOR.MIDDLE)
            p = tf.paragraphs[0]
            add_rich_text(p, label, th, 20 if row_h > 0.7 else 16, th["text"], True)
            if desc:
                p2 = tf.add_paragraph(); add_rich_text(p2, desc, th, 14, th["muted"])
        return s

    def t_bullets(self, d):
        s = self.new()
        top = self.chrome(s, d["title"], d.get("subtitle"))
        x, y, w, h = self.body_box(top)
        if d.get("callout"):
            h -= 0.95
        bullet_list(s, x, y, w, h, d.get("bullets", []), self.th, d.get("max_size", 22))
        if d.get("callout"):
            self.callout(s, d["callout"], y + h + 0.15)
        return s

    def _column(self, s, x, y, w, h, col):
        th = self.th
        if col.get("heading"):
            hb = rect(s, x, y, w, 0.5, th["light"])
            rect(s, x, y, 0.07, 0.5, th["primary"])
            shape_text(hb, col["heading"], th, 16, th["primary"], True, PP_ALIGN.LEFT)
            hb.text_frame.margin_left = Inches(0.2)
            y += 0.6; h -= 0.6
        if col.get("bullets"):
            bullet_list(s, x, y, w, h, col["bullets"], th, 18, 11, label="column")
        elif col.get("code"):
            self._code_box(s, x, y, w, h, col["code"])
        elif col.get("text"):
            put_text(s, x, y, w, h, col["text"], th, 16, fit_min=11, label="column text")

    def t_two_column(self, d):
        s = self.new()
        top = self.chrome(s, d["title"], d.get("subtitle"))
        x, y, w, h = self.body_box(top)
        if d.get("callout"):
            h -= 0.95
        gap = 0.4
        cw = (w - gap) / 2
        self._column(s, x, y, cw, h, d.get("left", {}))
        self._column(s, x + cw + gap, y, cw, h, d.get("right", {}))
        if d.get("callout"):
            self.callout(s, d["callout"], y + h + 0.15)
        return s

    def t_table(self, d):
        th = self.th
        s = self.new()
        top = self.chrome(s, d["title"], d.get("subtitle"))
        x, y, w, h = self.body_box(top)
        if d.get("callout"):
            h -= 0.95
        cols, rows = d["columns"], d["rows"]
        nrows = len(rows) + 1
        if nrows > 14:
            WARNINGS.append(f"slide {CUR_SLIDE[0]}: table has {len(rows)} rows — consider splitting (>13 gets cramped)")
        weights = d.get("col_widths") or [max(4, max(text_width_em(str(r[i]) if i < len(r) else "") for r in [cols] + rows))
                                          for i in range(len(cols))]
        tot = sum(weights)
        widths = [w * wt / tot for wt in weights]
        # font size from densest cell vs. available row height
        row_h = min(0.6, h / nrows)
        size = 16 if row_h >= 0.5 else 14 if row_h >= 0.4 else 12 if row_h >= 0.32 else 10
        for i, cw in enumerate(widths):
            longest = max(text_width_em(strip_markup(str(r[i]))) if i < len(r) else 0 for r in rows + [cols])
            while size > 9 and longest * size / 72 > (cw - 0.15) * 2.2:  # allow ~2 lines per cell
                size -= 1
        gt = s.shapes.add_table(nrows, len(cols), Inches(x), Inches(y), Inches(w), Inches(row_h * nrows))
        tbl = gt.table
        tblPr = tbl._tbl.tblPr
        style = tblPr.find(qn("a:tableStyleId"))
        if style is None:
            style = tblPr.makeelement(qn("a:tableStyleId"), {}); tblPr.append(style)
        style.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"  # "No Style, Table Grid"
        for i, cw in enumerate(widths):
            tbl.columns[i].width = Emu(int(Inches(cw)))
        for r in range(nrows):
            tbl.rows[r].height = Emu(int(Inches(row_h)))
            data = cols if r == 0 else rows[r - 1]
            for c in range(len(cols)):
                cell = tbl.cell(r, c)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(th["primary"] if r == 0 else (th["light"] if r % 2 == 0 else "#FFFFFF"))
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = cell.margin_right = Inches(0.08)
                cell.margin_top = cell.margin_bottom = Inches(0.03)
                tf = cell.text_frame; tf.word_wrap = True
                p = tf.paragraphs[0]
                val = str(data[c]) if c < len(data) else ""
                add_rich_text(p, val, th, size, "#FFFFFF" if r == 0 else th["text"], bold=(r == 0))
        if d.get("callout"):
            self.callout(s, d["callout"], y + h + 0.15)
        return s

    def _code_box(self, s, x, y, w, h, code):
        th = self.th
        lines = code.rstrip("\n").split("\n")
        size = 18
        longest = max((text_width_em(l, True) for l in lines), default=1)
        while size > 9 and (longest * size / 72 > w - 0.4 or len(lines) * size * 1.25 / 72 > h - 0.3):
            size -= 1
        if longest * size / 72 > w - 0.4 or len(lines) * size * 1.25 / 72 > h - 0.3:
            WARNINGS.append(f"slide {CUR_SLIDE[0]}: code block too long ({len(lines)} lines, "
                            f"longest {int(longest)} chars) — trim to the essential lines")
        h = min(h, len(lines) * size * 1.25 / 72 + 0.5)
        rect(s, x, y, w, h, th["code_bg"], shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.03)
        tb, tf = textbox(s, x + 0.2, y + 0.15, w - 0.4, h - 0.3, th)
        tf.word_wrap = False
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            r = p.add_run(); r.text = line or " "
            set_run_font(r, th, size, False, th["code_fg"], mono=True)
        return y + h

    # ------------------------------------------------------------ diagram
    KIND_STYLE = {
        # kind: (shape, fill, line, text color, dashed)
        "service":  (MSO_SHAPE.ROUNDED_RECTANGLE, "light", "border", "text", False),
        "module":   (MSO_SHAPE.ROUNDED_RECTANGLE, "#FFFFFF", "primary", "text", False),
        "client":   (MSO_SHAPE.ROUNDED_RECTANGLE, "primary", None, "#FFFFFF", False),
        "external": (MSO_SHAPE.ROUNDED_RECTANGLE, "#FFFFFF", "muted", "text", True),
        "db":       (MSO_SHAPE.CAN, "#E8EEF6", "primary", "text", False),
        "storage":  (MSO_SHAPE.CAN, "#F3F5F8", "muted", "text", False),
        "queue":    (MSO_SHAPE.FLOWCHART_PREDEFINED_PROCESS, "light", "border", "text", False),
        "note":     (MSO_SHAPE.RECTANGLE, "#FFF8E1", None, "text", False),
    }

    def _c(self, key):
        return self.th.get(key, key) if key else None

    def t_diagram(self, d):
        """Free-form architecture diagram on a grid: groups, typed nodes, labelled edges."""
        th = self.th
        s = self.new()
        top = self.chrome(s, d["title"], d.get("subtitle"))
        x, y, w, h = self.body_box(top)
        if d.get("callout"):
            h -= 0.95
        show_legend = d.get("legend", True)
        if show_legend:
            h -= 0.3
        nodes = d["nodes"]
        for n_ in nodes:   # text budget: a diagram box is a label + one short fact
            desc_ = strip_markup(str(n_.get("desc", "")))
            if len(desc_.split("\n")) > 2 or text_width_em(desc_.replace("\n", "")) > 22:
                WARNINGS.append(f"slide {CUR_SLIDE[0]}: diagram node '{n_.get('id')}' desc too long — "
                                f"≤ 2 short lines (~22 字), move detail to notes")
        cols = d.get("cols") or (max(n["col"] + n.get("w", 1) for n in nodes))
        rows = d.get("rows") or (max(n["row"] + n.get("h", 1) for n in nodes))
        groups = d.get("groups", [])
        gutter = 0.42 if any(g.get("label") for g in groups) else 0.0
        x += gutter; w -= gutter          # left strip reserved for vertical group labels
        cw, ch = w / cols, h / rows
        # widen gaps so edge labels sit between boxes instead of on top of them
        lab_w = max((0.2 + text_width_em(e["label"]) * 11 / 72 for e in d.get("edges", []) if e.get("label")),
                    default=0)
        mx = max(min(0.2, cw * 0.1), min(cw * 0.2, (lab_w + 0.08) / 2))
        my = max(min(0.16, ch * 0.12), min(ch * 0.18, 0.2 if lab_w else 0))

        # groups (drawn first, behind nodes): dashed frame + filled label strip on the left
        for g in groups:
            c0, r0, c1, r1 = g["cells"]  # inclusive start, exclusive end
            gx, gy = x + c0 * cw - (gutter if c0 == 0 else 0) + 0.02, y + r0 * ch + my * 0.25
            gw = x + c1 * cw - mx * 0.35 - gx
            gh = (r1 - r0) * ch - my * 0.5
            color = self._c(g.get("color", "primary"))
            box = rect(s, gx, gy, gw, gh, g.get("fill"), None, MSO_SHAPE.RECTANGLE)
            box.line.color.rgb = rgb(color); box.line.width = Pt(1.25)
            box.line.dash_style = 4  # dash
            lab = g.get("label")
            if lab:
                strip = rect(s, gx, gy, 0.34, gh, color)
                tf = strip.text_frame
                tf.word_wrap = True
                tf._txBody.bodyPr.set("vert", "eaVert")   # CJK vertical text
                for side in ("left", "right", "top", "bottom"):
                    setattr(tf, f"margin_{side}", Inches(0.02))
                size = 12
                while size > 8 and text_width_em(lab) * size / 72 > gh - 0.1:
                    size -= 1
                shape_text(strip, lab, th, size, "#FFFFFF", True)
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        # nodes — a node in the first column of a labelled group that doesn't start at
        # col 0 gives up room on its left for that group's vertical label strip
        strip_cols = [(g["cells"][0], g["cells"][1], g["cells"][3]) for g in groups
                      if g.get("label") and g["cells"][0] > 0]
        geo = {}
        for n in nodes:
            nx = x + n["col"] * cw + mx
            ny = y + n["row"] * ch + my
            nw = n.get("w", 1) * cw - 2 * mx
            nh = n.get("h", 1) * ch - 2 * my
            if any(n["col"] == c0 and r0 <= n["row"] < r1 for c0, r0, r1 in strip_cols):
                shift = max(0.0, 0.42 - mx)
                nx += shift; nw -= shift
            kind = n.get("kind", "service")
            shape, fill, line, tcolor, dashed = self.KIND_STYLE.get(kind, self.KIND_STYLE["service"])
            if n.get("highlight"):
                fill, line = "#FFF4E8", "accent"
            sh = rect(s, nx, ny, nw, nh, self._c(fill), self._c(line), shape,
                      0.12 if shape == MSO_SHAPE.ROUNDED_RECTANGLE else None)
            if shape == MSO_SHAPE.CAN:
                sh.adjustments[0] = min(0.2, 0.12 / max(nh, 0.1) * 1.0)
            if line:
                sh.line.width = Pt(1.5 if n.get("highlight") else 1)
            if dashed:
                sh.line.dash_style = 4
            label, desc = soft_breaks(n["label"]), soft_breaks(n.get("desc", ""))
            text = f"**{label}**" + (f"\n{desc}" if desc else "")
            paras = [(label, 0)] + [(l, 0) for l in desc.split("\n") if desc]
            top_pad = min(0.16, nh * 0.15) if shape == MSO_SHAPE.CAN else 0.0
            size = fit_size(paras, nw - 0.2, nh - 0.12 - top_pad, n.get("size", 14), 8, f"node {n['id']}")
            shape_text(sh, text, th, size, self._c(tcolor))
            if top_pad:
                sh.text_frame.margin_top = Inches(top_pad + 0.04)
            geo[n["id"]] = (nx, ny, nw, nh)

        # edges
        for e in d.get("edges", []):
            if e["from"] not in geo or e["to"] not in geo:
                WARNINGS.append(f"slide {CUR_SLIDE[0]}: edge {e['from']}→{e['to']} references unknown node")
                continue
            ax, ay, aw, ah = geo[e["from"]]
            bx, by, bw, bh = geo[e["to"]]
            acx, acy, bcx, bcy = ax + aw / 2, ay + ah / 2, bx + bw / 2, by + bh / 2
            dx, dy = bcx - acx, bcy - acy
            # overlap-aware anchors: if boxes overlap on an axis, connect straight along the other
            ov_x = min(ax + aw, bx + bw) - max(ax, bx)
            ov_y = min(ay + ah, by + bh) - max(ay, by)
            side = e.get("side")  # force "h" or "v"
            horiz = side == "h" or (side != "v" and (ov_y > 0.15 or (ov_x <= 0.15 and abs(dx) / cw >= abs(dy) / ch)))
            if horiz:
                ymid = (max(ay, by) + min(ay + ah, by + bh)) / 2 if ov_y > 0.15 else None
                x1 = ax + aw if dx > 0 else ax
                x2 = bx if dx > 0 else bx + bw
                y1 = ymid if ymid else acy
                y2 = ymid if ymid else bcy
            else:
                xmid = (max(ax, bx) + min(ax + aw, bx + bw)) / 2 if ov_x > 0.15 else None
                y1 = ay + ah if dy > 0 else ay
                y2 = by if dy > 0 else by + bh
                x1 = xmid if xmid else acx
                x2 = xmid if xmid else bcx
            color = self._c(e.get("color", "muted" if e.get("style") == "dashed" else "accent"))
            c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
            c.line.color.rgb = rgb(color); c.line.width = Pt(1.75)
            if e.get("style") == "dashed":
                c.line.dash_style = 4
            ln = c.line._get_or_add_ln()
            ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
            if e.get("both"):
                ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
            if e.get("label"):
                lab = e["label"]
                lw = 0.2 + text_width_em(lab) * 11 / 72
                lx, ly = (x1 + x2) / 2 - lw / 2, (y1 + y2) / 2 - 0.15
                if horiz and abs(y2 - y1) < 0.05:
                    ly -= 0.17  # sit just above a horizontal line
                tb = rect(s, lx, ly, lw, 0.3, "#FFFFFF", None)
                shape_text(tb, lab, th, 11, th["muted"])
                tb.text_frame.margin_left = tb.text_frame.margin_right = Inches(0.03)

        if show_legend:
            kinds = []
            for n in nodes:
                k = n.get("kind", "service")
                if k not in kinds and k != "note":
                    kinds.append(k)
            names = {"service": "服務/元件", "module": "模組", "client": "呼叫端", "external": "外部服務",
                     "db": "資料庫", "storage": "儲存", "queue": "佇列/排程"}
            if len(kinds) > 1:
                legend = "　".join(f"{'▭' if k not in ('db', 'storage') else '◫'} {names.get(k, k)}" for k in kinds)
                put_text(s, x, y + h + (0.95 if d.get("callout") else 0) + 0.02, w, 0.28,
                         legend + "　│　實線箭頭：呼叫　虛線箭頭：非同步/資料流", th, 10, th["muted"],
                         align=PP_ALIGN.RIGHT)
        if d.get("callout"):
            self.callout(s, d["callout"], y + h + 0.15)
        return s

    def t_code(self, d):
        s = self.new()
        top = self.chrome(s, d["title"], d.get("subtitle"))
        x, y, w, h = self.body_box(top)
        cap = d.get("caption")
        if cap:
            h -= 0.6
        bottom = self._code_box(s, x, y, w, h, d["code"])
        if cap:
            put_text(s, x, bottom + 0.1, w, 0.5, cap, self.th, 14, self.th["muted"], fit_min=10, label="caption")
        return s

    # Flow = numbered stepper. Text budget is deliberately tight: a flow slide is
    # read in seconds, so each step gets a short label + at most one short line;
    # anything longer belongs in the speaker notes.
    FLOW_LABEL_MAX = 8      # CJK-width characters
    FLOW_DESC_MAX = 16

    def _flow_items(self, steps):
        items = []
        for st in steps:
            label = st if isinstance(st, str) else st.get("label", "")
            desc = "" if isinstance(st, str) else st.get("desc", "")
            if isinstance(desc, list):
                desc = "\n".join(str(x) for x in desc)
            items.append((label, desc, bool(st_get(st, "highlight"))))
            if text_width_em(strip_markup(label)) > self.FLOW_LABEL_MAX + 0.5:
                WARNINGS.append(f"slide {CUR_SLIDE[0]}: flow label too long ('{label}') — keep ≤ {self.FLOW_LABEL_MAX} 字")
            if desc and (text_width_em(strip_markup(desc.replace("\n", ""))) > self.FLOW_DESC_MAX + 0.5
                         or "\n" in desc):
                WARNINGS.append(f"slide {CUR_SLIDE[0]}: flow desc too long ('{desc[:20]}…') — one line ≤ "
                                f"{self.FLOW_DESC_MAX} 字, move detail to notes")
        return items

    def _node(self, s, cx, cy, d, num, hi):
        th = self.th
        c = rect(s, cx - d / 2, cy - d / 2, d, d, th["accent"] if hi else th["primary"], "#FFFFFF", MSO_SHAPE.OVAL)
        c.line.width = Pt(4)
        shape_text(c, str(num), th, 24 if d >= 0.8 else 18, "#FFFFFF", True)

    def _line(self, s, x1, y1, x2, y2, width=3):
        ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        ln.line.color.rgb = rgb(self.th["border"]); ln.line.width = Pt(width)

    def t_flow(self, d):
        th = self.th
        s = self.new()
        top = self.chrome(s, d["title"], d.get("subtitle"))
        x, y, w, h = self.body_box(top)
        if d.get("callout"):
            h -= 0.95
        items = self._flow_items(d["steps"])
        n = len(items)
        if n <= 7:
            # horizontal: circles on a line, label + one line underneath
            d_ = 0.95
            span = w - 1.6
            xs = [x + 0.8 + (span * k / (n - 1) if n > 1 else span / 2) for k in range(n)]
            cell = min(2.6, span / max(n - 1, 1) + 0.2)
            block = d_ + 0.25 + 0.5 + 0.45
            cy = y + max(0.2, (h - block) / 2) + d_ / 2
            if n > 1:
                self._line(s, xs[0], cy, xs[-1], cy)
            for k, (label, desc, hi) in enumerate(items):
                self._node(s, xs[k], cy, d_, k + 1, hi)
                put_text(s, xs[k] - cell / 2, cy + d_ / 2 + 0.2, cell, 0.5, label, th,
                         fit_size([(label, 0)], cell, 0.5, 20, 13, "flow label"),
                         th["accent"] if hi else th["text"], True, align=PP_ALIGN.CENTER)
                if desc:
                    put_text(s, xs[k] - cell / 2, cy + d_ / 2 + 0.72, cell, 0.75, desc, th,
                             fit_size([(desc, 0)], cell, 0.75, 15, 11, "flow desc"), th["muted"],
                             align=PP_ALIGN.CENTER)
        else:
            # vertical stepper, two columns when long
            cols = 1 if n <= 6 else 2
            per = -(-n // cols)
            dy = min(0.95, h / per)
            d_ = min(0.62, dy - 0.2)
            colw = w / cols
            for k, (label, desc, hi) in enumerate(items):
                c_, r_ = divmod(k, per)
                cx = x + 0.45 + c_ * colw
                cy = y + 0.1 + r_ * dy + d_ / 2
                if r_ < per - 1 and k < n - 1:
                    self._line(s, cx, cy, cx, cy + dy)
                self._node(s, cx, cy, d_, k + 1, hi)
                lw = min(2.4, colw * 0.38)
                put_text(s, cx + 0.5, cy - 0.3, lw, 0.6, label, th,
                         fit_size([(label, 0)], lw, 0.6, 19, 13, "flow label"),
                         th["accent"] if hi else th["text"], True, anchor=MSO_ANCHOR.MIDDLE)
                if desc:
                    dw = colw - lw - 0.8
                    put_text(s, cx + 0.5 + lw, cy - 0.3, dw, 0.6, desc, th,
                             fit_size([(desc, 0)], dw, 0.6, 15, 11, "flow desc"), th["muted"],
                             anchor=MSO_ANCHOR.MIDDLE)
        if d.get("callout"):
            self.callout(s, d["callout"], y + h + 0.15)
        return s

    def _arrow(self, s, x1, y1, x2, y2, color=None):
        c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        c.line.color.rgb = rgb(color or self.th["accent"]); c.line.width = Pt(2.25)
        ln = c.line._get_or_add_ln()
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))

    def t_architecture(self, d):
        """Stacked layers, each a labelled row of component boxes (top = closest to user)."""
        th = self.th
        s = self.new()
        top = self.chrome(s, d["title"], d.get("subtitle"))
        x, y, w, h = self.body_box(top)
        if d.get("callout"):
            h -= 0.95
        layers = d["layers"]
        n = len(layers)
        gap = 0.28 if d.get("arrows", True) else 0.15
        lh = min(1.5, (h - gap * (n - 1)) / n)
        label_w = 2.0
        for i, layer in enumerate(layers):
            ly = y + i * (lh + gap)
            lab = rect(s, x, ly, label_w, lh, th["primary"], shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.1)
            shape_text(lab, layer["name"], th, fit_size([(layer["name"], 0)], label_w - 0.2, lh - 0.1, 16, 10, "layer"),
                       "#FFFFFF", True)
            items = layer.get("items", [])
            if not items:
                continue
            area_x = x + label_w + 0.2
            area_w = w - label_w - 0.2
            ig = 0.18
            iw = (area_w - ig * (len(items) - 1)) / len(items)
            for j, it in enumerate(items):
                name = it if isinstance(it, str) else it.get("name", "")
                desc = "" if isinstance(it, str) else it.get("desc", "")
                hl = (not isinstance(it, str)) and it.get("highlight")
                box = rect(s, area_x + j * (iw + ig), ly, iw, lh, "#FFF4E8" if hl else th["light"],
                           th["accent"] if hl else th["border"], MSO_SHAPE.ROUNDED_RECTANGLE, 0.1)
                text = f"**{name}**" + (f"\n{desc}" if desc else "")
                size = fit_size([(name, 0)] + ([(desc, 0)] if desc else []), iw - 0.2, lh - 0.1, 15, 9, "component")
                shape_text(box, text, th, size, th["text"])
            if d.get("arrows", True) and i < n - 1:
                self._arrow(s, x + label_w / 2, ly + lh + 0.02, x + label_w / 2, ly + lh + gap - 0.02, th["muted"])
        if d.get("callout"):
            self.callout(s, d["callout"], y + h + 0.15)
        return s

    def t_cards(self, d):
        th = self.th
        s = self.new()
        top = self.chrome(s, d["title"], d.get("subtitle"))
        x, y, w, h = self.body_box(top)
        if d.get("callout"):
            h -= 0.95
        cards = d["cards"]
        n = len(cards)
        cols = d.get("columns") or (n if n <= 4 else 3 if n in (5, 6, 9) else 4)
        nrows = -(-n // cols)
        gx, gy = 0.3, 0.3
        cw = (w - gx * (cols - 1)) / cols
        ch = min(3.4, (h - gy * (nrows - 1)) / nrows)
        for i, cd in enumerate(cards):
            r, c = divmod(i, cols)
            cx, cy = x + c * (cw + gx), y + r * (ch + gy)
            rect(s, cx, cy, cw, ch, th["light"], None, MSO_SHAPE.ROUNDED_RECTANGLE, 0.05)
            rect(s, cx, cy, cw, 0.08, th["accent"] if cd.get("highlight") else th["primary"])
            ty = cy + 0.2
            if cd.get("tag"):
                tag = rect(s, cx + 0.2, ty, min(cw - 0.4, 0.3 + text_width_em(cd["tag"]) * 11 / 72), 0.32,
                           th["accent"], shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
                shape_text(tag, cd["tag"], th, 11, "#FFFFFF", True)
                ty += 0.42
            put_text(s, cx + 0.2, ty, cw - 0.4, 0.6, cd["title"], th, 18, th["primary"], True,
                     fit_min=12, label="card title")
            body = cd.get("body", "")
            body_items = body if isinstance(body, list) else [body]
            by = ty + 0.65
            if isinstance(body, list):
                bullet_list(s, cx + 0.1, by, cw - 0.25, cy + ch - by - 0.1, body_items, th, 15, 10, "card body")
            elif body:
                put_text(s, cx + 0.2, by, cw - 0.4, cy + ch - by - 0.1, body, th, 15, th["text"],
                         fit_min=10, label="card body")
        if d.get("callout"):
            self.callout(s, d["callout"], y + h + 0.15)
        return s

    def t_stats(self, d):
        th = self.th
        s = self.new()
        top = self.chrome(s, d["title"], d.get("subtitle"))
        x, y, w, h = self.body_box(top)
        items = d["items"]
        n = len(items)
        cols = min(n, 4)
        nrows = -(-n // cols)
        gx = 0.3
        cw = (w - gx * (cols - 1)) / cols
        ch = min(2.4, (h - 0.3 * (nrows - 1) - (1.2 if d.get("bullets") else 0)) / nrows)
        for i, it in enumerate(items):
            r, c = divmod(i, cols)
            cx, cy = x + c * (cw + gx), y + r * (ch + 0.3)
            rect(s, cx, cy, cw, ch, th["light"], None, MSO_SHAPE.ROUNDED_RECTANGLE, 0.06)
            put_text(s, cx, cy + 0.15, cw, ch * 0.55, str(it["value"]), th, 44, th["accent"], True,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, fit_min=20, label="stat value")
            put_text(s, cx + 0.15, cy + ch * 0.6, cw - 0.3, ch * 0.38, it["label"], th, 15, th["text"],
                     align=PP_ALIGN.CENTER, fit_min=10, label="stat label")
        if d.get("bullets"):
            by = y + nrows * (ch + 0.3)
            bullet_list(s, x, by, w, y + h - by, d["bullets"], th, 18, 11)
        return s

    def t_timeline(self, d):
        th = self.th
        s = self.new()
        top = self.chrome(s, d["title"], d.get("subtitle"))
        x, y, w, h = self.body_box(top)
        ms = d["milestones"]
        n = len(ms)
        y += max(0, (h - 3.6) / 3)
        ly = y + 1.1
        line = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(ly), Inches(x + w), Inches(ly))
        line.line.color.rgb = rgb(th["border"]); line.line.width = Pt(3)
        step = w / n
        for i, m in enumerate(ms):
            cx = x + step * i + step / 2
            done = m.get("status", "done") == "done"
            rect(s, cx - 0.15, ly - 0.15, 0.3, 0.3, th["primary"] if done else "#FFFFFF",
                 th["primary"], MSO_SHAPE.OVAL)
            put_text(s, cx - step / 2 + 0.05, y, step - 0.1, 0.8, m.get("date", ""), th, 15,
                     th["accent"] if not done else th["primary"], True, align=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.BOTTOM, fit_min=10, label="milestone date")
            text = f"**{m['label']}**" + (f"\n{m['desc']}" if m.get("desc") else "")
            put_text(s, cx - step / 2 + 0.05, ly + 0.35, step - 0.1, 2.2, text, th, 16, th["text"],
                     align=PP_ALIGN.CENTER, fit_min=9, label="milestone")
        return s

    def t_image(self, d):
        th = self.th
        s = self.new()
        top = self.chrome(s, d["title"], d.get("subtitle"))
        x, y, w, h = self.body_box(top)
        if d.get("caption"):
            h -= 0.55
        p = Path(d["path"]).expanduser()
        if not p.is_absolute():
            p = self.base / p
        if not p.exists():
            WARNINGS.append(f"slide {CUR_SLIDE[0]}: image not found: {p}")
            put_text(s, x, y, w, h, f"[missing image: {p.name}]", th, 18, th["muted"], align=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE)
        else:
            pic = s.shapes.add_picture(str(p), Inches(x), Inches(y))
            ratio = min(Inches(w) / pic.width, Inches(h) / pic.height)
            pic.width, pic.height = int(pic.width * ratio), int(pic.height * ratio)
            pic.left = int(Inches(x) + (Inches(w) - pic.width) / 2)
            pic.top = int(Inches(y) + (Inches(h) - pic.height) / 2)
        if d.get("caption"):
            put_text(s, x, y + h + 0.1, w, 0.45, d["caption"], th, 13, th["muted"], align=PP_ALIGN.CENTER,
                     fit_min=10, label="caption")
        return s

    def t_closing(self, d):
        th = self.th
        s = self.new(th["primary"])
        rect(s, 0, H - 0.18, W, 0.18, th["accent"])
        put_text(s, 0.9, 1.0, W - 1.8, 1.2, d.get("title", "Q & A"), th, 40, "#FFFFFF", True,
                 anchor=MSO_ANCHOR.BOTTOM, fit_min=26, label="closing title")
        if d.get("bullets"):
            bullet_list(s, 0.9, 2.5, W - 1.8, H - 3.3, d["bullets"], th, 20, 12, "closing", color="#FFFFFF")
        return s

    def build(self, out: Path):
        for i, d in enumerate(self.spec["slides"], 1):
            t = d.get("type", "bullets")
            fn = getattr(self, f"t_{t}", None)
            if fn is None:
                sys.exit(f"slide {i}: unknown type '{t}'")
            try:
                s = fn(d)
            except KeyError as e:
                sys.exit(f"slide {i} ({t}): missing field {e}")
            self.notes(s, d.get("notes"))
        out.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(out)


def st_get(st, key):
    return isinstance(st, dict) and st.get(key)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec")
    ap.add_argument("-o", "--out")
    a = ap.parse_args()
    sp = Path(a.spec)
    spec = json.loads(sp.read_text(encoding="utf-8"))
    spec["_base"] = str(sp.parent.resolve())
    out = Path(a.out) if a.out else sp.with_suffix(".pptx")
    Deck(spec).build(out)
    print(f"wrote {out}  ({len(spec['slides'])} slides)")
    if WARNINGS:
        print("\nFIT WARNINGS:")
        for w_ in WARNINGS:
            print("  - " + w_)
    else:
        print("fit check: no overflow detected")


if __name__ == "__main__":
    main()
