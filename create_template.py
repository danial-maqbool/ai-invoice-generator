"""
Builds `template.docx` — the Jinja2-tagged Word template the app renders into.

The layout is a faithful recreation of the reference invoice: greyscale, a
two-column header, soft grey info cards, outlined status badges, a products
table with horizontal rules only, and a right-aligned order total.

Run once after cloning:

    python create_template.py

You can open template.docx in Word afterwards and restyle it freely; just keep
the {{ tags }} and the {%tr for item in items %} loop rows intact.
"""

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches

OUTPUT = "template.docx"

# Greyscale palette taken from the reference invoice.
INK = RGBColor(0x11, 0x11, 0x11)      # headings / primary text
BODY = RGBColor(0x2B, 0x2B, 0x2B)     # table body text
MUTED = RGBColor(0x73, 0x73, 0x73)    # labels, address, meta
PAPER = RGBColor(0xFF, 0xFF, 0xFF)    # text on the dark header band

CARD_FILL = "F5F5F5"                  # info card background
HEAD_FILL = "222222"                  # products table header band
RULE = "DCDCDC"                       # hairline row rules
STRONG = "222222"                     # heavy structural rules
BADGE_EDGE = "B0B0B0"                 # status badge outline

HAIRLINE = 4      # 0.5pt — row separators
MEDIUM = 8        # 1.0pt — card edges
HEAVY = 16        # 2.0pt — header underline, totals rule

CELL_PAD = 115    # products table cell padding, in twentieths of a point

CONTENT_WIDTH = Inches(7.0)


# --------------------------------------------------------------------------- #
#  Low-level docx helpers
# --------------------------------------------------------------------------- #
# w:tcPr children must appear in this order or Word renders the cell oddly —
# shading and borders in particular come out with hairline gaps at cell joins.
_TC_ORDER = [
    qn(t) for t in (
        "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders",
        "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
        "w:vAlign", "w:hideMark",
    )
]


def set_tc_property(cell, element):
    """Replace a w:tcPr child, keeping the schema's required element order."""
    tc_pr = cell._tc.get_or_add_tcPr()
    for existing in tc_pr.findall(element.tag):
        tc_pr.remove(existing)

    position = _TC_ORDER.index(element.tag)
    for child in tc_pr:
        if child.tag in _TC_ORDER and _TC_ORDER.index(child.tag) > position:
            child.addprevious(element)
            return
    tc_pr.append(element)


def shade(cell, hex_fill):
    """Solid background fill for a table cell."""
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear")
    el.set(qn("w:color"), "auto")
    el.set(qn("w:fill"), hex_fill)
    set_tc_property(cell, el)


def set_borders(cell, *, top=None, bottom=None, left=None, right=None, size=4):
    """
    Set individual cell borders. Pass a hex colour to draw an edge, or None to
    explicitly clear it — which is how the table gets horizontal rules only.
    """
    borders = OxmlElement("w:tcBorders")
    for edge, colour in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = OxmlElement(f"w:{edge}")
        if colour is None:
            el.set(qn("w:val"), "nil")
        else:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(size))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), colour)
        borders.append(el)
    set_tc_property(cell, borders)


def clear_borders(cell):
    set_borders(cell)


def card(cell):
    """Style a cell as a soft grey info card."""
    shade(cell, CARD_FILL)
    set_borders(cell, top=RULE, bottom=RULE, left=RULE, right=RULE, size=MEDIUM)
    pad(cell, 150)


def pad(cell, twips):
    """Inner margins for a cell, in twentieths of a point."""
    margins = OxmlElement("w:tcMar")
    for edge, value in (
        ("top", twips), ("bottom", twips),
        ("left", int(twips * 1.1)), ("right", int(twips * 1.1)),
    ):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"), str(value))
        el.set(qn("w:type"), "dxa")
        margins.append(el)
    set_tc_property(cell, margins)


def run(paragraph, text, *, size=10, bold=False, color=INK, caps=False, spacing=None):
    r = paragraph.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = "Calibri"
    if caps:
        r.font.all_caps = True
    if spacing:
        el = OxmlElement("w:spacing")
        el.set(qn("w:val"), str(spacing))
        r._element.get_or_add_rPr().append(el)
    return r


def para(container, *, space_before=0, space_after=0, align=None, first=False):
    """Fetch the first paragraph of a cell, or append a new one."""
    p = container.paragraphs[0] if first else container.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if align is not None:
        p.alignment = align
    return p


def fixed_width(table, widths):
    """
    Lock column widths. Word ignores cell.width unless autofit is off *and* the
    layout is declared fixed; without zeroed cell spacing it also leaves hairline
    white gaps at cell joins, which break up shaded bands and full-width rules.
    """
    table.autofit = False

    tbl_pr = table._tbl.tblPr

    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tbl_pr.append(layout)

    spacing = OxmlElement("w:tblCellSpacing")
    spacing.set(qn("w:w"), "0")
    spacing.set(qn("w:type"), "dxa")
    tbl_pr.append(spacing)

    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = widths[i]


def borderless(table):
    for row in table.rows:
        for cell in row.cells:
            clear_borders(cell)


def rule(doc, *, colour=STRONG, size=HEAVY, space_before=0, space_after=0):
    """
    A full-width horizontal rule drawn as a paragraph border.

    A vertically merged cell won't render its bottom border reliably, so the
    masthead rule is drawn this way instead — it always spans the text column.
    """
    p = doc.add_paragraph()

    # w:pBdr must precede w:spacing, so append the border before touching
    # paragraph_format (which is what creates the w:spacing element).
    p_pr = p._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), colour)
    borders.append(bottom)
    p_pr.append(borders)

    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.add_run().font.size = Pt(2)  # keeps the rule's own line height minimal
    return p


# --------------------------------------------------------------------------- #
#  Building blocks
# --------------------------------------------------------------------------- #
def card_label(cell, text):
    """Small uppercase letterspaced label — the eyebrow above each card value."""
    p = para(cell, first=True, space_after=4)
    run(p, text, size=7.5, bold=True, color=MUTED, caps=True, spacing=18)


def status_badge(cell, tag):
    """An outlined pill holding the status value, as in the reference."""
    badge = cell.add_table(rows=1, cols=1)
    badge.autofit = False
    b = badge.cell(0, 0)
    b.width = Inches(1.7)
    shade(b, "FFFFFF")
    set_borders(
        b, top=BADGE_EDGE, bottom=BADGE_EDGE, left=BADGE_EDGE,
        right=BADGE_EDGE, size=MEDIUM,
    )
    pad(b, 85)
    p = para(b, first=True)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, tag, size=10, bold=True, color=INK)


def two_cards(doc, build_left, build_right):
    """A row of two cards separated by an invisible gutter."""
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    fixed_width(table, [Inches(3.35), Inches(0.3), Inches(3.35)])
    borderless(table)

    left, gutter, right = table.cell(0, 0), table.cell(0, 1), table.cell(0, 2)
    card(left)
    card(right)
    clear_borders(gutter)

    build_left(left)
    build_right(right)
    return table


# --------------------------------------------------------------------------- #
#  Template
# --------------------------------------------------------------------------- #
def build():
    doc = Document()

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10)

    section = doc.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    # ------------------------------------------------------------------ header
    header = doc.add_table(rows=3, cols=2)
    fixed_width(header, [Inches(4.0), Inches(3.0)])
    borderless(header)

    # Company name is vertically centred against the stacked invoice block.
    masthead = header.cell(0, 0).merge(header.cell(2, 0))
    masthead.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    title = para(masthead, first=True)
    run(title, "{{ company_name }}", size=21, bold=True, color=INK)

    kind = para(header.cell(0, 1), first=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    run(kind, "Invoice", size=10, bold=True, color=MUTED, caps=True, spacing=60)

    number = para(header.cell(1, 1), first=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    run(number, "{{ invoice_number }}", size=17, bold=True, color=INK)

    date = para(header.cell(2, 1), first=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    run(date, "{{ invoice_date }}", size=9.5, color=MUTED)

    # Heavy rule anchoring the masthead.
    rule(doc, space_before=2, space_after=14)

    # ------------------------------------------------------- reseller / customer
    def reseller(cell):
        card_label(cell, "Reseller")
        p = para(cell)
        run(p, "{{ reseller_info }}", size=12, bold=True, color=INK)
        # Keeps this card the same height as the taller customer card.
        para(cell, space_before=2)
        para(cell, space_before=2)

    def customer(cell):
        card_label(cell, "Customer / Receiver")
        p = para(cell)
        run(p, "{{ customer_name }}", size=12, bold=True, color=INK)
        a = para(cell, space_before=3)
        run(a, "{{ customer_address }}", size=9.5, color=MUTED)

    two_cards(doc, reseller, customer)
    para(doc, space_after=6)

    # ---------------------------------------------------------------- statuses
    def payment(cell):
        card_label(cell, "Payment Status")
        status_badge(cell, "{{ payment_status }}")

    def shipping(cell):
        card_label(cell, "Shipping Status")
        status_badge(cell, "{{ shipping_status }}")

    two_cards(doc, payment, shipping)
    para(doc, space_after=10)

    # ---------------------------------------------------------------- products
    heading = para(doc, space_after=6)
    run(heading, "Products", size=11, bold=True, color=INK, caps=True, spacing=30)

    headers = ["CODE", "PRODUCT", "QTY", "UNIT PRICE", "LINE TOTAL"]
    widths = [Inches(1.05), Inches(2.95), Inches(0.7), Inches(1.15), Inches(1.15)]

    # Four rows: header, loop-open, the repeating data row, loop-close.
    #
    # docxtpl's row tag must be written as "{%tr ... %}" — NO space after "{%",
    # exactly one space after "tr" — and it must sit in a row of its own, because
    # docxtpl deletes the entire <w:tr> that contains it. Writing "{% tr ... %}"
    # leaves the tag in the XML and Jinja then fails with "unknown tag 'tr'".
    table = doc.add_table(rows=4, cols=5)
    fixed_width(table, widths)

    # Dark header band — the strongest horizontal element on the page.
    #
    # It is one merged cell rather than five: abutting shaded cells leave
    # hairline seams that show as vertical lines across the band, and painting
    # the cell edges to hide them just turns the seams into visible bars.
    # Labels are positioned with tab stops derived from the column widths so
    # they stay aligned with the data below.
    band = table.cell(0, 0).merge(table.cell(0, 4))
    shade(band, HEAD_FILL)
    set_borders(band, top=STRONG, bottom=STRONG, size=MEDIUM)
    pad(band, CELL_PAD)

    inset = Inches(CELL_PAD * 1.1 / 1440)
    edges = [sum(widths[: i + 1], Inches(0)) for i in range(len(widths))]

    p = para(band, first=True)
    stops = p.paragraph_format.tab_stops
    stops.add_tab_stop(widths[0], WD_TAB_ALIGNMENT.LEFT)             # PRODUCT
    for edge in edges[2:]:                                            # right-aligned
        stops.add_tab_stop(edge - inset * 2, WD_TAB_ALIGNMENT.RIGHT)

    for i, text in enumerate(headers):
        if i:
            run(p, "\t")
        run(p, text, size=8, bold=True, color=PAPER, caps=True, spacing=20)

    run(para(table.cell(1, 0), first=True), "{%tr for item in items %}", size=8)

    body = [
        "{{ item.code }}",
        "{{ item.product }}",
        "{{ item.qty }}",
        "{{ item.unit_price }}",
        "{{ item.line_total }}",
    ]
    for i, text in enumerate(body):
        cell = table.cell(2, i)
        set_borders(cell, bottom=RULE, size=HAIRLINE)  # horizontal rules only
        pad(cell, CELL_PAD)
        p = para(cell, first=True, align=WD_ALIGN_PARAGRAPH.RIGHT if i >= 2 else None)
        # Codes and money read as the anchors of each row, so they carry weight.
        run(p, text, size=10, bold=i in (0, 4), color=INK if i in (0, 4) else BODY)

    run(para(table.cell(3, 0), first=True), "{%tr endfor %}", size=8)

    for i in (1, 3):
        for cell in table.rows[i].cells:
            clear_borders(cell)

    # ------------------------------------------------------------------- total
    para(doc, space_after=2)

    # A single cell with a right tab stop, rather than two cells — abutting cells
    # leave a hairline gap that would show as a break in the heavy rule above.
    totals = doc.add_table(rows=1, cols=1)
    totals.alignment = WD_TABLE_ALIGNMENT.RIGHT
    fixed_width(totals, [Inches(3.4)])

    cell = totals.cell(0, 0)
    set_borders(cell, top=STRONG, size=HEAVY)
    pad(cell, 110)

    p = para(cell, first=True)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(3.2), WD_TAB_ALIGNMENT.RIGHT)
    run(p, "Order Total", size=10, bold=True, color=MUTED, caps=True, spacing=20)
    run(p, "\t")
    run(p, "{{ order_total }}", size=15, bold=True, color=INK)

    doc.save(OUTPUT)
    print(f"Created {OUTPUT}")


if __name__ == "__main__":
    build()
