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
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches

OUTPUT = "template.docx"

# Greyscale palette taken from the reference invoice.
INK = RGBColor(0x11, 0x11, 0x11)      # headings / primary text
BODY = RGBColor(0x33, 0x33, 0x33)     # table body text
MUTED = RGBColor(0x70, 0x70, 0x70)    # labels, address, meta
CARD_FILL = "F4F4F4"                  # info card background
HEAD_FILL = "EDEDED"                  # products table header
RULE = "D4D4D4"                       # hairline rules
BADGE_EDGE = "BFBFBF"                 # status badge outline

CONTENT_WIDTH = Inches(7.0)


# --------------------------------------------------------------------------- #
#  Low-level docx helpers
# --------------------------------------------------------------------------- #
def shade(cell, hex_fill):
    """Solid background fill for a table cell."""
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear")
    el.set(qn("w:fill"), hex_fill)
    cell._tc.get_or_add_tcPr().append(el)


def set_borders(cell, *, top=None, bottom=None, left=None, right=None, size=4):
    """
    Set individual cell borders. Pass a hex colour to draw an edge, or None to
    explicitly clear it — which is how the table gets horizontal rules only.
    """
    tc_pr = cell._tc.get_or_add_tcPr()
    existing = tc_pr.find(qn("w:tcBorders"))
    if existing is not None:
        tc_pr.remove(existing)

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
    tc_pr.append(borders)


def clear_borders(cell):
    set_borders(cell)


def card(cell):
    """Style a cell as a soft grey info card."""
    shade(cell, CARD_FILL)
    set_borders(cell, top=RULE, bottom=RULE, left=RULE, right=RULE)
    pad(cell, 130)


def pad(cell, twips):
    """Inner margins for a cell, in twentieths of a point."""
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = OxmlElement("w:tcMar")
    for edge, value in (
        ("top", twips), ("bottom", twips),
        ("left", int(twips * 1.1)), ("right", int(twips * 1.1)),
    ):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"), str(value))
        el.set(qn("w:type"), "dxa")
        margins.append(el)
    tc_pr.append(margins)


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
    """Lock column widths (Word ignores cell.width unless autofit is off)."""
    table.autofit = False
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = widths[i]


def borderless(table):
    for row in table.rows:
        for cell in row.cells:
            clear_borders(cell)


# --------------------------------------------------------------------------- #
#  Building blocks
# --------------------------------------------------------------------------- #
def card_label(cell, text):
    p = para(cell, first=True, space_after=3)
    run(p, text, size=9.5, bold=True, color=INK)


def status_badge(cell, tag):
    """An outlined pill holding the status value, as in the reference."""
    badge = cell.add_table(rows=1, cols=1)
    badge.autofit = False
    b = badge.cell(0, 0)
    b.width = Inches(1.55)
    shade(b, "FFFFFF")
    set_borders(b, top=BADGE_EDGE, bottom=BADGE_EDGE, left=BADGE_EDGE, right=BADGE_EDGE, size=4)
    pad(b, 70)
    p = para(b, first=True)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, tag, size=9.5, color=BODY)


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
    header = doc.add_table(rows=2, cols=2)
    fixed_width(header, [Inches(4.2), Inches(2.8)])
    borderless(header)

    title = para(header.cell(0, 0), first=True)
    run(title, "{{ company_name }}", size=18, bold=True, color=INK)

    number = para(header.cell(0, 1), first=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    run(number, "{{ invoice_number }}", size=18, bold=True, color=INK)

    kind = para(header.cell(1, 0), first=True)
    run(kind, "Invoice", size=9, color=MUTED)

    date = para(header.cell(1, 1), first=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    run(date, "{{ invoice_date }}", size=9, color=MUTED)

    # Hairline under the header block.
    for cell in header.rows[1].cells:
        set_borders(cell, bottom=RULE)

    para(doc, space_after=8)

    # ------------------------------------------------------- reseller / customer
    def reseller(cell):
        card_label(cell, "Reseller")
        p = para(cell)
        run(p, "{{ reseller_info }}", size=10.5, bold=True, color=INK)
        # Keeps this card the same height as the taller customer card.
        para(cell, space_before=2)
        para(cell, space_before=2)

    def customer(cell):
        card_label(cell, "Customer / Receiver")
        p = para(cell)
        run(p, "{{ customer_name }}", size=10.5, bold=True, color=INK)
        a = para(cell, space_before=2)
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
    heading = para(doc, space_after=5)
    run(heading, "Products", size=12.5, bold=True, color=INK)

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

    for i, text in enumerate(headers):
        cell = table.cell(0, i)
        shade(cell, HEAD_FILL)
        set_borders(cell, top=RULE, bottom=RULE)
        pad(cell, 90)
        p = para(cell, first=True, align=WD_ALIGN_PARAGRAPH.RIGHT if i >= 2 else None)
        run(p, text, size=7.5, bold=True, color=MUTED, caps=True, spacing=12)

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
        set_borders(cell, bottom=RULE)  # horizontal rules only — no vertical grid
        pad(cell, 105)
        p = para(cell, first=True, align=WD_ALIGN_PARAGRAPH.RIGHT if i >= 2 else None)
        run(p, text, size=10, color=BODY)

    run(para(table.cell(3, 0), first=True), "{%tr endfor %}", size=8)

    for i in (1, 3):
        for cell in table.rows[i].cells:
            clear_borders(cell)

    # ------------------------------------------------------------------- total
    para(doc, space_after=10)
    total = para(doc, align=WD_ALIGN_PARAGRAPH.RIGHT)
    run(total, "Order Total: ", size=14, bold=True, color=INK)
    run(total, "{{ order_total }}", size=14, bold=True, color=INK)

    doc.save(OUTPUT)
    print(f"Created {OUTPUT}")


if __name__ == "__main__":
    build()
