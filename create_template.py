"""
Builds `template.docx` — the Jinja2-tagged Word template the app renders into.

The layout mirrors the reference invoice (INV-0070, ML Trading International).
Run once after cloning:

    python create_template.py

You can then open template.docx in Word and restyle it freely; just keep the
{{ tags }} and the {% tr for item in items %} loop row intact.
"""

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches

OUTPUT = "template.docx"

INK = RGBColor(0x1F, 0x23, 0x28)
MUTED = RGBColor(0x6B, 0x72, 0x80)
ACCENT = RGBColor(0x4F, 0x46, 0xE5)


def shade(cell, hex_fill):
    """Apply a solid background fill to a table cell."""
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear")
    el.set(qn("w:fill"), hex_fill)
    cell._tc.get_or_add_tcPr().append(el)


def run(paragraph, text, *, size=10, bold=False, color=INK, caps=False):
    r = paragraph.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = "Calibri"
    if caps:
        r.font.all_caps = True
    return r


def label_value(cell, label, value_tag, *, value_size=10.5, bold_value=True):
    """A muted uppercase label with the Jinja2 value tag underneath."""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(1)
    run(p, label, size=7.5, bold=True, color=MUTED, caps=True)

    p2 = cell.add_paragraph()
    p2.paragraph_format.space_after = Pt(0)
    run(p2, value_tag, size=value_size, bold=bold_value)


def build():
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    # ---------------------------------------------------------------- header
    header = doc.add_table(rows=1, cols=2)
    header.autofit = True

    left = header.cell(0, 0).paragraphs[0]
    left.paragraph_format.space_after = Pt(0)
    run(left, "{{ company_name }}", size=19, bold=True, color=ACCENT)

    right = header.cell(0, 1).paragraphs[0]
    right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    right.paragraph_format.space_after = Pt(0)
    run(right, "INVOICE", size=19, bold=True, color=INK, caps=True)

    meta = header.cell(0, 1).add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    meta.paragraph_format.space_before = Pt(2)
    run(meta, "{{ invoice_number }}", size=10.5, bold=True, color=MUTED)

    date_p = header.cell(0, 1).add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run(date_p, "{{ invoice_date }}", size=10, color=MUTED)

    doc.add_paragraph()

    # ------------------------------------------------------- parties / status
    info = doc.add_table(rows=2, cols=2)
    info.style = "Table Grid"
    info.alignment = WD_TABLE_ALIGNMENT.CENTER

    label_value(info.cell(0, 0), "Customer / Receiver", "{{ customer_name }}")
    addr = info.cell(0, 0).add_paragraph()
    addr.paragraph_format.space_before = Pt(2)
    run(addr, "{{ customer_address }}", size=10, color=MUTED)

    label_value(info.cell(0, 1), "Reseller", "{{ reseller_info }}")

    label_value(info.cell(1, 0), "Payment Status", "{{ payment_status }}")
    label_value(info.cell(1, 1), "Shipping Status", "{{ shipping_status }}")

    for row in info.rows:
        for cell in row.cells:
            shade(cell, "F7F7FB")

    doc.add_paragraph()

    # --------------------------------------------------------------- products
    heading = doc.add_paragraph()
    heading.paragraph_format.space_after = Pt(4)
    run(heading, "Products", size=12, bold=True)

    headers = ["CODE", "PRODUCT", "QTY", "UNIT PRICE", "LINE TOTAL"]
    widths = [Inches(0.95), Inches(2.9), Inches(0.6), Inches(1.15), Inches(1.25)]

    # Four rows: header, loop-open, the repeating data row, loop-close.
    #
    # docxtpl's row tag must be written as "{%tr ... %}" — NO space after "{%",
    # exactly one space after "tr" — and it must sit in a row of its own, because
    # docxtpl deletes the entire <w:tr> that contains it. Writing "{% tr ... %}"
    # leaves the tag in the XML and Jinja then fails with "unknown tag 'tr'".
    table = doc.add_table(rows=4, cols=5)
    table.style = "Table Grid"

    for i, text in enumerate(headers):
        cell = table.cell(0, i)
        shade(cell, "4F46E5")
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(2)
        if i >= 2:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run(p, text, size=8.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), caps=True)

    run(table.cell(1, 0).paragraphs[0], "{%tr for item in items %}", size=9)

    body = [
        "{{ item.code }}",
        "{{ item.product }}",
        "{{ item.qty }}",
        "{{ item.unit_price }}",
        "{{ item.line_total }}",
    ]
    for i, text in enumerate(body):
        cell = table.cell(2, i)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(2)
        if i >= 2:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run(p, text, size=10, bold=(i == 0))

    run(table.cell(3, 0).paragraphs[0], "{%tr endfor %}", size=9)

    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = widths[i]

    # ------------------------------------------------------------------ total
    doc.add_paragraph()
    total_p = doc.add_paragraph()
    total_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run(total_p, "Order Total:  ", size=13, bold=True, color=MUTED)
    run(total_p, "{{ order_total }}", size=15, bold=True, color=ACCENT)

    # ----------------------------------------------------------------- footer
    doc.add_paragraph()
    foot = doc.add_paragraph()
    foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(
        foot,
        "Thank you for your order. Payment details available on request.",
        size=8.5,
        color=MUTED,
    )

    doc.save(OUTPUT)
    print(f"Created {OUTPUT}")


if __name__ == "__main__":
    build()
