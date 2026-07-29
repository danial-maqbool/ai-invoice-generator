"""
AI-Powered Invoice Generator
============================
Turns unstructured order text into a polished Word (.docx) invoice.

Pipeline:  Raw text  ->  GitHub Models (gpt-4o)  ->  strict JSON  ->  docxtpl  ->  .docx

Run with:  streamlit run app.py
"""

import os
import io
import json
import re
import shutil
import subprocess
import tempfile
from datetime import datetime

import streamlit as st
from openai import OpenAI
from docxtpl import DocxTemplate
from dotenv import load_dotenv

load_dotenv()

# --------------------------------------------------------------------------- #
#  Page configuration  (must be the first Streamlit call)
# --------------------------------------------------------------------------- #
st.set_page_config(
    page_title="AI Invoice Generator",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_URL = "https://models.github.ai/inference"
BUNDLED_TEMPLATE = "template.docx"
COMPANY_NAME = "ML Trading International"

# Tried in order when the model picker is set to "Auto"; any model the token is
# not entitled to is skipped automatically.
#
# Order is based on a benchmark of this exact invoice task against the reference
# INV-0070 output (see README). gpt-4.1, gpt-4.1-mini and mistral-medium each
# reproduced it byte-for-byte; gpt-4o got the corrected order total wrong
# (£3,250.00 instead of £3,200.00), so it sits below them. gpt-5 leads the list
# so the app upgrades itself if an account ever gains access to it.
MODEL_PREFERENCE = [
    "openai/gpt-5",
    "openai/gpt-4.1",
    "openai/gpt-4.1-mini",
    "mistral-ai/mistral-medium-2505",
    "openai/gpt-4o",
]
AUTO_MODEL = "Auto (best available)"
DEFAULT_MODEL = "openai/gpt-4.1"

SAMPLE_INPUT = """Reseller J/L
Shaq Jordan
Flat 21,
46 falcon road
SW11 2lr

70 × Retatrutide – £2,100 (£30 each)
10 × GLOW – £250 (£25 each)
10 × NAD Nasal – £250 (£25 each)
5 × GLOW Nasal – £100 (£20 each)
5 × Selank Nasal – £100 (£20 each)
5 × BPC/TB-500 Nasal – £100 (£20 each)
5 × BPC Nasal – £100 (£20 each)
5 × PT-141 Nasal – £100 (£20 each)
5 × Kisspeptin Nasal – £100 (£20 each)

Total: £3,300
(Payment awaiting) (Ready for delivery)"""


# --------------------------------------------------------------------------- #
#  Styling
# --------------------------------------------------------------------------- #
def inject_css() -> None:
    st.markdown(
        """
        <style>
            /* --- layout --- */
            .block-container { padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1250px; }
            #MainMenu, footer { visibility: hidden; }

            /* --- hero banner --- */
            .hero {
                background: linear-gradient(115deg, #4f46e5 0%, #7c3aed 45%, #db2777 100%);
                padding: 2.1rem 2.4rem;
                border-radius: 18px;
                color: #ffffff;
                box-shadow: 0 14px 40px -12px rgba(79, 70, 229, .55);
                margin-bottom: 1.6rem;
            }
            .hero h1 {
                font-size: 2.35rem;
                font-weight: 800;
                margin: 0 0 .35rem 0;
                letter-spacing: -.5px;
                color: #ffffff;
            }
            .hero p { font-size: 1.02rem; opacity: .92; margin: 0; line-height: 1.55; }
            .hero .pills { margin-top: 1.1rem; }
            .pill {
                display: inline-block;
                background: rgba(255,255,255,.16);
                border: 1px solid rgba(255,255,255,.28);
                border-radius: 999px;
                padding: .28rem .85rem;
                font-size: .78rem;
                font-weight: 600;
                margin-right: .5rem;
                letter-spacing: .3px;
            }

            /* --- section cards --- */
            .card {
                border: 1px solid rgba(128,128,128,.22);
                border-radius: 14px;
                padding: 1.15rem 1.35rem;
                margin-bottom: 1rem;
                background: rgba(128,128,128,.05);
            }
            .card h4 { margin: 0 0 .45rem 0; font-size: 1.02rem; font-weight: 700; }
            .card p  { margin: 0; font-size: .87rem; opacity: .8; }

            /* --- text area --- */
            .stTextArea textarea {
                font-family: "JetBrains Mono", "Cascadia Code", Consolas, monospace !important;
                font-size: .88rem !important;
                line-height: 1.55 !important;
                border-radius: 12px !important;
            }

            /* --- buttons --- */
            .stButton > button, .stDownloadButton > button {
                border-radius: 11px;
                font-weight: 700;
                letter-spacing: .3px;
                padding: .62rem 1.1rem;
                transition: transform .12s ease, box-shadow .12s ease;
            }
            .stButton > button[kind="primary"] {
                background: linear-gradient(100deg, #4f46e5, #7c3aed);
                border: none;
            }
            .stButton > button:hover, .stDownloadButton > button:hover {
                transform: translateY(-2px);
                box-shadow: 0 10px 22px -10px rgba(79,70,229,.85);
            }
            .stDownloadButton > button {
                background: linear-gradient(100deg, #059669, #10b981);
                color: #fff;
                border: none;
                width: 100%;
                padding: .9rem 1rem;
                font-size: 1rem;
            }
            .stDownloadButton > button:hover { box-shadow: 0 10px 22px -10px rgba(5,150,105,.9); }

            /* --- misc --- */
            .stTabs [data-baseweb="tab"] { font-weight: 600; }
            div[data-testid="stMetricValue"] { font-size: 1.45rem; font-weight: 700; }
            .footer-note { text-align: center; opacity: .55; font-size: .8rem; margin-top: 2.5rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )


# --------------------------------------------------------------------------- #
#  AI extraction  (GitHub Models, OpenAI-compatible endpoint)
# --------------------------------------------------------------------------- #
def get_client() -> OpenAI:
    """Build the OpenAI SDK client pointed at GitHub Models."""
    token = os.environ.get("GITHUB_TOKEN")
    if not token:
        raise RuntimeError(
            "GITHUB_TOKEN is not set. Create a .env file next to app.py containing:\n"
            "GITHUB_TOKEN=your_github_personal_access_token"
        )
    return OpenAI(base_url=BASE_URL, api_key=token)


def build_prompt(raw_text: str) -> str:
    today = datetime.now().strftime("%d/%m/%Y")
    return f"""
You are the invoicing engine for {COMPANY_NAME}. You convert raw, messy order notes
into a strict JSON invoice object that exactly matches the house style shown below.
You never invent products and you never trust the arithmetic written in the raw text.

Raw Text:
\"\"\"{raw_text}\"\"\"

--------------------------------------------------------------------
HOUSE STYLE — this is the authoritative reference invoice (INV-0070).
Reproduce this formatting exactly for equivalent inputs.
--------------------------------------------------------------------
Reseller: J/L
Customer / Receiver: Shaq Jordan
                     Flat 21, 46 Falcon Road
                     SW11 2LR
Payment Status: Awaiting Payment
Shipping Status: Ready for Dispatch

CODE     PRODUCT                QTY   UNIT PRICE   LINE TOTAL
RETA40   Retatrutide 40mg Pen    70   £30.00       £2,100.00
GLOW     Glow Pen                10   £25.00       £250.00
NADN     NAD+ Nasal              10   £25.00       £250.00
GLOWN    Glow Nasal               5   £20.00       £100.00
SELN     Selank Nasal             5   £20.00       £100.00
BPCTB    BPC/TB-500 Nasal         5   £20.00       £100.00
BPC157   BPC-157 Nasal            5   £20.00       £100.00
PT141    PT-141 Nasal             5   £20.00       £100.00
KISS     Kisspeptin Nasal         5   £20.00       £100.00

Order Total: £3,200.00
(The raw note for this order said "Total: £3,300" — that was wrong, and the
correct £3,200.00 was used instead. Always recompute.)
--------------------------------------------------------------------

--------------------------------------------------------------------
INPUT IS FREE-FORM. The order above is only ONE possible layout. Real notes
arrive in many shapes and you must handle all of them without dropping data.
--------------------------------------------------------------------

Work through the note line by line before you answer, and account for EVERY
line. Nothing that describes an ordered product may be omitted, merged with
another product, or silently skipped. If the note lists 14 products, the JSON
has exactly 14 items — never summarise, never truncate, never de-duplicate two
genuinely separate lines.

Expect and correctly parse any of these variations:

• Quantity/product/price written in any order or notation:
    "70 × Retatrutide – £2,100 (£30 each)"      "Retatrutide x70 @ £30"
    "70x Retatrutide £30ea"                      "Qty 70 - Retatrutide - 30.00"
    "- 70 Retatrutide ....... 2100"              "Retatrutide (70) 30.00 2100.00"
  Separators may be ×, x, X, *, @, -, –, —, |, tabs, dots or commas.
  Items may be bulleted (-, •, *, 1., a)) or plain lines, one or many per line.

• Prices with or without a symbol, with or without decimals, with or without
  thousands separators: "£30", "30", "30.00", "£2,100", "2100.00", "GBP 30".
  Currency may be £, $, € or a code; use whatever the note uses and report it
  in currency_symbol (default "£" if none is given).

• Missing figures — derive whatever is absent, never guess wildly:
    - unit price given, no line total  -> line_total = qty × unit_price
    - line total given, no unit price  -> unit_price = line_total ÷ qty
    - both given but inconsistent      -> TRUST qty × unit_price
    - qty absent entirely              -> qty = 1
  "each", "ea", "per", "@", "p/u" all introduce a UNIT price.
  A bare second number on a line is normally the LINE TOTAL, not a unit price.

• Labels that may or may not be present, in any wording or case:
    name      — "Customer:", "Client:", "Ship to:", "Deliver to:", "For:", or
                simply the first non-label line of the note
    address   — one line or many, comma-separated or line-broken
    reseller  — "Reseller J/L", "Reseller: J/L", "Ref J/L", "Agent: J/L", "J/L"
    status    — "(Payment awaiting)", "PAID", "unpaid", "awaiting funds",
                "ready for delivery", "ready to ship", "shipped", "on hold"
  Statuses may sit anywhere in the note, together on one line or apart.

• Noise to ignore: greetings, phone numbers, emails, order dates, tracking
  numbers, "thanks", running commentary, and any stated total (always wrong
  until you verify it). Never turn a note or comment into a line item.

• A stated total, subtotal or "balance" is NEVER copied. Recompute it.

Then produce each field as follows.

1. customer_name — the customer's full name, in Title Case.

2. customer_address — the postal address, cleaned up:
   - Title Case the street/town parts ("46 falcon road" -> "46 Falcon Road").
   - UPPERCASE the postcode ("sw11 2lr" -> "SW11 2LR").
   - Join the sub-premises and street onto ONE line separated by ", "
     ("Flat 21," + "46 falcon road" -> "Flat 21, 46 Falcon Road").
   - Put the postcode on its own final line.
   - Separate lines with "\\n". Strip trailing commas and stray whitespace.
   - Worked example: "Flat 21,\\n46 falcon road\\nSW11 2lr"
                  -> "Flat 21, 46 Falcon Road\\nSW11 2LR"

3. reseller_info — the reseller CODE ONLY, with the word "Reseller" stripped.
   "Reseller J/L" -> "J/L". If no reseller is mentioned, use "N/A".

4. payment_status — standardise:
      "Payment awaiting" / "awaiting payment" -> "Awaiting Payment"
      "paid" / "payment received"             -> "Paid"
   Default when absent: "Awaiting Payment".

5. shipping_status — standardise:
      "Ready for delivery" / "ready to ship"  -> "Ready for Dispatch"
      "shipped" / "sent" / "dispatched"       -> "Dispatched"
   Default when absent: "Pending".

6. items — one entry per ordered line, in the order given. For each item:

   - product: the FULL catalogue product name, not the shorthand in the note.
     Use this catalogue when the item matches (match case-insensitively):
        "Retatrutide"        -> "Retatrutide 40mg Pen"
        "GLOW" (no "Nasal")  -> "Glow Pen"
        "NAD Nasal"          -> "NAD+ Nasal"
        "GLOW Nasal"         -> "Glow Nasal"
        "Selank Nasal"       -> "Selank Nasal"
        "BPC/TB-500 Nasal"   -> "BPC/TB-500 Nasal"
        "BPC Nasal"          -> "BPC-157 Nasal"
        "PT-141 Nasal"       -> "PT-141 Nasal"
        "Kisspeptin Nasal"   -> "Kisspeptin Nasal"
     Match on meaning, not spelling: the note may abbreviate, misspell, change
     case, or omit the format ("retat", "Reta 40", "glow nasal spray",
     "kisspeptin ns", "bpc157"). Map it to the catalogue entry it clearly means.
     For any product NOT in the catalogue, write it in Title Case, keep any
     dosage/format wording present in the note, and append " Nasal" only if the
     note says nasal. Never invent a product that isn't in the note, and never
     drop one because you are unsure of its name — carry it through as written.

   - code: a SHORT uppercase alphanumeric code you generate for the product.
     Use these exact codes for catalogue items:
        Retatrutide 40mg Pen -> "RETA40"     (stem + dosage number)
        Glow Pen             -> "GLOW"
        NAD+ Nasal           -> "NADN"       (stem + "N" for nasal)
        Glow Nasal           -> "GLOWN"
        Selank Nasal         -> "SELN"
        BPC/TB-500 Nasal     -> "BPCTB"      (initials of both peptides)
        BPC-157 Nasal        -> "BPC157"     (stem + variant number)
        PT-141 Nasal         -> "PT141"
        Kisspeptin Nasal     -> "KISS"
     For anything else, derive one the same way: 3-6 leading letters of the
     product stem, plus a dosage/variant number if the name has one, plus a
     trailing "N" for a nasal product that has no distinguishing number.
     Codes must be UNIQUE within the invoice — add a digit if two would collide.

   - qty: the quantity, as an integer (not a string).
   - unit_price: per-unit price as a string, ALWAYS 2 decimal places, e.g. "£30.00".
   - line_total: qty x unit_price, RECALCULATED BY YOU, as a string with a
     thousands separator and 2 decimals, e.g. "£2,100.00".
     Do NOT copy a figure from the raw text — compute it.

7. CRITICAL MATH FIX: order_total is the exact sum of YOUR recalculated line
   totals, formatted like "£3,200.00". IGNORE any total written in the raw text;
   it is frequently wrong. In the reference order the note said "£3,300" while the
   items summed to £3,200 — the invoice must say £3,200.00.

8. invoice_number: a plausible random invoice number, format "INV-XXXX" with 4
   digits (e.g. "INV-0071"). Do not reuse "INV-0070".

9. invoice_date: today's date, which is {today} (DD/MM/YYYY).

10. company_name: "{COMPANY_NAME}".
11. currency_symbol: the currency symbol in use, e.g. "£".
12. total_items: the number of distinct line items (integer).
13. total_units: the sum of all quantities (integer).

BEFORE YOU ANSWER, re-read the raw note once more and confirm:
  - every product line in the note appears exactly once in "items";
  - no item has a missing or zero qty, unit_price or line_total;
  - every line_total equals qty × unit_price;
  - order_total equals the sum of the line totals;
  - the customer name, address, reseller and both statuses are populated
    (use the stated defaults only when genuinely absent from the note).
Fix anything that fails before returning. Return the corrected JSON only.

Return ONLY a valid JSON object with exactly these keys:
{{
    "invoice_number": "string",
    "invoice_date": "string",
    "company_name": "string",
    "customer_name": "string",
    "customer_address": "string",
    "reseller_info": "string",
    "payment_status": "string",
    "shipping_status": "string",
    "items": [
        {{
            "code": "string",
            "product": "string",
            "qty": 0,
            "unit_price": "string",
            "line_total": "string"
        }}
    ],
    "order_total": "string",
    "currency_symbol": "string",
    "total_items": 0,
    "total_units": 0
}}
""".strip()


SYSTEM_PROMPT = (
    "You are a precise data-extraction engine. You always return strict, valid JSON "
    "and you always recompute arithmetic yourself."
)


def _is_unavailable(exc: Exception) -> bool:
    """True when GitHub Models rejects a model the account isn't entitled to."""
    return "unavailable_model" in str(exc) or "Unavailable model" in str(exc)


def _call_model(client: OpenAI, model: str, raw_text: str) -> str:
    """One chat completion. Retries without `temperature` for reasoning models."""
    kwargs = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": build_prompt(raw_text)},
        ],
        "response_format": {"type": "json_object"},
    }
    try:
        # Low temperature keeps extraction deterministic.
        return client.chat.completions.create(temperature=0.1, **kwargs).choices[0].message.content
    except Exception as exc:  # noqa: BLE001
        # gpt-5 style reasoning models only accept the default temperature.
        if "temperature" not in str(exc):
            raise
        return client.chat.completions.create(**kwargs).choices[0].message.content


def extract_invoice_data(raw_text: str, model_choice: str = AUTO_MODEL) -> tuple[dict, str]:
    """
    Send the raw order text to GitHub Models and return (parsed_json, model_used).

    With "Auto", candidates are tried best-first and any model the token cannot
    access is skipped, so a free-tier PAT lands on gpt-4o without extra config.
    """
    client = get_client()
    candidates = MODEL_PREFERENCE if model_choice == AUTO_MODEL else [model_choice]

    last_error: Exception | None = None
    for model in candidates:
        try:
            return json.loads(_call_model(client, model, raw_text)), model
        except Exception as exc:  # noqa: BLE001
            if _is_unavailable(exc) and len(candidates) > 1:
                last_error = exc
                continue  # not entitled to this one — try the next tier down
            raise

    raise RuntimeError(
        f"None of the candidate models were available for this token. Last error: {last_error}"
    )


# --------------------------------------------------------------------------- #
#  Local verification of the AI's arithmetic
# --------------------------------------------------------------------------- #
def _money_to_float(value) -> float:
    """'£2,100.00' -> 2100.0 ; returns 0.0 when unparseable."""
    if isinstance(value, (int, float)):
        return float(value)
    match = re.search(r"-?\d[\d,]*(?:\.\d+)?", str(value or ""))
    return float(match.group(0).replace(",", "")) if match else 0.0


def _fmt_money(amount: float, symbol: str = "£") -> str:
    """2100.0 -> '£2,100.00'"""
    return f"{symbol}{amount:,.2f}"


# A line that looks like an order line: some quantity notation plus a price.
_QTY_LINE = re.compile(
    r"""(?:^|\s)(?:                       # quantity, written any of these ways
            \d+\s*(?:[x×*]|\bea\b|\bpcs\b)   # 70x / 70 ×
          | (?:[x×*])\s*\d+                  # x70
          | (?:qty|quantity)\s*[:=]?\s*\d+   # qty 70
        )""",
    re.IGNORECASE | re.VERBOSE,
)
_TOTAL_LINE = re.compile(r"^\s*(?:sub)?total\b|^\s*balance\b", re.IGNORECASE)


def count_probable_item_lines(raw_text: str) -> int:
    """
    Rough count of order lines in the raw note, used only to warn when the
    model appears to have dropped one. Deliberately conservative: it counts a
    line only when a quantity notation is present, so it under-counts rather
    than raising false alarms on prose.
    """
    lines = 0
    for line in raw_text.splitlines():
        if not line.strip() or _TOTAL_LINE.search(line):
            continue
        if _QTY_LINE.search(line):
            lines += 1
    return lines


def verify_and_correct_totals(data: dict, raw_text: str = "") -> dict:
    """
    Recompute every total in pure Python and overwrite the model's figures.

    The whole point of this app is that the arithmetic is right, so the totals
    must not depend on the LLM getting them right. Line totals are rebuilt from
    qty x unit_price and the order total from the sum of those, with the model's
    original values kept only so the UI can report what was corrected.
    """
    items = data.get("items") or []
    symbol = (data.get("currency_symbol") or "£").strip() or "£"
    corrections: list[str] = []

    running = 0.0
    for item in items:
        name = item.get("product", "item")
        qty = int(_money_to_float(item.get("qty")))
        unit = _money_to_float(item.get("unit_price"))
        stated_line = _money_to_float(item.get("line_total"))

        # Free-form notes often omit one of the three figures. Derive whichever
        # is missing rather than letting a zero silently wipe out the line.
        if qty <= 0:
            qty = round(stated_line / unit) if unit > 0 and stated_line > 0 else 1
            corrections.append(f"{name}: quantity missing → {qty}")
        if unit <= 0 and stated_line > 0:
            unit = stated_line / qty
            corrections.append(f"{name}: unit price derived → {_fmt_money(unit, symbol)}")

        expected = _fmt_money(qty * unit, symbol)
        if str(item.get("line_total", "")).strip() != expected:
            corrections.append(
                f"{name}: line total {item.get('line_total') or '—'} → {expected}"
            )
            item["line_total"] = expected

        item["qty"] = qty
        item["unit_price"] = _fmt_money(unit, symbol)
        running += qty * unit

    model_total = str(data.get("order_total", "")).strip()
    true_total = _fmt_money(running, symbol)
    if model_total != true_total:
        corrections.append(f"order total: {model_total or '—'} → {true_total}")
        data["order_total"] = true_total

    units = sum(int(_money_to_float(i.get("qty"))) for i in items)
    data["total_items"] = len(items)
    data["total_units"] = units

    detected = count_probable_item_lines(raw_text) if raw_text else 0
    missing = max(0, detected - len(items))

    return {
        "computed": running,
        "model_total": model_total,
        "final_total": true_total,
        "corrections": corrections,
        "units": units,
        "lines": len(items),
        "detected_lines": detected,
        "missing_lines": missing,
    }


# --------------------------------------------------------------------------- #
#  Document generation
# --------------------------------------------------------------------------- #
def generate_document(data_dict: dict, template_source) -> bytes:
    """Render the Jinja2-tagged .docx template and return it as raw bytes."""
    context = dict(data_dict)

    # Word needs \a (a vertical-tab style break) rather than \n for soft line breaks
    # inside a single paragraph — otherwise the address renders on one squashed line.
    for key, value in context.items():
        if isinstance(value, str) and "\n" in value:
            context[key] = value.replace("\r\n", "\n").replace("\n", "\a")

    doc = DocxTemplate(template_source)
    doc.render(context)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# --------------------------------------------------------------------------- #
#  PDF conversion
#
#  There is no reliable pure-Python .docx -> .pdf renderer, so we drive a real
#  office suite. Word is tried first because it renders the template's VML
#  rounded panels exactly; LibreOffice is the cross-platform fallback.
# --------------------------------------------------------------------------- #
SOFFICE_CANDIDATES = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
)


def _find_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    return next((p for p in SOFFICE_CANDIDATES if os.path.exists(p)), None)


def _pdf_via_word(src: str, dst: str) -> bool:
    """Microsoft Word via COM (Windows only). Best fidelity for the template."""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False

    # Streamlit runs the script on a worker thread, which must initialise COM.
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(src, False, True)
        doc.SaveAs(dst, 17)  # 17 = wdFormatPDF
        doc.Close(False)
        return os.path.exists(dst)
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:  # noqa: BLE001 - nothing useful to do on teardown
                pass
        pythoncom.CoUninitialize()


def _pdf_via_soffice(src: str, dst: str) -> bool:
    exe = _find_soffice()
    if not exe:
        return False
    subprocess.run(
        [exe, "--headless", "--norestore", "--convert-to", "pdf",
         "--outdir", os.path.dirname(dst), src],
        check=True, capture_output=True, timeout=180,
    )
    return os.path.exists(dst)


def convert_to_pdf(docx_bytes: bytes) -> bytes | None:
    """Render the .docx to PDF, or return None when no converter is available."""
    with tempfile.TemporaryDirectory() as workdir:
        src = os.path.join(workdir, "invoice.docx")
        dst = os.path.join(workdir, "invoice.pdf")
        with open(src, "wb") as f:
            f.write(docx_bytes)

        for backend in (_pdf_via_word, _pdf_via_soffice):
            try:
                if backend(src, dst):
                    with open(dst, "rb") as f:
                        return f.read()
            except Exception:  # noqa: BLE001 - fall through to the next backend
                continue
    return None


def pdf_converter_name() -> str | None:
    """Which converter this machine can use, for messaging in the UI."""
    try:
        import win32com  # noqa: F401
        if os.name == "nt":
            return "Microsoft Word"
    except ImportError:
        pass
    return "LibreOffice" if _find_soffice() else None


def safe_filename(customer_name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9 _-]", "", customer_name or "Customer").strip()
    cleaned = re.sub(r"\s+", "_", cleaned) or "Customer"
    return f"Invoice_{cleaned}.docx"


# --------------------------------------------------------------------------- #
#  Sidebar
# --------------------------------------------------------------------------- #
def render_sidebar():
    with st.sidebar:
        st.markdown("### ⚙️ Settings")

        if os.environ.get("GITHUB_TOKEN"):
            st.caption("🔑 GITHUB_TOKEN loaded")
        else:
            st.error("GITHUB_TOKEN missing — add it to your `.env`", icon="🔒")

        model_choice = st.selectbox(
            "Model",
            [AUTO_MODEL] + MODEL_PREFERENCE,
            index=0,
            help=(
                "Auto tries gpt-5 first and falls back to the best model your token "
                "can actually reach. Free GitHub Models tokens land on gpt-4o."
            ),
        )
        st.caption("**Provider:** GitHub Models · OpenAI-compatible endpoint")
        st.divider()

        uploaded_template = st.file_uploader(
            "Word template (.docx)",
            type="docx",
            help="Optional — the bundled template.docx is used if you don't upload one.",
        )

        if uploaded_template is not None:
            st.caption(f"Using **{uploaded_template.name}**")
        elif os.path.exists(BUNDLED_TEMPLATE):
            st.caption(f"Using bundled **{BUNDLED_TEMPLATE}**")
        else:
            st.warning("No template found. Run `python create_template.py`.", icon="⚠️")

        with st.expander("Template tags", expanded=False):
            st.code(
                "{{ company_name }}\n"
                "{{ invoice_number }}\n"
                "{{ invoice_date }}\n"
                "{{ customer_name }}\n"
                "{{ customer_address }}\n"
                "{{ reseller_info }}\n"
                "{{ payment_status }}\n"
                "{{ shipping_status }}\n"
                "{{ order_total }}",
                language="jinja2",
            )
            st.markdown(
                "Products table — 5 columns, 4 rows (header, loop-open, data, loop-close):"
            )
            st.code(
                "{%tr for item in items %}\n"
                "{{ item.code }} | {{ item.product }} | {{ item.qty }} | "
                "{{ item.unit_price }} | {{ item.line_total }}\n"
                "{%tr endfor %}",
                language="jinja2",
            )
            st.caption(
                "Write `{%tr` with **no space** after `{%`, in a row of its own — "
                "`{% tr %}` fails with *unknown tag 'tr'*. "
                "Or just run `python create_template.py`."
            )

        return uploaded_template, model_choice


# --------------------------------------------------------------------------- #
#  Results
# --------------------------------------------------------------------------- #
def render_results(
    data: dict, docx_bytes: bytes, check: dict, model_used: str,
    pdf_bytes: bytes | None = None,
):
    st.markdown("---")
    st.markdown("## 📊 Results")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Invoice No.", data.get("invoice_number", "—"))
    m2.metric("Line Items", check["lines"])
    m3.metric("Total Units", check["units"])
    m4.metric("Order Total", check["final_total"])

    st.caption(f"Extracted with `{model_used}` · totals recomputed locally in Python")

    if check.get("missing_lines"):
        st.error(
            f"**Possible missing items.** The note looks like it contains "
            f"{check['detected_lines']} order lines but only {check['lines']} were "
            f"extracted. Check the line items below before sending this invoice.",
            icon="🔎",
        )

    if check["corrections"]:
        st.warning(
            "**Arithmetic corrected.** The model's figures didn't add up, so they were "
            "recalculated:\n\n"
            + "\n".join(f"- {c}" for c in check["corrections"]),
            icon="🧮",
        )
    else:
        st.success(
            f"Arithmetic verified — the line items sum exactly to "
            f"**{check['final_total']}**.",
            icon="🧮",
        )

    left, right = st.columns([1.55, 1], gap="large")

    with left:
        st.markdown("### 🔍 Extracted Data")
        tab_table, tab_json, tab_raw = st.tabs(["📋 Line Items", "🧬 JSON", "📝 Details"])

        with tab_table:
            items = data.get("items") or []
            st.dataframe(
                [
                    {
                        "Code": i.get("code", ""),
                        "Product": i.get("product", ""),
                        "Qty": i.get("qty", ""),
                        "Unit Price": i.get("unit_price", ""),
                        "Line Total": i.get("line_total", ""),
                    }
                    for i in items
                ],
                use_container_width=True,
                hide_index=True,
            )

        with tab_json:
            st.json(data, expanded=True)

        with tab_raw:
            st.markdown(
                f"**Company:** {data.get('company_name', COMPANY_NAME)}  \n"
                f"**Customer:** {data.get('customer_name', '—')}  \n"
                f"**Reseller:** {data.get('reseller_info', '—')}  \n"
                f"**Date:** {data.get('invoice_date', '—')}  \n"
                f"**Payment:** `{data.get('payment_status', '—')}`  \n"
                f"**Shipping:** `{data.get('shipping_status', '—')}`"
            )
            st.markdown("**Address**")
            st.code(data.get("customer_address", "—"), language=None)

    with right:
        st.markdown("### 📥 Your Invoice")
        file_name = safe_filename(data.get("customer_name", ""))

        sizes = f"{len(docx_bytes) / 1024:.0f} KB Word"
        if pdf_bytes:
            sizes += f" · {len(pdf_bytes) / 1024:.0f} KB PDF"
        st.markdown(
            f"""
            <div class="card">
                <h4>✅ Document ready</h4>
                <p><b>{file_name.rsplit('.', 1)[0]}</b><br>{sizes}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if pdf_bytes:
            st.download_button(
                label="⬇️  Download PDF",
                data=pdf_bytes,
                file_name=file_name.replace(".docx", ".pdf"),
                mime="application/pdf",
                use_container_width=True,
            )

        st.download_button(
            label="⬇️  Download Word (.docx)",
            data=docx_bytes,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        st.download_button(
            label="⬇️  Download JSON",
            data=json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8"),
            file_name=file_name.replace(".docx", ".json"),
            mime="application/json",
            use_container_width=True,
        )

        if not pdf_bytes:
            converter = pdf_converter_name()
            st.caption(
                "PDF conversion failed — download the `.docx` and export from "
                f"{converter} instead."
                if converter else
                "PDF export needs Microsoft Word or LibreOffice installed. "
                "Install LibreOffice, or open the `.docx` and save as PDF."
            )


# --------------------------------------------------------------------------- #
#  Main
# --------------------------------------------------------------------------- #
def main():
    inject_css()

    st.markdown(
        """
        <div class="hero">
            <h1>🧾 AI Invoice Generator</h1>
            <p>Paste messy order notes — get a clean, correctly-costed Word invoice.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_template, model_choice = render_sidebar()

    st.markdown("### ✍️ Order Text")
    raw_text = st.text_area(
        "Raw order text",
        value=SAMPLE_INPUT,
        height=340,
        label_visibility="collapsed",
        placeholder="Paste the customer's order here…",
    )

    b1, b2, _ = st.columns([1.2, 1, 2.6])
    with b1:
        generate = st.button(
            "🚀  Generate Invoice", type="primary", use_container_width=True
        )
    with b2:
        if st.button("Clear", use_container_width=True):
            st.session_state.pop("result", None)
            st.rerun()

    if generate:
        template_source = uploaded_template
        if template_source is None:
            if os.path.exists(BUNDLED_TEMPLATE):
                template_source = BUNDLED_TEMPLATE
            else:
                st.error(
                    "No Word template found. Upload one in the sidebar, or run "
                    "`python create_template.py` to build a starter `template.docx`.",
                    icon="🚫",
                )
                st.stop()

        if not raw_text.strip():
            st.error("Paste some order text first.", icon="🚫")
            st.stop()

        try:
            with st.spinner("🧠 Querying GitHub Models — extracting structured data…"):
                data, model_used = extract_invoice_data(raw_text, model_choice)
        except json.JSONDecodeError as exc:
            st.error(f"The model returned something that wasn't valid JSON: {exc}", icon="🧨")
            st.stop()
        except Exception as exc:  # noqa: BLE001 - surface any API/auth failure to the user
            st.error(f"**AI request failed:** {exc}", icon="🧨")
            st.info(
                "Common causes: an expired or unscoped `GITHUB_TOKEN`, no GitHub Models "
                "access on the account, or hitting the free-tier rate limit. "
                "Check your `.env` and try again.",
                icon="💡",
            )
            st.stop()

        # Recompute all arithmetic in Python before it ever reaches the document.
        check = verify_and_correct_totals(data, raw_text)
        data.setdefault("company_name", COMPANY_NAME)

        try:
            with st.spinner("📄 Rendering your Word template…"):
                if hasattr(template_source, "seek"):
                    template_source.seek(0)
                docx_bytes = generate_document(data, template_source)
        except Exception as exc:  # noqa: BLE001 - template problems are user-fixable
            st.error(f"**Template rendering failed:** {exc}", icon="🧨")
            st.info(
                "Make sure every Jinja2 tag in your `.docx` matches the field names "
                "listed in the sidebar, and that the table loop uses `{% tr for ... %}`.",
                icon="💡",
            )
            st.stop()

        pdf_bytes = None
        if pdf_converter_name():
            with st.spinner("🖨️ Converting to PDF…"):
                pdf_bytes = convert_to_pdf(docx_bytes)

        st.session_state["result"] = {
            "data": data,
            "docx": docx_bytes,
            "pdf": pdf_bytes,
            "check": check,
            "model": model_used,
        }
        st.balloons()

    if "result" in st.session_state:
        result = st.session_state["result"]
        render_results(
            result["data"], result["docx"], result["check"],
            result["model"], result.get("pdf"),
        )

    st.markdown(
        '<div class="footer-note">Built with Streamlit · GitHub Models · docxtpl</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
