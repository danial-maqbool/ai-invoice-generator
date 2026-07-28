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

Instructions:

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
     For any product NOT in the catalogue, write it in Title Case, keep any
     dosage/format wording present in the note, and append " Nasal" only if the
     note says nasal.

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


def verify_and_correct_totals(data: dict) -> dict:
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
        qty = int(_money_to_float(item.get("qty")))
        unit = _money_to_float(item.get("unit_price"))
        expected = _fmt_money(qty * unit, symbol)

        if str(item.get("line_total", "")).strip() != expected:
            corrections.append(
                f"{item.get('product', 'item')}: line total "
                f"{item.get('line_total')} → {expected}"
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

    return {
        "computed": running,
        "model_total": model_total,
        "final_total": true_total,
        "corrections": corrections,
        "units": units,
        "lines": len(items),
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


def safe_filename(customer_name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9 _-]", "", customer_name or "Customer").strip()
    cleaned = re.sub(r"\s+", "_", cleaned) or "Customer"
    return f"Invoice_{cleaned}.docx"


# --------------------------------------------------------------------------- #
#  Sidebar
# --------------------------------------------------------------------------- #
def render_sidebar():
    with st.sidebar:
        st.markdown("## ⚙️ Configuration")

        token_ok = bool(os.environ.get("GITHUB_TOKEN"))
        if token_ok:
            st.success("GITHUB_TOKEN loaded", icon="🔑")
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

        st.markdown("### 📎 Invoice Template")
        uploaded_template = st.file_uploader(
            "Upload your `template.docx`",
            type="docx",
            help="A Word document containing Jinja2 placeholder tags.",
        )

        bundled_exists = os.path.exists(BUNDLED_TEMPLATE)
        if uploaded_template is not None:
            st.success(f"Using **{uploaded_template.name}**", icon="✅")
        elif bundled_exists:
            st.info(f"No upload — falling back to bundled **{BUNDLED_TEMPLATE}**", icon="📄")
        else:
            st.warning("No template available. Upload one, or run `python create_template.py`.", icon="⚠️")

        st.divider()

        st.markdown("### 🏷️ Required template tags")
        with st.expander("Header & footer fields", expanded=False):
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
        with st.expander("Products table loop", expanded=False):
            st.markdown(
                "Build a 5-column Word table (`CODE`, `PRODUCT`, `QTY`, "
                "`UNIT PRICE`, `LINE TOTAL`) with **4 rows** — header, loop-open, "
                "the data row, loop-close:"
            )
            st.code(
                "Row 1:  CODE | PRODUCT | QTY | UNIT PRICE | LINE TOTAL\n"
                "Row 2:  {%tr for item in items %}\n"
                "Row 3:  {{ item.code }} | {{ item.product }} | {{ item.qty }} | "
                "{{ item.unit_price }} | {{ item.line_total }}\n"
                "Row 4:  {%tr endfor %}",
                language="jinja2",
            )
            st.warning(
                "Write it as `{%tr for item in items %}` — **no space** after `{%`, "
                "and put it in a row of its own. `{% tr ... %}` with a space is not "
                "recognised by docxtpl and rendering fails with *unknown tag 'tr'*.",
                icon="⚠️",
            )
            st.caption(
                "Don't want to build this by hand? Run `python create_template.py` "
                "to generate a correctly-tagged `template.docx`."
            )

        st.divider()
        st.markdown("### 🧭 How it works")
        st.markdown(
            "1. Paste the raw order text\n"
            "2. GitHub Models extracts structured JSON\n"
            "3. Totals are **recomputed in Python**, never copied\n"
            "4. `docxtpl` fills your Word template\n"
            "5. Download the finished invoice"
        )

        return uploaded_template, model_choice


# --------------------------------------------------------------------------- #
#  Results
# --------------------------------------------------------------------------- #
def render_results(data: dict, docx_bytes: bytes, check: dict, model_used: str):
    st.markdown("---")
    st.markdown("## 📊 Results")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Invoice No.", data.get("invoice_number", "—"))
    m2.metric("Line Items", check["lines"])
    m3.metric("Total Units", check["units"])
    m4.metric("Order Total", check["final_total"])

    st.caption(f"Extracted with `{model_used}` · totals recomputed locally in Python")

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
        st.markdown(
            f"""
            <div class="card">
                <h4>✅ Document ready</h4>
                <p><b>{file_name}</b><br>{len(docx_bytes) / 1024:.1f} KB · Microsoft Word</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.download_button(
            label="⬇️  Download Invoice (.docx)",
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
        st.caption(
            "Open the `.docx` in Word to review before sending. Re-run the "
            "generator to get a fresh invoice number."
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
            <p>Paste messy order notes — get a clean, correctly-costed Word invoice.
            Powered by <b>GitHub Models</b>, rendered into your own
            <b>.docx</b> template with <b>docxtpl</b>.</p>
            <div class="pills">
                <span class="pill">⚡ GitHub Models</span>
                <span class="pill">🧠 gpt-4.1</span>
                <span class="pill">🧮 Totals verified in Python</span>
                <span class="pill">📄 Word output</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_template, model_choice = render_sidebar()

    c1, c2 = st.columns([2.1, 1], gap="large")

    with c1:
        st.markdown("### ✍️ Raw Order Text")
        st.caption(
            "Free-form is fine — names, address, quantities, prices and status notes "
            "in any layout. The model does the tidying."
        )
        raw_text = st.text_area(
            "Raw order text",
            value=SAMPLE_INPUT,
            height=380,
            label_visibility="collapsed",
            placeholder="Paste the customer's order here…",
        )

    with c2:
        st.markdown("### 🎯 What you get")
        st.markdown(
            """
            <div class="card">
                <h4>🏷️ Smart product codes</h4>
                <p>Generated per item — <code>Retatrutide → RETA40</code>,
                <code>NAD Nasal → NADN</code>.</p>
            </div>
            <div class="card">
                <h4>📚 Catalogue names</h4>
                <p>Shorthand is expanded — <code>GLOW → Glow Pen</code>,
                <code>BPC Nasal → BPC-157 Nasal</code>.</p>
            </div>
            <div class="card">
                <h4>🧮 Corrected maths</h4>
                <p>Every total is recomputed in Python, so a typo'd
                <code>£3,300</code> lands as the true <code>£3,200.00</code>.</p>
            </div>
            <div class="card">
                <h4>✨ Clean statuses</h4>
                <p><code>Payment awaiting → Awaiting Payment</code>,
                <code>Ready for delivery → Ready for Dispatch</code>.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("")
    b1, b2 = st.columns([1, 3])
    with b1:
        generate = st.button(
            "🚀  Generate Invoice", type="primary", use_container_width=True
        )
    with b2:
        if st.button("🧹  Clear results", use_container_width=False):
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
        check = verify_and_correct_totals(data)
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

        st.session_state["result"] = {
            "data": data,
            "docx": docx_bytes,
            "check": check,
            "model": model_used,
        }
        st.balloons()

    if "result" in st.session_state:
        result = st.session_state["result"]
        render_results(result["data"], result["docx"], result["check"], result["model"])

    st.markdown(
        '<div class="footer-note">Built with Streamlit · GitHub Models · docxtpl</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
