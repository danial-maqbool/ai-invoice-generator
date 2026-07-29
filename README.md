# 🧾 AI-Powered Invoice Generator

Turn messy, unstructured order notes into a clean, correctly-costed Microsoft Word invoice.

Paste something like this:

```
Reseller J/L
Shaq Jordan
Flat 21,
46 falcon road
SW11 2lr

70 × Retatrutide – £2,100 (£30 each)
10 × GLOW – £250 (£25 each)
...
Total: £3,300
(Payment awaiting) (Ready for delivery)
```

…and get back a formatted `.docx` invoice with tidy addresses, catalogue product
names, generated product codes, standardised statuses — and an **order total that
is actually correct** (`£3,300` in the note above is a typo; the real total is
`£3,200.00`).

**Pipeline:** `raw text → GitHub Models → strict JSON → Python math check → docxtpl → .docx`

---

## ✨ Features

| | |
|---|---|
| 🧠 **GitHub Models** | Uses the OpenAI-compatible GitHub Models endpoint. Auto-selects the best model your token can reach. |
| 🧾 **Handles any input format** | Bullets, tables, `x70`, `70 ×`, `@ £30`, `qty: 70`, missing unit prices, inline or multi-line addresses, `$`/`€`/`£`, and surrounding chit-chat. Missing figures are derived; a dropped line is flagged. |
| 🧮 **Arithmetic that's actually right** | Every line total and the order total are recomputed **in Python**, so a wrong figure never reaches the document — regardless of what the model says. |
| 🏷️ **Generated product codes** | `Retatrutide 40mg Pen → RETA40`, `NAD+ Nasal → NADN`, `BPC/TB-500 Nasal → BPCTB`. |
| 📚 **Catalogue expansion** | Shorthand becomes the full product name: `GLOW → Glow Pen`, `BPC Nasal → BPC-157 Nasal`. |
| ✨ **Standardised statuses** | `Payment awaiting → Awaiting Payment`, `Ready for delivery → Ready for Dispatch`. |
| 📮 **Address cleanup** | `flat 21, / 46 falcon road / sw11 2lr` → `Flat 21, 46 Falcon Road` / `SW11 2LR`. |
| 📄 **Your own template** | Bring any Jinja2-tagged `.docx`, or generate the included one. |
| 📥 **Word, PDF or JSON** | Download the invoice as `.docx`, as a ready-to-send `.pdf`, or grab the raw extracted JSON. |
| 🎨 **Polished UI** | Streamlit with a custom theme, live metrics, tabbed JSON/table views and one-click download. |

---

## 🚀 Quick Start

### 1. Clone

```bash
git clone https://github.com/<your-username>/ai-invoice-generator.git
cd ai-invoice-generator
```

### 2. Install — one command

**macOS / Linux / Git Bash**

```bash
bash setup.sh
```

**Windows CMD**

```cmd
setup.bat
```

This creates a virtual environment, installs everything in `requirements.txt`,
builds `template.docx`, and creates your `.env`.

<details>
<summary>Prefer to do it manually?</summary>

```bash
python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt
python create_template.py
cp .env.example .env          # Windows: copy .env.example .env
```
</details>

### 3. Add your GitHub token

Open `.env` and paste in a GitHub Personal Access Token:

```
GITHUB_TOKEN=github_pat_xxxxxxxxxxxxxxxxxxxx
```

Create one at **[github.com/settings/tokens](https://github.com/settings/tokens)**.
A fine-grained token needs the read-only **Models** account permission.
GitHub Models has a free tier — no billing setup required.

### 4. Run

```bash
streamlit run app.py
```

Your browser opens at `http://localhost:8501`. The sample order is pre-filled —
just hit **🚀 Generate Invoice**.

---

## 🤖 Model selection

The sidebar has a model picker. On **Auto** the app walks this list and uses the
first model your token can actually reach, skipping any it isn't entitled to:

```
openai/gpt-5  →  openai/gpt-4.1  →  openai/gpt-4.1-mini
              →  mistral-ai/mistral-medium-2505  →  openai/gpt-4o
```

`gpt-5` sits first so the app upgrades itself automatically if your account ever
gains access to it. Free-tier tokens currently land on **`gpt-4.1`**.

### Why this order?

Each candidate was benchmarked against the reference invoice (`INV-0070`),
scoring an exact field-by-field match across all 9 line items:

| Model | Match | Note |
|---|---|---|
| `openai/gpt-4.1` | **100%** | Exact reproduction |
| `openai/gpt-4.1-mini` | **100%** | Exact reproduction, faster |
| `mistral-ai/mistral-medium-2505` | **100%** | Exact reproduction |
| `openai/gpt-4o` | 98% | ❌ Got the corrected order total wrong (`£3,250.00` instead of `£3,200.00`) |
| `openai/gpt-5*`, `openai/o3*`, `o4-mini` | — | Not available on a free-tier token |

`gpt-4o` failing the arithmetic is precisely why the totals are **also** verified
in Python — see `verify_and_correct_totals()` in [`app.py`](app.py). Any figure the
model gets wrong is silently corrected and reported in the UI.

---

## 📄 The Word template

`create_template.py` generates a `template.docx` laid out like the reference
invoice. Open it in Word and restyle it however you like — just keep the tags.

**Header / footer fields**

```jinja2
{{ company_name }}      {{ invoice_number }}    {{ invoice_date }}
{{ customer_name }}     {{ customer_address }}  {{ reseller_info }}
{{ payment_status }}    {{ shipping_status }}   {{ order_total }}
```

**Products table** — a 5-column table with **four** rows:

| Row | Content |
|---|---|
| 1 | `CODE` · `PRODUCT` · `QTY` · `UNIT PRICE` · `LINE TOTAL` |
| 2 | `{%tr for item in items %}` |
| 3 | `{{ item.code }}` · `{{ item.product }}` · `{{ item.qty }}` · `{{ item.unit_price }}` · `{{ item.line_total }}` |
| 4 | `{%tr endfor %}` |

> ⚠️ **Write it as `{%tr for item in items %}` — no space after `{%`** — and put it
> in a row of its own. docxtpl matches the literal string `{%tr `; writing
> `{% tr ... %}` leaves the tag in the XML and rendering fails with
> `TemplateSyntaxError: Encountered unknown tag 'tr'`.

Multi-line values (the address) are converted from `\n` to `\a` before rendering,
which is how Word represents a soft line break inside a single paragraph.

### Rounded corners

Word table cells are always square, so the info panels are **VML rounded
rectangles** (`v:roundrect`) with their content inside the shape's text box.
`docxtpl` renders `{{ tags }}` inside `w:txbxContent` like any other text.

Two constraints are worth knowing if you edit `create_template.py`:

- **A text box inside a text box makes Word declare the file corrupt.** The
  status values were originally pills nested inside the status panel; they are
  now plain bold text for this reason. A *table* inside a text box is fine.
- Each row is **one full-width panel containing a two-column table**, not two
  side-by-side shapes. Separate shapes would need matching fixed heights and
  would clip any address that outgrew them.

---

## 📥 PDF export

The results panel offers **PDF**, **Word** and **JSON** downloads. There is no
reliable pure-Python `.docx → .pdf` renderer, so the app drives a real office
suite and picks whichever is installed:

| Platform | Requirement |
|---|---|
| Windows | Microsoft Word + `pywin32` (installed by `requirements.txt`) |
| macOS / Linux | LibreOffice (`brew install --cask libreoffice` / `apt install libreoffice`) |

Word is tried first — it renders the template's rounded VML panels exactly.
If neither is available the app still works normally; the PDF button is simply
replaced by a note telling you to export from Word yourself.

## 🗂 Project structure

```
├── app.py                # The Streamlit application
├── create_template.py    # Generates a correctly-tagged template.docx
├── requirements.txt      # Dependencies
├── setup.sh / setup.bat  # One-command setup
├── .env.example          # Token template — copy to .env
└── template.docx         # Generated by create_template.py
```

## 🧾 Extracted JSON shape

```json
{
  "invoice_number": "INV-0832",
  "invoice_date": "29/07/2026",
  "company_name": "ML Trading International",
  "customer_name": "Shaq Jordan",
  "customer_address": "Flat 21, 46 Falcon Road\nSW11 2LR",
  "reseller_info": "J/L",
  "payment_status": "Awaiting Payment",
  "shipping_status": "Ready for Dispatch",
  "items": [
    { "code": "RETA40", "product": "Retatrutide 40mg Pen",
      "qty": 70, "unit_price": "£30.00", "line_total": "£2,100.00" }
  ],
  "order_total": "£3,200.00",
  "currency_symbol": "£",
  "total_items": 9,
  "total_units": 120
}
```

---

## 🩺 Troubleshooting

| Problem | Fix |
|---|---|
| `GITHUB_TOKEN is not set` | Create `.env` from `.env.example` and add your token. |
| `unavailable_model` | Your account isn't entitled to that model — leave the picker on **Auto**. |
| `429 / rate limit` | GitHub Models free tier throttles requests. Wait a minute, or pick a `-mini` model. |
| `unknown tag 'tr'` | Your template uses `{% tr %}` with a space. Use `{%tr `, or regenerate with `python create_template.py`. |
| Address on one squashed line | The template tag must be `{{ customer_address }}` in its own paragraph. |

---

## 🔐 A note on the token

`.env` is git-ignored on purpose. Beyond the obvious, GitHub's push protection
actively blocks commits containing a Personal Access Token and will auto-revoke
any that slips through — so a committed token wouldn't work for the next person
anyway. Everyone who clones this repo supplies their own free token via `.env`.

---

## 📜 License

MIT
