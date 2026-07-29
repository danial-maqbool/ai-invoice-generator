# Setup Guide

A step-by-step guide for getting the AI Invoice Generator running on a Windows PC.
No prior experience needed — just follow the steps in order.

---

## Part 1 — One-time setup

You only ever do this once per computer.

### Step 1 — Install Python

Download from **https://www.python.org/downloads/** and run the installer.

> ⚠️ On the first screen of the installer, **tick the box that says
> "Add python.exe to PATH"** before clicking Install.
> If you miss it, the commands below will not work and you'll have to reinstall.

Python 3.9 or newer is fine.

### Step 2 — Install Git

Download from **https://git-scm.com/downloads** and run the installer.
Click Next through every screen — the defaults are all fine.

### Step 3 — Get your GitHub token

1. Go to **https://github.com/settings/tokens**
2. Click **Generate new token** → **Fine-grained token**
3. Under **Permissions → Account permissions**, find **Models** and set it to
   **Read-only**
4. Click **Generate token**, then **copy the token** — it starts with
   `github_pat_...`

> Copy it somewhere safe now. GitHub only shows it once.
> It's free — no billing or card required.

### Step 4 — Open a command prompt in the right folder

1. Open File Explorer and go to the folder where you want the app installed
   (e.g. `Documents`)
2. Click the **address bar** at the top, delete what's there
3. Type `cmd` and press **Enter**

A black command prompt window opens, already pointing at that folder.

### Step 5 — Download and install the app

Paste these commands into the command prompt, **one at a time**, pressing
**Enter** after each:

```
git clone https://github.com/danial-maqbool/ai-invoice-generator.git
```
```
cd ai-invoice-generator
```
```
setup.bat
```

`setup.bat` takes a couple of minutes. It installs everything the app needs.
Wait until you see **"Setup complete"** before continuing.

### Step 6 — Add your token

Open the settings file:

```
notepad .env
```

Notepad opens. Replace the whole contents with this single line, pasting your
own token after the `=` sign:

```
GITHUB_TOKEN=github_pat_your_token_here
```

Then **File → Save** and close Notepad.

> ⚠️ No spaces around the `=`. No quotes. No curly brackets `{ }`.
> It must look exactly like `GITHUB_TOKEN=github_pat_abc123...`

### Step 7 — Start the app

```
venv\Scripts\activate
```
```
streamlit run app.py
```

Your browser opens automatically at **http://localhost:8501**.

Setup is done. 🎉

---

## Part 2 — Normal use (every time after that)

1. Open File Explorer and go into the **`ai-invoice-generator`** folder
2. Click the **address bar**, type `cmd`, press **Enter**
3. Run these two commands:

```
venv\Scripts\activate
```
```
streamlit run app.py
```

> ⚠️ Don't skip `venv\Scripts\activate`. Without it you'll get
> *"'streamlit' is not recognized..."*

### Using the app

1. Paste your raw order text into the big box
2. Click **🚀 Generate Invoice**
3. Download the invoice as **PDF**, **Word** or **JSON**

**To stop the app:** press `Ctrl + C` in the command prompt, or just close the
window. Keep it open while you're using the app — closing it stops the app and
the browser page will say the connection was refused.

---

## Optional — PDF downloads

PDF export needs **Microsoft Word** or **LibreOffice** installed on the PC.

- If you already have Microsoft Word, nothing more to do.
- If not, install LibreOffice free from **https://www.libreoffice.org/download/**

Without either, the app still works normally — you just get Word and JSON
downloads instead of PDF.

---

## If something goes wrong

| What you see | What to do |
|---|---|
| `'python' is not recognized` | Python isn't on PATH. Reinstall Python and tick **"Add python.exe to PATH"**. |
| `'git' is not recognized` | Git isn't installed — go back to Step 2. |
| `'streamlit' is not recognized` | You skipped `venv\Scripts\activate`. Run it, then try again. |
| `GITHUB_TOKEN missing` in the app | The `.env` file is wrong. Run `notepad .env` and check Step 6 — no quotes, no brackets, no spaces. |
| `Too many requests` | You've hit the free usage limit. Wait about a minute and try again. If it keeps happening, pick **`openai/gpt-4.1-mini`** from the Model dropdown in the sidebar. |
| Browser says "connection refused" | The app isn't running. Make sure the command prompt is still open and you ran both commands from Part 2. |
| The page looks frozen or out of date | Press `Ctrl + F5` in the browser to force a refresh. |

---

## Usage limits

The app uses GitHub Models, which is free but rate-limited:

- roughly **10–15 requests per minute**
- roughly **50 invoices per day** on the default model

If you hit the limit, either wait, or switch the sidebar **Model** dropdown to
`openai/gpt-4.1-mini` — it has a higher daily allowance and produces identical
results on these invoices.
