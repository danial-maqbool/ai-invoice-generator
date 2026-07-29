# Setup Guide

## One Time Setup:

1. Install Python 3.9 or better if not already installed — tick **"Add python.exe to PATH"** in the installer.
2. Install Git from https://git-scm.com/downloads (defaults are fine).
3. Go to https://github.com/settings/tokens → Generate new token (fine-grained) → grant the read-only Models permission. Free tier, no billing needed.
4. Go to a folder you want to download the files in, in the address bar of that folder, click it, erase the text, write 'cmd' and press enter, a command prompt will open.
5. Paste these commands and enter:

        git clone https://github.com/danial-maqbool/ai-invoice-generator.git
        cd ai-invoice-generator
        setup.bat
        notepad .env

6. Notepad opens — paste your github API key, then save and close:

        GITHUB_TOKEN=github_pat_your_token_here

   (no quotes, no brackets, no spaces around the `=`)

7. Run these commands in the same command prompt window and it will open in browser:

        venv\Scripts\activate
        streamlit run app.py

---

## For normal use:

1. Open command prompt from that folder, by typing `cmd` in address bar.
2. Run these commands and it will open in browser:

        venv\Scripts\activate
        streamlit run app.py

   (`venv\Scripts\activate` is required, or you get "'streamlit' is not recognized")

---

PDF downloads need Microsoft Word or LibreOffice installed. Without either, Word and JSON downloads still work.

Free tier allows ~50 invoices/day. If you hit "Too many requests", wait a minute or pick `openai/gpt-4.1-mini` from the Model dropdown.
