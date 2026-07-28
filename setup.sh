#!/usr/bin/env bash
# One-command setup for the AI Invoice Generator.
#   bash setup.sh
set -e

PY=$(command -v python3 || command -v python)

echo "==> Creating virtual environment (venv/)"
"$PY" -m venv venv

if [ -f venv/bin/activate ]; then
    # macOS / Linux
    source venv/bin/activate
else
    # Windows (Git Bash)
    source venv/Scripts/activate
fi

echo "==> Installing dependencies"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo "==> Building template.docx"
python create_template.py

if [ ! -f .env ]; then
    cp .env.example .env
    echo "==> Created .env — open it and paste in your GITHUB_TOKEN"
fi

cat <<'EOF'

------------------------------------------------------------
Setup complete.

1. Put your GitHub token in .env:
       GITHUB_TOKEN=ghp_xxxxxxxxxxxx

2. Activate the environment:
       source venv/bin/activate      # macOS / Linux
       venv\Scripts\activate         # Windows

3. Run the app:
       streamlit run app.py
------------------------------------------------------------
EOF
