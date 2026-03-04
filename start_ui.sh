#!/bin/bash
# LFL Financial Projections — UI starten
# Verwendung: ./start_ui.sh [PORT]

PORT=${1:-8501}
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Abhängigkeiten prüfen / installieren
if ! python3 -c "import streamlit" 2>/dev/null; then
    echo "Installiere Abhängigkeiten..."
    pip install -r "$SCRIPT_DIR/requirements.txt" -q
fi

# API-Key prüfen
if [ -z "$ANTHROPIC_API_KEY" ]; then
    echo "HINWEIS: ANTHROPIC_API_KEY nicht gesetzt."
    echo "Der Assistent-Tab funktioniert nur mit gesetztem API-Key:"
    echo "  export ANTHROPIC_API_KEY=sk-ant-..."
    echo ""
fi

echo "Starte LFL Financial Projections UI auf http://localhost:$PORT"
echo "Zum Beenden: Strg+C"
echo ""

cd "$SCRIPT_DIR"
streamlit run scripts/ui_app.py \
    --server.port "$PORT" \
    --server.headless true \
    --theme.base dark \
    --theme.primaryColor "#7c6af7" \
    --theme.backgroundColor "#1e1e2e" \
    --theme.secondaryBackgroundColor "#181825" \
    --theme.textColor "#cdd6f4"
