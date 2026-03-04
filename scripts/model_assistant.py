"""
LFL Model Assistant — Anthropic API Integration
Wird vom Streamlit UI (ui_app.py) importiert.
"""

import os
import anthropic
from pathlib import Path

ROOT = Path(__file__).parent.parent

# .env Datei laden falls vorhanden (Windows-freundlich)
_env_file = ROOT / ".env"
if _env_file.exists():
    for line in _env_file.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, _, val = line.partition("=")
            os.environ.setdefault(key.strip(), val.strip())

# System-Prompt aus CLAUDE.md laden
_claude_md = ROOT / "CLAUDE.md"
if _claude_md.exists():
    MODEL_CONTEXT = _claude_md.read_text(encoding="utf-8")
else:
    MODEL_CONTEXT = "Du bist ein Finanzmodell-Assistent für das LoopforgeLab Financial Projection Model."

SYSTEM_PROMPT = f"""
{MODEL_CONTEXT}

---

## DEINE ROLLE IM UI-KONTEXT

Du wirst als interaktiver Assistent im Streamlit-Interface aufgerufen.
Der Nutzer kann dir Fragen zum Finanzmodell stellen. Du hast Zugriff auf:
- Die vollständige Modellarchitektur (oben beschrieben)
- Die aktuellen Baseline-Werte (werden im User-Message als Kontext übergeben)

Antworte präzise, auf Deutsch, mit konkreten Zahlen wo möglich.
Erkläre die Auswirkungen auf KPIs (ARR, Burn Rate, Runway, EBITDA, Ending Cash).
Wenn der Nutzer nach einem Szenario fragt, beschreibe die konkreten Zell-Änderungen.
Bleibe sachlich und zahlenorientiert. Max. 3-4 Absätze pro Antwort.
"""

def get_assistant_response(
    user_message: str,
    history: list[dict],
    model_context_summary: str,
) -> str:
    """
    Ruft die Anthropic API auf und gibt die Antwort als String zurück.

    Args:
        user_message: Aktuelle Nutzerfrage
        history: Bisheriger Chat-Verlauf [{"role": "user"|"assistant", "content": "..."}]
        model_context_summary: Aktuelle Modell-Werte als Kurztext
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        return (
            "**API-Key fehlt.** Setze die Umgebungsvariable:\n\n"
            "```bash\nexport ANTHROPIC_API_KEY=sk-ant-...\n```\n\n"
            "Dann Streamlit neu starten."
        )

    client = anthropic.Anthropic(api_key=api_key)

    # Kontext-Nachricht einbauen
    context_msg = (
        f"**Aktueller Modell-Zustand:**\n```\n{model_context_summary}\n```\n\n"
        f"**Frage:** {user_message}"
    )

    # Messages für die API aufbauen
    messages = []
    for msg in history[-8:]:  # Max. 8 vergangene Nachrichten
        if msg["role"] in ("user", "assistant"):
            messages.append({"role": msg["role"], "content": msg["content"]})

    messages.append({"role": "user", "content": context_msg})

    try:
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1024,
            system=SYSTEM_PROMPT,
            messages=messages,
        )
        return response.content[0].text
    except anthropic.AuthenticationError:
        return "**Authentifizierungsfehler.** Bitte prüfe deinen `ANTHROPIC_API_KEY`."
    except anthropic.RateLimitError:
        return "**Rate-Limit erreicht.** Bitte warte kurz und versuche es erneut."
    except Exception as e:
        return f"**API-Fehler:** {e}"
