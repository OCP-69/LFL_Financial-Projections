# Umsetzungsplan: LFL Financial Model in Claude Code
# Schritt-für-Schritt-Anleitung

---

## VORAUSSETZUNGEN

### Was du brauchst, bevor du startest:

1. **Claude Code installiert** (CLI-Tool von Anthropic)
   - Installation: `npm install -g @anthropic-ai/claude-code`
   - Prüfe Version: `claude --version`
   - Benötigt: Node.js 18+

2. **Anthropic API-Key**
   - Hinterlegt als Umgebungsvariable: `export ANTHROPIC_API_KEY=sk-ant-...`
   - Oder in `~/.claude/config.json`

3. **Projektordner anlegen**
   ```bash
   mkdir ~/lfl-scenario-engine
   cd ~/lfl-scenario-engine
   ```

4. **Excel-Template-Datei**
   - Kopiere `260304_LFL_SaaS_Startup_Financial_Model_v0_4.xlsx` in den Projektordner
   - Benenne sie um zu: `template_v0.4.xlsx`

---

## SCHRITT 1: Projektstruktur anlegen

Erstelle folgende Ordnerstruktur:

```bash
mkdir -p ~/lfl-scenario-engine/{templates,scenarios,reports,scripts}
cp 260304_LFL_SaaS_Startup_Financial_Model_v0_4.xlsx ~/lfl-scenario-engine/templates/template_v0.4.xlsx
```

Zielstruktur:
```
lfl-scenario-engine/
├── templates/
│   └── template_v0.4.xlsx          ← Original-Template (wird NIE verändert)
├── scenarios/
│   └── (hier landen generierte Szenario-Dateien)
├── reports/
│   └── (hier landen Delta-Berichte als .md)
├── scripts/
│   └── (hier liegen Hilfs-Skripte)
├── CLAUDE.md                        ← System-Prompt für Claude Code
└── requirements.txt
```

---

## SCHRITT 2: CLAUDE.md erstellen

Die Datei `CLAUDE.md` im Projektroot ist der System-Prompt, den Claude Code automatisch liest.

```bash
cd ~/lfl-scenario-engine
```

**Kopiere den gesamten Inhalt von `system-prompt-lfl-v2.md` in die Datei `CLAUDE.md`.**

Das ist der entscheidende Schritt: Claude Code liest `CLAUDE.md` automatisch beim Start und verwendet den Inhalt als Kontext für alle Interaktionen.

---

## SCHRITT 3: Python-Abhängigkeiten installieren

Erstelle `requirements.txt`:
```
openpyxl>=3.1.0
```

Claude Code hat Zugriff auf das Terminal und installiert Pakete bei Bedarf automatisch. Alternativ vorab:
```bash
pip install openpyxl
```

---

## SCHRITT 4: Claude Code starten und initialisieren

```bash
cd ~/lfl-scenario-engine
claude
```

Claude Code startet und liest automatisch die `CLAUDE.md`. Beim ersten Start gibst du folgenden Befehl ein:

```
Lies bitte das Template in templates/template_v0.4.xlsx aus und bestätige mir,
dass du die Modellstruktur korrekt erkennst. Liste alle Sheets, die Anzahl 
der editierbaren Inputs und die Formelketten auf.
```

**Erwartetes Ergebnis:** Claude Code liest die Datei, bestätigt die 9 Sheets und die ca. 85 editierbaren Input-Zellen.

---

## SCHRITT 5: Validierungslauf durchführen

Gib Claude Code folgenden Befehl:

```
Erstelle einen Validierungslauf: Kopiere das Template ohne Änderungen als 
"LFL_BM_Baseline_[Datum].xlsx" in den scenarios-Ordner. Lese dann die 
berechneten Werte für folgende KPIs aus (data_only=True):
- Total ARR Monat 12, 24, 36, 52
- Total Headcount Monat 12, 24, 36, 52
- EBITDA Monat 12, 24, 36, 52
- Ending Cash Monat 12, 24, 36, 52
- Runway Monat 12, 24, 36, 52
Speichere diese als Baseline-Referenz.
```

**WICHTIG:** Die `data_only=True`-Werte funktionieren nur, wenn die Datei vorher in Excel/Google Sheets geöffnet und gespeichert wurde (damit die berechneten Werte gecached sind). Falls die Werte `None` sind:

→ Öffne `template_v0.4.xlsx` einmal in Google Sheets oder Excel
→ Speichere sie erneut
→ Kopiere sie zurück in den templates-Ordner

---

## SCHRITT 6: Erstes Szenario erstellen

Teste mit einem einfachen Szenario:

```
Erstelle ein Szenario "HighChurn":
- Ändere die jährliche Churn Rate von 8% auf 15%
- Alle anderen Werte bleiben gleich
- Speichere die Datei und erstelle einen Delta-Bericht
```

Claude Code wird:
1. Das Template laden
2. `Inputs!B28` von 0.08 auf 0.15 ändern
3. Die Datei als `LFL_BM_HighChurn_[Datum].xlsx` speichern
4. Einen Bericht generieren

---

## SCHRITT 7: Sandbox-Szenario testen

```
Wechsle das aktive Szenario auf "stark" (Packaging-Fokus).
Erstelle die Datei und berichte, welche Inputs sich dadurch ändern.
```

Claude Code wird:
1. `00_Input_Sandbox!B1` auf "stark" setzen
2. Erklären, dass B20 (Startpreis) und B22 (Startmonat) sich per VLOOKUP ändern
3. Erklären, dass die KI-Strategie-Spalte F alle KI-Agent-Positionen auf 99 setzt

---

## SCHRITT 8: Komplexes Szenario

```
Erstelle ein Szenario "Conservative_NoAI":
- Szenario: "gering" (Automotive-Fokus)
- KI-Strategie: Setze ALLE Positionen in Spalte E auf "Fix" 
  (kein KI-Effekt auf Einstellungsplan)
- Churn Rate: 12%
- Preiserhöhung: 3% statt 8%
- Enterprise-Deals erst ab Monat 30
- Vergleiche mit dem Basis-Szenario
```

---

## SCHRITT 9: Iterativer Ausbau

Ab hier kannst du Claude Code interaktiv nutzen:

**Typische Befehle:**
```
# Schnelles Szenario
"Was passiert, wenn wir die Seed-Runde um 6 Monate verschieben?"

# Vergleich
"Vergleiche das Szenario 'stark' mit 'gering' - was sind die 3 größten Unterschiede?"

# Sensitivitätsanalyse
"Erstelle 3 Szenarien mit Churn 5%, 10%, 15% und vergleiche den Runway."

# Neue Sandbox-Parameter
"Füge in der Sandbox eine neue Zeile 'Cloud-Kosten-Faktor' hinzu mit Werten 0.8/1.0/1.5"
```

---

## SCHRITT 10: Erweiterte Features (optional)

Wenn die Grundfunktion steht, kannst du Claude Code bitten:

```
Erstelle ein Python-Skript scripts/calculate_model.py, das die gesamte 
Formellogik des Modells in Python nachrechnet. Das Skript soll:
1. Alle Inputs aus einer Excel-Datei lesen
2. Revenue, Costs, P&L, Cash Flow, Balance Sheet berechnen
3. Die Ergebnisse als JSON zurückgeben
So können wir Delta-Berichte OHNE Excel berechnen.
```

Das wäre der größte Qualitätssprung: Ein vollständiger Python-Rechenkern, der die Excel-Logik 1:1 abbildet und sofortige Vergleiche ermöglicht.

---

## ZUSAMMENFASSUNG: Minimale Schritte zum Start

| # | Was | Aufwand |
|---|-----|---------|
| 1 | Claude Code installieren (`npm install -g @anthropic-ai/claude-code`) | 2 min |
| 2 | API-Key setzen (`export ANTHROPIC_API_KEY=...`) | 1 min |
| 3 | Projektordner anlegen | 1 min |
| 4 | Template-Excel reinkopieren | 1 min |
| 5 | `CLAUDE.md` aus dem System-Prompt erstellen | 2 min |
| 6 | `claude` starten und "Lies das Template" | 5 min |
| 7 | Erstes Szenario testen | 5 min |
| **Total** | | **~17 min** |

---

## FEHLERBEHEBUNG

**Problem: `data_only=True` gibt None zurück**
→ Die Excel-Datei wurde nie in Excel geöffnet. Lösung: Einmal in Google Sheets/Excel öffnen und speichern.

**Problem: Formeln werden überschrieben**
→ Claude Code muss `load_workbook(data_only=False)` verwenden und darf nur Zellen in Spalte B des Inputs-Sheets und Sandbox-Zellen ändern.

**Problem: Claude Code kennt die Modellstruktur nicht**
→ Prüfe, ob `CLAUDE.md` im Projektroot liegt und ob Claude Code sie erkennt (`/status` oder "Was weißt du über das Modell?").

**Problem: Generierte Datei zeigt falsche Werte**
→ Die Datei muss nach dem Öffnen in Excel/Sheets einmal "Alles neuberechnen" (Strg+Shift+F9) durchlaufen.

---

*Umsetzungsplan v1.0 – 04.03.2026*
