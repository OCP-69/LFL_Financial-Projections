# LFL Financial Model v8 — IT-Abhängigkeits-Analyse

**Dokument:** Bewertung der IT-Spend- und Personalplan-Abhängigkeiten
**Basis:** `260307_LFL_BM_Vorlage_v8.xlsx` + `LFL_Benutzeranleitung_v8.pdf`
**Erstellt:** 10.03.2026
**Status:** Analyse-Bericht — Handlungsempfehlungen für Gründer

---

## 1. Executive Summary

Das Modell v8 zeigt eine solide Grundstruktur mit klarer Trennung von SaaS-Subscription- und Consulting-Revenue. Die IT-Kostenlogik ist in der richtigen Richtung modelliert, hat aber **vier kritische Lücken**, die das Modell strukturell verzerren:

1. **Kein erster Senior Engineer im Einstellungsplan** (Inputs Z48, Spalte B = leer): Die wichtigste Produktentwicklungsstelle fehlt. Alle Annahmen über Produktreife und First Revenue bauen auf einer Person auf, die formal nicht eingestellt ist.

2. **Kein Consulting-Personal modelliert**: Consulting-Revenue (Revenue Z28-29) wächst mit Kundenzahl × Wahrscheinlichkeit × Tage × Tagessatz — ohne Kapazitätslimit. Niemand liefert die Beratungstage. CTO und Gründer werden als unlimitierte Ressource implizit verbraucht.

3. **AI/ML API-Kosten sind zeitbasiert, nicht nutzungsbasiert**: `Inputs!B84 × (1.08)^(Monat-1)` eskaliert unabhängig von Kundenanzahl, Produktnutzung oder Projekten. Das überschätzt frühe Kosten und unterschätzt späte bei starkem Wachstum.

4. **First Revenue = strategische Annahme, kein Engineering-Constraint**: Der Startmonat der ersten zahlenden Kunden (Inputs B22, per VLOOKUP aus Sandbox) ist nicht aus der verfügbaren Engineering-Kapazität abgeleitet. Das Modell rechnet keine Entwicklerstunden für das MVP.

**Kernbotschaft an die Gründer:** Das Modell berechnet korrekt, WAS es kostet — aber nicht, ob die IT-Mannschaft rechtzeitig und ausreichend vorhanden ist, um den modellierten Umsatz zu liefern. Der strukturelle Zusammenhang „Mehr Projekte → mehr IT-Kapazität → mehr Kosten → mehr Umsatz" ist nur für die Subscription-Seite modelliert. Für Consulting und MVP-Entwicklung fehlen die Gegengewichte.

---

## 2. Modell-Architektur: IT-Abhängigkeiten im Überblick

```
Abhängigkeitsstruktur (Ist-Zustand v8):

Sandbox (Szenario: gering/normal/stark)
    ↓ VLOOKUP
Inputs (Einstellungsmonat, IT-Kosten, Consulting-Parameter)
    ↓                              ↓
Personalplan (Costs Z5-Z38)      Revenue-Treiber
    ↓                              ↓
TOTAL HEADCOUNT (Costs Z38)      Active Seats (Revenue Z8)
    ↓                              ↓
SaaS Tools (Costs Z84)    →      Cloud Variable (Costs Z82)
                                   ↓
                              Total Monthly Revenue
                                   ↓ (Subscription + Consulting)
                                P&L → Cash Flow → Balance Sheet

FEHLENDE Verbindungen:
  Engineering-Kapazität  ←✗→  First Revenue Month
  Consulting-Personal    ←✗→  Consulting Revenue
  Produktkomplexität     ←✗→  CS-Kapazität
  Nutzungsvolumen        ←✗→  AI/ML API Costs
```

---

## 3. IT-Spend Treiber — Arbeitsblatt Costs (Z80–Z87)

### 3.1 Cloud Infrastructure Basis — Costs!Z81

**Formel:** `=Inputs!$B$82` (für alle 52 Monate identisch)
**Wert:** €2.000/Monat (fix)
**Treiber:** Kein dynamischer Treiber. Konstante Grundlast.

**Bewertung:** Korrekt für die Frühphase. Ab M24+ (Enterprise-Kunden, mehr Datenvolumen) wird €2.000/Monat Grundlast zu niedrig für GPU-Inferenz und Vektordatenbanken. Kein Skalierungsfaktor eingebaut.

**Empfehlung:** Ab Seed-Phase (M17) einen zweiten Basiskostenpunkt einführen oder den Basiswert per VLOOKUP aus der Sandbox szenariobezogen steuern (ähnlich wie Büromiete-Upgrade bei M18).

---

### 3.2 Cloud Infrastructure Variable — Costs!Z82

**Formel:** `=Revenue!B8 × Inputs!$B$83`
**Treiber:** `Revenue!B8 = Total Active Seats` (der stärkste und korrekteste Treiber)
**Satz:** €80/Seat/Monat

**Kausalitätskette:**
```
Neue Kunden → mehr Seats (Revenue Z5-Z8)
                  ↓ direkte Formel
           Cloud Variable Costs (Costs Z82)
```

**Bewertung:** ✓ **Korrekt modelliert.** Dies ist die einzige IT-Kostenposition, die direkt mit dem Umsatzwachstum verknüpft ist. Bei 5 initialen Seats pro Kunde und €80/Seat entstehen €400/Monat pro neuem Kunden an Cloud-Zusatzkosten — das entspricht einer variablen COGS-Marge von ca. 23% auf MRR (bei €350/Seat/Monat Grundpreis im Normal-Szenario). Realistisch für AI-intensive Workloads.

**Kritische Frage für Gründer:** Ist der Satz €80/Seat/Monat für Packaging-Kunden (kleinere Datenmengen) und Automotive-Kunden (große CAD-Dateien, hohe Inferenzlast) identisch? Hier könnte ein Branchen-Faktor (Packaging: €50, Automotive: €120) die Marge realistischer abbilden.

---

### 3.3 AI/ML API Kosten — Costs!Z83

**Formel:** `=Inputs!$B$84 × (1 + Inputs!$B$85)^(Monat-1)`
**Parameter:** Basis €2.000/Mo, Wachstum 8%/Monat
**Treiber:** Ausschließlich die Zeit (Monatsnummer)

**Kostenentwicklung (berechnet):**
| Monat | AI/ML Kosten |
|-------|-------------|
| M1 | €2.000 |
| M6 | €2.939 |
| M12 | €5.036 |
| M24 | €12.669 |
| M36 | €31.868 |
| M52 | €107.233 |

**Bewertung:** ⚠️ **Kritische Schwäche.** Die 8%/Monat-Eskalation ist eine reine Zeitfunktion — sie wächst unabhängig davon, ob LFL 2 oder 200 Kunden hat, ob gerade ein Projekt aktiv ist oder nicht. Das führt zu:
- **Überschätzung früher Kosten** (M1-M8, keine Kunden, aber Kosten wachsen)
- **Unterschätzung später Kosten** bei starkem Projektwachstum ohne Skalierungsbegrenzung
- **Falscher COGS-Anteil**: AI/ML APIs sind direkte Produktionskosen (COGS), kein fixer OpEx

**Korrekte Logik wäre:** `=Basis × (1 + Wachstumsrate_per_Seat)^(TotalSeats[m]) + Basis_for_internal_dev`

**Empfehlung für Gründer (Entscheidung erforderlich):** Welcher Anteil der AI/ML-Kosten ist nutzungsbasiert (pro Seat/Anfrage) vs. fix für die Modellentwicklung? Diese Entscheidung bestimmt die wahre Bruttomarge bei Skalierung.

---

### 3.4 SaaS Tools Intern — Costs!Z84

**Formel:** `=Inputs!$B$86 + Costs!B38 × Inputs!$B$87`
**Treiber:** `Costs!B38 = TOTAL HEADCOUNT`
**Parameter:** Basis €400/Mo + €100/MA/Monat (inkl. KI-Coding-Tools)

**Kostenentwicklung (abhängig vom Szenario):**

Im Normal-Szenario (KI-Hebel = 6 Monate Verzögerung):
| Monat | Headcount | SaaS Tools |
|-------|-----------|-----------|
| M1 | 3 (Founder) | €700 |
| M8 | ~5 | €900 |
| M12 | ~7 | €1.100 |
| M24 | ~10 | €1.400 |
| M36 | ~14 | €1.800 |

**Bewertung:** ✓ **Korrekt modelliert.** Der headcount-basierte Treiber ist sinnvoll. €100/MA/Monat für KI-Coding-Tools (Cursor/Copilot) ist marktgerecht und ermöglicht gleichzeitig die KI-Hebelwirkung im Personalplan zu rechtfertigen.

---

### 3.5 Software-Lizenzen Dev — Costs!Z85 und Sicherheit — Costs!Z86

**Formel:** `=Inputs!$B$88/12` (€5.000/Jahr = €417/Monat, fix)
**Sicherheit:** `=Inputs!$B$89` (€3.500/Monat, fix)

**Bewertung:** ✓ Akzeptabel für die Frühphase. Ab TISAX-Zertifizierung (typisch M18-M24 für Automotive-Zulieferer) werden €3.500/Monat für Sicherheit knapp. Kein Upgrade-Mechanismus modelliert.

---

### 3.6 Zusammenfassung IT-Spend Treiber

| Kostenpunkt | Cells | Treiber | Korrektheit |
|-------------|-------|---------|-------------|
| Cloud Basis | Costs!Z81 | Fix | ✓ OK (früh); ab M24 zu niedrig |
| Cloud Variable | Costs!Z82 → Revenue!Z8 | Active Seats | ✓ Korrekt |
| AI/ML APIs | Costs!Z83 → Zeit | Zeitfaktor | ⚠️ Nicht nutzungsbasiert |
| SaaS Tools | Costs!Z84 → Costs!Z38 | Headcount | ✓ Korrekt |
| SW-Lizenzen | Costs!Z85 | Fix | ✓ OK |
| Sicherheit | Costs!Z86 | Fix | ⚠️ Kein TISAX-Eskalator |
| **FEHLEND** | — | Consulting-IT-Kapazität | ✗ Nicht modelliert |

**Einzige direkte Verbindung Umsatz → IT-Spend:** Cloud Variable (Z82) via Active Seats.

---

## 4. Personalplan IT — Timing, Rollen, Lücken

### 4.1 Einstellungsplan IT-Rollen (Ist-Zustand v8)

Alle Eintrittszeiten aus `Inputs-Sheet, Spalte B` (Normal-Szenario, KI-Hebel = +6 Monate):

| Rolle | Geplant (B) | Effektiv (F) | KI-Strategie | Bewertung |
|-------|-------------|-------------|-------------|-----------|
| CTO (Gründer) | M1 (Fix) | M1 | Fix | ✓ Kernrolle |
| **1. Senior Engineer** | **LEER** | **FEHLT** | Fix | ⛔ KRITISCH |
| 2. Senior Engineer | M10 | M16 (+6) | KI-Hebel | ⚠️ Sehr spät |
| 1. Junior Engineer | M9 | M15 (+6) | KI-Hebel | ⚠️ Spät |
| 1. ML/AI Engineer | M8 | M8 | Fix | ✓ Korrekt |
| 2. ML/AI Engineer | M14 | M20 (+6) | KI-Hebel | OK |
| 1. Sales Rep | M3 | M99 | KI-Agent | ✓ AI-SDR |
| 1. CS Manager | M12 | — | unklar | ⚠️ Zu spät |
| 1. Marketing Mgr | M11 | M17 (+6) | KI-Hebel | OK |
| 1. Product Mgr | M13 | M19 (+6) | KI-Hebel | OK |
| 3. Senior Engineer | M18 | M18 | Fix | ✓ |
| Finance Manager | M18 | M18 | Fix | ✓ |
| 2. CS Manager | M24 | M24 | Fix | OK |

**Gesamte Engineering-Kapazität bis First Revenue (Normal-Szenario = M8):**
- M1-M7: **Nur CTO** = 1 Entwickler
- M8: CTO + ML/AI Engineer = 2 Entwickler
- M8 = Monat des ersten zahlenden Kunden (Normal-Szenario)

**Kritische Frage:** Kann ein CTO allein in 7 Monaten ein MVP erstellen, das einen ersten zahlenden B2B-Kunden (Packaging/Automotive) überzeugt und den 6-Stufen-Onboarding-Prozess durchläuft?

---

### 4.2 Kritischer Befund: 1. Senior Engineer fehlt (Inputs!Z48, Spalte B)

**Zellbefund:** Inputs!B48 = `None` (kein Eintrittsmonat eingetragen)
**Formel in Costs:** `=IF(COLUMNS($B$6:B6) >= Inputs!$F$48, 1, 0)`

Da Inputs!B48 leer ist, bleibt Inputs!F48 = `IF(... Fix ..., B48, ...)` = leer → 0. Das bedeutet: Der erste Senior Engineer wird in keinem Monat als angestellt gezählt. **Die Headcount-Berechnung ist dadurch falsch**, da eine geplante Stelle fehlt.

**Handlungsbedarf:** Die Gründer müssen entscheiden und eintragen:
- Startmonat des 1. Senior Engineers (Empfehlung: M4-M6, parallel zu ML/AI Engineer)
- Wird diese Rolle als "Fix" oder "KI-Hebel" klassifiziert?

---

### 4.3 Consulting-Personal fehlt vollständig

**Problem:** Revenue!Z28 berechnet `Total Customers × Buchungswahrscheinlichkeit × Tage/Einsatz` ohne jede Kapazitätsbeschränkung. Im Normal-Szenario bedeutet das:

- **M4** (Consulting-Start): 0 Kunden → €0 Consulting-Revenue ✓
- **M8** (First SaaS): ~1 Kunde × 50% × 10 Tage = 5 Consulting-Tage/Monat → Wer liefert das?
- **M12**: ~3-5 Kunden × 50% × 10 Tage = 15-25 Tage/Monat → CTO + ML/AI Eng haben keine freie Kapazität mehr!
- **M18**: 8-12 Kunden × 50% × 10 Tage = 40-60 Tage/Monat → Vollzeit 2 Berater nötig

**Aktuelles Modell:** Consulting-Revenue wächst unbegrenzt. Keine Stelle liefert die Tage.

**Maximalkapazität ohne dedizierten Consulting-Engineer (Abschätzung):**
- CTO: 20% Zeit = ~4 Tage/Monat (Rest: Produktentwicklung + Management)
- ML/AI Engineer: 30% Zeit = ~6 Tage/Monat
- **Gesamtkapazität: ~10 Tage/Monat bis M18**

Ab M18 (mit 3. Senior Engineer + Finance Manager): Kapazität steigt, aber alle sind für Produktentwicklung eingeplant.

---

### 4.4 Customer Success als Flaschenhals

**Modelliert:** 1. CS Manager ab M12 (Normal-Szenario)
**Problem:**
- Onboarding-Aufwand: 60 Stunden/Kunde (Normal)
- 1 CS Manager = ca. 160 Arbeitsstunden/Monat = max. 2-3 neue Kunden/Monat onboarden
- Monatliches Seat-Wachstum: 5% = organische Expansion, aber neue Kunden kommen durch Sales (Enterprise ab M24)

**Kritisch:** Wer macht Onboarding M8-M12? Der CTO? Die Gründer? Das ist 60h/Kunde Aufwand ohne explizite Ressource.

**Szenario-Abhängigkeit (Onboarding-Aufwand aus Sandbox):**
- Gering: 120h/Kunde → bei 2 Kunden = 240h → braucht fast 1,5 FTE allein für Onboarding
- Normal: 60h/Kunde → handhabbar mit CTO + einem Gründer
- Stark: 30h/Kunde → efficient, Self-Service-optimiert

---

## 5. Consulting Revenue — Modelllogik und strukturelle Lücken

### 5.1 Formel-Kette (korrekt dokumentiert)

```
Szenarien_Analyse (Z5: Wskt., Z8: Tage, Z11: Startmonat)
    ↓ manuell übernommen
Sandbox (Z12: Wskt., Z11: Tage, Z15: Startmonat)
    ↓ VLOOKUP
Inputs (B138: Tagessatz, B139: Wskt., B140: Tage, B141: Startmonat)
    ↓ Formel
Revenue!Z28: Consulting-Tage/Monat = IF(m >= Startmonat, Kunden × Wskt. × Tage, 0)
Revenue!Z29: Consulting Revenue = Tage × Tagessatz
    ↓
P&L: Total Revenue (Z20)
```

**Positiv:** Die Datenfluss-Dokumentation in Inputs!A136 ist vorbildlich und zeigt, dass die Gründer den Mechanismus verstehen.

### 5.2 Fehlende Kausalitäten

**Lücke 1 — Kapazitätsdeckel fehlt:**
```
Ist:  Tage/Monat = Kunden × Wskt. × Tage/Einsatz (unbegrenzt)
Soll: Tage/Monat = MIN(Kunden × Wskt. × Tage/Einsatz, Verfügbare Berater-Kapazität)
```

**Lücke 2 — Qualitätskosten Consulting fehlen:**
Consulting erfordert:
- Reisekosten (€500-€2.000/Einsatz)
- Vor-Ort-Aufwand (ggf. Übernachtung)
- Dokumentationsaufwand
Diese sind nicht in COGS für Consulting modelliert. Die Consulting-COGS sind aktuell = €0.

**Lücke 3 — Revenue-Staffel fehlt:**
- Initialberatung (einmalig, höherer Aufwand) vs. Folgeberatung (geringer, wenn Produkt bekannt)
- Derzeit: jede Consulting-Interaktion = gleiche Tagessatz × gleiche Tage
- Realität: Wiederkehrende Kunden brauchen weniger Tage, Erstintegration braucht mehr

### 5.3 Was Gründer entscheiden müssen

Die folgende Tabelle zeigt die offenen Entscheidungen, die direkt in Sandbox/Inputs eingetragen werden müssen:

| Entscheidung | Aktuell | Empfehlung | Auswirkung |
|-------------|---------|------------|------------|
| Wer liefert Consulting M4-M12? | Nicht definiert | CTO (20%) + CCO (30%) | Revenue capped: ~10 Tage/Mo |
| Ab wann Consulting-Engineer einstellen? | Nicht modelliert | M12-M18 | +€65.000-€80.000/Jahr Personal |
| Consulting-COGS (Reise, etc.)? | €0 | ~15-20% des Consulting-Rev | Echte Marge = 80-85% |
| Buchungswahrsch. Normal (50%): realistisch? | Annahme | Aus Pilot-Daten validieren | ARR ±30% wenn falsch |
| Tage/Einsatz Normal (10 Tage): realistisch? | Annahme | Aus Pilot-Daten validieren | Revenue ±100% wenn falsch |

---

## 6. Umsatz → IT-Spend: Kausalitätsanalyse

### 6.1 Modellierte Abhängigkeiten (vollständig)

| Revenue-Treiber | IT-Cost-Effekt | Formel-Verbindung | Korrektheit |
|----------------|----------------|-------------------|-------------|
| Mehr Active Seats | Cloud Variable ↑ | Revenue!B8 → Costs!B82 | ✓ |
| Mehr Headcount | SaaS Tools ↑ | Costs!B38 → Costs!B84 | ✓ |
| Monat ↑ | AI/ML APIs ↑ | Zeitfunktion → Costs!B83 | ⚠️ |

### 6.2 Nicht modellierte Abhängigkeiten (fehlend)

| Revenue-Treiber | Fehlender IT-Cost-Effekt | Begründung |
|----------------|--------------------------|------------|
| Mehr Consulting-Projekte | Consulting-Personal-Kosten | Kein Personal modelliert |
| Enterprise-Deals ab M24 | Höhere Sicherheits-/Compliance-Kosten | TISAX-Anforderungen |
| Mehr Consulting-Projekte | Reisekosten COGS | Nicht in P&L |
| Höheres Nutzungsvolumen | AI/ML API-Kosten | Nicht nutzungsbasiert |
| Produktkomplexität | DevOps/SRE-Kosten | Keine DevOps-Rolle |
| Internationale Kunden (ab M30+) | Lokalisierungskosten | Nicht modelliert |

### 6.3 Kritischer Befund: IT-Invest → Revenue ist einseitig

Das Modell berechnet IT-Kosten als Funktion von Umsatz, aber **nicht** Revenue als Funktion von IT-Invest:
- Mehr Senior Engineers → Produkt fertig früher → First Revenue früher: **NICHT MODELLIERT**
- Mehr ML/AI Engineers → bessere KI-Features → höherer NRR: **NICHT MODELLIERT**
- Mehr CS-Personal → weniger Churn: **NICHT MODELLIERT**

Die Gründer entscheiden diese Zusammenhänge im Sandbox-Switch (Startmonat, Churn, NRR), aber das Modell berechnet sie nicht kausal aus dem Personalplan ab.

---

## 7. Kritische Feststellungen — Priorisiert

### KRITISCH (sofortiger Handlungsbedarf)

**K1: Inputs!B48 (1. Senior Engineer) = leer**
- **Zelle:** `Inputs!B48` (Eintrittsmonat 1. Senior Engineer)
- **Wirkung:** Costs!Z6 = permanent 0 → Headcount um 1 unterschätzt → Personal-Budget falsch → Cash-Flow falsch
- **Fix:** Eintrittsmonat eintragen (Empfehlung: M4-M6)

**K2: Consulting Revenue ohne Kapazitätsdeckel**
- **Zellen:** `Revenue!B28-B29` (Consulting-Tage/Revenue)
- **Wirkung:** Bis M18 überschätzt das Modell Consulting-Revenue um Faktor 3-5× vs. reale Kapazität
- **Fix:** `=MIN(Kunden × Wskt. × Tage, VerfügbareKapazität)` als neuer Deckel; neue Input-Variable „Max Consulting-Tage/Monat"

**K3: Kein Consulting-Personal im Hiring-Plan**
- **Zeile fehlt in:** `Inputs-Sheet`, `Costs-Sheet`
- **Wirkung:** Consulting-Marge erscheint als 100% — tatsächlich braucht jede Consulting-Stunde Personalkapazität
- **Fix:** Neue Rolle „Solutions Engineer/Consultant" ab M10-M14 in Inputs einfügen

---

### HOCH (innerhalb der nächsten Modelliteration)

**H1: AI/ML API Kosten zeitbasiert statt nutzungsbasiert**
- **Zelle:** `Inputs!B85` (8%/Monat Eskalation)
- **Wirkung:** Kosten eskalieren auch bei 0 Kunden; bei 100+ Kunden zu wenig
- **Fix:** Zweistufige Formel: Fix-Anteil (Entwicklungsmodell, intern) + Variable (pro Seat oder pro API-Call)

**H2: CS-Manager fehlt M8-M12 (kritische Onboarding-Lücke)**
- **Zelle:** `Inputs!B61` (1. CS Manager = M12)
- **Wirkung:** Wer führt Onboarding M8-M11 durch? (60h/Kunde im Normal-Szenario)
- **Fix:** Entweder CS auf M8-M10 vorziehen oder explizit modellieren, dass CTO 30% Onboarding macht (und damit Produktentwicklung verlangsamt wird)

**H3: Szenarien_Analyse Z7 (FTE-Annahme) vs. Einstellungsplan nicht abgestimmt**
- **Zelle:** `Szenarien_Analyse!D7` (Normal = 7 FTE, Stark = 12 FTE)
- **Wirkung:** Diese Werte stehen isoliert — sie fließen in keine Formel ein und werden nicht gegen den realen Headcount aus dem Personalplan geprüft
- **Fix:** Formel in Szenarien_Analyse!D7 = `=Costs!B38` (aktueller Headcount aus Personalplan) für direkten Abgleich

---

### MITTEL (strategische Modell-Erweiterung)

**M1: Cloud Basis kein Upgrade-Mechanismus**
- Analog zur Büromiete (ab M18 höhere Rate): Cloud-Basiskosten sollten ab M17-M24 eskalieren
- **Empfehlung:** `=IF(m >= 17, Inputs!B82_neu, Inputs!B82_alt)` in Costs!Z81

**M2: Sicherheitskosten kein TISAX-Eskalator**
- TISAX Level B (Automotive) kostet initial €15.000-€50.000 (Einmalaufwand) + höhere laufende Kosten
- **Empfehlung:** Einmalig im Finanzierungsmonat (M17, Seed-Close) als Professional-Service-Extrakosten

**M3: Consulting COGS fehlen**
- Consulting-Revenue erscheint in P&L als Fast-100%-Marge (nur Tagessatz, keine Kosten)
- **Empfehlung:** Reisekosten-Formel: `Consulting-Tage × 150 €/Tag` (Reisekostenbudget) als neue COGS-Zeile

**M4: Entwicklungs-Roadmap (Sheet) enthält unimplementierte P1-Items**
- Produkt-Tiers (Basic/Pro/Enterprise): Vom Modell noch nicht abgebildet
- Technologie-Stack-Faktor: Sandbox-Erweiterung möglich (Open Source -40% AI/ML Kosten)
- **Empfehlung:** Schritt 1: AI/ML-Stack-Faktor in Sandbox (Zeile 17) einbauen

---

## 8. Entscheidungen für die Gründer — Priorisiert

### Entscheidungspaket 1: Engineering-Kapazität (sofort, blockiert das Modell)

**Frage 1.1:** Wann startet der erste Senior Engineer?
- Option A: M4 (direkt nach Pre-Seed, parallel zu ML/AI Eng) → mehr MVP-Kapazität, höherer Burn
- Option B: M6 (nach erstem Pilot-Setup durch CTO) → geringerer früher Burn, riskanterer MVP-Zeitplan
- Option C: M1 (Mitgründer?) → günstigste Option wenn möglich

**Folge:** Eintrag in `Inputs!B48` + Prüfung ob das die First-Revenue-Annahme (Inputs!B22) rechtfertigt.

**Frage 1.2:** Ist der CTO in M1-M8 Vollzeit als Entwickler oder 50% auch als Geschäftsführer tätig?
- 100% Entwickler → MVP in 7 Monaten möglich (mit 1-2 Unterstützern)
- 50% Management → MVP in 14 Monaten (= Konservativ-Szenario logisch)

---

### Entscheidungspaket 2: Consulting-Strategie (beeinflusst Revenue um ±40%)

**Frage 2.1:** Wer liefert Consulting in M4-M12?
- Option A: Gründer (CEO/CCO) als Berater → kein zusätzliches Personal, aber Management-Zeit gebunden
- Option B: CTO als Teil-Berater (20%) → Produktentwicklung verlangsamt sich
- Option C: Sofort externen Freiberufler für Consulting → variable Kosten, kein Headcount

**Frage 2.2:** Ab wann wird ein dedizierter „Solutions Engineer" benötigt?
- Schwellenwert: ~10 Consulting-Tage/Monat = 1 Vollzeit-Berater
- Im Normal-Szenario tritt dieser Schwellenwert bei ca. 3-5 aktiven Kunden auf (M12-M16)
- **Empfehlung:** `Inputs!B** [neue Zeile]` Solutions Engineer ab M12 einplanen (€80.000/Jahr)

**Frage 2.3:** Welche Buchungswahrscheinlichkeit ist durch Pilot-Daten validiert?
- 50% (Normal) bedeutet: jeder zweite Lizenz-Kunde bucht auch Consulting
- Aus dem Dossier: Automotive-Kunden brauchen mehr IT-Begleitung → eher 70-80%
- Packaging-Kunden: Self-Service-afin → eher 20-30%
- **Empfehlung:** Buchungswahrscheinlichkeit branchenspezifisch aufteilen (nicht nur szenariobezogen)

---

### Entscheidungspaket 3: AI/ML Kostenmodell (beeinflusst Bruttomarge)

**Frage 3.1:** Welcher Anteil der AI/ML-API-Kosten ist nutzungsbasiert?
- Interne Modellentwicklung (fix, unabhängig von Kunden): ca. €1.000/Monat
- Kundenseitige Inferenz (variabel, pro Seat/Session): rest
- **Empfehlung:** Split-Parameter in `Inputs!B84_fix` (€1.000) + `Inputs!B84_var` (€X/Seat/Monat)

**Frage 3.2:** Ist 8%/Monat AI-Kostenwachstum realistisch oder zu hoch?
- M52-Wert bei 8%: €107.000/Monat für AI/ML allein
- Alternativ: 3%/Monat = €18.000/Monat bei M52 (deutlich realistischer)
- **Empfehlung:** Scenario-Wert in Sandbox einbauen: `AI-Stack-Faktor (Gering: 3% | Normal: 5% | Stark: 8%)`

---

## 9. Empfehlungen: Konkrete Modell-Änderungen

### Sofort-Maßnahmen (nächste Modellversion, v9)

**Änderung 1: 1. Senior Engineer Eintrittsmonat eintragen**
```
Zelle: Inputs!B48
Wert eintragen: 4 (empfohlen) oder nach Gründer-Entscheidung
KI-Strategie (Inputs!E48): bereits "Fix" — korrekt
```

**Änderung 2: Solutions Engineer Rolle ergänzen**
```
Neue Zeile in Inputs (nach dem letzten CS-Eintrag):
  Label: "Solutions Engineer / Consulting"
  Gehalt: Inputs!B** = 80.000
  Eintrittsmonat: B** = 12 (Normal)
  KI-Strategie: E** = "Fix" (unverzichtbar für Consulting-Delivery)
```

**Änderung 3: Consulting-Kapazitätsdeckel in Revenue!Z28**
```
Neue Input-Variable in Inputs: B142 = "Max. Consulting-Tage/Monat" (default: 10)
Neue Formel in Revenue!B28:
=IF(m >= Inputs!$B$141,
   MIN(B27 × Inputs!$B$139 × Inputs!$B$140, Inputs!$B$142),
   0)
```

**Änderung 4: AI/ML Kosten in Sandbox-Wachstum überführen**
```
Neue Sandbox-Zeile (Zeile 18):
  Label: "AI-Stack Wachstum/Monat"
  Gering: 0.03, Normal: 0.05, Stark: 0.08

Inputs!B85 VLOOKUP auf neue Sandbox-Zeile
(analog zu anderen VLOOKUP-Parametern)
```

### Mittelfristige Ergänzungen (v10, nach ersten Pilot-Daten)

**Änderung 5: Consulting COGS**
```
Neue COGS-Zeile in P&L (oder Costs):
  Label: "Consulting COGS (Reise & Aufwand)"
  Formel: =Revenue!B29 * 0.15  (15% der Consulting-Revenue)
```

**Änderung 6: Branchenspezifische Buchungswahrscheinlichkeit**
```
Neue Inputs-Zeile:
  B143: Anteil Automotive (aus Sandbox, heute Z4: 0.1/0.4/0.2)
  B144: Buchungswahrsch. Automotive = 0.7
  B145: Buchungswahrsch. Packaging = 0.3

Revenue!B28: Gewichtete Wskt. = B143 × B144 + (1-B143) × B145
```

**Änderung 7: Cloud Basis-Upgrade analog zur Büromiete**
```
Neues Input: B82b = "Cloud Basis ab Seed" = 4.000
Neues Input: B82_upgrade_Monat = 17

Costs!B81: =IF(m >= Inputs!B82_upgrade_Monat, Inputs!B82b, Inputs!B82)
```

---

## 10. Bewertungszusammenfassung

### Was gut ist (v8 gegenüber v0.4 verbessert)

- ✓ Consulting-Revenue als eigener Stream mit vollständigem Datenfluss (Szenarien → Sandbox → Inputs → Revenue → P&L)
- ✓ KI-Strategie im Personalplan (Fix/KI-Hebel/KI-Agent) gut durchdacht
- ✓ Cloud Variable direkt an Active Seats gekoppelt
- ✓ Szenarien_Analyse mit dokumentierten Treibern
- ✓ Entwicklungs-Roadmap-Sheet als Prioritätsliste für Modellerweiterungen
- ✓ AI-Personal-Hebel als Sandbox-Parameter szenariobezogen steuerbar

### Was fehlt (kritisch für Investoren-Präsentation)

- ✗ Engineering-Kapazität ist nicht kausal mit First Revenue verknüpft
- ✗ Consulting-Kapazitätsdeckel fehlt → Revenue ist strukturell überschätzt
- ✗ 1. Senior Engineer nicht eingetragen → Headcount-Kalkulation fehlerhaft
- ✗ Keine Consulting-COGS → Marge zu optimistisch
- ✗ AI/ML Kosten nicht nutzungsbasiert → COGS-Marge bei Skalierung falsch
- ✗ CS-Bottleneck nicht modelliert → Churn-Annahmen zu optimistisch bei schnellem Wachstum

### Kalibrierungs-Prioritäten für die Gründer

| Priorität | Entscheidung | Auswirkung auf ARR M24 |
|-----------|-------------|----------------------|
| 1 | 1. Senior Engineer eintragen (B48) | Headcount korrekt, Burn präzise |
| 2 | Consulting-Kapazitätsdeckel (max. Tage/Mo) | Revenue -30% bis -60% in frühen Monaten |
| 3 | AI/ML Szenario-Wachstum (3% vs. 8%) | EBITDA ±€20K/Monat bei M24 |
| 4 | Solutions Engineer ab M12 | Kosten +€80K/Jahr; Revenue realistisch |
| 5 | Buchungswahrsch. nach Branche aufteilen | Revenue-Qualität, nicht Quantität |

---

*Analyse erstellt: 10.03.2026 | Basis: LFL_BM_Vorlage_v8.xlsx (260307) | Branch: claude/analyze-lfl-bmw-template-Dfcl4*
