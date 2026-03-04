# LFL Financial Projections — Detailbericht: Szenario KONSERVATIV

**Dokument:** Szenario-Analyse, Schlüsseltreiber & Phasenentwicklung
**Basis:** 260304_LFL-Financial-Planning-and-Carbon-Case.pdf
**Erstellt:** 04.03.2026
**Status:** Szenario-Entwurf – offene Punkte aus Part G des Dossiers noch nicht final entschieden

---

## 1. Executive Summary

Das Konservativ-Szenario beschreibt einen **kapitaleffizienten, risikominimierenden Aufbau** von LFL über 52 Monate. Im Mittelpunkt steht Marktvalidierung vor Skalierung: Erst Produktreife beweisen, dann wachsen. Die wesentlichen Konsequenzen dieser Strategie:

- **Erster zahlender Kunde:** Monat 17 (September 2027) — nach 17 Monaten ohne Umsatz
- **ARR am Ende Seed-Phase (M29):** 180.000–300.000 EUR (20–25 SMB-Kunden)
- **ARR am Ende Series A (M42):** 450.000–750.000 EUR (inkl. erste Enterprise-Deals)
- **ARR am Ende Projektionszeitraum (M52):** 700.000–1.100.000 EUR
- **Headcount M52:** 18–22 Vollzeitstellen (kein KI-Delay beim Hiring)
- **Gesamtfinanzierungsbedarf:** 15,5 Mio. EUR (Pre-Seed 1,5M + Seed 4M + Series A 10M)
- **Kernrisiko:** 17 Monate ohne Umsatz = sehr langer Cash-Drain vor erstem Revenue-Signal

### Stärken des Szenarios
- Realistisches Risikoprofil für einen ungetesteten B2B-Markt
- Kapitaldisziplin in der frühen Phase
- Ausreichend Zeit für Produktreife und Marktverständnis

### Schwächen des Szenarios
- Sehr spät erster Umsatz → höhere Investoranforderungen für Proof-of-Revenue
- Preispositionierung (200 EUR/Seat/Monat) lässt Marge liegen
- Kein KI-Hebel → überproportionaler Personalaufbau
- 15% Churn bedeutet: ein Drittel der Kunden wird jährlich ersetzt — hoher Akquisitionsaufwand

---

## 2. Schlüsseltreiber — Abhängigkeiten & Hebel

### 2.1 Treiber 1: Zeit bis zum ersten zahlenden Kunden (kritischster Einzelfaktor)

**PDF-Quelle:** `months_to_first_revenue: ~17` — explizit als "most critical overall variable" markiert (G1.3)

**Konservative Annahme:** Monat 17 = September 2027, parallel zum Seed-Close

**Kausalitätskette:**
```
Pre-Seed-Funding (M6)
  → MVP-Entwicklung (M6-M14)
  → 3-5 Pilot-Kunden für Validierung (M6-M16, kein Revenue-Fokus)
  → Produktreife + Onboarding-Paket fertig (M16-M17)
  → Erste Zahlungsbereitschaft (M17)
  → Seed-Funding ermöglicht Sales-Team (M17)
  → Skalierter Revenue-Aufbau (M18+)
```

**Sensitivität:** Eine Verschiebung um ±3 Monate verändert den Cashbedarf um ca. ±120.000 EUR (bei 40K EUR/Monat Burn in Seed-Vorläuferphase). Jeder Monat früher = 120K EUR weniger Finanzierungsbedarf ODER 3 Monate mehr Runway.

**Abhängigkeiten:**
- Produktentwicklungsgeschwindigkeit (CTO + 2 Senior Engineers)
- Qualität des Pilot-Programms (reines Validierungs-PoC vs. zahlender Pilot)
- Onboarding-Paket Fertigstellung (6-Stufen-Prozess laut PDF)
- Bereitschaft der Packaging/Maschinenbau-Kunden zur Datenintegration

---

### 2.2 Treiber 2: Preis pro Seat (direkter ARR-Multiplikator)

**Konservative Annahme:** 200 EUR/Seat/Monat = 2.400 EUR/Seat/Jahr

**Berechnung Standard-Kunde:**
| Metrik | Wert | Berechnung |
|--------|------|-----------|
| Seats pro Kunde | 5 | PDF avg_seats_per_customer |
| Preis/Seat/Monat | 200 EUR | Benchmark-Low |
| MRR pro Kunde | 1.000 EUR | 5 × 200 |
| ARR pro Kunde | 12.000 EUR | 1.000 × 12 |
| ARR-Abstand zum PDF-Referenz | −9.000 EUR | PDF: 21.000 EUR |

**Abhängigkeit vom ROI-Nachweis:** Der Preis ist direkt an den nachweisbaren Mehrwert geknüpft. Olaf betont im Dossier: "Die CO₂-Impact-Story muss kausal erklärt werden." Bei fehlendem ROI-Nachweis ist 200 EUR/Monat das Maximal-Erreichbare. Mit klarem ROI (Stunden-Einsparung × Personalkostensatz > Lizenzkosten) ist 350 EUR (PDF-Referenz) oder höher möglich.

**ROI-Formel (PDF):** ROI = (Hours_saved × Personnel_cost_rate) − License_cost
Bei 200 EUR/Seat/Monat: Break-Even bei **1,7 eingesparten Arbeitsstunden/Monat/Seat** (Annahme: 120 EUR/Stunde Ingenieur-Stundensatz in Packaging-Industrie).

---

### 2.3 Treiber 3: Monatliche Seat-Wachstumsrate (Compounding-Hebel)

**Konservative Annahme:** 3% pro Monat = ~42% pro Jahr (Zinseszinseffekt)

**Wachstumspfad:**
| Monat | Seats (kumulativ) | MRR | ARR |
|-------|-----------------|-----|-----|
| M17 (Start) | 15 | 3.000 EUR | 36.000 EUR |
| M20 | 16,4 | 3.280 EUR | 39.360 EUR |
| M24 | 18,4 | 3.680 EUR | 44.160 EUR |
| M29 | 21,4 | 4.280 EUR | 51.360 EUR |
| M36 | 26,3 | 5.260 EUR | 63.120 EUR |
| M42 | 31,4 | 6.280 EUR | 75.360 EUR |
| M52 | 42,3 | 8.460 EUR | 101.520 EUR |

> **Wichtiger Hinweis:** Diese Berechnung bildet nur das organische Seat-Wachstum bestehender Kunden ab. Der Modell-Parameter beschreibt Seat-Expansion innerhalb bestehender Accounts, NICHT Neukunden. Für Neukunden-Akquisition ist der Enterprise-Kanal (ab M30) der zweite Treiber.

**Problem im Konservativ-Szenario:** SMB-Only bis M30 + nur 3% Seat-Growth = sehr begrenzter ARR ohne Enterprise-Beitrag. Der ARR bleibt bis M29 unter 52K EUR/Jahr aus dem Seat-Kanal.

---

### 2.4 Treiber 4: Churn Rate (ARR-Erosion)

**Konservative Annahme:** 15% jährliche Churn Rate

**Implikation:**
- Bei 50K ARR: jährlicher Verlust durch Abwanderung = 7.500 EUR/Jahr
- Um ARR netto zu wachsen, muss Neuwachstum > 15% des Bestands-ARR sein
- NRR von 105% bedeutet: Bestandskunden zahlen 5% mehr — aber 15% wandern ab
- **Netto-ARR-Wachstum aus Bestandskunden:** 105% − 15% = −10% der ARR-Base
  → **Das Konservativ-Szenario verliert netto ARR durch Churn!** Wachstum nur durch Neukunden.

**Kausalzusammenhang:**
```
Fehlender ROI-Nachweis / unklare CO₂-Kausalität (Olaf-Kritik)
  → Kunden zweifeln an Mehrwert nach 12 Monaten
  → Nicht-Verlängerung bei Vertragsende
  → 15% Churn
  → Neukunden-Akquisition muss Abwanderung kompensieren UND Wachstum liefern
  → Erhöhter Sales-Aufwand und CAC
```

---

### 2.5 Treiber 5: Enterprise-Deals (Umsatz-Katalysator ab M30)

**Konservative Annahme:** 1 Deal/Quartal × 25.000 EUR ARR ab M30

**Kumulativer Enterprise-ARR:**
| Monat | Deals kumulativ | Enterprise ARR | Anteil am Gesamt-ARR |
|-------|----------------|----------------|---------------------|
| M30 | 0 | 0 EUR | 0% |
| M33 | 1 | 25.000 EUR | ~29% |
| M36 | 2 | 50.000 EUR | ~44% |
| M42 | 4 | 100.000 EUR | ~57% |
| M52 | 7 | 175.000 EUR | ~63% |

Enterprise-Deals werden bis M52 der **dominante Umsatztreiber**, obwohl sie erst spät beginnen. Das zeigt die strukturelle Wichtigkeit dieser Pipeline.

---

## 3. Personalaufbau & Rollenabhängigkeiten

### 3.1 Philosophie im Konservativ-Szenario

**KI-Hebel = 0:** Alle Rollen werden zum ursprünglich geplanten Zeitpunkt eingestellt. Keine Verzögerung durch KI-Werkzeuge. Dies entspricht einem **traditionellen SaaS-Aufbau** mit vollständigem Funktionsteam.

### 3.2 Hiring-Roadmap nach Phase

#### Pre-Seed (M1–M17): Kernteam
| Monat | Rolle | Gehalt/Jahr | Begründung |
|-------|-------|------------|-----------|
| M1 | CEO | 72.000 EUR | Gründer, Strategic + Sales Lead |
| M1 | CTO | 72.000 EUR | Gründer, Produkt + Tech |
| M5 (Pre-Seed Close) | Senior Engineer | 90.000 EUR | Core-Produktentwicklung |
| M8 | Senior Engineer #2 | 90.000 EUR | Skalierung MVP |

**Personalkosten Pre-Seed:**
- M1–M4: (72K + 72K) / 12 × 1,21 = 14.520 EUR/Monat
- M5–M8: + Senior Eng: 14.520 + (90K/12×1,21) = 14.520 + 9.075 = 23.595 EUR/Monat
- M9–M17: +Senior Eng #2: 23.595 + 9.075 = 32.670 EUR/Monat

#### Seed (M17–M29): Go-to-Market Team
| Monat | Rolle | Gehalt/Jahr | Begründung |
|-------|-------|------------|-----------|
| M17 (Seed Close) | CCO | 72.000 EUR | Commercial Lead, erster Sales-Aufbau |
| M17 | Sales Representative | 66.000 EUR | Pipeline-Aufbau Packaging/Machinery |
| M20 | Customer Success | 60.000 EUR | Onboarding der ersten Kunden |
| M23 | Marketing Manager | 66.000 EUR | Content + Events (primärer Lead-Kanal) |
| M26 | ML/AI Engineer | 110.000 EUR | Produkterweiterungen, KI-Features |

**Personalkosten Seed (Ende, M29, ~7 Personen):**
Summe Jahresgehälter: 72+72+90+90+72+66+60 = 522K EUR/Jahr
Mit Lohnnebenkosten (21%): 522K × 1,21 = 631.620 EUR/Jahr = **52.635 EUR/Monat**

**Bottleneck Customer Success:** Das PDF identifiziert explizit: *"As customer count grows, demand for customer success, training, implementation, and sales support grows proportionally."* Bei 25–30 Kunden und 6-stufigem Onboarding-Prozess (Datenanalyse → Integration → Normalisierung → Training → Custom → Community) wird 1 Customer-Success-Person zur Engpassstelle.

#### Series A (M30–M42): Skalierung & Professionalisierung
| Monat | Rolle | Gehalt/Jahr | Begründung |
|-------|-------|------------|-----------|
| M30 | Finance Manager | 70.000 EUR | Series A Investor Reporting |
| M32 | Junior Engineer | 66.000 EUR | Tech-Skalierung |
| M34 | Sales Rep #2 | 66.000 EUR | Enterprise-Pipeline |
| M36 | Customer Success #2 | 60.000 EUR | Wachsende Kundenbasis |
| M38 | Office/Admin | 48.000 EUR | Back-Office-Entlastung |
| M40 | Product Manager | 78.000 EUR | Produkt-Roadmap |

**Personalkosten Series A (Ende, M42, ~13 Personen):**
Summe Jahresgehälter: ~880K EUR/Jahr
Mit Lohnnebenkosten: ~1.065K EUR/Jahr = **88.750 EUR/Monat**

#### Scale Phase (M43–M52): Weitere Expansion
| Monat | Rolle | Gehalt/Jahr | Begründung |
|-------|-------|------------|-----------|
| M43 | Sales Rep #3 | 66.000 EUR | Enterprise-Fokus |
| M45 | Customer Success #3 | 60.000 EUR | Skalierung |
| M47 | Senior Engineer #3 | 90.000 EUR | Platform-Reife |
| M49 | Marketing #2 | 66.000 EUR | Demand Generation |

**Personalkosten M52 (17–22 Personen):**
Gesamtpersonalkosten: ~1.100.000–1.320.000 EUR/Jahr = **91.000–110.000 EUR/Monat**

### 3.3 Kritische Rollenabhängigkeiten

**CEO → Revenue:** Im Konservativ-Szenario ist der CEO bis zum CCO-Hire (M17) der einzige Sales-Treiber. Alle ersten 3–5 Pilot-Kunden hängen direkt von der CEO-Netzwerk und Überzeugungskraft ab.

**CTO → Produktreife:** Der Zeitpunkt des ersten Umsatzes (M17) hängt direkt vom CTO und den 2 Senior Engineers ab. Verzögerungen hier = Verschiebung des Revenue-Starts.

**Customer Success → Churn:** Die 15% Churn-Annahme könnte auf 10–12% gesenkt werden, wenn Customer Success früher (M17 statt M20) eingestellt wird und proaktives Onboarding betreibt.

---

## 4. Umsatzannahmen & ihre Abhängigkeiten

### 4.1 Umsatzquellen (nach PDF)

| Stream | Status | Abhängigkeit | Konservativ-Beitrag |
|--------|--------|-------------|-------------------|
| SaaS-Lizenzen (Seats) | Primär | Produktreife + Sales | Hauptumsatz ab M17 |
| Consulting/Implementierung | Sekundär | Onboarding-Paket | Gelegentlich ab M20 |
| Community/Enablement | In Diskussion | Kundenbasis >30 | Nicht modelliert |
| Performance-Fees | Deferred | CO₂-Messung | Nicht in M1-M52 |

### 4.2 Preis-Volumen-Gleichung

```
ARR = Kunden × Seats/Kunde × Preis/Seat/Jahr × (1 + Preiserhöhung)^t × (1 - Churn)
    + Enterprise-Deals × Enterprise-ARR
```

**Konservativ M42 (Beispielrechnung):**
- SMB: ~32 Seats (organisches Wachstum, 3%/Monat ab M17) × 2.400 EUR/Seat ≈ 77K ARR
- Enterprise: 4 Deals × 25K = 100K ARR
- **Gesamt: ~177K ARR**

> **Hinweis:** Diese Schätzung berücksichtigt nicht den vollständigen Kunden-Akquisitions-Prozess. Tatsächliche Zahlen entstehen durch Excel-Neuberechnung (Formeln). Die Modellwerte liefern die berechneten Zeitreihenwerte für alle 52 Monate.

### 4.3 Kritische Umsatz-Abhängigkeiten

**Abhängigkeit 1: Zahlungsbereitschaft der Zielkunden**
Packaging-Maschinenbau-Kunden (primärer Markt: 2.255 Unternehmen nach Rene) sind B2B-Industrieunternehmen. Ihr Kaufentscheidungsprozess:
1. Identifikation des Problems (Effizienz/CO₂-Kosten)
2. Interne Freigabe (budget cycle, Q4-lastig)
3. Pilotphase (3–6 Monate, kostenlos oder reduziert)
4. Vollvertrag

**Risiko:** Bei 200 EUR/Seat/Monat könnte die interne Freigabe einfacher sein (unter Einkaufsrichtlinie), aber das ROI-Signal kommt erst nach 6–12 Monaten Nutzung.

**Abhängigkeit 2: Consulting-Revenue als Stabilisator**
Laut PDF ist Consulting (Datenintegration, Training, Customizing) ein Sekundärstrom. Im Konservativ-Szenario ist dieser nicht explizit modelliert. Könnte aber in der Seed-Phase 15.000–30.000 EUR pro Enterprise-Onboarding einbringen und den Cash-Burn reduzieren.

**Abhängigkeit 3: Preiserhöhung 5%/Jahr**
Die 5%ige jährliche Preiserhöhung setzt voraus, dass Bestandskunden die Erhöhung akzeptieren. Bei 15% Churn und fehlendem ROI-Nachweis besteht das Risiko, dass Preiserhöhungen die Churn-Rate weiter erhöhen. Diese Parameter sind **negativ korreliert**.

---

## 5. Zielunternehmen & Marktbezug

### 5.1 Primärmarkt: Verpackungsmaschinenbau

**PDF-Quelle:** Rene identifiziert ~2.255 Packaging/Machinery-Unternehmen als primären Markt

| Dimension | Wert | Quelle |
|-----------|------|--------|
| TAM | ~2.255 Unternehmen | PDF (Rene) |
| Segment-Priorität | Hoch (primary) | PDF |
| Sustainability-Relevanz | Hoch | PDF |
| Customer Readiness | Hoch | PDF |
| Lead Cycle | Standard | PDF |
| Margin Profile | Higher | PDF |

**Warum Packaging/Maschinenbau?**
- Hohe Sustainability-Relevanz: CO₂-Reduzierung durch Material-Substitution, Prozessoptimierung
- Standardisierter Produktionsprozess → leichtere Digitalisierung
- Keine extremen Compliance-Anforderungen wie Automotive
- Rene betont: Branchenkenntnisse vorhanden (Netzwerk für ersten Zugang)

**Konservativ-Spezifisch:** Bei 200 EUR/Seat/Monat und 5 Seats = 12.000 EUR ARR/Unternehmen/Jahr ist der Kaufwiderstand vergleichsweise gering. Aber: 2.255 × 12.000 EUR = 27 Mio. EUR theoretisches TAM (SMB-Tier). Realistisch erreichbar (5–10%): 1,35–2,7 Mio. EUR ARR.

### 5.2 Sekundärmarkt: Automotive

| Dimension | Wert | Quelle |
|-----------|------|--------|
| Lead Cycle | Länger | PDF |
| Deal Complexity | Höher | PDF |
| Market Size | TBD | PDF (offen) |
| Sustainability | Hoch | PDF |

**Konservativ-Entscheidung:** Automotive wird im Konservativ-Szenario explizit **nicht** als primärer Kanal adressiert (langer Lead-Cycle passt nicht zu Cashbedarf). Nur als Referenz- und Folgegeschäft nach M36.

### 5.3 Tertiärmarkt: Engineering Services

**PDF:** Low customer readiness, supplementary — im Konservativ-Szenario nicht aktiv adressiert.

---

## 6. Phasenentwicklung: Revenue & Kosten

### Phase 1: Pre-Seed — Aufbau & Validierung (M1–M17)
**Zeitraum:** April 2026 – August 2027
**Funding:** 1.500.000 EUR (Close ca. M6 = September 2026)

#### Ziele dieser Phase (PDF-basiert):
- 3–5 Pilot-Kunden für Produktvalidierung (kein Revenue-Fokus)
- MVP-Fertigstellung
- Team: CEO + CTO + 2 Senior Engineers
- Bereitstellung des Onboarding-Pakets (6-Stufen-Prozess)

#### Umsatzentwicklung:
| Monat | Revenue | Kumulativ |
|-------|---------|-----------|
| M1–M16 | 0 EUR | 0 EUR |
| M17 | 3.000 EUR/Monat | Erstumsatz |

**Revenue M17:** 3 Kunden × 5 Seats × 200 EUR = **3.000 EUR/Monat MRR**

#### Kostenentwicklung:
| Monat | Personalkosten | Tech/Cloud | Marketing | Office | Gesamt |
|-------|--------------|-----------|----------|--------|--------|
| M1–M4 | 14.520 EUR | 1.400 EUR | 2.083 EUR | 3.500 EUR | **21.503 EUR** |
| M5–M8 | 23.595 EUR | 1.400 EUR | 2.083 EUR | 3.500 EUR | **30.578 EUR** |
| M9–M17 | 32.670 EUR | 1.400 EUR | 2.083 EUR | 3.500 EUR | **39.653 EUR** |

**Gesamt-Cashbedarf Pre-Seed:**
- M1–M4: 4 × 21.503 = 86.012 EUR
- M5–M8: 4 × 30.578 = 122.312 EUR
- M9–M17: 9 × 39.653 = 356.877 EUR
- **Summe: ~565.000 EUR Burn** (von 1,5M Pre-Seed)
- **Residual Kapital bei Seed-Close (M17):** ~935.000 EUR

#### Kritische Annahmen in dieser Phase:
1. CTO + 2 Engineers bauen MVP in 12–14 Monaten (M5–M17)
2. 3–5 Packaging-Unternehmen akzeptieren Pilot ohne Zahlungsverpflichtung
3. Keine größeren unvorhergesehenen Tech-Kosten (Cloud-Infrastruktur, Compliance)
4. CEO findet Pre-Seed-Investoren bis M6 (September 2026)

#### Zielunternehmen in dieser Phase:
- **Anzahl:** 3–5 Pilot-Kunden
- **Profil:** Kleine bis mittlere Packaging-Unternehmen mit Nachhaltigkeitsagenda
- **Erwartung:** Kostenlos oder stark reduziert gegen Feedback und Referenz
- **Abhängigkeit:** CEO-Netzwerk in Packaging/Maschinenbau (Rene's Kontakte)

---

### Phase 2: Seed — Erste Revenue-Skalierung (M17–M29)
**Zeitraum:** September 2027 – August 2028
**Funding:** 4.000.000 EUR (Close M17)

#### Ziele dieser Phase (PDF-basiert):
- 50–60 Kunden-Target (PDF: "scale sales, go-to-customer") → Konservativ realistisch: 20–35 Kunden
- Revenue-Ramp aufbauen
- Go-to-Market-Team aufbauen: CCO, Sales, Customer Success, Marketing
- Konferenzen als primären Lead-Kanal etablieren (25.000 EUR/Jahr Events-Budget)

#### Umsatzentwicklung:
| Monat | SMB-MRR | Enterprise | Gesamt-MRR | ARR (ann.) |
|-------|---------|-----------|-----------|-----------|
| M17 | 3.000 EUR | 0 EUR | 3.000 EUR | 36.000 EUR |
| M20 | 3.280 EUR | 0 EUR | 3.280 EUR | 39.360 EUR |
| M24 | 3.680 EUR | 0 EUR | 3.680 EUR | 44.160 EUR |
| M29 | 4.280 EUR | 0 EUR | 4.280 EUR | 51.360 EUR |

> **Modell-Limit:** Die SMB-Zahlen oben zeigen nur organisches Seat-Wachstum (3%/Monat). Neukunden-Akquisition durch Sales-Team (ab M17 mit CCO + Sales Rep) ist im Modell über den Sandbox-Startmonat-Parameter und Enterprise-Deals gesteuert. Die tatsächliche Revenue-Kurve entsteht erst durch Excel-Neuberechnung.

**Realistische Schätzung mit Sales-Team-Effort (20–30 Neukunden bis M29):**
- 25 Kunden × 5 Seats × 200 EUR = 25.000 EUR/Monat MRR = **300.000 EUR ARR**

#### Kostenentwicklung Seed-Phase:
| Kategorie | M17 | M22 | M29 | Kommentar |
|-----------|-----|-----|-----|-----------|
| Personal (7 Pers.) | 38.500 EUR | 52.000 EUR | 54.000 EUR | Lohnnebenkosten inkl. |
| Cloud/AI | 1.400 EUR | 1.500 EUR | 1.600 EUR | Wächst mit Kundenzahl |
| Marketing/Events | 4.250 EUR | 4.250 EUR | 4.250 EUR | Konferenzen-Fokus |
| Office/Ops | 4.000 EUR | 4.200 EUR | 4.200 EUR | Berlin |
| **Gesamt/Monat** | **~48K EUR** | **~62K EUR** | **~64K EUR** | |

**Cash-Position Ende Seed (M29):**
- Seed-Funding: 4.000.000 EUR
- Residual aus Pre-Seed: +935.000 EUR
- Burn M17–M29 (12 Monate × ~56K avg.): −672.000 EUR
- Revenue M17–M29 (wachsend, avg. ~15K/Monat): +180.000 EUR
- **Cash M29: ~4.443.000 EUR** (auf dem Weg zur Series A, starke Position)

#### Kritische Annahmen in dieser Phase:
1. **Konferenz-First GTM:** Packaging/Machinery-Konferenzen sind der Haupt-Lead-Kanal (Rene). Mit 25.000 EUR/Jahr = 2–3 Premium-Konferenzen möglich. Konversion: 2–3 Kunden pro Konferenz = 6–9 neue Kunden/Jahr durch diesen Kanal allein.
2. **Sales-Zyklus:** Standard (Packaging/Machinery) = 3–6 Monate. Mit CCO ab M17: erste Sales-Abschlüsse in M20–M23.
3. **Onboarding-Kapazität:** Customer Success (ab M20) schafft max. 5–8 Kunden gleichzeitig onzuboarden (6-Stufen-Prozess). Engpassrisiko ab 15+ Kunden.
4. **Churn im 1. Jahr:** Piloten aus Pre-Seed müssen zu bezahlenden Kunden konvertieren — Risiko, dass 1–2 der 3 Piloten nicht verlängern.

#### Zielunternehmen in dieser Phase:
- **Anzahl:** 20–35 neue zahlende Kunden
- **Profil:** Mittelständische Packaging-/Maschinenbau-Unternehmen, 50–500 MA
- **Akquisitionskanal primär:** Industrie-Konferenzen (FachPack, Drupa, K-Messe)
- **Akquisitionskanal sekundär:** PoCs (Conversion Tool), Referenzkunden-Netzwerk
- **Deal-Größe:** 5 Seats × 200 EUR/Monat = 12.000 EUR ARR

---

### Phase 3: Series A — Professionalisierung & Enterprise (M30–M42)
**Zeitraum:** September 2028 – August 2029
**Funding:** 10.000.000 EUR (Close M36)

#### Ziele:
- Enterprise-Deals etablieren (ab M30: 1 Deal/Quartal)
- Team auf 13+ Personen skalieren
- Reporting- und Compliance-Strukturen für Series A aufbauen
- Internationalisierung vorbereiten (noch nicht aktiv)

#### Umsatzentwicklung:
| Monat | SMB-ARR | Enterprise-ARR | Gesamt-ARR | ∆ YoY |
|-------|---------|---------------|-----------|-------|
| M30 | ~55.000 EUR | 0 EUR | ~55.000 EUR | — |
| M33 | ~60.000 EUR | 25.000 EUR | ~85.000 EUR | +55% |
| M36 | ~65.000 EUR | 50.000 EUR | ~115.000 EUR | +35% |
| M39 | ~71.000 EUR | 75.000 EUR | ~146.000 EUR | +27% |
| M42 | ~78.000 EUR | 100.000 EUR | ~178.000 EUR | +22% |

> **Enterprise-Treiber dominiert:** Ab M36 macht Enterprise bereits 43% des ARR aus, obwohl erst 2 Deals abgeschlossen wurden. Jeder zusätzliche Enterprise-Deal hat einen ARR-Effekt von 25.000 EUR — deutlich mehr als SMB-Wachstum.

#### Kostenentwicklung:
| Kategorie | M30 | M36 | M42 |
|-----------|-----|-----|-----|
| Personal (10–13 Pers.) | 65.000 EUR | 80.000 EUR | 90.000 EUR |
| Tech/Cloud | 2.000 EUR | 2.500 EUR | 3.000 EUR |
| Marketing | 5.000 EUR | 5.500 EUR | 6.000 EUR |
| Office/Ops/Reise | 6.000 EUR | 7.000 EUR | 7.500 EUR |
| **Gesamt/Monat** | **~78K EUR** | **~95K EUR** | **~107K EUR** |

**EBITDA-Entwicklung (Schätzung, M42):**
- ARR: 178.000 EUR → MRR: 14.833 EUR
- Monatliche Kosten: ~107.000 EUR
- **EBITDA M42: −92.167 EUR/Monat** (noch deutlich negativ)
- Break-Even benötigt: ~105.000 EUR MRR = ~1,26 Mio. EUR ARR

#### Kritische Annahmen:
1. **Enterprise-Sales-Zyklus:** PDF nennt "longer in automotive". Im Packaging-Segment: 6–9 Monate von Lead zu Abschluss. Erster Enterprise-Deal M33 bedeutet: Lead-Qualifizierung muss in M24–M27 beginnen.
2. **Series A Investor-Erwartungen:** ARR von 115K bei M36 ist für typische Series A-Investoren (erwarten 1–3M ARR) **zu niedrig**. Konservativ-Szenario hat ein **Series A Fundraising-Risiko**.
3. **Referenz-Kunden für Enterprise:** 20–30 SMB-Kunden aus Seed als Referenz nötig. Ohne nachgewiesenen ROI kein Enterprise-Deal.

---

### Phase 4: Scale — Richtung Series B (M43–M52)
**Zeitraum:** September 2029 – Juni 2030
**Funding:** Series B (Timing unsicher, per PDF: "durch Aug 2029" — bereits in Sichtweite)

#### Umsatzentwicklung:
| Monat | SMB-ARR | Enterprise-ARR | Gesamt-ARR |
|-------|---------|---------------|-----------|
| M43 | ~85.000 EUR | 125.000 EUR | ~210.000 EUR |
| M47 | ~95.000 EUR | 150.000 EUR | ~245.000 EUR |
| M52 | ~105.000 EUR | 175.000 EUR | ~280.000 EUR |

**ARR Jahresziel M52: 280.000–350.000 EUR** (inkl. Preiserhöhungseffekte)

#### Kosten M52 (~20 Personen):
- Personal: ~105.000 EUR/Monat
- Tech/Infra: ~4.000 EUR/Monat
- Marketing: ~7.000 EUR/Monat
- Office/Ops: ~9.000 EUR/Monat
- **Gesamt: ~125.000 EUR/Monat**

**EBITDA M52:**
MRR: ~29.000 EUR | Kosten: ~125.000 EUR | **Monatlicher Verlust: −96.000 EUR**

> **Kritische Beobachtung:** Das Konservativ-Szenario erreicht auch nach 52 Monaten keinen Break-Even. Die Lücke zwischen ARR (~350K EUR/Jahr) und Jahreskosten (~1,5M EUR) ist groß. **Series B wäre zwingend notwendig**, muss aber mit ~350K ARR bei Series A-Investoren gerechtfertigt werden — schwierige Position.

---

## 7. Risikobewertung

### 7.1 Kritische Risiken (Wahrscheinlichkeit × Impact)

| Risiko | W'keit | Impact | Beschreibung |
|--------|-------|--------|-------------|
| First Revenue > M17 | Hoch | Sehr hoch | Jeder Monat Verzögerung kostet ~40K EUR zusätzlich |
| Churn > 15% | Mittel | Hoch | Schlechtes Onboarding oder fehlender ROI-Nachweis |
| Series A Funding-Lücke | Hoch | Sehr hoch | 115K ARR zu wenig für Series A Investoren |
| Customer Success Bottleneck | Hoch | Mittel | Skalierungsproblem bereits ab 15+ Kunden |
| Preiserosion | Mittel | Mittel | 200 EUR/Seat wird von Wettbewerbern unterboten |
| Enterprise-Deals verzögern | Mittel | Hoch | Lead-Zyklen länger als M30 erwartet |

### 7.2 Offene Punkte aus PDF (direkt relevant)

- **G1.1 (HIGH):** Timeline-Konsolidierung (Sept 2026 vs. April-Start)
- **G1.3 (HIGH):** First-Paying-Customer-Timing — die kritischste Einzelvariable
- **G1.4 (MED):** Deployment-Standard (Cloud vs. On-Prem) → On-Prem erhöht Onboarding-Aufwand massiv
- **G2.3 (HIGH):** Customer-Success-Skalierungsmodell — identifizierter Engpass
- **G3.1 (HIGH):** CO₂-Kausalitätskette — ohne diese kein Premium-Pricing möglich

---

## 8. Zusammenfassung Konservativ

Das Konservativ-Szenario ist **finanziell überlebensfähig**, aber **strategisch suboptimal**:

✅ **Stärken:**
- Realistisches Risikoprofil
- Ausreichend Funding für jede Phase
- Fokus auf Packaging-Nische gut begründet

❌ **Schwächen:**
- 280–350K ARR nach 52 Monaten ist zu wenig für die geplante Funding-Story
- 15% Churn gefährdet die ARR-Basis
- 200 EUR/Seat lässt 150 EUR/Seat Marge (zum PDF-Referenzpreis) ungenutzt
- Kein KI-Hebel = hohe und früh skalierenде Personalkosten
- Series A Fundraising bei 115K ARR strukturell schwierig

**Empfohlene Maßnahmen:**
1. Preis auf mindestens 300–350 EUR/Seat anheben (wenn ROI-Nachweis gelingt)
2. First-Revenue-Datum auf M14 vorziehen (PoC-Konversion früher)
3. Customer Success ab M17 (nicht M20) einstellen
4. KI-Hebel für Support- und Admin-Rollen einführen (mindestens D10=24)
5. Enterprise-Deals ab M24 (nicht M30) anstreben

---
*Erstellt auf Basis des LFL Financial Planning & Carbon Case Dossiers vom 04.03.2026*
*Alle Zahlen sind Schätzwerte. Genaue Werte entstehen durch Excel-Neuberechnung (Formeln aller nachgelagerten Sheets).*
