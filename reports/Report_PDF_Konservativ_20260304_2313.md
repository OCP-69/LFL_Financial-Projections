# LFL Financial Projections — Szenario KONSERVATIV: PDF_Konservativ
**Basis:** 260304_LFL-Financial-Planning-and-Carbon-Case.pdf
**Erstellt:** 04.03.2026 23:13

---

## Hintergrund

Dieses Szenario basiert auf dem LFL Financial Planning & Carbon Case Dossier (260304_LFL-Financial-Planning-and-Carbon-Case.pdf, Stand 4. März 2026). Alle Parameter wurden aus den strategischen Meeting-Transkripten von Rene und dem Finanzmodellierer abgeleitet. Externe Annahmen wurden nicht hinzugefügt. Offene Punkte (TBD/OPEN) aus dem Dossier sind im Report-Sheet markiert.

---

## 1. PREISGESTALTUNG — Konservative Positionierung

| Parameter | Bezeichnung | Baseline | NEU | Begründung |
|-----------|------------|---------|-----|------------|
| `Sandbox D6` | Preis/Seat/Monat | 800 EUR (gering-Standard) | **200 EUR** | Das PDF nennt 350 EUR/Monat als Referenzpreis (→ 4.200 EUR/Seat/Jahr) und 200 EUR als benchmark_low. Im konservativen Szenario wählen wir den Benchmark-Tiefpunkt, um das Preisrisiko in der Marktvalidi |
| `Inputs B21` | Preiserhöhung/Jahr | 8% | **5%** | Geringere jährliche Preissteigerung reflektiert schwächere Marktposition und Kundenwiderstand in der Frühphase. Packaging-Kunden (primärer Markt: 2.255 Unternehmen laut Rene) sind preissensitiv. |

**→ Hauptauswirkung:** ARR pro Standardkunde (5 Seats): 200 EUR × 5 × 12 = 12.000 EUR/Jahr (vs. 21.000 EUR im PDF-Referenzpreis). Runway und Break-Even verschlechtern sich deutlich. Höheres Funding-Risiko.

## 2. KUNDEN-TIMING & WACHSTUM — Langsame Marktdurchdringung

| Parameter | Bezeichnung | Baseline | NEU | Begründung |
|-----------|------------|---------|-----|------------|
| `Sandbox D5` | Startmonat erste Kunden | 14 | **17** | Das PDF nennt explizit '~17 months to first revenue'. Im konservativen Szenario halten wir an diesem Planwert fest. Die Pre-Seed-Phase (3-5 Kunden, Sept 2026-Sept 2027) dient reiner Validierung ohne R |
| `Inputs B23` | Initiale Seats | 5 | **3** | PDF nennt 3-5 als Pre-Seed-Kundenzahl. Konservativ: 3 Piloten. Jeder Pilot mit 5 Seats = 15 Seats total. |
| `Inputs B24` | Monatl. Seat-Wachstum | 5% | **3%** | Verlangsamtes Wachstum durch längere Sales-Zyklen (packaging/automotive lt. PDF: 'longer lead cycle in automotive'). 3% entspricht ca. 43% Jahreswachstum – konservativ aber realistisch für B2B-Manufac |
| `Inputs B25` | Enterprise-Deals ab Monat | 24 | **30** | Enterprise-Deals starten erst nach Monat 30 – nach erstem erfolgreichen Track Record. PDF: deal_complexity höher als Standard, lead_cycle länger. |

**→ Hauptauswirkung:** Seed-Ziel: 50-60 Kunden (PDF) → im Konservativ-Szenario realistisch 25-35 Kunden bis Monat 17-24. Gesamt-ARR M24: deutlich unter 500K EUR. Enterprise-Revenue startet erst ab M30.

## 3. CHURN & KUNDENBINDUNG — Höheres Abwanderungsrisiko

| Parameter | Bezeichnung | Baseline | NEU | Begründung |
|-----------|------------|---------|-----|------------|
| `Inputs B28` | Jährliche Churn Rate | 8% | **15%** | PDF: adoption_probability_low = 10%. Geringe Adoption bedeutet hohes Abwanderungsrisiko sobald Verträge auslaufen. 15% Churn entspricht einem 'gefährdeten' SaaS-Modell in der Early-Stage, typisch für  |
| `Inputs B29` | Net Revenue Retention | 118% | **105%** | Ohne starke Expansion durch Upsell oder neue Seats bleibt NRR knapp über 100%. PDF: Keine quantifizierten Upsell-Mechanismen – konservativ daher 105%. |

**→ Hauptauswirkung:** ARR-Erosion durch Churn: Bei 100K ARR und 15% Churn verliert das Unternehmen 15K/Jahr. Ohne starkes Neuwachstum schrumpft der Revenue-Pool. NRR von 105% rettet knapp die Nettowachstum-Rate.

## 4. FUNDING — Konservative Kapitalstrategie

| Parameter | Bezeichnung | Baseline | NEU | Begründung |
|-----------|------------|---------|-----|------------|
| `Inputs B12` | Pre-Seed Betrag | 1.500.000 EUR | **1.500.000 EUR** | PDF Lower End: 1,5M EUR. Im Konservativ-Szenario bleibt Pre-Seed beim Minimum. Runway: 12 Monate (Sept 2026 – Sept 2027). Reicht für MVP + 3-5 Pilotkunden. |
| `Inputs B14` | Seed Betrag | 6.000.000 EUR | **4.000.000 EUR** | PDF Seed Lower End: 4M EUR. Konservativ, da langsameres Wachstum weniger Kapital verbraucht. Seed-Runway: 12 Monate (Sept 2027 – Sept 2028), für 50-60 Kunden. |
| `Inputs B16` | Series A Betrag | 15.000.000 EUR | **10.000.000 EUR** | Reduzierte Series A durch geringeres Wachstumstempo. Schont Verwässerung. |
| `Inputs B17` | Series A Monat | 35 | **36** | Leicht später, da Meilensteine langsamer erreicht werden. |

**→ Hauptauswirkung:** Gesamtfinanzierung: 15,5M EUR (vs. 23,5M im Aggressiv). Niedrigerer Cash-Verbrauch durch geringere Headcount- und Marketing-Ausgaben.

## 5. HEADCOUNT & KI-HEBEL — Vollständige Personalplanung

| Parameter | Bezeichnung | Baseline | NEU | Begründung |
|-----------|------------|---------|-----|------------|
| `Sandbox D10` | AI-Personal-Hebel | 0 Monate (Gering-Standard) | **0 Monate** | Konservativ: Kein KI-Delay beim Hiring. Alle Rollen werden zum ursprünglich geplanten Monat eingestellt. Begründung: KI-Werkzeuge sind in der Praxis in Monat 1-24 noch nicht produktionsreif genug, um  |
| `Inputs B44` | Jährl. Gehaltserhöhung | 5% | **3%** | Geringere Lohnsteigerung reflektiert konservatives Budget und weniger Wettbewerb um Talent im konservativen Szenario (kleineres Unternehmen, Berlin-Markt). |

**→ Hauptauswirkung:** Konservativ hat MEHR Mitarbeiter als Aggressiv (kein KI-Delay). Personalkosten sind der größte Kostenblock. Typische Entwicklung:
  M1-5: CEO + CTO (Founder, Eigenfinanzierung)
  M5-17: +2 Senior Engineers nach Pre-Seed
  M17-24: +CCO, +1 Sales, +1 Customer Success nach Seed
  M24-36: +Marketing, +Finance, +2 weitere Engineers
  M36-52: +2-3 Sales Reps, +2 CS, +1 weitere
  GESAMT M52: ~18-22 Mitarbeiter

## 6. MARKETING & SALES — Budgetdisziplin

| Parameter | Bezeichnung | Baseline | NEU | Begründung |
|-----------|------------|---------|-----|------------|
| `Inputs B125` | Events & Messen/Jahr | 25.000 EUR | **25.000 EUR** | PDF-Placeholder-Wert beibehalten. Rene betont: Industrie-Konferenzen > Software-Konferenzen. 25K EUR als Baseline für 2-3 Packaging/Machinery-Konferenzen. Im Konservativ-Szenario kein Aufschlag. |
| `Inputs B122` | Paid Ads Initial | 500 EUR/Monat | **300 EUR/Monat** | Geringeres Ads-Budget in der Validierungsphase. B2B-Manufacturing kauft nicht über Ads. |
| `Inputs B127` | Sales-Provision | 10% | **8%** | Etwas geringere Provision, da weniger aggressives Sales-Ziel. |

**→ Hauptauswirkung:** Total Marketing-Spend M1-M52: ~30-40% geringer als Aggressiv. GTM über Konferenzen (primary), Content (sekundär), POCs (Conversion).

---

## Gesamtvergleich

```

╔══════════════════════════════════════════════════════════════════════╗
║         HAUPTUNTERSCHIEDE: KONSERVATIV vs. AGGRESSIV                ║
╠══════════════════════════════════════════════════════════════════════╣
║  Parameter          │ Konservativ  │ Aggressiv    │ Quelle          ║
╠══════════════════════════════════════════════════════════════════════╣
║  Preis/Seat/Monat   │ 200 EUR      │ 500 EUR      │ PDF benchmark   ║
║  Preis/Seat/Jahr    │ 2.400 EUR    │ 6.000 EUR    │ ×12             ║
║  First Revenue Mon. │ 17           │ 9            │ PDF + Anpassung ║
║  Initiale Seats     │ 3            │ 5            │ PDF avg         ║
║  Seat-Wachstum/Mo.  │ 3%           │ 8%           │ Scenario        ║
║  Enterprise ab Mon. │ 30           │ 18           │ Scenario        ║
║  Enterprise ARR     │ 25.000 EUR   │ 50.000 EUR   │ Scenario        ║
║  Deals/Quartal      │ 1            │ 2            │ Scenario        ║
║  Churn Rate/Jahr    │ 15%          │ 6%           │ PDF adoption    ║
║  NRR                │ 105%         │ 125%         │ Scenario        ║
║  Pre-Seed           │ 1,5M EUR     │ 2,0M EUR     │ PDF range       ║
║  Seed               │ 4,0M EUR     │ 6,0M EUR     │ PDF range       ║
║  Series A           │ 10M EUR      │ 20M EUR      │ Scenario        ║
║  Headcount M52      │ ~18-22       │ ~10-14       │ KI-Hebel-Effekt ║
║  Events/Jahr        │ 25.000 EUR   │ 50.000 EUR   │ PDF + ×2        ║
║  KI-Personal-Hebel  │ 0 Monate     │ 99 Monate    │ Maximaler Effekt║
║  ARR-Schätzung M24  │ 200-400K     │ 600K-1,2M    │ Modellschätzung ║
║  ARR-Schätzung M52  │ 500K-1M      │ 5-10M        │ Modellschätzung ║
╚══════════════════════════════════════════════════════════════════════╝

```

---

## Offene Punkte aus dem PDF


OFFENE PUNKTE AUS DEM PDF (Decision Register Part G):
─────────────────────────────────────────────────────
G1.1 HIGH: Timeline-Varianten konsolidieren (Sept 2026 vs. April-Start)
G1.3 HIGH: First-Paying-Customer-Timing final festlegen (kritischste Variable)
G1.4 MED:  Deployment-Standard (Cloud/On-Prem/Hybrid) → beeinflusst Kostenstruktur
G1.7 MED:  KI-Rollensubstitution quantifizieren (direkt auf Personalkosten)
G2.2 HIGH: Conversion-Rate-Annahmen per Kanal definieren
G2.3 HIGH: Customer-Success-Skalierungsmodell planen (identifizierter Engpass)
G3.1 HIGH: Kausale CO₂-Wirkungskette präzise definieren
G3.3 HIGH: Baseline für CO₂-Berechnungen festlegen


---

## Angewendete Änderungen (technisch)

| Sheet | Zelle | Bezeichnung | Alt → Neu |
|-------|-------|-------------|-----------|
| 00_Input_Sandbox | `B1` | Aktives Szenario | gering → **gering** |
| Inputs | `B12` | Pre-Seed Betrag | 1500000 → **1500000** |
| Inputs | `B14` | Seed Betrag | 6000000 → **4000000** |
| Inputs | `B15` | Seed Monat | 17 → **17** |
| Inputs | `B16` | Series A Betrag | 15000000 → **10000000** |
| Inputs | `B17` | Series A Monat | 35 → **36** |
| Inputs | `B21` | Preiserhöhung/Jahr | 0.08 → **0.05** |
| Inputs | `B23` | Initiale Seats | 5 → **3** |
| Inputs | `B24` | Monatliche Seat-Wachstumsrate | 0.05 → **0.03** |
| Inputs | `B25` | Enterprise-Deals ab Monat | 24 → **30** |
| Inputs | `B26` | Durchschnitt Enterprise ARR | 150000 → **25000** |
| Inputs | `B27` | Enterprise Deals pro Quartal | 1 → **1** |
| Inputs | `B28` | Jährliche Churn Rate | 0.08 → **0.15** |
| Inputs | `B29` | Net Revenue Retention | 1.18 → **1.05** |
| Inputs | `B44` | Jährliche Gehaltserhöhung | 0.05 → **0.03** |
| Inputs | `B82` | Cloud/Hosting Basis | 1200 → **800** |
| Inputs | `B84` | AI/ML APIs Basis | 1000 → **600** |
| Inputs | `B85` | AI Kosten Wachstum/Monat | 0.05 → **0.03** |
| Inputs | `B122` | Paid Ads Budget Initial | 500 → **300** |
| Inputs | `B123` | Ads Budget Wachstum/Monat | 0.05 → **0.03** |
| Inputs | `B124` | Content & SEO | 1500 → **1000** |
| Inputs | `B125` | Events & Messen/Jahr | 25000 → **25000** |
| Inputs | `B127` | Sales Provision | 0.1 → **0.08** |
| Inputs | `B128` | Reisekosten Sales/MA/Monat | 500 → **300** |
| 00_Input_Sandbox | `D5` | Startmonat Kunden (Gering) | 14 → **17** |
| 00_Input_Sandbox | `D6` | Startpreis Seat/Jahr (Gering) | 800 → **200** |
| 00_Input_Sandbox | `D7` | Consulting-Tagessatz (Gering) | 1200 → **1200** |
| 00_Input_Sandbox | `D9` | Onboarding-Aufwand (Gering) | 120 → **120** |
| 00_Input_Sandbox | `D10` | AI-Personal-Hebel (Gering) | 0 → **0** |
