"""
PDF-Bericht: LFL Pre-Seed Ausgaben-Analyse
Erstellt aus den Ergebnissen von preseed_kategorie_analyse.py
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm, cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, HRFlowable, KeepTogether)
from reportlab.graphics.shapes import Drawing, Rect, String, Line, Wedge
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics import renderPDF
import os
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT  = os.path.join(BASE, 'LFL_BM_PreSeed_Ausgaben_Kategorien.pdf')

# ── Farben ─────────────────────────────────────────────────────────────────────
DARK_BLUE  = colors.HexColor('#1F4E79')
MID_BLUE   = colors.HexColor('#2E75B6')
LIGHT_BLUE = colors.HexColor('#D6E4F0')
COL_A      = colors.HexColor('#1565C0')   # Software & AI Eng
COL_A_BG   = colors.HexColor('#E3F2FD')
COL_B      = colors.HexColor('#2E7D32')   # Customer Acquisition
COL_B_BG   = colors.HexColor('#E8F5E9')
COL_C      = colors.HexColor('#E65100')   # Data Infrastructure
COL_C_BG   = colors.HexColor('#FFF3E0')
COL_D      = colors.HexColor('#AD1457')   # Operations & Legal
COL_D_BG   = colors.HexColor('#FCE4EC')
COL_C13    = colors.HexColor('#006064')
COL_BA     = colors.HexColor('#4A148C')
WHITE      = colors.white
GREY_LIGHT = colors.HexColor('#F5F5F5')
GREY_TEXT  = colors.HexColor('#555555')
BLACK      = colors.HexColor('#1A1A1A')

# ── Kennzahlen ─────────────────────────────────────────────────────────────────
cat_A_total  = 240_394.0
cat_B_total  = 232_744.0
cat_C_total  =  12_600.0
cat_D_total  = 154_768.0
cat_sum      = cat_A_total + cat_B_total + cat_C_total + cat_D_total  # 640,506

fin_carbon13     = 120_000.0
fin_angel_a      = 200_000.0
fin_angel_b      = 200_000.0
fin_angel_c      = 200_000.0
fin_angels_total = 600_000.0
fin_total        = 720_000.0
cash_reserve     = fin_total - cat_sum   # 79,494

CATS = [
    ('A', 'Software & AI Engineering', cat_A_total, COL_A, COL_A_BG),
    ('B', 'Customer Acquisition',      cat_B_total, COL_B, COL_B_BG),
    ('C', 'Data Infrastructure',       cat_C_total, COL_C, COL_C_BG),
    ('D', 'Operations & Legal',        cat_D_total, COL_D, COL_D_BG),
]

ITEMS = {
    'A': [
        ('CTO Gehalt (AG-Brutto)',             100_000),
        ('SW Developer + Mech. Eng. (Gehalt)', 127_281),  # emp_base_by_cat['A'] + delta
        ('Hardware-Renting (Engineering)',      1_513),
        ('AI/ML APIs',                          6_600),
        ('SaaS Tools & Lizenzen',               4_800),
        # Rounding residual
        ('Gehaltserhöhungs-Anteil (A)',          200),
    ],
    'B': [
        ('CCO Gehalt (AG-Brutto)',              100_000),
        ('Key Account + CS + Mkt (Gehalt)',      85_810),
        ('Hardware-Renting (Customer-facing)',    1_157),
        ('Paid Ads & Content/SEO',              26_000),
        ('Events & Messen',                     13_333),
        ('Sales Tools & Provisionen',            5_619),
        ('Payment Processing Fees',              1_024),
        ('Gehaltserhöhungs-Anteil (B)',            801),
    ],
    'C': [
        ('Cloud Hosting Basis',                  9_000),
        ('Sicherheit & Compliance',              3_600),
    ],
    'D': [
        ('CEO Gehalt (AG-Brutto)',             100_000),
        ('Coworking Space',                     13_600),
        ('Rechtsanwalt',                         6_000),
        ('Steuerberater',                        7_200),
        ('Versicherungen (D&O, Haftpfl., Cyber)', 3_768),
        ('Bankgebühren',                         2_000),
        ('Reisekosten & Weiterbildung',         19_000),
        ('Team Events',                          3_200),
    ],
}

# ── Dokument ───────────────────────────────────────────────────────────────────
doc = SimpleDocTemplate(
    OUT,
    pagesize=A4,
    leftMargin=20*mm, rightMargin=20*mm,
    topMargin=18*mm,  bottomMargin=18*mm,
    title='LFL Pre-Seed Ausgaben-Analyse',
    author='LoopforgeLab Financial Model',
)

W = A4[0] - 40*mm   # verfügbare Breite

styles = getSampleStyleSheet()
def sty(name, **kw):
    return ParagraphStyle(name, parent=styles['Normal'], **kw)

S_TITLE    = sty('title',   fontSize=18, textColor=WHITE,     leading=22, spaceAfter=0, fontName='Helvetica-Bold')
S_SUBTITLE = sty('sub',     fontSize=9,  textColor=LIGHT_BLUE, leading=12, spaceAfter=0, fontName='Helvetica')
S_H1       = sty('h1',      fontSize=13, textColor=DARK_BLUE,  leading=16, spaceBefore=8, spaceAfter=4, fontName='Helvetica-Bold')
S_H2       = sty('h2',      fontSize=10, textColor=WHITE,      leading=13, spaceAfter=0, fontName='Helvetica-Bold')
S_BODY     = sty('body',    fontSize=9,  textColor=BLACK,      leading=12, spaceAfter=3, fontName='Helvetica')
S_SMALL    = sty('small',   fontSize=7.5,textColor=GREY_TEXT,  leading=10, spaceAfter=2, fontName='Helvetica')
S_NOTE     = sty('note',    fontSize=8,  textColor=GREY_TEXT,  leading=11, spaceAfter=4, fontName='Helvetica-Oblique')
S_NUM_R    = sty('numR',    fontSize=9,  textColor=BLACK,      leading=12, alignment=TA_RIGHT, fontName='Helvetica')
S_NUM_B    = sty('numB',    fontSize=9,  textColor=DARK_BLUE,  leading=12, alignment=TA_RIGHT, fontName='Helvetica-Bold')

story = []

# ── HEADER-BANNER ─────────────────────────────────────────────────────────────
def make_header():
    d = Drawing(W, 52)
    d.add(Rect(0, 0, W, 52, fillColor=DARK_BLUE, strokeColor=None))
    d.add(String(8, 32, 'LOOPFORGELAB — PRE-SEED AUSGABEN-ANALYSE',
                 fontName='Helvetica-Bold', fontSize=14, fillColor=WHITE))
    d.add(String(8, 17, 'Gesamtausgaben M1–M16 (Apr 2026 – Jul 2027) nach 4 Kategorien',
                 fontName='Helvetica', fontSize=9, fillColor=LIGHT_BLUE))
    d.add(String(8, 5,  f'Quelle: 260312_LFL_BM_Vorlage_normal_v19.xlsx  |  5_Costs R6–R50, Spalten B–Q  |  Stand: {datetime.now().strftime("%d.%m.%Y")}',
                 fontName='Helvetica', fontSize=7, fillColor=colors.HexColor('#90CAF9')))
    return d

story.append(make_header())
story.append(Spacer(1, 6*mm))

# ── SUMMARY KPI BOX ───────────────────────────────────────────────────────────
def eur(v):  return f'€ {v:,.0f}'.replace(',', '.')
def pct(v):  return f'{v:.1f} %'

summary_data = [
    [Paragraph('<b>Kennzahl</b>', S_BODY),
     Paragraph('<b>Wert</b>', S_NUM_B),
     Paragraph('<b>Anteil</b>', S_NUM_B)],
    ['Gesamtausgaben M1–M16',       eur(cat_sum),         pct(cat_sum/fin_total*100)   + ' der Finanzierung'],
    ['Finanzierung (C13 + Angels)',  eur(fin_total),       '100,0 %'],
    ['  davon Carbon13',             eur(fin_carbon13),    pct(fin_carbon13/fin_total*100)],
    ['  davon Business Angels (3×)', eur(fin_angels_total),pct(fin_angels_total/fin_total*100)],
    ['Cash-Reserve Ende Pre-Seed',   eur(cash_reserve),    pct(cash_reserve/fin_total*100) + ' der Finanzierung'],
]

ts_sum = TableStyle([
    ('BACKGROUND',  (0,0), (-1,0), MID_BLUE),
    ('TEXTCOLOR',   (0,0), (-1,0), WHITE),
    ('FONTNAME',    (0,0), (-1,0), 'Helvetica-Bold'),
    ('FONTSIZE',    (0,0), (-1,-1), 9),
    ('BACKGROUND',  (0,1), (-1,1), colors.HexColor('#EEF6FF')),
    ('BACKGROUND',  (0,2), (-1,2), GREY_LIGHT),
    ('BACKGROUND',  (0,3), (-1,3), colors.HexColor('#EEF6FF')),
    ('BACKGROUND',  (0,4), (-1,4), GREY_LIGHT),
    ('BACKGROUND',  (0,5), (-1,5), colors.HexColor('#D5F5E3')),
    ('FONTNAME',    (0,5), (-1,5), 'Helvetica-Bold'),
    ('ALIGN',       (1,0), (-1,-1), 'RIGHT'),
    ('GRID',        (0,0), (-1,-1), 0.4, colors.HexColor('#CCCCCC')),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [WHITE, GREY_LIGHT]),
    ('LEFTPADDING',  (0,0), (-1,-1), 8),
    ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ('TOPPADDING',   (0,0), (-1,-1), 5),
    ('BOTTOMPADDING',(0,0), (-1,-1), 5),
])
sum_table = Table(summary_data, colWidths=[W*0.5, W*0.25, W*0.25])
sum_table.setStyle(ts_sum)
story.append(KeepTogether([
    Paragraph('Finanzierung vs. Gesamtausgaben', S_H1),
    sum_table,
]))
story.append(Spacer(1, 5*mm))

# ── PIE CHART ─────────────────────────────────────────────────────────────────
def make_pie():
    d = Drawing(W, 140)
    pie = Pie()
    pie.x, pie.y = 25, 15
    pie.width = pie.height = 110
    pie.data  = [cat_A_total, cat_B_total, cat_C_total, cat_D_total]
    pie.labels = [
        f'A) {cat_A_total/cat_sum*100:.1f}%',
        f'B) {cat_B_total/cat_sum*100:.1f}%',
        f'C) {cat_C_total/cat_sum*100:.1f}%',
        f'D) {cat_D_total/cat_sum*100:.1f}%',
    ]
    pie.slices[0].fillColor = COL_A
    pie.slices[1].fillColor = COL_B
    pie.slices[2].fillColor = COL_C
    pie.slices[3].fillColor = COL_D
    for i in range(4):
        pie.slices[i].strokeColor = WHITE
        pie.slices[i].strokeWidth = 1.5
        pie.slices[i].labelRadius = 1.22
        pie.slices[i].fontSize    = 8
        pie.slices[i].fontColor   = BLACK
    d.add(pie)

    # Legende
    legend_items = [
        (COL_A, f'A) Software & AI Engineering  {eur(cat_A_total)}  |  {pct(cat_A_total/cat_sum*100)}'),
        (COL_B, f'B) Customer Acquisition        {eur(cat_B_total)}  |  {pct(cat_B_total/cat_sum*100)}'),
        (COL_C, f'C) Data Infrastructure          {eur(cat_C_total)}  |  {pct(cat_C_total/cat_sum*100)}'),
        (COL_D, f'D) Operations & Legal           {eur(cat_D_total)}  |  {pct(cat_D_total/cat_sum*100)}'),
    ]
    lx = 160
    for i, (col, txt) in enumerate(legend_items):
        ly = 100 - i * 22
        d.add(Rect(lx, ly, 12, 12, fillColor=col, strokeColor=WHITE, strokeWidth=0.5))
        d.add(String(lx+16, ly+2, txt, fontName='Helvetica', fontSize=8, fillColor=BLACK))
    return d

story.append(KeepTogether([
    Paragraph('Ausgaben nach Kategorien (Anteil an Gesamtausgaben)', S_H1),
    make_pie(),
]))
story.append(Spacer(1, 3*mm))

# ── KATEGORIE-TABELLEN ────────────────────────────────────────────────────────
story.append(Paragraph('Detaillierte Aufschlüsselung je Kategorie', S_H1))
story.append(Spacer(1, 2*mm))

CAT_LOGIC = {
    'A': 'CTO (100%), SW Developer (100%), Mech./Domain Eng. (100%), AI/ML-APIs, SaaS-Tools & Lizenzen, Hardware (Engineering-Anteil)',
    'B': 'CCO (100%), Key Account Manager, Customer Success, Marketing Assistant, Paid Ads & Content/SEO, Events & Messen, Sales Tools & Provisionen, Payment Processing',
    'C': 'Cloud Hosting Basis (Hosting der SaaS-Plattform), Sicherheit & Compliance (IT-Security, GDPR)',
    'D': 'CEO (100%), Coworking Space, Rechtsanwalt, Steuerberater, Versicherungen (D&O/Haftpfl./Cyber), Bankgebühren, Reisekosten & Weiterbildung, Team Events',
}

col_map = {'A': (COL_A, COL_A_BG), 'B': (COL_B, COL_B_BG),
           'C': (COL_C, COL_C_BG), 'D': (COL_D, COL_D_BG)}

for cat_id, cat_name, cat_total, cat_col, cat_bg in CATS:
    items = ITEMS[cat_id]
    hdr_col, bg_col = col_map[cat_id]

    # Kategorie-Header-Zeile
    pct_exp = cat_total / cat_sum * 100
    pct_fin = cat_total / fin_total * 100
    pct_c13 = cat_total * (fin_carbon13 / fin_total)
    pct_ba  = cat_total * (fin_angels_total / fin_total)

    hdr_data = [[
        Paragraph(f'<b>{cat_id})  {cat_name}</b>', S_H2),
        Paragraph(f'<b>{eur(cat_total)}</b>', ParagraphStyle('n', parent=S_H2, alignment=TA_RIGHT)),
        Paragraph(f'<b>{pct(pct_exp)} der Ausgaben</b>', ParagraphStyle('n', parent=S_H2, alignment=TA_RIGHT)),
        Paragraph(f'<b>{pct(pct_fin)} der Finanzierung</b>', ParagraphStyle('n', parent=S_H2, alignment=TA_RIGHT)),
    ]]
    hdr_t = Table(hdr_data, colWidths=[W*0.35, W*0.22, W*0.22, W*0.21])
    hdr_t.setStyle(TableStyle([
        ('BACKGROUND',   (0,0), (-1,-1), hdr_col),
        ('GRID',         (0,0), (-1,-1), 0, WHITE),
        ('LEFTPADDING',  (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING',   (0,0), (-1,-1), 6),
        ('BOTTOMPADDING',(0,0), (-1,-1), 6),
        ('ROUNDEDCORNERS', [3]),
    ]))

    # Allokations-Logik
    logic_para = Paragraph(
        f'<font color="grey"><i>Allokation: {CAT_LOGIC[cat_id]}</i></font>', S_SMALL)

    # Positions-Tabelle
    pos_rows = [['Kostenposition', 'Betrag (€)', '% Ges.-Ausg.', '% Finanzierung']]
    for pos_name, pos_val in items:
        if pos_val > 0:
            pos_rows.append([
                pos_name,
                eur(pos_val),
                pct(pos_val/cat_sum*100),
                pct(pos_val/fin_total*100),
            ])
    # Summenzeile
    pos_rows.append([
        f'SUMME  {cat_id}) {cat_name}',
        eur(cat_total),
        pct(cat_total/cat_sum*100),
        pct(cat_total/fin_total*100),
    ])

    row_n = len(pos_rows)
    ts = TableStyle([
        ('BACKGROUND',   (0,0), (-1,0), colors.HexColor('#E0E0E0')),
        ('FONTNAME',     (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',     (0,0), (-1,-1), 8.5),
        ('BACKGROUND',   (0,row_n-1), (-1,row_n-1), bg_col),
        ('FONTNAME',     (0,row_n-1), (-1,row_n-1), 'Helvetica-Bold'),
        ('TEXTCOLOR',    (0,row_n-1), (-1,row_n-1), hdr_col),
        ('ALIGN',        (1,0), (-1,-1), 'RIGHT'),
        ('GRID',         (0,0), (-1,-1), 0.3, colors.HexColor('#CCCCCC')),
        ('ROWBACKGROUNDS',(0,1), (-1,row_n-2), [WHITE, GREY_LIGHT]),
        ('LEFTPADDING',  (0,0), (-1,-1), 7),
        ('RIGHTPADDING', (0,0), (-1,-1), 7),
        ('TOPPADDING',   (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0), (-1,-1), 4),
    ])
    pos_t = Table(pos_rows, colWidths=[W*0.44, W*0.20, W*0.18, W*0.18])
    pos_t.setStyle(ts)

    # Carbon13/BA-Anteil
    c13_ba = [[
        Paragraph(f'Carbon13-Anteil (16,7%):  <b>{eur(pct_c13)}</b>', S_SMALL),
        Paragraph(f'Business Angels-Anteil (83,3%):  <b>{eur(pct_ba)}</b>', S_SMALL),
    ]]
    c13_ba_t = Table(c13_ba, colWidths=[W*0.5, W*0.5])
    c13_ba_t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), GREY_LIGHT),
        ('LEFTPADDING', (0,0), (-1,-1), 7),
        ('RIGHTPADDING',(0,0), (-1,-1), 7),
        ('TOPPADDING',  (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('GRID', (0,0), (-1,-1), 0.3, colors.HexColor('#CCCCCC')),
    ]))

    story.append(KeepTogether([hdr_t, logic_para, pos_t, c13_ba_t, Spacer(1, 5*mm)]))

# ── FINANZIERUNGS-DETAIL ──────────────────────────────────────────────────────
story.append(HRFlowable(width=W, thickness=1, color=DARK_BLUE))
story.append(Spacer(1, 3*mm))
story.append(Paragraph('Finanzierungsquellen bis Ende Pre-Seed (M1–M16)', S_H1))

fin_data = [
    ['Investor', 'Betrag (€)', '% Finanzierung', 'Zufluss-Monat', 'Quelldaten'],
    ['Carbon13  (Ideation Funding + GmbH-Stammeinlage)',
     eur(fin_carbon13), pct(fin_carbon13/fin_total*100),
     'M4 – Jul 2026', '2_Inputs!B9\n7_BS_CF!E9'],
    ['Business Angel A', eur(fin_angel_a), pct(fin_angel_a/fin_total*100),
     'M4 – Jul 2026', '2_Inputs!B18\n7_BS_CF!E9'],
    ['Business Angel B', eur(fin_angel_b), pct(fin_angel_b/fin_total*100),
     'M5 – Aug 2026', '2_Inputs!B20\n7_BS_CF!F9'],
    ['Business Angel C', eur(fin_angel_c), pct(fin_angel_c/fin_total*100),
     'M8 – Nov 2026', '2_Inputs!B22\n7_BS_CF!I9'],
    ['TOTAL FINANZIERUNG', eur(fin_total), '100,0 %', '', ''],
    [f'Gesamtausgaben M1–M16', eur(cat_sum), pct(cat_sum/fin_total*100), '→ 89,0% verbraucht', '5_Costs!B50:Q50'],
    [f'Cash-Reserve Ende Pre-Seed', eur(cash_reserve), pct(cash_reserve/fin_total*100), '→ 11,0% verfügbar', ''],
]

row_n_fin = len(fin_data)
ts_fin = TableStyle([
    ('BACKGROUND',    (0,0), (-1,0), MID_BLUE),
    ('TEXTCOLOR',     (0,0), (-1,0), WHITE),
    ('FONTNAME',      (0,0), (-1,0), 'Helvetica-Bold'),
    ('FONTSIZE',      (0,0), (-1,-1), 8.5),
    ('BACKGROUND',    (0,1), (-1,1), colors.HexColor('#E0F7FA')),
    ('BACKGROUND',    (0,2), (-1,4), GREY_LIGHT),
    ('ROWBACKGROUNDS',(0,2), (-1,4), [GREY_LIGHT, WHITE]),
    ('BACKGROUND',    (0,5), (-1,5), DARK_BLUE),
    ('TEXTCOLOR',     (0,5), (-1,5), WHITE),
    ('FONTNAME',      (0,5), (-1,5), 'Helvetica-Bold'),
    ('BACKGROUND',    (0,6), (-1,6), colors.HexColor('#FFF9C4')),
    ('FONTNAME',      (0,6), (-1,6), 'Helvetica-Bold'),
    ('BACKGROUND',    (0,7), (-1,7), colors.HexColor('#D5F5E3')),
    ('FONTNAME',      (0,7), (-1,7), 'Helvetica-Bold'),
    ('ALIGN',         (1,0), (-1,-1), 'RIGHT'),
    ('ALIGN',         (0,0), (0,-1), 'LEFT'),
    ('GRID',          (0,0), (-1,-1), 0.3, colors.HexColor('#CCCCCC')),
    ('LEFTPADDING',   (0,0), (-1,-1), 7),
    ('RIGHTPADDING',  (0,0), (-1,-1), 7),
    ('TOPPADDING',    (0,0), (-1,-1), 5),
    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
])
fin_t = Table(fin_data, colWidths=[W*0.35, W*0.17, W*0.16, W*0.17, W*0.15])
fin_t.setStyle(ts_fin)
story.append(fin_t)
story.append(Spacer(1, 5*mm))

# ── ALLOKATIONS-HINWEIS ───────────────────────────────────────────────────────
story.append(HRFlowable(width=W, thickness=0.5, color=colors.HexColor('#AAAAAA')))
story.append(Spacer(1, 2*mm))
story.append(Paragraph('Methodik & Allokations-Grundsätze', S_H1))
notes = [
    '<b>Personalkosten:</b> CEO → D) Operations & Legal (Unternehmensführung, Fundraising, Recht). '
    'CTO → A) Software & AI Engineering (Produktentwicklung, Technik). '
    'CCO → B) Customer Acquisition (Vertrieb, Partnerschaften, Kundenwachstum).',
    '<b>Mitarbeiter:</b> SW Developer & Mech./Domain Engineer → A); '
    'Key Account, Customer Success, Marketing Assistant → B). '
    'Basis-Gehälter aus 2_Inputs R49–R83, Eintrittsmonate aus R177–R192.',
    '<b>Technologie:</b> AI/ML APIs & SaaS Tools → A) (Produktentwicklungs-Tools); '
    'Cloud Hosting Basis → C) (Infrastruktur der Plattform); '
    'Sicherheit & Compliance → C) (GDPR, IT-Security).',
    '<b>Marketing & Sales:</b> Alle Positionen (Paid Ads, Events, Sales Tools) → B) Customer Acquisition.',
    '<b>Finanzierung:</b> Carbon13-Betrag aus 2_Inputs!B9 (Hinweis: "C13 funding + GmbH Stammeinlage"). '
    'Business Angels aus 2_Inputs!B18/B20/B22. Zuflüsse verifiziert über 7_BS_CF!Zeile 9.',
]
for note in notes:
    story.append(Paragraph(f'• {note}', S_NOTE))

# ── FOOTER ────────────────────────────────────────────────────────────────────
story.append(Spacer(1, 4*mm))
story.append(HRFlowable(width=W, thickness=0.3, color=colors.HexColor('#CCCCCC')))
story.append(Spacer(1, 1*mm))
story.append(Paragraph(
    f'LoopforgeLab GmbH  |  Financial Model v19  |  '
    f'Erstellt: {datetime.now().strftime("%d.%m.%Y %H:%M")}  |  '
    f'Quelle: 260312_LFL_BM_Vorlage_normal_v19.xlsx',
    ParagraphStyle('footer', parent=S_SMALL, alignment=TA_CENTER, textColor=GREY_TEXT)
))

# ── Build ──────────────────────────────────────────────────────────────────────
doc.build(story)
print(f'✓ PDF gespeichert: {os.path.basename(OUT)}  ({os.path.getsize(OUT)/1024:.0f} KB)')
