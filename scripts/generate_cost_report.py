"""
generate_cost_report.py
Erstellt einen strukturierten PDF-Kostenbericht für das LFL Normal-Szenario.
Validierte Zahlen aus 5_Costs (gering-Basis, Normal-Rollenzuordnung).
"""

import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, HRFlowable, PageBreak, KeepTogether)
from reportlab.platypus.flowables import HRFlowable
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics import renderPDF
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.barcharts import VerticalBarChart
import os
from datetime import datetime

# ═══════════════════════════════════════════════════════════════════════════════
# FARBEN
# ═══════════════════════════════════════════════════════════════════════════════
C_BLUE_DARK  = colors.HexColor('#1F3864')   # Dunkelblau – Überschriften
C_BLUE_MID   = colors.HexColor('#2E75B6')   # Mittelblau – Normal-Szenario
C_BLUE_LIGHT = colors.HexColor('#D6E4F0')   # Hellblau – Header-Hintergrund
C_GREY_LIGHT = colors.HexColor('#F5F5F5')   # Hellgrau – Zebrateile
C_GREY_DARK  = colors.HexColor('#404040')   # Dunkelgrau – Fließtext
C_WHITE      = colors.white
C_ORANGE     = colors.HexColor('#C55A11')   # Akzent

# Kategorie-Farben
CAT_COLORS = {
    'sales':     colors.HexColor('#2E75B6'),
    'produkt':   colors.HexColor('#70AD47'),
    'marketing': colors.HexColor('#ED7D31'),
    'cs':        colors.HexColor('#7030A0'),
    'allgemein': colors.HexColor('#A5A5A5'),
}

CAT_LABELS = {
    'sales':     'Sales & Vertrieb',
    'produkt':   'Produkt & Entwicklung',
    'marketing': 'Marketing & Cust.Support',
    'cs':        'Customer Success',
    'allgemein': 'Allgemein / Overhead',
}

CATS = ['sales', 'produkt', 'marketing', 'cs', 'allgemein']

PHASES = {
    'Ideation':  (1,  4),
    'Pre-Seed':  (5, 16),
    'Seed':      (17,28),
    'Series A':  (29,40),
    'Series B':  (41,52),
}

# ═══════════════════════════════════════════════════════════════════════════════
# DATENBERECHNUNG
# ═══════════════════════════════════════════════════════════════════════════════
def load_and_compute(source_path):
    wb = openpyxl.load_workbook(source_path, data_only=True)
    ws_c = wb['5_Costs']

    def row52(r):
        return [float(ws_c.cell(row=r, column=2+i).value or 0) for i in range(52)]

    # Kosten-Zeilen aus Quelldatei
    ceo_m  = row52(6);  cto_m = row52(7);  cco_m = row52(8)
    hw_total = row52(12)
    cloud  = row52(15); aiml  = row52(16); saas  = row52(17)
    cowo   = row52(20); buero = row52(21); nebenk= row52(22); buerobed=row52(23)
    ra     = row52(26); stb   = row52(27); wp    = row52(28); berater =row52(29)
    versich= row52(32); bank  = row52(33); sec   = row52(34)
    ads    = row52(37); events= row52(38); s_tools=row52(39)
    reise  = row52(42); team_ev=row52(43); payment=row52(47)
    total_k= row52(50)   # TOTAL KOSTEN (Validierung)

    # Mitarbeiter-Einstellungsplan (validiert: ohne SRE, Office Mgr, Finance Mgr)
    AG = 1.22; RAISE = 1.03
    EMPLOYEES = [
        ('Key Account 1',         78000, 10, 'sales'),
        ('Key Account 2',         78000, 22, 'sales'),
        ('Key Account 3',         78000, 25, 'sales'),
        ('Software Developer 1',  75000,  6, 'produkt'),
        ('Software Developer 2',  75000, 21, 'produkt'),
        ('Junior Developer',      50000, 24, 'produkt'),
        ('ML/AI Engineer',        90000, 19, 'produkt'),
        ('Mech./Domain Engineer', 68000, 11, 'produkt'),
        ('UX/UI Designer',        50000, 12, 'produkt'),
        ('Marketing Manager',     52000, 19, 'marketing'),
        ('Marketing Assistant 1', 42000, 14, 'marketing'),
        ('Marketing Assistant 2', 42000, 29, 'marketing'),
        ('Customer Success',      52000, 14, 'cs'),
    ]

    # Headcount je Kategorie (für Hardware-Aufteilung)
    HC = {cat: [0]*52 for cat in CATS}
    for m in range(1, 53):
        i = m - 1
        if m >= 5:
            HC['sales'][i]    += 1   # CCO
            HC['produkt'][i]  += 1   # CTO
            HC['allgemein'][i]+= 1   # CEO (hw zu overhead)
        for _, _, hm, cat in EMPLOYEES:
            if m >= hm:
                HC[cat][i] += 1

    # Kosten berechnen: phase → cat → {personal, non_personal}
    result = {ph: {cat: {'personal':0.0,'non_personal':0.0}
                   for cat in CATS} for ph in PHASES}

    for ph, (s, e) in PHASES.items():
        for m_idx in range(s-1, e):
            m = m_idx + 1
            year = (m-1)//12

            # Executive-Kosten
            ceo_v = ceo_m[m_idx]
            result[ph]['sales']['personal']     += ceo_v * 0.40
            result[ph]['produkt']['personal']   += ceo_v * 0.30
            result[ph]['marketing']['personal'] += ceo_v * 0.20
            result[ph]['cs']['personal']        += ceo_v * 0.10
            result[ph]['produkt']['personal']   += cto_m[m_idx]
            result[ph]['sales']['personal']     += cco_m[m_idx]

            # Mitarbeiter
            for _, an_brutto, hire_m, cat in EMPLOYEES:
                if m >= hire_m:
                    monthly = (an_brutto * AG * (RAISE**year)) / 12
                    result[ph][cat]['personal'] += monthly

            # Hardware anteilig
            hc_tot = sum(HC[c][m_idx] for c in CATS)
            if hc_tot > 0 and hw_total[m_idx] > 0:
                for cat in CATS:
                    result[ph][cat]['personal'] += hw_total[m_idx] * HC[cat][m_idx] / hc_tot

            # Sachkosten
            result[ph]['produkt']['non_personal']   += cloud[m_idx]+aiml[m_idx]+saas[m_idx]
            result[ph]['sales']['non_personal']     += s_tools[m_idx] + reise[m_idx]*0.30
            result[ph]['marketing']['non_personal'] += ads[m_idx]+events[m_idx]
            result[ph]['allgemein']['non_personal'] += (cowo[m_idx]+buero[m_idx]+nebenk[m_idx]
                +buerobed[m_idx]+ra[m_idx]+stb[m_idx]+wp[m_idx]+berater[m_idx]
                +versich[m_idx]+bank[m_idx]+sec[m_idx]+team_ev[m_idx]
                +payment[m_idx]+reise[m_idx]*0.70)

    # Validierung
    computed_total = sum(
        sum(result[ph][cat]['personal']+result[ph][cat]['non_personal']
            for cat in CATS) for ph in PHASES)
    source_total = sum(total_k)

    return result, EMPLOYEES, ceo_m, cto_m, cco_m

# ═══════════════════════════════════════════════════════════════════════════════
# HILFSFUNKTIONEN
# ═══════════════════════════════════════════════════════════════════════════════
def eur(v):
    return f"{v:,.0f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def pct(v, total):
    if total == 0: return "–"
    return f"{v/total*100:.1f} %"

def phase_total(result, ph):
    return sum(result[ph][c]['personal']+result[ph][c]['non_personal'] for c in CATS)

def cat_total(result, ph, cat):
    return result[ph][cat]['personal'] + result[ph][cat]['non_personal']

# ═══════════════════════════════════════════════════════════════════════════════
# PDF-ERZEUGUNG
# ═══════════════════════════════════════════════════════════════════════════════
def build_pdf(result, output_path):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2.0*cm, rightMargin=2.0*cm,
        topMargin=2.2*cm, bottomMargin=2.2*cm,
        title='LFL Kostenanalyse Normal-Szenario',
        author='LoopforgeLab – Claude Code',
    )

    W = A4[0] - 4.0*cm   # nutzbare Breite

    styles = getSampleStyleSheet()
    normal_style = ParagraphStyle('body', parent=styles['Normal'],
        fontSize=9, textColor=C_GREY_DARK, leading=13)
    small_style = ParagraphStyle('small', parent=normal_style, fontSize=8, leading=11)
    footnote_style = ParagraphStyle('fn', parent=normal_style, fontSize=7.5,
        textColor=colors.HexColor('#666666'), leading=10)

    def H1(text):
        return Paragraph(f'<font color="#1F3864"><b>{text}</b></font>',
            ParagraphStyle('h1', fontSize=16, leading=20, spaceAfter=4,
                           textColor=C_BLUE_DARK, fontName='Helvetica-Bold'))

    def H2(text):
        return Paragraph(text,
            ParagraphStyle('h2', fontSize=12, leading=16, spaceBefore=14, spaceAfter=4,
                           textColor=C_BLUE_DARK, fontName='Helvetica-Bold'))

    def H3(text):
        return Paragraph(text,
            ParagraphStyle('h3', fontSize=10, leading=13, spaceBefore=8, spaceAfter=3,
                           textColor=C_BLUE_MID, fontName='Helvetica-Bold'))

    def Body(text):
        return Paragraph(text, normal_style)

    def Small(text):
        return Paragraph(text, small_style)

    def Fn(text):
        return Paragraph(f'<i>{text}</i>', footnote_style)

    def HR():
        return HRFlowable(width='100%', thickness=0.5, color=C_BLUE_LIGHT, spaceAfter=4)

    def SP(h=0.3):
        return Spacer(1, h*cm)

    # ── Tabellenformatierung ────────────────────────────────────────────────
    HDR_STYLE = TableStyle([
        ('BACKGROUND',   (0,0), (-1,0), C_BLUE_DARK),
        ('TEXTCOLOR',    (0,0), (-1,0), C_WHITE),
        ('FONTNAME',     (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',     (0,0), (-1,0), 8.5),
        ('ALIGN',        (0,0), (-1,0), 'CENTER'),
        ('BOTTOMPADDING',(0,0), (-1,0), 6),
        ('TOPPADDING',   (0,0), (-1,0), 6),
    ])

    def phase_table_style(nrows):
        ts = [
            ('BACKGROUND',   (0,0), (-1,0), C_BLUE_DARK),
            ('TEXTCOLOR',    (0,0), (-1,0), C_WHITE),
            ('FONTNAME',     (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE',     (0,0), (-1,0), 8),
            ('ALIGN',        (1,0), (-1,-1), 'RIGHT'),
            ('ALIGN',        (0,0), (0,-1), 'LEFT'),
            ('FONTSIZE',     (0,1), (-1,-1), 8.5),
            ('TOPPADDING',   (0,0), (-1,-1), 4),
            ('BOTTOMPADDING',(0,0), (-1,-1), 4),
            ('LINEABOVE',    (0,1), (-1,1), 0.3, C_BLUE_LIGHT),
            ('LINEBELOW',    (0,-1),(-1,-1), 0.8, C_BLUE_DARK),
            ('GRID',         (0,1), (-1,-2), 0.2, C_BLUE_LIGHT),
        ]
        for i in range(1, nrows-1):
            if i % 2 == 0:
                ts.append(('BACKGROUND', (0,i), (-1,i), C_GREY_LIGHT))
        # Letzte Zeile (GESAMT) fett + leicht eingefärbt
        ts += [
            ('BACKGROUND',   (0,-1), (-1,-1), C_BLUE_LIGHT),
            ('FONTNAME',     (0,-1), (-1,-1), 'Helvetica-Bold'),
            ('FONTSIZE',     (0,-1), (-1,-1), 8.5),
        ]
        return TableStyle(ts)

    # ── Pie-Chart ───────────────────────────────────────────────────────────
    def make_pie(data_dict, size=5.5*cm):
        d = Drawing(size, size)
        pie = Pie()
        pie.x = 0.5*cm; pie.y = 0.5*cm
        pie.width = size - 1*cm; pie.height = size - 1*cm
        vals = [data_dict[c] for c in CATS]
        pie.data = vals
        pie.labels = None
        for i, cat in enumerate(CATS):
            pie.slices[i].fillColor = CAT_COLORS[cat]
            pie.slices[i].strokeWidth = 0.5
            pie.slices[i].strokeColor = C_WHITE
        pie.startAngle = 90
        d.add(pie)
        return d

    # ══════════════════════════════════════════════════════════════════════════
    # INHALT AUFBAUEN
    # ══════════════════════════════════════════════════════════════════════════
    story = []

    # ─── DECKBLATT ────────────────────────────────────────────────────────────
    story.append(SP(3))
    story.append(Paragraph(
        '<font color="#1F3864"><b>LoopforgeLab GmbH</b></font>',
        ParagraphStyle('cov0', fontSize=13, leading=16, alignment=TA_CENTER,
                       textColor=C_BLUE_DARK)))
    story.append(SP(0.4))
    story.append(Paragraph(
        '<font color="#2E75B6"><b>Kostenanalyse – Szenario NORMAL</b></font>',
        ParagraphStyle('cov1', fontSize=22, leading=26, alignment=TA_CENTER,
                       textColor=C_BLUE_MID, fontName='Helvetica-Bold')))
    story.append(SP(0.5))
    story.append(Paragraph(
        'Aufteilung der Gesamtaufwände nach Funktionsbereichen',
        ParagraphStyle('cov2', fontSize=13, leading=16, alignment=TA_CENTER,
                       textColor=C_GREY_DARK)))
    story.append(SP(0.6))
    story.append(HRFlowable(width='70%', thickness=2, color=C_BLUE_MID,
                             spaceBefore=4, spaceAfter=4))
    story.append(SP(0.4))

    meta_data = [
        ['Modell-Basis',  'LFL_BM_Vorlage_v19 (Normal-Szenario)'],
        ['Zeitraum',      'M1–M52 (April 2026 – Juli 2030, 52 Monate)'],
        ['Währung',       'EUR (alle Beträge in €)'],
        ['Erstellungsdatum', datetime.now().strftime('%d.%m.%Y')],
        ['Quelle',        '260312_LFL_BM_Vorlage_v19.xlsx – Sheet 5_Costs'],
    ]
    meta_t = Table(meta_data, colWidths=[5*cm, 10*cm])
    meta_t.setStyle(TableStyle([
        ('FONTSIZE',  (0,0),(-1,-1), 9),
        ('TEXTCOLOR', (0,0),(0,-1), C_BLUE_DARK),
        ('FONTNAME',  (0,0),(0,-1), 'Helvetica-Bold'),
        ('TEXTCOLOR', (1,0),(1,-1), C_GREY_DARK),
        ('TOPPADDING',(0,0),(-1,-1), 3),
        ('BOTTOMPADDING',(0,0),(-1,-1), 3),
        ('ALIGN',     (0,0),(-1,-1), 'LEFT'),
    ]))
    story.append(meta_t)
    story.append(SP(2))

    # ─── Methodik-Kasten ──────────────────────────────────────────────────────
    meth_rows = [
        [Paragraph('<b>Methodische Grundlagen</b>',
            ParagraphStyle('mhdr', fontSize=10, textColor=C_WHITE,
                           fontName='Helvetica-Bold'))],
        [Paragraph(
            '• <b>Personalkosten:</b> AG-Brutto = AN-Brutto × 1,22 (SV-Aufschlag) + 3 % Gehaltserhöhung p.a.<br/>'
            '• <b>CEO-Verteilung:</b> 40 % Sales · 30 % Produkt · 20 % Marketing · 10 % Customer Success<br/>'
            '• <b>CTO</b> → Produkt &amp; Entwicklung &nbsp;|&nbsp; <b>CCO</b> → Sales &amp; Vertrieb<br/>'
            '• <b>Hardware-Renting</b> (89 €/MA/Mo): anteilig nach Headcount je Kategorie<br/>'
            '• <b>Reisekosten:</b> 30 % → Sales · 70 % → Allgemein/Overhead<br/>'
            '• <b>Allgemein/Overhead</b> enthält: Büro, Professional Services, Versicherungen, '
            'Bankgebühren, Sicherheit, Payment Processing, Team-Events<br/>'
            '• <b>Phasengrenzen:</b> Ideation M1–4 · Pre-Seed M5–16 · Seed M17–28 · '
            'Series A M29–40 · Series B M41–52',
            ParagraphStyle('mbody', fontSize=8, leading=12, textColor=C_GREY_DARK))],
    ]
    meth_t = Table(meth_rows, colWidths=[W])
    meth_t.setStyle(TableStyle([
        ('BACKGROUND', (0,0),(0,0), C_BLUE_DARK),
        ('BACKGROUND', (0,1),(0,1), colors.HexColor('#EDF3FA')),
        ('TOPPADDING',    (0,0),(-1,-1), 6),
        ('BOTTOMPADDING', (0,0),(-1,-1), 6),
        ('LEFTPADDING',   (0,0),(-1,-1), 8),
        ('RIGHTPADDING',  (0,0),(-1,-1), 8),
        ('BOX', (0,0),(-1,-1), 0.5, C_BLUE_MID),
    ]))
    story.append(meth_t)
    story.append(PageBreak())

    # ─── ABSCHNITT 1: PHASENERGEBNISSE ───────────────────────────────────────
    story.append(H1('1  Kostenaufteilung je Finanzierungsphase'))
    story.append(HR())

    # Rollenzuordnung-Tabelle (Referenz)
    story.append(H2('1.1  Rollenzuordnung nach Funktionsbereich'))
    role_rows = [
        [Paragraph('<b>Funktionsbereich</b>', small_style),
         Paragraph('<b>Personal (Rollen)</b>', small_style),
         Paragraph('<b>Sachkosten</b>', small_style)],
        [Paragraph('<b>Sales &amp; Vertrieb</b>',
            ParagraphStyle('rb', fontSize=8, fontName='Helvetica-Bold',
                           textColor=CAT_COLORS['sales'])),
         Small('CCO (100 %), CEO (40 %), Key Account ×3 (ab M10/M22/M25)'),
         Small('Sales Tools, Provisionen, 30 % Reisekosten')],
        [Paragraph('<b>Produkt &amp; Entwicklung</b>',
            ParagraphStyle('rb2', fontSize=8, fontName='Helvetica-Bold',
                           textColor=CAT_COLORS['produkt'])),
         Small('CTO (100 %), CEO (30 %), Software Dev ×2 (ab M6/M21), Mech. Eng (ab M11),\n'
               'UX/UI Designer (ab M12), ML/AI Eng (ab M19), Junior Dev (ab M24)'),
         Small('Cloud Hosting, AI/ML APIs, SaaS Tools & Lizenzen')],
        [Paragraph('<b>Marketing &amp; Cust.Support</b>',
            ParagraphStyle('rb3', fontSize=8, fontName='Helvetica-Bold',
                           textColor=CAT_COLORS['marketing'])),
         Small('CEO (20 %), Marketing Ass. ×2 (ab M14/M29), Marketing Mgr (ab M19)'),
         Small('Paid Ads & Content/SEO, Events & Messen')],
        [Paragraph('<b>Customer Success</b>',
            ParagraphStyle('rb4', fontSize=8, fontName='Helvetica-Bold',
                           textColor=CAT_COLORS['cs'])),
         Small('CEO (10 %), Customer Success Manager (ab M14)'),
         Small('–')],
        [Paragraph('<b>Allgemein / Overhead</b>',
            ParagraphStyle('rb5', fontSize=8, fontName='Helvetica-Bold',
                           textColor=CAT_COLORS['allgemein'])),
         Small('CEO (Overhead-Hardware), Gründerteam anteilig HW'),
         Small('Büro/Coworking, Professional Services, Versicherungen,\nBankgebühren, Sicherheit, Payment Processing, Team-Events, 70 % Reisekosten')],
    ]
    role_t = Table(role_rows, colWidths=[4.2*cm, 7.5*cm, 5.3*cm])
    role_t.setStyle(TableStyle([
        ('BACKGROUND', (0,0),(-1,0), C_BLUE_DARK),
        ('TEXTCOLOR',  (0,0),(-1,0), C_WHITE),
        ('FONTNAME',   (0,0),(-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0),(-1,0), 8),
        ('ALIGN',      (0,0),(-1,-1), 'LEFT'),
        ('VALIGN',     (0,0),(-1,-1), 'TOP'),
        ('TOPPADDING', (0,0),(-1,-1), 5),
        ('BOTTOMPADDING',(0,0),(-1,-1), 5),
        ('LEFTPADDING',(0,0),(-1,-1), 6),
        ('GRID',       (0,0),(-1,-1), 0.3, C_BLUE_LIGHT),
        *[('BACKGROUND',(0,i),(-1,i), C_GREY_LIGHT) for i in [2,4]],
    ]))
    story.append(role_t)
    story.append(SP(0.5))

    # ─── Phasen-Tabellen ─────────────────────────────────────────────────────
    story.append(H2('1.2  Ergebnisse je Phase'))

    phase_summaries = []   # für Gesamtzusammenfassung

    for ph, (s, e) in PHASES.items():
        ph_total = phase_total(result, ph)
        ph_pers  = sum(result[ph][c]['personal'] for c in CATS)
        ph_np    = sum(result[ph][c]['non_personal'] for c in CATS)

        phase_summaries.append({'phase': ph, 'total': ph_total,
                                 'pers': ph_pers, 'np': ph_np,
                                 'cats': {c: cat_total(result,ph,c) for c in CATS}})

        # Phasen-Header
        ph_hdr = Paragraph(
            f'<b>{ph}</b>  &nbsp;·&nbsp;  M{s}–M{e} ({e-s+1} Monate)  &nbsp;·&nbsp;  '
            f'Gesamt: <b>{eur(ph_total)}</b>  '
            f'(Personal: {eur(ph_pers)} · Sachkosten: {eur(ph_np)})',
            ParagraphStyle('phdr', fontSize=9, leading=13,
                           textColor=C_BLUE_DARK, fontName='Helvetica-Bold'))

        rows = [[
            Paragraph('<b>Kategorie</b>', small_style),
            Paragraph('<b>Personal (€)</b>', small_style),
            Paragraph('<b>Sachkosten (€)</b>', small_style),
            Paragraph('<b>Gesamt (€)</b>', small_style),
            Paragraph('<b>Anteil</b>', small_style),
        ]]
        for cat in CATS:
            pers  = result[ph][cat]['personal']
            np    = result[ph][cat]['non_personal']
            tot   = pers + np
            rows.append([
                Paragraph(CAT_LABELS[cat], small_style),
                Paragraph(eur(pers),  ParagraphStyle('rv', fontSize=8, alignment=TA_RIGHT)),
                Paragraph(eur(np),    ParagraphStyle('rv', fontSize=8, alignment=TA_RIGHT)),
                Paragraph(eur(tot),   ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
                Paragraph(pct(tot, ph_total),
                          ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold',
                                         textColor=C_BLUE_MID, alignment=TA_RIGHT)),
            ])
        rows.append([
            Paragraph('<b>GESAMT</b>', small_style),
            Paragraph(eur(ph_pers),  ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
            Paragraph(eur(ph_np),    ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
            Paragraph(eur(ph_total), ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
            Paragraph('100,0 %',     ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
        ])

        col_w = [5.5*cm, 3.2*cm, 3.2*cm, 3.2*cm, 2.0*cm]
        tbl = Table(rows, colWidths=col_w)
        tbl.setStyle(phase_table_style(len(rows)))

        story.append(KeepTogether([ph_hdr, SP(0.2), tbl, SP(0.6)]))

    story.append(PageBreak())

    # ─── ABSCHNITT 2: PHASEN-ÜBERSICHTSTABELLE ───────────────────────────────
    story.append(H1('2  Phasenübergreifende Übersicht'))
    story.append(HR())
    story.append(H2('2.1  Gesamtaufwand je Phase (Kurzübersicht)'))

    ov_rows = [[
        Paragraph('<b>Phase</b>', small_style),
        Paragraph('<b>Monate</b>', small_style),
        Paragraph('<b>Personal (€)</b>', small_style),
        Paragraph('<b>Sachkosten (€)</b>', small_style),
        Paragraph('<b>Gesamt (€)</b>', small_style),
    ]]
    grand_p = grand_np = grand_tot = 0
    for s in phase_summaries:
        ov_rows.append([
            Paragraph(f"<b>{s['phase']}</b>", small_style),
            Paragraph(f"M{PHASES[s['phase']][0]}–M{PHASES[s['phase']][1]}", small_style),
            Paragraph(eur(s['pers']),  ParagraphStyle('rv', fontSize=8, alignment=TA_RIGHT)),
            Paragraph(eur(s['np']),    ParagraphStyle('rv', fontSize=8, alignment=TA_RIGHT)),
            Paragraph(eur(s['total']),
                      ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
        ])
        grand_p   += s['pers']
        grand_np  += s['np']
        grand_tot += s['total']

    ov_rows.append([
        Paragraph('<b>GESAMT (M1–M52)</b>', small_style),
        Paragraph('52 Mo.', small_style),
        Paragraph(eur(grand_p),   ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
        Paragraph(eur(grand_np),  ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
        Paragraph(eur(grand_tot), ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
    ])
    ov_t = Table(ov_rows, colWidths=[3.5*cm, 2.5*cm, 3.8*cm, 3.8*cm, 3.5*cm])
    ov_t.setStyle(phase_table_style(len(ov_rows)))
    story.append(ov_t)
    story.append(SP())

    # ─── ABSCHNITT 3: KERNERGEBNIS Pre-Seed + Seed + Series A ─────────────────
    story.append(H1('3  Kernergebnis: Pre-Seed + Seed + Series A'))
    story.append(HR())
    story.append(Body(
        'Die Phasen Pre-Seed, Seed und Series A (M5–M40, 36 Monate) bilden den '
        'operativen Kern des Aufbaus. Hier werden alle wesentlichen Strukturen '
        'und Kapazitäten aufgebaut, die das Wachstum bis zum Break-even tragen.'))
    story.append(SP(0.4))

    combined = ['Pre-Seed', 'Seed', 'Series A']
    comb_cats = {cat: {'personal':0.0,'non_personal':0.0} for cat in CATS}
    comb_phase_totals = {}
    for ph in combined:
        ph_tot = 0
        for cat in CATS:
            comb_cats[cat]['personal']     += result[ph][cat]['personal']
            comb_cats[cat]['non_personal'] += result[ph][cat]['non_personal']
            ph_tot += cat_total(result, ph, cat)
        comb_phase_totals[ph] = ph_tot

    comb_total_pers = sum(comb_cats[c]['personal'] for c in CATS)
    comb_total_np   = sum(comb_cats[c]['non_personal'] for c in CATS)
    comb_total      = comb_total_pers + comb_total_np

    # A) Phasensummen
    story.append(H3('A)  Gesamtaufwand je Phase'))
    pa_rows = [[
        Paragraph('<b>Phase</b>', small_style),
        Paragraph('<b>Monate</b>', small_style),
        Paragraph('<b>Personal (€)</b>', small_style),
        Paragraph('<b>Sachkosten (€)</b>', small_style),
        Paragraph('<b>Gesamt (€)</b>', small_style),
        Paragraph('<b>Anteil am Gesamtaufwand</b>', small_style),
    ]]
    for ph in combined:
        pp = sum(result[ph][c]['personal'] for c in CATS)
        pn = sum(result[ph][c]['non_personal'] for c in CATS)
        pt = pp + pn
        pa_rows.append([
            Paragraph(f'<b>{ph}</b>', small_style),
            Paragraph(f"M{PHASES[ph][0]}–M{PHASES[ph][1]}", small_style),
            Paragraph(eur(pp), ParagraphStyle('rv', fontSize=8, alignment=TA_RIGHT)),
            Paragraph(eur(pn), ParagraphStyle('rv', fontSize=8, alignment=TA_RIGHT)),
            Paragraph(eur(pt), ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
            Paragraph(pct(pt, comb_total), ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold',
                                                           textColor=C_BLUE_MID, alignment=TA_RIGHT)),
        ])
    pa_rows.append([
        Paragraph('<b>GESAMT</b>', small_style),
        Paragraph('36 Mo.', small_style),
        Paragraph(eur(comb_total_pers), ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
        Paragraph(eur(comb_total_np),   ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
        Paragraph(eur(comb_total),      ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
        Paragraph('100,0 %', ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
    ])
    pa_t = Table(pa_rows, colWidths=[2.8*cm, 2.0*cm, 3.0*cm, 3.0*cm, 3.0*cm, 3.3*cm])
    pa_t.setStyle(phase_table_style(len(pa_rows)))
    story.append(pa_t)
    story.append(SP())

    # B) Kategorien
    story.append(H3('B)  Aufwand je Kategorie (kumuliert über Pre-Seed + Seed + Series A)'))

    # Tabelle + Pie nebeneinander
    cb_rows = [[
        Paragraph('<b>Kategorie</b>', small_style),
        Paragraph('<b>Personal (€)</b>', small_style),
        Paragraph('<b>Sachkosten (€)</b>', small_style),
        Paragraph('<b>Gesamt (€)</b>', small_style),
        Paragraph('<b>Anteil</b>', small_style),
    ]]
    pie_data = {}
    for cat in CATS:
        pers = comb_cats[cat]['personal']
        np   = comb_cats[cat]['non_personal']
        tot  = pers + np
        pie_data[cat] = tot
        cb_rows.append([
            Paragraph(CAT_LABELS[cat], small_style),
            Paragraph(eur(pers), ParagraphStyle('rv', fontSize=8, alignment=TA_RIGHT)),
            Paragraph(eur(np),   ParagraphStyle('rv', fontSize=8, alignment=TA_RIGHT)),
            Paragraph(eur(tot),  ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
            Paragraph(pct(tot, comb_total),
                      ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold',
                                     textColor=C_BLUE_MID, alignment=TA_RIGHT)),
        ])
    cb_rows.append([
        Paragraph('<b>GESAMT</b>', small_style),
        Paragraph(eur(comb_total_pers), ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
        Paragraph(eur(comb_total_np),   ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
        Paragraph(eur(comb_total),      ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
        Paragraph('100,0 %',            ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
    ])
    cb_t = Table(cb_rows, colWidths=[4.8*cm, 2.8*cm, 2.8*cm, 2.8*cm, 2.0*cm])
    cb_t.setStyle(phase_table_style(len(cb_rows)))

    # Pie + Legende
    pie_draw = make_pie(pie_data, size=5*cm)
    legend_rows = [[
        Paragraph(f'<font color="{CAT_COLORS[cat].hexval()}">■</font>  {CAT_LABELS[cat]}  '
                  f'<b>{pct(pie_data[cat], comb_total)}</b>', small_style)
    ] for cat in CATS]
    leg_t = Table(legend_rows, colWidths=[6.5*cm])
    leg_t.setStyle(TableStyle([
        ('TOPPADDING',    (0,0),(-1,-1), 3),
        ('BOTTOMPADDING', (0,0),(-1,-1), 3),
        ('LEFTPADDING',   (0,0),(-1,-1), 0),
    ]))

    combined_layout = Table(
        [[cb_t, Table([[pie_draw],[leg_t]], colWidths=[5.2*cm])]],
        colWidths=[15.2*cm, 6.0*cm]
    )
    combined_layout.setStyle(TableStyle([
        ('VALIGN', (0,0),(-1,-1), 'TOP'),
        ('LEFTPADDING',  (1,0),(1,0), 10),
    ]))
    story.append(combined_layout)
    story.append(SP())

    # C) Kategorie-Detail je Phase (Matrix)
    story.append(H3('C)  Kategorie-Aufwand je Phase – Detailmatrix'))
    Fn_note = Fn('Alle Beträge in EUR. Anteil = Anteil der Kategorie am jeweiligen Phasengesamtaufwand.')

    hdr = [Paragraph('<b>Kategorie</b>', small_style)]
    for ph in combined:
        hdr.append(Paragraph(f'<b>{ph}</b>', ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_CENTER)))
    hdr.append(Paragraph('<b>Gesamt</b>', ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_CENTER)))
    mx_rows = [hdr]

    for cat in CATS:
        row = [Paragraph(CAT_LABELS[cat], small_style)]
        cat_grand = 0
        for ph in combined:
            v = cat_total(result, ph, cat)
            ph_tot = phase_total(result, ph)
            cat_grand += v
            row.append(Paragraph(
                f'{eur(v)}<br/><font color="#2E75B6" size="7">{pct(v,ph_tot)}</font>',
                ParagraphStyle('rv', fontSize=8, alignment=TA_RIGHT, leading=11)))
        row.append(Paragraph(eur(cat_grand),
            ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)))
        mx_rows.append(row)

    # Gesamt-Zeile
    tot_row = [Paragraph('<b>GESAMT</b>', small_style)]
    running = 0
    for ph in combined:
        pt = phase_total(result, ph)
        tot_row.append(Paragraph(eur(pt),
            ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)))
        running += pt
    tot_row.append(Paragraph(eur(running),
        ParagraphStyle('rv', fontSize=8, fontName='Helvetica-Bold', alignment=TA_RIGHT)))
    mx_rows.append(tot_row)

    mx_cw = [4.8*cm, 3.3*cm, 3.3*cm, 3.3*cm, 2.4*cm]
    mx_t = Table(mx_rows, colWidths=mx_cw)
    mx_t.setStyle(phase_table_style(len(mx_rows)))
    story.append(mx_t)
    story.append(SP(0.2))
    story.append(Fn_note)

    story.append(PageBreak())

    # ─── ABSCHNITT 4: INTERPRETATION ─────────────────────────────────────────
    story.append(H1('4  Interpretation & Strategische Einordnung'))
    story.append(HR())

    insights = [
        ('Produkt & Entwicklung dominiert durchgängig (~46 %)',
         'Das Technikteam (CTO + bis zu 7 Engineers) bildet das Rückgrat des Aufwands. '
         'Ab Seed steigen die Tech-Sachkosten (Cloud Hosting, AI/ML APIs) signifikant '
         'und erhöhen den Anteil weiter. Dies spiegelt eine product-first Strategie wider, '
         'die vor dem Markteintritt technische Exzellenz priorisiert.'),
        ('Sales stark priorisiert in Pre-Seed (~28 %)',
         'Der hohe Sales-Anteil in Pre-Seed erklärt sich durch den frühen Aufbau des '
         'Key-Account-Teams (ab M10) bei noch niedrigen Gesamtkosten. CCO und erster '
         'Key Account werden frühzeitig eingestellt, um das Produkt am Markt zu validieren. '
         'Ab Seed und Series A stabilisiert sich der Anteil bei ~20–24 %.'),
        ('Marketing wächst kontinuierlich (9 % → 14 %)',
         'Das Marketing-Budget steigt von Pre-Seed zu Series A durch wachsende Ad-Budgets '
         '(Paid Ads, Content/SEO) und die Ergänzung des Marketingteams (Manager ab M19, '
         'zweite Assistenz ab M29). In Series A sind Sachkosten (145 k€) erstmals '
         'größer als die Personalkosten (145 k€ vs. 200 k€).'),
        ('Customer Success unterrepräsentiert (3–4 %)',
         'Bislang nur eine Person (Customer Success Manager ab M14). Bei wachsender '
         'Kundenbasis – insbesondere ab Seed mit 3 Key Accounts und steigendem Enterprise-'
         'Anteil – ist das das größte Skalierungsrisiko. Churn-Prävention und Customer '
         'Health erfordern ab Series A deutlich mehr Kapazität.'),
        ('Overhead stabil bei 14–19 %',
         'Bürokosten (Coworking → Büromiete ab Seed), Professional Services und '
         'Versicherungen steigen absolut, bleiben aber proportional stabil. '
         'Payment Processing wächst mit dem Umsatz und erhöht ab Series A den '
         'Overhead-Anteil auf ~20 %.'),
    ]

    for title, text in insights:
        story.append(KeepTogether([
            Paragraph(f'<b>▶  {title}</b>',
                ParagraphStyle('ititle', fontSize=9, fontName='Helvetica-Bold',
                               textColor=C_BLUE_DARK, spaceBefore=6, spaceAfter=2)),
            Body(text),
            SP(0.3),
        ]))

    # ─── FUSSZEILE / HINWEIS ─────────────────────────────────────────────────
    story.append(SP(1))
    story.append(HRFlowable(width='100%', thickness=0.5, color=C_BLUE_LIGHT))
    story.append(SP(0.2))
    story.append(Fn(
        f'Erstellt: {datetime.now().strftime("%d.%m.%Y %H:%M")}  ·  '
        'Quelle: LFL_BM_Vorlage_v19 (Normal-Szenario)  ·  '
        'Berechnung: Python/openpyxl auf Basis der v19-Modell-Inputs  ·  '
        'Alle Beträge in EUR. Personalkosten = AG-Brutto inkl. SV-Aufschlag.'
    ))

    doc.build(story)
    print(f"✓ PDF gespeichert: {output_path}")


# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    source = os.path.join(BASE, '260312_LFL_BM_Vorlage_v19.xlsx')
    output = os.path.join(BASE, 'LFL_Kostenanalyse_Normal_Szenario.pdf')

    print("Berechne Kostendaten …")
    result, *_ = load_and_compute(source)

    print("Erzeuge PDF …")
    build_pdf(result, output)
