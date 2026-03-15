"""
run_szenario.py – Erzeugt C13-Datei für ein bestimmtes Szenario (normal / stark / gering).

Ablauf:
  1. Kopiert Quelldatei, setzt Szenario-Zelle C1 im Sheet '1_Szenario'
  2. LibreOffice headless: recalkuliert Formeln, speichert als xlsx
  3. Überschreibt die SOURCE-Konstante in create_c13_formatted.py temporär
     → ruft es als Modul auf
  4. Benennt Ausgabe nach Konvention: ...normal_C13.xlsx / ...stark_C13.xlsx

Aufruf:
  python scripts/run_szenario.py normal
  python scripts/run_szenario.py stark
  python scripts/run_szenario.py gering
"""

import sys
import os
import shutil
import subprocess
import importlib
import tempfile
import time

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def main():
    if len(sys.argv) < 2:
        print("Aufruf: python scripts/run_szenario.py <gering|normal|stark>")
        sys.exit(1)

    szenario = sys.argv[1].strip().lower()
    if szenario not in ("gering", "normal", "stark"):
        print(f"Unbekanntes Szenario: {szenario}. Erlaubt: gering, normal, stark")
        sys.exit(1)

    import openpyxl

    original_source = os.path.join(BASE_DIR, '260312_LFL_BM_Vorlage_v19.xlsx')

    # ── Schritt 1: Quelldatei kopieren und Szenario setzen ────────────────────
    tmp_dir = tempfile.mkdtemp(prefix="lfl_szen_")
    tmp_src = os.path.join(tmp_dir, f'LFL_BM_v19_{szenario}.xlsx')
    shutil.copy2(original_source, tmp_src)
    print(f"[1/3] Kopie erstellt: {tmp_src}")

    wb_tmp = openpyxl.load_workbook(tmp_src)
    ws_scen = wb_tmp['1_Szenario']
    ws_scen.cell(row=1, column=3).value = szenario   # C1 = Szenario-Label
    wb_tmp.save(tmp_src)
    wb_tmp.close()
    print(f"      Szenario-Zelle C1 → '{szenario}'")

    # ── Schritt 2: LibreOffice Formel-Neuberechnung ───────────────────────────
    print(f"[2/3] LibreOffice recalkuliert Formeln …")
    lo_out = os.path.join(tmp_dir, 'calc_out')
    os.makedirs(lo_out, exist_ok=True)

    result = subprocess.run(
        [
            'libreoffice',
            '--headless',
            '--norestore',
            '--calc',
            f'--infilter=Calc MS Excel 2007 XML',
            '--convert-to', 'xlsx',
            '--outdir', lo_out,
            tmp_src,
        ],
        capture_output=True, text=True, timeout=120
    )
    if result.returncode != 0:
        print("LibreOffice stderr:", result.stderr[-500:])
        sys.exit(1)

    lo_file = os.path.join(lo_out, os.path.basename(tmp_src))
    if not os.path.exists(lo_file):
        # LO manchmal ändert den Namen
        files = os.listdir(lo_out)
        if files:
            lo_file = os.path.join(lo_out, files[0])
        else:
            print("LibreOffice hat keine Ausgabedatei erzeugt!")
            sys.exit(1)
    print(f"      Recalkuliert: {lo_file}")

    # ── Schritt 3: create_c13_formatted mit überschriebener SOURCE aufrufen ───
    print(f"[3/3] Erzeuge C13-Datei für Szenario '{szenario}' …")

    # Ausgabe-Pfad nach Konvention
    szen_label = szenario  # gering / normal / stark
    output_path = os.path.join(BASE_DIR, f'260312_LFL_BM_Vorlage_v19_{szen_label}_C13.xlsx')

    # Umgebungsvariablen für den Subprocess
    env = os.environ.copy()
    env['LFL_SOURCE_OVERRIDE'] = lo_file
    env['LFL_OUTPUT_OVERRIDE'] = output_path

    result2 = subprocess.run(
        [sys.executable, os.path.join(BASE_DIR, 'scripts', 'create_c13_formatted.py')],
        env=env, capture_output=False, cwd=BASE_DIR
    )

    # Aufräumen
    shutil.rmtree(tmp_dir, ignore_errors=True)

    if result2.returncode != 0:
        print("Fehler beim Erzeugen der C13-Datei.")
        sys.exit(1)

    print(f"\n✓ Fertig: {output_path}")

if __name__ == '__main__':
    main()
