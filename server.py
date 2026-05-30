"""
HERM Compiler — Server locale con Flask
Servono i file statici + gestisce l'export Excel via openpyxl.
"""
import json, shutil, re, copy, os, sys
from pathlib import Path
from flask import Flask, request, jsonify, send_file, send_from_directory
from openpyxl import load_workbook
from openpyxl.styles import Alignment

app = Flask(__name__, static_folder='.', static_url_path='')

TEMPLATE = Path(__file__).parent / "TABELLA_RISCHIO_SECONDO_MODELLO_HERM__MO_3330_.xlsx"
EXPORT_DIR = Path(__file__).parent / "exports"
EXPORT_DIR.mkdir(exist_ok=True)

AREA_NAMES = {
    'A': 'Requisiti Generali', 'B': 'Personale', 'C': 'Acquisti',
    'D': 'Immobilizzazioni', 'E': 'Contabilità e Reporting Finanziario'
}

CAT_L1_MAP = {
    "Rischi clinico-sanitari": "Rischio Clinico Sanitario",
    "Rischi esterni": "Rischio Esterno", "Rischi finanziari": "Rischio Finanziario",
    "Rischi strategici": "Rischio Strategico", "Rischi operativi": "Rischio Operativo",
    "Rischi compliance": "Rischio di Compliance",
    "Rischio Finanziario": "Rischio Finanziario", "Rischio di Compliance": "Rischio di Compliance",
    "Rischio Operativo": "Rischio Operativo", "Rischio Strategico": "Rischio Strategico",
    "Rischio Esterno": "Rischio Esterno", "Rischio Clinico Sanitario": "Rischio Clinico Sanitario",
}


def normalizza_cat_l1(v):
    return CAT_L1_MAP.get(v, v or '')


def formatta_cause(c):
    if isinstance(c, dict):
        parti = []
        if c.get("processi"):
            parti.append(f"Processi:\n{c['processi']}")
        if c.get("persone"):
            parti.append(f"Persone:\n{c['persone']}")
        if c.get("fattori_esterni"):
            parti.append(f"Fattori esterni:\n{c['fattori_esterni']}")
        if c.get("strumenti"):
            parti.append(f"Strumenti:\n{c['strumenti']}")
        return "\n\n".join(parti) if parti else ""
    return str(c) if c else ""


def formatta_conseguenze(c):
    if isinstance(c, dict):
        parti = []
        if c.get("economiche"):
            parti.append(f"Economiche: {c['economiche']}")
        if c.get("reputazionali"):
            parti.append(f"Danno d'immagine: {c['reputazionali']}")
        if c.get("compliance"):
            parti.append(f"Compliance normativa: {c['compliance']}")
        if c.get("salute_sicurezza"):
            parti.append(f"Salute e sicurezza: {c['salute_sicurezza']}")
        else:
            parti.append("Salute e sicurezza: -")
        if c.get("gestionale_operativo"):
            parti.append(f"Gestionale - Operativo: {c['gestionale_operativo']}")
        if c.get("ambiente"):
            parti.append(f"Ambiente: {c['ambiente']}")
        return "\n\n".join(parti) if parti else ""
    return str(c) if c else ""


def clone_row(ws, template_row, new_row):
    offset = new_row - template_row
    for col in range(1, ws.max_column + 1):
        src = ws.cell(row=template_row, column=col)
        tgt = ws.cell(row=new_row, column=col)
        if src.has_style:
            tgt.font = copy.copy(src.font)
            tgt.fill = copy.copy(src.fill)
            tgt.border = copy.copy(src.border)
            tgt.alignment = copy.copy(src.alignment)
            tgt.number_format = src.number_format
        if src.value is None:
            tgt.value = None
        elif isinstance(src.value, str) and src.value.startswith("="):
            tgt.value = re.sub(
                r'(?<!\$)([A-Z]+)(\d+)',
                lambda m: f"{m.group(1)}{int(m.group(2)) + offset}",
                src.value
            )
        else:
            tgt.value = src.value
    rd = ws.row_dimensions.get(template_row)
    if rd and rd.height:
        ws.row_dimensions[new_row].height = rd.height


@app.route("/")
def index():
    return send_from_directory('.', 'index.html')


@app.route("/<path:path>")
def static_files(path):
    return send_from_directory('.', path)


@app.route("/api/health")
def health():
    return jsonify({"status": "ok", "message": "HERM Compiler API"})


@app.route("/api/export", methods=["POST", "OPTIONS"])
def export_excel():
    if request.method == "OPTIONS":
        return "", 204

    data = request.get_json()
    scenari = data.get("risks", [])
    pac_area = data.get("pacArea", "D")

    if not scenari:
        return jsonify({"error": "Nessun rischio da esportare"}), 400

    if not TEMPLATE.exists():
        return jsonify({"error": f"Template non trovato: {TEMPLATE}"}), 500

    try:
        area_title = f"PAC AREA {pac_area} {AREA_NAMES.get(pac_area, '').upper()}"
        filename = f"Risk_Register_PAC_Area_{pac_area}_{AREA_NAMES.get(pac_area, '').replace(' ', '_')}.xlsx"
        output_path = EXPORT_DIR / filename

        shutil.copy2(TEMPLATE, output_path)
        wb = load_workbook(output_path)

        if "Copertina" in wb.sheetnames:
            ws_cop = wb["Copertina"]
            for r in range(1, ws_cop.max_row + 1):
                for c in range(1, ws_cop.max_column + 1):
                    v = ws_cop.cell(row=r, column=c).value
                    if v and isinstance(v, str) and "RISK REGISTER" in v:
                        ws_cop.cell(row=r, column=c).value = f"RISK REGISTER - {area_title}"

        ws = wb["Registro dei Rischi"]
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=2, column=c).value
            if v and isinstance(v, str) and "RISK REGISTER" in v:
                ws.cell(row=2, column=c).value = f"RISK REGISTER - {area_title}"

        FIRST_DATA_ROW = 7
        n = len(scenari)
        righe_esistenti = ws.max_row - FIRST_DATA_ROW + 1
        if n > righe_esistenti:
            for i in range(righe_esistenti, n):
                clone_row(ws, FIRST_DATA_ROW, FIRST_DATA_ROW + i)

        for idx, s in enumerate(scenari):
            r = FIRST_DATA_ROW + idx
            ws.cell(row=r, column=3).value = idx + 1
            ws.cell(row=r, column=5).value = s.get("direzione", "")
            ws.cell(row=r, column=6).value = s.get("risk_owner", "")
            ws.cell(row=r, column=7).value = s.get("processo", "")
            ws.cell(row=r, column=8).value = normalizza_cat_l1(s.get("categoria_l1", ""))
            ws.cell(row=r, column=9).value = s.get("categoria_l2", "")
            ws.cell(row=r, column=10).value = s.get("categoria_l3", "N.A.") or "N.A."
            ws.cell(row=r, column=11).value = s.get("scenario", "")
            ws.cell(row=r, column=12).value = formatta_cause(s.get("cause", {}))
            ws.cell(row=r, column=13).value = formatta_conseguenze(s.get("conseguenze", {}))
            ws.cell(row=r, column=14).value = s.get("descrizione", "")
            ws.cell(row=r, column=15).value = s.get("impatto_economico")
            ws.cell(row=r, column=16).value = s.get("impatto_reputazionale")
            ws.cell(row=r, column=17).value = s.get("impatto_salute_sicurezza")
            ws.cell(row=r, column=18).value = s.get("impatto_compliance")
            ws.cell(row=r, column=19).value = s.get("impatto_gestionale_operativo")
            imp_amb = s.get("impatto_ambiente")
            if imp_amb is not None and imp_amb != "":
                ws.cell(row=r, column=20).value = imp_amb
            ws.cell(row=r, column=22).value = s.get("probabilita_inerente")
            ws.cell(row=r, column=24).value = s.get("controlli_correttivi", "")
            val_imp = s.get("valutazione_controlli_impatto", "")
            for col in [25, 26, 27, 28, 29, 30]:
                ws.cell(row=r, column=col).value = val_imp
            ws.cell(row=r, column=37).value = s.get("controlli_preventivi", "")
            ws.cell(row=r, column=38).value = s.get("valutazione_controlli_probabilita", "")
            ws.cell(row=r, column=49).value = s.get("azioni_mitigazione", "")
            conf = s.get("confidence_score", 0.5)
            eff = 2 if conf >= 0.7 else (1 if conf >= 0.4 else 0)
            ws.cell(row=r, column=50).value = eff
            ws.cell(row=r, column=52).value = eff
            for c in [11, 12, 13, 14, 24, 37, 49]:
                al = ws.cell(row=r, column=c).alignment
                ws.cell(row=r, column=c).alignment = Alignment(
                    wrap_text=True, vertical='top',
                    horizontal=al.horizontal if al else 'left'
                )

        wb.save(output_path)

        return send_file(
            output_path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get('PORT', 8080))
    print(f"HERM Compiler server: http://0.0.0.0:{port}")
    print(f"Template: {TEMPLATE}")
    print(f"Export dir: {EXPORT_DIR}")
    app.run(host='0.0.0.0', port=port, debug=False)
