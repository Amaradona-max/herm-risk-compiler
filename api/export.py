from flask import Flask, request, jsonify, send_file
import json, shutil, re, copy, os
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import tempfile

app = Flask(__name__)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), '..', 'TABELLA_RISCHIO_SECONDO_MODELLO_HERM__MO_3330_.xlsx')

AREA_NAMES = {
    'A': 'Requisiti Generali', 'B': 'Personale', 'C': 'Acquisti',
    'D': 'Immobilizzazioni', 'E': 'Contabilita e Reporting Finanziario'
}
AREA_FILE = {
    'A': 'Requisiti_Generali', 'B': 'Personale', 'C': 'Acquisti',
    'D': 'Immobilizzazioni', 'E': 'Contabilita_e_Reporting'
}
CAT_L1 = {
    "Rischi clinico-sanitari": "Rischio Clinico Sanitario",
    "Rischi esterni": "Rischio Esterno", "Rischi finanziari": "Rischio Finanziario",
    "Rischi strategici": "Rischio Strategico", "Rischi operativi": "Rischio Operativo",
    "Rischi compliance": "Rischio di Compliance",
    "Rischio Finanziario": "Rischio Finanziario", "Rischio di Compliance": "Rischio di Compliance",
    "Rischio Operativo": "Rischio Operativo", "Rischio Strategico": "Rischio Strategico",
    "Rischio Esterno": "Rischio Esterno", "Rischio Clinico Sanitario": "Rischio Clinico Sanitario",
}


def format_cause(c):
    if isinstance(c, dict):
        p = []
        if c.get("processi"): p.append("Processi:\n" + c["processi"])
        if c.get("persone"): p.append("Persone:\n" + c["persone"])
        if c.get("fattori_esterni"): p.append("Fattori esterni:\n" + c["fattori_esterni"])
        if c.get("strumenti"): p.append("Strumenti:\n" + c["strumenti"])
        return "\n\n".join(p)
    return str(c) if c else ""


def format_consequence(c):
    if isinstance(c, dict):
        p = []
        if c.get("economiche"): p.append("Economiche: " + c["economiche"])
        if c.get("reputazionali"): p.append("Danno d'immagine: " + c["reputazionali"])
        if c.get("compliance"): p.append("Compliance normativa: " + c["compliance"])
        if c.get("salute_sicurezza"): p.append("Salute e sicurezza: " + c["salute_sicurezza"])
        else: p.append("Salute e sicurezza: -")
        if c.get("gestionale_operativo"): p.append("Gestionale - Operativo: " + c["gestionale_operativo"])
        if c.get("ambiente"): p.append("Ambiente: " + c["ambiente"])
        return "\n\n".join(p)
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


@app.route("/api/export", methods=["POST", "OPTIONS"])
def export_excel():
    if request.method == "OPTIONS":
        return "", 204

    data = request.get_json()
    scenari = data.get("risks", [])
    pac = data.get("pacArea", "D")

    if not scenari:
        return jsonify({"error": "Nessun rischio"}), 400

    try:
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp_path = tmp.name
        tmp.close()

        shutil.copy2(TEMPLATE_PATH, tmp_path)
        wb = load_workbook(tmp_path)

        area_title = f"PAC AREA {pac} {AREA_NAMES.get(pac, '').upper()}"

        # Update Copertina title
        if "Copertina" in wb.sheetnames:
            ws_cop = wb["Copertina"]
            for r in range(1, ws_cop.max_row + 1):
                for c in range(1, ws_cop.max_column + 1):
                    v = ws_cop.cell(row=r, column=c).value
                    if v and isinstance(v, str) and "RISK REGISTER" in v:
                        ws_cop.cell(row=r, column=c).value = f"RISK REGISTER - {area_title}"

        # Update Registro dei Rischi title
        ws = wb["Registro dei Rischi"]
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=2, column=c).value
            if v and isinstance(v, str) and "RISK REGISTER" in v:
                ws.cell(row=2, column=c).value = f"RISK REGISTER - {area_title}"

        FIRST_DATA_ROW = 7
        n = len(scenari)
        existing_rows = ws.max_row - FIRST_DATA_ROW + 1
        if n > existing_rows:
            for i in range(existing_rows, n):
                clone_row(ws, FIRST_DATA_ROW, FIRST_DATA_ROW + i)

        for idx, s in enumerate(scenari):
            r = FIRST_DATA_ROW + idx
            ws.cell(row=r, column=3).value = idx + 1
            ws.cell(row=r, column=5).value = s.get("direzione", "")
            ws.cell(row=r, column=6).value = s.get("risk_owner", "")
            ws.cell(row=r, column=7).value = s.get("processo", "")
            ws.cell(row=r, column=8).value = CAT_L1.get(s.get("categoria_l1", ""), s.get("categoria_l1", ""))
            ws.cell(row=r, column=9).value = s.get("categoria_l2", "")
            ws.cell(row=r, column=10).value = s.get("categoria_l3", "N.A.") or "N.A."
            ws.cell(row=r, column=11).value = s.get("scenario", "")
            ws.cell(row=r, column=12).value = format_cause(s.get("cause", {}))
            ws.cell(row=r, column=13).value = format_consequence(s.get("conseguenze", {}))
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
            # Set wrap text on long text columns
            for col in [11, 12, 13, 14, 24, 37, 49]:
                cell = ws.cell(row=r, column=col)
                cell.alignment = Alignment(wrap_text=True, vertical='top')

        wb.save(tmp_path)

        filename = f"Risk_Register_PAC_Area_{pac}_{AREA_FILE.get(pac, 'Unknown')}.xlsx"
        return send_file(
            tmp_path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/health")
def health():
    return jsonify({"status": "ok"})
