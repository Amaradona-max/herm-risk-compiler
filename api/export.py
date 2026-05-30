import json, os, base64, tempfile, re, copy
from flask import Flask, request, jsonify, send_file
from openpyxl import load_workbook
from openpyxl.styles import Alignment

app = Flask(__name__)

# Load template from separate file
from template_b64 import TEMPLATE_B64

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


def fc(c):
    if isinstance(c, dict):
        p = []
        if c.get("processi"): p.append("Processi:\n" + c["processi"])
        if c.get("persone"): p.append("Persone:\n" + c["persone"])
        if c.get("fattori_esterni"): p.append("Fattori esterni:\n" + c["fattori_esterni"])
        if c.get("strumenti"): p.append("Strumenti:\n" + c["strumenti"])
        return "\n\n".join(p)
    return str(c) if c else ""


def fco(c):
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


def cr(ws, tr, nr):
    o = nr - tr
    for col in range(1, ws.max_column + 1):
        s = ws.cell(row=tr, column=col)
        t = ws.cell(row=nr, column=col)
        if s.has_style:
            t.font = copy.copy(s.font)
            t.fill = copy.copy(s.fill)
            t.border = copy.copy(s.border)
            t.alignment = copy.copy(s.alignment)
            t.number_format = s.number_format
        if s.value is None:
            t.value = None
        elif isinstance(s.value, str) and s.value.startswith("="):
            t.value = re.sub(r'(?<!\$)([A-Z]+)(\d+)', lambda m: f"{m.group(1)}{int(m.group(2))+o}", s.value)
        else:
            t.value = s.value
    rd = ws.row_dimensions.get(tr)
    if rd and rd.height:
        ws.row_dimensions[nr].height = rd.height


@app.route("/api/health")
def health():
    return jsonify({"status": "ok"})


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

        template_data = base64.b64decode(TEMPLATE_B64)
        with open(tmp_path, "wb") as f:
            f.write(template_data)

        wb = load_workbook(tmp_path)
        area = f"PAC AREA {pac} {AREA_NAMES.get(pac, '').upper()}"

        if "Copertina" in wb.sheetnames:
            ws_c = wb["Copertina"]
            for r in range(1, ws_c.max_row + 1):
                for c in range(1, ws_c.max_column + 1):
                    v = ws_c.cell(row=r, column=c).value
                    if v and isinstance(v, str) and "RISK REGISTER" in v:
                        ws_c.cell(row=r, column=c).value = f"RISK REGISTER - {area}"

        ws = wb["Registro dei Rischi"]
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=2, column=c).value
            if v and isinstance(v, str) and "RISK REGISTER" in v:
                ws.cell(row=2, column=c).value = f"RISK REGISTER - {area}"

        F = 7
        n = len(scenari)
        existing = ws.max_row - F + 1
        if n > existing:
            for i in range(existing, n):
                cr(ws, F, F + i)

        for i, s in enumerate(scenari):
            r = F + i
            ws.cell(r, 3).value = i + 1
            ws.cell(r, 5).value = s.get("direzione", "")
            ws.cell(r, 6).value = s.get("risk_owner", "")
            ws.cell(r, 7).value = s.get("processo", "")
            ws.cell(r, 8).value = CAT_L1.get(s.get("categoria_l1", ""), s.get("categoria_l1", ""))
            ws.cell(r, 9).value = s.get("categoria_l2", "")
            ws.cell(r, 10).value = s.get("categoria_l3", "N.A.") or "N.A."
            ws.cell(r, 11).value = s.get("scenario", "")
            ws.cell(r, 12).value = fc(s.get("cause", {}))
            ws.cell(r, 13).value = fco(s.get("conseguenze", {}))
            ws.cell(r, 14).value = s.get("descrizione", "")
            ws.cell(r, 15).value = s.get("impatto_economico")
            ws.cell(r, 16).value = s.get("impatto_reputazionale")
            ws.cell(r, 17).value = s.get("impatto_salute_sicurezza")
            ws.cell(r, 18).value = s.get("impatto_compliance")
            ws.cell(r, 19).value = s.get("impatto_gestionale_operativo")
            ia = s.get("impatto_ambiente")
            if ia is not None and ia != "":
                ws.cell(r, 20).value = ia
            ws.cell(r, 22).value = s.get("probabilita_inerente")
            ws.cell(r, 24).value = s.get("controlli_correttivi", "")
            vi = s.get("valutazione_controlli_impatto", "")
            for col in [25, 26, 27, 28, 29, 30]:
                ws.cell(r, col).value = vi
            ws.cell(r, 37).value = s.get("controlli_preventivi", "")
            ws.cell(r, 38).value = s.get("valutazione_controlli_probabilita", "")
            ws.cell(r, 49).value = s.get("azioni_mitigazione", "")
            cf = s.get("confidence_score", 0.5)
            ef = 2 if cf >= 0.7 else (1 if cf >= 0.4 else 0)
            ws.cell(r, 50).value = ef
            ws.cell(r, 52).value = ef
            for c in [11, 12, 13, 14, 24, 37, 49]:
                al = ws.cell(r, c).alignment
                ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical='top', horizontal=al.horizontal if al else 'left')

        wb.save(tmp_path)
        fn = f"Risk_Register_PAC_Area_{pac}_{AREA_FILE.get(pac, 'Unknown')}.xlsx"
        return send_file(tmp_path, as_attachment=True, download_name=fn,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        return jsonify({"error": str(e)}), 500
