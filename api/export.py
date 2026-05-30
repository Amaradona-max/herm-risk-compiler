from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "message": "HERM Compiler API"})

@app.route("/api/export", methods=["POST", "OPTIONS"])
def export_excel():
    if request.method == "OPTIONS":
        return "", 204
    return jsonify({
        "error": "Export Excel richiede il server locale. Avvia: python3 server.py",
        "instructions": "Apri http://localhost:8080 per usare l'app con export Excel"
    }), 501
