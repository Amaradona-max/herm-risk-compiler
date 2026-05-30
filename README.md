# HERM Risk Register Compiler v2.0

Web application per la compilazione intelligente del **Registro dei Rischi** secondo il modello **HERM** di Regione Lombardia — **MO 3330**.

---

## Accesso Online (per colleghi)

### URL pubblico
👉 **https://herm-compiler.vercel.app**

Aprire il link nel browser. Nessuna installazione necessaria.

### Cosa funziona online
- ✅ Upload documenti (PDF, Excel, CSV, DOCX, immagini)
- ✅ Analisi AI per estrarre scenari di rischio HERM
- ✅ Registro rischi interattivo con filtro e approvazione
- ✅ Dettaglio scenario con cause, conseguenze, metriche
- ✅ Export JSON, CSV, Markdown
- ⚠️ Export Excel MO 3330: richiede il server locale (vedi sotto)

### Per l'export Excel
L'export Excel con formule originali del template MO 3330 richiede il server Python locale (openpyxl). Se serve l'export Excel:

1. Apri il Terminale
2. Esegui:
   ```bash
   cd "/Users/prova/Desktop/Risck Herm2026/herm-compiler"
   pip install openpyxl flask
   python3 server.py
   ```
3. Apri **http://localhost:8080**
4. Usa l'app normalmente — l'export Excel funzionerà

---

## Avvio Server Locale (per sviluppo)

```bash
cd "/Users/prova/Desktop/Risck Herm2026/herm-compiler"
pip install openpyxl flask
python3 server.py
```
Poi apri: **http://localhost:8080**

---

## Utilizzo

1. **Seleziona Area PAC** (A-E) nella pagina Documenti
2. **Carica** documenti (PDF, Excel, note PAC, metodologia HERM)
3. **Avvia Analisi AI** — genera 6-8 scenari di rischio strutturati
4. **Rivedi e approva** gli scenari nel registro
5. **Esporta** in Excel MO 3330, JSON, CSV o Markdown

---

## Repository GitHub
https://github.com/Amaradona-max/herm-risk-compiler

---

*ASST della Brianza — HERM Risk Register Compiler v2.0*
