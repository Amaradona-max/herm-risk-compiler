# HERM Risk Register Compiler

Web application per la compilazione intelligente del **Registro dei Rischi** secondo il modello **HERM** (Healthcare Enterprise Risk Management) di Regione Lombardia — **MO 3330**.

Sviluppata per **ASST della Brianza**.

## Architettura

```
Frontend:    index.html (singolo file, vanilla JS, zero framework)
AI proxy:    worker.js (Cloudflare Worker — proxy Anthropic API)
Deploy:      Cloudflare Pages (frontend) + Cloudflare Workers (backend/proxy)
Excel:       SheetJS (xlsx) via CDN — generazione .xlsx lato browser
Python:      herm_json_to_xlsx.py — script locale per compilare template originale MO 3330
```

## Quick Start

### 1. Deploy Worker (API proxy)

```bash
npm install -g wrangler
wrangler login
wrangler deploy
wrangler secret put ANTHROPIC_API_KEY
```

### 2. Deploy Frontend

```bash
wrangler pages deploy . --project-name herm-compiler
```

### 3. Configura

Apri `https://herm-compiler.pages.dev`, inserisci l'URL del Worker nel campo dedicato.

## Uso

1. **Carica** documenti (PDF, Excel, note PAC, metodologia HERM)
2. **Classifica** automaticamente i ruoli dei file
3. **Avvia** l'analisi AI per estrarre scenari di rischio strutturati
4. **Rivedi** e approva gli scenari nel registro
5. **Esporta** in Excel MO 3330, JSON, CSV o Markdown

## File

| File | Descrizione |
|------|-------------|
| `index.html` | Webapp completa (unico file frontend) |
| `worker.js` | Cloudflare Worker proxy API |
| `wrangler.toml` | Configurazione Wrangler/Cloudflare |
| `herm_json_to_xlsx.py` | Script Python locale (compilazione template MO 3330) |
| `TABELLA_RISCHIO_SECONDO_MODELLO_HERM__MO_3330_.xlsx` | Template originale MO 3330 |

## Script Python locale

```bash
pip install openpyxl
# Dopo aver esportato il JSON dalla webapp:
python3 herm_json_to_xlsx.py
# → genera HERM_MO3330_Compilato.xlsx
```
