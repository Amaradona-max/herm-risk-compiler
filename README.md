# HERM Risk Register Compiler v2.0

Web application per la compilazione intelligente del **Registro dei Rischi** secondo il modello **HERM** (Healthcare Enterprise Risk Management) di Regione Lombardia — **MO 3330**.

Sviluppata per **ASST della Brianza**.

---

## Avvio Applicazione (per colleghi)

### Requisiti
- Python 3.9 o superiore
- Nessuna installazione npm/node necessaria

### Avvio rapido

1. Apri il **Terminale** (Mac) o **Prompt dei comandi** (Windows)
2. Naviga nella cartella del progetto:
   ```bash
   cd "/Users/prova/Desktop/Risck Herm2026/herm-compiler"
   ```
3. Avvia il server:
   ```bash
   python3 server.py
   ```
4. Apri il browser e vai a:
   ```
   http://localhost:8080
   ```

### Primo avvio (installazione dipendenze)

Se `python3 server.py` dà errore "No module named 'openpyxl'":

```bash
pip install openpyxl
python3 server.py
```

### Argo alternativo (porta diversa)

Se la porta 8080 è occupata:

```bash
python3 server.py 9090
```
Poi apri `http://localhost:9090`

---

## Utilizzo

### 1. Configurazione iniziale

- **Worker URL**: inserisci l'URL del Cloudflare Worker (già configurato)
- **Area PAC**: seleziona l'area del Percorso Attuativo della Certificabilità:
  - A — Requisiti Generali
  - B — Personale
  - C — Acquisti
  - D — Immobilizzazioni
  - E — Contabilità e Reporting Finanziario

### 2. Caricamento documenti

- Trascina i file nella zona di upload oppure clicca per selezionare
- Formati supportati: PDF, DOCX, Excel (XLSX/XLSM), CSV, TXT, MD, immagini
- I file vengono classificati automaticamente (puoi correggere il ruolo manualmente)

### 3. Analisi AI

- Clicca "Avvia Analisi AI" per generare gli scenari di rischio
- L'AI analizza i documenti e produce 6-8 scenari strutturati secondo il modello HERM
- Ogni scenario include: cause, conseguenze, impatti, probabilità, controlli, azioni

### 4. Registro rischi

- Visualizza tutti gli scenari nella tabella
- Filtra per categoria HERM
- Clicca su uno scenario per vedere il dettaglio completo
- Approva gli scenari pronti per l'esportazione

### 5. Export Excel

- Seleziona l'area PAC
- Clicca "Esporta Excel MO 3330"
- Il file viene generato con:
  - Tutte le formule originali del template (VLOOKUP, MAX, moltiplicazioni)
  - Formattazione e stili preservati
  - Titolo aggiornato con l'area PAC selezionata
  - Tutti i fogli del template (Copertina, Registro, Risk Guide, Risk Model, Metriche, Appoggio, Fonti)

---

## Formato file esportato

Il file Excel esportato (`Risk_Register_PAC_Area_[lettera]_[nome].xlsx`) contiene:

| Foglio | Contenuto |
|--------|-----------|
| Copertina | Titolo e indice |
| Registro dei Rischi | Dati compilati con 56 colonne e formule |
| Risk Guide | Guida ai campi |
| Risk Model | Tassonomia HERM |
| Metriche di valutazione | Scale di impatto, probabilità, controlli |
| Appoggio | Tabelle di lookup per VLOOKUP |
| Fonti e note | Documentazione di riferimento |

---

## Note tecniche

- **Frontend**: HTML5 + CSS3 + JavaScript vanilla (nessun framework)
- **Backend**: Python (server.py) con openpyxl per export Excel
- **AI**: Claude Sonnet via Cloudflare Worker proxy
- **Excel**: openpyxl preserva formule, stili e struttura del template originale

## File del progetto

| File | Descrizione |
|------|-------------|
| `index.html` | Webapp completa (frontend) |
| `server.py` | Server Python con export Excel via openpyxl |
| `worker.js` | Cloudflare Worker proxy API Anthropic |
| `wrangler.toml` | Configurazione Cloudflare |
| `Risk_Register_PAC_Area_D_Immobilizzazioni.xlsx` | Template di riferimento |
| `TABELLA_RISCHIO_SECONDO_MODELLO_HERM__MO_3330_.xlsx` | Template MO 3330 |
| `herm_json_to_xlsx.py` | Script Python alternativo per compilazione locale |

---

*ASST della Brianza — HERM Risk Register Compiler v2.0*
*QD_ENT_20265_3330 rev 0 del 20/05/2026*
