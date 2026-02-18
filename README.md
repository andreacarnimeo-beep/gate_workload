# Gate Workload – EasyMag (Streamlit)

Questa app mostra il carico di lavoro per **GATE di spedizione** in termini di:
- **Righe prelevate**
- **Colli creati**
con dettaglio per **Giro di consegna** e filtri per **periodo/date**.

## Modalità "Automatica" (schedulata)
L'integrazione automatica presuppone che EasyMag generi (o tu esporti) i file Excel in una cartella condivisa, ad esempio:

- `DATA_INBOX/`  (cartella dove arrivano gli export)
- `DATA_LAKE/`   (cartella dove l'app salva i dati normalizzati)

L'app legge SEMPRE i dati normalizzati (Parquet) in `DATA_LAKE/`.

### 1) Configurazione
Copia `.env.example` in `.env` e imposta le cartelle:

- `DATA_INBOX=/percorso/dove/arrivano/gli/export`
- `DATA_LAKE=/percorso/dove/salvare/i/parquet`
- `GATE_MAP=/percorso/GATE.xlsx`

### 2) Esecuzione ingestione (manuale)
```bash
python ingestion/run_ingestion.py
```

### 3) Schedulazione (Linux cron)
Esempio: ogni giorno alle 06:30
```cron
30 6 * * * /usr/bin/python3 /path/app/ingestion/run_ingestion.py >> /path/app/logs/ingestion.log 2>&1
```

### 4) Schedulazione (Windows Task Scheduler)
Crea un task che esegue:
```bat
C:\Python\python.exe C:\path\app\ingestion\run_ingestion.py
```

### 5) Docker (consigliato on-prem)
Usa `docker-compose.yml` per avviare:
- streamlit app
- container "scheduler" che esegue ingestion ogni N minuti (cron-like)

## Avvio app
```bash
streamlit run app.py
```
