# Gate Workload • EasyMag (single-file)

App Streamlit single-file per analizzare carico di lavoro per Gate (righe prelevate e/o colli creati) a partire da export EasyMag in formato pivot.

## File richiesti
- **Obbligatorio:** `GATE.xlsx` (mapping Giro -> Gate; Giro in colonna B, Gate in colonna J)
- **Opzionale:** 1 o 2 export EasyMag (`.xlsx`):
  - puoi caricare solo **Righe** oppure solo **Colli** oppure entrambi
  - l'app tronca automaticamente la seconda tabella duplicata dopo la riga `(*) Numero di Operazioni`

## Run locale
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy Streamlit Cloud
- Main file path: `app.py`
