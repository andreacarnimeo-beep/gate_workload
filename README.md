# Gate Workload • EasyMag (single-file v1.7)

Fix principali:
- Eliminati TUTTI i ternary con `st.metric(...) if ... else ...` (causavano stampa di testo/DeltaGenerator in alcune build).
- Normalizzazione Giro robusta (es: 99.0 -> 99).
- Esclusione colonne Tot/Tot: come "giro" (evita doppio conteggio e falsi missing).
- Se esistono giri senza Gate, l'app si blocca e ti obbliga ad assegnarli.

File:
- Obbligatorio: GATE.xlsx
- Opzionali: 1-2 export EasyMag (righe/colli)
