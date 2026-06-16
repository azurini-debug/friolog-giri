# Friolog — Gestione Sequenze Consegna

Applicazione web per la gestione delle sequenze di scarico dei giri di consegna.

## File del progetto

- `app.py` — Applicazione principale Flask
- `genera_password.py` — Script per generare l'hash della password admin
- `requirements.txt` — Dipendenze Python
- `render.yaml` — Configurazione per deploy su Render.com
- `templates/` — Pagine HTML

## Deploy su Render.com (passo per passo)

Vedi la guida dettagliata fornita separatamente.

## Variabili d'ambiente richieste

| Variabile | Descrizione |
|---|---|
| `SECRET_KEY` | Chiave segreta Flask (generata automaticamente da Render) |
| `ADMIN_PASSWORD_HASH` | Hash della password admin (generato con `genera_password.py`) |

## Formato file XLS atteso

File SpreadsheetML (.xls) esportato da AS400 con colonne:
- A: Sequenza originale
- B: Ragione sociale cliente
- C: Indirizzo
- D: Città
- E: Provincia
- G: Codice cliente (chiave per l'export)
- H: Telefono cliente
- I: Numero giro
- J: Note di consegna

## Formato CSV di output (per AS400)

```
GIRO;SEQ_AUTISTA;CODICE_CLIENTE;RAGIONE_SOCIALE
23;1;000687;DUE STELLE DI FICARELLI GILDA SRL
23;2;PA6597;BELAEA OLGA
...
```
