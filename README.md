# Vendor Folder Miner

Applicativo Streamlit semplice per scandagliare una cartella locale piena di documenti e produrre un Excel unico con fornitori.

## File supportati

- Excel: .xlsx, .xlsm
- CSV
- TXT
- PDF
- Word: .docx

## Installazione

```bash
pip3 install -r requirements.txt
```

## Configurazione

Apri `config.py` e inserisci la tua chiave OpenAI.

```python
OPENAI_API_KEY = "sk-..."
```

## Avvio

```bash
streamlit run vendor_folder_miner.py
```

## Uso

1. Inserisci il percorso della cartella locale.
2. Premi "Avvia analisi cartella".
3. Scarica l'Excel finale.
