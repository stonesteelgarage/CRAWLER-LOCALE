import os
import re
import json
import subprocess
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
from openai import OpenAI

from openpyxl import load_workbook
import pdfplumber
from docx import Document


# ======================================================
# CONFIG IBRIDA: config.py locale + Streamlit secrets
# ======================================================

def get_secret(key, default=None):
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass

    try:
        import config
        if hasattr(config, key):
            return getattr(config, key)
    except Exception:
        pass

    return default


OPENAI_API_KEY = get_secret("OPENAI_API_KEY")
LOGO_PATH = get_secret("LOGO_PATH", "logo.png")
APP_PASSWORD = get_secret("APP_PASSWORD", "")
MODEL_NAME = get_secret("MODEL_NAME", "gpt-4.1-mini")


# ======================================================
# BASE APP
# ======================================================

APP_NAME = "Vendor Folder Miner"

st.set_page_config(
    page_title=APP_NAME,
    layout="wide",
    initial_sidebar_state="expanded"
)

os.makedirs("output", exist_ok=True)
os.makedirs("logs", exist_ok=True)

if not OPENAI_API_KEY:
    st.error("OPENAI_API_KEY non trovata. Inseriscila in config.py oppure nei secrets di Streamlit.")
    st.stop()

client = OpenAI(api_key=OPENAI_API_KEY)


# ======================================================
# CSS
# ======================================================

st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #071526 0%, #0b2238 45%, #08111f 100%);
    color: white;
}

[data-testid="stSidebar"] {
    background-color: rgba(6, 18, 33, 0.98);
}

h1, h2, h3, h4, h5, h6, p, label, span {
    color: white !important;
}

.stButton > button,
.stDownloadButton > button,
button[kind="primary"],
button[kind="secondary"] {
    background-color: #b30000 !important;
    color: white !important;
    border: 1px solid #ff4d4d !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background-color: #d00000 !important;
    color: white !important;
}

div[data-testid="stMetric"] {
    background-color: rgba(8, 33, 61, 0.85);
    border: 1px solid rgba(0, 210, 255, 0.35);
    border-radius: 14px;
    padding: 18px;
}

.card {
    background: rgba(8, 33, 61, 0.82);
    border: 1px solid rgba(0, 210, 255, 0.25);
    border-radius: 18px;
    padding: 22px;
    margin-bottom: 18px;
}

input[type="password"],
input[type="text"],
input[type="number"],
textarea {
    background-color: white !important;
    color: black !important;
}

div[data-testid="stAlert"] p {
    color: black !important;
}
</style>
""", unsafe_allow_html=True)


# ======================================================
# LOGIN OPZIONALE
# ======================================================

if APP_PASSWORD:
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    if not st.session_state["logged_in"]:
        col1, col2, col3 = st.columns([1, 1.2, 1])
        with col2:
            if os.path.exists(LOGO_PATH):
                st.image(LOGO_PATH, width=160)

            st.markdown(f"""
            <div class="card">
                <h2>{APP_NAME}</h2>
                <p>Accesso riservato</p>
            </div>
            """, unsafe_allow_html=True)

            password = st.text_input("Password", type="password")

            if st.button("Accedi", use_container_width=True):
                if password == APP_PASSWORD:
                    st.session_state["logged_in"] = True
                    st.rerun()
                else:
                    st.error("Password errata")

        st.stop()


# ======================================================
# UTILITY PERCORSI LOCALI / ONEDRIVE / RETE
# ======================================================

def normalize_folder_path(raw_path):
    """
    Corregge i problemi tipici dei percorsi copiati da Windows:
    - virgolette iniziali/finali
    - spazi finali
    - percorsi OneDrive
    - percorsi Desktop
    - percorsi UNC di rete tipo \\server\share\cartella
    """
    if raw_path is None:
        raw_path = ""

    cleaned = str(raw_path).strip()
    cleaned = cleaned.strip('"').strip("'").strip()

    return Path(cleaned).expanduser()


def select_folder_with_tkinter():
    """
    Apre il selettore cartella nativo.
    Funziona quando Streamlit gira localmente sul PC.
    Se l'app è su Streamlit Cloud, non può vedere il disco locale dell'utente.
    """
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", 1)

        folder = filedialog.askdirectory(title="Seleziona cartella da analizzare")

        root.destroy()

        return folder or ""
    except Exception:
        return ""


def select_folder_with_powershell():
    """
    Fallback specifico Windows se tkinter non apre la finestra.
    """
    if os.name != "nt":
        return ""

    try:
        ps_script = '''
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
$dialog.Description = "Seleziona cartella da analizzare"
$dialog.ShowNewFolderButton = $false
$result = $dialog.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    Write-Output $dialog.SelectedPath
}
'''
        completed = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=60,
        )

        if completed.returncode == 0:
            return completed.stdout.strip()

    except Exception:
        pass

    return ""


def select_folder_dialog():
    """
    Prima prova tkinter, poi PowerShell.
    """
    folder = select_folder_with_tkinter()
    if folder:
        return folder

    folder = select_folder_with_powershell()
    if folder:
        return folder

    return ""


def diagnose_missing_folder(raw_path):
    """
    Produce suggerimenti quando la cartella non viene trovata.
    """
    folder = normalize_folder_path(raw_path)

    suggestions = []
    folder_name = folder.name

    if folder_name:
        home = Path.home()

        possible_paths = [
            home / "Desktop" / folder_name,
            home / "OneDrive - Ghella SpA" / "Desktop" / folder_name,
            home / "OneDrive - Ghella SpA" / folder_name,
            home / "OneDrive" / "Desktop" / folder_name,
            home / "OneDrive" / folder_name,
        ]

        for p in possible_paths:
            try:
                if p.exists() and p.is_dir():
                    suggestions.append(str(p))
            except Exception:
                pass

    return {
        "path": str(folder),
        "exists": folder.exists(),
        "is_dir": folder.is_dir() if folder.exists() else False,
        "parent": str(folder.parent),
        "parent_exists": folder.parent.exists(),
        "suggestions": suggestions,
    }


# ======================================================
# UTILITY TESTO
# ======================================================

SUPPORTED_EXTENSIONS = {
    ".xlsx", ".xlsm", ".csv", ".txt", ".pdf", ".docx"
}


def clean_text(value):
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    value = str(value)
    value = value.replace("\n", " ")
    value = value.replace("\t", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def chunk_text(text, max_chars=9000):
    text = clean_text(text)
    chunks = []

    while len(text) > max_chars:
        split_at = text.rfind(" ", 0, max_chars)
        if split_at == -1:
            split_at = max_chars
        chunks.append(text[:split_at])
        text = text[split_at:].strip()

    if text:
        chunks.append(text)

    return chunks


def normalize_supplier_name(name):
    name = clean_text(name).lower()
    name = name.replace(".", "")
    name = name.replace(",", "")
    name = name.replace(" srl", "")
    name = name.replace(" s r l", "")
    name = name.replace(" spa", "")
    name = name.replace(" s p a", "")
    name = name.replace(" società", "")
    name = name.replace(" societa", "")
    name = re.sub(r"\s+", " ", name)
    return name.strip()


def find_files(folder_path):
    folder = normalize_folder_path(folder_path)

    if not folder.exists():
        raise FileNotFoundError(f"Cartella non trovata: {folder}")

    if not folder.is_dir():
        raise NotADirectoryError(f"Il percorso esiste ma non è una cartella: {folder}")

    files = []
    for root, dirs, filenames in os.walk(folder):
        for filename in filenames:
            path = Path(root) / filename
            if path.suffix.lower() in SUPPORTED_EXTENSIONS:
                if not filename.startswith("~$"):
                    files.append(path)

    return files


# ======================================================
# LETTURA DOCUMENTI
# ======================================================

def read_excel_file(path):
    texts = []

    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            rows_text = []
            max_row = min(ws.max_row or 1, 2500)
            max_col = min(ws.max_column or 1, 80)

            for row in ws.iter_rows(
                min_row=1,
                max_row=max_row,
                max_col=max_col,
                values_only=True
            ):
                cells = [clean_text(v) for v in row if clean_text(v)]
                if cells:
                    rows_text.append(" | ".join(cells))

            if rows_text:
                texts.append(f"FOGLIO: {ws.title}\n" + "\n".join(rows_text))
    except Exception as e:
        texts.append(f"ERRORE LETTURA EXCEL: {e}")

    return "\n\n".join(texts)


def read_csv_file(path):
    try:
        df = pd.read_csv(path, dtype=str, encoding="utf-8", sep=None, engine="python").fillna("")
    except Exception:
        try:
            df = pd.read_csv(path, dtype=str, encoding="latin-1", sep=None, engine="python").fillna("")
        except Exception as e:
            return f"ERRORE LETTURA CSV: {e}"

    return df.to_csv(index=False, sep=" | ")


def read_txt_file(path):
    for enc in ["utf-8", "latin-1", "cp1252"]:
        try:
            return Path(path).read_text(encoding=enc, errors="ignore")
        except Exception:
            continue
    return ""


def read_pdf_file(path):
    texts = []

    try:
        with pdfplumber.open(path) as pdf:
            for i, page in enumerate(pdf.pages[:80], start=1):
                txt = page.extract_text() or ""
                if txt.strip():
                    texts.append(f"PAGINA {i}\n{txt}")
    except Exception as e:
        texts.append(f"ERRORE LETTURA PDF: {e}")

    return "\n\n".join(texts)


def read_docx_file(path):
    try:
        doc = Document(path)
        parts = []

        for p in doc.paragraphs:
            txt = clean_text(p.text)
            if txt:
                parts.append(txt)

        for table in doc.tables:
            for row in table.rows:
                cells = [clean_text(cell.text) for cell in row.cells if clean_text(cell.text)]
                if cells:
                    parts.append(" | ".join(cells))

        return "\n".join(parts)
    except Exception as e:
        return f"ERRORE LETTURA DOCX: {e}"


def read_any_file(path):
    suffix = path.suffix.lower()

    if suffix in [".xlsx", ".xlsm"]:
        return read_excel_file(path)

    if suffix == ".csv":
        return read_csv_file(path)

    if suffix == ".txt":
        return read_txt_file(path)

    if suffix == ".pdf":
        return read_pdf_file(path)

    if suffix == ".docx":
        return read_docx_file(path)

    return ""


# ======================================================
# OPENAI EXTRACTION
# ======================================================

def extract_suppliers_with_openai(text, source_file):
    chunks = chunk_text(text, max_chars=9000)
    all_suppliers = []

    for idx, chunk in enumerate(chunks[:8], start=1):
        prompt = f"""
Sei un procurement analyst.

Analizza il testo seguente, proveniente da un documento aziendale/procurement.
Devi estrarre tutti i fornitori, imprese, subappaltatori, aziende, produttori, distributori o società citate come soggetti economici.

Per ogni fornitore estrai, quando disponibile:
- supplier_name: nome fornitore / ragione sociale
- vat_number: partita IVA o codice fiscale aziendale
- email
- phone
- address
- website
- offering: cosa offre, cioè servizio/prodotto/materiale/lavorazione collegata
- category: categoria sintetica, es. impianti, carpenteria, cemento, trasporti, consulenza, noleggio, sicurezza, forniture, subappalto, ecc.
- confidence: da 0 a 100

Regole:
- Non inventare dati.
- Se non trovi un campo, lascia stringa vuota.
- Se un nome sembra persona fisica senza azienda, ignoralo.
- Se una società compare senza dettagli ma è chiaramente un fornitore, includila.
- Restituisci SOLO JSON valido.
- Il JSON deve avere questa forma:
{{
  "suppliers": [
    {{
      "supplier_name": "",
      "vat_number": "",
      "email": "",
      "phone": "",
      "address": "",
      "website": "",
      "offering": "",
      "category": "",
      "confidence": 0
    }}
  ]
}}

File origine: {source_file}
Chunk: {idx}/{len(chunks)}

TESTO:
{chunk}
"""

        try:
            response = client.responses.create(
                model=MODEL_NAME,
                input=prompt
            )

            raw = response.output_text.strip()
            raw = raw.replace("```json", "").replace("```", "").strip()

            try:
                data = json.loads(raw)
            except Exception:
                match = re.search(r"\{.*\}", raw, flags=re.DOTALL)
                if match:
                    data = json.loads(match.group(0))
                else:
                    data = {"suppliers": []}

            suppliers = data.get("suppliers", [])
            if isinstance(suppliers, list):
                for s in suppliers:
                    if clean_text(s.get("supplier_name", "")):
                        s["source_file"] = source_file
                        all_suppliers.append(s)

        except Exception as e:
            log_path = Path("logs") / "openai_errors.log"
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"\n--- {datetime.now()} | {source_file} | chunk {idx} ---\n{e}\n")

    return all_suppliers


# ======================================================
# DEDUPLICA
# ======================================================

def deduplicate_suppliers(rows):
    merged = {}

    for row in rows:
        name = clean_text(row.get("supplier_name", ""))
        vat = clean_text(row.get("vat_number", ""))
        email = clean_text(row.get("email", ""))
        website = clean_text(row.get("website", ""))

        if vat:
            key = f"vat::{vat.lower()}"
        elif email:
            key = f"email::{email.lower()}"
        elif website:
            key = f"web::{website.lower().replace('https://', '').replace('http://', '').replace('www.', '').strip('/')}"
        else:
            key = f"name::{normalize_supplier_name(name)}"

        if key not in merged:
            merged[key] = {
                "Fornitore": name,
                "Partita IVA": vat,
                "Email": email,
                "Telefono": clean_text(row.get("phone", "")),
                "Indirizzo": clean_text(row.get("address", "")),
                "Sito web": website,
                "Cosa offre": clean_text(row.get("offering", "")),
                "Categoria": clean_text(row.get("category", "")),
                "Confidenza AI": row.get("confidence", ""),
                "File origine": clean_text(row.get("source_file", "")),
            }
        else:
            existing = merged[key]

            for target_col, source_key in [
                ("Partita IVA", "vat_number"),
                ("Email", "email"),
                ("Telefono", "phone"),
                ("Indirizzo", "address"),
                ("Sito web", "website"),
                ("Cosa offre", "offering"),
                ("Categoria", "category"),
            ]:
                current = clean_text(existing.get(target_col, ""))
                new = clean_text(row.get(source_key, ""))

                if new and new not in current:
                    existing[target_col] = f"{current}; {new}".strip("; ")

            source = clean_text(row.get("source_file", ""))
            if source and source not in existing["File origine"]:
                existing["File origine"] += f"; {source}"

            try:
                existing_conf = int(existing.get("Confidenza AI") or 0)
                new_conf = int(row.get("confidence") or 0)
                existing["Confidenza AI"] = max(existing_conf, new_conf)
            except Exception:
                pass

    return list(merged.values())


# ======================================================
# ELABORAZIONE CARTELLA
# ======================================================

def process_folder(folder_path, max_files=None):
    files = find_files(folder_path)

    if max_files:
        files = files[:max_files]

    all_rows = []
    report_rows = []

    progress = st.progress(0)
    status = st.empty()

    total = len(files)

    if total == 0:
        return files, [], [], None

    for i, path in enumerate(files, start=1):
        status.info(f"ALMOND Intelligence sta lavorando... File {i}/{total}: {path.name}")

        try:
            text = read_any_file(path)
            chars = len(text)

            if chars < 30:
                report_rows.append({
                    "File": str(path),
                    "Esito": "Saltato",
                    "Dettaglio": "Testo insufficiente"
                })
            else:
                suppliers = extract_suppliers_with_openai(text, str(path))
                all_rows.extend(suppliers)

                report_rows.append({
                    "File": str(path),
                    "Esito": "Analizzato",
                    "Dettaglio": f"Fornitori estratti: {len(suppliers)}"
                })

        except Exception as e:
            report_rows.append({
                "File": str(path),
                "Esito": "Errore",
                "Dettaglio": str(e)
            })

        progress.progress(i / total if total else 1)

    status.success("Analisi cartella terminata")

    deduped = deduplicate_suppliers(all_rows)

    output_path = Path("output") / f"fornitori_estratti_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    df_suppliers = pd.DataFrame(deduped)
    df_report = pd.DataFrame(report_rows)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_suppliers.to_excel(writer, sheet_name="Fornitori", index=False)
        df_report.to_excel(writer, sheet_name="Log analisi", index=False)

        try:
            ws = writer.book["Fornitori"]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            widths = {
                "A": 38,
                "B": 18,
                "C": 32,
                "D": 22,
                "E": 40,
                "F": 32,
                "G": 60,
                "H": 30,
                "I": 18,
                "J": 80,
            }
            for col, width in widths.items():
                ws.column_dimensions[col].width = width
        except Exception:
            pass

    return files, deduped, report_rows, output_path


# ======================================================
# SESSION STATE
# ======================================================

if "selected_folder_path" not in st.session_state:
    st.session_state["selected_folder_path"] = ""

if "last_output_path" not in st.session_state:
    st.session_state["last_output_path"] = None


# ======================================================
# SIDEBAR
# ======================================================

with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=150)

    st.markdown(f"### {APP_NAME}")
    st.caption("Estrazione fornitori da cartelle locali, Desktop, OneDrive e percorsi di rete")

    st.divider()

    st.markdown("**File supportati**")
    st.write(".xlsx, .xlsm, .csv, .txt, .pdf, .docx")

    st.divider()

    st.markdown("**Configurazione**")
    st.write(f"Modello: `{MODEL_NAME}`")

    st.divider()

    st.markdown("**Percorsi validi**")
    st.code(r"C:\Users\samandorla\Desktop\DB COLMETO")
    st.code(r"C:\Users\samandorla\OneDrive - Ghella SpA\Desktop\DB COLMETO")
    st.code(r"\\server\share\cartella")


# ======================================================
# UI PRINCIPALE
# ======================================================

st.title(APP_NAME)

st.markdown("""
<div class="card">
<h3>Estrazione intelligente fornitori da cartella locale</h3>
<p>
Inserisci o seleziona il percorso di una cartella del PC. L'app scandaglia sottocartelle e documenti,
legge Excel, PDF, Word, CSV e TXT, usa OpenAI per individuare fornitori, contatti e servizi,
e produce un unico file Excel finale.
</p>
</div>
""", unsafe_allow_html=True)

col_path, col_button = st.columns([3, 1])

with col_path:
    folder_path = st.text_input(
        "Percorso cartella locale o percorso di rete",
        value=st.session_state["selected_folder_path"],
        placeholder=r"C:\Users\samandorla\OneDrive - Ghella SpA\Desktop\DB COLMETO oppure \\server\share\cartella",
        help="Puoi incollare un percorso Windows, OneDrive, Desktop o rete UNC. Esempio rete: \\\\server\\share\\cartella"
    )

with col_button:
    st.write("")
    st.write("")
    if st.button("📁 Seleziona cartella", use_container_width=True):
        selected = select_folder_dialog()

        if selected:
            st.session_state["selected_folder_path"] = selected
            st.rerun()
        else:
            st.warning(
                "Selettore cartella non disponibile. Se stai usando Streamlit Cloud è normale: "
                "devi eseguire l'app in locale oppure incollare manualmente il percorso."
            )

if folder_path:
    st.session_state["selected_folder_path"] = folder_path

if folder_path:
    with st.expander("Verifica percorso"):
        info = diagnose_missing_folder(folder_path)

        st.write(f"Percorso letto dall'app: `{info['path']}`")
        st.write(f"Cartella esistente: `{info['exists']}`")
        st.write(f"È una cartella: `{info['is_dir']}`")
        st.write(f"Cartella superiore: `{info['parent']}`")
        st.write(f"Cartella superiore esistente: `{info['parent_exists']}`")

        if not info["exists"]:
            st.warning(
                "La cartella non risulta raggiungibile da Python. "
                "Apri la cartella in Esplora file, clicca sulla barra del percorso in alto, "
                "copia il percorso completo e incollalo qui."
            )

            if info["suggestions"]:
                st.info("Ho trovato questi percorsi alternativi esistenti. Copiane uno nel campo sopra:")
                for s in info["suggestions"]:
                    st.code(s)

col_a, col_b = st.columns(2)

with col_a:
    limit_files = st.checkbox("Limita numero file per test", value=True)

with col_b:
    max_files = st.number_input("Numero massimo file", min_value=1, max_value=1000, value=20, step=1)

effective_max_files = int(max_files) if limit_files else None

if st.button("Avvia analisi cartella", use_container_width=True):
    if not folder_path:
        st.warning("Inserisci o seleziona il percorso della cartella.")
    else:
        try:
            folder = normalize_folder_path(folder_path)

            if not folder.exists():
                info = diagnose_missing_folder(folder_path)
                st.error(f"Cartella non trovata: {folder}")

                if info["parent_exists"]:
                    st.info(
                        "La cartella superiore esiste, ma non esiste la cartella finale. "
                        "Controlla il nome esatto della cartella."
                    )
                else:
                    st.info(
                        "Nemmeno la cartella superiore risulta raggiungibile. "
                        "Probabile percorso OneDrive/Desktop diverso, cartella non sincronizzata, "
                        "oppure percorso di rete non montato."
                    )

                if info["suggestions"]:
                    st.info("Percorsi alternativi trovati:")
                    for s in info["suggestions"]:
                        st.code(s)

                st.stop()

            if not folder.is_dir():
                st.error(f"Il percorso esiste ma non è una cartella: {folder}")
                st.stop()

            files, suppliers, report_rows, output_path = process_folder(folder, effective_max_files)
            st.session_state["last_output_path"] = str(output_path) if output_path else None

            st.success("Analisi completata")
            st.info(f"File analizzati: {len(files)}")
            st.info(f"Fornitori univoci trovati: {len(suppliers)}")

            if suppliers and output_path:
                st.subheader("Anteprima fornitori estratti")
                st.dataframe(pd.DataFrame(suppliers), use_container_width=True)

                with open(output_path, "rb") as f:
                    st.download_button(
                        "Scarica Excel fornitori",
                        data=f,
                        file_name=output_path.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            else:
                st.warning("Nessun fornitore trovato nei documenti analizzati.")

            with st.expander("Log analisi file"):
                st.dataframe(pd.DataFrame(report_rows), use_container_width=True)

        except Exception as e:
            st.error(f"Errore: {e}")
            st.info(
                "Controlla che il percorso sia accessibile dal PC dove gira Streamlit. "
                "Per percorsi di rete usa formato UNC, ad esempio: \\\\server\\share\\cartella"
            )
