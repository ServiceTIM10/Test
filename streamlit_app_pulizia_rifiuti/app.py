from __future__ import annotations

import io
import tempfile
from contextlib import redirect_stdout
from pathlib import Path

import streamlit as st

from pulisci_file_rifiuti import clean_excel_file


st.set_page_config(
    page_title="Pulizia file rifiuti nave",
    page_icon="🚢",
    layout="centered",
)

st.title("Pulizia file Excel - Rifiuti Nave 🚢")
st.markdown(
    """
Questa applicazione consente di caricare il file Excel relativo ai Rifiuti Nave, correggere automaticamente le colonne **Nr. Doc.**, **MC** e **Kg**, ordinare il database e scaricare il file pulito.

Il file viene elaborato temporaneamente durante la sessione e non viene salvato in modo permanente dall'applicazione.
"""
)

with st.expander("Cosa fa l'app", expanded=False):
    st.markdown(
        """
- Normalizza la colonna **Nr. Doc.** in numero intero.
- Corregge i casi troncati, ad esempio `5.64` → `5.640`.
- Converte **MC** e **Kg** in valori numerici e formato numero italiano
- Ordina l'intera tabella in ordine crescente sulla base di **Nr. Doc.**.
- Restituisce un nuovo file Excel pulito.
"""
    )

uploaded_file = st.file_uploader(
    "Carica il file Excel da pulire",
    type=["xlsx"],
    accept_multiple_files=False,
)

sheet_name = st.text_input(
    "Nome del foglio da pulire",
    value="",
    placeholder="Lascia vuoto per usare il primo foglio del file",
)

if uploaded_file is not None:
    st.info(f"File caricato: {uploaded_file.name}")

    if st.button("Pulisci file", type="primary"):
        with st.spinner("Pulizia del file in corso..."):
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    tmpdir_path = Path(tmpdir)
                    input_path = tmpdir_path / uploaded_file.name
                    output_path = tmpdir_path / f"pulito_{uploaded_file.name}"

                    input_path.write_bytes(uploaded_file.getbuffer())

                    log_buffer = io.StringIO()
                    with redirect_stdout(log_buffer):
                        clean_excel_file(
                            input_path=input_path,
                            output_path=output_path,
                            sheet_name=sheet_name.strip() or None,
                        )

                    output_bytes = output_path.read_bytes()
                    log_text = log_buffer.getvalue().strip()

                st.success("File pulito correttamente.")

                if log_text:
                    st.text_area("Report elaborazione", value=log_text, height=120)

                st.download_button(
                    label="Scarica file pulito",
                    data=output_bytes,
                    file_name=f"pulito_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            except Exception as exc:
                st.error("Si è verificato un errore durante la pulizia del file.")
                st.exception(exc)
else:
    st.warning("Carica un file Excel per iniziare.")
