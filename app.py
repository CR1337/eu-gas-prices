import streamlit as st
from eu_gas_prices import main

st.title("EU-Benzinpreise")

with st.spinner("Lade Benzinpreise herunter..."):
    zip_buffer = main()

st.success("Benzinpreise erfolgreich heruntergeladen!")

st.download_button(
    label="ZIP-Datei herunterladen",
    data=zip_buffer,
    file_name="benzinpreise.zip",
    mime="application/zip",
    width="stretch"
)