import pandas as pd
import streamlit as st
import re
from urllib.parse import urlparse
import io
import os

def setup_matplotlib():
    try:
        import matplotlib.pyplot as plt
        import numpy as np
        return True, plt, np
    except ModuleNotFoundError:
        st.warning("Modulo 'matplotlib' mancante. Per visualizzare i radar chart, esegui: pip install matplotlib")
        return False, None, None

MATPLOTLIB_OK, plt, np = setup_matplotlib()

# === FUNZIONI SEO ===

# [funzioni SEO già presenti sopra...]

# === INTERFACCIA STREAMLIT ===

st.title("SEO Audit Tool")

# Tab 1 - Singolo file
with st.tabs(["Singolo File", "Multi File"])[0]:
    file = st.file_uploader("Carica un file .xlsx (Screaming Frog)", type="xlsx", key="single")
    if file:
        xls = pd.ExcelFile(file)
        sheet_name = None
        for name in ['1 - HTML', '1 - All']:
            if name in xls.sheet_names:
                sheet_name = name
                break
        if sheet_name:
            df = xls.parse(sheet_name)
            kpi = estrai_kpi(df)
            st.subheader("Riepilogo SEO")
            st.dataframe(kpi)
            # Tabella riepilogativa sintetica
            kpi_short = kpi[['SEO Score', 'Penalità Status Code %', 'Penalità Canonical %', 'Penalità Tag HTML %', 'Penalità Contenuti Duplicati %', 'Penalità CWV %']].copy()
            kpi_short['Stato'] = kpi_short['SEO Score'].apply(lambda x: 'Critico' if x < 50 else 'Medio' if x < 70 else 'Buono')
            st.subheader("Sintesi SEO")
            st.dataframe(kpi_short)
            if MATPLOTLIB_OK:
                kpi_riepilogo = kpi[['Penalità Status Code %', 'Penalità Canonical %', 'Penalità Tag HTML %', 'Penalità Contenuti Duplicati %', 'Penalità CWV %']]
                labels = kpi_riepilogo.columns.tolist()
                values = kpi_riepilogo.iloc[0].tolist()
                values = [v if isinstance(v, (int, float)) else 0 for v in values]
                angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
                values += values[:1]
                angles += angles[:1]

                fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))
                ax.plot(angles, values, 'o-', linewidth=2)
                ax.fill(angles, values, alpha=0.25)
                ax.set_yticklabels([])
                ax.set_xticks(angles[:-1])
                ax.set_xticklabels(labels)
                ax.set_title("Penalità SEO Radar Chart")
                st.pyplot(fig)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                kpi.to_excel(writer, sheet_name='SEO Report', index=False)
            st.download_button("📥 Scarica il report Excel", buffer.getvalue(), file_name="seo_report.xlsx")
        else:
            st.warning("Il file non contiene un foglio valido ('1 - HTML' o '1 - All').")

# Tab 2 - Multi File
with st.tabs(["Singolo File", "Multi File"])[1]:
    files = st.file_uploader("Carica più file .xlsx (Screaming Frog)", type="xlsx", accept_multiple_files=True, key="multi")
    if files:
        risultati = []
        for f in files:
            xls = pd.ExcelFile(f)
            sheet_name = None
            for name in ['1 - HTML', '1 - All']:
                if name in xls.sheet_names:
                    sheet_name = name
                    break
            if not sheet_name:
                continue
            df = xls.parse(sheet_name)
            dominio = estrai_dominio(df) or os.path.splitext(f.name)[0].split("_")[0]
            kpi = estrai_kpi(df)
            kpi.insert(0, 'Dominio', dominio)
            risultati.append(kpi)
        if risultati:
            df_riepilogo = pd.concat(risultati, ignore_index=True)
            st.subheader("Riepilogo SEO per tutti i domini")
            st.dataframe(df_riepilogo)
            # Tabella riepilogativa sintetica
            df_short = df_riepilogo[['Dominio', 'SEO Score', 'Penalità Status Code %', 'Penalità Canonical %', 'Penalità Tag HTML %', 'Penalità Contenuti Duplicati %', 'Penalità CWV %']].copy()
            df_short['Stato'] = df_short['SEO Score'].apply(lambda x: 'Critico' if x < 50 else 'Medio' if x < 70 else 'Buono')
            st.subheader("Sintesi SEO per dominio")
            st.dataframe(df_short)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                for dominio, group in df_riepilogo.groupby('Dominio'):
                    group.drop(columns=['Dominio']).to_excel(writer, sheet_name=dominio[:31], index=False)
                df_riepilogo.to_excel(writer, sheet_name='Riepilogo', index=False)
            st.download_button("📥 Scarica il report Excel", buffer.getvalue(), file_name="seo_audit_multi.xlsx")
        else:
            st.warning("Nessun file valido con foglio '1 - HTML' o '1 - All' trovato.")