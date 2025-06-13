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

# Funzioni SEO già definite (estrai_dominio, calcola_score, estrai_kpi)...

st.title("SEO Audit Tool")

tab1, tab2 = st.tabs(["Singolo File", "Multi File"])

with tab1:
    file = st.file_uploader("Carica un file .xlsx (Screaming Frog)", type="xlsx", key="singolo")

    if file:
        xls = pd.ExcelFile(file)
        sheet_name = None
        for name in ['1 - HTML', '1 - All']:
            if name in xls.sheet_names:
                sheet_name = name
                break

        if not sheet_name:
            st.error("Foglio '1 - HTML' o '1 - All' non trovato nel file.")
        else:
            df = xls.parse(sheet_name)
            kpi_df = estrai_kpi(df)

            st.subheader("📊 Tabella KPI SEO")
            st.dataframe(kpi_df)

            st.subheader("🎯 SEO Score e Penalità")
            kpi_riepilogo = kpi_df[[
                'SEO Score',
                'Penalità Status Code %',
                'Penalità Canonical %',
                'Penalità Tag HTML %',
                'Penalità Contenuti Duplicati %',
                'Penalità CWV %'
            ]]

            st.dataframe(
                kpi_riepilogo.style.apply(
                    lambda row: [
                        'background-color: #f8d7da' if row['SEO Score'] <= 49 else
                        'background-color: #fff3cd' if row['SEO Score'] <= 69 else
                        'background-color: #d4edda' for _ in row
                    ],
                    axis=1
                )
            )

            if MATPLOTLIB_OK:
                labels = ['Status Code', 'Canonical', 'Tag HTML', 'Contenuti Duplicati', 'CWV']
                values = [
                    kpi_riepilogo['Penalità Status Code %'].iloc[0],
                    kpi_riepilogo['Penalità Canonical %'].iloc[0],
                    kpi_riepilogo['Penalità Tag HTML %'].iloc[0],
                    kpi_riepilogo['Penalità Contenuti Duplicati %'].iloc[0],
                    kpi_riepilogo['Penalità CWV %'].iloc[0]
                ]
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
                ax.set_title("Radar Chart delle Penalità SEO")
                st.pyplot(fig)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                kpi_df.to_excel(writer, sheet_name='Report SEO', index=False)
                kpi_riepilogo.to_excel(writer, sheet_name='Riepilogo Score', index=False)

            st.download_button("📥 Scarica il report Excel", buffer.getvalue(), file_name="seo_report_singolo.xlsx")

with tab2:
    files = st.file_uploader("Carica più file .xlsx (Screaming Frog)", type="xlsx", accept_multiple_files=True, key="multi")

    if files:
        output = {}
        report_completo = []

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
            dominio = urlparse(df['Address'].dropna().iloc[0]).netloc.replace("www.", "") if 'Address' in df.columns else f.name.split("_")[0]
            kpi_df = estrai_kpi(df)
            kpi_df.insert(0, 'Dominio', dominio)
            report_completo.append(kpi_df)
            output[dominio] = kpi_df.drop(columns=['Dominio'])

        if output:
            df_totale = pd.concat(report_completo, ignore_index=True)
            df_riepilogo = df_totale[[
                'Dominio',
                'SEO Score',
                'Penalità Status Code %',
                'Penalità Canonical %',
                'Penalità Tag HTML %',
                'Penalità Contenuti Duplicati %',
                'Penalità CWV %'
            ]]
            st.subheader("📈 Riepilogo Domini")
            st.dataframe(
                df_riepilogo.style.apply(
                    lambda row: [
                        'background-color: #f8d7da' if row['SEO Score'] <= 49 else
                        'background-color: #fff3cd' if row['SEO Score'] <= 69 else
                        'background-color: #d4edda' for _ in row
                    ],
                    axis=1
                )
            )

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                for dominio, df_kpi in output.items():
                    df_kpi.to_excel(writer, sheet_name=dominio[:31], index=False)
                df_totale.to_excel(writer, sheet_name='Riepilogo', index=False)
                df_riepilogo.to_excel(writer, sheet_name='Riepilogo Score', index=False)

            st.download_button("📥 Scarica il report Excel", buffer.getvalue(), file_name="multi_seo_audit.xlsx")