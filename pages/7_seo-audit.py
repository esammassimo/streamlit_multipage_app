import pandas as pd
import streamlit as st
import re
from urllib.parse import urlparse
import io
import os

def estrai_dominio(df):
    try:
        url_sample = df['Address'].dropna().iloc[0]
        dominio = urlparse(url_sample).netloc.replace('www.', '')
        return dominio if dominio else None
    except Exception:
        return None

def calcola_score(df, kpi):
    pagine_totali = df.shape[0]
    if pagine_totali == 0:
        return 0, {}

    # Status Code e robots
    penalita_status = (kpi['Pagine 3xx'] + kpi['Pagine 4xx'] + kpi['Bloccate da Robots.txt']) / pagine_totali

    # Canonical non self-referencing
    canonical_non_self = 0
    if 'Canonical Link Element 1' in df.columns and 'Canonical Link Element 1 Resolved' in df.columns:
        canon_df = df[['Address', 'Canonical Link Element 1 Resolved']].dropna()
        canonical_non_self = (canon_df['Address'] != canon_df['Canonical Link Element 1 Resolved']).sum() / pagine_totali

    # HTML Tag (Title, Desc, H1)
    html_penalita = (
        (kpi['Title Duplicati'] + kpi['Title Mancanti']) +
        (kpi['Meta Description Duplicati'] + kpi['Meta Description Mancanti']) +
        (kpi['H1 Duplicati'] + kpi['H1 Mancanti'])
    ) / (3 * pagine_totali)

    # Contenuti duplicati
    penalita_duplicate = kpi['Pagine Duplicate'] / pagine_totali if isinstance(kpi['Pagine Duplicate'], (int, float)) else 0

    # CWV penalty (valore tra 0 e 1)
    cwv_penalita = 0
    cwv_colonne = ['LCP', 'INP', 'CLS', 'FCP', 'TTFB']
    soglie = {'LCP': 2500, 'INP': 200, 'CLS': 0.1, 'FCP': 1800, 'TTFB': 800}
    penalita_cwv = []
    for metrica in cwv_colonne:
        if metrica in kpi:
            val = kpi[metrica]
            soglia = soglie[metrica]
            if val > 0:
                if metrica == 'CLS':
                    penalita_cwv.append(min(1.0, val / soglia))
                else:
                    penalita_cwv.append(min(1.0, (val - soglia) / soglia))
    if penalita_cwv:
        cwv_penalita = sum(penalita_cwv) / len(penalita_cwv)

    score_components = {
        'Penalità Status Code %': round(penalita_status * 100, 1),
        'Penalità Canonical %': round(canonical_non_self * 100, 1),
        'Penalità Tag HTML %': round(html_penalita * 100, 1),
        'Penalità Contenuti Duplicati %': round(penalita_duplicate * 100, 1),
        'Penalità CWV %': round(cwv_penalita * 100, 1)
    }

    score = 100 * (1 - (
        0.30 * penalita_status +
        0.15 * canonical_non_self +
        0.20 * html_penalita +
        0.10 * penalita_duplicate +
        0.20 * cwv_penalita
    ))
    return round(max(score, 0), 2), score_components

def estrai_kpi(df):
    df.columns = df.columns.str.strip()
    df['Status Code'] = pd.to_numeric(df['Status Code'], errors='coerce')

    status = {
        'Pagine 2xx': df['Status Code'].between(200, 299).sum(),
        'Pagine 3xx': df['Status Code'].between(300, 399).sum(),
        'Pagine 4xx': df['Status Code'].between(400, 499).sum(),
        'Bloccate da Robots.txt': df['Indexability'].str.contains("Blocked by Robots", na=False).sum(),
        'Pagine HTML Totali': df.shape[0]
    }

    def analizza(col):
        if col not in df.columns:
            return (0, 0, 0)
        valid = df[col].dropna()
        return (
            len(valid[valid.duplicated(keep=False)].unique()),
            df[col].isna().sum(),
            df[col].notna().sum()
        )

    title = analizza("Title 1")
    description = analizza("Meta Description 1")
    h1 = analizza("H1-1")

    html_tag = {
        'Title Duplicati': title[0], 'Title Mancanti': title[1], 'Totale Title': title[2],
        'Meta Description Duplicati': description[0], 'Meta Description Mancanti': description[1], 'Totale Meta Description': description[2],
        'H1 Duplicati': h1[0], 'H1 Mancanti': h1[1], 'Totale H1': h1[2]
    }

    pagine_duplicate = df['Duplicate Content'].sum() if 'Duplicate Content' in df.columns else 'N/D'
    immagini_senza_alt = df['Images Missing Alt Text'].sum() if 'Images Missing Alt Text' in df.columns else 'N/D'
    content = {
        'Pagine Duplicate': pagine_duplicate,
        'Immagini senza ALT': immagini_senza_alt,
        'Pagine Totali': df.shape[0]
    }

    cwv = {}
    for metrica in ['LCP', 'INP', 'CLS', 'FCP', 'TTFB']:
        col = f"{metrica} (ms)"
        if col in df.columns:
            cwv[metrica] = round(df[col].mean(), 2)

    kpi = {
        **status,
        **html_tag,
        **content,
        **cwv
    }
    score, components = calcola_score(df, kpi)
    kpi['SEO Score'] = score
    kpi.update(components)
    return pd.DataFrame([kpi])

st.title("SEO Audit Tool")
tab1, tab2 = st.tabs(["Singolo File", "Multi File"])

with tab1:
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
            import matplotlib.pyplot as plt
            import numpy as np

            # Radar chart per il singolo file
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
            ax.set_title("Penalità SEO Radar Chart")

            st.pyplot(fig)

            # Radar chart comparativo tra domini
            st.subheader("Confronto Radar tra Domini")
            domini_disponibili = df_riepilogo['Dominio'].tolist()
            domini_selezionati = st.multiselect("Seleziona uno o più domini da confrontare", domini_disponibili, default=domini_disponibili[:2])

            if domini_selezionati:
                fig2, ax2 = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))
                labels = ['Status Code', 'Canonical', 'Tag HTML', 'Contenuti Duplicati', 'CWV']
                angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
                angles += angles[:1]

                for dominio in domini_selezionati:
                    riga = df_riepilogo[df_riepilogo['Dominio'] == dominio].iloc[0]
                    values = [
                        riga['Penalità Status Code %'],
                        riga['Penalità Canonical %'],
                        riga['Penalità Tag HTML %'],
                        riga['Penalità Contenuti Duplicati %'],
                        riga['Penalità CWV %'] if 'Penalità CWV %' in riga else 0
                    ]
                    values = [v if isinstance(v, (int, float)) else 0 for v in values]
                    values += values[:1]
                    ax2.plot(angles, values, label=dominio)
                    ax2.fill(angles, values, alpha=0.1)

                ax2.set_yticklabels([])
                ax2.set_xticks(angles[:-1])
                ax2.set_xticklabels(labels)
                ax2.set_title("Radar Chart Comparativo")
                ax2.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1))
                st.pyplot(fig2)
            kpi_visual = kpi.copy()
            pagine = kpi_visual['Pagine Totali'].iloc[0]
            kpi_visual['Status Error %'] = round(((kpi_visual['Pagine 3xx'] + kpi_visual['Pagine 4xx'] + kpi_visual['Bloccate da Robots.txt']) / pagine) * 100, 1)
            kpi_visual['HTML Error %'] = round(((kpi_visual['Title Duplicati'] + kpi_visual['Title Mancanti'] + kpi_visual['Meta Description Duplicati'] + kpi_visual['Meta Description Mancanti'] + kpi_visual['H1 Duplicati'] + kpi_visual['H1 Mancanti']) / (3 * pagine)) * 100, 1)
            kpi_visual['Canonical Non-Self %'] = 0  # placeholder, no canonical check in single tab
            kpi_visual['Contenuti Duplicati %'] = round((kpi_visual['Pagine Duplicate'] / pagine) * 100, 1) if isinstance(kpi_visual['Pagine Duplicate'].iloc[0], (int, float)) else 'N/D'
            kpi_riepilogo = kpi_visual[['SEO Score', 'Penalità Status Code %', 'Penalità Canonical %', 'Penalità Tag HTML %', 'Penalità Contenuti Duplicati %', 'Penalità CWV %']]
            kpi_riepilogo['Stato'] = kpi_riepilogo['SEO Score'].apply(lambda x: 'Critico' if x < 50 else ('Medio' if x < 70 else 'Buono'))
            st.dataframe(
                kpi_riepilogo.style.apply(
                    lambda row: ['background-color: #f8d7da' if row['SEO Score'] <= 49
                                 else 'background-color: #fff3cd' if row['SEO Score'] <= 69
                                 else 'background-color: #d4edda' for _ in row],
                    axis=1
                )
            )

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                kpi.to_excel(writer, sheet_name='Report SEO', index=False)
                kpi_riepilogo.to_excel(writer, sheet_name='Riepilogo Score', index=False)

            st.download_button("📥 Scarica il report Excel", buffer.getvalue(), file_name="seo_report_singolo.xlsx")
        else:
            st.warning("Il file non contiene un foglio valido ('1 - HTML' o '1 - All').")

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
            dominio = estrai_dominio(df)
            if not dominio:
                dominio = os.path.splitext(f.name)[0].split("_")[0]
            kpi = estrai_kpi(df)
            kpi.insert(0, 'Dominio', dominio)
            report_completo.append(kpi)
            output[dominio] = kpi.drop(columns=['Dominio'])

        if output:
            df_totale = pd.concat(report_completo, ignore_index=True)
            st.subheader("Riepilogo Complessivo")
            import matplotlib.pyplot as plt
            import numpy as np

            # Radar chart per ogni dominio selezionabile
            domini_unici = df_riepilogo['Dominio'].tolist()
            dominio_scelto = st.selectbox("Seleziona un dominio per visualizzare il radar chart", domini_unici)
            df_dominio = df_riepilogo[df_riepilogo['Dominio'] == dominio_scelto]

            labels = ['Status Code', 'Canonical', 'Tag HTML', 'Contenuti Duplicati', 'CWV']
            values = [
                df_dominio['Penalità Status Code %'].iloc[0],
                df_dominio['Penalità Canonical %'].iloc[0],
                df_dominio['Penalità Tag HTML %'].iloc[0],
                df_dominio['Penalità Contenuti Duplicati %'].iloc[0],
                df_dominio['Penalità CWV %'].iloc[0]
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
            ax.set_title(f"Radar Chart - {dominio_scelto}")

            st.pyplot(fig)
            df_visual = df_totale.copy()
            df_visual['Status Error %'] = round(((df_visual['Pagine 3xx'] + df_visual['Pagine 4xx'] + df_visual['Bloccate da Robots.txt']) / df_visual['Pagine Totali']) * 100, 1)
            df_visual['HTML Error %'] = round(((df_visual['Title Duplicati'] + df_visual['Title Mancanti'] + df_visual['Meta Description Duplicati'] + df_visual['Meta Description Mancanti'] + df_visual['H1 Duplicati'] + df_visual['H1 Mancanti']) / (3 * df_visual['Pagine Totali'])) * 100, 1)
            df_visual['Canonical Non-Self %'] = 0  # placeholder, canonical check solo disponibile da parsing
            df_visual['Contenuti Duplicati %'] = round((df_visual['Pagine Duplicate'] / df_visual['Pagine Totali']) * 100, 1)
            df_riepilogo = df_visual[['Dominio', 'SEO Score', 'Penalità Status Code %', 'Penalità Canonical %', 'Penalità Tag HTML %', 'Penalità Contenuti Duplicati %', 'Penalità CWV %']]
            df_riepilogo['Stato'] = df_riepilogo['SEO Score'].apply(lambda x: 'Critico' if x < 50 else ('Medio' if x < 70 else 'Buono'))
            st.dataframe(
                df_riepilogo.style.apply(
                    lambda row: ['background-color: #f8d7da' if row['SEO Score'] <= 49
                                 else 'background-color: #fff3cd' if row['SEO Score'] <= 69
                                 else 'background-color: #d4edda' for _ in row],
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
        else:
            st.warning("Nessun file valido caricato. Assicurati che ogni file contenga il foglio '1 - HTML' o '1 - All'.")