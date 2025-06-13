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

    penalita = {}

    penalita['Penalità Status Code %'] = (kpi['Pagine 3xx'] + kpi['Pagine 4xx'] + kpi['Bloccate da Robots.txt']) / pagine_totali

    canonical_non_self = 0
    if 'Canonical Link Element 1 Resolved' in df.columns:
        canon_df = df[['Address', 'Canonical Link Element 1 Resolved']].dropna()
        canonical_non_self = (canon_df['Address'] != canon_df['Canonical Link Element 1 Resolved']).sum() / pagine_totali
    penalita['Penalità Canonical %'] = canonical_non_self

    penalita['Penalità Tag HTML %'] = (
        (kpi['Title Duplicati'] + kpi['Title Mancanti'] +
         kpi['Meta Description Duplicati'] + kpi['Meta Description Mancanti'] +
         kpi['H1 Duplicati'] + kpi['H1 Mancanti']) / (3 * pagine_totali)
    )

    penalita['Penalità Contenuti Duplicati %'] = kpi['Pagine Duplicate'] / pagine_totali if isinstance(kpi['Pagine Duplicate'], (int, float)) else 0

    penalita_cwv = []
    soglie = {'LCP': 2500, 'INP': 200, 'CLS': 0.1, 'FCP': 1800, 'TTFB': 800}
    for metrica, soglia in soglie.items():
        if metrica in kpi:
            val = kpi[metrica]
            if isinstance(val, (int, float)) and val > 0:
                if metrica == 'CLS':
                    penalita_cwv.append(min(1.0, val / soglia))
                else:
                    penalita_cwv.append(min(1.0, max(0, (val - soglia) / soglia)))
    penalita['Penalità CWV %'] = sum(penalita_cwv) / len(penalita_cwv) if penalita_cwv else 0

    score = 100 * (1 - (
        0.30 * penalita['Penalità Status Code %'] +
        0.15 * penalita['Penalità Canonical %'] +
        0.20 * penalita['Penalità Tag HTML %'] +
        0.10 * penalita['Penalità Contenuti Duplicati %'] +
        0.20 * penalita['Penalità CWV %']
    ))

    return round(max(score, 0), 2), {k: round(v * 100, 2) for k, v in penalita.items()}

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
    score, penalita = calcola_score(df, kpi)
    kpi['SEO Score'] = score
    kpi.update(penalita)
    return pd.DataFrame([kpi])

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