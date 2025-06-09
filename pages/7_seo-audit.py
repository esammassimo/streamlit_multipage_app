import pandas as pd
import streamlit as st
import re
from urllib.parse import urlparse
import io
import os

# Titolo app
st.title("SEO Audit")

def estrai_dominio(df):
    try:
        url_sample = df['Address'].dropna().iloc[0]
        dominio = urlparse(url_sample).netloc.replace('www.', '')
        return dominio if dominio else None
    except Exception:
        return None

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

    return pd.DataFrame([{
        **status,
        **html_tag,
        **content,
        **cwv
    }])

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
            st.dataframe(kpi)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                kpi.to_excel(writer, sheet_name='Report SEO', index=False)

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
            st.dataframe(df_totale)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                for dominio, df_kpi in output.items():
                    df_kpi.to_excel(writer, sheet_name=dominio[:31], index=False)
                df_totale.to_excel(writer, sheet_name='Riepilogo', index=False)

            st.download_button("📥 Scarica il report Excel", buffer.getvalue(), file_name="multi_seo_audit.xlsx")
        else:
            st.warning("Nessun file valido caricato. Assicurati che ogni file contenga il foglio '1 - HTML' o '1 - All'.")