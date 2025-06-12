import pandas as pd
import streamlit as st
import re
from urllib.parse import urlparse
import io
import os

try:
    import matplotlib.pyplot as plt
    import numpy as np
    MATPLOTLIB_OK = True
except ModuleNotFoundError:
    MATPLOTLIB_OK = False
    st.warning("Modulo 'matplotlib' mancante. Per visualizzare i radar chart, esegui: pip install matplotlib")

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

    penalita_status = (kpi['Pagine 3xx'] + kpi['Pagine 4xx'] + kpi['Bloccate da Robots.txt']) / pagine_totali

    canonical_non_self = 0
    if 'Canonical Link Element 1' in df.columns and 'Canonical Link Element 1 Resolved' in df.columns:
        canon_df = df[['Address', 'Canonical Link Element 1 Resolved']].dropna()
        canonical_non_self = canon_df.apply(
            lambda r: r['Address'].rstrip('/').lower() != r['Canonical Link Element 1 Resolved'].rstrip('/').lower(), axis=1
        ).sum() / pagine_totali

    html_penalita = (
        (kpi['Title Duplicati'] + kpi['Title Mancanti']) +
        (kpi['Meta Description Duplicati'] + kpi['Meta Description Mancanti']) +
        (kpi['H1 Duplicati'] + kpi['H1 Mancanti'])
    ) / (3 * pagine_totali)

    penalita_duplicate = kpi['Pagine Duplicate'] / pagine_totali if isinstance(kpi['Pagine Duplicate'], (int, float)) else 0

    cwv_penalita = 0
    cwv_colonne = ['LCP', 'INP', 'CLS', 'FCP', 'TTFB']
    soglie = {'LCP': 2500, 'INP': 200, 'CLS': 0.1, 'FCP': 1800, 'TTFB': 800}
    penalita_cwv = []
    for metrica in cwv_colonne:
        if metrica in kpi:
            val = kpi[metrica]
            soglia = soglie[metrica]
            if isinstance(val, (int, float)) and val > 0:
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
            valid[valid.duplicated(keep=False)].nunique(),
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

    pagine_duplicate = df['Duplicate Content'].sum() if 'Duplicate Content' in df.columns and pd.api.types.is_numeric_dtype(df['Duplicate Content']) else 0
    immagini_senza_alt = df['Images Missing Alt Text'].sum() if 'Images Missing Alt Text' in df.columns else 'N/D'
    content = {
        'Pagine Duplicate': pagine_duplicate,
        'Immagini senza ALT': immagini_senza_alt,
        'Pagine Totali': df.shape[0]
    }

    cwv = {}
    for metrica in ['LCP', 'INP', 'CLS', 'FCP', 'TTFB']:
        col = f"{metrica} (ms)"
        if col in df.columns and pd.api.types.is_numeric_dtype(df[col]):
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
