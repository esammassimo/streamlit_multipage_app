import os
import pandas as pd
import chardet
from urllib.parse import urlparse, unquote
import re
import streamlit as st
import tempfile

# === FUNZIONE DI RILEVAMENTO ENCODING ===
def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read(10000))
    return result['encoding']

# === VARIABILI GLOBALI DI SUPPORTO ===
def classify_domain_rating(rating):
    if pd.isna(rating): return "UNKNOWN"
    if rating < 10: return 'JUNK'
    elif rating < 30: return 'BOTTOM'
    elif rating < 70: return 'MEDIUM'
    else: return 'HIGH'

def classify_anchor(text, brand_keywords):
    if pd.isna(text) or not isinstance(text, str) or text.strip() == "":
        return "UNKNOWN"
    text_lower = text.lower().strip()
    word_count = len(text_lower.split())
    if re.search(r'https?://|www\.|\.it|\.com|\.net', text_lower):
        return "URL"
    if any(brand.lower() in text_lower for brand in brand_keywords):
        return "BRAND"
    distraction_phrases = [
        "clicca qui", "scopri di più", "approfondisci", "guarda questo",
        "leggi di più", "visita il sito", "vai alla pagina", "clicca per info"
    ]
    if any(phrase in text_lower for phrase in distraction_phrases):
        return "DISTRACTION"
    navigational_terms = [
        "chi siamo", "contatti", "servizi", "home", "galleria", "news", "prodotti",
        "i nostri", "scopri il", "maggiori informazioni"
    ]
    if word_count > 3 or any(term in text_lower for term in navigational_terms):
        return "PHRASE"
    return "EXACT"

def classify_url_structure(url):
    try:
        path = urlparse(url).path
        segments = [seg for seg in path.strip("/").split("/") if seg]
        if len(segments) == 0:
            return "Home"
        elif len(segments) == 1:
            return "Pagina 1° livello"
        else:
            return "Pagina oltre il 1° Livello"
    except:
        return "UNKNOWN"

def extract_domain_from_url(url):
    try:
        parsed_url = urlparse(url)
        return parsed_url.netloc.replace("www.", "")
    except:
        return "UNKNOWN"

# === STREAMLIT APP ===
st.title("🔎 Link Profiler (Tier, Anchor, Page Level)")
st.text("I nomi dei file devono essere nomedominio.it.csv e non contenere ulteriori caratteri")

brand_keywords_input = st.text_input("Parole chiave del brand (separate da virgola):")
brand_keywords = [kw.strip().lower() for kw in brand_keywords_input.split(",") if kw.strip() != ""]

uploaded_files = st.file_uploader("Carica uno o più file CSV (uno per brand) con le colonne 'Target URL', 'Anchor', 'Domain rating':", type=["csv"], accept_multiple_files=True)

if uploaded_files:
    dfs = []
    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, uploaded_file in enumerate(uploaded_files):
        status_text.text(f"Elaborazione file: {uploaded_file.name} ({idx+1}/{len(uploaded_files)})")
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name
            encoding = detect_encoding(tmp_path)
            df = pd.read_csv(tmp_path, encoding=encoding, sep=None, engine="python")

            if not all(col in df.columns for col in ["Target URL", "Anchor", "Domain rating"]):
                st.error(f"❌ Il file {uploaded_file.name} non contiene tutte le colonne richieste e sarà ignorato.")
                continue

            df["Dominio"] = df["Target URL"].apply(extract_domain_from_url)
            df["Domain rating class"] = df["Domain rating"].apply(classify_domain_rating)
            df["Anchor class"] = df["Anchor"].apply(lambda x: classify_anchor(x, brand_keywords))
            df["URL structure class"] = df["Target URL"].apply(classify_url_structure)

            dfs.append(df)
        except Exception as e:
            st.error(f"❌ Errore durante l'elaborazione del file {uploaded_file.name}: {e}")
            continue

        progress_bar.progress((idx + 1) / len(uploaded_files))

    progress_bar.empty()
    status_text.text("✅ Tutti i file sono stati elaborati.")

    if dfs:
        final_df = pd.concat(dfs, ignore_index=True)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_output:
            full_output_path = tmp_output.name
            final_df.to_excel(full_output_path, index=False)
            with open(full_output_path, "rb") as f:
                st.download_button("📥 Scarica file Excel unificato", f, file_name="link_profiler_classificato.xlsx")

        st.subheader("📊 Riepiloghi per Dominio")

        # 1. Numero di link per dominio
        link_count = final_df.groupby("Dominio").size().reset_index(name="Totale link")
        link_count = link_count.set_index("Dominio")
        st.markdown("**Numero di link per Dominio**")
        st.dataframe(link_count)

        # 2. Domain rating per dominio
        dr_summary = final_df.pivot_table(index="Dominio", columns="Domain rating class", aggfunc="size", fill_value=0)
        dr_summary.loc["TOTALE"] = dr_summary.sum()
        st.markdown("**Classificazione Domain Rating per Dominio**")
        st.dataframe(dr_summary)

        # 3. Anchor type per dominio
        anchor_summary = final_df.pivot_table(index="Dominio", columns="Anchor class", aggfunc="size", fill_value=0)
        anchor_summary.loc["TOTALE"] = anchor_summary.sum()
        st.markdown("**Anchor Type per Dominio**")
        st.dataframe(anchor_summary)

        # 4. URL structure per dominio
        structure_summary = final_df.pivot_table(index="Dominio", columns="URL structure class", aggfunc="size", fill_value=0)
        structure_summary.loc["TOTALE"] = structure_summary.sum()
        st.markdown("**Page Level per Dominio**")
        st.dataframe(structure_summary)