import os
import pandas as pd
import chardet
from urllib.parse import urlparse, unquote
import re
import openai
import time
import streamlit as st

st.title("🔍 Link Profiler")

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

# === GPT CLASSIFIER ===
def gpt_semantic_url_classification(urls, api_key, model="gpt-4", batch_size=10, pause=1.5):
    from openai import OpenAI
    client = OpenAI(api_key=api_key)
    result_map = {}

    def extract_path_description(url):
        path = urlparse(url).path
        return " / ".join([unquote(p) for p in path.split("/") if p])

    for i in range(0, len(urls), batch_size):
        batch = urls[i:i + batch_size]
        descrizioni = [extract_path_description(url) for url in batch]

        prompt = "Per ciascun percorso URL seguente, assegna una categoria tematica coerente (es. Scarpe > Sandali, Borse, Collezione Sposa, ecc.).\n"
        prompt += "Rispondi in formato JSON: {\"<descrizione URL>\": \"<categoria>\"}\n\n"
        for desc in descrizioni:
            prompt += f"- {desc}\n"

        try:
            response = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3
            )

            text = response.choices[0].message.content
            parsed = eval(text) if text.strip().startswith("{") else {}
            for url, desc in zip(batch, descrizioni):
                result_map[url] = parsed.get(desc, "Non classificato")

        except Exception as e:
            for url in batch:
                result_map[url] = f"Errore: {e}"

        time.sleep(pause)

    return result_map

# === STREAMLIT APP ===
st.title("🔎 Link Profiler per Brand")

api_key = st.text_input("Inserisci la tua OpenAI API Key:", type="password")
if not api_key:
    st.warning("⚠ Inserisci la tua OpenAI API Key per continuare.")
    st.stop()

brand_keywords_input = st.text_input("Parole chiave del brand (separate da virgola):")
brand_keywords = [kw.strip().lower() for kw in brand_keywords_input.split(",") if kw.strip() != ""]

uploaded_files = st.file_uploader("Carica uno o più file CSV (uno per brand) con le colonne 'Target URL', 'Anchor', 'Domain rating':", type=["csv"], accept_multiple_files=True)

if uploaded_files:
    dfs = []
    for uploaded_file in uploaded_files:
        st.markdown(f"### 📄 {uploaded_file.name}")
        encoding = detect_encoding(uploaded_file.name)
        df = pd.read_csv(uploaded_file, encoding=encoding, sep=None, engine="python")

        if not all(col in df.columns for col in ["Target URL", "Anchor", "Domain rating"]):
            st.error(f"❌ Il file {uploaded_file.name} deve contenere le colonne richieste.")
            continue

        with st.spinner(f"🔄 Elaborazione in corso per {uploaded_file.name}..."):
            df["Dominio"] = os.path.splitext(uploaded_file.name)[0]
            df["Domain rating class"] = df["Domain rating"].apply(classify_domain_rating)
            df["Anchor class"] = df["Anchor"].apply(lambda x: classify_anchor(x, brand_keywords))
            df["URL structure class"] = df["Target URL"].apply(classify_url_structure)

            url_list = df["Target URL"].dropna().unique().tolist()
            url_category_map = gpt_semantic_url_classification(url_list, api_key)
            df["URL Category"] = df["Target URL"].apply(lambda x: url_category_map.get(x, "UNKNOWN"))

            st.success("✅ Classificazione completata")
            st.dataframe(df.head(50))

            dfs.append(df)

    if dfs:
        final_df = pd.concat(dfs, ignore_index=True)
        full_output_path = "classificazione_completa_tutti_brand.xlsx"
        final_df.to_excel(full_output_path, index=False)
        with open(full_output_path, "rb") as f:
            st.download_button("📥 Scarica file Excel unificato", f, file_name=full_output_path)

        grouped = final_df.groupby(["Dominio", "URL Category"]).size().reset_index(name="Totale link")
        st.subheader("📊 Riepilogo per Dominio e Categoria")
        st.dataframe(grouped)