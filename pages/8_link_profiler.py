import os
import pandas as pd
import chardet
from urllib.parse import urlparse, unquote
import re
import streamlit as st
import tempfile


# === FUNZIONE DI RILEVAMENTO ENCODING ===
def detect_encoding(file_path):
    with open(file_path, "rb") as f:
        result = chardet.detect(f.read(10000))
    return result.get("encoding") or "utf-8"


# === UTILITÀ DI NORMALIZZAZIONE COLONNE ===
CANONICAL_COLS_MAP = {
    "target url": "Target URL",
    "url": "Target URL",
    "anchor": "Anchor",
    "domain rating": "Domain rating",
    "dr": "Domain rating",
    "first seen": "First seen",
    "first_seen": "First seen",
    "firstseen": "First seen",
}


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    new_cols = []
    for c in df.columns:
        c_norm = str(c).replace("\ufeff", "").strip().lower()
        new_cols.append(CANONICAL_COLS_MAP.get(c_norm, str(c).strip()))
    df.columns = new_cols
    return df


# === CLASSIFICAZIONI ===
def classify_domain_rating(rating):
    if pd.isna(rating):
        return "UNKNOWN"
    try:
        r = float(rating)
    except Exception:
        return "UNKNOWN"

    if r < 10:
        return "JUNK"
    elif r < 30:
        return "BOTTOM"
    elif r < 70:
        return "MEDIUM"
    else:
        return "HIGH"


def classify_anchor(text, brand_keywords):
    if pd.isna(text) or not isinstance(text, str) or text.strip() == "":
        return "UNKNOWN"

    text_lower = text.lower().strip()
    word_count = len(text_lower.split())

    if re.search(r"https?://|www\.|\.it|\.com|\.net", text_lower):
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
        "chi siamo", "contatti", "servizi", "home", "galleria", "news",
        "prodotti", "i nostri", "scopri il", "maggiori informazioni"
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
    except Exception:
        return "UNKNOWN"


def extract_second_level_domain(url):
    """
    Estrae il dominio a 2 livelli ignorando i sottodomini.
    Esempi:
      - https://blog.example.com/page -> example.com
      - https://news.example.co.uk/a -> co.uk (euristico, non PSL-based)
    """
    try:
        netloc = urlparse(url).netloc
        if ":" in netloc:
            netloc = netloc.split(":")[0]
        netloc = netloc.replace("www.", "").strip(".")
        parts = [p for p in netloc.split(".") if p]
        if len(parts) >= 2:
            return ".".join(parts[-2:])
        return netloc or "UNKNOWN"
    except Exception:
        return "UNKNOWN"


def normalize_url(u):
    if pd.isna(u) or not isinstance(u, str):
        return u
    return unquote(u.strip())


def normalize_anchor(a):
    if pd.isna(a) or not isinstance(a, str):
        return a
    return a.strip()


# === GESTIONE DATE ROBUSTA (Ahrefs / Screaming Frog / CSV misti) ===
DATE_ALREADY_CLEAN_REGEX = re.compile(r"^\d{2}/\d{2}/\d{4}$")


def is_already_clean_ddmmyyyy(series: pd.Series) -> bool:
    """
    Evita parsing inutile se tutti i valori non nulli sono già nel formato gg/mm/aaaa.
    """
    non_null = series.dropna().astype(str).str.strip()
    if non_null.empty:
        return False
    return non_null.str.match(DATE_ALREADY_CLEAN_REGEX).all()


def clean_date_string(value):
    """
    Normalizza stringhe data comuni in export Ahrefs / Screaming Frog:
    - 2024-06-18
    - 2024-06-18 14:25:31
    - 2024-06-18T14:25:31
    - 18/06/2024
    - 18-06-2024
    - Jun 18 2024
    - 2024/06/18
    """
    if pd.isna(value):
        return pd.NA

    s = str(value).strip()
    if s == "":
        return pd.NA

    s = s.replace("\u00a0", " ").strip()

    # se è già pulita gg/mm/aaaa, la restituiamo senza parsing
    if DATE_ALREADY_CLEAN_REGEX.match(s):
        return s

    # rimuove eventuale timezone Z finale
    s = re.sub(r"Z$", "", s)

    # gestisce date ISO con T
    s = s.replace("T", " ")

    return s


def parse_single_date(value):
    """
    Parsing robusto e compatibile con pandas 2.x.
    Restituisce stringa gg/mm/aaaa oppure pd.NA.
    """
    if pd.isna(value):
        return pd.NA

    s = clean_date_string(value)
    if pd.isna(s):
        return pd.NA

    # Tentativo 1: parsing generico/misto (molto utile con export vari)
    try:
        dt = pd.to_datetime(s, errors="coerce", format="mixed", dayfirst=False)
    except Exception:
        dt = pd.to_datetime(s, errors="coerce", dayfirst=False)

    # Tentativo 2: se fallisce, prova interpretazione dayfirst=True
    if pd.isna(dt):
        try:
            dt = pd.to_datetime(s, errors="coerce", format="mixed", dayfirst=True)
        except Exception:
            dt = pd.to_datetime(s, errors="coerce", dayfirst=True)

    # Tentativo 3: se c'è data + ora, prova solo la parte data
    if pd.isna(dt):
        try:
            only_date = str(s).split(" ")[0]
            try:
                dt = pd.to_datetime(only_date, errors="coerce", format="mixed", dayfirst=False)
            except Exception:
                dt = pd.to_datetime(only_date, errors="coerce", dayfirst=False)
        except Exception:
            return pd.NA

    if pd.isna(dt):
        return pd.NA

    return dt.strftime("%d/%m/%Y")


def parse_date_column(series: pd.Series) -> pd.Series:
    """
    Evita parsing se la colonna è già pulita.
    Altrimenti applica parsing robusto valore per valore.
    """
    if is_already_clean_ddmmyyyy(series):
        return series.astype("string").str.strip()

    return series.apply(parse_single_date)


# === STREAMLIT APP ===
st.title("🔎 Link Profiler (Tier, Anchor, Page Level)")
st.text("I nomi dei file devono essere nomedominio.it.csv e non contenere ulteriori caratteri")

brand_keywords_input = st.text_input("Parole chiave del brand (separate da virgola):")
brand_keywords = [kw.strip().lower() for kw in brand_keywords_input.split(",") if kw.strip()]

uploaded_files = st.file_uploader(
    "Carica uno o più file CSV (uno per brand) con le colonne 'Target URL', 'Anchor', 'Domain rating':",
    type=["csv"],
    accept_multiple_files=True
)

if uploaded_files:
    dfs = []
    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, uploaded_file in enumerate(uploaded_files):
        status_text.text(f"Elaborazione file: {uploaded_file.name} ({idx + 1}/{len(uploaded_files)})")

        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name

            encoding = detect_encoding(tmp_path)

            # prova lettura con separatore auto
            try:
                df = pd.read_csv(tmp_path, encoding=encoding, sep=None, engine="python")
            except Exception:
                # fallback comune
                df = pd.read_csv(tmp_path, encoding=encoding, sep=",", engine="python")

            # --- Normalizza intestazioni ---
            df = normalize_columns(df)

            # --- Verifica colonne minime richieste ---
            required = ["Target URL", "Anchor", "Domain rating"]
            missing = [c for c in required if c not in df.columns]
            if missing:
                st.error(
                    f"❌ Il file {uploaded_file.name} non contiene tutte le colonne richieste "
                    f"{required} (mancano: {missing}) e sarà ignorato."
                )
                progress_bar.progress((idx + 1) / len(uploaded_files))
                continue

            # --- Normalizza valori base ---
            df["Target URL"] = df["Target URL"].apply(normalize_url)
            df["Anchor"] = df["Anchor"].apply(normalize_anchor)

            # --- Domain rating numerico sicuro ---
            df["Domain rating"] = pd.to_numeric(df["Domain rating"], errors="coerce")

            # --- Enrichment colonne ---
            df["Dominio"] = df["Target URL"].apply(extract_second_level_domain)
            df["Domain rating class"] = df["Domain rating"].apply(classify_domain_rating)
            df["Anchor class"] = df["Anchor"].apply(lambda x: classify_anchor(x, brand_keywords))
            df["URL structure class"] = df["Target URL"].apply(classify_url_structure)

            # --- First seen (date) se presente ---
            if "First seen" in df.columns:
                df["First seen (date)"] = parse_date_column(df["First seen"])

            dfs.append(df)

        except Exception as e:
            st.error(f"❌ Errore durante l'elaborazione del file {uploaded_file.name}: {e}")

        progress_bar.progress((idx + 1) / len(uploaded_files))

    progress_bar.empty()
    status_text.text("✅ Tutti i file sono stati elaborati.")

    if dfs:
        final_df = pd.concat(dfs, ignore_index=True)

        # --- Download Excel unificato ---
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_output_xlsx:
            xlsx_path = tmp_output_xlsx.name
            final_df.to_excel(xlsx_path, index=False)
            with open(xlsx_path, "rb") as f:
                st.download_button(
                    "📥 Scarica file Excel unificato",
                    f,
                    file_name="link_profiler_classificato.xlsx"
                )

        # --- Download CSV unificato ---
        with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp_output_csv:
            csv_path = tmp_output_csv.name
            final_df.to_csv(csv_path, index=False)
            with open(csv_path, "rb") as f:
                st.download_button(
                    "📥 Scarica CSV unificato",
                    f,
                    file_name="link_profiler_classificato.csv"
                )

        st.subheader("📊 Riepiloghi per Dominio")

        # 1. Numero di link per dominio
        link_count = final_df.groupby("Dominio").size().reset_index(name="Totale link")
        link_count = link_count.set_index("Dominio")
        st.markdown("**Numero di link per Dominio**")
        st.dataframe(link_count)

        # 2. Domain rating per dominio
        dr_summary = final_df.pivot_table(
            index="Dominio",
            columns="Domain rating class",
            aggfunc="size",
            fill_value=0
        )
        dr_summary.loc["TOTALE"] = dr_summary.sum()
        st.markdown("**Classificazione Domain Rating per Dominio**")
        st.dataframe(dr_summary)

        # 3. Anchor type per dominio
        anchor_summary = final_df.pivot_table(
            index="Dominio",
            columns="Anchor class",
            aggfunc="size",
            fill_value=0
        )
        anchor_summary.loc["TOTALE"] = anchor_summary.sum()
        st.markdown("**Anchor Type per Dominio**")
        st.dataframe(anchor_summary)

        # 4. URL structure per dominio
        structure_summary = final_df.pivot_table(
            index="Dominio",
            columns="URL structure class",
            aggfunc="size",
            fill_value=0
        )
        structure_summary.loc["TOTALE"] = structure_summary.sum()
        st.markdown("**Page Level per Dominio**")
        st.dataframe(structure_summary)