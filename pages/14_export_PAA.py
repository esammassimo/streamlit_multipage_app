import time
import io
from datetime import date
import requests
import pandas as pd
import streamlit as st

API_URL = "https://serpapi.com/search.json"

# ==========================
# CACHE GIORNALIERA
# ==========================
_CACHE_SS = "_paa_cache"

def _cache_get(key):
    cache = st.session_state.get(_CACHE_SS, {})
    if cache.get("date") != date.today().isoformat():
        return None
    return cache.get("data", {}).get(key)

def _cache_set(key, value):
    today = date.today().isoformat()
    cache = st.session_state.get(_CACHE_SS, {})
    if cache.get("date") != today:
        cache = {"date": today, "data": {}}
    cache["data"][key] = value
    st.session_state[_CACHE_SS] = cache


# ==========================
# SUPPORTO CONFIGURAZIONE
# ==========================

def get_language_options():
    return {
        "Italiano": "it",
        "English": "en",
        "Español": "es",
        "Français": "fr",
        "Deutsch": "de",
        "Português": "pt",
        "Nederlands": "nl",
        "日本語 (Giapponese)": "ja"
    }


def get_country_options():
    return {
        "Italia": "it",
        "Stati Uniti": "us",
        "Regno Unito": "uk",
        "Spagna": "es",
        "Francia": "fr",
        "Germania": "de",
        "Portogallo": "pt",
        "Paesi Bassi": "nl",
        "Giappone": "jp"
    }


# ==========================
# FUNZIONI DI BACKEND
# ==========================

def get_top4_paa_with_meta_and_indicators(keyword, api_key, gl="it", hl="it"):
    """
    Chiama SerpAPI e restituisce le prime 4 PAA con title/link/snippet e flag booleani.
    """
    cache_key = f"{keyword}|{gl}|{hl}"
    cached = _cache_get(cache_key)
    if cached is not None:
        return cached

    params = {
        "engine": "google",
        "q": keyword,
        "api_key": api_key,
        "gl": gl,
        "hl": hl
    }

    try:
        resp = requests.get(API_URL, params=params, timeout=30)
    except requests.exceptions.RequestException as e:
        st.warning(f"❌ Errore di rete per '{keyword}': {e}")
        return []
    if resp.status_code != 200:
        st.warning(f"❌ Errore per '{keyword}' – status {resp.status_code}: {resp.text[:200]}")
        return []

    try:
        data = resp.json()
    except ValueError:
        st.warning(f"❌ Risposta non valida per '{keyword}' (JSON non parsabile).")
        return []
    rq = data.get("related_questions", [])

    results = []
    seen = set()
    for item in rq:
        q = item.get("question")
        if not q or q in seen:
            continue
        seen.add(q)
        title = item.get("title", "")
        link = item.get("link", "")
        snippet = item.get("snippet", "")
        results.append({
            "question": q,
            "title": title,
            "link": link,
            "snippet": snippet,
            "has_snippet": int(bool(snippet)),
            "has_link": int(bool(link))
        })
        if len(results) >= 4:
            break

    _cache_set(cache_key, results)
    return results


def read_keywords(file_obj, sheet_name):
    """
    Legge le keyword dalla prima colonna del foglio specificato.
    file_obj può essere sia un path sia un file caricato da Streamlit.
    """
    df = pd.read_excel(file_obj, sheet_name=sheet_name, engine="openpyxl")
    return df.iloc[:, 0].dropna().astype(str).tolist()


def scrape_keywords_to_dataframe(keywords, api_key, sleep_seconds=1, hl="it", gl="it"):
    """
    Esegue lo scraping delle PAA per una lista di keyword e restituisce un DataFrame.
    """
    rows = []
    progress_bar = st.progress(0)
    status_text = st.empty()

    total = len(keywords)

    for idx, kw in enumerate(keywords, start=1):
        is_cached = _cache_get(f"{kw}|{gl}|{hl}") is not None
        label = "(da cache) " if is_cached else ""
        status_text.write(f"🔍 {label}Elaboro keyword: **{kw}** ({idx}/{total})")

        items = get_top4_paa_with_meta_and_indicators(kw, api_key=api_key, gl=gl, hl=hl)
        if items:
            for it in items:
                rows.append({
                    "Keyword": kw,
                    "Domanda PAA": it["question"],
                    "Title": it["title"],
                    "Link": it["link"],
                    "Snippet": it["snippet"],
                    "Has Snippet": it["has_snippet"],
                    "Has Link": it["has_link"]
                })
        else:
            rows.append({
                "Keyword": kw,
                "Domanda PAA": "Nessuna domanda trovata",
                "Title": "",
                "Link": "",
                "Snippet": "",
                "Has Snippet": 0,
                "Has Link": 0
            })

        progress_bar.progress(idx / total)
        if not is_cached:
            time.sleep(sleep_seconds)

    status_text.write("✅ Elaborazione completata.")
    df = pd.DataFrame(rows)
    return df


def dataframe_to_excel_bytes(df, sheet_name="PAA"):
    """
    Converte un DataFrame in un file Excel in memoria (BytesIO) pronto per il download.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output


# ==========================
# PAGINE DELL'APP
# ==========================

def page_config():
    st.title("🔑 Configurazione API SerpAPI")

    st.write(
        """
        Inserisci la tua **API key di SerpAPI**.  
        Verrà usata per tutte le richieste nella pagina di scraping.
        """
    )

    api_key = st.text_input(
        "SerpAPI API Key",
        type="password",
        value=st.session_state.get("api_key", "")
    )

    if api_key:
        st.session_state["api_key"] = api_key
        st.success("API key salvata in sessione.")
    else:
        st.info("Inserisci la tua API key per iniziare.")

    st.markdown("---")
    st.subheader("Impostazioni lingua / country")

    use_custom_locale = st.checkbox(
        "Usa lingua e country personalizzati",
        value=st.session_state.get("use_custom_locale", False)
    )
    st.session_state["use_custom_locale"] = use_custom_locale

    language_options = get_language_options()
    country_options = get_country_options()

    if use_custom_locale:
        current_hl = st.session_state.get("hl", "it")
        current_gl = st.session_state.get("gl", "it")

        language_values = list(language_options.values())
        country_values = list(country_options.values())

        col1, col2 = st.columns(2)

        with col1:
            selected_language_label = st.selectbox(
                "Lingua risultati (hl)",
                options=list(language_options.keys()),
                index=language_values.index(current_hl) if current_hl in language_values else 0
            )

        with col2:
            selected_country_label = st.selectbox(
                "Country Google (gl)",
                options=list(country_options.keys()),
                index=country_values.index(current_gl) if current_gl in country_values else 0
            )

        st.session_state["hl"] = language_options[selected_language_label]
        st.session_state["gl"] = country_options[selected_country_label]

        st.info(
            f"Configurazione attiva: lingua = **{st.session_state['hl']}**, "
            f"country = **{st.session_state['gl']}**"
        )
    else:
        st.session_state["hl"] = "it"
        st.session_state["gl"] = "it"
        st.info("Configurazione standard attiva: lingua = **it**, country = **it**")


def page_scraping():
    st.title("📥 Upload file & Scraping PAA")

    # Controllo che l'API key sia stata impostata
    api_key = st.session_state.get("api_key")
    if not api_key:
        st.warning("⚠️ Prima imposta la tua SerpAPI API key nella pagina **Configurazione**.")
        return

    hl = st.session_state.get("hl", "it")
    gl = st.session_state.get("gl", "it")

    st.write(
        f"""
        Carica un file Excel con le keyword nella **prima colonna** del foglio selezionato.
        Poi avvia lo scraping delle People Also Ask (PAA).

        **Configurazione attuale**
        - Lingua (`hl`): **{hl}**
        - Country (`gl`): **{gl}**
        """
    )

    uploaded_file = st.file_uploader("Carica file Excel (.xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        # Mostro i fogli presenti
        try:
            xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
            sheet_name = st.selectbox("Seleziona il foglio con le keyword", xls.sheet_names)
        except Exception as e:
            st.error(f"Errore nella lettura del file: {e}")
            return

        sleep_seconds = st.slider(
            "Delay tra le richieste (sec) – aiuta a non saturare l'API",
            min_value=0.0,
            max_value=5.0,
            value=1.0,
            step=0.5
        )

        if st.button("🚀 Avvia scraping PAA"):
            try:
                keywords = read_keywords(uploaded_file, sheet_name=sheet_name)
                if not keywords:
                    st.warning("Nessuna keyword trovata nella prima colonna del foglio selezionato.")
                    return

                st.write(f"Trovate **{len(keywords)}** keyword. Inizio elaborazione…")

                df = scrape_keywords_to_dataframe(
                    keywords=keywords,
                    api_key=api_key,
                    sleep_seconds=sleep_seconds,
                    hl=hl,
                    gl=gl
                )

                st.session_state["paa_df"] = df

                st.success("Scraping completato. Vai alla pagina **Risultati** per esportare l'Excel.")
                st.dataframe(df.head())
            except Exception as e:
                st.error(f"Errore durante lo scraping: {e}")

    else:
        st.info("Carica un file Excel per procedere.")


def page_results():
    st.title("📊 Risultati & Download")

    df = st.session_state.get("paa_df")

    if df is None:
        st.info("Nessun risultato disponibile. Esegui prima lo scraping nella pagina **Upload & Scraping**.")
        return

    st.write("Anteprima risultati:")
    st.dataframe(df)

    default_filename = "PAA_export.xlsx"
    filename = st.text_input("Nome file di output", value=default_filename)

    excel_bytes = dataframe_to_excel_bytes(df, sheet_name="PAA")

    st.download_button(
        label="💾 Scarica Excel con PAA",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==========================
# MAIN: NAVIGAZIONE MULTIPAGINA
# ==========================

def main():
    st.set_page_config(
        page_title="PAA Scraper SerpAPI",
        page_icon="🔎",
        layout="wide"
    )

    if "hl" not in st.session_state:
        st.session_state["hl"] = "it"
    if "gl" not in st.session_state:
        st.session_state["gl"] = "it"
    if "use_custom_locale" not in st.session_state:
        st.session_state["use_custom_locale"] = False

    st.sidebar.title("Navigazione")
    page = st.sidebar.radio(
        "Vai a:",
        ("1. Configurazione API", "2. Upload & Scraping", "3. Risultati")
    )

    if page.startswith("1"):
        page_config()
    elif page.startswith("2"):
        page_scraping()
    elif page.startswith("3"):
        page_results()


if __name__ == "__main__":
    main()