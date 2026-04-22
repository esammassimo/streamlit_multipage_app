import time
import io
import pandas as pd
import streamlit as st
from serpapi import GoogleSearch

# ==========================
# FUNZIONI DI BACKEND
# ==========================

def fetch_ai_overview(query, api_key, hl="it", gl="it"):
    """
    Chiama SerpAPI e restituisce l'oggetto 'ai_overview' (se presente).
    """
    params = {
        "q": query,
        "engine": "google",
        "api_key": api_key,
        "hl": hl,
        "gl": gl,
        "no_cache": False  # usa la cache se la query è già stata fatta
    }
    search = GoogleSearch(params)
    results = search.get_dict()
    return results.get("ai_overview")


def read_keywords_from_excel(file_obj, column_name="Keyword"):
    """
    Legge le parole chiave dal file Excel nella colonna specificata.
    file_obj può essere sia un path che un file caricato da Streamlit.
    """
    df_keywords = pd.read_excel(file_obj, usecols=[column_name])
    keywords = df_keywords[column_name].dropna().astype(str).tolist()
    return keywords


def scrape_ai_overviews_to_dataframe(keywords, api_key, sleep_seconds=1.0, hl="it", gl="it"):
    """
    Esegue la ricerca AI Overview per una lista di keyword e restituisce un DataFrame.
    """
    data = []

    progress_bar = st.progress(0)
    status_text = st.empty()
    total = len(keywords)

    for idx, keyword in enumerate(keywords, start=1):
        status_text.write(f"🔍 Ricerca per: **{keyword}** ({idx}/{total})")

        try:
            ai_overview = fetch_ai_overview(
                keyword,
                api_key=api_key,
                hl=hl,
                gl=gl
            )

            if ai_overview and "references" in ai_overview:
                references = ai_overview["references"]
                for ref_idx, ref in enumerate(references, start=1):
                    data.append({
                        "keyword": keyword,
                        "title": ref.get("title", ""),
                        "link": ref.get("link", ""),
                        "snippet": ref.get("snippet", ""),
                        "source": ref.get("source", ""),
                        "index": ref_idx,
                        "language": hl,
                        "country": gl
                    })
                st.write(f"✅ {len(references)} riferimenti trovati per **{keyword}**.")
            else:
                st.write(f"⚠️ Nessun AI Overview o riferimenti trovati per **{keyword}**.")

        except Exception as e:
            st.error(f"❌ Errore durante la ricerca per '{keyword}': {e}")

        progress_bar.progress(idx / total)
        time.sleep(sleep_seconds)

    status_text.write("✅ Elaborazione completata.")
    df_results = pd.DataFrame(data)
    return df_results


def dataframe_to_excel_bytes(df, sheet_name="AI_Overview"):
    """
    Converte un DataFrame in un file Excel in memoria (BytesIO) pronto per il download.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output


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
        "Nederlands": "nl"
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
        "Paesi Bassi": "nl"
    }


# ==========================
# PAGINE DELL'APP
# ==========================

def page_config():
    st.title("🔑 Configurazione API SerpAPI (AI Overview)")

    st.write(
        """
        Inserisci la tua **API key di SerpAPI**.  
        Verrà usata per le richieste alla funzionalità **AI Overview** di Google.
        """
    )

    api_key = st.text_input(
        "SerpAPI API Key",
        type="password",
        value=st.session_state.get("ai_api_key", "")
    )

    if api_key:
        st.session_state["ai_api_key"] = api_key
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
    st.title("📥 Upload file & Scraping AI Overview")

    api_key = st.session_state.get("ai_api_key")
    if not api_key:
        st.warning("⚠️ Prima imposta la tua SerpAPI API key nella pagina **Configurazione**.")
        return

    hl = st.session_state.get("hl", "it")
    gl = st.session_state.get("gl", "it")

    st.write(
        f"""
        Carica un file Excel con la colonna **Keyword**  
        che contiene le query per cui vuoi estrarre gli **AI Overview references**.

        **Configurazione attuale**
        - Lingua (`hl`): **{hl}**
        - Country (`gl`): **{gl}**
        """
    )

    uploaded_file = st.file_uploader("Carica file Excel (.xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        column_name = st.text_input("Nome colonna che contiene le keyword", value="Keyword")

        sleep_seconds = st.slider(
            "Delay tra le richieste (sec) – per sicurezza verso i limiti API",
            min_value=0.0,
            max_value=5.0,
            value=1.0,
            step=0.5
        )

        if st.button("🚀 Avvia scraping AI Overview"):
            try:
                keywords = read_keywords_from_excel(uploaded_file, column_name=column_name)
                if not keywords:
                    st.warning("Nessuna keyword trovata nella colonna indicata.")
                    return

                st.write(f"Trovate **{len(keywords)}** keyword. Inizio elaborazione…")

                df_results = scrape_ai_overviews_to_dataframe(
                    keywords=keywords,
                    api_key=api_key,
                    sleep_seconds=sleep_seconds,
                    hl=hl,
                    gl=gl
                )

                st.session_state["ai_overview_df"] = df_results

                st.success("Scraping completato. Vai alla pagina **Risultati** per esportare l'Excel.")
                if not df_results.empty:
                    st.dataframe(df_results.head())
                else:
                    st.info("Nessun riferimento trovato per le keyword fornite.")
            except ValueError as e:
                st.error(f"Errore nel file Excel o nel nome colonna: {e}")
            except Exception as e:
                st.error(f"Errore durante lo scraping: {e}")
    else:
        st.info("Carica un file Excel per procedere.")


def page_results():
    st.title("📊 Risultati & Download AI Overview")

    df = st.session_state.get("ai_overview_df")

    if df is None:
        st.info("Nessun risultato disponibile. Esegui prima lo scraping nella pagina **Upload & Scraping**.")
        return

    st.write("Anteprima risultati:")
    st.dataframe(df)

    default_filename = "AI_Overview_export.xlsx"
    filename = st.text_input("Nome file di output", value=default_filename)

    excel_bytes = dataframe_to_excel_bytes(df, sheet_name="AI_Overview")

    st.download_button(
        label="💾 Scarica Excel con AI Overview references",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==========================
# MAIN: NAVIGAZIONE MULTIPAGINA
# ==========================

def main():
    st.set_page_config(
        page_title="AI Overview Scraper - SerpAPI",
        page_icon="🤖",
        layout="wide"
    )

    # valori di default sessione
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