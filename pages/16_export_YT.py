import time
import io
import pandas as pd
import streamlit as st
from serpapi import GoogleSearch

# ==========================
# FUNZIONI DI BACKEND
# ==========================

def fetch_youtube_results(query, api_key, hl="it", gl="it"):
    """
    Chiama SerpAPI (engine youtube) e restituisce l'array dei risultati video (se presente).
    """
    params = {
        "engine": "youtube",
        "search_query": query,   # per youtube engine si usa spesso search_query
        "api_key": api_key,
        "hl": hl,
        "gl": gl,
        "no_cache": False
    }
    search = GoogleSearch(params)
    results = search.get_dict()

    # SerpAPI può usare chiavi diverse a seconda del response schema:
    # - video_results
    # - organic_results
    # - results
    return (
        results.get("video_results")
        or results.get("organic_results")
        or results.get("results")
        or []
    )


def read_keywords_from_excel(file_obj, column_name="Keyword"):
    """
    Legge le parole chiave dal file Excel nella colonna specificata.
    file_obj può essere sia un path che un file caricato da Streamlit.
    """
    df_keywords = pd.read_excel(file_obj, usecols=[column_name])
    keywords = df_keywords[column_name].dropna().astype(str).tolist()
    return keywords


def scrape_youtube_to_dataframe(keywords, api_key, sleep_seconds=1.0, hl="it", gl="it", max_results_per_keyword=None):
    """
    Esegue la ricerca YouTube per una lista di keyword e restituisce un DataFrame.
    max_results_per_keyword: se None prende tutti quelli restituiti da SerpAPI, altrimenti limita.
    """
    data = []

    progress_bar = st.progress(0)
    status_text = st.empty()
    total = len(keywords)

    for idx, keyword in enumerate(keywords, start=1):
        status_text.write(f"🔍 Ricerca YouTube per: **{keyword}** ({idx}/{total})")

        try:
            results = fetch_youtube_results(keyword, api_key=api_key, hl=hl, gl=gl)

            if results:
                if max_results_per_keyword is not None:
                    results = results[:max_results_per_keyword]

                for pos, item in enumerate(results, start=1):
                    # Parsing robusto (chiavi comuni in SerpAPI YouTube)
                    title = item.get("title", "")
                    link = item.get("link", "") or item.get("video_link", "")
                    video_id = item.get("video_id", "") or item.get("id", "")

                    channel = (
                        item.get("channel", "")
                        or item.get("author", "")
                        or (item.get("channel", {}) or {}).get("name", "")
                        or (item.get("author", {}) or {}).get("name", "")
                    )

                    views = item.get("views", "") or item.get("view_count", "")
                    published_date = item.get("published_date", "") or item.get("published_time", "")
                    duration = item.get("duration", "") or item.get("length", "")

                    description = item.get("description", "") or item.get("snippet", "")

                    thumbnail = ""
                    thumbs = item.get("thumbnails") or item.get("thumbnail") or item.get("thumbnail_url")
                    if isinstance(thumbs, list) and len(thumbs) > 0:
                        thumbnail = thumbs[0].get("url", "") if isinstance(thumbs[0], dict) else str(thumbs[0])
                    elif isinstance(thumbs, dict):
                        thumbnail = thumbs.get("url", "")
                    elif isinstance(thumbs, str):
                        thumbnail = thumbs

                    data.append({
                        "keyword": keyword,
                        "position": item.get("position", pos),
                        "title": title,
                        "link": link,
                        "video_id": video_id,
                        "channel": channel,
                        "views": views,
                        "published_date": published_date,
                        "duration": duration,
                        "description": description,
                        "thumbnail": thumbnail
                    })

                st.write(f"✅ {len(results)} risultati trovati per **{keyword}**.")
            else:
                st.write(f"⚠️ Nessun risultato trovato per **{keyword}**.")

        except Exception as e:
            st.error(f"❌ Errore durante la ricerca per '{keyword}': {e}")

        progress_bar.progress(idx / total)
        time.sleep(sleep_seconds)

    status_text.write("✅ Elaborazione completata.")
    df_results = pd.DataFrame(data)
    return df_results


def dataframe_to_excel_bytes(df, sheet_name="YouTube_Search"):
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
    st.title("🔑 Configurazione API SerpAPI (YouTube Search)")

    st.write(
        """
        Inserisci la tua **API key di SerpAPI**.  
        Verrà usata per le richieste al motore **YouTube**.
        """
    )

    api_key = st.text_input(
        "SerpAPI API Key",
        type="password",
        value=st.session_state.get("yt_api_key", "")
    )

    if api_key:
        st.session_state["yt_api_key"] = api_key
        st.success("API key salvata in sessione.")
    else:
        st.info("Inserisci la tua API key per iniziare.")


def page_scraping():
    st.title("📥 Upload file & Scraping YouTube Search Results")

    api_key = st.session_state.get("yt_api_key")
    if not api_key:
        st.warning("⚠️ Prima imposta la tua SerpAPI API key nella pagina **Configurazione**.")
        return

    st.write(
        """
        Carica un file Excel con la colonna **Keyword**  
        che contiene le query per cui vuoi estrarre i **risultati di ricerca YouTube**.
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

        max_results = st.number_input(
            "Max risultati per keyword (0 = tutti quelli restituiti da SerpAPI)",
            min_value=0,
            max_value=100,
            value=0,
            step=5
        )
        max_results_per_keyword = None if max_results == 0 else int(max_results)

        hl = st.text_input("hl (lingua interfaccia)", value="it")
        gl = st.text_input("gl (paese)", value="it")

        if st.button("🚀 Avvia scraping YouTube"):
            try:
                keywords = read_keywords_from_excel(uploaded_file, column_name=column_name)
                if not keywords:
                    st.warning("Nessuna keyword trovata nella colonna indicata.")
                    return

                st.write(f"Trovate **{len(keywords)}** keyword. Inizio elaborazione…")

                df_results = scrape_youtube_to_dataframe(
                    keywords=keywords,
                    api_key=api_key,
                    sleep_seconds=sleep_seconds,
                    hl=hl,
                    gl=gl,
                    max_results_per_keyword=max_results_per_keyword
                )

                st.session_state["youtube_df"] = df_results

                st.success("Scraping completato. Vai alla pagina **Risultati** per esportare l'Excel.")
                if not df_results.empty:
                    st.dataframe(df_results.head())
                else:
                    st.info("Nessun risultato trovato per le keyword fornite.")
            except Exception as e:
                st.error(f"Errore durante lo scraping: {e}")
    else:
        st.info("Carica un file Excel per procedere.")


def page_results():
    st.title("📊 Risultati & Download YouTube Search")

    df = st.session_state.get("youtube_df")

    if df is None:
        st.info("Nessun risultato disponibile. Esegui prima lo scraping nella pagina **Upload & Scraping**.")
        return

    st.write("Anteprima risultati:")
    st.dataframe(df)

    default_filename = "YouTube_search_export.xlsx"
    filename = st.text_input("Nome file di output", value=default_filename)

    excel_bytes = dataframe_to_excel_bytes(df, sheet_name="YouTube_Search")

    st.download_button(
        label="💾 Scarica Excel con risultati YouTube",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==========================
# MAIN: NAVIGAZIONE MULTIPAGINA
# ==========================

def main():
    st.set_page_config(
        page_title="YouTube Search Scraper - SerpAPI",
        page_icon="📺",
        layout="wide"
    )

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