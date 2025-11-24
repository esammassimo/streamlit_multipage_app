import time
import io
import requests
import pandas as pd
import streamlit as st

API_URL = "https://serpapi.com/search.json"


# ==========================
# FUNZIONI DI BACKEND
# ==========================

def get_top4_paa_with_meta_and_indicators(keyword, api_key, gl="it", hl="it"):
    """
    Chiama SerpAPI e restituisce le prime 4 PAA con title/link/snippet e flag booleani.
    """
    params = {
        "engine": "google",
        "q": keyword,
        "api_key": api_key,
        "gl": gl,
        "hl": hl
    }

    resp = requests.get(API_URL, params=params)
    if resp.status_code != 200:
        # Ritorno lista vuota in caso di errore
        st.warning(f"❌ Errore per '{keyword}' – status {resp.status_code}: {resp.text}")
        return []

    data = resp.json()
    rq = data.get("related_questions", [])  # contiene anche title, link, snippet

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

    return results


def read_keywords(file_obj, sheet_name):
    """
    Legge le keyword dalla prima colonna del foglio specificato.
    file_obj può essere sia un path sia un file caricato da Streamlit.
    """
    df = pd.read_excel(file_obj, sheet_name=sheet_name, engine="openpyxl")
    return df.iloc[:, 0].dropna().astype(str).tolist()


def scrape_keywords_to_dataframe(keywords, api_key, sleep_seconds=1):
    """
    Esegue lo scraping delle PAA per una lista di keyword e restituisce un DataFrame.
    """
    rows = []
    progress_bar = st.progress(0)
    status_text = st.empty()

    total = len(keywords)

    for idx, kw in enumerate(keywords, start=1):
        status_text.write(f"🔍 Elaboro keyword: **{kw}** ({idx}/{total})")

        items = get_top4_paa_with_meta_and_indicators(kw, api_key=api_key)
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

        # Aggiorna barra di progresso
        progress_bar.progress(idx / total)

        # Rate limit minimale
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


def page_scraping():
    st.title("📥 Upload file & Scraping PAA")

    # Controllo che l'API key sia stata impostata
    api_key = st.session_state.get("api_key")
    if not api_key:
        st.warning("⚠️ Prima imposta la tua SerpAPI API key nella pagina **Configurazione**.")
        return

    st.write(
        """
        Carica un file Excel con le keyword nella **prima colonna** del foglio selezionato.
        Poi avvia lo scraping delle People Also Ask (PAA).
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
                    sleep_seconds=sleep_seconds
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