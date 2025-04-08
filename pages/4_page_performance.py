import streamlit as st
import requests
import pandas as pd

# Titolo della Web App
st.set_page_config(page_title="Confronto PageSpeed Insights", layout="wide")
st.title("🚀 Confronto PageSpeed Insights e Core Web Vitals")

# Inserimento della chiave API di Google PageSpeed Insights
st.subheader("Step 1: Inserisci la tua API Key di Google PageSpeed Insights")
google_api_key = st.text_input("Google API Key", type="password")

# Inserimento degli URL da analizzare
st.subheader("Step 2: Inserisci gli URL da analizzare")
urls = st.text_area("Inserisci gli URL (uno per riga)")

# Funzione per determinare lo stile delle celle
def style_metric(value, good_threshold, needs_improvement_threshold, higher_is_better=False):
    if value == "N/A":
        color = 'white'
    else:
        if higher_is_better:
            if value >= good_threshold:
                color = 'lightgreen'
            elif value >= needs_improvement_threshold:
                color = 'yellow'
            else:
                color = 'lightcoral'
        else:
            if value <= good_threshold:
                color = 'lightgreen'
            elif value <= needs_improvement_threshold:
                color = 'yellow'
            else:
                color = 'lightcoral'
    return f'background-color: {color}'

# Pulsante di avvio analisi
if st.button("🔍 Analizza le Pagine"):
    url_list = [url.strip() for url in urls.split("\n") if url.strip()]
    
    if len(url_list) < 1:
        st.error("⚠️ Inserisci almeno un URL per effettuare l'analisi.")
    elif not google_api_key:
        st.error("⚠️ Inserisci la tua Google API Key.")
    else:
        st.write("🔄 Estrazione dati PageSpeed Insights in corso...")
        
        page_data = []
        
        for url in url_list:
            try:
                # Costruzione URL API
                api_url = f"https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url={url}&key={google_api_key}&strategy=mobile"
                response = requests.get(api_url)
                data = response.json()
                
                # Estrazione metriche Core Web Vitals
                metrics = data.get("loadingExperience", {}).get("metrics", {})
                
                lcp = metrics.get("LARGEST_CONTENTFUL_PAINT_MS", {}).get("percentile", "N/A")
                fcp = metrics.get("FIRST_CONTENTFUL_PAINT_MS", {}).get("percentile", "N/A")
                cls = metrics.get("CUMULATIVE_LAYOUT_SHIFT_SCORE", {}).get("percentile", "N/A")
                fid = metrics.get("FIRST_INPUT_DELAY_MS", {}).get("percentile", "N/A")
                ttfb = metrics.get("EXPERIMENTAL_TIME_TO_FIRST_BYTE", {}).get("percentile", "N/A")
                
                # Estrazione del Performance Score generale
                performance_score = data.get("lighthouseResult", {}).get("categories", {}).get("performance", {}).get("score", "N/A")
                if performance_score != "N/A":
                    performance_score = performance_score * 100  # Convertiamo il punteggio in scala 0-100
                
                # Conversione dei valori in secondi o mantenimento di "N/A"
                lcp = lcp / 1000 if lcp != "N/A" else lcp
                fcp = fcp / 1000 if fcp != "N/A" else fcp
                fid = fid / 1000 if fid != "N/A" else fid
                ttfb = ttfb / 1000 if ttfb != "N/A" else ttfb
                
                page_data.append({
                    "URL": url,
                    "LCP (s)": lcp,
                    "FCP (s)": fcp,
                    "CLS": cls,
                    "FID (s)": fid,
                    "TTFB": ttfb,
                    "Performance Score": performance_score
                })
                
            except Exception as e:
                st.error(f"❌ Errore nell'analisi di {url}: {e}")
        
        # Creazione DataFrame con i risultati
        df = pd.DataFrame(page_data)
        
        # Applicazione degli stili condizionali
        def apply_styles(df):
            df_styles = df.style.applymap(
                lambda x: style_metric(x, 2.5, 4.0) if isinstance(x, (int, float)) else '',
                subset=['LCP (s)']
            ).applymap(
                lambda x: style_metric(x, 1.8, 3.0) if isinstance(x, (int, float)) else '',
                subset=['FCP (s)']
            ).applymap(
                lambda x: style_metric(x, 0.1, 0.25) if isinstance(x, (int, float)) else '',
                subset=['CLS']
            ).applymap(
                lambda x: style_metric(x, 0.1, 0.3) if isinstance(x, (int, float)) else '',
                subset=['FID (s)']
            ).applymap(
                lambda x: style_metric(x, 0.1, 0.3) if isinstance(x, (int, float)) else '',
                subset=['TTFB']
            ).applymap(
                lambda x: style_metric(x, 90, 50, higher_is_better=True) if isinstance(x, (int, float)) else '',
                subset=['Performance Score']
            )
            return df_styles
        
        # Visualizzazione dei risultati
        st.write("📊 **Risultati dell'analisi PageSpeed Insights:**")
        st.dataframe(apply_styles(df))
        
        st.write("✅ Analisi completata con successo! 🚀")