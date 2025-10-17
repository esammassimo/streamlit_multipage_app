import streamlit as st
import requests
import pandas as pd

# Titolo della Web App
st.set_page_config(page_title="Confronto PageSpeed Insights", layout="wide")
st.title("🚀 Confronto PageSpeed Insights e Core Web Vitals")

# Step 1: API Key
st.subheader("Step 1: Inserisci la tua API Key di Google PageSpeed Insights")
google_api_key = st.text_input("Google API Key", type="password")

# Step 2: Strategia Mobile/Desktop (default Mobile)
st.subheader("Step 2: Scegli il tipo di analisi")
strategy = st.radio("Strategia PageSpeed", ["mobile", "desktop"], index=0, horizontal=True)

# Step 3: URL
st.subheader("Step 3: Inserisci gli URL da analizzare")
urls = st.text_area("Inserisci gli URL (uno per riga)")

# ---------- Helpers ----------
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

def style_cwv_status(val):
    mapping = {
        "PASS": "lightgreen",
        "NEEDS IMPROVEMENT": "yellow",
        "FAIL": "lightcoral",
        "N/A": "white"
    }
    return f"background-color: {mapping.get(val, 'white')}"

def extract_metric_ms(metrics_dict, primary_key, fallback_key=None):
    """
    Estrae il percentile in ms da 'metrics' di CrUX.
    Se non disponibile, prova la chiave di fallback.
    Ritorna 'N/A' se non trovata.
    """
    if primary_key in metrics_dict:
        return metrics_dict.get(primary_key, {}).get("percentile", "N/A")
    if fallback_key and fallback_key in metrics_dict:
        return metrics_dict.get(fallback_key, {}).get("percentile", "N/A")
    return "N/A"

def status_from_thresholds(val, good_thr, ni_thr, higher_is_better=False):
    if val == "N/A":
        return "N/A"
    try:
        v = float(val)
    except Exception:
        return "N/A"
    if higher_is_better:
        if v >= good_thr:
            return "Good"
        elif v >= ni_thr:
            return "Needs Improvement"
        else:
            return "Poor"
    else:
        if v <= good_thr:
            return "Good"
        elif v <= ni_thr:
            return "Needs Improvement"
        else:
            return "Poor"

def cwv_overall(lcp_s, cls_s, inp_s):
    # Totale: PASS se tutte Good; FAIL se almeno una Poor; altrimenti Needs Improvement; N/A se manca qualcosa
    if "N/A" in (lcp_s, cls_s, inp_s):
        return "N/A"
    if lcp_s == "Good" and cls_s == "Good" and inp_s == "Good":
        return "PASS"
    if "Poor" in (lcp_s, cls_s, inp_s):
        return "FAIL"
    return "NEEDS IMPROVEMENT"

# ---------- Run ----------
if st.button("🔍 Analizza le Pagine"):
    url_list = [url.strip() for url in urls.split("\n") if url.strip()]
    
    if len(url_list) < 1:
        st.error("⚠️ Inserisci almeno un URL per effettuare l'analisi.")
    elif not google_api_key:
        st.error("⚠️ Inserisci la tua Google API Key.")
    else:
        st.write(f"🔄 Estrazione dati PageSpeed Insights in corso... (Strategia: **{strategy}**)")
        page_data = []
        
        for url in url_list:
            try:
                api_url = (
                    "https://www.googleapis.com/pagespeedonline/v5/runPagespeed"
                    f"?url={url}"
                    f"&key={google_api_key}"
                    f"&strategy={strategy}"
                )
                response = requests.get(api_url, timeout=60)
                response.raise_for_status()
                data = response.json()

                # Field data (CrUX)
                metrics = (data.get("loadingExperience", {}) or {}).get("metrics", {}) or {}

                # Metriche in ms (o N/A)
                lcp_ms  = extract_metric_ms(metrics, "LARGEST_CONTENTFUL_PAINT_MS")
                fcp_ms  = extract_metric_ms(metrics, "FIRST_CONTENTFUL_PAINT_MS")
                cls     = extract_metric_ms(metrics, "CUMULATIVE_LAYOUT_SHIFT_SCORE")
                inp_ms  = extract_metric_ms(metrics, "INTERACTION_TO_NEXT_PAINT", "EXPERIMENTAL_INTERACTION_TO_NEXT_PAINT")
                ttfb_ms = extract_metric_ms(metrics, "EXPERIMENTAL_TIME_TO_FIRST_BYTE")

                # Performance Score (0-1 → 0-100)
                performance_score = (
                    data.get("lighthouseResult", {})
                        .get("categories", {})
                        .get("performance", {})
                        .get("score", "N/A")
                )
                if performance_score != "N/A":
                    performance_score = performance_score * 100

                # ms → s dove opportuno
                lcp  = lcp_ms / 1000 if lcp_ms != "N/A" else lcp_ms
                fcp  = fcp_ms / 1000 if fcp_ms != "N/A" else fcp_ms
                inp  = inp_ms / 1000 if inp_ms != "N/A" else inp_ms
                ttfb = ttfb_ms / 1000 if ttfb_ms != "N/A" else ttfb_ms
                # CLS resta com'è

                # Status per il totale CWV (da metriche numeriche)
                lcp_status = status_from_thresholds(lcp, 2.5, 4.0)  # <=2.5 Good, <=4 NI, >4 Poor
                cls_status = status_from_thresholds(cls, 0.1, 0.25) # <=0.1 Good, <=0.25 NI, >0.25 Poor
                inp_status = status_from_thresholds(inp, 0.2, 0.5)  # <=0.2 Good, <=0.5 NI, >0.5 Poor
                cwv_status = cwv_overall(lcp_status, cls_status, inp_status)

                page_data.append({
                    "URL": url,
                    "Strategia": strategy.capitalize(),
                    "LCP (s)": lcp,
                    "FCP (s)": fcp,        # non CWV, ma utile
                    "CLS": cls,
                    "INP (s)": inp,
                    "TTFB (s)": ttfb,
                    "Performance Score": performance_score,
                    "CWV Status": cwv_status  # <<< SOLO risultato totale richiesto
                })

            except Exception as e:
                st.error(f"❌ Errore nell'analisi di {url}: {e}")

        # Tabella
        df = pd.DataFrame(page_data)

        # Stili condizionali
        def apply_styles(df):
            df_styles = (
                df.style
                # metriche numeriche
                .applymap(lambda x: style_metric(x, 2.5, 4.0) if isinstance(x, (int, float)) else '', subset=['LCP (s)'])
                .applymap(lambda x: style_metric(x, 1.8, 3.0) if isinstance(x, (int, float)) else '', subset=['FCP (s)'])
                .applymap(lambda x: style_metric(x, 0.1, 0.25) if isinstance(x, (int, float)) else '', subset=['CLS'])
                .applymap(lambda x: style_metric(x, 0.2, 0.5) if isinstance(x, (int, float)) else '', subset=['INP (s)'])
                .applymap(lambda x: style_metric(x, 0.8, 1.8) if isinstance(x, (int, float)) else '', subset=['TTFB (s)'])
                .applymap(lambda x: style_metric(x, 90, 50, higher_is_better=True) if isinstance(x, (int, float)) else '', subset=['Performance Score'])
                # CWV totale (badge)
                .applymap(style_cwv_status, subset=['CWV Status'])
            )
            return df_styles

        st.write("📊 **Risultati dell'analisi PageSpeed Insights:**")
        st.dataframe(apply_styles(df), use_container_width=True)

        # (Facoltativo) riepilogo contatori PASS/NI/FAIL
        counts = df['CWV Status'].value_counts(dropna=False).to_dict()
        st.caption(f"✅ PASS: {counts.get('PASS', 0)}  •  ⚠️ NEEDS IMPROVEMENT: {counts.get('NEEDS IMPROVEMENT', 0)}  •  ❌ FAIL: {counts.get('FAIL', 0)}  •  N/A: {counts.get('N/A', 0)}")

        st.write("✅ Analisi completata con successo! 🚀")
