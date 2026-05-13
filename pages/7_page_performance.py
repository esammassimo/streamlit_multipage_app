import streamlit as st
import requests
import pandas as pd
import concurrent.futures
import threading
import time
import io

st.set_page_config(page_title="Confronto PageSpeed Insights", layout="wide")
st.title("🚀 Confronto PageSpeed Insights e Core Web Vitals")

# Step 1: API Key
st.subheader("Step 1: Inserisci la tua API Key di Google PageSpeed Insights")
google_api_key = st.text_input("Google API Key", type="password")

# Step 2: Strategia
st.subheader("Step 2: Scegli il tipo di analisi")
strategy = st.radio("Strategia PageSpeed", ["mobile", "desktop"], index=0, horizontal=True)

# Step 3: URL
st.subheader("Step 3: Inserisci gli URL da analizzare")
urls = st.text_area("Inserisci gli URL (uno per riga)")

# Step 4: Parallelizzazione
st.subheader("Step 4: Impostazioni di parallelizzazione")
col1, col2 = st.columns(2)
with col1:
    max_workers = st.slider("Worker paralleli", min_value=1, max_value=5, value=2)
with col2:
    delay_seconds = st.slider("Delay tra richieste (secondi)", min_value=0.5, max_value=5.0, value=1.0, step=0.5)

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

def style_stato(val):
    if val.startswith("✅"):
        return "background-color: lightgreen"
    if val.startswith("❌"):
        return "background-color: lightcoral"
    return ""

def extract_metric_ms(metrics_dict, primary_key, fallback_key=None):
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
    if "N/A" in (lcp_s, cls_s, inp_s):
        return "N/A"
    if lcp_s == "Good" and cls_s == "Good" and inp_s == "Good":
        return "PASS"
    if "Poor" in (lcp_s, cls_s, inp_s):
        return "FAIL"
    return "NEEDS IMPROVEMENT"

def apply_styles(df):
    return (
        df.style
        .map(lambda x: style_metric(x, 2.5, 4.0) if isinstance(x, (int, float)) else '', subset=['LCP (s)'])
        .map(lambda x: style_metric(x, 1.8, 3.0) if isinstance(x, (int, float)) else '', subset=['FCP (s)'])
        .map(lambda x: style_metric(x, 0.1, 0.25) if isinstance(x, (int, float)) else '', subset=['CLS'])
        .map(lambda x: style_metric(x, 0.2, 0.5) if isinstance(x, (int, float)) else '', subset=['INP (s)'])
        .map(lambda x: style_metric(x, 0.8, 1.8) if isinstance(x, (int, float)) else '', subset=['TTFB (s)'])
        .map(lambda x: style_metric(x, 90, 50, higher_is_better=True) if isinstance(x, (int, float)) else '', subset=['Performance Score'])
        .map(style_cwv_status, subset=['CWV Status'])
    )

def fetch_url_data(url, api_key, strat, rate_lock, last_request_time, delay):
    """Worker: recupera i dati PageSpeed con rate limiting condiviso."""
    with rate_lock:
        now = time.time()
        elapsed = now - last_request_time[0]
        if elapsed < delay:
            time.sleep(delay - elapsed)
        last_request_time[0] = time.time()

    try:
        response = requests.get(
            "https://www.googleapis.com/pagespeedonline/v5/runPagespeed",
            params={"url": url, "key": api_key, "strategy": strat},
            timeout=60,
        )
        response.raise_for_status()
        data = response.json()

        metrics = (data.get("loadingExperience", {}) or {}).get("metrics", {}) or {}

        lcp_ms  = extract_metric_ms(metrics, "LARGEST_CONTENTFUL_PAINT_MS")
        fcp_ms  = extract_metric_ms(metrics, "FIRST_CONTENTFUL_PAINT_MS")
        cls     = extract_metric_ms(metrics, "CUMULATIVE_LAYOUT_SHIFT_SCORE")
        inp_ms  = extract_metric_ms(metrics, "INTERACTION_TO_NEXT_PAINT", "EXPERIMENTAL_INTERACTION_TO_NEXT_PAINT")
        ttfb_ms = extract_metric_ms(metrics, "EXPERIMENTAL_TIME_TO_FIRST_BYTE")

        performance_score = (
            data.get("lighthouseResult", {})
                .get("categories", {})
                .get("performance", {})
                .get("score", "N/A")
        )
        if performance_score != "N/A":
            performance_score = round(performance_score * 100, 1)

        lcp  = lcp_ms  / 1000 if lcp_ms  != "N/A" else lcp_ms
        fcp  = fcp_ms  / 1000 if fcp_ms  != "N/A" else fcp_ms
        inp  = inp_ms  / 1000 if inp_ms  != "N/A" else inp_ms
        ttfb = ttfb_ms / 1000 if ttfb_ms != "N/A" else ttfb_ms

        lcp_status = status_from_thresholds(lcp,  2.5,  4.0)
        cls_status = status_from_thresholds(cls,  0.1,  0.25)
        inp_status = status_from_thresholds(inp,  0.2,  0.5)
        cwv_status = cwv_overall(lcp_status, cls_status, inp_status)

        return {
            "status": "ok",
            "url": url,
            "row": {
                "URL": url,
                "Strategia": strat.capitalize(),
                "LCP (s)": lcp,
                "FCP (s)": fcp,
                "CLS": cls,
                "INP (s)": inp,
                "TTFB (s)": ttfb,
                "Performance Score": performance_score,
                "CWV Status": cwv_status,
            },
        }
    except Exception as e:
        return {"status": "error", "url": url, "error": str(e)}

# ---------- Run ----------
if st.button("🔍 Analizza le Pagine"):
    url_list = [u.strip() for u in urls.split("\n") if u.strip()]

    if not url_list:
        st.error("⚠️ Inserisci almeno un URL per effettuare l'analisi.")
    elif not google_api_key:
        st.error("⚠️ Inserisci la tua Google API Key.")
    else:
        st.info(
            f"🔄 Analisi avviata con **{max_workers}** worker(s), "
            f"delay **{delay_seconds}s**, strategia **{strategy}**."
        )

        rate_lock = threading.Lock()
        last_request_time = [0.0]

        # ---- UI live ----
        progress_bar    = st.progress(0.0, text="Avvio...")
        st.markdown("**Anteprima richieste in tempo reale:**")
        preview_ph      = st.empty()

        url_statuses = {url: "⏳ In coda" for url in url_list}
        page_data    = []

        def render_preview():
            rows = [{"URL": u, "Stato": s} for u, s in url_statuses.items()]
            preview_ph.dataframe(
                pd.DataFrame(rows).style.map(style_stato, subset=["Stato"]),
                use_container_width=True,
                hide_index=True,
            )

        render_preview()

        # ---- Esecuzione parallela ----
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(
                    fetch_url_data, url, google_api_key, strategy,
                    rate_lock, last_request_time, delay_seconds
                ): url
                for url in url_list
            }

            completed = 0
            for future in concurrent.futures.as_completed(futures):
                url    = futures[future]
                result = future.result()
                completed += 1

                if result["status"] == "ok":
                    url_statuses[url] = "✅ Completato"
                    page_data.append(result["row"])
                else:
                    url_statuses[url] = f"❌ Errore: {result['error']}"

                progress_bar.progress(
                    completed / len(url_list),
                    text=f"Completati {completed} / {len(url_list)} URL",
                )
                render_preview()

        progress_bar.progress(1.0, text="✅ Analisi completata!")

        # ---- Errori ----
        for url, s in url_statuses.items():
            if s.startswith("❌"):
                st.error(f"{s}  —  {url}")

        # ---- Tabella risultati ----
        df = pd.DataFrame(page_data)

        if df.empty:
            st.warning("Nessun dato raccolto. Controlla gli URL e riprova.")
        else:
            st.markdown("---")
            st.write("📊 **Risultati dell'analisi PageSpeed Insights:**")
            st.dataframe(apply_styles(df), use_container_width=True)

            counts = df["CWV Status"].value_counts(dropna=False).to_dict()
            st.caption(
                f"✅ PASS: {counts.get('PASS', 0)}  •  "
                f"⚠️ NEEDS IMPROVEMENT: {counts.get('NEEDS IMPROVEMENT', 0)}  •  "
                f"❌ FAIL: {counts.get('FAIL', 0)}  •  "
                f"N/A: {counts.get('N/A', 0)}"
            )

            # ---- Download ----
            st.markdown("---")
            dl_col1, dl_col2 = st.columns(2)

            csv_bytes = df.to_csv(index=False).encode("utf-8")
            dl_col1.download_button(
                label="⬇️ Download CSV",
                data=csv_bytes,
                file_name="pagespeed_results.csv",
                mime="text/csv",
            )

            excel_buf = io.BytesIO()
            with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="PageSpeed")
            dl_col2.download_button(
                label="⬇️ Download Excel",
                data=excel_buf.getvalue(),
                file_name="pagespeed_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )