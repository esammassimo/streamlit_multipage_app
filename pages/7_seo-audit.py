
# SEO Audit Tool – v3 (contributi pesati per riga nel benchmark)
# Avvio: streamlit run seo_audit_tool_v3.py

import io
import os
import numpy as np
import pandas as pd
import streamlit as st
from urllib.parse import urlparse

def setup_matplotlib():
    try:
        import matplotlib.pyplot as plt
        import numpy as np  # noqa: F401
        return True, plt
    except ModuleNotFoundError:
        st.warning("Modulo 'matplotlib' mancante. pip install matplotlib")
        return False, None

MATPLOTLIB_OK, plt = setup_matplotlib()

def estrai_dominio(df: pd.DataFrame):
    try:
        url_sample = df['Address'].dropna().iloc[0]
        dominio = urlparse(url_sample).netloc.replace('www.', '')
        return dominio if dominio else None
    except Exception:
        return None

def first_present_numeric(df: pd.DataFrame, *names: str):
    for n in names:
        if n in df.columns:
            s = pd.to_numeric(df[n], errors='coerce')
            if s.notna().any():
                return float(round(s.mean(), 2))
    return None

def winsorize_series(s: pd.Series, lower_q=0.01, upper_q=0.99):
    if s is None or len(s) == 0:
        return s
    lo, hi = s.quantile(lower_q), s.quantile(upper_q)
    return s.clip(lower=lo, upper=hi)

def smooth_proportion(p: float, n: float, prior_mu: float, prior_n: float = 200.0) -> float:
    if pd.isna(p) or n <= 0:
        return prior_mu
    return float((p*n + prior_mu*prior_n) / (n + prior_n))

def redistribute_component_weights(component_mask: dict, base_weights: dict) -> dict:
    avail = {k: w for k, w in base_weights.items() if component_mask.get(k, False)}
    if not avail:
        return {"status": 1.0}
    s = sum(avail.values())
    return {k: w/s for k, w in avail.items()}

def penalties_to_score(penalties: dict, weights: dict) -> float:
    tot = 0.0
    for k, w in weights.items():
        tot += w * penalties.get(k, 0.0)
    return max(0.0, 100.0 * (1.0 - tot))

COMPONENT_MAP = {
    "status": "Penalità Status Code %",
    "canonical": "Penalità Canonical %",
    "html": "Penalità Tag HTML %",
    "duplicate": "Penalità Contenuti %",
    "cwv": "Penalità CWV %",
}
BASE_WEIGHTS = {"status": 0.30, "canonical": 0.15, "html": 0.20, "duplicate": 0.10, "cwv": 0.20}

def calcola_score(df: pd.DataFrame, kpi: dict, contenuti_penalty: float | None = None):
    pagine_totali = df.shape[0]
    if pagine_totali == 0:
        return 0.0, {}

    penalita = {}
    penalita['Penalità Status Code %'] = (
        (kpi.get('Pagine 3xx', 0) + kpi.get('Pagine 4xx', 0) + kpi.get('Bloccate da Robots.txt', 0)) / max(pagine_totali, 1)
    )
    canonical_non_self = 0.0
    if 'Canonical Link Element 1 Resolved' in df.columns:
        canon_df = df[['Address', 'Canonical Link Element 1 Resolved']].dropna()
        if len(canon_df) > 0:
            canonical_non_self = (canon_df['Address'] != canon_df['Canonical Link Element 1 Resolved']).sum() / max(pagine_totali, 1)
    penalita['Penalità Canonical %'] = canonical_non_self
    penalita['Penalità Tag HTML %'] = (
        (kpi.get('Title Duplicati', 0) + kpi.get('Title Mancanti', 0) +
         kpi.get('Meta Description Duplicati', 0) + kpi.get('Meta Description Mancanti', 0) +
         kpi.get('H1 Duplicati', 0) + kpi.get('H1 Mancanti', 0)) / (3 * max(pagine_totali, 1))
    )
    if contenuti_penalty is None:
        contenuti_penalty = 0.0
    penalita['Penalità Contenuti %'] = float(min(1.0, max(0.0, contenuti_penalty)))

    penalita_cwv = []
    soglie = {'LCP': 2500, 'INP': 200, 'CLS': 0.1, 'FCP': 1800, 'TTFB': 800}
    for metrica, soglia in soglie.items():
        if metrica in kpi:
            val = kpi[metrica]
            if isinstance(val, (int, float)) and val > 0:
                if metrica == 'CLS':
                    penalita_cwv.append(min(1.0, val / soglia))
                else:
                    penalita_cwv.append(min(1.0, max(0.0, (val - soglia) / soglia)))
    base_cwv_pen = sum(penalita_cwv) / len(penalita_cwv) if penalita_cwv else 0.0

    pass_rate = kpi.get('CWV Assessment Pass %', None)
    if isinstance(pass_rate, (int, float)):
        if pass_rate > 60:
            base_cwv_pen *= 0.85
        elif pass_rate > 50:
            base_cwv_pen *= 0.95
    penalita['Penalità CWV %'] = base_cwv_pen

    score = 100 * (1 - (
        0.30 * penalita['Penalità Status Code %'] +
        0.15 * penalita['Penalità Canonical %'] +
        0.20 * penalita['Penalità Tag HTML %'] +
        0.10 * penalita['Penalità Contenuti %'] +
        0.20 * penalita['Penalità CWV %']
    ))
    return round(max(score, 0), 2), {k: round(v * 100, 2) for k, v in penalita.items()}

def estrai_kpi(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.str.strip()
    df['Status Code'] = pd.to_numeric(df.get('Status Code'), errors='coerce')
    status = {
        'Pagine 2xx': df['Status Code'].between(200, 299).sum(),
        'Pagine 3xx': df['Status Code'].between(300, 399).sum(),
        'Pagine 4xx': df['Status Code'].between(400, 499).sum(),
        'Bloccate da Robots.txt': df.get('Indexability', pd.Series(index=df.index)).astype(str).str.contains("Blocked by Robots", na=False).sum(),
        'Pagine HTML Totali': df.shape[0]
    }

    def analizza(col):
        if col not in df.columns:
            return (0, 0, 0)
        valid = df[col].dropna()
        return (len(valid[valid.duplicated(keep=False)].unique()), df[col].isna().sum(), df[col].notna().sum())

    title = analizza("Title 1")
    description = analizza("Meta Description 1")
    h1 = analizza("H1-1")
    html_tag = {
        'Title Duplicati': title[0], 'Title Mancanti': title[1], 'Totale Title': title[2],
        'Meta Description Duplicati': description[0], 'Meta Description Mancanti': description[1], 'Totale Meta Description': description[2],
        'H1 Duplicati': h1[0], 'H1 Mancanti': h1[1], 'Totale H1': h1[2]
    }

    pagine_totali = max(df.shape[0], 1)
    dup_pages = 0
    if 'Title 1' in df.columns:
        vc = df['Title 1'].dropna().value_counts()
        dup_pages = int(vc[vc > 1].sum())
    dup_rate = dup_pages / pagine_totali

    thin_rate = 0.0
    if 'Word Count' in df.columns:
        wc = pd.to_numeric(df['Word Count'], errors='coerce')
        wc_valid = wc.dropna()
        if len(wc_valid) > 0:
            media_wc = wc_valid.mean()
            thin_pages = (wc_valid < media_wc).sum()
            thin_rate = thin_pages / len(wc_valid)

    contenuti_penalty = (dup_rate + thin_rate) / 2.0
    content = {'Pagine Duplicate (da Title)': dup_pages, 'Thin Content %': round(thin_rate * 100, 2), 'Pagine Totali': pagine_totali}

    img_pages_pct = None
    if 'Images Missing Alt Text' in df.columns:
        missing_any = pd.to_numeric(df['Images Missing Alt Text'], errors='coerce').fillna(0) > 0
        img_pages_pct = round((missing_any.sum() / pagine_totali) * 100, 2)

    cwv_vals = {
        'FCP': first_present_numeric(df, 'First contentfull Paint Time (ms)', 'FCP (ms)', 'FCP'),
        'LCP': first_present_numeric(df, 'Largest contentfull Paint Time (ms)', 'LCP (ms)', 'LCP'),
        'CLS': first_present_numeric(df, 'Cumulative Layout Shift', 'CLS (ms)', 'CLS'),
        'INP': first_present_numeric(df, 'CrUX Interaction to Next Paint (ms)', 'INP (ms)', 'INP'),
        'TTFB': first_present_numeric(df, 'Server Response Times (TTFB) (ms)', 'TTFB (ms)', 'TTFB'),
    }
    cwv = {k: v for k, v in cwv_vals.items() if v is not None}

    assessment_pass_pct = None
    if 'Core Web Vitals Assessment' in df.columns:
        assess = df['Core Web Vitals Assessment'].astype(str).str.strip()
        total = len(assess)
        passes = (assess.str.lower() == 'pass').sum()
        assessment_pass_pct = round((passes / total) * 100, 2) if total else None

    kpi = {**status, **html_tag, **content, **cwv}
    if assessment_pass_pct is not None: kpi['CWV Assessment Pass %'] = assessment_pass_pct
    if img_pages_pct is not None: kpi['Pagine con Immagini senza ALT %'] = img_pages_pct

    score, penalita = calcola_score(df, kpi, contenuti_penalty=contenuti_penalty)
    kpi['SEO Score'] = score
    kpi.update(penalita)
    return pd.DataFrame([kpi])

def build_stable_domain_scores(df_riepilogo: pd.DataFrame) -> pd.DataFrame:
    df = df_riepilogo.copy()
    needed_cols = ["Pagine Totali", *COMPONENT_MAP.values()]
    for c in needed_cols:
        if c not in df.columns:
            raise ValueError(f"Colonna mancante nel riepilogo: {c}")
    for col in COMPONENT_MAP.values():
        df[col] = winsorize_series(pd.to_numeric(df[col], errors="coerce") / 100.0)
    priors = {}
    for key, col in COMPONENT_MAP.items():
        numer = (df[col] * df["Pagine Totali"].astype(float)).sum()
        denom = (df["Pagine Totali"].astype(float)).sum()
        priors[key] = float(numer / denom) if denom > 0 else float(pd.to_numeric(df[col], errors="coerce").mean())
    shrunk_cols = {}
    for key, col in COMPONENT_MAP.items():
        shrunk_name = f"{col} (shrunk)"
        shrunk_cols[key] = shrunk_name
        df[shrunk_name] = [smooth_proportion(p, n, priors[key]) for p, n in zip(df[col].astype(float), df["Pagine Totali"].astype(float))]
    component_mask = {k: df[COMPONENT_MAP[k]].notna().any() for k in COMPONENT_MAP}
    dyn_w = redistribute_component_weights(component_mask, BASE_WEIGHTS)
    scores, se_est = [], []
    for _, row in df.iterrows():
        penalties = {k: float(row[shrunk_cols[k]]) for k in dyn_w}
        score = penalties_to_score(penalties, dyn_w)
        scores.append(score)
        n = max(float(row["Pagine Totali"]), 1.0)
        var = 0.0
        for k, w in dyn_w.items():
            p = penalties[k]
            var += (w ** 2) * (p * (1 - p) / n)
        se_est.append(np.sqrt(var) * 100.0)
    df["SEO Score Stabile"] = np.round(scores, 2)
    df["Score SE (pt)"] = np.round(se_est, 2)
    df["Score CI95% Min"] = np.round(df["SEO Score Stabile"] - 1.96 * df["Score SE (pt)"], 2)
    df["Score CI95% Max"] = np.round(df["SEO Score Stabile"] + 1.96 * df["Score SE (pt)"], 2)
    df["Stato Stabile"] = pd.cut(df["SEO Score Stabile"], bins=[-0.1, 50, 70, 100.1], labels=["Critico", "Medio", "Buono"])
    return df

def build_portfolio_score(df_stable: pd.DataFrame) -> dict:
    comp_shrunk = {k: f"{v} (shrunk)" for k, v in COMPONENT_MAP.items()}
    component_mask = {k: df_stable.get(comp_shrunk[k]) is not None and df_stable[comp_shrunk[k]].notna().any() for k in comp_shrunk}
    dyn_w = redistribute_component_weights(component_mask, BASE_WEIGHTS)
    N = df_stable["Pagine Totali"].astype(float)
    penalties_port = {}
    for k, col in comp_shrunk.items():
        if k not in dyn_w: continue
        p = (df_stable[col].astype(float) * N).sum() / max(N.sum(), 1.0)
        penalties_port[k] = float(p)
    score_port = penalties_to_score(penalties_port, dyn_w)
    N_tot = max(N.sum(), 1.0)
    var = 0.0
    for k, w in dyn_w.items():
        p = penalties_port[k]
        var += (w ** 2) * (p * (1 - p) / N_tot)
    se = np.sqrt(var) * 100.0
    return {"Portfolio Score": round(score_port, 2), "Portfolio SE (pt)": round(se, 2), "Portfolio CI95%": (round(score_port - 1.96 * se, 2), round(score_port + 1.96 * se, 2)), "Pesi Usati": dyn_w}

# ============ UI ============
st.set_page_config(page_title="SEO Audit Tool v3", layout="wide")
st.title("SEO Audit Tool v3 – contributi pesati")

tab1, tab2 = st.tabs(["Singolo File", "Multi File"])

with tab1:
    file = st.file_uploader("Carica un file .xlsx (Screaming Frog)", type="xlsx", key="single")
    if file:
        xls = pd.ExcelFile(file)
        sheet_name = next((n for n in ['1 - HTML', '1 - All'] if n in xls.sheet_names), xls.sheet_names[0])
        df = xls.parse(sheet_name)
        kpi = estrai_kpi(df)
        st.subheader("Riepilogo SEO")
        st.dataframe(kpi, use_container_width=True)

        st.subheader("Sintesi SEO")
        pesi_df = pd.DataFrame({
            "Componente": ["Status Code e Robots","Canonical non self-ref","HTML Tag duplicati/mancanti","Contenuti (duplicati da Title + thin)","Core Web Vitals (con bonus Assessment)"],
            "Peso (%)": [30, 15, 20, 10, 20],
            "Risultato (Penalità %)": [
                kpi['Penalità Status Code %'].iloc[0],
                kpi['Penalità Canonical %'].iloc[0],
                kpi['Penalità Tag HTML %'].iloc[0],
                kpi['Penalità Contenuti %'].iloc[0],
                kpi['Penalità CWV %'].iloc[0]
            ]
        })
        st.dataframe(pesi_df, use_container_width=True)

        cols = st.columns(4)
        cols[0].metric("SEO Score", f"{kpi['SEO Score'].iloc[0]}")
        if 'CWV Assessment Pass %' in kpi.columns:
            cols[1].metric("CWV Assessment Pass %", f"{kpi['CWV Assessment Pass %'].iloc[0]}%")
        if 'Pagine con Immagini senza ALT %' in kpi.columns:
            cols[2].metric("Pagine con ≥1 immagine senza ALT", f"{kpi['Pagine con Immagini senza ALT %'].iloc[0]}%")
        if 'TTFB' in kpi.columns:
            cols[3].metric("TTFB medio", f"{kpi['TTFB'].iloc[0]} ms")

with tab2:
    files = st.file_uploader("Carica più file .xlsx (Screaming Frog)", type="xlsx", accept_multiple_files=True, key="multi")
    if files:
        risultati = []
        for f in files:
            xls = pd.ExcelFile(f)
            sheet_name = next((n for n in ['1 - HTML', '1 - All'] if n in xls.sheet_names), xls.sheet_names[0])
            df = xls.parse(sheet_name)
            dominio = estrai_dominio(df) or os.path.splitext(f.name)[0].split("_")[0]
            kpi = estrai_kpi(df); kpi.insert(0, 'Dominio', dominio)
            risultati.append(kpi)
        if risultati:
            df_riepilogo = pd.concat(risultati, ignore_index=True)
            if "Pagine Totali" not in df_riepilogo.columns and "Pagine HTML Totali" in df_riepilogo.columns:
                df_riepilogo = df_riepilogo.rename(columns={"Pagine HTML Totali": "Pagine Totali"})

            st.subheader("Riepilogo SEO per tutti i domini")
            st.dataframe(df_riepilogo, use_container_width=True)

            # BENCHMARK con contributi pesati
            st.subheader("Benchmark competitivo (valori e contributi pesati)")

            # Sidebar pesi default (usati solo se DR/LV incompleti)
            with st.sidebar:
                st.markdown("### Pesi SCORE (default – normalizzati)")
                w_indexability = st.slider("Peso INDEXABILITY", 0, 100, 20)
                w_cwv          = st.slider("Peso CWV", 0, 100, 20)
                w_html         = st.slider("Peso HTML TAGS", 0, 100, 20)
                w_content      = st.slider("Peso CONTENT EVALUATION", 0, 100, 20)
                w_dr           = st.slider("Peso DR", 0, 100, 10)
                w_lv           = st.slider("Peso LINK VELOCITY", 0, 100, 10)

            weights_default_raw = {"INDEXABILITY": w_indexability,"CWV": w_cwv,"HTML TAGS": w_html,"CONTENT EVALUATION": w_content,"DR": w_dr,"LINK VELOCITY": w_lv}
            wsum = sum(weights_default_raw.values()) or 1
            weights_default = {k: v / wsum for k, v in weights_default_raw.items()}

            # Form per l’input manuale con submit
            if "manual_ext_df" not in st.session_state:
                st.session_state.manual_ext_df = pd.DataFrame({"Dominio": df_riepilogo["Dominio"].tolist(), "DR": [None]*len(df_riepilogo), "LINK VELOCITY":[None]*len(df_riepilogo)})
                st.session_state.offpage_scale = "0-100"

            with st.form("offpage_form", clear_on_submit=False):
                st.markdown("Compila **DR** e **LINK VELOCITY** (se disponi dei dati):")
                manual_input = st.data_editor(st.session_state.manual_ext_df, use_container_width=True, hide_index=True)
                scale = st.selectbox("Scala valori DR/LINK VELOCITY", ["0-100","0-1"], index=0, help="Se 0–1 → *100")
                submitted = st.form_submit_button("Aggiorna benchmark")
            if submitted:
                st.session_state.manual_ext_df = manual_input.copy()
                st.session_state.offpage_scale = scale

            manual_df = st.session_state.manual_ext_df.copy()
            scale = st.session_state.offpage_scale

            def to_num(v):
                try: return float(v)
                except Exception: return np.nan

            all_offpage_ready = manual_df["DR"].apply(lambda x: np.isfinite(to_num(x))).all() and manual_df["LINK VELOCITY"].apply(lambda x: np.isfinite(to_num(x))).all()

            # Pesi finali
            if all_offpage_ready:
                weights = {"INDEXABILITY": 0.33,"HTML TAGS": 0.11,"CONTENT EVALUATION": 0.11,"CWV": 0.11,"DR": 0.165,"LINK VELOCITY": 0.165}
                st.caption("Schema pesi **OFF-PAGE 33%** attivo (DR e Link Velocity compilati per tutti i domini).")
            else:
                weights = weights_default
                st.caption("Schema pesi **default** attivo (mancano valori completi di DR/Link Velocity).")

            rows = ["INDEXABILITY","CWV","HTML TAGS","CONTENT EVALUATION","DR","LINK VELOCITY"]
            domains = df_riepilogo["Dominio"].tolist()
            cols = [d.upper().replace("WWW.", "") for d in domains]

            # Build raw values table
            bench_vals = pd.DataFrame(index=rows, columns=cols, dtype="float")
            manual_map = manual_df.set_index("Dominio").to_dict("index")

            for dom_raw, col in zip(domains, cols):
                row = df_riepilogo.loc[df_riepilogo["Dominio"] == dom_raw].iloc[0]
                indexability = 100.0 - float(row.get("Penalità Status Code %", 0.0))
                cwv = float(row.get("CWV Assessment Pass %", 0.0)) if "CWV Assessment Pass %" in df_riepilogo.columns else 0.0
                html_ok = 100.0 - float(row.get("Penalità Tag HTML %", 0.0))
                content_ok = 100.0 - float(row.get("Penalità Contenuti %", 0.0))
                dr_val = to_num(manual_map.get(dom_raw, {}).get("DR"))
                lv_val = to_num(manual_map.get(dom_raw, {}).get("LINK VELOCITY"))
                if np.isfinite(dr_val) and scale == "0-1": dr_val *= 100.0
                if np.isfinite(lv_val) and scale == "0-1": lv_val *= 100.0
                if np.isfinite(dr_val): dr_val = max(0.0, min(100.0, dr_val))
                if np.isfinite(lv_val): lv_val = max(0.0, min(100.0, lv_val))

                bench_vals.loc["INDEXABILITY", col] = round(indexability, 2)
                bench_vals.loc["CWV", col] = round(cwv, 2)
                bench_vals.loc["HTML TAGS", col] = round(html_ok, 2)
                bench_vals.loc["CONTENT EVALUATION", col] = round(content_ok, 2)
                bench_vals.loc["DR", col] = dr_val if np.isfinite(dr_val) else np.nan
                bench_vals.loc["LINK VELOCITY", col] = lv_val if np.isfinite(lv_val) else np.nan

            # Compute weighted contributions per row (value% * weight)
            bench_pts = bench_vals.copy()
            for r in rows:
                w = weights[r]
                bench_pts.loc[r] = bench_vals.loc[r].astype(float).fillna(0.0) * w

            # SCORE = somma punti per colonna
            score = bench_pts.sum(axis=0).round(0).astype(int)
            bench_vals.loc["SCORE"] = score  # keep for reference
            bench_pts.loc["SCORE"] = score   # identical by definition

            # Build a combined table with MultiIndex columns: (Domain, Value% / Pts)
            tuples = []
            for c in cols:
                tuples.append((c, "Value %"))
                tuples.append((c, "Pts"))
            multi_cols = pd.MultiIndex.from_tuples(tuples, names=["Dominio", "Metric"])
            combined = pd.DataFrame(index=rows + ["SCORE"], columns=multi_cols, dtype="float")

            for c in cols:
                combined[(c, "Value %")] = bench_vals[c]
                combined[(c, "Pts")] = bench_pts[c]

            # Render
            st.dataframe(combined, use_container_width=True)
            st.caption("Ogni riga mostra sia il **valore (%)** sia il **contributo pesato (Pts)**. La riga **SCORE** è la somma dei contributi sopra.")

            # Export
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                bench_vals.to_excel(writer, sheet_name="Benchmark_Values", index=True)
                bench_pts.to_excel(writer, sheet_name="Benchmark_WeightedPts", index=True)
                combined.to_excel(writer, sheet_name="Benchmark_Combined", index=True)
            st.download_button("📥 Scarica benchmark (valori + contributi)", buffer.getvalue(), file_name="benchmark_with_contributions.xlsx")
