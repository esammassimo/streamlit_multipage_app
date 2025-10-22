
# SEO Audit Tool – versione completa (Singolo/Multi File) con benchmark competitivo e pesi off-page
# Requisiti: streamlit, pandas, openpyxl, matplotlib, numpy
# Avvio: streamlit run seo_audit_tool.py

import io
import os
import numpy as np
import pandas as pd
import streamlit as st
from urllib.parse import urlparse

# =========================
# Setup Matplotlib (radar)
# =========================
def setup_matplotlib():
    try:
        import matplotlib.pyplot as plt
        import numpy as np  # noqa: F401
        return True, plt
    except ModuleNotFoundError:
        st.warning("Modulo 'matplotlib' mancante. Per visualizzare i radar chart, esegui: pip install matplotlib")
        return False, None

MATPLOTLIB_OK, plt = setup_matplotlib()

# =========================
# Helper comuni
# =========================
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

# =========================
# Scoring – funzioni robuste
# =========================
def winsorize_series(s: pd.Series, lower_q=0.01, upper_q=0.99):
    if s is None or len(s) == 0:
        return s
    lo, hi = s.quantile(lower_q), s.quantile(upper_q)
    return s.clip(lower=lo, upper=hi)

def smooth_proportion(p: float, n: float, prior_mu: float, prior_n: float = 200.0) -> float:
    """
    Smoothing bayesiano:
    p_s = (p*n + prior_mu*prior_n) / (n + prior_n)
    """
    if pd.isna(p) or n <= 0:
        return prior_mu
    return float((p*n + prior_mu*prior_n) / (n + prior_n))

def redistribute_component_weights(component_mask: dict, base_weights: dict) -> dict:
    """Se mancano componenti, ridistribuisci i pesi rimanenti in modo proporzionale."""
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

# Componenti/pesi (usati anche per lo score stabile multi-file)
COMPONENT_MAP = {
    "status": "Penalità Status Code %",
    "canonical": "Penalità Canonical %",
    "html": "Penalità Tag HTML %",
    "duplicate": "Penalità Contenuti %",
    "cwv": "Penalità CWV %",
}
BASE_WEIGHTS = {
    "status": 0.30,
    "canonical": 0.15,
    "html": 0.20,
    "duplicate": 0.10,
    "cwv": 0.20,
}

# =========================
# Estrattori KPI + Score base
# =========================
def calcola_score(df: pd.DataFrame, kpi: dict, contenuti_penalty: float | None = None):
    pagine_totali = df.shape[0]
    if pagine_totali == 0:
        return 0.0, {}

    penalita = {}

    # 1) Status + Robots
    penalita['Penalità Status Code %'] = (
        (kpi.get('Pagine 3xx', 0) + kpi.get('Pagine 4xx', 0) + kpi.get('Bloccate da Robots.txt', 0)) / max(pagine_totali, 1)
    )

    # 2) Canonical non self-ref
    canonical_non_self = 0.0
    if 'Canonical Link Element 1 Resolved' in df.columns:
        canon_df = df[['Address', 'Canonical Link Element 1 Resolved']].dropna()
        if len(canon_df) > 0:
            canonical_non_self = (canon_df['Address'] != canon_df['Canonical Link Element 1 Resolved']).sum() / max(pagine_totali, 1)
    penalita['Penalità Canonical %'] = canonical_non_self

    # 3) HTML tag
    penalita['Penalità Tag HTML %'] = (
        (kpi.get('Title Duplicati', 0) + kpi.get('Title Mancanti', 0) +
         kpi.get('Meta Description Duplicati', 0) + kpi.get('Meta Description Mancanti', 0) +
         kpi.get('H1 Duplicati', 0) + kpi.get('H1 Mancanti', 0)) / (3 * max(pagine_totali, 1))
    )

    # 4) Contenuti (duplicati da Title + thin da Word Count)
    if contenuti_penalty is None:
        contenuti_penalty = 0.0
    penalita['Penalità Contenuti %'] = float(min(1.0, max(0.0, contenuti_penalty)))

    # 5) CWV penalità + bonus da Assessment
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

    # Bonus/boost da Core Web Vitals Assessment (CM)
    pass_rate = kpi.get('CWV Assessment Pass %', None)
    if isinstance(pass_rate, (int, float)):
        if pass_rate > 60:
            base_cwv_pen *= 0.85   # bonus forte se >60%
        elif pass_rate > 50:
            base_cwv_pen *= 0.95   # bonus leggero se >50%
    penalita['Penalità CWV %'] = base_cwv_pen

    # Score finale (pesi invariati)
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

    # STATUS / ROBOTS
    df['Status Code'] = pd.to_numeric(df.get('Status Code'), errors='coerce')
    status = {
        'Pagine 2xx': df['Status Code'].between(200, 299).sum(),
        'Pagine 3xx': df['Status Code'].between(300, 399).sum(),
        'Pagine 4xx': df['Status Code'].between(400, 499).sum(),
        'Bloccate da Robots.txt': df.get('Indexability', pd.Series(index=df.index)).astype(str).str.contains("Blocked by Robots", na=False).sum(),
        'Pagine HTML Totali': df.shape[0]
    }

    # HTML TAG
    def analizza(col):
        if col not in df.columns:
            return (0, 0, 0)
        valid = df[col].dropna()
        return (
            len(valid[valid.duplicated(keep=False)].unique()),
            df[col].isna().sum(),
            df[col].notna().sum()
        )

    title = analizza("Title 1")
    description = analizza("Meta Description 1")
    h1 = analizza("H1-1")
    html_tag = {
        'Title Duplicati': title[0], 'Title Mancanti': title[1], 'Totale Title': title[2],
        'Meta Description Duplicati': description[0], 'Meta Description Mancanti': description[1], 'Totale Meta Description': description[2],
        'H1 Duplicati': h1[0], 'H1 Mancanti': h1[1], 'Totale H1': h1[2]
    }

    # CONTENUTI: Duplicati da Title + Thin (Word Count < media)
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

    content = {
        'Pagine Duplicate (da Title)': dup_pages,
        'Thin Content %': round(thin_rate * 100, 2),
        'Pagine Totali': pagine_totali
    }

    # Immagini senza ALT -> % di pagine con >=1 immagine senza ALT
    img_pages_pct = None
    if 'Images Missing Alt Text' in df.columns:
        missing_any = pd.to_numeric(df['Images Missing Alt Text'], errors='coerce').fillna(0) > 0
        img_pages_pct = round((missing_any.sum() / pagine_totali) * 100, 2)

    # CORE WEB VITALS (header testuali mappati da BM/BO/BR/CO/DC/CM)
    cwv_vals = {
        # BM
        'FCP': first_present_numeric(df, 'First contentfull Paint Time (ms)', 'FCP (ms)', 'FCP'),
        # BO
        'LCP': first_present_numeric(df, 'Largest contentfull Paint Time (ms)', 'LCP (ms)', 'LCP'),
        # BR (nota: CLS non è in ms)
        'CLS': first_present_numeric(df, 'Cumulative Layout Shift', 'CLS (ms)', 'CLS'),
        # CO
        'INP': first_present_numeric(df, 'CrUX Interaction to Next Paint (ms)', 'INP (ms)', 'INP'),
        # DC
        'TTFB': first_present_numeric(df, 'Server Response Times (TTFB) (ms)', 'TTFB (ms)', 'TTFB'),
    }
    cwv = {k: v for k, v in cwv_vals.items() if v is not None}

    # CM: Core Web Vitals Assessment
    assessment_pass_pct = None
    if 'Core Web Vitals Assessment' in df.columns:
        assess = df['Core Web Vitals Assessment'].astype(str).str.strip()
        total = len(assess)
        passes = (assess.str.lower() == 'pass').sum()
        assessment_pass_pct = round((passes / total) * 100, 2) if total else None

    # ASSEMBLA KPI
    kpi = {
        **status,
        **html_tag,
        **content,
        **cwv
    }
    if assessment_pass_pct is not None:
        kpi['CWV Assessment Pass %'] = assessment_pass_pct
    if img_pages_pct is not None:
        kpi['Pagine con Immagini senza ALT %'] = img_pages_pct

    # SCORE + PENALITÀ
    score, penalita = calcola_score(df, kpi, contenuti_penalty=contenuti_penalty)
    kpi['SEO Score'] = score
    kpi.update(penalita)
    return pd.DataFrame([kpi])

# =========================
# Score stabile – MULTI FILE
# =========================
def build_stable_domain_scores(df_riepilogo: pd.DataFrame) -> pd.DataFrame:
    """
    Stabilizza le penalità per dominio:
    - winsorization (1–99° percentile)
    - smoothing bayesiano verso media pesata per pagine
    - score stabile + IC 95% approssimato
    """
    df = df_riepilogo.copy()

    # Verifica colonne richieste
    needed_cols = ["Pagine Totali", *COMPONENT_MAP.values()]
    for c in needed_cols:
        if c not in df.columns:
            raise ValueError(f"Colonna mancante nel riepilogo: {c}")

    # Converti penalità (in % → proporzioni) + winsorization
    for col in COMPONENT_MAP.values():
        df[col] = winsorize_series(pd.to_numeric(df[col], errors="coerce") / 100.0)

    # Priors (media pesata per #pagine)
    priors = {}
    for key, col in COMPONENT_MAP.items():
        numer = (df[col] * df["Pagine Totali"].astype(float)).sum()
        denom = (df["Pagine Totali"].astype(float)).sum()
        priors[key] = float(numer / denom) if denom > 0 else float(pd.to_numeric(df[col], errors="coerce").mean())

    # Smoothing per dominio
    shrunk_cols = {}
    for key, col in COMPONENT_MAP.items():
        shrunk_name = f"{col} (shrunk)"
        shrunk_cols[key] = shrunk_name
        df[shrunk_name] = [
            smooth_proportion(p, n, priors[key])
            for p, n in zip(df[col].astype(float), df["Pagine Totali"].astype(float))
        ]

    # Pesi dinamici (se manca una componente)
    component_mask = {k: df[COMPONENT_MAP[k]].notna().any() for k in COMPONENT_MAP}
    dyn_w = redistribute_component_weights(component_mask, BASE_WEIGHTS)

    # Score stabile + IC
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
    df["Stato Stabile"] = pd.cut(
        df["SEO Score Stabile"],
        bins=[-0.1, 50, 70, 100.1],
        labels=["Critico", "Medio", "Buono"],
    )
    return df

def build_portfolio_score(df_stable: pd.DataFrame) -> dict:
    """Score complessivo su tutti i domini (media pesata sulle pagine) + IC 95%."""
    comp_shrunk = {k: f"{v} (shrunk)" for k, v in COMPONENT_MAP.items()}

    component_mask = {k: df_stable.get(comp_shrunk[k]) is not None and df_stable[comp_shrunk[k]].notna().any() for k in comp_shrunk}
    dyn_w = redistribute_component_weights(component_mask, BASE_WEIGHTS)

    N = df_stable["Pagine Totali"].astype(float)
    penalties_port = {}
    for k, col in comp_shrunk.items():
        if k not in dyn_w:
            continue
        p = (df_stable[col].astype(float) * N).sum() / max(N.sum(), 1.0)
        penalties_port[k] = float(p)

    score_port = penalties_to_score(penalties_port, dyn_w)

    N_tot = max(N.sum(), 1.0)
    var = 0.0
    for k, w in dyn_w.items():
        p = penalties_port[k]
        var += (w ** 2) * (p * (1 - p) / N_tot)
    se = np.sqrt(var) * 100.0

    return {
        "Portfolio Score": round(score_port, 2),
        "Portfolio SE (pt)": round(se, 2),
        "Portfolio CI95%": (round(score_port - 1.96 * se, 2), round(score_port + 1.96 * se, 2)),
        "Pesi Usati": dyn_w,
    }

# =========================
# UI – Streamlit
# =========================
st.set_page_config(page_title="SEO Audit Tool", layout="wide")
st.title("SEO Audit Tool")

tab1, tab2 = st.tabs(["Singolo File", "Multi File"])

# -------------------------
# Tab 1 - Singolo file
# -------------------------
with tab1:
    file = st.file_uploader("Carica un file .xlsx (Screaming Frog)", type="xlsx", key="single")
    if file:
        xls = pd.ExcelFile(file)
        sheet_name = None
        for name in ['1 - HTML', '1 - All']:
            if name in xls.sheet_names:
                sheet_name = name
                break
        if sheet_name:
            df = xls.parse(sheet_name)

            # KPI (versione aggiornata)
            kpi = estrai_kpi(df)

            st.subheader("Riepilogo SEO")
            st.dataframe(kpi, use_container_width=True)

            # Tabella pesi
            st.subheader("Sintesi SEO")
            st.markdown("### Pesi Score SEO")
            pesi_df = pd.DataFrame({
                "Componente": [
                    "Status Code e Robots",
                    "Canonical non self-ref",
                    "HTML Tag duplicati/mancanti",
                    "Contenuti (duplicati da Title + thin)",
                    "Core Web Vitals (con bonus Assessment)"
                ],
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

            # KPI sintetici
            cols = st.columns(4)
            cols[0].metric("SEO Score", f"{kpi['SEO Score'].iloc[0]}")
            if 'CWV Assessment Pass %' in kpi.columns:
                cols[1].metric("CWV Assessment Pass %", f"{kpi['CWV Assessment Pass %'].iloc[0]}%")
            if 'Pagine con Immagini senza ALT %' in kpi.columns:
                cols[2].metric("Pagine con ≥1 immagine senza ALT", f"{kpi['Pagine con Immagini senza ALT %'].iloc[0]}%")
            if 'TTFB' in kpi.columns:
                cols[3].metric("TTFB medio", f"{kpi['TTFB'].iloc[0]} ms")

            # Caption bonus CWV
            if 'CWV Assessment Pass %' in kpi.columns:
                pr = kpi['CWV Assessment Pass %'].iloc[0]
                if pr > 60:
                    st.caption("Bonus CWV applicato: penalità CWV ridotta del 15% (assessment > 60%).")
                elif pr > 50:
                    st.caption("Bonus CWV applicato: penalità CWV ridotta del 5% (assessment > 50%).")

            # Tabella compatta
            kpi_short = kpi[['SEO Score', 'Penalità Status Code %', 'Penalità Canonical %', 'Penalità Tag HTML %', 'Penalità Contenuti %', 'Penalità CWV %']].copy()
            kpi_short['Stato'] = kpi_short['SEO Score'].apply(lambda x: 'Critico' if x < 50 else 'Medio' if x < 70 else 'Buono')
            st.dataframe(kpi_short, use_container_width=True)

            # Radar chart
            if MATPLOTLIB_OK:
                import numpy as _np
                kpi_riepilogo = kpi[['Penalità Status Code %', 'Penalità Canonical %', 'Penalità Tag HTML %', 'Penalità Contenuti %', 'Penalità CWV %']]
                labels = kpi_riepilogo.columns.tolist()
                values = kpi_riepilogo.iloc[0].tolist()
                values = [v if isinstance(v, (int, float)) else 0 for v in values]
                angles = _np.linspace(0, 2 * _np.pi, len(labels), endpoint=False).tolist()
                values += values[:1]
                angles += angles[:1]

                fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))
                ax.plot(angles, values, 'o-', linewidth=2)
                ax.fill(angles, values, alpha=0.25)
                ax.set_yticklabels([])
                ax.set_xticks(angles[:-1])
                ax.set_xticklabels(labels)
                ax.set_title("Penalità SEO Radar Chart")
                st.pyplot(fig)

            # Export Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                kpi.to_excel(writer, sheet_name='SEO Report', index=False)
            st.download_button("📥 Scarica il report Excel", buffer.getvalue(), file_name="seo_report.xlsx")
        else:
            st.warning("Il file non contiene un foglio valido ('1 - HTML' o '1 - All').")

# -------------------------
# Tab 2 - Multi File
# -------------------------
with tab2:
    files = st.file_uploader("Carica più file .xlsx (Screaming Frog)", type="xlsx", accept_multiple_files=True, key="multi")
    if files:
        risultati = []
        for f in files:
            try:
                xls = pd.ExcelFile(f)
                sheet_name = None
                for name in ['1 - HTML', '1 - All']:
                    if name in xls.sheet_names:
                        sheet_name = name
                        break
                if not sheet_name:
                    continue
                df = xls.parse(sheet_name)
                dominio = estrai_dominio(df) or os.path.splitext(f.name)[0].split("_")[0]
                kpi = estrai_kpi(df)
                kpi.insert(0, 'Dominio', dominio)
                risultati.append(kpi)
            except Exception as e:
                st.warning(f"Errore nel file '{getattr(f, 'name', 'sconosciuto')}': {e}")

        if risultati:
            df_riepilogo = pd.concat(risultati, ignore_index=True)

            st.subheader("Riepilogo SEO per tutti i domini")
            st.dataframe(df_riepilogo, use_container_width=True)

            # Tabella sintetica aggiornata
            df_short = df_riepilogo[['Dominio', 'SEO Score', 'Penalità Status Code %', 'Penalità Canonical %', 'Penalità Tag HTML %', 'Penalità Contenuti %', 'Penalità CWV %', 'Pagine Totali']].copy()
            df_short['Stato'] = df_short['SEO Score'].apply(lambda x: 'Critico' if x < 50 else 'Medio' if x < 70 else 'Buono')

            st.subheader("Sintesi SEO per dominio")
            st.markdown("### Pesi Score SEO (per ogni dominio)")
            tabella_pesi = pd.DataFrame({
                "Componente": [
                    "Status Code e Robots",
                    "Canonical non self-ref",
                    "HTML Tag duplicati/mancanti",
                    "Contenuti (duplicati da Title + thin)",
                    "Core Web Vitals (con bonus Assessment)"
                ],
                "Peso (%)": [30, 15, 20, 10, 20]
            })
            st.dataframe(tabella_pesi, use_container_width=True)
            st.dataframe(df_short, use_container_width=True)

            # Score stabile multi-file
            try:
                # Normalizza nome "Pagine Totali" (alcuni export hanno "Pagine HTML Totali")
                if "Pagine Totali" not in df_riepilogo.columns and "Pagine HTML Totali" in df_riepilogo.columns:
                    df_riepilogo = df_riepilogo.rename(columns={"Pagine HTML Totali": "Pagine Totali"})
                df_stable = build_stable_domain_scores(df_riepilogo)
                st.subheader("Riepilogo SEO STABILE per tutti i domini")
                st.dataframe(df_stable[['Dominio', 'Pagine Totali', 'SEO Score Stabile', 'Score CI95% Min', 'Score CI95% Max', 'Stato Stabile']], use_container_width=True)

                # Score complessivo portafoglio
                port = build_portfolio_score(df_stable)
                st.markdown(f"### Score complessivo portafoglio: **{port['Portfolio Score']}** (IC95% {port['Portfolio CI95%'][0]} – {port['Portfolio CI95%'][1]})")
                st.caption(f"Pesi componenti usati: {port['Pesi Usati']}")

                # Confronto score tra domini (bar chart)
                st.markdown("### Confronto Score SEO tra domini")
                st.bar_chart(df_short.set_index("Dominio")["SEO Score"])

            except Exception as e:
                st.warning(f"Impossibile calcolare lo score stabile: {e}")

            # =============================
            # BENCHMARK COMPETITIVO (custom)
            # =============================
            st.subheader("Benchmark competitivo")

            # 1) Pesi di default (sliders)
            with st.sidebar:
                st.markdown("### Pesi SCORE (default – verranno normalizzati)")
                w_indexability = st.slider("Peso INDEXABILITY", 0, 100, 20)
                w_cwv          = st.slider("Peso CWV", 0, 100, 20)
                w_html         = st.slider("Peso HTML TAGS", 0, 100, 20)
                w_content      = st.slider("Peso CONTENT EVALUATION", 0, 100, 20)
                w_dr           = st.slider("Peso DR", 0, 100, 10)
                w_lv           = st.slider("Peso LINK VELOCITY", 0, 100, 10)

            weights_default_raw = {
                "INDEXABILITY": w_indexability,
                "CWV": w_cwv,
                "HTML TAGS": w_html,
                "CONTENT EVALUATION": w_content,
                "DR": w_dr,
                "LINK VELOCITY": w_lv,
            }
            wsum = sum(weights_default_raw.values()) or 1
            weights_default = {k: v / wsum for k, v in weights_default_raw.items()}

            # 2) Editor manuale DR / LINK VELOCITY
            _domains = df_riepilogo["Dominio"].tolist()
            manual_defaults = pd.DataFrame({
                "Dominio": _domains,
                "DR": [None] * len(_domains),
                "LINK VELOCITY": [None] * len(_domains),
            })
            st.markdown("Compila **DR** e **LINK VELOCITY** (se disponi dei dati):")
            manual_input = st.data_editor(
                manual_defaults,
                use_container_width=True,
                key="manual_ext",
                hide_index=True
            )

            # utility: tutti i domini hanno DR & LV numerici valorizzati?
            def _all_offpage_ready(dfm: pd.DataFrame) -> bool:
                if dfm.empty or any(c not in dfm.columns for c in ["DR", "LINK VELOCITY"]):
                    return False
                def _is_num(x):
                    try:
                        return np.isfinite(float(x))
                    except Exception:
                        return False
                return dfm["DR"].apply(_is_num).all() and dfm["LINK VELOCITY"].apply(_is_num).all()

            offpage_ready = _all_offpage_ready(manual_input)

            # 3) Scelta automatica dei pesi
            if offpage_ready:
                # schema richiesto quando DR & LV sono compilati per tutti i domini
                weights = {
                    "INDEXABILITY": 0.33,
                    "HTML TAGS": 0.11,
                    "CONTENT EVALUATION": 0.11,
                    "CWV": 0.11,
                    "DR": 0.165,
                    "LINK VELOCITY": 0.165,
                }
                st.caption("Schema pesi **OFF-PAGE 33%** attivo (DR e Link Velocity compilati per tutti i domini).")
            else:
                weights = weights_default
                st.caption("Schema pesi **default** attivo (mancano valori completi di DR/Link Velocity).")

            manual_map = manual_input.set_index("Dominio").to_dict("index")

            # 4) Costruzione tabella benchmark
            rows = ["INDEXABILITY", "CWV", "HTML TAGS", "CONTENT EVALUATION", "DR", "LINK VELOCITY"]
            cols = [d.upper().replace("WWW.", "") for d in _domains]
            bench = pd.DataFrame(index=rows, columns=cols, dtype="float")

            for dom_raw, dom_col in zip(_domains, cols):
                row = df_riepilogo.loc[df_riepilogo["Dominio"] == dom_raw].iloc[0]

                indexability = 100.0 - float(row.get("Penalità Status Code %", 0.0))
                cwv = float(row.get("CWV Assessment Pass %", 0.0)) if "CWV Assessment Pass %" in df_riepilogo.columns else 0.0
                html_tags_ok = 100.0 - float(row.get("Penalità Tag HTML %", 0.0))
                content_ok = 100.0 - float(row.get("Penalità Contenuti %", 0.0))

                # DR / LV: usa valori manuali (se non numerici → NaN)
                def _to_num(v):
                    try:
                        return float(v)
                    except Exception:
                        return np.nan
                dr_val = _to_num(manual_map.get(dom_raw, {}).get("DR"))
                lv_val = _to_num(manual_map.get(dom_raw, {}).get("LINK VELOCITY"))

                bench.loc["INDEXABILITY", dom_col] = round(indexability, 2)
                bench.loc["CWV", dom_col] = round(cwv, 2)
                bench.loc["HTML TAGS", dom_col] = round(html_tags_ok, 2)
                bench.loc["CONTENT EVALUATION", dom_col] = round(content_ok, 2)
                bench.loc["DR", dom_col] = dr_val
                bench.loc["LINK VELOCITY", dom_col] = lv_val

            # 5) SCORE pesato
            score = []
            for dom_col in bench.columns:
                total = 0.0
                for r in rows:
                    val = bench.loc[r, dom_col]
                    if pd.isna(val):
                        val = 0.0  # se manca, conta 0
                    total += float(val) * weights[r]
                score.append(int(round(total, 0)))
            bench.loc["SCORE"] = score

            # Stile e rendering
            def _hl_score(r):
                return ['background-color: #2F3B52; color: white; font-weight: 700;' if r.name == "SCORE" else '' for _ in r]

            bench_styled = bench.style.format(
                lambda v: f"{v:.2f}" if isinstance(v, (int, float)) and not pd.isna(v) and v != int(v) else ("" if pd.isna(v) else f"{int(v)}")
            ).apply(_hl_score, axis=1)

            st.dataframe(bench_styled, use_container_width=True)

            # Export Benchmark
            buf_bench = io.BytesIO()
            with pd.ExcelWriter(buf_bench, engine="openpyxl") as writer:
                bench.to_excel(writer, sheet_name="Benchmark", index=True)
            st.download_button("📥 Scarica tabella benchmark", buf_bench.getvalue(), file_name="benchmark_competitors.xlsx")

            # Export Excel multi-file (KPI + stabile + benchmark)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # Un foglio per dominio con KPI
                for dominio, group in df_riepilogo.groupby('Dominio'):
                    group.drop(columns=['Dominio']).to_excel(writer, sheet_name=dominio[:31], index=False)
                # Riepilogo
                df_riepilogo.to_excel(writer, sheet_name='Riepilogo', index=False)
                # Stabile
                try:
                    df_stable.to_excel(writer, sheet_name='Riepilogo Stabile', index=False)
                except Exception:
                    pass
                # Benchmark
                try:
                    bench.to_excel(writer, sheet_name='Benchmark', index=True)
                except Exception:
                    pass
            st.download_button("📥 Scarica il report Excel (multi-file)", buffer.getvalue(), file_name="seo_audit_multi.xlsx")

        else:
            st.warning("Nessun file valido con foglio '1 - HTML' o '1 - All' trovato.")
