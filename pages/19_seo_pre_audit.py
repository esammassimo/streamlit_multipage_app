"""
SEO Audit App — NVL Agency
Interfaccia Streamlit per seo_audit.py

Deploy su Streamlit Cloud:
  File principale : seo_audit_app.py
  Requirements    : requirements.txt  (include playwright)
  Pacchetti apt   : packages.txt      (dipendenze sistema Chromium)

Locale:
  streamlit run seo_audit_app.py
"""

import streamlit as st
import sys
import os
import json
import time
import tempfile
import subprocess
import pandas as pd
from datetime import datetime

# ─── IMPORT seo_audit ────────────────────────────────────────────────────────
# Prova la directory dello script e il cwd: su Cloud sono spesso diversi.

for _p in [os.path.dirname(os.path.abspath(__file__)), os.getcwd()]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

try:
    import seo_audit as sa
    AUDIT_OK  = True
    AUDIT_ERR = ""
except Exception as _e:
    AUDIT_OK  = False
    AUDIT_ERR = str(_e)


# ─── PLAYWRIGHT: installa browser una sola volta per istanza server ───────────
# Su Streamlit Cloud il pacchetto playwright è installato da requirements.txt,
# ma il browser Chromium va scaricato a runtime.
# @st.cache_resource garantisce che il download avvenga una volta sola
# (non ad ogni rerun o cambio utente).

@st.cache_resource(show_spinner=False)
def _install_chromium():
    """Scarica chromium-headless-shell se non già presente."""
    try:
        result = subprocess.run(
            ["python3", "-m", "playwright", "install", "chromium"],
            capture_output=True,
            text=True,
            timeout=300,   # 5 minuti — sufficiente anche su Cloud lento
        )
        success = result.returncode == 0
        return success, (result.stdout + result.stderr).strip()
    except Exception as exc:
        return False, str(exc)

_pw_installed, _pw_log = _install_chromium()


# ─── CONFIGURAZIONE ──────────────────────────────────────────────────────────
# Tutto qui. Nessuna config in sidebar o form utente.

# Playwright: attivo se il browser è stato installato con successo
USE_PLAYWRIGHT: bool = _pw_installed

# Timeout fetch HTTP statico (secondi)
FETCH_TIMEOUT: int = 15

# Numero massimo di URL analizzabili da file (protezione Cloud)
MAX_URLS: int = 20


# ─── COSTANTI UI ─────────────────────────────────────────────────────────────

SC  = {"OK": 100, "WARN": 50, "FAIL": 0, "ERROR": 0}
BC  = {"OK": "#16a34a", "WARN": "#b45309", "FAIL": "#dc2626", "ERROR": "#6b7280"}
BL  = {"OK": "OK",      "WARN": "WARN",    "FAIL": "FAIL",    "ERROR": "N/D"}

AREA_META = {
    "1.":  ("1",  "Rendering & Visibilità",  "Contenuti nascosti, lazy-load, delta DOM"),
    "2.":  ("2",  "Struttura Heading",        "H1, H2/H3, gerarchia, blocchi orfani"),
    "3.":  ("3",  "Robots.txt",               "Direttive, sitemap, configurazione crawler"),
    "8.":  ("8",  "E-E-A-T Signals",          "Autore, fonti, schema, YMYL detection"),
    "9.":  ("9",  "Performance",              "CLS, LCP, script bloccanti, framework JS"),
    "10.": ("10", "Topical Authority",        "Title, meta desc, canonical, citazioni, link"),
}

BADGE_COLS = ["Visibilità", "Heading", "Robots", "E-E-A-T", "Perf.", "Topical"]


# ─── HELPERS ─────────────────────────────────────────────────────────────────

def r_by(results, prefix):
    """Risultato di una specifica area."""
    return next((r for r in results if r.get("area", "").startswith(prefix)), {})

def page_score(results):
    s = [SC.get(r.get("overall", "ERROR"), 0) for r in results]
    return round(sum(s) / len(s)) if s else 0

def global_score(grouped):
    s = [page_score(pg) for pg in grouped]
    return round(sum(s) / len(s)) if s else 0

def score_color(s):
    return BC["OK"] if s >= 80 else (BC["WARN"] if s >= 50 else BC["FAIL"])

def badge(overall):
    color = BC.get(overall, "#6b7280")
    label = BL.get(overall, overall)
    return (f'<span style="background:{color};color:white;padding:2px 10px;'
            f'border-radius:4px;font-size:12px;font-weight:700">{label}</span>')

def color_cell(val):
    m = {
        "OK":   "background-color:#d1fae5;color:#065f46;font-weight:700",
        "WARN": "background-color:#fef3c7;color:#78350f;font-weight:700",
        "FAIL": "background-color:#fee2e2;color:#7f1d1d;font-weight:700",
    }
    return m.get(val, "")

def export_download(label, generate_fn, suffix, mime):
    """Genera un file temporaneo e mostra il bottone download."""
    if st.button(f"Genera e scarica {label}", use_container_width=True):
        with st.spinner("Generazione…"):
            try:
                with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                    tmp_path = tmp.name
                generate_fn(tmp_path)
                data = open(tmp_path, "rb").read()
                os.unlink(tmp_path)
                fname = f"seo_audit_{datetime.now().strftime('%Y%m%d_%H%M')}{suffix}"
                st.download_button(f"⬇️ Scarica {label}", data=data,
                                   file_name=fname, mime=mime,
                                   use_container_width=True)
            except Exception as exc:
                st.error(f"Errore: {exc}")


# ─── PAGE CONFIG ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="SEO Audit — NVL Agency",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
/* Badge inline */
.sig-ok   { color:#065f46; padding:4px 0; font-size:14px;
            border-bottom:1px solid #f0fdf4; display:block; }
.sig-warn { color:#78350f; padding:4px 0; font-size:14px;
            border-bottom:1px solid #fffbeb; display:block; }
/* Callout info */
.callout  { background:#eff6ff; border:1px solid #bfdbfe; border-radius:8px;
            padding:12px 16px; font-size:13px; color:#1e40af; margin-bottom:12px; }
/* Heading map */
.hmap     { padding:3px 0; font-size:13px; border-bottom:1px solid #f9fafb; }
/* Area card overview */
.acard    { background:white; border:1px solid #e5e7eb; border-radius:10px;
            padding:16px 18px; margin-bottom:12px; }
</style>
""", unsafe_allow_html=True)


# ─── HEADER ──────────────────────────────────────────────────────────────────

col_logo, col_hd, col_st = st.columns([1, 7, 4])
with col_logo:
    st.markdown("# 🔍")
with col_hd:
    st.title("SEO Audit")
    st.caption("NVL Agency · Analisi tecnica on-page · 9 aree di analisi")
with col_st:
    st.markdown("")  # spacing
    if AUDIT_OK:
        st.success("19_seo_pre_audit.py ✅")
    else:
        st.error(f"seo_audit.py non trovato\n`{AUDIT_ERR[:100]}`")

    if USE_PLAYWRIGHT:
        st.success("Playwright ✅ — rendering JS attivo")
    else:
        st.warning("Playwright ⚠️ — solo fetch statico")
        if _pw_log:
            with st.expander("Log Playwright"):
                st.code(_pw_log[-800:], language="text")

st.divider()


# ─── INPUT ───────────────────────────────────────────────────────────────────

tab_single, tab_sitemap, tab_file = st.tabs([
    "🌐 URL singola", "🗺️ Sitemap XML", "📄 File CSV / Excel"
])
urls_to_audit: list[str] = []
input_label: str = ""   # etichetta usata come audit_url nel report

# ── Tab 1: URL singola ────────────────────────────────────────────────────────
with tab_single:
    col_u, col_b = st.columns([6, 1])
    with col_u:
        raw_url = st.text_input(
            "URL",
            placeholder="https://tuosito.com/pagina",
            label_visibility="collapsed",
        )
    with col_b:
        run_single = st.button(
            "▶ Analizza", type="primary",
            use_container_width=True,
            disabled=not AUDIT_OK,
        )
    if run_single:
        if raw_url.strip().startswith("http"):
            urls_to_audit = [raw_url.strip()]
            input_label   = raw_url.strip()
        elif raw_url.strip():
            st.warning("L'URL deve iniziare con http:// o https://")

# ── Tab 2: Sitemap XML ────────────────────────────────────────────────────────
with tab_sitemap:
    st.caption(
        f"Inserisci l'URL della sitemap XML. "
        f"Verranno analizzate al massimo **{MAX_URLS}** URL."
    )
    col_sm, col_bsm = st.columns([6, 1])
    with col_sm:
        sitemap_url = st.text_input(
            "Sitemap URL",
            placeholder="https://tuosito.com/sitemap.xml",
            label_visibility="collapsed",
            key="sitemap_url",
        )
    with col_bsm:
        run_sitemap = st.button(
            "▶ Analizza", type="primary",
            use_container_width=True,
            key="run_sitemap",
            disabled=not AUDIT_OK,
        )

    if sitemap_url.strip():
        # Anteprima URL in sitemap
        if st.button("👁 Anteprima URL sitemap", key="preview_sitemap"):
            with st.spinner("Lettura sitemap…"):
                try:
                    preview_urls = sa.get_urls_from_sitemap(
                        sitemap_url.strip(), sa.get_session(), MAX_URLS
                    )
                    if preview_urls:
                        st.info(f"✅ {len(preview_urls)} URL trovati nella sitemap")
                        with st.expander(f"Mostra URL ({len(preview_urls)})"):
                            for u in preview_urls:
                                st.code(u, language="text")
                    else:
                        st.warning("Nessuna URL trovata. Verifica che la sitemap sia accessibile e in formato XML standard.")
                except Exception as exc:
                    st.error(f"Errore lettura sitemap: {exc}")

    if run_sitemap:
        if sitemap_url.strip().startswith("http"):
            with st.spinner("Lettura sitemap…"):
                try:
                    from_sitemap = sa.get_urls_from_sitemap(
                        sitemap_url.strip(), sa.get_session(), MAX_URLS
                    )
                    if from_sitemap:
                        urls_to_audit = from_sitemap
                        input_label   = sitemap_url.strip()
                    else:
                        st.error("Nessuna URL trovata nella sitemap.")
                except Exception as exc:
                    st.error(f"Errore lettura sitemap: {exc}")
        elif sitemap_url.strip():
            st.warning("L'URL della sitemap deve iniziare con http:// o https://")

# ── Tab 3: File CSV / Excel ───────────────────────────────────────────────────
with tab_file:
    st.caption(
        f"Il file deve avere una colonna chiamata **url**. "
        f"Massimo **{MAX_URLS}** URL per analisi."
    )
    col_up, col_bup = st.columns([6, 1])
    with col_up:
        uploaded = st.file_uploader(
            "File",
            type=["csv", "xlsx", "xls"],
            label_visibility="collapsed",
        )
    with col_bup:
        run_file = st.button(
            "▶ Analizza", type="primary",
            use_container_width=True,
            key="run_file",
            disabled=not AUDIT_OK,
        )

    if uploaded:
        try:
            df_up = (pd.read_csv(uploaded) if uploaded.name.endswith(".csv")
                     else pd.read_excel(uploaded))
            col_name = next(
                (c for c in df_up.columns
                 if c.lower() in ["url", "urls", "link", "links", "pagina", "page"]),
                None,
            )
            if col_name:
                from_file = (
                    df_up[col_name].dropna().astype(str).str.strip()
                    .pipe(lambda s: s[s.str.startswith("http")])
                    .tolist()[:MAX_URLS]
                )
                st.info(f"✅ {len(from_file)} URL trovati nel file")
                with st.expander(f"Anteprima URL ({min(5, len(from_file))} di {len(from_file)})"):
                    for u in from_file[:5]:
                        st.code(u, language="text")
                if run_file:
                    urls_to_audit = from_file
                    input_label   = uploaded.name
            else:
                st.error(
                    "Colonna 'url' non trovata. "
                    "Rinomina la colonna con gli URL in **url**."
                )
        except Exception as exc:
            st.error(f"Errore lettura file: {exc}")


# ─── ESECUZIONE AUDIT ────────────────────────────────────────────────────────

if urls_to_audit and AUDIT_OK:
    st.session_state["grouped"]   = []
    st.session_state["audit_url"] = input_label or urls_to_audit[0]
    st.session_state["audit_ts"]  = datetime.now().strftime("%d/%m/%Y %H:%M")

    total      = len(urls_to_audit)
    is_multi   = total > 1          # sitemap o file → mostra loader dettagliato
    pbar       = st.progress(0, text="Avvio…")
    info       = st.empty()

    # Contatore URL — visibile solo per analisi multi-URL
    url_counter = st.empty() if is_multi else None

    steps = [
        ("1.",  "Rendering & Visibilità",
         lambda u, h, rh, ss: sa.audit_content_visibility(u, h, rh)),
        ("2.",  "Struttura Heading",
         lambda u, h, rh, ss: sa.audit_heading_structure(u, h)),
        ("3.",  "Robots.txt",
         lambda u, h, rh, ss: sa.audit_robots_txt(u, ss)),
        ("8.",  "E-E-A-T Signals",
         lambda u, h, rh, ss: sa.audit_eeat_signals(u, h)),
        ("9.",  "Performance",
         lambda u, h, rh, ss: sa.audit_performance_signals(u, h)),
        ("10.", "Topical Authority",
         lambda u, h, rh, ss: sa.audit_topical_authority(u, h, ss)),
    ]
    n = len(steps)

    for idx, url in enumerate(urls_to_audit):
        info.info(f"🔍 {url}")

        # Loader URL-per-URL per sitemap e file
        if is_multi and url_counter is not None:
            url_counter.markdown(
                f"**Pagina {idx + 1} di {total}** analizzata &nbsp;·&nbsp; "
                f"{idx} completate &nbsp;·&nbsp; {total - idx - 1} rimanenti"
            )

        page_res = []

        try:
            session = sa.get_session()

            pbar.progress(idx / total, text=f"[{idx+1}/{total}] Fetch statico…")
            static_html, _, final_url = sa.fetch_html_static(url, session)

            if not static_html:
                info.error(f"❌ {url} — impossibile recuperare ({final_url})")
                pbar.progress((idx + 1) / total)
                continue

            rendered_html = None
            if USE_PLAYWRIGHT:
                pbar.progress(
                    (idx + 0.1) / total,
                    text=f"[{idx+1}/{total}] Rendering JavaScript…",
                )
                rendered_html = sa.fetch_html_rendered(url)

            for si, (pfx, name, fn) in enumerate(steps):
                pbar.progress(
                    (idx + (si + 1) / n) / total,
                    text=f"[{idx+1}/{total}] Area {pfx.rstrip('.')} — {name}",
                )
                r = fn(url, static_html, rendered_html, session)
                r["url"] = url
                page_res.append(r)

            st.session_state["grouped"].append(page_res)

        except Exception as exc:
            info.error(f"❌ Errore su {url}: {exc}")

        pbar.progress((idx + 1) / total, text=f"Completati {idx+1}/{total}")

    pbar.progress(1.0, text=f"✅ Completato — {len(st.session_state['grouped'])} URL analizzate")
    time.sleep(0.6)
    pbar.empty()
    info.empty()
    if url_counter is not None:
        url_counter.empty()
    st.rerun()


# ─── WELCOME ─────────────────────────────────────────────────────────────────

if "grouped" not in st.session_state or not st.session_state["grouped"]:
    st.markdown("### Come iniziare")
    st.markdown(
        "Inserisci un URL nella barra qui sopra oppure carica un file CSV/Excel "
        f"con una colonna **url** (max {MAX_URLS} righe), poi premi **▶ Analizza**."
    )
    st.divider()
    st.markdown("#### Aree di analisi")
    cols = st.columns(3)
    for i, (pfx, (num, name, desc)) in enumerate(AREA_META.items()):
        with cols[i % 3]:
            st.markdown(f"**Area {num} — {name}**")
            st.caption(desc)
    st.stop()


# ─── DATI SESSIONE ────────────────────────────────────────────────────────────

grouped   = st.session_state["grouped"]
audit_url = st.session_state.get("audit_url", "")
audit_ts  = st.session_state.get("audit_ts", "")
npages    = len(grouped)
gs        = global_score(grouped)
gc        = score_color(gs)

first_pg = grouped[0]
r1f  = r_by(first_pg, "1.")
r2f  = r_by(first_pg, "2.")
r3f  = r_by(first_pg, "3.")
r8f  = r_by(first_pg, "8.")
r9f  = r_by(first_pg, "9.")
r10f = r_by(first_pg, "10.")

st.divider()
st.caption(
    f"**{audit_url}** &nbsp;·&nbsp; {audit_ts} &nbsp;·&nbsp; "
    f"{npages} {'pagina' if npages == 1 else 'pagine'} &nbsp;·&nbsp; "
    f"Playwright {'attivo ✅' if USE_PLAYWRIGHT else '⚠️ disabilitato'}"
)


# ─── TAB PRINCIPALI ──────────────────────────────────────────────────────────

tabs = st.tabs([
    "📊 Overview", "📋 Pagine", "⚠️ Problemi",
    "1 Visibilità", "2 Heading", "3 Robots.txt",
    "8 E-E-A-T", "9 Performance", "10 Topical",
    "⬇️ Esporta",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 0 — OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════

with tabs[0]:
    n_fail = sum(1 for pg in grouped for r in pg if r.get("overall") == "FAIL")
    n_warn = sum(1 for pg in grouped for r in pg if r.get("overall") == "WARN")
    n_ok   = sum(1 for pg in grouped for r in pg if r.get("overall") == "OK")

    c_sc, c_f, c_w, c_ok, c_pg = st.columns([2, 1, 1, 1, 1])
    with c_sc:
        st.metric("Score globale", f"{gs} / 100")
        st.progress(gs / 100)
    c_f.metric("FAIL",   n_fail)
    c_w.metric("WARN",   n_warn)
    c_ok.metric("OK",    n_ok)
    c_pg.metric("Pagine", npages)

    st.divider()
    st.subheader("Esito per area — prima URL")

    cols_a = st.columns(3)
    for i, pfx in enumerate(AREA_META):
        r   = r_by(first_pg, pfx)
        ov  = r.get("overall", "ERROR")
        num, name, _ = AREA_META[pfx]
        bc_ = BC.get(ov, "#6b7280")
        txt = r.get("summary", "—").split(" | ")[0][:140]
        with cols_a[i % 3]:
            st.markdown(
                f'<div class="acard" style="border-top:4px solid {bc_}">'
                f'<div style="font-size:10px;color:#9ca3af;letter-spacing:.08em;'
                f'text-transform:uppercase;margin-bottom:2px">AREA {num}</div>'
                f'<div style="font-size:15px;font-weight:700;margin-bottom:8px">{name}</div>'
                f'{badge(ov)}'
                f'<div style="font-size:13px;color:#4b5563;margin-top:10px;line-height:1.6">{txt}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — PAGINE
# ══════════════════════════════════════════════════════════════════════════════

with tabs[1]:
    st.subheader(f"Pagine analizzate — {npages} URL")
    rows = []
    for pg in grouped:
        pg_url = pg[0].get("url", "—") if pg else "—"
        ovs = {r.get("area", "")[:3]: r.get("overall", "N/D") for r in pg}
        rows.append({
            "URL":        pg_url,
            "Visibilità": ovs.get("1. ", "N/D"),
            "Heading":    ovs.get("2. ", "N/D"),
            "Robots":     ovs.get("3. ", "N/D"),
            "E-E-A-T":    ovs.get("8. ", "N/D"),
            "Perf.":      ovs.get("9. ", "N/D"),
            "Topical":    ovs.get("10.", "N/D"),
            "Score":      page_score(pg),
        })
    df_pg = pd.DataFrame(rows)
    st.dataframe(
        df_pg.style
             .applymap(color_cell, subset=BADGE_COLS)
             .background_gradient(subset=["Score"], cmap="RdYlGn", vmin=0, vmax=100),
        use_container_width=True,
        height=min(600, 80 + 38 * len(df_pg)),
    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — PROBLEMI
# ══════════════════════════════════════════════════════════════════════════════

with tabs[2]:
    probs = []
    for pg in grouped:
        u = pg[0].get("url", "—") if pg else "—"
        for r in pg:
            area = r.get("area", "")[:28]
            # 1 — visibilità
            for b in r.get("hidden_blocks", []):
                probs.append({"Area": area, "URL": u, "Tipo": "Contenuto nascosto",
                              "Severity": "WARN",
                              "Dettaglio": f'{b["selector"]} → {b["text_preview"][:80]}'})
            if r.get("lazy_elements", 0):
                probs.append({"Area": area, "URL": u, "Tipo": "Lazy-load",
                              "Severity": "WARN",
                              "Dettaglio": f'{r["lazy_elements"]} elementi lazy-load'})
            for pat in r.get("js_patterns", []):
                probs.append({"Area": area, "URL": u, "Tipo": "Pattern JS",
                              "Severity": "WARN", "Dettaglio": pat})
            # 2 — heading
            if r.get("h1_count") == 0:
                probs.append({"Area": area, "URL": u, "Tipo": "H1 assente",
                              "Severity": "FAIL", "Dettaglio": "Nessun H1 trovato"})
            elif r.get("h1_count", 1) > 1:
                probs.append({"Area": area, "URL": u, "Tipo": "H1 multipli",
                              "Severity": "WARN",
                              "Dettaglio": f'{r["h1_count"]} H1 — deve essere uno solo'})
            for jmp in r.get("level_jumps", []):
                probs.append({"Area": area, "URL": u, "Tipo": "Salto heading",
                              "Severity": "WARN", "Dettaglio": jmp})
            # 3 — robots
            for f in r.get("findings", []):
                if isinstance(f, dict) and f.get("severity") in ("WARN", "FAIL"):
                    probs.append({"Area": area, "URL": u, "Tipo": f.get("type", "—"),
                                  "Severity": f["severity"],
                                  "Dettaglio": f.get("note", "")[:120]})
            # 8 — E-E-A-T
            for m in r.get("missing", []):
                probs.append({"Area": area, "URL": u, "Tipo": "E-E-A-T mancante",
                              "Severity": "WARN", "Dettaglio": m})
            if r.get("is_ymyl") and r.get("score", 100) < 50:
                probs.append({"Area": area, "URL": u, "Tipo": "YMYL + E-E-A-T basso",
                              "Severity": "FAIL",
                              "Dettaglio": f'Score {r.get("score", 0)}/100 su pagina YMYL'})
            # 9 — performance
            for f in r.get("findings", []):
                if isinstance(f, dict) and "impatto" in f and f.get("severity") in ("WARN", "FAIL"):
                    probs.append({"Area": area, "URL": u, "Tipo": f.get("tipo", "—"),
                                  "Severity": f["severity"],
                                  "Dettaglio": f'{f.get("nota","")[:100]} [{f.get("impatto","")}]'})
            # 10 — topical
            for opp in r.get("opportunities", []):
                probs.append({"Area": area, "URL": u, "Tipo": "Opportunità",
                              "Severity": "WARN", "Dettaglio": opp[:120]})

    probs.sort(key=lambda x: {"FAIL": 0, "WARN": 1, "INFO": 2}.get(x["Severity"], 9))
    st.subheader(f"Problemi rilevati — {len(probs)} finding")

    if probs:
        sev_f = st.multiselect("Filtra severity", ["FAIL", "WARN", "INFO"],
                               default=["FAIL", "WARN"])
        df_pr = pd.DataFrame(probs)
        df_pr = df_pr[df_pr["Severity"].isin(sev_f)]
        st.dataframe(
            df_pr.style.applymap(color_cell, subset=["Severity"]),
            use_container_width=True,
            height=min(700, 80 + 38 * len(df_pr)),
        )
    else:
        st.success("✅ Nessun problema rilevato.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — AREA 1: VISIBILITÀ
# ══════════════════════════════════════════════════════════════════════════════

with tabs[3]:
    r = r1f; ov = r.get("overall", "ERROR")
    ct, cb = st.columns([7, 1])
    ct.subheader("Rendering & Visibilità Contenuti")
    cb.markdown(f"<div style='padding-top:10px'>{badge(ov)}</div>", unsafe_allow_html=True)
    st.caption(r.get("summary", "").replace(" | ", " · "))
    st.divider()

    c1, c2, c3 = st.columns(3)
    c1.metric("Blocchi nascosti",   len(r.get("hidden_blocks", [])))
    c2.metric("Elementi lazy-load", r.get("lazy_elements", 0))
    c3.metric("Pattern JS",         len(r.get("js_patterns", [])))
    st.caption(f"Delta DOM statico / renderizzato: `{r.get('dom_delta', 'N/A')}`")

    if r.get("js_patterns"):
        st.markdown("**Pattern JS rilevati:**")
        st.code("\n".join(r["js_patterns"]), language="javascript")

    if r.get("hidden_blocks"):
        st.markdown("**Blocchi contenuto nascosto:**")
        df_hb = pd.DataFrame([{
            "Selettore":        b["selector"],
            "Tag":              f'<{b["tag"]}>',
            "Contiene heading": "⚠ Sì" if b["has_heading"] else "No",
            "Anteprima":        b["text_preview"],
        } for b in r.get("hidden_blocks", [])])
        st.dataframe(df_hb, use_container_width=True, hide_index=True)
    else:
        st.success("✅ Nessun blocco contenuto nascosto rilevato.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — AREA 2: HEADING
# ══════════════════════════════════════════════════════════════════════════════

with tabs[4]:
    r = r2f; ov = r.get("overall", "ERROR")
    ct, cb = st.columns([7, 1])
    ct.subheader("Struttura Heading H2/H3")
    cb.markdown(f"<div style='padding-top:10px'>{badge(ov)}</div>", unsafe_allow_html=True)
    st.caption(r.get("summary", "").replace(" | ", " · "))
    st.divider()

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("H1",             r.get("h1_count", 0))
    c2.metric("H2",             r.get("h2_count", 0))
    c3.metric("H3",             r.get("h3_count", 0))
    c4.metric("Salti livello",  len(r.get("level_jumps", [])))
    c5.metric("Blocchi orfani", len(r.get("orphan_blocks", [])))

    if r.get("level_jumps"):
        st.warning("**Salti di livello:**\n" + "\n".join(f"- {j}" for j in r["level_jumps"]))

    if r.get("heading_map"):
        st.markdown("**Mappa heading:**")
        icons = {"1": "🔵", "2": "🟢", "3": "🟡"}
        for h in r.get("heading_map", [])[:60]:
            icon   = icons.get(str(h["level"]), "⚪")
            indent = "&nbsp;" * (h["level"] - 1) * 6
            st.markdown(
                f'<div class="hmap">{indent}{icon} <strong>H{h["level"]}</strong>&nbsp; {h["text"]}</div>',
                unsafe_allow_html=True,
            )
    else:
        st.info("Nessun heading trovato.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — AREA 3: ROBOTS.TXT
# ══════════════════════════════════════════════════════════════════════════════

with tabs[5]:
    r = r3f; ov = r.get("overall", "ERROR")
    ct, cb = st.columns([7, 1])
    ct.subheader("Robots.txt")
    cb.markdown(f"<div style='padding-top:10px'>{badge(ov)}</div>", unsafe_allow_html=True)
    st.caption(r.get("summary", "").replace(" | ", " · "))
    st.caption(f"File: `{r.get('robots_url', '—')}`")
    st.divider()

    c1, c2, c3 = st.columns(3)
    c1.metric("Direttive non supportate", r.get("unsupported_count", 0))
    c2.metric("Typo Disallow",            r.get("typos_count", 0))
    c3.metric("Sitemap dichiarata",       "Sì" if r.get("sitemap_declared") else "No")

    if r.get("findings"):
        df_f = pd.DataFrame([{
            "Severity":  f.get("severity", ""),
            "Tipo":      f.get("type", ""),
            "Direttiva": f.get("directive", ""),
            "Linea":     str(f.get("line", "—")),
            "Nota":      f.get("note", "")[:120],
        } for f in r["findings"]])
        st.dataframe(
            df_f.style.applymap(color_cell, subset=["Severity"]),
            use_container_width=True, hide_index=True,
        )
    else:
        st.success("✅ Nessun problema rilevato.")

    if r.get("robots_content"):
        with st.expander("Contenuto robots.txt"):
            st.code(r["robots_content"][:3000], language="text")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — AREA 8: E-E-A-T
# ══════════════════════════════════════════════════════════════════════════════

with tabs[6]:
    r = r8f; ov = r.get("overall", "ERROR")
    ct, cb = st.columns([7, 1])
    ct.subheader("E-E-A-T Signals")
    cb.markdown(f"<div style='padding-top:10px'>{badge(ov)}</div>", unsafe_allow_html=True)
    st.caption(r.get("summary", "").replace(" | ", " · "))
    if r.get("is_ymyl"):
        st.error("⚠️ **Pagina YMYL rilevata** — Google e i sistemi AI applicano standard più severi per finanza, salute e diritto.")
    st.divider()

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Score",       f"{r.get('score', 0)}/100")
    c2.metric("Autori",      len(r.get("author_found", [])))
    c3.metric("Bio autore",  "Sì" if r.get("bio_found") else "No")
    c4.metric("Data pubbl.", "Sì" if r.get("date_pub") else "No")
    c5.metric("Schema Org.", "Sì" if r.get("org_schema") else "No")
    c6.metric("Fonti est.",  len(r.get("auth_links", [])))

    col_pos, col_miss = st.columns(2)
    with col_pos:
        st.markdown("**✅ Segnali positivi**")
        for s in r.get("signals", []):
            st.markdown(f'<div class="sig-ok">✓ {s}</div>', unsafe_allow_html=True)
        if not r.get("signals"):
            st.caption("Nessun segnale positivo.")
    with col_miss:
        st.markdown("**⚠️ Da migliorare**")
        for m in r.get("missing", []):
            st.markdown(f'<div class="sig-warn">→ {m}</div>', unsafe_allow_html=True)
        if not r.get("missing"):
            st.success("Nessuna criticità.")

    if r.get("auth_links"):
        st.divider()
        st.markdown("**Link verso fonti autorevoli:**")
        st.dataframe(
            pd.DataFrame([{"URL": l["url"][:100], "Testo": l["text"]} for l in r["auth_links"]]),
            use_container_width=True, hide_index=True,
        )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — AREA 9: PERFORMANCE
# ══════════════════════════════════════════════════════════════════════════════

with tabs[7]:
    r = r9f; ov = r.get("overall", "ERROR")
    ct, cb = st.columns([7, 1])
    ct.subheader("Performance Signals")
    cb.markdown(f"<div style='padding-top:10px'>{badge(ov)}</div>", unsafe_allow_html=True)
    st.caption(r.get("summary", "").replace(" | ", " · "))
    psi_url = r.get("url", audit_url)
    st.markdown(
        f'<div class="callout">ℹ Segnali dal DOM statico — misurazioni LCP/INP/CLS: '
        f'<a href="https://pagespeed.web.dev/analysis?url={psi_url}" target="_blank">'
        f'PageSpeed Insights →</a></div>',
        unsafe_allow_html=True,
    )
    st.divider()

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Score stimato",  f"{r.get('score', 0)}/100")
    c2.metric("Img senza dim.", r.get("imgs_no_size", 0))
    c3.metric("Img senza alt",  r.get("imgs_no_alt", 0))
    c4.metric("Script blocc.",  r.get("blocking_scripts", 0))
    c5.metric("Script esterni", r.get("ext_scripts", 0))
    c6.metric("Preconnect",     r.get("preconnect_count", 0))

    if r.get("frameworks"):
        st.info(f"🔧 Framework JS: **{', '.join(r['frameworks'])}**")

    pf = [f for f in r.get("findings", []) if isinstance(f, dict) and "impatto" in f]
    if pf:
        st.markdown("**Findings per metrica Core Web Vitals:**")
        st.dataframe(
            pd.DataFrame([{
                "Severity":        f.get("severity", ""),
                "Problema":        f.get("tipo", ""),
                "Metrica CWV":     f.get("impatto", ""),
                "Raccomandazione": f.get("nota", "")[:130],
            } for f in pf]).style.applymap(color_cell, subset=["Severity"]),
            use_container_width=True, hide_index=True,
        )

    if r.get("positives"):
        st.markdown("**✅ Elementi ottimizzati:**")
        for p in r["positives"]:
            st.markdown(f'<div class="sig-ok">✓ {p}</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 8 — AREA 10: TOPICAL AUTHORITY
# ══════════════════════════════════════════════════════════════════════════════

with tabs[8]:
    r = r10f; ov = r.get("overall", "ERROR")
    ct, cb = st.columns([7, 1])
    ct.subheader("Autorevolezza Topica")
    cb.markdown(f"<div style='padding-top:10px'>{badge(ov)}</div>", unsafe_allow_html=True)
    st.caption(r.get("summary", "").replace(" | ", " · "))
    st.divider()

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Score",            f"{r.get('score', 0)}/100")
    c2.metric("Title (car.)",     r.get("title_len", 0))
    c3.metric("Meta desc (car.)", r.get("desc_len", 0))
    c4.metric("Canonical",        "Sì" if r.get("has_canonical") else "No")
    c5.metric("Breadcrumb",       "Sì" if r.get("has_breadcrumb") else "No")
    c6.metric("Parole",           r.get("word_count", 0))

    if r.get("title_text"):
        st.markdown(f"**Title:** `{r['title_text'][:80]}`")
        tl = r.get("title_len", 0)
        if tl < 30:    st.warning(f"Troppo corto ({tl} car.) — ottimale 50–60")
        elif tl > 65:  st.warning(f"Troppo lungo ({tl} car.) — troncato in SERP")
        else:           st.success(f"Lunghezza ottimale ({tl} car.)")

    if r.get("desc_text"):
        st.markdown(f"**Meta description:** _{r['desc_text'][:200]}_")
        dl = r.get("desc_len", 0)
        if dl < 70:    st.warning(f"Corta ({dl} car.) — ottimale 120–160")
        elif dl > 165: st.warning(f"Lunga ({dl} car.) — troncata in SERP")
        else:           st.success(f"Lunghezza ottimale ({dl} car.)")

    st.divider()
    col_str, col_opp = st.columns(2)
    with col_str:
        st.markdown("**✅ Punti di forza**")
        for s in r.get("signals", []):
            st.markdown(f'<div class="sig-ok">✓ {s}</div>', unsafe_allow_html=True)
        if not r.get("signals"):
            st.caption("—")
    with col_opp:
        st.markdown("**💡 Opportunità**")
        for o in r.get("opportunities", []):
            st.markdown(f'<div class="sig-warn">→ {o}</div>', unsafe_allow_html=True)
        if not r.get("opportunities"):
            st.success("Nessuna opportunità critica.")

    if r.get("h2_list"):
        st.divider()
        st.markdown("**Sezioni H2 della pagina:**")
        for h in r.get("h2_list", []):
            st.markdown(
                f'<div style="padding:6px 14px;border-left:3px solid #4f46e5;margin-bottom:4px;'
                f'font-size:14px;background:#f5f3ff;border-radius:0 6px 6px 0">{h}</div>',
                unsafe_allow_html=True,
            )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 9 — ESPORTA
# ══════════════════════════════════════════════════════════════════════════════

with tabs[9]:
    st.subheader("Esporta risultati")
    st.caption("Genera i report e scaricali direttamente dal browser.")
    st.divider()

    col_h, col_x = st.columns(2)

    with col_h:
        st.markdown("**📄 Report HTML**")
        st.caption("Dashboard interattiva navigabile. Ideale per la presentazione al cliente.")
        export_download(
            "HTML",
            lambda p: sa.export_html(audit_url, grouped, p),
            ".html",
            "text/html",
        )

    with col_x:
        st.markdown("**📊 Report Excel**")
        st.caption("File multi-foglio con riepilogo, sintesi, pagine, problemi e dettagli.")
        export_download(
            "Excel",
            lambda p: sa.export_excel(audit_url, grouped, p),
            ".xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.divider()
    st.markdown("**📋 Dati grezzi JSON**")
    st.caption("Per integrazione con altri tool o archiviazione strutturata.")
    try:
        json_data = json.dumps(
            [[{k: v for k, v in r.items() if k != "robots_content"} for r in pg]
             for pg in grouped],
            ensure_ascii=False, indent=2, default=str,
        ).encode("utf-8")
        st.download_button(
            "⬇️ Scarica JSON",
            data=json_data,
            file_name=f"seo_audit_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
            mime="application/json",
        )
    except Exception as exc:
        st.error(f"Errore JSON: {exc}")
