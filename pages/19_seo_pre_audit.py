"""
SEO Audit App — NVL Agency
Interfaccia Streamlit per seo_audit.py

Deploy su Streamlit Cloud:
  1. Repository con seo_audit.py e seo_audit_app.py nella stessa cartella
  2. requirements.txt con tutte le dipendenze (senza playwright — opzionale in locale)
  3. Main file: seo_audit_app.py

Locale:
  streamlit run seo_audit_app.py
"""

import streamlit as st
import importlib
import sys
import os
import json
import time
import tempfile
import pandas as pd
from datetime import datetime

# ─── IMPORT ROBUSTO DI seo_audit ─────────────────────────────────────────────
# Su Streamlit Cloud __file__ può essere instabile; aggiungiamo il CWD come fallback.

for _p in [os.path.dirname(os.path.abspath(__file__)), os.getcwd()]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

try:
    import seo_audit as sa
    AUDIT_OK = True
    AUDIT_ERR = ""
except Exception as _e:
    AUDIT_OK = False
    AUDIT_ERR = str(_e)

# Playwright disponibile? Su Cloud di solito no.
try:
    from playwright.sync_api import sync_playwright as _spw  # noqa
    PLAYWRIGHT_OK = True
except Exception:
    PLAYWRIGHT_OK = False


# ─── COSTANTI ────────────────────────────────────────────────────────────────

SC  = {"OK": 100, "WARN": 50, "FAIL": 0, "ERROR": 0}
BC  = {"OK": "#16a34a", "WARN": "#b45309", "FAIL": "#dc2626", "ERROR": "#6b7280"}
BGC = {"OK": "#d1fae5", "WARN": "#fef3c7", "FAIL": "#fee2e2", "ERROR": "#f1f5f9"}
BL  = {"OK": "OK", "WARN": "WARN", "FAIL": "FAIL", "ERROR": "N/D"}

AREA_ORDER = ["1.", "2.", "3.", "8.", "9.", "10."]
AREA_META  = {
    "1.":  ("1",  "Rendering & Visibilità",  "Contenuti nascosti, lazy-load, delta DOM"),
    "2.":  ("2",  "Struttura Heading",        "H1, H2/H3, gerarchia, blocchi orfani"),
    "3.":  ("3",  "Robots.txt",               "Direttive, sitemap, configurazione crawler"),
    "8.":  ("8",  "E-E-A-T Signals",          "Autore, fonti, schema, YMYL detection"),
    "9.":  ("9",  "Performance",              "CLS, LCP, script bloccanti, framework JS"),
    "10.": ("10", "Topical Authority",        "Title, meta desc, canonical, citazioni, link"),
}


# ─── HELPERS ─────────────────────────────────────────────────────────────────

def r_by(results, prefix):
    """Recupera il risultato di una specifica area dai risultati di una pagina."""
    return next((r for r in results if r.get("area", "").startswith(prefix)), {})

def page_score(results):
    scores = [SC.get(r.get("overall", "ERROR"), 0) for r in results]
    return round(sum(scores) / len(scores)) if scores else 0

def global_score(grouped):
    scores = [page_score(pg) for pg in grouped]
    return round(sum(scores) / len(scores)) if scores else 0

def badge(overall):
    """Ritorna una stringa HTML con badge colorato."""
    color = BC.get(overall, "#6b7280")
    label = BL.get(overall, overall)
    return (f'<span style="background:{color};color:white;padding:2px 10px;'
            f'border-radius:4px;font-size:12px;font-weight:700;'
            f'font-family:monospace">{label}</span>')

def color_cell(val):
    """Colorazione celle per st.dataframe.style."""
    colors = {
        "OK":   "background-color:#d1fae5;color:#065f46;font-weight:700",
        "WARN": "background-color:#fef3c7;color:#78350f;font-weight:700",
        "FAIL": "background-color:#fee2e2;color:#7f1d1d;font-weight:700",
    }
    return colors.get(val, "")

def score_color(s):
    return BC["OK"] if s >= 80 else (BC["WARN"] if s >= 50 else BC["FAIL"])


# ─── CONFIG PAGINA ───────────────────────────────────────────────────────────

st.set_page_config(
    page_title="SEO Audit — NVL Agency",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# CSS minimale — niente Google Fonts, niente layout override pesanti
st.markdown("""
<style>
/* Badge e segnali */
.sig-ok   { color: #065f46; padding: 4px 0; font-size: 14px; border-bottom: 1px solid #f0fdf4; }
.sig-warn { color: #78350f; padding: 4px 0; font-size: 14px; border-bottom: 1px solid #fffbeb; }
.callout  { background:#eff6ff; border:1px solid #bfdbfe; border-radius:8px;
            padding:12px 16px; font-size:13px; color:#1e40af; margin-bottom:12px; }
.callout-warn { background:#fffbeb; border:1px solid #fde68a; border-radius:8px;
                padding:12px 16px; font-size:13px; color:#78350f; margin-bottom:12px; }
/* Area card in overview */
.area-card { background:white; border:1px solid #e5e7eb; border-radius:10px;
             padding:16px 18px; margin-bottom:12px; border-top:4px solid #e5e7eb; }
/* Heading map */
.hmap { padding:3px 0; font-size:13px; border-bottom:1px solid #f9fafb; }
</style>
""", unsafe_allow_html=True)


# ─── SIDEBAR ─────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("🔍 SEO Audit")
    st.caption("NVL Agency · Tool interno")
    st.divider()

    # ── Modalità input
    input_mode = st.radio(
        "Modalità input",
        ["URL singola", "File con URL (CSV / Excel)"],
        help="CSV/Excel deve avere una colonna chiamata 'url'",
    )

    urls_to_audit: list[str] = []

    if input_mode == "URL singola":
        raw_url = st.text_input(
            "URL da analizzare",
            placeholder="https://tuosito.com/pagina",
        )
        if raw_url.strip():
            urls_to_audit = [raw_url.strip()]

    else:
        uploaded = st.file_uploader(
            "Carica CSV o Excel",
            type=["csv", "xlsx", "xls"],
        )
        if uploaded:
            try:
                df_up = (pd.read_csv(uploaded) if uploaded.name.endswith(".csv")
                         else pd.read_excel(uploaded))
                col_url = next(
                    (c for c in df_up.columns
                     if c.lower() in ["url", "urls", "link", "links", "pagina", "page"]),
                    None,
                )
                if col_url:
                    urls_to_audit = (
                        df_up[col_url].dropna().astype(str).str.strip()
                        .pipe(lambda s: s[s.str.startswith("http")])
                        .tolist()
                    )
                    st.success(f"✅ {len(urls_to_audit)} URL trovati")
                else:
                    st.error("Colonna 'url' non trovata nel file.")
            except Exception as exc:
                st.error(f"Errore lettura file: {exc}")

    st.divider()

    # ── Opzioni
    st.markdown("**Opzioni**")

    use_playwright = st.toggle(
        "Rendering JavaScript (Playwright)",
        value=False,
        disabled=not PLAYWRIGHT_OK,
        help=(
            "Richiede Playwright installato. Attiva su server locale."
            if not PLAYWRIGHT_OK
            else "Lancia Chromium headless per rilevare contenuto iniettato via JS. Più lento."
        ),
    )
    if not PLAYWRIGHT_OK:
        st.caption("⚠️ Playwright non disponibile in questo ambiente — solo fetch statico.")

    max_urls = 10
    if input_mode != "URL singola" and urls_to_audit:
        max_urls = st.slider("Max URL da analizzare", 1, 50, 10)
        urls_to_audit = urls_to_audit[:max_urls]

    st.divider()

    run_btn = st.button(
        "▶ Avvia analisi",
        type="primary",
        use_container_width=True,
        disabled=(not urls_to_audit or not AUDIT_OK),
    )

    if not AUDIT_OK:
        st.error(f"seo_audit.py non trovato.\n\n`{AUDIT_ERR}`")

    st.divider()
    st.caption(f"seo_audit.py {'✅ caricato' if AUDIT_OK else '❌ non disponibile'}")
    st.caption(f"Playwright {'✅' if PLAYWRIGHT_OK else '❌ non disponibile'}")


# ─── ESECUZIONE AUDIT ────────────────────────────────────────────────────────

if run_btn and urls_to_audit and AUDIT_OK:
    st.session_state["grouped"] = []
    st.session_state["audit_url"] = urls_to_audit[0]
    st.session_state["audit_ts"]  = datetime.now().strftime("%d/%m/%Y %H:%M")

    total = len(urls_to_audit)
    pbar  = st.progress(0, text="Avvio…")
    info  = st.empty()

    for idx, url in enumerate(urls_to_audit):
        info.info(f"🔍 [{idx+1}/{total}] {url}")
        page_res = []

        try:
            session = sa.get_session()
            static_html, _, final_url = sa.fetch_html_static(url, session)

            if not static_html:
                info.error(f"❌ {url} — impossibile recuperare ({final_url})")
                pbar.progress((idx+1)/total)
                continue

            rendered_html = None
            if use_playwright and PLAYWRIGHT_OK:
                rendered_html = sa.fetch_html_rendered(url)

            audit_steps = [
                ("1.", lambda: sa.audit_content_visibility(url, static_html, rendered_html)),
                ("2.", lambda: sa.audit_heading_structure(url, static_html)),
                ("3.", lambda: sa.audit_robots_txt(url, session)),
                ("8.", lambda: sa.audit_eeat_signals(url, static_html)),
                ("9.", lambda: sa.audit_performance_signals(url, static_html)),
                ("10.",lambda: sa.audit_topical_authority(url, static_html, session)),
            ]

            for pfx, fn in audit_steps:
                name = AREA_META[pfx][1]
                pbar.progress(
                    (idx + audit_steps.index((pfx, fn)) / len(audit_steps)) / total,
                    text=f"[{idx+1}/{total}] Area {pfx.rstrip('.')} — {name}",
                )
                r = fn()
                r["url"] = url
                page_res.append(r)

            st.session_state["grouped"].append(page_res)

        except Exception as exc:
            info.error(f"❌ Errore su {url}: {exc}")

        pbar.progress((idx+1)/total, text=f"Completati {idx+1}/{total}")

    pbar.progress(1.0, text="✅ Analisi completata")
    time.sleep(0.6)
    pbar.empty()
    info.empty()
    st.rerun()


# ─── WELCOME (nessun dato) ────────────────────────────────────────────────────

if "grouped" not in st.session_state or not st.session_state["grouped"]:
    st.header("SEO Audit Tool")
    st.markdown(
        "Inserisci un URL o carica un file nella barra laterale, poi premi **Avvia analisi**."
    )
    st.divider()

    st.subheader("Aree analizzate")
    cols = st.columns(3)
    for i, (pfx, (num, name, desc)) in enumerate(AREA_META.items()):
        with cols[i % 3]:
            st.markdown(f"**Area {num} — {name}**")
            st.caption(desc)
            st.markdown("")
    st.stop()


# ─── DATI SESSIONE ───────────────────────────────────────────────────────────

grouped   = st.session_state["grouped"]
audit_url = st.session_state.get("audit_url", "")
audit_ts  = st.session_state.get("audit_ts", "")
npages    = len(grouped)
gs        = global_score(grouped)
gc        = score_color(gs)

first_pg  = grouped[0]
r1f  = r_by(first_pg, "1.")
r2f  = r_by(first_pg, "2.")
r3f  = r_by(first_pg, "3.")
r8f  = r_by(first_pg, "8.")
r9f  = r_by(first_pg, "9.")
r10f = r_by(first_pg, "10.")


# ─── NAVIGAZIONE A TAB ───────────────────────────────────────────────────────

tab_names = [
    "📊 Overview", "📋 Pagine", "⚠️ Problemi",
    "1 Visibilità", "2 Heading", "3 Robots.txt",
    "8 E-E-A-T", "9 Performance", "10 Topical",
    "⬇️ Esporta",
]
tabs = st.tabs(tab_names)


# ════════════════════════════════════════════════════════════════════════════
# TAB 0 — OVERVIEW
# ════════════════════════════════════════════════════════════════════════════

with tabs[0]:
    st.caption(f"**{audit_url}** · {audit_ts} · {npages} {'pagina' if npages==1 else 'pagine'}")

    # ── Metriche globali
    c_score, c_fail, c_warn, c_ok, c_pages = st.columns([2,1,1,1,1])

    # Conta problemi aggregati
    n_fail = sum(
        1 for pg in grouped for r in pg
        if r.get("overall") == "FAIL"
    )
    n_warn = sum(
        1 for pg in grouped for r in pg
        if r.get("overall") == "WARN"
    )
    n_ok = sum(
        1 for pg in grouped for r in pg
        if r.get("overall") == "OK"
    )

    with c_score:
        st.metric("Score globale", f"{gs} / 100")
        st.progress(gs / 100)
    c_fail.metric("FAIL",  n_fail,  delta=None)
    c_warn.metric("WARN",  n_warn,  delta=None)
    c_ok.metric("OK",    n_ok,    delta=None)
    c_pages.metric("Pagine", npages, delta=None)

    st.divider()
    st.subheader("Esito per area — prima URL")

    cols_a = st.columns(3)
    for i, pfx in enumerate(AREA_ORDER):
        r = r_by(first_pg, pfx)
        ov  = r.get("overall", "ERROR")
        num, name, desc = AREA_META[pfx]
        sc  = SC.get(ov, 0)
        bc_ = BC.get(ov, "#6b7280")
        sum_text = r.get("summary", "—").split(" | ")[0][:140]

        with cols_a[i % 3]:
            st.markdown(
                f'<div class="area-card" style="border-top-color:{bc_}">'
                f'<div style="font-size:11px;color:#6b7280;margin-bottom:2px">AREA {num}</div>'
                f'<div style="font-size:15px;font-weight:700;margin-bottom:6px">{name}</div>'
                f'{badge(ov)}'
                f'<div style="font-size:13px;color:#4b5563;margin-top:10px">{sum_text}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )


# ════════════════════════════════════════════════════════════════════════════
# TAB 1 — PAGINE ANALIZZATE
# ════════════════════════════════════════════════════════════════════════════

with tabs[1]:
    st.subheader(f"Pagine analizzate — {npages} URL")

    rows = []
    for pg in grouped:
        url_pg = pg[0].get("url", "—") if pg else "—"
        ovs = {}
        for r in pg:
            for pfx in AREA_ORDER:
                if r.get("area","").startswith(pfx):
                    ovs[pfx] = r.get("overall","ERR")
        sc = page_score(pg)
        rows.append({
            "URL":         url_pg,
            "Visibilità":  ovs.get("1.", "N/D"),
            "Heading":     ovs.get("2.", "N/D"),
            "Robots":      ovs.get("3.", "N/D"),
            "E-E-A-T":     ovs.get("8.", "N/D"),
            "Perf.":       ovs.get("9.", "N/D"),
            "Topical":     ovs.get("10.","N/D"),
            "Score":       sc,
        })

    df_pages = pd.DataFrame(rows)
    badge_cols = ["Visibilità","Heading","Robots","E-E-A-T","Perf.","Topical"]
    st.dataframe(
        df_pages.style
            .applymap(color_cell, subset=badge_cols)
            .background_gradient(subset=["Score"], cmap="RdYlGn", vmin=0, vmax=100),
        use_container_width=True,
        height=min(600, 80 + 38 * len(df_pages)),
    )


# ════════════════════════════════════════════════════════════════════════════
# TAB 2 — PROBLEMI
# ════════════════════════════════════════════════════════════════════════════

with tabs[2]:
    probs = []

    for pg in grouped:
        u = pg[0].get("url","—") if pg else "—"
        for r in pg:
            area = r.get("area","")[:30]
            ov   = r.get("overall","")

            # Area 1 — contenuto nascosto
            for b in r.get("hidden_blocks",[]):
                probs.append({"Area":area,"URL":u,"Tipo":"Contenuto nascosto","Severity":"WARN",
                              "Dettaglio":f'{b["selector"]} → {b["text_preview"][:80]}'})
            if r.get("lazy_elements",0):
                probs.append({"Area":area,"URL":u,"Tipo":"Lazy-load","Severity":"WARN",
                              "Dettaglio":f'{r["lazy_elements"]} elementi lazy-load rilevati'})
            for pat in r.get("js_patterns",[]):
                probs.append({"Area":area,"URL":u,"Tipo":"Pattern JS","Severity":"WARN",
                              "Dettaglio":pat})

            # Area 2 — heading
            if r.get("h1_count") == 0:
                probs.append({"Area":area,"URL":u,"Tipo":"H1 assente","Severity":"FAIL",
                              "Dettaglio":"Nessun H1 trovato nella pagina"})
            elif r.get("h1_count",1) > 1:
                probs.append({"Area":area,"URL":u,"Tipo":"H1 multipli","Severity":"WARN",
                              "Dettaglio":f'{r["h1_count"]} H1 trovati — deve essere uno solo'})
            for jmp in r.get("level_jumps",[]):
                probs.append({"Area":area,"URL":u,"Tipo":"Salto heading","Severity":"WARN",
                              "Dettaglio":jmp})

            # Area 3 — robots
            for f in r.get("findings",[]):
                if isinstance(f,dict) and f.get("severity") in ("WARN","FAIL"):
                    probs.append({"Area":area,"URL":u,"Tipo":f.get("type","—"),
                                  "Severity":f["severity"],
                                  "Dettaglio":f.get("note","")[:120]})

            # Area 8 — E-E-A-T
            for m in r.get("missing",[]):
                probs.append({"Area":area,"URL":u,"Tipo":"E-E-A-T mancante","Severity":"WARN",
                              "Dettaglio":m})
            if r.get("is_ymyl") and r.get("score",100) < 50:
                probs.append({"Area":area,"URL":u,"Tipo":"YMYL + E-E-A-T basso","Severity":"FAIL",
                              "Dettaglio":f'Score {r.get("score",0)}/100 su pagina YMYL'})

            # Area 9 — performance
            for f in r.get("findings",[]):
                if isinstance(f,dict) and "impatto" in f and f.get("severity") in ("WARN","FAIL"):
                    probs.append({"Area":area,"URL":u,"Tipo":f.get("tipo","—"),
                                  "Severity":f["severity"],
                                  "Dettaglio":f'{f.get("nota","")[:100]} [{f.get("impatto","")}]'})

            # Area 10 — topical
            for opp in r.get("opportunities",[]):
                probs.append({"Area":area,"URL":u,"Tipo":"Opportunità","Severity":"WARN",
                              "Dettaglio":opp[:120]})

    # Ordina FAIL → WARN → INFO
    probs.sort(key=lambda x: {"FAIL":0,"WARN":1,"INFO":2}.get(x["Severity"],9))

    st.subheader(f"Problemi rilevati — {len(probs)} finding")

    if probs:
        # Filtro rapido per severity
        sev_filter = st.multiselect(
            "Filtra per severity",
            ["FAIL","WARN","INFO"],
            default=["FAIL","WARN"],
        )
        df_probs = pd.DataFrame(probs)
        df_probs = df_probs[df_probs["Severity"].isin(sev_filter)]

        st.dataframe(
            df_probs.style.applymap(color_cell, subset=["Severity"]),
            use_container_width=True,
            height=min(700, 80 + 38 * len(df_probs)),
        )
    else:
        st.success("✅ Nessun problema rilevato.")


# ════════════════════════════════════════════════════════════════════════════
# TAB 3 — AREA 1: VISIBILITÀ
# ════════════════════════════════════════════════════════════════════════════

with tabs[3]:
    r = r1f
    ov = r.get("overall","ERROR")

    col_title, col_badge = st.columns([6,1])
    col_title.subheader("Rendering & Visibilità Contenuti")
    col_badge.markdown(f"<div style='padding-top:8px'>{badge(ov)}</div>", unsafe_allow_html=True)

    st.caption(r.get("summary","").replace(" | ", " · "))
    st.divider()

    c1, c2, c3 = st.columns(3)
    c1.metric("Blocchi nascosti",  len(r.get("hidden_blocks",[])))
    c2.metric("Elementi lazy-load", r.get("lazy_elements",0))
    c3.metric("Pattern JS",         len(r.get("js_patterns",[])))

    st.caption(f"Delta DOM statico / renderizzato: `{r.get('dom_delta','N/A')}`")

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
        } for b in r.get("hidden_blocks",[])])
        st.dataframe(df_hb, use_container_width=True)
    else:
        st.success("✅ Nessun blocco contenuto nascosto rilevato.")


# ════════════════════════════════════════════════════════════════════════════
# TAB 4 — AREA 2: HEADING
# ════════════════════════════════════════════════════════════════════════════

with tabs[4]:
    r = r2f
    ov = r.get("overall","ERROR")

    col_t, col_b = st.columns([6,1])
    col_t.subheader("Struttura Heading H2/H3")
    col_b.markdown(f"<div style='padding-top:8px'>{badge(ov)}</div>", unsafe_allow_html=True)

    st.caption(r.get("summary","").replace(" | ", " · "))
    st.divider()

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("H1",            r.get("h1_count",0))
    c2.metric("H2",            r.get("h2_count",0))
    c3.metric("H3",            r.get("h3_count",0))
    c4.metric("Salti livello", len(r.get("level_jumps",[])))
    c5.metric("Blocchi orfani",len(r.get("orphan_blocks",[])))

    if r.get("level_jumps"):
        st.warning("**Salti di livello:**\n" + "\n".join(f"- {j}" for j in r["level_jumps"]))

    if r.get("heading_map"):
        st.markdown("**Mappa heading:**")
        level_icon = {"1":"🔵","2":"🟢","3":"🟡"}
        for h in r.get("heading_map",[])[:60]:
            icon   = level_icon.get(str(h["level"]),"⚪")
            indent = "&nbsp;" * (h["level"]-1) * 6
            st.markdown(
                f'<div class="hmap">{indent}{icon} <strong>H{h["level"]}</strong> &nbsp;{h["text"]}</div>',
                unsafe_allow_html=True,
            )
    else:
        st.info("Nessun heading trovato nella pagina.")


# ════════════════════════════════════════════════════════════════════════════
# TAB 5 — AREA 3: ROBOTS.TXT
# ════════════════════════════════════════════════════════════════════════════

with tabs[5]:
    r = r3f
    ov = r.get("overall","ERROR")

    col_t, col_b = st.columns([6,1])
    col_t.subheader("Robots.txt")
    col_b.markdown(f"<div style='padding-top:8px'>{badge(ov)}</div>", unsafe_allow_html=True)

    st.caption(r.get("summary","").replace(" | ", " · "))
    st.caption(f"File: `{r.get('robots_url','—')}`")
    st.divider()

    c1,c2,c3 = st.columns(3)
    c1.metric("Direttive non supportate", r.get("unsupported_count",0))
    c2.metric("Typo Disallow",            r.get("typos_count",0))
    c3.metric("Sitemap dichiarata",       "Sì" if r.get("sitemap_declared") else "No")

    findings = r.get("findings",[])
    if findings:
        st.markdown("**Findings:**")
        df_f = pd.DataFrame([{
            "Severity":  f.get("severity",""),
            "Tipo":      f.get("type",""),
            "Direttiva": f.get("directive",""),
            "Linea":     str(f.get("line","—")),
            "Nota":      f.get("note","")[:120],
        } for f in findings])
        st.dataframe(
            df_f.style.applymap(color_cell, subset=["Severity"]),
            use_container_width=True,
        )
    else:
        st.success("✅ Nessun problema rilevato nel robots.txt.")

    if r.get("robots_content"):
        with st.expander("Contenuto robots.txt"):
            st.code(r["robots_content"][:3000], language="text")


# ════════════════════════════════════════════════════════════════════════════
# TAB 6 — AREA 8: E-E-A-T
# ════════════════════════════════════════════════════════════════════════════

with tabs[6]:
    r = r8f
    ov = r.get("overall","ERROR")

    col_t, col_b = st.columns([6,1])
    col_t.subheader("E-E-A-T Signals")
    col_b.markdown(f"<div style='padding-top:8px'>{badge(ov)}</div>", unsafe_allow_html=True)

    st.caption(r.get("summary","").replace(" | ", " · "))

    if r.get("is_ymyl"):
        st.error(
            "⚠️ **Pagina YMYL rilevata** — E-E-A-T è critico per questo settore. "
            "Google e i sistemi AI applicano standard più severi per contenuti su finanza, salute e diritto."
        )

    st.divider()

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("Score",        f"{r.get('score',0)}/100")
    c2.metric("Autori",       len(r.get("author_found",[])))
    c3.metric("Bio autore",   "Sì" if r.get("bio_found") else "No")
    c4.metric("Data pubbl.",  "Sì" if r.get("date_pub") else "No")
    c5.metric("Schema Org.",  "Sì" if r.get("org_schema") else "No")
    c6.metric("Fonti est.",   len(r.get("auth_links",[])))

    col_pos, col_miss = st.columns(2)

    with col_pos:
        st.markdown("**✅ Segnali positivi**")
        for s in r.get("signals",[]):
            st.markdown(f'<div class="sig-ok">✓ {s}</div>', unsafe_allow_html=True)
        if not r.get("signals"):
            st.caption("Nessun segnale positivo rilevato.")

    with col_miss:
        st.markdown("**⚠️ Da migliorare**")
        for m in r.get("missing",[]):
            st.markdown(f'<div class="sig-warn">→ {m}</div>', unsafe_allow_html=True)
        if not r.get("missing"):
            st.success("Nessuna criticità.")

    if r.get("auth_links"):
        st.divider()
        st.markdown("**Link verso fonti autorevoli:**")
        df_al = pd.DataFrame([{"URL": l["url"][:100], "Testo": l["text"]}
                               for l in r["auth_links"]])
        st.dataframe(df_al, use_container_width=True, hide_index=True)


# ════════════════════════════════════════════════════════════════════════════
# TAB 7 — AREA 9: PERFORMANCE
# ════════════════════════════════════════════════════════════════════════════

with tabs[7]:
    r = r9f
    ov = r.get("overall","ERROR")

    col_t, col_b = st.columns([6,1])
    col_t.subheader("Performance Signals")
    col_b.markdown(f"<div style='padding-top:8px'>{badge(ov)}</div>", unsafe_allow_html=True)

    st.caption(r.get("summary","").replace(" | ", " · "))

    psi_url = r.get("url", audit_url)
    st.markdown(
        f'<div class="callout">ℹ Segnali rilevati dal DOM statico — per LCP/INP/CLS verificati: '
        f'<a href="https://pagespeed.web.dev/analysis?url={psi_url}" target="_blank">'
        f'PageSpeed Insights →</a></div>',
        unsafe_allow_html=True,
    )

    st.divider()

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("Score stimato",   f"{r.get('score',0)}/100")
    c2.metric("Img senza dim.",  r.get("imgs_no_size",0))
    c3.metric("Img senza alt",   r.get("imgs_no_alt",0))
    c4.metric("Script blocc.",   r.get("blocking_scripts",0))
    c5.metric("Script esterni",  r.get("ext_scripts",0))
    c6.metric("Preconnect",      r.get("preconnect_count",0))

    if r.get("frameworks"):
        st.info(f"🔧 Framework JS rilevati: **{', '.join(r['frameworks'])}**")

    perf_findings = [f for f in r.get("findings",[]) if isinstance(f,dict) and "impatto" in f]
    if perf_findings:
        st.markdown("**Findings per metrica Core Web Vitals:**")
        df_pf = pd.DataFrame([{
            "Severity":  f.get("severity",""),
            "Problema":  f.get("tipo",""),
            "Metrica":   f.get("impatto",""),
            "Raccomandazione": f.get("nota","")[:130],
        } for f in perf_findings])
        st.dataframe(
            df_pf.style.applymap(color_cell, subset=["Severity"]),
            use_container_width=True,
            hide_index=True,
        )

    positives = r.get("positives",[])
    if positives:
        st.markdown("**✅ Elementi ottimizzati:**")
        for p in positives:
            st.markdown(f'<div class="sig-ok">✓ {p}</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# TAB 8 — AREA 10: TOPICAL AUTHORITY
# ════════════════════════════════════════════════════════════════════════════

with tabs[8]:
    r = r10f
    ov = r.get("overall","ERROR")

    col_t, col_b = st.columns([6,1])
    col_t.subheader("Autorevolezza Topica")
    col_b.markdown(f"<div style='padding-top:8px'>{badge(ov)}</div>", unsafe_allow_html=True)

    st.caption(r.get("summary","").replace(" | ", " · "))
    st.divider()

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("Score",          f"{r.get('score',0)}/100")
    c2.metric("Title (car.)",   r.get("title_len",0))
    c3.metric("Meta desc (car.)",r.get("desc_len",0))
    c4.metric("Canonical",      "Sì" if r.get("has_canonical") else "No")
    c5.metric("Breadcrumb",     "Sì" if r.get("has_breadcrumb") else "No")
    c6.metric("Parole",         r.get("word_count",0))

    # Title
    if r.get("title_text"):
        st.markdown(f"**Title:** `{r['title_text'][:80]}`")
        tl = r.get("title_len",0)
        if tl < 30:    st.warning(f"Troppo corto ({tl} car.) — ottimale 50–60")
        elif tl > 65:  st.warning(f"Troppo lungo ({tl} car.) — verrà troncato in SERP")
        else:           st.success(f"Lunghezza ottimale ({tl} car.)")

    # Meta description
    if r.get("desc_text"):
        st.markdown(f"**Meta description:** _{r['desc_text'][:200]}_")
        dl = r.get("desc_len",0)
        if dl < 70:    st.warning(f"Corta ({dl} car.) — ottimale 120–160")
        elif dl > 165: st.warning(f"Lunga ({dl} car.) — verrà troncata in SERP")
        else:           st.success(f"Lunghezza ottimale ({dl} car.)")

    st.divider()

    col_str, col_opp = st.columns(2)
    with col_str:
        st.markdown("**✅ Punti di forza**")
        for s in r.get("signals",[]):
            st.markdown(f'<div class="sig-ok">✓ {s}</div>', unsafe_allow_html=True)
        if not r.get("signals"):
            st.caption("—")

    with col_opp:
        st.markdown("**💡 Opportunità**")
        for o in r.get("opportunities",[]):
            st.markdown(f'<div class="sig-warn">→ {o}</div>', unsafe_allow_html=True)
        if not r.get("opportunities"):
            st.success("Nessuna opportunità critica.")

    if r.get("h2_list"):
        st.divider()
        st.markdown("**Sezioni H2 della pagina:**")
        for h in r.get("h2_list",[]):
            st.markdown(
                f'<div style="padding:6px 14px;border-left:3px solid #4f46e5;'
                f'margin-bottom:4px;font-size:14px;background:#f5f3ff;border-radius:0 6px 6px 0">'
                f'{h}</div>',
                unsafe_allow_html=True,
            )


# ════════════════════════════════════════════════════════════════════════════
# TAB 9 — ESPORTA
# ════════════════════════════════════════════════════════════════════════════

with tabs[9]:
    st.subheader("Esporta risultati")
    st.caption("Genera i report e scaricali direttamente dal browser.")
    st.divider()

    col_h, col_x = st.columns(2)

    # ── HTML
    with col_h:
        st.markdown("**📄 Report HTML**")
        st.caption("Dashboard interattiva navigabile. Ideale per la presentazione al cliente.")

        if st.button("Genera e scarica HTML", use_container_width=True):
            with st.spinner("Generazione…"):
                try:
                    with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as tmp:
                        tmp_path = tmp.name
                    sa.export_html(audit_url, grouped, tmp_path)
                    html_data = open(tmp_path, "rb").read()
                    os.unlink(tmp_path)
                    fname = f"seo_audit_{datetime.now().strftime('%Y%m%d_%H%M')}.html"
                    st.download_button(
                        "⬇️ Scarica HTML",
                        data=html_data,
                        file_name=fname,
                        mime="text/html",
                        use_container_width=True,
                    )
                except Exception as exc:
                    st.error(f"Errore: {exc}")

    # ── Excel
    with col_x:
        st.markdown("**📊 Report Excel**")
        st.caption("File multi-foglio con riepilogo, sintesi, pagine, problemi e dettagli.")

        if st.button("Genera e scarica Excel", use_container_width=True):
            with st.spinner("Generazione…"):
                try:
                    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                        tmp_path = tmp.name
                    sa.export_excel(audit_url, grouped, tmp_path)
                    xlsx_data = open(tmp_path, "rb").read()
                    os.unlink(tmp_path)
                    fname = f"seo_audit_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                    st.download_button(
                        "⬇️ Scarica Excel",
                        data=xlsx_data,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as exc:
                    st.error(f"Errore: {exc}")

    st.divider()

    # ── JSON
    st.markdown("**📋 Dati grezzi (JSON)**")
    st.caption("Per integrazione con altri tool o archiviazione strutturata.")

    try:
        json_data = json.dumps(
            [
                [{k: v for k, v in r.items() if k != "robots_content"}
                 for r in pg]
                for pg in grouped
            ],
            ensure_ascii=False, indent=2, default=str,
        ).encode("utf-8")

        st.download_button(
            "⬇️ Scarica JSON",
            data=json_data,
            file_name=f"seo_audit_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
            mime="application/json",
        )
    except Exception as exc:
        st.error(f"Errore serializzazione JSON: {exc}")
