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

import sys
import os

import streamlit as st
import json
import time
import tempfile
import pandas as pd
from datetime import datetime

# ─── IMPORT modulo audit ─────────────────────────────────────────────────────
# Il modulo può chiamarsi seo_audit.py oppure seo_pre_audit.py a seconda del
# repo. Prova entrambi i nomi, in entrambe le directory (script dir e cwd).

for _p in [os.path.dirname(os.path.abspath(__file__)), os.getcwd()]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

sa        = None
AUDIT_OK  = False
AUDIT_ERR = ""

def _try_import(name, paths):
    """Importa un modulo per nome o per nome-file (supporta nomi con cifre iniziali)."""
    import importlib, importlib.util
    # 1. prova import diretto (funziona per nomi Python validi)
    try:
        return importlib.import_module(name)
    except ModuleNotFoundError:
        pass
    # 2. prova caricamento da file (funziona anche per "19_seo_pre_audit.py")
    for base in paths:
        candidate = os.path.join(base, name + ".py")
        if os.path.isfile(candidate):
            spec = importlib.util.spec_from_file_location("_sa_module", candidate)
            mod  = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    raise ImportError(f"Modulo '{name}' non trovato")

_search_paths = [os.path.dirname(os.path.abspath(__file__)), os.getcwd()]

for _mod_name in ["seo_pre_audit", "seo_audit"]:
    try:
        sa = _try_import(_mod_name, _search_paths)
        AUDIT_OK  = True
        AUDIT_ERR = ""
        break
    except Exception as _e:
        AUDIT_ERR = str(_e)


# ─── PLAYWRIGHT: rileva e configura il browser ───────────────────────────────
# Su Streamlit Cloud il browser Chromium è installato in /opt/pw-browsers
# tramite packages.txt. La variabile PLAYWRIGHT_BROWSERS_PATH non è sempre
# propagata al processo Python — la impostiamo noi prima di qualsiasi chiamata.

@st.cache_resource(show_spinner=False)
def _detect_playwright():
    """Controlla se Playwright è disponibile nel modulo seo_pre_audit."""
    try:
        if sa and getattr(sa, 'PLAYWRIGHT_AVAILABLE', False):
            return True, "Playwright disponibile"
        return False, "Playwright non disponibile — solo fetch statico attivo"
    except Exception as exc:
        return False, str(exc)

_pw_ok, _pw_log = _detect_playwright()


# ─── CONFIGURAZIONE ──────────────────────────────────────────────────────────
# Tutto qui. Nessuna config in sidebar o form utente.

USE_PLAYWRIGHT: bool = _pw_ok

# Timeout fetch HTTP statico (secondi)
FETCH_TIMEOUT: int = 15

# Numero massimo di URL analizzabili da file (protezione Cloud)
MAX_URLS: int = 100


# ─── COSTANTI UI ─────────────────────────────────────────────────────────────

SC  = {"OK": 100, "WARN": 50, "FAIL": 0, "ERROR": 0}
BC  = {"OK": "#16a34a", "WARN": "#b45309", "FAIL": "#dc2626", "ERROR": "#6b7280"}
BL  = {"OK": "OK",      "WARN": "WARN",    "FAIL": "FAIL",    "ERROR": "N/D"}

AREA_META = {
    "1.":  ("1",  "Rendering & Visibilità",  "Contenuti nascosti, lazy-load, delta DOM"),
    "2.":  ("2",  "Struttura Heading",        "H1, H2/H3, gerarchia, blocchi orfani"),
    "3.":  ("3",  "Robots.txt",               "Direttive, sitemap, configurazione crawler"),
    "4.":  ("4",  "Dati Strutturati",         "Schema.org, JSON-LD, Open Graph, rich results"),
    "8.":  ("8",  "E-E-A-T Signals",          "Autore, fonti, schema, YMYL detection"),
    "9.":  ("9",  "Performance",              "CLS, LCP, script bloccanti, framework JS"),
    "10.": ("10", "Topical Authority",        "Title, meta desc, canonical, citazioni, link"),
}

BADGE_COLS = ["Visibilità", "Heading", "Robots", "Strutt.", "E-E-A-T", "Perf.", "Topical"]


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
    """Genera un file e salva i bytes in session_state per persistenza tra rerun."""
    key_data  = f"_export_data_{label}"
    key_fname = f"_export_fname_{label}"

    if st.button(f"Genera {label}", use_container_width=True):
        with st.spinner(f"Generazione {label}…"):
            try:
                with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                    tmp_path = tmp.name
                generate_fn(tmp_path)
                st.session_state[key_data]  = open(tmp_path, "rb").read()
                st.session_state[key_fname] = f"seo_audit_{datetime.now().strftime('%Y%m%d_%H%M')}{suffix}"
                os.unlink(tmp_path)
            except Exception as exc:
                st.error(f"Errore generazione {label}: {exc}")
                return

    if key_data in st.session_state:
        st.download_button(
            f"⬇️ Scarica {label}",
            data=st.session_state[key_data],
            file_name=st.session_state[key_fname],
            mime=mime,
            use_container_width=True,
        )


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
        st.success("seo_audit.py ✅")
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
    col_sm, col_num, col_bsm = st.columns([5, 2, 1])
    with col_sm:
        sitemap_url = st.text_input(
            "Sitemap URL",
            placeholder="https://tuosito.com/sitemap.xml",
            label_visibility="collapsed",
            key="sitemap_url",
        )
    with col_num:
        sitemap_max = st.number_input(
            "Max URL",
            min_value=1,
            max_value=MAX_URLS,
            value=10,
            step=1,
            help=f"Numero di URL da analizzare dalla sitemap (max {MAX_URLS})",
        )
    with col_bsm:
        run_sitemap = st.button(
            "▶ Analizza", type="primary",
            use_container_width=True,
            key="run_sitemap",
            disabled=not AUDIT_OK,
        )

    if sitemap_url.strip():
        col_prev, _ = st.columns([2, 6])
        with col_prev:
            if st.button("👁 Anteprima URL", key="preview_sitemap"):
                with st.spinner("Lettura sitemap…"):
                    try:
                        preview_urls = sa.get_urls_from_sitemap(
                            sitemap_url.strip(), sa.get_session(), MAX_URLS
                        )
                        if preview_urls:
                            st.info(f"✅ {len(preview_urls)} URL trovati nella sitemap — verranno analizzate le prime {sitemap_max}")
                            with st.expander(f"Mostra tutte le URL ({len(preview_urls)})"):
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
                        sitemap_url.strip(), sa.get_session(), int(sitemap_max)
                    )
                    if from_sitemap:
                        urls_to_audit = from_sitemap
                        input_label   = sitemap_url.strip()
                        st.info(f"✅ {len(from_sitemap)} URL caricate dalla sitemap")
                    else:
                        st.error("Nessuna URL trovata nella sitemap.")
                except Exception as exc:
                    st.error(f"Errore lettura sitemap: {exc}")
        elif sitemap_url.strip():
            st.warning("L'URL della sitemap deve iniziare con http:// o https://")

# ── Tab 3: File CSV / Excel ───────────────────────────────────────────────────
with tab_file:
    col_up, col_num_f, col_bup = st.columns([4, 2, 1])
    file_max = MAX_URLS  # default, aggiornato dopo l'upload
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
                all_file_urls = (
                    df_up[col_name].dropna().astype(str).str.strip()
                    .pipe(lambda s: s[s.str.startswith("http")])
                    .tolist()
                )
                total_in_file = len(all_file_urls)
                with col_num_f:
                    file_max = st.number_input(
                        "Max URL",
                        min_value=1,
                        max_value=min(total_in_file, MAX_URLS),
                        value=min(10, total_in_file),
                        step=1,
                        help=f"{total_in_file} URL nel file, max {MAX_URLS} analizzabili",
                    )
                from_file = all_file_urls[:int(file_max)]
                st.info(f"✅ {total_in_file} URL nel file — verranno analizzate le prime {len(from_file)}")
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
        ("4.",  "Dati Strutturati",
         lambda u, h, rh, ss: sa.audit_structured_data(u, h)),
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
r4f  = r_by(first_pg, "4.")
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
    "4 Dati Strutt.", "8 E-E-A-T", "9 Performance", "10 Topical",
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
            "Strutt.":    ovs.get("4. ", "N/D"),
            "E-E-A-T":    ovs.get("8. ", "N/D"),
            "Perf.":      ovs.get("9. ", "N/D"),
            "Topical":    ovs.get("10.", "N/D"),
            "Score":      page_score(pg),
        })
    df_pg = pd.DataFrame(rows)
    st.dataframe(
        df_pg.style
             .map(color_cell, subset=BADGE_COLS)
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
                opp_str = opp if isinstance(opp, str) else str(opp)
                probs.append({"Area": area, "URL": u, "Tipo": "Opportunità",
                              "Severity": "WARN", "Dettaglio": opp_str[:120]})

    probs.sort(key=lambda x: {"FAIL": 0, "WARN": 1, "INFO": 2}.get(x["Severity"], 9))
    st.subheader(f"Problemi rilevati — {len(probs)} finding")

    if probs:
        sev_f = st.multiselect("Filtra severity", ["FAIL", "WARN", "INFO"],
                               default=["FAIL", "WARN"])
        df_pr = pd.DataFrame(probs)
        df_pr = df_pr[df_pr["Severity"].isin(sev_f)]
        st.dataframe(
            df_pr.style.map(color_cell, subset=["Severity"]),
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
            df_f.style.map(color_cell, subset=["Severity"]),
            use_container_width=True, hide_index=True,
        )
    else:
        st.success("✅ Nessun problema rilevato.")

    if r.get("robots_content"):
        with st.expander("Contenuto robots.txt"):
            st.code(r["robots_content"][:3000], language="text")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — AREA 4: DATI STRUTTURATI
# ══════════════════════════════════════════════════════════════════════════════

with tabs[6]:
    if npages > 1:
        st.subheader(f"Riepilogo Dati Strutturati — {npages} pagine")
        rows_sd_all = []
        for pg in grouped:
            pg_url = pg[0].get("url", "—") if pg else "—"
            rr = r_by(pg, "4.")
            rows_sd_all.append({
                "URL":           pg_url,
                "Overall":       rr.get("overall", "ERROR"),
                "Schema trovati": len(rr.get("schemas", [])),
                "JSON-LD":       rr.get("jsonld_blocks", 0),
                "Score medio":   rr.get("avg_score", 0),
                "Rich result":   rr.get("rich_ready", 0),
                "OG completo":   "Sì" if rr.get("og_complete") else "No",
                "Twitter Card":  "Sì" if rr.get("twitter_card") else "No",
            })
        df_sd_all = pd.DataFrame(rows_sd_all)
        st.dataframe(
            df_sd_all.style
                .map(color_cell, subset=["Overall"])
                .background_gradient(subset=["Score medio"], cmap="RdYlGn", vmin=0, vmax=100),
            use_container_width=True,
            height=min(600, 80 + 38 * len(df_sd_all)),
        )
        st.divider()

        page_urls_sd = [pg[0].get("url", f"Pagina {i+1}") for i, pg in enumerate(grouped)]
        sel_idx_sd = st.selectbox(
            "Dettaglio pagina",
            range(len(page_urls_sd)),
            format_func=lambda i: page_urls_sd[i],
            key="sd_page_sel",
        )
        r = r_by(grouped[sel_idx_sd], "4.")
    else:
        r = r4f
    ov = r.get("overall", "ERROR")
    ct, cb = st.columns([7, 1])
    ct.subheader("Dati Strutturati — Schema.org & Open Graph")
    cb.markdown(f"<div style='padding-top:10px'>{badge(ov)}</div>", unsafe_allow_html=True)
    st.caption(r.get("summary", "").replace(" | ", " · "))
    st.divider()

    schemas = r.get("schemas", [])
    warnings_sd = r.get("warnings", [])
    opps_sd = r.get("opportunities", [])

    # KPI
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Schema trovati",    len(schemas))
    c2.metric("JSON-LD",           r.get("jsonld_blocks", 0))
    c3.metric("Score medio",       f"{r.get('avg_score', 0)}/100")
    c4.metric("Rich result ready", r.get("rich_ready", 0))
    c5.metric("Tag Open Graph",    r.get("og_tags_count", 0))
    c6.metric("Twitter Card",      "Sì" if r.get("twitter_card") else "No")

    # Open Graph status
    if r.get("og_complete"):
        st.success(f"✅ Open Graph completo — tutti i tag essenziali presenti")
    else:
        og_miss = r.get("og_missing", [])
        st.warning(f"⚠️ Open Graph incompleto — mancano: {', '.join('og:'+f for f in og_miss)}")

    st.divider()

    # Tabella schema trovati
    if schemas:
        st.markdown("**Schema rilevati:**")
        prio_color = {"high": "🟢", "medium": "🟡", "low": "⚪"}
        rows_s = []
        for s in schemas:
            miss_r = ", ".join(s.get("missing_required", [])) or "—"
            miss_rec = f"{len(s.get('missing_recommended', []))} campi"
            rows_s.append({
                "Tipo":            s["type"],
                "Formato":         s["format"],
                "Priorità":        prio_color.get(s["priority"],"⚪") + " " + s["priority"].upper(),
                "Score":           f"{s['schema_score']}/100",
                "Rich result":     "✅" if s["has_rich_result_potential"] else "⚠️",
                "Required mancanti": miss_r,
                "Recommended mancanti": miss_rec,
                "Nome/Titolo":     s.get("name","")[:50],
            })
        df_s = pd.DataFrame(rows_s)
        st.dataframe(df_s, use_container_width=True, hide_index=True)
    else:
        st.error("⚠️ Nessun dato strutturato JSON-LD o Microdata trovato nella pagina.")

    # Errori di validazione
    if warnings_sd:
        st.divider()
        st.markdown("**Errori e avvisi di validazione:**")
        sev_filter_sd = st.multiselect(
            "Filtra", ["FAIL", "WARN"], default=["FAIL", "WARN"], key="sd_sev"
        )
        rows_w = [{"Severity": w["severity"], "Schema": w["schema"],
                   "Tipo errore": w["type"], "Campo": w.get("field","—") or "—",
                   "Messaggio": w["message"]}
                  for w in warnings_sd if w["severity"] in sev_filter_sd]
        if rows_w:
            df_w = pd.DataFrame(rows_w)
            st.dataframe(
                df_w.style.map(color_cell, subset=["Severity"]),
                use_container_width=True, hide_index=True,
            )

    # Opportunità
    if opps_sd:
        st.divider()
        st.markdown("**💡 Opportunità schema non sfruttate:**")
        for op in opps_sd:
            prio = op.get("priority","medium")
            icon = "🔴" if prio == "high" else "🟡"
            with st.expander(f"{icon} **{op['type']}** — {op['reason'][:80]}"):
                st.markdown(f"**Beneficio:** {op['benefit']}")
                st.markdown(f"**Priorità:** {prio.upper()}")

    # Open Graph dettaglio
    og_fields = r.get("og_fields", {})
    if og_fields:
        st.divider()
        st.markdown("**Tag Open Graph presenti:**")
        df_og = pd.DataFrame(
            [{"Tag": f"og:{k}", "Valore": v[:100]} for k, v in og_fields.items()]
        )
        st.dataframe(df_og, use_container_width=True, hide_index=True)

    # Twitter Card
    tc = r.get("tc_fields", {})
    if tc:
        st.markdown(f"**Twitter Card:** `{r.get('tc_card_type','—')}` — "
                    f"{len(tc)} tag presenti")



# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — AREA 8: E-E-A-T
# ══════════════════════════════════════════════════════════════════════════════

with tabs[7]:
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

with tabs[8]:
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
            } for f in pf]).style.map(color_cell, subset=["Severity"]),
            use_container_width=True, hide_index=True,
        )

    if r.get("positives"):
        st.markdown("**✅ Elementi ottimizzati:**")
        for p in r["positives"]:
            st.markdown(f'<div class="sig-ok">✓ {p}</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 8 — AREA 10: TOPICAL AUTHORITY
# ══════════════════════════════════════════════════════════════════════════════

with tabs[9]:
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

with tabs[10]:
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
    if st.button("Genera JSON", use_container_width=False):
        try:
            st.session_state["_export_data_JSON"] = json.dumps(
                [[{k: v for k, v in r.items() if k != "robots_content"} for r in pg]
                 for pg in grouped],
                ensure_ascii=False, indent=2, default=str,
            ).encode("utf-8")
            st.session_state["_export_fname_JSON"] = f"seo_audit_{datetime.now().strftime('%Y%m%d_%H%M')}.json"
        except Exception as exc:
            st.error(f"Errore JSON: {exc}")

    if "_export_data_JSON" in st.session_state:
        st.download_button(
            "⬇️ Scarica JSON",
            data=st.session_state["_export_data_JSON"],
            file_name=st.session_state["_export_fname_JSON"],
            mime="application/json",
        )
