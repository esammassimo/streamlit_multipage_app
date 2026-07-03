"""
4CAST — WSX SEO Analyzer · Navla SEO Tools
==========================================
CASTfastframework: Content · Authority · Structure · Technical
Pagina integrata nell'app Navla · richiede core_4cast.py nella root
"""
import streamlit as st
import pandas as pd
import re
import io
import os
import json as _json
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from core_4cast import (
    normalize_url, canonical_url,
    parse_suggestion, classify_action, classify_from_metric_name,
    detect_lang, needs_translation,
    generate_seo_tag, translate_to_italian, translate_to_lang, page_language,
    _prompt_authority_action, _prompt_jsonld_adapt, _prompt_tech_action, _call_llm,
    clen, len_status, delta_chars,
    TITLE_MIN, TITLE_MAX, DESC_MIN, DESC_MAX,
    ANT_MODELS, OAI_MODELS,
)

# ─── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title='4CAST — WSX SEO Analyzer',
    page_icon='📊',
    layout='wide',
    initial_sidebar_state='expanded'
)

# ─── STYLES — light theme (Streamlit default) ─────────────────────────────────
st.markdown("""
<style>
/* Custom classes only — no background or color overrides */

/* Section labels */
.section-label {
    font-size: .7rem;
    font-weight: 700;
    letter-spacing: .1em;
    text-transform: uppercase;
    color: #6b7280;
    border-bottom: 1px solid #e5e7eb;
    padding-bottom: .35rem;
    margin-bottom: .9rem;
}

/* Status badges */
.badge { display:inline-block; padding:.15rem .55rem; border-radius:4px; font-size:.7rem; font-weight:600; font-family:monospace; }
.badge-ok    { background:#d1fae5; color:#065f46; }
.badge-long  { background:#fee2e2; color:#991b1b; }
.badge-short { background:#fef3c7; color:#92400e; }
.badge-en    { background:#dbeafe; color:#1e40af; }
.badge-it    { background:#ede9fe; color:#5b21b6; }
.badge-new   { background:#d1fae5; color:#065f46; }
.badge-miss  { background:#f3f4f6; color:#6b7280; }

/* Metric cards */
.metric-card {
    border: 1px solid #e5e7eb;
    border-radius: 10px;
    padding: .9rem 1rem;
    text-align: center;
}
.metric-card .val { font-size: 1.8rem; font-weight: 700; color: #1d4ed8; line-height: 1.1; }
.metric-card .lbl { font-size: .7rem; color: #6b7280; text-transform: uppercase; letter-spacing:.06em; }
</style>
""", unsafe_allow_html=True)

# ─── HEADER ───────────────────────────────────────────────────────────────────
st.title("📊 4CAST — WSX SEO Analyzer")
st.caption("Navla SEO Tools · Framework CAST: Content · Authority · Structure · Technical")
st.divider()



# ─── SIDEBAR — API KEYS ───────────────────────────────────────────────────────
# Le chiavi impostate qui sovrascrivono le variabili di ambiente per questa pagina.
# Restano vuote se già configurate nel file .env o nelle secrets di Streamlit Cloud.
with st.sidebar:
    st.image(
        "https://img.shields.io/badge/4CAST-WSX%20Analyzer-2563eb?style=for-the-badge",
        use_container_width=True,
    )
    st.markdown("---")
    st.header("🔐 API Keys")
    st.caption("Inserisci le chiavi LLM per attivare traduzione e ottimizzazione AI. "
               "Se già presenti nel file .env non è necessario reinserirle.")

    _sb_ant = st.text_input(
        "Anthropic API Key",
        type="password",
        placeholder="sk-ant-... (opzionale se in .env)",
        key="cast_ant_key_sb",
        help="Usata per Claude (Haiku / Sonnet). Richiesta in modalità Anthropic.",
    )
    _sb_oai = st.text_input(
        "OpenAI API Key",
        type="password",
        placeholder="sk-... (opzionale se in .env)",
        key="cast_oai_key_sb",
        value=st.session_state.get('openai_api_key', ''),
        help="Usata per GPT-4o / GPT-4o-mini. Puoi usare la stessa chiave della home.",
    )

    # Propaga le chiavi nelle variabili d'ambiente per questo processo
    if _sb_ant:
        os.environ['ANTHROPIC_API_KEY'] = _sb_ant
    if _sb_oai:
        os.environ['OPENAI_API_KEY'] = _sb_oai

    # Stato connessione
    _ant_ok = bool(os.getenv('ANTHROPIC_API_KEY', _sb_ant))
    _oai_ok = bool(os.getenv('OPENAI_API_KEY', _sb_oai))
    st.markdown(
        f"{'🟢' if _ant_ok else '🔴'} Anthropic  &nbsp;&nbsp;"
        f"{'🟢' if _oai_ok else '🔴'} OpenAI",
        unsafe_allow_html=True,
    )

    st.markdown("---")
    st.caption(
        "📖 **Documentazione**: ogni tab contiene un expander *ℹ️ File e API richiesti* "
        "con le istruzioni dettagliate per gli export necessari."
    )
    st.markdown("---")
    st.caption("Navla SEO Tools · 4CAST v1.0")

# ─── SESSION STATE ────────────────────────────────────────────────────────────
for key in [
    # C — Content
    'wsx_df', 'sf_df', 'result_df', 'comp_df', 'wsx_prev_df', 'diff_df',
    # A — Authority (A01 score + A02 internal link audit)
    'auth_df', 'auth_result_df',
    'il_wsx_df', 'il_sf_links_df', 'il_sf_pages_df', 'il_ahrefs_df', 'il_result_df',
    # S — Structure
    'struct_df', 'struct_result_df',
    # T — Technical
    'tech_df', 'tech_result_df',
    # LLM config (persisted across tabs)
    'llm_provider', 'llm_key', 'llm_model', 'llm_mode',
]:
    if key not in st.session_state:
        st.session_state[key] = None

# ─── HELPERS ──────────────────────────────────────────────────────────────────
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        df.to_excel(w, index=False)
    return buf.getvalue()

def badge(text, kind):
    cls = {'ok':'badge-ok','long':'badge-long','short':'badge-short',
           'en':'badge-en','it':'badge-it','new':'badge-new','miss':'badge-miss'}.get(kind,'badge-miss')
    return f'<span class="badge {cls}">{text}</span>'

def len_badge(n, lo, hi):
    if n is None: return badge('—', 'miss')
    if lo <= n <= hi: return badge(f'{n} ✓', 'ok')
    if n < lo: return badge(f'{n} short', 'short')
    return badge(f'{n} long', 'long')

def _wsx_score(row, col):
    """Estrae uno score numerico WSX; None se assente o non numerico."""
    v = row.get(col)
    try:
        return int(float(v)) if v is not None and pd.notna(v) else None
    except (ValueError, TypeError):
        return None

# ══════════════════════════════════════════════════════════════════════════════
#  ROW WORKERS  —  funzioni pure (no Streamlit), eseguibili in thread pool
# ══════════════════════════════════════════════════════════════════════════════

def _analyse_row(args: dict) -> tuple:
    """
    Elabora UNA riga WSX: parsing + LLM. Gira in un thread del pool.
    Restituisce (idx_originale, dict_risultato).
    """
    row       = args['row']
    translate = args['translate']
    llm_mode  = args['llm_mode']   # 'translate' | 'optimize' | 'complete'
    provider  = args['provider']
    api_key   = args['api_key']
    model     = args['model']

    sug_text  = str(row.get(args['col_sug'],    '')) if pd.notna(row.get(args['col_sug'],    '')) else ''
    metric    = str(row.get(args['col_metric'], '')) if pd.notna(row.get(args['col_metric'], '')) else ''
    goal      = str(row.get(args['col_goal'],   '')) if pd.notna(row.get(args['col_goal'],   '')) else ''
    must_have = str(row.get(args['col_must'],   '')) if pd.notna(row.get(args['col_must'],   '')) else ''

    p       = parse_suggestion(sug_text)
    actions = classify_from_metric_name(metric, sug_text)

    url = str(row.get('URL', ''))
    if url and not url.startswith('http'):
        url = 'https://www.' + url

    st_val  = p.get('suggested_title',       '')
    sd_val  = p.get('suggested_description', '')
    sh1_val = p.get('suggested_h1',          '')
    if metric == 'Heading Structure' and st_val and not sh1_val:
        sh1_val, st_val = st_val, ''

    if translate and sug_text:
        tgt    = page_language(row.get('Language', ''))
        lkw    = dict(provider=provider, api_key=api_key, model=model)
        ckw    = dict(suggestion_text=sug_text, target_lang=tgt,
                      url=url, metric_name=metric, must_have=must_have)
        # Modalità interna usata per generate_seo_tag
        gen_m  = 'translate' if llm_mode == 'translate' else 'optimize'

        # ── Decide se generare ciascun tag ──────────────────────────────────
        def _should(extracted, tag_type):
            if llm_mode == 'complete':
                return bool(sug_text)                    # sempre
            if llm_mode == 'optimize':
                return bool(extracted or sug_text)       # se c'è contesto
            # translate
            return bool(extracted and needs_translation(extracted, tgt))

        if _should(st_val, 'title'):
            v = generate_seo_tag(tag_type='title', extracted_value=st_val,
                                 mode=gen_m, **ckw, **lkw)
            if v: st_val = v

        if _should(sd_val, 'description'):
            v = generate_seo_tag(tag_type='description', extracted_value=sd_val,
                                 mode=gen_m, **ckw, **lkw)
            if v: sd_val = v

        # H1 solo per Heading Structure o se già estratto
        h1_needed = sh1_val or (llm_mode in ('optimize','complete') and metric == 'Heading Structure')
        if h1_needed:
            v = generate_seo_tag(tag_type='h1', extracted_value=sh1_val,
                                 mode=gen_m, **ckw, **lkw)
            if v: sh1_val = v

    st_len = clen(st_val)
    sd_len = clen(sd_val)
    if   st_val and sd_val:  cov = 'Title + Description'
    elif st_val:             cov = 'Title only'
    elif sd_val:             cov = 'Description only'
    elif sh1_val:            cov = 'H1 only'
    else:                    cov = 'No suggestion'

    return args['idx'], {
        'URL':                      url,
        'Brand':                    row.get('Brand',     ''),
        'Market':                   row.get('Market',    ''),
        'Language':                 row.get('Language',  ''),
        'Page Type':                row.get('Page Type', ''),
        'Score Context':            _wsx_score(row, 'Context'),
        'Δ Context MoM':            row.get('Context MoM', ''),
        'Score Meta Tags':          _wsx_score(row, 'Meta Tags'),
        'Score Heading':            _wsx_score(row, 'Heading'),
        'Score Relevance':          _wsx_score(row, 'Relevance'),
        'Score Unique Content':     _wsx_score(row, 'Unique Content'),
        'Metric Name':              metric,
        'Action Type':              ' + '.join(actions),
        'Goal':                     goal,
        'Must Have':                must_have,
        'Primary Keyword':          p.get('primary_keyword', ''),
        'suggested_title':          st_val,
        'sug_title_len':            st_len,
        'sug_title_status':         len_status(st_len, TITLE_MIN, TITLE_MAX),
        'suggested_description':    sd_val,
        'sug_desc_len':             sd_len,
        'sug_desc_status':          len_status(sd_len, DESC_MIN, DESC_MAX),
        'suggested_h1':             sh1_val,
        'suggestion_coverage':      cov,
        'original_suggestion_text': sug_text,
    }


def _compare_row(args: dict) -> tuple:
    """
    Elabora UNA riga WSX nel confronto con SF. Gira in thread pool.
    Restituisce (idx_originale, dict_risultato).
    """
    row       = args['row']
    translate = args['translate']
    llm_mode  = args['llm_mode']
    provider  = args['provider']
    api_key   = args['api_key']
    model     = args['model']
    sf_lookup = args['sf_lookup']

    norm = normalize_url(row.get('URL', ''))
    sug_text  = str(row.get(args['col_sug'],    '')) if pd.notna(row.get(args['col_sug'],    '')) else ''
    metric    = str(row.get(args['col_metric'], '')) if pd.notna(row.get(args['col_metric'], '')) else ''
    goal      = str(row.get(args['col_goal'],   '')) if pd.notna(row.get(args['col_goal'],   '')) else ''
    must_have = str(row.get(args['col_must'],   '')) if pd.notna(row.get(args['col_must'],   '')) else ''

    p  = parse_suggestion(sug_text)
    sf = sf_lookup.get(norm, {})

    sf_title     = sf.get('sf_title', '')
    sf_title_len = sf.get('sf_title_len')
    sf_desc      = sf.get('sf_desc', '')
    sf_desc_len  = sf.get('sf_desc_len')
    sf_h1        = sf.get('sf_h1', '')
    sf_h1_len    = sf.get('sf_h1_len')

    st_val  = p.get('suggested_title', '')
    sd_val  = p.get('suggested_description', '')
    sh1_val = p.get('suggested_h1', '')
    if metric == 'Heading Structure' and st_val and not sh1_val:
        sh1_val, st_val = st_val, ''

    if translate and sug_text:
        tgt   = page_language(row.get('Language', ''))
        lkw   = dict(provider=provider, api_key=api_key, model=model)
        ckw   = dict(suggestion_text=sug_text, target_lang=tgt,
                     url=canonical_url(norm), metric_name=metric, must_have=must_have)
        gen_m = 'translate' if llm_mode == 'translate' else 'optimize'

        def _should(extracted):
            if llm_mode == 'complete':  return bool(sug_text)
            if llm_mode == 'optimize':  return bool(extracted or sug_text)
            return bool(extracted and needs_translation(extracted, tgt))

        if _should(st_val):
            v = generate_seo_tag(tag_type='title', extracted_value=st_val,
                                 mode=gen_m, **ckw, **lkw)
            if v: st_val = v
        if _should(sd_val):
            v = generate_seo_tag(tag_type='description', extracted_value=sd_val,
                                 mode=gen_m, **ckw, **lkw)
            if v: sd_val = v
        if sh1_val or (llm_mode in ('optimize','complete') and metric == 'Heading Structure'):
            v = generate_seo_tag(tag_type='h1', extracted_value=sh1_val,
                                 mode=gen_m, **ckw, **lkw)
            if v: sh1_val = v

    st_len  = clen(st_val)
    sd_len  = clen(sd_val)
    if   st_val and sd_val: cov = 'Title + Description'
    elif st_val:             cov = 'Title only'
    elif sd_val:             cov = 'Description only'
    elif sh1_val:            cov = 'H1 only'
    else:                    cov = 'No suggestion'

    def _dif(a, b): return bool(a and b and a.strip().lower() != b.strip().lower())
    def _lbl(d, cur, sug):
        if not sug: return '— nessun suggerimento'
        if not cur: return '⚠ tag assente'
        return '⚠ Diverso' if d else '✓ Allineato'

    td = _dif(sf_title, st_val);  dd = _dif(sf_desc, sd_val);  hd = _dif(sf_h1, sh1_val)

    return args['idx'], {
        'URL':                    canonical_url(norm),
        'Brand':                  row.get('Brand',     ''),
        'Market':                 row.get('Market',    ''),
        'Language':               row.get('Language',  ''),
        'Page Type':              row.get('Page Type', ''),
        'Score Context':          _wsx_score(row, 'Context'),
        'Score Meta Tags':        _wsx_score(row, 'Meta Tags'),
        'Score Heading':          _wsx_score(row, 'Heading'),
        'Score Relevance':        _wsx_score(row, 'Relevance'),
        'Metric Name':            metric,
        'Goal':                   goal,
        'Must Have':              must_have,
        'SF Match':               'YES' if sf else 'NOT IN CRAWL',
        'Status Code':            sf.get('sf_status', ''),
        'Indexability':           sf.get('sf_indexability', ''),
        'current_title':          sf_title,
        'current_title_len':      sf_title_len,
        'current_title_status':   len_status(sf_title_len, TITLE_MIN, TITLE_MAX),
        'suggested_title':        st_val,
        'suggested_title_len':    st_len,
        'suggested_title_status': len_status(st_len, TITLE_MIN, TITLE_MAX),
        'title_delta':            delta_chars(sf_title, st_val),
        'title_diff':             _lbl(td, sf_title, st_val),
        'current_description':    sf_desc,
        'current_desc_len':       sf_desc_len,
        'current_desc_status':    len_status(sf_desc_len, DESC_MIN, DESC_MAX),
        'suggested_description':  sd_val,
        'suggested_desc_len':     sd_len,
        'suggested_desc_status':  len_status(sd_len, DESC_MIN, DESC_MAX),
        'desc_delta':             delta_chars(sf_desc, sd_val),
        'desc_diff':              _lbl(dd, sf_desc, sd_val),
        'current_h1':             sf_h1,
        'current_h1_len':         sf_h1_len,
        'suggested_h1':           sh1_val,
        'h1_diff':                _lbl(hd, sf_h1, sh1_val),
        'tags_differing':         sum([td, dd, hd]),
        'suggestion_coverage':    cov,
        'original_suggestion':    sug_text,
    }


# ── Helper parallelo generico ──────────────────────────────────────────────────

def _run_parallel(worker_fn, jobs: list, n_workers: int,
                  progress_label: str = 'Elaborazione') -> list:
    """
    Esegue worker_fn(job) in parallelo su n_workers thread.
    Restituisce la lista dei risultati nell'ordine originale dei job.
    """
    total   = len(jobs)
    results = [None] * total
    prog    = st.progress(0, text=f'{progress_label} 0/{total}…')
    done    = 0

    with ThreadPoolExecutor(max_workers=n_workers) as pool:
        future_map = {pool.submit(worker_fn, job): job['idx'] for job in jobs}
        for future in as_completed(future_map):
            done += 1
            prog.progress(done / total,
                          text=f'{progress_label} {done}/{total}…')
            try:
                idx, result = future.result()
                results[idx] = result
            except Exception as exc:
                idx = future_map[future]
                results[idx] = {'_error': str(exc)}

    prog.empty()
    return [r for r in results if r is not None]


# ══════════════════════════════════════════════════════════════════════════════
#  BUILD FUNCTIONS  (orchestrano i worker)
# ══════════════════════════════════════════════════════════════════════════════

def build_wsx_analysis(df: pd.DataFrame, translate: bool = False,
                       provider: str = 'anthropic', api_key: str = None,
                       mode: str = 'complete', model: str = None,
                       n_workers: int = 4) -> pd.DataFrame:
    """Tab 1 — parsing + LLM in parallelo."""
    col_sug    = _find_col(df, 'Top Recommendation - Customized Suggestion')
    col_metric = _find_col(df, 'Top Recommendation - Metric Name')
    col_goal   = _find_col(df, 'Top Recommendation - Goal')
    col_must   = _find_col(df, 'Top Recommendation - Must Have')

    jobs = [
        {'idx': i, 'row': row, 'translate': translate,
         'llm_mode': mode, 'provider': provider, 'api_key': api_key, 'model': model,
         'col_sug': col_sug, 'col_metric': col_metric,
         'col_goal': col_goal, 'col_must': col_must}
        for i, (_, row) in enumerate(df.iterrows())
    ]
    rows = _run_parallel(_analyse_row, jobs, n_workers,
                         progress_label='Analisi suggerimenti')
    return pd.DataFrame(rows)


def build_comparison(wsx_df: pd.DataFrame, sf_df: pd.DataFrame,
                     translate: bool = False,
                     provider: str = 'anthropic', api_key: str = None,
                     mode: str = 'complete', model: str = None,
                     n_workers: int = 4) -> pd.DataFrame:
    """Tab 2 — merge WSX + SF crawl in parallelo."""
    col_sug    = _find_col(wsx_df, 'Top Recommendation - Customized Suggestion')
    col_metric = _find_col(wsx_df, 'Top Recommendation - Metric Name')
    col_goal   = _find_col(wsx_df, 'Top Recommendation - Goal')
    col_must   = _find_col(wsx_df, 'Top Recommendation - Must Have')

    # Pre-build SF lookup (single-threaded, fast)
    sf_df['_norm'] = sf_df['Address'].apply(normalize_url)
    sf_lookup: dict = {}
    for _, r in sf_df.iterrows():
        k = r['_norm']
        if k not in sf_lookup:
            def _sf(c): return str(r.get(c,'')) if pd.notna(r.get(c)) else ''
            def _si(c): return int(r[c]) if pd.notna(r.get(c)) else None
            sf_lookup[k] = {
                'sf_title':        _sf('Title 1'),
                'sf_title_len':    _si('Title 1 Length'),
                'sf_desc':         _sf('Meta Description 1'),
                'sf_desc_len':     _si('Meta Description 1 Length'),
                'sf_h1':           _sf('H1-1'),
                'sf_h1_len':       _si('H1-1 Length'),
                'sf_status':       _sf('Status Code'),
                'sf_indexability': _sf('Indexability'),
            }

    jobs = [
        {'idx': i, 'row': row, 'translate': translate, 'sf_lookup': sf_lookup,
         'llm_mode': mode, 'provider': provider, 'api_key': api_key, 'model': model,
         'col_sug': col_sug, 'col_metric': col_metric,
         'col_goal': col_goal, 'col_must': col_must}
        for i, (_, row) in enumerate(wsx_df.iterrows())
    ]
    rows = _run_parallel(_compare_row, jobs, n_workers,
                         progress_label='Confronto con SF')
    return pd.DataFrame(rows)

def build_diff(df_new: pd.DataFrame, df_old: pd.DataFrame) -> pd.DataFrame:
    """Tab 3 — confronto NEW vs OLD: URL, tag suggeriti, score WSX."""
    col_sug_new = _find_col(df_new, 'Top Recommendation - Customized Suggestion')
    col_sug_old = _find_col(df_old, 'Top Recommendation - Customized Suggestion')
    col_met_new = _find_col(df_new, 'Top Recommendation - Metric Name')
    col_met_old = _find_col(df_old, 'Top Recommendation - Metric Name')

    def _sdelta(new_v, old_v):
        if new_v is None or old_v is None: return ''
        d = new_v - old_v
        return f'+{d}' if d > 0 else (str(d) if d < 0 else '=')

    df_new['_norm'] = df_new['URL'].apply(normalize_url)
    df_old['_norm'] = df_old['URL'].apply(normalize_url)
    old_map = {r['_norm']: r for _, r in df_old.iterrows()}

    rows = []
    for _, row in df_new.iterrows():
        norm    = row['_norm']
        new_sug = str(row.get(col_sug_new, '')) if pd.notna(row.get(col_sug_new)) else ''
        new_met = str(row.get(col_met_new, '')) if pd.notna(row.get(col_met_new)) else ''

        new_sc  = _wsx_score(row, 'Context')
        new_sm  = _wsx_score(row, 'Meta Tags')
        new_sh  = _wsx_score(row, 'Heading')
        new_sr  = _wsx_score(row, 'Relevance')

        if norm in old_map:
            old_row = old_map[norm]
            old_sug = str(old_row.get(col_sug_old, '')) if pd.notna(old_row.get(col_sug_old)) else ''
            old_met = str(old_row.get(col_met_old, '')) if pd.notna(old_row.get(col_met_old)) else ''

            old_sc = _wsx_score(old_row, 'Context')
            old_sm = _wsx_score(old_row, 'Meta Tags')
            old_sh = _wsx_score(old_row, 'Heading')
            old_sr = _wsx_score(old_row, 'Relevance')

            p_new = parse_suggestion(new_sug)
            p_old = parse_suggestion(old_sug)

            new_st  = p_new.get('suggested_title', '')
            old_st  = p_old.get('suggested_title', '')
            new_sd  = p_new.get('suggested_description', '')
            old_sd  = p_old.get('suggested_description', '')
            new_sh1 = p_new.get('suggested_h1', '')
            old_sh1 = p_old.get('suggested_h1', '')

            if new_met == 'Heading Structure' and new_st and not new_sh1:
                new_sh1, new_st = new_st, ''
            if old_met == 'Heading Structure' and old_st and not old_sh1:
                old_sh1, old_st = old_st, ''

            title_ch = (new_st.lower() != old_st.lower())   if new_st and old_st else bool(new_st or old_st)
            desc_ch  = (new_sd.lower() != old_sd.lower())   if new_sd and old_sd else bool(new_sd or old_sd)
            h1_ch    = (new_sh1.lower() != old_sh1.lower()) if new_sh1 and old_sh1 else bool(new_sh1 or old_sh1)

            if   not new_st and not new_sd and not new_sh1: status = 'no_suggestion_new'
            elif not old_st and not old_sd and not old_sh1: status = 'new_url_has_suggestion'
            elif title_ch or desc_ch or h1_ch:              status = 'suggestion_updated'
            else:                                           status = 'suggestion_unchanged'

            # Euristiche "tag aggiornati": punteggio migliorato rispetto a OLD
            meta_up = bool(new_sm and old_sm and new_sm > old_sm)
            head_up = bool(new_sh and old_sh and new_sh > old_sh)
            if meta_up or head_up:
                tags_status = 'Aggiornati ✓'
            elif status == 'suggestion_unchanged':
                tags_status = 'Non aggiornati'
            else:
                tags_status = '—'
        else:
            old_sug = old_met = ''
            old_sc = old_sm = old_sh = old_sr = None
            p_new   = parse_suggestion(new_sug)
            new_st  = p_new.get('suggested_title', '')
            new_sd  = p_new.get('suggested_description', '')
            new_sh1 = p_new.get('suggested_h1', '')
            if new_met == 'Heading Structure' and new_st and not new_sh1:
                new_sh1, new_st = new_st, ''
            old_st = old_sd = old_sh1 = ''
            title_ch = desc_ch = h1_ch = False
            status = 'new_url'
            tags_status = '—'

        rows.append({
            'URL':                   canonical_url(norm),
            'Brand':                 row.get('Brand',    ''),
            'Market':                row.get('Market',   ''),
            'Language':              row.get('Language', ''),
            'Metric Name':           new_met,
            'status':                status,
            'Tags Status':           tags_status,
            # ── Score comparison ──
            'Context (new)':         new_sc,    'Context (old)':     old_sc,    'Δ Context':     _sdelta(new_sc, old_sc),
            'Meta Tags (new)':       new_sm,    'Meta Tags (old)':   old_sm,    'Δ Meta Tags':   _sdelta(new_sm, old_sm),
            'Heading (new)':         new_sh,    'Heading (old)':     old_sh,    'Δ Heading':     _sdelta(new_sh, old_sh),
            'Relevance (new)':       new_sr,    'Relevance (old)':   old_sr,    'Δ Relevance':   _sdelta(new_sr, old_sr),
            # ── Tag suggestions ──
            'new_suggested_title':   new_st,    'old_suggested_title': old_st,  'title_changed': '●' if title_ch else '—',
            'new_suggested_desc':    new_sd,    'old_suggested_desc':  old_sd,  'desc_changed':  '●' if desc_ch  else '—',
            'new_suggested_h1':      new_sh1,   'old_suggested_h1':    old_sh1, 'h1_changed':    '●' if h1_ch    else '—',
            'new_suggestion_text':   new_sug,
            'old_suggestion_text':   old_sug,
        })

    # URL presenti in OLD ma non in NEW
    new_norms = set(df_new['_norm'])
    for _, row in df_old.iterrows():
        if row['_norm'] not in new_norms:
            old_sug = str(row.get(col_sug_old, '')) if pd.notna(row.get(col_sug_old)) else ''
            old_met = str(row.get(col_met_old, '')) if pd.notna(row.get(col_met_old)) else ''
            p_old   = parse_suggestion(old_sug)
            old_st  = p_old.get('suggested_title', '')
            old_sd  = p_old.get('suggested_description', '')
            old_sh1 = p_old.get('suggested_h1', '')
            if old_met == 'Heading Structure' and old_st and not old_sh1:
                old_sh1, old_st = old_st, ''
            rows.append({
                'URL':                   canonical_url(row['_norm']),
                'Brand':                 row.get('Brand',    ''),
                'Market':                row.get('Market',   ''),
                'Language':              row.get('Language', ''),
                'Metric Name':           old_met,
                'status':                'removed_from_new',
                'Tags Status':           '—',
                'Context (new)': None,   'Context (old)':   _wsx_score(row,'Context'),   'Δ Context':   '',
                'Meta Tags (new)': None, 'Meta Tags (old)': _wsx_score(row,'Meta Tags'), 'Δ Meta Tags': '',
                'Heading (new)': None,   'Heading (old)':   _wsx_score(row,'Heading'),   'Δ Heading':   '',
                'Relevance (new)': None, 'Relevance (old)': _wsx_score(row,'Relevance'), 'Δ Relevance': '',
                'new_suggested_title':   '', 'old_suggested_title': old_st,  'title_changed': '—',
                'new_suggested_desc':    '', 'old_suggested_desc':  old_sd,  'desc_changed':  '—',
                'new_suggested_h1':      '', 'old_suggested_h1':    old_sh1, 'h1_changed':    '—',
                'new_suggestion_text':   '',
                'old_suggestion_text':   old_sug,
            })

    return pd.DataFrame(rows)

def _find_col(df: pd.DataFrame, name: str) -> str:
    """
    Trova il nome reale di una colonna in modo case-insensitive.
    Prima prova il match esatto (lowercase), poi un match normalizzato
    che ignora spazi, trattini e underscore — così
    'top_recommendation_customized_suggestion' trova
    'Top Recommendation - Customized Suggestion'.
    """
    col_map = {c.lower(): c for c in df.columns}
    result = col_map.get(name.lower())
    if result:
        return result
    def _norm(s): return re.sub(r'[^a-z0-9]', '', s.lower())
    norm_map = {_norm(c): c for c in df.columns}
    return norm_map.get(_norm(name), name)

def metrics_row(df: pd.DataFrame, cols: dict):
    """Render a row of metric cards."""
    cols_st = st.columns(len(cols))
    for col_st, (label, val) in zip(cols_st, cols.items()):
        with col_st:
            st.markdown(f"""
            <div class="metric-card">
                <div class="val">{val}</div>
                <div class="lbl">{label}</div>
            </div>""", unsafe_allow_html=True)

def status_icon(status: str) -> str:
    icons = {
        'suggestion_unchanged': '🟡 Invariato',
        'suggestion_updated':   '🟢 Aggiornato',
        'new_url':              '🔵 Nuova URL',
        'removed_from_new':     '🔴 Rimossa',
        'no_suggestion_new':    '⚪ Nessun suggerimento',
        'new_url_has_suggestion':'🟢 Nuovo + Suggerimento',
    }
    return icons.get(status, status)

# ══════════════════════════════════════════════════════════════════════════════
#  LLM CONFIG  (globale, applicato a tutte le tab)
# ══════════════════════════════════════════════════════════════════════════════
with st.expander('⚙️ Configurazione LLM', expanded=True):
    _lc1, _lc2, _lc3, _lc4 = st.columns([1, 1.2, 1.5, 1])

    with _lc1:
        _provider = st.radio('Provider', ['Anthropic', 'OpenAI'],
                             key='llm_prov', horizontal=True)

    with _lc2:
        _model_opts = ANT_MODELS if _provider == 'Anthropic' else OAI_MODELS
        _llm_model  = st.selectbox('Modello', _model_opts, key='llm_modelsel')

    with _lc3:
        _llm_mode = st.radio(
            'Modalità', ['Traduzione', 'Ottimizzazione', 'Completa'],
            key='llm_modesel', horizontal=True,
            help=(
                'Traduzione: traduce i valori estratti (veloce, min token).\n'
                'Ottimizzazione: genera tag ottimizzati con contesto completo.\n'
                'Completa: genera Title + Description per ogni URL con suggerimento '
                '(max qualità, max token).'
            ),
        )

    with _lc4:
        _n_workers = st.slider('Workers', min_value=1, max_value=10, value=4, step=1,
                               key='llm_workers',
                               help='Richieste API parallele. Aumentare per velocizzare; '
                                    'ridurre in caso di errori di rate-limit.')

    # ── Mappa UI → valori interni ──────────────────────────────────────────
    _PROVIDER  = 'openai' if _provider == 'OpenAI' else 'anthropic'
    _MODE      = {'Traduzione': 'translate', 'Ottimizzazione': 'optimize',
                  'Completa': 'complete'}[_llm_mode]
    # API key: letta dalla sidebar (propagata in os.environ nel blocco sidebar)
    _API_KEY   = (_sb_ant if _PROVIDER == 'anthropic' else _sb_oai) or None
    _MODEL     = _llm_model
    _N_WORKERS = _n_workers

st.markdown('<br>', unsafe_allow_html=True)

def _authority_row_worker(args: dict) -> tuple:
    """
    Elabora UNA riga Authority: score + classificazione priorità + LLM opzionale.
    Thread-safe, nessuna chiamata Streamlit.
    """
    row      = args['row']
    use_llm  = args['use_llm']
    provider = args['provider']
    api_key  = args['api_key']
    model    = args['model']

    url = str(row.get('URL', ''))
    if url and not url.startswith('http'):
        url = 'https://www.' + url

    metric    = str(row.get('Top Recommendation - Metric Name', '')) \
                if pd.notna(row.get('Top Recommendation - Metric Name')) else ''
    goal      = str(row.get('Top Recommendation - Goal', '')) \
                if pd.notna(row.get('Top Recommendation - Goal')) else ''
    must_have = str(row.get('Top Recommendation - Must Have', '')) \
                if pd.notna(row.get('Top Recommendation - Must Have')) else ''

    authority = _wsx_score(row, 'Authority') or 0

    # Priority level based on Authority score
    if authority < 30:   priority = '🔴 Alta'
    elif authority < 60: priority = '🟡 Media'
    else:                priority = '🟢 Bassa'

    # Score for the specific recommendation metric
    _metric_col_map = {
        'Internal Linking':     'Internal Linking',
        'Backlinking Quality':  'Backlinking Quality',
        'Backlinking Quantity': 'Backlinking Quantity',
        'Reviews Count':        'Reviews Count',
        'Reviews Value':        'Reviews Value',
        'Content Freshness':    'Content Freshness',
    }
    metric_score = _wsx_score(row, _metric_col_map[metric]) if metric in _metric_col_map else None

    # LLM: Italian action recommendation (optional)
    llm_action = ''
    if use_llm and metric:
        try:
            prompt = _prompt_authority_action(
                url=url, authority_score=authority,
                metric_name=metric, metric_score=metric_score,
                must_have=must_have, goal=goal,
            )
            llm_action = _call_llm(prompt, provider=provider, api_key=api_key, model=model)
        except Exception:
            pass

    return args['idx'], {
        'URL':                  url,
        'Brand':                row.get('Brand',     ''),
        'Market':               row.get('Market',    ''),
        'Language':             row.get('Language',  ''),
        'Page Type':            row.get('Page Type', ''),
        'Priority':             priority,
        # Authority scores
        'Authority':            _wsx_score(row, 'Authority'),
        'Δ Authority MoM':      row.get('Authority MoM', ''),
        'Internal Linking':     _wsx_score(row, 'Internal Linking'),
        'Backlinking Quality':  _wsx_score(row, 'Backlinking Quality'),
        'Backlinking Quantity': _wsx_score(row, 'Backlinking Quantity'),
        'Content Freshness':    _wsx_score(row, 'Content Freshness'),
        'Reviews Count':        _wsx_score(row, 'Reviews Count'),
        'Reviews Value':        _wsx_score(row, 'Reviews Value'),
        # Top Recommendation
        'Metric Name':          metric,
        'Metric Score':         metric_score,
        'Goal':                 goal,
        'Must Have':            must_have,
        # Analytics context
        'SEO Impressions':      _wsx_score(row, 'SEO Impressions (Google Search Console)'),
        'All Sessions':         _wsx_score(row, 'All Sessions (Google Analytics)'),
        # LLM output
        'Azione consigliata':   llm_action,
    }


def build_authority_analysis(df: pd.DataFrame, use_llm: bool = False,
                              provider: str = 'anthropic', api_key: str = None,
                              model: str = None, n_workers: int = 4) -> pd.DataFrame:
    """A tab — Authority score analysis, prioritisation, optional LLM actions."""
    jobs = [
        {'idx': i, 'row': row, 'use_llm': use_llm,
         'provider': provider, 'api_key': api_key, 'model': model}
        for i, (_, row) in enumerate(df.iterrows())
    ]
    rows = _run_parallel(_authority_row_worker, jobs, n_workers,
                         progress_label='Analisi Authority')
    return pd.DataFrame(rows)


def _il_classify(inlinks, ur, sf_internal_outlinks, wsx_metric):
    """
    Classifica un URL per tipologia e livello di priorità.
    Restituisce (tipologia, priority_level, priority_score).
    """
    n   = inlinks or 0
    ur  = ur or 0
    out = sf_internal_outlinks if sf_internal_outlinks is not None else 999

    if n == 0:
        return 'Orfana', '🔴 Urgente', min(90 + int(ur * 0.1), 100)
    if n < 3:
        return 'Quasi-orfana', '🔴 Alta',    min(60 + int(ur * 0.3), 90)
    if ur > 30 and out < 5:
        return 'Hub silenzioso', '🟠 Media-alta', min(50 + int(ur * 0.4), 85)
    if n < 10 and wsx_metric == 'Internal Linking':
        return 'Sottolinkata (WSX)', '🟡 Media', min(40 + int(ur * 0.2), 70)
    return 'OK', '🟢 Bassa', max(5, int(ur * 0.05))


def _il_action_text(tipologia, n_inlinks, ur, top_anchors, page_type):
    """Genera raccomandazione rule-based per internal linking."""
    n   = n_inlinks or 0
    ur  = ur or 0
    anc = str(top_anchors or '').strip()
    pt  = str(page_type or '')

    if tipologia == 'Orfana':
        return (
            f"PRIORITÀ MASSIMA — 0 inlink interni rilevati da SF. "
            f"Inserire min. 3 link contestuali da PLP o editoriali pertinenti ({pt})."
        )
    if tipologia == 'Quasi-orfana':
        return (
            f"Solo {n} inlink interni. Target min. 5 da hub editoriali/PLP correlate. "
            f"Verificare anche anchor text (attuali: {anc or 'N/A'})."
        )
    if tipologia == 'Hub silenzioso':
        return (
            f"Pagina con UR {ur} non distribuisce equity (< 5 link interni in uscita). "
            f"Aggiungere link verso PDP/PLP correlate con anchor keyword-rich."
        )
    if tipologia == 'Sottolinkata (WSX)':
        return (
            f"WSX segnala Internal Linking come priorità ({n} inlink, target ≥ 10). "
            f"Priorità: editoriali di categoria, guide prodotto, pagine hub."
        )
    if not anc or len(anc) < 3:
        return "Anchor text non rilevati o assenti. Aggiungere anchor keyword-rich diversificati."
    return "Stato interno nella norma. Monitorare diversificazione anchor e MoM inlinks."


def build_internal_link_audit(
    wsx_df,
    sf_links_df,
    sf_pages_df   = None,
    ahrefs_df     = None,
):
    """
    A02 — Internal Link Flow Audit.

    Merge WSX Authority + SF All Inlinks + (opz.) SF HTML Pages + (opz.) Ahrefs
    per produrre una lista priorizzata di interventi di link interno.
    """
    prog = st.progress(0, text='Normalizzazione URL…')

    # ── 1. WSX base ────────────────────────────────────────────────────────────
    wsx = wsx_df.copy()
    wsx['_norm'] = wsx['URL'].apply(normalize_url)
    col_metric = _find_col(wsx, 'Top Recommendation - Metric Name')
    col_auth   = _find_col(wsx, 'Authority')
    # Safe column access
    wsx_slim = wsx[['_norm', 'URL', 'Brand', 'Market', 'Language', 'Page Type',
                     col_metric, col_auth]].copy()
    wsx_slim.rename(columns={col_metric: '_wsx_metric', col_auth: '_wsx_authority'}, inplace=True)

    prog.progress(0.15, text='Calcolo inlink da SF All Inlinks…')

    # ── 2. SF All Inlinks ─────────────────────────────────────────────────────
    sf_links = sf_links_df.copy()
    sf_links.columns = [c.strip() for c in sf_links.columns]

    dest_col   = next((c for c in sf_links.columns
                       if c.lower() in ('destination', 'to', 'dst')), sf_links.columns[1])
    src_col    = next((c for c in sf_links.columns
                       if c.lower() in ('source', 'from', 'src')), sf_links.columns[0])
    anchor_col = next((c for c in sf_links.columns
                       if 'anchor' in c.lower()), None)
    pos_col    = next((c for c in sf_links.columns
                       if 'position' in c.lower() or 'link position' in c.lower()), None)

    sf_links['_nd'] = sf_links[dest_col].apply(normalize_url)
    sf_links['_ns'] = sf_links[src_col].apply(normalize_url)

    # Unique source pages per destination (= meaningful inlinks)
    inlink_counts = (
        sf_links[sf_links['_nd'].notna()]
        .groupby('_nd')
        .agg(
            sf_inlinks   = ('_ns', 'nunique'),
            sf_link_occ  = ('_ns', 'count'),
        )
        .reset_index()
        .rename(columns={'_nd': '_norm'})
    )

    # Top-3 anchors per destination
    if anchor_col:
        def _top3(s):
            vals = s.dropna().str.strip().str.lower()
            vals = vals[vals.str.len() > 0]
            return ' | '.join(vals.value_counts().head(3).index.tolist()) if len(vals) else ''

        anchor_df = (
            sf_links.groupby('_nd')[anchor_col]
            .apply(_top3)
            .reset_index()
            .rename(columns={'_nd': '_norm', anchor_col: 'top_anchors'})
        )
    else:
        anchor_df = pd.DataFrame(columns=['_norm', 'top_anchors'])

    # Content-area links only (if position column available)
    if pos_col:
        content_links = sf_links[sf_links[pos_col].str.lower().isin(['content', 'body']) == True]
        content_counts = (
            content_links.groupby('_nd')
            .agg(sf_content_inlinks=('_ns', 'nunique'))
            .reset_index()
            .rename(columns={'_nd': '_norm'})
        )
    else:
        content_counts = pd.DataFrame(columns=['_norm', 'sf_content_inlinks'])

    # Unique destinations per source (outlinks per page, for hub detection)
    outlink_counts = (
        sf_links.groupby('_ns')
        .agg(sf_outlinks=('_nd', 'nunique'))
        .reset_index()
        .rename(columns={'_ns': '_norm'})
    )

    prog.progress(0.40, text='Merge con SF HTML pages…')

    # ── 3. SF HTML Pages (optional) ───────────────────────────────────────────
    sf_pg = None
    if sf_pages_df is not None:
        sf_p = sf_pages_df.copy()
        addr = next((c for c in sf_p.columns if c.lower() in ('address', 'url')), sf_p.columns[0])
        sf_p['_norm'] = sf_p[addr].apply(normalize_url)

        _col_map = {
            'Status Code':       'sf_status_code',
            'Indexability':      'sf_indexability',
            'Inlinks':           'sf_inlinks_rep',    # SF-reported (may differ)
            'Unique Inlinks':    'sf_unique_inlinks',
            'Internal Outlinks': 'sf_int_outlinks',
        }
        keep = {'_norm': '_norm'}
        for src_name, alias in _col_map.items():
            real = _find_col(sf_p, src_name)
            if real in sf_p.columns:
                keep[real] = alias

        sf_pg = sf_p[list(keep.keys())].rename(columns=keep).drop_duplicates('_norm')

    prog.progress(0.60, text='Merge con Ahrefs…')

    # ── 4. Ahrefs Best by Links (optional) ────────────────────────────────────
    ah_norm = None
    if ahrefs_df is not None:
        ah = ahrefs_df.copy()
        url_col = next((c for c in ah.columns if c.lower() in ('url', 'address')), ah.columns[0])
        ah['_norm'] = ah[url_col].apply(normalize_url)

        _ah_map = {
            'UR':                 'ahrefs_ur',
            'URL Rating':         'ahrefs_ur',
            'Referring domains':  'ahrefs_ref_domains',
            'Referring Domains':  'ahrefs_ref_domains',
            'Backlinks':          'ahrefs_backlinks',
        }
        keep_ah = {'_norm': '_norm'}
        used_aliases = set()
        for src_name, alias in _ah_map.items():
            real = _find_col(ah, src_name)
            if real in ah.columns and alias not in used_aliases:
                keep_ah[real] = alias
                used_aliases.add(alias)

        ah_norm = ah[list(keep_ah.keys())].rename(columns=keep_ah).drop_duplicates('_norm')

    prog.progress(0.75, text='Assemblaggio dataset…')

    # ── 5. Merge all on WSX base ───────────────────────────────────────────────
    r = wsx_slim.copy()
    r = r.merge(inlink_counts,  on='_norm', how='left')
    r = r.merge(anchor_df,      on='_norm', how='left')
    r = r.merge(content_counts, on='_norm', how='left')
    r = r.merge(outlink_counts, on='_norm', how='left')
    if sf_pg is not None:
        r = r.merge(sf_pg, on='_norm', how='left')
    if ah_norm is not None:
        r = r.merge(ah_norm, on='_norm', how='left')

    prog.progress(0.88, text='Classificazione e scoring…')

    # ── 6. Classify ────────────────────────────────────────────────────────────
    def _get_outlinks(row):
        # prefer SF reported, fall back to computed
        for col in ('sf_int_outlinks', 'sf_outlinks'):
            if col in row.index and pd.notna(row.get(col)):
                return int(row[col])
        return None

    tipologie, levels, scores, actions = [], [], [], []
    for _, row in r.iterrows():
        n   = int(row.get('sf_inlinks', 0) or 0)
        ur  = float(row.get('ahrefs_ur', 0) or 0)
        out = _get_outlinks(row)
        mn  = str(row.get('_wsx_metric', '') or '')
        anc = str(row.get('top_anchors', '') or '')
        pt  = str(row.get('Page Type', '') or '')

        tip, lvl, sc = _il_classify(n, ur, out, mn)
        act = _il_action_text(tip, n, ur, anc, pt)
        tipologie.append(tip);  levels.append(lvl)
        scores.append(sc);      actions.append(act)

    r['Tipologia']          = tipologie
    r['Priority']           = levels
    r['Priority Score']     = scores
    r['Azione consigliata'] = actions
    r['SF Inlinks']         = r.get('sf_inlinks', pd.Series(0, index=r.index)).fillna(0).astype(int)

    prog.progress(0.97, text='Finalizzazione…')

    # ── 7. Final column selection ──────────────────────────────────────────────
    out_cols = {
        'URL': 'URL', 'Brand': 'Brand', 'Market': 'Market',
        'Language': 'Language', 'Page Type': 'Page Type',
        '_wsx_authority': 'WSX Authority', '_wsx_metric': 'WSX Metric prioritaria',
        'SF Inlinks': 'SF Inlinks (unique src)',
        'sf_link_occ': 'SF Link occurrences',
        'sf_content_inlinks': 'SF Content Inlinks',
        'sf_inlinks_rep': 'SF Inlinks (reportati)',
        'sf_int_outlinks': 'SF Internal Outlinks',
        'sf_status_code': 'Status Code', 'sf_indexability': 'Indexability',
        'ahrefs_ur': 'Ahrefs UR', 'ahrefs_ref_domains': 'Referring Domains',
        'ahrefs_backlinks': 'Backlinks Ahrefs',
        'top_anchors': 'Top Anchor Text (top 3)',
        'Tipologia': 'Tipologia', 'Priority': 'Priority',
        'Priority Score': 'Priority Score',
        'Azione consigliata': 'Azione consigliata',
    }
    keep = [c for c in out_cols if c in r.columns]
    out  = r[keep].rename(columns={c: out_cols[c] for c in keep})

    prog.empty()
    return out.sort_values('Priority Score', ascending=False).reset_index(drop=True)



# ══════════════════════════════════════════════════════════════════════════════
# STRUCTURE ANALYSIS — funzioni (S)
# ══════════════════════════════════════════════════════════════════════════════

# Mappa Metric Name → colonna score schema corrispondente
_SCHEMA_SCORE_COL = {
    'How to Declaration':              'How To',
    'FAQ Declaration':                 'FAQ',
    'Author Declaration':              'Author',
    'Article Declaration':             'Article',
    'Product Info Declaration':        'Product Info',
    'Parent Organization Declaration': 'Parent Organization',
}

# Peso del tipo di schema per il priority score (impatto SEO beauty)
_SCHEMA_WEIGHT = {
    'Product Info Declaration':        40,
    'Parent Organization Declaration': 35,
    'How to Declaration':              25,
    'FAQ Declaration':                 20,
    'Author Declaration':              20,
    'Article Declaration':             15,
}

# Peso del page type per il priority score
_PAGETYPE_WEIGHT = {
    'PDP':          40,
    'Homepage':     35,
    'PLP':          25,
    'Landing Page': 15,
    'Content Page': 10,
}


def _struct_impl_status(schema_score):
    """Classifica lo stato di implementazione da uno score schema WSX."""
    if schema_score is None or pd.isna(schema_score):
        return '— N/A'
    if schema_score == 0:
        return '❌ Non implementato'
    if schema_score < 80:
        return '⚠️ Parziale'
    return '✅ Implementato'


def _struct_priority(status, structure_score, metric_name, page_type):
    """Calcola priorità e score (0-100) per un URL Structure."""
    sw = _SCHEMA_WEIGHT.get(metric_name, 15)
    pw = _PAGETYPE_WEIGHT.get(str(page_type), 10)
    ss = structure_score or 0

    if '❌' in status:
        # Non implementato: peso schema + page type + (100 - struct score) / 3
        score = 0.4 * sw + 0.3 * pw + 0.3 * max(0, 100 - ss) * 0.5
        level = '🔴 Alta' if score >= 28 else '🟡 Media'
    elif '⚠️' in status:
        score = 0.3 * sw + 0.2 * pw + 0.15 * max(0, 100 - ss)
        level = '🟡 Media'
    elif '✅' in status:
        score = 5
        level = '🟢 Bassa'
    else:  # N/A
        score = 0
        level = '— N/A'

    return level, min(int(score), 100)


def _struct_action(status, metric_name, schema_score, page_type):
    """Raccomandazione rule-based per structured data."""
    stype = metric_name.replace(' Declaration', '')
    pt    = str(page_type or '')

    if status == '❌ Non implementato':
        return (
            f"Implementare JSON-LD {stype} ({pt}). "
            f"Adattare il template WSX all'URL e inserire in <head> come <script type=\"application/ld+json\">."
        )
    if status == '⚠️ Parziale':
        return (
            f"Schema {stype} parziale (score: {schema_score}/100). "
            f"Completare le proprietà mancanti secondo i criteri Must Have WSX. "
            f"Validare con Google Rich Results Test."
        )
    if status == '✅ Implementato':
        return f"Schema {stype} implementato. Monitorare aggiornamenti Must Have e validità periodica."
    return f"Schema {stype} non rilevato da WSX per questo page type. Verificare se applicabile."


def _struct_row_worker(args: dict) -> tuple:
    """
    Elabora UNA riga Structure WSX: classificazione + LLM adattamento JSON-LD.
    Thread-safe, nessuna chiamata Streamlit.
    """
    row      = args['row']
    use_llm  = args['use_llm']
    provider = args['provider']
    api_key  = args['api_key']
    model    = args['model']

    url = str(row.get('URL', '') or '')
    if url and not url.startswith('http'):
        url = 'https://www.' + url

    metric    = str(row.get('Top Recommendation - Metric Name', '') or '')
    goal      = str(row.get('Top Recommendation - Goal', '')         or '')
    must_have = str(row.get('Top Recommendation - Must Have', '')     or '')
    jsonld_tmpl = str(row.get('Top Recommendation - JSON-LD Template (if applicable)', '') or '')
    if jsonld_tmpl == 'nan': jsonld_tmpl = ''

    # Schema-specific score
    schema_col   = _SCHEMA_SCORE_COL.get(metric)
    schema_score = _wsx_score(row, schema_col) if schema_col else None
    struct_score = _wsx_score(row, 'Structure')
    page_type    = str(row.get('Page Type', '') or '')

    status            = _struct_impl_status(schema_score)
    priority, p_score = _struct_priority(status, struct_score, metric, page_type)
    action            = _struct_action(status, metric, schema_score, page_type)

    # LLM: adatta JSON-LD all'URL italiana
    jsonld_adapted = ''
    if use_llm and jsonld_tmpl and status in ('❌ Non implementato', '⚠️ Parziale'):
        try:
            prompt = _prompt_jsonld_adapt(
                url=url, jsonld_template=jsonld_tmpl,
                metric_name=metric, must_have=must_have,
            )
            # JSON-LD richiede molti token — usiamo 2000 per evitare troncamenti
            jsonld_adapted = _call_llm(prompt, provider=provider, api_key=api_key,
                                       model=model, max_tokens=2000)
        except Exception:
            pass

    return args['idx'], {
        'URL':                 url,
        'Brand':               row.get('Brand',     ''),
        'Market':              row.get('Market',    ''),
        'Language':            row.get('Language',  ''),
        'Page Type':           page_type,
        # Structure scores
        'Structure Score':     struct_score,
        'Δ Structure MoM':     row.get('Structure MoM', ''),
        # Schema scores (colonne singole)
        'Breadcrumb':          _wsx_score(row, 'Breadcrumb'),
        'Product Info':        _wsx_score(row, 'Product Info'),
        'How To':              _wsx_score(row, 'How To'),
        'Rating':              _wsx_score(row, 'Rating'),
        'Article':             _wsx_score(row, 'Article'),
        'Author':              _wsx_score(row, 'Author'),
        'FAQ':                 _wsx_score(row, 'FAQ'),
        'Organization':        _wsx_score(row, 'Organization'),
        # Classification
        'Schema Type':         metric,
        'Schema Score':        schema_score,
        'Status':              status,
        'Priority':            priority,
        'Priority Score':      p_score,
        'Azione consigliata':  action,
        # Templates
        'JSON-LD Template WSX':    jsonld_tmpl,
        'JSON-LD Adattato (LLM)':  jsonld_adapted,
        # Context
        'Goal':                goal,
        'Must Have':           must_have,
        # Analytics
        'SEO Impressions':     _wsx_score(row, 'SEO Impressions (Google Search Console)'),
        'All Sessions':        _wsx_score(row, 'All Sessions (Google Analytics)'),
    }


def build_structure_analysis(df: pd.DataFrame, use_llm: bool = False,
                             provider: str = 'anthropic', api_key: str = None,
                             model: str = None, n_workers: int = 4) -> pd.DataFrame:
    """S — Structure Schema Analysis con classificazione e adattamento JSON-LD."""
    jobs = [
        {'idx': i, 'row': row, 'use_llm': use_llm,
         'provider': provider, 'api_key': api_key, 'model': model}
        for i, (_, row) in enumerate(df.iterrows())
    ]
    rows = _run_parallel(_struct_row_worker, jobs, n_workers,
                         progress_label='Analisi Structure')
    result = pd.DataFrame(rows)
    return result.sort_values('Priority Score', ascending=False).reset_index(drop=True)



# ══════════════════════════════════════════════════════════════════════════════
# TECHNICAL ANALYSIS — funzioni (T)
# ══════════════════════════════════════════════════════════════════════════════

# Mappa WSX Page Speed score → (status label, priority label, base score)
_PAGESPEED_WSX_MAP = {
    0:   ('❌ Critica',        '🔴 Critica',    90),
    25:  ('⚠️ Insufficiente',  '🔴 Alta',       70),
    50:  ('⚠️ Moderata',       '🟡 Media',      45),
    75:  ('📈 Quasi ottimale', '🟡 Bassa-med',  20),
    100: ('✅ Ottimale',       '🟢 OK',          5),
}

_PAGETYPE_WEIGHT_T = {
    'PDP':          35,
    'Homepage':     30,
    'PLP':          25,
    'Landing Page': 15,
    'Content Page': 10,
}


def _tech_mom_label(mom):
    if mom is None or pd.isna(mom): return '— N/D'
    if mom > 0:  return f'📈 +{mom:.1f}'
    if mom < 0:  return f'📉 {mom:.1f}'
    return '➡ Stabile'


def _tech_priority(speed_score, sitemap_score, metric_name, page_type, mom):
    pw = _PAGETYPE_WEIGHT_T.get(str(page_type), 10)
    if metric_name == 'Sitemap Declaration' and (sitemap_score or 100) == 0:
        return '🔴 Alta', min(55 + int(pw * 0.3), 85)
    sp_key   = int(speed_score or 100)
    base_sc  = _PAGESPEED_WSX_MAP.get(sp_key, _PAGESPEED_WSX_MAP[100])[2]
    mom_val  = float(mom) if (mom is not None and pd.notna(mom)) else 0.0
    mom_bonus = (15 if mom_val < -17.5 else
                  8 if mom_val < 0      else
                 -5 if mom_val > 17.5   else 0)
    total = base_sc + pw * 0.3 + mom_bonus
    if total >= 80: return '🔴 Critica',    min(int(total), 100)
    if total >= 60: return '🔴 Alta',       min(int(total), 100)
    if total >= 35: return '🟡 Media',      min(int(total), 100)
    if total >= 15: return '🟡 Bassa-med',  min(int(total), 100)
    return '🟢 OK', max(5, int(total))


def _tech_action(metric_name, speed_score, sitemap_score, page_type, mom):
    pt      = str(page_type or '')
    mom_val = float(mom) if (mom is not None and pd.notna(mom)) else 0.0
    trend   = (''  if mom_val == 0 else
               f' | ⚠ Regressione MoM: {mom_val:+.1f}.' if mom_val < -8.75 else
               f' | Trend: {mom_val:+.1f} MoM in miglioramento.' if mom_val > 0 else '')
    if metric_name == 'Sitemap Declaration':
        if (sitemap_score or 100) == 0:
            return ("URL assente dal sitemap XML. Verificare Status 200, "
                    "assenza di canonical redirect e robots.txt block.")
        return "URL nel sitemap. Monitorare a ogni rilascio strutturale."
    sp = int(speed_score or 100)
    if sp == 0:
        return (f"TTLB critico ({pt}). Verificare con Screaming Frog Custom Extraction + "
                f"Google CrUX. Azioni: CDN, server-side cache, lazy-load.{trend}")
    if sp == 25:
        return (f"TTLB insufficiente ({pt}). PageSpeed Insights: "
                f"JS/CSS deferral, compressione immagini, cache policy.{trend}")
    if sp == 50:
        return (f"Performance moderata ({pt}). Target TTLB <500ms. "
                f"Verificare LCP e INP in GSC Core Web Vitals.{trend}")
    if sp == 75:
        return (f"Quasi ottimale ({pt}). Focus su INP/CLS. "
                f"Analizzare con RUM data o Chrome DevTools.{trend}")
    return f"Performance ottimale ({pt}). Monitorare MoM per regressioni.{trend}"


def _tech_row_worker(args: dict) -> tuple:
    row      = args['row']
    use_llm  = args['use_llm']
    provider = args['provider']
    api_key  = args['api_key']
    model    = args['model']

    url = str(row.get('URL', '') or '')
    if url and not url.startswith('http'): url = 'https://www.' + url

    metric    = str(row.get('Top Recommendation - Metric Name', '') or '')
    goal      = str(row.get('Top Recommendation - Goal', '')         or '')
    must_have = str(row.get('Top Recommendation - Must Have', '')     or '')

    speed_score   = _wsx_score(row, 'Page Speed')
    sitemap_score = _wsx_score(row, 'Sitemap Declaration')
    tech_score    = _wsx_score(row, 'Technicals')
    page_type     = str(row.get('Page Type', '') or '')
    mom_raw       = row.get('Technicals MoM', None)
    mom           = float(mom_raw) if (mom_raw is not None and pd.notna(mom_raw)) else None

    sp_key       = int(speed_score or 100)
    speed_status = _PAGESPEED_WSX_MAP.get(sp_key, _PAGESPEED_WSX_MAP[100])[0]
    sitemap_st   = ('— N/D' if sitemap_score is None
                    else '❌ Assente' if sitemap_score == 0
                    else '✅ Presente')

    priority, p_score = _tech_priority(speed_score, sitemap_score, metric, page_type, mom)
    action = _tech_action(metric, speed_score, sitemap_score, page_type, mom)

    llm_action = ''
    if use_llm and metric:
        try:
            prompt = _prompt_tech_action(
                url=url, page_type=page_type,
                speed_score=speed_score or 100, mom=mom or 0.0,
                technicals_score=tech_score or 0,
                must_have=must_have, metric_name=metric,
            )
            llm_action = _call_llm(prompt, provider=provider, api_key=api_key, model=model)
        except Exception:
            pass

    return args['idx'], {
        'URL':               url,
        'Brand':             row.get('Brand',     ''),
        'Market':            row.get('Market',    ''),
        'Language':          row.get('Language',  ''),
        'Page Type':         page_type,
        'Technicals Score':  tech_score,
        'Δ MoM':             mom,
        'MoM Trend':         _tech_mom_label(mom),
        'Page Speed (WSX)':  speed_score,
        'Speed Status':      speed_status,
        'Sitemap Score':     sitemap_score,
        'Sitemap Status':    sitemap_st,
        'Metric Name':       metric,
        'Priority':          priority,
        'Priority Score':    p_score,
        'Azione consigliata': action,
        'Azione LLM':        llm_action,
        'SEO Impressions':   _wsx_score(row, 'SEO Impressions (Google Search Console)'),
        'All Sessions':      _wsx_score(row, 'All Sessions (Google Analytics)'),
        'Goal':              goal,
        'Must Have':         must_have,
    }


def build_technical_analysis(df: pd.DataFrame, use_llm: bool = False,
                             provider: str = 'anthropic', api_key: str = None,
                             model: str = None, n_workers: int = 4) -> pd.DataFrame:
    """T — Technical SEO analysis: Page Speed + Sitemap con priorità e trend MoM."""
    jobs = [
        {'idx': i, 'row': row, 'use_llm': use_llm,
         'provider': provider, 'api_key': api_key, 'model': model}
        for i, (_, row) in enumerate(df.iterrows())
    ]
    rows = _run_parallel(_tech_row_worker, jobs, n_workers,
                         progress_label='Analisi Technical')
    return pd.DataFrame(rows).sort_values('Priority Score', ascending=False).reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════════════════
#  CAST FRAMEWORK
# ══════════════════════════════════════════════════════════════════════════════
_tab_c, _tab_a, _tab_s, _tab_t = st.tabs([
    '📝 C — Content',
    '🔗 A — Authority',
    '🏗️ S — Structured',
    '⚙️ T — Technical',
])

with _tab_c:
    tab1, tab2, tab3 = st.tabs([
        '01 · Analisi Suggerimenti',
        '02 · Confronto Crawl SF',
        '03 · Confronto Export WSX',
    ])

    with tab1:
        with st.expander('ℹ️ File e API richiesti', expanded=False):
            st.markdown("""
| # | File | Formato | Colonne chiave |
|---|------|---------|----------------|
| 1 | **WSX Context/URL Performance export** | `.xlsx` | `URL`, `Language`, `Page Type`, `Meta Tags`, `Heading`, `Relevance`, `Top Recommendation - Customized Suggestion`, `Top Recommendation - Metric Name`, `Top Recommendation - Goal`, `Top Recommendation - Must Have` |

**API LLM** (opzionale — attiva con il toggle 🤖):
- **Anthropic**: `ANTHROPIC_API_KEY` in `.env` o campo API Key — usato per traduzione/ottimizzazione dei tag suggeriti
- **OpenAI**: `OPENAI_API_KEY` — alternativa ad Anthropic

**Modalità LLM:**
- *Traduzione*: traduce i valori estratti in italiano (veloce, ~1 call/tag)
- *Ottimizzazione*: genera tag SEO-ready dal contesto completo (qualità superiore)
- *Completa*: genera Title + Description per ogni URL con suggerimento
""")
        st.markdown('<div class="section-label">Input</div>', unsafe_allow_html=True)

        c1, _ = st.columns([3, 1])
        with c1:
            wsx_file = st.file_uploader('File WSX export (.xlsx)', type='xlsx', key='wsx1')

        if wsx_file:
            df = pd.read_excel(wsx_file)
            st.session_state['wsx_df'] = df

        if st.session_state['wsx_df'] is not None:
            df = st.session_state['wsx_df']

            use_llm = st.toggle('🤖 Attiva LLM', key='trans1', value=True,
                                help='Usa il provider configurato in alto per tradurre/ottimizzare i tag')

            if st.button('▶ Analizza suggerimenti', key='run1'):
                with st.spinner(''):
                    result = build_wsx_analysis(
                        df, translate=use_llm,
                        provider=_PROVIDER, api_key=_API_KEY,
                        mode=_MODE, model=_MODEL,
                        n_workers=_N_WORKERS,
                    )
                    st.session_state['result_df'] = result

            if st.session_state['result_df'] is not None:
                res = st.session_state['result_df']

                # ── Riepilogo ──
                st.markdown('<div class="section-label">Riepilogo</div>', unsafe_allow_html=True)
                cov_counts = res['suggestion_coverage'].value_counts()
                metrics_row(res, {
                    'URL totali':          len(res),
                    'Title + Desc':        cov_counts.get('Title + Description', 0),
                    'Solo Description':    cov_counts.get('Description only', 0),
                    'Solo Title':          cov_counts.get('Title only', 0),
                    'Solo H1':             cov_counts.get('H1 only', 0),
                    'Nessun suggerimento': cov_counts.get('No suggestion', 0),
                })
                st.markdown('<br>', unsafe_allow_html=True)

                # ── Distribuzione Metric Name ──
                st.markdown('<div class="section-label">Distribuzione per Metric Name</div>', unsafe_allow_html=True)
                mc1, mc2 = st.columns(2)
                with mc1:
                    mn_counts = res['Metric Name'].value_counts().reset_index()
                    mn_counts.columns = ['Metric Name', 'URL']
                    st.dataframe(mn_counts, use_container_width=True, hide_index=True)
                with mc2:
                    at_counts = res['Action Type'].value_counts().reset_index()
                    at_counts.columns = ['Action Type', 'URL']
                    st.dataframe(at_counts, use_container_width=True, hide_index=True)

                # ── Filtri ──
                st.markdown('<div class="section-label">Esplora</div>', unsafe_allow_html=True)
                fc1, fc2, fc3, fc4 = st.columns(4)
                with fc1:
                    f_metric = st.multiselect(
                        'Metric Name',
                        options=res['Metric Name'].unique().tolist(),
                        default=res['Metric Name'].unique().tolist(), key='f1_metric')
                with fc2:
                    f_cov = st.multiselect(
                        'Coverage',
                        options=res['suggestion_coverage'].unique().tolist(),
                        default=res['suggestion_coverage'].unique().tolist(), key='f1_cov')
                with fc3:
                    f_brand = st.multiselect(
                        'Brand',
                        options=res['Brand'].unique().tolist(),
                        default=res['Brand'].unique().tolist(), key='f1_brand')
                with fc4:
                    f_status = st.multiselect(
                        'Title status',
                        options=['OK ✓', 'TOO LONG', 'TOO SHORT', ''],
                        default=['OK ✓', 'TOO LONG', 'TOO SHORT', ''], key='f1_status')

                filtered = res[
                    res['Metric Name'].isin(f_metric) &
                    res['suggestion_coverage'].isin(f_cov) &
                    res['Brand'].isin(f_brand) &
                    res['sug_title_status'].isin(f_status)
                ]
                st.caption(f'{len(filtered)} URL selezionati')
                st.dataframe(filtered, use_container_width=True, hide_index=True,
                    column_config={
                        'URL':                   st.column_config.LinkColumn('URL', width=300),
                        'Metric Name':           st.column_config.TextColumn('Metric Name', width=160),
                        'Score Context':         st.column_config.NumberColumn('Context', width=70),
                        'Score Meta Tags':       st.column_config.NumberColumn('Meta Tags', width=70),
                        'Score Heading':         st.column_config.NumberColumn('Heading', width=70),
                        'Score Relevance':       st.column_config.NumberColumn('Relevance', width=70),
                        'suggested_title':       st.column_config.TextColumn('Suggested Title', width=280),
                        'suggested_description': st.column_config.TextColumn('Suggested Desc', width=380),
                        'suggested_h1':          st.column_config.TextColumn('Suggested H1', width=240),
                        'Goal':                  st.column_config.TextColumn('Goal', width=300),
                        'Must Have':             st.column_config.TextColumn('Must Have', width=300),
                    })

                # Goal / Must Have expandable per metric
                if len(filtered):
                    with st.expander('📋 Goal & Must Have per Metric Name'):
                        for mn in filtered['Metric Name'].unique():
                            sub = filtered[filtered['Metric Name'] == mn]
                            if sub.empty: continue
                            first = sub.iloc[0]
                            st.markdown(f"**{mn}**")
                            st.markdown(f"*Goal:* {first.get('Goal','—')}")
                            st.markdown(f"*Must Have:* {first.get('Must Have','—')}")
                            st.divider()

                st.download_button(
                    '⬇ Scarica risultati (.xlsx)', data=to_excel_bytes(filtered),
                    file_name='meta_tag_analysis.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # ──────────────────────────────────────────────────────────────────────────────
    # TAB 2  ·  Confronto Crawl SF
    # ──────────────────────────────────────────────────────────────────────────────
    with tab2:
        with st.expander('ℹ️ File e API richiesti', expanded=False):
            st.markdown("""
| # | File | Formato | Colonne chiave |
|---|------|---------|----------------|
| 1 | **WSX Context/URL Performance export** | `.xlsx` | stesso file di C01 |
| 2 | **Screaming Frog — Internal HTML export** | `.xlsx` | `Address`, `Title 1`, `Title 1 Length`, `Meta Description 1`, `Meta Description 1 Length`, `H1-1`, `H1-1 Length`, `Status Code`, `Indexability` |

**Come esportare da Screaming Frog:**
1. Completa il crawl del dominio
2. Tab `Internal` → filtra per `HTML`
3. `Export` → salva come `.xlsx`

**API LLM**: stessa configurazione di C01.
""")
        st.markdown('<div class="section-label">Input</div>', unsafe_allow_html=True)

        r1, r2 = st.columns(2)
        with r1:
            wsx_file2 = st.file_uploader('File WSX export (.xlsx)', type='xlsx', key='wsx2')
        with r2:
            sf_file2  = st.file_uploader('File Screaming Frog crawl (.xlsx)', type='xlsx', key='sf2')

        use_llm2 = st.toggle('🤖 Attiva LLM', key='trans2', value=True)

        if wsx_file2:
            st.session_state['wsx_df'] = pd.read_excel(wsx_file2)
        if sf_file2:
            st.session_state['sf_df'] = pd.read_excel(sf_file2)

        wsx2 = st.session_state.get('wsx_df')
        sf2  = st.session_state.get('sf_df')

        if wsx2 is not None and sf2 is not None:
            if st.button('▶ Avvia confronto', key='run2'):
                with st.spinner(''):
                    comp = build_comparison(
                        wsx2, sf2, translate=use_llm2,
                        provider=_PROVIDER, api_key=_API_KEY,
                        mode=_MODE, model=_MODEL,
                        n_workers=_N_WORKERS,
                    )
                    st.session_state['comp_df'] = comp

            comp_df = st.session_state.get('comp_df')
            if comp_df is not None:
                st.markdown('<div class="section-label">Riepilogo</div>', unsafe_allow_html=True)

                matched   = (comp_df['SF Match'] == 'YES').sum()
                unmatched = (comp_df['SF Match'] == 'NOT IN CRAWL').sum()
                t_diff    = (comp_df['title_diff'] == '⚠ Diverso').sum()
                d_diff    = (comp_df['desc_diff']  == '⚠ Diverso').sum()
                h1_diff_n = (comp_df['h1_diff']    == '⚠ Diverso').sum()
                any_diff  = (comp_df['tags_differing'] > 0).sum()

                metrics_row(comp_df, {
                    'URL matchate SF':    matched,
                    'Non nel crawl':      unmatched,
                    'Title diverso':      t_diff,
                    'Desc diversa':       d_diff,
                    'H1 diverso':         h1_diff_n,
                    'Con almeno 1 diff':  any_diff,
                })
                st.markdown('<br>', unsafe_allow_html=True)

                # ── Filtri ──
                st.markdown('<div class="section-label">Esplora</div>', unsafe_allow_html=True)
                f1, f2, f3, f4, f5 = st.columns(5)
                with f1:
                    f_match  = st.multiselect('SF Match', comp_df['SF Match'].unique().tolist(),
                                              default=comp_df['SF Match'].unique().tolist(), key='f2_match')
                with f2:
                    f_met2   = st.multiselect('Metric Name', comp_df['Metric Name'].unique().tolist(),
                                              default=comp_df['Metric Name'].unique().tolist(), key='f2_met')
                with f3:
                    f_cov2   = st.multiselect('Coverage', comp_df['suggestion_coverage'].unique().tolist(),
                                              default=comp_df['suggestion_coverage'].unique().tolist(), key='f2_cov')
                with f4:
                    f_tstat  = st.multiselect('Title status (attuale)',
                                              options=['OK ✓','TOO LONG','TOO SHORT',''],
                                              default=['OK ✓','TOO LONG','TOO SHORT',''], key='f2_ts')
                with f5:
                    f_diff_only = st.toggle('Solo tag diversi', key='f2_diff', value=False)

                mask = (
                    comp_df['SF Match'].isin(f_match) &
                    comp_df['Metric Name'].isin(f_met2) &
                    comp_df['suggestion_coverage'].isin(f_cov2) &
                    comp_df['current_title_status'].isin(f_tstat)
                )
                if f_diff_only:
                    mask = mask & (comp_df['tags_differing'] > 0)

                filtered2 = comp_df[mask]
                st.caption(f'{len(filtered2)} URL selezionati')

                st.dataframe(filtered2, use_container_width=True, hide_index=True,
                    column_config={
                        'URL':                   st.column_config.LinkColumn('URL', width=280),
                        'Metric Name':           st.column_config.TextColumn('Metric', width=150),
                        'Score Meta Tags':       st.column_config.NumberColumn('⭐ Meta Tags', width=80),
                        'Score Heading':         st.column_config.NumberColumn('⭐ Heading', width=70),
                        'tags_differing':        st.column_config.NumberColumn('# Diff', width=55),
                        'title_diff':            st.column_config.TextColumn('Title', width=130),
                        'current_title':         st.column_config.TextColumn('Title Attuale', width=260),
                        'suggested_title':       st.column_config.TextColumn('Title Suggerito', width=260),
                        'desc_diff':             st.column_config.TextColumn('Desc', width=130),
                        'current_description':   st.column_config.TextColumn('Desc Attuale', width=320),
                        'suggested_description': st.column_config.TextColumn('Desc Suggerita', width=320),
                        'h1_diff':               st.column_config.TextColumn('H1', width=130),
                        'current_h1':            st.column_config.TextColumn('H1 Attuale', width=220),
                        'suggested_h1':          st.column_config.TextColumn('H1 Suggerito', width=220),
                    })

                # ── Report tag diversi ──
                tag_diff_df = comp_df[comp_df['tags_differing'] > 0]
                if len(tag_diff_df):
                    with st.expander(f'⚠ Report {len(tag_diff_df)} URL con tag diversi da SF'):
                        st.dataframe(tag_diff_df[[
                            'URL', 'Metric Name', 'tags_differing',
                            'title_diff', 'current_title', 'suggested_title',
                            'desc_diff', 'current_description', 'suggested_description',
                            'h1_diff', 'current_h1', 'suggested_h1',
                        ]], use_container_width=True, hide_index=True)

                st.download_button(
                    '⬇ Scarica confronto (.xlsx)', data=to_excel_bytes(filtered2),
                    file_name='meta_tag_comparison.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            st.info('Carica entrambi i file per procedere al confronto.')

    # ──────────────────────────────────────────────────────────────────────────────
    # TAB 3  ·  Confronto Export WSX
    # ──────────────────────────────────────────────────────────────────────────────
    with tab3:
        with st.expander('ℹ️ File richiesti', expanded=False):
            st.markdown("""
| # | File | Formato | Note |
|---|------|---------|------|
| 1 | **WSX Context export NEW** (mese corrente) | `.xlsx` | Export più recente |
| 2 | **WSX Context export OLD** (mese precedente) | `.xlsx` | Export del mese precedente per il confronto |

Nessuna API richiesta. Lo script confronta: URL nuove/rimosse, variazione degli score (Meta Tags, Heading, Context, Relevance), variazione dei tag suggeriti e stato dell'implementazione (score migliorato = tag implementato).
""")
        st.markdown('<div class="section-label">Input — confronta due elaborazioni WSX</div>', unsafe_allow_html=True)
        st.caption('Carica il file più recente (NEW) e quello precedente (OLD) per identificare URL invariati, aggiornati o nuovi.')

        d1, d2 = st.columns(2)
        with d1:
            st.markdown('**Export NEW** *(elaborazione corrente)*')
            wsx_new_file = st.file_uploader('File WSX nuovo (.xlsx)', type='xlsx', key='wsx_new')
        with d2:
            st.markdown('**Export OLD** *(elaborazione precedente)*')
            wsx_old_file = st.file_uploader('File WSX precedente (.xlsx)', type='xlsx', key='wsx_old')

        if wsx_new_file: st.session_state['wsx_df']      = pd.read_excel(wsx_new_file)
        if wsx_old_file: st.session_state['wsx_prev_df'] = pd.read_excel(wsx_old_file)

        df_new = st.session_state.get('wsx_df')
        df_old = st.session_state.get('wsx_prev_df')

        if df_new is not None and df_old is not None:
            if st.button('▶ Confronta export', key='run3'):
                with st.spinner(''):
                    diff = build_diff(df_new, df_old)
                    st.session_state['diff_df'] = diff

            diff_df = st.session_state.get('diff_df')
            if diff_df is not None:
                st.markdown('<div class="section-label">Riepilogo</div>', unsafe_allow_html=True)
                sc = diff_df['status'].value_counts()
                metrics_row(diff_df, {
                    'URL totali (new)':        (diff_df['status'] != 'removed_from_new').sum(),
                    'Nuove URL':               sc.get('new_url', 0),
                    'Suggerimento aggiornato': sc.get('suggestion_updated', 0),
                    'Tag aggiornati ✓':        (diff_df['Tags Status'] == 'Aggiornati ✓').sum(),
                    'Invariati':               sc.get('suggestion_unchanged', 0),
                    'URL rimosse (old)':       sc.get('removed_from_new', 0),
                })
                st.markdown('<br>', unsafe_allow_html=True)

                diff_df['Status Label'] = diff_df['status'].apply(status_icon)

                st.markdown('<div class="section-label">Filtro</div>', unsafe_allow_html=True)
                df1, df2, df3 = st.columns(3)
                with df1:
                    all_statuses = diff_df['Status Label'].unique().tolist()
                    f_status3 = st.multiselect('Status', all_statuses, default=all_statuses, key='f3_status')
                with df2:
                    f_met3 = st.multiselect(
                        'Metric Name',
                        options=diff_df['Metric Name'].unique().tolist(),
                        default=diff_df['Metric Name'].unique().tolist(), key='f3_met')
                with df3:
                    f_tags3 = st.multiselect(
                        'Tags Status',
                        options=diff_df['Tags Status'].unique().tolist(),
                        default=diff_df['Tags Status'].unique().tolist(), key='f3_tags')

                filtered3 = diff_df[
                    diff_df['Status Label'].isin(f_status3) &
                    diff_df['Metric Name'].isin(f_met3) &
                    diff_df['Tags Status'].isin(f_tags3)
                ]
                st.caption(f'{len(filtered3)} URL selezionati')

                st.dataframe(filtered3[[
                    'URL', 'Brand', 'Market', 'Metric Name', 'Status Label', 'Tags Status',
                    'Δ Meta Tags', 'Meta Tags (new)', 'Meta Tags (old)',
                    'Δ Heading',   'Heading (new)',   'Heading (old)',
                    'Δ Context',   'Context (new)',   'Context (old)',
                    'title_changed', 'new_suggested_title', 'old_suggested_title',
                    'desc_changed',  'new_suggested_desc',  'old_suggested_desc',
                    'h1_changed',    'new_suggested_h1',    'old_suggested_h1',
                ]], use_container_width=True, hide_index=True,
                column_config={
                    'URL':                 st.column_config.LinkColumn('URL', width=260),
                    'Status Label':        st.column_config.TextColumn('Status', width=160),
                    'Tags Status':         st.column_config.TextColumn('Tag', width=110),
                    'Metric Name':         st.column_config.TextColumn('Metric', width=140),
                    'Δ Meta Tags':         st.column_config.TextColumn('Δ MT', width=50),
                    'Δ Heading':           st.column_config.TextColumn('Δ H', width=50),
                    'Δ Context':           st.column_config.TextColumn('Δ Ctx', width=55),
                    'title_changed':       st.column_config.TextColumn('Δ T', width=40),
                    'desc_changed':        st.column_config.TextColumn('Δ D', width=40),
                    'h1_changed':          st.column_config.TextColumn('Δ H1', width=45),
                    'new_suggested_title': st.column_config.TextColumn('Title NEW', width=240),
                    'old_suggested_title': st.column_config.TextColumn('Title OLD', width=240),
                    'new_suggested_desc':  st.column_config.TextColumn('Desc NEW', width=300),
                    'old_suggested_desc':  st.column_config.TextColumn('Desc OLD', width=300),
                    'new_suggested_h1':    st.column_config.TextColumn('H1 NEW', width=220),
                    'old_suggested_h1':    st.column_config.TextColumn('H1 OLD', width=220),
                })

                # URL con suggerimento invariato (da verificare manualmente)
                unchanged = diff_df[diff_df['status'] == 'suggestion_unchanged']
                if len(unchanged):
                    with st.expander(f'🟡 {len(unchanged)} URL con suggerimento invariato — verifica se i tag sono stati implementati'):
                        st.dataframe(unchanged[[
                            'URL', 'Metric Name', 'Tags Status',
                            'Δ Meta Tags', 'Meta Tags (new)', 'Meta Tags (old)',
                            'Δ Heading', 'Heading (new)', 'Heading (old)',
                            'new_suggested_title', 'new_suggested_desc', 'new_suggested_h1',
                        ]], use_container_width=True, hide_index=True)

                st.download_button(
                    '⬇ Scarica diff completo (.xlsx)', data=to_excel_bytes(filtered3),
                    file_name='meta_tag_diff.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            st.info('Carica entrambi gli export WSX per procedere al confronto.')

# ──────────────────────────────────────────────────────────────────────────────
# A — Authority
# ──────────────────────────────────────────────────────────────────────────────
with _tab_a:
    _a01, _a02 = st.tabs(['A01 · Authority Score Analysis', 'A02 · Internal Link Flow Audit'])

    # ─────────────────────────────────────────────────────────────────────────
    # A01  ·  Authority Score Analysis (existing)
    # ─────────────────────────────────────────────────────────────────────────
    with _a01:
        with st.expander('ℹ️ File e API richiesti', expanded=False):
            st.markdown("""
| # | File | Formato | Colonne chiave |
|---|------|---------|----------------|
| 1 | **WSX Authority/URL Performance export** | `.xlsx` | `URL`, `Brand`, `Market`, `Language`, `Page Type`, `Authority`, `Authority MoM`, `Internal Linking`, `Backlinking Quality`, `Backlinking Quantity`, `Content Freshness`, `Reviews Count`, `Reviews Value`, `Top Recommendation - Metric Name`, `Top Recommendation - Goal`, `Top Recommendation - Must Have` |

**API LLM** (opzionale — toggle 🤖):
- Genera una raccomandazione Authority in italiano per ogni URL (max 150 caratteri)
- Usa il provider e il modello configurati nel pannello ⚙️
""")
        st.markdown('<div class="section-label">Input</div>', unsafe_allow_html=True)

        _a1, _a2 = st.columns([3, 1])
        with _a1:
            auth_file = st.file_uploader(
                'File WSX Authority export (.xlsx)', type='xlsx', key='auth_file')
        with _a2:
            use_llm_a = st.toggle('🤖 Attiva LLM', key='trans_a', value=False,
                                  help='Genera azioni consigliate in italiano per ogni URL')

        if auth_file:
            st.session_state['auth_df'] = pd.read_excel(auth_file)

        if st.session_state.get('auth_df') is not None:
            _adf = st.session_state['auth_df']

            if st.button('▶ Analizza Authority', key='run_a'):
                with st.spinner(''):
                    _ar = build_authority_analysis(
                        _adf, use_llm=use_llm_a,
                        provider=_PROVIDER, api_key=_API_KEY,
                        model=_MODEL, n_workers=_N_WORKERS,
                    )
                    st.session_state['auth_result_df'] = _ar
        if st.session_state.get('auth_result_df') is not None:
            _res_a = st.session_state['auth_result_df']

            # ── Riepilogo ──
            st.markdown('<div class="section-label">Riepilogo</div>', unsafe_allow_html=True)
            _pr = _res_a['Priority'].value_counts()
            metrics_row(_res_a, {
                'URL totali':           len(_res_a),
                '🔴 Alta priorità':     _pr.get('🔴 Alta', 0),
                '🟡 Media priorità':    _pr.get('🟡 Media', 0),
                '🟢 Bassa priorità':    _pr.get('🟢 Bassa', 0),
                'Authority media':      round(_res_a['Authority'].dropna().mean(), 1) if 'Authority' in _res_a else '—',
            })
            st.markdown('<br>', unsafe_allow_html=True)

            # ── Distribuzione per Metric Name ──
            st.markdown('<div class="section-label">Distribuzione per Metric Name</div>', unsafe_allow_html=True)
            _mc1, _mc2 = st.columns(2)
            with _mc1:
                _mn_a = _res_a['Metric Name'].value_counts().reset_index()
                _mn_a.columns = ['Metric Name', 'URL']
                st.dataframe(_mn_a, use_container_width=True, hide_index=True)
            with _mc2:
                _pr_a = _res_a['Priority'].value_counts().reset_index()
                _pr_a.columns = ['Priority', 'URL']
                st.dataframe(_pr_a, use_container_width=True, hide_index=True)

            # ── Filtri ──
            st.markdown('<div class="section-label">Esplora</div>', unsafe_allow_html=True)
            _fa1, _fa2, _fa3, _fa4 = st.columns(4)
            with _fa1:
                _f_mn = st.multiselect('Metric Name',
                    options=_res_a['Metric Name'].unique().tolist(),
                    default=_res_a['Metric Name'].unique().tolist(), key='fa_mn')
            with _fa2:
                _f_pr = st.multiselect('Priority',
                    options=_res_a['Priority'].unique().tolist(),
                    default=_res_a['Priority'].unique().tolist(), key='fa_pr')
            with _fa3:
                _f_br = st.multiselect('Brand',
                    options=_res_a['Brand'].unique().tolist(),
                    default=_res_a['Brand'].unique().tolist(), key='fa_br')
            with _fa4:
                _auth_min, _auth_max = st.slider(
                    'Authority score range', 0, 100, (0, 100), key='fa_auth')

            _filtered_a = _res_a[
                _res_a['Metric Name'].isin(_f_mn) &
                _res_a['Priority'].isin(_f_pr) &
                _res_a['Brand'].isin(_f_br) &
                _res_a['Authority'].between(_auth_min, _auth_max)
            ]
            st.caption(f'{len(_filtered_a)} URL selezionati')

            st.dataframe(_filtered_a, use_container_width=True, hide_index=True,
                column_config={
                    'URL':                  st.column_config.LinkColumn('URL', width=280),
                    'Priority':             st.column_config.TextColumn('Priority', width=90),
                    'Authority':            st.column_config.NumberColumn('Authority', width=75),
                    'Δ Authority MoM':      st.column_config.NumberColumn('Δ MoM', width=60),
                    'Internal Linking':     st.column_config.NumberColumn('Int. Link', width=75),
                    'Backlinking Quality':  st.column_config.NumberColumn('BL Quality', width=80),
                    'Backlinking Quantity': st.column_config.NumberColumn('BL Qty', width=70),
                    'Content Freshness':    st.column_config.NumberColumn('Freshness', width=75),
                    'Reviews Count':        st.column_config.NumberColumn('Rev. #', width=60),
                    'Reviews Value':        st.column_config.NumberColumn('Rev. Val', width=70),
                    'Metric Name':          st.column_config.TextColumn('Metric da migliorare', width=160),
                    'Metric Score':         st.column_config.NumberColumn('Score metrica', width=95),
                    'Azione consigliata':   st.column_config.TextColumn('Azione consigliata', width=300),
                    'Must Have':            st.column_config.TextColumn('Must Have', width=300),
                })

            st.download_button(
                '⬇ Scarica analisi Authority (.xlsx)',
                data=to_excel_bytes(_filtered_a),
                file_name='authority_analysis.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        else:
            st.info("Carica il file WSX Authority export per procedere all'analisi.")


    with _a02:

        # ── Documentazione export ────────────────────────────────────────────
        with st.expander('📋 Documentazione export richiesti — leggi prima di caricare i file', expanded=True):
            st.markdown("""
### Export richiesti per l'analisi

---

#### 🕷️ Screaming Frog — Export 1: All Inlinks (OBBLIGATORIO)
Contiene tutte le coppie sorgente→destinazione dei link interni: è la base dell'analisi.

**Come esportarlo:**
1. Esegui il crawl completo del dominio (`File > Crawl`)
2. Attendi completamento crawl
3. Vai su **`Bulk Export`** (menu in alto) → **`All Inlinks`**
4. Salva come **.xlsx** o .csv

**Colonne necessarie** (SF le include di default):
| Colonna | Descrizione |
|---------|------------|
| `Source` | URL della pagina che contiene il link |
| `Destination` | URL della pagina destinataria del link |
| `Anchor` | Testo del link (anchor text) |
| `Link Position` | Posizione nel DOM: Content / Nav / Header / Footer |
| `Follow` | True = link seguito dai bot, False = nofollow |
| `Type` | Content / Image |

> ⚠️ Per siti con 800+ pagine questo export può avere decine di migliaia di righe: è normale.

---

#### 🕷️ Screaming Frog — Export 2: Internal HTML (OPZIONALE — arricchisce l'analisi)
Fornisce per ogni pagina il numero di inlink segnalati da SF, lo status code e l'indexability.

**Come esportarlo:**
1. Stesso crawl del punto precedente
2. Vai sulla tab **`Internal`** (in alto)
3. Filtra per **`HTML`** nel dropdown "Content Type"
4. Clicca **`Export`** (icona freccia) → salva come **.xlsx**

**Colonne necessarie:**
| Colonna | Descrizione |
|---------|------------|
| `Address` | URL della pagina |
| `Status Code` | 200, 301, 404… |
| `Indexability` | Indexable / Non-Indexable |
| `Inlinks` | N° pagine che linkano questa |
| `Unique Inlinks` | N° pagine uniche (deduplicato) |
| `Internal Outlinks` | N° link interni in uscita da questa pagina |

---

#### 🔍 Ahrefs — Export: Best by Links (OPZIONALE — aggiunge URL Rating)
Fornisce l'autorità a livello URL (UR) e i referring domain per pagina.

**Come esportarlo:**
1. Apri **Ahrefs Site Explorer** → inserisci il dominio
2. Menu laterale → **`Pages`** → **`Best by Links`**
3. Imposta **Mode: Exact URL** e **Date: oggi**
4. Clicca **`Export`** → **Full Export (.xlsx)**

**Colonne necessarie:**
| Colonna | Descrizione |
|---------|------------|
| `URL` | URL della pagina |
| `UR` | URL Rating 0-100 (equivalente di Trust Flow per Ahrefs) |
| `Referring domains` | N° domini unici che linkano questa pagina |
| `Backlinks` | N° totale backlink |
| `Traffic` | Traffico organico stimato (opzionale) |

> 📌 **Nota su CF/TF vs UR:** i Must Have WSX citano CF/TF (metriche Majestic). In Ahrefs l'equivalente funzionale è **UR (URL Rating)**. Un UR > 30 corrisponde approssimativamente a TF > 20.
""", unsafe_allow_html=False)

        st.markdown('<div class="section-label">Carica i file</div>', unsafe_allow_html=True)
        _il1, _il2, _il3 = st.columns(3)
        with _il1:
            _wsx_a02  = st.file_uploader('WSX Authority (.xlsx)', type='xlsx', key='il_wsx',
                                          help='Lo stesso file caricato in A01')
        with _il2:
            _sf_links = st.file_uploader('SF All Inlinks (.xlsx/.csv)', type=['xlsx','csv'], key='il_sf_links',
                                          help='Bulk Export > All Inlinks da Screaming Frog')
            _sf_pages = st.file_uploader('SF Internal HTML (.xlsx) — opzionale', type=['xlsx','csv'], key='il_sf_pages',
                                          help='Export tab Internal > HTML da Screaming Frog')
        with _il3:
            _ah_file  = st.file_uploader('Ahrefs Best by Links (.xlsx) — opzionale', type=['xlsx','csv'], key='il_ahrefs',
                                          help='Pages > Best by Links da Ahrefs Site Explorer')

        # Load into session state
        if _wsx_a02:
            st.session_state['il_wsx_df'] = pd.read_excel(_wsx_a02)
        if _sf_links:
            st.session_state['il_sf_links_df'] = (
                pd.read_excel(_sf_links) if _sf_links.name.endswith('.xlsx')
                else pd.read_csv(_sf_links, low_memory=False)
            )
        if _sf_pages:
            st.session_state['il_sf_pages_df'] = (
                pd.read_excel(_sf_pages) if _sf_pages.name.endswith('.xlsx')
                else pd.read_csv(_sf_pages, low_memory=False)
            )
        if _ah_file:
            st.session_state['il_ahrefs_df'] = (
                pd.read_excel(_ah_file) if _ah_file.name.endswith('.xlsx')
                else pd.read_csv(_ah_file, low_memory=False)
            )

        _il_wsx_ready   = st.session_state.get('il_wsx_df') is not None
        _il_links_ready = st.session_state.get('il_sf_links_df') is not None

        if _il_wsx_ready and _il_links_ready:
            if st.button('▶ Avvia Internal Link Audit', key='run_il'):
                with st.spinner(''):
                    _il_result = build_internal_link_audit(
                        wsx_df     = st.session_state['il_wsx_df'],
                        sf_links_df= st.session_state['il_sf_links_df'],
                        sf_pages_df= st.session_state.get('il_sf_pages_df'),
                        ahrefs_df  = st.session_state.get('il_ahrefs_df'),
                    )
                    st.session_state['il_result_df'] = _il_result

            if st.session_state.get('il_result_df') is not None:
                _ilr = st.session_state['il_result_df']

                # ── Riepilogo ──
                st.markdown('<div class="section-label">Riepilogo</div>', unsafe_allow_html=True)
                _tc = _ilr['Tipologia'].value_counts() if 'Tipologia' in _ilr.columns else {}
                _pc = _ilr['Priority'].value_counts()  if 'Priority'  in _ilr.columns else {}
                metrics_row(_ilr, {
                    'URL analizzate':          len(_ilr),
                    '🔴 Orfane':               _tc.get('Orfana', 0),
                    '🔴 Quasi-orfane':          _tc.get('Quasi-orfana', 0),
                    '🟠 Hub silenziosi':        _tc.get('Hub silenzioso', 0),
                    '🟡 Sottolinkatе (WSX)':    _tc.get('Sottolinkata (WSX)', 0),
                    '🟢 OK':                    _tc.get('OK', 0),
                })
                st.markdown('<br>', unsafe_allow_html=True)

                # ── Filtri ──
                st.markdown('<div class="section-label">Esplora</div>', unsafe_allow_html=True)
                _if1, _if2, _if3, _if4 = st.columns(4)
                with _if1:
                    _tip_opts = _ilr['Tipologia'].unique().tolist() if 'Tipologia' in _ilr.columns else []
                    _f_tip = st.multiselect('Tipologia', _tip_opts, default=_tip_opts, key='il_f_tip')
                with _if2:
                    _pri_opts = _ilr['Priority'].unique().tolist() if 'Priority' in _ilr.columns else []
                    _f_pri = st.multiselect('Priority', _pri_opts, default=_pri_opts, key='il_f_pri')
                with _if3:
                    _br_opts = _ilr['Brand'].unique().tolist() if 'Brand' in _ilr.columns else []
                    _f_br2 = st.multiselect('Brand', _br_opts, default=_br_opts, key='il_f_br')
                with _if4:
                    _show_only_fix = st.toggle('Solo da correggere', key='il_fix', value=False,
                                               help='Nasconde le URL con tipologia OK')

                _il_mask = (
                    _ilr['Tipologia'].isin(_f_tip) &
                    _ilr['Priority'].isin(_f_pri) &
                    _ilr['Brand'].isin(_f_br2)
                )
                if _show_only_fix:
                    _il_mask = _il_mask & (_ilr['Tipologia'] != 'OK')

                _il_filtered = _ilr[_il_mask]
                st.caption(f'{len(_il_filtered)} URL selezionati · {(_il_mask & (_ilr["Tipologia"]=="Orfana")).sum()} orfane')

                _col_cfg = {
                    'URL':                       st.column_config.LinkColumn('URL', width=280),
                    'Priority':                  st.column_config.TextColumn('Priority', width=100),
                    'Priority Score':             st.column_config.NumberColumn('Score', width=55),
                    'Tipologia':                 st.column_config.TextColumn('Tipologia', width=140),
                    'WSX Authority':             st.column_config.NumberColumn('WSX Auth', width=75),
                    'WSX Metric prioritaria':    st.column_config.TextColumn('WSX Metric', width=150),
                    'SF Inlinks (unique src)':   st.column_config.NumberColumn('SF Inlinks', width=80),
                    'SF Content Inlinks':        st.column_config.NumberColumn('Content IL', width=80),
                    'SF Internal Outlinks':      st.column_config.NumberColumn('SF Outlinks', width=85),
                    'Ahrefs UR':                 st.column_config.NumberColumn('UR', width=55),
                    'Referring Domains':         st.column_config.NumberColumn('Ref. Dom.', width=70),
                    'Top Anchor Text (top 3)':   st.column_config.TextColumn('Anchor Text', width=220),
                    'Azione consigliata':        st.column_config.TextColumn('Azione', width=340),
                }

                st.dataframe(_il_filtered, use_container_width=True, hide_index=True,
                             column_config=_col_cfg)

                # ── Focus: orfane ──
                _orphans = _ilr[_ilr['Tipologia'] == 'Orfana']
                if len(_orphans):
                    with st.expander(f'🔴 {len(_orphans)} pagine orfane — 0 inlink rilevati da SF'):
                        st.dataframe(_orphans[['URL', 'Page Type', 'WSX Authority',
                                               'WSX Metric prioritaria', 'Ahrefs UR',
                                               'Referring Domains', 'Azione consigliata']],
                                     use_container_width=True, hide_index=True)

                st.download_button(
                    '⬇ Scarica Internal Link Audit (.xlsx)',
                    data=to_excel_bytes(_il_filtered),
                    file_name='internal_link_audit.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        else:
            _missing = []
            if not _il_wsx_ready:   _missing.append('WSX Authority export')
            if not _il_links_ready: _missing.append('SF All Inlinks export')
            st.info(f"Carica i file obbligatori per avviare l'audit: {', '.join(_missing)}.")


with _tab_s:
    with st.expander('ℹ️ File e API richiesti', expanded=False):
        st.markdown("""
| # | File | Formato | Colonne chiave |
|---|------|---------|----------------|
| 1 | **WSX Structure/URL Performance export** | `.xlsx` | `URL`, `Brand`, `Market`, `Language`, `Page Type`, `Structure`, `Structure MoM`, `Breadcrumb`, `Product Info`, `How To`, `Rating`, `Article`, `Author`, `FAQ`, `Organization`, `Top Recommendation - Metric Name`, `Top Recommendation - Goal`, `Top Recommendation - Must Have`, `Top Recommendation - JSON-LD Template (if applicable)` |

**Score schema** (colonne numeriche): `0` = schema assente, `0-80` = parziale, `≥80` = implementato, `null` = non applicabile per questo page type.

**API LLM** (opzionale — toggle 🤖 *Adatta JSON-LD*):
- Adatta il template JSON-LD WSX all'URL italiana reale (sostituisce `@id`, `mainEntityOfPage`, dominio)
- Richiede **2000 token** per output — usa il modello più capace disponibile
- **Anthropic** claude-sonnet-4-6 o **OpenAI** gpt-4o sono consigliati per questa funzione
- Output: JSON-LD pronto per il copy-paste in `<script type="application/ld+json">`
""")
    st.markdown('<div class="section-label">Input</div>', unsafe_allow_html=True)

    _s1, _s2 = st.columns([3, 1])
    with _s1:
        struct_file = st.file_uploader(
            'File WSX Structure export (.xlsx)', type='xlsx', key='struct_file')
    with _s2:
        use_llm_s = st.toggle('🤖 Adatta JSON-LD con LLM', key='trans_s', value=False,
                              help=(
                                  'Adatta il template JSON-LD WSX all\'URL italiana reale. '
                                  'Usa il provider configurato in ⚙️ Configurazione LLM.'
                              ))

    if struct_file:
        st.session_state['struct_df'] = pd.read_excel(struct_file)

    if st.session_state.get('struct_df') is not None:
        _sdf = st.session_state['struct_df']

        if st.button('▶ Analizza Structure', key='run_s'):
            with st.spinner(''):
                _sr = build_structure_analysis(
                    _sdf, use_llm=use_llm_s,
                    provider=_PROVIDER, api_key=_API_KEY,
                    model=_MODEL, n_workers=_N_WORKERS,
                )
                st.session_state['struct_result_df'] = _sr

        if st.session_state.get('struct_result_df') is not None:
            _res_s = st.session_state['struct_result_df']

            # ── Riepilogo ──────────────────────────────────────────────────────
            st.markdown('<div class="section-label">Riepilogo</div>', unsafe_allow_html=True)
            _st_cnt = _res_s['Status'].value_counts()
            _pr_cnt = _res_s['Priority'].value_counts()
            metrics_row(_res_s, {
                'URL analizzate':         len(_res_s),
                '❌ Non implementati':    _st_cnt.get('❌ Non implementato', 0),
                '⚠️ Parziali':           _st_cnt.get('⚠️ Parziale', 0),
                '✅ Implementati':        _st_cnt.get('✅ Implementato', 0),
                '🔴 Alta priorità':       _pr_cnt.get('🔴 Alta', 0),
                'Structure media':        round(_res_s['Structure Score'].dropna().mean(), 1),
            })
            st.markdown('<br>', unsafe_allow_html=True)

            # ── Matrice Schema × Page Type ──────────────────────────────────
            st.markdown('<div class="section-label">Matrice Schema × Page Type</div>', unsafe_allow_html=True)
            _sm1, _sm2 = st.columns(2)
            with _sm1:
                _schema_dist = _res_s['Schema Type'].value_counts().reset_index()
                _schema_dist.columns = ['Schema Type', 'URL']
                st.dataframe(_schema_dist, use_container_width=True, hide_index=True)
            with _sm2:
                _status_dist = (
                    _res_s.groupby(['Schema Type', 'Status'])
                    .size().reset_index(name='URL')
                )
                st.dataframe(_status_dist, use_container_width=True, hide_index=True)

            # ── Filtri ──────────────────────────────────────────────────────────
            st.markdown('<div class="section-label">Esplora</div>', unsafe_allow_html=True)
            _sf1, _sf2, _sf3, _sf4 = st.columns(4)
            with _sf1:
                _f_schema = st.multiselect('Schema Type',
                    options=_res_s['Schema Type'].unique().tolist(),
                    default=_res_s['Schema Type'].unique().tolist(), key='sf_schema')
            with _sf2:
                _f_status_s = st.multiselect('Status',
                    options=_res_s['Status'].unique().tolist(),
                    default=_res_s['Status'].unique().tolist(), key='sf_status')
            with _sf3:
                _f_pr_s = st.multiselect('Priority',
                    options=_res_s['Priority'].unique().tolist(),
                    default=_res_s['Priority'].unique().tolist(), key='sf_pr')
            with _sf4:
                _f_only_missing = st.toggle('Solo da implementare', key='sf_miss', value=False,
                                            help='Mostra solo ❌ Non implementato e ⚠️ Parziale')

            _s_mask = (
                _res_s['Schema Type'].isin(_f_schema) &
                _res_s['Status'].isin(_f_status_s) &
                _res_s['Priority'].isin(_f_pr_s)
            )
            if _f_only_missing:
                _s_mask = _s_mask & _res_s['Status'].isin(['❌ Non implementato', '⚠️ Parziale'])

            _s_filtered = _res_s[_s_mask]
            st.caption(f'{len(_s_filtered)} URL selezionati')

            _s_col_cfg = {
                'URL':                  st.column_config.LinkColumn('URL', width=270),
                'Page Type':            st.column_config.TextColumn('Page Type', width=100),
                'Schema Type':          st.column_config.TextColumn('Schema', width=185),
                'Schema Score':         st.column_config.NumberColumn('Schema Score', width=95),
                'Status':               st.column_config.TextColumn('Status', width=145),
                'Priority':             st.column_config.TextColumn('Priority', width=90),
                'Priority Score':       st.column_config.NumberColumn('P.Score', width=60),
                'Structure Score':      st.column_config.NumberColumn('Struct.', width=60),
                'Breadcrumb':           st.column_config.NumberColumn('BreadC.', width=65),
                'Product Info':         st.column_config.NumberColumn('ProdInfo', width=70),
                'How To':               st.column_config.NumberColumn('HowTo', width=58),
                'Rating':               st.column_config.NumberColumn('Rating', width=58),
                'Article':              st.column_config.NumberColumn('Article', width=60),
                'Author':               st.column_config.NumberColumn('Author', width=58),
                'FAQ':                  st.column_config.NumberColumn('FAQ', width=50),
                'Azione consigliata':   st.column_config.TextColumn('Azione', width=320),
            }

            # Columns to display (exclude JSON-LD text columns from table)
            _display_cols = [c for c in _s_filtered.columns
                             if c not in ('JSON-LD Template WSX', 'JSON-LD Adattato (LLM)',
                                          'Goal', 'Must Have')]
            st.dataframe(_s_filtered[_display_cols], use_container_width=True,
                         hide_index=True, column_config=_s_col_cfg)

            # ── JSON-LD Inspector ──────────────────────────────────────────────
            st.markdown('<div class="section-label">JSON-LD Inspector</div>', unsafe_allow_html=True)
            st.caption('Seleziona un URL per visualizzare il template JSON-LD WSX e la versione adattata dall\'LLM.')

            _jl_urls = _s_filtered['URL'].tolist()
            if _jl_urls:
                _sel_url = st.selectbox('URL', _jl_urls, key='jl_url_sel')
                _sel_row = _s_filtered[_s_filtered['URL'] == _sel_url].iloc[0]

                _jl_c1, _jl_c2 = st.columns(2)
                with _jl_c1:
                    st.markdown('**Template WSX** (originale)')
                    _tmpl = _sel_row.get('JSON-LD Template WSX', '')
                    if _tmpl and str(_tmpl).strip() and str(_tmpl) != 'nan':
                        try:
                            _pretty = _json.dumps(_json.loads(str(_tmpl)), indent=2, ensure_ascii=False)
                        except Exception:
                            _pretty = str(_tmpl)
                        st.code(_pretty, language='json')
                    else:
                        st.info('Nessun template disponibile per questa URL.')

                with _jl_c2:
                    _adapted = _sel_row.get('JSON-LD Adattato (LLM)', '')
                    if _adapted and str(_adapted).strip() and str(_adapted) != 'nan':
                        st.markdown('**JSON-LD Adattato** (LLM — pronto per l\'implementazione)')
                        try:
                            _pretty_a = _json.dumps(_json.loads(str(_adapted)), indent=2, ensure_ascii=False)
                        except Exception:
                            _pretty_a = str(_adapted)
                        st.code(_pretty_a, language='json')
                    else:
                        st.markdown('**JSON-LD Adattato** (LLM)')
                        st.info(
                            'Attiva il toggle **🤖 Adatta JSON-LD con LLM** e riesegui '
                            'l\'analisi per ottenere il JSON-LD adattato all\'URL italiana.'
                        )

                # Must Have e Goal nell'expander
                with st.expander(f'📋 Must Have & Goal — {_sel_row.get("Schema Type", "")}'):
                    st.markdown(f"**Goal:** {_sel_row.get('Goal', '—')}")
                    st.markdown('---')
                    st.markdown(f"**Must Have:**\n\n{_sel_row.get('Must Have', '—')}")

            # ── Focus: non implementati alta priorità ──────────────────────────
            _critical = _res_s[
                (_res_s['Status'] == '❌ Non implementato') &
                (_res_s['Priority'] == '🔴 Alta')
            ]
            if len(_critical):
                with st.expander(f'🔴 {len(_critical)} URL — Schema mancante e alta priorità'):
                    st.dataframe(
                        _critical[['URL', 'Page Type', 'Schema Type', 'Structure Score',
                                   'Priority Score', 'Azione consigliata']],
                        use_container_width=True, hide_index=True)

            st.download_button(
                '⬇ Scarica analisi Structure (.xlsx)',
                data=to_excel_bytes(_s_filtered),
                file_name='structure_analysis.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    else:
        st.info("Carica il file WSX Structure export per procedere all'analisi.")


# ──────────────────────────────────────────────────────────────────────────────
# T — Technical
# ──────────────────────────────────────────────────────────────────────────────
with _tab_t:
    with st.expander('ℹ️ File e API richiesti', expanded=False):
        st.markdown("""
| # | File | Formato | Colonne chiave |
|---|------|---------|----------------|
| 1 | **WSX Technicals/URL Performance export** | `.xlsx` | `URL`, `Brand`, `Market`, `Language`, `Page Type`, `Technicals`, `Technicals MoM`, `Page Speed`, `Sitemap Declaration`, `Page Type Tagging`, `Valid Inlinks`, `Top Recommendation - Metric Name`, `Top Recommendation - Goal`, `Top Recommendation - Must Have` |

**Score `Page Speed`** (scala WSX): `0` = TTLB critico (>2s), `25` = insufficiente, `50` = moderato, `75` = quasi ottimale, `100` = TTLB <500ms.
**Score `Sitemap Declaration`**: `0` = URL assente dal sitemap XML, `100` = presente.
**Colonne semplificate** (sempre 100/100 in questo dataset): `Page Type Tagging`, `Valid Inlinks` — presenti nell'analisi ma non influenzano il priority score.

**API LLM** (opzionale — toggle 🤖):
- Genera raccomandazioni tecniche specifiche per URL in italiano (max 180 caratteri)
- Tool citati nel prompt: Screaming Frog Custom Extraction, Google CrUX, PageSpeed Insights, GSC Core Web Vitals
""")
    st.markdown('<div class="section-label">Input</div>', unsafe_allow_html=True)

    _t1, _t2 = st.columns([3, 1])
    with _t1:
        tech_file = st.file_uploader(
            'File WSX Technical export (.xlsx)', type='xlsx', key='tech_file')
    with _t2:
        use_llm_t = st.toggle('🤖 Attiva LLM', key='trans_t', value=False,
                              help='Genera raccomandazioni tecniche specifiche per URL')

    if tech_file:
        st.session_state['tech_df'] = pd.read_excel(tech_file)

    if st.session_state.get('tech_df') is not None:
        _tdf = st.session_state['tech_df']

        if st.button('▶ Analizza Technical', key='run_t'):
            with st.spinner(''):
                _tr = build_technical_analysis(
                    _tdf, use_llm=use_llm_t,
                    provider=_PROVIDER, api_key=_API_KEY,
                    model=_MODEL, n_workers=_N_WORKERS,
                )
                st.session_state['tech_result_df'] = _tr

        if st.session_state.get('tech_result_df') is not None:
            _res_t = st.session_state['tech_result_df']

            # ── Riepilogo ──────────────────────────────────────────────────────
            st.markdown('<div class="section-label">Riepilogo</div>', unsafe_allow_html=True)
            _sp_cnt = _res_t['Page Speed (WSX)'].value_counts().sort_index()
            _pr_cnt = _res_t['Priority'].value_counts()
            _mom_neg = (_res_t['Δ MoM'].dropna() < 0).sum()
            _sitemap_ko = (_res_t['Sitemap Status'] == '❌ Assente').sum()

            metrics_row(_res_t, {
                'URL analizzate':        len(_res_t),
                '❌ Page Speed critico':  int(_sp_cnt.get(0, 0)) + int(_sp_cnt.get(25, 0)),
                '⚠️ Page Speed moderato': int(_sp_cnt.get(50, 0)),
                '✅ Page Speed ottimale': int(_sp_cnt.get(100, 0)),
                '📉 MoM in regressione': int(_mom_neg),
                '🗺 Sitemap mancante':   int(_sitemap_ko),
            })
            st.markdown('<br>', unsafe_allow_html=True)

            # ── Page Speed × Page Type matrix ──────────────────────────────────
            st.markdown('<div class="section-label">Page Speed × Page Type</div>',
                        unsafe_allow_html=True)
            _tm1, _tm2 = st.columns([2, 1])
            with _tm1:
                _pt_matrix = (
                    _res_t.groupby(['Page Type', 'Page Speed (WSX)'])
                    .size().unstack(fill_value=0)
                    .reindex(columns=[0, 25, 50, 75, 100], fill_value=0)
                )
                _pt_matrix.columns = ['Score 0 ❌', 'Score 25 ⚠️', 'Score 50 ⚠️',
                                      'Score 75 📈', 'Score 100 ✅']
                st.dataframe(_pt_matrix, use_container_width=True)
            with _tm2:
                _mom_dist = _res_t['MoM Trend'].value_counts().reset_index()
                _mom_dist.columns = ['MoM Trend', 'URL']
                st.dataframe(_mom_dist, use_container_width=True, hide_index=True)

            # ── Filtri ──────────────────────────────────────────────────────────
            st.markdown('<div class="section-label">Esplora</div>', unsafe_allow_html=True)
            _tf1, _tf2, _tf3, _tf4, _tf5 = st.columns(5)
            with _tf1:
                _f_mn_t = st.multiselect('Metric Name',
                    options=_res_t['Metric Name'].unique().tolist(),
                    default=_res_t['Metric Name'].unique().tolist(), key='tf_mn')
            with _tf2:
                _f_sp = st.multiselect('Page Speed (WSX)',
                    options=sorted(_res_t['Page Speed (WSX)'].dropna().unique().tolist()),
                    default=sorted(_res_t['Page Speed (WSX)'].dropna().unique().tolist()),
                    key='tf_sp')
            with _tf3:
                _f_pr_t = st.multiselect('Priority',
                    options=_res_t['Priority'].unique().tolist(),
                    default=_res_t['Priority'].unique().tolist(), key='tf_pr')
            with _tf4:
                _f_pt_t = st.multiselect('Page Type',
                    options=_res_t['Page Type'].unique().tolist(),
                    default=_res_t['Page Type'].unique().tolist(), key='tf_pt')
            with _tf5:
                _f_mom_neg = st.toggle('Solo in regressione', key='tf_mom', value=False,
                                       help='Mostra solo URL con MoM < 0')

            _t_mask = (
                _res_t['Metric Name'].isin(_f_mn_t) &
                _res_t['Page Speed (WSX)'].isin(_f_sp) &
                _res_t['Priority'].isin(_f_pr_t) &
                _res_t['Page Type'].isin(_f_pt_t)
            )
            if _f_mom_neg:
                _t_mask = _t_mask & (_res_t['Δ MoM'].fillna(0) < 0)

            _t_filtered = _res_t[_t_mask]
            st.caption(f'{len(_t_filtered)} URL selezionati')

            _t_col_cfg = {
                'URL':              st.column_config.LinkColumn('URL', width=270),
                'Page Type':        st.column_config.TextColumn('Page Type', width=100),
                'Priority':         st.column_config.TextColumn('Priority', width=105),
                'Priority Score':   st.column_config.NumberColumn('P.Score', width=60),
                'Page Speed (WSX)': st.column_config.NumberColumn('Speed WSX', width=80),
                'Speed Status':     st.column_config.TextColumn('Speed', width=140),
                'Δ MoM':            st.column_config.NumberColumn('Δ MoM', width=60),
                'MoM Trend':        st.column_config.TextColumn('Trend', width=100),
                'Sitemap Status':   st.column_config.TextColumn('Sitemap', width=100),
                'Technicals Score': st.column_config.NumberColumn('Technical', width=75),
                'Metric Name':      st.column_config.TextColumn('Metrica', width=150),
                'Azione consigliata': st.column_config.TextColumn('Azione', width=360),
                'Azione LLM':       st.column_config.TextColumn('Azione LLM', width=300),
            }

            _display_t = [c for c in _t_filtered.columns
                          if c not in ('Goal', 'Must Have')]
            st.dataframe(_t_filtered[_display_t], use_container_width=True,
                         hide_index=True, column_config=_t_col_cfg)

            # ── Focus: critici in regressione ─────────────────────────────────
            _critical_t = _res_t[
                (_res_t['Page Speed (WSX)'].isin([0, 25])) &
                (_res_t['Δ MoM'].fillna(0) < 0)
            ].sort_values('Priority Score', ascending=False)

            if len(_critical_t):
                with st.expander(
                    f'🔴 {len(_critical_t)} URL — Page Speed critico/insufficiente + regressione MoM',
                    expanded=True
                ):
                    st.dataframe(
                        _critical_t[['URL', 'Page Type', 'Page Speed (WSX)', 'Speed Status',
                                     'Δ MoM', 'MoM Trend', 'Technicals Score',
                                     'Priority', 'Azione consigliata']],
                        use_container_width=True, hide_index=True)

            # ── Sitemap audit ──────────────────────────────────────────────────
            _sitemap_ko_df = _res_t[_res_t['Sitemap Status'] == '❌ Assente']
            if len(_sitemap_ko_df):
                with st.expander(
                    f'🗺 {len(_sitemap_ko_df)} URL assenti dal sitemap XML'
                ):
                    st.markdown(
                        "**Must Have WSX:** ogni URL indexable deve essere presente "
                        "nel sitemap XML. Solo Status 200, no redirect 301 o pagine 404."
                    )
                    st.dataframe(
                        _sitemap_ko_df[['URL', 'Page Type', 'Technicals Score',
                                        'Priority', 'Azione consigliata']],
                        use_container_width=True, hide_index=True)

            # ── Must Have & Goal expander ──────────────────────────────────────
            _unique_metrics_t = _res_t['Metric Name'].dropna().unique().tolist()
            with st.expander('📋 Must Have & Goal per metrica'):
                for mn in _unique_metrics_t:
                    sub = _res_t[_res_t['Metric Name'] == mn]
                    if sub.empty: continue
                    r0  = sub.iloc[0]
                    st.markdown(f"**{mn}**")
                    st.markdown(f"*Goal:* {r0.get('Goal', '—')}")
                    st.markdown(f"*Must Have:* {r0.get('Must Have', '—')}")
                    st.divider()

            st.download_button(
                '⬇ Scarica analisi Technical (.xlsx)',
                data=to_excel_bytes(_t_filtered),
                file_name='technical_analysis.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    else:
        st.info("Carica il file WSX Technical export per procedere all'analisi.")

