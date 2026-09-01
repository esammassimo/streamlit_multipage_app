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
from typing import Dict, List, Optional, Tuple
import requests
from bs4 import BeautifulSoup
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
    # ALL — Recommendations (registro cross-CAST)
    'rec_files', 'rec_all_df', 'rec_tags_df',
    'rec_clu_items', 'rec_clu_summary',
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
#  RECOMMENDATION ENGINE  —  esplosione di TUTTE le raccomandazioni WSX
# ══════════════════════════════════════════════════════════════════════════════
#  Il formato export WSX 2026 non contiene più le colonne
#  `Top Recommendation - *` (una sola raccomandazione per URL): contiene
#  invece lo score di OGNI sotto-metrica del pilastro. Ogni sotto-metrica
#  sotto soglia è a tutti gli effetti una raccomandazione.
#
#  Questo blocco trasforma il formato wide (1 riga = 1 URL, N colonne score)
#  in formato long (1 riga = 1 raccomandazione), calcola severità e priorità
#  e genera il testo dell'azione consigliata.
#
#  Funzioni pure: nessuna chiamata Streamlit, eseguibili anche in thread pool.
# ══════════════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════════════
#  PILASTRI CAST
# ══════════════════════════════════════════════════════════════════════════════

PILLARS: Dict[str, dict] = {
    'C': {
        'code': 'C', 'name': 'Context',
        'label': 'C — Context',
        'score_col': 'Context - Score',
        'mom_col': 'Context MoM',
        'file_token': 'Context',
        'icon': '📝',
        'desc': 'Qualità e pertinenza del contenuto on-page: meta tag, '
                'heading, rilevanza, unicità, correttezza linguistica.',
    },
    'A': {
        'code': 'A', 'name': 'Authority',
        'label': 'A — Authority',
        'score_col': 'Authority - Score',
        'mom_col': 'Authority MoM',
        'file_token': 'Authority',
        'icon': '🔗',
        'desc': "Autorevolezza della pagina: link interni, profilo backlink, "
                "segnali di fiducia (recensioni, rating) e freschezza.",
    },
    'S': {
        'code': 'S', 'name': 'Structure',
        'label': 'S — Structure',
        'score_col': 'Structure - Score',
        'mom_col': 'Structure MoM',
        'file_token': 'Structure',
        'icon': '🏗️',
        'desc': 'Dati strutturati JSON-LD: copertura e completezza degli '
                'schema richiesti dal page type.',
    },
    'T': {
        'code': 'T', 'name': 'Technicals',
        'label': 'T — Technicals',
        'score_col': 'Technicals - Score',
        'mom_col': 'Technicals MoM',
        'file_token': 'Technicals',
        'icon': '⚙️',
        'desc': 'Salute tecnica dell’URL: presenza in sitemap, performance '
                'di caricamento, validità dei link in ingresso.',
    },
}

# Prefissi file → brand (estendibile)
BRAND_PREFIX: Dict[str, str] = {
    'BIO': 'Biotherm',
    'VIC': 'Vichy',
    'LRP': 'La Roche-Posay',
    'CER': 'CeraVe',
    'SKC': 'SkinCeuticals',
    'LAN': 'Lancôme',
    'KIE': 'Kiehl’s',
}

# Prefisso file → pilastro
PILLAR_PREFIX: Dict[str, str] = {'C': 'C', 'A': 'A', 'S': 'S', 'T': 'T'}

# Peso del page type sulla priorità (0-100)
PAGETYPE_WEIGHT: Dict[str, int] = {
    'Homepage':     40,
    'PDP':          35,
    'PLP':          30,
    'Landing Page': 20,
    'Content Page': 15,
}
PAGETYPE_WEIGHT_DEFAULT = 10


# ══════════════════════════════════════════════════════════════════════════════
#  CATALOGO RACCOMANDAZIONI  —  24 metriche WSX classificate per CAST
# ══════════════════════════════════════════════════════════════════════════════
#  Ogni voce:
#    pillar      : C | A | S | T
#    id          : codice raccomandazione stabile (es. C01)
#    weight      : peso della metrica nel calcolo di priorità (0-100)
#    kind        : 'content' | 'link' | 'schema' | 'tech'  (tipo di intervento)
#    owner       : team che esegue l'intervento
#    applies_to  : page type per cui la metrica è attesa ([] = tutti)
#    goal        : obiettivo WSX della metrica
#    must_have   : criterio di conformità WSX
#    action      : template azione consigliata; supporta {score} {gap}
#                  {page_type} {url} {status}
# ══════════════════════════════════════════════════════════════════════════════

METRIC_CATALOG: Dict[str, dict] = {

    # ─────────────────────────────── C — CONTEXT ──────────────────────────────
    'Meta Tags': {
        'pillar': 'C', 'id': 'C01', 'weight': 30, 'kind': 'content',
        'owner': 'SEO / Content',
        'applies_to': [],
        'goal': 'Title e meta description unici, di lunghezza corretta e '
                'allineati alla keyword primaria della pagina.',
        'must_have': 'Title 50-60 caratteri, Meta Description 140-155 caratteri, '
                     'keyword primaria presente, nessun duplicato nel sito.',
        'action': 'Riscrivere Title e Meta Description (score {score}/100). '
                  'Portare il Title a 50-60 caratteri e la Description a 140-155, '
                  'inserire la keyword primaria in apertura ed evitare duplicazioni '
                  'con le altre {page_type}.',
    },
    'Relevance': {
        'pillar': 'C', 'id': 'C02', 'weight': 25, 'kind': 'content',
        'owner': 'Content',
        'applies_to': [],
        'goal': 'Il contenuto della pagina copre in modo esaustivo l’intento '
                'di ricerca della keyword primaria.',
        'must_have': 'Keyword primaria e varianti semantiche nel body, copertura '
                     'delle domande correlate, volume di testo utile adeguato al page type.',
        'action': 'Ampliare e riallineare il contenuto all’intento di ricerca '
                  '(score {score}/100, gap {gap} punti). Integrare varianti semantiche, '
                  'sezioni FAQ e copy descrittivo sopra la fold della {page_type}.',
    },
    'Heading': {
        'pillar': 'C', 'id': 'C03', 'weight': 20, 'kind': 'content',
        'owner': 'SEO / Front-end',
        'applies_to': [],
        'goal': 'Gerarchia dei titoli corretta e semanticamente descrittiva.',
        'must_have': 'Un solo H1 per pagina, gerarchia H1→H2→H3 senza salti di livello, '
                     'keyword primaria nell’H1 e varianti negli H2.',
        'action': 'Correggere la struttura degli heading (score {score}/100). '
                  'Verificare la presenza di un unico H1, eliminare i salti di livello '
                  'e rendere gli H2 descrittivi rispetto alle sezioni di contenuto.',
    },
    'Grammar': {
        'pillar': 'C', 'id': 'C04', 'weight': 10, 'kind': 'content',
        'owner': 'Content / Copy',
        'applies_to': [],
        'goal': 'Testo privo di errori ortografici, grammaticali e di refusi.',
        'must_have': 'Nessun errore rilevato dal controllo linguistico WSX sulla '
                     'lingua dichiarata nella pagina.',
        'action': 'Revisione linguistica del copy (score {score}/100). '
                  'Correggere refusi ed errori grammaticali; verificare che la lingua '
                  'del testo coincida con quella dichiarata nell’attributo lang.',
    },
    'Unique Content': {
        'pillar': 'C', 'id': 'C05', 'weight': 15, 'kind': 'content',
        'owner': 'Content / SEO',
        'applies_to': [],
        'goal': 'Contenuto originale, non duplicato rispetto ad altre URL del dominio.',
        'must_have': 'Nessuna sovrapposizione sostanziale di testo con altre pagine '
                     'del sito o con la scheda prodotto di partenza.',
        'action': 'Riscrivere il contenuto duplicato (score {score}/100). '
                  'Differenziare descrizione e copy rispetto alle pagine simili; '
                  'in alternativa valutare canonical o consolidamento dell’URL.',
    },

    # ────────────────────────────── A — AUTHORITY ─────────────────────────────
    'Internal Linking': {
        'pillar': 'A', 'id': 'A01', 'weight': 30, 'kind': 'link',
        'owner': 'SEO',
        'applies_to': [],
        'goal': 'La pagina riceve un numero adeguato di link interni con anchor pertinenti.',
        'must_have': 'Inlink da pagine tematicamente affini, anchor text descrittiva '
                     'e non generica, profondità di click contenuta rispetto alla home.',
        'action': 'Rafforzare il link interno verso l’URL (score {score}/100, '
                  'gap {gap} punti). Aggiungere link contestuali da PLP, articoli e hub '
                  'di territorio correlati, con anchor descrittive e non generiche.',
    },
    'Backlinking Quality': {
        'pillar': 'A', 'id': 'A02', 'weight': 20, 'kind': 'link',
        'owner': 'SEO / Digital PR',
        'applies_to': [],
        'goal': 'Il profilo backlink della pagina proviene da domini autorevoli e pertinenti.',
        'must_have': 'Domini referenti con autorevolezza e affinità tematica adeguate, '
                     'assenza di pattern di link a basso valore.',
        'action': 'Migliorare la qualità del profilo backlink (score {score}/100). '
                  'Pianificare attività di digital PR su domini autorevoli del settore '
                  'beauty/skincare e disconoscere i referenti a basso valore.',
    },
    'Backlinking Quantity': {
        'pillar': 'A', 'id': 'A03', 'weight': 15, 'kind': 'link',
        'owner': 'SEO / Digital PR',
        'applies_to': [],
        'goal': 'La pagina dispone di un numero di domini referenti adeguato alla concorrenza.',
        'must_have': 'Numero di referring domain in linea con le pagine competitor '
                     'posizionate sulle stesse keyword.',
        'action': 'Aumentare il numero di domini referenti (score {score}/100). '
                  'Inserire l’URL nei piani di link earning e nelle citazioni di prodotto; '
                  'valutare redirect e consolidamento dei link storici.',
    },
    'Reviews Count': {
        'pillar': 'A', 'id': 'A04', 'weight': 15, 'kind': 'content',
        'owner': 'E-commerce / CRM',
        'applies_to': ['PDP'],
        'goal': 'La scheda prodotto raccoglie un volume di recensioni sufficiente '
                'a generare segnali di fiducia e rich result.',
        'must_have': 'Recensioni verificate visibili in pagina e dichiarate '
                     'in AggregateRating nel JSON-LD.',
        'action': 'Incrementare la raccolta recensioni sulla scheda (score {score}/100). '
                  'Attivare i flussi post-acquisto di sollecito e verificare che le '
                  'recensioni raccolte siano esposte anche nel markup AggregateRating.',
    },
    'Rating Value': {
        'pillar': 'A', 'id': 'A05', 'weight': 10, 'kind': 'content',
        'owner': 'E-commerce',
        'applies_to': ['PDP'],
        'goal': 'Il rating medio del prodotto è competitivo e correttamente esposto.',
        'must_have': 'Valore medio del rating presente in pagina e coerente con '
                     'il valore dichiarato nello structured data.',
        'action': 'Verificare ed esporre il rating medio (score {score}/100). '
                  'Controllare la coerenza tra il valore mostrato in pagina e '
                  'ratingValue nel JSON-LD; presidiare i prodotti con media bassa.',
    },
    'Content Freshness': {
        'pillar': 'A', 'id': 'A06', 'weight': 10, 'kind': 'content',
        'owner': 'Content',
        'applies_to': ['Content Page', 'Landing Page'],
        'goal': 'Il contenuto è aggiornato di recente e ne è dichiarata la data.',
        'must_have': 'dateModified valorizzato e coerente, aggiornamento sostanziale '
                     'del contenuto entro l’orizzonte previsto per la tipologia.',
        'action': 'Aggiornare il contenuto e la data di modifica (score {score}/100). '
                  'Rivedere dati, immagini e riferimenti stagionali, poi valorizzare '
                  'dateModified nel JSON-LD e nel CMS.',
    },

    # ────────────────────────────── S — STRUCTURE ─────────────────────────────
    'Organization': {
        'pillar': 'S', 'id': 'S01', 'weight': 20, 'kind': 'schema',
        'owner': 'Front-end / SEO',
        'applies_to': ['Homepage'],
        'goal': 'Entità di brand dichiarata a motori e LLM tramite schema Organization.',
        'must_have': 'JSON-LD Organization con name, url, logo, sameAs dei profili '
                     'ufficiali e contactPoint.',
        'action': 'Implementare o completare il JSON-LD Organization (score {score}/100). '
                  'Valorizzare name, url, logo, contactPoint e sameAs verso i profili '
                  'social e le property ufficiali del brand.',
    },
    'Collection Page': {
        'pillar': 'S', 'id': 'S02', 'weight': 20, 'kind': 'schema',
        'owner': 'Front-end / SEO',
        'applies_to': ['PLP'],
        'goal': 'Le pagine di listing sono dichiarate come CollectionPage/ItemList.',
        'must_have': 'JSON-LD CollectionPage con ItemList degli elementi in listing, '
                     'position e url di ciascun item.',
        'action': 'Implementare il JSON-LD CollectionPage sulla {page_type} '
                  '(score {score}/100). Includere ItemList con position e url '
                  'dei prodotti esposti nel listing.',
    },
    'Breadcrumb': {
        'pillar': 'S', 'id': 'S03', 'weight': 25, 'kind': 'schema',
        'owner': 'Front-end',
        'applies_to': [],
        'goal': 'Il percorso di navigazione è dichiarato e mostrabile nei rich result.',
        'must_have': 'JSON-LD BreadcrumbList con itemListElement ordinati, '
                     'coerenti con il percorso reale e con URL assoluti.',
        'action': 'Implementare il JSON-LD BreadcrumbList (score {score}/100). '
                  'Dichiarare itemListElement ordinati con URL assoluti, coerenti '
                  'con la struttura di navigazione visibile in pagina.',
    },
    'Product Basic': {
        'pillar': 'S', 'id': 'S04', 'weight': 35, 'kind': 'schema',
        'owner': 'Front-end / E-commerce',
        'applies_to': ['PDP'],
        'goal': 'Il prodotto è dichiarato con le proprietà base dello schema Product.',
        'must_have': 'JSON-LD Product con name, image, description, sku, brand.',
        'action': 'Completare le proprietà base dello schema Product '
                  '(score {score}/100). Verificare name, image, description, sku e brand '
                  'e la loro coerenza con il contenuto visibile della PDP.',
    },
    'Product Advanced': {
        'pillar': 'S', 'id': 'S05', 'weight': 30, 'kind': 'schema',
        'owner': 'Front-end / E-commerce',
        'applies_to': ['PDP'],
        'goal': 'Il prodotto espone anche le proprietà avanzate utili ai rich result.',
        'must_have': 'gtin/mpn, color, size, material, additionalProperty e '
                     'isVariantOf/hasVariant quando applicabile.',
        'action': 'Estendere lo schema Product alle proprietà avanzate '
                  '(score {score}/100). Aggiungere gtin/mpn, attributi di variante '
                  'e additionalProperty per formato e INCI.',
    },
    'Rating & Values': {
        'pillar': 'S', 'id': 'S06', 'weight': 20, 'kind': 'schema',
        'owner': 'Front-end / E-commerce',
        'applies_to': ['PDP'],
        'goal': 'Recensioni e valutazioni sono dichiarate in modo idoneo ai rich result.',
        'must_have': 'AggregateRating con ratingValue, reviewCount, bestRating, '
                     'coerenti con le recensioni visibili in pagina.',
        'action': 'Implementare AggregateRating nel JSON-LD (score {score}/100). '
                  'Dichiarare ratingValue, reviewCount e bestRating allineati alle '
                  'recensioni realmente mostrate nella PDP.',
    },
    'Offer': {
        'pillar': 'S', 'id': 'S07', 'weight': 25, 'kind': 'schema',
        'owner': 'Front-end / E-commerce',
        'applies_to': ['PDP'],
        'goal': 'Prezzo e disponibilità sono dichiarati e aggiornati.',
        'must_have': 'Offer con price, priceCurrency, availability, priceValidUntil '
                     'e url, coerenti con il prezzo esposto.',
        'action': 'Implementare il nodo Offer nello schema Product '
                  '(score {score}/100). Valorizzare price, priceCurrency, availability '
                  'e priceValidUntil, sincronizzati con il feed e-commerce.',
    },
    'Article': {
        'pillar': 'S', 'id': 'S08', 'weight': 20, 'kind': 'schema',
        'owner': 'Front-end / Content',
        'applies_to': ['Content Page'],
        'goal': 'I contenuti editoriali sono dichiarati come Article.',
        'must_have': 'JSON-LD Article con headline, image, datePublished, '
                     'dateModified e publisher.',
        'action': 'Implementare il JSON-LD Article sulla {page_type} '
                  '(score {score}/100). Valorizzare headline, image, datePublished, '
                  'dateModified e publisher.',
    },
    'Author': {
        'pillar': 'S', 'id': 'S09', 'weight': 15, 'kind': 'schema',
        'owner': 'Front-end / Content',
        'applies_to': ['Content Page'],
        'goal': 'La paternità del contenuto è dichiarata — segnale E-E-A-T rilevante '
                'anche per la citazione da parte degli LLM.',
        'must_have': 'Nodo author di tipo Person con name, url e, dove disponibile, '
                     'jobTitle e sameAs verso una pagina autore.',
        'action': 'Dichiarare l’autore nello structured data (score {score}/100). '
                  'Aggiungere author di tipo Person con name, url della pagina autore '
                  'e sameAs; creare la pagina autore se assente.',
    },
    'FAQ': {
        'pillar': 'S', 'id': 'S10', 'weight': 15, 'kind': 'schema',
        'owner': 'Front-end / Content',
        'applies_to': [],
        'goal': 'Le domande frequenti sono dichiarate e utilizzabili come risposta '
                'diretta da motori e assistenti generativi.',
        'must_have': 'JSON-LD FAQPage con coppie Question/acceptedAnswer '
                     'realmente visibili in pagina.',
        'action': 'Implementare il JSON-LD FAQPage (score {score}/100). '
                  'Pubblicare in pagina almeno tre coppie domanda/risposta pertinenti '
                  'e dichiararle in mainEntity.',
    },

    # ───────────────────────────── T — TECHNICALS ─────────────────────────────
    'Sitemap Declaration': {
        'pillar': 'T', 'id': 'T01', 'weight': 30, 'kind': 'tech',
        'owner': 'SEO / IT',
        'applies_to': [],
        'goal': 'Ogni URL indicizzabile è dichiarata nella sitemap XML.',
        'must_have': 'URL presente in sitemap, status 200, nessun redirect 301, '
                     'nessuna 404, coerenza con il canonical.',
        'action': 'Inserire l’URL nella sitemap XML (score {score}/100). '
                  'Verificare che restituisca 200 senza redirect, che il canonical '
                  'sia autoreferenziale e risottomettere la sitemap in Search Console.',
    },
    'Page Speed': {
        'pillar': 'T', 'id': 'T02', 'weight': 40, 'kind': 'tech',
        'owner': 'Front-end / IT',
        'applies_to': [],
        'goal': 'La pagina rispetta le soglie Core Web Vitals su mobile.',
        'must_have': 'LCP ≤ 2,5 s, INP ≤ 200 ms, CLS ≤ 0,1 su dati field mobile.',
        'action': 'Intervenire sulle performance di caricamento (score {score}/100). '
                  'Ottimizzare LCP (immagine hero, preload, formati next-gen), '
                  'ridurre il JS bloccante e stabilizzare il layout per il CLS.',
    },
    'Valid Inlinks': {
        'pillar': 'T', 'id': 'T03', 'weight': 30, 'kind': 'tech',
        'owner': 'SEO / IT',
        'applies_to': [],
        'goal': 'I link in ingresso puntano a URL valide, senza catene di redirect.',
        'must_have': 'Inlink verso URL con status 200, senza 301/302 intermedi '
                     'e senza destinazioni 404.',
        'action': 'Bonificare i link interni in ingresso (score {score}/100). '
                  'Aggiornare gli href che passano da redirect o puntano a 404, '
                  'sostituendoli con l’URL finale canonica.',
    },
}


# ── Alias colonne raccomandazione ─────────────────────────────────────────────
#  Nell'export WSX il nome della colonna score e quello della colonna
#  raccomandazione non sempre coincidono: 'Meta Tags - Score' ma
#  'Meta Tags Optimisation - Recommendation'.
REC_COL_ALIAS: Dict[str, str] = {
    'Meta Tags':      'Meta Tags Optimisation',
    'Relevance':      'Relevance & UX',
    'Heading':        'Heading Structure',
    'Grammar':        'Grammar & Spelling',
}

#  Sentinella WSX: la metrica non richiede intervento anche se lo score < 100.
ALREADY_OK = 'already optimized'

#  Metriche il cui intervento modifica testo o contenuto visibile in pagina:
#  per queste il registro può ospitare un suggerimento concreto e pubblicabile.
FRONTEND_METRICS = {
    'Meta Tags',        # title + meta description
    'Relevance',        # struttura e copy del contenuto
    'Heading',          # H1 e gerarchia heading
    'Grammar',          # correzioni ortografiche puntuali
    'Unique Content',   # UVP e paragrafo differenziante
    'FAQ',              # coppie domanda/risposta visibili in pagina
}


# ── Valori già proposti da WSX ───────────────────────────────────────────────
#  Il testo della raccomandazione contiene spesso il valore pronto, scritto in
#  lingua e citato fra virgolette ("Change the meta description to ... e.g.,
#  'Scopri Day Control Protection 48h, ...'"). Estrarlo costa zero: niente
#  crawl, niente LLM. Va usato come punto di partenza, non come verità: WSX
#  non garantisce lunghezza né unicità.
#
#  L'apostrofo italiano (dell', l', un') NON è un delimitatore: l'apice
#  aprente deve essere preceduto da spazio o punteggiatura e il chiudente
#  seguito da spazio, punteggiatura o fine stringa.

_PROP_OPEN  = r"(?<=[\s(:\[,])"
_PROP_CLOSE = r"(?=[\s,.;:)\]]|$)"
_PROP_PATTERNS = [
    re.compile(r'[\u201c]([^\u201c\u201d]{8,300})[\u201d]'),
    re.compile(r'"([^"]{8,300})"'),
    re.compile(_PROP_OPEN + r"'([^']{8,300})'" + _PROP_CLOSE),
    re.compile(_PROP_OPEN + r'\u2018([^\u2018\u2019]{8,300})\u2019' + _PROP_CLOSE),
]
#  Frammenti di istruzione inglese finiti dentro le virgolette
_PROP_JUNK = re.compile(
    r'^(with|to |and |or |s |the |a |an |for |such as|e\.g|i\.e|that|which|'
    r'this|these|it |its )', re.I)


#  Cue che introducono un valore PROPOSTO. Senza uno di questi prima delle
#  virgolette, la stringa citata è quasi sempre il valore ATTUALE portato come
#  esempio del problema ("Replace 'Centro preferenze sulla privacy' with…").
_PROP_CUE = re.compile(
    r"(such as|for example|e\.g\.,?|i\.e\.,?|replace .{0,60}? with|rename .{0,60}? to|"
    r"rewrite .{0,60}? (?:as|to)|change .{0,60}? to|update .{0,60}? to|use|add|"
    r"introduce|consider|propose[d]?|suggest(?:ed)?|like|with a|to a)\s*[:,]?\s*$",
    re.I)

#  Cue negativi: la citazione è il PROBLEMA, non la proposta
_PROP_NEGCUE = re.compile(
    r"(remove|duplicate[d]?|repeat(?:s|ed)?|appears? \d+ times|generic|vague|"
    r"non-content|redundant|relocate|merge|eliminate|avoid|delete)"
    r"[^'\"\u2018\u201c]{0,60}$", re.I)

#  Arredo di sito: mai un valore proposto, sempre il problema da rimuovere
_PROP_FURNITURE = re.compile(
    r'privacy|cookie|consent|preferenz|newsletter|carrello|accedi|login|'
    r'sembra che sei in|elenco dei cookie', re.I)


def extract_wsx_proposals(text: str, cued_only: bool = True) -> List[str]:
    """
    Valori concreti già proposti dalla raccomandazione WSX.

    WSX cita fra virgolette sia il valore ATTUALE (come esempio del problema)
    sia quello PROPOSTO. Si distinguono dal contesto che precede le virgolette:
    "Replace 'X' with a descriptive H2 such as 'Y'" → X è il problema, Y la
    proposta. Con cued_only=True vengono restituiti solo i candidati
    introdotti da un cue di proposta; se non ce n'è nessuno si ricade su
    tutti i candidati, esclusi quelli di arredo di sito.
    """
    t = _clean_text(text)
    if not t or is_already_ok(t):
        return []

    cand = []          # (posizione, testo, è_proposta)
    for pat in _PROP_PATTERNS:
        for m in pat.finditer(t):
            c = m.group(1).strip(' .,;:')
            if len(c) < 8 or _PROP_JUNK.match(c) or _PROP_FURNITURE.search(c):
                continue
            before = t[max(0, m.start() - 70):m.start()]
            if _PROP_NEGCUE.search(before) and not _PROP_CUE.search(before):
                continue          # citazione del problema, non una proposta
            cand.append((m.start(), c, bool(_PROP_CUE.search(before))))

    if not cand:
        return []
    cand.sort(key=lambda x: x[0])

    cued = [c for _, c, ok in cand if ok]
    pool = cued if (cued and cued_only) else [c for _, c, _ in cand]

    seen, out = set(), []
    for c in pool:
        if c.lower() not in seen:
            seen.add(c.lower())
            out.append(c)
    return out


def classify_proposals(metric: str, proposals: List[str]) -> dict:
    """
    Assegna i valori proposti al campo giusto in base a metrica e lunghezza.
    Le soglie sono quelle WSX, con tolleranza: un title fuori range resta un
    title, va solo segnalato.
    """
    out = {'title': '', 'description': '', 'h1': '', 'headings': [], 'altro': []}
    if not proposals:
        return out
    if metric == 'Meta Tags':
        for c in proposals:
            n = clen(c)
            if not out['description'] and n >= 100:
                out['description'] = c
            elif not out['title'] and 20 <= n < 100:
                out['title'] = c
            else:
                out['altro'].append(c)
    elif metric == 'Heading':
        # il primo candidato lungo con separatore di brand è tipicamente un H1
        for c in proposals:
            if not out['h1'] and (clen(c) >= 30 or '|' in c):
                out['h1'] = c
            else:
                out['headings'].append(c)
    else:
        out['altro'] = list(proposals)
    return out


def proposals_digest(metric: str, text: str) -> str:
    """Versione leggibile in cella dei valori proposti da WSX."""
    p = classify_proposals(metric, extract_wsx_proposals(text))
    parts = []
    if p['title']:
        parts.append(f"TITLE: {p['title']}  [{clen(p['title'])} car.]")
    if p['description']:
        parts.append(f"DESCRIPTION: {p['description']}  [{clen(p['description'])} car.]")
    if p['h1']:
        parts.append(f"H1: {p['h1']}")
    if p['headings']:
        parts.append('HEADING: ' + ' | '.join(p['headings'][:10]))
    if p['altro']:
        parts.append('ALTRO: ' + ' | '.join(p['altro'][:6]))
    return '\n'.join(parts)


# ══════════════════════════════════════════════════════════════════════════════
#  CATALOGO → DataFrame
# ══════════════════════════════════════════════════════════════════════════════

def catalog_dataframe() -> pd.DataFrame:
    """Elenco completo delle raccomandazioni WSX classificate per tipo CAST."""
    rows = []
    for metric, meta in METRIC_CATALOG.items():
        p = PILLARS[meta['pillar']]
        rows.append({
            'ID':               meta['id'],
            'CAST':             meta['pillar'],
            'Pilastro':         p['name'],
            'Metrica WSX':      metric,
            'Tipo intervento':  meta['kind'],
            'Front-end':        'Sì' if metric in FRONTEND_METRICS else 'No',
            'Owner':            meta['owner'],
            'Peso':             meta['weight'],
            'Page type attesi': ', '.join(meta['applies_to']) if meta['applies_to'] else 'Tutti',
            'Goal':             meta['goal'],
            'Must Have':        meta['must_have'],
            'Azione tipo':      re.sub(r'\s*\(score \{score\}/100(, gap \{gap\} punti)?\)', '',
                                       meta['action']).replace('{page_type}', 'pagina'),
        })
    df = pd.DataFrame(rows)
    order = {'C': 0, 'A': 1, 'S': 2, 'T': 3}
    df['_o'] = df['CAST'].map(order)
    return df.sort_values(['_o', 'ID']).drop(columns='_o').reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════════════════
#  DETECTION  —  pilastro e brand
# ══════════════════════════════════════════════════════════════════════════════

def detect_pillar(df: pd.DataFrame, filename: str = '') -> Optional[str]:
    """
    Determina il pilastro CAST di un export WSX.
    Priorità: colonna score di pilastro → token nel nome file → prefisso file.
    """
    cols = {str(c).strip().lower() for c in df.columns}
    for code, meta in PILLARS.items():
        if meta['score_col'].lower() in cols:
            return code
    # fallback: nome del pilastro senza suffisso (formato export legacy)
    for code, meta in PILLARS.items():
        if meta['name'].lower() in cols:
            return code

    base = os.path.basename(str(filename or ''))
    for code, meta in PILLARS.items():
        if meta['file_token'].lower() in base.lower():
            return code

    m = re.match(r'^[A-Z]{2,4}_([CAST])_', base)
    if m:
        return PILLAR_PREFIX.get(m.group(1))
    return None


def detect_brand(df: pd.DataFrame, filename: str = '') -> str:
    """Brand dal contenuto del file; in fallback dal prefisso del nome file."""
    if 'Brand' in df.columns:
        vals = df['Brand'].dropna().astype(str)
        if len(vals):
            return vals.mode().iloc[0]
    base = os.path.basename(str(filename or ''))
    m = re.match(r'^([A-Z]{2,4})_', base)
    if m:
        return BRAND_PREFIX.get(m.group(1), m.group(1))
    return '—'


def score_columns(df: pd.DataFrame, pillar: str) -> Dict[str, str]:
    """
    Mappa {nome metrica catalogo → nome reale colonna score} per le sotto-metriche
    del pilastro presenti nel DataFrame. Gestisce sia il formato 2026
    ('Meta Tags - Score') sia quello legacy ('Meta Tags').
    """
    out: Dict[str, str] = {}
    overall = PILLARS[pillar]['score_col']
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for metric, meta in METRIC_CATALOG.items():
        if meta['pillar'] != pillar:
            continue
        for cand in (f'{metric} - Score', metric, f'{metric} Score'):
            real = lookup.get(cand.strip().lower())
            if real and real != overall:
                out[metric] = real
                break
    return out


def text_columns(df: pd.DataFrame, pillar: str) -> Dict[str, dict]:
    """
    Mappa {metrica → {'rec', 'nice', 'jsonld'}} con i nomi reali delle colonne
    testuali dell'export. Chiavi assenti se la colonna non esiste nel file.

    Gli export precedenti al rilascio di settembre 2026 non hanno queste
    colonne: in quel caso il dizionario torna vuoto e l'estrazione ricade
    sulla sola soglia di score.
    """
    out: Dict[str, dict] = {}
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for metric, meta in METRIC_CATALOG.items():
        if meta['pillar'] != pillar:
            continue
        base = REC_COL_ALIAS.get(metric, metric)
        found = {}
        for key, suffixes in (
            ('rec',    [' - Recommendation']),
            ('nice',   [' - Nice to Have Recommendation']),
            ('jsonld', [' - JSON-LD Template']),
        ):
            for name in {base, metric}:
                for suf in suffixes:
                    real = lookup.get(f'{name}{suf}'.strip().lower())
                    if real:
                        found[key] = real
                        break
                if key in found:
                    break
        if found:
            out[metric] = found
    return out


def _clean_text(v) -> str:
    """Normalizza una cella testuale WSX: NaN e 'nan' diventano stringa vuota."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    s = str(v).replace('\u200b', '').strip()
    return '' if s.lower() in ('nan', 'none', '') else s


def is_already_ok(text: str) -> bool:
    """True se WSX dichiara la metrica già a posto per quella URL."""
    return _clean_text(text).lower() == ALREADY_OK


# ══════════════════════════════════════════════════════════════════════════════
#  SEVERITÀ  E  PRIORITÀ
# ══════════════════════════════════════════════════════════════════════════════

# soglia sotto la quale una sotto-metrica genera raccomandazione
DEFAULT_THRESHOLD = 100

SEVERITY_BANDS: List[Tuple[float, float, str, str]] = [
    #  min,   max,  label,                 badge
    (0,     0.001, '❌ Assente',           'critical'),
    (0.001, 50,    '🔴 Insufficiente',     'high'),
    (50,    80,    '🟠 Da migliorare',     'medium'),
    (80,    100,   '🟡 Quasi ottimale',    'low'),
    (100,   1e9,   '✅ Conforme',          'ok'),
]


def severity_of(score: Optional[float]) -> str:
    """Etichetta di severità a partire dallo score 0-100."""
    if score is None or pd.isna(score):
        return '— N/A'
    s = float(score)
    for lo, hi, label, _ in SEVERITY_BANDS:
        if lo <= s < hi:
            return label
    return '✅ Conforme'


def _traffic_weight(seo_impr, sessions) -> float:
    """Peso 0-100 del valore di business dell’URL, da impression e sessioni."""
    def _num(v):
        try:
            return float(v) if v is not None and pd.notna(v) else 0.0
        except (TypeError, ValueError):
            return 0.0
    tot = _num(seo_impr) + _num(sessions) * 3  # le sessioni pesano di più
    if tot <= 0:    return 0.0
    if tot < 10:    return 15.0
    if tot < 100:   return 35.0
    if tot < 1000:  return 60.0
    if tot < 10000: return 80.0
    return 100.0


def priority_of(score: Optional[float], metric: str, page_type: str,
                seo_impr=None, sessions=None, mom=None) -> Tuple[str, int]:
    """
    Priorità di intervento.
      45%  gap rispetto a 100, pesato per l’importanza della metrica
      25%  peso del page type
      30%  valore di traffico dell’URL
      ±    correzione per trend MoM negativo
    Ritorna (label, score 0-100).
    """
    if score is None or pd.isna(score):
        return '— N/A', 0

    gap = max(0.0, 100.0 - float(score)) / 100.0
    mw  = METRIC_CATALOG.get(metric, {}).get('weight', 15) / 40.0   # normalizza su max 40
    pw  = PAGETYPE_WEIGHT.get(str(page_type), PAGETYPE_WEIGHT_DEFAULT) / 40.0
    tw  = _traffic_weight(seo_impr, sessions) / 100.0

    raw = 100.0 * (0.45 * gap * min(mw, 1.0) + 0.25 * min(pw, 1.0) + 0.30 * tw)

    # trend in peggioramento → priorità più alta
    try:
        if mom is not None and pd.notna(mom) and float(mom) < 0:
            raw += min(abs(float(mom)), 10.0)
    except (TypeError, ValueError):
        pass

    raw = max(0.0, min(100.0, raw))
    if   raw >= 55: label = '🔴 Critica'
    elif raw >= 38: label = '🟠 Alta'
    elif raw >= 22: label = '🟡 Media'
    else:           label = '🟢 Bassa'
    return label, int(round(raw))


def action_text(metric: str, score: Optional[float], page_type: str) -> str:
    """Testo dell’azione consigliata per la coppia (metrica, score)."""
    meta = METRIC_CATALOG.get(metric)
    if not meta:
        return f'Verificare la metrica {metric} (score {score}).'
    gap = '' if score is None or pd.isna(score) else str(int(round(100 - float(score))))
    sc  = '—' if score is None or pd.isna(score) else f'{float(score):.0f}'
    return meta['action'].format(score=sc, gap=gap, page_type=page_type or 'pagina')


# ══════════════════════════════════════════════════════════════════════════════
#  ESPLOSIONE  wide → long
# ══════════════════════════════════════════════════════════════════════════════

_META_COLS = [
    'Brand', 'Market', 'Language', 'URL', 'Page Type',
    'Territories', 'Core Territories',
    'SEO Impressions (Google Search Console)',
    'SEA Impressions (Google Ads)',
    'All Sessions (Google Analytics)',
    'Engaged Sessions (Google Analytics)',
    'Engagement Rate (Google Analytics)',
]


def _num(v):
    try:
        return float(v) if v is not None and pd.notna(v) else None
    except (TypeError, ValueError):
        return None


def explode_recommendations(df: pd.DataFrame, pillar: Optional[str] = None,
                            filename: str = '',
                            threshold: float = DEFAULT_THRESHOLD,
                            include_compliant: bool = False,
                            include_na: bool = False,
                            trust_wsx_text: bool = True) -> pd.DataFrame:
    """
    Trasforma un export WSX (1 riga = 1 URL) in formato long:
    1 riga = 1 raccomandazione (URL × sotto-metrica).

    Criterio di inclusione, in ordine di precedenza:

    1. **Colonna `<metrica> - Recommendation` presente** (export ≥ set. 2026):
       fa fede il verdetto di WSX. Testo valorizzato e diverso da
       "Already optimized" → raccomandazione. Vale anche con score 100,
       e non vale con score basso ma testo "Already optimized".
    2. **Colonna assente** (export legacy): ricade sulla soglia di score.

    trust_wsx_text=False forza il criterio 2 anche quando il testo c'è.
    """
    pillar = pillar or detect_pillar(df, filename)
    if not pillar:
        return pd.DataFrame()

    brand_fb = detect_brand(df, filename)
    pmeta    = PILLARS[pillar]
    smap     = score_columns(df, pillar)
    tmap     = text_columns(df, pillar) if trust_wsx_text else {}
    if not smap and not tmap:
        return pd.DataFrame()

    overall_col = pmeta['score_col'] if pmeta['score_col'] in df.columns else None
    mom_col     = pmeta['mom_col']   if pmeta['mom_col']   in df.columns else None
    seo_col     = 'SEO Impressions (Google Search Console)'
    sess_col    = 'All Sessions (Google Analytics)'

    metrics = sorted(set(smap) | set(tmap),
                     key=lambda m: METRIC_CATALOG[m]['id'])

    rows: List[dict] = []
    for _, r in df.iterrows():
        url       = str(r.get('URL', '') or '').strip()
        page_type = str(r.get('Page Type', '') or '').strip()
        seo_impr  = r.get(seo_col)  if seo_col  in df.columns else None
        sessions  = r.get(sess_col) if sess_col in df.columns else None
        mom       = _num(r.get(mom_col)) if mom_col else None
        overall   = _num(r.get(overall_col)) if overall_col else None

        for metric in metrics:
            meta  = METRIC_CATALOG[metric]
            score = _num(r.get(smap[metric])) if metric in smap else None
            cols  = tmap.get(metric, {})

            wsx_rec  = _clean_text(r.get(cols['rec']))    if 'rec'    in cols else ''
            wsx_nice = _clean_text(r.get(cols['nice']))   if 'nice'   in cols else ''
            jsonld   = _clean_text(r.get(cols['jsonld'])) if 'jsonld' in cols else ''

            has_text_col = 'rec' in cols
            already_ok   = is_already_ok(wsx_rec)
            _prop = (classify_proposals(metric, extract_wsx_proposals(wsx_rec))
                     if wsx_rec and not already_ok
                     else {'title': '', 'description': '', 'h1': ''})

            # ── decide se la riga è una raccomandazione ─────────────────────
            if has_text_col:
                fonte = 'WSX Recommendation'
                if already_ok:
                    if not include_compliant:
                        continue
                    sev, prio, pscore = '✅ Conforme', '🟢 Bassa', 0
                    act = f'{metric}: WSX dichiara la metrica già ottimizzata. Monitorare.'
                elif not wsx_rec:
                    if not include_na:
                        continue
                    sev, prio, pscore = '— N/A', '— N/A', 0
                    act = (f'Metrica {metric} non valutata da WSX per il page type '
                           f'{page_type or "n/d"}. Verificare se applicabile.')
                else:
                    sev = severity_of(score) if score is not None else '🟠 Da valutare'
                    prio, pscore = priority_of(score if score is not None else 50,
                                               metric, page_type, seo_impr, sessions, mom)
                    act = action_text(metric, score, page_type)
            else:
                fonte = 'Score < soglia'
                if score is None:
                    if not include_na:
                        continue
                    sev, prio, pscore = '— N/A', '— N/A', 0
                    act = (f'Metrica {metric} non rilevata da WSX per il page type '
                           f'{page_type or "n/d"}. Verificare se applicabile.')
                elif score >= threshold and not include_compliant:
                    continue
                elif score >= threshold:
                    sev, prio, pscore = '✅ Conforme', '🟢 Bassa', 0
                    act = f'{metric} conforme. Mantenere e monitorare nel tempo.'
                else:
                    sev = severity_of(score)
                    prio, pscore = priority_of(score, metric, page_type,
                                               seo_impr, sessions, mom)
                    act = action_text(metric, score, page_type)

            rows.append({
                'CAST':               pillar,
                'Pilastro':           pmeta['name'],
                'ID':                 meta['id'],
                'Brand':              r.get('Brand', brand_fb) or brand_fb,
                'Market':             r.get('Market', ''),
                'Language':           r.get('Language', ''),
                'URL':                url,
                'Page Type':          page_type,
                'Metrica':            metric,
                'Front-end':          'Sì' if metric in FRONTEND_METRICS else 'No',
                'Tipo intervento':    meta['kind'],
                'Owner':              meta['owner'],
                'Score metrica':      score,
                'Gap (100-score)':    None if score is None else round(100 - score, 2),
                'Score pilastro':     overall,
                'Δ MoM pilastro':     mom,
                'Fonte':              fonte,
                'Severità':           sev,
                'Priorità':           prio,
                'Priority Score':     pscore,
                'Raccomandazione WSX': wsx_rec,
                'Valore proposto da WSX': proposals_digest(metric, wsx_rec),
                'Title proposto WSX':     _prop.get('title', ''),
                'Description proposta WSX': _prop.get('description', ''),
                'H1 proposto WSX':        _prop.get('h1', ''),
                'Nice to Have WSX':    wsx_nice,
                'JSON-LD Template':    jsonld,
                'Suggerimento':        '',
                'Azione consigliata': act,
                'Goal':               meta['goal'],
                'Must Have':          meta['must_have'],
                'SEO Impressions':    seo_impr,
                'All Sessions':       sessions,
                'Territories':        r.get('Territories', ''),
                'Core Territories':   r.get('Core Territories', ''),
                'Source file':        os.path.basename(str(filename or '')),
            })

    if not rows:
        return pd.DataFrame()
    out = pd.DataFrame(rows)
    return out.sort_values(['Priority Score', 'Gap (100-score)'],
                           ascending=[False, False]).reset_index(drop=True)


def build_all_recommendations(files: List[Tuple[str, pd.DataFrame]],
                              threshold: float = DEFAULT_THRESHOLD,
                              include_compliant: bool = False,
                              include_na: bool = False,
                              trust_wsx_text: bool = True) -> pd.DataFrame:
    """
    Esplode e concatena più export WSX (brand × pilastro) in un unico
    registro di raccomandazioni.

    files : lista di tuple (filename, DataFrame)
    """
    parts = []
    for fname, df in files:
        if df is None or df.empty:
            continue
        ex = explode_recommendations(df, filename=fname, threshold=threshold,
                                     include_compliant=include_compliant,
                                     include_na=include_na,
                                     trust_wsx_text=trust_wsx_text)
        if not ex.empty:
            parts.append(ex)
    if not parts:
        return pd.DataFrame()
    allr = pd.concat(parts, ignore_index=True)
    order = {'C': 0, 'A': 1, 'S': 2, 'T': 3}
    allr['_o'] = allr['CAST'].map(order)
    allr = (allr.sort_values(['Priority Score', '_o', 'Gap (100-score)'],
                             ascending=[False, True, False])
                .drop(columns='_o').reset_index(drop=True))
    return allr


# ══════════════════════════════════════════════════════════════════════════════
#  AGGREGAZIONI
# ══════════════════════════════════════════════════════════════════════════════

def _drop_empty_suggestion(df: pd.DataFrame) -> pd.DataFrame:
    """
    Toglie la colonna 'Suggerimento' quando nessuna riga è valorizzata.
    Una colonna vuota in un export sembra un guasto: i suggerimenti generati
    arrivano dalla tab «03 · Suggerimenti front-end», non da questa.
    """
    if df is None or df.empty or 'Suggerimento' not in df.columns:
        return df
    if df['Suggerimento'].astype(str).str.strip().replace(
            {'nan': '', 'None': '', '<NA>': ''}).eq('').all():
        return df.drop(columns='Suggerimento')
    return df


def summarize_by_cast(recs: pd.DataFrame) -> pd.DataFrame:
    """Riepilogo raccomandazioni per pilastro CAST e metrica."""
    if recs.empty:
        return pd.DataFrame()
    g = (recs.groupby(['CAST', 'Pilastro', 'ID', 'Metrica'], dropna=False)
             .agg(**{
                 'Raccomandazioni': ('URL', 'count'),
                 'URL distinte':    ('URL', 'nunique'),
                 'Score medio':     ('Score metrica', 'mean'),
                 'Gap medio':       ('Gap (100-score)', 'mean'),
                 'Priority medio':  ('Priority Score', 'mean'),
                 'Critiche':        ('Priorità', lambda s: int((s == '🔴 Critica').sum())),
                 'Alte':            ('Priorità', lambda s: int((s == '🟠 Alta').sum())),
             })
             .reset_index())
    for c in ('Score medio', 'Gap medio', 'Priority medio'):
        g[c] = g[c].round(1)
    order = {'C': 0, 'A': 1, 'S': 2, 'T': 3}
    g['_o'] = g['CAST'].map(order)
    return (g.sort_values(['_o', 'Raccomandazioni'], ascending=[True, False])
             .drop(columns='_o').reset_index(drop=True))


def pivot_metric_by_pagetype(recs: pd.DataFrame) -> pd.DataFrame:
    """Matrice metrica × page type con il conteggio delle raccomandazioni."""
    if recs.empty:
        return pd.DataFrame()
    p = pd.pivot_table(recs, index=['CAST', 'Metrica'], columns='Page Type',
                       values='URL', aggfunc='count', fill_value=0)
    p['Totale'] = p.sum(axis=1)
    return p.sort_values('Totale', ascending=False).reset_index()


def to_excel_workbook(sheets: Dict[str, pd.DataFrame]) -> bytes:
    """Serializza più DataFrame in un unico workbook .xlsx."""
    import io
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        for name, d in sheets.items():
            if d is None or (hasattr(d, 'empty') and d.empty):
                continue
            d.to_excel(w, sheet_name=str(name)[:31], index=False)
            ws = w.sheets[str(name)[:31]]
            ws.freeze_panes = 'A2'
            for i, col in enumerate(d.columns, start=1):
                try:
                    width = min(max(12, int(d[col].astype(str).str.len().quantile(0.9)) + 3), 60)
                except Exception:
                    width = 18
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
#  FRONT-END SUGGESTION GENERATOR
# ══════════════════════════════════════════════════════════════════════════════
#  Per le metriche che richiedono una modifica al testo o al contenuto visibile
#  della pagina (FRONTEND_METRICS) il registro non si limita a riportare la
#  raccomandazione WSX: produce il valore già scritto, pronto da pubblicare.
#
#  Input della generazione:
#    · la raccomandazione WSX per quella URL e quella metrica (specifica,
#      in inglese, con riferimenti al contenuto reale della pagina)
#    · il Nice to Have WSX, quando presente
#    · i valori attualmente in pagina, letti col crawl (opzionale ma consigliato)
#
#  Output per metrica:
#    Meta Tags       → title + meta description
#    Heading         → H1 + outline H2/H3
#    Relevance       → paragrafo di apertura + sezioni con bullet
#    Grammar         → tabella di correzioni da → a
#    Unique Content  → UVP + paragrafo differenziante
#    FAQ             → coppie domanda/risposta
#
#  Le metriche non front-end restano con la sola raccomandazione WSX.
# ══════════════════════════════════════════════════════════════════════════════

H1_MIN, H1_MAX = 30, 70

_UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
       '(KHTML, like Gecko) Chrome/124.0 Safari/537.36')


def fetch_page_tags(url: str, timeout: int = 15) -> dict:
    """
    Legge dalla pagina live i contenuti on-page utili alla riscrittura.
    Non solleva eccezioni: in caso di errore valorizza 'http_status', così la
    riga resta nel report invece di sparire.
    """
    out = {'http_status': '', 'title_now': '', 'desc_now': '', 'h1_now': '',
           'h1_count': 0, 'outline_now': '', 'body_excerpt': '',
           'lang_attr': '', 'canonical': ''}
    if not url:
        out['http_status'] = 'URL mancante'
        return out
    try:
        r = requests.get(url, timeout=timeout, headers={'User-Agent': _UA},
                         allow_redirects=True)
        out['http_status'] = str(r.status_code)
        if r.status_code != 200:
            return out
        r.encoding = r.encoding or 'utf-8'
        soup = BeautifulSoup(r.text, 'html.parser')

        if soup.title and soup.title.string:
            out['title_now'] = soup.title.string.strip()

        md = (soup.find('meta', attrs={'name': 'description'})
              or soup.find('meta', attrs={'property': 'og:description'}))
        if md and md.get('content'):
            out['desc_now'] = md['content'].strip()

        h1s = [h.get_text(' ', strip=True) for h in soup.find_all('h1')]
        h1s = [h for h in h1s if h]
        out['h1_count'] = len(h1s)
        out['h1_now'] = ' ⏐ '.join(h1s[:3])

        outline = []
        for tag in soup.find_all(['h1', 'h2', 'h3']):
            txt = tag.get_text(' ', strip=True)
            if txt:
                outline.append(f'{tag.name.upper()}: {txt}')
        out['outline_now'] = '\n'.join(outline[:40])

        for junk in soup(['script', 'style', 'noscript', 'nav', 'footer']):
            junk.decompose()
        body = re.sub(r'\s+', ' ', soup.get_text(' ', strip=True))
        out['body_excerpt'] = body[:3000]

        html_tag = soup.find('html')
        if html_tag and html_tag.get('lang'):
            out['lang_attr'] = html_tag['lang']

        can = soup.find('link', attrs={'rel': 'canonical'})
        if can and can.get('href'):
            out['canonical'] = can['href']

    except requests.Timeout:
        out['http_status'] = 'timeout'
    except Exception as exc:
        out['http_status'] = f'errore: {type(exc).__name__}'
    return out


# ─── Prompt per metrica ───────────────────────────────────────────────────────

_COMMON_RULES = (
    "- Non inventare proprietà, ingredienti, percentuali, prezzi, promozioni o "
    "certificazioni che non compaiano già nei contenuti forniti.\n"
    "- Mantieni invariati nome prodotto, formato, concentrazione e claim "
    "regolamentati.\n"
    "- Nessuna parola in inglese salvo nomi prodotto, brand e termini tecnici "
    "consolidati (SPF, INCI, retinolo)."
)


def _ctx_block(url, brand, market, page_type, score, page) -> str:
    lines = [
        f"URL: {url}",
        f"Brand: {brand or 'n/d'} · Mercato: {market or 'n/d'} · "
        f"Tipo pagina: {page_type or 'n/d'}",
    ]
    if score is not None and pd.notna(score):
        lines.append(f"Score WSX della metrica: {float(score):.0f}/100")
    if page.get('http_status') == '200':
        lines += ['', 'CONTENUTI ATTUALI IN PAGINA:',
                  f"Title: {page.get('title_now') or '(assente)'}",
                  f"Meta description: {page.get('desc_now') or '(assente)'}",
                  f"H1 ({page.get('h1_count', 0)}): {page.get('h1_now') or '(nessuno)'}"]
        if page.get('outline_now'):
            lines += ['Struttura heading:', page['outline_now'][:1500]]
        if page.get('body_excerpt'):
            lines += ['Estratto del testo:', page['body_excerpt'][:1800]]
    else:
        lines += ['', '(pagina non letta: lavora sulla sola raccomandazione WSX '
                  'e sullo slug URL, restando prudente)']
    return '\n'.join(lines)


def _prompt_frontend(metric, url, target_lang, brand, market, page_type,
                     score, wsx_rec, wsx_nice, page) -> str:
    """Costruisce il prompt di riscrittura per una metrica front-end."""
    spec = {
        'Meta Tags': (
            f"- Title: {TITLE_MIN}-{TITLE_MAX} caratteri, keyword principale nelle "
            f"prime 3 parole, separatore | prima del brand.\n"
            f"- Description: {DESC_MIN}-{DESC_MAX} caratteri, con call to action "
            f"finale.\n"
            "- Il testo deve essere unico rispetto alle altre pagine dello stesso tipo.",
            '{"title": "...", "description": "..."}',
        ),
        'Heading': (
            f"- H1: uno solo, {H1_MIN}-{H1_MAX} caratteri, con la keyword "
            "principale, senza call to action, non identico al meta title.\n"
            "- Outline: sequenza logica di H2/H3 che sostituisce quella attuale, "
            "eliminando heading di servizio (privacy, cookie) e duplicati.\n"
            "- Massimo 12 voci di outline.",
            '{"h1": "...", "outline": [{"level": "H2", "text": "..."}]}',
        ),
        'Relevance': (
            "- Paragrafo di apertura: risponde all'intento primario nelle prime "
            "due frasi, con dati concreti se disponibili nei contenuti forniti.\n"
            "- Sezioni: da 2 a 5, ciascuna con titolo e 2-4 bullet.\n"
            "- Riordina e riformula i contenuti già presenti; non aggiungerne di nuovi.",
            '{"intro": "...", "sections": [{"heading": "...", "bullets": ["..."]}]}',
        ),
        'Grammar': (
            "- Elenca solo correzioni puntuali e verificabili nei contenuti forniti.\n"
            "- Ogni voce ha il testo errato esatto e la sua correzione.\n"
            "- Se non trovi errori certi, restituisci una lista vuota.",
            '{"corrections": [{"from": "...", "to": "...", "note": "..."}]}',
        ),
        'Unique Content': (
            "- UVP: una frase che dichiara perché questa pagina esiste e in cosa "
            "differisce dalle pagine simili del sito.\n"
            "- Paragrafo: 60-100 parole di contenuto differenziante, da inserire "
            "sopra la piega.",
            '{"uvp": "...", "paragraph": "..."}',
        ),
        'FAQ': (
            "- Da 3 a 5 coppie domanda/risposta pertinenti alla pagina.\n"
            "- Risposte di 40-60 parole, con la risposta nella prima frase.\n"
            "- Le domande devono essere formulate come le porrebbe un utente.",
            '{"faqs": [{"q": "...", "a": "..."}]}',
        ),
    }
    rules, schema = spec.get(metric, ('', '{"suggerimento": "..."}'))

    lines = [
        f"Sei un consulente SEO senior. Applica a questa pagina la "
        f"raccomandazione WSX sulla metrica «{metric}», scrivendo il contenuto "
        f"definitivo in {target_lang}.",
        '',
        _ctx_block(url, brand, market, page_type, score, page),
        '',
        'RACCOMANDAZIONE WSX DA APPLICARE:',
        (wsx_rec or '(non fornita)')[:1500],
    ]
    if wsx_nice:
        lines += ['', 'NICE TO HAVE WSX:', wsx_nice[:600]]
    _prop = proposals_digest(metric, wsx_rec)
    if _prop:
        lines += ['', 'VALORI GIÀ PROPOSTI DA WSX (punto di partenza, da '
                  'verificare e correggere su lunghezza e unicità):', _prop[:900]]
    lines += [
        '', 'REGOLE:',
        f"- Scrivi interamente in {target_lang}.",
        "- Se i valori proposti da WSX sono validi, riusali correggendo solo "
        "quel che non rispetta i limiti; non riscrivere da zero senza motivo.",
        rules,
        _COMMON_RULES,
        '',
        'Restituisci SOLO un oggetto JSON valido, senza backtick e senza commenti:',
        schema,
    ]
    return '\n'.join(l for l in lines if l)


def _parse_json_reply(raw: str) -> dict:
    """Estrae un oggetto JSON dalla risposta LLM, tollerando fence e preamboli."""
    if not raw:
        return {}
    txt = re.sub(r'^```(?:json)?|```$', '', str(raw).strip(), flags=re.M).strip()
    try:
        return _json.loads(txt)
    except Exception:
        pass
    m = re.search(r'\{.*\}', txt, re.S)
    if m:
        try:
            return _json.loads(m.group(0))
        except Exception:
            pass
    return {}


def _flatten_suggestion(metric: str, data: dict) -> str:
    """Rende leggibile in una cella di foglio il suggerimento strutturato."""
    if not data:
        return ''
    try:
        if metric == 'Meta Tags':
            return (f"TITLE: {data.get('title','')}\n"
                    f"DESCRIPTION: {data.get('description','')}")
        if metric == 'Heading':
            out = [f"H1: {data.get('h1','')}"]
            for h in data.get('outline', []):
                out.append(f"{h.get('level','H2')}: {h.get('text','')}")
            return '\n'.join(out)
        if metric == 'Relevance':
            out = [f"APERTURA: {data.get('intro','')}"]
            for s in data.get('sections', []):
                out.append(f"\n{s.get('heading','')}")
                out += [f"  · {b}" for b in s.get('bullets', [])]
            return '\n'.join(out)
        if metric == 'Grammar':
            return '\n'.join(
                f"«{c.get('from','')}» → «{c.get('to','')}»"
                + (f"  ({c.get('note')})" if c.get('note') else '')
                for c in data.get('corrections', [])) or 'Nessuna correzione certa individuata'
        if metric == 'Unique Content':
            return f"UVP: {data.get('uvp','')}\n\n{data.get('paragraph','')}"
        if metric == 'FAQ':
            return '\n\n'.join(f"D: {f.get('q','')}\nR: {f.get('a','')}"
                               for f in data.get('faqs', []))
    except Exception:
        pass
    return _json.dumps(data, ensure_ascii=False)


def _faq_jsonld(data: dict) -> str:
    """Costruisce il JSON-LD FAQPage dalle coppie generate."""
    faqs = data.get('faqs') or []
    if not faqs:
        return ''
    return _json.dumps({
        '@context': 'https://schema.org',
        '@type': 'FAQPage',
        'mainEntity': [{
            '@type': 'Question', 'name': f.get('q', ''),
            'acceptedAnswer': {'@type': 'Answer', 'text': f.get('a', '')},
        } for f in faqs],
    }, ensure_ascii=False, indent=2)


def _frontend_row_worker(args: dict) -> tuple:
    """
    Elabora UNA url: crawl (una volta sola) + generazione per ciascuna
    metrica front-end con raccomandazione aperta. Thread-safe.
    """
    url  = args['url']
    page = (fetch_page_tags(url, timeout=args.get('timeout', 15))
            if args['do_crawl'] else {'http_status': 'non crawlata'})

    results, notes = {}, []

    for metric, item in args['items'].items():
        if not args['use_llm']:
            continue
        try:
            reply = _call_llm(
                _prompt_frontend(
                    metric=metric, url=url, target_lang=args['target_lang'],
                    brand=args['brand'], market=args['market'],
                    page_type=args['page_type'], score=item.get('score'),
                    wsx_rec=item.get('rec', ''), wsx_nice=item.get('nice', ''),
                    page=page),
                provider=args['provider'], api_key=args['api_key'],
                model=args['model'], max_tokens=1500)
            data = _parse_json_reply(reply)
            if not data:
                notes.append(f'{metric}: risposta LLM non parsabile')
            results[metric] = data
        except Exception as exc:
            notes.append(f'{metric}: {type(exc).__name__}')
            results[metric] = {}

    if args['do_crawl'] and page.get('http_status') == '200':
        if page.get('h1_count', 0) > 1:
            notes.append(f"{page['h1_count']} H1 in pagina — rimuovere i duplicati")
        if page.get('h1_count', 0) == 0:
            notes.append('nessun H1 in pagina')
        if not page.get('desc_now'):
            notes.append('meta description assente')

    meta_d = results.get('Meta Tags', {}) or {}
    head_d = results.get('Heading', {}) or {}
    faq_d  = results.get('FAQ', {}) or {}

    title_new = str(meta_d.get('title', '') or '').strip()
    desc_new  = str(meta_d.get('description', '') or '').strip()
    h1_new    = str(head_d.get('h1', '') or '').strip().strip('"\u201c\u201d')

    # Fallback sui valori già proposti da WSX: copre il caso LLM spento e
    # quello di una generazione fallita. La colonna 'Origine' dice sempre da
    # dove arriva ciascun valore, perché la qualità non è la stessa.
    _origine = []
    for _m, _key, _cur in (('Meta Tags', 'title', title_new),
                           ('Meta Tags', 'description', desc_new),
                           ('Heading', 'h1', h1_new)):
        if _cur or _m not in args['items']:
            continue
        _p = classify_proposals(_m, extract_wsx_proposals(
            args['items'][_m].get('rec', '')))
        if _p.get(_key):
            if _key == 'title':        title_new = _p['title']
            elif _key == 'description': desc_new = _p['description']
            else:                       h1_new   = _p['h1']
            _origine.append(f'{_key} da WSX')
    if _origine:
        notes.append('valori ripresi da WSX: ' + ', '.join(
            o.split(' da ')[0] for o in _origine))

    # heading proposti da WSX, utili anche quando l'outline LLM manca
    _head_props = classify_proposals(
        'Heading', extract_wsx_proposals(args['items'].get('Heading', {}).get('rec', '')))

    row = {
        'Brand':      args['brand'],
        'Market':     args['market'],
        'Language':   args['language'],
        'URL':        url,
        'Page Type':  args['page_type'],
        'Metriche front-end': ' · '.join(sorted(args['items'])),
        'Priority Score':     args['priority'],
        'HTTP':               page.get('http_status', ''),
        # Meta tags
        'Title attuale':      page.get('title_now', ''),
        'Len title now':      clen(page.get('title_now', '')),
        'Title suggerito':    title_new,
        'Len title new':      clen(title_new),
        'Title status':       len_status(clen(title_new), TITLE_MIN, TITLE_MAX),
        'Description attuale':   page.get('desc_now', ''),
        'Len desc now':          clen(page.get('desc_now', '')),
        'Description suggerita': desc_new,
        'Len desc new':          clen(desc_new),
        'Desc status':           len_status(clen(desc_new), DESC_MIN, DESC_MAX),
        # Heading
        'H1 attuale':     page.get('h1_now', ''),
        'N. H1':          page.get('h1_count', 0),
        'H1 suggerito':   h1_new,
        'Len h1 new':     clen(h1_new),
        'H1 status':      len_status(clen(h1_new), H1_MIN, H1_MAX),
        'Outline suggerito': _flatten_suggestion('Heading', head_d).split('\n', 1)[-1]
                             if head_d.get('outline') else '',
        'Heading proposti da WSX': ' | '.join(_head_props.get('headings', [])[:12]),
        # Altre metriche front-end
        'Relevance — copy suggerito':   _flatten_suggestion('Relevance', results.get('Relevance', {})),
        'Grammar — correzioni':         _flatten_suggestion('Grammar', results.get('Grammar', {})),
        'Unique Content — suggerito':   _flatten_suggestion('Unique Content', results.get('Unique Content', {})),
        'FAQ — coppie suggerite':       _flatten_suggestion('FAQ', faq_d),
        'FAQ — JSON-LD pronto':         _faq_jsonld(faq_d),
        # Contesto
        'lang attr':  page.get('lang_attr', ''),
        'Canonical':  page.get('canonical', ''),
        'Note':       ' ; '.join(notes),
        '_raw':       {m: d for m, d in results.items()},
    }
    return args['idx'], row


def build_frontend_suggestions(recs: pd.DataFrame, limit: int = 30,
                               metrics: Optional[List[str]] = None,
                               do_crawl: bool = True, use_llm: bool = True,
                               provider: str = 'anthropic', api_key: str = None,
                               model: str = None, n_workers: int = 4,
                               timeout: int = 15,
                               target_lang_override: str = '') -> pd.DataFrame:
    """
    Genera i contenuti pronti da pubblicare per le URL che hanno almeno una
    raccomandazione aperta su una metrica front-end.

    Una URL viene crawlata UNA volta sola, poi si genera solo per le metriche
    che per quella URL hanno effettivamente una raccomandazione WSX aperta.
    """
    if recs is None or recs.empty:
        return pd.DataFrame()

    wanted = set(metrics) if metrics else set(FRONTEND_METRICS)
    sub = recs[recs['Metrica'].isin(wanted)]
    if 'Front-end' in sub.columns:
        sub = sub[sub['Front-end'] == 'Sì']
    if sub.empty:
        return pd.DataFrame()

    grouped = []
    for (url, brand), g in sub.groupby(['URL', 'Brand'], dropna=False):
        r0 = g.iloc[0]
        items = {}
        for _, rr in g.iterrows():
            items[rr['Metrica']] = {
                'score': rr.get('Score metrica'),
                'rec':   rr.get('Raccomandazione WSX', '') or '',
                'nice':  rr.get('Nice to Have WSX', '') or '',
            }
        grouped.append({
            'url': url, 'brand': brand, 'items': items,
            'market':    r0.get('Market', ''),
            'language':  r0.get('Language', ''),
            'page_type': r0.get('Page Type', ''),
            'priority':  int(g['Priority Score'].max()),
        })

    grouped.sort(key=lambda d: d['priority'], reverse=True)
    if limit:
        grouped = grouped[:limit]

    jobs = []
    for i, d in enumerate(grouped):
        tgt = target_lang_override or page_language(d['language']) or 'italiano'
        jobs.append({'idx': i, **d, 'target_lang': tgt,
                     'do_crawl': do_crawl, 'use_llm': use_llm,
                     'provider': provider, 'api_key': api_key,
                     'model': model, 'timeout': timeout})

    rows = _run_parallel(_frontend_row_worker, jobs, n_workers,
                         progress_label='Generazione suggerimenti')
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    out = out.drop(columns=['_raw'], errors='ignore')
    return out.sort_values('Priority Score', ascending=False).reset_index(drop=True)


def merge_suggestions_into_register(recs: pd.DataFrame,
                                    sugg: pd.DataFrame) -> pd.DataFrame:
    """Riporta i suggerimenti generati nella colonna 'Suggerimento' del registro."""
    if recs is None or recs.empty or sugg is None or sugg.empty:
        return recs
    col_by_metric = {
        'Meta Tags':      lambda r: (f"TITLE: {r['Title suggerito']}\\n"
                                     f"DESCRIPTION: {r['Description suggerita']}"
                                     if r['Title suggerito'] or r['Description suggerita'] else ''),
        'Heading':        lambda r: (f"H1: {r['H1 suggerito']}\\n{r['Outline suggerito']}"
                                     if r['H1 suggerito'] else ''),
        'Relevance':      lambda r: r['Relevance — copy suggerito'],
        'Grammar':        lambda r: r['Grammar — correzioni'],
        'Unique Content': lambda r: r['Unique Content — suggerito'],
        'FAQ':            lambda r: r['FAQ — coppie suggerite'],
    }
    out = recs.copy()
    idx = {(r['URL'], r['Brand']): r for _, r in sugg.iterrows()}
    for i, row in out.iterrows():
        m = row['Metrica']
        if m not in col_by_metric:
            continue
        s = idx.get((row['URL'], row['Brand']))
        if s is None:
            continue
        try:
            out.at[i, 'Suggerimento'] = col_by_metric[m](s) or ''
        except Exception:
            pass
    return out


# ══════════════════════════════════════════════════════════════════════════════
#  CLUSTERING RACCOMANDAZIONI
# ══════════════════════════════════════════════════════════════════════════════
#  Le raccomandazioni WSX sono per-URL ma non per questo tutte diverse: molte
#  ripetono la stessa azione su centinaia di pagine, spesso perché il problema
#  sta in un componente condiviso (banner cookie, widget, blocco CTA) e non
#  nel contenuto della singola pagina.
#
#  Il clustering serve a decidere COME lavorare prima di lavorare:
#    · azioni di TEMPLATE  → un ticket al front-end risolve N URL insieme
#    · azioni di PAGINA    → richiedono intervento editoriale una per una
#
#  Metodo: ogni raccomandazione viene spezzata nelle sue azioni atomiche, i
#  riferimenti specifici (testi fra virgolette, numeri) vengono sostituiti da
#  segnaposto per far emergere il pattern d'azione, poi TF-IDF + clustering
#  agglomerativo su distanza coseno.
# ══════════════════════════════════════════════════════════════════════════════

#  Abbreviazioni da proteggere: contengono '.,' e romperebbero lo split
_ABBR_GUARD = [('e.g.,', '\x01EG\x01'), ('i.e.,', '\x01IE\x01'),
               ('etc.,', '\x01ET\x01'), ('vs.,', '\x01VS\x01')]

_QUOTED = re.compile(r"['\u2018\u2019\"\u201c\u201d]([^'\u2018\u2019\"\u201c\u201d]{2,80})"
                     r"['\u2018\u2019\"\u201c\u201d]")

#  Entità che indicano arredo di sito condiviso anziché contenuto di pagina
_TEMPLATE_HINT = re.compile(
    r'privacy|cookie|consent|preferenz|skinconsult|acquista online|newsletter|'
    r'accedi|login|carrello|breadcrumb|menu|footer|header|banner|pop-?up',
    re.I)


def split_recommendation_items(text: str) -> List[str]:
    """Spezza una raccomandazione WSX nelle singole azioni atomiche."""
    t = _clean_text(text)
    if not t or is_already_ok(t):
        return []
    for a, ph in _ABBR_GUARD:
        t = t.replace(a, ph)
    parts = re.split(r'\.,\s+|\.\s+(?=[A-Z])|\n+', t)
    out = []
    for p in parts:
        for a, ph in _ABBR_GUARD:
            p = p.replace(ph, a)
        p = p.strip(' .;·-')
        if len(p) > 15:
            out.append(p)
    return out


def item_entities(item: str) -> List[str]:
    """Testi fra virgolette citati dalla raccomandazione (heading, CTA, label)."""
    return [m.strip() for m in _QUOTED.findall(str(item)) if len(m.strip()) > 2]


def normalize_item(item: str) -> str:
    """Rimuove i riferimenti specifici per isolare il pattern d'azione."""
    s = _QUOTED.sub(' <X> ', str(item).lower())
    s = re.sub(r'\d+', ' <N> ', s)
    s = re.sub(r'[^a-z<>\s]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def is_template_action(item: str) -> bool:
    """True se l'azione riguarda un componente condiviso e non il contenuto."""
    return bool(_TEMPLATE_HINT.search(str(item)))


def explode_recommendation_items(recs: pd.DataFrame,
                                 metrics: Optional[List[str]] = None
                                 ) -> pd.DataFrame:
    """Da 1 riga = 1 raccomandazione a 1 riga = 1 azione atomica."""
    if recs is None or recs.empty or 'Raccomandazione WSX' not in recs.columns:
        return pd.DataFrame()
    sel = recs[recs['Metrica'].isin(metrics)] if metrics else recs
    rows = []
    for _, r in sel.iterrows():
        for it in split_recommendation_items(r.get('Raccomandazione WSX', '')):
            ents = item_entities(it)
            rows.append({
                'CAST':     r['CAST'],
                'Brand':    r['Brand'],
                'Metrica':  r['Metrica'],
                'URL':      r['URL'],
                'Page Type': r.get('Page Type', ''),
                'Priority Score': r.get('Priority Score', 0),
                'Azione atomica': it,
                'Livello':  'Template' if is_template_action(it) else 'Pagina',
                'Entità citate': ' | '.join(ents[:4]),
            })
    return pd.DataFrame(rows)


def cluster_recommendation_items(items: pd.DataFrame,
                                 n_clusters: int = 12) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Raggruppa le azioni atomiche per pattern.
    Ritorna (items con colonna Cluster, tabella riassuntiva dei cluster).
    """
    if items is None or items.empty:
        return pd.DataFrame(), pd.DataFrame()
    try:
        import numpy as np
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.cluster import AgglomerativeClustering
    except ImportError:
        return items.assign(Cluster=-1), pd.DataFrame()

    work = items.copy().reset_index(drop=True)
    work['_norm'] = work['Azione atomica'].map(normalize_item)
    work = work[work['_norm'].str.len() > 5].reset_index(drop=True)
    if len(work) < 4:
        return work.assign(Cluster=0), pd.DataFrame()

    # min_df adattivo: su corpora piccoli 2 azzererebbe il vocabolario
    for min_df in (2, 1):
        try:
            vec = TfidfVectorizer(ngram_range=(1, 3), min_df=min_df, max_df=0.85,
                                  stop_words='english', sublinear_tf=True)
            X = vec.fit_transform(work['_norm'])
            if X.shape[1] >= 5:
                break
        except ValueError:
            X = None
    if X is None or X.shape[1] < 2:
        return work.assign(Cluster=0), pd.DataFrame()

    Xd = X.toarray()
    keep = Xd.sum(axis=1) > 0          # righe senza termini utili → fuori
    work = work[keep].reset_index(drop=True)
    Xd = Xd[keep]
    n = max(2, min(int(n_clusters), len(work) - 1))

    model = AgglomerativeClustering(n_clusters=n, metric='cosine', linkage='average')
    work['Cluster'] = model.fit_predict(Xd)

    terms = np.array(vec.get_feature_names_out())
    tot_url = work['URL'].nunique() or 1

    summary = []
    for c, g in work.groupby('Cluster'):
        idxs = g.index.to_numpy()
        centroid = Xd[idxs].mean(axis=0)
        top_terms = terms[centroid.argsort()[::-1][:6]]
        rep = g.iloc[int(np.argmax(Xd[idxs] @ centroid))]['Azione atomica']
        ents = pd.Series([e for row in g['Entità citate']
                          for e in str(row).split(' | ') if e]).value_counts()
        n_tpl = int((g['Livello'] == 'Template').sum())
        summary.append({
            'Cluster':   f'C{c}',
            'Azioni':    len(g),
            'URL':       g['URL'].nunique(),
            '% URL':     round(g['URL'].nunique() / tot_url * 100, 1),
            'Livello':   'Template' if n_tpl > len(g) / 2 else 'Pagina',
            '% template': round(n_tpl / len(g) * 100),
            'Metriche':  ' · '.join(sorted(g['Metrica'].unique())),
            'Pattern':   ' · '.join(top_terms),
            'Esempio':   rep[:300],
            'Entità ricorrenti': ' | '.join(
                f'{k} ({v})' for k, v in ents.head(5).items()),
        })

    sm = (pd.DataFrame(summary)
            .sort_values('Azioni', ascending=False)
            .reset_index(drop=True))
    return work.drop(columns='_norm'), sm


def cluster_effort_summary(items: pd.DataFrame) -> pd.DataFrame:
    """
    Per ogni metrica quantifica quanto lavoro si risolve col template e quanto
    resta davvero da fare pagina per pagina.
    """
    if items is None or items.empty:
        return pd.DataFrame()
    out = []
    for (metric, brand), g in items.groupby(['Metrica', 'Brand'], dropna=False):
        per_url = g.groupby('URL')['Livello'].agg(
            tpl=lambda s: int((s == 'Template').sum()), tot='count')
        only_tpl = int((per_url['tpl'] == per_url['tot']).sum())
        some_tpl = int((per_url['tpl'] > 0).sum())
        n_url = len(per_url)
        out.append({
            'Metrica':   metric,
            'Brand':     brand,
            'URL':       n_url,
            'Azioni':    len(g),
            'Azioni template':  int((g['Livello'] == 'Template').sum()),
            '% azioni template': round((g['Livello'] == 'Template').mean() * 100),
            'URL toccate dal template': some_tpl,
            'URL risolte dal solo template': only_tpl,
            'URL da lavorare a mano': n_url - only_tpl,
        })
    return (pd.DataFrame(out)
              .sort_values('Azioni', ascending=False)
              .reset_index(drop=True))


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
_tab_c, _tab_a, _tab_s, _tab_t, _tab_all = st.tabs([
    '📝 C — Content',
    '🔗 A — Authority',
    '🏗️ S — Structured',
    '⚙️ T — Technical',
    '📋 ALL — Recommendations',
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


# ══════════════════════════════════════════════════════════════════════════════
#  ALL — RECOMMENDATIONS  ·  registro completo cross-CAST
# ══════════════════════════════════════════════════════════════════════════════
#  Estrae TUTTE le raccomandazioni presenti negli export WSX, non solo la
#  "Top Recommendation": ogni sotto-metrica sotto soglia diventa una riga.
#  Accetta N file di N brand contemporaneamente; pilastro e brand sono
#  riconosciuti automaticamente da colonne e nome file.
# ══════════════════════════════════════════════════════════════════════════════

with _tab_all:

    _r_setup, _r_reg, _r_tag, _r_clu, _r_cat = st.tabs([
        '01 · Caricamento & Registro',
        '02 · Analisi per CAST',
        '03 · Suggerimenti front-end',
        '04 · Cluster raccomandazioni',
        '05 · Catalogo raccomandazioni',
    ])

    # ──────────────────────────────────────────────────────────────────────────
    #  01 · CARICAMENTO & REGISTRO
    # ──────────────────────────────────────────────────────────────────────────
    with _r_setup:
        with st.expander('ℹ️ File accettati', expanded=False):
            st.markdown("""
Carica **tutti** gli export WSX *URL Performance / recommendations* che vuoi
consolidare, anche di brand diversi e in un'unica volta.

| Pilastro | Token nel nome file | Colonna riconosciuta |
|---|---|---|
| **C** — Context | `_C_` · `Context` | `Context - Score` |
| **A** — Authority | `_A_` · `Authority` | `Authority - Score` |
| **S** — Structure | `_S_` · `Structure` | `Structure - Score` |
| **T** — Technicals | `_T_` · `Technicals` | `Technicals - Score` |

Il brand viene letto dalla colonna `Brand`; in mancanza, dal prefisso del nome
file (`BIO_` → Biotherm, `VIC_` → Vichy, …).

**Cosa conta come raccomandazione:** ogni sotto-metrica del pilastro con score
inferiore alla soglia. Le metriche con score assente (`NaN`) non sono
applicabili al page type e vengono escluse, salvo diversa impostazione.
""")

        st.markdown('<div class="section-label">Input</div>', unsafe_allow_html=True)

        _rec_uploads = st.file_uploader(
            'Export WSX (.xlsx) — selezione multipla',
            type='xlsx', accept_multiple_files=True, key='rec_up',
            help='Puoi caricare insieme i 4 file CAST di più brand.',
        )

        if _rec_uploads:
            _loaded = []
            for _f in _rec_uploads:
                try:
                    _d = pd.read_excel(_f)
                    _loaded.append((_f.name, _d))
                except Exception as _e:
                    st.error(f'Errore nella lettura di {_f.name}: {_e}')
            st.session_state['rec_files'] = _loaded

        _files = st.session_state.get('rec_files') or []

        if _files:
            # ── Riconoscimento file ───────────────────────────────────────────
            st.markdown('<div class="section-label">File riconosciuti</div>',
                        unsafe_allow_html=True)
            _rows = []
            for _n, _d in _files:
                _p = detect_pillar(_d, _n)
                _rows.append({
                    'File':     _n,
                    'Brand':    detect_brand(_d, _n),
                    'CAST':     _p or '⚠️ non riconosciuto',
                    'Pilastro': PILLARS[_p]['name'] if _p else '—',
                    'URL':      len(_d),
                    'Colonne':  len(_d.columns),
                })
            _recog = pd.DataFrame(_rows)
            st.dataframe(_recog, use_container_width=True, hide_index=True)

            if (_recog['CAST'] == '⚠️ non riconosciuto').any():
                st.warning(
                    'Alcuni file non sono stati associati a un pilastro CAST: '
                    'verifica che contengano la colonna di score del pilastro '
                    'o che il nome file rispetti la convenzione `BRAND_X_...`.'
                )

            st.markdown('<br>', unsafe_allow_html=True)

            # ── Parametri di estrazione ───────────────────────────────────────
            st.markdown('<div class="section-label">Parametri</div>',
                        unsafe_allow_html=True)
            _p1, _p2, _p3 = st.columns([1.4, 1, 1])
            with _p1:
                _thr = st.slider(
                    'Soglia score', min_value=50, max_value=100, value=100, step=5,
                    key='rec_thr',
                    help='Una sotto-metrica genera raccomandazione se il suo score '
                         'è inferiore a questo valore. 100 = tutto ciò che non è '
                         'pienamente conforme.',
                )
            with _p2:
                _inc_ok = st.toggle('Includi conformi', value=False, key='rec_ok',
                                    help='Aggiunge anche le metriche già a norma, '
                                         'per un audit di copertura completo.')
            with _p3:
                _inc_na = st.toggle('Includi N/A', value=False, key='rec_na',
                                    help='Aggiunge le metriche non applicabili al '
                                         'page type (score assente).')
                _trust = st.toggle(
                    'Usa il verdetto WSX', value=True, key='rec_trust',
                    help='Attivo: fa fede la colonna <metrica> - Recommendation, '
                         'ignorando le righe marcate "Already optimized". '
                         'Disattivo: usa solo la soglia di score (comportamento '
                         'necessario sugli export privi di quelle colonne).')

            if st.button('▶ Estrai tutte le raccomandazioni', key='rec_run',
                         type='primary'):
                with st.spinner('Esplosione delle metriche in corso…'):
                    st.session_state['rec_all_df'] = build_all_recommendations(
                        _files, threshold=_thr,
                        include_compliant=_inc_ok, include_na=_inc_na,
                        trust_wsx_text=_trust,
                    )

        _recs = st.session_state.get('rec_all_df')

        if _recs is not None and not _recs.empty:
            st.markdown('<br>', unsafe_allow_html=True)
            st.markdown('<div class="section-label">Riepilogo</div>',
                        unsafe_allow_html=True)

            _pc = _recs['Priorità'].value_counts()
            metrics_row(_recs, {
                'Raccomandazioni': len(_recs),
                'URL coinvolte':   _recs['URL'].nunique(),
                'Front-end':       int((_recs['Front-end'] == 'Sì').sum()),
                'Metriche':        _recs['Metrica'].nunique(),
                'Con valore WSX':  int((_recs['Valore proposto da WSX']
                                        .astype(str).str.len() > 0).sum()),
                'Critiche':        int(_pc.get('🔴 Critica', 0)),
            })
            st.markdown('<br>', unsafe_allow_html=True)

            # ── Distribuzione per CAST ────────────────────────────────────────
            st.markdown('<div class="section-label">Distribuzione per pilastro CAST</div>',
                        unsafe_allow_html=True)
            _cc = st.columns(4)
            for _i, _code in enumerate(['C', 'A', 'S', 'T']):
                _sub = _recs[_recs['CAST'] == _code]
                _pm  = PILLARS[_code]
                with _cc[_i]:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="val">{len(_sub)}</div>
                        <div class="lbl">{_pm['icon']} {_pm['label']}</div>
                    </div>""", unsafe_allow_html=True)
            st.markdown('<br>', unsafe_allow_html=True)

            # ── Filtri ────────────────────────────────────────────────────────
            st.markdown('<div class="section-label">Filtri</div>',
                        unsafe_allow_html=True)
            _f1, _f2, _f3, _f4 = st.columns(4)
            with _f1:
                _fb = st.multiselect('Brand', sorted(_recs['Brand'].dropna().unique()),
                                     key='rec_fb')
            with _f2:
                _fc = st.multiselect('CAST', ['C', 'A', 'S', 'T'], key='rec_fc')
            with _f3:
                _fp = st.multiselect('Page Type',
                                     sorted(_recs['Page Type'].dropna().unique()),
                                     key='rec_fp')
            with _f4:
                _fpr = st.multiselect('Priorità',
                                      ['🔴 Critica', '🟠 Alta', '🟡 Media', '🟢 Bassa'],
                                      key='rec_fpr')

            _f5, _f6 = st.columns([2, 1])
            with _f5:
                _fm = st.multiselect('Metrica', sorted(_recs['Metrica'].unique()),
                                     key='rec_fm')
            with _f6:
                _fmin = st.number_input('Priority Score minimo', 0, 100, 0, 5,
                                        key='rec_fmin')
                _ffe = st.checkbox('Solo metriche front-end', key='rec_ffe')

            _flt = _recs.copy()
            if _fb:  _flt = _flt[_flt['Brand'].isin(_fb)]
            if _fc:  _flt = _flt[_flt['CAST'].isin(_fc)]
            if _fp:  _flt = _flt[_flt['Page Type'].isin(_fp)]
            if _fpr: _flt = _flt[_flt['Priorità'].isin(_fpr)]
            if _fm:  _flt = _flt[_flt['Metrica'].isin(_fm)]
            if _fmin: _flt = _flt[_flt['Priority Score'] >= _fmin]
            if _ffe:  _flt = _flt[_flt['Front-end'] == 'Sì']

            st.caption(f'{len(_flt)} raccomandazioni su {len(_recs)} · '
                       f"{_flt['URL'].nunique()} URL distinte")

            _show = ['CAST', 'ID', 'Brand', 'Page Type', 'Metrica', 'Front-end',
                     'Score metrica', 'Severità', 'Priorità', 'Priority Score',
                     'URL', 'Raccomandazione WSX', 'Valore proposto da WSX',
                     'Azione consigliata']
            st.dataframe(_flt[_show], use_container_width=True, hide_index=True,
                         height=460)

            # ── Export ────────────────────────────────────────────────────────
            st.markdown('<div class="section-label">Export</div>',
                        unsafe_allow_html=True)
            _e1, _e2 = st.columns(2)
            with _e1:
                st.download_button(
                    '⬇ Vista filtrata (.xlsx)',
                    data=to_excel_bytes(_drop_empty_suggestion(_flt)),
                    file_name='wsx_recommendations_filtered.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True)
            with _e2:
                _exp = _drop_empty_suggestion(_recs)
                _book = {
                    'Riepilogo CAST':  summarize_by_cast(_exp),
                    'Tutte le racc.':  _exp,
                    'C - Context':     _exp[_exp['CAST'] == 'C'],
                    'A - Authority':   _exp[_exp['CAST'] == 'A'],
                    'S - Structure':   _exp[_exp['CAST'] == 'S'],
                    'T - Technicals':  _exp[_exp['CAST'] == 'T'],
                    'Front-end':       _exp[_exp['Front-end'] == 'Sì'],
                    'Metrica x PageType': pivot_metric_by_pagetype(_recs),
                    'Catalogo':        catalog_dataframe(),
                }
                st.download_button(
                    '⬇ Registro completo multi-foglio (.xlsx)',
                    data=to_excel_workbook(_book),
                    file_name='wsx_all_recommendations.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    type='primary', use_container_width=True)

        elif _files:
            st.info('Imposta i parametri e avvia l’estrazione per generare il registro.')
        else:
            st.info('Carica uno o più export WSX per costruire il registro '
                    'completo delle raccomandazioni.')

    # ──────────────────────────────────────────────────────────────────────────
    #  02 · ANALISI PER CAST
    # ──────────────────────────────────────────────────────────────────────────
    with _r_reg:
        _recs = st.session_state.get('rec_all_df')

        if _recs is None or _recs.empty:
            st.info('Nessun registro in memoria: esegui prima l’estrazione '
                    'nella tab «01 · Caricamento & Registro».')
        else:
            st.markdown('<div class="section-label">Raccomandazioni per pilastro e metrica</div>',
                        unsafe_allow_html=True)
            _summary = summarize_by_cast(_recs)
            st.dataframe(_summary, use_container_width=True, hide_index=True)

            st.markdown('<br>', unsafe_allow_html=True)

            # ── Brand × CAST ──────────────────────────────────────────────────
            _b1, _b2 = st.columns(2)
            with _b1:
                st.markdown('<div class="section-label">Brand × CAST</div>',
                            unsafe_allow_html=True)
                _bx = pd.crosstab(_recs['Brand'], _recs['CAST'], margins=True,
                                  margins_name='Totale')
                st.dataframe(_bx, use_container_width=True)
            with _b2:
                st.markdown('<div class="section-label">Priorità × CAST</div>',
                            unsafe_allow_html=True)
                _px = pd.crosstab(_recs['Priorità'], _recs['CAST'], margins=True,
                                  margins_name='Totale')
                st.dataframe(_px, use_container_width=True)

            st.markdown('<br>', unsafe_allow_html=True)

            # ── Matrice metrica × page type ───────────────────────────────────
            st.markdown('<div class="section-label">Metrica × Page Type</div>',
                        unsafe_allow_html=True)
            st.dataframe(pivot_metric_by_pagetype(_recs),
                         use_container_width=True, hide_index=True)

            st.markdown('<br>', unsafe_allow_html=True)

            # ── Top URL per carico di intervento ──────────────────────────────
            st.markdown('<div class="section-label">URL con il maggior carico di intervento</div>',
                        unsafe_allow_html=True)
            _top = (_recs.groupby(['Brand', 'URL', 'Page Type'], dropna=False)
                         .agg(**{
                             'Raccomandazioni': ('Metrica', 'count'),
                             'CAST coinvolti':  ('CAST', lambda s: ' · '.join(sorted(set(s)))),
                             'Priority totale': ('Priority Score', 'sum'),
                             'Gap medio':       ('Gap (100-score)', 'mean'),
                         })
                         .reset_index()
                         .sort_values('Priority totale', ascending=False)
                         .head(50))
            _top['Gap medio'] = _top['Gap medio'].round(1)
            st.dataframe(_top, use_container_width=True, hide_index=True)

            # ── Dettaglio per pilastro ────────────────────────────────────────
            st.markdown('<br>', unsafe_allow_html=True)
            st.markdown('<div class="section-label">Dettaglio per pilastro</div>',
                        unsafe_allow_html=True)
            for _code in ['C', 'A', 'S', 'T']:
                _sub = _recs[_recs['CAST'] == _code]
                if _sub.empty:
                    continue
                _pm = PILLARS[_code]
                with st.expander(f"{_pm['icon']} {_pm['label']} — {len(_sub)} raccomandazioni"):
                    st.caption(_pm['desc'])
                    for _mn in _sub['Metrica'].value_counts().index:
                        _ms = _sub[_sub['Metrica'] == _mn]
                        _meta = METRIC_CATALOG[_mn]
                        st.markdown(
                            f"**{_meta['id']} · {_mn}** — {len(_ms)} URL · "
                            f"score medio {_ms['Score metrica'].mean():.0f}/100 · "
                            f"owner: {_meta['owner']}"
                        )
                        st.markdown(f"*Goal:* {_meta['goal']}")
                        st.markdown(f"*Must Have:* {_meta['must_have']}")
                        st.divider()


    # ──────────────────────────────────────────────────────────────────────────
    #  03 · SUGGERIMENTI FRONT-END
    # ──────────────────────────────────────────────────────────────────────────
    with _r_tag:
        _recs = st.session_state.get('rec_all_df')

        with st.expander('ℹ️ Come funziona', expanded=False):
            st.markdown("""
Per le metriche che richiedono una modifica al **testo o al contenuto visibile**
della pagina, il registro non si ferma alla raccomandazione WSX: produce il
contenuto già scritto nella lingua della pagina, pronto da pubblicare.

| Metrica | Output generato |
|---|---|
| **C01 Meta Tags** | Title + Meta Description |
| **C02 Relevance** | Paragrafo di apertura + sezioni con bullet |
| **C03 Heading** | H1 + outline H2/H3 |
| **C04 Grammar** | Tabella di correzioni «da → a» |
| **C05 Unique Content** | UVP + paragrafo differenziante |
| **S10 FAQ** | Coppie D/R + JSON-LD FAQPage pronto |

Tutte le altre metriche (backlink, sitemap, page speed, schema Product…)
restano con la raccomandazione WSX: l'intervento non è testuale.

La generazione parte dalla **raccomandazione WSX di quella specifica URL**, che
negli export da settembre 2026 è puntuale e cita il contenuto reale della
pagina. Il **crawl** aggiunge i valori attualmente pubblicati: senza, il modello
lavora solo sulla raccomandazione e va riletto tutto a mano.

⚠️ Il crawl richiede che le URL siano raggiungibili dalla macchina su cui gira
l'app. In caso di 403 sistematici (bot protection su CDN) esegui l'app in locale.
""")

        if _recs is None or _recs.empty:
            st.info('Nessun registro in memoria: esegui prima l\'estrazione '
                    'nella tab «01 · Caricamento & Registro».')
        else:
            _cand = _recs[_recs['Metrica'].isin(FRONTEND_METRICS)]
            if 'Front-end' in _cand.columns:
                _cand = _cand[_cand['Front-end'] == 'Sì']

            _has_wsx_text = (_recs['Raccomandazione WSX'].astype(str).str.len() > 0).any()
            if not _has_wsx_text:
                st.warning(
                    'I file caricati non contengono le colonne '
                    '`<metrica> - Recommendation`: sono export precedenti a '
                    'settembre 2026. I suggerimenti verranno generati dal solo '
                    'contenuto della pagina, con qualità inferiore.')

            if _cand.empty:
                st.warning('Nessuna raccomandazione front-end aperta nel registro attuale.')
            else:
                _n_url = _cand['URL'].nunique()
                _by_m = _cand['Metrica'].value_counts()
                metrics_row(_cand, {
                    'URL candidate':  _n_url,
                    'Meta Tags':      int(_by_m.get('Meta Tags', 0)),
                    'Heading':        int(_by_m.get('Heading', 0)),
                    'Relevance':      int(_by_m.get('Relevance', 0)),
                    'FAQ':            int(_by_m.get('FAQ', 0)),
                    'Altre':          int(_by_m.get('Grammar', 0) + _by_m.get('Unique Content', 0)),
                })
                st.markdown('<br>', unsafe_allow_html=True)

                st.markdown('<div class="section-label">Parametri</div>',
                            unsafe_allow_html=True)

                _sel_m = st.multiselect(
                    'Metriche da generare',
                    sorted(_cand['Metrica'].unique()),
                    default=sorted(_cand['Metrica'].unique()),
                    key='fe_metrics',
                    help='Ogni metrica selezionata costa una chiamata LLM per URL.')

                _t1, _t2, _t3, _t4 = st.columns([1.3, 1, 1, 1])
                with _t1:
                    _lim = st.number_input(
                        'URL da elaborare (per priorità)', 1, int(_n_url),
                        min(30, int(_n_url)), 10, key='fe_lim')
                with _t2:
                    _crawl = st.toggle('Crawl pagine', value=True, key='fe_crawl')
                with _t3:
                    _tllm = st.toggle('Genera con LLM', value=True, key='fe_llm')
                with _t4:
                    _tto = st.number_input('Timeout (s)', 5, 60, 15, 5, key='fe_to')

                _l1, _l2 = st.columns([1, 2])
                with _l1:
                    _lang_ovr = st.text_input(
                        'Forza lingua output', value='', key='fe_lang',
                        placeholder='es. italiano — vuoto = lingua della pagina')
                with _l2:
                    _tbrand = st.multiselect(
                        'Limita ai brand', sorted(_cand['Brand'].dropna().unique()),
                        key='fe_brand')

                _cand_f = _cand[_cand['Brand'].isin(_tbrand)] if _tbrand else _cand
                if _sel_m:
                    _cand_f = _cand_f[_cand_f['Metrica'].isin(_sel_m)]

                _n_calls = len(_cand_f[_cand_f['URL'].isin(
                    _cand_f.drop_duplicates('URL')['URL'].head(int(_lim)))])
                st.caption(f'Stima: ~{_n_calls} chiamate LLM '
                           f'({int(_lim)} URL × metriche aperte su ciascuna)')

                if _tllm and not _API_KEY and not os.getenv(
                        'ANTHROPIC_API_KEY' if _PROVIDER == 'anthropic' else 'OPENAI_API_KEY'):
                    st.warning('Nessuna API key rilevata per il provider selezionato: '
                               'inseriscila nella sidebar o disattiva «Genera con LLM».')

                if st.button('▶ Genera suggerimenti', key='fe_run', type='primary'):
                    with st.spinner('Crawl e generazione in corso…'):
                        st.session_state['rec_tags_df'] = build_frontend_suggestions(
                            _cand_f, limit=int(_lim), metrics=_sel_m or None,
                            do_crawl=_crawl, use_llm=_tllm,
                            provider=_PROVIDER, api_key=_API_KEY, model=_MODEL,
                            n_workers=_N_WORKERS, timeout=int(_tto),
                            target_lang_override=_lang_ovr.strip())

            _tags = st.session_state.get('rec_tags_df')

            if _tags is not None and not _tags.empty:
                st.markdown('<br>', unsafe_allow_html=True)
                st.markdown('<div class="section-label">Esito</div>',
                            unsafe_allow_html=True)

                def _filled(col):
                    return int((_tags[col].astype(str).str.len() > 0).sum()) \
                        if col in _tags.columns else 0

                _ok_http = int((_tags['HTTP'] == '200').sum())
                metrics_row(_tags, {
                    'URL elaborate':  len(_tags),
                    'Crawl OK':       _ok_http,
                    'Title':          _filled('Title suggerito'),
                    'H1':             _filled('H1 suggerito'),
                    'Copy Relevance': _filled('Relevance — copy suggerito'),
                    'FAQ':            _filled('FAQ — coppie suggerite'),
                })
                st.markdown('<br>', unsafe_allow_html=True)

                if _ok_http < len(_tags):
                    _bad = _tags[_tags['HTTP'] != '200']['HTTP'].value_counts().to_dict()
                    st.warning(f'Crawl non riuscito su {len(_tags) - _ok_http} URL: {_bad}. '
                               'I suggerimenti di queste righe nascono dalla sola '
                               'raccomandazione WSX e vanno verificati uno a uno.')

                _vw = st.radio(
                    'Vista',
                    ['Meta tags', 'Heading', 'Relevance', 'Grammar & Unique', 'FAQ', 'Completa'],
                    horizontal=True, key='fe_view')
                _views = {
                    'Meta tags': ['Brand', 'Page Type', 'URL', 'HTTP',
                                  'Title attuale', 'Len title now', 'Title suggerito',
                                  'Len title new', 'Title status',
                                  'Description attuale', 'Description suggerita',
                                  'Len desc new', 'Desc status'],
                    'Heading':   ['Brand', 'Page Type', 'URL', 'N. H1', 'H1 attuale',
                                  'H1 suggerito', 'Len h1 new', 'H1 status',
                                  'Outline suggerito', 'Heading proposti da WSX',
                                  'Note'],
                    'Relevance': ['Brand', 'Page Type', 'URL', 'Relevance — copy suggerito'],
                    'Grammar & Unique': ['Brand', 'URL', 'Grammar — correzioni',
                                         'Unique Content — suggerito'],
                    'FAQ':       ['Brand', 'Page Type', 'URL', 'FAQ — coppie suggerite',
                                  'FAQ — JSON-LD pronto'],
                    'Completa':  list(_tags.columns),
                }
                _cols = [c for c in _views[_vw] if c in _tags.columns]
                st.dataframe(_tags[_cols], use_container_width=True,
                             hide_index=True, height=480)

                if 'Title status' in _tags.columns:
                    _fuori = _tags[
                        (_tags['Title status'].isin(['TOO SHORT', 'TOO LONG'])) |
                        (_tags['Desc status'].isin(['TOO SHORT', 'TOO LONG']))]
                    if len(_fuori):
                        with st.expander(f'⚠️ {len(_fuori)} meta tag fuori dai limiti '
                                         f'di lunghezza — correggere prima di pubblicare'):
                            st.dataframe(
                                _fuori[['URL', 'Title suggerito', 'Len title new',
                                        'Title status', 'Description suggerita',
                                        'Len desc new', 'Desc status']],
                                use_container_width=True, hide_index=True)

                st.markdown('<div class="section-label">Export</div>',
                            unsafe_allow_html=True)
                _x1, _x2, _x3 = st.columns(3)
                with _x1:
                    _up = _tags[[c for c in
                                 ['URL', 'Title suggerito', 'Description suggerita',
                                  'H1 suggerito'] if c in _tags.columns]].copy()
                    _up.columns = ['URL', 'Title', 'Meta Description', 'H1'][:len(_up.columns)]
                    st.download_button(
                        '⬇ Tracciato CMS (.xlsx)', data=to_excel_bytes(_up),
                        file_name='suggerimenti_meta_h1.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        type='primary', use_container_width=True)
                with _x2:
                    st.download_button(
                        '⬇ Report completo (.xlsx)', data=to_excel_bytes(_tags),
                        file_name='suggerimenti_frontend_report.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True)
                with _x3:
                    _merged = merge_suggestions_into_register(_recs, _tags)
                    st.download_button(
                        '⬇ Registro + suggerimenti (.xlsx)',
                        data=to_excel_workbook({
                            'Riepilogo CAST': summarize_by_cast(_merged),
                            'Tutte le racc.': _merged,
                            'Front-end':      _merged[_merged['Front-end'] == 'Sì'],
                            'Suggerimenti':   _tags,
                            'Catalogo':       catalog_dataframe(),
                        }),
                        file_name='wsx_registro_con_suggerimenti.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True)

    # ──────────────────────────────────────────────────────────────────────────
    #  04 · CLUSTER RACCOMANDAZIONI
    # ──────────────────────────────────────────────────────────────────────────
    with _r_clu:
        _recs = st.session_state.get('rec_all_df')

        with st.expander('ℹ️ A cosa serve', expanded=False):
            st.markdown("""
Le raccomandazioni WSX sono scritte per singola URL, ma molte ripetono la
stessa azione su centinaia di pagine — di solito perché il problema sta in un
**componente condiviso** (banner cookie, widget, blocco CTA) e non nel
contenuto della pagina.

Il clustering separa i due casi prima di aprire i ticket:

- **Template** → un intervento al front-end chiude N URL insieme
- **Pagina** → serve lavoro editoriale una per una

Ogni raccomandazione viene spezzata nelle sue **azioni atomiche**; i
riferimenti specifici (testi fra virgolette, numeri) diventano segnaposto così
da far emergere il pattern; poi TF-IDF e clustering agglomerativo su distanza
coseno.

Richiede gli export da settembre 2026: senza la colonna
`<metrica> - Recommendation` non c'è testo da raggruppare.
""")

        if _recs is None or _recs.empty:
            st.info('Nessun registro in memoria: esegui prima l\'estrazione '
                    'nella tab «01 · Caricamento & Registro».')
        elif 'Raccomandazione WSX' not in _recs.columns or \
                not (_recs['Raccomandazione WSX'].astype(str).str.len() > 0).any():
            st.warning('I file caricati non contengono testo di raccomandazione: '
                       'sono export precedenti a settembre 2026. Il clustering '
                       'non è applicabile.')
        else:
            _with_text = _recs[_recs['Raccomandazione WSX'].astype(str).str.len() > 0]

            st.markdown('<div class="section-label">Parametri</div>',
                        unsafe_allow_html=True)
            _q1, _q2, _q3 = st.columns([2, 1, 1])
            with _q1:
                _cm = st.multiselect(
                    'Metriche da raggruppare',
                    sorted(_with_text['Metrica'].unique()),
                    default=[m for m in ['Heading', 'Relevance']
                             if m in set(_with_text['Metrica'])],
                    key='clu_m',
                    help='Ha senso sulle metriche con raccomandazioni per-URL. '
                         'Su quelle con testo di playbook identico per tutti '
                         'produce un cluster solo.')
            with _q2:
                _nc = st.slider('Numero cluster', 3, 25, 12, 1, key='clu_n')
            with _q3:
                _cb = st.multiselect('Brand',
                                     sorted(_with_text['Brand'].dropna().unique()),
                                     key='clu_b')

            if st.button('▶ Calcola cluster', key='clu_run', type='primary'):
                _base = _with_text[_with_text['Brand'].isin(_cb)] if _cb else _with_text
                with st.spinner('Estrazione azioni e clustering…'):
                    _items = explode_recommendation_items(_base, metrics=_cm or None)
                    _it, _sm = cluster_recommendation_items(_items, n_clusters=_nc)
                    st.session_state['rec_clu_items']   = _it
                    st.session_state['rec_clu_summary'] = _sm

            _it = st.session_state.get('rec_clu_items')
            _sm = st.session_state.get('rec_clu_summary')

            if _it is not None and not _it.empty:
                st.markdown('<br>', unsafe_allow_html=True)
                st.markdown('<div class="section-label">Riepilogo</div>',
                            unsafe_allow_html=True)
                _n_tpl = int((_it['Livello'] == 'Template').sum())
                metrics_row(_it, {
                    'Azioni atomiche': len(_it),
                    'URL':             _it['URL'].nunique(),
                    'Cluster':         _it['Cluster'].nunique(),
                    'Azioni template': _n_tpl,
                    '% template':      f'{round(_n_tpl / max(len(_it), 1) * 100)}%',
                    'Azioni di pagina': len(_it) - _n_tpl,
                })
                st.markdown('<br>', unsafe_allow_html=True)

                st.markdown('<div class="section-label">Carico di lavoro reale</div>',
                            unsafe_allow_html=True)
                _eff = cluster_effort_summary(_it)
                st.dataframe(_eff, use_container_width=True, hide_index=True)
                st.caption(
                    '«URL risolte dal solo template»: pagine in cui **tutte** le '
                    'azioni riguardano componenti condivisi — si chiudono senza '
                    'toccare il contenuto. «Da lavorare a mano»: le restanti.')

                st.markdown('<br>', unsafe_allow_html=True)
                st.markdown('<div class="section-label">Cluster individuati</div>',
                            unsafe_allow_html=True)
                if _sm is not None and not _sm.empty:
                    st.dataframe(_sm, use_container_width=True, hide_index=True,
                                 height=420)

                    st.markdown('<br>', unsafe_allow_html=True)
                    _pick = st.selectbox('Ispeziona un cluster',
                                         _sm['Cluster'].tolist(), key='clu_pick')
                    _cid = int(str(_pick).lstrip('C'))
                    _cg = _it[_it['Cluster'] == _cid]
                    _r0 = _sm[_sm['Cluster'] == _pick].iloc[0]
                    st.markdown(f"**{_r0['Livello']}** · {_r0['Azioni']} azioni su "
                                f"{_r0['URL']} URL · pattern: *{_r0['Pattern']}*")
                    if _r0['Entità ricorrenti']:
                        st.markdown(f"Entità citate: {_r0['Entità ricorrenti']}")
                    st.dataframe(
                        _cg[['Brand', 'Page Type', 'Metrica', 'Livello',
                             'Azione atomica', 'Entità citate', 'URL']],
                        use_container_width=True, hide_index=True, height=340)

                st.markdown('<div class="section-label">Export</div>',
                            unsafe_allow_html=True)
                st.download_button(
                    '⬇ Cluster e azioni (.xlsx)',
                    data=to_excel_workbook({
                        'Carico di lavoro': _eff,
                        'Cluster':          _sm if _sm is not None else pd.DataFrame(),
                        'Azioni atomiche':  _it,
                        'Solo template':    _it[_it['Livello'] == 'Template'],
                        'Solo pagina':      _it[_it['Livello'] == 'Pagina'],
                    }),
                    file_name='wsx_cluster_raccomandazioni.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    type='primary')

    # ──────────────────────────────────────────────────────────────────────────
    #  05 · CATALOGO RACCOMANDAZIONI
    # ──────────────────────────────────────────────────────────────────────────
    with _r_cat:
        st.markdown('<div class="section-label">Catalogo delle raccomandazioni WSX per tipo CAST</div>',
                    unsafe_allow_html=True)
        st.caption(
            'Tassonomia di riferimento: tutte le sotto-metriche rilevate dagli '
            'export WSX, con obiettivo, criterio di conformità e azione tipo. '
            'Indipendente dai file caricati.'
        )

        _cat = catalog_dataframe()

        _k = st.columns(4)
        for _i, _code in enumerate(['C', 'A', 'S', 'T']):
            _n = int((_cat['CAST'] == _code).sum())
            _pm = PILLARS[_code]
            with _k[_i]:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="val">{_n}</div>
                    <div class="lbl">{_pm['icon']} {_pm['label']}</div>
                </div>""", unsafe_allow_html=True)
        st.markdown('<br>', unsafe_allow_html=True)

        _cf = st.multiselect('Filtra per pilastro', ['C', 'A', 'S', 'T'],
                             key='rec_catf')
        _cat_v = _cat[_cat['CAST'].isin(_cf)] if _cf else _cat

        st.dataframe(_cat_v, use_container_width=True, hide_index=True, height=560)

        st.download_button(
            '⬇ Scarica catalogo (.xlsx)',
            data=to_excel_bytes(_cat),
            file_name='wsx_recommendation_catalog.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
