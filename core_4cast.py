"""
core.py  —  logica condivisa: parsing, normalizzazione, traduzione, classificazione
"""
import re
import os
import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional

# Carica .env se presente (locale). In produzione (Streamlit Cloud) usa i secrets.
try:
    from dotenv import load_dotenv
    _env_path = Path(__file__).parent / '.env'
    if _env_path.exists():
        load_dotenv(_env_path)
except ImportError:
    pass  # python-dotenv non installato, si usa solo env di sistema

# ─── COSTANTI ─────────────────────────────────────────────────────────────────
TITLE_MIN, TITLE_MAX = 50, 60
DESC_MIN,  DESC_MAX  = 140, 155

# ─── URL NORMALIZER ───────────────────────────────────────────────────────────
def normalize_url(url: str) -> str:
    url = str(url).strip().lower()
    url = re.sub(r'^https?://', '', url)
    url = re.sub(r'^www\.', '', url)
    url = url.rstrip('/')
    return url

def canonical_url(norm: str) -> str:
    return 'https://www.' + norm + '/'

# ─── LANGUAGE DETECTOR ────────────────────────────────────────────────────────
_IT = ['scopri','della','degli','delle','per','con','una','uno','che',
       'come','nella','nei','sui','questa','questi','anche','sono','puoi',
       'capelli','pelle','viso','prodotti','crema','trattamento','guida',
       'tua','tuo','nostro','nostra','ora','subito','acquista','leggi',
       'segreti','benefici','idratante','garnier','loreal','loréal']
_EN = ['discover','your','the','and','for','with','our','how','best','get',
       'learn','shop','find','skin','hair','care','achieve','unlock',
       'explore','improve','boost','glow','radiant','long-lasting',
       'moisturize','hydrate','flawless','coverage','anti-aging','read',
       'struggling','tips','routine','guide','secrets','benefits']

def detect_lang(text: str) -> str:
    if not text or not str(text).strip():
        return 'empty'
    tl = str(text).lower()
    it = sum(1 for w in _IT if re.search(r'\b' + re.escape(w) + r'\b', tl))
    en = sum(1 for w in _EN if re.search(r'\b' + re.escape(w) + r'\b', tl))
    if it > en: return 'it'
    if en > it: return 'en'
    return 'mixed'


def needs_translation(text: str, target_lang: str = 'italiano') -> bool:
    """
    True se il testo deve essere tradotto in target_lang.
    Tratta 'mixed' come inglese per pagine non-inglesi — evita contenuti
    bilingue parziale che sfuggono al controllo detect_lang == 'en'.
    """
    if not text or not str(text).strip():
        return False
    detected   = detect_lang(text)
    target_iso = _LABEL_TO_ISO.get(target_lang, 'it')
    if detected == target_iso:
        return False                           # già nella lingua target
    if detected in ('en', 'mixed'):
        return target_iso != 'en'             # EN/misto → traduci se target ≠ inglese
    return False

# ─── LANGUAGE LABEL MAP ───────────────────────────────────────────────────────
# Mappa codici lingua (ISO 639-1/2, nomi estesi, locale) → etichetta in italiano
# usata nei prompt di traduzione.
_LANG_LABELS: Dict[str, str] = {
    # ISO 639-1
    'it': 'italiano',    'fr': 'francese',     'de': 'tedesco',
    'es': 'spagnolo',    'pt': 'portoghese',   'nl': 'olandese',
    'pl': 'polacco',     'ru': 'russo',         'cs': 'ceco',
    'sk': 'slovacco',    'hu': 'ungherese',    'ro': 'rumeno',
    'el': 'greco',       'tr': 'turco',         'ar': 'arabo',
    'zh': 'cinese',      'ja': 'giapponese',   'ko': 'coreano',
    'sv': 'svedese',     'da': 'danese',        'fi': 'finlandese',
    'nb': 'norvegese',   'no': 'norvegese',    'en': 'inglese',
    # ISO 639-2
    'ita': 'italiano',   'fra': 'francese',    'deu': 'tedesco',
    'spa': 'spagnolo',   'por': 'portoghese',  'nld': 'olandese',
    'pol': 'polacco',    'rus': 'russo',        'eng': 'inglese',
    # Nomi estesi (lowercase)
    'italian': 'italiano',   'french': 'francese',    'german': 'tedesco',
    'spanish': 'spagnolo',   'portuguese': 'portoghese', 'dutch': 'olandese',
    'polish': 'polacco',     'russian': 'russo',       'english': 'inglese',
    'czech': 'ceco',         'slovak': 'slovacco',     'hungarian': 'ungherese',
    'romanian': 'rumeno',    'greek': 'greco',          'turkish': 'turco',
    'swedish': 'svedese',    'danish': 'danese',        'finnish': 'finlandese',
    'norwegian': 'norvegese',
}

# ISO 639-1 → codice detect_lang, per evitare ritraduzione dell'invariato
_LABEL_TO_ISO: Dict[str, str] = {v: k for k, v in _LANG_LABELS.items() if len(k) == 2}


def page_language(raw: str) -> str:
    """
    Normalizza il valore grezzo della colonna Language del file WSX
    (es. 'it', 'IT', 'Italian', 'it-IT', 'fr_FR') in un'etichetta
    in italiano da usare nei prompt (es. 'italiano', 'francese').
    Se il codice non è riconosciuto, restituisce 'italiano' come default.
    """
    if not raw or not str(raw).strip():
        return 'italiano'
    code = str(raw).strip().lower()
    # Gestisce locale compositi: 'it-IT', 'fr_FR', 'pt-BR' → prende la prima parte
    code = re.split(r'[-_]', code)[0]
    return _LANG_LABELS.get(code, 'italiano')


# ─── LLM PROVIDERS ────────────────────────────────────────────────────────────
import anthropic as _anthropic

# Client cache (keyed by api_key string, avoids re-instantiation)
_ant_clients: Dict[str, object] = {}
_oai_clients: Dict[str, object] = {}

# Modelli disponibili (esposti all'UI)
ANT_MODELS = ['claude-haiku-4-5-20251001', 'claude-sonnet-4-6']
OAI_MODELS = ['gpt-4o-mini', 'gpt-4o']
_ANT_DEFAULT = ANT_MODELS[0]
_OAI_DEFAULT = OAI_MODELS[0]


def _get_anthropic(api_key: str = None) -> object:
    key = (api_key or '').strip() or os.getenv('ANTHROPIC_API_KEY', '')
    if not key:
        raise EnvironmentError(
            'ANTHROPIC_API_KEY non trovata. '
            'Aggiungila in .env oppure inseriscila nel campo API Key.'
        )
    if key not in _ant_clients:
        _ant_clients[key] = _anthropic.Anthropic(api_key=key)
    return _ant_clients[key]


def _get_openai(api_key: str = None) -> object:
    key = (api_key or '').strip() or os.getenv('OPENAI_API_KEY', '')
    if not key:
        raise EnvironmentError('OPENAI_API_KEY non trovata. Inseriscila nel campo API Key.')
    if key not in _oai_clients:
        try:
            from openai import OpenAI as _OAI
        except ImportError:
            raise ImportError('Esegui: pip install openai --break-system-packages')
        _oai_clients[key] = _OAI(api_key=key)
    return _oai_clients[key]


def _get_client():
    """Backward-compat alias."""
    return _get_anthropic()


def _call_llm(prompt: str, provider: str = 'anthropic',
              api_key: str = None, model: str = None,
              max_tokens: int = 300) -> str:
    """Esegue la chiamata LLM e restituisce il testo generato."""
    if provider == 'openai':
        m = model or _OAI_DEFAULT
        c = _get_openai(api_key)
        resp = c.chat.completions.create(
            model=m, max_tokens=max_tokens, temperature=0.3,
            messages=[{'role': 'user', 'content': prompt}]
        )
        return resp.choices[0].message.content.strip()
    else:
        m = model or _ANT_DEFAULT
        c = _get_anthropic(api_key)
        msg = c.messages.create(
            model=m, max_tokens=max_tokens,
            messages=[{'role': 'user', 'content': prompt}]
        )
        return msg.content[0].text.strip()


# ── Prompt builders ────────────────────────────────────────────────────────────

def _prompt_translate(text: str, tag_type: str, target_lang: str) -> str:
    specs = {
        'title':       'meta title SEO (50-60 caratteri, keyword nelle prime parole)',
        'description': 'meta description SEO (140-155 caratteri, CTA coinvolgente)',
        'h1':          'H1 SEO (30-70 caratteri, keyword principale)',
    }
    hint = specs.get(tag_type, 'testo SEO')
    return (
        f"Traduci il seguente {hint} in {target_lang}.\n"
        f"Rispetta i vincoli di lunghezza. "
        f"Mantieni brand names, nomi prodotto e simboli (|, +, &) invariati.\n"
        f"Restituisci SOLO il testo tradotto, senza virgolette ne spiegazioni.\n\n"
        f"Testo: {text}"
    )


def _prompt_authority_action(url: str, authority_score: int, metric_name: str,
                             metric_score, must_have: str, goal: str = '') -> str:
    """Prompt per generare una raccomandazione Authority sintetica in italiano."""
    score_str = f'{metric_score}/100' if metric_score is not None else 'N/A'
    lines = [
        'Sei un esperto SEO specializzato in Authority e Link Building.',
        'Genera UNA raccomandazione concisa e actionable in italiano per questa pagina.',
        '',
        f'URL: {url}',
        f'Score Authority globale: {authority_score}/100',
        f'Metrica prioritaria da migliorare: {metric_name} (score attuale: {score_str})',
    ]
    if goal:
        lines += [f'Obiettivo: {goal[:200]}']
    if must_have:
        lines += ['', 'Criteri Must Have (rispettali):', must_have[:350]]
    lines += [
        '',
        'La raccomandazione deve:',
        '- Essere in italiano',
        '- Essere max 150 caratteri',
        '- Essere specifica per la metrica indicata',
        '- Essere immediatamente actionable',
        '',
        'Restituisci SOLO la raccomandazione, senza virgolette né spiegazioni.',
    ]
    return '\n'.join(l for l in lines if l is not None)


def _prompt_jsonld_adapt(url: str, jsonld_template: str,
                         metric_name: str, must_have: str) -> str:
    """Prompt per adattare un template JSON-LD WSX all'URL italiana specificata."""
    schema_type = metric_name.replace(' Declaration', '')
    return (
        f"Sei un esperto di Structured Data / Schema.org per SEO.\n"
        f"Adatta il seguente JSON-LD {schema_type} per la pagina italiana indicata.\n\n"
        f"URL della pagina: {url}\n\n"
        f"Template JSON-LD WSX:\n{jsonld_template[:1800]}\n\n"
        f"Criteri Must Have:\n{must_have[:400]}\n\n"
        f"Istruzioni:\n"
        f"1. Sostituisci @id e mainEntityOfPage con l'URL reale della pagina\n"
        f"2. Adatta tutti i campi URL al dominio della pagina\n"
        f"3. Mantieni tutti i @type e le proprietà obbligatorie del Must Have\n"
        f"4. Lascia i placeholder (<product-name>, <step-text>, ecc.) per i dati da compilare\n"
        f"5. Assicurati che il JSON risultante sia valido\n\n"
        f"Restituisci SOLO il JSON-LD adattato, senza markdown né spiegazioni."
    )


def _prompt_tech_action(url: str, page_type: str, speed_score: int,
                        mom: float, technicals_score: int,
                        must_have: str, metric_name: str) -> str:
    """Prompt per raccomandazione Technical SEO specifica in italiano."""
    mom_str = f'{mom:+.1f}' if mom is not None else 'N/D'
    trend_ctx = (
        'forte regressione in corso' if mom is not None and mom < -17 else
        'regressione in corso' if mom is not None and mom < 0 else
        'trend in miglioramento' if mom is not None and mom > 0 else
        'stabile'
    )
    return (
        f"Sei un Technical SEO specialist esperto in performance web e mercato beauty italiano.\n"
        f"Genera una raccomandazione tecnica concisa e actionable in italiano.\n\n"
        f"URL: {url}\n"
        f"Page Type: {page_type}\n"
        f"Metrica prioritaria: {metric_name}\n"
        f"WSX Page Speed Score: {speed_score}/100 (0=critico, 100=ottimale)\n"
        f"Trend MoM: {mom_str} punti ({trend_ctx})\n"
        f"WSX Technicals Score: {technicals_score}/100\n\n"
        f"Must Have WSX:\n{must_have[:350]}\n\n"
        f"La raccomandazione deve:\n"
        f"- Essere in italiano\n"
        f"- Essere max 180 caratteri\n"
        f"- Indicare l'azione specifica e il tool da usare (Screaming Frog, GSC, PageSpeed Insights)\n"
        f"- Tenere conto del trend MoM\n\n"
        f"Restituisci SOLO la raccomandazione, senza virgolette né spiegazioni."
    )


def _prompt_optimize(tag_type: str, target_lang: str, url: str,
                     suggestion_text: str, must_have: str,
                     extracted_value: str) -> str:
    specs = {
        'title': (
            'meta title', '50-60 caratteri',
            'keyword principale nelle prime 3 parole; separatore | prima del brand'
        ),
        'description': (
            'meta description', '140-155 caratteri',
            'CTA finale (Scopri/Acquista/Leggi ora); tono diretto e coinvolgente'
        ),
        'h1': (
            'H1', '30-70 caratteri',
            'keyword principale; tono autorevole; senza CTA'
        ),
    }
    tag_name, char_range, style = specs.get(tag_type, ('tag SEO', '50-150 caratteri', ''))

    lines = [f"Sei un esperto SEO. Genera un {tag_name} ottimizzato in {target_lang}.", '']
    if url:
        lines += [f"URL della pagina: {url}", '']
    lines += [
        'Raccomandazione originale (fonte WSX, in inglese):',
        (suggestion_text or '')[:500], '',
    ]
    if must_have:
        lines += ['Criteri Must Have:', must_have[:350], '']
    if extracted_value:
        lines += ['Valore di partenza (ottimizza e traduci):', extracted_value, '']
    lines += [
        f"Genera UN SOLO {tag_name} che:",
        f"- Sia scritto interamente in {target_lang}",
        f"- Abbia esattamente {char_range}",
        f"- {style}" if style else '',
        "- Mantenga brand names, nomi prodotto e simboli (|, +, &) invariati",
        '',
        f"Restituisci SOLO il {tag_name}, senza virgolette ne spiegazioni.",
    ]
    return '\n'.join(l for l in lines if l is not None)


# ── Funzione principale ────────────────────────────────────────────────────────

def generate_seo_tag(    suggestion_text: str = '',
    tag_type: str        = 'description',
    target_lang: str     = 'italiano',
    url: str             = '',
    metric_name: str     = '',
    must_have: str       = '',
    extracted_value: str = '',
    mode: str            = 'optimize',
    provider: str        = 'anthropic',
    api_key: str         = None,
    model: str           = None,
) -> str:
    """
    Genera o traduce un tag SEO usando il provider LLM scelto.

    mode='translate': traduce extracted_value in target_lang (veloce).
    mode='optimize' : usa URL + suggestion + must_have per generare un tag
                      SEO-ready in target_lang (qualita superiore).
    """
    if mode == 'translate':
        if not extracted_value:
            return ''
        if not needs_translation(extracted_value, target_lang):
            return extracted_value
        prompt = _prompt_translate(extracted_value, tag_type, target_lang)
    else:
        if not suggestion_text and not extracted_value:
            return ''
        prompt = _prompt_optimize(
            tag_type=tag_type,
            target_lang=target_lang,
            url=url,
            suggestion_text=suggestion_text,
            must_have=must_have,
            extracted_value=extracted_value,
        )

    try:
        result = _call_llm(prompt, provider=provider, api_key=api_key, model=model)
        return result if len(result) > 5 else (extracted_value or '')
    except Exception:
        return extracted_value or ''


def translate_to_lang(text: str, target_lang: str = 'italiano',
                      context: str = '',
                      provider: str = 'anthropic',
                      api_key: str = None,
                      model: str = None) -> str:
    """Traduce text in target_lang. Wrapper di generate_seo_tag(mode='translate')."""
    return generate_seo_tag(
        suggestion_text=text,
        tag_type=context or 'description',
        target_lang=target_lang,
        extracted_value=text,
        mode='translate',
        provider=provider,
        api_key=api_key,
        model=model,
    )


def translate_to_italian(text: str, context: str = '') -> str:
    """Backward-compatible wrapper."""
    return translate_to_lang(text, 'italiano', context)


def translate_batch(records: List[Dict], lang_col_title: str = 'suggested_title',
                    lang_col_desc: str = 'suggested_description',
                    target_lang: str = 'italiano') -> List[Dict]:
    """Traduce in batch title/desc se in inglese."""
    for r in records:
        for col, ctx in [(lang_col_title, 'title'), (lang_col_desc, 'description')]:
            val = r.get(col, '')
            if val and needs_translation(val, target_lang):
                r[col + '_translated'] = translate_to_lang(val, target_lang, ctx)
                r[col + '_lang'] = 'translated'
            elif val:
                r[col + '_translated'] = val
                r[col + '_lang'] = detect_lang(val)
            else:
                r[col + '_translated'] = ''
                r[col + '_lang'] = 'empty'
    return records

# ─── ACTION CLASSIFIER ────────────────────────────────────────────────────────
def classify_action(text: str) -> List[str]:
    if pd.isna(text) or not str(text).strip():
        return ['no_suggestion']
    t = str(text).lower()
    actions = []
    if any(p in t for p in ['ensure 100%','populate','add a meta','add meta','missing','both title']):
        actions.append('populate_missing')
    if any(p in t for p in ['50-60','50–60','140-155','140–155','character','length','shorten','concise','optimal','standard length']):
        actions.append('fix_length')
    if any(p in t for p in ['keyword','search intent','primary keyword']):
        actions.append('add_keyword')
    if any(p in t for p in ['call to action','cta','click','engaging','compelling','action-oriented','persuasive']):
        actions.append('improve_cta')
    if any(p in t for p in ['change ','rewrite','revise','refine','enhance','optimize','improve']):
        actions.append('rewrite')
    return actions if actions else ['generic_advice']

# ─── METRIC NAME CLASSIFIER ───────────────────────────────────────────────────
_METRIC_TO_ACTION: Dict[str, str] = {
    'Meta Tags Optimisation': 'meta_tags_optimisation',
    'Relevance & UX':         'relevance_ux',
    'Unique Content':         'unique_content',
    'Heading Structure':      'heading_structure',
}

def classify_from_metric_name(metric_name: str, suggestion_text: str = '') -> List[str]:
    """
    Classifica l'azione SEO a partire da Top Recommendation - Metric Name,
    arricchendola con sub-azioni derivate dal testo della suggestion.
    """
    mn   = str(metric_name).strip() if metric_name and str(metric_name).strip() else ''
    base = [_METRIC_TO_ACTION[mn]] if mn in _METRIC_TO_ACTION else []

    # Sub-azioni testuali (length, keyword, CTA…) — solo per meta tags
    sub = classify_action(suggestion_text)
    if sub in (['generic_advice'], ['no_suggestion']):
        return base if base else sub

    seen, out = set(), []
    for a in (base + sub):
        if a not in seen:
            seen.add(a)
            out.append(a)
    return out if out else ['generic_advice']

# ─── SUGGESTION PARSER ────────────────────────────────────────────────────────
def parse_suggestion(text: str) -> dict:
    """
    Estrae il testo suggerito (title, description, h1) da una suggestion WSX.

    Gestisce:
    - virgolette singole  'VALUE'  e doppie  "VALUE"
    - pattern  change 'OLD' [paren] to [something like] 'NEW'
    - pattern  instead of 'OLD', try/use/consider [something like] 'NEW'
    - apostrofi interni a parole italiane/francesi (L'Oréal, dell', l'altra)
    """
    if pd.isna(text): return {}
    text = str(text).strip()
    if not text: return {}
    result = {}

    # ── Quote normalisation ───────────────────────────────────────────────────
    # Sostituisce "..." (≥ 10 caratteri) con '...' per processamento uniforme.
    # Usa un placeholder temporaneo (\x02...\x03) per evitare doppia sostituzione.
    norm = re.sub(
        r'"([^"]{10,})"',
        lambda m: '\x02' + m.group(1) + '\x03',
        text
    )
    norm = norm.replace('\x02', "'").replace('\x03', "'")

    # ── Closing-quote detector ────────────────────────────────────────────────
    # Chiusura valida: apostrofo NON seguito da lettera (ASCII o accentata).
    _Q_CLOSE = re.compile(r"'(?![a-zA-Z\u00c0-\u024f])")

    def _trim(val: str) -> Optional[str]:
        """Ritaglia val all'ultima chiusura logica. None se nessuna trovata."""
        closes = [m.start() for m in _Q_CLOSE.finditer(val)]
        if not closes:
            return None
        trimmed = val[:closes[-1]].strip()
        return trimmed if len(trimmed) >= 10 else None

    def _clean(val: str) -> Optional[str]:
        """Per valori già delimitati dal regex: pulisce senza esigere chiusura."""
        val = val.strip()
        # rimuove un eventuale apostrofo finale isolato
        if val.endswith("'") and not (len(val) > 1 and val[-2].isalpha()):
            val = val[:-1].strip()
        return val if len(val) >= 10 else None

    def _best(raw: str) -> Optional[str]:
        """Prova _trim; se non trova una chiusura valida, usa _clean."""
        return _trim(raw) or _clean(raw)

    # ── Pattern A ─────────────────────────────────────────────────────────────
    # change ['it from' | 'the meta X from' | ''] 'OLD' [optional paren] to [something like] 'NEW'
    change_re = re.compile(
        r"change\s+(?:(?:it|the\s+\w+(?:\s+\w+)?)\s+)?(?:from\s+)?"
        r"'((?:[^']|'(?=[a-zA-Z\u00c0-\u024f]))+)'\s*"
        r"(?:\([^)]*\))?\s*"
        r"to\s+(?:something like\s+)?"
        r"'((?:[^']|'(?=[a-zA-Z\u00c0-\u024f]))+)",
        re.IGNORECASE
    )
    for m in change_re.finditer(norm):
        new_val = _best(m.group(2))
        if not new_val:
            continue
        ctx = norm[:m.start()].lower()
        t_pos = ctx.rfind('title')
        d_pos = max(ctx.rfind('description'), ctx.rfind('meta desc'), 0)
        tag = 'title' if t_pos > d_pos else 'description'
        result.setdefault(f'current_{tag}', _best(m.group(1)) or m.group(1).strip())
        result.setdefault(f'suggested_{tag}', new_val)

    # ── Pattern B ─────────────────────────────────────────────────────────────
    # instead of 'OLD'[,/.] try/use/consider [something like] 'NEW'
    instead_re = re.compile(
        r"instead of\s+'(?:[^']|'(?=[a-zA-Z\u00c0-\u024f]))+'[,.]?\s*"
        r"(?:try|use|consider)\s*(?:something like\s+)?"
        r"'((?:[^']|'(?=[a-zA-Z\u00c0-\u024f]))+)",
        re.IGNORECASE
    )
    for m in instead_re.finditer(norm):
        new_val = _best(m.group(1))
        if not new_val or len(new_val) < 10:
            continue
        ctx = norm[:m.start()].lower()
        t_pos = ctx.rfind('title')
        d_pos = max(ctx.rfind('description'), ctx.rfind('meta desc'), 0)
        tag = 'title' if t_pos > d_pos else 'description'
        result.setdefault(f'suggested_{tag}', new_val)

    # ── Section split ─────────────────────────────────────────────────────────
    tl = norm.lower()
    desc_start = -1
    for phrase in ['meta description', 'description to be', 'description to create',
                   'improve the description', 'refine the description',
                   'enhance the description', 'the description', 'meta desc']:
        idx = tl.find(phrase)
        if idx > -1 and (desc_start == -1 or idx < desc_start):
            desc_start = idx

    if desc_start > 0 and 'title' in tl[:desc_start]:
        sections = [('title', norm[:desc_start]), ('description', norm[desc_start:])]
    elif 'title' in tl and desc_start == -1:
        sections = [('title', norm)]
    else:
        sections = [('description', norm)]

    def _extract(section: str) -> Optional[str]:
        """Estrae un valore suggerito da una sezione (title o description)."""

        # (a) change [it / the meta X / this X] to: 'VALUE'
        m = re.search(
            r"change\s+(?:it|the|this)(?:\s+[\w\s]{0,30})?\s+to:?\s*"
            r"'((?:[^']|'(?=[a-zA-Z\u00c0-\u024f]))+)",
            section, re.IGNORECASE
        )
        if m:
            v = _best(m.group(1))
            if v: return v

        # (b) consider/perhaps/e.g.: 'VALUE'
        m2 = re.search(
            r"(?:consider|perhaps|e\.g\.)[^:]*:\s*"
            r"'((?:[^']|'(?=[a-zA-Z\u00c0-\u024f])){15,})",
            section, re.IGNORECASE
        )
        if m2:
            v = _best(m2.group(1))
            if v: return v

        # (c) For example / Example: / e.g., / such as → 'VALUE'
        for marker in ['For example,', 'For example:', 'Example:', 'e.g.,', 'e.g.', 'such as']:
            idx = section.rfind(marker)
            if idx == -1: continue
            after = section[idx + len(marker):].strip()
            if re.match(r"change\s+", after, re.I): continue
            # strip leading "try [something like]"
            after = re.sub(r'^try\s+(?:something like\s+)?', '', after, flags=re.IGNORECASE)
            q = after.find("'")
            if q > -1:
                content = after[q + 1:]
                v = _trim(content)          # richiede chiusura valida
                if v and len(v) > 10: return v

        # (d) try [something like] 'VALUE'  (senza marker precedente)
        m3 = re.search(
            r"try\s+(?:something like\s+)?"
            r"'((?:[^']|'(?=[a-zA-Z\u00c0-\u024f]))+)",
            section, re.IGNORECASE
        )
        if m3:
            v = _best(m3.group(1))
            if v and len(v) > 10: return v

        # (e) colon terminale + 'VALUE'
        m4 = re.search(
            r":\s*'((?:[^']|'(?=[a-zA-Z\u00c0-\u024f])){15,})'",
            section
        )
        if m4:
            return m4.group(1).strip()

        return None

    for tag_type, section in sections:
        if f'suggested_{tag_type}' in result:
            continue
        val = _extract(section)
        if val:
            result[f'suggested_{tag_type}'] = val

    # ── Primary keyword ───────────────────────────────────────────────────────
    kw = re.search(r"primary keyword[s]?\s+'([^']+)'", norm, re.I)
    if not kw:
        kw = re.search(r'primary keyword[s]?\s+"([^"]+)"', text, re.I)
    if kw:
        result['primary_keyword'] = kw.group(1).strip()

    return result

# ─── LENGTH HELPERS ───────────────────────────────────────────────────────────
def clen(s) -> Optional[int]:
    return len(str(s)) if s and pd.notna(s) and str(s).strip() else None

def len_status(n, lo, hi) -> str:
    if n is None: return ''
    if lo <= n <= hi: return 'OK ✓'
    return 'TOO SHORT' if n < lo else 'TOO LONG'

def delta_chars(a, b) -> str:
    if not a or not b: return ''
    d = len(str(b)) - len(str(a))
    return f'+{d}' if d >= 0 else str(d)
