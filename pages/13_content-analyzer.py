import streamlit as st
import requests
from requests.adapters import HTTPAdapter, Retry
from bs4 import BeautifulSoup
import json
from urllib.parse import urlparse
import pandas as pd
from collections import Counter
import re
from datetime import datetime
from typing import Optional, List, Dict, Tuple, Any
from rapidfuzz import fuzz  # pip install rapidfuzz
from sklearn.feature_extraction.text import TfidfVectorizer  # pip install scikit-learn

# =====================================
# Configurazione pagina
# =====================================
st.set_page_config(
    page_title="SERP Analyzer Pro v3",
    page_icon="🔍",
    layout="wide"
)

# =====================================
# Costanti e configurazioni
# =====================================
MARKETS: Dict[str, Dict[str, str]] = {
    "Italia": {"code": "it", "lang": "it", "google": "google.it", "location": "Italy"},
    "Francia": {"code": "fr", "lang": "fr", "google": "google.fr", "location": "France"},
    "Germania": {"code": "de", "lang": "de", "google": "google.de", "location": "Germany"},
    "Spagna": {"code": "es", "lang": "es", "google": "google.es", "location": "Spain"},
    "Regno Unito": {"code": "uk", "lang": "en", "google": "google.co.uk", "location": "United Kingdom"},
}

# Stopword minime multilingua (fallback)
STOPWORDS: Dict[str, set] = {
    "it": {
        'il','lo','la','i','gli','le','un','uno','una','di','da','a','in','su','per','con','tra','fra','e','o','ma','se',
        'che','chi','cui','come','quando','dove','quale','quanto','essere','avere','fare','dire','stare','questo','quello',
        'più','anche','molto','tutto','altro','sono','dei','nel','della','alla','delle','nella','sulla','dai','dalle','alla'
    },
    "en": {
        'the','a','an','of','to','in','on','for','and','or','but','if','is','are','be','been','being','this','that','these','those',
        'with','as','by','at','from','it','its','into','than','then','so','very','more','most','such','about'
    },
    "fr": {'le','la','les','un','une','des','de','du','au','aux','et','ou','mais','si','est','sont','être','avec','par','à','dans','pour','sur','ce','cet','cette','ces'},
    "de": {'der','die','das','ein','eine','und','oder','aber','wenn','ist','sind','sein','mit','von','zu','im','am','für','auf','als','dies','diese','dieser'},
    "es": {'el','la','los','las','un','una','unos','unas','de','del','y','o','pero','si','es','son','ser','con','por','para','en','al','como','este','esta','estos','estas'}
}

REQUESTS_TIMEOUT = 20

# =====================================
# Cache compat con Streamlit vecchi
# =====================================
try:
    cache_data = st.cache_data  # Streamlit >= 1.18
except AttributeError:
    cache_data = st.cache

# =====================================
# Utility HTTP
# =====================================

def _build_http_session() -> requests.Session:
    retries = Retry(total=3, backoff_factor=0.5, status_forcelist=(429, 500, 502, 503, 504))
    session = requests.Session()
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
    })
    session.mount('https://', HTTPAdapter(max_retries=retries))
    session.mount('http://', HTTPAdapter(max_retries=retries))
    return session

http = _build_http_session()

@cache_data(show_spinner=False, ttl=60*60)
def cached_get(url: str, params: Optional[Dict[str, Any]] = None) -> Any:
    r = http.get(url, params=params, timeout=REQUESTS_TIMEOUT)
    r.raise_for_status()
    return r.json() if 'application/json' in r.headers.get('Content-Type','') else r.text

@cache_data(show_spinner=False, ttl=60*30)
def cached_fetch_html(url: str) -> str:
    r = http.get(url, timeout=REQUESTS_TIMEOUT)
    r.raise_for_status()
    return r.text

# =====================================
# SERP Analyzer
# =====================================
class SERPAnalyzer:
    def __init__(self, api_key: str):
        self.api_key = api_key

    def get_serp_data(
        self,
        keyword: str,
        market: str = "it",
        lang: str = "it",
        google_domain: str = "google.it",
        location: Optional[str] = None
    ) -> Optional[Dict[str, Any]]:
        params: Dict[str, Any] = {
            "engine": "google",
            "q": keyword,
            "google_domain": google_domain,
            "gl": market,
            "hl": lang,
            "num": 20,
            "api_key": self.api_key,
        }
        if location:
            params["location"] = location
        try:
            data = cached_get("https://serpapi.com/search", params)
            if isinstance(data, str):
                data = json.loads(data)
            return data
        except Exception as e:
            st.error(f"Errore nel recupero dati SERP: {e}")
            return None

    @staticmethod
    def extract_serp_features(serp_data: Dict[str, Any]) -> Dict[str, Any]:
        features: Dict[str, Any] = {
            "organic_results": [],
            "people_also_ask": [],
            "related_searches": [],
            "images": [],
            "videos": [],
            "local_pack": [],
            "knowledge_graph": None,
            "featured_snippet": None
        }
        if not serp_data:
            return features

        # Organici
        for result in serp_data.get("organic_results", []):
            link = result.get("link") or result.get("url")
            features["organic_results"].append({
                "position": result.get("position"),
                "title": result.get("title"),
                "link": link,
                "snippet": result.get("snippet") or result.get("snippet_highlighted_words", [""]),
                "domain": urlparse(link or "").netloc
            })

        # PAA
        for q in serp_data.get("related_questions", []):
            features["people_also_ask"].append({
                "question": q.get("question"),
                "snippet": q.get("snippet"),
                "source": (q.get("source") or (q.get("link") and urlparse(q.get("link")).netloc))
            })

        # Related searches
        for s in serp_data.get("related_searches", []):
            q = s.get("query") if isinstance(s, dict) else s
            if q:
                features["related_searches"].append(q)

        # Featured snippet
        ab = serp_data.get("answer_box") or serp_data.get("featured_snippet")
        if ab:
            features["featured_snippet"] = {
                "type": ab.get("type"),
                "snippet": ab.get("snippet") or ab.get("answer"),
                "source": ab.get("link") or ab.get("source")
            }

        # Knowledge graph
        kg = serp_data.get("knowledge_graph")
        if kg:
            features["knowledge_graph"] = {
                "title": kg.get("title"),
                "type": kg.get("type"),
                "description": kg.get("description")
            }

        # Video
        for v in serp_data.get("inline_videos", []):
            features["videos"].append({"title": v.get("title"), "link": v.get("link")})

        # Immagini
        for img in serp_data.get("inline_images", [])[:10]:
            features["images"].append({"title": img.get("title"), "source": img.get("source")})

        # Local pack
        local = serp_data.get("local_results")
        places: List[Dict[str, Any]] = []
        if isinstance(local, dict):
            places = local.get("places", [])
        elif isinstance(local, list):
            places = local
        for loc in places:
            features["local_pack"].append({
                "title": loc.get("title"),
                "rating": loc.get("rating"),
                "reviews": loc.get("reviews"),
                "type": loc.get("type")
            })

        return features

# =====================================
# Content Analyzer
# =====================================
class ContentAnalyzer:
    @staticmethod
    def extract_content_from_url(url: str) -> Optional[Dict[str, Any]]:
        try:
            html = cached_fetch_html(url)
            soup = BeautifulSoup(html, 'html.parser')

            # Rimuovi elementi non utili
            for tag in soup(['script', 'style', 'nav', 'footer', 'noscript', 'iframe']):
                tag.decompose()

            # Titolo e meta
            title_el = soup.find('title')
            meta_desc_el = soup.find('meta', attrs={'name': 'description'})

            # Headings
            h1 = [h.get_text(strip=True) for h in soup.find_all('h1')]
            h2 = [h.get_text(strip=True) for h in soup.find_all('h2')]

            # Structured data (ld+json)
            schemas: List[Dict[str, Any]] = []
            for ld in soup.find_all('script', type='application/ld+json'):
                try:
                    schemas.append(json.loads(ld.string))
                except Exception:
                    pass

            # Testo
            text = soup.get_text(separator=' ', strip=True)
            text = re.sub(r"\s+", " ", text)

            return {
                "text": text,
                "title": title_el.string.strip() if title_el and title_el.string else "",
                "meta_description": meta_desc_el.get('content').strip() if meta_desc_el and meta_desc_el.get('content') else "",
                "word_count": len(text.split()),
                "h1": h1,
                "h2": h2,
                "schemas": schemas
            }
        except Exception as e:
            st.error(f"Errore nell'estrazione contenuto: {e}")
            return None

    @staticmethod
    def _normalize_text(text: str, lang: str) -> List[str]:
        text = text.lower()
        # Conserva lettere accentate europee
        text = re.sub(r"[^a-zA-Zàèéìòùäöüßçñãõâêîôûáéíóúüïœæ\s]", " ", text)
        tokens = text.split()
        stop = STOPWORDS.get(lang, set())
        tokens = [t for t in tokens if t not in stop and len(t) > 3]
        return tokens

    @staticmethod
    def extract_keywords_freq(text: str, lang: str = 'it', top_n: int = 30) -> List[Tuple[str, int]]:
        tokens = ContentAnalyzer._normalize_text(text, lang)
        freq = Counter(tokens)
        return freq.most_common(top_n)

    @staticmethod
    def extract_keywords_tfidf(corpus: List[str], lang: str = 'it', top_n: int = 30) -> List[Tuple[str, float]]:
        """TF‑IDF robusto con fallback multipli per corpora piccoli o testi rumorosi."""
        # Sanitize corpus
        cleaned: List[str] = []
        for doc in corpus or [""]:
            if not isinstance(doc, str):
                doc = str(doc or "")
            cleaned.append(doc.strip())
        if not any(cleaned):
            return []

        stop = list(STOPWORDS.get(lang, []))

        # Heuristic per max_df in base alla dimensione del corpus
        max_df = 1.0 if len(cleaned) < 3 else 0.85

        def _rank(vec: TfidfVectorizer, docs: List[str]) -> List[Tuple[str, float]]:
            X = vec.fit_transform(docs)
            means = X.toarray().mean(axis=0)
            terms = vec.get_feature_names_out()
            pairs = [(str(t), float(s)) for t, s in zip(terms, means)]
            # rimuovi token numerici puri
            pairs = [(t, s) for t, s in pairs if not re.fullmatch(r"[\d\.]+", t)]
            pairs.sort(key=lambda x: x[1], reverse=True)
            return pairs[:top_n]

        # Tentativo 1: bi-grammi, filtro stopwords, min_df=1, max_df dinamico, cap features
        try:
            vec = TfidfVectorizer(stop_words=stop, ngram_range=(1, 2), min_df=1, max_df=max_df, max_features=5000)
            return _rank(vec, cleaned)
        except ValueError:
            pass

        # Tentativo 2: solo unigrammi
        try:
            vec = TfidfVectorizer(stop_words=stop, ngram_range=(1, 1), min_df=1, max_df=1.0, max_features=5000)
            return _rank(vec, cleaned)
        except ValueError:
            pass

        # Tentativo 3: nessuno stopword (testi molto brevi)
        try:
            vec = TfidfVectorizer(ngram_range=(1, 1), min_df=1, max_df=1.0, max_features=3000)
            return _rank(vec, cleaned)
        except Exception:
            pass

        # Fallback finale: frequenze
        text_all = " ".join(cleaned)
        freq = Counter(ContentAnalyzer._normalize_text(text_all, lang))
        return [(k, float(v)) for k, v in freq.most_common(top_n)]

    @staticmethod
    def analyze_content_structure(text: str) -> Dict[str, Any]:
        paragraphs = [p.strip() for p in re.split(r"\n+", text) if p.strip()]
        return {
            "word_count": len(text.split()),
            "char_count": len(text),
            "paragraph_count": len(paragraphs),
            "avg_words_per_paragraph": (len(text.split()) / max(1, len(paragraphs)))
        }

# =====================================
# Competitor Analyzer
# =====================================
class CompetitorAnalyzer:
    def __init__(self, content_analyzer: ContentAnalyzer, lang: str, use_tfidf: bool):
        self.ca = content_analyzer
        self.lang = lang
        self.use_tfidf = use_tfidf

    def _extract_keywords(self, text: str, top_n: int = 30) -> List[Tuple[str, float]]:
        if self.use_tfidf:
            out = self.ca.extract_keywords_tfidf([text], self.lang, top_n)
            return [(k, float(s)) for k, s in out]
        freq = self.ca.extract_keywords_freq(text, self.lang, top_n)
        return [(k, float(s)) for k, s in freq]

    def analyze(self, serp_features: Dict[str, Any], user_domain: Optional[str] = None, top_n: int = 3) -> pd.DataFrame:
        rows: List[Dict[str, Any]] = []
        count = 0
        for r in serp_features.get("organic_results", []):
            if count >= top_n:
                break
            link = r.get("link")
            if not link:
                continue
            dom = urlparse(link).netloc
            if user_domain and dom == user_domain:
                continue
            data = ContentAnalyzer.extract_content_from_url(link)
            if not data:
                continue
            kw = self._extract_keywords(data.get("text", ""), 30)
            rows.append({
                "position": r.get("position"),
                "domain": dom,
                "url": link,
                "title": r.get("title"),
                "words": data.get("word_count", 0),
                "h1": "; ".join(data.get("h1", [])[:3]),
                "h2_count": len(data.get("h2", [])),
                "top_keywords": ", ".join([k for k, _ in kw[:10]])
            })
            count += 1
        return pd.DataFrame(rows)

    @staticmethod
    def keyword_overlap(user_keywords: List[Tuple[str, float]], competitor_df: pd.DataFrame) -> pd.DataFrame:
        user_set = set([k.lower() for k, _ in user_keywords])
        overlaps: List[int] = []
        for _, row in competitor_df.iterrows():
            comp_set = set([k.strip().lower() for k in (row.get("top_keywords") or "").split(",") if k.strip()])
            inter = user_set & comp_set
            overlaps.append(len(inter))
        competitor_df = competitor_df.copy()
        competitor_df["keyword_overlap"] = overlaps
        return competitor_df

# =====================================
# Snippet Scorer & FAQ Schema
# =====================================
class SnippetScorer:
    @staticmethod
    def compute_score(serp_features: Dict[str, Any], content_data: Dict[str, Any], content_text: str, paa_coverage: Dict[str, Any]) -> Tuple[int, List[str]]:
        score = 0
        reasons: List[str] = []
        if serp_features.get("featured_snippet"):
            score += 20; reasons.append("FS presente in SERP (+20)")
        if content_data.get("h2"):
            score += 10; reasons.append("Sottotitoli presenti (+10)")
        has_faq = any((isinstance(s, dict) and (s.get("@type") == "FAQPage" or (isinstance(s.get("@type"), list) and "FAQPage" in s.get("@type")))) for s in content_data.get("schemas", []))
        if has_faq:
            score += 15; reasons.append("Schema FAQPage presente (+15)")
        cov = float(paa_coverage.get("coverage_rate", 0))
        add = 15 if cov >= 50 else (8 if cov >= 25 else 0)
        score += add; reasons.append("Copertura PAA {:.0f}% (+{})".format(cov, add))
        has_bullets = bool(re.search(r"(^|\n)[\-\*\d]+\s+", content_text))
        if has_bullets:
            score += 10; reasons.append("Liste puntate/numerate (+10)")
        avg_words = len(content_text.split()) / max(1, len(re.split(r"\n+", content_text)))
        if avg_words <= 80:
            score += 10; reasons.append("Paragrafi brevi (+10)")
        wc = int(content_data.get("word_count", 0))
        if 600 <= wc <= 2000:
            score += 10; reasons.append("Lunghezza adeguata (+10)")
        score = max(0, min(100, score))
        return int(score), reasons

    @staticmethod
    def build_faq_schema(paa_not_covered: List[str], max_q: int = 6) -> str:
        items = [{
            "@type": "Question",
            "name": q,
            "acceptedAnswer": {"@type": "Answer", "text": "Aggiungi qui una risposta sintetica e diretta (40-60 parole)."}
        } for q in paa_not_covered[:max_q]]
        schema = {
            "@context": "https://schema.org",
            "@type": "FAQPage",
            "mainEntity": items
        }
        return json.dumps(schema, ensure_ascii=False, indent=2)

# =====================================
# UI — Sidebar
# =====================================
st.title("🔍 SERP Analyzer Pro - v3")
with st.sidebar:
    st.header("⚙️ Configurazioni")
    serpapi_key = st.text_input("SerpApi Key", type="password", help="Ottieni la tua chiave su serpapi.com")

    st.subheader("🌍 Mercato")
    selected_market = st.selectbox("Seleziona mercato", list(MARKETS.keys()))
    market_config = MARKETS[selected_market]

    st.subheader("⚗️ Opzioni Analisi")
    use_tfidf = st.checkbox("TF‑IDF (consigliato)", value=True)
    paa_threshold = st.slider("Soglia match PAA (fuzzy)", 50, 95, 70)
    st.subheader("🏁 Competitor")
    n_competitors = st.slider("Numero competitor da analizzare", 2, 10, 5)

# =====================================
# UI — Main inputs
# =====================================
col1, col2 = st.columns([1, 1])
with col1:
    st.subheader("🎯 Keyword da Analizzare")
    keyword = st.text_input("Inserisci la keyword", placeholder="es: consulente seo milano")
with col2:
    st.subheader("📄 Contenuto da Analizzare")
    content_source = st.radio("Sorgente contenuto", ["URL", "Testo diretto"], horizontal=True)

if content_source == "URL":
    content_input = st.text_input("URL del tuo sito", placeholder="https://example.com/pagina")
else:
    content_input = st.text_area("Incolla il testo del tuo contenuto", height=200)

# =====================================
# Azione
# =====================================
if st.button("🚀 Avvia Analisi", type="primary", use_container_width=True):
    if not serpapi_key:
        st.error("⚠️ Inserisci la tua SerpApi Key nella sidebar")
    elif not keyword:
        st.error("⚠️ Inserisci una keyword da analizzare")
    elif not content_input:
        st.error("⚠️ Fornisci il contenuto da analizzare (URL o testo)")
    else:
        with st.spinner("🔄 Analisi in corso..."):
            serp_analyzer = SERPAnalyzer(serpapi_key)
            content_analyzer = ContentAnalyzer()

            # 1) SERP
            st.info("📊 Recupero dati SERP…")
            serp_data = serp_analyzer.get_serp_data(
                keyword,
                market_config["code"],
                market_config["lang"],
                market_config["google"],
                market_config.get("location")
            )
            if not serp_data:
                st.error("❌ Impossibile recuperare dati SERP. Verifica la tua API key e riprova.")
                st.stop()
            serp_features = serp_analyzer.extract_serp_features(serp_data)

            # 2) Contenuto
            st.info("📝 Analisi contenuto…")
            if content_source == "URL":
                content_data = content_analyzer.extract_content_from_url(content_input)
                user_domain = urlparse(content_input).netloc
            else:
                content_data = {
                    "text": content_input,
                    "title": "",
                    "meta_description": "",
                    "word_count": len(content_input.split()),
                    "h1": [],
                    "h2": [],
                    "schemas": []
                }
                user_domain = None
            if not content_data:
                st.error("❌ Impossibile analizzare il contenuto fornito")
                st.stop()
            content_text = content_data["text"]

            # 3) Keyword
            st.info("🔑 Estrazione keyword…")
            lang = market_config["lang"]
            serp_text = " ".join([
                (r.get("title") or "") + " " + (" ".join(r.get("snippet")) if isinstance(r.get("snippet"), list) else (r.get("snippet") or ""))
                for r in serp_features["organic_results"]
            ])
            if not serp_text.strip():
                serp_text = keyword

            if use_tfidf:
                serp_keywords = content_analyzer.extract_keywords_tfidf([serp_text], lang, 50)
                content_keywords = content_analyzer.extract_keywords_tfidf([content_text], lang, 50)
                serp_keywords = [(k, float(s)) for k, s in serp_keywords]
                content_keywords = [(k, float(s)) for k, s in content_keywords]
            else:
                serp_keywords = content_analyzer.extract_keywords_freq(serp_text, lang, 50)
                content_keywords = content_analyzer.extract_keywords_freq(content_text, lang, 50)
                serp_keywords = [(k, float(s)) for k, s in serp_keywords]
                content_keywords = [(k, float(s)) for k, s in content_keywords]

            related_searches_text = " ".join(serp_features["related_searches"]) if serp_features["related_searches"] else ""
            related_keywords = content_analyzer.extract_keywords_freq(related_searches_text, lang, 100) if related_searches_text else []

            # 4) Gap + PAA
            st.info("🎯 Identificazione gap…")
            serp_set = set([kw.lower() for kw, _ in serp_keywords])
            content_set = set([kw.lower() for kw, _ in content_keywords])
            missing_keywords = sorted(list(serp_set - content_set))
            common_keywords = sorted(list(serp_set & content_set))
            coverage_rate = (len(common_keywords) / len(serp_set) * 100) if serp_set else 0.0

            content_lower = content_text.lower()
            covered: List[str] = []
            not_covered: List[str] = []
            for paa in serp_features["people_also_ask"]:
                q = paa.get("question", "")
                if not q:
                    continue
                score = fuzz.token_set_ratio(q.lower(), content_lower)
                (covered if score >= paa_threshold else not_covered).append(q)
            total = len(covered) + len(not_covered)
            paa_coverage = {
                "covered": covered,
                "not_covered": not_covered,
                "coverage_rate": (len(covered) / total * 100) if total else 0.0
            }

            # 5) Competitor
            comp = CompetitorAnalyzer(content_analyzer, lang, use_tfidf)
            competitors_df = comp.analyze(serp_features, user_domain=user_domain, top_n=n_competitors)
            if not competitors_df.empty:
                competitors_df = comp.keyword_overlap(content_keywords, competitors_df)

            # 6) Snippet score + FAQ schema
            fs_score, fs_reasons = SnippetScorer.compute_score(serp_features, content_data, content_text, paa_coverage)
            faq_schema = SnippetScorer.build_faq_schema(paa_coverage.get("not_covered", [])) if paa_coverage.get("not_covered") else ""

            # =====================================
            # RISULTATI
            # =====================================
            st.success("✅ Analisi completata!")
            st.markdown("---")

            st.header("📊 Metriche Principali")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Risultati Organici", len(serp_features["organic_results"]))
            m2.metric("People Also Ask", len(serp_features["people_also_ask"]))
            m3.metric("Ricerche Correlate", len(serp_features["related_searches"]))
            m4.metric("Parole Contenuto", content_data.get("word_count", 0))

            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                "🎯 Gap Analisi",
                "🔍 SERP Features",
                "📝 Contenuto",
                "🏁 Competitor",
                "⭐ Snippet & FAQ",
                "📊 Export Dati"
            ])

            with tab1:
                st.header("🎯 Analisi Gap Contenuti")
                a1, a2 = st.columns(2)
                a1.metric("Copertura Keyword", f"{coverage_rate:.1f}%", delta=f"{len(common_keywords)} in comune")
                a2.metric("Copertura PAA", f"{paa_coverage['coverage_rate']:.1f}%", delta=f"{len(paa_coverage['covered'])} coperte")

                st.subheader("❌ Keyword Mancanti (priorità: correlate)")
                related_set = set([kw.lower() for kw, _ in related_keywords])
                missing_in_related = [kw for kw in missing_keywords if kw in related_set]
                if missing_in_related:
                    st.dataframe(pd.DataFrame({"Keyword": missing_in_related[:20]}), use_container_width=True)
                    st.info("💡 Queste keyword provengono dalle **Ricerche Correlate** e non sono presenti nel contenuto")
                else:
                    st.success("✅ Le keyword delle ricerche correlate risultano coperte")

                st.subheader("❓ Domande PAA Non Coperte")
                if paa_coverage['not_covered']:
                    for q in paa_coverage['not_covered']:
                        st.warning(f"❓ {q}")
                else:
                    st.success("✅ Tutte le PAA risultano coperte")

            with tab2:
                st.header("🔍 SERP Features")
                if serp_features["featured_snippet"]:
                    st.subheader("⭐ Featured Snippet")
                    fs = serp_features["featured_snippet"]
                    st.info(f"**Tipo:** {fs['type']}")
                    st.write(fs.get('snippet', ''))
                    if fs.get('source'):
                        st.caption(f"Fonte: {fs['source']}")
                if serp_features["knowledge_graph"]:
                    st.subheader("📚 Knowledge Graph")
                    kg = serp_features["knowledge_graph"]
                    st.write(f"**{kg['title']}** ({kg.get('type','')})")
                    st.write(kg.get('description',''))
                st.subheader("🔗 Top 10 Risultati Organici")
                if serp_features["organic_results"]:
                    st.dataframe(pd.DataFrame(serp_features["organic_results"][:10]), use_container_width=True)
                if serp_features["people_also_ask"]:
                    st.subheader("❓ People Also Ask")
                    for paa in serp_features["people_also_ask"]:
                        with st.expander(paa.get("question", "Domanda")):
                            st.write(paa.get("snippet", ""))
                            st.caption(f"Fonte: {paa.get('source', 'N/A')}")
                if serp_features["related_searches"]:
                    st.subheader("🔎 Ricerche Correlate")
                    cols = st.columns(3)
                    for idx, search in enumerate(serp_features["related_searches"]):
                        cols[idx % 3].info(search)

            with tab3:
                st.header("📝 Analisi del Tuo Contenuto")
                structure = content_analyzer.analyze_content_structure(content_text)
                d1, d2, d3 = st.columns(3)
                d1.metric("Parole Totali", structure["word_count"])
                d2.metric("Caratteri", structure["char_count"])
                d3.metric("Paragrafi", structure["paragraph_count"])

                st.subheader("🔑 Top 20 Keyword nel Contenuto")
                top20 = content_keywords[:20]
                st.dataframe(pd.DataFrame(top20, columns=["Keyword", "Score/Freq"]), use_container_width=True)

                if content_data.get("title") or content_data.get("meta_description"):
                    st.subheader("📋 Meta Informazioni")
                    if content_data.get("title"):
                        st.write(f"**Title:** {content_data['title']}")
                    if content_data.get("meta_description"):
                        st.write(f"**Meta Description:** {content_data['meta_description']}")
                if content_data.get("h1") or content_data.get("h2"):
                    st.subheader("#️⃣ Headings trovati")
                    if content_data.get("h1"):
                        st.write("**H1:**", ", ".join(content_data["h1"]))
                    if content_data.get("h2"):
                        st.write("**H2:**", ", ".join(content_data["h2"]))

            with tab4:
                st.header("🏁 Analisi Competitor (Top SERP)")
                if competitors_df.empty:
                    st.info("Nessun competitor analizzato (SERP vuota o errori di crawling)")
                else:
                    st.dataframe(competitors_df, use_container_width=True)
                    st.caption("Overlap = numero di keyword top condivise con il tuo contenuto")

            with tab5:
                st.header("⭐ Probabilità Featured Snippet & FAQ")
                c1, c2 = st.columns([1, 1])
                with c1:
                    st.metric("Snippet Score", f"{fs_score}/100")
                    for r in fs_reasons:
                        st.write("• ", r)
                with c2:
                    if faq_schema:
                        st.subheader("FAQPage JSON-LD suggerito")
                        st.code(faq_schema, language="json")
                    else:
                        st.info("Tutte le PAA risultano coperte: nessuna FAQ suggerita.")

            with tab6:
                st.header("📊 Export Dati")
                export_data = {
                    "keyword": keyword,
                    "market": selected_market,
                    "analysis_date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "metrics": {
                        "keyword_coverage": coverage_rate,
                        "paa_coverage": paa_coverage['coverage_rate'],
                        "content_word_count": content_data['word_count'],
                        "organic_results": len(serp_features["organic_results"]),
                        "paa_count": len(serp_features["people_also_ask"]),
                        "snippet_score": fs_score
                    },
                    "missing_keywords": missing_keywords[:30],
                    "uncovered_paa": paa_coverage['not_covered'],
                    "competitors_table": competitors_df.to_dict(orient='records') if not competitors_df.empty else [],
                }
                e1, e2 = st.columns(2)
                with e1:
                    st.download_button(
                        label="📥 Download Report JSON",
                        data=json.dumps(export_data, indent=2, ensure_ascii=False),
                        file_name=f"serp_analysis_{keyword.replace(' ', '_')}.json",
                        mime="application/json"
                    )
                with e2:
                    if serp_features["organic_results"]:
                        csv_data = pd.DataFrame(serp_features["organic_results"]).to_csv(index=False)
                        st.download_button(
                            label="📥 Download Risultati CSV",
                            data=csv_data,
                            file_name=f"organic_results_{keyword.replace(' ', '_')}.csv",
                            mime="text/csv"
                        )
                with st.expander("👁️ Preview Report JSON"):
                    st.json(export_data)

st.markdown("---")
st.caption("🔍 **SERP Analyzer Pro v3** | ES - SEO - TOOLS")