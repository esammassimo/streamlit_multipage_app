import streamlit as st
from openai import OpenAI
import re
from urllib.parse import urlparse, parse_qs
from docx import Document
import io
import zipfile
from datetime import datetime
import time
import importlib

# ============================
# Import youtube-transcript-api dinamico + introspezione
# ============================
YTA = None
TranscriptsDisabled = Exception
NoTranscriptFound = Exception
YTA_VERSION = "unknown"
YTA_FILE = "unknown"

try:
    yta_mod = importlib.import_module("youtube_transcript_api")
    YTA_FILE = getattr(yta_mod, "__file__", "unknown")
    YTA_VERSION = getattr(yta_mod, "__version__", "unknown")
    YTA = getattr(yta_mod, "YouTubeTranscriptApi", None)
    TranscriptsDisabled = getattr(yta_mod, "TranscriptsDisabled", Exception)
    NoTranscriptFound = getattr(yta_mod, "NoTranscriptFound", Exception)
except Exception as e:
    st.error(f"❌ Impossibile importare 'youtube_transcript_api': {type(e).__name__}: {e}")
    st.stop()

# ============================
# Helpers: estrazione ID e retry
# ============================
def extract_video_id(url: str):
    """Estrae l'ID (11 char) da tutte le varianti comuni di URL YouTube."""
    if not url:
        return None
    try:
        p = urlparse(url)
        if p.netloc:
            qs = parse_qs(p.query)
            if 'v' in qs and len(qs['v'][0]) == 11:
                return qs['v'][0]
            m = re.search(r"/shorts/([A-Za-z0-9_-]{11})", p.path)
            if m: return m.group(1)
            m = re.search(r"/embed/([A-Za-z0-9_-]{11})", p.path)
            if m: return m.group(1)
            m = re.search(r"youtu\.be/([A-Za-z0-9_-]{11})", p.netloc + p.path)
            if m: return m.group(1)
    except Exception:
        pass
    for pat in [
        r"(?:v=|/v/|&v=|watch\?v=)([A-Za-z0-9_-]{11})",
        r"youtu\.be/([A-Za-z0-9_-]{11})",
        r"/embed/([A-Za-z0-9_-]{11})",
        r"/shorts/([A-Za-z0-9_-]{11})",
    ]:
        m = re.search(pat, url)
        if m: return m.group(1)
    return None

def _with_retry(fn, *args, max_attempts=4, **kwargs):
    """Retry con backoff esponenziale (0.5, 1, 2s) per errori temporanei."""
    delay = 0.5
    for attempt in range(max_attempts):
        try:
            return fn(*args, **kwargs)
        except TranscriptsDisabled:
            raise
        except NoTranscriptFound:
            # errore strutturale, non ha senso ritentare
            raise
        except Exception:
            if attempt == max_attempts - 1:
                raise
            time.sleep(delay)
            delay = min(2.0, delay * 2)

# ============================
# Fallback via yt_dlp (opzionale ma consigliato)
# ============================
def fetch_transcript_via_ytdlp(video_id: str, language_code: str, allow_fallback: bool):
    """
    Ottiene sottotitoli con yt_dlp senza scaricare il video.
    Ritorna lista di dict [{'text': '...', 'start': float, 'duration': float}].
    """
    try:
        import yt_dlp
        import requests
    except Exception:
        raise NoTranscriptFound("yt_dlp/requests non installati per il fallback.")

    url = f"https://www.youtube.com/watch?v={video_id}"
    ydl_opts = {
        "skip_download": True,
        "quiet": True,
        "no_warnings": True,
    }
    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        info = ydl.extract_info(url, download=False)

        candidates = []

        def collect(subs_dict, label):
            if not subs_dict:
                return
            # esatta
            if language_code in subs_dict:
                candidates.append((subs_dict[language_code], f"{label}:{language_code}"))
            # base es. en-US -> en
            base = language_code.split("-")[0].lower()
            if base in subs_dict:
                candidates.append((subs_dict[base], f"{label}:{base}"))
            # qualsiasi se fallback
            if allow_fallback:
                for lang, entries in subs_dict.items():
                    candidates.append((entries, f"{label}:{lang}"))

        collect(info.get("subtitles"), "subs")
        collect(info.get("automatic_captions"), "auto")

        if not candidates:
            raise NoTranscriptFound("Nessun sottotitolo disponibile via yt_dlp.")

        tracks, _desc = candidates[0]
        track = tracks[0]  # tipicamente .vtt
        srt_url = track.get("url")
        if not srt_url:
            raise NoTranscriptFound("URL sottotitoli non disponibile.")

        r = requests.get(srt_url, timeout=20)
        r.raise_for_status()
        vtt = r.text

        entries = []
        for block in re.split(r"\n\s*\n", vtt.strip()):
            lines = [ln.strip() for ln in block.splitlines() if ln.strip()]
            if len(lines) < 2:
                continue
            time_line = None
            for ln in lines:
                if "-->" in ln:
                    time_line = ln
                    break
            if not time_line:
                continue
            tm = re.match(r"(?P<start>[\d:.]+)\s*-->\s*(?P<end>[\d:.]+)", time_line)
            if not tm:
                continue
            def to_seconds(s):
                parts = s.split(":")
                parts = [float(p.replace(",", ".")) for p in parts]
                while len(parts) < 3:
                    parts = [0.0] + parts
                h, m, sec = parts[-3], parts[-2], parts[-1]
                return h*3600 + m*60 + sec
            start = to_seconds(tm.group("start"))
            end = to_seconds(tm.group("end"))
            duration = max(0.0, end - start)

            text_lines = []
            for ln in lines:
                if ln == time_line:
                    continue
                if re.fullmatch(r"\d+", ln):
                    continue
                text_lines.append(ln)
            text = " ".join(text_lines).strip()
            if text:
                entries.append({"text": text, "start": start, "duration": duration})

        if not entries:
            raise NoTranscriptFound("Impossibile parsare i sottotitoli VTT via yt_dlp.")
        return entries

# ============================
# Fetch transcript – percorsi modern/compat + orchestratore
# ============================
def fetch_transcript_modern(video_id: str, language_code: str, allow_fallback: bool = True):
    """
    Percorso moderno (quando esiste list_transcripts). NON usa get_transcript
    se il metodo non è presente nella tua versione della libreria.
    """
    base_lang = language_code.split("-")[0].lower() if language_code else "en"
    prefer_list = []
    if language_code:
        prefer_list.append(language_code)
    if base_lang and base_lang not in prefer_list:
        prefer_list.append(base_lang)

    # 1) Tentativo diretto SOLO se get_transcript esiste
    if YTA and hasattr(YTA, "get_transcript"):
        try:
            return _with_retry(getattr(YTA, "get_transcript"), video_id, languages=prefer_list)
        except NoTranscriptFound:
            pass
        except Exception:
            pass

    # 2) Percorso list_transcripts (sappiamo che esiste per questa funzione)
    transcripts = _with_retry(getattr(YTA, "list_transcripts"), video_id)

    # 2a) match lingua esatta/base
    try:
        return transcripts.find_transcript(prefer_list).fetch()
    except NoTranscriptFound:
        if not allow_fallback:
            raise

    # 2b) traduzione verso lingua richiesta/base (se disponibile)
    if allow_fallback:
        for tr in transcripts:
            try:
                if getattr(tr, "is_translatable", False):
                    for tgt in prefer_list:
                        try:
                            return tr.translate(tgt).fetch()
                        except Exception:
                            continue
            except Exception:
                continue

    # 2c) prima disponibile
    for tr in transcripts:
        try:
            return tr.fetch()
        except Exception:
            continue

    raise NoTranscriptFound(f"Nessuna trascrizione recuperabile per video {video_id}.")

def fetch_transcript_any(video_id: str, language_code: str, allow_fallback: bool = True):
    """
    Ordine tentativi:
    1) list_transcripts (se disponibile)
    2) get_transcript (se disponibile)
    3) yt_dlp fallback
    Ritorna: (entries, used_path)
    """
    base_lang = language_code.split("-")[0].lower() if language_code else "en"
    prefer_list = [lc for lc in [language_code, base_lang] if lc]

    # 1) list_transcripts path
    if YTA and hasattr(YTA, "list_transcripts"):
        entries = fetch_transcript_modern(video_id, language_code, allow_fallback)
        return entries, "list_transcripts"

    # 2) get_transcript path
    if YTA and hasattr(YTA, "get_transcript"):
        langs = prefer_list[:]
        if allow_fallback:
            for extra in ["en", "en-US", "en-GB", "it", "es", "de", "fr", "pt", "pt-BR"]:
                if extra not in langs:
                    langs.append(extra)
        for lc in langs:
            try:
                entries = _with_retry(getattr(YTA, "get_transcript"), video_id, languages=[lc])
                return entries, f"get_transcript({lc})"
            except TranscriptsDisabled:
                raise
            except NoTranscriptFound:
                continue
            except Exception:
                continue

    # 3) Fallback yt_dlp
    entries = fetch_transcript_via_ytdlp(video_id, language_code, allow_fallback)
    return entries, "yt_dlp"

# ============================
# Helpers: pulizia testo e sommario
# ============================
def clean_transcript_text(text: str) -> str:
    text = re.sub(r"\[(?:Music|Applause|Laughter)\]", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def summarize_text(client: OpenAI, text: str, model: str, temperature: float, min_words: int = 400) -> str:
    MAX_CHARS = 12000
    chunks = [text[i:i + MAX_CHARS] for i in range(0, len(text), MAX_CHARS)]

    if len(chunks) == 1:
        prompt = (
            f"Summarize the following transcript, keeping the key points and context. "
            f"Ensure the summary is at least {min_words} words, clear, concise, non-repetitive, and highlights the most important aspects.\n\n"
            f"Transcript:\n{chunks[0]}\n\n"
            "Rules:\n- Keep it clear and structured (use short paragraphs or bullet points if helpful).\n"
            "- Avoid redundancy and trivial details.\n- Preserve all critical explanations, definitions, and conclusions."
        )
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are an expert at synthesizing video transcripts into accurate, readable long summaries."},
                {"role": "user", "content": prompt},
            ],
            temperature=temperature,
            max_tokens=4000,
        )
        return resp.choices[0].message.content

    partial_summaries = []
    for idx, ch in enumerate(chunks, 1):
        prompt = (
            f"Summarize part {idx} of a longer transcript. Focus on key points, arguments, data, and conclusions. "
            f"Length: ~250-350 words.\n\nPart {idx}:\n{ch}"
        )
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are an expert note-taker."},
                {"role": "user", "content": prompt},
            ],
            temperature=temperature,
            max_tokens=1200,
        )
        partial_summaries.append(resp.choices[0].message.content)

    synthesis_prompt = (
        f"You are given {len(partial_summaries)} partial summaries from a video transcript. "
        f"Merge them into a single cohesive summary of at least {min_words} words. "
        "Avoid duplication, keep a logical flow, and highlight the most critical insights."
        "\n\nPartial summaries:\n" + "\n".join(f"Part {i+1}:\n{ps}" for i, ps in enumerate(partial_summaries))
    )
    final_resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": "You are an expert editor."},
            {"role": "user", "content": synthesis_prompt},
        ],
        temperature=temperature,
        max_tokens=4000,
    )
    return final_resp.choices[0].message.content

# ============================
# UI
# ============================
st.set_page_config(page_title="YouTube Transcript & Summary Generator", page_icon="🎥", layout="centered")
st.title("🎥 YouTube Transcript & Summary Generator")

with st.sidebar:
    st.subheader("🔐 OpenAI Settings")
    if 'oai_api_key' not in st.session_state:
        st.session_state['oai_api_key'] = ""
    oai_key = st.text_input("Insert your OpenAI API KEY", type="password", value=st.session_state['oai_api_key'])
    st.session_state['oai_api_key'] = oai_key

    # Modelli (incluso GPT-5)
    model = st.selectbox(
        "Model",
        options=[
            "gpt-5",
            "gpt-5-chat-latest",
            "gpt-5-mini",
            "gpt-4o",
            "gpt-4o-mini",
            "gpt-4.1",
            "gpt-4.1-mini",
            "gpt-4",
        ],
        index=0,
        help="Scegli un modello compatibile con l'API chat.completions."
    )
    temperature = st.slider("Creativity (temperature)", 0.0, 1.0, 0.7, 0.1)
    min_words = st.number_input("Minimum summary length (words)", min_value=150, max_value=2000, value=400, step=50)

# Diagnostica della libreria caricata
st.caption(f"📦 youtube-transcript-api — version: {YTA_VERSION}, file: {YTA_FILE}")
st.caption(f"🔍 capabilities → list_transcripts: {hasattr(YTA, 'list_transcripts') if YTA else False}, get_transcript: {hasattr(YTA, 'get_transcript') if YTA else False}")

youtube_url = st.text_input("📺 Enter YouTube video URL")
language_code = st.text_input("🌍 Enter language code (default 'en')", value="en")
allow_fallback = st.checkbox("If not found, try other available languages", value=True)

with st.expander("🧭 How to use"):
    st.markdown(
        """
        1. Enter your OpenAI API key in the sidebar.
        2. Paste a YouTube video URL.
        3. Optionally set the transcript language code (e.g., `en`, `it`, `es`).
        4. Adjust model, creativity, and minimum summary length.
        5. Click **Generate Transcript and Summary**.
        6. Download the ZIP with `.txt` transcript and `.docx` summary.

        Notes:
        - Alcuni video (shorts/livestream/age-restricted/geo-blocked) potrebbero non avere transcript.
        - Se la lingua richiesta non è disponibile e il fallback è attivo, si proveranno alternative e traduzioni.
        - Video lunghi sono gestiti con riassunto a chunk + sintesi finale.
        """
    )

if not oai_key:
    st.warning("⚠️ Insert your API KEY in the sidebar to continue.")
    st.stop()

client = OpenAI(api_key=oai_key)

if st.button("📄 Generate Transcript and Summary", use_container_width=True):
    if not youtube_url:
        st.error("⚠️ You must enter the video URL!")
        st.stop()

    video_id = extract_video_id(youtube_url)
    if not video_id:
        st.error("❌ Invalid YouTube URL: unable to extract a valid video ID.")
        st.stop()

    st.caption(f"🎯 Detected Video ID: `{video_id}`")

    try:
        # orchestration: modern -> compat -> yt_dlp
        entries, used_path = fetch_transcript_any(video_id, language_code, allow_fallback=allow_fallback)
        st.info(f"✅ Transcript fetched via: **{used_path}**")

        full_text = " ".join([e.get("text", "") for e in entries])
        full_text = clean_transcript_text(full_text)

        if not full_text.strip():
            st.error("❌ Transcript retrieved but empty.")
            st.stop()

        with st.spinner("🧠 Generating summary with OpenAI..."):
            summary = summarize_text(client, full_text, model=model, temperature=temperature, min_words=int(min_words))

        # Costruisci ZIP (TXT transcript + DOCX summary)
        transcript_bytes = full_text.encode('utf-8')
        transcript_filename = f"transcript_{video_id}.txt"

        doc = Document()
        doc.add_heading("Riassunto del Video", level=1)
        doc.add_paragraph(summary)
        docx_buffer = io.BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)
        summary_filename = f"summary_{video_id}.docx"

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(transcript_filename, transcript_bytes)
            zf.writestr(summary_filename, docx_buffer.getvalue())
        zip_buffer.seek(0)
        zip_filename = f"youtube_summary_{video_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

        st.success("✅ Transcript and Summary successfully generated!")
        st.download_button(
            "📥 Download ZIP with Transcript & Summary",
            data=zip_buffer,
            file_name=zip_filename,
            mime="application/zip",
            use_container_width=True,
        )

        with st.expander("👀 Preview summary"):
            st.write(summary)

    except TranscriptsDisabled:
        st.error("❌ Transcripts are disabled for this video.")
    except NoTranscriptFound as e:
        st.error(f"❌ No transcript found: {e}")
    except Exception as e:
        st.error(f"❌ Error during extraction: {type(e).__name__}: {e}")