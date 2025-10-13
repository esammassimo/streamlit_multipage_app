import streamlit as st
from openai import OpenAI
import re
from urllib.parse import urlparse, parse_qs
from docx import Document
import io
import zipfile
from datetime import datetime
import time

# ====== youtube-transcript-api (import + version/capability detection) ======
try:
    from youtube_transcript_api import YouTubeTranscriptApi as YTA  # type: ignore
    try:
        from youtube_transcript_api import TranscriptsDisabled, NoTranscriptFound  # type: ignore
    except Exception:
        # Fallback: definisci eccezioni se assenti nelle versioni antiche (evita crash)
        class TranscriptsDisabled(Exception): ...
        class NoTranscriptFound(Exception): ...
    try:
        from youtube_transcript_api import __version__ as YTA_VERSION  # type: ignore
    except Exception:
        YTA_VERSION = "unknown"
except Exception as e:
    st.error(f"Impossibile importare youtube-transcript-api: {type(e).__name__}: {e}")
    st.stop()

# ============================
# Helpers: ID & transcript
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
            m = re.search(r"/shorts/([a-zA-Z0-9_-]{11})", p.path)
            if m:
                return m.group(1)
            m = re.search(r"/embed/([a-zA-Z0-9_-]{11})", p.path)
            if m:
                return m.group(1)
            m = re.search(r"youtu\.be/([a-zA-Z0-9_-]{11})", p.netloc + p.path)
            if m:
                return m.group(1)
    except Exception:
        pass
    # Fallback regex grezza
    patterns = [
        r"(?:v=|/v/|&v=|watch\?v=)([a-zA-Z0-9_-]{11})",
        r"youtu\.be/([a-zA-Z0-9_-]{11})",
        r"/embed/([a-zA-Z0-9_-]{11})",
        r"/shorts/([a-zA-Z0-9_-]{11})",
    ]
    for p in patterns:
        m = re.search(p, url)
        if m:
            return m.group(1)
    return None

def _with_retry(fn, *args, max_attempts=4, **kwargs):
    """Retry helper con backoff esponenziale (0.5, 1, 2 secondi) per errori temporanei."""
    delay = 0.5
    for attempt in range(max_attempts):
        try:
            return fn(*args, **kwargs)
        except TranscriptsDisabled:
            raise
        except NoTranscriptFound:
            # errore 'strutturale': lascia gestire al chiamante
            raise
        except Exception as e:
            if attempt == max_attempts - 1:
                raise
            time.sleep(delay)
            delay = min(2.0, delay * 2)

def fetch_transcript_compat(video_id: str, language_code: str, allow_fallback: bool = True):
    """
    Percorso compatibile con versioni senza list_transcripts:
    tenta get_transcript con una lista di lingue ordinate per priorità.
    """
    base_lang = language_code.split("-")[0].lower() if language_code else "en"
    prefer_list = []
    if language_code:
        prefer_list.append(language_code)
    if base_lang and base_lang not in prefer_list:
        prefer_list.append(base_lang)

    common_langs = ["en", "en-US", "en-GB", "it", "es", "de", "fr", "pt", "pt-BR"]
    for l in common_langs:
        if l not in prefer_list:
            prefer_list.append(l)

    # Tenta in ordine: appena trova una trascrizione valida, la ritorna
    for lang in prefer_list:
        try:
            return _with_retry(YTA.get_transcript, video_id, languages=[lang])
        except NoTranscriptFound:
            if not allow_fallback and lang != language_code:
                # se non è consentito fallback, fermati dopo il primo tentativo
                break
            continue
    raise NoTranscriptFound(f"Nessuna trascrizione recuperabile per video {video_id} (compat).")

def fetch_transcript_modern(video_id: str, language_code: str, allow_fallback: bool = True):
    """
    Percorso moderno (con list_transcripts disponibile):
    1) get_transcript con lingua richiesta e base
    2) list_transcripts: match lingua -> translate (se possibile) -> prima disponibile
    """
    base_lang = language_code.split("-")[0].lower() if language_code else "en"
    prefer_list = []
    if language_code:
        prefer_list.append(language_code)
    if base_lang and base_lang not in prefer_list:
        prefer_list.append(base_lang)

    # 1) Tentativo diretto
    try:
        return _with_retry(YTA.get_transcript, video_id, languages=prefer_list)
    except NoTranscriptFound:
        pass

    # 2) list_transcripts path
    transcripts = _with_retry(YTA.list_transcripts, video_id)

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

def fetch_transcript(video_id: str, language_code: str, allow_fallback: bool = True):
    """Se list_transcripts esiste usa il percorso moderno, altrimenti quello compat."""
    if hasattr(YTA, "list_transcripts"):
        return fetch_transcript_modern(video_id, language_code, allow_fallback)
    return fetch_transcript_compat(video_id, language_code, allow_fallback)

# ============================
# Helpers: text & summary
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

    # Multi-chunk: riassumi i pezzi e poi sintetizza
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

st.caption(f"📦 youtube-transcript-api version: {YTA_VERSION} — capability(list_transcripts) = {hasattr(YTA, 'list_transcripts')}")
if not hasattr(YTA, "list_transcripts"):
    st.info("ℹ️ Modalità compat: uso di get_transcript con lista di lingue. Per risultati migliori: `pip install --upgrade youtube-transcript-api`.")

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
        # Transcript (usa percorso moderno o compat a seconda della versione)
        try:
            transcript = fetch_transcript(video_id, language_code, allow_fallback=allow_fallback)
        except NoTranscriptFound:
            if allow_fallback:
                st.warning(f"⚠️ No transcript found in '{language_code}'. Trying other languages...")
                transcript = fetch_transcript(video_id, language_code, allow_fallback=True)
            else:
                raise
        except TranscriptsDisabled:
            st.error("❌ Transcripts are disabled for this video.")
            st.stop()

        if not transcript:
            st.error("❌ No transcript available for this video.")
            st.stop()

        full_text = " ".join([entry.get('text', '') for entry in transcript])
        full_text = clean_transcript_text(full_text)

        if not full_text.strip():
            st.error("❌ Transcript retrieved but empty. The video may not contain usable captions.")
            st.stop()

        with st.spinner("🧠 Generating summary with OpenAI..."):
            summary = summarize_text(client, full_text, model=model, temperature=temperature, min_words=int(min_words))

        # Files (TXT + DOCX in ZIP)
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

    except NoTranscriptFound:
        st.error(f"❌ No transcript found for this video (requested language: '{language_code}').")
    except Exception as e:
        st.error(f"❌ Error during extraction: {type(e).__name__}: {str(e)}")