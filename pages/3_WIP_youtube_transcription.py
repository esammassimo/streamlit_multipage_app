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
import sys

# ============= Tentativo import youtube-transcript-api + introspezione =============
YTA = None
TranscriptsDisabled = Exception
NoTranscriptFound = Exception
YTA_VERSION = "unknown"
YTA_FILE = "unknown"

try:
    yta_mod = importlib.import_module("youtube_transcript_api")
    YTA_FILE = getattr(yta_mod, "__file__", "unknown")
    YTA_VERSION = getattr(yta_mod, "__version__", "unknown")
    # In alcune versioni l'API sta su youtube_transcript_api.YouTubeTranscriptApi
    YTA = getattr(yta_mod, "YouTubeTranscriptApi", None)
    # Eccezioni (se presenti)
    TranscriptsDisabled = getattr(yta_mod, "TranscriptsDisabled", Exception)
    NoTranscriptFound = getattr(yta_mod, "NoTranscriptFound", Exception)
except Exception as e:
    st.error(f"❌ Impossibile importare 'youtube_transcript_api': {type(e).__name__}: {e}")
    st.stop()

# ============= Helper estrazione ID =============
def extract_video_id(url: str):
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

# ============= Helper retry/backoff =============
def _with_retry(fn, *args, max_attempts=4, **kwargs):
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

# ============= Fallback via yt_dlp (opzionale) =============
def fetch_transcript_via_ytdlp(video_id: str, language_code: str, allow_fallback: bool):
    """
    Prova a ottenere sottotitoli con yt_dlp senza scaricare il video.
    Restituisce lista di dict [{'text': '...', 'start': ..., 'duration': ...}] in stile youtube-transcript-api.
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
        # Priorità: sottotitoli nella lingua richiesta, poi auto-generated
        candidates = []

        def collect(subs_dict, label):
            # subs_dict: {'en': [{'ext': 'vtt', 'url': '...'}, ...], 'it': [...]}
            if not subs_dict:
                return
            # 1) esatta
            if language_code in subs_dict:
                candidates.append((subs_dict[language_code], f"{label}:{language_code}"))
            # 2) base
            base = language_code.split("-")[0].lower()
            if base in subs_dict:
                candidates.append((subs_dict[base], f"{label}:{base}"))
            # 3) qualsiasi lingua se allow_fallback
            if allow_fallback:
                for lang, entries in subs_dict.items():
                    candidates.append((entries, f"{label}:{lang}"))

        collect(info.get("subtitles"), "subs")
        collect(info.get("automatic_captions"), "auto")

        if not candidates:
            raise NoTranscriptFound("Nessun sottotitolo disponibile via yt_dlp.")

        # prendi il primo candidato e scarica il primo formato disponibile
        tracks, desc = candidates[0]
        track = tracks[0]  # tipicamente .vtt
        srt_url = track.get("url")
        if not srt_url:
            raise NoTranscriptFound("URL sottotitoli non disponibile.")

        # scarica il testo VTT e converti a lista tipo transcript_api
        r = requests.get(srt_url, timeout=20)
        r.raise_for_status()
        vtt = r.text

        # parsing VTT minimale: estrai blocchi "HH:MM:SS.mmm --> HH:MM:SS.mmm" + testo
        # (parsing semplice; per robustezza totale si potrebbe usare 'webvtt' lib)
        entries = []
        # separa per doppia newline
        for block in re.split(r"\n\s*\n", vtt.strip()):
            lines = [ln.strip() for ln in block.splitlines() if ln.strip()]
            if len(lines) < 2:
                continue
            # trova linea tempo
            time_line = None
            for ln in lines:
                if "-->" in ln:
                    time_line = ln
                    break
            if not time_line:
                continue
            # parse tempi
            tm = re.match(r"(?P<start>[\d:.]+)\s*-->\s*(?P<end>[\d:.]+)", time_line)
            if not tm:
                continue
            def to_seconds(s):
                # 00:01:02.345
                parts = s.split(":")
                parts = [float(p.replace(",", ".")) for p in parts]
                while len(parts) < 3:  # mm:ss -> 0:mm:ss
                    parts = [0.0] + parts
                h, m, sec = parts[-3], parts[-2], parts[-1]
                return h*3600 + m*60 + sec
            start = to_seconds(tm.group("start"))
            end = to_seconds(tm.group("end"))
            duration = max(0.0, end - start)
            # testo = tutto ciò che non è time_line/indice
            text_lines = []
            for ln in lines:
                if ln == time_line: 
                    continue
                # salta indici numerici
                if re.fullmatch(r"\d+", ln):
                    continue
                text_lines.append(ln)
            text = " ".join(text_lines).strip()
            if text:
                entries.append({"text": text, "start": start, "duration": duration})
        if not entries:
            raise NoTranscriptFound("Impossibile parsare i sottotitoli VTT via yt_dlp.")
        return entries

# ============= Fetch transcript – percorso adattivo =============
def fetch_transcript_any(video_id: str, language_code: str, allow_fallback: bool = True):
    """
    Prova in sequenza:
    1) youtube_transcript_api.list_transcripts (se disponibile)
    2) youtube_transcript_api.get_transcript (se disponibile)
    3) Fallback yt_dlp (se disponibile)
    """
    base_lang = language_code.split("-")[0].lower() if language_code else "en"
    prefer_list = [lc for lc in [language_code, base_lang] if lc]

    used_path = None

    # 1) Percorso moderno con list_transcripts
    if YTA and hasattr(YTA, "list_transcripts"):
        transcripts = _with_retry(getattr(YTA, "list_transcripts"), video_id)
        # a) lingua preferita
        try:
            tr = transcripts.find_transcript(prefer_list)
            used_path = f"list_transcripts:match({tr.language_code})"
            return transcripts.find_transcript(prefer_list).fetch(), used_path
        except Exception:
            pass
        # b) translate
        if allow_fallback:
            for tr in transcripts:
                try:
                    if getattr(tr, "is_translatable", False):
                        for tgt in prefer_list:
                            try:
                                used_path = f"list_transcripts:translate({tr.language_code}->{tgt})"
                                return tr.translate(tgt).fetch(), used_path
                            except Exception:
                                continue
                except Exception:
                    continue
        # c) prima disponibile
        for tr in transcripts:
            try:
                used_path = f"list_transcripts:first({tr.language_code})"
                return tr.fetch(), used_path
            except Exception:
                continue

    # 2) Percorso classico get_transcript
    if YTA and hasattr(YTA, "get_transcript"):
        # prova prefer_list, poi qualche lingua comune se allow_fallback
        langs = prefer_list[:]
        if allow_fallback:
            for extra in ["en", "en-US", "en-GB", "it", "es", "de", "fr", "pt", "pt-BR"]:
                if extra not in langs:
                    langs.append(extra)
        for lc in langs:
            try:
                used_path = f"get_transcript({lc})"
                return _with_retry(getattr(YTA, "get_transcript"), video_id, languages=[lc]), used_path
            except TranscriptsDisabled:
                raise
            except NoTranscriptFound:
                continue
            except Exception:
                continue

    # 3) Fallback yt_dlp
    try:
        entries = fetch_transcript_via_ytdlp(video_id, language_code, allow_fallback)
        used_path = "yt_dlp"
        return entries, used_path
    except Exception as e:
        raise NoTranscriptFound(f"Nessun metodo disponibile: {type(e).__name__}: {e}")

# ============= Text helpers =============
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

    partial = []
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
        partial.append(resp.choices[0].message.content)

    synthesis_prompt = (
        f"You are given {len(partial)} partial summaries from a video transcript. "
        f"Merge them into a single cohesive summary of at least {min_words} words. "
        "Avoid duplication, keep a logical flow, and highlight the most critical insights."
        "\n\nPartial summaries:\n" + "\n".join(f"Part {i+1}:\n{ps}" for i, ps in enumerate(partial))
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

# ============= UI =============
st.set_page_config(page_title="YouTube Transcript & Summary Generator", page_icon="🎥", layout="centered")
st.title("🎥 YouTube Transcript & Summary Generator")

with st.sidebar:
    st.subheader("🔐 OpenAI Settings")
    if 'oai_api_key' not in st.session_state:
        st.session_state['oai_api_key'] = ""
    oai_key = st.text_input("Insert your OpenAI API KEY", type="password", value=st.session_state['oai_api_key'])
    st.session_state['oai_api_key'] = oai_key

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

# Diagnostica modulo importato
st.caption(f"📦 youtube-transcript-api — version: {YTA_VERSION}, file: {YTA_FILE}")
st.caption(f"🔍 attributes → list_transcripts: {hasattr(YTA, 'list_transcripts') if YTA else False}, get_transcript: {hasattr(YTA, 'get_transcript') if YTA else False}")

youtube_url = st.text_input("📺 Enter YouTube video URL")
language_code = st.text_input("🌍 Enter language code (default 'en')", value="en")
allow_fallback = st.checkbox("If not found, try other available languages", value=True)

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
        entries, used_path = fetch_transcript_any(video_id, language_code, allow_fallback=allow_fallback)
        st.info(f"✅ Transcript fetched via: **{used_path}**")

        # entries → [{'text': '...', 'start': float, 'duration': float}, ...]
        full_text = " ".join([e.get("text", "") for e in entries])
        full_text = clean_transcript_text(full_text)
        if not full_text.strip():
            st.error("❌ Transcript retrieved but empty.")
            st.stop()

        with st.spinner("🧠 Generating summary with OpenAI..."):
            summary = summarize_text(client, full_text, model=model, temperature=temperature, min_words=int(min_words))

        # Build ZIP (txt + docx)
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