import streamlit as st
from openai import OpenAI
from youtube_transcript_api import YouTubeTranscriptApi as YTA, TranscriptsDisabled, NoTranscriptFound
import re
from urllib.parse import urlparse, parse_qs
from docx import Document
import io
import zipfile
from datetime import datetime

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
            m_shorts = re.search(r"/shorts/([a-zA-Z0-9_-]{11})", p.path)
            if m_shorts:
                return m_shorts.group(1)
            m_embed = re.search(r"/embed/([a-zA-Z0-9_-]{11})", p.path)
            if m_embed:
                return m_embed.group(1)
            m_be = re.search(r"youtu\.be/([a-zA-Z0-9_-]{11})", p.netloc + p.path)
            if m_be:
                return m_be.group(1)
    except Exception:
        pass
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


def fetch_transcript(video_id: str, language_code: str, allow_fallback: bool = True):
    """
    Strategia robusta:
    1) get_transcript con lingua richiesta e base (es. en-US -> en)
    2) list_transcripts:
       - match lingua
       - se consentito, translate verso la lingua richiesta/base
       - altrimenti prima disponibile
    """
    base_lang = language_code.split("-")[0].lower() if language_code else "en"
    prefer_list = []
    if language_code:
        prefer_list.append(language_code)
    if base_lang and base_lang not in prefer_list:
        prefer_list.append(base_lang)

    try:
        return YTA.get_transcript(video_id, languages=prefer_list)
    except TranscriptsDisabled:
        raise
    except NoTranscriptFound:
        pass
    except Exception:
        pass

    transcripts = YTA.list_transcripts(video_id)

    try:
        return transcripts.find_transcript(prefer_list).fetch()
    except NoTranscriptFound:
        if not allow_fallback:
            raise

    if allow_fallback:
        for tr in transcripts:
            if getattr(tr, "is_translatable", False):
                for tgt in prefer_list:
                    try:
                        return tr.translate(tgt).fetch()
                    except Exception:
                        continue

    for tr in transcripts:
        try:
            return tr.fetch()
        except Exception:
            continue

    raise NoTranscriptFound(f"Nessuna trascrizione recuperabile per video {video_id}.")


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

# --- Sidebar: OpenAI settings ---
with st.sidebar:
    st.subheader("🔐 OpenAI Settings")
    if 'oai_api_key' not in st.session_state:
        st.session_state['oai_api_key'] = ""

    oai_key = st.text_input("Insert your OpenAI API KEY", type="password", value=st.session_state['oai_api_key'])
    st.session_state['oai_api_key'] = oai_key

    # >>> MODELLI AGGIORNATI (incluso GPT-5) <<<
    model = st.selectbox(
        "Model",
        options=[
            "gpt-5",            # flagship
            "gpt-5-chat-latest",
            "gpt-5-mini",
            "gpt-4o",
            "gpt-4o-mini",
            "gpt-4.1",
            "gpt-4.1-mini",
            "gpt-4",
        ],
        index=0,
        help="Scegli un modello compatibile con l'API chat.completions. Per risultati migliori, prova 'gpt-5' o 'gpt-4o'."
    )
    temperature = st.slider("Creativity (temperature)", 0.0, 1.0, 0.7, 0.1)
    min_words = st.number_input("Minimum summary length (words)", min_value=150, max_value=2000, value=400, step=50)

# --- Main inputs ---
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

# --- Action button ---
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
        # 1) Transcript
        try:
            transcript = fetch_transcript(video_id, language_code, allow_fallback=allow_fallback)
        except NoTranscriptFound:
            if allow_fallback:
                st.warning(f"⚠️ No transcript found in '{language_code}'. Trying other available languages...")
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

        # 2) Summary
        with st.spinner("🧠 Generating summary with OpenAI..."):
            summary = summarize_text(client, full_text, model=model, temperature=temperature, min_words=int(min_words))

        # 3) Files (TXT + DOCX in ZIP)
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
        st.error(f"❌ No transcript found for this video (requested language: '{language_code}'). Try enabling fallback or choosing a different language.")
    except Exception as e:
        st.error(f"❌ Error during extraction: {type(e).__name__}: {str(e)}")