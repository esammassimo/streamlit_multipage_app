import streamlit as st
import openai
from youtube_transcript_api import YouTubeTranscriptApi, TranscriptsDisabled, NoTranscriptFound
import re
from docx import Document
import os
import zipfile

def extract_video_id(url):
    match = re.search(r"(?:v=|youtu\.be/|embed/|/v/|/e/|watch\?v=|&v=)([a-zA-Z0-9_-]{11})", url)
    return match.group(1) if match else None

st.title("🎥 YouTube Transcript & Summary Generator")

# Input API Key
oai_key = st.text_input("🔑 Enter your OpenAI API Key", type="password")

# Input YouTube video URL and language
youtube_url = st.text_input("📺 Enter YouTube video URL")
language_code = st.text_input("🌍 Enter language code (default 'en')", "en")

if st.button("📄 Generate Transcript and Summary"):
    if not oai_key or not youtube_url:
        st.error("⚠️ You must enter both the API Key and the video URL!")
    else:
        video_id = extract_video_id(youtube_url)
        if not video_id:
            st.error("❌ Invalid YouTube URL. Please try again.")
        else:
            try:
                # Extract transcript with fallback
                try:
                    transcript = YouTubeTranscriptApi.get_transcript(video_id, languages=[language_code])
                except NoTranscriptFound:
                    st.warning(f"⚠️ No transcript found in '{language_code}', trying other available languages...")
                    transcript = YouTubeTranscriptApi.get_transcript(video_id)  # Tries default languages
                except TranscriptsDisabled:
                    st.error("❌ Transcripts are disabled for this video.")
                    transcript = None
                
                if transcript:
                    full_text = " ".join([entry['text'] for entry in transcript])
                    
                    # Save transcript as a .txt file
                    transcript_file = f"transcript_{video_id}.txt"
                    with open(transcript_file, "w", encoding="utf-8") as txt_file:
                        txt_file.write(full_text)
                    
                    # Generate summary with OpenAI
                    client = openai.OpenAI(api_key=oai_key)
                    prompt = f"""
                    Riassumi il seguente testo, mantenendo i punti principali e il contesto chiave.
                    Assicurati che il riassunto abbia almeno 400 parole.
                    
                    {full_text}
                    
                    **Regole:**
                    - Mantieni il riassunto chiaro e conciso.
                    - Evita ripetizioni e dettagli non necessari.
                    - Evidenzia gli aspetti più importanti della discussione.
                    """
                    response = client.chat.completions.create(
                        model="gpt-4",
                        messages=[{"role": "system", "content": "Sei un esperto nel sintetizzare contenuti di video."},
                                  {"role": "user", "content": prompt}],
                        temperature=0.7,
                        max_tokens=4000  # Increased tokens for more complete summaries
                    )
                    summary = response.choices[0].message.content
                    
                    # Save summary as a Word document
                    summary_file = f"summary_{video_id}.docx"
                    doc = Document()
                    doc.add_heading("Riassunto del Video", level=1)
                    doc.add_paragraph(summary)
                    doc.save(summary_file)
                    
                    # Create a ZIP file with both transcript and summary
                    zip_file = f"youtube_summary_{video_id}.zip"
                    with zipfile.ZipFile(zip_file, "w") as zipf:
                        zipf.write(transcript_file)
                        zipf.write(summary_file)
                    
                    st.success("✅ Transcript and Summary successfully generated!")
                    with open(zip_file, "rb") as zip_data:
                        st.download_button("📥 Download ZIP with Transcript & Summary", data=zip_data, file_name=zip_file, mime="application/zip")
                    
                    # Clean up temporary files
                    os.remove(transcript_file)
                    os.remove(summary_file)
                    os.remove(zip_file)
                else:
                    st.error("❌ No transcript available for this video.")
                
            except Exception as e:
                st.error(f"❌ Error during extraction: {str(e)}")