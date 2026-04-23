import streamlit as st
import requests
import anthropic
import openai
import base64
import io
import re
from bs4 import BeautifulSoup
from PIL import Image

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Content Alchemist",
    page_icon="✍️",
    layout="wide",
)

st.title("✍️ Content Alchemist")
st.caption("Estrai, riscrivi e visualizza — powered by AI")

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configurazione")

    # --- Testo ---
    st.subheader("Modello testo")
    text_provider = st.selectbox(
        "Provider testo",
        ["OpenAI", "Anthropic Claude"],
        label_visibility="collapsed",
    )

    if text_provider == "OpenAI":
        openai_key = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")
        anthropic_key = ""
        text_model = st.selectbox("Modello testo", ["gpt-4o", "gpt-4o-mini", "gpt-4-turbo"])
    else:
        anthropic_key = st.text_input("Anthropic API Key", type="password", placeholder="sk-ant-...")
        openai_key = ""
        text_model = st.selectbox(
            "Modello testo",
            ["claude-opus-4-5", "claude-sonnet-4-5", "claude-haiku-4-5-20251001"],
        )

    st.divider()

    # --- Immagine ---
    st.subheader("Generazione immagine (DALL-E)")
    dalle_key = st.text_input(
        "OpenAI API Key per DALL-E",
        type="password",
        placeholder="sk-...",
        help="Se usi già OpenAI per il testo inserisci la stessa chiave",
    )
    image_model = st.selectbox(
        "Modello",
        ["dall-e-3", "dall-e-2"],
        help="DALL-E 3 = qualità superiore | DALL-E 2 = più veloce",
    )
    img_size = st.selectbox(
        "Dimensione",
        ["1024x1024", "1792x1024", "1024x1792"],
        help="1792x1024 = landscape | 1024x1792 = portrait",
    )
    img_quality = st.selectbox(
        "Qualità",
        ["hd", "standard"],
        help="hd = più dettaglio (costo doppio)",
    )

    st.divider()

    # --- Opzioni ---
    st.subheader("Opzioni pipeline")
    do_text  = st.checkbox("Riscrivi testo", value=True)
    do_image = st.checkbox("Genera immagine", value=True)

# ── Main inputs ───────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.subheader("Input")
    url = st.text_input(
        "URL da estrarre",
        placeholder="https://esempio.com/articolo",
    )
    text_prompt = st.text_area(
        "Prompt di riscrittura",
        placeholder=(
            "Es: Riscrivi come articolo di blog in tono informale per un pubblico "
            "italiano 25-40 anni. Mantieni i dati principali ma rendi il testo "
            "coinvolgente e scorrevole..."
        ),
        height=120,
    )
    img_prompt_custom = st.text_area(
        "Prompt immagine (opzionale — se vuoto viene generato automaticamente)",
        placeholder=(
            "Es: Photorealistic close-up of a modern laptop on a wooden desk, "
            "natural light, shallow depth of field, 8K..."
        ),
        height=80,
    )
    run = st.button("Avvia pipeline", type="primary", use_container_width=True)


# ── Helpers ───────────────────────────────────────────────────────────────────

def extract_content(page_url: str) -> str:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    }
    resp = requests.get(page_url, headers=headers, timeout=15)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    for tag in soup(["script", "style", "nav", "footer", "header", "aside", "form"]):
        tag.decompose()
    article = soup.find("article") or soup.find("main") or soup.body
    if not article:
        raise ValueError("Impossibile trovare contenuto nella pagina.")
    text = article.get_text(separator=" ", strip=True)
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) < 100:
        raise ValueError("Contenuto troppo breve — la pagina potrebbe bloccare il crawling.")
    return text[:6000]


def rewrite_openai(content: str, prompt: str, model: str, api_key: str) -> str:
    client = openai.OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=model,
        messages=[
            {
                "role": "system",
                "content": "Sei un content writer esperto. Riscrivi il testo seguendo le istruzioni dell'utente.",
            },
            {
                "role": "user",
                "content": f"Istruzioni: {prompt}\n\n---\nTesto originale:\n{content}",
            },
        ],
        max_tokens=2000,
    )
    return response.choices[0].message.content


def rewrite_anthropic(content: str, prompt: str, model: str, api_key: str) -> str:
    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model=model,
        max_tokens=2000,
        system="Sei un content writer esperto. Riscrivi il testo seguendo le istruzioni dell'utente.",
        messages=[
            {
                "role": "user",
                "content": f"Istruzioni: {prompt}\n\n---\nTesto originale:\n{content}",
            }
        ],
    )
    return message.content[0].text


def auto_image_prompt(text: str, text_provider: str, openai_key: str, anthropic_key: str) -> str:
    system_msg = (
        "Sei un esperto di prompt engineering per generatori di immagini AI. "
        "Basandoti sul testo fornito, crea un prompt in inglese ottimizzato per DALL-E 3 "
        "che produca un'immagine fotorealistica e non artificiale. "
        "Includi: soggetto principale, ambientazione, illuminazione, stile fotografico, "
        "qualità tecnica (es: 8K, shallow depth of field, natural light). "
        "Max 120 parole. Solo il prompt, niente altro."
    )
    if text_provider == "OpenAI":
        client = openai.OpenAI(api_key=openai_key)
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": system_msg},
                {"role": "user",   "content": text[:1500]},
            ],
            max_tokens=150,
        )
        return resp.choices[0].message.content
    else:
        client = anthropic.Anthropic(api_key=anthropic_key)
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=150,
            system=system_msg,
            messages=[{"role": "user", "content": text[:1500]}],
        )
        return msg.content[0].text


def generate_image_dalle(
    prompt: str,
    model: str,
    api_key: str,
    size: str,
    quality: str,
) -> tuple:
    """Genera immagine con DALL-E 3 e restituisce (PIL Image, revised_prompt)."""
    client = openai.OpenAI(api_key=api_key)
    kwargs = dict(
        model=model,
        prompt=prompt,
        n=1,
        size=size,
        response_format="b64_json",
    )
    if model == "dall-e-3":
        kwargs["quality"] = quality

    response = client.images.generate(**kwargs)
    img_data  = response.data[0]
    img_bytes = base64.b64decode(img_data.b64_json)
    revised   = getattr(img_data, "revised_prompt", prompt)
    img = Image.open(io.BytesIO(img_bytes))

    # Upscale a minimo 1200px di larghezza mantenendo le proporzioni
    min_width = 1200
    if img.width < min_width:
        ratio = min_width / img.width
        new_h = int(img.height * ratio)
        img   = img.resize((min_width, new_h), Image.LANCZOS)

    return img, revised


# ── Pipeline ──────────────────────────────────────────────────────────────────
with col2:
    st.subheader("Output")

    if run:
        # Validazione
        errors = []
        if not url:
            errors.append("Inserisci un URL.")
        if do_text and not text_prompt:
            errors.append("Inserisci un prompt di riscrittura.")
        if do_text and text_provider == "OpenAI" and not openai_key:
            errors.append("Inserisci la chiave OpenAI per il testo.")
        if do_text and text_provider == "Anthropic Claude" and not anthropic_key:
            errors.append("Inserisci la chiave Anthropic.")
        if do_image and not dalle_key:
            errors.append("Inserisci la chiave OpenAI per DALL-E.")

        if errors:
            for e in errors:
                st.error(e)
            st.stop()

        rewritten_text = ""

        # Step 1 — Estrazione
        with st.status("Estrazione contenuto dalla pagina...", expanded=True) as status:
            try:
                content = extract_content(url)
                st.write(f"Estratti ~{len(content)} caratteri")
                status.update(label="Contenuto estratto", state="complete")
            except Exception as e:
                st.error(f"Errore estrazione: {e}")
                st.stop()

        # Step 2 — Riscrittura
        if do_text:
            with st.status("Riscrittura testo in corso...", expanded=True) as status:
                try:
                    if text_provider == "OpenAI":
                        rewritten_text = rewrite_openai(content, text_prompt, text_model, openai_key)
                    else:
                        rewritten_text = rewrite_anthropic(content, text_prompt, text_model, anthropic_key)
                    status.update(label=f"Testo riscritto con {text_model}", state="complete")
                except Exception as e:
                    st.error(f"Errore riscrittura: {e}")
                    st.stop()

            with st.expander("Testo riscritto", expanded=True):
                st.write(rewritten_text)
                st.download_button(
                    "Scarica testo",
                    data=rewritten_text,
                    file_name="testo_riscritto.txt",
                    mime="text/plain",
                )

        # Step 3 — Immagine
        if do_image:
            with st.status(f"Generazione immagine con {image_model}...", expanded=True) as status:
                try:
                    # Determina prompt immagine
                    if img_prompt_custom.strip():
                        final_img_prompt = img_prompt_custom.strip()
                        st.write("Usando prompt personalizzato")
                    elif rewritten_text:
                        st.write("Generazione automatica del prompt immagine...")
                        final_img_prompt = auto_image_prompt(
                            rewritten_text, text_provider, openai_key, anthropic_key
                        )
                        st.write(f"Prompt generato: *{final_img_prompt}*")
                    else:
                        final_img_prompt = (
                            "Photorealistic editorial photograph for a digital article, "
                            "professional studio lighting, sharp focus, 8K resolution, "
                            "natural colors, no text overlays."
                        )

                    img, revised_prompt = generate_image_dalle(
                        final_img_prompt, image_model, dalle_key, img_size, img_quality
                    )
                    status.update(label=f"Immagine generata con {image_model}", state="complete")

                except Exception as e:
                    st.error(f"Errore generazione immagine: {e}")
                    st.stop()

            st.image(img, caption=f"{image_model} • {img_size} • {img_quality}", use_container_width=True)

            if revised_prompt != final_img_prompt:
                with st.expander("Prompt rivisto da DALL-E"):
                    st.caption(revised_prompt)

            buf = io.BytesIO()
            img.save(buf, format="PNG")
            st.download_button(
                "Scarica immagine (PNG)",
                data=buf.getvalue(),
                file_name="immagine_generata.png",
                mime="image/png",
            )

        st.success("Pipeline completata!")

    else:
        st.info("Configura le API key nella sidebar e avvia la pipeline.")
        with st.expander("Come ottenere le chiavi API"):
            st.markdown("""
**OpenAI (testo + DALL-E):** [platform.openai.com/api-keys](https://platform.openai.com/api-keys)

**Anthropic:** [console.anthropic.com/keys](https://console.anthropic.com/keys)

**Prezzi DALL-E 3:**
- 1024×1024 standard: $0.040/img
- 1024×1024 HD: $0.080/img
- 1792×1024 / 1024×1792 standard: $0.080/img
- 1792×1024 / 1024×1792 HD: $0.120/img
""")
