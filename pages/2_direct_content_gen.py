import streamlit as st
import openai
import os
import tempfile
from docx import Document

# 📌 Configurazione API OpenAI
st.set_page_config(page_title="AI Content Generator", layout="wide")
st.title("📝 AI Content Generator - Articles & Recipes")

# 📌 Inserimento della chiave API di OpenAI
def get_openai_api_key():
    with st.sidebar:
        st.subheader("🔐 API Key OpenAI")
        if 'oai_api_key' not in st.session_state:
            st.session_state['oai_api_key'] = ""

        oai_api_key = st.text_input(
            "Insert your OpenAI API KEY",
            type="password",
            value=st.session_state['oai_api_key']
        )

        if not oai_api_key:
            st.warning("⚠️ Insert your API KEY to continue.")
            st.stop()

        # Salva in sessione
        st.session_state['oai_api_key'] = oai_api_key
        return oai_api_key
    oai_client = openai.Client(api_key=oai_api_key)

# 📌 Creazione delle due Tab
tab1, tab2 = st.tabs(["📝 Articles", "🍽 Recipes"])

# 📌 📝 TAB 1 - Generazione Articoli
with tab1:
    st.header("📄 Generate Articles")

    with st.form("article_form"):
        title = st.text_input("Title of the content")
        seo_title = st.text_input("SEO Title")
        meta_description = st.text_area("Meta Description")
        min_words = st.number_input("Minimum number of words:", min_value=100, value=800, step=100)

        # 📌 Tone of Voice
        tone_of_voice = st.selectbox("Tone of Voice", ["Informale", "Professionale", "Formale"])

        # 📌 Fonte autorevole (non citata nel testo)
        source_text = st.text_area("Fonte autorevole (non verrà citata direttamente)")

        st.subheader("📝 Paragraphs (max 5)")
        paragraphs = []
        for i in range(1, 6):
            col1, col2 = st.columns([1, 2])
            with col1:
                p_title = st.text_input(f"Paragrafo {i} - Titolo", key=f"title_{i}")
            with col2:
                p_desc = st.text_area(f"Paragrafo {i} - Descrizione", key=f"desc_{i}")

            if p_title and p_desc:
                paragraphs.append((p_title, p_desc))

        submit_article = st.form_submit_button("Generate Article")

    if submit_article:
        if not title or not seo_title or not meta_description or len(paragraphs) == 0:
            st.error("❌ Please fill out all required fields.")
        else:
            st.write(f"📝 Generating article: {title}...")

            full_text = []
            for p_title, p_desc in paragraphs:
                paragraph_text = f"## {p_title}\n\n{p_desc}"
                full_text.append(paragraph_text)

            # 📌 Creazione del file Word
            temp_dir = tempfile.mkdtemp()
            file_path = os.path.join(temp_dir, f"{title.replace(' ', '_')}.docx")

            doc = Document()
            doc.add_heading(title, level=1)
            doc.add_paragraph(f"SEO Title: {seo_title}")
            doc.add_paragraph(f"Meta Description: {meta_description}")
            doc.add_paragraph(f"Tone of Voice: {tone_of_voice}")

            for paragraph in full_text:
                doc.add_paragraph(paragraph)

            doc.save(file_path)

            with open(file_path, "rb") as docx_file:
                st.download_button("📥 Download Article (.docx)", docx_file, file_name=f"{title.replace(' ', '_')}.docx")

# 📌 🍽 TAB 2 - Generazione Ricette
with tab2:
    st.header("🍽 Generate Recipes")

    with st.form("recipe_form"):
        recipe_title = st.text_input("Recipe Title")
        min_words_recipe = st.number_input("Minimum number of words:", min_value=100, value=500, step=50)

        # 📌 Tone of Voice
        recipe_tone = st.selectbox("Tone of Voice", ["Informale", "Professionale", "Formale"], key="recipe_tone")

        # 📌 Fonte autorevole (non citata direttamente)
        recipe_source = st.text_area("Fonte autorevole (non verrà citata direttamente)", key="recipe_source")

        # 📌 Elenco ingredienti (dato in input, da non modificare)
        ingredients = st.text_area("Insert ingredients (one per line)", key="ingredients")

        # 📌 Descrizione della preparazione
        preparation_desc = st.text_area("Describe the preparation process", key="preparation_desc")

        submit_recipe = st.form_submit_button("Generate Recipe")

    if submit_recipe:
        if not recipe_title or not ingredients or not preparation_desc:
            st.error("❌ Please fill out all required fields.")
        else:
            st.write(f"🍽 Generating recipe: {recipe_title}...")

            # 📌 Generazione del testo
            intro_prompt = f"""
            Scrivi una breve introduzione per la ricetta **{recipe_title}**. 
            Il tono di voce deve essere **{recipe_tone.lower()}**. 
            La fonte autorevole è fornita solo per contesto e **non deve essere citata direttamente**.
            """

            preparation_prompt = f"""
            Scrivi un paragrafo dettagliato per la preparazione della ricetta **{recipe_title}**.
            Il tono di voce deve essere **{recipe_tone.lower()}**.
            Deve contenere almeno {min_words_recipe} parole.
            """

            response_intro = oai_client.chat.completions.create(
                model="gpt-4",
                messages=[{"role": "system", "content": "Sei un esperto di cucina."},
                          {"role": "user", "content": intro_prompt}],
                temperature=0.7,
                max_tokens=1500
            )

            response_prep = oai_client.chat.completions.create(
                model="gpt-4",
                messages=[{"role": "system", "content": "Sei un esperto di cucina."},
                          {"role": "user", "content": preparation_prompt}],
                temperature=0.7,
                max_tokens=2500
            )

            recipe_intro = response_intro.choices[0].message.content
            recipe_prep = response_prep.choices[0].message.content

            # 📌 Creazione del file Word
            temp_dir = tempfile.mkdtemp()
            file_path = os.path.join(temp_dir, f"{recipe_title.replace(' ', '_')}.docx")

            doc = Document()
            doc.add_heading(recipe_title, level=1)
            doc.add_paragraph(f"Tone of Voice: {recipe_tone}")

            doc.add_paragraph(recipe_intro)
            doc.add_heading("Ingredienti", level=2)
            for ingredient in ingredients.split("\n"):
                doc.add_paragraph(f"• {ingredient}")

            doc.add_heading("Preparazione", level=2)
            doc.add_paragraph(recipe_prep)

            doc.save(file_path)

            with open(file_path, "rb") as docx_file:
                st.download_button("📥 Download Recipe (.docx)", docx_file, file_name=f"{recipe_title.replace(' ', '_')}.docx")
