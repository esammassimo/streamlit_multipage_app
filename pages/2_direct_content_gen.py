import streamlit as st
import openai
import os
import tempfile
from docx import Document

# 📌 Configurazione pagina
st.set_page_config(page_title="AI Content Generator", layout="wide")
st.title("📝 AI Content Generator - Articles & Recipes")

# 📌 Inserimento API Key OpenAI
def get_openai_api_key():
    with st.sidebar:
        st.subheader("🔐 API Key OpenAI")
        api_key = st.text_input("Insert your OpenAI API KEY", type="password")
        if not api_key:
            st.warning("⚠️ Insert your API KEY to continue.")
            st.stop()
        return api_key

openai.api_key = get_openai_api_key()

# 📌 Funzione generica per generazione contenuto
def generate_openai_content(system_prompt, user_prompt, model="gpt-4", temperature=0.7, max_tokens=1500):
    response = openai.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        temperature=temperature,
        max_tokens=max_tokens
    )
    return response.choices[0].message.content.strip()

# 📌 Creazione Tabs
tab1, tab2 = st.tabs(["📝 Articles", "🍽 Recipes"])

# 📝 TAB ARTICOLI
with tab1:
    st.header("📄 Generate Articles")

    with st.form("article_form"):
        title = st.text_input("Title of the article")
        seo_title = st.text_input("SEO Title")
        meta_description = st.text_area("Meta Description")
        min_words = st.number_input("Minimum number of words:", min_value=100, value=800, step=100)
        tone_of_voice = st.selectbox("Tone of Voice", ["Informale", "Professionale", "Formale"])
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
        if not title or not seo_title or not meta_description or not paragraphs:
            st.error("❌ Please fill out all required fields.")
        else:
            st.success(f"📝 Generating article: {title}...")

            generated_sections = []
            for p_title, p_desc in paragraphs:
                prompt = f"""
                Scrivi un paragrafo con il titolo "{p_title}".
                Il tono deve essere {tone_of_voice.lower()}.
                Il contenuto deve sviluppare questa idea: "{p_desc}".
                Deve contenere almeno {min_words // len(paragraphs)} parole.
                La fonte seguente può servire da contesto (non citarla): {source_text}
                """
                content = generate_openai_content(
                    "Sei un esperto redattore SEO.", prompt,
                    temperature=0.7, max_tokens=1800
                )
                generated_sections.append((p_title, content))

            # 📄 Anteprima del contenuto
            st.subheader("📖 Anteprima Articolo")
            st.markdown(f"# {title}")
            st.markdown(f"**SEO Title**: {seo_title}\n\n**Meta Description**: {meta_description}")
            for p_title, p_text in generated_sections:
                st.markdown(f"## {p_title}\n\n{p_text}")

            # 📁 Esporta in Word
            temp_dir = tempfile.mkdtemp()
            file_path = os.path.join(temp_dir, f"{title.replace(' ', '_')}.docx")

            doc = Document()
            doc.add_heading(title, level=1)
            doc.add_paragraph(f"SEO Title: {seo_title}")
            doc.add_paragraph(f"Meta Description: {meta_description}")
            doc.add_paragraph(f"Tone of Voice: {tone_of_voice}")

            for p_title, p_text in generated_sections:
                doc.add_heading(p_title, level=2)
                doc.add_paragraph(p_text)

            doc.save(file_path)
            with open(file_path, "rb") as f:
                st.download_button("📥 Download Article (.docx)", f, file_name=os.path.basename(file_path))

# 🍽 TAB RICETTE
with tab2:
    st.header("🍽 Generate Recipes")

    with st.form("recipe_form"):
        recipe_title = st.text_input("Recipe Title")
        min_words_recipe = st.number_input("Minimum number of words:", min_value=100, value=500, step=50)
        recipe_tone = st.selectbox("Tone of Voice", ["Informale", "Professionale", "Formale"], key="recipe_tone")
        recipe_source = st.text_area("Fonte autorevole (non verrà citata direttamente)", key="recipe_source")
        ingredients = st.text_area("Insert ingredients (one per line)", key="ingredients")
        preparation_desc = st.text_area("Describe the preparation process", key="preparation_desc")

        submit_recipe = st.form_submit_button("Generate Recipe")

    if submit_recipe:
        if not recipe_title or not ingredients or not preparation_desc:
            st.error("❌ Please fill out all required fields.")
        else:
            st.success(f"🍽 Generating recipe: {recipe_title}...")

            intro_prompt = f"""
            Scrivi una breve introduzione per la ricetta "{recipe_title}".
            Il tono deve essere {recipe_tone.lower()}.
            La fonte è solo per contesto (non citarla): {recipe_source}
            """

            prep_prompt = f"""
            Scrivi un paragrafo dettagliato per spiegare come preparare "{recipe_title}".
            Deve contenere almeno {min_words_recipe} parole.
            Gli ingredienti sono: {ingredients}
            Il tono deve essere {recipe_tone.lower()}.
            {preparation_desc}
            """

            recipe_intro = generate_openai_content("Sei un esperto di cucina.", intro_prompt)
            recipe_prep = generate_openai_content("Sei un esperto di cucina.", prep_prompt, max_tokens=2500)

            # 📄 Anteprima
            st.subheader("📖 Anteprima Ricetta")
            st.markdown(f"# {recipe_title}")
            st.markdown(recipe_intro)
            st.markdown("## Ingredienti")
            st.markdown("\n".join([f"- {i}" for i in ingredients.strip().split("\n") if i.strip()]))
            st.markdown("## Preparazione")
            st.markdown(recipe_prep)

            # 📁 Word export
            temp_dir = tempfile.mkdtemp()
            file_path = os.path.join(temp_dir, f"{recipe_title.replace(' ', '_')}.docx")

            doc = Document()
            doc.add_heading(recipe_title, level=1)
            doc.add_paragraph(f"Tone of Voice: {recipe_tone}")
            doc.add_paragraph(recipe_intro)
            doc.add_heading("Ingredienti", level=2)
            for ing in ingredients.strip().split("\n"):
                if ing.strip():
                    doc.add_paragraph(f"• {ing.strip()}")
            doc.add_heading("Preparazione", level=2)
            doc.add_paragraph(recipe_prep)

            doc.save(file_path)
            with open(file_path, "rb") as f:
                st.download_button("📥 Download Recipe (.docx)", f, file_name=os.path.basename(file_path))