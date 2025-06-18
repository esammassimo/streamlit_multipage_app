import streamlit as st
import pandas as pd
import openai
from io import BytesIO

st.set_page_config(page_title="Keyword Clustering", layout="centered")
st.title("🔗 Keyword Clustering con OpenAI")

# Inserimento API key
api_key = st.text_input("🔐 Inserisci la tua OpenAI API Key", type="password")

# Metodo di inserimento
input_method = st.radio("📥 Scegli il metodo di inserimento delle keyword:", ["Carica file Excel", "Incolla testo manualmente"])

keywords = []

if input_method == "Carica file Excel":
    uploaded_file = st.file_uploader("📄 Carica un file Excel con le parole chiave (colonna A)", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        if df.empty or df.shape[1] < 1:
            st.error("Il file deve contenere almeno una colonna con le keyword.")
        else:
            keywords = df.iloc[:, 0].dropna().astype(str).tolist()

elif input_method == "Incolla testo manualmente":
    manual_input = st.text_area("✍️ Incolla le parole chiave, una per riga")
    if manual_input:
        keywords = [line.strip() for line in manual_input.splitlines() if line.strip()]

# Etichette
labels_input = st.text_input("🏷️ Inserisci le etichette di clustering separate da virgola",
                              placeholder="es. birra artigianale, ricette, marchi")

if keywords and labels_input and api_key:
    labels = [label.strip() for label in labels_input.split(",") if label.strip()]

    results = []
    openai.api_key = api_key

    with st.spinner("🔍 Sto classificando le parole chiave..."):
        for keyword in keywords:
            try:
                response = openai.ChatCompletion.create(
                    model="gpt-4o",
                    messages=[
                        {"role": "system", "content": "Assegna ogni parola chiave alla categoria più pertinente tra quelle fornite."},
                        {"role": "user", "content": f"Assegna la parola chiave '{keyword}' a una delle seguenti categorie: {', '.join(labels)}. Rispondi solo con il nome della categoria più pertinente."}
                    ],
                    max_tokens=10
                )
                category = response.choices[0].message.content.strip()
            except Exception as e:
                category = f"Errore: {e}"
            results.append((keyword, category))

    output_df = pd.DataFrame(results, columns=["Keyword", "Cluster"])
    st.success("✅ Classificazione completata!")
    st.dataframe(output_df)

    # Esporta file Excel
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        output_df.to_excel(writer, index=False)
    st.download_button("⬇️ Scarica il file con i cluster", data=buffer.getvalue(), file_name="keyword_clustering.xlsx")