import streamlit as st
import pandas as pd
import openai
from openai import OpenAI
from io import BytesIO
import xlsxwriter

st.set_page_config(page_title="Domain Clustering", layout="centered")
st.title("🌐 Domain Clustering con OpenAI")

# Inserimento API key
api_key = st.text_input("🔐 Inserisci la tua OpenAI API Key", type="password")

# Metodo di inserimento domini
input_method = st.radio("📥 Scegli il metodo di inserimento dei domini:", ["Carica file Excel", "Incolla testo manualmente"])

domains = []

if input_method == "Carica file Excel":
    uploaded_file = st.file_uploader("📄 Carica un file Excel con i domini (colonna A)", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        if df.empty or df.shape[1] < 1:
            st.error("Il file deve contenere almeno una colonna con i domini.")
        else:
            domains = df.iloc[:, 0].dropna().astype(str).tolist()

elif input_method == "Incolla testo manualmente":
    manual_input = st.text_area("✍️ Incolla i domini, uno per riga")
    if manual_input:
        domains = [line.strip() for line in manual_input.splitlines() if line.strip()]

# Metodo di definizione etichette
label_mode = st.radio("🏷️ Come vuoi fornire le etichette?", ["Inserimento manuale", "Usa file predefinito category_domain.txt"])

labels = []

if label_mode == "Inserimento manuale":
    labels_input = st.text_input("✏️ Inserisci le etichette di clustering separate da virgola",
                                  placeholder="es. blog, ecommerce, istituzionale, news")
    if labels_input:
        labels = [label.strip() for label in labels_input.split(",") if label.strip()]

elif label_mode == "Usa file predefinito category_domain.txt":
    try:
        with open("category_domain.txt", "r", encoding="utf-8") as f:
            labels = [line.strip() for line in f.readlines() if line.strip()]
        st.markdown("### 📂 Etichette caricate da category_domain.txt:")
        st.text("".join(labels[:30]))
        if len(labels) > 30:
            st.text(f"...e altre {len(labels) - 30} categorie")
    except FileNotFoundError:
        st.error("❌ Il file 'category_domain.txt' non è stato trovato nella directory dello script.")

# Scelta del modello
model_choice = st.selectbox("🤖 Seleziona il modello OpenAI da utilizzare", ["gpt-4o", "gpt-3.5-turbo"])

# Stima preliminare dei token (approssimazione)
if domains and labels:
    st.markdown("---")
    st.subheader("📈 Stima preliminare dei token")
    avg_prompt_tokens = 15 + len(", ".join(labels).split()) + 5  # stima base
    estimated_total_prompt_tokens = avg_prompt_tokens * len(domains)
    estimated_completion_tokens = 5 * len(domains)
    estimated_total = estimated_total_prompt_tokens + estimated_completion_tokens
    st.markdown(f"🔢 **Token stimati:** Prompt: `{estimated_total_prompt_tokens}`, Completion: `{estimated_completion_tokens}`, Totale: `{estimated_total}`")
    if model_choice == "gpt-4o":
        cost = (estimated_total_prompt_tokens / 1000 * 0.005) + (estimated_completion_tokens / 1000 * 0.015)
    else:
        cost = (estimated_total_prompt_tokens / 1000 * 0.001) + (estimated_completion_tokens / 1000 * 0.002)
    st.markdown(f"💰 **Costo stimato:** ~${cost:.4f}")

# Pulsante di conferma prima dell'avvio
confirm_run = st.checkbox("✅ Conferma e avvia la classificazione")

# Avvio solo se c'è anche l'API key e la conferma
if domains and labels and api_key and confirm_run:
    results = []
    client = OpenAI(api_key=api_key)

    total_prompt_tokens = 0
    total_completion_tokens = 0

    with st.spinner("🔍 Sto classificando i domini..."):
        for domain in domains:
            try:
                response = client.chat.completions.create(
                    model=model_choice,
                    messages=[
                        {"role": "system", "content": "Assegna ogni dominio alla categoria più pertinente tra quelle fornite."},
                        {"role": "user", "content": f"Assegna il dominio '{domain}' a una delle seguenti categorie: {', '.join(labels)}. Rispondi solo con il nome della categoria più pertinente."}
                    ],
                    max_tokens=10
                )
                category = response.choices[0].message.content.strip()

                usage = response.usage
                total_prompt_tokens += usage.prompt_tokens
                total_completion_tokens += usage.completion_tokens

            except Exception as e:
                category = f"Errore: {e}"
            results.append((domain, category))

    output_df = pd.DataFrame(results, columns=["Dominio", "Cluster"])
    st.success("✅ Classificazione completata!")
    st.dataframe(output_df)

    st.markdown(f"**📊 Token utilizzati:** Prompt: `{total_prompt_tokens}`, Completion: `{total_completion_tokens}`, Totale: `{total_prompt_tokens + total_completion_tokens}`")

    # Esporta file Excel
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        output_df.to_excel(writer, index=False)
    st.download_button("⬇️ Scarica il file con i cluster", data=buffer.getvalue(), file_name="domain_clustering.xlsx")