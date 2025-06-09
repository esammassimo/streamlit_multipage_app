import streamlit as st
import pandas as pd
import numpy as np
from sklearn.cluster import KMeans
from sentence_transformers import SentenceTransformer
import openai
import os

# Titolo app
st.title("🔠 Semantic Keyword Clustering")

# Upload file CSV
uploaded_file = st.file_uploader("Carica un file CSV con una colonna 'keyword'", type="csv")

if uploaded_file:
    try:
        df = pd.read_csv(uploaded_file)
    except UnicodeDecodeError:
        df = pd.read_csv(uploaded_file, encoding='latin1')

    if 'keyword' not in df.columns:
        st.error("Il file deve contenere una colonna chiamata 'keyword'")
    else:
        keywords = df['keyword'].dropna().unique().tolist()
        st.success(f"Sono state caricate {len(keywords)} keyword uniche.")

        # Selezione numero cluster
        n_clusters = st.slider("Seleziona il numero di cluster", min_value=2, max_value=15, value=5)

        # Inizializzazione modello
        model = SentenceTransformer('all-MiniLM-L6-v2')
        embeddings = model.encode(keywords)

        # Clustering
        kmeans = KMeans(n_clusters=n_clusters, random_state=42)
        labels = kmeans.fit_predict(embeddings)

        # Calcolo keyword rappresentativa per cluster
        cluster_labels = []
        for i in range(n_clusters):
            cluster_indices = np.where(labels == i)[0]
            cluster_embeddings = [embeddings[j] for j in cluster_indices]
            center = np.mean(cluster_embeddings, axis=0)
            distances = [np.linalg.norm(embeddings[j] - center) for j in cluster_indices]
            closest_idx = cluster_indices[np.argmin(distances)]
            cluster_labels.append(keywords[closest_idx])

        # Mappa dei cluster
        label_map = {i: f"{i} - {cluster_labels[i]}" for i in range(n_clusters)}

        # Opzionale: Etichettatura AI con OpenAI
        if st.checkbox("🧠 Usa OpenAI per etichettare i cluster"):
            openai_api_key = st.text_input("Inserisci la tua API Key OpenAI", type="password")
            if openai_api_key:
                openai.api_key = openai_api_key
                ai_labels = []
                for i in range(n_clusters):
                    kws = [keywords[j] for j in np.where(labels == i)[0]]
                    prompt = f"Qual è l'argomento principale comune alle seguenti keyword? {kws}"
                    try:
                        response = openai.chat.completions.create(
                            model="gpt-3.5-turbo",
                            messages=[
                                {"role": "user", "content": prompt}
                            ],
                            max_tokens=20
                        )
                        label_map[i] = response.choices[0].message.content.strip()
                    except Exception as e:
                        st.warning(f"Errore durante la generazione dell'etichetta AI per il cluster {i}: {e}")

        # Output finale
        df_result = pd.DataFrame({
            'keyword': keywords,
            'cluster_label': [label_map[label] for label in labels]
        })

        st.subheader("📊 Risultato del Clustering")
        st.dataframe(df_result)

        csv = df_result.to_csv(index=False).encode('utf-8')
        st.download_button("💾 Scarica CSV con i cluster", data=csv, file_name="clustered_keywords.csv", mime="text/csv")