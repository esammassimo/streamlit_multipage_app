import streamlit as st
import requests
import openai
import json

def get_paa_questions(keyword, serpapi_api_key):
    # Endpoint e parametri per SerpAPI
    url = "https://serpapi.com/search.json"
    params = {
        "engine": "google",
        "q": keyword,
        "gl": "it",            # Paese: Italia
        "hl": "it",            # Lingua: Italiano
        # "google_domain": "google.it",  # Opzionale, dominio Google Italia
        "api_key": serpapi_api_key
    }
    # Esegui la richiesta GET a SerpAPI
    try:
        response = requests.get(url, params=params, timeout=30)
    except requests.exceptions.RequestException as e:
        st.error(f"❌ Errore di rete nella richiesta SerpAPI: {e}")
        return []
    if response.status_code == 200:
        try:
            data = response.json()
        except ValueError:
            st.error("❌ Risposta SerpAPI non valida (JSON non parsabile).")
            return []
        return [q["question"] for q in data.get("related_questions", [])]
    else:
        st.error(f"❌ Errore nella richiesta SerpAPI: {response.status_code}")
        return []

def generate_answer(question, openai_api_key):
    # (Funzione invariata - utilizza OpenAI per generare risposta alla domanda)
    client = openai.OpenAI(api_key=openai_api_key)
    prompt = f"""
    Rispondi in modo chiaro e dettagliato alla seguente domanda:
    {question}
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "Sei un esperto che fornisce risposte chiare e dettagliate."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        st.error(f"❌ Errore generazione risposta: {e}")
        return "Nessuna risposta generata."

st.title("🔍 People Also Ask (PAA) Generator")

# Input delle API Key
serpapi_api_key = st.text_input("🔑 Enter your SerpAPI API Key", type="password")
oai_api_key = st.text_input("🔑 Enter your OpenAI API Key", type="password")

# Input keyword di ricerca
keyword = st.text_input("🔎 Enter a keyword to fetch PAA questions")

if st.button("📄 Generate PAA Questions and Answers"):
    if not serpapi_api_key or not oai_api_key or not keyword:
        st.error("⚠️ Devi inserire entrambe le API Key e una keyword!")
    else:
        st.info(f"Fetching PAA questions for: {keyword}...")
        paa_questions = get_paa_questions(keyword, serpapi_api_key)
        
        if not paa_questions:
            st.warning("⚠️ Nessuna domanda PAA trovata.")
        else:
            results = {}
            for question in paa_questions:
                st.write(f"✨ Generating answer for: {question}")
                results[question] = generate_answer(question, oai_api_key)
            
            # Mostra i risultati
            st.subheader("📌 Generated Questions & Answers")
            combined_text = ""
            for q, a in results.items():
                st.markdown(f"**Q:** {q}")
                st.markdown(f"**A:** {a}")
                combined_text += f"Domanda: {q}\nRisposta: {a}\n\n"
            
            # Pulsante per copiare i risultati
            st.download_button("📋 Copy to Clipboard", combined_text, file_name="paa_results.txt", mime="text/plain")