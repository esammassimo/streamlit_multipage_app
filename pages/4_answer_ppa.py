import streamlit as st
import requests
import openai
import json

def get_paa_questions(keyword, serper_api_key):
    url = "https://google.serper.dev/search"
    headers = {"X-API-KEY": serper_api_key, "Content-Type": "application/json"}
    payload = json.dumps({"q": keyword, "gl": "it", "hl": "it", "type": "search", "engine": "google"})
    
    response = requests.post(url, headers=headers, data=payload)
    if response.status_code == 200:
        data = response.json()
        return [q['question'] for q in data.get("peopleAlsoAsk", [])]
    else:
        st.error(f"❌ Error in Serper.dev request: {response.status_code}")
        return []

def generate_answer(question, openai_api_key):
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
        st.error(f"❌ Error generating answer: {e}")
        return "Nessuna risposta generata."

st.title("🔍 People Also Ask (PAA) Generator")

# API Keys Inputs
serper_api_key = st.text_input("🔑 Enter your Serper.dev API Key", type="password")
oai_api_key = st.text_input("🔑 Enter your OpenAI API Key", type="password")

# Keyword Input
keyword = st.text_input("🔎 Enter a keyword to fetch PAA questions")

if st.button("📄 Generate PAA Questions and Answers"):
    if not serper_api_key or not oai_api_key or not keyword:
        st.error("⚠️ You must enter both API Keys and a keyword!")
    else:
        st.info(f"Fetching PAA questions for: {keyword}...")
        paa_questions = get_paa_questions(keyword, serper_api_key)
        
        if not paa_questions:
            st.warning("⚠️ No PAA questions found.")
        else:
            results = {}
            for question in paa_questions:
                st.write(f"✨ Generating answer for: {question}")
                results[question] = generate_answer(question, oai_api_key)
            
            # Display results
            st.subheader("📌 Generated Questions & Answers")
            combined_text = ""
            for q, a in results.items():
                st.markdown(f"**Q:** {q}")
                st.markdown(f"**A:** {a}")
                combined_text += f"Domanda: {q}\nRisposta: {a}\n\n"
            
            # Copy to clipboard button
            st.download_button("📋 Copy to Clipboard", combined_text, file_name="paa_results.txt", mime="text/plain")