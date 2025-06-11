import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
from io import BytesIO
from time import sleep

# Funzione per verificare il link
def verifica_link(url, anchor_text, target_url):
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code != 200:
            return f"❌ HTTP {response.status_code}"

        soup = BeautifulSoup(response.text, 'html.parser')
        links = soup.find_all('a', href=True)
        for link in links:
            href = link['href'].strip()
            text = link.get_text().strip()
            if href == target_url and anchor_text.lower() in text.lower():
                return "✅ Link corretto"
        return "❌ Link non trovato"
    except Exception as e:
        return f"❌ Errore: {str(e)}"

# Funzione per generare bozza email
def genera_bozza_email(nome_sito, url, anchor_text, target_url, validazione):
    if validazione.startswith("✅"):
        return ""
    return f"""Oggetto: Correzione link su {nome_sito}

Buongiorno,

abbiamo notato che il link presente su {url} non è configurato correttamente.

Dovrebbe puntare a: {target_url}
con anchor text: "{anchor_text}"

Ti chiediamo cortesemente di aggiornare il link nel più breve tempo possibile.

Grazie per la collaborazione!
"""

# Streamlit UI
st.title("🔗 Controllo Link in Tempo Reale con Bozze Email")

uploaded_file = st.file_uploader("📤 Carica il file Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    sheet_name = "Report Mensile"
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        df_check = df[['URL', 'Anchor Text', 'Target URL', 'Sito']].dropna().copy()

        results = []
        drafts = []

        st.subheader("🟢 Stato dei Link")

        for idx, row in df_check.iterrows():
            with st.container():
                col1, col2 = st.columns([0.05, 0.95])
                result = verifica_link(row['URL'], row['Anchor Text'], row['Target URL'])
                icon = "✅" if result.startswith("✅") else "❌"
                col1.markdown(icon)
                col2.markdown(f"[{row['URL']}]({row['URL']}) — `{result}`")

                results.append(result)
                drafts.append(genera_bozza_email(
                    nome_sito=row['Sito'],
                    url=row['URL'],
                    anchor_text=row['Anchor Text'],
                    target_url=row['Target URL'],
                    validazione=result
                ))

                sleep(1)  # per evitare rate limit

        # Aggiunge i risultati al dataframe
        df_check['Validazione Link'] = results
        df_check['Bozza Email'] = drafts
        df_finale = df.merge(df_check[['URL', 'Validazione Link', 'Bozza Email']], on='URL', how='left')

        # Esporta il risultato
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_finale.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)

        st.success("✅ Analisi completata.")
        st.download_button(
            label="📥 Scarica il file con i risultati",
            data=output,
            file_name="BTK-GAR-Report_Validazione.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Errore durante la lettura del file: {e}")