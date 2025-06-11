
import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
from io import BytesIO
from time import sleep

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

st.title("🔗 Validazione Link e Bozze Email")

uploaded_file = st.file_uploader("📤 Carica il file Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    sheet_name = "Report Mensile"
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        df_check = df[['URL', 'Anchor Text', 'Target URL', 'Sito']].copy()
        df_check.dropna(inplace=True)

        results = []
        email_drafts = []

        with st.spinner("🔍 Analisi dei link in corso..."):
            for idx, row in df_check.iterrows():
                st.text(f"Controllo {idx+1}/{len(df_check)}: {row['URL']}")
                result = verifica_link(row['URL'], row['Anchor Text'], row['Target URL'])
                results.append(result)
                draft = genera_bozza_email(
                    nome_sito=row['Sito'],
                    url=row['URL'],
                    anchor_text=row['Anchor Text'],
                    target_url=row['Target URL'],
                    validazione=result
                )
                email_drafts.append(draft)
                sleep(1)

        df_check['Validazione Link'] = results
        df_check['Bozza Email'] = email_drafts

        df_finale = df.merge(df_check[['URL', 'Validazione Link', 'Bozza Email']], on='URL', how='left')

        st.success("✅ Analisi completata.")
        st.dataframe(df_finale)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_finale.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)

        st.download_button(
            label="📥 Scarica il file con i risultati",
            data=output,
            file_name="BTK-GAR-Report_Validazione.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Errore durante la lettura o elaborazione del file: {e}")
