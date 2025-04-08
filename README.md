# AI SEO Tools

Una raccolta di strumenti SEO basati su intelligenza artificiale, sviluppati con Streamlit e OpenAI API.  
Questo progetto include diverse funzionalità utili per la generazione di contenuti, analisi delle performance, trascrizione da YouTube e altro.

## 📁 Struttura del progetto

- `app.py`: file principale dell'app Streamlit.
- `pages/`: directory contenente le varie pagine del tool:
  - `1_excel_content_gen.py`: generazione contenuti da Excel.
  - `2_direct_content_gen.py`: generazione contenuti diretta.
  - `3_youtube_transcription.py`: trascrizione video YouTube.
  - `4_page_performance.py`: analisi prestazioni di pagina.
  - `5_answer_ppa.py`: generazione risposte per "People Also Ask".

## 🛠️ Requisiti

Assicurati di avere Python 3.8+ installato.

Installa i pacchetti necessari con:

```bash
pip install -r requirements.txt
```

## 🚀 Avvio dell'app

Per avviare l'app:

```bash
streamlit run app.py
```

## 🔑 Configurazione API

Assicurati di impostare la tua chiave API di OpenAI. Puoi farlo tramite variabile d'ambiente:

```bash
export OPENAI_API_KEY="la_tua_chiave_api"
```

Oppure modifica il codice nei file dove viene usata `openai.api_key`.

## 📌 Note

- La pagina "Page Performance" potrebbe richiedere l’integrazione con le API di Google PageSpeed Insights (non inclusa nella versione attuale).
- Alcune pagine usano file Excel in input: assicurati che siano formattati correttamente.
