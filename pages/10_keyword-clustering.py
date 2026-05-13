import streamlit as st
import pandas as pd
from openai import OpenAI
from io import BytesIO
import xlsxwriter
import time
import re

# ===================== COSTANTI =====================
MAX_ROWS = 2000  # limite righe analizzabili per volta

# ===================== APP =====================
st.set_page_config(page_title="Keyword Clustering", layout="centered")
st.title("🔎 Keyword Clustering con OpenAI")

# ===================== UTILITIES =====================
def normalize_keyword(s: str, to_lower: bool = True) -> str:
    """
    Normalizza una keyword:
      - rimuove spazi iniziali/finali
      - comprime spazi multipli interni
      - opzionale: lowercase
    """
    if not isinstance(s, str):
        return ""
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    if to_lower:
        s = s.lower()
    return s

def batch(iterable, size):
    for i in range(0, len(iterable), size):
        yield iterable[i:i+size]

def backoff_sleep(attempt):
    # esponenziale: 0.5, 1, 2, 4 (cap a 4s)
    time.sleep(min(4, 0.5 * (2 ** attempt)))

# ===================== INPUT =====================
api_key = st.text_input("🔐 Inserisci la tua OpenAI API Key", type="password")
st.caption(f"ℹ️ Limite massimo righe analizzabili per esecuzione: {MAX_ROWS:,}")

input_method = st.radio("📥 Metodo di inserimento:", ["Carica file Excel", "Incolla testo manualmente"])

raw_keywords = []
if input_method == "Carica file Excel":
    uploaded = st.file_uploader("📄 Carica un file Excel con le keyword (colonna A)", type=["xlsx"])
    if uploaded:
        df = pd.read_excel(uploaded)
        if df.empty or df.shape[1] < 1:
            st.error("Il file deve contenere almeno una colonna con le keyword.")
        else:
            col = df.iloc[:, 0].dropna().astype(str).tolist()
            total = len(col)
            if total > MAX_ROWS:
                st.warning(f"⚠️ Hai caricato {total:,} righe. Verranno analizzate solo le prime {MAX_ROWS:,}.")
                col = col[:MAX_ROWS]
            raw_keywords = col
else:
    manual = st.text_area("✍️ Incolla le keyword, una per riga")
    if manual:
        lines = [ln.strip() for ln in manual.splitlines() if ln.strip()]
        total = len(lines)
        if total > MAX_ROWS:
            st.warning(f"⚠️ Hai inserito {total:,} righe. Verranno analizzate solo le prime {MAX_ROWS:,}.")
            lines = lines[:MAX_ROWS]
        raw_keywords = lines

to_lower = st.checkbox("🔤 Converti a minuscolo (consigliato)", value=True)

# Normalizzazione + dedup (manteniamo una mappa per riportare l'originale se serve)
norm_keywords = [normalize_keyword(k, to_lower=to_lower) for k in raw_keywords]
norm_keywords = [k for k in norm_keywords if k]  # rimuovi vuoti
# dedup preservando l'ordine
seen = set()
keywords = []
for k in norm_keywords:
    if k not in seen:
        seen.add(k)
        keywords.append(k)

st.write(f"**Keyword uniche da processare:** {len(keywords)}")

# Etichette (cluster)
label_mode = st.radio("🏷️ Come vuoi fornire le etichette di cluster?", ["Inserimento manuale", "Usa file predefinito category_keyword.txt"])
labels = []
if label_mode == "Inserimento manuale":
    labels_input = st.text_input(
        "✏️ Etichette separate da virgola",
        placeholder="es. informazionale, transazionale, navigazionale, brand, prodotto"
    )
    if labels_input:
        labels = [lbl.strip() for lbl in labels_input.split(",") if lbl.strip()]
else:
    try:
        with open("category_keyword.txt", "r", encoding="utf-8") as f:
            labels = [ln.strip() for ln in f.readlines() if ln.strip()]
        st.markdown("### 📂 Etichette da category_keyword.txt:")
        st.text("\n".join(labels[:30]))
        if len(labels) > 30:
            st.text(f"...e altre {len(labels) - 30} etichette")
    except FileNotFoundError:
        st.error("❌ Il file 'category_keyword.txt' non è stato trovato nella directory dello script.")

# Modello e batch size
model_choice = st.selectbox("🤖 Modello OpenAI", ["gpt-4o", "gpt-3.5-turbo"])
batch_size = st.slider("📦 Dimensione batch (keyword per richiesta)", min_value=10, max_value=200, value=100, step=10)

# Campo opzionale per lingua/mercato
locale_hint = st.text_input("🌍 Contesto lingua/mercato (opzionale)", placeholder="es. it-IT (Italia)")

# ===================== STIMA TOKEN (indicativa) =====================
if keywords and labels:
    st.markdown("---")
    st.subheader("📈 Stima preliminare dei token (indicativa)")
    avg_prompt_tokens_per_kw = 10
    labels_tokens = len(" ".join(labels).split()) + 5
    locale_tokens = len(locale_hint.split()) if locale_hint else 0
    estimated_prompt = (labels_tokens + locale_tokens + avg_prompt_tokens_per_kw) * len(keywords)
    estimated_completion = 4 * len(keywords)  # risposta breve (solo etichetta)
    estimated_total = estimated_prompt + estimated_completion
    st.markdown(f"🔢 **Token stimati:** Prompt: `{estimated_prompt}`, Completion: `{estimated_completion}`, Totale: `{estimated_total}`")
    if model_choice == "gpt-4o":
        cost = (estimated_prompt / 1000 * 0.005) + (estimated_completion / 1000 * 0.015)
    else:
        cost = (estimated_prompt / 1000 * 0.001) + (estimated_completion / 1000 * 0.002)
    st.markdown(f"💰 **Costo stimato:** ~${cost:.4f}")

    token_threshold = 100000
    if estimated_total > token_threshold:
        st.warning(f"⚠️ Attenzione: la stima dei token supera i {token_threshold:,}. Questo potrebbe generare costi elevati!")

confirm_run = st.checkbox("✅ Conferma e avvia la classificazione")

# ===================== RUN =====================
if keywords and labels and api_key and confirm_run:
    client = OpenAI(api_key=api_key)
    results = []
    total_prompt_tokens = 0
    total_completion_tokens = 0

    sys_prompt = (
        "Sei un assistente di classificazione SEO. Riceverai una lista di keyword e una lista di cluster (etichette). "
        "Per ciascuna keyword scegli esattamente UNA etichetta dall'elenco fornito, in base all'intento e al significato. "
        "Rispondi SOLO con righe 'keyword, cluster' in CSV, nell'ordine dato. "
        "Non aggiungere commenti, spiegazioni o righe vuote."
    )

    if locale_hint:
        sys_prompt += f" Considera il contesto linguistico/mercato: {locale_hint}."

    with st.spinner("🔍 Sto classificando le keyword..."):
        st.markdown("---")
        st.markdown("### 🧭 Avanzamento della classificazione")

        for idx, kw_batch in enumerate(batch(keywords, batch_size), 1):
            st.markdown(f"➡️ **Batch {idx}**: {len(kw_batch)} keyword")
            user_prompt = (
                f"Etichette disponibili: {', '.join(labels)}.\n"
                "Keyword da classificare (una per riga):\n"
                + "\n".join(kw_batch) +
                "\n\nRispondi con una riga per keyword nel formato 'keyword, cluster'."
            )

            max_attempts = 5
            for attempt in range(max_attempts):
                try:
                    response = client.chat.completions.create(
                        model=model_choice,
                        messages=[
                            {"role": "system", "content": sys_prompt},
                            {"role": "user", "content": user_prompt},
                        ],
                        temperature=0,
                        max_tokens=6 * len(kw_batch) + 100,  # margine sufficiente
                    )
                    if not response.choices:
                        raise ValueError("Risposta OpenAI vuota.")
                    content = response.choices[0].message.content.strip()
                    usage = response.usage
                    if usage:
                        total_prompt_tokens += getattr(usage, "prompt_tokens", 0)
                        total_completion_tokens += getattr(usage, "completion_tokens", 0)

                    # parsing CSV "keyword, cluster"
                    lines = [l.strip() for l in content.splitlines() if l.strip()]
                    parsed = []
                    for line in lines:
                        line = line.replace(";", ",")
                        parts = [p.strip() for p in line.split(",")]
                        if len(parts) >= 2:
                            k = normalize_keyword(parts[0], to_lower=to_lower)
                            c = parts[1]
                            if k:
                                parsed.append((k, c))
                    mapped = {k: c for k, c in parsed}
                    for k in kw_batch:
                        results.append((k, mapped.get(k, "Uncategorized")))
                    break
                except Exception as e:
                    if attempt < max_attempts - 1:
                        backoff_sleep(attempt)
                        continue
                    for k in kw_batch:
                        results.append((k, f"Errore: {e}"))

    output_df = pd.DataFrame(results, columns=["Keyword", "Cluster"])
    output_df = output_df.sort_values(["Cluster", "Keyword"]).reset_index(drop=True)

    st.success("✅ Classificazione completata!")
    st.dataframe(output_df.head(500))

    st.markdown(f"**📊 Token utilizzati:** Prompt: `{total_prompt_tokens}`, Completion: `{total_completion_tokens}`, Totale: `{total_prompt_tokens + total_completion_tokens}`")

    # Summary per cluster
    summary_df = output_df.groupby("Cluster").size().reset_index(name="Conteggio").sort_values("Conteggio", ascending=False)

    # Export Excel con due fogli
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        output_df.to_excel(writer, index=False, sheet_name="Risultati")
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
    st.download_button(
        "⬇️ Scarica il file con i cluster",
        data=buffer.getvalue(),
        file_name="keyword_clustering.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )