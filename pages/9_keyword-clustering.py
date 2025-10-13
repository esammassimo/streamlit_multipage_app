import streamlit as st
import pandas as pd
from openai import OpenAI
from io import BytesIO
import xlsxwriter
import time
from urllib.parse import urlparse
import re

st.set_page_config(page_title="Domain Clustering", layout="centered")
st.title("🌐 Domain Clustering con OpenAI")

# ===================== COSTANTI =====================
MAX_ROWS = 2000  # limite righe analizzabili per volta

# ===================== UTILITIES =====================
def normalize_domain(raw: str) -> str:
    """
    Normalizza una stringa (URL o dominio) al dominio registrabile:
      - rimuove schema, path, porta
      - rimuove 'www.'
      - fallback agli ultimi 2 label (example.co.uk -> example.co.uk)
    """
    if not isinstance(raw, str) or raw.strip() == "":
        return ""
    s = raw.strip()
    if re.match(r"^[a-zA-Z]+://", s):
        netloc = urlparse(s).netloc
    else:
        netloc = s.split("/")[0]
    netloc = netloc.split(":")[0]
    netloc = netloc.lower().strip(".")
    netloc = re.sub(r"^www\.", "", netloc)
    parts = [p for p in netloc.split(".") if p]
    if len(parts) >= 2:
        return ".".join(parts[-2:])
    return netloc

def batch(iterable, size):
    for i in range(0, len(iterable), size):
        yield iterable[i:i+size]

def backoff_sleep(attempt):
    time.sleep(min(4, 0.5 * (2 ** attempt)))

# ===================== INPUTS =====================
api_key = st.text_input("🔐 Inserisci la tua OpenAI API Key", type="password")
st.caption(f"ℹ️ Limite massimo righe analizzabili per esecuzione: {MAX_ROWS:,}")

input_method = st.radio("📥 Scegli il metodo di inserimento dei domini:", ["Carica file Excel", "Incolla testo manualmente"])

domains_raw = []
if input_method == "Carica file Excel":
    uploaded_file = st.file_uploader("📄 Carica un file Excel con i domini (colonna A)", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        if df.empty or df.shape[1] < 1:
            st.error("Il file deve contenere almeno una colonna con i domini.")
        else:
            col = df.iloc[:, 0].dropna().astype(str).tolist()
            total_rows = len(col)
            if total_rows > MAX_ROWS:
                st.warning(f"⚠️ Hai caricato {total_rows:,} righe. Verranno analizzate solo le prime {MAX_ROWS:,}.")
                col = col[:MAX_ROWS]
            domains_raw = col

else:
    manual_input = st.text_area("✍️ Incolla i domini o URL, uno per riga")
    if manual_input:
        lines = [line.strip() for line in manual_input.splitlines() if line.strip()]
        total_rows = len(lines)
        if total_rows > MAX_ROWS:
            st.warning(f"⚠️ Hai inserito {total_rows:,} righe. Verranno analizzate solo le prime {MAX_ROWS:,}.")
            lines = lines[:MAX_ROWS]
        domains_raw = lines

# Normalizzazione + dedup
domains = [normalize_domain(d) for d in domains_raw]
domains = [d for d in domains if d]
domains = sorted(set(domains))

# Etichette
label_mode = st.radio("🏷️ Come vuoi fornire le etichette?", ["Inserimento manuale", "Usa file predefinito category_domain.txt"])
labels = []
if label_mode == "Inserimento manuale":
    labels_input = st.text_input(
        "✏️ Inserisci le etichette di clustering separate da virgola",
        placeholder="es. blog, ecommerce, istituzionale, news"
    )
    if labels_input:
        labels = [label.strip() for label in labels_input.split(",") if label.strip()]
else:
    try:
        with open("category_domain.txt", "r", encoding="utf-8") as f:
            labels = [line.strip() for line in f.readlines() if line.strip()]
        st.markdown("### 📂 Etichette caricate da category_domain.txt:")
        st.text("\n".join(labels[:30]))
        if len(labels) > 30:
            st.text(f"...e altre {len(labels) - 30} categorie")
    except FileNotFoundError:
        st.error("❌ Il file 'category_domain.txt' non è stato trovato nella directory dello script.")

# Modello e batch size
model_choice = st.selectbox("🤖 Seleziona il modello OpenAI da utilizzare", ["gpt-4o", "gpt-3.5-turbo"])
batch_size = st.slider("📦 Dimensione batch (domini per richiesta)", min_value=5, max_value=200, value=50, step=5)

# ===================== STIMA TOKEN =====================
if domains and labels:
    st.markdown("---")
    st.subheader("📈 Stima preliminare dei token (indicativa)")
    avg_prompt_tokens_per_domain = 10
    labels_tokens = len(" ".join(labels).split()) + 5
    estimated_prompt = (labels_tokens + avg_prompt_tokens_per_domain) * len(domains)
    estimated_completion = 4 * len(domains)
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
if domains and labels and api_key and confirm_run:
    client = OpenAI(api_key=api_key)
    results = []
    total_prompt_tokens = 0
    total_completion_tokens = 0

    sys_prompt = (
        "Sei un assistente di classificazione. Ti verrà passato un elenco di domini "
        "e una lista di categorie. Per ciascun dominio scegli esattamente UNA categoria "
        "dall'elenco fornito. Rispondi SOLO con righe 'dominio, categoria' in CSV, nell'ordine dato."
    )

    with st.spinner("🔍 Sto classificando i domini..."):
        st.markdown("---")
        st.markdown("### 🧭 Avanzamento della classificazione")

        for idx, dom_batch in enumerate(batch(domains, batch_size), 1):
            st.markdown(f"➡️ **Batch {idx}**: {len(dom_batch)} domini")
            user_prompt = (
                f"Categorie disponibili: {', '.join(labels)}.\n"
                "Domini da classificare (uno per riga):\n"
                + "\n".join(dom_batch) +
                "\n\nRispondi con una riga per dominio nel formato 'dominio, categoria'."
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
                        max_tokens=4 * len(dom_batch) + 50,
                    )
                    content = response.choices[0].message.content.strip()
                    usage = response.usage
                    if usage:
                        total_prompt_tokens += getattr(usage, "prompt_tokens", 0)
                        total_completion_tokens += getattr(usage, "completion_tokens", 0)

                    lines = [l.strip() for l in content.splitlines() if l.strip()]
                    parsed = []
                    for line in lines:
                        line = line.replace(";", ",")
                        parts = [p.strip() for p in line.split(",")]
                        if len(parts) >= 2:
                            d = normalize_domain(parts[0])
                            c = parts[1]
                            if d:
                                parsed.append((d, c))
                    mapped = {d: c for d, c in parsed}
                    for d in dom_batch:
                        results.append((d, mapped.get(d, "Uncategorized")))
                    break
                except Exception as e:
                    if attempt < max_attempts - 1:
                        backoff_sleep(attempt)
                        continue
                    for d in dom_batch:
                        results.append((d, f"Errore: {e}"))

    output_df = pd.DataFrame(results, columns=["Dominio", "Cluster"])
    output_df = output_df.sort_values(["Cluster", "Dominio"]).reset_index(drop=True)

    st.success("✅ Classificazione completata!")
    st.dataframe(output_df.head(500))

    st.markdown(f"**📊 Token utilizzati:** Prompt: `{total_prompt_tokens}`, Completion: `{total_completion_tokens}`, Totale: `{total_prompt_tokens + total_completion_tokens}`")

    summary_df = output_df.groupby("Cluster").size().reset_index(name="Conteggio").sort_values("Conteggio", ascending=False)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        output_df.to_excel(writer, index=False, sheet_name="Risultati")
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
    st.download_button(
        "⬇️ Scarica il file con i cluster",
        data=buffer.getvalue(),
        file_name="domain_clustering.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
