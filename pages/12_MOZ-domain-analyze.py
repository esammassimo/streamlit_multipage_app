import streamlit as st
import pandas as pd
import requests
import time

# Titolo dell'app
st.title("Moz API - Analisi Domini e Link Intersection")

# Sidebar per inserire Access ID e Secret Key manualmente
st.sidebar.subheader("Credenziali Moz API")
access_id = st.sidebar.text_input("Access ID", type="password")
secret_key = st.sidebar.text_input("Secret Key", type="password")

# Se le credenziali non sono disponibili, interrompi l'esecuzione finché non vengono fornite
if not access_id or not secret_key:
    st.warning("Inserisci Access ID e Secret Key di Moz API nella barra laterale per continuare.")
    st.stop()

# Configura l'autenticazione per le chiamate API Moz (Basic Auth con Access ID e Secret Key)
auth = (access_id, secret_key)

# Creazione delle due tab principali nell'interfaccia Streamlit
tab1, tab2 = st.tabs(["Metriche Dominio", "Link Intersection"])

# Tab 1: Metriche Dominio
with tab1:
    st.header("Metriche Dominio")
    st.write("Inserisci uno o più domini (uno per riga) per recuperare le metriche Moz:")
    domains_input = st.text_area("Domini (uno per riga)", height=150)
    if st.button("Calcola metriche Moz", key="metrics_button"):
        domains = [line.strip().split("/")[0].replace("https://", "").replace("http://", "") 
                   for line in domains_input.splitlines() if line.strip()]
        if not domains:
            st.error("Inserisci almeno un dominio valido.")
        else:
            payload = {"targets": domains}
            response = requests.post("https://lsapi.seomoz.com/v2/url_metrics", 
                                     json=payload, auth=auth)
            if response.status_code != 200:
                st.error(f"Errore API Moz: {response.status_code} - {response.text}")
            else:
                data = response.json().get("results", [])
                results = []
                for r in data:
                    results.append({
                        "Dominio": r.get("root_domain"),
                        "DA": r.get("domain_authority"),
                        "PA": r.get("page_authority"),
                        "Backlink": r.get("external_pages_to_root_domain"),
                        "Domini referenti": r.get("root_domains_to_root_domain"),
                        "Spam Score": r.get("spam_score")
                    })
                df = pd.DataFrame(results)
                st.dataframe(df)

# Tab 2: Link Intersection
with tab2:
    st.header("Link Intersection (Intersezione di domini colleganti)")
    st.write("Inserisci da 2 a 5 domini (uno per riga) per trovare i domini che li collegano in comune:")
    targets_input = st.text_area("Domini target (2-5, uno per riga)", height=100)
    if st.button("Calcola Link Intersection", key="intersection_button"):
        targets = [line.strip().split("/")[0].replace("https://", "").replace("http://", "") 
                   for line in targets_input.splitlines() if line.strip()]
        if not (2 <= len(targets) <= 5):
            st.error("Inserisci da 2 a 5 domini.")
        else:
            all_links = []
            for dom in targets:
                payload = {
                    "target": dom,
                    "target_scope": "root_domain",
                    "filter": "external",
                    "limit": 50
                }
                r = requests.post("https://lsapi.seomoz.com/v2/linking_root_domains", 
                                  json=payload, auth=auth)
                if r.status_code != 200:
                    st.error(f"Errore per {dom}: {r.status_code} - {r.text}")
                    st.stop()
                data = r.json().get("results", [])
                source_domains = set(x.get("root_domain") for x in data if x.get("root_domain"))
                all_links.append(source_domains)
                time.sleep(1)

            common_links = set.intersection(*all_links)
            domain_count = {}
            for s in all_links:
                for d in s:
                    domain_count[d] = domain_count.get(d, 0) + 1
            partial_links = [d for d, c in domain_count.items() if c > 1 and d not in common_links]

            st.subheader("Link comuni a tutti i domini")
            st.dataframe(pd.DataFrame(sorted(common_links), columns=["Dominio"]))

            st.subheader("Link parziali (a più di un dominio)")
            df_partial = pd.DataFrame([{"Dominio": d, "Target linkati": domain_count[d]} 
                                       for d in partial_links])
            df_partial = df_partial.sort_values(by="Target linkati", ascending=False)
            st.dataframe(df_partial)

            st.subheader("Distribuzione domini linkanti")
            distrib = {}
            for val in domain_count.values():
                distrib[val] = distrib.get(val, 0) + 1
            dist_df = pd.DataFrame({
                "Numero target linkati": list(distrib.keys()),
                "Numero domini": list(distrib.values())
            }).sort_values(by="Numero target linkati")
            dist_df = dist_df.set_index("Numero target linkati")
            st.bar_chart(dist_df)