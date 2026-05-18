"""
NVL Agency — DOCX → HTML Converter
Pagina Streamlit multipage: upload, config e input nel main (flusso verticale unico).
Nessun st.set_page_config() — lo gestisce il main dell'app.
"""

import io
import tempfile
from pathlib import Path
from typing import List, Optional

import streamlit as st

# ── CSS (iniettato una volta per pagina) ───────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&family=DM+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.block-container { padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1060px; }

/* ── header ── */
.nvl-header {
    display: flex; align-items: baseline; gap: 0.75rem;
    border-bottom: 2px solid #111; padding-bottom: 0.65rem; margin-bottom: 1.75rem;
}
.nvl-header .brand {
    font-family: 'DM Mono', monospace; font-size: 0.9rem; font-weight: 500;
    letter-spacing: 0.14em; color: #111; text-transform: uppercase;
}
.nvl-header .sep { color: #ccc; }
.nvl-header .title { font-size: 0.95rem; font-weight: 400; color: #555; }

/* ── section labels ── */
.sec-label {
    font-family: 'DM Mono', monospace; font-size: 0.7rem; font-weight: 500;
    letter-spacing: 0.14em; text-transform: uppercase; color: #999;
    margin: 1.75rem 0 0.6rem; border-bottom: 1px solid #efefef; padding-bottom: 0.35rem;
}

/* ── upload zone ── */
.stFileUploader > label {
    font-family: 'DM Mono', monospace; font-size: 0.75rem;
    font-weight: 500; letter-spacing: 0.08em; text-transform: uppercase; color: #555;
}
.stFileUploader section {
    border: 1.5px dashed #d0d0d0 !important; border-radius: 6px !important;
    background: #fafafa !important; transition: border-color 0.18s;
}
.stFileUploader section:hover { border-color: #111 !important; }

/* ── options row ── */
.opt-grid {
    display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
    gap: 0.6rem; margin: 0.5rem 0 0.25rem;
}
.opt-card {
    border: 1px solid #e8e8e8; border-radius: 6px; padding: 0.7rem 0.9rem;
    background: #fff; font-size: 0.82rem; color: #333;
}
.opt-card strong {
    display: block; font-family: 'DM Mono', monospace;
    font-size: 0.68rem; letter-spacing: 0.1em; text-transform: uppercase;
    color: #aaa; margin-bottom: 0.3rem;
}

/* ── buttons ── */
.stButton > button {
    font-family: 'DM Mono', monospace; font-size: 0.78rem; font-weight: 500;
    letter-spacing: 0.1em; text-transform: uppercase;
    background: #111; color: #fff; border: none; border-radius: 4px;
    padding: 0.5rem 1.5rem; transition: background 0.15s;
}
.stButton > button:hover { background: #333; }
.stButton > button:disabled { background: #ccc !important; color: #fff !important; }

.stDownloadButton > button {
    font-family: 'DM Mono', monospace; font-size: 0.76rem; font-weight: 500;
    letter-spacing: 0.08em; text-transform: uppercase;
    background: transparent; color: #111;
    border: 1.5px solid #111; border-radius: 4px;
    padding: 0.45rem 1rem; transition: all 0.15s;
}
.stDownloadButton > button:hover { background: #111; color: #fff; }

/* ── status cards ── */
.scard {
    border-radius: 5px; padding: 0.75rem 1rem; margin: 0.4rem 0;
    font-family: 'DM Mono', monospace; font-size: 0.79rem; line-height: 1.6;
}
.sc-ok   { background: #f0faf4; border-left: 3px solid #2d9e5f; color: #1a5c38; }
.sc-warn { background: #fffbf0; border-left: 3px solid #e8a020; color: #7a4f00; }
.sc-info { background: #f0f4ff; border-left: 3px solid #4a6cf7; color: #1a2e8a; }

/* ── stat chips ── */
.stat-row { display: flex; gap: 0.6rem; flex-wrap: wrap; margin: 0.6rem 0; }
.stat-chip {
    font-family: 'DM Mono', monospace; font-size: 0.72rem;
    background: #f5f5f5; border: 1px solid #e4e4e4; border-radius: 20px;
    padding: 0.18rem 0.7rem; color: #444;
}
.stat-chip strong { color: #111; }

/* ── block pills ── */
.bpill {
    display: inline-block; font-family: 'DM Mono', monospace;
    font-size: 0.68rem; font-weight: 500; padding: 0.13rem 0.4rem;
    border-radius: 3px; letter-spacing: 0.05em; white-space: nowrap;
}
.p-h1       { background: #111;    color: #fff; }
.p-intro    { background: #e8e8e8; color: #333; }
.p-s3       { background: #e0eaff; color: #1a3a8a; }
.p-img      { background: #fff7d6; color: #7a5200; }
.p-carousel { background: #e6faef; color: #1a5c38; }
.p-meta     { background: #f0e8ff; color: #4a1a8a; }

/* ── preview table ── */
.ptable-wrap { border: 1px solid #e8e8e8; border-radius: 6px; overflow: hidden; margin-top: 0.75rem; }
.ptable { width: 100%; border-collapse: collapse; font-size: 0.8rem; }
.ptable th {
    background: #111; color: #fff; font-family: 'DM Mono', monospace;
    font-size: 0.7rem; letter-spacing: 0.08em; text-transform: uppercase;
    padding: 0.5rem 0.85rem; text-align: left;
}
.ptable td {
    padding: 0.5rem 0.85rem; border-bottom: 1px solid #f2f2f2;
    vertical-align: top; line-height: 1.5;
}
.ptable tr:last-child td { border-bottom: none; }
.ptable tr:hover td { background: #fafafa; }
.html-cell {
    font-family: 'DM Mono', monospace; font-size: 0.72rem; color: #555;
    white-space: pre-wrap; word-break: break-all; max-width: 580px;
}

/* ── soc flow ── */
.soc-flow { display: flex; flex-wrap: wrap; gap: 0.3rem; align-items: center; margin: 0.5rem 0 1rem; }
.soc-arrow { color: #ccc; font-size: 0.7rem; }

/* ── divider ── */
.thin-hr { border: none; border-top: 1px solid #efefef; margin: 1.5rem 0; }

/* ── empty state ── */
.empty-state {
    text-align: center; padding: 3rem 1rem; color: #ccc;
    font-family: 'DM Mono', monospace; font-size: 0.82rem; letter-spacing: 0.08em;
}
.empty-state .icon { font-size: 2.2rem; margin-bottom: 0.5rem; }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="nvl-header">
    <span class="brand">NVL Agency</span>
    <span class="sep">·</span>
    <span class="title">DOCX → HTML Converter</span>
</div>
""", unsafe_allow_html=True)

# ── Import motore ──────────────────────────────────────────────────────────────
try:
    from docx_to_html_engine import (
        parse_input_docx,
        load_ped_lookup,
        lookup_article_extras,
        build_html_rows,
        build_structure_of_content,
        write_output_docx,
    )
except ImportError as e:
    st.error(f"❌ Impossibile importare `docx_to_html_engine.py`: {e}")
    st.stop()

# ── Session state defaults ─────────────────────────────────────────────────────
if "dth_results"    not in st.session_state: st.session_state.dth_results    = []
if "dth_ped_lookup" not in st.session_state: st.session_state.dth_ped_lookup = {}
if "dth_xlsx_hash"  not in st.session_state: st.session_state.dth_xlsx_hash  = None


# ══════════════════════════════════════════════════════════════════════════════
#  1. INPUT — UPLOAD FILE
# ══════════════════════════════════════════════════════════════════════════════

st.markdown('<p class="sec-label">1 · File di input</p>', unsafe_allow_html=True)

docx_files = st.file_uploader(
    "Brief editoriali (.docx) — puoi caricarne più di uno",
    type=["docx"],
    accept_multiple_files=True,
    key="dth_docx_upload",
    help="Trascina uno o più file .docx dei brief da convertire.",
)

xlsx_file = st.file_uploader(
    "Piano editoriale PED (.xlsx) — opzionale, per placeholder immagini e carosello",
    type=["xlsx"],
    accept_multiple_files=False,
    key="dth_xlsx_upload",
    help="Foglio 'PED 2026 - Short View', colonne Q (n° immagini) e R (ID prodotti carosello).",
)

# Carica / invalida PED quando cambia il file
if xlsx_file is not None:
    raw = xlsx_file.read()
    xlsx_hash = hash(raw)
    xlsx_file.seek(0)

    if st.session_state.dth_xlsx_hash != xlsx_hash:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(raw)
            tmp_xlsx = Path(tmp.name)
        st.session_state.dth_ped_lookup = load_ped_lookup(tmp_xlsx)
        st.session_state.dth_xlsx_hash  = xlsx_hash

    n_entries = len([k for k in st.session_state.dth_ped_lookup if not k.startswith("__h1__")])
    st.markdown(
        f'<div class="scard sc-ok">✓ PED caricato — <strong>{n_entries}</strong> articoli indicizzati</div>',
        unsafe_allow_html=True,
    )
else:
    # Reset se il file viene rimosso
    if st.session_state.dth_xlsx_hash is not None:
        st.session_state.dth_ped_lookup = {}
        st.session_state.dth_xlsx_hash  = None


# ══════════════════════════════════════════════════════════════════════════════
#  2. CONFIGURAZIONE
# ══════════════════════════════════════════════════════════════════════════════

st.markdown('<p class="sec-label">2 · Configurazione</p>', unsafe_allow_html=True)

cfg_col1, cfg_col2, cfg_col3 = st.columns(3, gap="medium")

with cfg_col1:
    img_strategy = st.selectbox(
        "Distribuzione placeholder immagini",
        options=[
            "Distribuiti uniformemente",
            "Solo dopo l'intro",
            "Solo in fondo al contenuto",
        ],
        index=0,
        key="dth_img_strategy",
    )

with cfg_col2:
    show_preview = st.checkbox("Mostra anteprima blocchi", value=True, key="dth_preview")
    show_html    = st.checkbox("HTML completo nella preview", value=False, key="dth_show_html")

with cfg_col3:
    merge_output = st.checkbox(
        "Unisci output in un unico .docx",
        value=False,
        key="dth_merge",
        help="Tutti i brief convertiti vengono concatenati in un solo file con page break.",
    )


# ══════════════════════════════════════════════════════════════════════════════
#  3. CONVERTI
# ══════════════════════════════════════════════════════════════════════════════

st.markdown('<p class="sec-label">3 · Conversione</p>', unsafe_allow_html=True)

btn_convert = st.button(
    "▶ Converti",
    disabled=not bool(docx_files),
    key="dth_btn_convert",
)

if btn_convert and docx_files:
    results = []
    lookup  = st.session_state.dth_ped_lookup
    prog    = st.progress(0, text="Elaborazione in corso…")

    for i, uploaded in enumerate(docx_files):
        prog.progress(i / len(docx_files), text=f"Elaboro: {uploaded.name}")
        try:
            # Salva in temp file (python-docx richiede path o file-like seekable)
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_in:
                tmp_in.write(uploaded.read())
                tmp_in_path = Path(tmp_in.name)

            parsed = parse_input_docx(tmp_in_path)

            # Lookup PED
            slug   = (parsed.get("meta", {}).get("URL") or "").strip()
            h1     = (parsed.get("h1") or "").strip()
            extras = lookup_article_extras(lookup, slug=slug, h1=h1)

            n_images    = extras["n_images"]
            product_ids = extras["product_ids"]

            # Applica strategia immagini
            if img_strategy == "Solo dopo l'intro":
                # Solo la prima sezione riceve IMAGE
                n_images_for_build = min(1, n_images)
                html_rows = build_html_rows(parsed, n_images=n_images_for_build, product_ids=product_ids)
            elif img_strategy == "Solo in fondo al contenuto":
                # Costruiamo senza immagini, poi appendiamo i placeholder prima di Related Product
                html_rows = build_html_rows(parsed, n_images=0, product_ids=product_ids)
                img_blocks = [
                    ("S{} IMAGE".format(j + 1), "")
                    for j in range(n_images)
                ]
                related_idx = next(
                    (j for j, (b, _) in enumerate(html_rows) if "Related Product" in b), None
                )
                if related_idx is not None:
                    html_rows = html_rows[:related_idx] + img_blocks + html_rows[related_idx:]
                else:
                    html_rows = html_rows + img_blocks
            else:  # Distribuiti uniformemente (default)
                html_rows = build_html_rows(parsed, n_images=n_images, product_ids=product_ids)

            # Genera docx output
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_out:
                tmp_out_path = Path(tmp_out.name)
            write_output_docx(parsed, tmp_out_path, n_images=n_images, product_ids=product_ids)
            docx_bytes = tmp_out_path.read_bytes()

            results.append({
                "name":       uploaded.name,
                "ok":         True,
                "parsed":     parsed,
                "extras":     extras,
                "html_rows":  html_rows,
                "docx_bytes": docx_bytes,
                "error":      None,
            })

        except Exception as exc:
            results.append({
                "name":       uploaded.name,
                "ok":         False,
                "parsed":     {},
                "extras":     {},
                "html_rows":  [],
                "docx_bytes": b"",
                "error":      str(exc),
            })

    prog.progress(1.0, text=f"✓ Completati {len(results)} file")
    st.session_state.dth_results = results


# ══════════════════════════════════════════════════════════════════════════════
#  4. RISULTATI
# ══════════════════════════════════════════════════════════════════════════════

results = st.session_state.dth_results

if results:
    ok_count  = sum(1 for r in results if r["ok"])
    err_count = len(results) - ok_count

    st.markdown('<p class="sec-label">4 · Risultati</p>', unsafe_allow_html=True)

    # Riepilogo
    chips = (
        f'<div class="stat-row">'
        f'<span class="stat-chip">📄 <strong>{len(results)}</strong> file</span>'
        f'<span class="stat-chip">✅ <strong>{ok_count}</strong> convertiti</span>'
        + (f'<span class="stat-chip">⚠️ <strong>{err_count}</strong> errori</span>' if err_count else "")
        + "</div>"
    )
    st.markdown(chips, unsafe_allow_html=True)

    # ── Download ───────────────────────────────────────────────────────────────
    if ok_count > 0:
        if not merge_output:
            dl_cols = st.columns(min(ok_count, 4), gap="small")
            ci = 0
            for res in results:
                if not res["ok"]:
                    continue
                with dl_cols[ci % len(dl_cols)]:
                    st.download_button(
                        label=f"⬇ {Path(res['name']).stem[:28]}",
                        data=res["docx_bytes"],
                        file_name="output_" + res["name"],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dth_dl_{res['name']}_{ci}",
                    )
                ci += 1
        else:
            # Merge di tutti i docx in un unico file
            from docx import Document as DocxDoc
            merged = DocxDoc()
            first  = True
            for res in results:
                if not res["ok"]:
                    continue
                if not first:
                    merged.add_page_break()
                first = False
                sub = DocxDoc(io.BytesIO(res["docx_bytes"]))
                for elem in sub.element.body:
                    merged.element.body.append(elem.__class__(elem.xml))
            buf = io.BytesIO()
            merged.save(buf)
            st.download_button(
                label="⬇ Scarica tutto (file unico)",
                data=buf.getvalue(),
                file_name="output_merged.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dth_dl_merged",
            )

    st.markdown('<hr class="thin-hr">', unsafe_allow_html=True)

    # ── Preview per file ───────────────────────────────────────────────────────
    if show_preview:

        # Mappe pill — i blocchi IMAGE hanno nome dinamico tipo "S1 IMAGE"
        PILL_CLS = {
            "H1": "p-h1", "Intro": "p-intro",
        }
        PILL_LBL = {
            "H1": "H1", "Intro": "INTRO",
        }

        def pill(block: str) -> str:
            if block in PILL_CLS:
                return f'<span class="bpill {PILL_CLS[block]}">{PILL_LBL[block]}</span>'
            if "IMAGE" in block:
                return f'<span class="bpill p-img">🖼 {block}</span>'
            if "Related Product" in block:
                return f'<span class="bpill p-carousel">➡️ Related Product</span>'
            # Sezioni Sx
            return f'<span class="bpill p-s3">{block}</span>'

        for res in results:
            exp_label = ("✅ " if res["ok"] else "❌ ") + res["name"]
            with st.expander(exp_label, expanded=(len(results) == 1)):

                if not res["ok"]:
                    st.markdown(
                        f'<div class="scard sc-warn">⚠️ {res["error"]}</div>',
                        unsafe_allow_html=True,
                    )
                    continue

                parsed    = res["parsed"]
                extras    = res["extras"]
                html_rows = res["html_rows"]
                meta      = parsed.get("meta", {})

                # Status PED
                if extras.get("n_images") or extras.get("product_ids"):
                    imgs  = extras.get("n_images", 0)
                    prods = ", ".join(extras.get("product_ids", [])) or "—"
                    ped_html = (
                        f'<div class="scard sc-ok">'
                        f'PED ✓ — <strong>{imgs}</strong> immagini &nbsp;·&nbsp; '
                        f'Carosello: <strong>{prods}</strong></div>'
                    )
                elif st.session_state.dth_ped_lookup:
                    ped_html = (
                        '<div class="scard sc-warn">'
                        '⚠️ Articolo non trovato nel PED — nessun placeholder aggiunto</div>'
                    )
                else:
                    ped_html = (
                        '<div class="scard sc-info">'
                        'ℹ️ Nessun PED caricato — placeholder immagini/carosello disabilitati</div>'
                    )

                # Metadati
                meta_rows = "".join(
                    f"<tr>"
                    f"<td style='width:130px'><span class='bpill p-meta'>{k}</span></td>"
                    f"<td class='html-cell'>{v or '<em style=\"color:#ccc\">—</em>'}</td>"
                    f"</tr>"
                    for k, v in meta.items()
                    if k in ["Title", "Description", "URL", "Target Keyword"]
                )

                # Structure of content
                structure  = build_structure_of_content(html_rows)
                soc_items  = []
                for idx, b in enumerate(structure):
                    # build_structure_of_content restituisce già H1 / Intro / ✏️ S3
                    label = b.replace("✏️ ", "").replace("🖼️ ", "").replace("🎠 ", "")
                    if label == "H1":
                        soc_items.append('<span class="bpill p-h1">H1</span>')
                    elif label == "Intro":
                        soc_items.append('<span class="bpill p-intro">INTRO</span>')
                    else:
                        soc_items.append('<span class="bpill p-s3">S3</span>')
                    if idx < len(structure) - 1:
                        soc_items.append('<span class="soc-arrow">›</span>')
                soc_html = "\n".join(soc_items)

                # Blocchi output
                MAX_PREVIEW = 320
                block_rows = "".join(
                    f"<tr>"
                    f"<td style='width:95px'>{pill(block)}</td>"
                    f"<td class='html-cell'>"
                    + (html_content if show_html else
                       (html_content[:MAX_PREVIEW] + "…" if len(html_content) > MAX_PREVIEW else html_content))
                    + "</td></tr>"
                    for block, html_content in html_rows
                )

                st.markdown(f"""
{ped_html}

<p class="sec-label" style="margin-top:1.25rem">Metadati estratti</p>
<div class="ptable-wrap">
  <table class="ptable">{meta_rows}</table>
</div>

<p class="sec-label" style="margin-top:1.25rem">Structure of content</p>
<div class="soc-flow">{soc_html}</div>

<p class="sec-label">Blocchi output</p>
<div class="ptable-wrap">
  <table class="ptable">
    <thead><tr><th style="width:95px">Block</th><th>HTML Output</th></tr></thead>
    <tbody>{block_rows}</tbody>
  </table>
</div>
""", unsafe_allow_html=True)

elif not docx_files:
    st.markdown("""
<div class="empty-state">
    <div class="icon">⚙️</div>
    Carica uno o più .docx per iniziare
</div>
""", unsafe_allow_html=True)
