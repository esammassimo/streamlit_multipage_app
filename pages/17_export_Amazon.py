import io
import json
import time
from datetime import datetime, date

import pandas as pd
import requests
import streamlit as st

API_URL = "https://serpapi.com/search.json"

# =========================================================
# CACHE GIORNALIERA
# =========================================================
_CACHE_SS = "_amazon_cache"

def _cache_get(key):
    cache = st.session_state.get(_CACHE_SS, {})
    if cache.get("date") != date.today().isoformat():
        return None
    return cache.get("data", {}).get(key)

def _cache_set(key, value):
    today = date.today().isoformat()
    cache = st.session_state.get(_CACHE_SS, {})
    if cache.get("date") != today:
        cache = {"date": today, "data": {}}
    cache["data"][key] = value
    st.session_state[_CACHE_SS] = cache


# =========================================================
# CONFIG
# =========================================================
STANDARD_FIELDS = [
    "title",
    "brand",
    "short_description",
    "long_description",
    "bullet_points",
    "main_image",
    "price",
    "old_price",
    "currency",
    "availability",
    "merchant",
    "seller",
    "offer",
    "buybox",
    "rating",
    "reviews_count",
    "amazon_url",
]

MARKETPLACES = [
    "amazon.it",
    "amazon.de",
    "amazon.fr",
    "amazon.es",
    "amazon.co.uk",
    "amazon.com",
]

FIELD_LABELS = {
    "title": "Titolo",
    "brand": "Brand",
    "short_description": "Descrizione breve",
    "long_description": "Descrizione lunga",
    "bullet_points": "Bullet point",
    "main_image": "Immagine principale",
    "price": "Prezzo",
    "old_price": "Prezzo barrato",
    "currency": "Valuta",
    "availability": "Disponibilità",
    "merchant": "Merchant",
    "seller": "Seller",
    "offer": "Offerta / coupon",
    "buybox": "Buy Box",
    "rating": "Rating",
    "reviews_count": "Numero recensioni",
    "amazon_url": "URL Amazon",
}


# =========================================================
# CONFIG · RICERCA DA KEYWORD (engine=amazon)
# =========================================================
SEARCH_FIELDS = [
    "organic_rank",
    "sponsored_rank",
    "sponsored",
    "asin",
    "title",
    "brand",
    "price",
    "price_raw",
    "currency",
    "old_price",
    "rating",
    "reviews_count",
    "bought_last_month",
    "prime",
    "availability",
    "delivery",
    "badges",
    "options",
    "thumbnail",
    "product_url",
]

SEARCH_FIELD_LABELS = {
    "organic_rank": "Ranking organico",
    "sponsored_rank": "Ranking sponsorizzato",
    "sponsored": "Sponsorizzato",
    "asin": "ASIN",
    "title": "Nome prodotto",
    "brand": "Brand",
    "price": "Prezzo",
    "price_raw": "Prezzo (testo)",
    "currency": "Valuta",
    "old_price": "Prezzo barrato",
    "rating": "Rating",
    "reviews_count": "Numero recensioni",
    "bought_last_month": "Acquisti ultimo mese",
    "prime": "Prime",
    "availability": "Disponibilità",
    "delivery": "Consegna",
    "badges": "Badge",
    "options": "Opzioni / varianti",
    "thumbnail": "Immagine",
    "product_url": "URL prodotto",
}

SORT_OPTIONS = {
    "Rilevanza (default Amazon)": "",
    "Prezzo crescente": "price-asc-rank",
    "Prezzo decrescente": "price-desc-rank",
    "Recensioni migliori": "review-rank",
    "Novità": "date-desc-rank",
    "Più venduti": "exact-aware-popularity-rank",
}


# =========================================================
# SESSION STATE
# =========================================================
def init_session_state():
    defaults = {
        "serpapi_key": "",
        "uploaded_file_name": None,
        "input_df": None,
        "sheet_names": [],
        "selected_sheet": None,
        "asin_column": None,
        "asin_list": [],
        "amazon_domain": "amazon.it",
        "delay_seconds": 1.0,
        "deduplicate_asins": True,
        "show_logs": True,
        "save_raw_json": False,
        "show_image_preview": True,
        "selected_fields": STANDARD_FIELDS.copy(),
        "results_df": None,
        "raw_results": [],
        "run_completed": False,
        # --- Tab "Ricerca da keyword" ---
        "kw_uploaded_file_name": None,
        "kw_input_df": None,
        "kw_sheet_names": [],
        "kw_selected_sheet": None,
        "kw_column": None,
        "kw_list": [],
        "kw_deduplicate": True,
        "kw_amazon_domain": "amazon.it",
        "kw_sort": "",
        "kw_delay_seconds": 1.0,
        "kw_include_sponsored": True,
        "kw_max_results": 0,
        "kw_show_logs": True,
        "kw_selected_fields": SEARCH_FIELDS.copy(),
        "kw_results_df": None,
        "kw_run_completed": False,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


# =========================================================
# HELPER GENERICI
# =========================================================
def safe_str(value):
    if value is None:
        return ""
    return str(value).strip()


def get_nested(data, *keys, default=None):
    current = data
    for key in keys:
        if isinstance(current, dict):
            current = current.get(key, default)
        else:
            return default
    return current


def first_non_empty(values, default=""):
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and value.strip():
            return value.strip()
        if isinstance(value, (int, float)):
            return value
        if isinstance(value, list) and len(value) > 0:
            return value
        if isinstance(value, dict) and len(value) > 0:
            return value
    return default


def normalize_asin(value):
    asin = safe_str(value).upper()
    asin = asin.replace(" ", "")
    return asin


def list_to_pipe_text(items):
    if not items:
        return ""
    cleaned = []
    for item in items:
        if isinstance(item, dict):
            text = first_non_empty(
                [
                    item.get("text"),
                    item.get("item"),   # chiave usata da SerpAPI per feature_bullets
                    item.get("title"),
                    item.get("name"),
                    item.get("value"),
                    # fallback: primo valore stringa non vuoto del dict
                    next((v for v in item.values() if isinstance(v, str) and v.strip()), None),
                ],
                default="",
            )
            if text:
                cleaned.append(safe_str(text))
        else:
            text = safe_str(item)
            if text:
                cleaned.append(text)
    return " | ".join(cleaned)


def number_from_any(value):
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)):
        return value
    text = safe_str(value).replace(",", ".")
    filtered = "".join(ch for ch in text if ch.isdigit() or ch == ".")
    if filtered.count(".") > 1:
        first_dot = filtered.find(".")
        filtered = filtered[: first_dot + 1] + filtered[first_dot + 1 :].replace(".", "")
    try:
        return float(filtered)
    except Exception:
        return safe_str(value)


def make_arrow_safe(df):
    """
    Uniforma i tipi delle colonne object: Streamlit (pyarrow) va in errore
    quando una colonna mescola numeri e stringhe vuote.
    """
    for col in df.columns:
        if df[col].dtype != object:
            continue
        series = df[col]
        values = [v for v in series.tolist() if v is not None and v != ""]

        if not values:
            df[col] = series.apply(lambda v: "" if v is None else str(v))
            continue

        if all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in values):
            numeric = pd.to_numeric(series.replace("", None), errors="coerce")
            if all(isinstance(v, int) for v in values):
                try:
                    numeric = numeric.astype("Int64")
                except Exception:
                    pass
            df[col] = numeric
        elif any(not isinstance(v, str) for v in values):
            df[col] = series.apply(lambda v: "" if v is None or v == "" else str(v))

    return df


def json_dumps_safe(data):
    try:
        return json.dumps(data, ensure_ascii=False)
    except Exception:
        return ""


# =========================================================
# INPUT FILE
# =========================================================
def load_input_file(uploaded_file):
    """
    Restituisce:
    - df preview / default
    - sheet_names (solo per Excel)
    - file_type ('csv' o 'xlsx')
    """
    file_name = uploaded_file.name.lower()
    file_bytes = uploaded_file.getvalue()

    if file_name.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes))
        return df, [], "csv"

    if file_name.endswith(".xlsx"):
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        first_sheet = xls.sheet_names[0]
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=first_sheet, engine="openpyxl")
        return df, xls.sheet_names, "xlsx"

    raise ValueError("Formato file non supportato. Usa CSV o XLSX.")


def read_dataframe_from_uploaded_file(uploaded_file, selected_sheet=None):
    file_name = uploaded_file.name.lower()
    file_bytes = uploaded_file.getvalue()

    if file_name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(file_bytes))

    if file_name.endswith(".xlsx"):
        if not selected_sheet:
            raise ValueError("Seleziona un foglio Excel.")
        return pd.read_excel(io.BytesIO(file_bytes), sheet_name=selected_sheet, engine="openpyxl")

    raise ValueError("Formato file non supportato. Usa CSV o XLSX.")


def extract_asins_from_df(df, asin_column, deduplicate=True):
    if asin_column not in df.columns:
        raise ValueError(f"La colonna '{asin_column}' non esiste nel file.")

    raw_values = df[asin_column].dropna().tolist()
    cleaned = [normalize_asin(v) for v in raw_values]
    cleaned = [x for x in cleaned if x]

    duplicates_removed = 0
    if deduplicate:
        original_count = len(cleaned)
        cleaned = list(dict.fromkeys(cleaned))
        duplicates_removed = original_count - len(cleaned)

    return cleaned, duplicates_removed


# =========================================================
# SERPAPI
# =========================================================
def get_amazon_product_data(asin, api_key, amazon_domain="amazon.it"):
    cache_key = f"{asin}|{amazon_domain}"
    cached = _cache_get(cache_key)
    if cached is not None:
        return cached

    params = {
        "engine": "amazon_product",
        "api_key": api_key,
        "amazon_domain": amazon_domain,
        "asin": asin,
        "output": "json",
    }

    response = requests.get(API_URL, params=params, timeout=60)

    if response.status_code != 200:
        raise RuntimeError(
            f"Errore API per ASIN {asin} - HTTP {response.status_code}: {response.text[:300]}"
        )

    try:
        data = response.json()
    except ValueError:
        raise RuntimeError(f"Risposta non valida per ASIN {asin} (JSON non parsabile).")

    if "error" in data:
        raise RuntimeError(f"Errore API per ASIN {asin}: {data['error']}")

    _cache_set(cache_key, data)
    return data


def parse_price_info(product_results, root_data):
    candidate_prices = [
        get_nested(product_results, "buybox_winner", "price", "value"),
        get_nested(product_results, "buybox_winner", "price", "raw"),
        get_nested(product_results, "buybox_winner", "price"),
        get_nested(product_results, "featured_offer", "price", "value"),
        get_nested(product_results, "featured_offer", "price"),
        get_nested(product_results, "price", "value"),
        get_nested(product_results, "price"),
        get_nested(root_data, "price", "value"),
        get_nested(root_data, "price"),
    ]

    price = first_non_empty(candidate_prices, default="")

    old_price_candidates = [
        get_nested(product_results, "buybox_winner", "old_price", "value"),
        get_nested(product_results, "buybox_winner", "old_price"),
        get_nested(product_results, "featured_offer", "old_price", "value"),
        get_nested(product_results, "featured_offer", "old_price"),
        get_nested(product_results, "old_price", "value"),
        get_nested(product_results, "old_price"),
    ]
    old_price = first_non_empty(old_price_candidates, default="")

    currency_candidates = [
        get_nested(product_results, "buybox_winner", "price", "currency"),
        get_nested(product_results, "featured_offer", "price", "currency"),
        get_nested(product_results, "price", "currency"),
        get_nested(root_data, "price", "currency"),
    ]
    currency = first_non_empty(currency_candidates, default="")

    prices_list = first_non_empty(
        [
            get_nested(product_results, "prices"),
            get_nested(root_data, "prices"),
        ],
        default=[],
    )
    if not price and isinstance(prices_list, list) and prices_list:
        first_price_item = prices_list[0]
        if isinstance(first_price_item, dict):
            price = first_non_empty(
                [
                    first_price_item.get("value"),
                    first_price_item.get("price"),
                    first_price_item.get("raw"),
                    first_price_item.get("text"),
                ],
                default="",
            )
            if not old_price:
                old_price = first_non_empty(
                    [
                        first_price_item.get("old_price"),
                        get_nested(first_price_item, "old_price", "value"),
                    ],
                    default="",
                )
            if not currency:
                currency = first_non_empty(
                    [
                        first_price_item.get("currency"),
                        get_nested(first_price_item, "price", "currency"),
                    ],
                    default="",
                )

    return price, old_price, currency


def parse_bullets(product_results, root_data):
    bullets = first_non_empty(
        [
            get_nested(product_results, "about_item"),       # chiave reale SerpAPI amazon_product
            get_nested(product_results, "feature_bullets"),
            get_nested(product_results, "about_this_item"),
            get_nested(product_results, "features"),
            get_nested(root_data, "about_item"),
            get_nested(root_data, "feature_bullets"),
        ],
        default=[],
    )

    if isinstance(bullets, list):
        return list_to_pipe_text(bullets)

    if isinstance(bullets, str):
        return bullets.strip()

    return ""


def parse_short_description(product_results, root_data):
    short_desc = first_non_empty(
        [
            get_nested(product_results, "short_description"),
            get_nested(product_results, "subtitle"),
            get_nested(root_data, "short_description"),
        ],
        default="",
    )

    if short_desc:
        return safe_str(short_desc)

    # Fallback: primo elemento di about_item come descrizione breve
    about_item = first_non_empty(
        [
            get_nested(product_results, "about_item"),
            get_nested(root_data, "about_item"),
        ],
        default=[],
    )
    if isinstance(about_item, list) and about_item:
        first_item = about_item[0]
        if isinstance(first_item, dict):
            text = first_non_empty([first_item.get("text"), first_item.get("item")], default="")
            if text:
                return safe_str(text)
        elif isinstance(first_item, str) and first_item.strip():
            return first_item.strip()

    # Fallback finale: primo bullet point
    bullet_text = parse_bullets(product_results, root_data)
    if bullet_text:
        return bullet_text.split(" | ")[0].strip()

    return ""


def parse_long_description(product_results, root_data):
    # Prova prima i campi semplici
    long_desc = first_non_empty(
        [
            get_nested(product_results, "description"),
            get_nested(product_results, "product_description"),
            get_nested(root_data, "description"),
            get_nested(root_data, "product_description"),
        ],
        default="",
    )

    # product_description da SerpAPI è una lista di blocchi con features[].text
    if isinstance(long_desc, list):
        texts = []
        for block in long_desc:
            if not isinstance(block, dict):
                continue
            features = block.get("features")
            if isinstance(features, list):
                for feat in features:
                    if isinstance(feat, dict):
                        text = feat.get("text", "")
                        if text and isinstance(text, str) and text.strip():
                            texts.append(text.strip())
            # blocchi senza features ma con testo diretto
            title = block.get("title", "")
            if title and isinstance(title, str) and title.strip() and not features:
                texts.append(title.strip())
        if texts:
            return " | ".join(texts)
        # fallback generico
        return list_to_pipe_text(long_desc)

    if isinstance(long_desc, dict):
        return list_to_pipe_text(long_desc.values())

    return safe_str(long_desc)


def parse_main_image(product_results, root_data):
    image = first_non_empty(
        [
            get_nested(product_results, "main_image"),
            get_nested(product_results, "thumbnail"),
            get_nested(product_results, "images", 0),
            get_nested(root_data, "main_image"),
            get_nested(root_data, "images", 0),
        ],
        default="",
    )

    if isinstance(image, dict):
        return first_non_empty(
            [
                image.get("link"),
                image.get("image"),
                image.get("url"),
                image.get("thumbnail"),
            ],
            default="",
        )

    return safe_str(image)


def parse_availability(product_results, root_data):
    return safe_str(
        first_non_empty(
            [
                get_nested(product_results, "availability", "status"),
                get_nested(product_results, "availability"),
                get_nested(product_results, "stock"),
                get_nested(root_data, "availability"),
            ],
            default="",
        )
    )


def parse_merchant_and_seller(product_results, root_data):
    merchant = first_non_empty(
        [
            get_nested(product_results, "buybox_winner", "merchant_info"),
            get_nested(product_results, "buybox_winner", "merchant"),
            get_nested(product_results, "featured_offer", "merchant_info"),
            get_nested(product_results, "featured_offer", "merchant"),
            get_nested(product_results, "merchant_info"),
            get_nested(product_results, "merchant"),
        ],
        default="",
    )

    seller = first_non_empty(
        [
            get_nested(product_results, "buybox_winner", "seller", "name"),
            get_nested(product_results, "buybox_winner", "seller"),
            get_nested(product_results, "featured_offer", "seller", "name"),
            get_nested(product_results, "featured_offer", "seller"),
            get_nested(product_results, "seller", "name"),
            get_nested(product_results, "seller"),
            get_nested(root_data, "seller", "name"),
            get_nested(root_data, "seller"),
        ],
        default="",
    )

    if isinstance(merchant, dict):
        merchant = first_non_empty(
            [merchant.get("name"), merchant.get("text"), merchant.get("value")],
            default="",
        )

    if isinstance(seller, dict):
        seller = first_non_empty(
            [seller.get("name"), seller.get("text"), seller.get("value")],
            default="",
        )

    return safe_str(merchant), safe_str(seller)


def parse_offer(product_results, root_data):
    promo = first_non_empty(
        [
            get_nested(product_results, "coupon"),
            get_nested(product_results, "coupons"),
            get_nested(product_results, "promotions"),
            get_nested(product_results, "special_offers"),
            get_nested(root_data, "coupon"),
            get_nested(root_data, "promotions"),
        ],
        default="",
    )

    if isinstance(promo, list):
        return list_to_pipe_text(promo)
    if isinstance(promo, dict):
        return list_to_pipe_text(promo.values())

    return safe_str(promo)


def parse_buybox(product_results, root_data):
    has_buybox = any(
        [
            bool(get_nested(product_results, "buybox_winner")),
            bool(get_nested(product_results, "featured_offer")),
            bool(get_nested(product_results, "buybox")),
            bool(get_nested(root_data, "buybox_winner")),
        ]
    )
    return "Yes" if has_buybox else "No"


def parse_rating_and_reviews(product_results, root_data):
    rating = first_non_empty(
        [
            get_nested(product_results, "rating"),
            get_nested(root_data, "rating"),
        ],
        default="",
    )

    reviews_count = first_non_empty(
        [
            get_nested(product_results, "reviews"),
            get_nested(product_results, "reviews_count"),
            get_nested(product_results, "ratings_total"),
            get_nested(root_data, "reviews"),
            get_nested(root_data, "reviews_count"),
        ],
        default="",
    )

    return rating, reviews_count


def parse_brand(product_results, root_data):
    brand = first_non_empty(
        [
            get_nested(product_results, "brand"),
            get_nested(product_results, "manufacturer"),
            get_nested(root_data, "brand"),
            get_nested(root_data, "manufacturer"),
        ],
        default="",
    )

    if isinstance(brand, dict):
        brand = first_non_empty(
            [brand.get("name"), brand.get("text"), brand.get("value")],
            default="",
        )

    return safe_str(brand)


def parse_amazon_url(root_data, asin, amazon_domain):
    url = first_non_empty(
        [
            get_nested(root_data, "search_metadata", "amazon_product_url"),
            get_nested(root_data, "product_results", "link"),
            get_nested(root_data, "product_results", "url"),
            get_nested(root_data, "link"),
        ],
        default="",
    )

    if url:
        return safe_str(url)

    return f"https://{amazon_domain}/dp/{asin}"


def parse_product_data(data, asin, amazon_domain="amazon.it", selected_fields=None, save_raw_json=False):
    selected_fields = selected_fields or STANDARD_FIELDS
    product_results = data.get("product_results", data)

    row = {
        "ASIN": asin,
        "Marketplace": amazon_domain,
        "Status": "OK",
        "Error": "",
        "Extraction Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    title = safe_str(
        first_non_empty(
            [
                get_nested(product_results, "title"),
                get_nested(data, "title"),
            ],
            default="",
        )
    )
    brand = parse_brand(product_results, data)
    short_description = parse_short_description(product_results, data)
    long_description = parse_long_description(product_results, data)
    bullet_points = parse_bullets(product_results, data)
    main_image = parse_main_image(product_results, data)
    price, old_price, currency = parse_price_info(product_results, data)
    availability = parse_availability(product_results, data)
    merchant, seller = parse_merchant_and_seller(product_results, data)
    offer = parse_offer(product_results, data)
    buybox = parse_buybox(product_results, data)
    rating, reviews_count = parse_rating_and_reviews(product_results, data)
    amazon_url = parse_amazon_url(data, asin, amazon_domain)

    field_values = {
        "title": title,
        "brand": brand,
        "short_description": short_description,
        "long_description": long_description,
        "bullet_points": bullet_points,
        "main_image": main_image,
        "price": price,
        "old_price": old_price,
        "currency": currency,
        "availability": availability,
        "merchant": merchant,
        "seller": seller,
        "offer": offer,
        "buybox": buybox,
        "rating": rating,
        "reviews_count": reviews_count,
        "amazon_url": amazon_url,
    }

    for field in selected_fields:
        row[FIELD_LABELS[field]] = field_values.get(field, "")

    if save_raw_json:
        row["Raw JSON"] = json_dumps_safe(data)

    return row


def process_asin_list(
    asins,
    api_key,
    amazon_domain,
    selected_fields,
    delay_seconds=1.0,
    show_logs=True,
    save_raw_json=False,
):
    rows = []
    raw_results = []

    progress_bar = st.progress(0)
    status_placeholder = st.empty()
    log_placeholder = st.empty()

    total = len(asins)
    log_lines = []

    for idx, asin in enumerate(asins, start=1):
        is_cached = _cache_get(f"{asin}|{amazon_domain}") is not None
        label = "(da cache) " if is_cached else ""
        status_placeholder.write(f"🔍 {label}Elaboro ASIN **{asin}** ({idx}/{total})")

        try:
            data = get_amazon_product_data(
                asin=asin,
                api_key=api_key,
                amazon_domain=amazon_domain,
            )
            row = parse_product_data(
                data=data,
                asin=asin,
                amazon_domain=amazon_domain,
                selected_fields=selected_fields,
                save_raw_json=save_raw_json,
            )
            rows.append(row)

            if save_raw_json:
                raw_results.append({"asin": asin, "data": data})

            if show_logs:
                log_lines.append(f"✅ {asin} - OK")

        except Exception as exc:
            error_message = safe_str(exc)
            error_row = {
                "ASIN": asin,
                "Marketplace": amazon_domain,
                "Status": "ERROR",
                "Error": error_message,
                "Extraction Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

            for field in selected_fields:
                error_row[FIELD_LABELS[field]] = ""

            if save_raw_json:
                error_row["Raw JSON"] = ""

            rows.append(error_row)

            if show_logs:
                log_lines.append(f"❌ {asin} - {error_message}")

        progress_bar.progress(idx / total)

        if show_logs:
            log_placeholder.text("\n".join(log_lines[-15:]))

        if not is_cached and delay_seconds > 0 and idx < total:
            time.sleep(delay_seconds)

    status_placeholder.write("✅ Estrazione completata.")
    df = make_arrow_safe(pd.DataFrame(rows))
    return df, raw_results


# =========================================================
# RICERCA DA KEYWORD · INPUT
# =========================================================
def normalize_keyword(value):
    return " ".join(safe_str(value).split())


def extract_keywords_from_df(df, keyword_column, deduplicate=True):
    if keyword_column not in df.columns:
        raise ValueError(f"La colonna '{keyword_column}' non esiste nel file.")

    raw_values = df[keyword_column].dropna().tolist()
    cleaned = [normalize_keyword(v) for v in raw_values]
    cleaned = [x for x in cleaned if x]

    duplicates_removed = 0
    if deduplicate:
        original_count = len(cleaned)
        seen = {}
        for kw in cleaned:
            seen.setdefault(kw.lower(), kw)
        cleaned = list(seen.values())
        duplicates_removed = original_count - len(cleaned)

    return cleaned, duplicates_removed


# =========================================================
# RICERCA DA KEYWORD · SERPAPI (engine=amazon)
# =========================================================
def get_amazon_search_data(keyword, api_key, amazon_domain="amazon.it", sort_by="", page=1):
    cache_key = f"SEARCH|{keyword.lower()}|{amazon_domain}|{sort_by}|{page}"
    cached = _cache_get(cache_key)
    if cached is not None:
        return cached

    params = {
        "engine": "amazon",
        "api_key": api_key,
        "amazon_domain": amazon_domain,
        "k": keyword,
        "page": page,
        "output": "json",
    }
    if sort_by:
        params["s"] = sort_by

    response = requests.get(API_URL, params=params, timeout=60)

    if response.status_code != 200:
        raise RuntimeError(
            f"Errore API per la keyword '{keyword}' - HTTP {response.status_code}: {response.text[:300]}"
        )

    try:
        data = response.json()
    except ValueError:
        raise RuntimeError(f"Risposta non valida per la keyword '{keyword}' (JSON non parsabile).")

    if "error" in data:
        raise RuntimeError(f"Errore API per la keyword '{keyword}': {data['error']}")

    _cache_set(cache_key, data)
    return data


def extract_currency_symbol(price_text):
    text = safe_str(price_text)
    if not text:
        return ""
    symbols = "".join(
        ch for ch in text if not ch.isdigit() and ch not in {".", ",", " ", " "}
    )
    return symbols.strip()


def price_number_from_text(value):
    """
    Converte un prezzo testuale in numero gestendo sia il formato europeo
    (1.299,00 €) sia quello anglosassone ($1,299.00).
    """
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)):
        return value

    text = safe_str(value)
    digits = "".join(ch for ch in text if ch.isdigit() or ch in {".", ","})
    if not digits:
        return ""

    last_dot = digits.rfind(".")
    last_comma = digits.rfind(",")

    if last_dot > last_comma:
        # separatore decimale = punto, virgole = migliaia
        digits = digits.replace(",", "")
    elif last_comma > last_dot:
        # separatore decimale = virgola, punti = migliaia
        digits = digits.replace(".", "").replace(",", ".")
    else:
        digits = digits.replace(",", "").replace(".", "")

    try:
        return float(digits)
    except Exception:
        return safe_str(value)


def bool_to_label(value):
    if value is True:
        return "Sì"
    if value is False:
        return "No"
    if value in (None, ""):
        return ""
    return safe_str(value)


def parse_search_brand(item):
    """
    L'engine 'amazon' di SerpAPI non espone sempre il brand:
    lo recuperiamo solo se presente (best effort), senza chiamate API aggiuntive.
    """
    brand = first_non_empty(
        [
            item.get("brand"),
            item.get("manufacturer"),
            item.get("byline"),
            get_nested(item, "brand", "name"),
        ],
        default="",
    )
    if isinstance(brand, dict):
        brand = first_non_empty(
            [brand.get("name"), brand.get("text"), brand.get("value")],
            default="",
        )
    return safe_str(brand)


def collect_search_items(data, include_sponsored=True):
    """
    Raccoglie i prodotti dalla SERP Amazon in ordine di apparizione.
    Ritorna una lista di tuple (item_dict, is_sponsored).
    """
    items = []

    organic = data.get("organic_results") or []
    if isinstance(organic, list):
        for item in organic:
            if not isinstance(item, dict):
                continue
            is_sponsored = bool(item.get("sponsored"))
            if is_sponsored and not include_sponsored:
                continue
            items.append((item, is_sponsored))

    if include_sponsored:
        # Blocchi pubblicitari separati (Sponsored Brands / Product Ads)
        for block_key in ("product_ads", "sponsored_brands", "ads"):
            block = data.get(block_key)
            if not block:
                continue
            candidates = []
            if isinstance(block, dict):
                candidates = block.get("products") or []
            elif isinstance(block, list):
                for entry in block:
                    if isinstance(entry, dict) and isinstance(entry.get("products"), list):
                        candidates.extend(entry["products"])
                    elif isinstance(entry, dict):
                        candidates.append(entry)
            for item in candidates:
                if isinstance(item, dict):
                    items.append((item, True))

    return items


def parse_search_row(item, is_sponsored, keyword, amazon_domain, organic_rank, sponsored_rank, selected_fields):
    asin = normalize_asin(item.get("asin"))

    price_raw = safe_str(first_non_empty([item.get("price"), item.get("price_unit")], default=""))
    price_value = first_non_empty(
        [item.get("extracted_price"), item.get("extracted_price_unit")],
        default="",
    )
    if price_value == "" and price_raw:
        price_value = price_number_from_text(price_raw)

    old_price = first_non_empty(
        [item.get("extracted_old_price"), item.get("old_price")],
        default="",
    )
    if isinstance(old_price, str) and old_price:
        old_price = price_number_from_text(old_price)

    link = safe_str(first_non_empty([item.get("link_clean"), item.get("link")], default=""))
    if not link and asin:
        link = f"https://{amazon_domain}/dp/{asin}"

    badges = first_non_empty([item.get("badges"), item.get("tags"), item.get("badge")], default=[])
    if isinstance(badges, (list, tuple)):
        badges_text = list_to_pipe_text(list(badges))
    else:
        badges_text = safe_str(badges)

    options = item.get("options")
    if isinstance(options, (list, tuple)):
        options_text = list_to_pipe_text(list(options))
    else:
        options_text = safe_str(options)

    delivery = item.get("delivery")
    if isinstance(delivery, (list, tuple)):
        delivery_text = list_to_pipe_text(list(delivery))
    else:
        delivery_text = safe_str(delivery)

    field_values = {
        "organic_rank": organic_rank if organic_rank else "",
        "sponsored_rank": sponsored_rank if sponsored_rank else "",
        "sponsored": "Sì" if is_sponsored else "No",
        "asin": asin,
        "title": safe_str(item.get("title")),
        "brand": parse_search_brand(item),
        "price": price_value,
        "price_raw": price_raw,
        "currency": extract_currency_symbol(price_raw),
        "old_price": old_price,
        "rating": first_non_empty([item.get("rating")], default=""),
        "reviews_count": first_non_empty([item.get("reviews"), item.get("reviews_count")], default=""),
        "bought_last_month": safe_str(item.get("bought_last_month")),
        "prime": bool_to_label(first_non_empty([item.get("prime"), item.get("has_prime")], default="")),
        "availability": safe_str(first_non_empty([item.get("stock"), item.get("availability")], default="")),
        "delivery": delivery_text,
        "badges": badges_text,
        "options": options_text,
        "thumbnail": safe_str(item.get("thumbnail")),
        "product_url": link,
    }

    now = datetime.now()
    row = {
        "Keyword": keyword,
        "Marketplace": amazon_domain,
        "Status": "OK",
        "Error": "",
        "Data estrazione": now.strftime("%Y-%m-%d"),
        "Ora estrazione": now.strftime("%H:%M:%S"),
    }
    for field in selected_fields:
        row[SEARCH_FIELD_LABELS[field]] = field_values.get(field, "")

    return row


def process_keyword_list(
    keywords,
    api_key,
    amazon_domain,
    selected_fields,
    include_sponsored=True,
    sort_by="",
    max_results_per_keyword=0,
    delay_seconds=1.0,
    show_logs=True,
):
    rows = []

    progress_bar = st.progress(0)
    status_placeholder = st.empty()
    log_placeholder = st.empty()

    total = len(keywords)
    log_lines = []

    for idx, keyword in enumerate(keywords, start=1):
        cache_key = f"SEARCH|{keyword.lower()}|{amazon_domain}|{sort_by}|1"
        is_cached = _cache_get(cache_key) is not None
        label = "(da cache) " if is_cached else ""
        status_placeholder.write(f"🔎 {label}Cerco **{keyword}** ({idx}/{total})")

        try:
            data = get_amazon_search_data(
                keyword=keyword,
                api_key=api_key,
                amazon_domain=amazon_domain,
                sort_by=sort_by,
            )
            items = collect_search_items(data, include_sponsored=include_sponsored)

            if max_results_per_keyword and max_results_per_keyword > 0:
                items = items[:max_results_per_keyword]

            if not items:
                now = datetime.now()
                empty_row = {
                    "Keyword": keyword,
                    "Marketplace": amazon_domain,
                    "Status": "NO RESULTS",
                    "Error": "",
                    "Data estrazione": now.strftime("%Y-%m-%d"),
                    "Ora estrazione": now.strftime("%H:%M:%S"),
                }
                for field in selected_fields:
                    empty_row[SEARCH_FIELD_LABELS[field]] = ""
                rows.append(empty_row)
                if show_logs:
                    log_lines.append(f"⚠️ {keyword} - nessun risultato")
            else:
                organic_counter = 0
                sponsored_counter = 0
                for item, is_sponsored in items:
                    if is_sponsored:
                        sponsored_counter += 1
                    else:
                        organic_counter += 1
                    rows.append(
                        parse_search_row(
                            item=item,
                            is_sponsored=is_sponsored,
                            keyword=keyword,
                            amazon_domain=amazon_domain,
                            organic_rank=0 if is_sponsored else organic_counter,
                            sponsored_rank=sponsored_counter if is_sponsored else 0,
                            selected_fields=selected_fields,
                        )
                    )
                if show_logs:
                    log_lines.append(f"✅ {keyword} - {len(items)} risultati")

        except Exception as exc:
            error_message = safe_str(exc)
            now = datetime.now()
            error_row = {
                "Keyword": keyword,
                "Marketplace": amazon_domain,
                "Status": "ERROR",
                "Error": error_message,
                "Data estrazione": now.strftime("%Y-%m-%d"),
                "Ora estrazione": now.strftime("%H:%M:%S"),
            }
            for field in selected_fields:
                error_row[SEARCH_FIELD_LABELS[field]] = ""
            rows.append(error_row)

            if show_logs:
                log_lines.append(f"❌ {keyword} - {error_message}")

        progress_bar.progress(idx / total)

        if show_logs:
            log_placeholder.text("\n".join(log_lines[-15:]))

        if not is_cached and delay_seconds > 0 and idx < total:
            time.sleep(delay_seconds)

    status_placeholder.write("✅ Ricerca completata.")
    return make_arrow_safe(pd.DataFrame(rows))


# =========================================================
# EXPORT
# =========================================================
def dataframe_to_csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8-sig")


def dataframe_to_excel_bytes(df, sheet_name="Amazon_ASIN"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        header_format = workbook.add_format({"bold": True, "text_wrap": True})
        wrap_format = workbook.add_format({"text_wrap": True, "valign": "top"})

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            max_len = max(
                len(str(value)),
                int(df.iloc[:, col_num].map(lambda v: len(str(v))).max()) if not df.empty else 10,
            )
            worksheet.set_column(col_num, col_num, min(max(max_len + 2, 15), 50), wrap_format)

        worksheet.freeze_panes(1, 0)

    output.seek(0)
    return output


# =========================================================
# UI FASES
# =========================================================
def render_phase_1_api_config():
    st.header("Fase 1 · Configurazione API")

    with st.container(border=True):
        st.write("Inserisci la tua **SerpAPI API Key** per abilitare l’estrazione.")
        api_key = st.text_input(
            "SerpAPI API Key",
            type="password",
            value=st.session_state.get("serpapi_key", ""),
            help="La chiave viene salvata nella sessione Streamlit.",
        )

        if st.button("Salva configurazione API", use_container_width=True):
            st.session_state["serpapi_key"] = api_key.strip()
            if st.session_state["serpapi_key"]:
                st.success("API key salvata in sessione.")
            else:
                st.warning("Inserisci una API key valida.")

        if st.session_state.get("serpapi_key"):
            st.success("Configurazione API disponibile.")
        else:
            st.info("Configura la API key per proseguire.")


def render_phase_2_upload_input():
    st.header("Fase 2 · Upload input file")

    with st.container(border=True):
        uploaded_file = st.file_uploader(
            "Carica un file CSV o XLSX con gli ASIN",
            type=["csv", "xlsx"],
        )

        if uploaded_file is None:
            st.info("Carica un file per vedere anteprima e selezionare la colonna ASIN.")
            return

        try:
            preview_df, sheet_names, file_type = load_input_file(uploaded_file)
            st.session_state["uploaded_file_name"] = uploaded_file.name
            st.session_state["sheet_names"] = sheet_names

            if file_type == "xlsx":
                selected_sheet = st.selectbox(
                    "Seleziona il foglio Excel",
                    options=sheet_names,
                    index=0,
                )
                st.session_state["selected_sheet"] = selected_sheet
                preview_df = read_dataframe_from_uploaded_file(uploaded_file, selected_sheet=selected_sheet)
            else:
                st.session_state["selected_sheet"] = None

            st.session_state["input_df"] = preview_df

            st.write("Anteprima file:")
            st.dataframe(preview_df.head(20), use_container_width=True)

            asin_column = st.selectbox(
                "Seleziona la colonna contenente gli ASIN",
                options=list(preview_df.columns),
                index=0 if len(preview_df.columns) > 0 else None,
            )
            st.session_state["asin_column"] = asin_column

            deduplicate = st.checkbox(
                "Rimuovi ASIN duplicati",
                value=st.session_state.get("deduplicate_asins", True),
            )
            st.session_state["deduplicate_asins"] = deduplicate

            asins, duplicates_removed = extract_asins_from_df(
                preview_df,
                asin_column=asin_column,
                deduplicate=deduplicate,
            )
            st.session_state["asin_list"] = asins

            col1, col2, col3 = st.columns(3)
            col1.metric("ASIN validi", len(asins))
            col2.metric("Duplicati rimossi", duplicates_removed)
            col3.metric("File", st.session_state["uploaded_file_name"])

            if asins:
                st.write("Prime 20 occorrenze ASIN pulite:")
                st.code("\n".join(asins[:20]))
            else:
                st.warning("Nessun ASIN valido trovato nella colonna selezionata.")

        except Exception as exc:
            st.error(f"Errore nella lettura del file: {exc}")


def render_phase_3_extraction_params():
    st.header("Fase 3 · Parametri di estrazione")

    with st.container(border=True):
        col1, col2 = st.columns(2)

        with col1:
            amazon_domain = st.selectbox(
                "Marketplace Amazon",
                options=MARKETPLACES,
                index=MARKETPLACES.index(st.session_state.get("amazon_domain", "amazon.it")),
            )
            st.session_state["amazon_domain"] = amazon_domain

            delay_seconds = st.slider(
                "Delay tra le richieste (secondi)",
                min_value=0.0,
                max_value=5.0,
                value=float(st.session_state.get("delay_seconds", 1.0)),
                step=0.5,
            )
            st.session_state["delay_seconds"] = delay_seconds

        with col2:
            st.session_state["show_logs"] = st.checkbox(
                "Mostra log dettagliato",
                value=st.session_state.get("show_logs", True),
            )
            st.session_state["save_raw_json"] = st.checkbox(
                "Salva risposta JSON grezza nel file finale",
                value=st.session_state.get("save_raw_json", False),
            )
            st.session_state["show_image_preview"] = st.checkbox(
                "Mostra anteprima immagini nei risultati",
                value=st.session_state.get("show_image_preview", True),
            )

        st.subheader("Campi da estrarre")

        default_fields = st.session_state.get("selected_fields", STANDARD_FIELDS.copy())
        selected_fields = []

        cols = st.columns(3)
        for idx, field in enumerate(STANDARD_FIELDS):
            with cols[idx % 3]:
                checked = st.checkbox(
                    FIELD_LABELS[field],
                    value=field in default_fields,
                    key=f"field_{field}",
                )
                if checked:
                    selected_fields.append(field)

        if st.button("Seleziona tutti i campi standard"):
            st.session_state["selected_fields"] = STANDARD_FIELDS.copy()
            st.rerun()

        st.session_state["selected_fields"] = selected_fields

        if not selected_fields:
            st.warning("Seleziona almeno un campo da estrarre.")
        else:
            st.success(f"Campi selezionati: {len(selected_fields)}")


def render_phase_4_extraction():
    st.header("Fase 4 · Estrazione e preparazione file")

    with st.container(border=True):
        api_key = st.session_state.get("serpapi_key", "").strip()
        input_df = st.session_state.get("input_df")
        asins = st.session_state.get("asin_list", [])
        selected_fields = st.session_state.get("selected_fields", [])
        amazon_domain = st.session_state.get("amazon_domain", "amazon.it")
        delay_seconds = st.session_state.get("delay_seconds", 1.0)
        show_logs = st.session_state.get("show_logs", True)
        save_raw_json = st.session_state.get("save_raw_json", False)

        checks = {
            "API key configurata": bool(api_key),
            "File caricato": input_df is not None,
            "ASIN disponibili": len(asins) > 0,
            "Campi selezionati": len(selected_fields) > 0,
        }

        for label, ok in checks.items():
            if ok:
                st.success(label)
            else:
                st.error(label)

        if st.button("🚀 Avvia estrazione", use_container_width=True):
            if not all(checks.values()):
                st.warning("Completa prima tutte le fasi precedenti.")
                return

            try:
                results_df, raw_results = process_asin_list(
                    asins=asins,
                    api_key=api_key,
                    amazon_domain=amazon_domain,
                    selected_fields=selected_fields,
                    delay_seconds=delay_seconds,
                    show_logs=show_logs,
                    save_raw_json=save_raw_json,
                )

                st.session_state["results_df"] = results_df
                st.session_state["raw_results"] = raw_results
                st.session_state["run_completed"] = True

                ok_count = int((results_df["Status"] == "OK").sum()) if "Status" in results_df.columns else 0
                err_count = int((results_df["Status"] == "ERROR").sum()) if "Status" in results_df.columns else 0

                st.success(f"Estrazione completata. OK: {ok_count} · ERROR: {err_count}")
                st.dataframe(results_df.head(20), use_container_width=True)

            except Exception as exc:
                st.error(f"Errore durante l’estrazione: {exc}")


def render_phase_5_download():
    st.header("Fase 5 · Preparazione file e download")

    with st.container(border=True):
        df = st.session_state.get("results_df")

        if df is None or df.empty:
            st.info("Nessun file pronto. Esegui prima la fase di estrazione.")
            return

        st.write("Anteprima risultati finali:")
        st.dataframe(df, use_container_width=True)

        show_image_preview = st.session_state.get("show_image_preview", True)
        if show_image_preview and "Immagine principale" in df.columns:
            image_rows = df[df["Immagine principale"].astype(str).str.startswith("http", na=False)].head(12)
            if not image_rows.empty:
                st.subheader("Anteprima immagini")
                cols = st.columns(4)
                for idx, (_, row) in enumerate(image_rows.iterrows()):
                    with cols[idx % 4]:
                        st.image(row["Immagine principale"], caption=row.get("ASIN", ""), use_container_width=True)

        base_name = st.text_input("Nome base file output", value="amazon_asin_export")

        csv_bytes = dataframe_to_csv_bytes(df)
        xlsx_bytes = dataframe_to_excel_bytes(df, sheet_name="Amazon_ASIN")

        col1, col2 = st.columns(2)

        with col1:
            st.download_button(
                label="⬇️ Scarica CSV",
                data=csv_bytes,
                file_name=f"{base_name}.csv",
                mime="text/csv",
                use_container_width=True,
            )

        with col2:
            st.download_button(
                label="⬇️ Scarica XLSX",
                data=xlsx_bytes,
                file_name=f"{base_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


# =========================================================
# UI · TAB RICERCA DA KEYWORD
# =========================================================
def render_keyword_upload():
    st.header("Fase 1 · Upload lista keyword")

    with st.container(border=True):
        uploaded_file = st.file_uploader(
            "Carica un file XLSX o CSV con le parole chiave",
            type=["csv", "xlsx"],
            key="kw_file_uploader",
        )

        if uploaded_file is None:
            st.info("Carica un file per selezionare la colonna delle keyword.")
            return

        try:
            preview_df, sheet_names, file_type = load_input_file(uploaded_file)
            st.session_state["kw_uploaded_file_name"] = uploaded_file.name
            st.session_state["kw_sheet_names"] = sheet_names

            if file_type == "xlsx":
                selected_sheet = st.selectbox(
                    "Seleziona il foglio Excel",
                    options=sheet_names,
                    index=0,
                    key="kw_sheet_select",
                )
                st.session_state["kw_selected_sheet"] = selected_sheet
                preview_df = read_dataframe_from_uploaded_file(uploaded_file, selected_sheet=selected_sheet)
            else:
                st.session_state["kw_selected_sheet"] = None

            st.session_state["kw_input_df"] = preview_df

            st.write("Anteprima file:")
            st.dataframe(preview_df.head(20), use_container_width=True)

            keyword_column = st.selectbox(
                "Seleziona la colonna contenente le keyword",
                options=list(preview_df.columns),
                index=0 if len(preview_df.columns) > 0 else None,
                key="kw_column_select",
            )
            st.session_state["kw_column"] = keyword_column

            deduplicate = st.checkbox(
                "Rimuovi keyword duplicate",
                value=st.session_state.get("kw_deduplicate", True),
                key="kw_dedup_checkbox",
            )
            st.session_state["kw_deduplicate"] = deduplicate

            keywords, duplicates_removed = extract_keywords_from_df(
                preview_df,
                keyword_column=keyword_column,
                deduplicate=deduplicate,
            )
            st.session_state["kw_list"] = keywords

            col1, col2, col3 = st.columns(3)
            col1.metric("Keyword valide", len(keywords))
            col2.metric("Duplicati rimossi", duplicates_removed)
            col3.metric("Chiamate API stimate", len(keywords))

            if keywords:
                st.write("Prime 20 keyword pulite:")
                st.code("\n".join(keywords[:20]))
            else:
                st.warning("Nessuna keyword valida trovata nella colonna selezionata.")

        except Exception as exc:
            st.error(f"Errore nella lettura del file: {exc}")


def render_keyword_options():
    st.header("Fase 2 · Opzioni di ricerca")

    with st.container(border=True):
        col1, col2 = st.columns(2)

        with col1:
            amazon_domain = st.selectbox(
                "Marketplace Amazon",
                options=MARKETPLACES,
                index=MARKETPLACES.index(st.session_state.get("kw_amazon_domain", "amazon.it")),
                key="kw_marketplace_select",
            )
            st.session_state["kw_amazon_domain"] = amazon_domain

            sort_labels = list(SORT_OPTIONS.keys())
            current_sort = st.session_state.get("kw_sort", "")
            current_label = next(
                (label for label, value in SORT_OPTIONS.items() if value == current_sort),
                sort_labels[0],
            )
            sort_label = st.selectbox(
                "Ordinamento risultati",
                options=sort_labels,
                index=sort_labels.index(current_label),
                key="kw_sort_select",
            )
            st.session_state["kw_sort"] = SORT_OPTIONS[sort_label]

            delay_seconds = st.slider(
                "Delay tra le richieste (secondi)",
                min_value=0.0,
                max_value=5.0,
                value=float(st.session_state.get("kw_delay_seconds", 1.0)),
                step=0.5,
                key="kw_delay_slider",
            )
            st.session_state["kw_delay_seconds"] = delay_seconds

        with col2:
            st.session_state["kw_include_sponsored"] = st.checkbox(
                "Includi risultati sponsorizzati (Sponsored / Product Ads)",
                value=st.session_state.get("kw_include_sponsored", True),
                help="I risultati sponsorizzati restano distinguibili grazie alla colonna «Sponsorizzato».",
                key="kw_sponsored_checkbox",
            )

            st.session_state["kw_max_results"] = int(
                st.number_input(
                    "Max risultati per keyword (0 = tutti quelli della prima pagina)",
                    min_value=0,
                    max_value=100,
                    value=int(st.session_state.get("kw_max_results", 0)),
                    step=1,
                    key="kw_max_results_input",
                )
            )

            st.session_state["kw_show_logs"] = st.checkbox(
                "Mostra log dettagliato",
                value=st.session_state.get("kw_show_logs", True),
                key="kw_logs_checkbox",
            )

        st.subheader("Colonne da esportare")

        default_fields = st.session_state.get("kw_selected_fields", SEARCH_FIELDS.copy())
        selected_fields = []

        cols = st.columns(3)
        for idx, field in enumerate(SEARCH_FIELDS):
            with cols[idx % 3]:
                checked = st.checkbox(
                    SEARCH_FIELD_LABELS[field],
                    value=field in default_fields,
                    key=f"kw_field_{field}",
                )
                if checked:
                    selected_fields.append(field)

        st.session_state["kw_selected_fields"] = selected_fields

        if not selected_fields:
            st.warning("Seleziona almeno una colonna da esportare.")
        else:
            st.success(f"Colonne selezionate: {len(selected_fields)}")


def render_keyword_extraction():
    st.header("Fase 3 · Esecuzione ricerca")

    with st.container(border=True):
        api_key = st.session_state.get("serpapi_key", "").strip()
        keywords = st.session_state.get("kw_list", [])
        selected_fields = st.session_state.get("kw_selected_fields", [])

        checks = {
            "API key configurata": bool(api_key),
            "Keyword disponibili": len(keywords) > 0,
            "Colonne selezionate": len(selected_fields) > 0,
        }

        for label, ok in checks.items():
            if ok:
                st.success(label)
            else:
                st.error(label)

        if st.button("🚀 Avvia ricerca da keyword", use_container_width=True, key="kw_run_button"):
            if not all(checks.values()):
                st.warning("Completa prima tutte le fasi precedenti.")
                return

            try:
                results_df = process_keyword_list(
                    keywords=keywords,
                    api_key=api_key,
                    amazon_domain=st.session_state.get("kw_amazon_domain", "amazon.it"),
                    selected_fields=selected_fields,
                    include_sponsored=st.session_state.get("kw_include_sponsored", True),
                    sort_by=st.session_state.get("kw_sort", ""),
                    max_results_per_keyword=st.session_state.get("kw_max_results", 0),
                    delay_seconds=st.session_state.get("kw_delay_seconds", 1.0),
                    show_logs=st.session_state.get("kw_show_logs", True),
                )

                st.session_state["kw_results_df"] = results_df
                st.session_state["kw_run_completed"] = True

                ok_count = int((results_df["Status"] == "OK").sum()) if "Status" in results_df.columns else 0
                err_count = int((results_df["Status"] == "ERROR").sum()) if "Status" in results_df.columns else 0

                st.success(
                    f"Ricerca completata. Righe prodotto: {ok_count} · Keyword in errore: {err_count}"
                )
                st.dataframe(results_df.head(30), use_container_width=True)

            except Exception as exc:
                st.error(f"Errore durante la ricerca: {exc}")


def render_keyword_download():
    st.header("Fase 4 · Preparazione file e download")

    with st.container(border=True):
        df = st.session_state.get("kw_results_df")

        if df is None or df.empty:
            st.info("Nessun file pronto. Esegui prima la ricerca.")
            return

        col1, col2, col3 = st.columns(3)
        col1.metric("Righe totali", len(df))
        col2.metric("Keyword elaborate", df["Keyword"].nunique() if "Keyword" in df.columns else 0)
        if "Sponsorizzato" in df.columns:
            col3.metric("Risultati sponsorizzati", int((df["Sponsorizzato"] == "Sì").sum()))

        st.write("Anteprima risultati:")
        st.dataframe(df, use_container_width=True)

        base_name = st.text_input(
            "Nome base file output",
            value="amazon_keyword_search",
            key="kw_base_name",
        )

        csv_bytes = dataframe_to_csv_bytes(df)
        xlsx_bytes = dataframe_to_excel_bytes(df, sheet_name="Amazon_Keyword")

        col_a, col_b = st.columns(2)

        with col_a:
            st.download_button(
                label="⬇️ Scarica CSV",
                data=csv_bytes,
                file_name=f"{base_name}.csv",
                mime="text/csv",
                use_container_width=True,
                key="kw_download_csv",
            )

        with col_b:
            st.download_button(
                label="⬇️ Scarica XLSX",
                data=xlsx_bytes,
                file_name=f"{base_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="kw_download_xlsx",
            )


def render_keyword_search_tab():
    st.write(
        """
        Carica una lista di **parole chiave** e scarica i risultati di ricerca Amazon
        (ranking, ASIN, nome prodotto, brand, prezzo, rating, recensioni, disponibilità…)
        in un unico foglio CSV/XLSX.
        """
    )

    render_keyword_upload()
    st.divider()

    render_keyword_options()
    st.divider()

    render_keyword_extraction()
    st.divider()

    render_keyword_download()


def render_asin_tab():
    st.write(
        """
        Estrai i dati di dettaglio prodotto a partire da una lista di **ASIN**,
        con workflow verticale in 5 fasi e output finale in **CSV** e **XLSX**.
        """
    )

    render_phase_2_upload_input()
    st.divider()

    render_phase_3_extraction_params()
    st.divider()

    render_phase_4_extraction()
    st.divider()

    render_phase_5_download()


# =========================================================
# MAIN
# =========================================================
def main():
    st.set_page_config(
        page_title="Amazon Extractor",
        page_icon="🛒",
        layout="wide",
    )

    init_session_state()

    st.title("🛒 Amazon Extractor · SerpAPI")
    st.write(
        """
        Tool Streamlit per estrarre dati da Amazon tramite **SerpAPI**:
        dettaglio prodotto a partire da una lista di **ASIN** oppure
        risultati di ricerca a partire da una lista di **keyword**.
        """
    )

    st.sidebar.title("Stato sessione")
    st.sidebar.write(f"API key configurata: {'Sì' if st.session_state.get('serpapi_key') else 'No'}")
    st.sidebar.write(f"ASIN caricati: {len(st.session_state.get('asin_list', []))}")
    st.sidebar.write(f"Keyword caricate: {len(st.session_state.get('kw_list', []))}")
    st.sidebar.write(f"Risultati ASIN: {'Sì' if st.session_state.get('results_df') is not None else 'No'}")
    st.sidebar.write(f"Risultati keyword: {'Sì' if st.session_state.get('kw_results_df') is not None else 'No'}")

    render_phase_1_api_config()
    st.divider()

    tab_asin, tab_keyword = st.tabs(
        ["📦 Estrazione da ASIN", "🔎 Ricerca da keyword"]
    )

    with tab_asin:
        render_asin_tab()

    with tab_keyword:
        render_keyword_search_tab()


if __name__ == "__main__":
    main()