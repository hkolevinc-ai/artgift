#!/usr/bin/env python3
"""Scrape Art-Gift T-shirts and bodysuits into the supplied Temu bulk-upload workbook.

The script preserves the original workbook structure/styles by editing the XLSX XML
package directly. Each colour/size combination is exported as a separate SKU row and
all rows of one product share the same Contribution Goods identifier.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import html as html_lib
import json
import logging
import re
import shutil
import sys
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Iterator
from urllib.parse import unquote, urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from lxml import etree
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://art-gift.net/"
LOAD_MORE_URL = urljoin(BASE_URL, "bg/products/loadMore")
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"m": NS_MAIN, "r": NS_REL}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

DEFAULT_CONFIG: dict[str, Any] = {
    "template_path": "template/TEMU_ARTGIFT_TEMPLATE.xlsx",
    "output_dir": "output",
    "manifest_path": "",
    "product_offset": 0,
    "product_limit": 0,
    "output_prefix": "",
    "categories": [
        {"url": "https://art-gift.net/тениски", "slug": "тениски"},
        {"url": "https://art-gift.net/бодита", "slug": "бодита"},
    ],
    "request_delay_seconds": 0.8,
    "max_category_pages": 200,
    "max_products": 0,
    "rows_per_workbook": 1990,
    "quantity": 100,
    "shipping_template": "",
    "handling_time": "2 Days",
    "fulfillment_channel": "I will ship this item myself",
    "country_of_origin": "Bulgaria",
    "province_of_origin": "",
    "manufacturer": "Webshop Bulgaria Ltd",
    "eu_responsible_person": "",
    "item_tax_code": "",
    "product_guide_url": "",
    "default_bodysuit_category": 26402,
    "default_tshirt_category": 30469,
    "default_kids_tshirt_category": 30843,
    "exclude_complex_sets": True,
    "strict_validation": True,
    "exact_variants_with_browser": True,
    "browser_timeout_ms": 15000,
    "fallback_to_colour_size_cross_product": True,
    "force_customizable": True,
    "packaging": {
        "tshirt": {"weight_g": 250, "length_cm": 30, "width_cm": 25, "height_cm": 3},
        "bodysuit": {"weight_g": 180, "length_cm": 25, "width_cm": 20, "height_cm": 3},
    },
}

COLOR_MAP = {
    "white": "White", "black": "Black", "red": "Red", "yellow": "Yellow",
    "green": "Green", "light green": "Light Green", "dark green": "Dark Green",
    "blue": "Blue", "light blue": "Light Blue", "navy blue": "Navy Blue",
    "grey": "Grey", "gray": "Grey", "light grey": "Light Grey", "dark grey": "Dark Grey",
    "pink": "Pink", "light pink": "Light Pink", "orange": "Orange", "purple": "Purple",
    "brown": "Brown", "beige": "Beige", "cream": "Cream", "burgundy": "Burgundy",
    "бял": "White", "бяла": "White", "бяло": "White",
    "черен": "Black", "черна": "Black", "черно": "Black",
    "червен": "Red", "червена": "Red", "червено": "Red",
    "жълт": "Yellow", "жълта": "Yellow", "жълто": "Yellow",
    "зелен": "Green", "зелена": "Green", "зелено": "Green",
    "светло зелен": "Light Green", "тъмно зелен": "Dark Green",
    "син": "Blue", "синя": "Blue", "синьо": "Blue",
    "светло син": "Light Blue", "светлосин": "Light Blue",
    "тъмно син": "Navy Blue", "тъмносин": "Navy Blue",
    "небесно син": "Light Blue", "кралско син": "Royal Blue",
    "сив": "Grey", "сива": "Grey", "сиво": "Grey",
    "светло сив": "Light Grey", "тъмно сив": "Dark Grey",
    "розов": "Pink", "розова": "Pink", "розово": "Pink",
    "светло розов": "Light Pink", "циклама": "Hot Pink",
    "оранжев": "Orange", "оранжева": "Orange", "оранжево": "Orange",
    "лилав": "Purple", "лилава": "Purple", "лилаво": "Purple",
    "виолетов": "Purple", "лавандула": "Lavender",
    "кафяв": "Brown", "кафява": "Brown", "кафяво": "Brown",
    "бежов": "Beige", "бежова": "Beige", "бежово": "Beige",
    "екрю": "Apricot", "крем": "Cream", "кремав": "Cream",
    "бордо": "Burgundy", "винен": "Burgundy",
    "ментa": "Mint Green", "мента": "Mint Green", "тюркоаз": "Turquoise",
    "златист": "Gold", "сребрист": "Silver",
}


CATEGORY_NAMES: dict[int, str] = {
    26325: "Baby Products / Apparel & Accessories / Baby Girls / Clothing / Bodysuits",
    26402: "Baby Products / Apparel & Accessories / Baby Boys / Clothing / Bodysuits",
    29069: "Clothing, Shoes & Jewelry / Women / Clothing / Tops, Tees & Blouses / T-Shirts",
    29553: "Clothing, Shoes & Jewelry / Girls / Clothing / Tops, Tees & Blouses / Tees",
    29848: "Clothing, Shoes & Jewelry / Baby / Baby Girls / Clothing / Tops / Tees",
    30006: "Clothing, Shoes & Jewelry / Baby / Baby Boys / Clothing / Tops / Tees",
    30469: "Clothing, Shoes & Jewelry / Men / Clothing / Shirts / T-Shirts",
    30843: "Clothing, Shoes & Jewelry / Boys / Clothing / Tops, Tees & Shirts / Tees",
}

ALPHA_SIZES = {"XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL", "3XL", "4XL", "5XL", "6XL"}

MONTH_SIZE_MAP = {
    "0-3": "0-3M", "3-6": "3-6M", "6-9": "6-9M", "9-12": "9-12M",
    "12-18": "12-18M", "18-24": "18-24M", "24-36": "2-3Y",
}

@dataclass
class Variant:
    size_raw: str
    color_raw: str
    variation_id: str = ""
    price_eur: float | None = None
    images: list[str] = field(default_factory=list)
    sleeve_raw: str = ""

@dataclass
class Product:
    url: str
    site_id: str
    code: str
    title: str
    description: str
    short_description: str
    category_id: int
    product_kind: str
    size_title: str
    images: list[str]
    base_price_eur: float
    customizable: bool
    customization_text: str
    variants: list[Variant]


class ExcludedProductError(RuntimeError):
    def __init__(self, *, url: str, title: str, code: str, reason: str):
        super().__init__(reason)
        self.url = url
        self.title = title
        self.code = code
        self.reason = reason


def deep_merge(base: dict[str, Any], overlay: dict[str, Any]) -> dict[str, Any]:
    result = copy.deepcopy(base)
    for key, value in overlay.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def normalize_space(value: str | None) -> str:
    return re.sub(r"\s+", " ", html_lib.unescape(value or "")).strip()


def safe_slug(value: str, limit: int = 45) -> str:
    translit = str.maketrans({
        "а":"a","б":"b","в":"v","г":"g","д":"d","е":"e","ж":"zh","з":"z","и":"i","й":"y",
        "к":"k","л":"l","м":"m","н":"n","о":"o","п":"p","р":"r","с":"s","т":"t","у":"u",
        "ф":"f","х":"h","ц":"ts","ч":"ch","ш":"sh","щ":"sht","ъ":"a","ь":"","ю":"yu","я":"ya",
        "А":"A","Б":"B","В":"V","Г":"G","Д":"D","Е":"E","Ж":"Zh","З":"Z","И":"I","Й":"Y",
        "К":"K","Л":"L","М":"M","Н":"N","О":"O","П":"P","Р":"R","С":"S","Т":"T","У":"U",
        "Ф":"F","Х":"H","Ц":"Ts","Ч":"Ch","Ш":"Sh","Щ":"Sht","Ъ":"A","Ь":"","Ю":"Yu","Я":"Ya",
    })
    value = value.translate(translit).lower()
    value = re.sub(r"[^a-z0-9]+", "-", value).strip("-")
    return value[:limit] or "value"


def make_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(total=5, connect=5, read=5, status=5, backoff_factor=1.0,
                  status_forcelist=(429, 500, 502, 503, 504), allowed_methods=("GET", "POST"))
    session.mount("https://", HTTPAdapter(max_retries=retry))
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/150 Safari/537.36",
        "Accept-Language": "bg-BG,bg;q=0.9,en;q=0.7",
    })
    return session


def fetch(session: requests.Session, url: str, *, delay: float, method: str = "GET", data: dict[str, Any] | None = None) -> str:
    response = session.request(method, url, data=data, timeout=45)
    response.raise_for_status()
    response.encoding = response.apparent_encoding or "utf-8"
    if delay:
        time.sleep(delay)
    return response.text


def extract_product_links(html: str) -> list[str]:
    soup = BeautifulSoup(html, "lxml")
    links: list[str] = []
    for a in soup.select(".products_list a.topoffer[href], .product_item a[href]"):
        href = urljoin(BASE_URL, a.get("href", ""))
        if href.endswith(".htm") and href not in links:
            links.append(href)
    return links


def discover_products(session: requests.Session, category: dict[str, str], cfg: dict[str, Any]) -> list[str]:
    slug = category["slug"]
    delay = float(cfg["request_delay_seconds"])
    first_html = fetch(session, category["url"], delay=delay)
    links = extract_product_links(first_html)
    seen = set(links)
    logging.info("%s: %d products on first page", slug, len(links))
    if cfg.get("max_products") and len(links) >= int(cfg["max_products"]):
        return links[: int(cfg["max_products"])]

    for page in range(2, int(cfg["max_category_pages"]) + 1):
        payload = {
            "page": str(page), "action": "index", "cat_link": slug, "category": slug,
            "subcategory": "", "subsubcategory": "", "subsubsubcategory": "", "named": "[]",
        }
        chunk = fetch(session, LOAD_MORE_URL, method="POST", data=payload, delay=delay)
        page_links = extract_product_links(chunk)
        new_links = [u for u in page_links if u not in seen]
        if not new_links:
            logging.info("%s: pagination stopped at page %d", slug, page)
            break
        links.extend(new_links)
        seen.update(new_links)
        logging.info("%s: page %d added %d products (total %d)", slug, page, len(new_links), len(links))
        if cfg.get("max_products") and len(links) >= int(cfg["max_products"]):
            return links[: int(cfg["max_products"])]
    return links


def parse_json_ld(soup: BeautifulSoup) -> list[Any]:
    values: list[Any] = []
    for script in soup.select('script[type="application/ld+json"]'):
        raw = script.string or script.get_text(" ", strip=True)
        try:
            values.append(json.loads(raw))
        except Exception:
            continue
    return values


def find_product_json(values: Iterable[Any]) -> dict[str, Any]:
    stack = list(values)
    while stack:
        value = stack.pop(0)
        if isinstance(value, dict):
            if value.get("@type") == "Product":
                return value
            stack.extend(value.values())
        elif isinstance(value, list):
            stack.extend(value)
    return {}


def parse_price(product_json: dict[str, Any], soup: BeautifulSoup) -> float:
    inp = soup.select_one("#price, #base_price")
    if inp and inp.get("value"):
        try:
            return float(str(inp["value"]).replace(",", "."))
        except ValueError:
            pass
    offers = product_json.get("offers", {})
    if isinstance(offers, list):
        offers = offers[0] if offers else {}
    try:
        return float(str(offers.get("price", "0")).replace(",", "."))
    except ValueError:
        return 0.0


def url_product_id(url: str) -> str:
    path = unquote(urlparse(url).path)
    match = re.search(r"-(\d+)\.htm$", path, re.I)
    return match.group(1) if match else ""


def product_identity_from_url(url: str) -> str:
    product_id = url_product_id(url)
    if product_id:
        return f"id:{product_id}"
    canonical_path = unquote(urlparse(url).path).rstrip("/").casefold()
    return "url:" + hashlib.sha1(canonical_path.encode("utf-8")).hexdigest()[:16]


def infer_kind_from_url(url: str, fallback: str) -> str:
    path = unquote(urlparse(url).path).lower()
    filename = path.rsplit("/", 1)[-1]
    has_body = "боди" in filename
    has_shirt = "тениск" in filename
    if has_body and not has_shirt:
        return "bodysuit"
    if has_shirt and not has_body:
        return "tshirt"
    if path.startswith("/бодита/"):
        return "bodysuit"
    if path.startswith("/тениски/"):
        return "tshirt"
    return fallback


def infer_category(url: str, title: str, size_title: str, product_kind: str, cfg: dict[str, Any]) -> int:
    path = unquote(urlparse(url).path).lower()
    text = normalize_space(" ".join([path, title, size_title])).lower()

    girl_terms = r"момич|girl|принцес|дъщер|красива|госпожиц|сладуран|кралица|фея|бебка"
    boy_terms = r"момч|boy|синът|син |юнак|принц(?!ес)"
    female_adult_terms = r"дамск|женск|women|woman|майка|мама|баба|съпруга|леля|сестра|кралица"
    male_adult_terms = r"мъжк|men|man|баща|татко|дядо|съпруг|крал|лов|риболов"

    if product_kind == "bodysuit":
        if re.search(girl_terms, text):
            return 26325
        return int(cfg["default_bodysuit_category"])

    # URL subcategories are the strongest signal.
    if "/дамски-тениски/" in path:
        return 29069
    if "/мъжки-тениски/" in path:
        return 30469
    if "/детски-тениски/" in path:
        if re.search(girl_terms, text):
            return 29553
        if re.search(boy_terms, text):
            return 30843
        return int(cfg["default_kids_tshirt_category"])

    if re.search(r"бебешк.*тениск|тениск.*беб", text):
        return 29848 if re.search(girl_terms, text) else 30006
    if re.search(girl_terms, text):
        return 29553 if re.search(r"детск|дете|момич", text) else 29069
    if re.search(boy_terms, text):
        return 30843 if re.search(r"детск|дете|момч", text) else 30469
    if re.search(female_adult_terms, text):
        return 29069
    if re.search(male_adult_terms, text):
        return 30469

    # Generic/unisex adult T-shirts must not default to a boys category.
    return int(cfg.get("default_tshirt_category", 30469))


def looks_like_size(raw: str) -> bool:
    value = normalize_space(raw).upper().replace("–", "-").replace("—", "-")
    compact = re.sub(r"\s+", "", value)
    if compact in ALPHA_SIZES:
        return True
    if re.search(r"\b\d{2,3}\s*(?:CM|СМ)\b", value):
        return True
    if re.search(r"\b\d{1,2}\s*-\s*\d{1,2}\s*(?:М|M|Г|Y)", value):
        return True
    if re.search(r"\b(?:0|1|2)\s*(?:МЕС|М\.|ГОД|Г\.)", value):
        return True
    return False


def looks_like_sleeve(raw: str) -> bool:
    value = normalize_space(raw).lower()
    return any(token in value for token in ("ръкав", "къс", "дълъг", "short sleeve", "long sleeve"))


def looks_like_color(raw: str) -> bool:
    cleaned = normalize_space(raw).lower().replace("цвят", "").strip(" :-")
    if cleaned in COLOR_MAP:
        return True
    return any(key in cleaned for key in COLOR_MAP)


def classify_option_values(values: list[str]) -> str:
    cleaned = [normalize_space(v) for v in values if normalize_space(v)]
    if not cleaned:
        return "unknown"
    size_score = sum(looks_like_size(v) for v in cleaned)
    sleeve_score = sum(looks_like_sleeve(v) for v in cleaned)
    color_score = sum(looks_like_color(v) for v in cleaned)
    threshold = max(1, (len(cleaned) + 1) // 2)
    if sleeve_score >= threshold:
        return "sleeve"
    if size_score >= threshold:
        return "size"
    if color_score >= threshold:
        return "color"
    return "unknown"


def is_complex_set(soup: BeautifulSoup, title: str, url: str) -> tuple[bool, str]:
    text = normalize_space(" ".join([unquote(urlparse(url).path), title])).lower()
    if "комплект" in text:
        return True, "Complex apparel set with two or more independently selected garment sizes"

    labels: list[str] = []
    for node in soup.select(".custom_field_title, .main_filter_title"):
        label = normalize_space(node.get_text(" ", strip=True)).lower()
        if label and label not in labels:
            labels.append(label)
    garment_labels = [
        label for label in labels
        if any(token in label for token in ("мъжка тениска", "дамска тениска", "детска тениска", "боди"))
    ]
    if len(garment_labels) >= 2:
        return True, "Multiple independent garment selectors detected"
    return False, ""


def map_color(raw: str) -> str:
    cleaned = normalize_space(raw).lower().replace("цвят", "").strip(" :-")
    if cleaned in COLOR_MAP:
        return COLOR_MAP[cleaned]
    for key in sorted(COLOR_MAP, key=len, reverse=True):
        if key in cleaned:
            return COLOR_MAP[key]
    logging.warning("Unknown colour '%s'; exporting normalized source value", raw)
    return normalize_space(raw)


def map_size(raw: str) -> tuple[str, str, str]:
    value = normalize_space(raw).upper().replace("ГОДИНИ", "Г.").replace("ГОД.", "Г.")
    value = value.replace("МЕСЕЦА", "М.").replace("МЕС.", "М.")
    value = value.replace("–", "-").replace("—", "-")
    alpha = re.sub(r"\s+", "", value)
    if alpha in {"XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL", "3XL", "4XL", "5XL", "6XL"}:
        return "2 - Regular Size", "10 - Alpha", "XXXL" if alpha == "3XL" else alpha
    months = re.search(r"(\d{1,2})\s*-\s*(\d{1,2})\s*М", value)
    if months:
        key = f"{months.group(1)}-{months.group(2)}"
        return "2 - Regular Size", "8 - Age", MONTH_SIZE_MAP.get(key, f"{key}M")
    years = re.search(r"(\d{1,2})\s*-\s*(\d{1,2})\s*(?:Г|Y)", value)
    if years:
        return "2 - Regular Size", "8 - Age", f"{years.group(1)}-{years.group(2)}Y"
    one_year = re.fullmatch(r"(\d{1,2})\s*(?:Г|Y).*", value)
    if one_year:
        return "2 - Regular Size", "8 - Age", f"{one_year.group(1)}Y"
    numeric = re.search(r"(\d{2,3})(?:\s*CM|\s*СМ)?", value)
    if numeric:
        return "2 - Regular Size", "7 - Numeric", numeric.group(1)
    return "2 - Regular Size", "10 - Alpha", normalize_space(raw)


def extract_customization(soup: BeautifulSoup) -> tuple[bool, str]:
    parts: list[str] = []
    for field in soup.select(".custom_field"):
        text = normalize_space(field.get_text(" ", strip=True))
        if text and text not in parts:
            parts.append(text)
    badges = [normalize_space(x.get_text(" ", strip=True)) for x in soup.select(".customization, .custom_label")]
    parts.extend(x for x in badges if x and x not in parts)
    full_text = normalize_space(soup.get_text(" ", strip=True)).lower()
    customizable = bool(parts) or any(k in full_text for k in ("име по избор", "персонализиран", "персонализирана"))
    return customizable, " | ".join(parts)[:1000]


def static_variants(soup: BeautifulSoup, base_price: float, product_kind: str) -> list[Variant]:
    """Parse visible variation selectors without assuming main/add roles.

    Art-Gift is not consistent across all product pages: on some pages the first
    selector contains sizes and the second contains colours, while on others the
    order is reversed. The role is therefore inferred from the option values.
    """
    def collect(selector: str) -> list[tuple[str, str]]:
        result: list[tuple[str, str]] = []
        for a in soup.select(selector):
            value = normalize_space(a.get("title") or a.get_text(" ", strip=True))
            item = (value, a.get("data-variation_id", ""))
            if value and value not in [x[0] for x in result]:
                result.append(item)
        return result

    main_options = collect("a.variations_products_links.main_filter")
    add_options = collect("a.variations_products_links.add_filter")
    main_role = classify_option_values([x[0] for x in main_options])
    add_role = classify_option_values([x[0] for x in add_options])

    if product_kind == "bodysuit":
        size_options: list[tuple[str, str]] = []
        sleeve_options: list[tuple[str, str]] = []
        for role, options in ((main_role, main_options), (add_role, add_options)):
            if role == "size" and not size_options:
                size_options = options
            elif role == "sleeve" and not sleeve_options:
                sleeve_options = options
        # Known storefront fallback if labels are sparse but selector classes are standard.
        if not size_options:
            size_options = add_options if add_options and add_role != "color" else []
        if not sleeve_options:
            sleeve_options = main_options if main_options and main_role != "size" else []
        sizes = size_options or [("One Size", "")]
        sleeves = sleeve_options or [("Standard Sleeve", "")]
        return [
            Variant(size, "White", sleeve_id or size_id, base_price, sleeve_raw=sleeve)
            for size, size_id in sizes
            for sleeve, sleeve_id in sleeves
        ]

    size_options: list[tuple[str, str]] = []
    color_options: list[tuple[str, str]] = []
    for role, options in ((main_role, main_options), (add_role, add_options)):
        if role == "size" and not size_options:
            size_options = options
        elif role == "color" and not color_options:
            color_options = options

    # Fallbacks are only used when value-based detection cannot decide.
    if not size_options:
        if main_options and main_role != "color":
            size_options = main_options
        elif add_options and add_role != "color":
            size_options = add_options
    if not color_options:
        if add_options and add_options is not size_options and add_role != "size":
            color_options = add_options
        elif main_options and main_options is not size_options and main_role != "size":
            color_options = main_options

    sizes = size_options or [("One Size", "")]
    colors = color_options or [("White", "")]
    return [
        Variant(size, color, size_id or color_id, base_price)
        for size, size_id in sizes
        for color, color_id in colors
    ]


def browser_variants(url: str, timeout_ms: int, base_price: float, product_kind: str) -> list[Variant]:
    """Enumerate exact storefront variations with role detection.

    Selector roles are inferred from the displayed values, so pages where Art-Gift
    swaps the order of size and colour are handled correctly.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError("Playwright is not installed") from exc

    def stable_price(page: Any) -> float:
        price_text = str(base_price)
        if page.locator("#price").count():
            last = None
            stable_reads = 0
            for _ in range(12):
                current = page.locator("#price").input_value().strip()
                if current and current == last:
                    stable_reads += 1
                    if stable_reads >= 2:
                        price_text = current
                        break
                else:
                    stable_reads = 0
                if current:
                    price_text = current
                last = current
                page.wait_for_timeout(250)
        try:
            return float(price_text.replace(",", "."))
        except Exception:
            return base_price

    def option_payload(page: Any, selector: str) -> list[dict[str, str]]:
        return page.locator(selector).evaluate_all(
            """els => els
                .filter(e => e.offsetParent !== null && !e.classList.contains('disabled'))
                .map(e => ({
                    value: (e.title || e.textContent || '').trim(),
                    id: e.dataset.variation_id || ''
                }))
                .filter(x => x.value)
                .filter((x,i,a) => a.findIndex(y => y.value === x.value) === i)"""
        )

    def selector_role(page: Any, selector: str) -> tuple[str, list[dict[str, str]]]:
        payload = option_payload(page, selector)
        return classify_option_values([x["value"] for x in payload]), payload

    def click_option(page: Any, selector: str, name: str) -> bool:
        is_active = page.evaluate(
            """([selector,name]) => {
                const e=[...document.querySelectorAll(selector)]
                    .find(x => (x.title||x.textContent||'').trim()===name);
                return !!(e && e.classList.contains('current_sell'));
            }""",
            [selector, name],
        )
        if is_active:
            return True
        clicked = page.evaluate(
            """([selector,name]) => {
                const e=[...document.querySelectorAll(selector)]
                    .find(x => (x.title||x.textContent||'').trim()===name);
                if(e){e.click(); return true;} return false;
            }""",
            [selector, name],
        )
        if not clicked:
            return False
        try:
            page.wait_for_function(
                """([selector,name]) => {
                    const e=[...document.querySelectorAll(selector)]
                        .find(x => (x.title||x.textContent||'').trim()===name);
                    return !!(e && e.classList.contains('current_sell'));
                }""",
                [selector, name],
                timeout=5000,
            )
        except Exception:
            page.wait_for_timeout(800)
        try:
            page.wait_for_load_state("networkidle", timeout=2500)
        except Exception:
            pass
        return True

    def images(page: Any) -> list[str]:
        return page.locator(".main_pic_container a.main_pic").evaluate_all(
            "els => [...new Set(els.map(e => e.href).filter(Boolean))]"
        )

    main_selector = "a.variations_products_links.main_filter"
    add_selector = "a.variations_products_links.add_filter"
    variants: dict[tuple[str, str, str], Variant] = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(locale="bg-BG")
        page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
        page.wait_for_selector("a.variations_products_links", timeout=timeout_ms)
        main_role, main_items = selector_role(page, main_selector)
        add_role, add_items = selector_role(page, add_selector)

        if product_kind == "bodysuit":
            size_selector = main_selector if main_role == "size" else add_selector if add_role == "size" else add_selector
            sleeve_selector = main_selector if main_role == "sleeve" else add_selector if add_role == "sleeve" else main_selector
            size_items = option_payload(page, size_selector) or [{"value": "One Size", "id": ""}]

            for size_item in size_items:
                size_name = size_item["value"]
                page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
                page.wait_for_selector("a.variations_products_links", timeout=timeout_ms)
                if size_name != "One Size" and not click_option(page, size_selector, size_name):
                    continue
                sleeve_items = option_payload(page, sleeve_selector) or [{"value": "Standard Sleeve", "id": ""}]
                for sleeve_item in sleeve_items:
                    sleeve_name = sleeve_item["value"]
                    if sleeve_name != "Standard Sleeve" and not click_option(page, sleeve_selector, sleeve_name):
                        continue
                    variation_id = (
                        page.locator("#variation_id").input_value()
                        if page.locator("#variation_id").count()
                        else sleeve_item.get("id") or size_item.get("id") or ""
                    )
                    variants[(size_name, "White", sleeve_name)] = Variant(
                        size_name,
                        "White",
                        variation_id,
                        stable_price(page),
                        images(page),
                        sleeve_raw=sleeve_name,
                    )
        else:
            size_selector = main_selector if main_role == "size" else add_selector if add_role == "size" else main_selector
            color_selector = main_selector if main_role == "color" else add_selector if add_role == "color" else ""
            color_items = option_payload(page, color_selector) if color_selector else []
            if not color_items:
                color_items = [{"value": "White", "id": ""}]

            # Reload before each colour because the site replaces selector elements dynamically.
            for color_item in color_items:
                color_name = color_item["value"]
                page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
                page.wait_for_selector("a.variations_products_links", timeout=timeout_ms)
                if color_selector and color_name != "White" and not click_option(page, color_selector, color_name):
                    logging.warning("Could not activate colour %s on %s", color_name, url)
                    continue
                size_items = option_payload(page, size_selector)
                if not size_items:
                    current_id = (
                        page.locator("#variation_id").input_value()
                        if page.locator("#variation_id").count()
                        else ""
                    )
                    size_items = [{"value": "One Size", "id": current_id}]
                price = stable_price(page)
                image_urls = images(page)
                for size_item in size_items:
                    size_name = size_item["value"]
                    variants[(size_name, color_name, "")] = Variant(
                        size_name,
                        color_name,
                        size_item.get("id") or color_item.get("id") or "",
                        price,
                        image_urls,
                    )
        browser.close()
    return list(variants.values())


def _variant_key(variant: Variant, product_kind: str) -> tuple[str, str, str]:
    size = map_size(variant.size_raw)[2].casefold()
    color = "white" if product_kind == "bodysuit" else map_color(variant.color_raw).casefold()
    sleeve = normalize_sleeve(variant.sleeve_raw)[0] if product_kind == "bodysuit" else ""
    return size, color, sleeve


def merge_exact_with_static(static: list[Variant], exact: list[Variant], product_kind: str) -> list[Variant]:
    """Keep the complete static option matrix and enrich it with browser values.

    The browser is preferred for exact price, variation ID and colour-specific images.
    If the dynamic storefront returns an incomplete set, no colour/size combination is
    silently discarded.
    """
    exact_by_key = {_variant_key(v, product_kind): v for v in exact}
    merged: list[Variant] = []
    for fallback in static:
        resolved = exact_by_key.get(_variant_key(fallback, product_kind))
        merged.append(resolved if resolved is not None else fallback)
    static_keys = {_variant_key(v, product_kind) for v in static}
    merged.extend(v for v in exact if _variant_key(v, product_kind) not in static_keys)
    return merged


def distinct_variant_values(variants: list[Variant], attr: str) -> set[str]:
    return {
        normalize_space(getattr(v, attr, "")).casefold()
        for v in variants
        if normalize_space(getattr(v, attr, ""))
    }


def parse_product(html: str, url: str, product_kind: str, cfg: dict[str, Any]) -> Product:
    soup = BeautifulSoup(html, "lxml")
    product_json = find_product_json(parse_json_ld(soup))
    title_node = soup.select_one("h1.product_title_view") or soup.title
    title = normalize_space(title_node.get_text(" ", strip=True) if title_node else "")
    code_node = soup.select_one(".custom_code")
    code = normalize_space(code_node.get_text(" ", strip=True) if code_node else str(product_json.get("sku", "")))
    site_id_node = soup.select_one("#product_id")
    site_id = str(site_id_node.get("value", "")).strip() if site_id_node else ""

    if cfg.get("exclude_complex_sets", True):
        excluded, reason = is_complex_set(soup, title, url)
        if excluded:
            raise ExcludedProductError(url=url, title=title, code=code, reason=reason)

    product_kind = infer_kind_from_url(url, product_kind)
    short_node = soup.select_one(".short_description")
    additional_node = soup.select_one(".additional_description")
    short = normalize_space(short_node.get_text(" ", strip=True) if short_node else str(product_json.get("description", "")))
    additional = normalize_space(additional_node.get_text(" ", strip=True) if additional_node else "")
    description = normalize_space(" ".join(x for x in (short, additional) if x))[:2000]

    images: list[str] = []
    for a in soup.select(".main_pic_container a.main_pic[href]"):
        image = urljoin(BASE_URL, a.get("href", ""))
        if image and image not in images:
            images.append(image)
    if not images and product_json.get("image"):
        values = product_json["image"] if isinstance(product_json["image"], list) else [product_json["image"]]
        for value in values:
            image = urljoin(BASE_URL, str(value))
            if image and image not in images:
                images.append(image)

    price = parse_price(product_json, soup)
    size_title_node = soup.select_one(".main_filter_title")
    size_title = normalize_space(size_title_node.get_text(" ", strip=True) if size_title_node else "")
    customizable, customization_text = extract_customization(soup)
    category_id = infer_category(url, title, size_title, product_kind, cfg)
    variants = static_variants(soup, price, product_kind)
    return Product(
        url=url,
        site_id=site_id,
        code=code or site_id or url_product_id(url),
        title=title,
        description=description,
        short_description=short,
        category_id=category_id,
        product_kind=product_kind,
        size_title=size_title,
        images=images[:10],
        base_price_eur=price,
        customizable=customizable,
        customization_text=customization_text,
        variants=variants,
    )


def normalize_sleeve(raw: str) -> tuple[str, str]:
    text = normalize_space(raw).lower()
    if "дъл" in text or "long" in text:
        return "long-sleeve", "Дълъг ръкав"
    if "къс" in text or "short" in text:
        return "short-sleeve", "Къс ръкав"
    return safe_slug(raw or "standard-sleeve", 20), normalize_space(raw or "Стандартен ръкав")


def product_rows(product: Product, cfg: dict[str, Any]) -> list[dict[str, Any]]:
    url_id = url_product_id(product.url)
    site_part = safe_slug(product.site_id or "site", 18)
    url_part = safe_slug(url_id or hashlib.sha1(product.url.encode("utf-8")).hexdigest()[:10], 18)
    code_part = safe_slug(product.code or "product", 28)
    base_parent_code = f"AG-{site_part}-{url_part}-{code_part}"[:78]
    pack = cfg["packaging"][product.product_kind]
    custom_detail = product.customization_text or "Buyer can provide the personalization requested on the product page."
    bullet_custom = f"Customizable product: {custom_detail}"[:500]
    mark_customizable = bool(cfg.get("force_customizable", True) or product.customizable)
    category_name = CATEGORY_NAMES.get(product.category_id, "")
    rows: list[dict[str, Any]] = []
    seen_skus: set[str] = set()
    seen_variant_combinations: set[tuple[str, str, str]] = set()

    for idx, variant in enumerate(product.variants, start=1):
        if product.product_kind == "bodysuit":
            sleeve_slug, sleeve_label = normalize_sleeve(variant.sleeve_raw)
            parent_code = f"{base_parent_code}-{sleeve_slug}"[:90]
            product_name = f"{product.title} – {sleeve_label}"[:500]
            size_family, sub_size_family, size = map_size(variant.size_raw)
            color = "White"
            sleeve_length = "Long Sleeve" if sleeve_slug == "long-sleeve" else "Short Sleeve" if sleeve_slug == "short-sleeve" else ""
            variation_bullet = "Each garment size is a separate SKU; sleeve type is grouped as a separate Temu product."
        else:
            parent_code = base_parent_code
            product_name = product.title[:500]
            size_family, sub_size_family, size = map_size(variant.size_raw)
            color = map_color(variant.color_raw)
            sleeve_length = ""
            variation_bullet = "Each colour and size is a separate SKU"

        combination_key = (parent_code, size.casefold(), color.casefold())
        if combination_key in seen_variant_combinations:
            logging.warning(
                "Duplicate mapped variant skipped for %s: parent=%s size=%s color=%s",
                product.url, parent_code, size, color,
            )
            continue
        seen_variant_combinations.add(combination_key)

        raw_variant_key = "|".join([
            product.url,
            parent_code,
            normalize_space(variant.sleeve_raw),
            color,
            size,
        ])
        variant_hash = hashlib.sha1(raw_variant_key.encode("utf-8")).hexdigest()[:8]
        suffix = f"{safe_slug(color,14)}-{safe_slug(size,14)}-{variant_hash}"
        sku = f"{parent_code}-{suffix}"[:90]
        if sku in seen_skus:
            sku = f"{sku[:81]}-{idx:04d}"
        seen_skus.add(sku)

        price = variant.price_eur if variant.price_eur is not None else product.base_price_eur
        images = (variant.images or product.images)[:10]
        detail_images = product.images[:10]

        rows.append({
            "category": product.category_id,
            "category_name": category_name,
            "product_type": "Custom product" if mark_customizable else "Normal product",
            "customization_mode": "Single technique" if mark_customizable else "",
            "primary_technique": "Leather/fabric customization technique" if mark_customizable else "",
            "secondary_technique": "Digital printing" if mark_customizable else "",
            "product_name": product_name,
            "parent_sku": parent_code,
            "sku": sku,
            "operation": "Add",
            "description": product.description,
            "bullet_points": [bullet_custom, "Material: 100% cotton", "Direct digital print", variation_bullet],
            "detail_images": detail_images,
            "variation_theme": "Color × Size",
            "size_family": size_family,
            "sub_size_family": sub_size_family,
            "size": size,
            "color": color,
            "sleeve_length": sleeve_length,
            "sku_images": images,
            "quantity": int(cfg["quantity"]),
            "base_price_eur": round(float(price), 2),
            "reference_link": product.url,
            "list_price_na": "N/A",
            "weight_g": pack["weight_g"], "length_cm": pack["length_cm"],
            "width_cm": pack["width_cm"], "height_cm": pack["height_cm"],
            "sku_type": "Single set", "individually_packed": "Yes",
            "packaging_quantity": 1, "packaging_unit": "piece",
            "item_quantity": 1, "item_unit": "piece",
            "shipping_template": cfg["shipping_template"], "handling_time": cfg["handling_time"],
            "fulfillment_channel": cfg["fulfillment_channel"], "item_tax_code": cfg["item_tax_code"],
            "country_origin": cfg["country_of_origin"], "province_origin": cfg["province_of_origin"],
            "product_guide": cfg["product_guide_url"],
            "product_identification": sku,
            "manufacturer": cfg["manufacturer"], "eu_responsible_person": cfg["eu_responsible_person"],
        })
    return rows


def audit_rows(rows: list[dict[str, Any]]) -> dict[str, Any]:
    sku_seen: dict[str, str] = {}
    parent_to_urls: dict[str, set[str]] = {}
    duplicate_skus: list[dict[str, str]] = []
    parent_collisions: list[dict[str, Any]] = []
    size_color_issues: list[dict[str, str]] = []
    duplicate_variant_combinations: list[dict[str, str]] = []
    variant_combinations_seen: dict[tuple[str, str, str], str] = {}
    set_rows: list[dict[str, str]] = []
    missing_category_names = 0

    for row in rows:
        sku = str(row.get("sku") or "")
        url = str(row.get("reference_link") or "")
        parent = str(row.get("parent_sku") or "")
        title = str(row.get("product_name") or "")
        size = str(row.get("size") or "")
        color = str(row.get("color") or "")
        if sku in sku_seen and sku_seen[sku] != url:
            duplicate_skus.append({"sku": sku, "first_url": sku_seen[sku], "second_url": url})
        else:
            sku_seen[sku] = url
        parent_to_urls.setdefault(parent, set()).add(url)
        combination = (parent, size.casefold(), color.casefold())
        if combination in variant_combinations_seen:
            duplicate_variant_combinations.append({
                "first_sku": variant_combinations_seen[combination],
                "second_sku": sku,
                "parent_sku": parent,
                "size": size,
                "color": color,
                "url": url,
            })
        else:
            variant_combinations_seen[combination] = sku
        if size.casefold() == "one size" and looks_like_size(color):
            size_color_issues.append({"sku": sku, "size": size, "color": color, "url": url})
        if "комплект" in normalize_space(title).lower():
            set_rows.append({"sku": sku, "title": title, "url": url})
        if not row.get("category_name"):
            missing_category_names += 1

    for parent, urls in parent_to_urls.items():
        if len(urls) > 1:
            parent_collisions.append({"parent_sku": parent, "urls": sorted(urls)})

    return {
        "row_count": len(rows),
        "unique_skus": len(sku_seen),
        "duplicate_skus": duplicate_skus,
        "parent_collisions": parent_collisions,
        "size_color_issues": size_color_issues,
        "duplicate_variant_combinations": duplicate_variant_combinations,
        "set_rows": set_rows,
        "missing_category_names": missing_category_names,
        "passed": not any((duplicate_skus, parent_collisions, size_color_issues, duplicate_variant_combinations, set_rows)) and missing_category_names == 0,
    }


def col_to_num(col: str) -> int:
    value = 0
    for char in col:
        value = value * 26 + ord(char) - 64
    return value


def num_to_col(value: int) -> str:
    result = ""
    while value:
        value, rem = divmod(value - 1, 26)
        result = chr(65 + rem) + result
    return result


def cell_col(ref: str) -> int:
    match = re.match(r"([A-Z]+)", ref)
    return col_to_num(match.group(1)) if match else 0


class TemuWorkbookWriter:
    TARGETS = {
        "category": ("t_1_Category", 0),
        "category_name": ("t_1_Category Name", 0),
        "product_type": ("t_1_Product type", 0),
        "customization_mode": ("t_1_Customization processing technique", 0),
        "primary_technique": ("t_1_Primary technique", 0),
        "secondary_technique": ("t_1_Secondary technique", 0),
        "product_name": ("t_1_Product Name", 0),
        "parent_sku": ("t_1_Contribution Goods", 0),
        "sku": ("t_1_Contribution SKU", 0),
        "operation": ("t_1_Update or Add", 0),
        "description": ("t_2_Product Description", 0),
        "variation_theme": ("t_4_Variation Theme", 0),
        "size_family": ("t_4_Size Family", 0),
        "sub_size_family": ("t_4_Sub-Size Family", 0),
        "size": ("t_4_Size:3001", 0),
        "color": ("t_4_Sale Property:1001", 0),
        "sleeve_length": ("t_3_Property:29", 0),
        "quantity": ("t_6_Quantity", 0),
        "base_price_eur": ("t_6_Base Price - EUR", 0),
        "reference_link": ("t_6_Reference Link", 0),
        "list_price_na": ("t_6_Not available for List price", 0),
        "weight_g": ("t_6_Weight - g", 0),
        "length_cm": ("t_6_Length - cm", 0),
        "width_cm": ("t_6_Width - cm", 0),
        "height_cm": ("t_6_Height - cm", 0),
        "sku_type": ("t_6_SKU type", 0),
        "individually_packed": ("t_6_Individually packed", 0),
        "packaging_quantity": ("t_6_Total packaging quantity", 0),
        "packaging_unit": ("t_6_Packaging unit", 0),
        "item_quantity": ("t_6_Total item quantity", 0),
        "item_unit": ("t_6_Item unit", 0),
        "shipping_template": ("t_7_Shipping Template", 0),
        "handling_time": ("t_7_Handling Time", 0),
        "fulfillment_channel": ("t_7_Fulfillment Channel", 0),
        "item_tax_code": ("t_7_Item Tax Code", 0),
        "country_origin": ("t_8_Country/Region of Origin", 0),
        "province_origin": ("t_8_Province of Origin", 0),
        "product_guide": ("t_8_Product Guide", 0),
        "product_identification": ("t_8_Governance Property:1100100115", 0),
        "manufacturer": ("t_8_Governance Property:3", 0),
        "eu_responsible_person": ("t_8_Governance Property:2", 0),
    }

    def __init__(self, template: Path):
        self.template = template
        with zipfile.ZipFile(template) as zf:
            self.files = {name: zf.read(name) for name in zf.namelist()}
        self.shared_strings = self._load_shared_strings()
        self.sheet_path = self._find_sheet_path("Template")
        self.sheet_root = etree.fromstring(self.files[self.sheet_path])
        self.header_cols = self._read_headers(4)
        self.style_by_col = self._styles_from_row(5)

    def _load_shared_strings(self) -> list[str]:
        raw = self.files.get("xl/sharedStrings.xml")
        if not raw:
            return []
        root = etree.fromstring(raw)
        values = []
        for si in root.findall(f"{{{NS_MAIN}}}si"):
            values.append("".join(si.itertext()))
        return values

    def _find_sheet_path(self, sheet_name: str) -> str:
        workbook = etree.fromstring(self.files["xl/workbook.xml"])
        rel_id = None
        for sheet in workbook.findall(".//m:sheet", NS):
            if sheet.get("name") == sheet_name:
                rel_id = sheet.get(f"{{{NS_REL}}}id")
                break
        if not rel_id:
            raise ValueError(f"Worksheet '{sheet_name}' not found")
        rels = etree.fromstring(self.files["xl/_rels/workbook.xml.rels"])
        target = None
        for rel in rels.findall(f"{{{NS_PKG_REL}}}Relationship"):
            if rel.get("Id") == rel_id:
                target = rel.get("Target")
                break
        if not target:
            raise ValueError(f"Worksheet relationship for '{sheet_name}' not found")
        target = target.lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        return str(Path(target).as_posix())

    def _cell_value(self, cell: etree._Element) -> str:
        cell_type = cell.get("t")
        if cell_type == "inlineStr":
            return "".join(cell.itertext())
        value = cell.find(f"{{{NS_MAIN}}}v")
        if value is None or value.text is None:
            return ""
        if cell_type == "s":
            try:
                return self.shared_strings[int(value.text)]
            except Exception:
                return ""
        return value.text

    def _row(self, row_number: int, create: bool = False) -> etree._Element | None:
        sheet_data = self.sheet_root.find("m:sheetData", NS)
        if sheet_data is None:
            raise ValueError("Invalid worksheet: sheetData missing")
        row = sheet_data.find(f'm:row[@r="{row_number}"]', NS)
        if row is None and create:
            row = etree.Element(f"{{{NS_MAIN}}}row", r=str(row_number))
            inserted = False
            for existing in sheet_data.findall("m:row", NS):
                if int(existing.get("r", "0")) > row_number:
                    existing.addprevious(row)
                    inserted = True
                    break
            if not inserted:
                sheet_data.append(row)
        return row

    def _read_headers(self, row_number: int) -> dict[str, list[int]]:
        row = self._row(row_number)
        result: dict[str, list[int]] = {}
        if row is None:
            return result
        for cell in row.findall("m:c", NS):
            value = self._cell_value(cell)
            if value:
                result.setdefault(value, []).append(cell_col(cell.get("r", "")))
        return result

    def _styles_from_row(self, row_number: int) -> dict[int, str]:
        row = self._row(row_number)
        return {cell_col(c.get("r", "")): c.get("s", "") for c in row.findall("m:c", NS)} if row is not None else {}

    def _cell(self, row: etree._Element, col_num: int, row_num: int) -> etree._Element:
        ref = f"{num_to_col(col_num)}{row_num}"
        cell = row.find(f'm:c[@r="{ref}"]', NS)
        if cell is None:
            attrs = {"r": ref}
            if self.style_by_col.get(col_num):
                attrs["s"] = self.style_by_col[col_num]
            cell = etree.Element(f"{{{NS_MAIN}}}c", **attrs)
            row.append(cell)
            cells = sorted(row.findall("m:c", NS), key=lambda c: cell_col(c.get("r", "")))
            for c in row.findall("m:c", NS):
                row.remove(c)
            row.extend(cells)
        return cell

    def set_value(self, row_num: int, col_num: int, value: Any) -> None:
        if value is None or value == "":
            return
        row = self._row(row_num, create=True)
        assert row is not None
        cell = self._cell(row, col_num, row_num)
        for child in list(cell):
            cell.remove(child)
        if isinstance(value, bool):
            cell.set("t", "b")
            etree.SubElement(cell, f"{{{NS_MAIN}}}v").text = "1" if value else "0"
        elif isinstance(value, (int, float)) and not isinstance(value, bool):
            cell.attrib.pop("t", None)
            etree.SubElement(cell, f"{{{NS_MAIN}}}v").text = str(value)
        else:
            cell.set("t", "inlineStr")
            is_el = etree.SubElement(cell, f"{{{NS_MAIN}}}is")
            t = etree.SubElement(is_el, f"{{{NS_MAIN}}}t")
            t.set(XML_SPACE, "preserve")
            t.text = str(value)

    def col(self, machine_header: str, occurrence: int = 0) -> int:
        columns = self.header_cols.get(machine_header, [])
        if occurrence >= len(columns):
            raise KeyError(f"Header not found: {machine_header} occurrence {occurrence}")
        return columns[occurrence]

    def write_rows(self, rows: list[dict[str, Any]]) -> None:
        for offset, data in enumerate(rows, start=5):
            for key, (header, occurrence) in self.TARGETS.items():
                self.set_value(offset, self.col(header, occurrence), data.get(key))
            for i, text in enumerate(data.get("bullet_points", [])[:6]):
                self.set_value(offset, self.col("t_2_Bullet Point", i), text)
            for i, url in enumerate(data.get("detail_images", [])[:10]):
                self.set_value(offset, self.col("t_2_Detail Images URL", i), url)
            for i, url in enumerate(data.get("sku_images", [])[:10]):
                self.set_value(offset, self.col("t_6_SKU Images URL", i), url)
        dimension = self.sheet_root.find("m:dimension", NS)
        if dimension is not None:
            dimension.set("ref", f"A1:YZ{max(3000, len(rows)+4)}")
        self.files[self.sheet_path] = etree.tostring(self.sheet_root, xml_declaration=True, encoding="UTF-8", standalone=True)
        self._set_recalculation()

    def _set_recalculation(self) -> None:
        root = etree.fromstring(self.files["xl/workbook.xml"])
        calc = root.find("m:calcPr", NS)
        if calc is None:
            calc = etree.SubElement(root, f"{{{NS_MAIN}}}calcPr")
        calc.set("calcMode", "auto")
        calc.set("fullCalcOnLoad", "1")
        calc.set("forceFullCalc", "1")
        self.files["xl/workbook.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def save(self, output: Path) -> None:
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, content in self.files.items():
                zf.writestr(name, content)


def chunks(values: list[dict[str, Any]], size: int) -> Iterator[list[dict[str, Any]]]:
    for start in range(0, len(values), size):
        yield values[start:start + size]


def validate_config(cfg: dict[str, Any]) -> None:
    template = Path(cfg["template_path"])
    if not template.exists():
        raise FileNotFoundError(f"Template not found: {template}")
    if not cfg.get("shipping_template"):
        logging.warning("Shipping Template is blank. Fill it in config.json before final Temu upload.")
    if not cfg.get("eu_responsible_person"):
        logging.warning("EU Responsible person is blank. Confirm whether Temu requires it for this seller/product setup.")


def load_config(config_path: Path) -> dict[str, Any]:
    user_cfg = json.loads(config_path.read_text(encoding="utf-8")) if config_path.exists() else {}
    cfg = deep_merge(DEFAULT_CONFIG, user_cfg)
    root = config_path.parent.resolve()
    cfg["template_path"] = str((root / cfg["template_path"]).resolve()) if not Path(cfg["template_path"]).is_absolute() else cfg["template_path"]
    cfg["output_dir"] = str((root / cfg["output_dir"]).resolve()) if not Path(cfg["output_dir"]).is_absolute() else cfg["output_dir"]
    manifest_path = str(cfg.get("manifest_path") or "").strip()
    if manifest_path and not Path(manifest_path).is_absolute():
        cfg["manifest_path"] = str((root / manifest_path).resolve())
    validate_config(cfg)
    return cfg


def choose_preferred_source(existing: tuple[str, str], candidate: tuple[str, str]) -> tuple[str, str]:
    """Choose the URL/category representation that best matches the product slug."""
    def score(item: tuple[str, str]) -> int:
        url, kind = item
        inferred = infer_kind_from_url(url, kind)
        path = unquote(urlparse(url).path).lower()
        value = 0
        if inferred == kind:
            value += 3
        if kind == "bodysuit" and path.startswith("/бодита/"):
            value += 2
        if kind == "tshirt" and path.startswith("/тениски/"):
            value += 2
        return value
    return candidate if score(candidate) > score(existing) else existing


def discover_product_sources(cfg: dict[str, Any]) -> tuple[list[tuple[str, str]], list[dict[str, str]], list[dict[str, str]]]:
    session = make_session()
    raw_sources: list[tuple[str, str]] = []
    for category in cfg["categories"]:
        kind = "bodysuit" if category["slug"] == "бодита" else "tshirt"
        links = discover_products(session, category, cfg)
        raw_sources.extend((url, kind) for url in links)

    excluded: list[dict[str, str]] = []
    duplicates: list[dict[str, str]] = []
    by_identity: dict[str, tuple[str, str]] = {}
    for url, fallback_kind in raw_sources:
        kind = infer_kind_from_url(url, fallback_kind)
        decoded_path = unquote(urlparse(url).path).lower()
        if cfg.get("exclude_complex_sets", True) and "комплект" in decoded_path:
            excluded.append({
                "url": url,
                "kind": kind,
                "reason": "Obvious complex apparel set excluded during discovery",
            })
            continue
        identity = product_identity_from_url(url)
        candidate = (url, kind)
        if identity in by_identity:
            preferred = choose_preferred_source(by_identity[identity], candidate)
            dropped = candidate if preferred == by_identity[identity] else by_identity[identity]
            kept = preferred
            by_identity[identity] = preferred
            duplicates.append({
                "identity": identity,
                "kept_url": kept[0],
                "removed_url": dropped[0],
                "reason": "Same product numeric ID discovered in more than one category path",
            })
        else:
            by_identity[identity] = candidate

    product_sources = list(by_identity.values())
    if cfg.get("max_products"):
        product_sources = product_sources[: int(cfg["max_products"])]
    return product_sources, excluded, duplicates


def products_by_category(product_sources: list[tuple[str, str]]) -> dict[str, int]:
    result: dict[str, int] = {}
    for _, kind in product_sources:
        result[kind] = result.get(kind, 0) + 1
    return result


def save_manifest(config_path: Path, manifest_path: Path) -> Path:
    cfg = load_config(config_path)
    product_sources, excluded, duplicates = discover_product_sources(cfg)
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "products_found": len(product_sources),
        "products_by_category": products_by_category(product_sources),
        "excluded_during_discovery": excluded,
        "duplicates_removed_during_discovery": duplicates,
        "products": [{"url": url, "kind": kind} for url, kind in product_sources],
    }
    manifest_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    logging.info(
        "Saved manifest %s with %d products (%d complex sets excluded, %d duplicate paths removed)",
        manifest_path,
        len(product_sources),
        len(excluded),
        len(duplicates),
    )
    return manifest_path


def load_manifest(manifest_path: Path) -> list[tuple[str, str]]:
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    raw_products = payload.get("products", payload if isinstance(payload, list) else [])
    by_identity: dict[str, tuple[str, str]] = {}
    for item in raw_products:
        if isinstance(item, dict):
            url = str(item.get("url") or "").strip()
            kind = str(item.get("kind") or "").strip()
        elif isinstance(item, (list, tuple)) and len(item) >= 2:
            url, kind = str(item[0]).strip(), str(item[1]).strip()
        else:
            continue
        if not url or kind not in {"tshirt", "bodysuit"}:
            continue
        kind = infer_kind_from_url(url, kind)
        identity = product_identity_from_url(url)
        candidate = (url, kind)
        if identity in by_identity:
            by_identity[identity] = choose_preferred_source(by_identity[identity], candidate)
        else:
            by_identity[identity] = candidate
    return list(by_identity.values())


def output_stem(prefix: str, base: str) -> str:
    clean = re.sub(r"[^A-Za-z0-9_-]+", "-", prefix.strip()).strip("-")
    return f"{base}_{clean}" if clean else base


def run(
    config_path: Path,
    *,
    manifest_path: Path | None = None,
    product_offset: int | None = None,
    product_limit: int | None = None,
    output_prefix: str | None = None,
) -> list[Path]:
    cfg = load_config(config_path)

    configured_manifest = str(cfg.get("manifest_path") or "").strip()
    effective_manifest = manifest_path or (Path(configured_manifest) if configured_manifest else None)
    discovery_excluded: list[dict[str, str]] = []
    discovery_duplicates: list[dict[str, str]] = []
    if effective_manifest:
        product_sources_all = load_manifest(effective_manifest)
        try:
            manifest_payload = json.loads(effective_manifest.read_text(encoding="utf-8"))
            discovery_excluded = list(manifest_payload.get("excluded_during_discovery", []))
            discovery_duplicates = list(manifest_payload.get("duplicates_removed_during_discovery", []))
        except Exception:
            pass
        logging.info("Loaded %d products from manifest %s", len(product_sources_all), effective_manifest)
    else:
        product_sources_all, discovery_excluded, discovery_duplicates = discover_product_sources(cfg)

    total_discovered = len(product_sources_all)
    offset = max(0, int(product_offset if product_offset is not None else cfg.get("product_offset", 0) or 0))
    limit = max(0, int(product_limit if product_limit is not None else cfg.get("product_limit", 0) or 0))
    end = min(total_discovered, offset + limit) if limit else total_discovered
    product_sources = product_sources_all[offset:end]
    logging.info(
        "Selected product range %d:%d from %d discovered products (%d products in this batch)",
        offset,
        end,
        total_discovered,
        len(product_sources),
    )

    session = make_session()
    all_rows: list[dict[str, Any]] = []
    failures: list[dict[str, str]] = []
    excluded_runtime: list[dict[str, str]] = []
    processed_products = 0
    for index, (url, kind) in enumerate(product_sources, start=1):
        try:
            logging.info("[%d/%d] %s", index, len(product_sources), url)
            source = fetch(session, url, delay=float(cfg["request_delay_seconds"]))
            product = parse_product(source, url, kind, cfg)
            if cfg.get("exact_variants_with_browser"):
                try:
                    static = list(product.variants)
                    exact = browser_variants(url, int(cfg["browser_timeout_ms"]), product.base_price_eur, product.product_kind)
                    if exact:
                        product.variants = merge_exact_with_static(static, exact, product.product_kind)
                except Exception as exc:
                    logging.warning("Browser variation enumeration failed for %s: %s", url, exc)
                    if not cfg.get("fallback_to_colour_size_cross_product"):
                        raise
            rows = product_rows(product, cfg)
            all_rows.extend(rows)
            processed_products += 1
            logging.info("  %s (%s): %d SKU rows", product.title, product.code, len(rows))
        except ExcludedProductError as exc:
            logging.info("Excluded complex set %s: %s", url, exc.reason)
            excluded_runtime.append({
                "url": exc.url,
                "title": exc.title,
                "code": exc.code,
                "reason": exc.reason,
            })
        except Exception as exc:
            logging.exception("Failed product %s", url)
            failures.append({"url": url, "error": str(exc)})

    output_dir = Path(cfg["output_dir"])
    output_dir.mkdir(parents=True, exist_ok=True)
    prefix = output_prefix if output_prefix is not None else str(cfg.get("output_prefix") or "")
    workbook_stem = output_stem(prefix, "TEMU_ARTGIFT_UPLOAD")
    report_stem = output_stem(prefix, "run_report")
    excluded_stem = output_stem(prefix, "excluded_products")
    audit_stem = output_stem(prefix, "validation")

    validation = audit_rows(all_rows)
    (output_dir / f"{audit_stem}.json").write_text(
        json.dumps(validation, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (output_dir / f"{excluded_stem}.json").write_text(
        json.dumps(excluded_runtime, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    report = {
        "products_discovered_total": total_discovered,
        "product_offset": offset,
        "product_limit": limit,
        "product_range_end": end,
        "products_selected_in_batch": len(product_sources),
        "products_processed": processed_products,
        "products_by_category": products_by_category(product_sources),
        "sku_rows": len(all_rows),
        "excluded_complex_sets_in_batch": excluded_runtime,
        "discovery_excluded_total": len(discovery_excluded),
        "discovery_duplicates_removed_total": len(discovery_duplicates),
        "validation": validation,
        "sku_index": [
            {
                "sku": row.get("sku", ""),
                "parent_sku": row.get("parent_sku", ""),
                "url": row.get("reference_link", ""),
            }
            for row in all_rows
        ],
        "workbooks": [],
        "failures": failures,
        "warnings": [
            "Complex multi-garment sets are excluded because the Temu template supports one Size and one Color axis, not two or more independently selected garment sizes.",
            "Packaging weight/dimensions are configurable defaults and must be verified.",
            "Shipping Template and EU Responsible person are store-specific and may need configuration.",
        ],
    }

    if cfg.get("strict_validation", True) and not validation["passed"]:
        (output_dir / f"{report_stem}.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        raise RuntimeError(
            "Strict validation failed: duplicate SKU, parent collision, size/color swap, complex-set row, or missing category name detected."
        )

    outputs: list[Path] = []
    for part, row_chunk in enumerate(chunks(all_rows, int(cfg["rows_per_workbook"])), start=1):
        writer = TemuWorkbookWriter(Path(cfg["template_path"]))
        writer.write_rows(row_chunk)
        path = output_dir / f"{workbook_stem}_part{part:02d}.xlsx"
        writer.save(path)
        outputs.append(path)
        logging.info("Saved %s with %d rows", path, len(row_chunk))

    report["workbooks"] = [str(p) for p in outputs]
    (output_dir / f"{report_stem}.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return outputs


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", type=Path, default=Path("config.json"))
    parser.add_argument("--log-level", default="INFO", choices=("DEBUG", "INFO", "WARNING", "ERROR"))
    parser.add_argument("--discover-only", action="store_true", help="Discover product URLs and save a manifest without opening a browser.")
    parser.add_argument("--manifest", type=Path, help="Manifest to write in discover mode or read in scrape mode.")
    parser.add_argument("--product-offset", type=int, default=None, help="Zero-based start position inside the manifest.")
    parser.add_argument("--product-limit", type=int, default=None, help="Number of products to process from the selected offset; 0 means all remaining.")
    parser.add_argument("--output-prefix", default=None, help="Unique suffix used in workbook and report filenames.")
    args = parser.parse_args()
    logging.basicConfig(level=getattr(logging, args.log_level), format="%(asctime)s | %(levelname)s | %(message)s")

    if args.discover_only:
        manifest = args.manifest or Path("product_manifest.json")
        try:
            save_manifest(args.config, manifest)
        except Exception:
            logging.exception("Product discovery failed")
            return 1
        return 0

    try:
        outputs = run(
            args.config,
            manifest_path=args.manifest,
            product_offset=args.product_offset,
            product_limit=args.product_limit,
            output_prefix=args.output_prefix,
        )
    except Exception:
        logging.exception("Scraper failed")
        return 1
    if not outputs:
        # A batch can legitimately contain only products excluded by the safe
        # multi-garment-set rule. Treat that as a successful empty batch so the
        # matrix workflow can continue and the finalizer can collect the
        # exclusion reports. Genuine fetch/parser failures still return 2.
        cfg = load_config(args.config)
        prefix = args.output_prefix if args.output_prefix is not None else str(cfg.get("output_prefix") or "")
        report_path = Path(cfg["output_dir"]) / f"{output_stem(prefix, 'run_report')}.json"
        try:
            report = json.loads(report_path.read_text(encoding="utf-8"))
        except Exception:
            report = {}
        failures = report.get("failures", [])
        excluded = report.get("excluded_complex_sets_in_batch", [])
        if excluded and not failures:
            logging.warning("No workbook was needed: all selected products were safely excluded as complex sets")
            return 0
        logging.error("No workbook was created because no SKU rows were collected")
        return 2
    print("\n".join(str(p) for p in outputs))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
