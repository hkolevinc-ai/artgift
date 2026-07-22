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
from urllib.parse import urljoin, urlparse

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
    "default_kids_tshirt_category": 30843,
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


def infer_category(url: str, title: str, size_title: str, product_kind: str, cfg: dict[str, Any]) -> int:
    text = normalize_space(" ".join([urlparse(url).path, title, size_title])).lower()
    if product_kind == "bodysuit":
        if re.search(r"момич|girl|принцес|дъщер", text):
            return 26325
        return int(cfg["default_bodysuit_category"])
    if re.search(r"дамск|женск|women|woman", text):
        return 29069
    if re.search(r"мъжк|men|man", text):
        return 30469
    if re.search(r"момич|girl", text):
        return 29553
    if re.search(r"момч|boy", text):
        return 30843
    return int(cfg["default_kids_tshirt_category"])


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


def static_variants(soup: BeautifulSoup, base_price: float) -> list[Variant]:
    sizes: list[tuple[str, str]] = []
    for a in soup.select("a.variations_products_links.main_filter"):
        value = normalize_space(a.get_text(" ", strip=True))
        if value and value not in [x[0] for x in sizes]:
            sizes.append((value, a.get("data-variation_id", "")))
    colors: list[tuple[str, str]] = []
    for a in soup.select("a.variations_products_links.add_filter"):
        value = normalize_space(a.get("title") or a.get_text(" ", strip=True))
        if value and value not in [x[0] for x in colors]:
            colors.append((value, a.get("data-variation_id", "")))
    if not sizes:
        sizes = [("One Size", "")]
    if not colors:
        colors = [("As shown", "")]
    return [Variant(size, color, sid or cid, base_price) for size, sid in sizes for color, cid in colors]


def browser_variants(url: str, timeout_ms: int, base_price: float) -> list[Variant]:
    """Enumerate exact offered size/colour combinations with the storefront browser.

    For each colour, the site refreshes the size anchors and embeds the matching
    variation IDs in them. Reading those anchors is much faster than clicking every
    individual size while still excluding combinations the site does not offer.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError("Playwright is not installed") from exc

    variants: dict[tuple[str, str], Variant] = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(locale="bg-BG")
        page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
        page.wait_for_selector("a.variations_products_links", timeout=timeout_ms)

        color_names = page.locator("a.variations_products_links.add_filter").evaluate_all(
            "els => els.filter(e => e.offsetParent !== null).map(e => (e.title || e.textContent || '').trim()).filter(Boolean)"
        )
        if not color_names:
            color_names = ["As shown"]

        for color_name in color_names:
            if color_name != "As shown":
                is_active = page.evaluate(
                    "name => { const e=[...document.querySelectorAll('a.variations_products_links.add_filter')].find(x => (x.title||x.textContent||'').trim()===name); return !!(e && e.classList.contains('current_sell')); }",
                    color_name,
                )
                if not is_active:
                    clicked = page.evaluate(
                        "name => { const e=[...document.querySelectorAll('a.variations_products_links.add_filter')].find(x => (x.title||x.textContent||'').trim()===name); if(e){e.click(); return true;} return false; }",
                        color_name,
                    )
                    if not clicked:
                        continue
                try:
                    page.wait_for_function(
                        "name => { const options=[...document.querySelectorAll('a.variations_products_links.add_filter')]; const target=options.find(x => (x.title||x.textContent||'').trim()===name); if(!target || !target.classList.contains('current_sell')) return false; const current=document.querySelector('#variation_id'); return !current || !target.dataset.variation_id || current.value===target.dataset.variation_id; }",
                        color_name,
                        timeout=4000,
                    )
                except Exception:
                    page.wait_for_timeout(700)
                try:
                    page.wait_for_load_state("networkidle", timeout=2500)
                except Exception:
                    pass

            # The storefront updates the variation id before the displayed price.
            # Read until the value is stable twice to avoid carrying over the
            # previous colour's price on slow AJAX responses.
            price_text = str(base_price)
            if page.locator("#price").count():
                last = None
                stable_reads = 0
                for _ in range(8):
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
                price = float(price_text.replace(",", "."))
            except Exception:
                price = base_price
            image_urls = page.locator(".main_pic_container a.main_pic").evaluate_all("els => els.map(e => e.href)")
            sizes = page.locator("a.variations_products_links.main_filter").evaluate_all(
                "els => els.filter(e => e.offsetParent !== null && !e.classList.contains('disabled')).map(e => ({size:(e.textContent||'').trim(), id:e.dataset.variation_id||''})).filter(x => x.size)"
            )
            if not sizes:
                current_id = page.locator("#variation_id").input_value() if page.locator("#variation_id").count() else ""
                sizes = [{"size": "One Size", "id": current_id}]
            for item in sizes:
                variants[(item["size"], color_name)] = Variant(item["size"], color_name, item["id"], price, image_urls)
        browser.close()
    return list(variants.values())


def parse_product(html: str, url: str, product_kind: str, cfg: dict[str, Any]) -> Product:
    soup = BeautifulSoup(html, "lxml")
    product_json = find_product_json(parse_json_ld(soup))
    title = normalize_space((soup.select_one("h1.product_title_view") or soup.title).get_text(" ", strip=True))
    code = normalize_space(soup.select_one(".custom_code").get_text(" ", strip=True) if soup.select_one(".custom_code") else str(product_json.get("sku", "")))
    site_id = (soup.select_one("#product_id") or {}).get("value", "") if soup.select_one("#product_id") else ""
    short = normalize_space(soup.select_one(".short_description").get_text(" ", strip=True) if soup.select_one(".short_description") else str(product_json.get("description", "")))
    additional = normalize_space(soup.select_one(".additional_description").get_text(" ", strip=True) if soup.select_one(".additional_description") else "")
    description = normalize_space(" ".join(x for x in (short, additional) if x))[:2000]
    images = []
    for a in soup.select(".main_pic_container a.main_pic[href]"):
        image = urljoin(BASE_URL, a.get("href", ""))
        if image and image not in images:
            images.append(image)
    if not images and product_json.get("image"):
        values = product_json["image"] if isinstance(product_json["image"], list) else [product_json["image"]]
        images.extend(urljoin(BASE_URL, x) for x in values)
    price = parse_price(product_json, soup)
    size_title = normalize_space(soup.select_one(".main_filter_title").get_text(" ", strip=True) if soup.select_one(".main_filter_title") else "")
    customizable, customization_text = extract_customization(soup)
    category_id = infer_category(url, title, size_title, product_kind, cfg)
    variants = static_variants(soup, price)
    return Product(url, site_id, code or site_id, title, description, short, category_id, product_kind,
                   size_title, images[:10], price, customizable, customization_text, variants)


def product_rows(product: Product, cfg: dict[str, Any]) -> list[dict[str, Any]]:
    parent_code = f"AG-{safe_slug(product.code or product.site_id, 60)}"
    pack = cfg["packaging"][product.product_kind]
    custom_detail = product.customization_text or "Buyer can provide the personalization requested on the product page."
    bullet_custom = f"Customizable product: {custom_detail}"[:500]
    mark_customizable = bool(cfg.get("force_customizable", True) or product.customizable)
    rows: list[dict[str, Any]] = []
    seen_skus: set[str] = set()
    for idx, variant in enumerate(product.variants, start=1):
        size_family, sub_size_family, size = map_size(variant.size_raw)
        color = map_color(variant.color_raw)
        suffix = f"{safe_slug(color,18)}-{safe_slug(size,18)}"
        sku = f"{parent_code}-{suffix}"[:90]
        if sku in seen_skus:
            sku = f"{sku[:82]}-{idx:04d}"
        seen_skus.add(sku)
        price = variant.price_eur if variant.price_eur is not None else product.base_price_eur
        images = (variant.images or product.images)[:10]
        detail_images = product.images[:10]
        rows.append({
            "category": product.category_id,
            "product_type": "Custom product" if mark_customizable else "Normal product",
            "customization_mode": "Single technique" if mark_customizable else "",
            "primary_technique": "Leather/fabric customization technique" if mark_customizable else "",
            "secondary_technique": "Digital printing" if mark_customizable else "",
            "product_name": product.title[:500],
            "parent_sku": parent_code,
            "sku": sku,
            "operation": "Add",
            "description": product.description,
            "bullet_points": [bullet_custom, "Material: 100% cotton", "Direct digital print", "Each colour and size is a separate SKU"],
            "detail_images": detail_images,
            "variation_theme": "Color × Size",
            "size_family": size_family,
            "sub_size_family": sub_size_family,
            "size": size,
            "color": color,
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
            # Use the merchant SKU as the traceable product identification.
            # Storefront variation IDs can repeat across colours on Art-Gift.
            "product_identification": sku,
            "manufacturer": cfg["manufacturer"], "eu_responsible_person": cfg["eu_responsible_person"],
        })
    return rows


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


def run(config_path: Path) -> list[Path]:
    user_cfg = json.loads(config_path.read_text(encoding="utf-8")) if config_path.exists() else {}
    cfg = deep_merge(DEFAULT_CONFIG, user_cfg)
    root = config_path.parent.resolve()
    cfg["template_path"] = str((root / cfg["template_path"]).resolve()) if not Path(cfg["template_path"]).is_absolute() else cfg["template_path"]
    cfg["output_dir"] = str((root / cfg["output_dir"]).resolve()) if not Path(cfg["output_dir"]).is_absolute() else cfg["output_dir"]
    validate_config(cfg)

    session = make_session()
    product_sources: list[tuple[str, str]] = []
    for category in cfg["categories"]:
        kind = "bodysuit" if category["slug"] == "бодита" else "tshirt"
        links = discover_products(session, category, cfg)
        product_sources.extend((url, kind) for url in links)
    product_sources = list(dict.fromkeys(product_sources))
    if cfg.get("max_products"):
        product_sources = product_sources[: int(cfg["max_products"])]
    logging.info("Total unique products: %d", len(product_sources))

    all_rows: list[dict[str, Any]] = []
    failures: list[dict[str, str]] = []
    for index, (url, kind) in enumerate(product_sources, start=1):
        try:
            logging.info("[%d/%d] %s", index, len(product_sources), url)
            source = fetch(session, url, delay=float(cfg["request_delay_seconds"]))
            product = parse_product(source, url, kind, cfg)
            if cfg.get("exact_variants_with_browser"):
                try:
                    exact = browser_variants(url, int(cfg["browser_timeout_ms"]), product.base_price_eur)
                    if exact:
                        product.variants = exact
                except Exception as exc:
                    logging.warning("Browser variation enumeration failed for %s: %s", url, exc)
                    if not cfg.get("fallback_to_colour_size_cross_product"):
                        raise
            rows = product_rows(product, cfg)
            all_rows.extend(rows)
            logging.info("  %s (%s): %d SKU rows", product.title, product.code, len(rows))
        except Exception as exc:
            logging.exception("Failed product %s", url)
            failures.append({"url": url, "error": str(exc)})

    output_dir = Path(cfg["output_dir"])
    output_dir.mkdir(parents=True, exist_ok=True)
    outputs: list[Path] = []
    for part, row_chunk in enumerate(chunks(all_rows, int(cfg["rows_per_workbook"])), start=1):
        writer = TemuWorkbookWriter(Path(cfg["template_path"]))
        writer.write_rows(row_chunk)
        path = output_dir / f"TEMU_ARTGIFT_UPLOAD_part{part:02d}.xlsx"
        writer.save(path)
        outputs.append(path)
        logging.info("Saved %s with %d rows", path, len(row_chunk))

    products_by_category: dict[str, int] = {}
    for _, kind in product_sources:
        products_by_category[kind] = products_by_category.get(kind, 0) + 1

    report = {
        "products_found": len(product_sources), "products_by_category": products_by_category,
        "sku_rows": len(all_rows),
        "workbooks": [str(p) for p in outputs], "failures": failures,
        "warnings": [
            "Packaging weight/dimensions are configurable defaults and must be verified.",
            "Unisex bodysuits default to category 26402; review products that should use 26325.",
            "Shipping Template and EU Responsible person are store-specific and may need configuration.",
        ],
    }
    (output_dir / "run_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return outputs


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", type=Path, default=Path("config.json"))
    parser.add_argument("--log-level", default="INFO", choices=("DEBUG", "INFO", "WARNING", "ERROR"))
    args = parser.parse_args()
    logging.basicConfig(level=getattr(logging, args.log_level), format="%(asctime)s | %(levelname)s | %(message)s")
    try:
        outputs = run(args.config)
    except Exception:
        logging.exception("Scraper failed")
        return 1
    if not outputs:
        logging.error("No workbook was created because no SKU rows were collected")
        return 2
    print("\n".join(str(p) for p in outputs))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
