# ============================================================
# TRADEINDIA BUSINESS WEBSITE SCRAPER
# Generic Scraper for IndiaMART Powered Business Websites
# ============================================================

import os
import re
import html
import json
import requests

from bs4 import BeautifulSoup

from urllib.parse import urljoin, urlparse

from requests.adapters import HTTPAdapter

from urllib3.util.retry import Retry

from openpyxl import Workbook

from urllib.parse import urlparse
# ============================================================
# CONFIGURATION
# ============================================================

BASE_URL = "https://www.mkpetro.com/"

TIMEOUT = 30

HEADERS = {

    "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/138.0 Safari/537.36"

}


# ============================================================
# SESSION
# ============================================================

session = requests.Session()

retry = Retry(

    total=3,

    backoff_factor=1,

    status_forcelist=[429,500,502,503,504]

)

adapter = HTTPAdapter(max_retries=retry)

session.mount("http://", adapter)

session.mount("https://", adapter)

session.headers.update(HEADERS)


# ============================================================
# HELPERS
# ============================================================

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.google.com/",
}

def get_html(url):
    print(f"\nOpening : {url}")

    r = session.get(
        url,
        headers=HEADERS,
        timeout=30,
        verify=False
    )

    r.raise_for_status()
    
    return r.text

    return ""

def get_soup(url):

    html_text = get_html(url)

    return BeautifulSoup(

        html_text,

        "html.parser"

    )
def extract_eval_array(script_text):

    pos = script_text.find("eval([")

    if pos == -1:
        return None

    start = script_text.find("[", pos)

    depth = 0

    for i in range(start, len(script_text)):

        ch = script_text[i]

        if ch == "[":
            depth += 1

        elif ch == "]":
            depth -= 1

            if depth == 0:
                return script_text[start:i+1]

    return None
def extract_products(url):

    html_text = get_html(url)

    soup = BeautifulSoup(html_text, "html.parser")

    category = extract_category(soup)

    html_tables = extract_html_tables(soup)

    products = []

    script_text = ""

    for script in soup.find_all("script"):

        txt = script.string or script.get_text()

        if "dataref" in txt.lower():

            print("\n" + "=" * 80)
            print("FOUND SCRIPT")
            print("=" * 80)
            print(txt[:5000])      # print first 5000 characters
            print("=" * 80)

            script_text = txt
            break

    if not script_text:
        print("No dataref found.")
        return products

    array_text = extract_eval_array(script_text)
    if not array_text:
        print("Could not locate eval array.")
        return products

    # Remove JavaScript trailing commas
    array_text = re.sub(
        r',(\s*[}\]])',
        r'\1',
        array_text
    )

    try:
        data = json.loads(array_text)

    except Exception as e:
        print("JSON Error:", e)
        return products

    description_blocks = []

    for div in soup.find_all(id=re.compile(r"^prdcont\d+$")):

        desc = clean_text(div.get_text("\n", strip=True))

        # Remove heading
        desc = re.sub(r"^Product Details\s*", "", desc, flags=re.I)

        # Remove Request Callback
        desc = re.sub(r"Request Callback.*$", "", desc, flags=re.I)

        # Split into lines
        lines = [x.strip() for x in desc.split("\n") if x.strip()]

        cleaned = []

        skip = True

        for line in lines:

            # The real description is usually a long sentence.
            if skip and len(line.split()) > 8:
                skip = False

            if not skip:
                cleaned.append(line)

        description_blocks.append(" ".join(cleaned))

    for i, item in enumerate(data, start=1):

        if int(item.get("img_id", 0)) >= 1000:
            continue

        product = {}

        product["url"] = url
        product["name"] = clean_text(item.get("prd_name", ""))
        product["image"] = item.get("img_path", "")

        if i <= len(description_blocks):
            product["description"] = description_blocks[i-1]
        else:
            product["description"] = ""

        attrs = []

        # JSON attributes
        for spec in item.get("isq_det_form", []):

            key = clean_text(spec.get("FK_IM_SPEC_MASTER_DESC", ""))
            value = clean_text(spec.get("SUPPLIER_RESPONSE_DETAIL", ""))

            if key and value:
                attrs.append((key, value))

        # HTML attributes
        if i <= len(html_tables):

            for pair in html_tables[i-1]:

                if pair not in attrs:
                    attrs.append(pair)

        product["attributes"] = attrs
        desc = product["description"]
        
        # Remove Product Details heading
        desc = re.sub(r"^Product Details\s*", "", desc, flags=re.I)

        # Remove Request Callback
        desc = re.sub(r"Request Callback.*$", "", desc, flags=re.I)

        # Remove every attribute name and value
        for key, value in attrs:

            pattern = re.escape(key) + r"\s*" + re.escape(value)

            desc = re.sub(pattern, "", desc, flags=re.I)

        # Remove extra spaces
        desc = re.sub(r"\s+", " ", desc).strip()
        # Remove leading colon(s), dash(es), spaces
        desc = re.sub(r'^\s*[:\-–]+\s*', '', desc)
        # Remove extra whitespace
        desc = re.sub(r'\s+', ' ', desc).strip()
        product["description"] = desc
        product["category"] = category

        products.append(product)

    return products


def clean_text(text):

    if not text:

        return ""

    text = html.unescape(text)

    text = re.sub(

        r"\s+",

        " ",

        text

    )

    return text.strip()

def extract_category(soup):

    h1 = soup.find("h1")

    if h1:
        return clean_text(h1.get_text())

    title = soup.find("title")

    if title:
        return clean_text(title.get_text())

    return ""
# ============================================================
# PRODUCT IMAGE
# ============================================================

def extract_product_image(soup, page_url):

    for img in soup.find_all("img"):

        src = (
            img.get("data-original")
            or img.get("data-large")
            or img.get("data-src")
            or img.get("data-lazy")
            or img.get("data-zoom-image")
            or img.get("src")
        )

        if not src:
            continue

        if src.startswith("data:image"):
            continue

        src = urljoin(page_url, src)

        lower = src.lower()

        if any(x in lower for x in (
            "logo",
            "icon",
            "facebook",
            "twitter",
            "linkedin",
            "youtube",
            "banner",
            "loader",
            "loading"
        )):
            continue

        return src

    return ""
# ============================================================
# PRODUCT DESCRIPTION
# ============================================================

def extract_description(soup):

    container = soup.find(
        id=lambda x: x and x.startswith("prdcont")
    )

    if not container:
        return ""

    text = clean_text(
        container.get_text(" ", strip=True)
    )

    ignore = [

        "Thank you",

        "Your Enquiry has been sent successfully."

    ]

    for word in ignore:
        text = text.replace(word, "")

    return text.strip()
# ============================================================
# PRODUCT ATTRIBUTES
# ============================================================

def extract_attributes(soup):

    attributes = []

    for tr in soup.find_all("tr"):

        cells = tr.find_all(["td", "th"])

        if len(cells) != 2:
            continue

        key = clean_text(
            cells[0].get_text()
        )

        value = clean_text(
            cells[1].get_text()
        )

        if not key or not value:
            continue

        attributes.append(
            (key, value)
        )

    return attributes

def extract_html_tables(soup):

    tables = []

    for table in soup.find_all("table"):

        attrs = []

        for tr in table.find_all("tr"):

            cells = tr.find_all(["td", "th"])

            if len(cells) != 2:
                continue

            key = clean_text(cells[0].get_text())

            value = clean_text(cells[1].get_text())

            if key and value:
                attrs.append((key, value))

        if attrs:
            tables.append(attrs)

    return tables

# ============================================================
# DISCOVER ALL INTERNAL LINKS
# ============================================================

IGNORE_PATTERNS = [

    "about",

    "profile",

    "contact",

    "enquiry",

    "testimonial",

    "client",

    "quality",

    "infrastructure",

    "brochure",

    "privacy",

    "terms",

    "career",

    "feedback",

    "video",

    "sitemap"

]


# ============================================================
# DISCOVER PRODUCT CATEGORY PAGES
# ============================================================

def discover_links():

    soup = get_soup(BASE_URL)

    domain = urlparse(BASE_URL).netloc

    pages = []
    seen = set()

    for a in soup.find_all("a", href=True):

        href = a["href"].strip()

        if not href:
            continue

        full = urljoin(BASE_URL, href).split("#")[0]

        parsed = urlparse(full)

        # Same website only
        if parsed.netloc != domain:
            continue

        # Only HTML pages
        if not parsed.path.endswith(".html"):
            continue

        text = clean_text(a.get_text())

        lower = text.lower()

        # Ignore common non-product pages
        if any(x in lower for x in IGNORE_PATTERNS):
            continue

        # Ignore these pages by URL
        ignore_urls = [

            "profile.html",
            "testimonial.html",
            "quality.html",
            "clientele.html",
            "corporate-video.html",
            "corporate-brochure.html",
            "infrastructure.html",
            "enquiry.html",
            "sitemap.html"

        ]

        if any(x in full.lower() for x in ignore_urls):
            continue

        if full in seen:
            continue

        seen.add(full)

        pages.append({

            "title": text,
            "url": full

        })

    return pages

def save_to_excel(products, filename=None):

    if filename is None:
        domain = urlparse(BASE_URL).netloc.replace("www.", "")
        website = domain.split(".")[0]
        filename = f"{website}_products.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Products"

    ws.append([
        "Category",
        "Page URL",
        "Product",
        "Image",
        "Description",
        "Attribute Name",
        "Attribute Value"
    ])

    for product in products:

        # If product has attributes
        if product["attributes"]:

            for key, value in product["attributes"]:

                ws.append([
                    product.get("category", ""),
                    product["url"],
                    product["name"],
                    product["image"],
                    product["description"],
                    key,
                    value
                ])

        # Product without attributes
        else:

            ws.append([
                product.get("category", ""),
                product["url"],
                product["name"],
                product["image"],
                product["description"],
                "",
                ""
            ])

    wb.save(filename)

    print(f"\nExcel saved as {filename}")

# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    pages = discover_links()

    print(f"\nDiscovered {len(pages)} pages.\n")

    all_products = []

    for page in pages:

        print("=" * 80)
        print("Processing :", page["url"])

        try:
            products = extract_products(page["url"])
        except Exception as e:
            print(f"Failed: {page['url']}")
            print(e)
            continue

        print("Products Found :", len(products))

        all_products.extend(products)

    save_to_excel(all_products)