import re
import time
from urllib.parse import urlparse

import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright


# =====================================================
# Helpers
# =====================================================

def seller_name(url):
    parsed = urlparse(url)

    if parsed.path.strip("/"):
        return parsed.path.strip("/")

    return parsed.netloc.replace("www.", "")


def get_all_links(page):
    return page.eval_on_selector_all(
        "a",
        """
        els => els.map(e => e.href).filter(Boolean)
        """
    )


def extract_product_data(page, url):

    print(f"\nExtracting Product: {url}")

    page.goto(url, wait_until="networkidle", timeout=60000)

    soup = BeautifulSoup(page.content(), "html.parser")

    # =====================================
    # CATEGORY
    # =====================================
    category = ""

    category_xpaths = [
        "/html/body/div[1]/main/div[1]/section/div[2]/nav",
        "/html/body/div[1]/main/div[1]/section/section[2]/nav"
    ]

    for xp in category_xpaths:
        try:
            locator = page.locator(f"xpath={xp}")

            if locator.count() > 0:
                txt = locator.first.inner_text(timeout=3000).strip()

                if txt:
                    category = txt
                    break

        except Exception as e:
            print(f"Category XPath Error: {e}")

    # Breadcrumb fallback
    if not category:
        try:
            crumbs = page.locator("nav a").all_inner_texts()

            crumbs = [x.strip() for x in crumbs if x.strip()]

            if crumbs:
                category = crumbs[-1]

        except Exception as e:
            print(f"Breadcrumb Error: {e}")

    if category:
        parts = [
            x.strip()
            for x in re.split(r">|/|\|", category)
            if x.strip()
        ]

        if parts:
            category = parts[-1]

    # =====================================
    # PRODUCT NAME
    # =====================================
    name = ""

    h1 = soup.find("h1")

    if h1:
        name = h1.get_text(" ", strip=True)

    if not name and soup.title:
        name = soup.title.get_text(" ", strip=True)

    # =====================================
    # DESCRIPTION
    # =====================================
    description = ""

    description_xpaths = [
        "/html/body/div[1]/main/div[1]/section/div[2]/div/div/div/div/div/div/div/div/p",
        "/html/body/div[1]/main/div[1]/section/section[2]/div/div/div/div/div/div/div/p",
        "/html/body/div[1]/main/div[1]/section/div[2]/div[1]/div/div[2]/div/div/div/div/div",
        "/html/body/div[1]/main/div[1]/section/div[2]/div/div/div/div/div/div/div/div",
        "/html/body/div[1]/main/div[1]/section/section[2]/div/div/div/div/div/div/div"
    ]

    for xp in description_xpaths:

        try:
            locator = page.locator(f"xpath={xp}")

            if locator.count() > 0:

                txt = locator.first.inner_text(timeout=5000).strip()
                txt = re.sub(r"\s+", " ", txt)

                if txt and len(txt) > 10:
                    description = txt
                    break

        except Exception:
            pass

    if not description:
        try:
            meta_desc = soup.find(
                "meta",
                attrs={"name": "description"}
            )

            if meta_desc:
                description = meta_desc.get("content", "")

        except Exception:
            pass

    # =====================================
    # IMAGE
    # =====================================
    image_url = ""

    try:
        og_image = soup.find(
            "meta",
            property="og:image"
        )

        if og_image:
            image_url = og_image.get("content", "")

    except Exception:
        pass

    # =====================================
    # ATTRIBUTES
    # =====================================
    attributes = []

    for tr in soup.select("tr"):

        cells = tr.find_all(["td", "th"])

        if len(cells) >= 2:

            attr = cells[0].get_text(" ", strip=True)
            val = cells[1].get_text(" ", strip=True)

            if attr and val:
                attributes.append((attr, val))

    print(
        f"Extracted => Name: {name}, "
        f"Category: {category}, "
        f"Attributes: {len(attributes)}"
    )

    return {
        "category": category,
        "name": name,
        "url": url,
        "image": image_url,
        "description": description,
        "attributes": attributes
    }


# =====================================================
# Read Seller URLs
# =====================================================

with open("websites.txt", "r", encoding="utf-8") as f:
    sellers = [x.strip() for x in f if x.strip()]

print(f"\nTotal Sellers: {len(sellers)}")

with sync_playwright() as p:

    browser = p.chromium.launch(headless=False)

    page = browser.new_page()

    for seller in sellers:

        print("\n" + "=" * 80)
        print("SELLER:", seller)
        print("SELLER NAME:", seller_name(seller))
        print("=" * 80)

        try:

            page.goto(
                seller,
                wait_until="networkidle",
                timeout=60000
            )

            page.wait_for_timeout(5000)

        except Exception as e:

            print("Seller Page Error:", e)
            continue

        # =====================================================
        # Seller Links
        # =====================================================

        all_links = set(get_all_links(page))

        print("\nTOTAL LINKS FOUND:", len(all_links))

        print("\nSAMPLE LINKS:")
        for x in list(all_links)[:20]:
            print(x)

        category_links = {
            x for x in all_links
            if x.endswith(".html")
            and seller_name(seller) in x
        }

        print("\nCATEGORY LINKS FOUND:", len(category_links))

        print("\nSAMPLE CATEGORY LINKS:")
        for x in list(category_links)[:20]:
            print(x)

        product_group_links = set()
        product_detail_links = set()

        # =====================================================
        # Category -> Product Group
        # =====================================================

        for idx, cat in enumerate(category_links, start=1):

            print(
                f"\n[{idx}/{len(category_links)}] "
                f"Processing Category:"
            )
            print(cat)

            try:

                page.goto(
                    cat,
                    wait_until="networkidle",
                    timeout=60000
                )

                page.wait_for_timeout(3000)

                links = set(get_all_links(page))

                print(
                    f"Links Found In Category: {len(links)}"
                )

                sample = list(links)[:20]

                print("\nSample Category Page Links:")

                for s in sample:
                    print(s)

                for link in links:

                    if "/proddetail/" in link:

                        product_detail_links.add(link)

                    elif link.endswith(".html"):

                        product_group_links.add(link)

            except Exception as e:

                print("Category Error:", e)

        print(
            "\nPRODUCT GROUP LINKS AFTER CATEGORY SCAN:",
            len(product_group_links)
        )

        print("\nSample Product Group Links:")

        for x in list(product_group_links)[:20]:
            print(x)

        # =====================================================
        # Product Group -> Product Detail
        # =====================================================

        for idx, group in enumerate(product_group_links, start=1):

            print(
                f"\n[{idx}/{len(product_group_links)}] "
                f"Processing Product Group:"
            )
            print(group)

            try:

                page.goto(
                    group,
                    wait_until="networkidle",
                    timeout=60000
                )

                page.wait_for_timeout(3000)

                links = set(get_all_links(page))

                print(
                    f"Links Found In Group: {len(links)}"
                )

                for link in links:

                    if "/proddetail/" in link:

                        product_detail_links.add(link)

            except Exception as e:

                print("Group Error:", e)

        # =====================================================
        # Final Debug
        # =====================================================

        print("\n" + "=" * 80)
        print("FINAL COUNTS")
        print("=" * 80)

        print("Category Links :", len(category_links))
        print("Group Links    :", len(product_group_links))
        print("Product Links  :", len(product_detail_links))

        print("\nSample Product Links:")

        for x in list(product_detail_links)[:20]:
            print(x)

        rows = []

        # =====================================================
        # Extract Product Data
        # =====================================================

        for i, product_url in enumerate(product_detail_links, start=1):

            try:

                print(
                    f"\nProcessing Product "
                    f"{i}/{len(product_detail_links)}"
                )

                data = extract_product_data(
                    page,
                    product_url
                )

                if not data["attributes"]:

                    rows.append([
                        data["category"],
                        data["name"],
                        data["url"],
                        data["image"],
                        data["description"],
                        "",
                        ""
                    ])

                else:

                    for attr, val in data["attributes"]:

                        rows.append([
                            data["category"],
                            data["name"],
                            data["url"],
                            data["image"],
                            data["description"],
                            attr,
                            val
                        ])

                time.sleep(1)

            except Exception as e:

                print(
                    f"Product Extraction Error "
                    f"({product_url}): {e}"
                )

        # =====================================================
        # Save Excel
        # =====================================================

        outfile = f"{seller_name(seller)}.xlsx"

        print("\nRows Generated:", len(rows))

        df = pd.DataFrame(
            rows,
            columns=[
                "Category",
                "Product Name",
                "Product URL",
                "Image URL",
                "Description",
                "Attribute",
                "Value"
            ]
        )

        df.to_excel(outfile, index=False)

        print("Excel Saved:", outfile)

    browser.close()

print("\nDONE")