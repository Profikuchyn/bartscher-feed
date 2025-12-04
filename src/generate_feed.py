import pandas as pd
import requests
import xml.etree.ElementTree as ET
from datetime import datetime
from currency import get_czk_rate


# --------------------------------------------------
# 1) Nastavení cest k souborům
# --------------------------------------------------

XLS_FILE = "produktyBartscherCZ.xlsx"
OUTPUT_XML = "output/bartscher.xml"
SKLAD_URL = "https://www.bartscher.com/download/availability-list"


# --------------------------------------------------
# 2) Funkce pro stažení skladu Bartscher
# --------------------------------------------------

def download_sklad():
    print("📦 Stahuji sklad z Bartscheru...")
    r = requests.get(SKLAD_URL)
    r.raise_for_status()

    lines = r.text.splitlines()
    sklad_map = {}

    # najdeme sloupce podle hlavičky
    header = lines[0].split("\t")
    idx_kod = header.index("Artikel Nr. / Item No.")
    idx_avail = header.index("Verfügbarkeit / Availability")

    for row in lines[1:]:
        cols = row.split("\t")
        if len(cols) <= idx_avail:
            continue
        kod = cols[idx_kod].strip()
        avail = cols[idx_avail].strip().lower()

        sklad_map[kod] = (2 if avail == "yes" else 0)

    return sklad_map


# --------------------------------------------------
# 3) Hlavní funkce pro generování XML
# --------------------------------------------------

def generate_xml():
    print("📘 Načítám XLS...")
    df = pd.read_excel(XLS_FILE)

    print("💱 Stahuji kurz EUR → CZK…")
    eur_rate = get_czk_rate()

    print(f"💶 Aktuální kurz EUR: {eur_rate} CZK")

    print("📦 Stahuji dostupnost...")
    sklad_data = download_sklad()

    root = ET.Element("products")

    # projdeme každý produkt v XLS
    for _, row in df.iterrows():
        produkt = ET.SubElement(root, "product")

        kod = str(row["kód"]).strip()
        nazev = str(row["Název"]).strip()
        ean = str(row["gtin"]).strip() if not pd.isna(row["gtin"]) else ""
        popis_text = str(row["popisText"]).strip() if not pd.isna(row["popisText"]) else ""

        # ceny
        nakup_eur = float(row["Nákupní cena do eshopu s DPH v EUR"])
        prodej_eur = float(row["Cena s DPH eshop v EUR"])

        nakup_czk = round(nakup_eur * eur_rate, 2)
        prodej_czk = round(prodej_eur * eur_rate, 2)

        # sklad
        sklad = sklad_data.get(kod, 0)
        dostupnost = "Skladem 2 ks" if sklad > 0 else "Do 5 dnů"

        # --------------------------------------------------------------------
        # Název produktu (Bartscher | ...)
        # --------------------------------------------------------------------
        full_name = f"Bartscher | {nazev}"
        ET.SubElement(produkt, "kod_produktu").text = kod
        ET.SubElement(produkt, "ean").text = ean
        ET.SubElement(produkt, "nazev_vyrobku").text = full_name
        ET.SubElement(produkt, "vyrobce").text = "Bartscher"
    # Funkce pro generování HTML popisu
    def build_html_description(name, short_desc, attributes, documentation_links):
        html = []
        html.append(f"<h2>{name}</h2>")
        html.append(f"<p><strong>{name}</strong> — {short_desc}</p>")

        # Atributy
        if attributes:
            html.append("<h3>Technické parametry</h3>")
            html.append("<ul>")
            for a in attributes:
                if a["value"]:
                    html.append(f"<li><strong>{a['name']}:</strong> {a['value']}</li>")
            html.append("</ul>")

        # Dokumentace
        if documentation_links:
            html.append("<h3>Dokumentace</h3>")
            html.append("<ul>")
            for d in documentation_links:
                if d["url"]:
                    html.append(f'<li><a href="{d["url"]}" target="_blank">{d["label"]}</a></li>')
            html.append("</ul>")

        html.append("<p>Produkt lze zakoupit u Profikuchyn.cz</p>")
        return "\n".join(html)

    # Funkce pro stavbu XML uzlu produktu
    def build_xml_product(row, eur_rate):
        product = ET.Element("product")

        def add(tag, value):
            el = ET.SubElement(product, tag)
            el.text = str(value).strip() if value not in ("", None) else ""

        # Název
        name = f'Bartscher | {row["Název"]}'.strip()

        # Krátký popis
        short_desc = f'{name} – profesionální gastronomické zařízení značky Bartscher.'

        # Identifikace produktu
        add("kod_produktu", row["kód"])
        add("nazev_vyrobku", name)
        add("ean", row["gtin"])

        # Ceny + přepočet
        purchase_eur = safe_float(row["Celková cena včetně dopravy pro distributora bez DPH v EUR (nákupní cena bez DPH v EUR)"])
        price_eur = safe_float(row["Sleva 20 procent na eshop včetně dopravy (výsledná prodejní cena bez DPH v EUR"])

        purchase_czk = round(purchase_eur * eur_rate, 2)
        price_czk = round(price_eur * eur_rate, 2)

        add("nakupni_cena", purchase_czk)
        add("prodejni_cena", price_czk)

        # Obrázky
        images = []
        for col in ["Image1", "Image2", "Image3", "Image4", "Image5", "Image6"]:
            if col in row and str(row[col]).startswith("http"):
                images.append(row[col])

        for img in images:
            add("obrazek", img)

        # Atributy → spojení do popisu
        attributes = []
        for col in row.index:
            if col.startswith("Atribut") and str(row[col]).strip():
                attributes.append({"name": col, "value": row[col]})

        # Dokumentace
        documentation_links = []
        docs = {
            "datový list": "datový list",
            "rozložený pohled": "rozložený pohled",
            "schéma zapojení": "schéma zapojení",
            "návod k obsluze": "návod k obsluze",
            "prohlášení o shodě CE": "prohlášení o shodě CE",
        }
        for col, lbl in docs.items():
            if col in row and str(row[col]).startswith("http"):
                documentation_links.append({"label": lbl, "url": row[col]})

        # HTML popis
        html_description = build_html_description(name, short_desc, attributes, documentation_links)
        add("popis_html", html_description)

        # Textový popis (SEO)
        text_description = short_desc + "\n\n" + "\n".join(
            [f"{a['name']}: {a['value']}" for a in attributes]
        )
        add("popis", text_description)

        # dostupnost / sklad
        add("sklad", row.get("sklad", "0"))
        add("dostupnost", "do 5 dní")

        return product
def generate_xml(products, exchange_rate):
    root = ET.Element("products")

    for row in products:
        product_xml = build_product_xml(row, exchange_rate)
        root.append(product_xml)

    tree = ET.ElementTree(root)
    tree.write("bartscher.xml", encoding="utf-8", xml_declaration=True)
    print("XML soubor bartscher.xml byl úspěšně vytvořen.")


def main():
    xls_path = "produktyBartscherCZ.xlsx"

    exchange_rate = get_exchange_rate()

    products = load_bartscher_xls(xls_path)

    generate_xml(products, exchange_rate)


if __name__ == "__main__":
    main()
