from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
from bs4 import NavigableString
import pandas as pd
import time
from datetime import datetime
import json
import re
import os
import base64
import mimetypes
from dotenv import load_dotenv
from xlsxwriter.utility import xl_col_to_name

load_dotenv()
SMTP_USERNAME = os.getenv("SMTP_USERNAME")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
FROM_ADDR = os.getenv("FROM_ADDR")
import ssl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

urls_studentdepot = {
    "Warszawa Wilanowska(Student depot)": "https://studentdepot.pl/pl/akademik-warszawa-wilanowska",
    "Warszawa Suwak(Student depot)": "https://studentdepot.pl/pl/akademik-warszawa-suwak",
    "Łódź Wróblewskiego(Student depot)": "https://studentdepot.pl/pl/akademik-lodz-wroblewskiego",
    "Lódź Wigury(Student depot)": "https://studentdepot.pl/pl/akademik-lodz-wigury",
    "Gdańsk(Student depot)": "https://studentdepot.pl/pl/akademik-gdansk",
    "Kraków(Student depot)": "https://studentdepot.pl/pl/akademik-krakow",
    "Lublin(Student depot)": "https://studentdepot.pl/pl/akademik-lublin",
    "Poznań A(Student depot)": "https://studentdepot.pl/pl/akademik-poznan",
    "Poznań B(Student depot)": "https://studentdepot.pl/pl/akademik-poznan-2",
    "Wrocław(Student depot)": "https://studentdepot.pl/pl/akademik-wroclaw"
}

urls_Basecamp = {
    "Warszawa(Basecamp)": "https://www.basecampstudent.com/student/warsaw-wenedow/#rooms-types-details-block",
    "Lódź Rewolucji(Basecamp)": "https://www.basecampstudent.com/student/lodz-rewolucji/#rooms-types-details-block",
    "Lódź Rembelińskiego(Basecamp)": "https://www.basecampstudent.com/student/lodz-rembielinskiego/#rooms-types-details-block",
    "Katowice(Basecamp)": "https://www.basecampstudent.com/student/katowice/#rooms-types-details-block",
    "Kraków(Basecamp)": "https://www.basecampstudent.com/student/krakow/#rooms-types-details-block",
    "Wrocław(Basecamp)": "https://www.basecampstudent.com/student/wroclaw/#rooms-types-details-block",
}

urls_Nextdoor = {
    "kraków(Nextdoor)": [
        "https://nextdoor-housing.pl/pokoj/pokoj-dwuosobowy-extra-w-prywatnym-akademiku/?pa_dlugosc-wynajmu=12-miesiecy",
        "https://nextdoor-housing.pl/pokoj/studio-standard-parter-akademik-krakow/",
        "https://nextdoor-housing.pl/pokoj/studio-jednoosobowe-standard-akademik-krakow/",
        "https://nextdoor-housing.pl/pokoj/komfortowe-studio-dwuosobowe-akademik-krakow/",
        "https://nextdoor-housing.pl/pokoj/premium-studio-krakow-luksusowy-akademik/",
        "https://nextdoor-housing.pl/pokoj/studio-dla-par-prywatny-akadademik-krakow/",
    ]
}

urls_Shed = {
    "Kraków(Shed)": "https://shedcoliving.com/krakow/",
    "Warszawa(Shed)": "https://shedcoliving.com/warsaw-campusliving/",
    "Warszawa Ochota(Shed)": "https://shedcoliving.com/warsaw-skyliving/"
}

urls_Zeitraum = {
    "Kraków Koszykarska(Zeitraum)": "https://students.zeitraum.re/pl/location/koszykarska/",
    "Kraków Racławicka(Zeitraum)": "https://students.zeitraum.re/pl/location/raclawicka/",
    "Warszawa Solec(Zeitraum)": "https://students.zeitraum.re/pl/location/solec/",
}

urls_Milestone = {
    "Wrocław Ołbin(Milestone)": "https://triberaliving.com/wroclaw/wroclaw-olbin-student-accommodation/",
    "Wrocław Fabryczna(Milestone)": "https://triberaliving.com/wroclaw/wroclaw-fabryczna-student-accommodation/",
    "Gdańsk(Milestone)": "https://triberaliving.com/gdansk/gdansk-center-student-accommodation/",
    "Kraków(Milestone)": "https://triberaliving.com/krakow/krakow-center-student-accommodation/",

}

urls_Zeus = {
    "Lublin(Zeus)": [
        "https://zeusapartments.pl/pokoj/studio-1-osobowe-komfort-20-m2",
        "https://zeusapartments.pl/pokoj/studio-2-osobowe-20-m2",
        "https://zeusapartments.pl/pokoj/studio-2-osobowe-25-m2",
        "https://zeusapartments.pl/pokoj/studio-2-osobowe-standard-plus-27-m2",
        "https://zeusapartments.pl/pokoj/apartament-dwupokojowy-37-m2",
        "https://zeusapartments.pl/pokoj/apartament-dwupokojowy-42-m2",
        "https://zeusapartments.pl/pokoj/pokoj-2-osobowy-w-apartamencie-dwupokojowym-45m2-z-prywatna-lazienka-i-aneksem-kuchennym",
        "https://zeusapartments.pl/pokoj/pokoj-1-osobowy-w-apartamencie-dwupokojowym-45m2-z-prywatna-lazienka-i-aneksem-kuchennym"

    ]
}

urls_MagisRent = {
    "Poznań(MagisRent)": "https://www.magisrent.pl/for-rent?filterEagle=true&priceMin=0&priceMax=2200&orderOption=price_asc"
}

urls_Collegia = {
    "Gdańsk(Collegia)": "https://www.collegia.pl/pl/akademik-sobieskiego/cennik/"
}

urls_Collegiate = {
    "Milan North(Collegiate)": "https://www.collegiate.it/en/student-accommodation/milan/collegiate-milan-north/",
    "Milan Bovisa(Collegiate)": "https://www.collegiate.it/en/student-accommodation/milan/collegiate-milan-bovisa/"
}

urls_Joivy = {
    "Bologna(Joivy)": "https://coliving.joivy.com/en/rent-room-bologna/?areas=%5B%7B%22min_latitude%22%3A44.43810125763809%2C%22max_latitude%22%3A44.53865628414486%2C%22max_longitude%22%3A11.441977891842704%2C%22min_longitude%22%3A11.294084146665911%7D%5D&orderBy=availability_asc",
    "Milan(Joivy)": "https://coliving.joivy.com/en/rent-room-milan/?orderBy=availability_asc",
}

urls_CXplaces = {
    "Bari(CXplaces)": "https://www.cx-place.com/ecommerce/step/"
}

urls_TSH = {
    "bologna(TSH)": "https://semester.thesocialhub.co/en/ibe/results/?hotelId=BOL01&arrival=2025-09-29&departure=2026-06-15",
    "Florence Belfiore(TSH)": "https://semester.thesocialhub.co/en/ibe/results/?hotelId=FLO02&arrival=2025-09-29&departure=2026-06-15",
    "Florence Lavagnini(TSH)": "https://semester.thesocialhub.co/en/ibe/results/?hotelId=FLO01&arrival=2025-09-1&departure=2026-01-31",
    "Rome(TSH)": "https://semester.thesocialhub.co/en/ibe/results/?hotelId=ROM01&arrival=2025-09-1&departure=2026-06-30"
}

urls_Studentspace = {
    "Kraków Al.29 Listopada(Studentspace)": "https://www.studentspace.pl/akademiki-krakow/29-listopada",
    "Kraków Wita Stwosza A(Studentspace)": "https://www.studentspace.pl/akademiki-krakow/wita-stwosza-a",
    "Kraków Wita Stwosza B(Studentspace)": "https://www.studentspace.pl/akademiki-krakow/wita-stwosza-b"
}


headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}


def parse_studentdepot_page(soup):
    room_types = []
    prices = []


    items = soup.select("div.room.listing-item")
    if not items:
        items = soup.select("div.room-listing-item, div.room-listing .room")

    for item in items:
        title_el = item.select_one("h3.item-title")
        room_type = title_el.get_text(strip=True) if title_el else "brak"

        price_el = item.select_one("span.price-val")
        price_text = price_el.get_text(" ", strip=True) if price_el else "brak"
        room_types.append(room_type)
        prices.append(price_text)

    return pd.DataFrame({"Room Type": room_types, "Price": prices})

def parse_basecamp_page(soup):
    rooms_section = soup.find_all("h4", class_="rooms-details-list__title")
    room_types = []
    prices = []
    for room in rooms_section:
        spans = room.find_all("span")
        if len(spans) >= 2:
            room_type = spans[0].text.strip()
            price = spans[1].text.strip()
        else:
            room_type = "brak"
            price = "brak"
        room_types.append(room_type)
        prices.append(price)
    return pd.DataFrame({'Room Type': room_types, 'Price': prices})


def parse_Shed_page(soup):
    room_names = soup.find_all("span", class_="text-2xl sm:text-[32px] lg:text-[40px]")
    room_types = [room.get_text(strip=True) for room in room_names]

    price_spans = soup.find_all("span", class_="text-2xl pl-2 font-medium")
    cleaned_prices = []
    for span in price_spans:
        b = span.find("b")
        if b and b.get_text(strip=True):
            raw = b.get_text(strip=True)
        else:
            txt = span.get_text(" ", strip=True)
            nums = re.findall(r"\d[\d\s\.,]*", txt)
            raw = nums[-1] if nums else ""
        raw = raw.replace("\u00A0", " ")
        raw = raw.replace(" ", "")
        raw = re.sub(r"\.(?=\d{3}\b)", "", raw)
        raw = raw.replace(",", ".")
        raw = re.sub(r"[^\d]", "", raw)

        price_text = f"{raw} zł" if raw else "brak"
        cleaned_prices.append(price_text)

    max_len = max(len(room_types), len(cleaned_prices))
    room_types.extend(["brak"] * (max_len - len(room_types)))
    cleaned_prices.extend(["brak"] * (max_len - len(cleaned_prices)))
    return pd.DataFrame({"Room Type": room_types, "Price": cleaned_prices})


def parse_nextdoor_page(soup):
    room_title = soup.find("h1")
    if not room_title:
        room_title = soup.find("h2")
    room_type = room_title.get_text(strip=True) if room_title else "brak"
    form = soup.find("form", class_="variations_form")
    if not form:
        return pd.DataFrame({'Room Type': ["brak"], 'Price': ["brak"]})
    data_variations = form.get("data-product_variations")
    if not data_variations:
        return pd.DataFrame({'Room Type': [room_type], 'Price': ["brak"]})
    variations = json.loads(data_variations)
    price_text = "brak"
    for var in variations:
        attrs = var.get("attributes", {})
        if attrs.get("attribute_pa_dlugosc-wynajmu") == "12-miesiecy":
            price_html = var.get("price_html", "")
            soup_price = BeautifulSoup(price_html, "html.parser")
            bdi = soup_price.find("bdi")
            if bdi:
                raw_text = bdi.get_text(separator="", strip=True)
                price_text = raw_text
            break
    return pd.DataFrame({'Room Type': [room_type], 'Price': [price_text]})


def parse_Zeitraum_page(soup):
    room_types = []
    prices = []
    room_cards = soup.find_all("li", class_="space-list__item")
    for card in room_cards:
        title_elem = card.find("h3")
        if not title_elem:
            title_elem = card.find("p", class_="-mb-1 h5")
        room_type = title_elem.get_text(strip=True) if title_elem else "brak"
        price_elem = card.find("ul", class_="mt-2 md:flex space-list__attributes list")
        if price_elem:
            raw_price = price_elem.get_text(strip=True)
            raw_price = re.sub(r'^[Oo]d\s*', '', raw_price)
            final_price = " ".join(raw_price.split())
        else:
            final_price = "brak"
        room_types.append(room_type)
        prices.append(final_price)
    return pd.DataFrame({"Room Type": room_types, "Price": prices})


def parse_Milestone_page(soup):
    room_types = []
    prices = []

    room_blocks = soup.select("div.room-details")
    for block in room_blocks:
        name_el = block.select_one("h3.room-name")
        room_type = name_el.get_text(strip=True) if name_el else "brak"
        price_el = block.select_one("p.room-price")
        price_text = "brak"

        if price_el:
            raw = price_el.get_text(" ", strip=True)
            match = re.search(r'(\d[\d\s]*)\s*zł', raw, re.IGNORECASE)
            if match:
                value = match.group(1)
                value = value.replace(" ", "")
                price_text = f"{value} zł"

        room_types.append(room_type)
        prices.append(price_text)

    return pd.DataFrame({
        "Room Type": room_types,
        "Price": prices
    })


def parse_Zeus_page(soup):
    h1 = soup.find("h1")
    room_type = h1.get_text(" ", strip=True) if h1 else "brak"
    price_text = "brak"
    p_price = soup.find("p", class_="fs200")
    if p_price:
        i_tag = p_price.find("i")
        if i_tag:
            value = i_tag.get_text(strip=True)
            value = re.sub(r"[^\d]", "", value)
            if value:
                price_text = f"{value} zł"

    return pd.DataFrame({"Room Type": [room_type], "Price": [price_text]})


def get_total_pages(soup):
    page_buttons = soup.find_all("button", attrs={"wire:click": True})
    page_numbers = []
    for btn in page_buttons:
        text = btn.get_text(strip=True)
        if text.isdigit():
            page_numbers.append(int(text))
    return max(page_numbers) if page_numbers else 1


def parse_MagisRent_page(soup):
    results = []
    title_divs = soup.find_all("div", class_="title")
    for title_div in title_divs:
        h3_tag = title_div.find("h3")
        if not h3_tag:
            continue
        room_type = h3_tag.get_text(strip=True)
        price_elem = title_div.find_next("span", class_="price")
        if not price_elem:
            price_text = "brak"
        else:
            raw_text = price_elem.get_text(" ", strip=True)
            raw_text = raw_text.replace('\xa0', ' ')
            match_euro = re.search(r'€\s*([\d,\.]+)', raw_text)
            if match_euro:
                euro_value = match_euro.group(1).strip()
                euro_formatted = f"{euro_value} EUR"
            else:
                euro_formatted = "brak"
            match_pln = re.search(r'PLN\s*([\d,\.]+)', raw_text)
            if match_pln:
                pln_value = match_pln.group(1).strip()
                pln_formatted = f"{pln_value} zł"
            else:
                pln_formatted = "brak"
            if euro_formatted == "brak" and pln_formatted == "brak":
                price_text = "brak"
            else:
                price_text = f"{euro_formatted} / {pln_formatted}"
        results.append({"Room Type": room_type, "Price": price_text})
    return pd.DataFrame(results)


def parse_Collegia_page(soup):
    rows = []
    containers = soup.select("div.py-4.py-md-5.my-md-5.text-center")
    for cont in containers:
        h3 = cont.find("h3")
        room_name = h3.get_text(strip=True) if h3 else "brak"
        price = "brak"
        for p in cont.find_all("p"):
            if "cena" in p.get_text(" ", strip=True).lower():
                strong = p.find("strong")
                if strong:
                    raw = "".join(strong.stripped_strings)
                    raw = re.sub(r'(?i)\bcena[:\s]*', '', raw).strip()
                    price = raw or "brak"
                break
        rows.append({"Room Type": room_name, "Price": price})
    return pd.DataFrame(rows)


def parse_TSH_page(soup):
    room_panels = soup.find_all("div", class_="RoomPanel cf")
    room_types = []
    prices = []
    for panel in room_panels:
        title_div = panel.find("div", class_="RoomPanel__DetailsTitle")
        if title_div:
            h2_elem = title_div.find("h2")
            room_type = h2_elem.get_text(strip=True) if h2_elem else "brak"
        else:
            room_type = "brak"
        price_div = panel.find("div", class_="RoomPanel__DetailsPrice")
        if price_div:
            price_number_p = price_div.find("p", class_="RoomPanel__PriceNumber")
            price_text = price_number_p.get_text(strip=True) if price_number_p else "brak"
        else:
            price_text = "brak"
        room_types.append(room_type)
        prices.append(price_text)
    return pd.DataFrame({"Room Type": room_types, "Price": prices})


def parse_studentspace_page(soup):
    rows = []
    title_divs = soup.select("div.room_title-wrapper")
    for title_div in title_divs:
        h3 = title_div.find("h3")
        room_type = h3.get_text(strip=True) if h3 else ""
        action_wrap = None
        anc = title_div
        for _ in range(6):
            if not anc:
                break
            action_wrap = anc.find("div", class_="rooms_action-wrapper")
            if action_wrap:
                break
            anc = anc.parent

        price = ""
        if action_wrap:
            price_elem = action_wrap.find("div", class_="rooms_price-element")
            # jeśli wrapper jest pusty -> zostaw puste pole
            if price_elem and price_elem.get_text(strip=True):
                texts = price_elem.select("div.rooms_price-text:not(.is-old)")
                chosen = None
                for t in texts:
                    s = t.get_text(" ", strip=True)
                    if re.search(r"\d", s):
                        chosen = s
                        break
                if chosen:
                    num = re.sub(r"[^\d]", "", chosen)
                    price = f"{num} zł" if num else ""

        rows.append({"Room Type": room_type, "Price": price})

    return pd.DataFrame(rows)


def normalize_text(s: str) -> str:
    if not isinstance(s, str):
        return s
    s = s.replace("\u00A0", " ")
    return s.strip().lower()


def normalize_room_type(rt: str) -> str:
    if not isinstance(rt, str):
        return rt
    rt = normalize_text(rt)
    rt_cleaned = re.sub(r"\s+", "", rt)
    return rt_cleaned


def clean_price(price_str: str) -> str:
    if not price_str or price_str.lower() == "brak":
        return "brak"
    match = re.search(r'([\d\s,\.]+)\s*(zł|eur|€)', price_str, re.IGNORECASE)
    if match:
        number_str, currency = match.groups()
        number_str = number_str.replace(" ", "").replace("\u00A0", "").replace(".", "")
        number_str = number_str.replace(',', '.')
        try:
            number = float(number_str)
        except ValueError:
            return price_str
        number_int = int(round(number))
        if currency.lower() in ["€", "eur"]:
            currency_str = "EUR"
        else:
            currency_str = "zł"
        return f"{number_int} {currency_str}"
    else:
        currency_str = "EUR" if ("€" in price_str or "eur" in price_str.lower()) else "zł"
        cleaned_number = re.sub(r'[^\d,\.]', '', price_str)
        cleaned_number = cleaned_number.replace('.', '').replace(',', '.')
        if not cleaned_number:
            return "brak"
        try:
            number = float(cleaned_number)
        except ValueError:
            return price_str
        number_int = int(round(number))
        return f"{number_int} {currency_str}"


def extract_academy_info(academy_key):
    owner_match = re.search(r'\((.*?)\)', academy_key)
    owner = owner_match.group(1) if owner_match else ""
    name_without_owner = academy_key.split('(')[0].strip()
    parts = name_without_owner.split()
    if len(parts) >= 2:
        location = parts[0]
        object_name = " ".join(parts[1:])
    elif len(parts) == 1:
        location = parts[0]
        object_name = ""
    else:
        location = ""
        object_name = ""
    owner = normalize_text(owner)
    location = normalize_text(location)
    object_name = normalize_text(object_name)
    if not object_name:
        object_name = location
    return location.title(), object_name.title(), owner.title()

def send_email_outlook(from_addr, to_addrs, subject, body_text, attachment_path, smtp_username, smtp_password,
    smtp_server="smtp-mail.outlook.com", smtp_port=587):
    msg = MIMEMultipart()
    msg['From'] = from_addr
    if isinstance(to_addrs, list):
        msg['To'] = ", ".join(to_addrs)
    else:
        msg['To'] = to_addrs
    msg['Subject'] = subject
    msg.attach(MIMEText(body_text, 'plain'))
    with open(attachment_path, 'rb') as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
    encoders.encode_base64(part)
    filename = os.path.basename(attachment_path)
    part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
    msg.attach(part)
    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls(context=context)
        server.login(smtp_username, smtp_password)
        server.send_message(msg)
    print("Wiadomość wysłana!")


def load_page_soup(page, url):
    try:
        print(f"Pobieranie: {url}")
        page.goto(url, timeout=60000)
        page.wait_for_load_state("domcontentloaded")
        content = page.content()
        return BeautifulSoup(content, "html.parser")
    except Exception as e:
        print(f"❌ Błąd pobierania: {url}, błąd: {e}")
        return None

def refresh_data():
    current_snapshot = []
    def add_data(academy_key, df: pd.DataFrame):
        loc, obj, owner = extract_academy_info(academy_key)
        for _, row in df.iterrows():
            price_clean = clean_price(row["Price"])
            raw_room_type = row["Room Type"]

            room_key = normalize_room_type(raw_room_type)

            current_snapshot.append({
                'Lokalizacja': loc,
                'Właściciel': owner,
                'Obiekt': obj,
                'Typ pokoju': raw_room_type,
                'Typ pokoju (key)': room_key,
                'Price': price_clean
            })

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(user_agent=headers["User-Agent"])
        page = context.new_page()
        for academy_key, url in urls_studentdepot.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Student Depot)...")
            soup = load_page_soup(page, url)
            if soup is None:
                continue
            df = parse_studentdepot_page(soup)
            add_data(academy_key, df)


        for academy_key, url in urls_Basecamp.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Basecamp)...")
            soup = load_page_soup(page, url)
            if soup is None:
                continue
            df = parse_basecamp_page(soup)
            add_data(academy_key, df)


        for academy_key, url in urls_Shed.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Shed)...")
            soup = load_page_soup(page, url)
            if soup is None:
                continue
            df = parse_Shed_page(soup)
            add_data(academy_key, df)


        for academy_key, urls_list in urls_Nextdoor.items():
            if not isinstance(urls_list, list):
                urls_list = [urls_list]
            all_rooms = []
            for url in urls_list:
                print(f"🔄 Pobieranie danych dla: {academy_key} z {url} (Nextdoor)...")
                soup = load_page_soup(page, url)
                if soup is None:
                    continue
                df = parse_nextdoor_page(soup)
                all_rooms.append(df)
            if all_rooms:
                df_concat = pd.concat(all_rooms, ignore_index=True)
                add_data(academy_key, df_concat)


        for academy_key, url in urls_Zeitraum.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Zeitraum)...")
            soup = load_page_soup(page, url)
            if soup is None:
                continue
            df = parse_Zeitraum_page(soup)
            add_data(academy_key, df)


        for academy_key, url in urls_Milestone.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Milestone)...")
            soup = load_page_soup(page, url)
            if soup is None:
                continue
            df = parse_Milestone_page(soup)
            add_data(academy_key, df)


        for academy_key, url in urls_Zeus.items():
            if isinstance(url, list):
                for single_url in url:
                    print(f"🔄 Pobieranie danych dla: {academy_key} z {single_url} (Zeus)...")
                    soup = load_page_soup(page, single_url)
                    if soup is None:
                        continue
                    df = parse_Zeus_page(soup)
                    add_data(academy_key, df)
            else:
                print(f"🔄 Pobieranie danych dla: {academy_key} (Zeus)...")
                soup = load_page_soup(page, url)
                if soup is None:
                    continue
                df = parse_Zeus_page(soup)
                add_data(academy_key, df)


        for academy_key, url in urls_MagisRent.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (MagisRent)...")
            soup = load_page_soup(page, url)
            if soup is None:
                continue
            total_pages = get_total_pages(soup)
            print(f"Znaleziono {total_pages} stron dla {academy_key}.")
            for page_num in range(1, total_pages + 1):
                page_url = f"{url}&page={page_num}"
                print(f"Pobieranie: {page_url}")
                soup_page = load_page_soup(page, page_url)
                if soup_page is None:
                    continue
                df_page = parse_MagisRent_page(soup_page)
                add_data(academy_key, df_page)


        for key, url in urls_Collegia.items():
            print(f"🔄 Pobieranie danych dla: {key}…")
            page.goto(url, timeout=60000)
            page.wait_for_load_state("domcontentloaded")
            soup = BeautifulSoup(page.content(), "html.parser")
            df_nv2 = parse_Collegia_page(soup)
            if df_nv2 is None:
                print(f"  ❗ Parser dla {key} zwrócił None — pomijam.")
                continue
            print(f"  → znaleziono {len(df_nv2)} pokoi")
            add_data(key, df_nv2)

        for academy_key, url in urls_TSH.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (TSH)...")
            soup = load_page_soup(page, url)
            if soup is None:
                continue
            df = parse_TSH_page(soup)
            add_data(academy_key, df)


        for academy_key, url in urls_Studentspace.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Studentspace)…")
            soup = load_page_soup(page, url)
            if soup is None:
                continue
            df = parse_studentspace_page(soup)
            add_data(academy_key, df)

        browser.close()

    if not current_snapshot:
        print("Brak danych do zapisania.")
        return

    key_cols_merge = ['Właściciel', 'Lokalizacja', 'Obiekt', 'Typ pokoju (key)']

    current_df = pd.DataFrame(current_snapshot)
    current_df = current_df.drop_duplicates(subset=key_cols_merge, keep="last")
    room_name_map = (
        current_df
        .dropna(subset=["Typ pokoju"])
        .drop_duplicates(subset=key_cols_merge, keep="last")[key_cols_merge + ["Typ pokoju"]]
        .set_index(key_cols_merge)["Typ pokoju"]
    )

    current_pivot = current_df.pivot_table(index=key_cols_merge, values='Price', aggfunc='first')

    current_date = datetime.now().strftime("%Y-%m-%d_%H-%M")
    new_col_name = current_date
    current_pivot.rename(columns={'Price': new_col_name}, inplace=True)
    current_pivot["Typ pokoju"] = room_name_map
    excel_file = "StudentDepot_All_Academies_Room_Prices7.xlsx"
    key_cols = key_cols_merge  # dla funkcji merge

    def _canon(s):
        if not isinstance(s, str):
            return ""
        s = s.replace("\u00A0", " ")
        s = re.sub(r"\s+", " ", s.strip())
        return s.lower()

    def _with_key(df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        if "Typ pokoju (key)" not in df.columns:
            if "Typ pokoju" in df.columns:
                df["Typ pokoju (key)"] = df["Typ pokoju"].map(normalize_room_type)
            else:
                df["Typ pokoju (key)"] = ""

        for c in key_cols:
            df[f"__{c}_canon__"] = df[c].map(_canon)
        df["__key__"] = df[[f"__{c}_canon__" for c in key_cols]].agg("|".join, axis=1)
        return df

    def _prepare_index(df: pd.DataFrame, label="") -> pd.DataFrame:
        if df.empty:
            return df
        if not set(key_cols).issubset(df.columns):
            df = df.reset_index()
        df = _with_key(df)

        if df["__key__"].duplicated().any():
            dup_cnt = int(df["__key__"].duplicated().sum())
            print(f"⚠️ {label} ma {dup_cnt} zduplikowanych kluczy – usuwam duplikaty (keep='last').")
            df = df.drop_duplicates(subset="__key__", keep="last")

        return df.set_index("__key__")

    current_reset = _prepare_index(current_pivot.reset_index(), label="current")
    if os.path.exists(excel_file):
        try:
            master_df = pd.read_excel(excel_file)
        except Exception as e:
            print(f"Błąd odczytu pliku: {e}\nTworzę nowy DataFrame.")
            master_df = pd.DataFrame()
    else:
        master_df = pd.DataFrame()

    if master_df.empty:
        merged = current_reset[[*key_cols, "Typ pokoju", new_col_name]].copy()
    else:
        master_idx = _prepare_index(master_df, label="master")
        merged = master_idx.reindex(master_idx.index.union(current_reset.index))
        merged[new_col_name] = current_reset[new_col_name]

        for c in key_cols:
            if c not in merged.columns:
                merged[c] = None
            merged[c] = merged[c].fillna(current_reset[c])


        if "Typ pokoju" not in merged.columns:
            merged["Typ pokoju"] = None
        merged["Typ pokoju"] = merged["Typ pokoju"].fillna(current_reset["Typ pokoju"])
    merged = merged[[col for col in merged.columns if not col.startswith("__")]].copy()

    cols_seen, cols_to_drop = {}, []
    for col in merged.columns:
        base = re.sub(r"\.\d+$", "", col)
        if base in cols_seen:
            cols_to_drop.append(cols_seen[base])  # wcześniejszą wersję odrzucamy
        cols_seen[base] = col
    if cols_to_drop:
        merged.drop(columns=cols_to_drop, inplace=True, errors="ignore")

    display_cols = ['Właściciel', 'Lokalizacja', 'Obiekt', 'Typ pokoju']
    for c in display_cols:
        if c not in merged.columns:
            merged[c] = ""

    date_cols = [c for c in merged.columns if c not in display_cols and c != 'Typ pokoju (key)']
    date_cols_sorted = sorted(date_cols)
    merged = merged[display_cols + date_cols_sorted]
    merged_df = merged.reset_index(drop=True)

    if "index" in merged_df.columns:
        merged_df = merged_df.drop(columns=["index"])
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        merged_df.to_excel(writer, index=False, sheet_name='Dane')
        workbook = writer.book
        worksheet = writer.sheets['Dane']

        header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
        for col_num, value in enumerate(merged_df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        for i, col in enumerate(merged_df.columns):
            column_len = merged_df[col].fillna("").astype(str).str.len().max()
            column_len = max(column_len, len(col)) + 2
            worksheet.set_column(i, i, column_len)
        key_cols_display = ['Właściciel', 'Lokalizacja', 'Obiekt', 'Typ pokoju']
        snapshot_cols = [c for c in merged_df.columns if c not in key_cols_display]

        red_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})

        start_row = 1
        end_row = len(merged_df)
        first_excel_row = start_row + 1

        if len(snapshot_cols) >= 2:
            for i in range(1, len(snapshot_cols)):
                prev_idx = merged_df.columns.get_loc(snapshot_cols[i - 1])
                new_idx = merged_df.columns.get_loc(snapshot_cols[i])

                prev_letter = xl_col_to_name(prev_idx)
                new_letter = xl_col_to_name(new_idx)

                greater_formula = (
                    f'=IFERROR('
                    f'VALUE(LEFT(${new_letter}{first_excel_row},FIND(" ",${new_letter}{first_excel_row}&" ")-1))>'
                    f'VALUE(LEFT(${prev_letter}{first_excel_row},FIND(" ",${prev_letter}{first_excel_row}&" ")-1))'
                    f',FALSE)'
                )
                less_formula = (
                    f'=IFERROR('
                    f'VALUE(LEFT(${new_letter}{first_excel_row},FIND(" ",${new_letter}{first_excel_row}&" ")-1))<'
                    f'VALUE(LEFT(${prev_letter}{first_excel_row},FIND(" ",${prev_letter}{first_excel_row}&" ")-1))'
                    f',FALSE)'
                )

                worksheet.conditional_format(start_row, new_idx, end_row, new_idx,
                                             {'type': 'formula', 'criteria': greater_formula, 'format': red_fmt})
                worksheet.conditional_format(start_row, new_idx, end_row, new_idx,
                                             {'type': 'formula', 'criteria': less_formula, 'format': green_fmt})

    print(f"🎉 Dane zapisane do {excel_file} w kolumnie: {current_date}")


    send_email_outlook(
        from_addr=FROM_ADDR,
        to_addrs=["jakub.swiniarski@studentdepot.pl","adam.swiniarski@studentdepot.pl", "sylwia.rogalska@studentdepot.pl"],
        subject="Tygodniowa aktualizacja cen akademików",
        body_text="Cześć,\nW załączniku przesyłam najnowszy plik Excel z cenami akademików.",
        attachment_path=excel_file,
        smtp_username=SMTP_USERNAME,
        smtp_password=SMTP_PASSWORD
    )


if __name__ == "__main__":
    refresh_data()













