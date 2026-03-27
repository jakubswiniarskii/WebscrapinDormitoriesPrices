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
    "Łódź Wigury(Student depot)": "https://studentdepot.pl/pl/akademik-lodz-wigury",
    "Gdańsk(Student depot)": "https://studentdepot.pl/pl/akademik-gdansk",
    "Kraków(Student depot)": "https://studentdepot.pl/pl/akademik-krakow",
    "Lublin(Student depot)": "https://studentdepot.pl/pl/akademik-lublin",
    "Poznań A(Student depot)": "https://studentdepot.pl/pl/akademik-poznan",
    "Poznań B(Student depot)": "https://studentdepot.pl/pl/akademik-poznan-2",
    "Wrocław(Student depot)": "https://studentdepot.pl/pl/akademik-wroclaw"
}

urls_Basecamp = {
    "Warszawa(Basecamp)": "https://www.basecampstudent.com/student/warsaw-wenedow/#rooms-types-details-block",
    "Łódź Rewolucji(Basecamp)": "https://www.basecampstudent.com/student/lodz-rewolucji/#rooms-types-details-block",
    "Łódź Rembelińskiego(Basecamp)": "https://www.basecampstudent.com/student/lodz-rembielinskiego/#rooms-types-details-block",
    "Katowice(Basecamp)": "https://www.basecampstudent.com/student/katowice/#rooms-types-details-block",
    "Kraków(Basecamp)": "https://www.basecampstudent.com/student/krakow/#rooms-types-details-block",
    "Wrocław(Basecamp)": "https://www.basecampstudent.com/student/wroclaw/#rooms-types-details-block",
}

urls_Nextdoor = {
    "Kraków(Nextdoor)": [
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
    "Warszawa Ochota(Shed)": "https://shedcoliving.com/warsaw-skyliving/",
    "Riga(Shed)": "https://shedcoliving.com/riga/",
    "Wilno(Shed)":"https://shedcoliving.com/vilnius/",
}

urls_Zeitraum = {
    "Kraków Koszykarska(Zeitraum)": "https://students.zeitraum.re/pl/location/koszykarska/",
    "Kraków Racławicka(Zeitraum)": "https://students.zeitraum.re/pl/location/raclawicka/",
    "Warszawa Solec(Zeitraum)": "https://students.zeitraum.re/pl/location/solec/",
    "Prague U Průhonu(Zeitraum)": "https://students.zeitraum.re/pl/location/u-pruhonu/",
    "Prague Seifertova(Zeitraum)": "https://students.zeitraum.re/pl/location/seifertova/",
    "Prague Na Šachtě(Zeitraum)": "https://students.zeitraum.re/pl/location/na-sachte/"
}

urls_Milestone = {
    "Wrocław Ołbin(Milestone)": "https://triberaliving.com/wroclaw/wroclaw-olbin-student-accommodation/",
    "Wrocław Fabryczna(Milestone)": "https://triberaliving.com/wroclaw/wroclaw-fabryczna-student-accommodation/",
    "Gdańsk(Milestone)": "https://triberaliving.com/gdansk/gdansk-center-student-accommodation/",
    "Kraków(Milestone)": "https://triberaliving.com/krakow/krakow-center-student-accommodation/",
    "Warszawa Mokotów(Milestone)":"https://triberaliving.com/warsaw/warsaw-mokotow-student-accommodation/",

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

urls_collegiate = {
    "Milan Bovisa(Collegiate)": "https://www.collegiate.it/en/student-accommodation/milan/collegiate-milan-bovisa/",
    "Milan North(Collegiate)": "https://www.collegiate.it/en/student-accommodation/milan/collegiate-milan-north/",
}

urls_CXplaces = {
    "Turyn Vanchiglia(CX places)": "https://www.cx-place.com/cxturin-vanchiglia-campus.html",
    "Turyn Marconi(CX places)":"https://www.cx-place.com/cxturin-marconi-campus.html",
    "Milan Bicocca(CX places)":"https://www.cx-place.com/cxmilan-bicocca-campus.html",
    "Milan NoM(CX places)":"https://www.cx-place.com/cxmilan-nom-campus.html",
}

urls_TSH = {
    "Bologna(TSH)": "https://semester.thesocialhub.co/en/ibe/results/?hotelId=BOL01&arrival=2025-09-29&departure=2026-06-15",
    "Florence Belfiore(TSH)": "https://semester.thesocialhub.co/en/ibe/results/?hotelId=FLO02&arrival=2025-09-29&departure=2026-06-15",
    "Florence Lavagnini(TSH)": "https://semester.thesocialhub.co/en/ibe/results/?hotelId=FLO01&arrival=2025-09-1&departure=2026-01-31",
    "Rome(TSH)": "https://semester.thesocialhub.co/en/ibe/results/?hotelId=ROM01&arrival=2025-09-1&departure=2026-06-30"
}

urls_Studentspace = {
    "Kraków Al.29 Listopada(Studentspace)": "https://www.studentspace.pl/akademiki-krakow/29-listopada",
    "Kraków Wita Stwosza A(Studentspace)": "https://www.studentspace.pl/akademiki-krakow/wita-stwosza-a",
    "Kraków Wita Stwosza B(Studentspace)": "https://www.studentspace.pl/akademiki-krakow/wita-stwosza-b",
    "Warszawa Mokotów(Studentspace)": "https://www.studentspace.pl/akademik-warszawa-woloska",
}

urls_FizzPrague = {
    "Prague(TheFizz)" : "https://www.the-fizz.com/en/student-accommodation/prague/",
    "Berlin friedrichshain(TheFizz)" : "https://www.the-fizz.com/en/student-accommodation/berlin-friedrichshain/",
    "Berlin kreuzberg(TheFizz)" : "https://www.the-fizz.com/en/student-accommodation/berlin-kreuzberg/",
    "Monachium(TheFizz)":"https://www.the-fizz.com/en/student-accommodation/munich/"
}
urls_chillhills = {
    "Brno Kunzova(Chillhills)" : "https://www.chillhills.cz/en/nase-komplexy/kunzova/",
    "Brno Dominikanske Square(Chillhills)" : "https://www.chillhills.cz/en/nase-komplexy/dominikanske-namesti/",
    "Brno Jeronýmova street(Chillhills)" : "https://www.chillhills.cz/en/nase-komplexy/jeronymova/",
    "Brno Behounska street(Chillhills)" : "https://www.chillhills.cz/en/nase-komplexy/behounska/",
    "Brno Cihlarska(Chillhills)" : "https://www.chillhills.cz/en/nase-komplexy/cihlarska/",
    "Brno Spolkova 7(Chillhills)" : "https://www.chillhills.cz/en/nase-komplexy/spolkova/",
    "Brno Pradlacka(Chillhills)" : "https://www.chillhills.cz/en/nase-komplexy/pradlacka/",
    "Brno Masarykova(Chillhills)" : "https://www.chillhills.cz/en/masarykova/",
    "Brno Obla(Chillhills)" : "https://www.chillhills.cz/en/nase-komplexy/obla/",
    "Brno Spolkova 4,6(Chillhills)" : "https://www.chillhills.cz/en/nase-komplexy/spolkova-2/",
    "Brno Kovarska(Chillhills)" : "https://www.chillhills.cz/en/nase-komplexy/kovarska/",
    "Brno Uzbecka street(Chillhills)" : "https://www.chillhills.cz/en/nase-komplexy/uzbecka/",
    "Brno Rumiste(Chillhills)" : "https://www.chillhills.cz/en-nase-komplexy-rumiste/",
    "Brno Novy Cejl(Chillhills)" : "https://www.chillhills.cz/en/novy-cejl/",
    "Brno Dornych(Chillhills)" : "https://www.chillhills.cz/en/dornych/",
    "Brno Bratislavska(Chillhills)" : "https://www.chillhills.cz/en/bratislavska/",
    "Brno Krenova(Chillhills)" : "https://www.chillhills.cz/en/krenova/",
    "Brno Oderska(Chillhills)" : "https://www.chillhills.cz/en/oderska/",
    "Brno Slatina(Chillhills)" : "https://www.chillhills.cz/en/slatina/",
    "Brno Zidlochovice(Chillhills)" : "https://www.chillhills.cz/en/zidlochovice/",
}

urls_scandium = {

    "Tallinn lava(Scandium living)": "https://scandiumliving.ee/en/building/laava-apartments/",
    "Tallinn Järve 40(Scandium living)": "https://scandiumliving.ee/en/building/jarve-40/",
    "Tallinn Newton(Scandium living)": "https://scandiumliving.ee/en/building/newton-studios/",
    "Tallinn Magma(Scandium living)": "https://scandiumliving.ee/en/building/magma-studios/",
    "Tallinn Marati(Scandium living)": "https://scandiumliving.ee/en/building/new-marati-kvartal/",
}

urls_neonwood = {
    "Berlin Mitte-Wedding(neonwood)": "https://neonwood.com/location/berlin-mitte-wedding-2/",
    "Berlin Frankfurter Tor(neonwood)": "https://neonwood.com/location/frankfurter-tor/",
    "Berlin Adlershof(neonwood)": "https://neonwood.com/location/adlershof/",
    "Berlin Neukölln(neonwood)": "https://neonwood.com/location/th-neukolln/",
}

urls_youston = {

    "Riga Kr.Valdemāra 38(Youston)":"https://www.youstonliving.com/latvia/co-living/valdemara",
    "Wilno Slucko(Youston)":"https://youstonliving.com/lithuania/co-living/slucko",
    "Wilno Smolensko 14(Youston)":"https://youstonliving.com/lithuania/co-living/smolensko",
    "Wilno Smolensko 10(Youston)":"https://youstonliving.com/lithuania/apartments/smolensko",
}

urls_duckrepublic = {

    "Riga Lauvas(Duck republic)": "https://duckrepublik.eu/find-your-room/",
}

urls_Duckrepublic = {
    "Riga Slokas(Duck Republic)":"https://slokas.duckrepublik.eu/",
}

urls_solosociety = {
    "Wilno(Solo Society)": "https://solosociety.lt/city-house-vilnius/room-types/",
    "Wilno Kaunas(Solo Society)": "https://solosociety.lt/student-house-kaunas/rooms/",

}

urls_livin = {
    "Wilno newtown(LivIn)":"https://liv-in.lt/newtown/apartments",
    "Wilno Ozas(LivIn)": "https://liv-in.lt/ozas/apartments",
}

urls_camplus = {
    "Turyn(Camplus)":"https://www.camplus.it/prezzi-e-disponibilita/?city=torino",
    "Bologna(Camplus)":"https://www.camplus.it/prezzi-e-disponibilita/?city=bologna",
    "Milan(Camplus)":"https://www.camplus.it/prezzi-e-disponibilita/?city=milano",

}

urls_relife = {
    "Turyn(Relife)":"https://relifenation.com/en/struttura/torino/",
}

urls_campus_sanpaolo = {
    "Turyn(CampusSanPaolo)": "https://campussanpaolo.it/csp/#camere"
}

urls_Beyoo = {
    "Bologna(Beyoo)": "https://yugo.com/en-us/global/italy/bologna/laude-living-bologna/rooms",
    "Turyn(Beyoo)": "https://yugo.com/en-us/global/italy/turin/beyoo-taurasia-living-turin/rooms",
}

urls_indomus = {
    "Milano Internazionale(In-Domus)": "https://in-domus.it/en/campus-milano-internazionale",
    "Milano Olympia(In-Domus)": "https://in-domus.it/en/campus-milano-olympia",
    "Milano Monneret(In-Domus)": "https://in-domus.it/en/campus-milano-monneret",
}

urls_aparto = {
    "Milan(Aparto Students)": "https://apartostudent.com/find-a-room?tab=student&where=milan"
}

urls_sbsstudent = {
    "Karlstad(SBS Student)": "https://sbsstudent.se/en/available-accommodations/?qt_mll_search_tags=Karlstad",
    "Jönköping(SBS Student)": "https://sbsstudent.se/en/available-accommodations/?qt_mll_search_tags=J%C3%B6nk%C3%B6ping",
    "Malmö(SBS Student)": "https://sbsstudent.se/en/available-accommodations/?qt_mll_search_tags=Malm%C3%B6",
    "Stockholm(SBS Student)": "https://sbsstudent.se/en/available-accommodations/?qt_mll_search_tags=Stockholm",
}

urls_rikshem = {
    "Kalmar(Rikshem)":"https://minasidor.rikshem.se/ledigt/lagenhet",
}

urls_K2A = {
    "Orebro(K2A)": "https://minasidor.k2a.se/market/KBvv4jwRFvFgjb9xmJ8pHhvp?areas=HYrHFcTCrfywMfd7cTYFRd8C&pageSize=100&page=1"
}

urls_livetogrow = {
    "Stockholm Huddinge(Live to Grow)": "https://minasidor.byggvesta.se/market/THpPbVKc9yb4ftJDFwVdXYW7?areas=r68bqGQrJyDqyJGwcxcYMrbv",
    "linköping(Live to Grow)": "https://minasidor.byggvesta.se/market/THpPbVKc9yb4ftJDFwVdXYW7?areas=bgbBPV6McXdbTwTYbvqVmDg8",
    "Stockholm Kista(Live to Grow)": "https://minasidor.byggvesta.se/market/THpPbVKc9yb4ftJDFwVdXYW7?areas=hQpGHMc44T9dRf7GQpwqgTFm",
    "Stockholm Nacka(Live to Grow)": "https://minasidor.byggvesta.se/market/THpPbVKc9yb4ftJDFwVdXYW7?areas=wxDpphcjhytGvxdWP9pDTbdY",
    "Stockholm Norra(Live to Grow)": "https://minasidor.byggvesta.se/market/THpPbVKc9yb4ftJDFwVdXYW7?areas=XvmhQwVtP7RV9FWwbjG6W9rv",

}

urls_campusviva = {
    "Monachium Munich V(Campus Viva)": "https://www.campusviva.de/en/renting/munich/",
    "Monachium Munich VI(Campus Viva)": "https://www.campusviva.de/en/renting/muenchen-vi/",
    "Monachium Munich II(Campus Viva)": "https://www.campusviva.de/en/renting/munich-ii/",
}

urls_bookinghomeandco = {
    "Munich(Booking Home & Co)": "https://bookinghomeand.co/en/location/munich"
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

        if not spans:
            room_types.append("brak")
            prices.append("brak")
            continue
        room_type = spans[0].get_text(strip=True)
        price = "brak"

        if len(spans) >= 2:
            price_span = spans[1]
            for d in price_span.find_all("del"):
                d.decompose()

            price_text = price_span.get_text(" ", strip=True)
            m = re.search(r"\d[\d\s]*\s*(zł|€|eur|czk)", price_text, re.IGNORECASE)

            if m:
                price = m.group(0)
            else:
                price = price_text.strip()

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
            price_source = b.get_text(" ", strip=True)
        else:
            span_copy = BeautifulSoup(str(span), "html.parser")
            for s_tag in span_copy.find_all("s"):
                s_tag.decompose()
            price_source = span_copy.get_text(" ", strip=True)

        price_source = (price_source or "").replace("\u00A0", " ").strip()
        low = price_source.lower()


        if "€" in price_source or "eur" in low:
            currency = "€"
        elif "pln" in low or "zł" in low:
            currency = "zł"
        elif "sek" in low:
            currency = "SEK"
        elif "czk" in low:
            currency = "CZK"
        else:
            whole = span.get_text(" ", strip=True).replace("\u00A0", " ")
            whole_low = whole.lower()
            if "€" in whole or "eur" in whole_low:
                currency = "€"
            elif "pln" in whole_low or "zł" in whole_low:
                currency = "zł"
            elif "sek" in whole_low:
                currency = "SEK"
            elif "czk" in whole_low:
                currency = "CZK"
            else:
                currency = "zł"

        # --- LICZBA ---
        m = re.search(r"\d[\d\s\.,]*", price_source)
        raw = m.group(0) if m else ""

        raw = raw.replace(" ", "")
        raw = re.sub(r"\.(?=\d{3}\b)", "", raw)  # usuń separatory tys.
        raw = raw.replace(",", ".")
        raw = re.sub(r"[^\d]", "", raw)

        if raw:
            price_text = f"{raw} {currency}"
        else:
            price_text = "brak"

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

def parse_FizzPrague_page(soup):
    rows = []

    def eur_value(txt: str):
        m = re.search(r"€\s*([\d\.,]+)", txt)
        if not m:
            return None
        num = m.group(1).replace(",", "")
        try:
            return float(num)
        except:
            return None


    panels = soup.select("div.pex-room-type.panel.panel-default")
    if not panels:
        panels = soup.select("div.pex-room-type")

    for panel in panels:
        name_el = panel.select_one("div.text-center.below-image")
        name = name_el.get_text(" ", strip=True) if name_el else ""


        if not name:
            name_el = panel.select_one("h2, h3, .room-type-headings")
            name = name_el.get_text(" ", strip=True) if name_el else "brak"


        no_offer_el = panel.select_one("div.room-type-information-no-offer")
        if no_offer_el:
            rows.append({"Room Type": name, "Price": "brak"})
            continue


        candidates = []
        for sp in panel.select("span.pex-room-price"):
            txt = sp.get_text(" ", strip=True)
            low = txt.lower()

            if "€" not in txt:
                continue
            if ("monthly" not in low) and ("per month" not in low) and ("month" not in low) and ("/mo" not in low):
                continue

            v = eur_value(txt)
            if v is not None and v >= 100:  # filtr na śmieci
                candidates.append(v)

        if candidates:
            best = min(candidates)
            price_out = f"{best:.2f} EUR".replace(".00", "")
        else:
            price_out = "brak"

        rows.append({"Room Type": name, "Price": price_out})


    rows = [r for r in rows if r["Room Type"] and r["Room Type"] != "brak"]


    best_by_name = {}
    for r in rows:
        n, p = r["Room Type"], r["Price"]
        if p == "brak":
            best_by_name.setdefault(n, "brak")
            continue
        v = eur_value(p)
        if v is None:
            best_by_name.setdefault(n, p)
            continue
        if n not in best_by_name or best_by_name[n] == "brak" or eur_value(best_by_name[n]) > v:
            best_by_name[n] = p

    out = [{"Room Type": n, "Price": pr} for n, pr in best_by_name.items()]
    return pd.DataFrame(out)

def parse_chillhills_page(soup):

    def norm(t: str) -> str:
        return re.sub(r"\s+", " ", (t or "").replace("\u00A0", " ")).strip()


    start = soup.find(lambda tag: tag.name in ("h1", "h2", "h3")
                                 and norm(tag.get_text(" ", strip=True)).lower() in ("price list", "ceník", "ceník"))
    if not start:
        print("[ChillHills] ❌ Nie znalazłem sekcji 'Price list/Ceník'.")
        return pd.DataFrame([])
    rows = []
    expecting_name = True
    current_name = None

    for h3 in start.find_all_next("h3"):
        text = norm(h3.get_text(" ", strip=True))
        if not text:
            continue

        if expecting_name:

            if "deposit" in text.lower() or "kauce" in text.lower():
                current_name = None
                expecting_name = True
                continue

            current_name = text
            expecting_name = False
        else:

            price_txt = text


            if "deposit" in price_txt.lower() or "kauce" in price_txt.lower():
                current_name = None
                expecting_name = True
                continue


            if current_name:
                rows.append({"Room Type": current_name, "Price": price_txt})


            current_name = None
            expecting_name = True

    return pd.DataFrame(rows)

def parse_scandium_page(soup):
    rows = []


    cards = soup.select("li.object-card, li[class*='object-card']")
    if not cards:

        cards = soup.select("li:has(h3 a)")

    for card in cards:

        name_el = card.select_one("h3 a")
        room_type = name_el.get_text(" ", strip=True) if name_el else "brak"


        price_el = card.select_one("p.text-xs.font-bold")
        price_text = price_el.get_text(" ", strip=True) if price_el else "brak"


        if not price_text or price_text.strip() == "":
            price_text = "brak"


        price_text = price_text.split("/")[0].strip()


        if price_text != "brak" and not re.search(r"\d", price_text):
            price_text = "brak"


        if price_text != "brak" and re.search(r"\bdeposit\b|\bkauce\b", price_text, re.IGNORECASE):
            price_text = "brak"

        rows.append({"Room Type": room_type, "Price": price_text})


    rows = [r for r in rows if r["Room Type"] and r["Room Type"] != "brak"]
    return pd.DataFrame(rows)

def parse_new_neonwood_page(soup):
    rows = []
    items = soup.select("li.navigation-item")
    for it in items:
        title_el = it.select_one("h3.navigation-title")
        raw_title = title_el.get_text(" ", strip=True) if title_el else ""

        price_el = it.select_one("span.h2.price-figure")
        raw_price = price_el.get_text(" ", strip=True) if price_el else "brak"
        room_short = shorten_room_title_to_name_and_m2(raw_title) if raw_title else "brak"
        if room_short == "brak":
            continue
        rows.append({
            "Room Type": room_short,
            "Price": raw_price
        })
    return pd.DataFrame(rows)

def shorten_room_title_to_name_and_m2(title: str) -> str:

    if not title:
        return "brak"

    t = title.replace("\u00A0", " ").strip()


    t = t.replace("m²", "m2").replace("㎡", "m2")


    m = re.search(r"\b\d+\s*m\s*2\b", t, re.IGNORECASE)
    if not m:

        m = re.search(r"\b\d+\s*m\b", t, re.IGNORECASE)
        if not m:

            return t.split(" - ")[0].strip()

        end = m.end()
        left = t[:end].strip()
        left = left.replace("m", "m2")  # bardzo awaryjnie
        return left

    end = m.end()


    left = t[:end].strip()


    left = re.sub(r"\s+", " ", left)


    left = re.sub(r"(\d+)\s*m2\b", r"\1m2", left, flags=re.IGNORECASE)

    return left

def parse_youston_page(soup):
    rows = []


    cards = soup.select("div.stagger-item")
    for card in cards:

        name_el = card.select_one("div.apart-new h3, div.apart-contentTop h3, h3")
        room_type = name_el.get_text(" ", strip=True) if name_el else "brak"


        span = card.select_one("div.apart-price--month span")
        price_text = "brak"
        if span:
            visible = span.get_text(" ", strip=True)
            data_discount = (span.get("data-discount") or "").strip()


            if visible and re.search(r"\d", visible):
                price_text = visible
            elif data_discount and re.search(r"\d", data_discount):
                price_text = data_discount


        if price_text != "brak" and not re.search(r"\d", price_text):
            price_text = "brak"

        if room_type and room_type != "brak":
            rows.append({"Room Type": room_type, "Price": price_text})

    return pd.DataFrame(rows).drop_duplicates(subset=["Room Type"], keep="first")

def parse_duckrepublik_page(soup):
    rows = []

    cards = soup.select("div.grid-item__wrapper")
    for card in cards:

        name_el = card.select_one("span.grid-item__title")
        room_type = name_el.get_text(" ", strip=True) if name_el else "brak"


        price_el = card.select_one("div.grid-item__subtitle span")
        raw = price_el.get_text(" ", strip=True) if price_el else ""

        price_text = "brak"
        if raw:
            raw_norm = raw.replace("\u00A0", " ").strip()
            low = raw_norm.lower()


            if "€" in raw_norm or "eur" in low:
                m = re.search(r"\d[\d\s\.,]*", raw_norm)  # bierze pierwszą liczbę
                if m:
                    num = m.group(0)
                    num = num.replace(" ", "").replace(".", "")
                    num = num.split(",")[0]
                    num = re.sub(r"[^\d]", "", num)
                    if num:
                        price_text = f"{num} €"

        if room_type and room_type != "brak":
            rows.append({"Room Type": room_type, "Price": price_text})

    return pd.DataFrame(rows).drop_duplicates(subset=["Room Type"], keep="first")

def parse_duckrepublic2_page(soup):
    rows = []


    heads = soup.select("h3.wp-block-heading strong")
    for h in heads:
        room_type = h.get_text(" ", strip=True)
        if not room_type:
            continue


        p = h.find_parent("h3")
        price_text = "brak"

        if p:

            nxt = p.find_next("p")

            for _ in range(6):
                if not nxt:
                    break
                strong = nxt.find("strong")
                txt = strong.get_text(" ", strip=True) if strong else nxt.get_text(" ", strip=True)
                txt = (txt or "").replace("\u00A0", " ").strip()

                if "€" in txt or "eur" in txt.lower():
                    # "from 699 € / mo." -> "699 €"
                    m = re.search(r"\d[\d\s\.,]*", txt)
                    if m:
                        num = m.group(0).replace(" ", "").replace(".", "")
                        num = num.split(",")[0]
                        num = re.sub(r"[^\d]", "", num)
                        if num:
                            price_text = f"{num} €"
                    break

                nxt = nxt.find_next("p")

        if room_type and price_text != "brak":
            rows.append({"Room Type": room_type, "Price": price_text})

    return pd.DataFrame(rows).drop_duplicates(subset=["Room Type"], keep="first")

def parse_solosociety_page(soup):
    rows = []


    rooms = soup.select("div.divided-row")
    for room in rooms:

        name_el = room.select_one("h3.divided-title")
        room_type = name_el.get_text(" ", strip=True) if name_el else "brak"
        if not room_type or room_type == "brak":
            continue

        price_text = "brak"


        infos = room.select("div.divided-booking-info")
        for info in infos:
            unit_el = info.select_one("span.small-text")
            unit = unit_el.get_text(" ", strip=True).lower() if unit_el else ""


            if "month" not in unit:
                continue

            price_el = info.select_one("p.divided-booking-price")
            raw = price_el.get_text(" ", strip=True) if price_el else ""
            raw = raw.replace("\u00A0", " ").strip()


            m = re.search(r"(\d[\d\s\.,]*)\s*(€|eur|zł|pln|czk)", raw, re.IGNORECASE)
            if m:
                num, cur = m.group(1), m.group(2)
                num = num.replace(" ", "").replace(".", "")
                num = num.split(",")[0]
                num = re.sub(r"[^\d]", "", num)

                if num:
                    cur_low = cur.lower()
                    if cur_low in ["€", "eur"]:
                        price_text = f"{num} €"
                    elif cur_low in ["pln", "zł"]:
                        price_text = f"{num} zł"
                    else:
                        price_text = f"{num} CZK"
            else:

                num_m = re.search(r"\d[\d\s\.,]*", raw)
                if num_m and ("€" in raw or "eur" in raw.lower()):
                    num = num_m.group(0).replace(" ", "").replace(".", "")
                    num = num.split(",")[0]
                    num = re.sub(r"[^\d]", "", num)
                    if num:
                        price_text = f"{num} €"


            if price_text != "brak":
                break


        if price_text != "brak":
            rows.append({"Room Type": room_type, "Price": price_text})

    return pd.DataFrame(rows).drop_duplicates(subset=["Room Type"], keep="first")

def parse_livin_page(soup):
    rows = []

    cards = soup.select("div.apartments-card")
    for card in cards:
        title_el = card.select_one("div.apartments__title")
        size_el = card.select_one("div.apartments__size")
        price_el = card.select_one("p.apartments__price span.apartments__price--bold")

        title = title_el.get_text(" ", strip=True) if title_el else ""
        size = size_el.get_text(" ", strip=True) if size_el else ""
        price_raw = price_el.get_text(" ", strip=True) if price_el else ""

        if not title:
            continue


        size = size.replace("\u00A0", " ").strip()
        size = size.replace("m²", "m2").replace("㎡", "m2")
        size = re.sub(r"\s+", " ", size)
        size = re.sub(r"(\d+)\s*m2", r"\1m2", size, flags=re.IGNORECASE)  # "20 m2" -> "20m2"

        room_type = f"{title} {size}".strip() if size else title


        price_text = "brak"
        if price_raw:
            pr = price_raw.replace("\u00A0", " ").strip()

            m = re.search(r"\d[\d\s\.,]*", pr)
            if m:
                num = m.group(0).replace(" ", "").replace(".", "")
                num = num.split(",")[0]
                num = re.sub(r"[^\d]", "", num)

                if num:
                    if "€" in pr or "eur" in pr.lower():
                        price_text = f"{num} €"
                    elif "zł" in pr.lower() or "pln" in pr.lower():
                        price_text = f"{num} zł"
                    elif "czk" in pr.lower():
                        price_text = f"{num} CZK"

        if price_text != "brak":
            rows.append({"Room Type": room_type, "Price": price_text})

    return pd.DataFrame(rows).drop_duplicates(subset=["Room Type"], keep="first")

def parse_camplus(page) -> pd.DataFrame:
    rows = []


    page.wait_for_selector('div[class*="RatesAdsItem-root"]', timeout=60000, state="attached")

    cards = page.locator('div[class*="RatesAdsItem-root"]')
    n = cards.count()
    if n == 0:
        return pd.DataFrame([])

    for i in range(n):
        card = cards.nth(i)


        title_el = card.locator('h6[class*="RatesAdsItem-title"]').first
        room_type = title_el.inner_text().strip() if title_el.count() else "brak"
        if not room_type or room_type == "brak":
            continue


        price_el = card.locator('p[class*="RatesAdsItem-price"]').first
        raw = price_el.inner_text().replace("\u00A0", " ").strip().lower() if price_el.count() else ""
        if not raw:
            continue


        is_eur = ("€" in raw) or ("eur" in raw) or ("euro" in raw)


        m_range = re.search(r"\b(?:da|od)\s*(\d[\d\s\.,]*)\s*(?:a|do|-)\s*(\d[\d\s\.,]*)", raw)
        if m_range:
            chosen = m_range.group(2)
        else:

            m_one = re.search(r"\d[\d\s\.,]*", raw)
            chosen = m_one.group(0) if m_one else ""

        num = chosen.replace(" ", "").replace(".", "")
        num = num.split(",")[0]
        num = re.sub(r"[^\d]", "", num)

        if not num:
            continue

        price_text = f"{num} €" if is_eur else f"{num}"
        rows.append({"Room Type": room_type, "Price": price_text})

    return pd.DataFrame(rows).drop_duplicates(subset=["Room Type"], keep="first")

def parse_relifenation(page) -> pd.DataFrame:
    rows = []

    page.wait_for_selector("div.campus-room-info h3", timeout=60000, state="attached")
    page.wait_for_timeout(1200)


    slides = page.locator("div.swiper-slide")
    n = slides.count()
    if n == 0:
        return pd.DataFrame([])

    for i in range(n):
        slide = slides.nth(i)

        name_el = slide.locator("div.campus-room-info h3").first
        if name_el.count() == 0:
            continue
        room_type = name_el.inner_text().replace("\u00A0", " ").strip()
        if not room_type:
            continue


        price_text = "brak"
        price_el = slide.locator("div.campus-room-popup-info p").first
        if price_el.count():
            txt = price_el.inner_text().replace("\u00A0", " ").strip()

            if "€" in txt or "eur" in txt.lower() or "euro" in txt.lower():
                m = re.search(r"\d[\d\s\.,]*", txt)
                if m:
                    num = m.group(0).replace(" ", "").replace(".", "")
                    num = num.split(",")[0]
                    num = re.sub(r"[^\d]", "", num)
                    if num:
                        price_text = f"{num} €"

        rows.append({"Room Type": room_type, "Price": price_text})

    return pd.DataFrame(rows).drop_duplicates(subset=["Room Type"], keep="first")

def parse_cx_places(page) -> pd.DataFrame:
    rows = []


    page.wait_for_selector("div.fieldvalue.f1.sf0", timeout=60000, state="attached")
    page.wait_for_selector("div.fieldvalue.f10.sf0", timeout=60000, state="attached")
    page.wait_for_timeout(800)


    slides = page.locator("div.slick-slide:has(div.fieldvalue.f1.sf0)")
    n = slides.count()
    if n == 0:
        return pd.DataFrame([])

    for i in range(n):
        slide = slides.nth(i)

        name_el = slide.locator("div.fieldvalue.f1.sf0").first
        price_el = slide.locator("div.fieldvalue.f10.sf0").first

        room_type = name_el.inner_text().replace("\u00A0", " ").strip() if name_el.count() else ""
        raw_price = price_el.inner_text().replace("\u00A0", " ").strip() if price_el.count() else ""

        if not room_type:
            continue


        price_text = "brak"
        if raw_price and re.search(r"\d", raw_price):
            is_eur = ("€" in raw_price) or ("eur" in raw_price.lower()) or ("euro" in raw_price.lower())
            m = re.search(r"\d[\d\s\.,]*", raw_price)
            if m:
                num = m.group(0).replace(" ", "").replace(".", "")
                num = num.split(",")[0]
                num = re.sub(r"[^\d]", "", num)
                if num:
                    price_text = f"{num} €" if is_eur else num

        rows.append({"Room Type": room_type, "Price": price_text})

    return pd.DataFrame(rows).drop_duplicates(subset=["Room Type"], keep="first")

def parse_campus_sanpaolo_page(soup) -> pd.DataFrame:
    rows = []


    headers = soup.select("h2.av-special-heading-tag")
    if not headers:
        return pd.DataFrame([])

    for h in headers:
        room_type = h.get_text(" ", strip=True).replace("\u00A0", " ").strip()
        if not room_type:
            continue


        container = h.parent
        for _ in range(8):
            if not container:
                break
            if container.select_one("a.avia-button span.avia_iconbox_title"):
                break
            container = container.parent

        price_text = "brak"
        if container:
            price_el = container.select_one("a.avia-button span.avia_iconbox_title")
            if price_el:
                raw = price_el.get_text(" ", strip=True).replace("\u00A0", " ").strip()

                m = re.search(r"(\d[\d\s\.,]*)\s*€", raw)
                if m:
                    num = m.group(1).replace(" ", "").replace(".", "")
                    num = num.split(",")[0]
                    num = re.sub(r"[^\d]", "", num)
                    if num:
                        price_val = int(num)
                        price_text = f"{price_val} €"


        rt_low = room_type.lower()
        is_double = (
            "dwuosob" in rt_low
            or "2 osob" in rt_low
            or "2-osob" in rt_low
            or "2osob" in rt_low
            or re.search(r"\b2\s*(person|persons|people)\b", rt_low)
            or re.search(r"\bfor\s*2\b", rt_low)
        )

        if is_double and price_text != "brak":
            m2 = re.search(r"(\d+)\s*€", price_text)
            if m2:
                v = int(m2.group(1))
                price_text = f"{v * 2} €"

        rows.append({"Room Type": room_type, "Price": price_text})


    rows = [r for r in rows if r["Room Type"]]
    return pd.DataFrame(rows).drop_duplicates(subset=["Room Type"], keep="first")

def parse_beyoo_rooms(soup) -> pd.DataFrame:
    rows = []

    cards = soup.select("article")
    if not cards:

        cards = soup.select("h4.product-tile__title")


    def extract_eur_price(text: str) -> str:
        if not text:
            return "brak"
        t = text.replace("\u00A0", " ").strip()


        m = re.search(r"€\s*([\d\.,]+)", t)
        if not m:
            return "brak"

        num = m.group(1)
        num = num.replace(",", "")
        try:
            value = float(num)
        except ValueError:
            return "brak"

        return f"{int(round(value))} EUR"


    if cards and hasattr(cards[0], "select"):
        for card in cards:
            name_el = card.select_one("h4.product-tile__title")
            room_type = name_el.get_text(" ", strip=True) if name_el else "brak"

            price_el = card.select_one("h5.product-tile__subtitle")
            raw_price = price_el.get_text(" ", strip=True) if price_el else ""

            price_text = extract_eur_price(raw_price)

            if room_type != "brak":
                rows.append({"Room Type": room_type, "Price": price_text})

    else:

        for h4 in soup.select("h4.product-tile__title"):
            room_type = h4.get_text(" ", strip=True) or "brak"
            price_el = h4.find_next("h5", class_="product-tile__subtitle")
            raw_price = price_el.get_text(" ", strip=True) if price_el else ""
            price_text = extract_eur_price(raw_price)
            if room_type != "brak":
                rows.append({"Room Type": room_type, "Price": price_text})

    return pd.DataFrame(rows).drop_duplicates(subset=["Room Type"], keep="first")

def parse_indomus_page(soup):
    rows = []

    price_blocks = soup.select('div[id^="accommodation-"][id$="-price"]')

    for pb in price_blocks:
        card = pb.find_parent("div", class_=lambda c: c and "h-100" in c and "flex-column" in c)
        if not card:
            continue


        name_el = card.select_one('h3[id^="accommodation-"], h3.content-title')
        room_type = name_el.get_text(" ", strip=True) if name_el else "brak"


        price_el = pb.select_one("span.price")
        price_text = "brak"
        if price_el:
            raw = price_el.get_text(" ", strip=True).replace("\u00A0", " ").strip()

            m = re.search(r"(\d{1,3}(?:\.\d{3})*(?:,\d{2})?|\d+)", raw)
            if m:
                num = m.group(1)
                num = num.replace(".", "")
                num = num.split(",")[0]
                num = re.sub(r"[^\d]", "", num)

                if num:
                    price_text = f"{int(num)} €"

        if room_type != "brak" and price_text != "brak":
            rows.append({
                "Room Type": room_type,
                "Price": price_text
            })

    return pd.DataFrame(rows).drop_duplicates(subset=["Room Type"], keep="first")


def parse_collegiate_page(page) -> pd.DataFrame:
    rows = []

    page.wait_for_timeout(2500)

    cards = page.locator('[data-room][data-price]')
    count = cards.count()
    print(f"[Collegiate] cards in DOM: {count}")

    for i in range(count):
        card = cards.nth(i)

        room_type = (card.get_attribute("data-room") or "").strip()
        raw_price = (card.get_attribute("data-price") or "").strip()

        if not room_type:
            continue

        price_text = "brak"
        if raw_price:
            try:
                value = float(raw_price.replace(",", ""))
                price_text = f"{int(round(value))} €"
            except ValueError:
                price_text = "brak"

        if price_text != "brak":
            rows.append({
                "Room Type": room_type,
                "Price": price_text
            })

    return pd.DataFrame(rows).drop_duplicates(subset=["Room Type"], keep="first")

def parse_sbsstudent_page(soup):
    rows = []

    cards = soup.select("a.card.card--listing")
    for card in cards:

        title_el = card.select_one("h2.card__title")
        title_text = title_el.get_text(" ", strip=True).replace("\u00A0", " ").strip() if title_el else "brak"


        price_text = "brak"
        details = card.select("ul.listing__details li")

        for li in details:
            txt = li.get_text(" ", strip=True).replace("\u00A0", " ").strip()
            low = txt.lower()

            if "sek" in low:
                m = re.search(r"(\d[\d\s\.,]*)\s*sek", txt, re.IGNORECASE)
                if m:
                    num = m.group(1)
                    num = num.replace(" ", "").replace(".", "")
                    num = num.split(",")[0]
                    num = re.sub(r"[^\d]", "", num)
                    if num:
                        price_text = f"{num} SEK"
                        break


        area_text = ""
        for li in details:
            txt = li.get_text(" ", strip=True).replace("\u00A0", " ").strip()

            m = re.search(r"(\d+(?:[.,]\d+)?)\s*m(?:²|2)", txt, re.IGNORECASE)
            if m:
                area = m.group(1).replace(",", ".")
                if area.endswith(".0"):
                    area = area[:-2]
                area_text = f"{area} m2"
                break


        room_type = "brak"
        if title_text != "brak":
            if area_text:
                room_type = f"{title_text} - {area_text}"
            else:
                room_type = title_text

        if room_type != "brak" and price_text != "brak":
            rows.append({
                "Room Type": room_type,
                "Price": price_text
            })

    return pd.DataFrame(rows).drop_duplicates(subset=["Room Type", "Price"], keep="first")


def parse_livetogrow_page(page) -> pd.DataFrame:
    rows = []

    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(2500)

    title_locs = page.locator(
        'h2[data-mom-test="MinaSidor_MittSökande_AnnonsRubrik"], h2.mat-h2'
    )
    title_count = title_locs.count()

    # jeśli brak ofert -> zwróć pusty DF bez błędu
    if title_count == 0:
        print("[Live to Grow] Brak ofert na stronie - skip")
        return pd.DataFrame(columns=["Room Type", "Price", "Room Type Key"])

    titles = []
    for i in range(title_count):
        txt = title_locs.nth(i).inner_text().replace("\u00A0", " ").strip()
        if txt:
            titles.append(txt)

    item_locs = page.locator("div.details-item")
    item_count = item_locs.count()

    prices = []
    for i in range(item_count):
        item = item_locs.nth(i)

        label_el = item.locator("div.details-item-label").first
        value_el = item.locator("div.details-item-value").first

        label = label_el.inner_text().replace("\u00A0", " ").strip().lower() if label_el.count() else ""
        value = value_el.inner_text().replace("\u00A0", " ").strip() if value_el.count() else ""

        label_norm = (
            label.replace("å", "a")
                 .replace("ä", "a")
                 .replace("ö", "o")
                 .replace(" ", "")
        )

        if "kr" in label_norm and "man" in label_norm:
            m = re.search(r"\d[\d\s.,]*", value)
            if m:
                num = m.group(0)
                num = num.replace(" ", "").replace(".", "")
                num = num.split(",")[0]
                num = re.sub(r"[^\d]", "", num)

                if num:
                    prices.append(f"{num} SEK")

    print(f"[Live to Grow] titles: {len(titles)}")
    print(f"[Live to Grow] prices: {len(prices)}")

    n = min(len(titles), len(prices))
    for i in range(n):
        rows.append({
            "Room Type": titles[i],
            "Price": prices[i]
        })

    df = pd.DataFrame(rows)

    if df.empty:
        return pd.DataFrame(columns=["Room Type", "Price", "Room Type Key"])


    df = df.drop_duplicates(subset=["Room Type", "Price"], keep="first").reset_index(drop=True)


    df["_dup_no"] = df.groupby("Room Type").cumcount() + 1
    df["Room Type Key"] = df.apply(
        lambda row: f'{normalize_room_type(row["Room Type"])}__{row["_dup_no"]}',
        axis=1
    )
    df = df.drop(columns=["_dup_no"])

    return df

def parse_k2a_page(page) -> pd.DataFrame:
    rows = []

    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(2500)

    title_locs = page.locator(
        'h2[data-mom-test="MinaSidor_MittSökande_AnnonsRubrik"], h2.mat-h2'
    )
    title_count = title_locs.count()

    if title_count == 0:
        print("[K2A] Brak ofert na stronie - skip")
        return pd.DataFrame(columns=["Room Type", "Price", "Room Type Key"])

    titles = []
    for i in range(title_count):
        txt = title_locs.nth(i).inner_text().replace("\u00A0", " ").strip()
        if txt:
            titles.append(txt)

    item_locs = page.locator("div.details-item")
    item_count = item_locs.count()

    prices = []
    for i in range(item_count):
        item = item_locs.nth(i)

        label_el = item.locator("div.details-item-label").first
        value_el = item.locator("div.details-item-value").first

        label = label_el.inner_text().replace("\u00A0", " ").strip().lower() if label_el.count() else ""
        value = value_el.inner_text().replace("\u00A0", " ").strip() if value_el.count() else ""

        label_norm = (
            label.replace("å", "a")
                 .replace("ä", "a")
                 .replace("ö", "o")
                 .replace(" ", "")
        )

        if "kr" in label_norm and "man" in label_norm:
            m = re.search(r"\d[\d\s.,]*", value)
            if m:
                num = m.group(0)
                num = num.replace(" ", "").replace(".", "")
                num = num.split(",")[0]
                num = re.sub(r"[^\d]", "", num)

                if num:
                    prices.append(f"{num} SEK")

    print(f"[K2A] titles: {len(titles)}")
    print(f"[K2A] prices: {len(prices)}")

    n = min(len(titles), len(prices))
    for i in range(n):
        rows.append({
            "Room Type": titles[i],
            "Price": prices[i]
        })

    df = pd.DataFrame(rows)

    if df.empty:
        return pd.DataFrame(columns=["Room Type", "Price", "Room Type Key"])


    df = df.drop_duplicates(subset=["Room Type", "Price"], keep="first").reset_index(drop=True)


    df["_dup_no"] = df.groupby("Room Type").cumcount() + 1
    df["Room Type Key"] = df.apply(
        lambda row: f'{normalize_room_type(row["Room Type"])}__{row["_dup_no"]}',
        axis=1
    )
    df = df.drop(columns=["_dup_no"])

    return df








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

    price_str = price_str.replace("\u00A0", " ").strip()

    # specjalny przypadek typu: 1.215,- €
    price_str = price_str.replace(",- €", " EUR")
    price_str = price_str.replace(",-€", " EUR")
    price_str = price_str.replace(" €", " EUR")
    price_str = price_str.replace("€", " EUR")

    match = re.search(r'([\d\s,\.\-]+)\s*(zł|eur|czk|sek)', price_str, re.IGNORECASE)

    if match:
        number_str, currency = match.groups()

        number_str = number_str.replace(" ", "").replace("\u00A0", "")
        number_str = number_str.replace(".-", "")
        number_str = number_str.replace(",-", "")
        number_str = number_str.replace(".", "")
        number_str = number_str.replace(",", ".")

        try:
            number = float(number_str)
        except ValueError:
            return price_str

        number_int = int(round(number))

        cur = currency.lower()
        if cur == "eur":
            currency_str = "EUR"
        elif cur == "czk":
            currency_str = "CZK"
        elif cur == "sek":
            currency_str = "SEK"
        else:
            currency_str = "zł"

        return f"{number_int} {currency_str}"

    else:
        low = price_str.lower()

        if "eur" in low:
            currency_str = "EUR"
        elif "czk" in low:
            currency_str = "CZK"
        elif "sek" in low or "kr" in low:
            currency_str = "SEK"
        else:
            currency_str = "zł"

        cleaned_number = re.sub(r"[^\d,.\-]", "", price_str)
        cleaned_number = cleaned_number.replace(".-", "")
        cleaned_number = cleaned_number.replace(",-", "")
        cleaned_number = cleaned_number.replace(".", "")
        cleaned_number = cleaned_number.replace(",", ".")

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
"""
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
"""

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

COUNTRY_CITIES = {
    "Polska": {
        "warszawa", "kraków", "katowice", "gdańsk", "lublin", "poznań", "wrocław",
        "łódź", "lódź"
    },
    "Szwecja": {
        "stockholm", "linkoping", "linköping", "jönköping", "jonkoping", "malmö", "malmo",
        "karlstad", "kalmar", "orebro", "örebro"
    },
    "Litwa": {
        "wilno", "vilnius", "kaunas"
    },
    "Łotwa": {
        "riga"
    },
    "Estonia": {
        "tallinn"
    },
    "Niemcy": {
        "berlin", "monachium", "munich"
    },
    "Czechy": {
        "prague", "praga", "brno"
    },
    "Włochy": {
        "milan", "milano", "turyn", "torino", "bologna", "florence", "florencja", "rome", "rzym"
    }
}

def get_country_from_location(location: str) -> str:
    if not isinstance(location, str):
        return "Inne"

    loc = normalize_text(location)

    for country, cities in COUNTRY_CITIES.items():
        if loc in cities:
            return country

    return "Inne"

def refresh_data():
    current_snapshot = []

    def add_data(academy_key, df: pd.DataFrame):
        loc, obj, owner = extract_academy_info(academy_key)
        for _, row in df.iterrows():
            price_clean = clean_price(str(row["Price"]))
            raw_room_type = str(row["Room Type"]).strip()

            if "Room Type Key" in df.columns and pd.notna(row.get("Room Type Key")):
                room_key = str(row["Room Type Key"]).strip()
            else:
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


        for academy_key, url in urls_FizzPrague.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (TheFizz)...")
            page.goto(url, timeout=60000, wait_until="networkidle")


            page.wait_for_selector("span.pex-room-price", timeout=20000, state="attached")

            soup = BeautifulSoup(page.content(), "html.parser")
            df = parse_FizzPrague_page(soup)

            print(f"[TheFizz] rekordów: {len(df)}")
            if df.empty:
                print("[TheFizz] ❗ Nic nie znaleziono (pusty DF)")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_chillhills.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (ChillHills)...")
            soup = load_page_soup(page, url)
            if soup is None:
                continue
            df = parse_chillhills_page(soup)
            add_data(academy_key, df)

        for academy_key, url in urls_scandium.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Scandium)...")


            page.goto(url, timeout=60000, wait_until="networkidle")


            page.wait_for_selector("li.object-card h3 a", timeout=20000)

            soup = BeautifulSoup(page.content(), "html.parser")
            df = parse_scandium_page(soup)

            print(f"[Scandium] rekordów: {len(df)}")
            if df.empty:
                print("[Scandium] ❗ Nic nie znaleziono (pusty DF)")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_neonwood.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Neonwood)...")

            page.goto(url, timeout=60000, wait_until="networkidle")
            page.wait_for_selector("h3.navigation-title", timeout=30000, state="attached")
            page.wait_for_selector("span.h2.price-figure", timeout=30000, state="attached")

            html = page.content()
            soup = BeautifulSoup(html, "html.parser")
            df = parse_new_neonwood_page(soup)
            add_data(academy_key, df)

            print(f"[Neonwood] rekordów: {len(df)}")
            if df.empty:
                print("[Neonwood] ❗ Nic nie znaleziono (pusty DF)")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_youston.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Youston)...")

            page.goto(url, timeout=60000, wait_until="networkidle")

            # Youston często ładuje sekcję po scrollu
            page.mouse.wheel(0, 2500)
            page.wait_for_timeout(1500)

            soup = BeautifulSoup(page.content(), "html.parser")
            df = parse_youston_page(soup)

            print(f"[Youston] rekordów: {len(df)}")
            if df.empty:
                print("[Youston] ❗ Nic nie znaleziono (pusty DF)")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_duckrepublic.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Duck Republik)...")

            page.goto(url, timeout=60000, wait_until="domcontentloaded")

            page.goto(url, timeout=60000, wait_until="domcontentloaded")
            page.wait_for_selector("div.grid-item__wrapper", timeout=30000, state="attached")
            page.wait_for_selector("div.grid-item__subtitle span", timeout=30000, state="attached")
            page.mouse.wheel(0, 4000)
            page.wait_for_timeout(1200)

            soup = BeautifulSoup(page.content(), "html.parser")
            df = parse_duckrepublik_page(soup)

            print(f"[Duck Republik] rekordów: {len(df)}")
            if df.empty:
                print("[Duck Republik] ❗ Nic nie znaleziono (pusty DF)")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_Duckrepublic.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Duckrepublic)...")

            page.goto(url, timeout=60000, wait_until="domcontentloaded")
            page.wait_for_timeout(1200)

            soup = BeautifulSoup(page.content(), "html.parser")
            df = parse_duckrepublic2_page(soup)

            print(f"[Duckrepublic] rekordów: {len(df)}")
            if df.empty:
                print("[Duckrepublic] ❗ Nic nie znaleziono (pusty DF)")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_solosociety.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (solosociety)...")

            page.goto(url, timeout=60000, wait_until="domcontentloaded")
            page.wait_for_timeout(1500)

            soup = BeautifulSoup(page.content(), "html.parser")
            df = parse_solosociety_page(soup)

            print(f"[solosociety] rekordów: {len(df)}")
            if df.empty:
                print("[solosociety] ❗ Nic nie znaleziono (pusty DF)")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_livin.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (livin)...")

            page.goto(url, timeout=60000, wait_until="domcontentloaded")
            page.wait_for_timeout(1200)

            soup = BeautifulSoup(page.content(), "html.parser")
            df = parse_livin_page(soup)

            print(f"[livin] rekordów: {len(df)}")
            if df.empty:
                print("[livin] ❗ Nic nie znaleziono (pusty DF)")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_camplus.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Camplus)...")

            page.goto(url, timeout=60000, wait_until="domcontentloaded")
            page.wait_for_timeout(2000)

            df = parse_camplus(page)

            print(f"[Camplus] rekordów: {len(df)}")
            if df.empty:
                print("[Camplus] ❗ Nic nie znaleziono (pusty DF)")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_relife.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (ReLifeNation)...")
            page.goto(url, timeout=60000, wait_until="domcontentloaded")
            page.wait_for_timeout(2000)

            df = parse_relifenation(page)

            print(f"[ReLifeNation] rekordów: {len(df)}")
            if df.empty:
                print("[ReLifeNation] ❗ Nic nie znaleziono (pusty DF)")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_CXplaces.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (CX)...")

            page.goto(url, timeout=60000, wait_until="domcontentloaded")
            page.wait_for_timeout(1500)

            df = parse_cx_places(page)

            print(f"[CX] rekordów: {len(df)}")
            if df.empty:
                print("[CX] ❗ Nic nie znaleziono (pusty DF)")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_campus_sanpaolo.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (CampusSanPaolo)...")
            soup = load_page_soup(page, url)
            if soup is None:
                continue

            df = parse_campus_sanpaolo_page(soup)

            print(f"[CampusSanPaolo] rekordów: {len(df)}")
            if df.empty:
                print("[CampusSanPaolo] ❗ Nic nie znaleziono (pusty DF)")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_Beyoo.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Beyoo)...")

            page.goto(url, timeout=90000, wait_until="networkidle")
            page.wait_for_timeout(1500)

            soup = BeautifulSoup(page.content(), "html.parser")
            df = parse_beyoo_rooms(soup)

            print(f"[Beyoo] rekordów: {len(df)}")
            if df.empty:
                print("[Beyoo] ❗ Nic nie znaleziono")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_indomus.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (In-Domus)...")

            page.goto(url, timeout=60000, wait_until="domcontentloaded")
            page.wait_for_timeout(1500)

            soup = BeautifulSoup(page.content(), "html.parser")
            df = parse_indomus_page(soup)

            print(f"[In-Domus] rekordów: {len(df)}")
            if df.empty:
                print("[In-Domus] ❗ Nic nie znaleziono (pusty DF)")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_collegiate.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Collegiate)...")

            page.goto(url, timeout=90000, wait_until="domcontentloaded")
            page.wait_for_timeout(3000)


            page.mouse.wheel(0, 2500)
            page.wait_for_timeout(1500)

            df = parse_collegiate_page(page)

            print(f"[Collegiate] rekordów: {len(df)}")
            if df.empty:
                print("[Collegiate] ❗ Nic nie znaleziono (pusty DF)")
                continue

            add_data(academy_key, df)

        for academy_key, url in urls_sbsstudent.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (SBSstudent)...")

            all_dfs = []

            # strona 1
            page.goto(url, timeout=90000, wait_until="domcontentloaded")
            page.wait_for_timeout(2000)

            soup = BeautifulSoup(page.content(), "html.parser")
            df1 = parse_sbsstudent_page(soup)
            if not df1.empty:
                all_dfs.append(df1)

            # strona 2
            try:
                page.goto("https://sbsstudent.se/en/available-accommodations/?qt_mll_search_tags=Karlstad&paged=2",
                          timeout=90000, wait_until="domcontentloaded")
                page.wait_for_timeout(2000)

                soup = BeautifulSoup(page.content(), "html.parser")
                df2 = parse_sbsstudent_page(soup)

                if not df2.empty:
                    all_dfs.append(df2)

            except Exception as e:
                print(f"[SBSstudent] ⚠️ Nie udało się pobrać strony 2: {e}")

            if not all_dfs:
                print("[SBSstudent] ❗ Nic nie znaleziono (pusty DF)")
                continue

            df = pd.concat(all_dfs, ignore_index=True).drop_duplicates(subset=["Room Type", "Price"], keep="first")

            print(f"[SBSstudent] rekordów: {len(df)}")
            add_data(academy_key, df)


        for academy_key, url in urls_livetogrow.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (Live to Grow)...")

            try:
                page.goto(url, timeout=90000, wait_until="networkidle")
                page.wait_for_timeout(1500)

                df = parse_livetogrow_page(page)

                print(f"[Live to Grow] rekordów: {len(df)}")
                print(df)

                if df.empty:
                    print("[Live to Grow] ❗ Nic nie znaleziono (pusty DF)")
                    continue

                add_data(academy_key, df)

            except Exception as e:
                print(f"[Live to Grow] ❌ Błąd: {e}")

        for academy_key, url in urls_K2A.items():
            print(f"🔄 Pobieranie danych dla: {academy_key} (K2A)...")

            try:
                page.goto(url, timeout=90000, wait_until="networkidle")
                page.wait_for_timeout(1500)

                df = parse_k2a_page(page)

                print(f"[K2A] rekordów: {len(df)}")
                print(df)

                if df.empty:
                    print("[K2A] ❗ Nic nie znaleziono (pusty DF)")
                    continue

                add_data(academy_key, df)

            except Exception as e:
                print(f"[K2A] ❌ Błąd: {e}")







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
    excel_file = "StudentDepot_All_Academies_Room_Prices1.xlsx"
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
    merged_df["Kraj"] = merged_df["Lokalizacja"].apply(get_country_from_location)
    cols = merged_df.columns.tolist()

    cols.remove("Kraj")
    loc_index = cols.index("Lokalizacja")
    cols.insert(loc_index + 1, "Kraj")

    merged_df = merged_df[cols]
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        workbook = writer.book

        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D7E4BC',
            'border': 1
        })

        red_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})

        # opcjonalnie: zakładka zbiorcza
        all_df = merged_df.copy()
        all_df = all_df.sort_values(by=["Kraj", "Lokalizacja", "Obiekt", "Typ pokoju"])
        all_df.to_excel(writer, index=False, sheet_name='Wszystkie dane')
        worksheet_all = writer.sheets['Wszystkie dane']

        for col_num, value in enumerate(all_df.columns.values):
            worksheet_all.write(0, col_num, value, header_format)

        for i, col in enumerate(all_df.columns):
            column_len = all_df[col].fillna("").astype(str).str.len().max()
            column_len = max(column_len, len(col)) + 2
            worksheet_all.set_column(i, i, column_len)

        # osobne zakładki dla krajów
        country_order = ["Polska", "Szwecja", "Litwa", "Łotwa", "Estonia", "Niemcy", "Czechy", "Włochy", "Inne"]

        for country in country_order:
            country_df = merged_df[merged_df["Kraj"] == country].copy()

            if country_df.empty:
                continue

            country_df = country_df.sort_values(by=["Lokalizacja", "Obiekt", "Typ pokoju"])

            # możesz usunąć kolumnę Kraj z arkusza kraju, bo i tak nazwa zakładki mówi jaki to kraj
            country_df_to_save = country_df.drop(columns=["Kraj"])

            sheet_name = country[:31]  # Excel ma limit 31 znaków
            country_df_to_save.to_excel(writer, index=False, sheet_name=sheet_name)
            worksheet = writer.sheets[sheet_name]

            # format nagłówków
            for col_num, value in enumerate(country_df_to_save.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # szerokości kolumn
            for i, col in enumerate(country_df_to_save.columns):
                column_len = country_df_to_save[col].fillna("").astype(str).str.len().max()
                column_len = max(column_len, len(col)) + 2
                worksheet.set_column(i, i, column_len)

            # kolorowanie zmian cen
            key_cols_display = ['Właściciel', 'Lokalizacja', 'Obiekt', 'Typ pokoju']
            snapshot_cols = [c for c in country_df_to_save.columns if c not in key_cols_display]

            start_row = 1
            end_row = len(country_df_to_save)
            first_excel_row = start_row + 1

            if len(snapshot_cols) >= 2:
                for i in range(1, len(snapshot_cols)):
                    prev_idx = country_df_to_save.columns.get_loc(snapshot_cols[i - 1])
                    new_idx = country_df_to_save.columns.get_loc(snapshot_cols[i])

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

                    worksheet.conditional_format(
                        start_row, new_idx, end_row, new_idx,
                        {'type': 'formula', 'criteria': greater_formula, 'format': red_fmt}
                    )
                    worksheet.conditional_format(
                        start_row, new_idx, end_row, new_idx,
                        {'type': 'formula', 'criteria': less_formula, 'format': green_fmt}
                    )

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

"""
    send_email_outlook(
        from_addr=FROM_ADDR,
        to_addrs=["jakub.swiniarski@studentdepot.pl", "adam.swiniarski@studentdepot.pl","sylwia.rogalska@studentdepot.pl"],
        subject="Tygodniowa aktualizacja cen akademików",
        body_text="Cześć,\nW załączniku przesyłam najnowszy plik Excel z cenami akademików.",
        attachment_path=excel_file,
        smtp_username=SMTP_USERNAME,
        smtp_password=SMTP_PASSWORD
    )
"""

if __name__ == "__main__":
    refresh_data()













