import xlwings as xw
from datetime import datetime
from tkinter import *
from tkinter import messagebox
import logging

# Excel fájl megnyitása
file = xw.Book("D:\\Leltar\\IT_leltar.xlsm")
sheet = file.sheets["Leltár"]

# Naplózás
log_file = open(r"D:\Leltar\log.txt")
logging.basicConfig(filename=r"D:\Leltar\log.txt", level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s', filemode='a')

# Alert üzenet
msgbox = Tk()

# Kategóriák és alkategóriák szótárak
categories = {'Munka állomás': 'WST', 'Kijelző': 'DSP', 'Nyomtató': 'PRT', 'Telefon': 'PHO',
              'Fülhallgató': 'HDS', 'Szerver': 'SRV', 'Hálózati eszköz': 'NET', 'PDA': 'PDA',
              'Egyéb': 'ETC', 'Tartozék': 'ACC'}
subCategories = {'PC': 'PCO', 'Laptop': 'LTP', 'Tablet': 'TAB', 'Egér': 'MOU', 'Billentyűzet': 'KEY',
                 'Monitor': 'MON', 'TV': 'TVI'}

# Adatok beolvasása
lastRow = sheet.range(f"N{str(sheet.cells.last_cell.row)}").end('up').row
barCodes = sheet.range(f"A5:A{lastRow}").value
descriptions = sheet.range(f"E5:E{lastRow}").value
purchDates = sheet.range(f"Q5:Q{lastRow}").value
logging.info(
    f"A vonalkód generáláshoz tartozó adatok beolvasása: utolsó sor-{lastRow}, vonalkód-{barCodes}, leírás-{descriptions}, beszerzési dátum-{purchDates}")

# Kétdimenziós listák átalakítása egy listává
barCodes = [code[0] if isinstance(code, list) else code for code in barCodes]
descriptions = [desc[0] if isinstance(desc, list) else desc for desc in descriptions]
purchDates = [date[0] if isinstance(date, list) else date for date in purchDates]
logging.info(
    f"A vonalkód generáláshoz tartozó adatok átalakítása: utolsó sor-{lastRow}, vonalkód-{barCodes}, leírás-{descriptions}, beszerzési dátum-{purchDates}")

# Minták a kategóriák beazonosításához
patterns = {
    'Munka állomás': ['PC', 'Laptop', 'Tablet', 'Egér', 'Billentyűzet'],
    'Kijelző': ['TV', 'Monitor'],
    'Hálózati eszköz': ['Router', 'Switch', 'Patch panel', 'Modem'],
    'Telefon': ['Mobil'],
    'Tartozék': ['Tok', 'Töltő', 'Táska', 'Hálókártya', 'Szünetmentes tápegység', 'Winchester', 'Dokkoló'],
    'Szerver': ['Szerver', 'Kamera szerver']
}

# Meglévő vonalkódok feldolgozása
existing_codes = {}  # Melyik kategóriának mi az utolsó sorszáma

for code in barCodes:
    if code and "-" in code:
        parts = code.split("-")
        if len(parts) == 4:  # Pl. 2025-WS-PC-003
            base_code = "-".join(parts[:3])  # 2025-WS-PC
            serial = int(parts[3])  # 003
            existing_codes[base_code] = max(existing_codes.get(base_code, 0), serial)

logging.info(f"Meglévő vonalkódok feldolgozása")

# Új vonalkódok generálása
updated_codes = []
for i in range(len(barCodes)):
    if barCodes[i]:  # Ha már van vonalkód, ne változtassuk meg
        updated_codes.append(barCodes[i])
        continue

    category_code = "ETC"
    subcategory_code = "000"

    for key, value in categories.items():
        for category, items in patterns.items():

            if descriptions[i] in items:
                category_code = categories[category]
                subcategory_code = subCategories.get(descriptions[i], "000")
            elif descriptions[i] == key:
                category_code = categories[key]
                subcategory_code = subCategories.get(descriptions[i], "000")

    year = datetime.strftime(purchDates[i], "%Y") if purchDates[i] else "0000"
    base_code = f"{year}-{category_code}-{subcategory_code}"

    # Új sorszám keresése
    existing_codes[base_code] = existing_codes.get(base_code, 0) + 1
    new_code = f"{base_code}-{existing_codes[base_code]:03}"  # Pl. 2025-WS-PC-004

    updated_codes.append(new_code)

logging.info(f"Új vonalkódok generálása")

# Vonalkódok visszaírása az Excelbe
sheet.range(f"A5:A{lastRow}").value = [[code] for code in updated_codes]
logging.info(f"Új vonalkódok generálása")

messagebox.showinfo(message="Vonalkódok frissítve.")

file.save()