import xlwings as xw
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from collections import defaultdict
import logging
from tkinter import *
from tkinter import messagebox

# Excel fájl beolvasása
file = xw.Book("D:\\Leltar\\IT_leltar.xlsm")
sheet = file.sheets["Leltár"]

# Naplózás
log_file = open(r"D:\Leltar\log.txt")
logging.basicConfig(filename=r"D:\Leltar\log.txt", level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s', filemode='a')

#Alert üzenet
msgbox = Tk()

# Szükséges adatok beolvasása
lastRow = sheet.range(f"N{str(sheet.cells.last_cell.row)}").end('up').row
sheetTable = sheet.range(f"A4:Q{lastRow}").value
sheetTable = [row for row in sheetTable if any(cell is not None for cell in row)]
pdfmetrics.registerFont(TTFont("Arial", "Fonts/Arial.ttf"))
pdfmetrics.registerFont(TTFont("Arial-Bold", "Fonts/Arial_Bold.ttf"))
logging.info(f"A PDF generáláshoz tartozó adatok beolvasása: utolsó sor-{lastRow}, táblázat-{sheetTable}")

styles = getSampleStyleSheet()
normal_style = styles["Normal"]


# Adatok csoportosítása helyek szerint
def group_by_location(sheetTable):
    grouped_data = defaultdict(list)

    for index, row in enumerate(sheetTable):
        if index == 0:
            continue
        elif row[12]:  # Ha nem None
            location = row[12]
            grouped_data[location].append(row)
            print(grouped_data)

    logging.info(f"Adatok csoportosítása: {grouped_data}")
    return grouped_data

# Helyiségnév táblázat létrehozása
def localization(location):
    bold_style = styles["Title"]
    bold_style.fontName = "Arial-Bold"

    table_data = [[Paragraph(location, bold_style)]]
    style = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 0), (-1, -1), colors.lightgrey),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10)
    ])
    table = Table(table_data, colWidths=[500])
    table.setStyle(style)

    logging.info(f"Lokalizálás: {table}")
    return table


# Adattábla létrehozása
def create_table(rows):
    formatted_rows = []
    col_names = ['Vonalkód', 'Gyártó', 'Modell', 'Eszköz leírás', 'Db', 'Beszerzés\ndátuma']
    formatted_rows.insert(0, col_names)

    # Felesleges oszlopok törlése

    wanted_columns = [0, 2, 3, 4, 13, 16]  # Az oszlopok indexei
    filtered_rows = [[row[i] for i in wanted_columns] for row in rows]

    for selected_row in filtered_rows:

        # Cellák formázása és sortörés beállítása
        for i in range(len(selected_row)):
            if selected_row[i] is not None:
                if isinstance(selected_row[i], float):
                    selected_row[i] = int(selected_row[i])
                if isinstance(selected_row[i], datetime):
                    selected_row[i] = datetime.strftime(selected_row[i], '%Y.%m.%d.')
                elif isinstance(selected_row[i], str) and len(selected_row[i]) > 15:
                    selected_row[i] = Paragraph(selected_row[i], normal_style)

        formatted_rows.append(selected_row)

    col_width = [110, 75, 120, 80, 40, 70]

    style = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Arial-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('FONTNAME', (0, 1), (-1, -1), 'Arial'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ])

    table = Table(formatted_rows, colWidths=col_width, repeatRows=2)
    table.setStyle(style)
    return table


# PDF generálása

def generate_pdf(sheetTable):
    grouped_data = group_by_location(sheetTable)

    fileName = f'Tárgyi eszközleltár {datetime.today().strftime("%Y-%m-%d")}.pdf'

    pdf = SimpleDocTemplate(fileName, pagesize=A4)
    pdf.title = "Tárgyi eszközleltár"
    title_style = styles["Title"]
    title = Paragraph("Tárgyi eszközleltár", title_style)
    subTitle = Paragraph(f"{datetime.today().strftime('%Y-%m-%d')}", title_style)

    elements = []
    elements.insert(0, title)
    elements.insert(1, subTitle )
    for location, rows in grouped_data.items():
        elements.append(localization(location))  # Helyiség neve
        elements.append(Spacer(1, 10))  # Kis térköz
        elements.append(create_table(rows))  # Tábla
        elements.append(Spacer(1, 20))  # Nagyobb térköz az új helyiség előtt

    pdf.build(elements)


# PDF létrehozásának meghívása
generate_pdf(sheetTable)
messagebox.showinfo(message="A PDF létre jött.")
