import xlwings as xw
from datetime import datetime
from tkinter import *
from tkinter import messagebox
import logging

# Excel fájl beolvasása
file = xw.Book("D:\\Leltar\\IT_leltar.xlsm")

# Munkalapok
inventory_sheet = file.sheets["Leltár"]
decommission_sheet = file.sheets["Selejtezési jegyzőkönyv"]

# Kijelölt sor
selected_row = xw.apps.active.selection
print(selected_row)
msgbox = Tk()

# Naplózás
log_file = open(r"D:\Leltar\log.txt")
logging.basicConfig(filename=r"D:\Leltar\log.txt", level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s', filemode='a')

logging.info(f"A selejtezendő eszköz sora: {selected_row}")

def transport_data(selection, worksheet):
    # Üres sor keresése a Selejtezési jegyzőkönyvben
    empty_row = worksheet.range("A" + str(worksheet.cells.last_cell.row)).end("up").row + 1
    decommission_date = datetime.now()

    logging.info(f"A selejtezendő eszköz sor értékei: {selection.value}")
    wanted_columns = [0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13]  # Az oszlopok indexei
    filtered_cells = [selection.value[i] for i in wanted_columns] # A kiválasztott sor elemeinek megszűrése
    filtered_cells.append(decommission_date.strftime("%Y.%m.%d")) # Dátum hozzá fűzése
    logging.info(f"A selejtezendő eszköz megszűrt sora: {filtered_cells}")
    worksheet.range(f"A{empty_row}").value = filtered_cells # Beillesztés a jegyzőkönyvbe

try:
    if messagebox.askyesno(message=f"Tényleg törölni szeretné ezt az eszközt?\nsor: {selected_row.row}"): # Alert
        transport_data(selected_row, decommission_sheet)
        selected_row.delete() # Sor törlése
        logging.info(f"A selejtezés végbe ment")
except:
    messagebox.showinfo(message="A selejtezés nem ment végbe.") # Ha hiba keletkezik be
    logging.warning(f"A selejtezés során hiba történt")
