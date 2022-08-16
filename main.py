import os
from time import sleep
from pywhatkit.core import core
import pywhatkit
import webbrowser as web
import pyautogui as pg
from openpyxl import load_workbook
from urllib.parse import quote


def whatsapp_automation(path):
    # * Validate if data file exists
    if(os.path.isfile(path) == False):
        print("No se encontró el archivo data.xlsx, por favor provisionarlo en la carpeta data")
        return

    # * Validate if the communication needs an image and validate it
    img_name = input(
        "Hola! Por favor tipea el nombre y extensión de la imagen (ejm: imagen.png)\nEn caso esta comunicación no tenga imagen no escribas nada y presiona ENTER.\n> ")
    img_path = f"data/{img_name}"
    if(img_name != "" and os.path.isfile(img_path) == False):
        print("\nLa imagen ingresada es errónea, por favor intenta de nuevo.")
        return

    # * Read workbook
    wb = load_workbook(filename=path)
    sheet = wb.active
    for idx, row in enumerate(sheet.rows):
        # * Ignore headers row
        if idx == 0:
            continue

        phone_no = f"+51{row[0].value}"
        message = row[1].value

        if(img_name == ""):
            web.open(
                f"https://web.whatsapp.com/send?phone={phone_no}&text={quote(message)}")
            sleep(10)
            width, height = pg.size()
            pg.click(width / 2, height / 2)
            sleep(3)
            pg.press("enter")
            core.close_tab(wait_time=5)
        else:
            pywhatkit.sendwhats_image(
                phone_no, img_path, message, tab_close=True, close_time=7)


if __name__ == "__main__":
    whatsapp_automation("data/data.xlsx")
