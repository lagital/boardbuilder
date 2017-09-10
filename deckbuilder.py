import argparse
import parms
import pandas as pd
import textwrap
import os
from PIL import Image, ImageDraw
from pathlib import Path
from Card import Card
from cust import cust_title
from cust import cust_description

FILE_EXT = parms.EXT_XLSX()
SHEETS = []
EXCEL = None

def build():
    global FILE_EXT
    global SHEETS
    global EXCEL

    # parse args
    parser = argparse.ArgumentParser(description='Building decks')
    parser.add_argument('-s', "--source", type=str, action='store', dest='source', help='Excel or CSV source')
    parser.add_argument('-o', "--output", type=str, action='store', dest='output', help='Output folder')
    parser.add_argument('-f', "--format", type=str, action='store', dest='format', help='Only PDF for now')
    parser.add_argument('-t', "--tabletop", type=bool, action='store', dest='tabletop',
                        help='Export for Tabletop Simulator')
    args = parser.parse_args()

    # redefine global parameters
    parms.FILE_SOURCE = args.source
    parms.DIR_OUTPUT = args.output
    parms.FORMAT = args.format
    parms.FLAG_TABLETOP = args.tabletop

    print("[Validating parameters]")
    if not validate_parameters():
        return

    print("[Processing sheets]")
    process_sheets()


def validate_parameters():
    if parms.FILE_SOURCE is None:
        print("Source file path is invalid")
        return False

    if not Path(parms.FILE_SOURCE).is_file():
        print("Source file path is invalid")
        return False

    filename, ext = parms.FILE_SOURCE.split(".")
    if ext.lower() not in (parms.EXT_XLS(), parms.EXT_XLSX(), parms.EXT_CSV()):
        print("Source file type is not supported")
        return False
    else:
        global FILE_EXT
        FILE_EXT = ext

    return True


def process_sheets():
    global SHEETS
    global EXCEL

    # excel
    if FILE_EXT in (parms.EXT_XLS(), parms.EXT_XLSX()):
        EXCEL = pd.ExcelFile(parms.FILE_SOURCE)
        for sn in EXCEL.sheet_names:
            sheet = EXCEL.parse(sn)
            SHEETS.append(sheet)
            process_sheet(sheet, sn)


def process_sheet(sheet, sheet_title):
    print("Processing", '"' + sheet_title + '"', "...")

    deck = []

    if parms.COLUMN_TITLE() not in sheet.keys() or parms.COLUMN_DESCRIPTION() not in sheet.keys():
        print("SKIPPED: Title and Description columns must be defined on the sheet.")
        return

    for index, row in sheet.iterrows():
        print(row)
        card_title = cust_title.do(row, sheet_title, row[parms.COLUMN_TITLE()])
        card_description = cust_description.do(row, sheet_title, row[parms.COLUMN_DESCRIPTION()])

        print("Generating card", '"' + card_title + '"')
        card = Card(card_title, card_description, generate_card_image(card_title, card_description))
        deck.append(card)

    save_sheet(sheet_title, deck)


def generate_card_image(title, description):
    # scheme, size, background color
    img = Image.new('RGB', (parms.DIM_CARD_WIDTH(), parms.DIM_CARD_HEIGHT()), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    # draw title
    draw.text((parms.DIM_TEXT_LEFT_MARGIN(),parms.DIM_TEXT_TOP_MARGIN()), title, fill=(0, 0, 0))

    # draw description
    lines = textwrap.wrap(description, width=parms.DIM_TEXT_WIDTH())
    y_text = parms.DIM_TEXT_TOP_MARGIN() + parms.DIM_TEXT_HEIGHT()
    for line in lines:
        draw.text((parms.DIM_TEXT_LEFT_MARGIN(), y_text), line, fill=(0, 0, 0))
        y_text += parms.DIM_TEXT_HEIGHT()

    return img


def save_sheet(title, deck):
    directory = parms.DIR_OUTPUT + "/" + title
    if not os.path.exists(directory):
        os.makedirs(directory)

    for card in deck:
        card.image.save(directory + "/" + card.title.replace(" ", "_") + "." + parms.EXT_PNG(), parms.EXT_PNG())

    print('"' + title + '"', "saved.")


if __name__ == "__main__":
    build()
