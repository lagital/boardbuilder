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

FILE_EXT  = parms.EXT_XLSX()
SHEETS    = []
EXCEL     = None
MASK_DICT = {}

def build():
    global FILE_EXT
    global SHEETS
    global EXCEL

    # parse args
    parser = argparse.ArgumentParser(description='Building decks')
    parser.add_argument('-s', "--source", type=str, action='store', dest='source', help='Excel source')
    parser.add_argument('-o', "--output", type=str, action='store', dest='output', help='Output folder')
    parser.add_argument('-f', "--format", type=str, action='store', dest='format', help='Only PDF for now')
    parser.add_argument('-t', "--tabletop", type=bool, action='store', dest='tabletop',
                        help='Export for Tabletop Simulator')
    args = parser.parse_args()

    # redefine global parameters
    parms.FILE_SOURCE   = args.source
    parms.DIR_OUTPUT    = nvl(args.output, parms.DIR_OUTPUT)
    parms.FORMAT        = nvl(args.format, parms.FORMAT)
    parms.FLAG_TABLETOP = nvl(args.tabletop, parms.FLAG_TABLETOP)

    print("[Validating parameters]")
    if not valid_parameters():
        return

    print("[Validating masks]")
    if not valid_masks():
        return

    print("[Processing sheets]")
    process_sheets()


def valid_parameters():
    if parms.FILE_SOURCE is None:
        print("ERROR: Source file path is invalid")
        return False

    if not Path(parms.FILE_SOURCE).is_file():
        print("ERROR: Source file path is invalid")
        return False

    filename, ext = parms.FILE_SOURCE.split(".")
    if ext.lower() not in (parms.EXT_XLS(), parms.EXT_XLSX(), parms.EXT_CSV()):
        print("ERROR: Source file type is not supported")
        return False
    else:
        global FILE_EXT
        FILE_EXT = ext

    if parms.FORMAT not in [parms.FORMAT_PDF()]:
        print(parms.FORMAT, parms.FORMAT_PDF())
        print("ERROR: Export format not supported")
        return False

    return True

def valid_masks():
    global MASK_DICT
    for m in parms.MASKS().split(parms.MASK_SEPARATOR()):
        if m.count(".", 1, len(m) - 1) != 1:
            print(m.count(".", 1, len(m) - 1))
            print("ERROR: Mask", '"' + m + '"', "is invalid")
            return False
        else:
            sheet_title, value = m.split(parms.MASK_DOT())
            if sheet_title not in MASK_DICT.keys():
                MASK_DICT[sheet_title] = []
            MASK_DICT[sheet_title].append(value)

    print("Masks:", MASK_DICT)
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

    CommonCount = None
    deck = []

    if parms.COLUMN_TITLE() not in sheet.keys() or parms.COLUMN_DESCRIPTION() not in sheet.keys():
        print("WARNING:", parms.COLUMN_TITLE(), "and", parms.COLUMN_DESCRIPTION(),
              "columns must be defined on the sheet. Skipping.")
        return

    if parms.COLUMN_COUNT() not in sheet.keys():
        print("WARNING:", parms.COLUMN_COUNT(), "column not defined on sheet", sheet_title + ".",
              "Generating one copy for each card")
        CommonCount = 1

    for index, row in sheet.iterrows():

        card_title = cust_title.do(row, sheet_title, row[parms.COLUMN_TITLE()])

        if card_included(sheet_title, card_title):
            card_description = cust_description.do(row, sheet_title, row[parms.COLUMN_DESCRIPTION()])
            card_image = generate_card_image(card_title, card_description)
            card_count = nvl(CommonCount, row[parms.COLUMN_COUNT()])

            card = Card(card_title, card_description, card_image, card_count)
            deck.append(card)

            print(card_count, '"' + card_title + '" cards have been generated.')

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


def save_sheet(sheet_title, deck):
    directory = parms.DIR_OUTPUT + "/" + sheet_title
    if not os.path.exists(directory):
        os.makedirs(directory)

    for card in deck:
        for i in range(card.count):
            card.image.save(directory + "/"
                            + card.title.replace(" ", "_") + "_" + str(i)
                            + "." + parms.EXT_PNG(), parms.EXT_PNG())

    print('"' + sheet_title + '"', "finished.")


def card_included(sheet_title, card_title):
    global MASK_DICT

    if sheet_title not in MASK_DICT.keys():
        return False
    elif parms.MASK_ALL() in MASK_DICT[sheet_title] or card_title in MASK_DICT[sheet_title]:
        return True
    else:
        return False


def nvl(var, val):
  if var is None:
    return val
  return var


if __name__ == "__main__":
    build()
