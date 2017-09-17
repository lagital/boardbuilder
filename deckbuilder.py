import argparse
import parms
import pandas as pd
import textwrap
import os
import subprocess
import sys
from PIL import Image, ImageDraw, ImageFont
from fpdf import FPDF
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
    parser.add_argument('-p', "--print", type=bool, action='store', dest='print',
                        help='Print generated files')
    args = parser.parse_args()

    # redefine global parameters
    parms.FILE_SOURCE   = args.source
    parms.DIR_OUTPUT    = nvl(args.output, parms.DIR_OUTPUT)
    parms.FORMAT        = nvl(args.format, parms.FORMAT)
    parms.FLAG_TABLETOP = nvl(args.tabletop, parms.FLAG_TABLETOP)
    parms.PRINT         = nvl(args.print, parms.PRINT)

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

    deck = []

    if parms.COLUMN_TITLE() not in sheet.keys() or parms.COLUMN_DESCRIPTION() not in sheet.keys():
        print("WARNING:", parms.COLUMN_TITLE(), "and", parms.COLUMN_DESCRIPTION(),
              "columns must be defined on the sheet. Skipping.")
        return

    if parms.COLUMN_COUNT() not in sheet.keys():
        print("WARNING:", parms.COLUMN_COUNT(), "column not defined on sheet", sheet_title + ".",
              "Generating one copy for each card")
        sheet["Count"] = pd.Series(1, index=sheet.index)

    for index, row in sheet.iterrows():

        card_title = cust_title.do(row, sheet_title, row[parms.COLUMN_TITLE()])

        if card_included(sheet_title, card_title):
            card_description = cust_description.do(row, sheet_title, row[parms.COLUMN_DESCRIPTION()])
            card_image = generate_card_image(card_title, card_description)
            card_count = row[parms.COLUMN_COUNT()]

            card = Card(card_title, card_description, card_image, card_count)
            deck.append(card)

            print(card_count, '"' + card_title + '" cards have been generated.')

    save_sheet(sheet_title, deck)


def generate_card_image(title, description):
    # scheme, size, background color
    img = Image.new('RGB', (parms.DIM_CARD_WIDTH(), parms.DIM_CARD_HEIGHT()), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    # draw title
    unicode_font = ImageFont.truetype("Arial.ttf")
    y_text = draw_lines(draw, unicode_font, title, parms.DIM_TEXT_TOP_MARGIN())

    # space between title and description
    y_text += parms.DIM_TEXT_TOP_MARGIN()

    # draw description
    for p in str.split(description, "\p"):
        for n in str.split(p, "\\n"):
            y_text = draw_lines(draw, unicode_font, n, y_text)

        y_text += parms.DIM_TEXT_TOP_MARGIN()

    # border
    img = apply_card_border(img)

    return img


def draw_lines(draw, font, text, y_text):
    lines = textwrap.wrap(text, width=(parms.DIM_CARD_WIDTH() // parms.DIM_CHAR_WIDTH()))
    for line in lines:
        draw.text((parms.DIM_TEXT_LEFT_MARGIN(), y_text), line, fill=(0, 0, 0), font=font)
        y_text += parms.DIM_TEXT_HEIGHT()
    return y_text


def apply_card_border(img):
    new_size = (img.size[0] + parms.DIM_CARD_BORDER() * 2, img.size[1] + parms.DIM_CARD_BORDER() * 2)
    bordered_img = Image.new("RGB", new_size)
    bordered_img.paste(img, (parms.DIM_CARD_BORDER(), parms.DIM_CARD_BORDER()))

    return bordered_img


def save_sheet(sheet_title, deck):
    main_directory = generate_sheet_directories(sheet_title)

    pdf = None

    if parms.FORMAT == parms.FORMAT_PDF():
        pdf = FPDF()

    card_paths = []
    card_total_count = 0
    for c in deck:
        card_total_count += c.count
    card_counter = 0

    for i, card in enumerate(deck):
        for j in range(card.count):

            # separate images
            card_path = main_directory + "/" + card.title.replace(" ", "_") + "_" + str(j) + "." + parms.EXT_PNG()
            card_paths.append(card_path)
            card.image.save(card_path, parms.EXT_PNG())

            card_counter += 1

            # combine in one page
            if (card_total_count - card_counter) % (parms.CARDS_IN_ROW() * parms.CARDS_IN_COLUMN()) == 0:
                print("Page added", card_total_count - card_counter)
                sheet_page_image = Image.new('RGB',
                                             (parms.CARDS_IN_ROW() * (parms.DIM_CARD_WIDTH() + parms.DIM_CARD_BORDER() * 2),
                                              parms.CARDS_IN_COLUMN() * (parms.DIM_CARD_HEIGHT() + parms.DIM_CARD_BORDER() * 2 )),
                                             (255,255,255,0))
                x_offset = 0
                for k, img in enumerate(map(Image.open, card_paths)):
                    sheet_page_image.paste(img, ((k % parms.CARDS_IN_ROW()) * img.size[0],
                                                 (k // parms.CARDS_IN_COLUMN()) * img.size[1]))
                    x_offset += img.size[0]

                sheet_page_image_path = main_directory + "/" + parms.DIR_PAGES() + "/"\
                                        + parms.FILE_PAGE() + str(card_total_count - card_counter)\
                                        + "." +  parms.EXT_PNG()
                sheet_page_image.save(sheet_page_image_path)

                # pdf
                if parms.FORMAT == parms.FORMAT_PDF():
                    pdf.add_page()
                    pdf.image(sheet_page_image_path, x=parms.DIM_PDF_LEFT_MARGIN(), y=parms.DIM_PDF_TOP_MARGIN())

                card_paths = []

    printing_file = None

    if parms.FORMAT == parms.FORMAT_PDF():
        printing_file = main_directory + "/" + parms.DIR_PRINT() + "/" + sheet_title.replace(" ", "_")\
                        + "." + parms.FORMAT_PDF()
        pdf.output(printing_file, "F")

    if parms.PRINT is True:
        print_sheet(printing_file)

    print('"' + sheet_title + '"', "finished.")


def generate_sheet_directories(sheet_title):
    main_directory = parms.DIR_OUTPUT + "/" + sheet_title
    if not os.path.exists(main_directory):
        os.makedirs(main_directory)
    for d in [parms.DIR_PAGES(), parms.DIR_PRINT(), parms.DIR_TABLETOP()]:
        directory = main_directory + "/" + d
        if not os.path.exists(directory):
            os.makedirs(directory)

    return main_directory


def card_included(sheet_title, card_title):
    global MASK_DICT

    if sheet_title not in MASK_DICT.keys():
        return False
    elif parms.MASK_ALL() in MASK_DICT[sheet_title] or card_title in MASK_DICT[sheet_title]:
        return True
    else:
        return False


def print_sheet(sheet_path):
    if sheet_path is not None:
        print("Printing ...")
        if sys.platform == "win32":
            os.startfile(sheet_path, "print")
        else:
            lpr = subprocess.Popen("/usr/bin/lpr", stdin=subprocess.PIPE)
            lpr.stdin.write(open(sheet_path, "rb").read())


def nvl(var, val):
  if var is None:
    return val
  return var


if __name__ == "__main__":
    build()
