import os


# card generation mask
def MASKS():                  return "Example.Card 1,Example.Card 2,Example2.*"
def MASK_SEPARATOR():         return ","
def MASK_DOT():               return "."
def MASK_ALL():               return "*"


# source data columns
def COLUMN_TITLE():           return "Title"
def COLUMN_DESCRIPTION():     return "Description"
def COLUMN_COUNT():           return "Count"

# directories
DIR_OUTPUT                    = os.getcwd()

def DIR_PRINT():              return "print"
def DIR_TABLETOP():           return "tabletop"

# files
FILE_SOURCE                   = "dummy"

def FILE_TABLETOP_TEMPLATE(): return "template"
def FILE_TABLETOP_DECK():     return "deck_"


# formats
def FORMAT_PDF():             return "pdf"

FORMAT                        = "pdf"

# flags
FLAG_TABLETOP                 = True

# file extensions
def EXT_XLS():                return "xls"
def EXT_XLSX():               return"xlsx"
def EXT_CSV():                return "csv"
def EXT_PNG():                return "png"

# dimensions
def DIM_CARD_WIDTH():         return 222
def DIM_CARD_HEIGHT():        return 319
def DIM_TEXT_WIDTH():         return 100
def DIM_TEXT_HEIGHT():        return 40
def DIM_TEXT_LEFT_MARGIN():   return 20
def DIM_TEXT_TOP_MARGIN():    return 20
