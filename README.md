# deckbuilder
Card pictures generating and printing tool. Accepts Excel file as a source and produces a deck of cards for each Excel sheet inside. Good for card game playtests and card proxies.

**No ponies, no dragons**: the tool doesn't support card illustrations.
Life shows that text presence is much more important. Especially for tests.
# Generation Example
PDF output file: https://github.com/lagital/deckbuilder/blob/master/example/Example/print/Example.pdf

It has been constructed using this Excel workbook: https://github.com/lagital/deckbuilder/blob/master/example/Example.xlsx

And using these parameters: https://github.com/lagital/deckbuilder/blob/master/example/parms.py
# Installation
Refer to *install.txt* for the list of required Python packages.
# Usage
Source Excel file has to have **Title** and **Description** columns on each sheet.
Sheets not passing the check will be skipped by the builder.

Run *deckbuilder.py* with options:

**-s, --source:** Excel source file. **Mandatory**

**-o, --output:** Output folder. **Default:** execution directory.
 
**-f, --format:** Output document format. **Default:** pdf.
 
**-t, --tabletop:** Export for Tabletop Simulator - True/False. **Default:** False. **TBD**

**-p, --print:** Print cards on default printer - True/False. **Default:** False.
# Configuration (parms.py)
 - **Masks:**
```python
# card generation mask
def MASKS():                  return "Example.Card 1,Example.Card 2,Example2.*"
def MASK_SEPARATOR():         return ","
def MASK_DOT():               return "."
def MASK_ALL():               return "*"
```
Define what cards from what Excel sheets you would like to generate. E.g., saying "Men.Jon Snow,Men.Daenerys,Others.*" you denote that only cards with Titles "Jon Snow" and "Daenerys" will be printed from sheet "Men" + all cards from sheet "Others".
 - **Dimentions**
```python
# dimensions
def DIM_CARD_WIDTH():         return 200
def DIM_CARD_HEIGHT():        return 289
def DIM_TEXT_WIDTH():         return 100
def DIM_TEXT_HEIGHT():        return 10
def DIM_CHAR_WIDTH():         return 7
def DIM_TEXT_TOP_MARGIN():    return 20
def DIM_TEXT_LEFT_MARGIN():   return 20
def DIM_PDF_TOP_MARGIN():     return 5
def DIM_PDF_LEFT_MARGIN():    return 5
```
Define sizes for cards, text and margins.
 - **Printing**
```python
 # printing
PRINT                         = False
def CARDS_IN_ROW():           return 3
def CARDS_IN_COLUMN():        return 3
```
With PRINT = True the scrpit will always try to print generated data using default printer. You can also define the number of cards in row/column on one PDF page.
 - **Columns**
```python
# source data columns
def COLUMN_TITLE():           return "Title"
def COLUMN_DESCRIPTION():     return "Description"
def COLUMN_COUNT():           return "Count"
def COLUMN_IDT():             return "N"
```
By default the script searches for Title, Description and Count columns on each sheet. You can redefine column names using these parameters.
IDT (identifier) column is used in case of cards with the same titles but different descriptions. In IDT column you define some unique identifier for each card so they will be saved and generated correctly.
# Customization
It is possible to customize card Title and Description generation using **cust/cust_title.py** and **cust/cust_description.py** files respectively.
