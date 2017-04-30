''' Ultimate Card Generater by Kurtis Raymond (kraymond (at) sfu.ca)
    Excel .xlsx file + Template .svg -> Generated Cards

    This script requires the each data ID matches
    is the column is labeled .xlsx, then the variable
    should be the same in the template file.
    ie,
    1 | VAR_NAME
    2 | data_val1
    3 | data_val2
    . |     .
    . |     .

    And VAR_NAME should be the same text field in the card template
    The actual chosen name for VAR_NAME should be unique and
    not match any of the other xml code in the template file.

    Future work:
    - Extend this to work with different card types, ie. VAR_TEMP in xlsx file
    - Create a more robust way of changing data in a .svg file
'''


# Set the filename of the .xlsx file SET THIS!
SpreadSheetFile = 'Card_Data.xlsx'

# Our Template File SET THIS!
template = 'Printerstudio_Template.svg'

# Import needed modules
import os
from shutil import copyfile
import re

# Install Dependencies
try:
    import pandas as pd
    import xlrd
except:
    print('Pandas not Install, Installing.', end='')
    import pip
    print('.', end='')
    pip.main(['install', 'pandas'])
    pip.main(['install', 'xlrd'])
    print('. Installed')
    import pandas as pd
    del pip

# === PROGRAM HELPER FUNCTIONS ===
def fileNamify(string):
    ''' Removes all Spaces and whitespace characters with dashes '''
    string = re.sub(r"[^\w\s]", '', string)
    string = re.sub(r"\s+", '-', string)
    return string

def mkdir(dir):
    if not os.path.exists(dir):
        os.makedirs(dir)
mkdir('bin')

# Import data from excel using pandas
x1 = pd.ExcelFile(SpreadSheetFile)
dfCards = x1.parse('Sheet1')
print('Sheet 1 Parsed')

templateNames = list(dfCards)
dataNew = []
print("Data Type Names from .xlsx", end='')
print(templateNames)

# Export all the cards
for index, row in dfCards.iterrows():
    fName = 'bin/' + fileNamify(row[templateNames[0]]) + '.svg'
    for dataParam in templateNames:
        dataNew.append(row[dataParam])

    # Copy Our Template File to our new location.
    copyfile(template, fName)

    # Read Data
    with open(fName, 'r') as card:
        Card_Data = card.read()

    # Change our Fields
    for dataParam, fieldData in zip(templateNames, dataNew):
        Card_Data = Card_Data.replace(dataParam, fieldData)

    # Write Cards
    with open(fName, 'w') as card:
        card.write(Card_Data)
