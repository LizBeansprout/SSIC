from tkinter import filedialog
import openpyxl as pxl
import math


def importExcel():
    try:
        # Load Excel (not include formula, Only read as actual data)
        file = filedialog.askopenfilename()
        excel = pxl.load_workbook(file, data_only=True)

        # Selet Sheet (This select active one, latest active)
        data = excel.active
        #data = excel['Sheet1']

        # Iterate over table to get value
        for row in range(0, data.max_row):
            for col in data.iter_cols(1, data.max_column):
                print(col[row].value)
    except:
        print("Err")
    finally:
        pass
    
def isAnyData(sheet):
    data = sheet.get_sheet_data()
    for row in data:
        for cell in row:
            if cell != '':
                print("Data in Table!")
                return True
    print("None in Table!")
    return False

#Ozone, Tawan
def analyze(sheet):
    if isAnyData(sheet):
        try:
            pass
        except:
            print("Err")
    else:
        print("Import your data or filled the table")
#Sun
def drawGraph(optiaml_input):
    pass

def exportReport():
    pass

def processInput():
    pass

def processOuput():
    pass
