from tkinter import filedialog
import openpyxl as pxl

import math



def importExcel():
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
    

def exportReport():
    pass

def processInput():
    pass

def processOuput():
    pass
