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
def analyze(sheet, L, AVG, z, K, h, STD, AVGL, STDL, r, R, Q, D):
    if isAnyData(sheet):
        try:
            # (Q, R) Policy
            AverageDemandDuringLeadTime = L * AVG 
            SafetyStock = z * STD * math.sqrt(L)
            ReorderLevel = (L * AVG) + (z * STD * math.sqrt(L))
            OrderQuantity = math.sqrt(((2 * K) * AVG) / h)
            InventoryLevelBeforeReceivingAnOrder = z * STD * math.sqrt(L)
            InventoryLevelAfterReceivingAnOrder = (Q + z) * STD * math.sqrt(L)
            AverageInventory = Q / (2 + (z * STD * math.sqrt(L)))
            ReorderPoint = (AVG * AVGL) + z * math.sqrt((AVGL * STD^2) + (AVG^2 * STDL^2))
            DemandDuringLeadTime = AVG * AVGL
            StandardDeviationOfDemandDuringLeadTime = math.sqrt((AVGL * STD^2) + (AVG^2 * STDL^2))
            AmountOfSafetyStock = z * math.sqrt((AVGL * STD^2) + (AVG^2 * STDL^2))
            # (s, S) Policy
            s = R 
            S = R + Q
            # Base-stock level Policy
            AverageDemandDuringAnIntervalOfRplusLDays = (r + L) * AVG
            SafetyStockBS = z * STD * math.sqrt(r + L)
            BasestockLevelS = (r + L) * AVG + (z * STD * math.sqrt(r + L))
            AverageInventoryBS = (r * D) / (2 + (z * STD * math.sqrt(r + L)))
            # Base-stock level Policy (Lead time = uncertain, Normally distributed with lead time of AVGL, and STDL)
            AverageDemandDuringAnIntervalOfRplusLDaysUncertain = (r + AVGL) * AVG
            StandardDeviationOfDemandDuringAnIntervalOfRplusLDaysUncertain = math.sqrt((r + AVGL) * STD^2 + (AVG^2 * STDL^2))
            SafetyStockBSUncertain = z * math.sqrt((r + AVGL) * STD^2 + (AVG^2 * STDL^2))
            BasestockLevelSUncertain = (r + AVGL) * AVG + SafetyStockBSUncertain
        except:
            print("Error")
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
