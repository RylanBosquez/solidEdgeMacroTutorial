import os
import openpyxl
import comtypes.client

from comtypes.automation import VARIANT


def exportPartlistToExcel():
    
    savePath = './assets/partlist.xlsx' # <------------ Change this to your save path
    
    try:
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
    except:
        print("No running Solid Edge instance found.")
        return
    
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    
    headerRow = ['Item', 'Part Number', 'Qty', 'Description']
    worksheet.append(headerRow)

    seApp.Visible = True
    seDoc = seApp.ActiveDocument
    
    partsList = seDoc.PartsLists
    
    partList = partsList.Item(1) # Grab the first parts list if only part list, else loop over
    
    for rowIndex in range(partList.Rows.Count):
        
        for colIndex in range(partList.Columns.Count):
        
            try:
                worksheet.cell(row=rowIndex+2, column=colIndex+1).value = partList.Cell(VARIANT(rowIndex + 1), VARIANT(colIndex + 1)).value
            except:
                pass
        
    workbook.save(savePath)


if __name__ == "__main__":
    
    os.system('cls')
    exportPartlistToExcel()