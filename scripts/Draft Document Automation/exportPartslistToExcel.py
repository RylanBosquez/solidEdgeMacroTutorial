import os
import comtypes.client

from comtypes.automation import VARIANT


def exportPartlistToExcel():
    
    try:
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
    except:
        print("No running Solid Edge instance found.")
        return

    seApp.Visible = True
    seDoc = seApp.ActiveDocument
    
    print(f'Name: {seDoc.Name}\n')

    partsList = seDoc.PartsLists
    
    for partsListIndex in range(partsList.Count):
        
        partList = partsList.Item(partsListIndex+1)
        
        rows = partList.Rows
        columns = partList.Columns
        
        for rowIndex in range(rows.Count):
            
            for colIndex in range(columns.Count):
            
                try:
                    print(partList.Cell(VARIANT(rowIndex + 1), VARIANT(colIndex + 1)).value, end='\t')
                except:
                    pass
                
            print()


if __name__ == "__main__":
    
    os.system('cls')
    exportPartlistToExcel()