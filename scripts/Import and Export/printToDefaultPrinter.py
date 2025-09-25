import os
import comtypes.client
from comtypes.automation import VARIANT
from comtypes.gen._3E2B3BDC_F0B9_11D1_BDFD_080036B4D502_0_1_0 import DraftDocument


def printToDefaultLandscape(printerName, copies, orientation):
    
    # 'PrintOut',
    #     (['in', 'optional'], VARIANT, 'Printer'),
    #     (['in', 'optional'], VARIANT, 'NumCopies'),
    #     (['in', 'optional'], VARIANT, 'Orientation'),
    #     (['in', 'optional'], VARIANT, 'PaperSize'),
    #     (['in', 'optional'], VARIANT, 'Scale'),
    #     (['in', 'optional'], VARIANT, 'PrintToFile'),
    #     (['in', 'optional'], VARIANT, 'OutputFileName'),
    #     (['in', 'optional'], VARIANT, 'PrintRange'),
    #     (['in', 'optional'], VARIANT, 'Sheets'),
    #     (['in', 'optional'], VARIANT, 'ColorAsBlack'),
    #     (['in', 'optional'], VARIANT, 'Collate')
    
    try:
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
    except:
        print("Solid Edge not running.")
        return

    seApp.Visible = True
    activeDoc = seApp.ActiveDocument
    draftDoc = activeDoc.QueryInterface(DraftDocument)

    try:
        
        draftDoc.PrintOut(
            VARIANT(printerName),
            VARIANT(copies),
            VARIANT(orientation),
        )
        
    except Exception as e:
        print(f"PrintOut failed: {e}")


if __name__ == "__main__":
    
    printerName = "Microsoft Print to PDF" # <-- change this to your printer name
    copies = 1 # <-- change this to the number of copies
    orientation = 2 # 2 for Horizontal 1 for Vertical Layout
    
    os.system('cls')
    printToDefaultLandscape(printerName, copies, orientation)
