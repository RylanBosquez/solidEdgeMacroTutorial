import os
import comtypes.client
import comtypes.gen.SolidEdgeFramework as sef


def getPartListFromDraft():
    
    try:
        
        # Connect to running Solid Edge instance
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
        seDoc = seApp.ActiveDocument

        # Get document type and extension
        docType = seDoc.Type
        fileExt = os.path.splitext(seDoc.Name)[1].lower()

        # Override type if extension is .dft but type is incorrect
        if fileExt == ".dft" and docType != sef.igDraftDocument:
            docType = sef.igDraftDocument

        # Abort if document is not a Draft
        if docType != sef.igDraftDocument:
            print("❌ Active document is not a Draft (.dft).")
            return

        # Access PartsLists collection
        partsLists = seDoc.PartsLists

        # Iterate over each PartsList in the document
        for i in range(1, partsLists.Count + 1):
            
            partsList = partsLists.Item(i)
            rowCount = partsList.Rows.Count
            colCount = partsList.Columns.Count

            # Iterate over each row
            for rowIndex in range(1, rowCount + 1):

                # Iterate over each column in the row
                for colIndex in range(1, colCount + 1):
                    
                    try:
                        
                        # Access cell using 2D indexer
                        cell = partsList.Cell[rowIndex, colIndex]
                        cellValue = str(cell.Value).upper()

                        # Print cell value inline
                        print(cellValue, end=" ")

                    except Exception:
                        
                        # Skip cell if access fails
                        continue

                # Newline after each row
                print()

    except Exception as e:
        
        # Print any unexpected error
        print(f"Error: {e}")

    finally:
        
        # Release COM objects
        seDoc = None
        seApp = None


# Entry point
if __name__ == "__main__":
    
    os.system('cls')  # Clear console
    getPartListFromDraft()
