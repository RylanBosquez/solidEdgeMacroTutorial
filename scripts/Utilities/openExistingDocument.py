import os
import comtypes.client
from tkinter import filedialog

def openExistingDocument(documentPath):
    
    try:
        
        # Connect to running Solid Edge instance
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
        seApp.Visible = True

        # Access the Documents collection
        seDocuments = seApp.Documents

        # Validate the file path
        if not os.path.exists(documentPath):
            print(f"File not found: {documentPath}")
            return
        # Open the document
        sePartDoc = seDocuments.Open(documentPath)

        # Activate the document to bring it to the foreground
        sePartDoc.Activate()

        print(f"Opened and activated: {documentPath}")

    except Exception as e:
        
        print(f"Error: {e}")

    finally:
        
        # Clean up COM references
        sePartDoc = None
        seDocuments = None
        seApp = None
        

if __name__ == "__main__":
    
    # Get absolute path to the file
    assetsFolder = os.path.join(os.getcwd(), "assets") 
    os.makedirs(assetsFolder, exist_ok=True)
    
    documentPath = os.path.join(assetsFolder, "newPart.par")

    if documentPath:
        openExistingDocument(documentPath)
    else:
        print("No file selected.")
