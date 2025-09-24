import os
import comtypes.client


def batchFileProcessor(folderPath):
    
    try:
        
        # Connect to running Solid Edge
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")

        # Iterate over all Solid Edge files in the folder
        for fileName in os.listdir(folderPath):
            
            if fileName.lower().endswith(('.par', '.asm', '.psm', '.dft')):
                
                fullPath = os.path.join(folderPath, fileName)
                print(f"Processing: {fullPath}")

                # Open the document
                seDoc = seApp.Documents.Open(fullPath)

                # Example: Print document name
                print(f"Opened: {seDoc.Name}")
                seDoc.Close()

    except Exception as e:
        print(f"Error: {e}")

    finally:
        seApp = None


if __name__ == "__main__":
    
    folderPath = os.path.abspath(os.path.join(os.getcwd(), "assets"))
    
    os.system('cls')
    batchFileProcessor(folderPath)
