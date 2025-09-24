import os
import comtypes.client


def getSessionDocumentList():
    
    try:
        
        # Connect to running Solid Edge
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
        seDocs = seApp.Documents

        # Iterate over each document in the Documents collection
        for i in range(1, seDocs.Count + 1):
            
            sessionDoc = seDocs.Item(i)
            print(f"Session Document: {sessionDoc.Name}")

    except Exception as e:
        print(f"Error: {e}")

    finally:
        
        seDocs = None
        seApp = None


if __name__ == "__main__":
    
    os.system('cls')
    getSessionDocumentList()
