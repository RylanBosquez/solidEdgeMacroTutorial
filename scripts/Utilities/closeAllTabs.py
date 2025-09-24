import os
import comtypes.client


def closeAllTabs():
    
    # Connect to running Solid Edge instance
    seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
    seApp.Visible = True
    
    # Access the Documents collection
    seDocuments = seApp.Documents
    
    # Iterate through all documents
    for seDoc in seDocuments:
        
        # Close the document
        seDoc.Close()
        
    print("All documents closed.")

if __name__ == "__main__":
    
    os.system('cls')
    closeAllTabs()