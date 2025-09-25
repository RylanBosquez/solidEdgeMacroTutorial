import os
import comtypes.client


def createSolidEdgePart(partName, savePath):
    
    try:
        
        # Connect to running Solid Edge instance
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")

        # Ensure Solid Edge is visible
        seApp.Visible = True

        # Access the Documents collection
        seDocuments = seApp.Documents

        # Create a new Part document
        sePartDoc = seDocuments.Add("SolidEdge.PartDocument")

        # Activate the new document to bring it to the foreground
        sePartDoc.Activate()
        
        
        # Optional: Save the new part to a specific location
        
        assetsFolder = os.path.join(os.getcwd(), "assets") 
        os.makedirs(assetsFolder, exist_ok=True)
        
        sePartDoc.SaveAs(savePath)
        print(f"New Part Document created and saved to: {savePath}")

    except Exception as e:
        print(f"Error: {e}")

    finally:
        
        # Clean up COM references to avoid locking Solid Edge
        sePartDoc = None
        seDocuments = None
        seApp = None


if __name__ == "__main__":

    # Run the macro 
    partName = "examplePart.par"
    savePath = os.path.abspath(os.path.join(os.getcwd(), "assets/partFolder"))
    createSolidEdgePart(partName, savePath)
