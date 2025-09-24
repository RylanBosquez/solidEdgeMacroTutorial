import os
import comtypes.client


def listPartFeatures():
    
    try:
        
        # Connect to running Solid Edge
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
        seDoc = seApp.ActiveDocument

        # Ensure it's a Part document
        if seDoc.Type != 1:  # 1 = igPartDocument
            print("⚠️ Active document is not a Part.")
            return

        # Access the model
        model = seDoc.Models.Item(1)

        # Access the features collection
        features = model.Features

        print(f"🔍 Found {features.Count} features in part:")
        for i in range(1, features.Count + 1):
            
            feature = features.Item(i)
            print(f" - {feature.Name} ({feature.Type})")

    except Exception as e:
        print(f"Error: {e}")

    finally:
        
        seDoc = None
        seApp = None


if __name__ == "__main__":
    
    os.system('cls')
    listPartFeatures()
