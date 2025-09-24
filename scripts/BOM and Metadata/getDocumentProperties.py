import os
import comtypes.client


def getAllDocumentProperties():
    
    try:
        
        # Connect to running Solid Edge
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
        seDoc = seApp.ActiveDocument

        # Access all property sets
        propertySets = seDoc.Properties

        # Iterate over each property set
        for i in range(1, propertySets.Count + 1):
            
            propertySet = propertySets.Item(i)
            print(f"\n--- Property Set: {propertySet.Name} ---")

            # Iterate over each property in the set
            for j in range(1, propertySet.Count + 1):
                
                prop = propertySet.Item(j)
                try:
                    
                    print(f"{prop.Name}: {prop.Value}")
                except Exception:
                    
                    continue

    except Exception as e:
        
        print(f"Error: {e}")

    finally:
        
        seDoc = None
        seApp = None
        

if __name__ == "__main__":
    
    os.system('cls')
    getAllDocumentProperties()
