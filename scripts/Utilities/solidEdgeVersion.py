import os
import comtypes.client


def getSolidEdgeVersion():
    
    try:
        
        # Connect to running Solid Edge
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
        seVersion = seApp.Version
        
        # Print Version
        print(f"Solid Edge Version: {seVersion}")
        
    except Exception as e:
        
        print(f"Error: {e}")

    finally:
        
        seApp = None


if __name__ == "__main__":
    
    os.system('cls')
    getSolidEdgeVersion()