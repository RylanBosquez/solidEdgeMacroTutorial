import os
import comtypes.client


def launchSolidEdge():
    
    try:
        
        # Try to connect to an already running Solid Edge instance
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
        seApp.Visible = True  # Ensure the application window is visible
        
        print("Connected to Solid Edge.")
        
    except:
        
        # If not running, launch a new instance of Solid Edge
        seApp = comtypes.client.CreateObject("SolidEdge.Application")
        seApp.Visible = True
        
        print("Launched Solid Edge.")

    finally:
        
        # Release COM object reference to avoid memory leaks
        seApp = None


if __name__ == "__main__":
    
    os.system('cls')  # Clear the console window
    launchSolidEdge()
