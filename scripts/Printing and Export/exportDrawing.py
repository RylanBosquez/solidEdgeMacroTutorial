import os
import comtypes.client
import comtypes.gen.SolidEdgeFramework as sef


# Export the active Solid Edge document to PDF
def exportToPdf():
    
    try:
        
        # Get running instance of Solid Edge
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
        seDoc = seApp.ActiveDocument
        
        baseName = os.path.splitext(os.path.basename(seDoc.FullName))[0]
        pdfPath = os.path.join(os.getcwd(), "assets", baseName + ".pdf")
        
        # Export to PDF
        seDoc.SaveAsPDF(pdfPath)
        print(f"✅ Exported to PDF: {pdfPath}")
        
    except Exception as e:
        print(f"Error: {e}")
        
    finally:
        
        # Release COM objects
        seDoc = None
        seApp = None


# Export the active Solid Edge Draft document to DXF
def exportToDxf():
    
    try:
        
        # Get running instance of Solid Edge
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
        seDoc = seApp.ActiveDocument
        
        baseName = os.path.splitext(os.path.basename(seDoc.FullName))[0]
        dxfPath = os.path.join(os.getcwd(), "assets", baseName + ".dxf")
        
        # Only export if document is a Draft
        if seDoc.Type == sef.igDraftDocument:
            
            seDoc.SaveAsDXF(dxfPath)
            print(f"✅ Exported to DXF: {dxfPath}")
            
        else:
            print("⚠️ DXF export skipped: not a Draft document.")
            
    except Exception as e:
        print(f"Error: {e}")
        
    finally:
        
        seDoc = None
        seApp = None
        
        
# Export the active Solid Edge Draft document to DWG
def exportToDwg():
    
    try:
        
        # Get running instance of Solid Edge
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
        seDoc = seApp.ActiveDocument
        
        baseName = os.path.splitext(os.path.basename(seDoc.FullName))[0]
        dwgPath = os.path.join(os.getcwd(), "assets", baseName + ".dwg")
        
        # Only export if document is a Draft
        if seDoc.Type == sef.igDraftDocument:
            
            seDoc.SaveAsDWG(dwgPath)
            print(f"✅ Exported to DWG: {dwgPath}")
            
        else:
            print("⚠️ DWG export skipped: not a Draft document.")
            
    except Exception as e:
        print(f"Error: {e}")
        
    finally:
        
        seDoc = None
        seApp = None
        

# Export the active Solid Edge Draft document to JT
def exportToJt():
    
    try:
        
        # Get running instance of Solid Edge
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
        seDoc = seApp.ActiveDocument
        
        baseName = os.path.splitext(os.path.basename(seDoc.FullName))[0]
        jtPath = os.path.join(os.getcwd(), "assets", baseName + ".jt")
        
        # Only export if document is a Draft
        if seDoc.Type == sef.igDraftDocument:
            
            seDoc.SaveAsJT(jtPath)
            print(f"✅ Exported to JT: {jtPath}")
            
        else:
            print("⚠️ JT export skipped: not a Draft document.")
            
    except Exception as e:
        print(f"Error: {e}")
        
    finally:
        
        seDoc = None
        seApp = None


# Main execution block
if __name__ == "__main__":
    
    os.system('cls')  # Clear console
    exportToPdf()
    # exportToDxf()
    # exportToDwg()
    # exportToJt()
