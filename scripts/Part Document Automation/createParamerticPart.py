# =============================== Script In Progress ==============================

import os
import comtypes.client
from ctypes import c_float
from comtypes.automation import IDispatch, VARIANT, VT_DISPATCH, VT_I4, VT_R4, DISPATCH_METHOD

# Load Solid Edge type libraries
comtypes.client.GetModule(r"C:\Program Files\Siemens\Solid Edge 2019\Program\framewrk.tlb")
comtypes.client.GetModule(r"C:\Program Files\Siemens\Solid Edge 2019\Program\Part.tlb")

from comtypes.gen import SolidEdgeFramework, SolidEdgePart

def getDispIdByName(dispatch, methodName):
    return dispatch.GetIDsOfNames(methodName)[0]

def createParametricCylinder():
    try:
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
        print("Connected to running Solid Edge instance.")
    except:
        seApp = comtypes.client.CreateObject("SolidEdge.Application", interface=SolidEdgeFramework.Application)
        print("Started new Solid Edge instance.")

    seApp.Visible = True

    seDocuments = seApp.Documents
    sePartDoc = seDocuments.Add("SolidEdge.PartDocument")
    sePartDoc = sePartDoc.QueryInterface(SolidEdgePart.PartDocument)
    sePartDoc.Activate()

    refPlanes = sePartDoc.RefPlanes
    frontPlane = refPlanes.Item(1)

    profileSet = sePartDoc.ProfileSets.Add()
    profiles = profileSet.Profiles
    profile = profiles.Add(frontPlane)
    profile = profile.QueryInterface(SolidEdgePart.Profile)

    circles2d = profile.Circles2d
    circles2d.AddByCenterRadius(0.0, 0.0, 0.025)
    profile.End(False)

    # Extrude



    print("Done.")

if __name__ == "__main__":
    os.system('cls')
    createParametricCylinder()
