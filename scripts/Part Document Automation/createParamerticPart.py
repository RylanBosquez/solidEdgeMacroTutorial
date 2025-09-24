import os
import ctypes
import comtypes.client
from comtypes.safearray import _midlSAFEARRAY
from comtypes.automation import IDispatch

def createParametricCylinder():
    try:
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
        print("Connected to running Solid Edge instance.")
    except:
        seApp = comtypes.client.CreateObject("SolidEdge.Application")
        print("Started new Solid Edge instance.")

    seApp.Visible = True

    seDocuments = seApp.Documents
    sePartDoc = seDocuments.Add("SolidEdge.PartDocument")
    sePartDoc.Activate()

    refPlanes = sePartDoc.RefPlanes
    frontPlane = refPlanes.Item(1)

    profileSet = sePartDoc.ProfileSets.Add()
    profiles = profileSet.Profiles
    profile = profiles.Add(frontPlane)

    circles2d = profile.Circles2d
    circles2d.AddByCenterRadius(0.0, 0.0, 0.025)
    profile.End(False)

    print(f"Profile Type: {type(profile)}")

    # Force profile into IDispatch
    profile_dispatch = profile.QueryInterface(IDispatch)

    # Build SAFEARRAY of IDispatch
    SAFEARRAY_Dispatch = _midlSAFEARRAY(ctypes.POINTER(IDispatch))
    psa = SAFEARRAY_Dispatch.from_param([profile_dispatch])

    models = sePartDoc.Models

    models.AddFiniteExtrudedProtrusion(
        ctypes.c_long(1),     # NumberOfProfiles
        psa,                  # SAFEARRAY of profiles
        ctypes.c_long(2),     # Direction: 1 = Front, 2 = Back, 3 = Both
        ctypes.c_double(0.1)  # ExtrusionDistance
    )

    print("Done.")

if __name__ == "__main__":
    os.system('cls')
    createParametricCylinder()
