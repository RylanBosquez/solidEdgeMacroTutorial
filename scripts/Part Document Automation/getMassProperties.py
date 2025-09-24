import os
import comtypes.client
import pythoncom
from comtypes.safearray import _midlSAFEARRAY
import ctypes

def getMassProperties():
    pythoncom.CoInitialize()

    # Connect to Solid Edge
    try:
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
    except:
        seApp = comtypes.client.CreateObject("SolidEdge.Application")

    seApp.Visible = True
    sePartDoc = seApp.ActiveDocument
    model = sePartDoc.Models.Item(1)

    # Scalars
    lngStatus = ctypes.c_long()
    dblDensity = ctypes.c_double()
    dblAccuracyIn = ctypes.c_double()
    dblVolume = ctypes.c_double()
    dblArea = ctypes.c_double()
    dblMass = ctypes.c_double()
    dblAccuracyOut = ctypes.c_double()

    # Create empty SAFEARRAYs of double for vector outputs
    SafeDouble = _midlSAFEARRAY(ctypes.c_double)
    emptyVector3 = SafeDouble.from_param([0.0, 0.0, 0.0])   # 3-element vectors
    emptyVector9 = SafeDouble.from_param([0.0]*9)           # 9-element vectors (axes)
    
    # Call GetPhysicalProperties
    model.GetPhysicalProperties(
        ctypes.byref(lngStatus),
        ctypes.byref(dblDensity),
        ctypes.byref(dblAccuracyIn),
        ctypes.byref(dblVolume),
        ctypes.byref(dblArea),
        ctypes.byref(dblMass),
        emptyVector3,    # CenterOfGravity
        emptyVector3,    # CenterOfVolume
        emptyVector3,    # GlobalMomentsOfInertia
        emptyVector3,    # PrincipalMomentsOfInertia
        emptyVector9,    # PrincipalAxes
        emptyVector3,    # RadiiOfGyration
        ctypes.byref(dblAccuracyOut)
    )

    # Print results
    print("Mass Properties of Active Model:")
    print(f"  Status: {lngStatus.value}")
    print(f"  Density: {dblDensity.value}")
    print(f"  Accuracy In: {dblAccuracyIn.value}")
    print(f"  Volume: {dblVolume.value}")
    print(f"  Surface Area: {dblArea.value}")
    print(f"  Mass: {dblMass.value}")
    print(f"  Center of Gravity: {list(emptyVector3)}")
    print(f"  Center of Volume: {list(emptyVector3)}")
    print(f"  Global Moments of Inertia: {list(emptyVector3)}")
    print(f"  Principal Moments of Inertia: {list(emptyVector3)}")
    print(f"  Principal Axes: {list(emptyVector9)}")
    print(f"  Radii of Gyration: {list(emptyVector3)}")
    print(f"  Relative Accuracy Achieved: {dblAccuracyOut.value}")

if __name__ == "__main__":
    os.system('cls')
    getMassProperties()
