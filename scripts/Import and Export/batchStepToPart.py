import os
import comtypes.client

def batchStepToPart(stepFolder):
    try:
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
    except:
        seApp = comtypes.client.CreateObject("SolidEdge.Application", interface=comtypes.gen.SolidEdgeFramework.Application)

    seApp.Visible = True
    documents = seApp.Documents

    # Manually specify the Part template path
    partTemplatePath = os.path.abspath(os.path.join(os.getcwd(), "assets/iso metric part.par"))

    for fileName in os.listdir(stepFolder):
        if fileName.lower().endswith(".step") or fileName.lower().endswith(".stp"):
            stepPath = os.path.join(stepFolder, fileName)
            documents.OpenWithTemplate(stepPath, partTemplatePath)

if __name__ == "__main__":
    folderPath = os.path.abspath(os.path.join(os.getcwd(), "assets/stepFolder"))
    os.system('cls')
    batchStepToPart(folderPath)
