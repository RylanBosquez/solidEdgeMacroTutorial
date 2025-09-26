import os
import comtypes.client
from comtypes.automation import VARIANT
import comtypes.gen.SolidEdgeDraft as sedraft

def weldCalloutIdentifier():
    
    # igWeldSymbol = -42581280
    
    try:
        seApp = comtypes.client.GetActiveObject("SolidEdge.Application")
    except OSError:
        print("No running Solid Edge instance found.")
        return

    seApp.Visible = True
    seActiveDoc = seApp.ActiveDocument
    seActiveDoc.Activate()

    activeSheet = seActiveDoc.ActiveSheet
    weldSymbols = activeSheet.WeldSymbols
    
    weldSymbolList = []
    for i in range(weldSymbols.Count):
        
        weldSymbolDict = {}
        weldSymbol = weldSymbols.Item(i+1)
        
        weldSymbolDict['ZSymbol'] = weldSymbol.ZSymbol
        weldSymbolDict['Tail'] = True if weldSymbol.Tail == 1 else False
        weldSymbolDict['WeldInField'] = weldSymbol.WeldInField
        weldSymbolDict['OffsetTopBottom'] = weldSymbol.OffsetTopBottom
        weldSymbolDict['WeldAllAround'] = weldSymbol.WeldAllAround
        
        weldSymbolDict['TopNote1'] = weldSymbol.TopNote1
        weldSymbolDict['TopNote2'] = weldSymbol.TopNote2
        weldSymbolDict['TopNote3'] = weldSymbol.TopNote3
        weldSymbolDict['TopNoteZ'] = weldSymbol.TopNoteZ
        weldSymbolDict['TopNoteAngle'] = weldSymbol.TopNoteAngle
        weldSymbolDict['TopNoteDepth'] = weldSymbol.TopNoteDepth
        weldSymbolDict['TopNoteCSize'] = weldSymbol.TopNoteCSize
        weldSymbolDict['TopType'] = weldSymbol.TopType
        weldSymbolDict['TopTreatmentType'] = weldSymbol.TopTreatmentType
        weldSymbolDict['TopPosOffset'] = weldSymbol.TopPosOffset
        weldSymbolDict['TopWeldModifier'] = weldSymbol.TopWeldModifier
        
        weldSymbolDict['TailNote'] = weldSymbol.TailNote
        weldSymbolDict['TailNote2'] = weldSymbol.TailNote2
        
        weldSymbolDict['BottomNote1'] = weldSymbol.BottomNote1
        weldSymbolDict['BottomNote2'] = weldSymbol.BottomNote2
        weldSymbolDict['BottomNote3'] = weldSymbol.BottomNote3
        weldSymbolDict['BottomNoteZ'] = weldSymbol.BottomNoteZ
        weldSymbolDict['BottomNoteAngle'] = weldSymbol.BottomNoteAngle
        weldSymbolDict['BottomNoteDepth'] = weldSymbol.BottomNoteDepth
        weldSymbolDict['BottomNoteCSize'] = weldSymbol.BottomNoteCSize
        weldSymbolDict['BottomType'] = weldSymbol.BottomType
        weldSymbolDict['BottomTreatmentType'] = weldSymbol.BottomTreatmentType
        weldSymbolDict['BottomPosOffset'] = weldSymbol.BottomPosOffset
        weldSymbolDict['BottomWeldModifier'] = weldSymbol.BottomWeldModifier
        
        weldSymbolDict['BreakLineDistance'] = weldSymbol.BreakLineDistance
        weldSymbolDict['BreakLine'] = weldSymbol.BreakLine
        weldSymbolDict['BreakLineDirection'] = weldSymbol.BreakLineDirection
        weldSymbolDict['Leader'] = weldSymbol.Leader
        
        weldSymbolDict['WeldInFieldFlagDirection'] = 'Right' if weldSymbol.WeldInFieldFlagDirection == 0 else 'Left'
        
        weldSymbolList.append(weldSymbolDict)
        
    return weldSymbolList
        

if __name__ == "__main__":
    
    os.system('cls')
    weldCalloutIdentifier()
    
    weldList = weldCalloutIdentifier()
    for i, weldSymbol in enumerate(weldList):
        
        print(f"\nWeld Symbol: {i}")
        for key, value in weldSymbol.items():
            print(f"   {key}: {value}")
