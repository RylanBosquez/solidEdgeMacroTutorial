import os
import inspect
import comtypes.client
import comtypes.safearray
import comtypes.automation
import comtypes.gen.SolidEdgeDraft as sed
import comtypes.gen.SolidEdgeFramework as sef

from comtypes.automation import _midlSAFEARRAY


def listTlbFiles(solidEdgeFolder):

    for _, _, files in os.walk(solidEdgeFolder):
        for file in files:
            if file.endswith(".tlb"):
                print(file)
                

def inspectTlbs(searchList, solidEdgeFolder):
    
    tlbToModuleMap = {}
    for root, _, files in os.walk(solidEdgeFolder):
        for file in files:
            if file.endswith(".tlb"):
                tlbPath = os.path.join(root, file)
                try:
                    module = comtypes.client.GetModule(tlbPath)
                    print(f"Loaded: {file}")
                    tlbToModuleMap[file] = module.__name__  # Map TLB to module name
                except Exception as e:
                    print(f"Error loading {file}: {e}")

    classDict = {}
    functionDict = {}

    for _, module in comtypes.gen.__dict__.items():
        if inspect.ismodule(module):
            # Find which TLB this module came from
            tlbFile = next((tlb for tlb, modName in tlbToModuleMap.items() if modName == module.__name__), None)
            if not tlbFile:
                continue

            for memberName, memberObj in inspect.getmembers(module):
                if inspect.isclass(memberObj) and (memberName in searchList or len(searchList) == 0):
                    classDict.setdefault(tlbFile, []).append(memberName)
                elif inspect.isfunction(memberObj) and (memberName in searchList or len(searchList) == 0):
                    functionDict.setdefault(tlbFile, []).append(memberName)

    print(f"\nFound Classes:")
    for tlbFile, classes in classDict.items():
        print(f"{tlbFile}: {classes}")

    print(f"\nFound Functions:")
    for tlbFile, functions in functionDict.items():
        print(f"{tlbFile}: {functions}")
    

def inspectImport(module):

    for name, member in inspect.getmembers(module):
        if inspect.isfunction(member):
            print('Function', name)
        elif inspect.ismethod(member):
            print('Method', name)
        elif inspect.isclass(member):
            print('Class', name)
        else:
            print('Other', name)
            

def getMethodsFromClasses(tlbPath, classList):
    
    moduleName = comtypes.client.GetModule(tlbPath).__name__.split('.')[-1]
    
    try:
        module = getattr(comtypes.gen, moduleName)
    except AttributeError:
        print(f"Module '{moduleName}' not found in comtypes.gen.")
        return

    if len(classList) > 0:
        
        for className in classList:
            try:
                cls = getattr(module, className)
                print(f"\nClass: {className}")
                for name, member in inspect.getmembers(cls):
                    if inspect.ismethod(member):
                        print(f"  Method: {name}")
                    elif inspect.isfunction(member):
                        print(f"  Function: {name}")
            except AttributeError:
                print(f"Class '{className}' not found in module '{moduleName}'.")
                
    else:
        
        for i, (name, member) in enumerate(inspect.getmembers(module)):
            if inspect.isclass(member):
                print(f"Class: {name}")
            
            
            
def inspectComMethodSignatures(tlbPath, className, methodList):
    
    moduleName = comtypes.client.GetModule(tlbPath).__name__.split('.')[-1]
    try:
        module = getattr(comtypes.gen, moduleName)
        cls = getattr(module, className)
    except AttributeError:
        print(f"Class '{className}' not found in module '{moduleName}'.")
        return

    
    if hasattr(cls, '_methods_'):
        for method in cls._methods_:
            
            if method.name not in methodList and len(methodList) > 0:
                continue
            
            print(f"\nMethod: {method.name}")
            print(f"  Return Type: {method.restype}")
            for i, (argType, paramFlag) in enumerate(zip(method.argtypes, method.paramflags)):
                direction = paramFlag[0]
                name = paramFlag[1]
                print(f"  Param {i+1}: {name}, Type: {argType}, Direction: {direction}")
    else:
        print(f"No _methods_ found for class '{className}'.")
        
        
def inspectSeMethodSignatures(tlbPath, className, methodList):
    moduleName = comtypes.client.GetModule(tlbPath).__name__.split('.')[-1]
    try:
        module = getattr(comtypes.gen, moduleName)
        cls = getattr(module, className)
    except AttributeError:
        print(f"Class '{className}' not found in module '{moduleName}'.")
        return

    # Try to find the implemented interface (not just IDispatch)
    interfaces = [base for base in cls.__bases__ if hasattr(base, '_methods_') and base._methods_]
    if not interfaces:
        print(f"No usable interface found for class '{className}'.")
        return

    iface = interfaces[0]

    for method in iface._methods_:
        if methodList and method.name not in methodList:
            continue
        print(f"\nMethod: {method.name}")
        print(f"  Return Type: {method.restype}")
        argTypes = method.argtypes or []
        paramFlags = method.paramflags or []
        for i, (argType, paramFlag) in enumerate(zip(argTypes, paramFlags)):
            direction = paramFlag[0]
            name = paramFlag[1]
            print(f"  Param {i+1}: {name}, Type: {argType}, Direction: {direction}")



def getMethodParameters(tlbPath, classList):
    comtypes.client.GetModule(tlbPath)
    for cls in classList:
        print(f"\nInspecting: {cls.__name__}")
        if hasattr(cls, '_methods_') and cls._methods_:
            for method in cls._methods_:
                print(f"  Method: {method.name}")
                print(f"    Return Type: {method.restype}")
                for i, (argType, paramFlag) in enumerate(zip(method.argtypes, method.paramflags)):
                    direction = paramFlag[0]
                    name = paramFlag[1]
                    print(f"    Param {i+1}: {name}, Type: {argType}, Direction: {direction}")
        else:
            print("  No _methods_ found.")


def listAllModules(folderPath):
    
    for file in os.listdir(folderPath):
        if file.endswith(".py"):
            print(file)
            
            
                    
if __name__ == "__main__":
    
    os.system('cls')
    
    solidEdgeFolder = r"C:\Program Files\Siemens\Solid Edge 2019" # <-- change this to your Solid Edge folder
    tlbPath = r"C:\Program Files\Siemens\Solid Edge 2019\Program\fwksupp.tlb"# <-- change this to your TLB file
    moduleFolderPath = r"C:\Users\rbosquez\.conda\envs\solidEdgeEnv\Lib\site-packages\comtypes\gen" # <-- change this to your module folder
    
    #listTlbFiles(solidEdgeFolder)
    
    #inspectTlbs(['WeldSymbol'], solidEdgeFolder)
    
    #inspectImport(comtypes.automation)
    
    #getMethodsFromClasses(tlbPath, ['DisplayData', '_IDisplayDataAuto'])
    
    #getMethodParameters(tlbPath, ['WeldSymbol'])

    #findModuleContainingClass(tlbPath, 'DraftDocument')
    
    #listAllModules(moduleFolderPath)
    
    #inspectComMethodSignatures(tlbPath, 'DisplayData', ['GetTextAtIndex', 'GetTextAtIndexEx', 'GetTextAndFontAtIndex', 'GetTextAndFontAtIndexEx', 'GetLineAtIndex'])

    inspectComMethodSignatures(tlbPath, 'DisplayData', ['GetTextAtIndexEx'])
