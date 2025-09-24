import os
import inspect
import comtypes.client
import comtypes.automation
import comtypes.safearray

from comtypes.automation import _midlSAFEARRAY


def listTlbFiles(solidEdgeFolder):

    for _, _, files in os.walk(solidEdgeFolder):
        for file in files:
            if file.endswith(".tlb"):
                print(file)
                

def inspectTlbs(tlbList, searchList, solidEdgeFolder):
    
    tlbToModuleMap = {}
    for root, _, files in os.walk(solidEdgeFolder):
        for file in files:
            if file.endswith(".tlb") and file in tlbList:
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
                if inspect.isclass(memberObj) and memberName in searchList:
                    classDict.setdefault(tlbFile, []).append(memberName)
                elif inspect.isfunction(memberObj) and memberName in searchList:
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

    for className in classList:
        try:
            cls = getattr(module, className)
            print(f"\nClass: {className}")
            for name, member in inspect.getmembers(cls):
                if inspect.isfunction(member) or inspect.ismethod(member):
                    print(f"  Method: {name}")
        except AttributeError:
            print(f"Class '{className}' not found in module '{moduleName}'.")
            
            
def getMethodParameters(tlbPath, className, methodList):
    moduleName = comtypes.client.GetModule(tlbPath).__name__.split('.')[-1]

    try:
        module = getattr(comtypes.gen, moduleName)
    except AttributeError:
        print(f"Module '{moduleName}' not found in comtypes.gen.")
        return

    try:
        cls = getattr(module, className)
    except AttributeError:
        print(f"Class '{className}' not found in module '{moduleName}'.")
        return

    for methodName in methodList:
        try:
            method = getattr(cls, methodName)
            sig = inspect.signature(method)
            print(f"\nMethod: {methodName}")
            for param in sig.parameters.values():
                print(f"  Param: {param.name} (Type: {param.annotation})")
        except AttributeError:
            print(f"\nMethod '{methodName}' not found in class '{className}'.")
        except ValueError:
            print(f"\nMethod '{methodName}' has no inspectable signature.")



if __name__ == "__main__":
    
    os.system('cls')
    
    solidEdgeFolder = r"C:\Program Files\Siemens\Solid Edge 2019\Program" # <-- change this to your Solid Edge folder
    tlbPath = "C:\Program Files\Siemens\Solid Edge 2019\Program\Part.tlb"
    
    #listTlbFiles(solidEdgeFolder)
    
    # inspectTlbs(
    #     ['framewrk.tlb', 'Part.tlb', 'geometry.tlb', 'constants.tlb', 'SolidEdgeGateway.tlb'],
    #     ['PartDocument', 'Model', 'Body', 'MassProperties'],
    #     solidEdgeFolder)
    
    #inspectImport(comtypes.automation)
    
    #getMethodsFromClasses(tlbPath, ['PartDocument', 'Model'])
    
    getMethodParameters(tlbPath, '_IModelAuto', [
        'ComputePhysicalProperties',
        'ComputePhysicalPropertiesRelativeToCoordinateSystem',
        'ComputePhysicalPropertiesWithSpecifiedDensity',
        'GetPhysicalProperties',
        'GetPhysicalPropertiesRelativeToCoordinateSystem'
    ])
