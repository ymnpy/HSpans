import win32com.client
import pandas as pd
import os
from pyNastran.bdf.bdf import read_bdf

def main(hdb_path,project_name):
    """
    open up hs,
    import fem,
    make project,
    make assembly,
    assign pid and concept,
    take out spans
    """
    hs = win32com.client.Dispatch('HyperSizer.Application')
    hs.Login("Hypersizer Admin","")
    hs.OpenDatabase(hdb_path)
    
    project_name+="_bucklingSpan"
    assembly_name="buckbuck"
    try: hs.Projects.create(project_name)
    except: pass
    
    project=hs.Projects.GetProject(project_name)
    rundeck=project.Rundecks.Item(1)
    rundeck.PathFem = bdf_path
    rundeck.PathResults = op2_path
    project.Save()
    project.ImportFEM()

    try: assembly=project.Assemblies.Create(assembly_name,3,3) #materialmode -> 3 any | analysismode -> 3 detail
    except: assembly=project.Assemblies.GetAssembly(assembly_name)
    
    for pid,prop in bdf.properties.items():
        if prop.type=="PCOMP" or prop.type=="PSHELL":
            try: assembly=project.Assemblies.Create(assembly_name,3,3) #materialmode -> 3 any | analysismode -> 3 detail
            except: assembly=project.Assemblies.GetAssembly(assembly_name)
            assembly.ComponentIds.Add(pid)
        
    assembly.SetGroupConcepts([1], 1, 0) #conceptIds,category,componentType -> smeared
    assembly.ComponentIds.Save()
    project.Save()
    
    buckling_assembly=project.Assemblies.GetAssembly(assembly_name)
    compids=buckling_assembly.ComponentIDs.toArray()
    
    storage=[]
    for compid in compids:
        comp=project.Components.GetComponent(compid)
        xspan=comp.PanelProperty(21)
        yspan=comp.PanelProperty(22)
        storage.append((compid,xspan,yspan))
    
    return pd.DataFrame(storage,columns=["PID","X_SPAN","Y_SPAN"])
    

if __name__=="__main__":
    hdb_path=r"..."
    bdf_path=r"..."
    op2_path=r"..."
    
    bdf=read_bdf(bdf_path,xref=False)
    file_name=os.path.split(bdf_path)[-1].rstrip(".bdf")
    
    df_out=main(hdb_path,file_name)
    df_out.to_excel(f"{file_name}_HSpans.xlsx",index=False)




