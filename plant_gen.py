import pandas as pd
from time import process_time
import matplotlib.pyplot as plt
import numpy as np
import os


def readfile(file,rows_skiped, colmns):
    if file[-4:] == ".csv":
        if colmns == []:
            df = pd.DataFrame(pd.read_csv(file,skiprows= rows_skiped))
        else:
            df = pd.DataFrame(pd.read_csv(file,usecols= colmns,skiprows=rows_skiped))
    elif file[-5:] == ".xlsx":
        if colmns == []:
            df = pd.DataFrame(pd.read_excel(file,skiprows= rows_skiped))
        else:
            df = pd.DataFrame(pd.read_excel(file,usecols= colmns,skiprows=rows_skiped))
    df.head(10)
    return df


#### user inputs


file_root = input("Please add file path's: ",)
output_file_name = "plant_gen.xlsx"

PowerPlants = ["EOL_Totoral","EOL_SanJuan"]

print("Reading generations")
gergnd = readfile(os.path.join(file_root,"gergnd.csv"),3,[])
gerhid = readfile(os.path.join(file_root,"gerhid.csv"),3,[])
gerter = readfile(os.path.join(file_root,"gerter.csv"),3,[])
gergnd = gergnd.rename(columns= {col: col.replace(" ","") for col in gergnd.columns})
gerhid = gerhid.rename(columns= {col: col.replace(" ","") for col in gerhid.columns})
gerter = gerter.rename(columns= {col: col.replace(" ","") for col in gerter.columns})
gergnd = gergnd.set_index(["Stag","Seq.","Blck"])
gerhid = gerhid.set_index(["Stag","Seq.","Blck"])
gerter = gerter.set_index(["Stag","Seq.","Blck"])

print("Reading curtailments")
vergnd = readfile(os.path.join(file_root,"vergnd.csv"),3,[])
vergnd = vergnd.rename(columns= {col: col.replace(" ","") for col in vergnd.columns})
vergnd = vergnd.set_index(["Stag","Seq.","Blck"])

print("Reading Mg Costs")
cmgbus = readfile(os.path.join(file_root, "cmgbus.csv"), 3, [])
cmgbus = cmgbus.rename(columns= {col: col.replace(" ","") for col in cmgbus.columns})
cmgbus = cmgbus.set_index(["Stag","Seq.","Blck"])

print("Reading buses definition")
dbus = readfile(os.path.join(file_root, "dbus.csv"), 1, [])
dbus = dbus.iloc[:,[1,6]]
dbus = dbus.fillna("-")

total_gen = gerter.join(gerhid).join(gergnd,rsuffix = "_gnd")

mg_costs = pd.DataFrame(cmgbus.loc[:,"Polpaico220"])

gen_out = total_gen.loc[:,PowerPlants]
vert_out = pd.DataFrame(0, columns = PowerPlants, index = gen_out.index)
for col in vert_out.columns:
    mg_costs[dbus.loc[dbus["Gen. name"] == col,"Name"]] = cmgbus.loc[:,dbus.loc[dbus["Gen. name"] == col,"Name"]]
    if col in vergnd.columns:
        vert_out.loc[:,col] = vergnd.loc[:,col]

gen_out = gen_out.reset_index()
vert_out = vert_out.reset_index()
mg_costs = mg_costs.reset_index()

print("exporting file...")
with pd.ExcelWriter(output_file_name, engine = 'xlsxwriter') as writer:
    print("loading generation...")
    gen_out.to_excel(writer, sheet_name = "Generation", index = False)
    print("loading verts...")
    vert_out.to_excel(writer, sheet_name = "Vert", index = False)
    print("loading Mg Costs...")
    mg_costs.to_excel(writer, sheet_name = "Mg Costs", index = False)
    print(f"Generation analysis data exported to: {output_file_name}.")
