import pandas as pd
import numpy as np
import os
from time import process_time


def readfile(file,rows_skiped, colmns, sheet_name: str or int or list or None = 0):
    if file[-4:] == ".csv":
        if colmns == []:
            df = pd.DataFrame(pd.read_csv(file,skiprows= rows_skiped))
        else:
            df = pd.DataFrame(pd.read_csv(file,usecols= colmns,skiprows=rows_skiped))
    elif file[-5:] == ".xlsx":
        if colmns == []:
            df = pd.DataFrame(pd.read_excel(file,skiprows= rows_skiped, sheet_name = sheet_name))
        else:
            df = pd.DataFrame(pd.read_excel(file,usecols = colmns,skiprows=rows_skiped, sheet_name = sheet_name))
    df.head(10)
    return df

def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    # Print New Line on Complete
    if iteration == total: 
        print()

def exectime(text,flag,t0):
    if flag:
        print("______________________________________________")
        print(text)
        t0 = process_time()
        return t0
    else:
        t1 = process_time()
        print("Execution Time: ",t1-t0," sec.")
        print("______________________________________________")
        return None

def month_year_index(data, date,start_date,last_date):
    data["Year"] = date[1]
    data["Month"] = date[0]
    month = date[0]
    year = date[1]

    for stage in data["Stag"].unique():
        data.loc[data["Stag"] == stage, "Month"] = month
        data.loc[data["Stag"] == stage, "Year"] = year
        month += 1
        if month > 12:
            month = 1
            year += 1
    data = data.drop(columns = ["Stag"])
    data = data.set_index(["Year","Month","Seq.","Blck"])
    if last_date == []:
        data = data.loc[start_date:,:]
    else:
        data = data.loc[start_date:last_date,:]
    return data

def init_date(file):
    df = pd.DataFrame(pd.read_csv(file,nrows=1))
    df = df.rename(columns = {i:i.replace(" ","") for i in df.columns})
    return [int(df.columns[-2]),int(df.columns[-1])]

def gen_bus_adder(dbus, column_names, df_index, *args):
    df = pd.DataFrame(0,columns = column_names, index = df_index)
    for data in args:
        for col in data:
            df.loc[:,dbus.loc[dbus["Gen. name"] == col,"Name"].all()] += data.loc[:,col]
    return df

def data_col_multiplier(data1, data2,df_cols, df_index):
    df = pd.DataFrame(0, columns = df_cols, index = df_index)
    for col in data1:
        if col in data2.columns:
            df.loc[:,col] = data1.loc[:,col].values*data2.loc[:,col].values
    return df


#### Main Code


file_root = input("Please add file's root: ",)
file = "Min_Tec.xlsx"

last_date = (2050,12)

t0 = exectime("Reading Data...", True, 0)
date = init_date(os.path.join(file_root,"duraci.csv"))

cmgbus = readfile(os.path.join(file_root,"cmgbus.csv"),3,[])
demand = readfile(os.path.join(file_root,"demand.csv"),3,[])
gerter = readfile(os.path.join(file_root,"gerter.csv"),3,[])

dbus = readfile(os.path.join(file_root,"dbus.csv"),1,[1,6])
dbus = dbus.fillna("-")
dbus["Gen. name"] = dbus["Gen. name"].str.replace(" ","")

ctermise = readfile(os.path.join(file_root,"ctermise.csv"), 0, [])
ctermise = ctermise.iloc[:,[1,9,11,13]]

combmese = readfile(os.path.join(file_root,"combmese.csv"), 0, [])
combmese = combmese.drop(index = [0])
combmese = combmese.rename(columns = {combmese.columns[0]:"Year",combmese.columns[1]:"Month"})
combmese = combmese.astype("float")
combmese.loc[:,["Year","Month"]] = combmese.loc[:,["Year","Month"]].astype("int")
combmese = combmese.set_index(["Year","Month"])

cmgbus = cmgbus.rename(columns= {col: col.replace(" ","") for col in cmgbus.columns})
demand = demand.rename(columns= {col: col.replace(" ","") for col in demand.columns})
gerter = gerter.rename(columns= {col: col.replace(" ","") for col in gerter.columns})

cmgbus = month_year_index(cmgbus,date,(date[1],date[0]),())
demand = month_year_index(demand,date,(date[1],date[0]),())
gerter = month_year_index(gerter,date,(date[1],date[0]),())

Var_cost = pd.DataFrame(0,columns = gerter.columns,index = gerter.index)

combmese = combmese.loc[gerter.groupby(level = [0,1,3]).mean().groupby(level = [0,1]).sum().index,:]

for col in Var_cost.columns:
    Var_cost.loc[:,col] = combmese.loc[:,str(ctermise.loc[ctermise["...Nombre..."] == col,"Comb"].values[0])]*ctermise.loc[ctermise["...Nombre..."] == col,".CEsp.1"].values[0] + ctermise.loc[ctermise["...Nombre..."] == col,".CVaria"].values[0]


for col in gerter.columns:
    gerter.loc[cmgbus[dbus.loc[dbus["Gen. name"] == col,"Name"].all()] >= Var_cost[col],col] = 0

Gen_x_VC = gerter*Var_cost

# Add gen by bus
t0 = exectime("Adding gen by buses...", True, 0)
gen_by_bus = gen_bus_adder(dbus,cmgbus.columns,gerter.index,gerter)
Gen_x_VC = gen_bus_adder(dbus, cmgbus.columns,Gen_x_VC.index,Gen_x_VC)
Gen_x_VC = Gen_x_VC.groupby(level = [0,1,3]).mean().groupby(level = [0,1]).sum()
t0 = exectime(None, False, t0)

# Gen X MgC
t0 = exectime("Calculating MgC x Generation...", True, 0)
gen_x_MgC = data_col_multiplier(gen_by_bus, cmgbus,gen_by_bus.columns, gen_by_bus.index)
gen_x_MgC = gen_x_MgC.groupby(level = [0,1,3]).mean().groupby(level = [0,1]).sum()
t0 = exectime(None, False, t0)

gen_x_MgC["Total"] = gen_x_MgC.sum(axis = 1)
Gen_x_VC["Total"] = Gen_x_VC.sum(axis = 1)

df_out = pd.DataFrame(0,columns = ["MinTec_x_MgC","Cost_Ter","Demand","Compensation"], index = gen_x_MgC.index)
df_out["MinTec_x_MgC"] = gen_x_MgC["Total"]
df_out["Cost_Ter"] = Gen_x_VC["Total"]
df_out["Demand"] = demand.groupby(level = [0,1]).sum()
df_out["Compensation"] = (df_out["Cost_Ter"] - df_out["MinTec_x_MgC"])/df_out["Demand"]

df_out.to_excel(os.path.join(file_root,file))