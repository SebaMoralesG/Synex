import pandas as pd
import os
from time import process_time
import matplotlib as plt


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
            df = pd.DataFrame(pd.read_excel(file,usecols= colmns,skiprows=rows_skiped, sheet_name = sheet_name))
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

def gen_bus_adder(dbus, column_names, df_index, *args):
    df = pd.DataFrame(0,columns = column_names, index = df_index)
    for data in args:
        for col in data:
            df.loc[:,dbus.loc[dbus["Gen. name"] == col,"Name"].all()] += data.loc[:,col]
    return df

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

def data_col_multiplier(data1, data2,df_cols, df_index):
    df = pd.DataFrame(0, columns = df_cols, index = df_index)
    for col in data1:
        if col in data2.columns:
            df.loc[:,col] = data1.loc[:,col].values*data2.loc[:,col].values
    return df

def xlsx_chart(chart_type,chart_book, sheet_name, chart_sheet,chart_size):
    chart = chart_book.add_chart(chart_type)
    for row in range(1, chart_size[0] + 1):
        chart.add_series({
            'name':       [sheet_name, row, 0],
            'categories': [sheet_name, 0, 1, 0,chart_size[1]],
            'values':     [sheet_name, row, 1, row, chart_size[1]]
        })

    chart.set_legend({'position': 'bottom', 'font': {'size': 9, 'Arial Narrow': True}})
    chart.set_size({'width': 700, 'height': 450})
    chart_sheet[sheet_name].insert_chart(1, 10, chart)


### Read Data

file_root = input("Please add file's root: ",)
file = "Tariff_income.xlsx"

last_date = (2050,12)

t0 = exectime("Reading Data...", True, 0)
date = init_date(os.path.join(file_root,"duraci.csv"))

cmgbus = readfile(os.path.join(file_root,"cmgbus.csv"),3,[])

demxba = readfile(os.path.join(file_root,"demxba.csv"),3,[])

gerhid = readfile(os.path.join(file_root,"gerhid.csv"),3,[])
gerter = readfile(os.path.join(file_root,"gerter.csv"),3,[])
gergnd = readfile(os.path.join(file_root,"gergnd.csv"),3,[])

dbus = readfile(os.path.join(file_root,"dbus.csv"),1,[1,6])
dbus = dbus.fillna("-")
dbus["Gen. name"] = dbus["Gen. name"].str.replace(" ","")

cmgbus = cmgbus.rename(columns= {col: col.replace(" ","") for col in cmgbus.columns})
demxba = demxba.rename(columns= {col: col.replace(" ","") for col in demxba.columns})
gerter = gerter.rename(columns= {col: col.replace(" ","") for col in gerter.columns})
gerhid = gerhid.rename(columns= {col: col.replace(" ","") for col in gerhid.columns})
gergnd = gergnd.rename(columns= {col: col.replace(" ","") for col in gergnd.columns})

cmgbus = month_year_index(cmgbus,date,(date[1],date[0]),())
gerhid = month_year_index(gerhid,date,(date[1],date[0]),())
gerter = month_year_index(gerter,date,(date[1],date[0]),())
gergnd = month_year_index(gergnd,date,(date[1],date[0]),())
demxba = month_year_index(demxba,date,(date[1],date[0]),(2050,12))
t0 = exectime(None, False, t0)

# Add gen by bus
t0 = exectime("Adding gen by buses...", True, 0)
gen_by_bus = gen_bus_adder(dbus,cmgbus.columns,gerter.index,gerter,gerhid,gergnd)
t0 = exectime(None, False, t0)

# Dem x MgC
t0 = exectime("Calculating MgC x Demand...", True, 0)
cmgbus_blck = cmgbus.groupby(level = [0,1,3]).mean()
dem_x_MgC = data_col_multiplier(demxba, cmgbus_blck, demxba.columns, cmgbus_blck.index)
dem_x_MgC = dem_x_MgC.groupby(level = [0,1]).sum()
t0 = exectime(None, False, t0)

# Gen X MgC
t0 = exectime("Calculating MgC x Generation...", True, 0)
gen_x_MgC = data_col_multiplier(gen_by_bus, cmgbus,gen_by_bus.columns, gen_by_bus.index)
gen_x_MgC = gen_x_MgC.groupby(level = [0,1,3]).mean().groupby(level = [0,1]).sum()
t0 = exectime(None, False, t0)

# total tariff income
t0 = exectime("Tariff Income...", True, 0)
dem_x_MgC["Total"] = dem_x_MgC.sum(axis = 1)
gen_x_MgC["Total"] = gen_x_MgC.sum(axis = 1)

Tariff_income = pd.DataFrame(0, columns = ["    Dem_x_MgC [M USD]","Gen_x_MgC [M USD]","Tariff_Income [M USD]"],index = dem_x_MgC.index)
Tariff_income["Dem_x_MgC [M USD]"] = dem_x_MgC["Total"]/1000
Tariff_income["Gen_x_MgC [M USD]"] = gen_x_MgC["Total"]/1000
Tariff_income["Tariff_Income [M USD]"] = (dem_x_MgC["Total"] - gen_x_MgC["Total"])/1000
Tariff_income["Demand [GWh]"] = demxba.groupby(level = [0,1,3]).mean().groupby(level = [0,1]).sum().sum(axis = 1)

Tariff_income_y = Tariff_income.groupby(level = 0).sum()
t0 = exectime(None, False, t0)

gen_by_bus.loc[:,(gen_by_bus != 0).any(axis = 0)].columns
mean_cmg_gen = cmgbus.loc[:,gen_by_bus.loc[:,(gen_by_bus != 0).any(axis = 0)].columns].groupby(level = [0,1,3]).mean().mean(axis = 1)
mean_cmg_dem = cmgbus.loc[:,demxba.loc[:,(demxba != 0).any(axis = 0)].columns].groupby(level = [0,1,3]).mean().mean(axis = 1)
fig = (mean_cmg_dem - mean_cmg_gen).to_csv(os.path.join(file_root,"MgC_diff.csv"))
df_out = pd.DataFrame(0,index = mean_cmg_dem.index, columns = ["MgC_Dem","MgC_gen","Diff"])
df_out["MgC_Dem"] = mean_cmg_dem
df_out["MgC_gen"] = mean_cmg_gen
df_out["Diff"] = mean_cmg_dem - mean_cmg_gen
df_out.to_csv(os.path.join(file_root,"MgC_diff.csv"))

gen_by_bus = gen_by_bus.loc[:,(gen_by_bus != 0).any(axis = 0)].mean(axis = 1)
demxba = demxba.loc[(demxba != 0).any(axis = 0)]
cmgbus_gen = cmgbus.loc[:,gen_by_bus.columns]
cmgbus_dem = cmgbus_blck.loc[:,demxba.columns]

(gen_by_bus.groupby(level = [0,1,3]).mean().groupby(level = [0,1]).sum().sum(axis = 1) - demxba.groupby(level = [0,1,3]).mean().groupby(level = [0,1]).sum().sum(axis = 1))/demxba.groupby(level = [0,1,3]).mean().groupby(level = [0,1]).sum().sum(axis = 1)

# Export data
t0 = exectime("Exporting Data...", True, 0)
with pd.ExcelWriter(os.path.join(file_root, file), engine = 'xlsxwriter') as writer:
    gen_by_bus.to_excel(writer, sheet_name = "Gen")
    demxba.to_excel(writer, sheet_name = "Demxba")
    cmgbus_dem.to_excel(writer, sheet_name = "MgC_dem")
    cmgbus_gen.to_excel(writer, sheet_name = "Mgc gen")
    Tariff_income.to_excel(writer, sheet_name = "Tariff_Income_month")
    Tariff_income_y.to_excel(writer, sheet_name = "Tariff_Income_Year")
    Tariff_income["Tariff_Income [M USD]"].unstack(level = 1).to_excel(writer, sheet_name = "Result")
    xlsx_chart({'type':'line'},writer.book,"Result",writer.sheets,Tariff_income["Tariff_Income [M USD]"].unstack(level = 1).shape)
t0 = exectime(None, False, t0)
print("Data exported.")