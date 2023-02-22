import pandas as pd
import os
from time import process_time
import matplotlib.pyplot as plt
import numpy as np


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

def exectime(file,flag,t0):
    if flag:
        print("______________________________________________")
        print("Processing File ", file)
        t0 = process_time()
        return t0
    else:
        t1 = process_time()
        print("Execution Time: ",t1-t0," sec.")
        print("______________________________________________")
        return 0

def init_date(file):
    df = pd.DataFrame(pd.read_csv(file,nrows=1))
    df = df.rename(columns = {i:i.replace(" ","") for i in df.columns})
    return [int(df.columns[-2]),int(df.columns[-1])]

def month_year_index(data, date,start_date):
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
    data = data.loc[start_date:,:]
    return data

def df_row_add(serie,first_value):
    first_row = first_value
    df_out = serie + serie.shift(1)
    serie.iloc[0] = first_row
    return df_out

def df_row_substract(serie, config, name):
    if not config.loc[config["Name"] == name,"Type"].all():
        first_row = serie.iloc[0] - config.loc[config["Name"] == name,"Pot"].values[0]
    else:
        first_row = serie.iloc[0]
    serie -= serie.shift(1)
    serie.iloc[0] = first_row
    return serie

def Gen_work_schedule(mgnd,mhidro,mtermi,dictionary):
    mgnd_out = pd.DataFrame(mgnd)
    mgnd_out["Technology"] = None
    mgnd_out["Substation"] = None
    mgnd_out["Developer"] = None
    for index, row in mgnd_out.iterrows():
        mgnd_out.loc[index, "Substation"] = dictionary.loc[dictionary["SDDP"] == mgnd_out.loc[index,"Name"],"Barra"].all()
        mgnd_out.loc[index, "Developer"] = dictionary.loc[dictionary["SDDP"] == mgnd_out.loc[index,"Name"],"Desarrollador"].all()
        if mgnd_out.loc[index,"Name"][:3] == "HIB":
            mgnd_out.loc[index,"Technology"] = "Hybrid Solar PV"
        else:
            mgnd_out.loc[index,"Technology"] = dictionary.loc[dictionary["SDDP"] == mgnd_out.loc[index,"Name"],"Tecnología Inglés"].all()
        if dictionary.loc[dictionary["SDDP"] == mgnd_out.loc[index,"Name"],"Agrupacion en Informe?"].all() != "-":
            mgnd_out.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == mgnd_out.loc[index,"Name"],"Agrupacion en Informe?"].all()
        else:
            mgnd_out.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == mgnd_out.loc[index,"Name"],"Informe Inglés"].all()

    mhidro_out = pd.DataFrame(mhidro)
    mhidro_out["Technology"] = None
    mhidro_out["Substation"] = None
    mhidro_out["Developer"] = None
    for index, row in mhidro_out.iterrows():
        mhidro_out.loc[index, "Substation"] = dictionary.loc[dictionary["SDDP"] == mhidro_out.loc[index,"Name"],"Barra"].all()
        mhidro_out.loc[index, "Developer"] = dictionary.loc[dictionary["SDDP"] == mhidro_out.loc[index,"Name"],"Desarrollador"].all()
        if mhidro_out.loc[index,"Name"][:3] == "HIB":
            mhidro_out.loc[index,"Technology"] = "Hybrid Storage"
        else:
            mhidro_out.loc[index,"Technology"] = dictionary.loc[dictionary["SDDP"] == mhidro_out.loc[index,"Name"],"Tecnología Inglés"].all()
        if dictionary.loc[dictionary["SDDP"] == mhidro_out.loc[index,"Name"],"Agrupacion en Informe?"].all() != "-":
            mhidro_out.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == mhidro_out.loc[index,"Name"],"Agrupacion en Informe?"].all()
        else:
            mhidro_out.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == mhidro_out.loc[index,"Name"],"Informe Inglés"].all()
        

    mtermi_out = pd.DataFrame(mtermi)
    mtermi_out["Technology"] = None
    mtermi_out["Substation"] = None
    mhidro_out["Developer"] = None
    for index, row in mtermi_out.iterrows():
        mtermi_out.loc[index, "Substation"] = dictionary.loc[dictionary["SDDP"] == mtermi_out.loc[index,"Name"],"Barra"].all()
        mtermi_out.loc[index, "Developer"] = dictionary.loc[dictionary["SDDP"] == mtermi_out.loc[index,"Name"],"Desarrollador"].all()
        mtermi_out.loc[index,"Technology"] = dictionary.loc[dictionary["SDDP"] == mtermi_out.loc[index,"Name"],"Tecnología Inglés"].all()
        mtermi_out.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == mtermi_out.loc[index,"Name"],"Informe Inglés"].all()

    return mgnd_out, mhidro_out, mtermi_out

def mod_net_pot(mod, config):
    for unit in mod["Name"].unique():
        net_power = df_row_substract(mod.loc[mod["Name"] == unit,"Pot"], config, unit)
        mod.loc[mod["Name"] == unit,"Pot"] = net_power.values
    return mod

def synex_style(data, workbook, worksheet,type_chart, type_flag, language):
    
    # Header table color
    header_format = workbook.add_format({"fg_color":"#16365C",'font_name':"Arial Narrow",'font_color':"white",'align':"center"})

    # row and col index
    row_idx, col_idx = data.shape

    if type_chart:
        if language:
            worksheet.write(0,0,"Year",header_format)
        else:
            worksheet.write(0,0,"Año",header_format)
    
    for col_num, value in enumerate(data.columns.values):
        worksheet.write(0,col_num + 1, value, header_format)
    
    for r in range(row_idx):
        worksheet.write( r+1, 0, data.index[r], workbook.add_format({'num_format':1,'align':"center",'font_name':"Arial Narrow",'right':1}))
        if r == row_idx - 1:
            worksheet.write( r+1, 0, data.index[r], workbook.add_format({'num_format':1,'align':"center",'font_name':"Arial Narrow",'bottom':6,'right':1}))
        for c in range(col_idx):
            if c == 0 and type_flag == 0 and r == row_idx - 1:
                worksheet.write( r+1, c+1, data.loc[ data.index[r], data.columns[c]], workbook.add_format({'num_format':'mmm-yy','align':"center",
                    'font_name':"Arial Narrow",'bottom':6,'right':1}))
            elif c == 0 and type_flag == 0 and r < row_idx - 1:
                worksheet.write( r+1, c+1, data.loc[ data.index[r], data.columns[c]], workbook.add_format({'num_format':'mmm-yy','align':"center",
                    'font_name':"Arial Narrow",'right':1}))
            elif c < col_idx - 1 and r < row_idx - 1:
                worksheet.write( r+1, c+1, data.loc[ data.index[r], data.columns[c]], workbook.add_format({'num_format':3,'align':"center",
                    'font_name':"Arial Narrow",'right':1}))
            elif r == row_idx - 1 and c < col_idx - 1:
                worksheet.write( r+1, c+1, data.loc[ data.index[r], data.columns[c]], workbook.add_format({'num_format':3,'align':"center",
                    'font_name':"Arial Narrow",'right':1,'bottom':6}))
            elif c == col_idx - 1 and r == row_idx - 1:
                worksheet.write( r+1, c+1, data.loc[data.index[r], data.columns[c]], workbook.add_format({'num_format':3,'align':"center",
                    'font_name':"Arial Narrow",'bottom':6}))
            else:
                worksheet.write( r+1, c+1, data.loc[ data.index[r], data.columns[c]], workbook.add_format({'num_format':3,'align':"center",
                    'font_name':"Arial Narrow"}))

def xlsx_chart(chart_type, chart_colors,chart_book, sheet_name, chart_sheet,chart_size, language,type_flag):
    chart = chart_book.add_chart(chart_type)
    for col in range(1, chart_size[1]):
        chart.add_series({
            'name':       [sheet_name, 0, col],
            'categories': [sheet_name, 1, 0,   chart_size[0], 0],
            'values':     [sheet_name, 1, col, chart_size[0], col],
            'fill':       {'color': chart_colors[col - 1]},
            'line':       {'color': chart_colors[col - 1]},
            'gap':        50
        })

    chart.set_legend({'position': 'right', 'font': {'size': 9, 'Arial Narrow': True}})
    chart.set_size({'width': 625, 'height': 345})
    chart_sheet[sheet_name].insert_chart(1, 15, chart)

def export_excel(sheet_list, type_flag, file_root, filename, language,*args):

    t0 = exectime("Exporting data to: " + filename, True, 0)
    printProgressBar(0, len(sheet_list), prefix = 'Progress:', suffix = 'Complete', length = 50)

    with pd.ExcelWriter(os.path.join(file_root,filename), engine = 'xlsxwriter') as writer:
        for i in range(len(sheet_list)):
            if type_flag[i] == 0:              # Generation incoming capacity
                args[i].to_excel(writer, sheet_name = sheet_list[i], index = False)
                synex_style(args[i], writer.book, writer.sheets[sheet_list[i]],False, type_flag[i], language)

            elif type_flag[i] == 1:            # Generation works plan
                args[i].to_excel(writer, sheet_name = sheet_list[i])
                synex_style(args[i], writer.book, writer.sheets[sheet_list[i]], True, type_flag[i], language)
                xlsx_chart({'type':'column', "subtype":"stacked"},
                    ["orange","red","black","gray","brown","green","lime","blue","purple","yellow","cyan","magenta"],
                    writer.book,sheet_list[i],writer.sheets,args[i].shape,language,type_flag)

            elif type_flag[i] == 2:
                args[i].to_excel(writer, sheet_name = sheet_list[i])
                synex_style(args[i], writer.book, writer.sheets[sheet_list[i]], True, type_flag[i], language)
                xlsx_chart({'type':'column', "subtype":"stacked"},
                    ["orange","red","black","gray","brown","green","lime","blue","purple","yellow","cyan","magenta"],
                        writer.book,sheet_list[i],writer.sheets,args[i].shape,language,type_flag)


            printProgressBar(i+1, len(sheet_list), prefix = sheet_list[i], suffix = 'Complete', length = 50)           # 'Progress:'
    t0 = exectime("Exporting data to: "+filename, False, t0)


##### MAIN CODE

start_date = (1,2023)

technologies  = ["Biomass","Coal","Diesel","Gas","Geothermal","Hybrid Storage","Hybrid Solar PV","Hydro","Solar CSP","Solar PV","Wind Farm","Battery"]
technologies_esp  = ["Biomasa","Carbón","Diesel","Gas","Geotérmica","Híbrido Almacenamiento","Híbrido Solar FV","Hidro","Solar CSP","Solar FV","Eólica","Batería"]

file_root = input("Please add file path's: ",)
output_file_name = "WorkPlan.xlsx"
english = True

date = init_date(os.path.join(file_root,"duraci.csv")) 
start_date_output = (date[0],date[1])

print("Reading inital capacities [SEN]...")
init_capacity = readfile("CEN_RealGen.xlsx",3,[21,22])
init_capacity = init_capacity.dropna()
init_capacity = init_capacity.rename(columns = {"Technology.1":"Technology"})
init_capacity = init_capacity.set_index("Technology")


print("Reading power plant Dictionary...")
dictionary = readfile("Power_Plant_Dictionary.xlsx",0,[0,1,2,3,4,5,6,7])
dictionary = dictionary.fillna("-")

dbus = readfile(os.path.join(file_root,"dbus.csv"),1,[1,6])
dbus = dbus.fillna("-")
dbus["Gen. name"] = dbus["Gen. name"].str.replace(" ","")

print("Reading Power Plant config...")
ctermise = readfile(os.path.join(file_root,"ctermise.csv"), 0, [])
cgndse = readfile(os.path.join(file_root,"cgndse.csv"), 0, [])
chidrose = readfile(os.path.join(file_root,"chidrose.csv"), 0, [])
ctermise = ctermise.iloc[:,[1,3,4]]
chidrose = chidrose.iloc[:,[1,6,7]]
cgndse = cgndse.iloc[:,[1,3,5]]

ctermise = ctermise.rename(columns = {"...Nombre...":"Name",".PotIns":"Pot", "Tipo":"Type"})
cgndse = cgndse.rename(columns = {"Name":"Name","PotIns":"Pot"})
chidrose = chidrose.rename(columns = {"...Nombre...":"Name","....Pot":"Pot", "Tipo":"Type"})
cgndse["Name"] = cgndse["Name"].str.replace(" ","")
chidrose["Name"] = chidrose["Name"].str.replace(" ","")
ctermise["Name"] = ctermise["Name"].str.replace(" ","")

print("Reading Power plants modifications...")
mtermise = readfile(os.path.join(file_root,"mtermise.csv"), 0, [])
mgndse = readfile(os.path.join(file_root,"mgndse.csv"), 0, [])
mhidrose = readfile(os.path.join(file_root,"mhidrose.csv"), 0, [])

mtermise = mtermise.iloc[:,[1,2,5]]
mtermise = mtermise.rename(columns= {"Nombre":"Name","GerMax":"Pot"})
mtermise = mtermise.sort_values(by = "Data")
mtermise["Name"] = mtermise["Name"].str.replace(" ","")

mgndse = mgndse.iloc[:,[1,2,4]]
mgndse = mgndse.rename(columns= {"Nombre":"Name",".PotIns":"Pot"})
mgndse = mgndse.sort_values(by = "Data")
mgndse["Name"] = mgndse["Name"].str.replace(" ","")

mhidrose = mhidrose.iloc[:,[1,2,9]]
mhidrose = mhidrose.rename(columns= {"Nombre":"Name"})
mhidrose = mhidrose.sort_values(by = "Data")
mhidrose["Name"] = mhidrose["Name"].str.replace(" ","")

## Clean Data

mgndse["year"] = None
mgndse["month"] = None
mhidrose["year"] = None
mhidrose["month"] = None
mtermise["year"] = None
mtermise["month"] = None

mgndse["year"] = mgndse["Data"].str.split("-").str[0].astype("int")
mgndse["month"] = mgndse["Data"].str.split("-").str[1].astype("int")
mhidrose["year"] = mhidrose["Data"].str.split("-").str[0].astype("int")
mhidrose["month"] = mhidrose["Data"].str.split("-").str[1].astype("int")
mtermise["year"] = mtermise["Data"].str.split("-").str[0].astype("int")
mtermise["month"] = mtermise["Data"].str.split("-").str[1].astype("int")

mgndse = mod_net_pot(mgndse, cgndse)
mhidrose = mod_net_pot(mhidrose, chidrose)
mtermise = mod_net_pot(mtermise, ctermise)

print("Loading Work's plan...")
mgnd,mhidro,mtermi = Gen_work_schedule(mgndse, mhidrose, mtermise,dictionary)

work_plan = mgnd.append(mhidro).append(mtermi)
work_plan = work_plan.loc[work_plan["Pot"] != 0, ["year","month","Name","Technology","Pot","Substation","Developer"]].sort_values(by = ["year","month"])
work_plan = work_plan.loc[work_plan["year"] >= start_date[1],:]
work_plan["day"] = 1
df_index = pd.to_datetime(work_plan.loc[:,["year","month","day"]])
work_plan["COD"] = df_index



print("Loading Power plants additions...")
gen_additions = work_plan.loc[:,["year","month","Technology","Pot"]]
gen_additions = gen_additions.groupby(["year","month","Technology"]).sum()
gen_additions = gen_additions.unstack(level = [2])
gen_additions = gen_additions.fillna(0)
gen_additions = gen_additions.groupby(level = 0).sum()
gen_additions.loc[start_date[1]-1] = 0
gen_additions = gen_additions.sort_index()
gen_additions.columns = gen_additions.columns.droplevel(0)

print("Loading Power plants expantion...")
gen_expantion = pd.DataFrame(0,columns = init_capacity.index, index = gen_additions.index)

init_capacity.loc["Coal","Capacity[MW]"] = -gen_additions.loc[:,"Coal"].sum()

for col in gen_expantion.columns:
    if col not in gen_additions.columns:
        gen_additions[col] = 0
gen_additions = gen_additions.loc[:,gen_expantion.columns]

for index in gen_expantion.index:
    for col in gen_expantion.columns:
        if index == start_date[1] - 1:
            gen_expantion.loc[index, col] = init_capacity.loc[col, "Capacity[MW]"]
        else:
            gen_expantion.loc[index,col] = gen_additions.loc[index,col] + gen_expantion.loc[index - 1, col]

gen_additions = gen_additions.loc[start_date[1]:,:]
gen_expantion["Total"] = gen_expantion.sum(axis = 1)
gen_additions["Total"] = gen_additions.sum(axis = 1)



work_plan = work_plan.fillna("-")
work_plan = work_plan.drop(columns = ["year","month","day"])
work_plan = work_plan.rename(columns = {"Name":"Project", "Pot":"MW"})
work_plan = work_plan.groupby(["COD","Project","Technology","Substation","Developer"], as_index = False).sum()
work_plan = work_plan.loc[:,["COD","Project","Technology","Substation","Developer","MW"]]


export_excel(["WorksPlan","AddedCap","GenCap"],[0,1,2], file_root, output_file_name, english, work_plan, gen_additions, gen_expantion)