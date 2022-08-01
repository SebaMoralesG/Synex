import pandas as pd
from time import process_time
import matplotlib.pyplot as plt
import numpy as np



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

def exportexcel(listnames, file, adderlist, *args):
    t0 = exectime("Exporting data to: "+file, True, 0)
    printProgressBar(0, len(args), prefix = 'Progress:', suffix = 'Complete', length = 50)
    with pd.ExcelWriter(file, engine = 'xlsxwriter') as writer:
        for i in range(len(args)):
            args[i].to_excel(writer, sheet_name = listnames[i])
            printProgressBar(i+1, len(args), prefix = listnames[i], suffix = 'Complete', length = 50)           # 'Progress:'
    t0 = exectime("Exporting data to: "+file, False, t0)

def exportexcelunstack(listnames, file_root, file, level_unstacked, data, type_data_flag, language,cmg_lim):

    t0 = exectime("Exporting data to: "+file, True, 0)
    printProgressBar(0, len(listnames), prefix = 'Progress:', suffix = 'Complete', length = 50)
    with pd.ExcelWriter(file_root + file, engine = 'xlsxwriter') as writer:
        for i in range(len(listnames)):
            data[listnames[i]].unstack(level = level_unstacked).to_excel(writer, sheet_name = listnames[i])
            synex_style(data[listnames[i]].unstack(level = level_unstacked), listnames[i], writer.book, writer.sheets[listnames[i]], type_data_flag)
            plot_CMg_3d(data, listnames[i],language,cmg_lim,file_root)
            writer.sheets[listnames[i]].insert_image('AA1', file_root + listnames[i] + ".png")
            printProgressBar(i+1, len(listnames), prefix = listnames[i], suffix = 'Complete', length = 50)           # 'Progress:'

    t0 = exectime("Exporting data to: "+file, False, t0)

def exectime(file,flag,t0):
    if flag == True:
        print("______________________________________________")
        print("Processing File ", file)
        t0 = process_time()
        return t0
    else:
        t1 = process_time()
        print("Execution Time: ",t1-t0," sec.")
        print("______________________________________________")
        return 0

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

def init_date(file):
    df = pd.DataFrame(pd.read_csv(file,nrows=1))
    df = df.rename(columns = {i:i.replace(" ","") for i in df.columns})
    return [int(df.columns[-2]),int(df.columns[-1])]

def month_year_index(data, date):
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
    return data

def daily_format(data, working_blocks, Nonworking_blocks, time_data, divide_flag):

    hours = [23,24,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22]
    if divide_flag == True:
        data = data.div(time_data.values)
    data_workday = data[data.index.get_level_values("Blck").isin(working_blocks)]
    data_workday = pd.DataFrame(np.repeat( data_workday.values, 2, axis=0), columns = data_workday.columns, index = np.repeat(data_workday.index,2))
    data_workday = data_workday.reset_index(level = 2)

    for index, row in data_workday.iterrows():
        data_workday.loc[index,"Blck"] = hours
    data_workday = data_workday.reset_index()
    data_workday = data_workday.set_index(["Year","Month","Blck"])
    data_workday = data_workday.sort_index()


    data_nonworkday = data[data.index.get_level_values("Blck").isin(Nonworking_blocks)]
    data_nonworkday = pd.DataFrame(np.repeat( data_nonworkday.values, 2, axis=0), columns = data_nonworkday.columns, index = np.repeat(data_nonworkday.index,2))
    data_nonworkday = data_nonworkday.reset_index(level = 2)

    for index, row in data_nonworkday.iterrows():
        data_nonworkday.loc[index,"Blck"] = hours
    data_nonworkday = data_nonworkday.reset_index()
    data_nonworkday = data_nonworkday.set_index(["Year","Month","Blck"])
    data_nonworkday = data_nonworkday.sort_index()

    data_meanday = (data_workday + data_nonworkday)/2

    data_meanday = data_meanday.groupby(level = [0,2]).mean()
    data_workday = data_workday.groupby(level = [0,2]).mean()
    data_nonworkday = data_nonworkday.groupby(level = [0,2]).mean()

    return data_workday, data_nonworkday, data_meanday

def synex_style(data, workbook, worksheet,type_chart, language):
    
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
            if c < col_idx - 1 and r < row_idx - 1:
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

def polygon_under_graph(x, y):
    """
    Construct the vertex list which defines the polygon filling the space under
    the (x, y) line graph. This assumes x is in ascending order.
    """
    return [(x[0], 0.), *zip(x, y), (x[-1], 0.)]

def data_expand(conf_data, mod_data, start_date):

    date_index = [f"{start_date-1}-12"]

    for y in range(2022,2051):
        for m in range(1,13):
            if m < 10:
                date_index.append(f"{y}-0{m}")
            else:
                date_index.append(f"{y}-{m}")

    gen_expantion = pd.DataFrame(0, columns = conf_data["Nombre"], index = date_index)

    gen_expantion.loc[gen_expantion.index[0],conf_data[conf_data["Type"] == 0]["Nombre"]] = conf_data[conf_data["Type"] == 0]["Pot"].values

    gen_expantion.loc[gen_expantion.index[0],mod_data[mod_data["Data"] <= gen_expantion.index[0]]["Name"]] = mod_data[mod_data["Data"] <= gen_expantion.index[0]]["Pot"].values

    for index, row in mod_data[mod_data["Data"] > gen_expantion.index[0]].iterrows():
        if row["Pot"] == 0:
            gen_expantion.loc[row["Data"],row["Name"]] = -1
        else:
            gen_expantion.loc[row["Data"],row["Name"]] = row["Pot"]
    
    last_index = gen_expantion.index[0]
    for r in gen_expantion.index:
        if r == gen_expantion.index[0]:
            pass
        else:
            for c in gen_expantion.columns:
                if gen_expantion.loc[ r, c] == 0:
                    gen_expantion.loc[ r, c] = gen_expantion.loc[ last_index, c]
                elif gen_expantion.loc[ r, c] == -1:
                    gen_expantion.loc[ r, c] = 0
            last_index = r

    return gen_expantion

def export_excel(sheet_list, type_flag, file_root, filename, language,*args):

    t0 = exectime("Exporting data to: " + filename, True, 0)
    printProgressBar(0, len(sheet_list), prefix = 'Progress:', suffix = 'Complete', length = 50)

    with pd.ExcelWriter(file_root + filename, engine = 'xlsxwriter') as writer:
        for i in range(len(sheet_list)):
            if type_flag[i] == 0:              # Generation incoming capacity
                args[i].to_excel(writer, sheet_name = sheet_list[i], index = False)
                synex_style(args[i], writer.book, writer.sheets[sheet_list[i]],False, language)

            elif type_flag[i] == 1:            # Generation works plan
                args[i].to_excel(writer, sheet_name = sheet_list[i])
                synex_style(args[i], writer.book, writer.sheets[sheet_list[i]], True, language)
                xlsx_chart({'type':'column', "subtype":"stacked"},
                    ["orange","red","black","gray","brown","green","lime","blue","purple","yellow","cyan","magenta"],
                    writer.book,sheet_list[i],writer.sheets,args[i].shape,language,type_flag)

            elif type_flag[i] == 2:
                args[i].to_excel(writer, sheet_name = sheet_list[i])
                synex_style(args[i], writer.book, writer.sheets[sheet_list[i]], True, language)
                xlsx_chart({'type':'column', "subtype":"stacked"},
                    ["orange","red","black","gray","brown","green","lime","blue","purple","yellow","cyan","magenta"],
                        writer.book,sheet_list[i],writer.sheets,args[i].shape,language,type_flag)


            printProgressBar(i+1, len(sheet_list), prefix = sheet_list[i], suffix = 'Complete', length = 50)           # 'Progress:'
    t0 = exectime("Exporting data to: "+filename, False, t0)

def work_plan_loader(mtermise, mgndse, mhidrose, mod_ter, mod_hid, mod_gn, dictionary, language, start_date):
    
    mtermise = mtermise.loc[mtermise["Data"] >= start_date,:]
    mhidrose = mhidrose.loc[mhidrose["Data"] >= start_date,:]
    mgndse = mgndse.loc[mgndse["Data"] >= start_date,:]

    mtermise.loc[:,"Substation"] = "-"
    mtermise.loc[:,"Developer"] = "-"
    mtermise.loc[:,"Technology"] = "-"
    for index, row in mtermise.iterrows():
        mtermise.loc[index,"Substation"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Barra"].all()
        mtermise.loc[index,"Developer"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Desarrollador"].all()
        if language:
            mtermise.loc[index,"Technology"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Tecnología Inglés"].all()
            mtermise.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Informe Inglés"].all()
        else:
            mtermise.loc[index,"Technology"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Tecnología Español"].all()
            mtermise.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Informe Español"].all()
        mtermise.loc[index,"Pot"] -= mod_ter.loc[mod_ter.index[mod_ter.index.get_loc(row["Data"]) - 1],row["Name"]]
        
    
    mhidrose.loc[:,"Substation"] = "-"
    mhidrose.loc[:,"Developer"] = "-"
    mhidrose.loc[:,"Technology"] = "-"
    for index, row in mhidrose.iterrows():
        if row["Name"][:3] == "HIB":
            if language:
                mhidrose.loc[index,"Technology"] = "Hybrid Storage"
            else:
                mhidrose.loc[index,"Technology"] = "Híbrido Almacenamiento"
        else:
            if language:
                mhidrose.loc[index,"Technology"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Tecnología Inglés"].all()
                mhidrose.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Informe Inglés"].all()
            else:
                mhidrose.loc[index,"Technology"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Tecnología Español"].all()
                mhidrose.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Informe Español"].all()
        mhidrose.loc[index,"Substation"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Barra"].all()
        mhidrose.loc[index,"Developer"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Desarrollador"].all()
        mhidrose.loc[index,"Pot"] -= mod_hid.loc[mod_hid.index[mod_hid.index.get_loc(row["Data"]) - 1],row["Name"]]
        
    
    mgndse.loc[:,"Substation"] = "-"
    mgndse.loc[:,"Developer"] = "-"
    mgndse.loc[:,"Technology"] = "-"
    for index, row in mgndse.iterrows():
        if row["Name"][:3] == "HIB":
            if language:
                mgndse.loc[index,"Technology"] = "Hybrid Solar PV"
            else:
                mgndse.loc[index,"Technology"] = "Híbrido Solar FV"
            mgndse.loc[index,"Developer"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Desarrollador"].all()
            mgndse.loc[index,"Pot"] -= mod_gn.loc[mod_gn.index[mod_gn.index.get_loc(row["Data"]) - 1],row["Name"]]
            if dictionary[dictionary["SDDP"] == row["Name"]]["Agrupacion en Informe?"].all() != "-":
                mgndse.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Agrupacion en Informe?"].all()
                mgndse.loc[index,"Substation"] = "-"

            else:
                if language:
                    mgndse.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Informe Inglés"].all()
                else:
                    mgndse.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Informe Español"].all()
                mgndse.loc[index,"Substation"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Barra"].all()
        elif row["Name"] == "EOL_Da�ical":
            mgndse.loc[index,"Substation"] = "EntreRios220"
            mgndse.loc[index,"Developer"] = "-"
            if language:
                mgndse.loc[index,"Technology"] = "Wind Farm"
                mgndse.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Informe Inglés"].all()
            else:
                mgndse.loc[index,"Technology"] = "Eólica"
                mgndse.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Informe Español"].all()
            mgndse.loc[index,"Pot"] -= mod_gn.loc[mod_gn.index[mod_gn.index.get_loc(row["Data"]) - 1],row["Name"]]
            
        else:
            if language:
                mgndse.loc[index,"Technology"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Tecnología Inglés"].all()
            else:
                mgndse.loc[index,"Technology"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Tecnología Español"].all()
            mgndse.loc[index,"Substation"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Barra"].all()
            mgndse.loc[index,"Developer"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Desarrollador"].all()
            mgndse.loc[index,"Pot"] -= mod_gn.loc[mod_gn.index[mod_gn.index.get_loc(row["Data"]) - 1],row["Name"]]
            if dictionary[dictionary["SDDP"] == row["Name"]]["Agrupacion en Informe?"].all() != "-":
                mgndse.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Agrupacion en Informe?"].all()
                mgndse.loc[index,"Substation"] = "-"

            else:
                if language:
                    mgndse.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Informe Inglés"].all()
                else:
                    mgndse.loc[index,"Name"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Informe Español"].all()
                mgndse.loc[index,"Substation"] = dictionary.loc[dictionary["SDDP"] == row["Name"],"Barra"].all()

    work_plan = mtermise.append(mgndse).append(mhidrose)
    work_plan = work_plan.sort_values(by = "Data")
    work_plan = work_plan.rename(columns = {"Name":"Project","Data":"COD","Pot":"MW"})
    work_plan = work_plan.iloc[:,[1,0,5,3,4,2]]
    work_plan = work_plan.groupby(["COD","Project","Technology","Substation","Developer"], as_index = False).sum()
    return work_plan

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

    if language:
        chart.set_x_axis({'name': 'Date', 'name_font' : {'size' : 9}, 'num_font' : {'arial narrow' : True}})
        if type_flag == 2:
            chart.set_y_axis({'name': 'Capacity [MW]', 'major_gridlines': {'visible': True}, 'name_font' : {'size' : 9},
                        'num_font' : {'arial narrow' : True},'num_format':'#,##0','min':0})
        else:
            chart.set_y_axis({'name': 'Capacity [MW]', 'major_gridlines': {'visible': True}, 'name_font' : {'size' : 9},
                        'num_font' : {'arial narrow' : True},'num_format':'#,##0'})
    else:
        chart.set_x_axis({'name': 'Fecha', 'name_font' : {'size' : 9}, 'num_font' : {'arial narrow' : True}})
        if type_flag == 2:
            chart.set_y_axis({'name': 'Capacidad [MW]', 'major_gridlines': {'visible': True}, 'name_font' : {'size' : 9}, 
                            'num_font' : {'arial narrow' : True},'num_format':'#,##0','min':0})
        else:
            chart.set_y_axis({'name': 'Capacidad [MW]', 'major_gridlines': {'visible': True}, 'name_font' : {'size' : 9}, 
                            'num_font' : {'arial narrow' : True},'num_format':'#,##0'})
    chart.set_legend({'position': 'bottom', 'font': {'size': 9, 'Arial Narrow': True}})
    chart.set_size({'width': 700, 'height': 450})
    chart_sheet[sheet_name].insert_chart(1, 15, chart)

def added_gen(w_plan, technologies):

    add_gen = pd.DataFrame(w_plan)
    add_gen.loc[:,"COD"] = add_gen.loc[:,"COD"].str[:4].astype('int')
    add_gen = add_gen.loc[:,["COD","Technology","MW"]]
    add_gen = add_gen.groupby(["COD","Technology"], as_index = False).sum()
    add_gen = add_gen.pivot(index = "COD", columns = "Technology", values = "MW")
    add_gen = add_gen.fillna(0)
    for tec in technologies:
        if tec not in add_gen.columns:
            add_gen[tec] = 0
    add_gen = add_gen.loc[:,technologies]
    add_gen["Total"] = add_gen.sum(axis = 1)
    return add_gen


####################################################################################################
########################################     MAIN CODE      ########################################
####################################################################################################

### Init user's values

start_date = 2022
date = str(start_date)+"-01"

technologies  = ["Biomass","Coal","Diesel","Gas","Geothermal","Hybrid Storage","Hybrid Solar PV","Hydro","Solar CSP","Solar PV","Wind Farm","Battery"]
technologies_esp  = ["Biomasa","Carbón","Diesel","Gas","Geotérmica","Híbrido Almacenamiento","Híbrido Solar FV","Hidro","Solar CSP","Solar FV","Eólica","Batería"]

file_root = "C:/aCasosSynex/20_93_1513_09032022 - EBS LP - CMB - BASE comb Junio"
output_file_name = "/WorkPlan.xlsx"
english = True

init_capacity = {"Battery":0,
                "Biomass":253,
                "Coal":4486,
                "Diesel":3306,
                "Gas":5031,
                "Geothermal":78,
                "Hybrid Storage":0,
                "Hybrid Solar PV":0,
                "Hydro":7113,
                "Solar CSP":110,
                "Solar PV":6198,
                "Wind Farm":3536}


# Read data

ctermise = readfile(file_root + "/ctermise.csv",0, [])
ctermise = ctermise[ctermise["...1"] != 0]
ctermise = ctermise.iloc[:,[1,3,4,11]]

chidrose = readfile(file_root + "/chidrose.csv",0, [])
chidrose = chidrose.iloc[:,[1,6,7]]

cgndse = readfile(file_root + "/cgndse.csv",0, [])
cgndse = cgndse.iloc[:,[1,3,5]]

ctermise = ctermise.rename(columns = {"...Nombre...":"Nombre",".PotIns":"Pot", "Tipo":"Type"})
cgndse = cgndse.rename(columns = {"Name":"Nombre","PotIns":"Pot"})
chidrose = chidrose.rename(columns = {"...Nombre...":"Nombre","....Pot":"Pot", "Tipo":"Type"})
ctermise[ctermise["Type"] == 0]
chidrose[chidrose["Type"] == 0]
cgndse[cgndse["Type"] == 0]

dbus = readfile(file_root + "/dbus.csv", 1, [])
dbus = dbus.iloc[:,[1,6]]
dbus = dbus.fillna("-")

dictionary = readfile(file_root + "/Power_Plant_Dictionary.xlsx",0,[])
dictionary = dictionary.fillna("-")

mtermise = readfile(file_root + "/mtermise.csv", 0, [])
mtermise = mtermise.iloc[:,[1,2,5]]
mtermise = mtermise.rename(columns= {"Nombre":"Name","GerMax":"Pot"})
mtermise = mtermise.sort_values(by = "Data")

mgndse = readfile(file_root + "/mgndse.csv", 0, [])
mgndse = mgndse.iloc[:,[1,2,4]]
mgndse = mgndse.rename(columns= {"Nombre":"Name",".PotIns":"Pot"})
mgndse = mgndse.sort_values(by = "Data")

mhidrose = readfile(file_root + "/mhidrose.csv", 0, [])
mhidrose = mhidrose.iloc[:,[1,2,9]]
mhidrose = mhidrose.rename(columns= {"Nombre":"Name"})
mhidrose = mhidrose.sort_values(by = "Data")

mtermise.loc[:,"Data"] = mtermise.loc[:,"Data"].str[:7]
mhidrose.loc[:,"Data"] = mhidrose.loc[:,"Data"].str[:7]
mgndse.loc[:,"Data"] = mgndse.loc[:,"Data"].str[:7]


### Generation park updates

mod_termi = data_expand(ctermise, mtermise, start_date)
mod_hidro = data_expand(chidrose, mhidrose, start_date)
mod_gnd = data_expand(cgndse, mgndse, start_date)

# Generation park Works Plan

work_plan = work_plan_loader(mtermise, mgndse,mhidrose, mod_termi, mod_hidro, mod_gnd,dictionary,english,date)

coal_power = -work_plan.loc[work_plan["Technology"] == "Coal","MW"].sum()
init_capacity["Coal"] = coal_power
print(work_plan)

add_gen = added_gen(work_plan, technologies)

gen_expantion = pd.DataFrame(0,columns = technologies, index = [i for i in range(start_date-1,add_gen.index[-1] + 1)])
for col in gen_expantion.columns:
    for index in gen_expantion.index:
        if index == start_date - 1:
            gen_expantion.loc[index,col] = init_capacity[col]
        else:
            if col not in add_gen.columns:
                gen_expantion.loc[index,col] = gen_expantion.loc[index - 1,col]
            else:
                gen_expantion.loc[index,col] = gen_expantion.loc[index - 1,col] + add_gen.loc[index,col]

gen_expantion["Total"] = gen_expantion.sum(axis = 1)
print(gen_expantion)

work_plan = work_plan_loader(mtermise, mgndse,mhidrose, mod_termi, mod_hidro, mod_gnd,dictionary,english,date)

if not english:
        work_plan = work_plan.rename(columns = {"Project":"Proyecto"})
        add_gen = add_gen.rename(columns = {technologies[i]:technologies_esp[i] for i in range(len(technologies))})


print(work_plan)
export_excel(["WorksPlan","AddedCap","GenCap"],[0,1,2], file_root, output_file_name, english, work_plan, add_gen, gen_expantion)