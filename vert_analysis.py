import pandas as pd
import os
import numpy as np
from time import process_time

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

def init_date(file):
    df = pd.DataFrame(pd.read_csv(file,nrows=1))
    df = df.rename(columns = {i:i.replace(" ","") for i in df.columns})
    return [int(df.columns[-2]),int(df.columns[-1])]

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

def export_excel(sheet_list, type_flag, file_root, filename, language,*args):

    t0 = exectime("Exporting data to: " + filename, True, 0)
    printProgressBar(0, len(sheet_list), prefix = 'Progress:', suffix = 'Complete', length = 50)

    with pd.ExcelWriter(os.path.join(file_root,filename), engine = 'xlsxwriter') as writer:
        for i in range(len(sheet_list)):
            if type_flag[i] == 0:              # Generation incoming capacity
                args[i].to_excel(writer, sheet_name = sheet_list[i], index = True)
                # synex_style(args[i], writer.book, writer.sheets[sheet_list[i]],False, language)

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
                    ["orange","red","black","gray","brown","silver","green","lime","blue","purple","yellow","cyan","magenta"],
                        writer.book,sheet_list[i],writer.sheets,args[i].shape,language,type_flag)


            printProgressBar(i+1, len(sheet_list), prefix = sheet_list[i], suffix = 'Complete', length = 50)           # 'Progress:'
    t0 = exectime("Exporting data to: "+filename, False, t0)

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

def weighted_mean(data, time, Serie_flag):

    t0 = exectime("cmgbus.csv", True, 0)
    if Serie_flag:
        df_seq_mean = data.groupby(level = [0,1,3]).mean()                                            # data with year-month-block mean
        df_out = df_seq_mean.mul(time.loc[:,"SEN"], axis = 0).div(time.loc[:,"Total_month"], axis = 0)
        df_out = df_out.groupby(level = [0,1]).sum()
        df_out2 = df_out.groupby(level = [0]).mean()
        t0 = exectime("cmgbus.csv", False, t0)
        return df_out, df_out2
    else:
        df_out = pd.DataFrame(0,columns = data.columns,index = data.index)
        printProgressBar(0, len(df_out.index.get_level_values("Seq.").unique()), prefix = 'Progress:', suffix = 'Complete', length = 50)
        for serie in df_out.index.get_level_values("Seq.").unique():
            df_out.loc[df_out.index.get_level_values("Seq.").isin([serie]),:] = data.loc[data.index.get_level_values("Seq.").isin([serie]),:].mul(time.loc[:,"SEN"].values, axis = 0).div(time.loc[:,"Total_month"].values, axis = 0)
            printProgressBar(serie, len(df_out.index.get_level_values("Seq.").unique()), prefix = 'Progress:', suffix = 'Complete', length = 50)
        df_out = df_out.groupby(level = [0,1,2]).sum()
        df_out2 = df_out.groupby(level = [0,2]).mean()
        t0 = exectime("cmgbus.csv", False, t0)
        return df_out, df_out2

def case_hydro_condition(init_hydro_date, last_hydro_date, data_index, hydro_seq):
    
    hydro_condition = pd.DataFrame(index = data_index, columns = [i for i in range(1,94)])

    first_row_year = init_hydro_date
    last_row_year = init_hydro_date

    for c in hydro_condition.columns:
        for r in hydro_condition.index:
            if c == hydro_condition.columns[0] and  r == hydro_condition.index[0]:              # First cell (0,0)
                if r[1] > 4:
                    first_row_year += 1
                hydro_condition.loc[r,c] = first_row_year
                last_row_year = hydro_condition.loc[r,c]
                first_row_year += 1
            elif c != hydro_condition.columns[0] and  r == hydro_condition.index[0]:            # anything but first col (r,0)
                if first_row_year > last_hydro_date:
                    first_row_year = init_hydro_date
                hydro_condition.loc[r,c] = first_row_year
                last_row_year = hydro_condition.loc[r,c]
                first_row_year += 1
            else:                                                                               # any other cell (r,c)
                if r[1] == 5:
                    if last_row_year == last_hydro_date:
                        last_row_year = init_hydro_date
                    else:
                        last_row_year += 1
                hydro_condition.loc[r,c] = last_row_year
                last_row_year = hydro_condition.loc[r,c]
    for c in hydro_condition.columns:
        for r in hydro_condition.index:
            hydro_condition.loc[r,c] = hydro_seq[hydro_seq["Year"] == hydro_condition.loc[r,c]].index[0]
    return hydro_condition

def gen_bus_adder(dbus, column_names, df_index, *args):
    df = pd.DataFrame(0,columns = column_names, index = df_index)
    for data in args:
        for col in data:
            df.loc[:,dbus.loc[dbus["Gen. name"] == col,"Name"].all()] += data.loc[:,col]
    return df


file_root = input("Please add file root: ",)

output_file_name = "vert_generation.xlsx"

English = True

bus_list = ["CPinto220","Cardones220","NVaCardon500","Maitenci220","PAzucar220","Polpaico220","Crucero220","DAlmagro220"]

print("Reading files...")
dbus = readfile(os.path.join(file_root,"dbus.csv"),1,[1,6])
dbus = dbus.fillna("-")
dbus["Gen. name"] = dbus["Gen. name"].str.replace(" ","")

dictionary = readfile("Power_Plant_Dictionary.xlsx",0,[])
dictionary = dictionary.fillna("-")
dictionary["Tecnología Inglés"].unique()

start_date_output = (2023,1)
date = init_date(os.path.join(file_root,"duraci.csv"))
duraci = readfile(os.path.join(file_root,"duraci.csv"), 3, [0,1,2,3])   
vergnd = readfile(os.path.join(file_root,"vergnd.csv"), 3, [])
gergnd = readfile(os.path.join(file_root,"gergnd.csv"), 3, [])

gergnd = gergnd.rename(columns= {col: col.replace(" ","") for col in gergnd.columns})
vergnd = vergnd.rename(columns= {col: col.replace(" ","") for col in vergnd.columns})

vergnd = month_year_index(vergnd, date, start_date_output)
gergnd = month_year_index(gergnd, date, start_date_output)

vergnd = vergnd.groupby(level = [0,1,3]).mean().groupby(level = [0,1]).sum()
gergnd = gergnd.groupby(level = [0,1,3]).mean().groupby(level = [0,1]).sum()

# Vert by bus
print("Calculating vert by bus...")
gen_by_bus = gen_bus_adder(dbus,dbus["Name"].unique(),gergnd.index,gergnd)
vert_by_bus = gen_bus_adder(dbus,dbus["Name"].unique(),vergnd.index,vergnd)
gen_by_bus = gen_by_bus.loc[:,bus_list]
vert_by_bus = vert_by_bus.loc[:,bus_list]

vert_perc_by_bus = vert_by_bus/(gen_by_bus + vert_by_bus)
vert_by_bus_annual = vert_by_bus.groupby(level = [0]).sum()/(gen_by_bus.groupby(level = [0]).sum() + vert_by_bus.groupby(level = [0]).sum())

# Vert by technology
print("Calculating vert by technology...")
rw_gen = pd.DataFrame(0,columns=["Solar PV","Wind Farm","Solar CSP"],index = gergnd.index)
vert_gen = pd.DataFrame(0,columns=["Solar PV","Wind Farm","Solar CSP"],index = gergnd.index)

for col in gergnd.columns:
    if dictionary.loc[dictionary["SDDP"] == col,"Tecnología Inglés"].all() == "Solar PV" or dictionary.loc[dictionary["SDDP"] == col,"Tecnología Inglés"].all() == "Hybrid":
        rw_gen.loc[:,"Solar PV"] += gergnd.loc[:,col]
        vert_gen.loc[:,"Solar PV"] += vergnd.loc[:,col]
    elif dictionary.loc[dictionary["SDDP"] == col,"Tecnología Inglés"].all() == "Wind Farm":
        rw_gen.loc[:,"Wind Farm"] += gergnd.loc[:,col]
        vert_gen.loc[:,"Wind Farm"] += vergnd.loc[:,col]
    elif dictionary.loc[dictionary["SDDP"] == col,"Tecnología Inglés"].all() == "Solar CSP":
        rw_gen.loc[:,"Solar CSP"] += gergnd.loc[:,col]
        vert_gen.loc[:,"Solar CSP"] += vergnd.loc[:,col]
    else:
        print(dictionary.loc[dictionary["SDDP"] == col,"Tecnología Inglés"].all())

rw_gen["Total"] = rw_gen.sum(axis = 1)
vert_gen["Total"] = vert_gen.sum(axis = 1)

vert_percentage = vert_gen/(rw_gen + vert_gen)

vert_percentage_annual = vert_gen.groupby(level = [0]).sum()/(rw_gen.groupby(level = [0]).sum() + vert_gen.groupby(level = [0]).sum())

export_excel(["Vert_percentage","vert_perc_annual","Vert_by_bus","Vert_by_bus_annual"],[0,0,0,0], file_root, output_file_name, English, 
            vert_percentage,vert_perc_by_bus,vert_percentage_annual,vert_by_bus_annual)
