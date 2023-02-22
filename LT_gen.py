import pandas as pd
from time import process_time
import matplotlib.pyplot as plt
import numpy as np
import os


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

def export_excel(sheet_list, type_flag, file_root, filename, language,colors,*args):

    t0 = exectime("Exporting data to: " + filename, True, 0)
    printProgressBar(0, len(sheet_list), prefix = 'Progress:', suffix = 'Complete', length = 50)

    with pd.ExcelWriter(os.path.join(file_root,filename), engine = 'xlsxwriter') as writer:
        for i in range(len(sheet_list)):
            if type_flag[i] == 0:              # Generation incoming capacity
                args[i].to_excel(writer, sheet_name = sheet_list[i], index = False)
                synex_style(args[i], writer.book, writer.sheets[sheet_list[i]],False, language)

            elif type_flag[i] == 1:            # Generation works plan
                args[i].to_excel(writer, sheet_name = sheet_list[i])
                synex_style(args[i], writer.book, writer.sheets[sheet_list[i]], True, language)
                xlsx_chart({'type':'column', "subtype":"stacked"}, colors,
                    writer.book,sheet_list[i],writer.sheets,args[i].shape,language,type_flag)

            elif type_flag[i] == 2:
                args[i].to_excel(writer, sheet_name = sheet_list[i])
                synex_style(args[i], writer.book, writer.sheets[sheet_list[i]], True, language)
                xlsx_chart({'type':'area', "subtype":"stacked"}, colors,
                        writer.book,sheet_list[i],writer.sheets,args[i].shape,language,type_flag)


            printProgressBar(i+1, len(sheet_list), prefix = sheet_list[i], suffix = 'Complete', length = 50)           # 'Progress:'
    t0 = exectime("Exporting data to: "+filename, False, t0)

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


#####     User inputs

file_root = input("Please add file path's: ",)
output_file_name = "GenerationExpantion.xlsx"

start_date_output = (2023,2)                                                                                # Start date for output data
ST_flag = False
short_term_hydro_cond = [28,28]                                                                             # Hydro condition for the shorttest time
short_term_date = (2023,4)                                                                                  # Short term hydro date
short_mid_term_hydro_cond = [28,28]                                                                         # hydro condition for the short time
short_mid_term_date = (2023,4)                                                                              # Short-mid term hydro date
English = True

previous_gen_flag = True

Hydro_limit_condition = [3,29]                                                                              # 3rd hydro as the wettest limit, 29th as the driest limit
init_hydro_date = 1988
last_hydro_date = 2018

HW = [i for i in range(1,94,3)]                                                                             # High - wind series
MW = [i for i in range(2,94,3)]                                                                             # Middle - wind series
LW = [i for i in range(3,94,3)]                                                                             # Low - wind series
FW = [i for i in range(1,94)]

wind_cond = LW

technologies  = ["Coal","Hydro","Gas","Solar PV","Wind Farm","Hybrid Storage","Hybrid Solar PV","Solar CSP","Battery","Biomass","Diesel","Geothermal","Fuel Oil"]
technologies_esp  = ["Carbón","Hidro","Gas","Solar FV","Eólica","Híbrido Almacenamiento","Híbrido Solar FV","Solar CSP","Batería","Biomasa","Diesel","Geotérmica","Fuel Oil"]

tech_colors = {"Biomass":"orange",
            "Coal":"red",
            "Diesel":"black",
            "Gas":"gray",
            "Geothermal":"brown",
            "Hybrid Storage":"green",
            "Hybrid Solar PV":"lime",
            "Hydro":"blue",
            "Solar CSP":"purple",
            "Solar PV":"yellow",
            "Wind Farm":"cyan",
            "Battery":"magenta",
            "Fuel Oil":"silver"}

colors = [tech_colors[i] for i in technologies]


## Actualizado a Ene23(considera Ene)

previous_gen = readfile("CEN_RealGen.xlsx",3,[18,19])
previous_gen = previous_gen.dropna()
previous_gen = previous_gen.set_index("Technology")
previous_gen

#####     Read data
print("Reading generations")
gergnd = readfile(os.path.join(file_root,"gergnd.csv"),3,[])
gerhid = readfile(os.path.join(file_root,"gerhid.csv"),3,[])
gerter = readfile(os.path.join(file_root,"gerter.csv"),3,[])
gergnd = gergnd.rename(columns= {col: col.replace(" ","") for col in gergnd.columns})
gerhid = gerhid.rename(columns= {col: col.replace(" ","") for col in gerhid.columns})
gerter = gerter.rename(columns= {col: col.replace(" ","") for col in gerter.columns})

dictionary = readfile("Power_Plant_Dictionary.xlsx",0,[])
dictionary = dictionary.fillna("-")
dictionary

date = init_date(os.path.join(file_root, "duraci.csv"))
duraci = readfile(os.path.join(file_root,"duraci.csv"),3,[0,1,2,3])
duraci = month_year_index(duraci, date, start_date_output)                                                  # duraci data with year-month values
duraci = duraci.rename(columns = {duraci.columns[0]:"SEN"})
duraci = duraci.groupby(level = ["Year","Month"]).sum()

gergnd = month_year_index(gergnd, date, start_date_output)
gerter = month_year_index(gerter, date, start_date_output)
gerhid = month_year_index(gerhid, date, start_date_output)

print("Reading hydro files")
hidroseq = readfile("Orden_hidrologias.xlsx", 0, [])                                                        # read hydro seq file
hidroseq = hidroseq[(hidroseq["Year"] >= init_hydro_date) & (hidroseq["Year"] <= last_hydro_date)]          # take just hydro sequence used
hidroseq = hidroseq.sort_values(by = "Total", ascending = False)                                            # sort hydrologies frow wet to dry
hidroseq["index"] = [i for i in range(1,32)]                                                                # index values by preview order
hidroseq = hidroseq.set_index("index")                                                                      # Index it's setted

hydro_condition = case_hydro_condition(init_hydro_date, last_hydro_date, duraci.index, hidroseq)      # Get hydro condition for every stage

print("Gen by year")
gergnd = gergnd.groupby(level = ["Year","Month","Seq."]).sum()
gerter = gerter.groupby(level = ["Year","Month","Seq."]).sum()
gerhid = gerhid.groupby(level = ["Year","Month","Seq."]).sum()


if ST_flag:
    for index, row in gergnd.iterrows():
        if (index[0],index[1]) <= short_term_date:
            hydro_cond_index = (index[0],index[1],(i for i in hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] >= short_term_hydro_cond[0]].index.intersection(hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] <= short_term_hydro_cond[1]].index) if i in wind_cond))
            gergnd.loc[index,:] = gergnd.loc[hydro_cond_index,:].mean()
            gerter.loc[index,:] = gerter.loc[hydro_cond_index,:].mean()
            gerhid.loc[index,:] = gerhid.loc[hydro_cond_index,:].mean()
        else:
            break

gnd_year = gergnd.groupby(level = ["Year","Month"]).mean().groupby(level = ["Year"]).sum().stack()
ter_year = gerter.groupby(level = ["Year","Month"]).mean().groupby(level = ["Year"]).sum().stack()
hid_year = gerhid.groupby(level = ["Year","Month"]).mean().groupby(level = ["Year"]).sum().stack()
gnd_year = gnd_year.reset_index()
gnd_year = gnd_year.rename(columns = {0:"Values","level_1":"Plant"})
gnd_year["Technology"] = None

hid_year = hid_year.reset_index()
hid_year = hid_year.rename(columns = {0:"Values","level_1":"Plant"})
hid_year["Technology"] = None

ter_year = ter_year.reset_index()
ter_year = ter_year.rename(columns = {0:"Values","level_1":"Plant"})
ter_year["Technology"] = None


for plant in gnd_year["Plant"].unique():
    gnd_year.loc[gnd_year["Plant"] == plant, "Technology"] = dictionary.loc[dictionary["SDDP"] == plant, "Tecnología Inglés"].all()
    if gnd_year.loc[gnd_year["Plant"] == plant, "Technology"].all() == "Hybrid":
        gnd_year.loc[gnd_year["Plant"] == plant, "Technology"] = "Hybrid Solar PV"
    elif gnd_year.loc[gnd_year["Plant"] == plant, "Technology"].all() == True:
        print(plant, gnd_year.loc[gnd_year["Plant"] == plant, "Technology"].all())

for plant in hid_year["Plant"].unique():
    hid_year.loc[hid_year["Plant"] == plant, "Technology"] = dictionary.loc[dictionary["SDDP"] == plant, "Tecnología Inglés"].all()
    if hid_year.loc[hid_year["Plant"] == plant, "Technology"].all() == "Hybrid":
        hid_year.loc[hid_year["Plant"] == plant, "Technology"] = "Hybrid Storage"
    elif hid_year.loc[hid_year["Plant"] == plant, "Technology"].all() == True:
        print(plant, gnd_year.loc[hid_year["Plant"] == plant, "Technology"].all())

for plant in ter_year["Plant"].unique():
    ter_year.loc[ter_year["Plant"] == plant, "Technology"] = dictionary.loc[dictionary["SDDP"] == plant, "Tecnología Inglés"].all()
    if ter_year.loc[ter_year["Plant"] == plant, "Technology"].all() == True:
        print(plant, ter_year.loc[ter_year["Plant"] == plant, "Technology"].all())


generation = pd.DataFrame(pd.concat([gnd_year,ter_year,hid_year]))
generation = generation.groupby(["Year","Technology"]).sum()
generation = generation.unstack(level = 1)
generation.columns = generation.columns.droplevel(0)
if previous_gen_flag:
    for col in generation.columns:
        if col in previous_gen.index:
            generation.loc[generation.index[0],col] += previous_gen.loc[col,:].values
    generation = generation.loc[:,technologies]
    generation["Total"] = generation.sum(axis = 1)

print(generation)

export_excel(["Generation"],[2], file_root, output_file_name, English, colors, generation)