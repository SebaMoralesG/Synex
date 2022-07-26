import pandas as pd
from time import process_time
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.collections import PolyCollection



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

def export_excel(bus_list, data, data_aux1, data_aux2, type_flag, file_root, filename, language):

    t0 = exectime("Exporting data to: " + filename, True, 0)
    printProgressBar(0, len(bus_list), prefix = 'Progress:', suffix = 'Complete', length = 50)

    with pd.ExcelWriter(file_root + filename, engine = 'xlsxwriter') as writer:
        for i in range(len(bus_list)):
            if type_flag == 0:              # Year-Month average
                data_out = plot_year_month_data(data,bus_list[i], language)
                data_out.to_excel(writer, sheet_name = bus_list[i])
                synex_style(data_out, bus_list[i], writer.book, writer.sheets[bus_list[i]], type_flag)

                (max_row, max_col) = data_out.shape
                chart = writer.book.add_chart({'type':'line'})
                line_color = ['blue', 'red']
                for col in range(1, max_col + 1):
                    chart.add_series({
                        'name':       [bus_list[i], 0, col],
                        'categories': [bus_list[i], 1, 0,   max_row, 0],
                        'values':     [bus_list[i], 1, col, max_row, col],
                        'line':       {'color': line_color[col - 1]}
                    })
                if language:
                    chart.set_x_axis({'name': 'Date', 'name_font' : {'size' : 9}, 'num_font' : {'arial narrow' : True}})
                else:
                    chart.set_x_axis({'name': 'Fecha', 'name_font' : {'size' : 9}, 'num_font' : {'arial narrow' : True}})
                chart.set_y_axis({'min': 0, 'name': 'CMg [USD/MWh]', 'major_gridlines': {'visible': True}, 'name_font' : {'size' : 9}, 'num_font' : {'arial narrow' : True}})
                chart.set_legend({'position': 'bottom', 'font': {'size': 9, 'Arial Narrow': True}})
                chart.set_size({'width': 700, 'height': 450})
                writer.sheets[bus_list[i]].insert_chart(1, 6, chart)

            elif type_flag == 1:            # Hour-Month average
                data[bus_list[i]].unstack(level = 1).to_excel(writer, sheet_name = bus_list[i])
                synex_style(data[bus_list[i]].unstack(level = 1), bus_list[i], writer.book, writer.sheets[bus_list[i]], type_flag)
                plot_CMg_3d(data, bus_list[i], language, 100, file_root)
                writer.sheets[bus_list[i]].insert_image('AA1', file_root + bus_list[i] + ".png")

            elif type_flag == 2:            # Hydro condition
                data_out = plot_hydro_cmg(data, data_aux1, data_aux2, bus_list[i], language)
                data_out.to_excel(writer, sheet_name = bus_list[i])
                synex_style(data_out, bus_list[i], writer.book, writer.sheets[bus_list[i]], type_flag)

                (max_row, max_col) = data_out.shape
                chart = writer.book.add_chart({'type':'line'})
                line_color = ["green","blue","red"]
                for col in range(1, max_col):
                    chart.add_series({
                        'name':       [bus_list[i], 0, col],
                        'categories': [bus_list[i], 1, 0,   max_row, 0],
                        'values':     [bus_list[i], 1, col, max_row, col],
                        'line':       {'color': line_color[col - 1]}
                    })
                if language:
                    chart.set_x_axis({'name': 'Date', 'name_font' : {'size' : 9}, 'num_font' : {'arial narrow' : True}})
                else:
                    chart.set_x_axis({'name': 'Fecha', 'name_font' : {'size' : 9}, 'num_font' : {'arial narrow' : True}})
                chart.set_y_axis({'min': 0, 'name': 'CMg [USD/MWh]', 'major_gridlines': {'visible': True}, 'name_font' : {'size' : 9}, 'num_font' : {'arial narrow' : True}})
                chart.set_legend({'position': 'bottom', 'font': {'size': 9, 'Arial Narrow': True}})
                chart.set_size({'width': 700, 'height': 450})
                writer.sheets[bus_list[i]].insert_chart(1, 6, chart)

            elif type_flag == 3:            # Wind condition
                data_out = wind_cmg(data, data_aux1, data_aux2, bus_list[i], language)
                data_out.to_excel(writer, sheet_name = bus_list[i])
                synex_style(data_out, bus_list[i], writer.book, writer.sheets[bus_list[i]], type_flag)
                
                (max_row, max_col) = data_out.shape
                chart = writer.book.add_chart({'type':'line'})
                line_color = ["green","blue","red"]
                for col in range(1, max_col):
                    chart.add_series({
                        'name':       [bus_list[i], 0, col],
                        'categories': [bus_list[i], 1, 0,   max_row, 0],
                        'values':     [bus_list[i], 1, col, max_row, col],
                        'line':       {'color': line_color[col - 1]}
                    })
                if language:
                    chart.set_x_axis({'name': 'Date', 'name_font' : {'size' : 9}, 'num_font' : {'arial narrow' : True}})
                else:
                    chart.set_x_axis({'name': 'Fecha', 'name_font' : {'size' : 9}, 'num_font' : {'arial narrow' : True}})
                chart.set_y_axis({'min': 0, 'name': 'CMg [USD/MWh]', 'major_gridlines': {'visible': True}, 'name_font' : {'size' : 9}, 'num_font' : {'arial narrow' : True}})
                chart.set_legend({'position': 'bottom', 'font': {'size': 9, 'Arial Narrow': True}})
                chart.set_size({'width': 700, 'height': 450})
                writer.sheets[bus_list[i]].insert_chart(1, 6, chart)

            elif type_flag == 4:            # Hour-Month average
                data[bus_list[i]].unstack(level = 1).to_excel(writer, sheet_name = bus_list[i])
                synex_style(data[bus_list[i]].unstack(level = 1), bus_list[i], writer.book, writer.sheets[bus_list[i]], type_flag)

            printProgressBar(i+1, len(bus_list), prefix = bus_list[i], suffix = 'Complete', length = 50)           # 'Progress:'
    t0 = exectime("Exporting data to: "+filename, False, t0)

def exportexcelunstack(listnames, file_root, file, level_unstacked, data, type_data_flag, language,cmg_lim):

    t0 = exectime("Exporting data to: " + file, True, 0)
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

def plot_hydro_cmg(wet_data,mid_data,dry_data,node, language):
    
    if node not in dry_data.columns:
        print("Error: Node doesn't exist, please check node name")
        return None
    else:
        if language:
            df_out = pd.DataFrame(index = dry_data.index, columns = [ "Wet", "Medium", "Dry","Average"])
            df_out["Dry"] = dry_data[node]
            df_out["Medium"] = mid_data[node]
            df_out["Wet"] = wet_data[node]
            df_out["Average"] = (3*dry_data[node] + 25*mid_data[node] + 3*wet_data[node])/31
        else:
            df_out = pd.DataFrame(index = dry_data.index, columns = [ "Húmedo", "Medio", "Seco","Promedio"])
            df_out["Seco"] = dry_data[node]
            df_out["Medio"] = mid_data[node]
            df_out["Húmedo"] = wet_data[node]
            df_out["Promedio"] = (3*dry_data[node] + 25*mid_data[node] + 3*wet_data[node])/31

        df_out = df_out.reset_index()
        df_out["Month"] = df_out["Month"].astype("str") + "-" + df_out["Year"].astype("str")
        df_out = df_out.drop(columns = ["Year"])
        if language:
            df_out = df_out.rename(columns = {"Month":"Date"})
            df_out = df_out.set_index("Date")
        else:
            df_out = df_out.rename(columns = {"Month":"Fecha"})
            df_out = df_out.set_index("Fecha")
    return df_out

def wind_cmg(HW_data,MW_data,LW_data,node, language):

    if node not in HW_data.columns:
        print("Error: Node doesn't exist, please check node name")
        return None
    else:
        if language:
            df_out = pd.DataFrame(index = HW_data.index, columns = [ "High Wind", "Medium Wind", "Low Wind","Average"])
            df_out["High Wind"] = HW_data[node]
            df_out["Medium Wind"] = MW_data[node]
            df_out["Low Wind"] = LW_data[node]
            df_out["Average"] = (3*HW_data[node] + 25*MW_data[node] + 3*LW_data[node])/31
        else:
            df_out = pd.DataFrame(index = HW_data.index, columns = [ "Viento Alto", "Viento Medio", "Viento Bajo","Promedio"])
            df_out["Viento Alto"] = HW_data[node]
            df_out["Viento Medio"] = MW_data[node]
            df_out["Viento Bajo"] = LW_data[node]
            df_out["Promedio"] = (3*HW_data[node] + 25*MW_data[node] + 3*LW_data[node])/31

        df_out = df_out.reset_index()
        df_out["Month"] = df_out["Month"].astype("str") + "-" + df_out["Year"].astype("str")
        df_out = df_out.drop(columns = ["Year"])
        if language:
            df_out = df_out.rename(columns = {"Month":"Date"})
            df_out = df_out.set_index("Date")
        else:
            df_out = df_out.rename(columns = {"Month":"Fecha"})
            df_out = df_out.set_index("Fecha")
    return df_out

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

def daily_format(data, working_blocks, Nonworking_blocks, time_data, divide_flag, hydro_condition, ST_flag, ST_date1, ST_serie1, ST_date2, ST_serie2):

    data_serie_mean = data.groupby(level = [0,1,3]).mean()

    if ST_flag:
        for index, row in hydro_condition.iterrows():
                if (index[0],index[1]) <= ST_date1:
                    data_serie_mean.loc[(index[0],index[1]),:] = data.loc[(index[0],index[1],1,1):(index[0],index[1],93,20),:].loc[data.loc[(index[0],index[1],1,1):(index[0],index[1],93,20),:].index.get_level_values("Seq.").isin(hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] >= ST_serie1[0]].index) & data.loc[(index[0],index[1],1,1):(index[0],index[1],93,20),:].index.get_level_values("Seq.").isin(hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] <= ST_serie1[1]].index),:].groupby(level = [0,1,3]).mean()
                elif (index[0],index[1]) <= ST_date2:
                    data_serie_mean.loc[(index[0],index[1]),:] = data.loc[(index[0],index[1],1,1):(index[0],index[1],93,20),:].loc[data.loc[(index[0],index[1],1,1):(index[0],index[1],93,20),:].index.get_level_values("Seq.").isin(hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] >= ST_serie2[0]].index) & data.loc[(index[0],index[1],1,1):(index[0],index[1],93,20),:].index.get_level_values("Seq.").isin(hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] <= ST_serie2[1]].index),:].groupby(level = [0,1,3]).mean()
                else:
                    break
    hours = [24,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23]
    if divide_flag:
        data_serie_mean = data_serie_mean.div(time_data.values)
    data_workday = data_serie_mean[data_serie_mean.index.get_level_values("Blck").isin(working_blocks)]
    data_workday = pd.DataFrame(np.repeat( data_workday.values, 2, axis=0), columns = data_workday.columns, index = np.repeat(data_workday.index,2))
    data_workday = data_workday.reset_index(level = 2)

    for index, row in data_workday.iterrows():
        data_workday.loc[index,"Blck"] = hours
    data_workday = data_workday.reset_index()
    data_workday = data_workday.set_index(["Year","Month","Blck"])
    data_workday = data_workday.sort_index()

    # [1,2,3,13,14,15,16,17,18,19,20,12]
    data_nonworkday = data_serie_mean[data_serie_mean.index.get_level_values("Blck").isin(Nonworking_blocks)]
    data_nonworkday = pd.DataFrame(np.repeat( data_nonworkday.values, 2, axis=0), columns = data_nonworkday.columns, index = np.repeat(data_nonworkday.index,2))
    data_nonworkday = data_nonworkday.reset_index(level = 2)

    hours = [24,1,2,3,4,5,22,23,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21]

    for index, row in data_nonworkday.iterrows():
        data_nonworkday.loc[index,"Blck"] = hours
    data_nonworkday = data_nonworkday.reset_index()
    data_nonworkday = data_nonworkday.set_index(["Year","Month","Blck"])
    data_nonworkday = data_nonworkday.sort_index()

    time_data = time_data.loc[time_data.index.get_level_values("Blck").isin([4,13])].unstack(level = 3).groupby(level = [1]).mean()

    data_meanday = data_workday
    for month in time_data.index:
        data_meanday.loc[data_meanday.index.get_level_values("Month").isin([month]),:] = (data_meanday.loc[data_meanday.index.get_level_values("Month").isin([month]),:]*time_data.loc[month,4] + data_nonworkday.loc[data_nonworkday.index.get_level_values("Month").isin([month]),:]*time_data.loc[month,13])/(time_data.loc[month,13] + time_data.loc[month,4])

    data_meanday_y = data_meanday.groupby(level = [0,2]).mean()
    data_workday_y = data_workday.groupby(level = [0,2]).mean()
    data_nonworkday_y = data_nonworkday.groupby(level = [0,2]).mean()

    data_meanday = data_meanday.reset_index()
    data_workday = data_workday.reset_index()
    data_nonworkday = data_nonworkday.reset_index()
    for index, row in data_meanday.iterrows():
        if data_meanday.loc[index,"Month"] < 10:
            data_meanday.loc[index,"Year"] = str(data_meanday.loc[index,"Year"]) + "-" + "0"+str(data_meanday.loc[index,"Month"])
            data_workday.loc[index,"Year"] = str(data_workday.loc[index,"Year"]) + "-" + "0"+str(data_workday.loc[index,"Month"])
            data_nonworkday.loc[index,"Year"] = str(data_nonworkday.loc[index,"Year"]) + "-" + "0"+str(data_nonworkday.loc[index,"Month"])
        else:
            data_meanday.loc[index,"Year"] = str(data_meanday.loc[index,"Year"]) + "-" +str(data_meanday.loc[index,"Month"])
            data_workday.loc[index,"Year"] = str(data_workday.loc[index,"Year"]) + "-" +str(data_workday.loc[index,"Month"])
            data_nonworkday.loc[index,"Year"] = str(data_nonworkday.loc[index,"Year"]) + "-" +str(data_nonworkday.loc[index,"Month"])
    
    data_meanday = data_meanday.drop(columns = ["Month"])
    data_workday = data_workday.drop(columns = ["Month"])
    data_nonworkday = data_nonworkday.drop(columns = ["Month"])
    data_meanday = data_meanday.set_index(["Year","Blck"]) # .sort_index()
    data_workday = data_workday.set_index(["Year","Blck"]) # .sort_index()
    data_nonworkday = data_nonworkday.set_index(["Year","Blck"]) # .sort_index()

    return data_workday, data_nonworkday, data_meanday, data_workday_y, data_nonworkday_y, data_meanday_y

def synex_style(data, name, workbook, worksheet, type_data_flag):
    
    # Header table color
    header_format = workbook.add_format({"fg_color":"#16365C",'font_name':"Arial Narrow",'font_color':"white",'align':"center", 'valign': 'vcenter',"font_size":9, 'bold': True, 'text_wrap':True})

    # row and col index
    row_idx, col_idx = data.shape                                                                       # Data shape

    worksheet.write(0,0,name,header_format)                                                             # Header Format
    
    if type_data_flag == 1 or type_data_flag == 4:                                                                             # heat map
        worksheet.set_column(0, 0, 10)
        worksheet.set_column(1, col_idx, 2)
        worksheet.conditional_format(1,1,row_idx,col_idx,{'type':'3_color_scale'})

    for col_num, value in enumerate(data.columns.values):
        worksheet.write(0,col_num + 1, value, header_format)
    
    for r in range(row_idx):
        worksheet.write( r+1, 0, data.index[r], workbook.add_format({'num_format':1,'align':"center",'font_name':"Arial Narrow",'right':1,"font_size":9}))
        if r == row_idx - 1:
            worksheet.write( r+1, 0, data.index[r], workbook.add_format({'num_format':1,'align':"center",'font_name':"Arial Narrow",'bottom':6,'right':1,"font_size":9}))
        for c in range(col_idx):
            if c < col_idx - 1 and r < row_idx - 1:
                worksheet.write( r+1, c+1, data.loc[ data.index[r], data.columns[c]], workbook.add_format({'num_format':1,'align':"center",
                    'font_name':"Arial Narrow",'right':1,"font_size":9}))
            elif r == row_idx - 1 and c < col_idx - 1:
                worksheet.write( r+1, c+1, data.loc[ data.index[r], data.columns[c]], workbook.add_format({'num_format':1,'align':"center",
                    'font_name':"Arial Narrow",'right':1,'bottom':6,"font_size":9}))
            elif c == col_idx - 1 and r == row_idx - 1:
                worksheet.write( r+1, c+1, data.loc[data.index[r], data.columns[c]], workbook.add_format({'num_format':1,'align':"center",
                    'font_name':"Arial Narrow",'bottom':6,"font_size":9}))
            else:
                worksheet.write( r+1, c+1, data.loc[ data.index[r], data.columns[c]], workbook.add_format({'num_format':1,'align':"center",
                    'font_name':"Arial Narrow","font_size":9}))

def polygon_under_graph(x, y):
    """
    Construct the vertex list which defines the polygon filling the space under
    the (x, y) line graph. This assumes x is in ascending order.
    """
    return [(x[0], 0.), *zip(x, y), (x[-1], 0.)]

def plot_CMg_3d(data,bus,language,cmg_lim, file_root):

    bus_selected = data[bus].unstack(level = 1)
    year_period = [bus_selected.index[0], bus_selected.index[-1]]

    ax = plt.figure(figsize = (9.5/2.54, 7.75/2.54), dpi = 200.0).add_subplot(projection='3d')
    verts = [polygon_under_graph(bus_selected.columns.values, bus_selected.loc[year,:].values) for year in bus_selected.index]
    facecolors = plt.get_cmap("Blues")(np.linspace(0, 1, 2*len(verts)))[-len(verts):]
    poly = PolyCollection(verts, facecolors = facecolors, alpha=0.7, edgecolors = '#002060')
    ax.add_collection3d(poly, zs=bus_selected.index.values, zdir='y')
    ax.view_init(15,-70)
    

    xlabel_ticks = [i for i in range(1,25,4)]

    ylabel_ticks = [i for i in range(year_period[0],year_period[1] + 1, 6)]

    if language:
        ax.set(xlim=(1, 24), ylim=(year_period[0],year_period[1]), zlim=(0, cmg_lim),
        xlabel='Hours', ylabel='Years', zlabel='USD/MWh',
            xticks = xlabel_ticks, yticks = ylabel_ticks)
    else:
        ax.set(xlim=(1, 24), ylim=(year_period[0],year_period[1]), zlim=(0, cmg_lim),
        xlabel='Horas', ylabel='Años', zlabel='USD/MWh',
            xticks = xlabel_ticks, yticks = bus_selected.index.values)
    
    ax.set_aspect('auto')

    plt.savefig(file_root + bus + ".png")
    plt.close()

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

def plot_year_month_data( data_month, bus, language):

    if language:
        df_out = pd.DataFrame(0, index = data_month.index, columns = ["Month Average Spot Price", "Year Average Spot Price"])
        df_out["Month Average Spot Price"] = data_month[bus]
        data_year = data_month[bus].groupby(level = [0]).mean()
        for index, row in df_out.iterrows():
            df_out.loc[index,"Year Average Spot Price"] = data_year[index[0]]

    else:
        df_out = pd.DataFrame(0, index = data_month.index, columns = ["Promedio Mensual Costo marginal", "Promedio Anual Costo marginal"])
        df_out["Promedio Mensual Costo marginal"] = data_month[bus]
        data_year = data_month[bus].groupby(level = [0]).mean()
        for index, row in df_out.iterrows():
            df_out.loc[index,"Promedio Anual Costo marginal"] = data_year[index[0]]

    df_out = df_out.reset_index()
    df_out["Month"] = df_out["Month"].astype("str") + "-" + df_out["Year"].astype("str")
    df_out = df_out.drop(columns = ["Year"])
    if language:
        df_out = df_out.rename(columns = {"Month":"Date"})
        df_out = df_out.set_index("Date")
    else:
        df_out = df_out.rename(columns = {"Month":"Fecha"})
        df_out = df_out.set_index("Fecha")

    return df_out

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

def hydro_cmg(data, out_index, hydro_condition, hydro_limit, short_term_flag, first_date, first_series, second_date, second_series):

    t0 = exectime("Cmgbus by hydro condition", True, 0)
    CMG_wet_hydrology = pd.DataFrame(0, index = out_index, columns = data.columns)
    CMG_dry_hydrology = pd.DataFrame(0, index = out_index, columns = data.columns)
    CMG_mid_hydrology = pd.DataFrame(0, index = out_index, columns = data.columns)

    counter = 0
    printProgressBar(0, len(out_index), prefix = 'Progress:', suffix = 'Complete', length = 50)

    if short_term_flag:                     # Changes in short-term hydro condition
        for index, row in CMG_wet_hydrology.iterrows():
            
            if index < first_date:
                CMG_dry_hydrology.loc[index,:] = data.loc[(index[0],index[1]),:].loc[(hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] >= first_series[0]].index) & (hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] <= first_series[1]].index),:].mean()
                CMG_wet_hydrology.loc[index,:] = CMG_dry_hydrology.loc[index,:]
                CMG_mid_hydrology.loc[index,:] = CMG_dry_hydrology.loc[index,:]

            elif index >= first_date and index <= second_date:
                CMG_dry_hydrology.loc[index,:] = data.loc[(index[0],index[1]),:].loc[(hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] >= second_series[0]].index) & (hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] <= second_series[1]].index),:].mean()

                CMG_wet_hydrology.loc[index,:] = CMG_dry_hydrology.loc[index,:]

                CMG_mid_hydrology.loc[index,:] = CMG_dry_hydrology.loc[index,:]
            else: 
                CMG_wet_hydrology.loc[index,:] = data.loc[(index[0],index[1]),:].loc[hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] <= hydro_limit[0]].index,:].mean()

                CMG_dry_hydrology.loc[index,:] = data.loc[(index[0],index[1]),:].loc[hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] >= hydro_limit[1]].index,:].mean()

                CMG_mid_hydrology.loc[index,:] = data.loc[(index[0],index[1]),:].loc[(hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] < hydro_limit[1]].index) & (hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] > hydro_limit[0]].index),:].mean()

            counter += 1
            printProgressBar(counter, len(out_index), prefix = 'Progress:', suffix = 'Complete', length = 50)
    else:
        for index, row in CMG_wet_hydrology.iterrows():
            
            CMG_wet_hydrology.loc[index,:] = data.loc[(index[0],index[1]),:].loc[hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] <= hydro_limit[0]].index,:].mean()

            CMG_dry_hydrology.loc[index,:] = data.loc[(index[0],index[1]),:].loc[hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] >= hydro_limit[1]].index,:].mean()

            CMG_mid_hydrology.loc[index,:] = data.loc[(index[0],index[1]),:].loc[(hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] < hydro_limit[1]].index) & (hydro_condition.loc[(index[0],index[1]),:][hydro_condition.loc[(index[0],index[1]),:] > hydro_limit[0]].index),:].mean()
            counter += 1
            printProgressBar(counter, len(out_index), prefix = 'Progress:', suffix = 'Complete', length = 50)
    t0 = exectime("Cmgbus by hydro condition", False, t0)
    return CMG_wet_hydrology, CMG_mid_hydrology, CMG_dry_hydrology

def wind_out_data(data, hydro_condition, ST_flag, ST_date1, ST_serie1, ST_date2, ST_serie2):
    
    HW = [i for i in range(1,93,3)]                                                                             # High - wind series
    MW = [i for i in range(2,93,3)]                                                                             # Middle - wind series
    LW = [i for i in range(3,94,3)]                                                                             # Low - wind series

    cmgbus_HW = data[data.index.get_level_values("Seq.").isin(HW)].groupby(level = [0,1]).mean()                                              # Stage-Block High-wind CMg
    cmgbus_MW = data[data.index.get_level_values("Seq.").isin(MW)].groupby(level = [0,1]).mean()                                              # Stage-Block Middle-wind CMg
    cmgbus_LW = data[data.index.get_level_values("Seq.").isin(LW)].groupby(level = [0,1]).mean()                                              # Stage-Block Low-wind CMg

    if ST_flag:
        for index, row in cmgbus_HW.iterrows():
            if index < ST_date1:
                cmgbus_HW.loc[index,:] = data.loc[index,:].loc[(i for i in hydro_condition.loc[index,:][hydro_condition.loc[index,:] >= ST_serie1[0]].index.intersection(hydro_condition.loc[index,:][hydro_condition.loc[index,:] <= ST_serie1[1]].index) if i in HW),:].mean()
                cmgbus_MW.loc[index,:] = data.loc[index,:].loc[(i for i in hydro_condition.loc[index,:][hydro_condition.loc[index,:] >= ST_serie1[0]].index.intersection(hydro_condition.loc[index,:][hydro_condition.loc[index,:] <= ST_serie1[1]].index) if i in MW),:].mean()
                cmgbus_LW.loc[index,:] = data.loc[index,:].loc[(i for i in hydro_condition.loc[index,:][hydro_condition.loc[index,:] >= ST_serie1[0]].index.intersection(hydro_condition.loc[index,:][hydro_condition.loc[index,:] <= ST_serie1[1]].index) if i in LW),:].mean()
            elif index >= ST_date1 and index <= ST_date2:
                cmgbus_HW.loc[index,:] = data.loc[index,:].loc[(i for i in hydro_condition.loc[index,:][hydro_condition.loc[index,:] >= ST_serie2[0]].index.intersection(hydro_condition.loc[index,:][hydro_condition.loc[index,:] <= ST_serie2[1]].index) if i in HW),:].mean()
                cmgbus_MW.loc[index,:] = data.loc[index,:].loc[(i for i in hydro_condition.loc[index,:][hydro_condition.loc[index,:] >= ST_serie2[0]].index.intersection(hydro_condition.loc[index,:][hydro_condition.loc[index,:] <= ST_serie2[1]].index) if i in MW),:].mean()
                cmgbus_LW.loc[index,:] = data.loc[index,:].loc[(i for i in hydro_condition.loc[index,:][hydro_condition.loc[index,:] >= ST_serie2[0]].index.intersection(hydro_condition.loc[index,:][hydro_condition.loc[index,:] <= ST_serie2[1]].index) if i in LW),:].mean()
            else:
                break
    
    return cmgbus_HW, cmgbus_MW, cmgbus_LW



##########      MAIN CODE      ##########


### Short term conditions

start_date_output = (2022,6)                                                                                # Start date for output data
ST_flag = True
short_term_hydro_cond = [26,31]                                                                             # Hydro condition for the shorttest time
short_term_date = (2023,4)                                                                                  # Short term hydro date
short_mid_term_hydro_cond = [26,31]                                                                         # hydro condition for the short time
short_mid_term_date = (2024,4)                                                                              # Short-mid term hydro date
English = True


Hydro_limit_condition = [3,29]                                                                              # 3rd hydro as the wettest limit, 29th as the driest limit
init_hydro_date = 1988
last_hydro_date = 2018
file_root = "C:/Users/Seba Morales/OneDrive - Synex Ingenieros Consultores Ltda/Casos2022/PlanDeObras/"

typical_buses = ["PAzucar220","Crucero220", "Polpaico220","Cardones220","Charrua220","PMontt220"]           # Typical SEN buses
# typical_buses = cmgbus.columns.values                                                                           # All Sen buses

## Read data

date = init_date(file_root + "cmgbus.csv")                                                                  # Starting Cases date
cmgbus = readfile( file_root + "cmgbus.csv", 3, [])                                                         # read file cmgbus
cmgbus = cmgbus.rename(columns= {col: col.replace(" ","") for col in cmgbus.columns})                       # drop spaces in columns

duraci = readfile(file_root + "duraci.csv", 3, [0,1,2,3])                                                   # read file duraci
duraci = month_year_index(duraci, date, start_date_output)                                                  # duraci data with year-month values
duraci = duraci.rename(columns = {duraci.columns[0]:"SEN"})
duraci["Total_month"] = 0
for index,row in duraci.iterrows():
    duraci.loc[index, "Total_month"] = duraci.loc[(index[0],index[1],index[2]),"SEN"].sum()

hidroseq = readfile("Orden_hidrologias.xlsx", 0, [])                                                        # read hydro seq file
hidroseq = hidroseq[(hidroseq["Year"] >= init_hydro_date) & (hidroseq["Year"] <= last_hydro_date)]          # take just hydro sequence used
hidroseq = hidroseq.sort_values(by = "Total", ascending = False)                                            # sort hydrologies frow wet to dry
hidroseq["index"] = [i for i in range(1,32)]                                                                # index values by preview order
hidroseq = hidroseq.set_index("index")                                                                      # Index it's setted

cmgbus = month_year_index(cmgbus, date, start_date_output)                                                  # cmgbus data with year-month values
cmgbus_stage_block_mean = cmgbus.groupby(level = [0,1,3]).mean()                                            # Data with year-month-block mean
cmgbus_month, cmgbus_year = weighted_mean(cmgbus, duraci, True)                                             # Data with month and year average
cmgbus_month_serie, cmgbus_year_serie = weighted_mean(cmgbus, duraci, False)                                # Data with block average


###     Data by hydro condition     ###

hydro_condition = case_hydro_condition(init_hydro_date, last_hydro_date, cmgbus_month.index, hidroseq)      # Get hydro condition for every stage

wet_cmgbus, mid_cmgbus, dry_cmgbus = hydro_cmg(cmgbus_month_serie, cmgbus_month.index,
                                                hydro_condition, Hydro_limit_condition, ST_flag,
                                                short_term_date, short_term_hydro_cond,
                                                short_mid_term_date, short_mid_term_hydro_cond)             # CMg by hydro condition


###     Wind condition     ###

cmgbus_HW, cmgbus_MW, cmgbus_LW = wind_out_data(cmgbus_month_serie, hydro_condition,                        # CMg by wind condition (high, medium, low)
                                    ST_flag, short_term_date, short_term_hydro_cond,                        # If Flag = True -> short-term variation by hydrology
                                    short_mid_term_date, short_mid_term_hydro_cond)                         # output indexed by month


###     Daily format     ###

working_day = [1,2,3,4,5,6,7,8,9,10,11,12]                                                                  # Blocks for working day
Non_working_day = [1,2,3,13,14,15,16,17,18,19,20,12]                                                        # Blocks for Non working day

CMgbus_WD, CMgbus_NWD, CMgbus_Mean, CMg_WD_y, CMg_NWD_y, CMg_Mean_y = daily_format(cmgbus, working_day,
                                        Non_working_day, 
                                        duraci["SEN"], False, hydro_condition,
                                        ST_flag, short_term_date, short_term_hydro_cond,
                                        short_mid_term_date, short_mid_term_hydro_cond)                     # CMg x hour in WD, NWD and meanday


### output files

export_excel(typical_buses, cmgbus_month, None, None, 0, file_root,"CMg_month.xlsx", English)             # Export excel with month-year CMg values

export_excel(typical_buses, CMg_Mean_y, None, None, 1, file_root,"CMg_Hour.xlsx", English)                # Export excel with hour-year CMg values

export_excel(typical_buses, wet_cmgbus, mid_cmgbus, dry_cmgbus, 2, file_root,"CMg_hydro.xlsx", English)   # Export excel with hydro condition CMg values

export_excel(typical_buses, cmgbus_HW, cmgbus_MW, cmgbus_LW, 3, file_root,"CMg_wind.xlsx", English)       # Export excel with wind condition CMg values

export_excel(typical_buses, CMgbus_Mean, None, None, 4, file_root,"CMg_Month_Hour.xlsx", English)         # Export excel with hour-year CMg values