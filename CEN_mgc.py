
import pandas as pd
from time import process_time
import os
from zipfile import ZipFile
import zipfile



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

def readfile(file,file_name,rows_skiped, colmns,sheet_name):
    if file_name[-4:] == ".csv":
        if colmns == []:
            df = pd.DataFrame(pd.read_csv(file,skiprows= rows_skiped))
        else:
            df = pd.DataFrame(pd.read_csv(file,usecols= colmns,skiprows=rows_skiped))
    elif file_name[-5:] == ".xlsx" or file_name[-5:] == ".xlsm":
        if colmns == []:
            if sheet_name == None:
                df = pd.DataFrame(pd.read_excel(file,skiprows= rows_skiped))
            else:
                df = pd.DataFrame(pd.read_excel(file,skiprows= rows_skiped, sheet_name = sheet_name))
        else:
            if sheet_name == None:
                df = pd.DataFrame(pd.read_excel(file,usecols= colmns,skiprows=rows_skiped))
            else:
                df = pd.DataFrame(pd.read_excel(file,usecols= colmns,skiprows=rows_skiped, sheet_name = sheet_name))
    return df

def zip_file_reader(file_root, zip_file, file_name, sheet_name, buses,minute_list):

    with ZipFile(os.path.join(file_root, zip_file)) as z:
        with z.open(file_name) as file:
            if file_name[:3] == "cmg":
                print("Reading MgC data: ",file_name)
                df = readfile(file,file_name, 0,[],sheet_name)
                df = df.rename(columns = {"Mes":"Month", "Día":"Day","Hora":"Hour"})
                df = df.pivot(index = ["Month", "Day","Hour"], columns = "Barra", values = "CMg [mills/kWh]")
                df = df.loc[:,(bus for bus in buses if bus in df.columns)]
            else:
                print("Reading Marginal Power plants: ",file_name)
                df = readfile(file,file_name, 8,[],sheet_name)
                df = df.rename(columns = {df.columns[0]:"Month",df.columns[1]:"Day",df.columns[2]:"Hour",df.columns[3]:"Minute"})
                df = df.drop(index = 0)
                df = df.set_index(["Month", "Day","Hour","Minute"])
                df = df.loc[df.index.get_level_values("Minute").isin(minute_list),(bus for bus in buses if bus in df.columns)]
            return df

def export_excel(sheetnames, filename, filepath, *args):
    t0 = exectime("output file:",True, 0)
    printProgressBar(0, len(sheetnames), prefix = 'Progress:', suffix = 'Complete', length = 50)
    with pd.ExcelWriter(os.path.join(filepath, filename), engine = 'xlsxwriter') as writer:
        for i in range(len(sheetnames)):
            args[i].to_excel(writer, sheet_name = sheetnames[i])
            printProgressBar( i + 1, len(sheetnames), prefix = 'Progress:', suffix = 'Complete', length = 50)
        t0 = exectime("output file:",False, t0)


#### MAIN CODE ####

file_root = "../../Casos2022/CMg"

bus_list = ["TARAPACA______220","CRUCERO_______220","ATACAMA_______220","CARDONES______220",
            "P.AZUCAR______220","QUILLOTA______220","CHARRUA_______220","P.MONTT_______220"]

CMg_list = os.listdir(file_root)


for zipped in CMg_list:
    actual_zip = os.path.join(file_root, zipped)
    file_list = ZipFile(os.path.join(file_root, zipped)).namelist()
    file_list.sort(reverse = True)
    cmg = zip_file_reader(file_root, zipped, file_list[0], "CMG Barras",bus_list,[1,10,20,30,40,50,60])
    marginal_PowerPlants = zip_file_reader(file_root, zipped, file_list[1],None,bus_list,[1,10,20,30,40,50,60])
    if zipped == CMg_list[0]:
        cmg_out = cmg
        marginal_PP_out = marginal_PowerPlants
    else:
        cmg_out = pd.concat([cmg_out,cmg])
        marginal_PP_out = pd.concat([marginal_PP_out, marginal_PowerPlants])

cmg_out = cmg_out.reset_index()
cmg_out.loc[:,"Year"] = cmg_out.loc[:,"Month"].astype('str').str[:4].astype('int')
cmg_out.loc[:,"Month"] = cmg_out.loc[:,"Month"].astype('str').str[-2:].astype('int')
marginal_PP_out = marginal_PP_out.reset_index()
marginal_PP_out["Year"] = marginal_PP_out.loc[:,"Month"].astype('str').str[:4].astype('int')
marginal_PP_out["Month"] = marginal_PP_out.loc[:,"Month"].astype('str').str[-2:].astype('int')
cmg_out = cmg_out.loc[:,["Year","Month","Day","Hour"] + bus_list]
marginal_PP_out = marginal_PP_out.loc[:,["Year","Month","Day","Hour","Minute"] + bus_list]


export_excel(["CMg","CentralesMarginales"],"Mgc_CEN.xlsx",file_root,cmg_out,marginal_PP_out)