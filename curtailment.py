

#### Libraries


import pandas as pd
from time import process_time


#### Functions 

def datreader(file, colnums,headernames,rows_skip):
    t0 = exectime(file, True, 0)
    df = pd.read_csv(file, header = 0,sep="\s+", usecols = colnums,skiprows = rows_skip)
    # for i in range(len(colnums)):
    #     df = df.rename(columns = {df.columns[i]: headernames[i]})
    print(df)
    t0 = exectime(file, False, t0)
    return df

def gennames(file,rows_skip):
    skip = []
    for i in range(rows_skip):
        skip.append(i)
    df = pd.read_csv(file, skiprows = skip,nrows= 0)
    df = df.rename(columns = {i:i.replace(" ","") for i in df.columns})
    return df

def central_bus_linker(dbus, cgndse):
    t0 = exectime(cgndse, True, 0)
    data = [[],[]]
    df1 = pd.DataFrame(dbus["Gen_Name"].dropna(),columns = ["Gen_Name"])
    hid = gennames("..\gerhid.csv", 3)
    gnd = gennames("..\gergnd.csv", 3)
    for index,row in df1.iterrows():                #   Append Gen, busname and if it's cuttable
        data[0].append(dbus[dbus["Gen_Name"] == row["Gen_Name"]]["Bus_Name"].values[0])
        if row["Gen_Name"] in hid.columns:
            data[1].append(1)
        else:
            data[1].append(0)
    df2 = pd.DataFrame(cgndse["Gen_Name"].dropna(),columns = ["Gen_Name"])
    for index,row in cgndse.iterrows():
        data[0].append(dbus[dbus["Bus_Num"] == row["Bus_Num"]]["Bus_Name"].values[0])
        if (row["Gen_Name"] in gnd.columns) and ("PMGD" not in row["Gen_Name"]) and ("_Res_" not in row["Gen_Name"]):
            data[1].append(1)
        else:
            data[1].append(0)
    df3 = pd.concat([df1,df2],ignore_index = True)
    df3["Bus_Name"] = data[0]
    df3["Gen_Cuttable"] = data[1]
    df3 = df3.sort_values(by = ["Gen_Name"], ignore_index = True)
    print("Dbus\n",df1,"\ncgndse\n",df2,"\nGen\n",df3)
    t0 = exectime(cgndse, False, t0)
    return df3

def connectmatrix(dbus,dcirc):
    t0 = exectime("dcirc", True, 0)
    titles = []
    df = dbus.sort_values(by = ["Bus_Num"])
    for index,row in df.iterrows():
        if row["Bus_Name"] not in titles:
            titles.append(row["Bus_Name"])
    df = pd.DataFrame(columns = titles, index = titles)
    for index, row in dcirc.iterrows():
        df.at[dbus[dbus["Bus_Num"] == row["Bus_Origin_Num"]]["Bus_Name"],dbus[dbus["Bus_Num"] == row["Bus_Destiny_Num"]]["Bus_Name"]] = 1
        df.at[dbus[dbus["Bus_Num"] == row["Bus_Destiny_Num"]]["Bus_Name"],dbus[dbus["Bus_Num"] == row["Bus_Origin_Num"]]["Bus_Name"]] = 1
    df = df.replace({None: 0})
    print(df)
    t0 = exectime("", False, t0)
    return df

def readdatacsv(file,skip):
    t0 = exectime(file,True, 0)
    skp_rows = []
    for i in range(skip):
        skp_rows.append(i)
    df = pd.DataFrame(pd.read_csv(file, skiprows = skp_rows,header = 0))
    df = df.rename(columns = {i:i.replace(" ","") for i in df.columns})
    print(df)
    t0 = exectime(file,False, t0)
    return df

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

def cmgcomparator(cmgbus):
    t0 = exectime("cmgbus", True, 0)
    df = cmgbus
    # for index, row in df.iterrows():
    for i in df.columns:
        df.loc[(df[i] <1), i] = 1
        df.loc[(df[i] != 1), i] = 0
    print(df)
    t0 = exectime("cmgbus", False, t0)
    return df

def vertfinder(cmgbus):
    t0 = exectime("cmgbus reduction", True, 0)
    vertbus = []
    for index, row in cmgbus.iterrows():
        vertbus.append([])
        vertbus[len(vertbus)-1].append(cmgbus.loc[0,cmgbus.loc[cmgbus.index[0]] < 1].index.values)
        # vertbus.append([])
        # for i in range(3,len(data.columns)):
        #     if row[data.columns[i]] == 1:
        #         vertbus[len(vertbus)-1].append(data.columns[i])
    t0 = exectime("cmgbus reduction", False, t0)
    print(vertbus[0])
    return vertbus

def vertbusfinder(cmgbus, matrix):
    cmgbus = cmgbus.reset_index()
    t0 = exectime("CMg for vert bus extraction", True, 0)
    printProgressBar(0, len(cmgbus), prefix = 'Progress:', suffix = 'Complete', length = 50)
    output = []
    out_bloq = []
    for index,row in cmgbus.iterrows():
        output.append(cmgbus.loc[index,cmgbus.loc[cmgbus.index[index]] < 1].index.values.tolist())
        if len(cmgbus.loc[index,cmgbus.loc[cmgbus.index[index]] < 1].index.values.tolist()) == len(matrix.index.values):
            out_bloq.append([cmgbus.loc[index,cmgbus.loc[cmgbus.index[index]] < 1].index.values.tolist()])
        elif cmgbus.loc[index,cmgbus.loc[cmgbus.index[index]] < 1].index.values.tolist() != []:
            out_bloq.append(recursivebus(matrix, cmgbus.loc[index,cmgbus.loc[cmgbus.index[index]] < 1].index.values.tolist(), []))
        else:
            out_bloq.append([])
        printProgressBar(index + 1, len(cmgbus), prefix = 'Progress:', suffix = 'Complete', length = 50)
        # if index > 50:
        #     break
    # print(output[0])
    # print(out_bloq[0])
    t0 = exectime("CMG", False, t0)
    return output, out_bloq

def cmgfinder(cmgbus):
    t0 = exectime("CMg extraction", True, 0)
    listnames = []
    printProgressBar(0, len(cmgbus), prefix = 'Progress:', suffix = 'Complete', length = 50)
    for i in cmgbus.columns:
        listnames.append(cmgbus[cmgbus[i] < 1][i].name)
        printProgressBar(len(listnames), len(cmgbus.columns), prefix = 'Progress:', suffix = 'Complete', length = 50)
    return listnames

def recursivezone(node,nodelist, matrix,out,list2):
    list1 = [item for item in matrix.loc[matrix[node] == 1,node].index.values if (item in nodelist) and (item not in out)]
    list2 += list1
    if node not in out:
        out.append(node)
    out += list1
    nodelist.remove(node)
    if node in list2:
        list2.remove(node)
    if list2 != []:
        for i in list2:
            return recursivezone(i, nodelist, matrix, out, list2)
    else:
        return out, nodelist

def recursivebus(matrix,nodelist,zone):
    out, restnodes = recursivezone(nodelist[0], nodelist, matrix, [],[])
    zone.append(out)
    if restnodes != []:
        return recursivebus(matrix, nodelist, zone)
    else:
        return zone

def vertzone(vertbus, matrix):
    t0 = exectime("vertzone recursion",True, 0)
    printProgressBar(0, len(vertbus), prefix = 'Progress:', suffix = 'Complete', length = 50)
    out = []
    for i in range(len(vertbus)):       # len(vertbus)
        if len(vertbus[i]) == len(matrix.index.values):
            out.append(vertbus[i])
        elif vertbus[i] != []:
            out.append(recursivebus(matrix, vertbus[i], []))
        else:
            out.append([])
        printProgressBar(i + 1, len(vertbus), prefix = 'Progress:', suffix = 'Complete', length = 50)
    print(out[0])
    t0 = exectime("vertzone recursion",False, t0)
    return out

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

def vertBusgnd(gen_bus_names, vergnd,name, busnames):
    t0 = exectime(name, True, 0)
    df = pd.DataFrame(vergnd["Stag"])
    df = df.join([vergnd["Seq."],vergnd["Blck"]])
    for i in range(3,len(vergnd.columns)):
        if gen_bus_names[gen_bus_names["Gen_Name"] == vergnd.columns[i]]["Gen_Cuttable"].values[0] == 1:
            if gen_bus_names[gen_bus_names["Gen_Name"] == vergnd.columns[i]]["Bus_Name"].values[0] not in df.columns:
                df[gen_bus_names[gen_bus_names["Gen_Name"] == vergnd.columns[i]]["Bus_Name"].values[0]] = vergnd[vergnd.columns[i]]
            else:
                df[gen_bus_names[gen_bus_names["Gen_Name"] == vergnd.columns[i]]["Bus_Name"].values[0]] += vergnd[vergnd.columns[i]]
    # for i in busnames:
    #     if i not in df.columns:
    #         df[i] = df.apply(lambda x: 0, axis=1)
    print(df)
    t0 = exectime(name, False, t0)
    return df

def totalvertbus(regen,hidgen,tergen,revert):
    t0 = exectime("total vert",True,0)
    dftotalvert = pd.DataFrame(regen["Stag"])
    dftotalvert = dftotalvert.join([regen["Seq."],regen["Blck"]])
    for i in range(3,len(revert.columns)):
        if revert.columns[i] not in dftotalvert.columns:
            dftotalvert[revert.columns[i]] = revert[revert.columns[i]]
        else:
            dftotalvert[revert.columns[i]] += revert[revert.columns[i]]
    for i in range(3,len(regen.columns)):
        if regen.columns[i] in dftotalvert.columns:
            dftotalvert[regen.columns[i]] += regen[regen.columns[i]]
        else:
            dftotalvert[regen.columns[i]] = regen[regen.columns[i]]
    for i in range(3,len(hidgen.columns)):
        if hidgen.columns[i] in dftotalvert.columns:
            dftotalvert[hidgen.columns[i]] += hidgen[hidgen.columns[i]]
        else:
            dftotalvert[hidgen.columns[i]] = hidgen[hidgen.columns[i]]
    for i in range(3,len(tergen.columns)):
        if tergen.columns[i] in dftotalvert.columns:
            dftotalvert[tergen.columns[i]] += tergen[tergen.columns[i]]
        else:
            dftotalvert[tergen.columns[i]] = tergen[tergen.columns[i]]
    print(dftotalvert)
    t0 = exectime("total vert",False,t0)
    return dftotalvert

def totalvertbus1(regen,hidgen,tergen,revert):
    t0 = exectime("total vert",True,0)
    dftotalvert = pd.DataFrame(regen["Stag"])
    dftotalvert = dftotalvert.join([regen["Seq."],regen["Blck"]])
    for i in range(3,len(regen.columns)):
        if regen.columns[i] not in dftotalvert.columns:
            dftotalvert[regen.columns[i]] = regen[regen.columns[i]]
        else:
            dftotalvert[regen.columns[i]] += regen[regen.columns[i]]
    for i in range(3,len(hidgen.columns)):
        if hidgen.columns[i] not in dftotalvert.columns:
            dftotalvert[hidgen.columns[i]] = hidgen[hidgen.columns[i]]
        else:
            dftotalvert[hidgen.columns[i]] += hidgen[hidgen.columns[i]]
    for i in range(3,len(tergen.columns)):
        if tergen.columns[i] not in dftotalvert.columns:
            dftotalvert[tergen.columns[i]] = tergen[tergen.columns[i]]
        else:
            dftotalvert[tergen.columns[i]] += tergen[tergen.columns[i]]
    for i in range(3,len(revert.columns)):
        if revert.columns[i] not in dftotalvert.columns:
            dftotalvert[revert.columns[i]] = revert[revert.columns[i]]
        else:
            dftotalvert[revert.columns[i]] += revert[revert.columns[i]]
    print(dftotalvert)
    t0 = exectime("total vert",False,t0)
    return dftotalvert

def vertpercentage(totalvert,revert,vert):
    t0 = exectime("vertzone percentage",True, 0)
    printProgressBar(0, len(totalvert.index), prefix = 'Progress:', suffix = 'Complete', length = 50)
    df = pd.DataFrame(columns = totalvert.columns, index = totalvert.index)
    df["Stag"] = totalvert["Stag"]
    df["Seq."] = totalvert["Seq."]
    df["Blck"] = totalvert["Blck"]
    for index, row in totalvert.iterrows():
        for zones in vert[index]:
            num = 0
            den = 0
            for i in zones:
                if i in revert.columns and i in totalvert.columns:    
                    num += revert.loc[index,i]
                    den += totalvert.loc[index,i]
                    print(num,den)
            if den != 0:
                df.loc[index, (j for j in zones)] = num/den
            else:
                df.loc[index, (j for j in zones)] = 0
        # printProgressBar(index + 1, len(totalvert.index), prefix = 'Progress:', suffix = 'Complete', length = 50)
        if index > 1860:
            break
    print(df)
    print(df.loc[1,df.loc[df.index[1]] > 0])
    print(df.loc[len(df.index)-1,df.loc[df.index[len(df.index)-1]] > 0])
    # print(df.loc[df[df.columns[(i for i in df.columns)] > 1]])
    print(df.loc[df[df.columns[5]] > 1])
    t0 = exectime("vertzone percentage",False, t0)
    return df

def exportexcel(listnames, file, *args):
    t0 = exectime("Exporting data to: "+file, True, 0)
    printProgressBar(0, len(args), prefix = 'Progress:', suffix = 'Complete', length = 50)
    with pd.ExcelWriter(file, engine = 'xlsxwriter') as writer:
        for i in range(len(args)):
            args[i][listnames[i]].unstack(level = -2).to_excel(writer, sheet_name = listnames[i])
            printProgressBar(i+1, len(args), prefix = listnames[i], suffix = 'Complete', length = 50)           # 'Progress:'
    t0 = exectime("Exporting data to: "+file, False, t0)

def curtailment_percentage(data_total,data_vert,databus):
    t0 = exectime(" for percentage calculation",True, 0)
    print(data_vert)
    print(data_total)
    df_num = pd.DataFrame(columns = data_total.columns, index = data_total.index)
    df_den = pd.DataFrame(columns = data_total.columns, index = data_total.index)
    printProgressBar(0, len(databus), prefix = 'Progress:', suffix = 'Complete', length = 50)
    for block in range(len(databus)):
        # print(databus[block])
        for zone in databus[block]:
            num = 0
            den = 0
            # print(zone,data_vert.index[block])
            if len(zone) == 1:
                df_num.loc[df_num.index[block],zone] = data_vert.loc[data_vert.index[block],zone]
                df_den.loc[df_den.index[block],zone] = data_total.loc[data_total.index[block],zone]
            else:
                for i in zone:
                    num += data_vert.loc[data_vert.index[block],i]
                    den += data_total.loc[data_total.index[block],i]
                    # print(data_vert.loc[data_vert.index[block],i])
                for i in zone:
                    df_num.loc[df_num.index[block],i] = num
                    df_den.loc[df_den.index[block],i] = den
            # print(num,den)
            printProgressBar(0, block + 1, prefix = 'Progress:', suffix = 'Complete', length = 50)
    # df_perc = df_num/df_den
    print(df_num)
    print(df_den)
    # print(df_perc)
    # print(df_num)
    t0 = exectime(" for percentage calculation",False, t0)
    return df_num, df_den

def bus_adder(names,*args):
    for i in range(len(args)):
        if i == 0:
            df = pd.DataFrame(args[i]["Stag"])
            df = df.join(args[i]["Seq."])
            df = df.join(args[i]["Blck"])
        for j in range(3,len(args[i].columns)):
            if names[args[i].columns[j]] not in df.columns:
                df[names[args[i].columns[j]]] = args[i][args[i].columns[j]]
            else:
                df[names[args[i].columns[j]]] += args[i][args[i].columns[j]] # df[args[i].columns[j]]
    df = df.set_index(["Stag","Seq.","Blck"])
    print(df)
    return df

def nodefinder(file, names,skip):
    print("______________________________________________")
    print("Processing File ", file)
    t0 = process_time()
    skp_rows = []
    output = []
    output1 = []
    output2 = []
    output3 = {}
    for i in range(skip):
        skp_rows.append(i)
    if file[-4:] == ".csv":
        df = pd.read_csv(file, skiprows = skp_rows,header = 0)
    elif file[-4:] == ".xlsx":
        df = pd.read_excel(file, skiprows = skp_rows, header = 0)
    else:
        print(" Not compatible extension.")
        return None
    for i in range(len(names)):
        for index, row in df.iterrows():
            if row["Gen. name"] == names[i] and names[i] not in output1:
                output.append([names[i],row["Name"].replace(" ","")])
                output1.append(names[i])
                output3[names[i]] = row["Name"].replace(" ","")
                if row["Name"] not in output2:
                    output2.append(row["Name"].replace(" ",""))
        # if len(output) != i+1:
        #     output.append([names[i],"Not Found"])
    t1 = process_time()
    print("Execution Time: ",t1-t0," sec.")
    print("______________________________________________")
    return output,output1, output2, output3

def data_generator(file1, file2, end1):
    data1 = readdatacsv(file1,3)
    data2 = readdatacsv(file2,3)
    data_out = df_combine(data1,data2,end1)
    return data_out

def df_combine(df_CP,df_LP,end1):
    t0 = exectime("Appending DataFrames",True, 0)
    df_CP = df_CP[df_CP["Stag"] < end1 + 1]
    df_out = df_CP.append(df_LP,ignore_index = True)
    df_out = df_out.fillna(0)
    print(df_out)
    t0 = exectime("Appending DataFrames",False, t0)
    return df_out




#### MAIN CODE ####

# endpoint = "C:/Seba/20_93_1513_18102021 - EBS LP - Arroyo"

print("Welcome to Curtailment calculator, please add case's path:\n")
endpoint = input()

dbus = readdatacsv(endpoint + "/dbus.csv",1)
dbus = dbus.rename(columns = {"!Code":"Bus_Num","Name":"Bus_Name","Gen.name":"Gen_Name"})
dbus = dbus[["Bus_Num","Bus_Name","Gen_Name"]]

buses = []
list_ready = False
print("Please add buses to export:")
while not list_ready:
    bus2add = input()
    if bus2add not in dbus["Bus_Name"].values:
        print("Bus not found on dbus.csv, please try again.")
    else:
        buses.append(bus2add)
        print("Bus added to list, Do you want to add another? (1/0)")
        list_ready = int(input())
        if list_ready == 0:
            list_ready = True
        else:
            list_ready = False

### Cambiar rutas
cmgbus = readdatacsv(endpoint + "/cmgbus.csv", 3)

gerhid = readdatacsv(endpoint + "/gerhid.csv", 3)

gerter = readdatacsv(endpoint + "/gerter.csv", 3)

gergnd = readdatacsv(endpoint + "/gergnd.csv", 3)

vergnd = readdatacsv(endpoint + "/vergnd.csv", 3)


names = []
for i in range(3,len(gerhid.columns)):
    if gerhid.columns[i] not in names:
        names.append(gerhid.columns[i])

for i in range(3,len(gergnd.columns)):
    if gergnd.columns[i] not in names:
        names.append(gergnd.columns[i])

for i in range(3,len(gerter.columns)):
    if gerter.columns[i] not in names:
        names.append(gerter.columns[i])


for i in range(3,len(vergnd.columns)):
    if vergnd.columns[i] not in names:
        print(vergnd.columns[i])
        names.append(vergnd.columns[i])

gennodes, genames, nodenames, dictnodes = nodefinder(endpoint + "/dbus.csv", names,1)

vert_gen = bus_adder(dictnodes,vergnd)
total_gen = bus_adder(dictnodes,gergnd,vergnd,gerhid,gerter)

dcirc = readdatacsv(endpoint + "/dcirc.csv",0)
dcirc = dcirc.rename(columns= {"#BOR.":"Bus_Origin_Num","#BDE.":"Bus_Destiny_Num","Nome........":"Line_Name"})
dcirc = dcirc[["Bus_Origin_Num","Bus_Destiny_Num","Line_Name"]]

matrix = connectmatrix(dbus,dcirc)

cmgbus = cmgbus.set_index(["Stag","Seq.","Blck"])

vertbus, vert = vertbusfinder(cmgbus, matrix)

for i in vertbus:
    for j in i:
        if j not in total_gen.columns:
            total_gen[j] = 0
        if j not in vert_gen.columns:
            vert_gen[j] = 0

total_gen = total_gen.reset_index()
vert_gen = vert_gen.reset_index()
num_out, den_out = curtailment_percentage(total_gen, vert_gen,vert)
num_out1 = num_out
den_out1 = den_out


num_out = num_out.fillna(1000000000)
den_out = den_out.fillna(1000000000)
df_vert = num_out/den_out
df_vert = df_vert.replace(1,0)

df_vert.Stag = gergnd.Stag
df_vert["Seq."] = gergnd["Seq."]
df_vert["Blck"] = gergnd["Blck"]
df_vert = df_vert.set_index(["Stag","Seq.","Blck"])

### Cambiar la ruta de archivo salida
exportexcel(buses,endpoint + "/Curtailment_buses.xlsx",[],df_vert)
