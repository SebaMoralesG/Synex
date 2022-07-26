##################################################################################################################################################################################

####  Calculo compensaciones PMGD    ###

## Descripción y procedimiento:
# - Generar documentos .csv para potencia nominal, barras sistema, generacion renovable, CMg por energia renovable y demanda SEN
# - Entregar nombres de archivos en consola / corroborar nombre de archivo en documento
# - Correr código con la sentencia "python pmgd_bloque.py"                 (python [nombre archivo].py)
# - Abrir archivo PN-Calculo cuando sea requerido y elegir barra en cuestión segun indica consola, lueg  o de elegir barra guardar archivo PN-Calculo
# - Resultados se encontraran en archivo CompensacionPMGD.xslx

##################################################################################################################################################################################

### Librerias importadas
import pandas as pd
import numpy as np
import os
from time import process_time
import xlwings as xw

# stage = Mes , seq = serie, blck = bloque hora SDDP(1-20)
# Definicion Horarios PN estabilizado:
# Intervalos: 0-4 | 4-8 | 8-12 | 12-16 | 16-20 | 20-24 |


def datamaker(file, skip, tags):
    print("______________________________________________")
    print("Processing File ", file)
    t0 = process_time()
    skp_rows = []
    for i in range(skip):
        skp_rows.append(i)
    dotpos = file.find(".")
    if file[dotpos:] == ".csv":
        df = pd.read_csv(file, skiprows = skp_rows,header = 0)
    elif file[dotpos:] == ".xlsx":
        df = pd.read_excel(file, skiprows = skp_rows, header = 0)
    else:
        print(" Not compatible extension.")
        return None
    
    output = df.filter(items = ["Stag", "Seq.", "Blck"])
    for i in tags:
        df1 = df.filter(like = i, axis = 1)

        if (df1.columns[0] in output.columns):
            pass
        else:
            output = output.join(df1)
    for i in range(len(output.columns)):
        output = output.rename(columns = {output.columns[i]:output.columns[i].replace(" ","")})
    t1 = process_time()
    print("Execution Time: ",t1-t0," sec.")
    print("______________________________________________")
    return output

def pmgdfinder(file, word, skip):
    print("______________________________________________")
    print("Processing File ", file)
    t0 = process_time()
    skp_rows = []
    counter = 0
    output = []
    for i in range(skip):
        skp_rows.append(i)
    dotpos = file.find(".")
    if file[dotpos:] == ".csv":
        df = pd.read_csv(file, skiprows = skp_rows,header = 0)
    elif file[dotpos:] == ".xlsx":
        df = pd.read_excel(file, skiprows = skp_rows, header = 0)
    else:
        print(" Not compatible extension.")
        return None
    for label, content in df.items():
        if label == "Seq." or label == "Blck" or label == "SFV_Res_I   ":
            pass
        else:
            for i in content:
                if (i <= 9) and (i != 0):
                    counter += 1
            if counter == len(content):
                output.append(label.replace(" ",""))
            counter = 0
    df = df.filter(like = word)
    for i in df.columns:
        output.append(i.replace(" ",""))
    t1 = process_time()
    print("Execution Time: ",t1-t0," sec.")
    print("______________________________________________")
    return output

def nodefinder(file, names,skip):
    print("______________________________________________")
    print("Processing File ", file)
    t0 = process_time()
    skp_rows = []
    counter = 0
    output = []
    output1 = []
    for i in range(skip):
        skp_rows.append(i)
    dotpos = file.find(".")
    if file[dotpos:] == ".csv":
        df = pd.read_csv(file, skiprows = skp_rows,header = 0)
    elif file[dotpos:] == ".xlsx":
        df = pd.read_excel(file, skiprows = skp_rows, header = 0)
    else:
        print(" Not compatible extension.")
        return None
    for i in range(len(names)):
        for index, row in df.iterrows():
            if row["Gen. name"] == names[i]:
                output.append([names[i],row["Name"]])
                output1.append(names[i])
        if len(output) != i+1:
            output.append([names[i],"Not Found"])
    t1 = process_time()
    print("Execution Time: ",t1-t0," sec.")
    print("______________________________________________")
    return output,output1

def nodepricefinder(file,names,skip):
    print("______________________________________________")
    print("Processing File ", file)
    print("Please open file and choose the next bus in AInforme_OtrasBarras_Nuevo: \n")
    t0 = process_time()
    df = pd.read_excel(file, sheet_name = "AInforme_OtrasBarras_Nuevo", skiprows = [0,1,2,4], header = 0,usecols = "I:O", nrows = 40)
    t1 = process_time()
    print("Execution Time: ",t1-t0," sec.")
    print("______________________________________________")
    return df

def blockpricemaker(file,nodes, lendata, initmonth):
    print("______________________________________________")
    print("Processsing file:  ", file,)
    t0 = process_time()
    prices = []
    outputprices = [[],[]]
    outputdata = [[],[],[]]
    for i in nodes:
        nodprice = nodepricefinder(file,nodes, 1)
        prices.append(nodprice)
        # state = True
        # print("please choose bus ",i[1]," - ",i[0])
        # while state:
        #     update = input("Are you ready?(Y/N): ")
        #     if update == "Y":
        #         nodprice = nodepricefinder(file,nodes, 1)
        #         prices.append(nodprice)
        #         state = False
    for i in range(len(prices)):
        outputprices.append([])
        outputdata.append([])
        for index, row in prices[i].iterrows():
            if i == 0:
                outputprices[0].append(row[prices[i].columns[0]])         
                outputprices[0].append(row[prices[i].columns[0]])
                outputprices[0].append(row[prices[i].columns[0]])         
                outputprices[0].append(row[prices[i].columns[0]])
                outputprices[0].append(row[prices[i].columns[0]])         
                outputprices[0].append(row[prices[i].columns[0]])
                outputprices[0].append(row[prices[i].columns[0]])         
                outputprices[0].append(row[prices[i].columns[0]])
                outputprices[0].append(row[prices[i].columns[0]])         
                outputprices[0].append(row[prices[i].columns[0]])
                outputprices[0].append(row[prices[i].columns[0]])         
                outputprices[0].append(row[prices[i].columns[0]])
                outputprices[0].append(row[prices[i].columns[0]])         
                outputprices[0].append(row[prices[i].columns[0]])
                outputprices[0].append(row[prices[i].columns[0]])         
                outputprices[0].append(row[prices[i].columns[0]])
                outputprices[0].append(row[prices[i].columns[0]])         
                outputprices[0].append(row[prices[i].columns[0]])
                outputprices[0].append(row[prices[i].columns[0]])         
                outputprices[0].append(row[prices[i].columns[0]])
                outputprices[1].extend([1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20])
            outputprices[len(outputprices)-1].append(row["Interval 0-4"])           # Precio Nudo para bloque 1 (24-1)      check
            outputprices[len(outputprices)-1].append(row["Interval 0-4"])           # Precio Nudo para bloque 2 (2-3)       check   
            outputprices[len(outputprices)-1].append(row["Interval 4-8"])           # Precio Nudo para bloque 3 (4-5)       check
            outputprices[len(outputprices)-1].append(row["Interval 4-8"])           # Precio Nudo para bloque 4 (6-7)       check
            outputprices[len(outputprices)-1].append(row["Interval 8-12"])           # Precio Nudo para bloque 5 (8-9)      check
            outputprices[len(outputprices)-1].append(row["Interval 8-12"])           # Precio Nudo para bloque 6 (10-11)    check
            outputprices[len(outputprices)-1].append(row["Interval 12-16"])           # Precio Nudo para bloque 7 (12-13)   check
            outputprices[len(outputprices)-1].append(row["Interval 12-16"])          # Precio Nudo para bloque 8 (14-15)    check
            outputprices[len(outputprices)-1].append(row["Interval 16-20"])          # Precio Nudo para bloque 9 (16-17)    check
            outputprices[len(outputprices)-1].append(row["Interval 16-20"])          # Precio Nudo para bloque 10 (18-19)   check
            outputprices[len(outputprices)-1].append(row["Interval 20-24"])          # Precio Nudo para bloque 11 (20-21)   check
            outputprices[len(outputprices)-1].append(row["Interval 20-24"])         # Precio Nudo para bloque 12 (22-23)    check
            outputprices[len(outputprices)-1].append(row["Interval 4-8"])         # Precio Nudo para bloque 13 (6-7)        check
            outputprices[len(outputprices)-1].append(row["Interval 8-12"])         # Precio Nudo para bloque 14 (8-9)       check
            outputprices[len(outputprices)-1].append(row["Interval 8-12"])         # Precio Nudo para bloque 15 (10-11)     check
            outputprices[len(outputprices)-1].append(row["Interval 12-16"])         # Precio Nudo para bloque 16 (12-13)    check
            outputprices[len(outputprices)-1].append(row["Interval 12-16"])         # Precio Nudo para bloque 17 (14-15)    check
            outputprices[len(outputprices)-1].append(row["Interval 16-20"])         # Precio Nudo para bloque 18 (16-17)    check
            outputprices[len(outputprices)-1].append(row["Interval 16-20"])         # Precio Nudo para bloque 19 (18-19)    check
            outputprices[len(outputprices)-1].append(row["Interval 20-24"])         # Precio Nudo para bloque 20 (20-21)    check
    counter = [initmonth - outputprices[0][0].month,0]
    for i in range(1,int(lendata/(93*20))+1):                                       # Iteración por mes - Stage
        for j in range(1,94):                                                       # Iteración por serie
            for k in range(1,21):                                                   # Iteración por bloque
                outputdata[0].append(i)
                outputdata[1].append(j)
                outputdata[2].append(k)
                if counter[0] < 0:
                    for l in range(3,len(outputdata)):
                        outputdata[l].append(outputprices[l-1][k-1])
                elif (counter[0] == 0 and counter[1] == 0):
                    for l in range(3,len(outputdata)):
                        outputdata[l].append(outputprices[l-1][k-1])
                elif counter[1] == 6:
                    for l in range(3,len(outputdata)):
                        outputdata[l].append(outputprices[l-1][(counter[0])*20+(k-1)])
                else:
                    for l in range(3,len(outputdata)):
                        outputdata[l].append(outputprices[l-1][counter[0]*20+(k-1)])
        if counter[0] < 0:
            counter[0] += 1
        elif counter[0] == 0 and counter[1] == 0:
            counter[1] = 1
        elif counter[1] == 6:
            counter[1] == 1
            counter[0] += 1
            if counter[0] > (len(outputprices[0])/20-1):
                counter[0] = int((len(outputprices[0])/20-1))
        else:
            counter[1] += 1
    dictData = {"Stag": outputdata[0], "Seq.": outputdata[1],"Blck":outputdata[2]}
    for i in range(3,len(outputdata)):
        dictData[nodes[i-3][0]] = outputdata[i]
    df = pd.DataFrame(data = dictData)
    return df

    t1 = process_time()
    print("Execution Time: ",t1-t0," sec.")
    print("______________________________________________")
    return df

def exportexcel(genxcmg, demsen, genpmgd, price,  genbloq, gencmgbloq, pricebloq, diffbloq, compbloq):   # bloqcomp, monthcomp
    print("______________________________________________")
    print("Exporting data to CompensacionPMGD.xlsx")
    t0 = process_time()
    with pd.ExcelWriter("CompensacionPMGD.xlsx") as writer:
        print("Exporting: Ingresos_CMg")
        genxcmg.to_excel(writer, sheet_name = "Ingresos_CMg")
        print("Exporting: Generacion PMGD")
        genpmgd.to_excel(writer, sheet_name = "Gen_PMGD")
        print("Exporting:  Precio nudo estabilizado")
        price.to_excel(writer, sheet_name = "Precio_Nudo")  
        print("Exporting: Demanda")
        demsen.to_excel(writer, sheet_name = "Demanda_SEN")
        print("Exporting:  Generacion por bloque")
        genbloq.to_excel(writer, sheet_name = "Gen_blck")
        print("Exporting:  Ingresos CMg por bloque")
        gencmgbloq.to_excel(writer, sheet_name = "Ingreso_Cmg_blck")
        print("Exporting:  Precio nudo por bloque")
        pricebloq.to_excel(writer, sheet_name = "price_blck")
        print("Exporting:  Diferencia por bloque")
        diffbloq.to_excel(writer, sheet_name = "diff_blck")
        print("Exporting:  Compensacion por bloque")
        compbloq.to_excel(writer, sheet_name = "Comp_blck")  
    t1 = process_time()
    print("Execution Time: ",t1-t0," sec.")
    print("files exported")
    print("______________________________________________")

def serieaverage(genxcmg, genpmgd, price,nodes,demsen):
    print("______________________________________________")
    print("Calculating average per serie in each file.\n")
    t0 = process_time()
    gencmgdata = []
    pricedata = []
    genpmgddata = []
    stage = []
    block = []
    counter = [1, 0, 0]             # Stage, Serie, Block
    stage = np.zeros(int(len(genpmgd.index)/93),dtype = float)
    block = np.zeros(int(len(genpmgd.index)/93),dtype = float)
    for i in range(len(nodes)):
        gencmgdata.append(np.zeros(int(len(genxcmg.index)/93),dtype = float))
        pricedata.append(np.zeros(int(len(price.index)/93),dtype = float))
        genpmgddata.append(np.zeros(int(len(genpmgd.index)/93),dtype = float))
    print("Calculating average in: PMGD Generation x CMg")
    for index, row in genxcmg.iterrows():
        counter[2] += 1
        counter[1] += 1
        if counter[1] == 94 and counter[2] == 21:
            counter[0] += 1
            # stage.append(counter[0])
        if counter[2] == 21:
            counter[2] = 1
        if counter[1] == 94:
            counter[1] = 1
        for i in range(len(nodes)):
            gencmgdata[i][20*(counter[0]-1)+counter[2]-1] += row[nodes[i][0]]/93
            stage[20*(counter[0]-1)+counter[2]-1] = counter[0]
            block[20*(counter[0]-1)+counter[2]-1] = counter[2]
            # stage.append(counter[0])
            # block.append(counter[2])
    counter = [1, 0, 0] 
    print("Calculating average in: PMGD generation")
    for index, row in genpmgd.iterrows():
        counter[2] += 1
        counter[1] += 1
        if counter[1] == 94 and counter[2] == 21:
            counter[0] += 1
        if counter[2] == 21:
            counter[2] = 1
        if counter[1] == 94:
            counter[1] = 1
        for i in range(len(nodes)):
            genpmgddata[i][20*(counter[0]-1)+counter[2]-1] += row[nodes[i][0]]/93
    counter = [1, 0, 0] 
    print("Calculating average in: Node Price")
    for index, row in price.iterrows():
        counter[2] += 1
        counter[1] += 1
        if counter[1] == 94 and counter[2] == 21:
            counter[0] += 1
        if counter[2] == 21:
            counter[2] = 1
        if counter[1] == 94:
            counter[1] = 1
        for i in range(len(nodes)):
            pricedata[i][20*(counter[0]-1)+counter[2]-1] += row[nodes[i][0]]/93

    dictdata1 = {"Stage": stage}
    for i in range(len(nodes)):
        dictdata1[nodes[i][0]] = genpmgddata[i]
    df1 = pd.DataFrame(data = dictdata1)

    dictdata2 = {"Stage": stage}
    for i in range(len(nodes)):
        dictdata2[nodes[i][0]] = gencmgdata[i]
    df2 = pd.DataFrame(data = dictdata2)

    dictdata3 = {"Stage": stage}
    for i in range(len(nodes)):
        dictdata3[nodes[i][0]] =pricedata[i]
    df3 = pd.DataFrame(data = dictdata3)

    gennudo = df1*df3
    # print(gennudo.head(5))
    diff = gennudo-df2
    diff["Stage"] = stage
    # print(diff.head(5))

    df4 = diff
    for i in nodes:
        df4[i[0]] = df4[i[0]]//demsen["SEN"]
    
    df4["Stage"] = stage
    # print(df4.head(10))
    # print(df4.tail(10))
    t1 = process_time()
    print("Execution Time: ",t1-t0," sec.")
    print("______________________________________________")
    return df1,df2,df3,diff,df4

def monthprice(file, sheetname,nodes, start,numrows):
    df_ref = pd.read_excel(file, sheet_name = sheetname, skiprows = [0,1,2,4], header = 0,usecols = "B:C", index_col = 0,nrows = 40)
    df_ref = df_ref.rename(columns = {df_ref.columns[0]:"Referencia"})
    df_nodes = pd.read_excel(file, sheet_name = sheetname, skiprows = [0,1,2,3], header = 0,usecols = "V:BO", nrows = 40)
    for i in nodes:
        df_actual = df_nodes.filter(like = i[1][0:-3])
        for j in df_actual.columns:
            if j in df_ref:
                pass
            else:
                df_actual = df_actual.rename(columns = {j: j.replace(" ","")})
                df_ref[j.replace(" ","")] = df_actual[j.replace(" ","")].to_numpy()
                df_ref[j.replace(" ","")] = df_ref[j.replace(" ","")]*df_ref["Referencia"]
    counter = [start[0]-df_ref.index[0].month +12*(start[1]-df_ref.index[0].year),0]
    for i in range(numrows):
        if i == 0:
            df1 = df_ref.iloc[[0]]
            if counter[0] < 0:
                counter[0] += 1
            else:
                counter[1] += 1
        else:
            if counter[0] < 0:
                df1 = df1.append(df_ref.iloc[[0]], ignore_index = True)
                counter[0] += 1
            else:
                df1 = df1.append(df_ref.iloc[[counter[0]]], ignore_index = True)
            counter[1] += 1
            if counter[1] == 6:
                counter[1] = 0
                counter[0] += 1
                if counter[0] == 40:
                    counter[0] = 39
    for i in df1.columns:
        if i == "Referencia":
            pass
        else:
            df1 = df1.rename(columns = {i:i[0:-3]})
    return df1

def monthyearaverage(file,start):
    df = file
    df_out = pd.DataFrame(columns = df.columns)
    df_out2 = df_out
    for i in range(int(len(file.index)/20)):
        df1 = df.iloc[20*i:20*(i+1)]
        df_out = df_out.append(df1.sum(axis = 0),ignore_index = True)
    for i in reversed(df_out.index):
        df_out = df_out.rename(index = {i:i+1})
    delay = start -1
    delay1 = 0
    for i in range(int(round(len(df_out.index)/12))):
        df1 = df_out.iloc[12*i - delay1:12*(i+1) - delay]
        if i == 0:
            delay1 = delay
        df_out2 = df_out2.append(df1.sum(axis = 0),ignore_index = True)
    return df_out, df_out2

def appenddata(data,row,counter,numrows, index,df,start):
    if counter[0] < 0:
        for i in range(-counter[0]):
            data[len(data)-1].append(row[1])           # Precio Nudo para bloque 1 (24-1)
            data[len(data)-1].append(row[1])           # Precio Nudo para bloque 2 (2-3)
            data[len(data)-1].append(row[2])           # Precio Nudo para bloque 3 (4-5)
            data[len(data)-1].append(row[2])           # Precio Nudo para bloque 4 (6-7)
            data[len(data)-1].append(row[3])           # Precio Nudo para bloque 5 (8-9)
            data[len(data)-1].append(row[3])           # Precio Nudo para bloque 6 (10-11)
            data[len(data)-1].append(row[4])           # Precio Nudo para bloque 7 (12-13)
            data[len(data)-1].append(row[4])          # Precio Nudo para bloque 8 (14-15)
            data[len(data)-1].append(row[5])          # Precio Nudo para bloque 9 (16-17)
            data[len(data)-1].append(row[5])          # Precio Nudo para bloque 10 (18-19)
            data[len(data)-1].append(row[6])          # Precio Nudo para bloque 11 (20-21)
            data[len(data)-1].append(row[6])         # Precio Nudo para bloque 12 (22-23)
            data[len(data)-1].append(row[2])         # Precio Nudo para bloque 13 (6-7)
            data[len(data)-1].append(row[3])         # Precio Nudo para bloque 14 (8-9)
            data[len(data)-1].append(row[3])         # Precio Nudo para bloque 15 (10-11)
            data[len(data)-1].append(row[4])         # Precio Nudo para bloque 16 (12-13)
            data[len(data)-1].append(row[4])         # Precio Nudo para bloque 17 (14-15)
            data[len(data)-1].append(row[5])         # Precio Nudo para bloque 18 (16-17)
            data[len(data)-1].append(row[5])         # Precio Nudo para bloque 19 (18-19)
            data[len(data)-1].append(row[6])         # Precio Nudo para bloque 20 (20-21)
            counter[0] += 1
    if numrows - counter[0] >= 6 and counter[0] <= 240:
        data[len(data)-1].append(row[1])           # Precio Nudo para bloque 1 (24-1)
        data[len(data)-1].append(row[1])           # Precio Nudo para bloque 2 (2-3)
        data[len(data)-1].append(row[2])           # Precio Nudo para bloque 3 (4-5)
        data[len(data)-1].append(row[2])           # Precio Nudo para bloque 4 (6-7)
        data[len(data)-1].append(row[3])           # Precio Nudo para bloque 5 (8-9)
        data[len(data)-1].append(row[3])           # Precio Nudo para bloque 6 (10-11)
        data[len(data)-1].append(row[4])           # Precio Nudo para bloque 7 (12-13)
        data[len(data)-1].append(row[4])          # Precio Nudo para bloque 8 (14-15)
        data[len(data)-1].append(row[5])          # Precio Nudo para bloque 9 (16-17)
        data[len(data)-1].append(row[5])          # Precio Nudo para bloque 10 (18-19)
        data[len(data)-1].append(row[6])          # Precio Nudo para bloque 11 (20-21)
        data[len(data)-1].append(row[6])         # Precio Nudo para bloque 12 (22-23)
        data[len(data)-1].append(row[2])         # Precio Nudo para bloque 13 (6-7)
        data[len(data)-1].append(row[3])         # Precio Nudo para bloque 14 (8-9)
        data[len(data)-1].append(row[3])         # Precio Nudo para bloque 15 (10-11)
        data[len(data)-1].append(row[4])         # Precio Nudo para bloque 16 (12-13)
        data[len(data)-1].append(row[4])         # Precio Nudo para bloque 17 (14-15)
        data[len(data)-1].append(row[5])         # Precio Nudo para bloque 18 (16-17)
        data[len(data)-1].append(row[5])         # Precio Nudo para bloque 19 (18-19)
        data[len(data)-1].append(row[6])         # Precio Nudo para bloque 20 (20-21)

        data[len(data)-1].append(row[1])           # Precio Nudo para bloque 1 (24-1)
        data[len(data)-1].append(row[1])           # Precio Nudo para bloque 2 (2-3)
        data[len(data)-1].append(row[2])           # Precio Nudo para bloque 3 (4-5)
        data[len(data)-1].append(row[2])           # Precio Nudo para bloque 4 (6-7)
        data[len(data)-1].append(row[3])           # Precio Nudo para bloque 5 (8-9)
        data[len(data)-1].append(row[3])           # Precio Nudo para bloque 6 (10-11)
        data[len(data)-1].append(row[4])           # Precio Nudo para bloque 7 (12-13)
        data[len(data)-1].append(row[4])          # Precio Nudo para bloque 8 (14-15)
        data[len(data)-1].append(row[5])          # Precio Nudo para bloque 9 (16-17)
        data[len(data)-1].append(row[5])          # Precio Nudo para bloque 10 (18-19)
        data[len(data)-1].append(row[6])          # Precio Nudo para bloque 11 (20-21)
        data[len(data)-1].append(row[6])         # Precio Nudo para bloque 12 (22-23)
        data[len(data)-1].append(row[2])         # Precio Nudo para bloque 13 (6-7)
        data[len(data)-1].append(row[3])         # Precio Nudo para bloque 14 (8-9)
        data[len(data)-1].append(row[3])         # Precio Nudo para bloque 15 (10-11)
        data[len(data)-1].append(row[4])         # Precio Nudo para bloque 16 (12-13)
        data[len(data)-1].append(row[4])         # Precio Nudo para bloque 17 (14-15)
        data[len(data)-1].append(row[5])         # Precio Nudo para bloque 18 (16-17)
        data[len(data)-1].append(row[5])         # Precio Nudo para bloque 19 (18-19)
        data[len(data)-1].append(row[6])         # Precio Nudo para bloque 20 (20-21)

        data[len(data)-1].append(row[1])           # Precio Nudo para bloque 1 (24-1)
        data[len(data)-1].append(row[1])           # Precio Nudo para bloque 2 (2-3)
        data[len(data)-1].append(row[2])           # Precio Nudo para bloque 3 (4-5)
        data[len(data)-1].append(row[2])           # Precio Nudo para bloque 4 (6-7)
        data[len(data)-1].append(row[3])           # Precio Nudo para bloque 5 (8-9)
        data[len(data)-1].append(row[3])           # Precio Nudo para bloque 6 (10-11)
        data[len(data)-1].append(row[4])           # Precio Nudo para bloque 7 (12-13)
        data[len(data)-1].append(row[4])          # Precio Nudo para bloque 8 (14-15)
        data[len(data)-1].append(row[5])          # Precio Nudo para bloque 9 (16-17)
        data[len(data)-1].append(row[5])          # Precio Nudo para bloque 10 (18-19)
        data[len(data)-1].append(row[6])          # Precio Nudo para bloque 11 (20-21)
        data[len(data)-1].append(row[6])         # Precio Nudo para bloque 12 (22-23)
        data[len(data)-1].append(row[2])         # Precio Nudo para bloque 13 (6-7)
        data[len(data)-1].append(row[3])         # Precio Nudo para bloque 14 (8-9)
        data[len(data)-1].append(row[3])         # Precio Nudo para bloque 15 (10-11)
        data[len(data)-1].append(row[4])         # Precio Nudo para bloque 16 (12-13)
        data[len(data)-1].append(row[4])         # Precio Nudo para bloque 17 (14-15)
        data[len(data)-1].append(row[5])         # Precio Nudo para bloque 18 (16-17)
        data[len(data)-1].append(row[5])         # Precio Nudo para bloque 19 (18-19)
        data[len(data)-1].append(row[6])         # Precio Nudo para bloque 20 (20-21)

        data[len(data)-1].append(row[1])           # Precio Nudo para bloque 1 (24-1)
        data[len(data)-1].append(row[1])           # Precio Nudo para bloque 2 (2-3)
        data[len(data)-1].append(row[2])           # Precio Nudo para bloque 3 (4-5)
        data[len(data)-1].append(row[2])           # Precio Nudo para bloque 4 (6-7)
        data[len(data)-1].append(row[3])           # Precio Nudo para bloque 5 (8-9)
        data[len(data)-1].append(row[3])           # Precio Nudo para bloque 6 (10-11)
        data[len(data)-1].append(row[4])           # Precio Nudo para bloque 7 (12-13)
        data[len(data)-1].append(row[4])          # Precio Nudo para bloque 8 (14-15)
        data[len(data)-1].append(row[5])          # Precio Nudo para bloque 9 (16-17)
        data[len(data)-1].append(row[5])          # Precio Nudo para bloque 10 (18-19)
        data[len(data)-1].append(row[6])          # Precio Nudo para bloque 11 (20-21)
        data[len(data)-1].append(row[6])         # Precio Nudo para bloque 12 (22-23)
        data[len(data)-1].append(row[2])         # Precio Nudo para bloque 13 (6-7)
        data[len(data)-1].append(row[3])         # Precio Nudo para bloque 14 (8-9)
        data[len(data)-1].append(row[3])         # Precio Nudo para bloque 15 (10-11)
        data[len(data)-1].append(row[4])         # Precio Nudo para bloque 16 (12-13)
        data[len(data)-1].append(row[4])         # Precio Nudo para bloque 17 (14-15)
        data[len(data)-1].append(row[5])         # Precio Nudo para bloque 18 (16-17)
        data[len(data)-1].append(row[5])         # Precio Nudo para bloque 19 (18-19)
        data[len(data)-1].append(row[6])         # Precio Nudo para bloque 20 (20-21)

        data[len(data)-1].append(row[1])           # Precio Nudo para bloque 1 (24-1)
        data[len(data)-1].append(row[1])           # Precio Nudo para bloque 2 (2-3)
        data[len(data)-1].append(row[2])           # Precio Nudo para bloque 3 (4-5)
        data[len(data)-1].append(row[2])           # Precio Nudo para bloque 4 (6-7)
        data[len(data)-1].append(row[3])           # Precio Nudo para bloque 5 (8-9)
        data[len(data)-1].append(row[3])           # Precio Nudo para bloque 6 (10-11)
        data[len(data)-1].append(row[4])           # Precio Nudo para bloque 7 (12-13)
        data[len(data)-1].append(row[4])          # Precio Nudo para bloque 8 (14-15)
        data[len(data)-1].append(row[5])          # Precio Nudo para bloque 9 (16-17)
        data[len(data)-1].append(row[5])          # Precio Nudo para bloque 10 (18-19)
        data[len(data)-1].append(row[6])          # Precio Nudo para bloque 11 (20-21)
        data[len(data)-1].append(row[6])         # Precio Nudo para bloque 12 (22-23)
        data[len(data)-1].append(row[2])         # Precio Nudo para bloque 13 (6-7)
        data[len(data)-1].append(row[3])         # Precio Nudo para bloque 14 (8-9)
        data[len(data)-1].append(row[3])         # Precio Nudo para bloque 15 (10-11)
        data[len(data)-1].append(row[4])         # Precio Nudo para bloque 16 (12-13)
        data[len(data)-1].append(row[4])         # Precio Nudo para bloque 17 (14-15)
        data[len(data)-1].append(row[5])         # Precio Nudo para bloque 18 (16-17)
        data[len(data)-1].append(row[5])         # Precio Nudo para bloque 19 (18-19)
        data[len(data)-1].append(row[6])         # Precio Nudo para bloque 20 (20-21)

        data[len(data)-1].append(row[1])           # Precio Nudo para bloque 1 (24-1)
        data[len(data)-1].append(row[1])           # Precio Nudo para bloque 2 (2-3)
        data[len(data)-1].append(row[2])           # Precio Nudo para bloque 3 (4-5)
        data[len(data)-1].append(row[2])           # Precio Nudo para bloque 4 (6-7)
        data[len(data)-1].append(row[3])           # Precio Nudo para bloque 5 (8-9)
        data[len(data)-1].append(row[3])           # Precio Nudo para bloque 6 (10-11)
        data[len(data)-1].append(row[4])           # Precio Nudo para bloque 7 (12-13)
        data[len(data)-1].append(row[4])          # Precio Nudo para bloque 8 (14-15)
        data[len(data)-1].append(row[5])          # Precio Nudo para bloque 9 (16-17)
        data[len(data)-1].append(row[5])          # Precio Nudo para bloque 10 (18-19)
        data[len(data)-1].append(row[6])          # Precio Nudo para bloque 11 (20-21)
        data[len(data)-1].append(row[6])         # Precio Nudo para bloque 12 (22-23)
        data[len(data)-1].append(row[2])         # Precio Nudo para bloque 13 (6-7)
        data[len(data)-1].append(row[3])         # Precio Nudo para bloque 14 (8-9)
        data[len(data)-1].append(row[3])         # Precio Nudo para bloque 15 (10-11)
        data[len(data)-1].append(row[4])         # Precio Nudo para bloque 16 (12-13)
        data[len(data)-1].append(row[4])         # Precio Nudo para bloque 17 (14-15)
        data[len(data)-1].append(row[5])         # Precio Nudo para bloque 18 (16-17)
        data[len(data)-1].append(row[5])         # Precio Nudo para bloque 19 (18-19)
        data[len(data)-1].append(row[6])         # Precio Nudo para bloque 20 (20-21)
        counter[0] += 6
        if counter[0] > 240:
            counter[0] = 240
    if index == df.index[len(df.index)-1]:
        # print("last: ",index)
        while counter[0] != 358- start + 1:
            # print("Counter in: ",counter[0])
            if numrows - counter[0] >= 6:
                data[len(data)-1].append(row[1])           # Precio Nudo para bloque 1 (24-1)
                data[len(data)-1].append(row[1])           # Precio Nudo para bloque 2 (2-3)
                data[len(data)-1].append(row[2])           # Precio Nudo para bloque 3 (4-5)
                data[len(data)-1].append(row[2])           # Precio Nudo para bloque 4 (6-7)
                data[len(data)-1].append(row[3])           # Precio Nudo para bloque 5 (8-9)
                data[len(data)-1].append(row[3])           # Precio Nudo para bloque 6 (10-11)
                data[len(data)-1].append(row[4])           # Precio Nudo para bloque 7 (12-13)
                data[len(data)-1].append(row[4])          # Precio Nudo para bloque 8 (14-15)
                data[len(data)-1].append(row[5])          # Precio Nudo para bloque 9 (16-17)
                data[len(data)-1].append(row[5])          # Precio Nudo para bloque 10 (18-19)
                data[len(data)-1].append(row[6])          # Precio Nudo para bloque 11 (20-21)
                data[len(data)-1].append(row[6])         # Precio Nudo para bloque 12 (22-23)
                data[len(data)-1].append(row[2])         # Precio Nudo para bloque 13 (6-7)
                data[len(data)-1].append(row[3])         # Precio Nudo para bloque 14 (8-9)
                data[len(data)-1].append(row[3])         # Precio Nudo para bloque 15 (10-11)
                data[len(data)-1].append(row[4])         # Precio Nudo para bloque 16 (12-13)
                data[len(data)-1].append(row[4])         # Precio Nudo para bloque 17 (14-15)
                data[len(data)-1].append(row[5])         # Precio Nudo para bloque 18 (16-17)
                data[len(data)-1].append(row[5])         # Precio Nudo para bloque 19 (18-19)
                data[len(data)-1].append(row[6])         # Precio Nudo para bloque 20 (20-21)

                data[len(data)-1].append(row[1])           # Precio Nudo para bloque 1 (24-1)
                data[len(data)-1].append(row[1])           # Precio Nudo para bloque 2 (2-3)
                data[len(data)-1].append(row[2])           # Precio Nudo para bloque 3 (4-5)
                data[len(data)-1].append(row[2])           # Precio Nudo para bloque 4 (6-7)
                data[len(data)-1].append(row[3])           # Precio Nudo para bloque 5 (8-9)
                data[len(data)-1].append(row[3])           # Precio Nudo para bloque 6 (10-11)
                data[len(data)-1].append(row[4])           # Precio Nudo para bloque 7 (12-13)
                data[len(data)-1].append(row[4])          # Precio Nudo para bloque 8 (14-15)
                data[len(data)-1].append(row[5])          # Precio Nudo para bloque 9 (16-17)
                data[len(data)-1].append(row[5])          # Precio Nudo para bloque 10 (18-19)
                data[len(data)-1].append(row[6])          # Precio Nudo para bloque 11 (20-21)
                data[len(data)-1].append(row[6])         # Precio Nudo para bloque 12 (22-23)
                data[len(data)-1].append(row[2])         # Precio Nudo para bloque 13 (6-7)
                data[len(data)-1].append(row[3])         # Precio Nudo para bloque 14 (8-9)
                data[len(data)-1].append(row[3])         # Precio Nudo para bloque 15 (10-11)
                data[len(data)-1].append(row[4])         # Precio Nudo para bloque 16 (12-13)
                data[len(data)-1].append(row[4])         # Precio Nudo para bloque 17 (14-15)
                data[len(data)-1].append(row[5])         # Precio Nudo para bloque 18 (16-17)
                data[len(data)-1].append(row[5])         # Precio Nudo para bloque 19 (18-19)
                data[len(data)-1].append(row[6])         # Precio Nudo para bloque 20 (20-21)

                data[len(data)-1].append(row[1])           # Precio Nudo para bloque 1 (24-1)
                data[len(data)-1].append(row[1])           # Precio Nudo para bloque 2 (2-3)
                data[len(data)-1].append(row[2])           # Precio Nudo para bloque 3 (4-5)
                data[len(data)-1].append(row[2])           # Precio Nudo para bloque 4 (6-7)
                data[len(data)-1].append(row[3])           # Precio Nudo para bloque 5 (8-9)
                data[len(data)-1].append(row[3])           # Precio Nudo para bloque 6 (10-11)
                data[len(data)-1].append(row[4])           # Precio Nudo para bloque 7 (12-13)
                data[len(data)-1].append(row[4])          # Precio Nudo para bloque 8 (14-15)
                data[len(data)-1].append(row[5])          # Precio Nudo para bloque 9 (16-17)
                data[len(data)-1].append(row[5])          # Precio Nudo para bloque 10 (18-19)
                data[len(data)-1].append(row[6])          # Precio Nudo para bloque 11 (20-21)
                data[len(data)-1].append(row[6])         # Precio Nudo para bloque 12 (22-23)
                data[len(data)-1].append(row[2])         # Precio Nudo para bloque 13 (6-7)
                data[len(data)-1].append(row[3])         # Precio Nudo para bloque 14 (8-9)
                data[len(data)-1].append(row[3])         # Precio Nudo para bloque 15 (10-11)
                data[len(data)-1].append(row[4])         # Precio Nudo para bloque 16 (12-13)
                data[len(data)-1].append(row[4])         # Precio Nudo para bloque 17 (14-15)
                data[len(data)-1].append(row[5])         # Precio Nudo para bloque 18 (16-17)
                data[len(data)-1].append(row[5])         # Precio Nudo para bloque 19 (18-19)
                data[len(data)-1].append(row[6])         # Precio Nudo para bloque 20 (20-21)

                data[len(data)-1].append(row[1])           # Precio Nudo para bloque 1 (24-1)
                data[len(data)-1].append(row[1])           # Precio Nudo para bloque 2 (2-3)
                data[len(data)-1].append(row[2])           # Precio Nudo para bloque 3 (4-5)
                data[len(data)-1].append(row[2])           # Precio Nudo para bloque 4 (6-7)
                data[len(data)-1].append(row[3])           # Precio Nudo para bloque 5 (8-9)
                data[len(data)-1].append(row[3])           # Precio Nudo para bloque 6 (10-11)
                data[len(data)-1].append(row[4])           # Precio Nudo para bloque 7 (12-13)
                data[len(data)-1].append(row[4])          # Precio Nudo para bloque 8 (14-15)
                data[len(data)-1].append(row[5])          # Precio Nudo para bloque 9 (16-17)
                data[len(data)-1].append(row[5])          # Precio Nudo para bloque 10 (18-19)
                data[len(data)-1].append(row[6])          # Precio Nudo para bloque 11 (20-21)
                data[len(data)-1].append(row[6])         # Precio Nudo para bloque 12 (22-23)
                data[len(data)-1].append(row[2])         # Precio Nudo para bloque 13 (6-7)
                data[len(data)-1].append(row[3])         # Precio Nudo para bloque 14 (8-9)
                data[len(data)-1].append(row[3])         # Precio Nudo para bloque 15 (10-11)
                data[len(data)-1].append(row[4])         # Precio Nudo para bloque 16 (12-13)
                data[len(data)-1].append(row[4])         # Precio Nudo para bloque 17 (14-15)
                data[len(data)-1].append(row[5])         # Precio Nudo para bloque 18 (16-17)
                data[len(data)-1].append(row[5])         # Precio Nudo para bloque 19 (18-19)
                data[len(data)-1].append(row[6])         # Precio Nudo para bloque 20 (20-21)

                data[len(data)-1].append(row[1])           # Precio Nudo para bloque 1 (24-1)
                data[len(data)-1].append(row[1])           # Precio Nudo para bloque 2 (2-3)
                data[len(data)-1].append(row[2])           # Precio Nudo para bloque 3 (4-5)
                data[len(data)-1].append(row[2])           # Precio Nudo para bloque 4 (6-7)
                data[len(data)-1].append(row[3])           # Precio Nudo para bloque 5 (8-9)
                data[len(data)-1].append(row[3])           # Precio Nudo para bloque 6 (10-11)
                data[len(data)-1].append(row[4])           # Precio Nudo para bloque 7 (12-13)
                data[len(data)-1].append(row[4])          # Precio Nudo para bloque 8 (14-15)
                data[len(data)-1].append(row[5])          # Precio Nudo para bloque 9 (16-17)
                data[len(data)-1].append(row[5])          # Precio Nudo para bloque 10 (18-19)
                data[len(data)-1].append(row[6])          # Precio Nudo para bloque 11 (20-21)
                data[len(data)-1].append(row[6])         # Precio Nudo para bloque 12 (22-23)
                data[len(data)-1].append(row[2])         # Precio Nudo para bloque 13 (6-7)
                data[len(data)-1].append(row[3])         # Precio Nudo para bloque 14 (8-9)
                data[len(data)-1].append(row[3])         # Precio Nudo para bloque 15 (10-11)
                data[len(data)-1].append(row[4])         # Precio Nudo para bloque 16 (12-13)
                data[len(data)-1].append(row[4])         # Precio Nudo para bloque 17 (14-15)
                data[len(data)-1].append(row[5])         # Precio Nudo para bloque 18 (16-17)
                data[len(data)-1].append(row[5])         # Precio Nudo para bloque 19 (18-19)
                data[len(data)-1].append(row[6])         # Precio Nudo para bloque 20 (20-21)

                data[len(data)-1].append(row[1])           # Precio Nudo para bloque 1 (24-1)
                data[len(data)-1].append(row[1])           # Precio Nudo para bloque 2 (2-3)
                data[len(data)-1].append(row[2])           # Precio Nudo para bloque 3 (4-5)
                data[len(data)-1].append(row[2])           # Precio Nudo para bloque 4 (6-7)
                data[len(data)-1].append(row[3])           # Precio Nudo para bloque 5 (8-9)
                data[len(data)-1].append(row[3])           # Precio Nudo para bloque 6 (10-11)
                data[len(data)-1].append(row[4])           # Precio Nudo para bloque 7 (12-13)
                data[len(data)-1].append(row[4])          # Precio Nudo para bloque 8 (14-15)
                data[len(data)-1].append(row[5])          # Precio Nudo para bloque 9 (16-17)
                data[len(data)-1].append(row[5])          # Precio Nudo para bloque 10 (18-19)
                data[len(data)-1].append(row[6])          # Precio Nudo para bloque 11 (20-21)
                data[len(data)-1].append(row[6])         # Precio Nudo para bloque 12 (22-23)
                data[len(data)-1].append(row[2])         # Precio Nudo para bloque 13 (6-7)
                data[len(data)-1].append(row[3])         # Precio Nudo para bloque 14 (8-9)
                data[len(data)-1].append(row[3])         # Precio Nudo para bloque 15 (10-11)
                data[len(data)-1].append(row[4])         # Precio Nudo para bloque 16 (12-13)
                data[len(data)-1].append(row[4])         # Precio Nudo para bloque 17 (14-15)
                data[len(data)-1].append(row[5])         # Precio Nudo para bloque 18 (16-17)
                data[len(data)-1].append(row[5])         # Precio Nudo para bloque 19 (18-19)
                data[len(data)-1].append(row[6])         # Precio Nudo para bloque 20 (20-21)
                counter[0] += 6
            else:
                data[len(data)-1].append(row[1])           # Precio Nudo para bloque 1 (24-1)
                data[len(data)-1].append(row[1])           # Precio Nudo para bloque 2 (2-3)
                data[len(data)-1].append(row[2])           # Precio Nudo para bloque 3 (4-5)
                data[len(data)-1].append(row[2])           # Precio Nudo para bloque 4 (6-7)
                data[len(data)-1].append(row[3])           # Precio Nudo para bloque 5 (8-9)
                data[len(data)-1].append(row[3])           # Precio Nudo para bloque 6 (10-11)
                data[len(data)-1].append(row[4])           # Precio Nudo para bloque 7 (12-13)
                data[len(data)-1].append(row[4])          # Precio Nudo para bloque 8 (14-15)
                data[len(data)-1].append(row[5])          # Precio Nudo para bloque 9 (16-17)
                data[len(data)-1].append(row[5])          # Precio Nudo para bloque 10 (18-19)
                data[len(data)-1].append(row[6])          # Precio Nudo para bloque 11 (20-21)
                data[len(data)-1].append(row[6])         # Precio Nudo para bloque 12 (22-23)
                data[len(data)-1].append(row[2])         # Precio Nudo para bloque 13 (6-7)
                data[len(data)-1].append(row[3])         # Precio Nudo para bloque 14 (8-9)
                data[len(data)-1].append(row[3])         # Precio Nudo para bloque 15 (10-11)
                data[len(data)-1].append(row[4])         # Precio Nudo para bloque 16 (12-13)
                data[len(data)-1].append(row[4])         # Precio Nudo para bloque 17 (14-15)
                data[len(data)-1].append(row[5])         # Precio Nudo para bloque 18 (16-17)
                data[len(data)-1].append(row[5])         # Precio Nudo para bloque 19 (18-19)
                data[len(data)-1].append(row[6])         # Precio Nudo para bloque 20 (20-21)
                counter[0] += 1
            # print("Counter out: ",counter[0])
    return data, counter

def blockpricemaker1(file, sheetname,nodes,start,numrows):
    sheet = xw.Book(file).sheets(sheetname)
    busnames = sheet.range("Q2").expand('down')
    df_prices = pd.DataFrame()
    prices = []
    for i in nodes:
        for j in range(3,len(busnames.value)):
            if i[1] == busnames.value[j].replace(" ",""):
                if i[1] not in df_prices.columns:
                    df_prices.insert(0,i[1],None,allow_duplicates=False)
                    prices.append([])
                    sheet.range("O2").value = busnames.value[j]
                    busprices = sheet.range("I4:O45").value
                    df = pd.DataFrame(busprices)
                    df = df.drop(index = [0,1])
                    counter = [start - df.iloc[0,0].month,0]
                    for index, row in df.iterrows():
                        prices, counter = appenddata(prices,row,counter,numrows,index,df,start)
                    df_prices[i[1]] = prices[len(prices)-1]
    return df_prices


#### MAIN CODE


# print("__________________")
# print("COMPENSACION PMGD")
# print("__________________")
# names = pmgdfinder("C:/Seba/20_93_1513_11032021 - EBS/pnomnd.csv","PMGD", 3)            # Se obtienen nombres para cada central PMGD no contabilizada   

# nodes, genames = nodefinder("C:/Seba/20_93_1513_11032021 - EBS/dbus.csv", names,1)      # se obtiene barra de ubicación de cada PMGD

# genpmgd = datamaker("C:/Seba/20_93_1513_11032021 - EBS/gergnd.csv",3, genames)          # Se obtiene generación Renovable PMGD en GWh

# genxcmg = datamaker("C:/Seba/20_93_1513_11032021 - EBS/ggcmgb.csv",3, genames)          # Se obtiene generación Renovable X CMg  PMGD

# demsen = datamaker("C:/Seba/20_93_1513_11032021 - EBS/demand.csv",3, "SEN")             # Se obtiene Demanda SEN en GWh

# blockprice = blockpricemaker("C:/Seba/20_93_1513_11032021 - EBS/Adicionales/Precio Estabilizado - PEstabilizado Bloques/Cálculo PN final-Base.xlsx",nodes,len(genpmgd.index),3)  # 

# genbloq, gencmgbloq, pricebloq, diffbloq, compbloq = serieaverage(genxcmg, genpmgd, blockprice,nodes,demsen)
# # print(genbloq)

# pricestage = monthprice("C:/Seba/20_93_1513_11032021 - EBS/Adicionales/Precio Estabilizado - PEstabilizado Bloques/Cálculo PN final-Base.xlsx","AInforme_Actual",nodes, [3,2021],358)
# # print(mnprice)
# start = 3
# genstage, genyear = monthyearaverage(genbloq,start)
# genxcmgstage, genxcmgyear = monthyearaverage(gencmgbloq,start)

# print("Generacion por mes\n",genstage)
# print("Generacion x CMg por mes\n",genxcmgstage)
# print("Precio Nudo por mes\n",pricestage)


# diffstage = genstage*pricestage - genxcmgstage
# print("Diferencia por Mes")
# print(diffstage)

# print(os.getcwd())

# exportexcel(genxcmg, demsen, genpmgd, price, genbloq, gencmgbloq, pricebloq, diffbloq, compbloq)      # bloqcomp, monthcomp


# df = pd.read_excel("C:/Users/Seba Morales/Documents/Synex/Cálculo PN final-Base.xlsx",sheet_name = "AInforme_OtrasBarras_Nuevo", skiprows = [0],  usecols= "N:O",nrows=2)
# print(df.columns)

nodes = [['EOL_Lebu', 'Hualpen154'], ['EOL_Raki', 'Hualpen154'], ['EOL_Huajache', 'Charrua220'], ['EOL_LasPe�as', 'Charrua220'], ['EOL_Lebu3', 'Hualpen154'], ['SFV_Solarpa2', 'PAlmonte110'], ['SFV_PMGD_CEN', 'CNavia110'], ['SFV_PMGD_NG', 'Crucero220'], ['SFV_PMGD_NOR', 'PAzucar110']]
# sheet = xw.Book("C:/Users/Seba Morales/Documents/Synex/Cálculo PN final-Base.xlsx").sheets("AInforme_OtrasBarras_Nuevo")
# print(sheet.range("O2").value)
# sheet.range("O2").value = "Charrua220                "
# df = pd.DataFrame(sheet.range("I4:O45").value)
print(nodes)

blockpricemaker1("C:/Users/Seba Morales/Documents/Synex/Cálculo PN final-Base.xlsx","AInforme_OtrasBarras_Nuevo",nodes,3,358)



# ['Likananta220' 'Parinas500' 'BajaCordi220' 'BajaCordi110' 'Hualqui220',
# 'Dichato220' 'NvaCauque220' 'NvaNirivi220' 'Mataquito220' 'NvaCasabl220' 'LaPolvora220' 'JMA220' 'Centinela220' 'Nogales500' 'Lagunas500']