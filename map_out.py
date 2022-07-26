
import pandas as pd
import requests
import re
import plotly.graph_objects as go
import math
import itertools
import os
from time import process_time




def API_request(url_buses,columns_droped):
    response = requests.get(url_buses)
    if response.status_code != 200:
        print("error reading ",url_buses," :",response.status_code)
    else:
        df = pd.DataFrame(response.json())
        df = df.drop(columns = [df.columns[i] for i in columns_droped])
        return df

def API_geocode(df, url_geocode):
    printProgressBar(0, df.shape[0], prefix = 'Progress:', suffix = 'Complete', length = 50)
    df["East_UTM"] = ""
    df["North_UTM"] = ""
    df["Zone"] = ""
    for index, row in df.iterrows():
        df_geocode = API_request(url_geocode.format(df.loc[index,"id"]),[])
        df_geocode = df_geocode.T
        df_geocode = df_geocode.iloc[:,[7,9,16]]
        geo_code_utm = [df_geocode[df_geocode[df_geocode.columns[0]] == "Coordenada Este"][df_geocode.columns[1]][0], 
                        df_geocode[df_geocode[df_geocode.columns[0]] == "Coordenada Norte"][df_geocode.columns[1]][0],
                        df_geocode[df_geocode[df_geocode.columns[0]] == "Zona o Huso [Ej: 18H-19J etc.]"][df_geocode.columns[1]][0]]
        df.loc[index,"East_UTM"] = geo_code_utm[0]
        df.loc[index,"North_UTM"] = geo_code_utm[1]
        df.loc[index,"Zone"] = geo_code_utm[2]
        printProgressBar(index + 1, df.shape[0], prefix = 'Progress:', suffix = 'Complete', length = 50)
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

def Map_online(lat_center,long_center,text_center,lat_bus,long_bus,bus_names):

    fig = go.Figure(go.Scattermapbox(
    mode = "markers",
    lon = [long_center],
    lat = [lat_center],
    marker = {'size': 10},
    text=[text_center]))

    fig.add_trace(go.Scattermapbox(
    mode = "markers",
    lon = long_bus,
    lat = lat_bus,
    marker = {'size': 10},
    text = bus_names))

    fig.add_trace(go.Scattermapbox(
    mode = "markers",
    lon = long_bus,
    lat = lat_bus,
    marker = {'size': 10},
    text = bus_names))

    fig.update_layout(
    margin ={'l':0,'t':0,'b':0,'r':0},
    mapbox = {
        'center': {'lon': 10, 'lat': 10},
        'style': "stamen-terrain",
        'center': {'lon': -20, 'lat': -20},
        'zoom': 1})

    fig.show()

def Map_lines_online(lat_center,long_center,text_center,lat_bus,long_bus,bus_names,bus_demand,data_lines):
    fig = go.Figure(go.Scattermapbox(
    mode = "markers",
    lon = [long_center],
    lat = [lat_center],
    marker = {'size': 20},
    text=[text_center],
    name = text_center))

    for index,row in data_lines.iterrows():
        fig.add_trace(go.Scattermapbox(
            mode = "lines",
            lon = [data_lines.loc[index, "Lon_Or"], data_lines.loc[index, "Lon_De"]],
            lat = [data_lines.loc[index, "Lat_Or"], data_lines.loc[index, "Lat_De"]],
            showlegend = False,
            line = {"color": 'rgb(255, 0, 0)'},
            # text = data_lines.loc[index, "Name"]
        ))
    
    fig.add_trace(go.Scattermapbox(
    mode = "markers",
    lon = long_bus,
    lat = lat_bus,
    marker = {'size': 20},
    text = bus_names,
    hovertext = bus_names + ": " + bus_demand,
    name = "Synex SEN Representation"))

    fig.update_layout(
    margin ={'l':0,'t':0,'b':0,'r':0},
    mapbox = {
        'center': {'lon': 10, 'lat': 10},
        'style': "stamen-terrain",
        'center': {'lon': -20, 'lat': -20},
        'zoom': 1})

    fig.show()

def Map_online_mixed(lat_center,long_center,text_center,lat_bus,long_bus,bus_names,lat_bus_synex,long_bus_synex,bus_names_synex):

    fig = go.Figure(go.Scattermapbox(
    mode = "markers",
    lon = [long_center],
    lat = [lat_center],
    marker = {'size': 10},
    text=[text_center]))

    fig.add_trace(go.Scattermapbox(
    mode = "markers",
    lon = long_bus,
    lat = lat_bus,
    marker = {'size': 10},
    text = bus_names))

    fig.add_trace(go.Scattermapbox(
    mode = "markers",
    lon = long_bus_synex,
    lat = lat_bus_synex,
    marker = {'size': 10},
    text = bus_names_synex))

    fig.update_layout(
    margin ={'l':0,'t':0,'b':0,'r':0},
    mapbox = {
        'center': {'lon': 10, 'lat': 10},
        'style': "stamen-terrain",
        'center': {'lon': -20, 'lat': -20},
        'zoom': 1})

    fig.show()

def utmToLatLong(utmNorthing, utmEasting, utmZone):
    eastingOffset = 500000.0
    northingOffset = 10000000.0
    k0 = 0.9996
    equatorialRadius = 6378137.0
    eccSquared = 0.006694380023
    eccPrimeSquared = eccSquared / (1 - eccSquared)
    e1 = (1 - math.sqrt(1 - eccSquared)) / (1 + math.sqrt(1 - eccSquared));
    rad2deg = 180.0/math.pi

    # Casts input from string to floats or ints
    # Removes 500,000 metre offset for longitude
    xUTM = float(utmEasting) - eastingOffset
    yUTM = float(utmNorthing) - northingOffset
    zoneNumber = int(utmZone)

    # This line below is for debug purposes only, remove for batch processes.
    # print('The input is: ' + str(utmEasting) + 'm E, ' + str(utmNorthing) + 'm N in NAD83 UTM Zone ' + str(utmZone) + '\n')

    # Finds the origin longitude for the zone
    lonOrigin = (zoneNumber - 1) * 6 - 180 + 3 # +3 puts in zone centre

    M = yUTM / k0 #This finds the meridional arc
    mu = M / (equatorialRadius * (1- eccSquared / 4 - 3 * eccSquared * eccSquared / 64 -5 * eccSquared * eccSquared * eccSquared /256))

    # Calculates the footprint latitude
    phi1Rad = mu + (3 * e1 / 2 - 27 * e1 * e1 * e1 /32) * math.sin(2*mu) + ( 21 * e1 * e1 / 16 - 55 * e1 * e1 * e1 * e1 / 32) * math.sin( 4 * mu) + (151 * e1 * e1 * e1 / 96) * math.sin(6 * mu)
    phi1 = phi1Rad * rad2deg

    # Variables for conversion equations
    N1 = equatorialRadius / math.sqrt( 1 - eccSquared * math.sin(phi1Rad) *  math.sin(phi1Rad))
    T1 = math.tan(phi1Rad) * math.tan(phi1Rad)
    C1 = eccPrimeSquared * math.cos(phi1Rad) * math.cos(phi1Rad)
    R1 = equatorialRadius * (1 - eccSquared) / math.pow(1 - eccSquared * math.sin(phi1Rad) * math.sin(phi1Rad), 1.5)
    D = xUTM / (N1 * k0)

    # Calculate latitude, in decimal degrees
    lat = phi1Rad - ( N1 * math.tan(phi1Rad) / R1) * (D * D / 2 - (5 + 3 * T1 + 10 * C1 - 4 * C1 * C1 - 9 * eccPrimeSquared) * D * D * D * D / 24 + (61 + 90 * T1 + 298 * C1 + 45 * T1 * T1 - 252 * eccPrimeSquared - 3 * C1 * C1) * D * D * D * D * D * D / 720)
    lat = lat * rad2deg
    
    # Calculate longitude, in decimal degrees
    lon = (D - (1 + 2 * T1 + C1) * D * D * D / 6 + (5 - 2 * C1 + 28 * T1 - 3 * C1 * C1 + 8 * eccPrimeSquared + 24 * T1 * T1) * D * D * D * D * D / 120) / math.cos(phi1Rad)
    lon = lonOrigin + lon * rad2deg

    # Print function below is for debug purposes
    # Note: THIS IS THE LOCATION WHERE THE NUMBERS ARE ROUNDED TO 5 DECIMAL PLACES
    # print( "Lat: " + str(round(lat, 5)) + ", Long: " + str(round(lon,5)))
    
    return lat, lon

def readexcel(file,rows_skiped, colmns):
    if colmns == []:
        df = pd.DataFrame(pd.read_excel(file,skiprows=rows_skiped))
    else:
        df = pd.DataFrame(pd.read_excel(file,usecols= colmns,skiprows=rows_skiped))
    print(df)
    return df

def read_sub_data_excel(file,rows_skiped, colmns,header_rows):
    if colmns == []:
        df = pd.DataFrame(pd.read_excel(file,skiprows=rows_skiped, header = header_rows))
    else:
        df = pd.DataFrame(pd.read_excel(file,usecols= colmns,skiprows=rows_skiped, header = header_rows))
    print(df)
    return df

def measurement_data_loader(meter_info_root, data_meter_root):   
    # Meter information
    socket = readexcel(meter_info_root,0,[3,6,8,9,10,12,13])
    socket = socket.sort_values(by = ["Subestación"], ignore_index = True)
    socket["Descripción_Región"] = socket["Descripción_Región"].fillna(value = "Otros")
    socket["Región"] = socket["Región"].fillna(0)
    socket.head(10)

    # Init dataframe for meters
    data_index = [[2021],[i for i in range(1,13)],[i for i in range(1,32)],[i for i in range(1,25)]]
    list_name = socket["Medidor"].drop_duplicates()

    list_name = list(itertools.chain.from_iterable(itertools.repeat(x,4) for x in list_name))
    suffix = ["RA","RR","IA","IR"]*int(len(list_name)/4)
    for index in range(len(list_name)):
        list_name[index] = list_name[index] + "_" + suffix[index]
    df = pd.DataFrame(columns = list_name, index = pd.MultiIndex.from_product(data_index, names = ["AÑO","MES","DIA","FIN INTERVALO"]))

    # iterate over excel files and add data
    for year in [2021]:
        for month in [1,2,3,5,6,7,8,9,10,11,12]: # [1,2,3,5,6,7,8,9,10,11,12]
            for filename in os.listdir(data_meter_root.format(year,month)):
                f = os.path.join(data_meter_root.format(year,month),filename)
                if os.path.isfile(f) and f[-4:] == "xlsx":
                    print(f)    
                    df_meter = read_sub_data_excel(f,[0,1,2,3,4,6,7,8,9,10,11],[],[0])
                    index_name = ["AÑO","MES","DIA","HORA UTC","INICIO INTERVALO","FIN INTERVALO"]
                    df_meter = df_meter.rename(columns = {df_meter.columns[i]:index_name[i] for i in range(6)})
                    df_meter = df_meter.drop(df_meter[(df_meter["AÑO"] != year) & (df_meter["MES"] != month)].index)
                    df_meter = df_meter.set_index(["AÑO","MES","DIA","FIN INTERVALO"])
                    df_meter = df_meter.drop(columns = ["HORA UTC","INICIO INTERVALO"])
                    print(df_meter)
                    
                    last_name = ""
                    for id in range(len(df_meter.columns)):
                        if df_meter.columns[id] in socket["Medidor"].drop_duplicates().values:
                            # print(id)
                            last_name = df_meter.columns[id]
                        elif df_meter.columns[id][:7] != "Unnamed":
                            last_name = df_meter.columns[id]
                            print(last_name)
                        df_meter = df_meter.rename(columns= {df_meter.columns[id]:last_name + "_" + suffix[id]})
                        df.loc[df_meter.index,df_meter.columns[id]] = df_meter.loc[df_meter.index,df_meter.columns[id]]
                    # df.loc[df_meter.index.to_list(),df_meter.columns.to_list()] = df_meter
                    # print("next file")
    return df

def year_month_sum(data):
    sum_year_df = data.sum(axis = 0)
    sum_month_df = pd.DataFrame(index = [i for i in range(1,12+1)], columns = data.columns)
    for index, row in sum_month_df.iterrows():
        sum_month_df.loc[index,:] = data.loc[(2021,index)].sum(axis = 0)
    data_index = [[i for i in range(1,3)], [i for i in range(1,25)]]
    average_day_df = pd.DataFrame(index = pd.MultiIndex.from_product(data_index, names = ["MES","HORA"]), columns= data.columns)
    # for index, row in average_day_df.iterrows():
    #     average_day_df.loc[index,:] = average_day_df.loc[:]
    return sum_year_df, sum_month_df, average_day_df

def readcsv(file,rows_skiped, colmns):
    if colmns == []:
        df = pd.DataFrame(pd.read_csv(file,skiprows= rows_skiped))
    else:
        df = pd.DataFrame(pd.read_csv(file,usecols= colmns,skiprows=rows_skiped))
    print(df)
    return df

def subs_adder(socket,year_data, month_data):
    year_subestation = pd.DataFrame(columns = socket["Subestación"].drop_duplicates(), index = [2021])
    year_subestation = year_subestation.fillna(0)
    print("Adding year meter data by substation...")
    for medidor in year_data.index:
        if isinstance(year_data.loc[medidor],float) == True or isinstance(year_data.loc[medidor],int) == True:
            year_subestation.loc[2021,socket[socket["Medidor"] == medidor]["Subestación"]] += year_data.loc[medidor]
        else:
            year_subestation.loc[2021,socket[socket["Medidor"] == medidor]["Subestación"]] += year_data.loc[medidor][0]

    month_substation = pd.DataFrame(columns= socket["Subestación"].drop_duplicates(), index = [1,2,3,4,5,6,7,8,9,10,11,12])
    print("Adding month meter data by substation...")
    month_substation = month_substation.fillna(0)
    for medidor in year_data.index:
        for month in month_substation.index:
            if isinstance(month_data.loc[month,medidor],float) == True or isinstance(month_data.loc[month,medidor],int) == True:
                month_substation.loc[month,socket[socket["Medidor"] == medidor]["Subestación"]] += month_data.loc[month,medidor]
            else:
                month_substation.loc[month,socket[socket["Medidor"] == medidor]["Subestación"]] += month_data.loc[month,medidor][0]
    return year_subestation, month_substation

def exportexcel(listnames, file, adderlist, *args):
    t0 = exectime("Exporting data to: "+file, True, 0)
    printProgressBar(0, len(args), prefix = 'Progress:', suffix = 'Complete', length = 50)
    with pd.ExcelWriter(file, engine = 'xlsxwriter') as writer:
        for i in range(len(args)):
            args[i].to_excel(writer, sheet_name = listnames[i])
            # plotexceladder(listnames[i], writer, adderlist, [0, 1, len(args[i].index), len(args[i].columns)])
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

#### MAIN CODE



northlat = -17.29
southlat = -56.3212
eastlong = -67.0428
westlong = -75.38

google_key = "AIzaSyD1w-BSfdInI2NHZ01IUZ35nbUNhAUjdTE"

Bus = API_request("https://api-infotecnica.coordinador.cl/v1/subestaciones/", [1,2,3,4,5,7,8,9,10,11,12,13,14])

bar = API_geocode(Bus,"https://api-infotecnica.coordinador.cl/v1/subestaciones/{0}/fichas-tecnicas/general/")
bar = bar.fillna(value="0.0")

sub_extended = API_request("https://api-infotecnica.coordinador.cl/v1/subestaciones/extended/", [2,3,4,5,8,9,10,11,12,13,14,15,16,19,20,21,22,23,24])
sub_extended = sub_extended.fillna(value = 0)

for index, row in bar.iterrows():
    if "," in bar.loc[index,"East_UTM"]:
        bar.loc[index,"East_UTM"] = bar.loc[index,"East_UTM"].replace(",",".")
    if not isinstance(bar.loc[index,"East_UTM"],float):
        if re.findall("\d+\.\d+",bar.loc[index,"East_UTM"]) != []:
            bar.loc[index,"East_UTM"] = re.findall("\d+\.\d+",bar.loc[index,"East_UTM"])[0]
    if "," in bar.loc[index,"North_UTM"]:
        bar.loc[index,"North_UTM"] = bar.loc[index,"North_UTM"].replace(",",".")
    if not isinstance(bar.loc[index,"North_UTM"],float):
        if re.findall("\d+\.\d+",bar.loc[index,"North_UTM"]) != []:
            bar.loc[index,"North_UTM"] = re.findall("\d+\.\d+",bar.loc[index,"North_UTM"])[0]
    if bar.loc[index,"North_UTM"] == "" or bar.loc[index,"North_UTM"] == "0.0":
        bar.loc[index,"North_UTM"] = 0.0
    if bar.loc[index,"East_UTM"] == "" or bar.loc[index,"East_UTM"] == "0.0":
        bar.loc[index,"East_UTM"] = 0.0
    if not isinstance(bar.loc[index,"East_UTM"],float) and bar.loc[index,"East_UTM"][0].isalpha():
        bar.loc[index,"East_UTM"] = 0.0
        bar.loc[index,"North_UTM"] = 0.0
    if re.findall("\d+",bar.loc[index,"Zone"]) == []:
        bar.loc[index,"Zone"] = "19"
    elif re.findall("\d+",bar.loc[index,"Zone"])[0] == "18" or re.findall("\d+",bar.loc[index,"Zone"])[0] == "19":
        bar.loc[index,"Zone"] = re.findall("\d+",bar.loc[index,"Zone"])[0]
    else:
        bar.loc[index,"Zone"] = "18"

bar["East_UTM"].astype(float)
bar["North_UTM"].astype(float)
bar["Zone"].astype(float)

for index, row in sub_extended.iterrows():
    if (sub_extended.loc[index,"longitud"] == 0 or sub_extended.loc[index,"latitud"] == 0) and len(bar[bar["nombre"] == sub_extended.loc[index,"nombre"]]["East_UTM"]) !=0:
        if float(bar[bar["nombre"] == sub_extended.loc[index,"nombre"]]["North_UTM"]) !=0:
            lat,long = utmToLatLong(float(bar[bar["nombre"] == sub_extended.loc[index,"nombre"]]["North_UTM"]), float(bar[bar["nombre"] == sub_extended.loc[index,"nombre"]]["East_UTM"]), float(bar[bar["nombre"] == sub_extended.loc[index,"nombre"]]["Zone"]))
            if lat < northlat and lat > southlat:
                sub_extended.loc[index,"latitud"] = lat
            if long < eastlong and long > westlong:
                sub_extended.loc[index,"longitud"] = long

Map_online(-33.45694,-70.64827,"Santiago",sub_extended["latitud"],sub_extended["longitud"],sub_extended["nombre"])



## Synex coordinator location

demxba = readcsv("../../Transmision/dem x barra anual.csv",0,[])
demxba = demxba.rename(columns = {demxba.columns[0]:"ETAPA"})
demxba = demxba.set_index("ETAPA")
demxba = demxba.loc[2021,:]

dgbus = readcsv("../../Transmision/dgbus.csv",0,[])
dbus = readcsv("../../Transmision/dbus.csv",1,[0,1])
dgbus["Name"] = ""
dgbus["Demand"] = ""
for index, row in dgbus.iterrows():
    dgbus.loc[index, "Name"] = dbus.loc[dbus["!Code"] == dgbus.loc[index,"!Bus code"],"Name"].values[0]
    if dgbus.loc[index, "Name"] in demxba.index:
        dgbus.loc[index, "Demand"] = demxba.loc[dgbus.loc[index, "Name"]]
    else:
        dgbus.loc[index, "Demand"] = "0%"
dgbus.loc[1, "Name"] in demxba.index

dcirc = readcsv("../../Transmision/dcirc.csv",0,[0,1,2,3,4])
dcirc = dcirc.rename(columns = {"#BOR.  ":"BOR","#BDE.   ":"BDE",".RESIS":"R","REACT":"X","Nome........":"Name"})
dcirc["Lat_Or"] = 0.0
dcirc["Lon_Or"] = 0.0
dcirc["Lat_De"] = 0.0
dcirc["Lon_De"] = 0.0
for index, row in dcirc.iterrows():
    dcirc.loc[index,"Lat_Or"] = dgbus[dgbus["!Bus code"] == dcirc.loc[index,"BOR"]]["latitude"].values[0]
    dcirc.loc[index,"Lon_Or"] = dgbus[dgbus["!Bus code"] == dcirc.loc[index,"BOR"]]["longitude"].values[0]
    dcirc.loc[index,"Lat_De"] = dgbus[dgbus["!Bus code"] == dcirc.loc[index,"BDE"]]["latitude"].values[0]
    dcirc.loc[index,"Lon_De"] = dgbus[dgbus["!Bus code"] == dcirc.loc[index,"BDE"]]["longitude"].values[0]


Map_lines_online(-33.45694,-70.64827,"Santiago",dgbus["latitude"],dgbus["longitude"],dgbus["Name"], dgbus["Demand"], dcirc)

Map_online(-33.45694,-70.64827,"Santiago",dgbus["latitude"],dgbus["longitude"],dgbus["Name"])

Map_online_mixed(-33.45694,-70.64827,"Santiago",dgbus["latitude"],dgbus["longitude"],dgbus["Name"],sub_extended["latitud"],sub_extended["longitud"],sub_extended["nombre"])
