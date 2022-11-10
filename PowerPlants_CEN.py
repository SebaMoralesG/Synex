import pandas as pd
import requests
import re
from time import process_time
import math
import plotly.graph_objects as go


def Map_online(lat_center,long_center,text_center,lat_bus,long_bus,bus_names):

    fig = go.Figure(go.Scattermapbox(
    mode = "markers",
    lon = [long_center],
    lat = [lat_center],
    marker = {'size': 10},
    text=[text_center],
    name = text_center))

    fig.add_trace(go.Scattermapbox(
    mode = "markers",
    lon = long_bus,
    lat = lat_bus,
    marker = {'size': 10},
    text = bus_names,
    hovertext = bus_names,
    name = "SEN power plants"))

    fig.update_layout(
    margin ={'l':0,'t':0,'b':0,'r':0},
    mapbox = {
        'center': {'lon': 10, 'lat': 10},
        'style': "stamen-terrain",
        'center': {'lon': -20, 'lat': -20},
        'zoom': 1})

    fig.show()

def API_request(url):
    response = requests.request("GET", url)
    if response.status_code != 200:
        print(response.status_code)
        return response.status_code
    else:
        df = pd.DataFrame(response.json())
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

def readfile(file,rows_skiped, colmns, sheet_name: str | int | list | None = 0):
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


power_plants = API_request("https://api-infotecnica.coordinador.cl/v1/centrales")
power_plants = power_plants.iloc[:,[0,9,10,14]]
power_plants["Bus"] = ""
power_plants["Net Capacity"] = 0
power_plants["Minimum Capacity"] = 0
power_plants["COD"] = ""
power_plants["Technology"] = ""
power_plants["Fuel"] = ""
power_plants["Subtype"] = ""
power_plants["Coord E"] = ""
power_plants["Coord W"] = ""
power_plants["Zone"] = ""
power_plants["Lat"] = 0
power_plants["Long"] = 0

t0 = exectime("Buses data",True,0)
printProgressBar(0, power_plants.index[-1] + 1, prefix = 'Progress:', suffix = 'Complete', length = 50)
for index, row in power_plants.iterrows():
    id = power_plants.loc[index,"id"]
    power_plants_data = API_request(f"https://api-infotecnica.coordinador.cl/v1/centrales/{id}/fichas-tecnicas/general")
    power_plants_data = power_plants_data.iloc[[7,9],[1,3,4,5,6,7,8,12,14,15,17,18,19,27,28,29]].T

    if power_plants_data.iloc[15,1] == None or re.findall("\d+", power_plants_data.iloc[15,1]) == []:
        power_plants_data.iloc[15,1] = "19"
    else:
        power_plants_data.iloc[15,1] = re.findall("\d+", power_plants_data.iloc[15,1])[0]

    if power_plants_data.iloc[14,1] == "" or power_plants_data.iloc[14,1] == None or re.findall("\d+\.\d+",power_plants_data.iloc[14,1].replace(",",".")) == []:
        power_plants_data.iloc[14,1] = "0"
    else:
        power_plants_data.iloc[14,1] = re.findall("\d+\.\d+",power_plants_data.iloc[14,1].replace(",","."))[0]

    if power_plants_data.iloc[13,1] == "" or power_plants_data.iloc[13,1] == None or re.findall("\d+\.\d+",power_plants_data.iloc[13,1].replace(",",".")) == []:
        power_plants_data.iloc[13,1] = "0"
    else:
        power_plants_data.iloc[13,1] = re.findall("\d+\.\d+",power_plants_data.iloc[13,1].replace(",","."))[0]

    lat, long = utmToLatLong(float(power_plants_data.iloc[14,1]), float(power_plants_data.iloc[13,1]),int(power_plants_data.iloc[15,1]))

    power_plants.loc[index,"Bus"] = power_plants_data.iloc[1,1]
    power_plants.loc[index,"Net Capacity"] = power_plants_data.iloc[4,1].replace("No aplica","0").replace(",",".")
    power_plants.loc[index,"Minimum Capacity"] = power_plants_data.iloc[5,1].replace("No aplica","0").replace(",",".")
    power_plants.loc[index,"COD"] = power_plants_data.iloc[6,1]
    power_plants.loc[index,"Technology"] = power_plants_data.iloc[7,1]
    power_plants.loc[index,"Fuel"] = power_plants_data.iloc[8,1]
    power_plants.loc[index,"Subtype"] = power_plants_data.iloc[9,1]
    power_plants.loc[index,"Coord E"] = power_plants_data.iloc[13,1]
    power_plants.loc[index,"Coord W"] = power_plants_data.iloc[14,1]
    power_plants.loc[index,"Zone"] = power_plants_data.iloc[15,1]
    power_plants.loc[index,"Long"] = long
    power_plants.loc[index,"Lat"] = lat
    printProgressBar(index + 1, power_plants.index[-1] + 1, prefix = 'Progress:', suffix = 'Complete', length = 50)

t0 = exectime("Buses data",False,t0)


power_plants.loc[:,"Minimum Capacity"] = power_plants.loc[:,"Minimum Capacity"].str.lower().replace("no aplica","0").replace(",",".").replace("","0")
power_plants.loc[:,"Net Capacity"] = power_plants.loc[:,"Net Capacity"].str.lower().replace("no aplica","0").replace(",",".").replace("","0")
power_plants.loc[:,"Coord W"] = power_plants.loc[:,"Coord W"].astype(float)
power_plants.loc[:,"Coord E"] = power_plants.loc[:,"Coord E"].astype(float)
power_plants.loc[:,"Zone"] = power_plants.loc[:,"Zone"].astype(int)

Map_online(-33.45694,-70.64827,"Santiago", power_plants["Lat"], power_plants["Long"],power_plants["nombre"])

print("Reading API-Infotecnica: Buses")
buses = API_request("https://api-infotecnica.coordinador.cl/v1/subestaciones")



print("exporting file...")
with pd.ExcelWriter("PowerPlants_CEN.xlsx", engine = 'xlsxwriter') as writer:
    power_plants.to_excel(writer, sheet_name = "PowerPlants")
    buses.to_excel(writer, sheet_name = "Substations")
    print("File CEN Power plants exported to: main root.")


power_plants = readfile("PowerPlants_CEN.xlsx", 0, [], sheet_name = "PowerPlants")
power_plants = power_plants.drop(columns = "Unnamed: 0")
