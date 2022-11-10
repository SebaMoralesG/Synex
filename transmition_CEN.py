import pandas as pd
import requests
import re
from time import process_time
import math


def GET_API_SIP(url, token, date: str | None = None, offset = 0, limit = 20):
    headers = {'Authorization': f'Token {token}'}
    if date == None:
        response = requests.request("GET", url + f"?limit={limit}&offset={offset}", headers = headers)
    else:
        response = requests.request("GET", url + f"?limit={limit}&offset={offset}&fecha={date}", headers = headers)
    if response.status_code != 200:
        print(response.status_code)
        return response.status_code
    else:
        df = pd.DataFrame(response.json()["results"])
        print(df)
        return df

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


## CEN Buses

print("Reading API-Infotecnica: Buses")
buses = API_request("https://api-infotecnica.coordinador.cl/v1/subestaciones")
buses = buses.iloc[:,[0,4,5,6]]
buses["Coord E"] = ""
buses["Coord W"] = ""
buses["Zone"] = ""
t0 = exectime("Buses data",True,0)
printProgressBar(0, buses.index[-1] + 1, prefix = 'Progress:', suffix = 'Complete', length = 50)
for index, row in buses.iterrows():
    bus_id = buses.loc[index,"id"]
    buses_data = API_request(f"https://api-infotecnica.coordinador.cl/v1/subestaciones/{bus_id}/fichas-tecnicas/general")
    buses_data = buses_data.iloc[[7,9],[1,2,3,16,17,18,19]].T

    if buses_data.iloc[6,1] == None or re.findall("\d+", buses_data.iloc[6,1]) == []:
        buses_data.iloc[6,1] = "19"
    else:
        buses_data.iloc[6,1] = re.findall("\d+", buses_data.iloc[6,1])[0]

    if buses_data.iloc[4,1] == "" or buses_data.iloc[4,1] == None or re.findall("\d+\.\d+",buses_data.iloc[4,1].replace(",",".")) == []:
        buses_data.iloc[4,1] = "0"
    else:
        buses_data.iloc[4,1] = re.findall("\d+\.\d+",buses_data.iloc[4,1].replace(",","."))[0]

    if buses_data.iloc[5,1] == "" or buses_data.iloc[5,1] == None or re.findall("\d+\.\d+",buses_data.iloc[5,1].replace(",",".")) == []:
        buses_data.iloc[5,1] = "0"
    else:
        buses_data.iloc[5,1] = re.findall("\d+\.\d+",buses_data.iloc[5,1].replace(",","."))[0]

    buses.loc[index, "Coord E"] = buses_data.iloc[4,1]
    buses.loc[index, "Coord W"] = buses_data.iloc[5,1]
    buses.loc[index, "Zone"] = buses_data.iloc[6,1]
    printProgressBar(index + 1, buses.index[-1] + 1, prefix = 'Progress:', suffix = 'Complete', length = 50)

buses.loc[:,"Coord E"] = buses.loc[:,"Coord E"].astype(float)
buses.loc[:,"Coord W"] = buses.loc[:,"Coord W"].astype(float)
buses.loc[:,"Zone"] = buses.loc[:,"Zone"].astype(int)
t0 = exectime("Buses data",False,t0)

## Lines data
print("Reading API-Infotecnica: Lines")
lines = API_request("https://api-infotecnica.coordinador.cl/v1/secciones-tramos")
lines = lines.iloc[:,[0,1,2,5,6,7,10,11]]
lines["linea_nombre"] = lines["linea_nombre"].str.upper()
lines["OrBus"] = ""
lines["DesBus"] = ""
lines["kV"] = 0
lines["km"] = 0
lines["R+"] = 0
lines["X+"] = 0
lines["R0"] = 0
lines["X0"] = 0

t0 = exectime("Lines data",True,0)
printProgressBar(0, lines.index[-1] + 1, prefix = 'Progress:', suffix = 'Complete', length = 50)
for index, row in lines.iterrows():
    line_id = lines.loc[index,"id"]
    lines_data = API_request(f"https://api-infotecnica.coordinador.cl/v1/secciones-tramos/{line_id}/fichas-tecnicas/general")
    lines_data = lines_data.iloc[[7,8,9,13],[1,2,3,4,6,7]].T
    lines_data = lines_data.fillna("0")
    if re.findall("\d+KV",lines.loc[index,"linea_nombre"]) == [] and re.findall("\d+\ KV",lines.loc[index,"linea_nombre"]) != []:
        if "–" in lines.loc[index,"linea_nombre"].replace(re.findall("\d+\ KV",lines.loc[index,"linea_nombre"])[0],""):
            lines.loc[index,"OrBus"] = lines.loc[index,"linea_nombre"].replace(re.findall("\d+\ KV",lines.loc[index,"linea_nombre"])[0],"").split("–",2)[0].strip()
            lines.loc[index,"DesBus"] = lines.loc[index,"linea_nombre"].replace(re.findall("\d+\ KV",lines.loc[index,"linea_nombre"])[0],"").split("–",2)[1].strip()
        else:
            lines.loc[index,"OrBus"] = lines.loc[index,"linea_nombre"].replace(re.findall("\d+\ KV",lines.loc[index,"linea_nombre"])[0],"").split("-",2)[0].strip()
            lines.loc[index,"DesBus"] = lines.loc[index,"linea_nombre"].replace(re.findall("\d+\ KV",lines.loc[index,"linea_nombre"])[0],"").split("-",2)[1].strip()
    elif len(re.findall("\d+KV",lines.loc[index,"linea_nombre"])) == 1:
        if "–" in lines.loc[index,"linea_nombre"].replace(re.findall("\d+KV",lines.loc[index,"linea_nombre"])[0],""):
            lines.loc[index,"OrBus"] = lines.loc[index,"linea_nombre"].replace(re.findall("\d+KV",lines.loc[index,"linea_nombre"])[0],"").split("–",2)[0].strip()
            lines.loc[index,"DesBus"] = lines.loc[index,"linea_nombre"].replace(re.findall("\d+KV",lines.loc[index,"linea_nombre"])[0],"").split("–",2)[1].strip()
        else:
            lines.loc[index,"OrBus"] = lines.loc[index,"linea_nombre"].replace(re.findall("\d+KV",lines.loc[index,"linea_nombre"])[0],"").split("-",2)[0].strip()
            lines.loc[index,"DesBus"] = lines.loc[index,"linea_nombre"].replace(re.findall("\d+KV",lines.loc[index,"linea_nombre"])[0],"").split("-",2)[1].strip()
    else:
        if "–" in lines.loc[index,"linea_nombre"]:
            lines.loc[index,"OrBus"] = lines.loc[index,"linea_nombre"].split("–",2)[0].strip()
            lines.loc[index,"DesBus"] = lines.loc[index,"linea_nombre"].split("–",2)[1].strip()
        else:
            lines.loc[index,"OrBus"] = lines.loc[index,"linea_nombre"].split("-",2)[0].strip()
            lines.loc[index,"DesBus"] = lines.loc[index,"linea_nombre"].split("-",2)[1].strip()

    for element in lines_data.index:
        if lines_data.loc[element,"valor_texto"] == True:
            lines_data.loc[element,"valor_texto"] = "1"
        elif lines_data.loc[element,"valor_texto"] == False:
            lines_data.loc[element,"valor_texto"] = "0"
        elif all(i.isalpha() for i in lines_data.loc[element,"valor_texto"].replace(" ","")):
            lines_data.loc[element,"valor_texto"] = "0"
        elif lines_data.loc[element,"valor_texto"] == "":
            lines_data.loc[element,"valor_texto"] = "0"
        else:
            lines_data.loc[element,"valor_texto"] = lines_data.loc[element,"valor_texto"].replace(",",".")
            if not lines_data.loc[element,"valor_texto"].isnumeric():
                lines_data.loc[element,"valor_texto"] = re.findall('\d+\.\d+',lines_data.loc[element,"valor_texto"])[0]
    lines_data.iloc[:,2] = lines_data.iloc[:,2].astype(float)
    lines.loc[index,"kV"] = lines_data.iloc[0,2]
    lines.loc[index,"km"] = lines_data.iloc[1,2]
    lines.loc[index,"R+"] = lines_data.iloc[2,2]
    lines.loc[index,"X+"] = lines_data.iloc[3,2]
    lines.loc[index,"R0"] =lines_data.iloc[4,2]
    lines.loc[index,"X0"] =lines_data.iloc[5,2]
    printProgressBar(index + 1, lines.index[-1] + 1, prefix = 'Progress:', suffix = 'Complete', length = 50)

t0 = exectime("Buses data",False,t0)

print("exporting file...")
with pd.ExcelWriter("transmition_CEN.xlsx", engine = 'xlsxwriter') as writer:
    lines.to_excel(writer, sheet_name = "Transmition")
    buses.to_excel(writer, sheet_name = "Buses")
    print("File CEN transmition lines exported to: main root.")