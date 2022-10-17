import pandas as pd
import requests
import re


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


# buses = API_request("https://api-infotecnica.coordinador.cl/v1/subestaciones")

lines = API_request("https://api-infotecnica.coordinador.cl/v1/secciones-tramos")
lines = lines.iloc[:,[0,1,2,5,6,7,10,11]]
lines["kV"] = 0
lines["km"] = 0
lines["R+"] = 0
lines["X+"] = 0
lines["R0"] = 0
lines["X0"] = 0

printProgressBar(0, lines.index[-1] + 1, prefix = 'Progress:', suffix = 'Complete', length = 50)
for index, row in lines.iterrows():
    line_id = lines.loc[0,"id"]
    line_id = lines.loc[index,"id"]
    lines_data = API_request(f"https://api-infotecnica.coordinador.cl/v1/secciones-tramos/{line_id}/fichas-tecnicas/general")
    lines_data = lines_data.iloc[[7,8,9,13],[1,2,3,4,6,7]].T
    lines_data = lines_data.fillna("0")
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
                print(re.findall('\d+.\d+',lines_data.loc[element,"valor_texto"]), lines_data.loc[element,"valor_texto"])
                lines_data.loc[element,"valor_texto"] = re.findall('\d+.\d+',lines_data.loc[element,"valor_texto"])[0]
    lines_data.iloc[:,2] = lines_data.iloc[:,2].astype(float)
    lines.loc[index,"kV"] = lines_data.iloc[0,2]
    lines.loc[index,"km"] = lines_data.iloc[1,2]
    lines.loc[index,"R+"] = lines_data.iloc[2,2]
    lines.loc[index,"X+"] = lines_data.iloc[3,2]
    lines.loc[index,"R0"] =lines_data.iloc[4,2]
    lines.loc[index,"X0"] =lines_data.iloc[5,2]
    printProgressBar(index + 1, lines.index[-1] + 1, prefix = 'Progress:', suffix = 'Complete', length = 50)

print("exporting file...")
with pd.ExcelWriter("transmition_CEN.xlsx", engine = 'xlsxwriter') as writer:
    lines.to_excel(writer, sheet_name = "Transmition")
    print("File CEN transmition lines exported to: main root.")

