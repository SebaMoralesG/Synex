import pandas as pd
import os
from zipfile import ZipFile

#####################################################################################################################################
################################################### UPDATE THE FOLLOWING VALUES #####################################################
#####################################################################################################################################
print("MgC Reshape: Please type the file root.")
file_root = input()

while True:
    zip_file = input("Add zipfile name: ",)
    if zip_file not in os.listdir(file_root):
        print("Zip file not found, please try again.")
    else:
        break

export_filename = "MgC_matrix.xlsx"

#####################################################################################################################################
#####################################################################################################################################

# Skip the following code 

SheetName = "CMG Barras"

print("Loading Buses...")

bus_list = ["PARINACOTA____220","P.ALMONTE_____220","CONDORES______220","TARAPACA______220","COLLAHUASI____220","LAGUNAS_______220",
            "CRUCERO_______220","CHUQUICAMATA__220","CHACAYA_______220","ENCUENTRO_____220","ELCOBRE_______220","LABERINTO_____220",
            "N.ZALDIVAR____220","ESCONDIDA_____220","MEJILLONES____220","ATACAMA_______220","CUMBRES_______500","D.ALMAGRO_____220",
            "C.PINTO_______220","CARDONES______220","MAITENCILLO___220","P.COLORADA____220","P.AZUCAR______220","L.PALMAS______220",
            "L.VILOS_______220","NOGALES_______220","QUILLOTA______220","POLPAICO______220","FLORIDA_______110","A.JAHUEL______220",
            "ITAHUE________154","ANCOA_________220","HUALPEN_______220","CHARRUA_______220","TEMUCO________220","CIRUELOS______220",
            "RAHUE_________220", "P.MONTT_______220","ANCUD_________110","CHILOE________110"]

print("Reading zip file...")
with ZipFile(os.path.join(file_root, zip_file)) as z:
    with z.open(zip_file[:-4] + ".xlsm") as file:

        df = pd.DataFrame(pd.read_excel(file, header = 0, sheet_name = SheetName,usecols = "A:G"))
        df = df.drop_duplicates()
        df = df.pivot(index = ["Mes", "Día","Hora"], columns = "Barra", values = "CMg [mills/kWh]")
        df = df.loc[:,(bus for bus in bus_list if bus in df.columns)]
        df = df.reset_index()

        for bus in bus_list:
            if bus not in df.columns:
                df[bus] = 0
        df = df.loc[:,bus_list]

        posible_cols = ["J:P", "S:Y", "AB:AH"]

        for option in posible_cols:
            df1 = pd.DataFrame(pd.read_excel(file, header = 0, sheet_name = SheetName,usecols = option))
            if not df1.empty:
                df1 = df1.drop_duplicates()
                df1 = df1.rename(columns = {i:i.replace(".1","") for i in df1.columns})
                df1 = df1.pivot(index = ["Mes", "Día","Hora"], columns = "Barra", values = "CMg [mills/kWh]")
                df1 = df1.loc[:,(bus for bus in bus_list if bus in df1.columns)]
                df1 = df1.reset_index()
                for bus in bus_list:
                    if bus not in df1.columns:
                        df1[bus] = 0
                df1 = df1.loc[:,bus_list]

                df = pd.concat([df,df1],ignore_index = True)

        print("exporting file...")
        with pd.ExcelWriter(os.path.join(file_root, export_filename), engine = 'xlsxwriter') as writer:
            df.to_excel(writer, sheet_name = SheetName)
            print(f"File {zip_file} exported to: {export_filename}.")