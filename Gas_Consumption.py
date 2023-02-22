import pandas as pd
import os

def readfile(file,rows_skiped, colmns, sheet_name: str or int or list or None = 0):

    if file[-4:] == ".csv":
        if colmns == []:
            df = pd.DataFrame(pd.read_csv(file,skiprows= rows_skiped))
        else:
            df = pd.DataFrame(pd.read_csv(file,usecols= colmns,skiprows=rows_skiped))
    elif file[-5:] == ".xlsx":
        if colmns == []:
            df = pd.DataFrame(pd.read_excel(file,skiprows= rows_skiped, sheet_name = sheet_name))
        else:
            df = pd.DataFrame(pd.read_excel(file,usecols = colmns,skiprows=rows_skiped, sheet_name = sheet_name))
    df.head(10)
    return df

def month_year_index(data, date,start_date,last_date):
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
    if last_date == []:
        data = data.loc[start_date:,:]
    else:
        data = data.loc[start_date:last_date,:]
    return data

def init_date(file):
    df = pd.DataFrame(pd.read_csv(file,nrows=1))
    df = df.rename(columns = {i:i.replace(" ","") for i in df.columns})
    return [int(df.columns[-2]),int(df.columns[-1])]

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

def data_hydro_cond(data, hydro_condition, hydro_bound, df_index):
    wet_data = pd.DataFrame(index = df_index, columns = data.columns)
    med_data = pd.DataFrame(index = df_index, columns = data.columns)
    dry_data = pd.DataFrame(index = df_index, columns = data.columns)
    for index, row in data.iterrows():
        wet_index = hydro_condition.loc[(index[0],index[1]),:].loc[hydro_condition.loc[(index[0],index[1]),:] <= hydro_bound[0]].index
        med_index = hydro_condition.loc[(index[0],index[1]),:].loc[(hydro_condition.loc[(index[0],index[1]),:] < hydro_bound[1]) &
                    (hydro_condition.loc[(index[0],index[1]),:] > hydro_bound[0])].index
        dry_index = hydro_condition.loc[(index[0],index[1]),:].loc[hydro_condition.loc[(index[0],index[1]),:] >= hydro_bound[1]].index
        wet_data.loc[(index[0],index[1]),:] = data.loc[(index[0],index[1]),:].loc[data.loc[(index[0],index[1]),:].index.get_level_values(0).isin(wet_index)].mean()
        med_data.loc[(index[0],index[1]),:] = data.loc[(index[0],index[1]),:].loc[data.loc[(index[0],index[1]),:].index.get_level_values(0).isin(med_index)].mean()
        dry_data.loc[(index[0],index[1]),:] = data.loc[(index[0],index[1]),:].loc[data.loc[(index[0],index[1]),:].index.get_level_values(0).isin(dry_index)].mean()
    return wet_data, med_data, dry_data


# Main Code

file_root = input("Please add file's root: ",)

Hydro_limit_condition = [3,29]
init_hydro_date = 1988
last_hydro_date = 2018

date = init_date(os.path.join(file_root,"duraci.csv"))

print("Reading power plant Dictionary...")
dictionary = readfile("Power_Plant_Dictionary.xlsx",0,[0,1,2,3,4,5,6,7])
dictionary = dictionary.fillna("-")

gerter = readfile(os.path.join(file_root,"gerter.csv"),3,[])
gerter = gerter.rename(columns= {col: col.replace(" ","") for col in gerter.columns})
gerter = month_year_index(gerter,date,(date[1],date[0]),())

ctermise = readfile(os.path.join(file_root,"ctermise.csv"),0,[])
ctermise = ctermise.rename(columns = {"...Nombre...":"Name",".PotIns":"Pot", "Tipo":"Type"})
ctermise["Name"] = ctermise["Name"].str.replace(" ","")
ctermise = ctermise.iloc[:,[1,11,13]]
ctermise["Unit"] = None

ctermise.loc[ctermise["Name"] == "NRencaGN",:]

ccombuse = readfile(os.path.join(file_root,"ccombuse.csv"),0,[0,1,2])
ccombuse = ccombuse.rename(columns = {"Nome":"Name","UComb":"Unit","!Num":"Num"})
for index, row in ctermise.iterrows():
    ctermise.loc[index,"Unit"] = ccombuse.loc[ccombuse["Num"] == row["Comb"],"Unit"].values[0]

gas_list = dictionary.loc[dictionary["Tecnología Inglés"] == "Gas","SDDP"].unique()

gerter = gerter.loc[:,gas_list]
gerter = gerter.groupby(level = [0,1,2]).sum()

for col in gerter.columns:
    gerter.loc[:, col] = gerter.loc[:, col]*ctermise.loc[ctermise["Name"] == col,".CEsp.1"].values[0]

hidroseq = readfile("Orden_hidrologias.xlsx", 0, [])                                                        # read hydro seq file
hidroseq = hidroseq[(hidroseq["Year"] >= init_hydro_date) & (hidroseq["Year"] <= last_hydro_date)]          # take just hydro sequence used
hidroseq = hidroseq.sort_values(by = "Total", ascending = False)                                            # sort hydrologies frow wet to dry
hidroseq["index"] = [i for i in range(1,32)]                                                                # index values by preview order
hidroseq = hidroseq.set_index("index")                                                                      # Index it's setted

hydro_condition = case_hydro_condition(init_hydro_date, last_hydro_date, gerter.groupby(level = [0,1]).mean().index, hidroseq)

wet_gerter, med_gerter, dry_gerter = data_hydro_cond(gerter, hydro_condition, Hydro_limit_condition, gerter.groupby(level = [0,1]).mean().index)

wet_gerter["Total Wet"] = wet_gerter.sum(axis = 1)
med_gerter["Total Medium"] = med_gerter.sum(axis = 1)
dry_gerter["Total Dry"] = dry_gerter.sum(axis = 1)

hydro_gerter = pd.concat([wet_gerter["Total Wet"],med_gerter["Total Medium"],dry_gerter["Total Dry"]], axis = 1)

with pd.ExcelWriter(os.path.join(file_root,"Gas_consumption.xlsx"), engine = 'xlsxwriter') as writer:
    wet_gerter.to_excel(writer, sheet_name = "Wet", index = True)
    med_gerter.to_excel(writer, sheet_name = "medium", index = True)
    dry_gerter.to_excel(writer, sheet_name = "Dry", index = True)
    hydro_gerter.to_excel(writer, sheet_name = "Hydro", index = True)