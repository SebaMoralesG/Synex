import pandas as pd

def csv_datamaker(file):
    newdata = pd.read_csv(file,skiprows=[0,1,2])
    print(newdata)

csv_datamaker("pnomnd.csv")         # C:\Users\Seba Morales\Documents\Synex\

# data = pd.read_excel("taller1.xlsx", skiprows=[0,1])
# print(data)

# iterating the columns

# data1 = data.filter(["FC","Cupon","Vencimiento"])
# print(data1)
# data2 = data1.to_numpy()

# for i in data2:
#     print(i)