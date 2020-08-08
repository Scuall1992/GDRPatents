import pandas as pd 
import os 

files = [i for i in os.listdir(".") if ".xlsx" in i]

df = pd.DataFrame()

for f in files:
    temp = pd.read_excel(f)
    df = pd.concat([df, temp])

df.to_excel("1.xlsx")
