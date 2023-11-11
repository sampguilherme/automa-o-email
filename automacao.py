import pandas as pd

#Lê a planilha do excel
df = pd.read_excel("teste.xlsx")



for i in range(df.__len__()):
    print(df.loc[i])
    df = df.drop(i) 
    df.to_excel("teste.xlsx", index=False)
    
    