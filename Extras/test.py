import pandas as pd

df=pd.read_excel("../Input/output_data.xlsx")

df_input=pd.read_excel("../Input/input_data.xlsx")

b=df_input["Task Name"].to_list()
print(b)