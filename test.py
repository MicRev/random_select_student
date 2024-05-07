import pandas as pd
import random

df = pd.read_excel("PyProgs\RandomSelect\练习名单.xlsx")
df_female = df[df["性别"]=="女"]
print(df_female.shape[0])