from io import BytesIO
import requests
import pandas as pd


r = requests.get('https://docs.google.com/spreadsheet/ccc?key=1fIP8dwH5hwSDMKEmOUdbnvwbwMZyAOClAm_4HDVVe5k&output=csv')
data = r.content
df = pd.read_csv(BytesIO(data), index_col=0)
print(df)