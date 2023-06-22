import vnquant.data as dt

loader = dt.DataLoader(symbols=["VND"], start="2022-01-01", end="2023-06-06", minimal=False, data_source="VND")
data = loader.download()
print(data.head())
