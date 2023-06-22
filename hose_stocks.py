from vnquant import DataLoader

dl = DataLoader()
hose_stocks = dl.get_data('HOSE')
print(hose_stocks)
