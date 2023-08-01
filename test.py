from vnstock import *
df = financial_ratio('VGC', 'yearly', False)
df
pe = df.loc[0, 'priceToEarning']