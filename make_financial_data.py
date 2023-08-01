import pandas as pd
import numpy as np
from vnstock import *
import json
from openpyxl import load_workbook

# Read the CSV file
df = pd.read_csv("vn_stock_companies.csv")

# Filter the rows where group_code is 'VNINDEX' or 'HNX'
vnindex_df = df[df["group_code"] != "UpcomIndex"]

# Get the list of tickers
symbollist = vnindex_df["ticker"].tolist()

# Filter the symbol list
filtered_symbollist = [symbol for symbol in symbollist if len(symbol) <= 4]

financial_data = []  # List to store the financial data objects

for symbol in filtered_symbollist:
    net_income = {}
    compound_rate = {}
    dividend_yield = 0.0  # Default value for dividend_yield
    try:
        data = financial_report(
            symbol=symbol, report_type="IncomeStatement", frequency="Yearly"
        )

        # Remove the 'Q5 ' from years
        data.columns = [col.replace("Q5 ", "") for col in data.columns]

        # Find the row index where "CHỈ TIÊU" matches one of the three values
        row_index = data[
            data["CHỈ TIÊU"].str.contains(
                "Lợi nhuận của Cổ đông của Công ty mẹ|Lợi nhuận sau thuế của chủ sở hữu, tập đoàn|Lợi nhuận sau thuế phân bổ cho chủ sở hữu|Lợi nhuận sau thuế"
            )
        ].index[0]

        # Use the row index to get the net income data
        net_income = data.loc[row_index].to_dict()

        # Remove the first entry in the dictionary
        del net_income["CHỈ TIÊU"]

        years = list(net_income.keys())[5:]
        compound_rate = {}

        for year in years:
            year_int = int(year)
            sum_net_income_5_years = sum(
                net_income[str(y)] for y in range(year_int - 5, year_int)
            )
            compound_rate[year] = (
                net_income[year] - net_income[str(year_int - 5)]
            ) / sum_net_income_5_years
    except Exception as e:
        print(f"Error fetching data for symbol {symbol} in financial_report: {str(e)}")

    final_average_roe = 0.0
    pe = None
    try:
        # Calculate average 5 year ROE
        df = financial_ratio(symbol, "yearly", True)

        if "roe" in df.columns and len(df) >= 5:
            average_roe = df["roe"].head(5).mean()
            if not np.isnan(average_roe):
                final_average_roe = average_roe

        # Get latest Price to Earning ratio and dividend yield
        pe = df.loc[0, "priceToEarning"]

        dividend_yield = df.loc[0, "dividend"]
            
    except Exception as e:
        print(f"Error fetching data for symbol {symbol} in financial_ratio: {str(e)}")

    financial_data.append(
        {
            "ticker": symbol,
            "net_income": net_income,
            "compound_rate": compound_rate,
            "average_5y_roe": final_average_roe,
            "pe": pe,
            "dividend_yield": dividend_yield,
        }
    )

# Process financial_data to calculate average_5y_compound_rate
for entry in financial_data:
    compound_rates = entry["compound_rate"]
    latest_years = list(compound_rates.keys())[-5:]
    average_5y_compound_rate = sum(compound_rates[year] for year in latest_years) / 5
    entry["average_5y_compound_rate"] = average_5y_compound_rate

# sort financial_data based on the highest roe and average_5y_compound_rate
financial_data = sorted(
    financial_data,
    key=lambda x: (x["average_5y_roe"], x["average_5y_compound_rate"]),
    reverse=True,
)

# Define the file path to save the data
file_path = "financial_data.json"

# Write financial_data to the file in JSON format
with open(file_path, "w") as file:
    json.dump(financial_data, file)

# Assuming financial_data is a dictionary
df = pd.DataFrame(financial_data)

# Write the DataFrame to a CSV file
df.to_csv("financial_data.csv", index=False)

# Replacing the financial_data sheet in the excel. Removed for now.

# # Load the existing workbook
# book = load_workbook('stock-valuation-2023.xlsx')
# # Create an Excel writer object with the loaded workbook
# writer = pd.ExcelWriter('stock-valuation-2023.xlsx', engine='openpyxl') 
# writer.book = book

# # Delete the 'financial_data' sheet if it exists
# if 'financial_data' in book.sheetnames:
#     std = book['financial_data']
#     book.remove(std)

# # Write the DataFrame to the 'financial_data' sheet
# df.to_excel(writer, sheet_name='financial_data', index=False)

# # Save the workbook
# writer.save()
