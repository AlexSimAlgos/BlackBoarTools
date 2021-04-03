from bs4 import BeautifulSoup
import requests
import re
import json
import pandas as pd
import tabulate
pd.set_option("max_rows", None)


# ----------------- Functions ----------------- #
# Fetches html file from morningstar and stores as a beautifulsoup
def get_financials_html(ticker):
    url1 = 'http://financials.morningstar.com/finan/financials/getFinancePart.html?&callback=xxx&t=' + ticker
    soup1 = BeautifulSoup(json.loads(re.findall(r'xxx\((.*)\)', requests.get(url1).text)[0])['componentData'], 'lxml')
    return soup1
# Fetches html file from morningstar and stores as a beautifulsoup
def get_keystats_html(ticker):
    url2 = 'http://financials.morningstar.com/finan/financials/getKeyStatPart.html?&callback=xxx&t=' + ticker
    soup2 = BeautifulSoup(json.loads(re.findall(r'xxx\((.*)\)', requests.get(url2).text)[0])['componentData'], 'lxml')
    return soup2
# Parses html file and stores values as dataframes
def make_dataframes(soup1, soup2):
    financial_data_df = pd.read_html(soup1.prettify())
    key_stats_df = pd.read_html(soup2.prettify())

    dict['financial_data'] = financial_data_df[0].dropna()
    dict['margin_percentage_revenue'] = key_stats_df[0].dropna()
    dict['profitability'] = key_stats_df[1].dropna()
    dict['yoy_percentages'] = key_stats_df[2].dropna()
    dict['cash_flow_ratios'] = key_stats_df[3].dropna()
    dict['balance_sheet_items'] = key_stats_df[4].dropna()
    dict['financial_health'] = key_stats_df[5].dropna()
    dict['efficiency'] = key_stats_df[6].dropna()

def print_to_console():
    for key, value in dict.items():
        print(key, '\n', value)

def to_excel():
    with pd.ExcelWriter(ticker +'_output.xlsx') as writer:  
        for key, value in dict.items():
            value.to_excel(writer, sheet_name=key)


# ----------------- Main ----------------- #
# Dataframe
dict = {}

# User input for stock ticker
ticker = input("Enter Stock Ticker: ").upper()

# Make dataframes
make_dataframes(get_financials_html(ticker), get_keystats_html(ticker))


# # Uncomment to write to excel
# to_excel()

# # Uncomment to print to console
# print_to_console()


