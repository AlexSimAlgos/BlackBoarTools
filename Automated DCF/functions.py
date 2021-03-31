import urllib2
import pandas as pd
import math

#TODO: Fix revenue Financial Firms (GS, WFC, etc.)
# dates = ["2012", "2013","2014","2015","2016"]
dates = ["2016", "2015","2014","2013","2012"]
empty = [0,0,0,0,0]
def print_urls(ticker):
    print( "IS: ", 'http://financials.morningstar.com/ajax/ReportProcess4CSV.html?&t='+ticker+'&region=usa&culture=en-US&cur=USD&reportType=is&period=12&dataType=R&order=desc&columnYear=5&rounding=3&view=raw&r=640081&denominatorView=raw&number=3')
    print( "BS: ", 'http://financials.morningstar.com/ajax/ReportProcess4CSV.html?&t='+ticker+'&region=usa&culture=en-US&cur=USD&reportType=bs&period=12&dataType=R&order=desc&columnYear=5&rounding=3&view=raw&r=640081&denominatorView=raw&number=3')
    print( "CF: ", 'http://financials.morningstar.com/ajax/ReportProcess4CSV.html?&t='+ticker+'&region=usa&culture=en-US&cur=USD&reportType=cf&period=12&dataType=R&order=desc&columnYear=5&rounding=3&view=raw&r=640081&denominatorView=raw&number=3')
# URL = is just a way to download the csv files from morningstar. Url's utilize variables in their address's to download
# the corresponding files. IF MORNINGSTAR CHANGES THEIR URL ADDRESSES THIS ENTIRE PROGRAM IS BROKEN.

def financials_download(ticker,report,frequency):
    if frequency == "A" or frequency == "a":
        frequency = "12"
    elif frequency == "Q" or frequency == "q":
        frequency = "3"
    url = 'http://financials.morningstar.com/ajax/ReportProcess4CSV.html?&t='+ticker+'&region=usa&culture=en-US&cur=USD&reportType='+report+'&period='+frequency+'&dataType=R&order=desc&columnYear=5&rounding=3&view=raw&r=640081&denominatorView=raw&number=3'
    # print(report, "         ", url)
    df = pd.read_csv(url, skiprows=1, index_col=0)
    #df.to_csv('test.csv')
    return df

def ratios_download(ticker):
    url = 'http://financials.morningstar.com/ajax/exportKR2CSV.html?&callback=?&t='+ticker+'&region=usa&culture=en-US&cur=USD&order=desc'
    print(url)
    df = pd.read_csv(url, skiprows=2, index_col=0)
    #df.to_csv('test.csv')
    return df

def y_stats(symbol, stat):
    url = 'http://finance.yahoo.com/d/quotes.csv?s=%s&f=%s' % (symbol, stat)
    return urllib2.urlopen(url).read().strip().strip('"')

def cash_and_equivalents(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Cash and cash equivalents',time]

def short_term_investments(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    historical_short_term_investments = {}
    for date in range(0, time):
        try:
            historical_short_term_investments[date] = df.ix["Short-term investments", date]
        except:
            try:
                historical_short_term_investments[date] = df.ix["Equity and other investments", date]
            except Exception as e:
                print("Problem with:")
                print(e)
                print("In: ", date)
    return historical_short_term_investments

def total_cash(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Total cash',time]

def receivables(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Receivables',time]

def inventories(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    historical_inventories = {}
    for date in range(0, time):
        try:
            historical_inventories[date] = df.ix["Inventories", date]
        except Exception as e:
            print("Problem with:")
            print(e)
            print("In: ", date)
            print("Set inventories to 0")
            historical_inventories[date] = 0


    return historical_inventories

def prepaid_expenses(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Prepaid expenses',time]

def current_assets(ticker,frequency,time):
    df = financials_download(ticker, 'bs',frequency)
    historical_current_assets = {}
    for date in range(0, time):
        historical_current_assets[date] = df.ix["Total current assets", date]
    return historical_current_assets

def other_current_assets(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Other current assets',time]

def gross_ppe(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Gross property, plant and equipment',time]

def accumulated_depreciation(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Accumulated Depreciation',time]

def net_ppe(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Net property, plant and equipment',time]

def equity_and_other_investments(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Equity and other investments',time]

def goodwill(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Goodwill',time]

def intangible_assets(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Intangible assets',time]

def other_longterm_assets(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Other long-term assets',time]

def total_noncurrent_assets(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Total non-current assets',time]

def total_assets(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    historical_assests = {}
    for date in range(0,time):
        historical_assests[date] = df.ix["Total assets",date]
    return historical_assests

def capital_expenditures(ticker,frequency,time):
    df = financials_download(ticker, 'cf', frequency)
    historical_capital_expenditures = {}
    for date in range(0, time):
        historical_capital_expenditures[date] = df.ix["Capital expenditure", date]
    return historical_capital_expenditures


def accounts_payable(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Accounts payable',time]

def shortterm_debt(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    historical_short_term_debt = {}
    for date in range(0, time):
        try:
            historical_short_term_debt[date] = df.ix["Short-term debt", date]
        except Exception as e:
            historical_short_term_debt[date] = '0'
            print("Problem with:")
            print(e)
            print("Set short-term debt to 0")
    print(historical_short_term_debt)
    return historical_short_term_debt


def accrued_liabilities(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Accrued liabilities',time]

def other_current_liabilities(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Other current liabilities',time]

def total_current_liabilities(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    historical_current_long_term_debt = {}
    for date in range(0, time):
        historical_current_long_term_debt[date] = df.ix["Total current liabilities", date]
    return historical_current_long_term_debt

def longterm_debt(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    historical_Long_term_debt = {}
    for date in range(0, time):
        try:
            historical_Long_term_debt[date] = df.ix["Long-term debt", date]
        except Exception as e:
            historical_Long_term_debt[date] = 0
            print("Problem with:")
            print(e)
            print("In: ", date)
            print("Set long-term as 0")
    return historical_Long_term_debt


def deferred_taxes_liabilities(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Deferred taxes liabilities',time]

def pensions_postretirement_benefits(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Pensions and other postretirement benefits',time]

def minority_interest(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Minority interest',time]

def other_longterm_liabilities(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Minority interest',time]

def total_noncurrent_liabilities(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Total non-current liabilities',time]

def total_liabilities(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    historical_total_liabilities = {}
    for date in range(0, time):
        historical_total_liabilities[date] = df.ix["Total liabilities", date]
    return historical_total_liabilities

def common_stock(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Common stock',time]

def additional_paidin_capital(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Additional_paidin_capital',time]

def retained_earnings(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Retained earnings',time]

def treasury_stock(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Treasury stock',time]

def accumulated_other_comprehensive_income(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc['Accumulated other comprehensive income',time]

def total_stockholder_equity(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    historical_stockholder_equity = {}
    for date in range(0, time):
        try:
            historical_stockholder_equity[date] = df.ix["Total stockholders' equity", date]
        except Exception as e:
            print("Problem with:")
            print(e)
            print("In: ", date)
            try:
                historical_stockholder_equity[date] = df.ix["Total Stockholders' equity", date]
            except Exception as e:
                print("Problem with:")
                print(e)
                print("In: ", date)
    return historical_stockholder_equity

def total_liabilities_and_stockholders_equity(ticker,frequency,time):
    df = financials_download(ticker,'bs',frequency)
    return df.loc["Total liabilities and stockholders' equity",time]

def revenue_5yr(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    return df.loc["Revenue",time]

def cost_of_goods(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    historical_cost_of_goods = {}
    for date in range(0, time):
        try:
            historical_cost_of_goods[date] = df.ix["Cost of revenue", date]
        except Exception as e:
            print("Problem with:")
            print(e)
            print("In: ", date)
            try:
                if math.isnan(df.ix["Other expenses", date]):
                    historical_cost_of_goods[date] = df.ix["Other expenses", date]
                elif math.isnan(df.ix["Interest expense", date]):
                    historical_cost_of_goods[date] = df.ix["Interest expense", date]
                else:
                    historical_cost_of_goods[date] = df.ix["Other expenses", date] + df.ix["Interest expense", date]

            except Exception as e:
                print("Problem with:")
                print(e)
                print("In: ", date)
    return historical_cost_of_goods

def gross_income(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    historical_gross_income = {}
    for date in range(0,time):
        historical_gross_income[date] = df.ix["Gross profit",date]
    return historical_gross_income

def cash_and_short_term_investments(ticker, frequency, time):
    df = financials_download(ticker,'bs',frequency)
    historical_cash_and_investments = {}
    for date in range(0,time):
        try:
            if math.isnan(df.ix["Cash and cash equivalents", date]):
                historical_cash_and_investments[date] = df.ix["Short-term investments",date]
            elif math.isnan(df.ix["Short-term investments",date]):
                historical_cash_and_investments[date] = df.ix["Cash and cash equivalents", date]
            else:
                historical_cash_and_investments[date] = df.ix["Cash and cash equivalents", date] + \
                                                    df.ix["Short-term investments",date]

        except Exception as e:
            try:
                if math.isnan(df.ix["Cash and cash equivalents", date]):
                    historical_cash_and_investments[date] = df.ix["Equity and other investments", date]
                elif math.isnan(df.ix["Equity and other investments", date]):
                    historical_cash_and_investments[date] = df.ix["Cash and cash equivalents", date]
                else:
                    historical_cash_and_investments[date] = df.ix["Cash and cash equivalents", date] + \
                                                            df.ix["Equity and other investments", date]
                print("Problem with:")
                print(e)
                print("In: ", date)
            except Exception as e:
                try:
                    print("Problem with:")
                    print(e)
                    print("In: ", date)
                    historical_cash_and_investments[date] = df.ix["Cash and cash equivalents", date]
                except Exception as e:
                    print("Problem with:")
                    print(e)
                    print("In: ", date)
    print(historical_cash_and_investments)
    return historical_cash_and_investments


def sales_administrative_expense(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    historical_SG_A = {}
    for date in range(0, time):
        historical_SG_A[date] = df.ix["Sales, General and administrative", date]
    return historical_SG_A

def sga_and_other_expenses(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    historical_sga_and_other = {}
    for date in range(0, time):
        try:
            historical_sga_and_other[date] = df.ix["Total operating expenses", date]
        except Exception as e:
            print("Problem with:")
            print(e)
            print("In: ", date)
            try:
                historical_sga_and_other[date] = df.ix["Total costs and expenses", date]
            except Exception as e:
                print("Problem with:")
                print(e)
                print("In: ", date)
    return historical_sga_and_other

def depreciation_amort_expense(ticker,frequency,time):
    df = financials_download(ticker,'cf',frequency)
    historical_depreciation = {}
    for date in range(0, time):
        historical_depreciation[date] = df.ix["Depreciation & amortization", date]
    return historical_depreciation



def interest_expense(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    return df.loc["Interest expense",time]

def other_operating_expenses(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    return df.loc["Other operating expenses",time]

def total_costs_and_expenses(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    return df.loc["Total costs and expenses",time]

def provision_for_income_expense(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    historical_provisions_for_income_tax = {}
    for date in range(0, time):
        historical_provisions_for_income_tax[date] = df.ix["Provision for income taxes", date]
    return historical_provisions_for_income_tax

def other_income_expense(ticker, frequency, time):
    df = financials_download(ticker, 'is', frequency)
    historical_other_income_expense = {}
    for date in range(0, time):
        try:
            historical_other_income_expense[date] = df.ix["Other income (expense)", date]
        except:
            historical_other_income_expense[date] = df.ix["Other", date]
    return historical_other_income_expense

def income_before_taxed(ticker, frequency, time):
    df = financials_download(ticker, 'is', frequency)
    historical_income_before_taxes = {}
    for date in range(0, time):
        try:
            historical_income_before_taxes[date] = df.ix["Income before taxes", date]
        except:
            historical_income_before_taxes[date] = df.ix["Income before income taxes", date]
    return historical_income_before_taxes

def net_income_contin_ops(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    return df.loc["Net income from continuing operations",time]

def net_income(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    historical_net_income = {}
    for date in range(0,time):
        historical_net_income[date] = df.ix["Net income",date]
    return historical_net_income

def preferred_dividend(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    return df.loc["Preferred dividend",time]

def net_income_to_common_shareholders(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    return df.loc["Net income available to common shareholders",time]

def EPS_basic(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    df = df[0:len(df)-3]
    return df.loc["Basic",time]

def EPS_diluted(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    df = df[0:len(df)-3]
    return df.loc["Diluted",time]

def EBITDA(ticker,frequency,time):
    df = financials_download(ticker,'is',frequency)
    return df.loc["EBITDA",time]

def net_cash_from_ops(ticker,frequency,time):
    df = financials_download(ticker,'cf',frequency)
    return df.loc["Net cash provided by operating activities",time]

def investments_in_ppe(ticker,frequency,time):
    df = financials_download(ticker,'cf',frequency)
    return df.loc["Investments in property, plant, and equipment",time]

def net_cash_used_for_investing_activities(ticker,frequency,time):
    df = financials_download(ticker,'cf',frequency)
    return df.loc["Net cash used for investing activities",time]

def debt_issued(ticker,frequency,time):
    df = financials_download(ticker,'cf',frequency)
    return df.loc["Debt issued",time]

def debt_repayment(ticker,frequency,time):
    df = financials_download(ticker,'cf',frequency)
    return df.loc["Debt repayment",time]

def net_cash_from_financing_activities(ticker,frequency,time):
    df = financials_download(ticker,'cf',frequency)
    return df.loc["Net cash provided by (used for) financing activities",time]

def net_change_in_cash(ticker,frequency,time):
    df = financials_download(ticker,'cf',frequency)
    return df.loc["Net change in cash",time]

def cash_at_beginning_period(ticker,frequency,time):
    df = financials_download(ticker,'cf',frequency)
    return df.loc["Cash at beginning of period",time]

def cash_at_end_period(ticker,frequency,time):
    df = financials_download(ticker,'cf',frequency)
    return df.loc["Cash at end of period",time]

def operating_cash_flow_5yr(ticker,frequency,time):
    df = financials_download(ticker,'cf',frequency)
    return df.loc["Operating cash flow",time]

def capital_expenditure_5yr(ticker,frequency,time):
    df = financials_download(ticker,'cf',frequency)
    return df.loc["Capital expenditure",time]

def free_cash_flow_5yr(ticker,frequency,time):
    df = financials_download(ticker,'cf',frequency)
    return df.loc["Free cash flow",time]

def gross_margin(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Gross Margin %",time]

def operating_margin(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Operating Margin %",time]

def EPS(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Earnings Per Share USD",time]

def DPS(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Dividends USD",time]

def payout_ratio(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Payout Ratio %",time]

def BVPS(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Book Value Per Share USD",time]

def operating_cash_flow(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Operating Cash Flow USD Mil",time]

def revenue(ticker,time):
    df = financials_download(ticker,'is','A')
    historical_revenue = {}
    for date in range(0, time):
        try:
            historical_revenue[date] = df.ix["Revenue", date]
        except Exception as e:
            print("Problem with:")
            print(e)
            print("In: ", date)
            try:
                historical_revenue[date] = df.ix["Total revenues", date]
            except Exception as e:
                print("Problem with:")
                print(e)
                print("In: ", date)

    return historical_revenue

def operating_income(ticker,time):
    df = financials_download(ticker,'is','A')
    historical_operating_income = {}
    for date in range(0,time):
        historical_operating_income[date] = df.ix["Operating income",date]
    return historical_operating_income

def capital_expenditure(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Cap Spending USD Mil",time]

def free_cash_flow(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Free Cash Flow USD Mil",time]

def FCFPS(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Free Cash Flow Per Share USD",time]

def working_capital(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Working Capital USD Mil",time]

def tax_rate(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Tax Rate %",time]

def net_margin(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Net Margin %",time]

def asset_turnover(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Asset Turnover (Average)",time]

def ROA(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Return on Assets %",time]

def financial_leverage(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Financial Leverage (Average)",time]

def ROE(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Return on Equity %",time]

def ROIC(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Return on Invested Capital %",time]

def interest_coverage(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Interest Coverage",time]

def revenue_yoy_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[41:45]
    return df.loc["Year over Year",time+1]

def revenue_3yr_avg_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[41:45]
    return df.loc["3-Year Average",time+1]

def revenue_5yr_avg_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[41:45]
    return df.loc["5-Year Average",time+1]

def revenue_10yr_avg_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[41:45]
    return df.loc["10-Year Average",time+1]

def operating_income_yoy_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[46:50]
    return df.loc["Year over Year",time+1]

def operating_income_3yr_avg_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[46:50]
    return df.loc["3-Year Average",time+1]

def operating_income_5yr_avg_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[46:50]
    return df.loc["5-Year Average",time+1]

def operating_income_10yr_avg_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[46:50]
    return df.loc["10-Year Average",time+1]

def net_income_yoy_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[51:55]
    return df.loc["Year over Year",time+1]

def net_income_3yr_avg_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[51:55]
    return df.loc["3-Year Average",time+1]

def net_income_5yr_avg_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[51:55]
    return df.loc["5-Year Average",time+1]

def net_income_10yr_avg_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[51:55]
    return df.loc["10-Year Average",time+1]

def EPS_yoy_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[56:60]
    return df.loc["Year over Year",time+1]

def EPS_3yr_avg_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[56:60]
    return df.loc["3-Year Average",time+1]

def EPS_5yr_avg_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[56:60]
    return df.loc["5-Year Average",time+1]

def EPS_10yr_avg_growth(ticker,time):
    df = ratios_download(ticker)
    df = df[56:60]
    return df.loc["10-Year Average",time+1]

def current_ratio(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Current Ratio",time]

def debt_to_equity(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Debt/Equity",time]

def quick_ratio(ticker,time):
    df = ratios_download(ticker)
    return df.loc["Quick Ratio",time]

def stock_exchange(symbol):
    return y_stats(symbol, 'x')

def price_current(symbol):
    return y_stats(symbol, 'l1')

def market_cap(symbol):
    return y_stats(symbol, 'j1')

def book_value(symbol):
    return y_stats(symbol, 'b4')

def dividend_yield(symbol):
    return y_stats(symbol, 'y')

def high_52_week(symbol):
    return y_stats(symbol, 'k')

def low_52_week(symbol):
    return y_stats(symbol, 'j')

def moving_average_50(symbol):
    return y_stats(symbol, 'm3')

def moving_average_200(symbol):
    return y_stats(symbol, 'm4')

def PE_ratio(symbol):
    return y_stats(symbol, 'r')

def forward_PE(symbol):
    return y_stats(symbol, 'r')

def PEG_ratio(symbol):
    return y_stats(symbol, 'r5')

def price_to_sales(symbol):
    return y_stats(symbol, 'p5')

def price_to_book(symbol):
    return y_stats(symbol, 'p6')

def short_ratio(symbol):
    return y_stats(symbol, 's7')

