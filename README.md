# equal_weight_sp500
# This program uses pandas and the free Yahoo Finance API to retrieve price and market cap data for a list of S&P500 stocks in order to calculate the number of shares to purchase to create an equal weight index of those stocks given a portfolio size.
# You may use the provided CSV of S&P500 stocks which is now out of date. If you choose to use your own file, please follow the program prompts to enter the file name, and make sure all tickers are in one column under the header "Ticker" in the same format
# as the provided sp500_companies.csv.
# The program will take a couple of minutes to retrieve all of the data (since batch calls were not available), calculate and insert the relevant data in a pandas dataframe, and use xlsxwriter to generate a formatted Excel Spreadsheet with recommended trades to
# create the equal weight index.
