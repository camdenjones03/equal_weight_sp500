import pandas as pd
import yfinance as yf
import time

# Adds some formatting to the market cap by displaying the full number followed by the abbreviated market cap that is easier to see at a glance
def format_market_cap(value):
    if value >= 1_000_000_000_000:
        return f"{value:,} (${value / 1000000000000:.2f}T)"
    elif value >= 1_000_000_000:
        return f"{value:,} (${value / 1000000000:.2f}B)"
    elif value >= 1_000_000:
        return f"{value:,} (${value / 1000000:.2f}M)"
    else:
        return value

print('This program uses live Yahoo Finance stock prices to calculate how many shares (fractional included)\nto purchase of each stock to create an equal weight index of the S&P500 stocks.')
print('\nIf you would like to use your own updated list of stocks, please use a CSV with just the stock tickers in one column under the header "Ticker".')

# Loop to handle file name input for user list of stocks
while True:
    user_data = input('If you have your own CSV, please enter the file name here (Press enter to use default stocks): ')
    if user_data.strip() == '':
        stocks = pd.read_csv('sp500_companies.csv')
        break
    else:
        try:
            stocks = pd.read_csv(user_data)
            break
        except:
            print('There was an error accessing your file. Please try again or use the default list of stocks.')

# Loop to handle portfolio size input
while True:
    try:
        portfolio_size = float(input('Please enter your portfolio size: $'))
        break
    except ValueError:
        print('Not a valid number! Please enter a number.')
print('This may take up to a couple of minutes.')
print('Calculating...')

# Initialize data frame with column titles and data types
df_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
df = pd.DataFrame(columns=df_columns)
df = df.astype({'Ticker': 'str', 'Stock Price': 'float', 'Market Capitalization': 'int', 'Number of Shares to Buy': 'str'})

# Loop through ticker list in csv file, try up to 3 times to get info (use get() to avoid key error), format and concatenate onto data frame, notify user of failed tickers
for symbol in stocks['Ticker']:
    ticker = yf.Ticker(symbol.strip())
    delay = 5
    for attempt in range(3):
        try:
            info = ticker.info
            val = info.get('currentPrice', None)
            
            val2 = info.get('marketCap', None)
            
        
            if val is not None and val2 is not None:
                market_cap = format_market_cap(val2)
                price = round(val, 2)

                new_row = {
                    'Ticker': symbol,
                    'Stock Price': price,
                    'Market Capitalization': market_cap,
                    'Number of Shares to Buy': 'N/A'
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            break
        except Exception:
            print(f"Error retrieving info for {symbol}")
#            time.sleep(delay)
#            delay *= 2
#        time.sleep(1)
# The above exponential backoff delay is typically not necessary, but sometimes it helps if the Yahoo Finance API is giving 429 errors for too many requests.

# Determine position size, use position size to calculate shares to buy and insert in data frame using loc[]
position_size = portfolio_size / len(df.index)
for i in range(len(df)):
    df.loc[i, 'Number of Shares to Buy'] = position_size/df.loc[i, 'Stock Price']

# Display results in terminal
print(df.to_string(index=False))

# Create excel writer
writer = pd.ExcelWriter('Equal_Weight_Trades.xlsx', engine = 'xlsxwriter')
df.to_excel(writer, sheet_name= 'Equal Weight Trades', index = False)

background = '#000000'
font = '149414'

# Create formats for each relevant type, align right on string format to match number formats
string_format = writer.book.add_format(
    {
        'font_color': font,
        'bg_color': background,
        'border' : 1,
        'align' : 'right'
    }
)

dollar_format = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color': font,
        'bg_color': background,
        'border' : 1
    }
)

share_format = writer.book.add_format(
    {
        'num_format': '0.000',
        'font_color': font,
        'bg_color': background,
        'border' : 1
    }
)

# Create dictionary to be used for loop that applies formats, one key/value pair for each column of output
column_formats = {
    'A' : ['Ticker', string_format],
    'B' : ['Stock Price', dollar_format],
    'C' : ['Market Capitalization', string_format],
    'D' : ['Number of Shares to Buy', share_format]
}

# Loop through columns and apply formats, both to the data and the header
for column in column_formats.keys():
    writer.sheets['Equal Weight Trades'].set_column(f'{column}:{column}', 24, column_formats[column][1])
    writer.sheets['Equal Weight Trades'].write(f'{column}1', column_formats[column][0], string_format)

# Save and close excel file
writer.close()