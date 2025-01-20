# Import credentials from config.py
from api_credential import credential
from SmartApi.smartWebSocketV2 import SmartWebSocketV2
from get_filtered_tokens import get_filtered_tokens
from fetch_ltp import fetch_ltp
import xlwings as xw
from datetime import datetime
from logzero import logger


# Fetch credentials
credentials = credential()
API_KEY = credentials["API_KEY"]
CLIENT_CODE = credentials["CLIENT_CODE"]
FEED_TOKEN = credentials["FEED_TOKEN"]
AUTH_TOKEN = credentials["AUTH_TOKEN"]
CORRELATION_ID = credentials["CORRELATION_ID"]

# Use the imported variables
# print(f"API Key: {API_KEY}")
# print(f"Client Code: {CLIENT_CODE}")
# print(f"Auth Token: {AUTH_TOKEN}")

# Fetch LTP
strike_price_roundup_ltp = fetch_ltp()

## Getting started with SmartAPI Websocket's
    ####### Websocket V2 sample code #######
    
# WebSocket setup
token_value = get_filtered_tokens()
correlation_id = CORRELATION_ID
action = 1
mode = 3

token_list = [
    {
        "exchangeType": 2,
        "tokens": token_value
    }
]


# Excel setup
# excel_file_path = r'C:\Users\neera\OneDrive\Desktop\start.xlsx'
excel_file_path ="option_chain.xlsm"
wb = xw.Book(excel_file_path)
sheet_name = "nifty"
if sheet_name not in [s.name for s in wb.sheets]:
    sheet = wb.sheets.add(sheet_name)
else:
    sheet = wb.sheets[sheet_name]

# Rearrange tokens and set up initial rows
base_row = 12
tokens = sorted(map(int, token_value))
for i in range(0, len(tokens), 2):
    call_token = tokens[i]
    put_token = tokens[i + 1] if i + 1 < len(tokens) else None
    sheet[f'F{base_row}'].value = call_token
    if put_token:
        sheet[f'AB{base_row}'].value = put_token
    base_row += 1

# Function to mimic Excel's MROUND
def mround(number, multiple):
    return round(number / multiple) * multiple

number = strike_price_roundup_ltp # Original number
multiple = 50 # Adjust the step to 50000 for desired output

# Apply mround function to get nearest multiple
nifty_strike_price = mround(number, multiple)
print(nifty_strike_price)

# Assign the same value to two cells
sheet["Q32"].value = nifty_strike_price
sheet["P8"].value = strike_price_roundup_ltp

# WebSocket callbacks
def on_data(wsapp, msg):
    try:
        token = int(msg.get("token")) # Ensure token is treated as an integer
        ltp = msg.get("last_traded_price")/100
        oi = msg.get("open_interest")/75
        volume = msg.get("volume_trade_for_the_day")/75
        open = msg.get("open_price_of_the_day")/75
        high = msg.get("high_price_of_the_day")/100
        low = msg.get("low_price_of_the_day")/100
        close = msg.get("closed_price")/100
        ltp_chg = ltp - close
        
        timestamp = datetime.fromtimestamp(msg["exchange_timestamp"] / 1000).strftime('%Y-%m-%d %H:%M:%S')

        # Logging the received data
        logger.info(f"Received token: {token}, LTP: {ltp}, OI: {oi}, Timestamp: {timestamp}, VOLUME: {volume}")
        
        
        # Iterate through the rows to find the matching token in F or R
        row = 12  # Start at the base row
        while sheet[f'F{row}'].value or sheet[f'AB{row}'].value:  # Stop when both are empty
            call_token = sheet[f'F{row}'].value
            put_token = sheet[f'AB{row}'].value


            if call_token == token:  # If the token matches the call token
                sheet[f'P11'].value = "LTP"  
                sheet[f'P{row}'].value = ltp  # Update LTP for the call
                sheet[f'M11'].value = "OI"   
                sheet[f'M{row}'].value = oi   # Update OI for the call
                sheet[f'N11'].value = "VOLUME"
                sheet[f'N{row}'].value = volume   # Update VOLUME for the 
                sheet[f'C11'].value = "OPEN"
                sheet[f'C{row}'].value = open  # Update OPEN for the call
                sheet[f'D11'].value = "HIGH"
                sheet[f'D{row}'].value = high   # Update HIGH for the 
                sheet[f'E11'].value = "LOW"
                sheet[f'E{row}'].value = low   # Update LOW for the call
                sheet[f'B11'].value = "CLOSE"
                sheet[f'B{row}'].value = close   # Update CLOSE for the 
                sheet[f'O11'].value = "LTP CHG"
                sheet[f'O{row}'].value = ltp_chg   # Update LTP CHANGE for the call
                sheet[f'A11'].value = "TIMESTAMP"
                sheet[f'A{row}'].value = timestamp  # Update timestamp for the call
                break

            elif put_token == token:  # If the token matches the put token
                sheet[f'R11'].value = "LTP"  
                sheet[f'R{row}'].value = ltp  # Update LTP for the put
                sheet[f'U11'].value = "OI"   
                sheet[f'U{row}'].value = oi   # Update OI for the put
                sheet[f'T11'].value = "VOLUME"
                sheet[f'T{row}'].value = volume   # Update VOLUME for the call
                sheet[f'AE11'].value = "OPEN"
                sheet[f'AE{row}'].value = open  # Update OPEN for the 
                sheet[f'AD11'].value = "HIGH"
                sheet[f'AD{row}'].value = high   # Update HIGH for the call
                sheet[f'AC11'].value = "LOW"
                sheet[f'AC{row}'].value = low   # Update LOW for the call
                sheet[f'AF11'].value = "CLOSE"
                sheet[f'AF{row}'].value = close   # Update CLOSE for the 
                sheet[f'S11'].value = "LTP CHG"
                sheet[f'S{row}'].value = ltp_chg   # Update LTP CHANGE for the 
                sheet[f'AG11'].value = "TIMESTAMP"
                sheet[f'AG{row}'].value = timestamp  # Update timestamp for the put
                break

            row += 1  # Move to the next row

    except Exception as e:
        logger.error(f"Error in on_data: {e}")


# def on_data(wsapp, message):
#     logger.info("Ticks: {}".format(message))
    # close_connection()

def on_open(wsapp):
    logger.info("on open")
    sws.subscribe(correlation_id, mode, token_list)
    # sws.unsubscribe(correlation_id, mode, token_list1)


def on_error(wsapp, error):
    logger.error(error)


def on_close(wsapp):
    logger.info("Close")



def close_connection():
    sws.close_connection()


# Assign the callbacks.
sws = SmartWebSocketV2(AUTH_TOKEN, API_KEY, CLIENT_CODE, FEED_TOKEN)
sws.on_open = on_open
sws.on_data = on_data
sws.on_error = on_error
sws.on_close = on_close

# Start WebSocket connection
try:
    logger.info("Starting WebSocket...")
    sws.connect()
except KeyboardInterrupt:
    logger.info("WebSocket connection terminated by user.")
finally:
    logger.info("Closing WebSocket connection...")
    sws.close_connection()
# sws.connect()
####### Websocket V2 sample code ENDS Here #######






#? Important Notes:- 
#* Safe Approach
# ltp = msg.get("last_traded_price", 0) / 100  # Default to 0 if the key doesn't exist #*but ye hamesa exit karega so, no tension
# OR

# ltp_raw = msg.get("last_traded_price")
# if ltp_raw is not None:
#     ltp = ltp_raw / 100
# else:
#     ltp = None  # Or any default value you want



#? RUNNING MARKET STRIKE PRICE(NIFTY=23200 EXPIRY 23 JAN 2024 ) DATA:-
#* {'subscription_mode': 3, 'exchange_type': 2, 'token': '54786', 'sequence_number': 32056818, 'exchange_timestamp': 1737356263000, 'last_traded_price': 25820, 'subscription_mode_val': 'SNAP_QUOTE', 'last_traded_quantity': 75, 'average_traded_price': 19815, 'volume_trade_for_the_day': 70813950, 'total_buy_quantity': 605175.0, 'total_sell_quantity': 281850.0, 'open_price_of_the_day': 20550, 'high_price_of_the_day': 26900, 'low_price_of_the_day': 16710, 'closed_price': 18380, 'last_traded_timestamp': 1737356261, 'open_interest': 3118800, 'open_interest_change_percentage': 4602115189163712890, 'upper_circuit_limit': 73165, 'lower_circuit_limit': 5, '52_week_high_price': 72990, '52_week_low_price': 0, 'best_5_buy_data': [{'flag': 1, 'quantity': 450, 'price': 25815, 'no of orders': 3}, {'flag': 1, 'quantity': 375, 'price': 25810, 'no of orders': 4}, {'flag': 1, 'quantity': 600, 'price': 25805, 'no of orders': 7}, {'flag': 1, 'quantity': 4200, 'price': 25800, 'no of orders': 19}, {'flag': 1, 'quantity': 150, 'price': 25795, 'no of orders': 2}], 'best_5_sell_data': [{'flag': 0, 'quantity': 150, 'price': 25870, 'no of orders': 2}, {'flag': 0, 'quantity': 1425, 'price': 25875, 'no of orders': 5}, {'flag': 0, 'quantity': 300, 'price': 25880, 'no of orders': 3}, {'flag': 0, 'quantity': 1725, 'price': 25890, 'no of orders': 10}, {'flag': 0, 'quantity': 825, 'price': 25895, 'no of orders': 8}]}

#? FORMULA OF DIFFERENT VARIABLES IN OPTION CHAIN
# ltp_change= current running ltp - previous day's closed price/ltp