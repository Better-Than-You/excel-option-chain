import requests
import json
from api_credential import credential


# Fetch credentials
credentials = credential()
API_KEY = credentials["API_KEY"]
AUTH_TOKEN = credentials["AUTH_TOKEN"]


# Request data payload
data = {
    "mode": "FULL",
    "exchangeTokens": {
        "NSE": ["99926000"]
    }
}

# Request headers
headers = {
    'X-PrivateKey': API_KEY,
    'Accept': 'application/json, application/json',
    'X-SourceID': 'WEB',
    'X-ClientLocalIP': '192.168.1.1',
    'X-ClientPublicIP': '203.0.113.1',
    'X-MACAddress': '00:1A:2B:3C:4D:5E',
    'X-UserType': 'USER',
    'Authorization': AUTH_TOKEN,
    'Content-Type': 'application/json'
}


# Function to fetch LTP
def fetch_ltp():
    url = 'https://apiconnect.angelone.in/rest/secure/angelbroking/market/v1/quote/'
    try:
        response = requests.post(url, headers=headers, data=json.dumps(data))
        if response.status_code == 200:
            response_data = response.json()
            ltp = response_data['data']['fetched'][0]['ltp']
            return ltp
        else:
            print(f"Error: {response.status_code} - {response.text}")
            return None
    except Exception as e:
        print(f"Request failed: {e}")
        return None

# Call the function
# send_request()

LTP =fetch_ltp()
print(LTP)



























# main.py
# from fetch_ltp import fetch_ltp

# def perform_analysis(ltp):
#     """Perform some analysis with the LTP."""
#     if ltp > 1000:
#         print(f"LTP of {ltp} is high, consider selling.")
#     else:
#         print(f"LTP of {ltp} is low, consider buying.")

# def main():
#     # Fetch the LTP
#     ltp = fetch_ltp()
#     if ltp is not None:
#         print(f"The Last Traded Price (LTP) is: {ltp}")
#         # Use LTP for analysis
#         perform_analysis(ltp)
#         # Use LTP for other tasks
#         print(f"LTP squared is: {ltp ** 2}")
#     else:
#         print("Failed to fetch LTP.")

# if __name__ == "__main__":
#     main()
