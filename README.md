# Excel Option Chain - AngelOne WebSocket v2 (Python)

[![Python](https://img.shields.io/badge/Python-3.12-blue)](https://www.python.org/)
[![WebSocket](https://img.shields.io/badge/WebSocket-✔-green)](https://developer.mozilla.org/en-US/docs/Web/API/WebSocket)
[![HTTPS](https://img.shields.io/badge/HTTPS-Secure-orange)](https://en.wikipedia.org/wiki/HTTPS)

📈 This repository provides a **pure Python implementation** for fetching **real-time option chain data** from **AngelOne WebSocket v2** and saving it to **Excel**. 


---

💡 **This project also serves as one of the best live examples demonstrating the key differences between WebSocket and HTTPS.**  

---


## 📡 AngelOne API Endpoints Used  

### ✅ **Fetching LTP (Last Traded Price)**  
🔗 **API:** [AngelOne LTP API](https://apiconnect.angelone.in/rest/secure/angelbroking/market/v1/quote/)  

### ✅ **Fetching Instrument Tokens**  
🔗 **API:** [AngelOne OpenAPI ScripMaster](https://margincalculator.angelbroking.com/OpenAPI_File/files/OpenAPIScripMaster.json/)  

### ✅ **WebSocket v2 for Live Data**  
🔗 **GitHub:** [AngelOne Smart WebSocket V2](https://github.com/angel-one/smartapi-python/blob/main/SmartApi/smartWebSocketV2.py)  

---

## 🎥 YouTube Video  
📌 Watch the complete tutorial on YouTube:  
🔗 [Click Here to Watch](https://youtu.be/JbpYethcRF4)  

---

## 🚀 Features
✅ **Live Option Chain Data Fetching** from **AngelOne WebSocket v2**  
✅ **Secure HTTPS API Integration** for additional data retrieval  
✅ **Excel File Handling** - Save and analyze data in Excel format  
✅ **Asynchronous WebSocket Implementation** for efficient processing  
✅ **Easy to Modify & Extend** for trading strategies  

---

## 📌 Installation

Make sure you have **Python 3.12+** installed.  

```sh
git clone https://github.com/Neeraj-prajapat/excel-option-chain.git
cd excel-option-chain
pip install -r requirements.txt
