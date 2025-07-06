# Stock Alert

This Google Apps Script monitors stock prices and sends email alerts when stocks reach specified buy or sell prices

## Sheet1

Here is the structure of the first sheet:
Header row (row 1)

| Ticker | Current Price              | Buy Price | Sell Price | Last Notified |
| ------ | -------------------------- | --------- | ---------- | ------------- |
| AAPL   | =GOOGLEFINANCE(A2,"price") | 150       | 170        |               |
| TSLA   | =GOOGLEFINANCE(A3,"price") | 600       | 750        |               |
| MSFT   | =GOOGLEFINANCE(A4,"price") | 250       | 300        |               |

## Sheet2

your email address in cell A1
