# The-VBA-of-Wall-Street
[StockAnalysis.bas](StockAnalysis.bas) is a VBA script that aggregates and analyzes stock market data.

The script uses an Excel file daily stock data in the following format:
>![stock data](readme-images/stock-data.jpg?raw=true "Stock Data")

The script goes through each stock by its and calculates the following:
* Change in price from opening on the first day to closing on the last day
* Percent change in price for the same time period (Stocks that opened at 0 on the first day will show a 100% change if any increase by closing on the last day)
* Total trading volume
>![stock calculations](readme-images/aggregated-stock-data-1.jpg?raw=true "Calculations")

The script will also output the stocks that had the greatest percent change increase, percent change decrease, and total volume.
>![stock calculations](readme-images/aggregated-stock-data-2.jpg?raw=true "Calculations")


 