# The-VBA-of-Wall-Street
[StockAnalysis.bas](StockAnalysis.bas) is a VBA script that aggregates and analyzes stock market data. The original dataset can be found [here](https://drive.google.com/file/d/1zPX7jH2TjhoMoLyTYBhQA8YMLQ0V4ofF/view?usp=sharing).

The script uses an Excel spreadsheet of daily stock data in the following format:
>![stock data](readme-images/stock-data.jpg?raw=true "Stock Data")

The script calculates and outputs the following for each stock:
* Change in price from open on the first day to close on the last day
* Percent change in price for the same time period (Stocks that open at 0 on the first day will show change of 100% if any increase by close on the last day)
* Total trading volume
>![stock calculations](readme-images/stock-calculations-1.jpg?raw=true "Calculations")

The script also outputs the stocks that had the greatest percent change increase, percent change decrease, and total volume.
>![stock calculations](readme-images/stock-calculations-2.jpg?raw=true "Calculations")


 