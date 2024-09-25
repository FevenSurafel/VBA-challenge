# VBA-Stock Market Analysis

This project contains a VBA script that automates the analysis of stock market data across multiple worksheets, each representing a different quarter. The script processes each worksheet, calculating the following key metrics for each stock:

-Ticker symbol: The stock’s unique identifier.
-Quarterly change: The price change from the opening price at the start of the quarter to the closing price at the end.
-Percentage change: The percentage difference between the opening and closing prices.
-Total stock volume: The total number of shares traded for each stock during the quarter.

Additionally, the script identifies the stocks with:

-The Greatest % Increase.
-The Greatest % Decrease.
-The Greatest Total Volume.

The script loops through all worksheets in the workbook and applies conditional formatting to highlight positive and negative quarterly changes automatically. It outputs the results in an organized format on each sheet.
