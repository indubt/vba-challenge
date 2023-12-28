# vba-challenge
VBA Challenge for Stock Data Analysis


This repository contains the [VBA scipt](./Stock%20Data/Multi_year_Stock_Data_Report.vbs) for analyzing stock market data. 
The repository also contains the [screenshots](./Images/) of the results when executing the VBA script on the [Stock Data](./Stock%20Data/Multiple_year_stock_data_solution.xlsm) Provided.


Solution:

VBA Script loops through all the worksheets in the given stock data and performs the below:

    1. Inserts the headers and Labels for
        a. Ticker
        b. Yearly Change
        c. Percent Change
        d. Total Stock Volume
        e. Greatest % Increase
        f. Greatest % Decrease
        g. Greatest Total Volume
        h. Ticker and Value for Greatest Values

    2. For Each Ticker 
        a. find the open and close values
        b. calculate the yearly change
        c. calculate the percent change
        d. calculate the total stock volume
        e. populate the stock report summary table

    3. With Stock Report summary table created from step 2
        a. calculate Greatest % Increase
        b. calculate Greatest % Decrease
        c. calculate Greatest Total Volume
        d. find the corresponding ticker symbol and populate Greatest Values

    4. Conditional Formatting for Yearly Change and Percent Change


References:

1. Conditional formatting:
    a. https://learn.microsoft.com/en-us/office/vba/api/excel.xlformatconditionoperator
    b. https://www.statology.org/vba-conditional-formatting/

2. Google Search for formatting Total Stock Volume and percent change columns

