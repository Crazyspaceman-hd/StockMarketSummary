# Stock Market Summary
The goal of this project was to create a VBA script that would run through a large spreadsheet of stocks date and output information.

The program I wrote checks the stock ticker in the first column, saves the opening cost of that ticker then proceeds down the column adding the daily volume. Each line it checks the ticker it has saved against the first column. If the ticker changes it will add the final volume, subtract the final closing price from the saved opening price and print out the data in columns to the right. It will also check to see if the ticker has a price of 0 to prevent errors but if the ticker has an opening stock price above 0 on a lower row the program will pick that up as the opening price for the year.

It also creates conditional formatting rules on the worksheets to indicate a positive or negative yearly change, as well as resizing the columns for ease of viewing.

The program only needs to be run once and will proceed through all worksheets in the workbook.


