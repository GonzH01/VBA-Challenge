# VBA-Challenge
The provided VBA script, Stockdata, processes each stock's data, calculates the required values, and populates them in the specified columns of the Excel sheet.
Purpose:
Calculates the yearly change in stock price.
Computes the percentage change in stock price over the year.
Sums up the total stock volume for each ticker over the year.
Identifies the ticker with the greatest percent decrease & increase over the year.
Applies conditional formatting to highlight positive and negative changes.

How it works:
    1) VBA Sheet Preparation: Ensure that within the VBA code ThisWorkbook.Sheets("X") is replaced with the title of your targeted sheet.
    2) VBA Date Preparation: Ensure the date within the VBA code matches with your targeted start & end date format from the Date data (ex: YYYYMMDD)
    3) Run each VBA code with respect to each sheet
    4) Apply Conditional Formatting: Manually apply conditional formatting to the 'Yearly Change' column to highlight increases (>0 green) and decreases (<0 red) in prices.
    5)Greatest Percent Decrease: Use Excel formulas (ex: MAX/MIN function) to find the greatest percent decrease and the corresponding ticker symbol.
