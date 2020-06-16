# VBA-challenge
We use VBA scripting to analyze metadata

Stock data : We find the yearly price change and total stock volume for each ticker and then find the ticker who has the greatest increase, decrease or total volume.

### The instructions: (from the README.md for the vba-challenge homework of BootCamp)

First, we create a script that will loop through all the stocks for each year and output the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.
  
  * The yearly change cells are colored green for 'increase' or red for 'decrease'.
  
  (VBAStocks/yearlychange_all.vbs)
  
For the challenge, for each year (worksheet) we also return the stock with 

  * the "Greatest % increase", 
  
  * "Greatest % decrease" and 
  
  * "Greatest total volume".
  
  (VBAStocks/yearlychange_all_hard.vbs)
  
### Extra : monthly summary

I wonder how the stock volumes change in a year. For each year (worksheet), we creat a monthly summary table with

  * tickers count for each month

  * total volume for each month
  
  * average volume for each month
  
  in a separate sheet 'Monthly Summary'.
  
  A line chart for this table can be created to see the timeline of the volume changes.
  
  (VBAStocks/extra_monthly_summary.vbs)
