# Stock-Analysis
## Project Overview

The client Steve has data for the stock performance of twelve different stocks in the years 2017 and 2018, with each year's data listed in a separate worksheet in an excel workbook.  This data includes the opening and closing prices and the daily volume for each stock for each day of the year.  Steve wants to compare the performance of each stock for each year, especially focusing on stock DQ as his parents are invested in it.  Steve wishes to determine if staying invested in DQ is wise, as well as any alternative stocks that could be good invenstments.

To analyze the performance of the stocks, I wrote a VBA script that will take the desired year for analysis as an input, then populate a table with each stock's total daily volume for the entire year as well as the percentage change in the Closing Price from the first day to the last day of the year, and finally format the color of the cell with that percentage green if the percentage increased and red if it decreased.  I then refactored the script to make it more efficient and reduce the run time; that refactoring and the advantages and disadvantages will be discussed later in this report.

## Results

### Comparing the original and refactored scripts

![2017original](/resources/Not_refactored_2017.png)
![2018original](/resources/Not_refactored_2018.png)
![2017refactor](/resources/VBA_Challenge_2017.png)
![2018refactor](/resources/VBA_Challenge_2018.png)

###

