# Stock-Analysis
## Project Overview

The client Steve has data for the stock performance of twelve different stocks in the years 2017 and 2018, with each year's data listed in a separate worksheet in an excel workbook.  This data includes the opening and closing prices and the daily volume for each stock for each day of the year.  Steve wants to compare the performance of each stock for each year, especially focusing on stock DQ as his parents are invested in it.  Steve wishes to determine if staying invested in DQ is wise, as well as any alternative stocks that could be good invenstments.

To analyze the performance of the stocks, I wrote a VBA script that will take the desired year for analysis as an input, then populate a table with each stock's total daily volume for the entire year as well as the percentage change in the Closing Price from the first day to the last day of the year, and finally format the color of the cell with that percentage green if the percentage increased and red if it decreased.  I then refactored the script to make it more efficient and reduce the run time; that refactoring and the advantages and disadvantages will be discussed later in this report.

## Results

### Conclusions

While stock DQ had the largest increase in stock price in 2017 of the twelve stocks analyzed with a growth of 199.4%, it had the largest price decrease in 2018 falling by 62.6%

![2017original](/Resources/Not_refactored_2017.png)

![2018original](/resources/Not_refactored_2018.png)

In 2017, all but one stock increased in price through the year-most very significantly, while in 2018 all but two dropped in price.  This general downward trend indicates overall volatility in the market, not a drop in quality of DQ specifically.  For some savvy traders who can monitor trends closely, DQ may still have potential.  However for the average person like Steve's parents, choosing the least volatile stocks with steady growth is a more secure option.  Therefore, Steve's parent should look to invest in ENPH and RUN, as they were the only stocks to grow in price both years.  They may also want to include a stock like AY in their portfolio.  It had was among the smallest growing stocks in 2017, and while it did decrease price in 2018, its drop was among the smallest.  So while the opportunity for large gains are not there, the steadiness of the stock performance makes it a safer option than others.

### Methods of Scripting and Inproving Script performance Through Refactoring

In order to analyze and summarize the large data tables for each year, the initial VBA script created an array conaining each of the twelve stock tickers, then looped over that array to analyze the full data set for each ticker one by one.

![OriginalOuter](/resources/Original_outer_loop.png)

I then used a nested loop with that first loop.

![OriginalNested](/resources/Original_nested_loop.png)

In this inner loop, each time the ticker in the current row matched the ticker being tested the ticker's value for that day was added to its total value.  I also used If Then statements to record the Closing price in the first row that matched the given ticker as the Beginning Price, and the Closing Price of the last row to match as the End Price.  These two prices could then be used to determine the percentage increase

By using this nested loop structure, the script examined each of the over 3,000 rows of the data set 12 times, once each for each of the 12 stocks.  The resulting script took almost 1.3 seconds to run, as seen here for each year:

![2017original](/resources/Not_refactored_2017.png)

![2018original](/resources/Not_refactored_2018.png)

After refactoring the code, the nested loop structure was eliminated.  Instead, before analyizing the rows of the data set, the script set a tickerIndex variable to 0, matching the index of the first stock in the tickers array.  Additionally, Output arrays were created to store the total volumes of each stock, initally all set to 0, as well as the Starting and Ending Prices

![OutputArrays](/resources/Output_arrays.png)

As the script analyzed each row, the daily volume was added to the total volume,  and If Then statemenst were again used to note when a row was first or last to display a given ticker and record the Starting and Ending prices to the arrays.  In this script, when a row was the last row for a given ticker, the ticker Index being tested was increased so that the next row would record the new stock's data appropriatley.

![RefactoredLoop](/resources/Refactored_loop.png)

This structure results in the script examining the entire data set only one time, instead of twelve.  As a result, the run times were reduced to under 0.18 seconds as seen here:

![2017refactor](/resources/VBA_Challenge_2017.png)

![2018refactor](/resources/VBA_Challenge_2018.png)
## Summary

### Advantages

Refactoring code after the initial scripting process can greatly improve the code's performance.  The refactoring process allows the opportunity to simplify or streamline the logic behind the code as well as remove any unecessary steps that may have been there while experimenting in the initial coding process; both of which can decrease the run time for a given script.  Additionally, there is an opportunity to add logic that will make the script more flexible.  For instance, variables can be assigned where values may have initially been hard coded.

In the refactored version of the script used in this project, the run time was reduced by about 85%.  


### Disadvantages

Refactoring code after a working version has been established does present a new opportunity for bugs and errors.  Changes to the overall logic used may require slight edits to the logic or structure of several lines of code.  Each edit required is an opportunity to make an error.  Additionally, new code blocks may be needed to offset lost functions from the original version being streamlined.  Having to write new code again gives the opportunity faulty logic or typing errors to result in bugs.

In the scripts for this project, the If Then statements used to determine the first and last rows of a given ticker had to be slightly adjusted as different variables were used to accomodate the different loop structures, as can be seen below first in the original script followed by the refactored:

![OriginalNested](/resources/Original_nested_loop.png)

![RefactoredLoop](/resources/Refactored_loop.png)

Each edit was a chance to make a mistake.

Additionally, the original version outer loop finished by recording the total value and Starting and Ending Prices to the output table before looping back over the next ticker.  

![OriginalOutputs](/resources/Original_outputs.png)

In the refactored version, after completing the loop over the entire data set, a new For loop was necessary to take the valuse stored in each of the out put arrays and use the to populate the output table:

![OutputArraysLoop](/resources/Output_arrays_loop.png)

Having to write all new code again gave opportunity for error.  
