# Stock Analysis with VBA

## Overview of Project

This project involves financial data and analysis of the stock market from 2017 and 2018 by using analysis tools such as Excel and VBA. The project analysis different stocks and displays the total daily volume and the yearly return of each stock. The code in VBA has two versions. The on in Module 1 is the initial code and the other one in Module 2 is the refactored code. The initial code completes the analysis, but the code is refactored to ensure better performance.

### Purpose

A client, Steve, wants stocks analyzed to find out the best stock options to invent in for his parents. By analyzing the stock market data from the years 2017 and 2018, this project provides the client with the analysis that would help him make the best decision. After writing the code to do so, this project seeks to refactor the code to ensure code efficiency and faster runtime.

## Analysis and Results

The original VBA code performs the analysis on stock data perfectly. However, the refactored code does it faster. One of the main difference the refactored code has is that it uses arrays to hold the total volumes of each stock, its starting prices and ending prices. This enabled me to use three separate for loops for the code instead of nested for loops. A sample of the original code is shown below:

![Main block of old code](https://github.com/Zarif601/stock-analysis/blob/main/Resources/MainBlockOld.png)

As we can see, this older version of the code uses nested for loops and single variables, whereas the new refactored code uses arrays and single for loops as shown below:

![Main block of refactored code](https://github.com/Zarif601/stock-analysis/blob/main/Resources/MainBlockNew.png)

The result was that the new code ran faster than the old code. The total elapsed time to run the old code is:

![Elapsd time for 2017 for old code](https://github.com/Zarif601/stock-analysis/blob/main/Resources/Old_stocks_2017.png)
![Elapsed time for 2018 for old code](https://github.com/Zarif601/stock-analysis/blob/main/Resources/Old_stocks_2018.png)

And here is the total elapsed time for the refactored code:

![Elapsed time for 2017 for new code](https://github.com/Zarif601/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![Elapsed time for 2018 for new code](https://github.com/Zarif601/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

As we can see, the refactored code runs much faster, and hence, is more efficient. The code was timed using the following blocks of code before and after the main block:

![Start timer](https://github.com/Zarif601/stock-analysis/blob/main/Resources/startTimer.png)
![End timer](https://github.com/Zarif601/stock-analysis/blob/main/Resources/endTimer.png)

## Summary

1. Advantages and disadvantages of refactoring code:
- The main advantage of refactoring code is that the code runs raster and it also becomes easier to read. After refactoring, the code usually has fewer steps and looks much cleaner as well. This makes it easier for someone else to follow the code and work with it.

- The advantges far outweight any possible disadvantage of refactoring code. However, it may be the case that someone is on a tight schedule and might not have the time to go back to a running code and tweak it to run raster or make it more readable, when there are other pending tasks. Having said that, it is always best to come back to the code when time permits and perform some refactoring.

2. How these pros and cons apply to refactoring the original VBA script:
- The original VBA script was performing the analysis corrently, however, the code had nested loops, which is always more complex than single loops. The refactored code, on the other hand, has no nested loops. It uses three separate loops to perform the same tasks and uses arrays to hold all the data. This makes the code look cleaner, more readable and easier to follow. It also made the code run faster.
