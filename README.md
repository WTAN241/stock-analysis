# VBA of Wall Street

## Overview of Project

### Purpose
The purpose of this analysis is to explain my findings of the All Stocks Analysis and rationalize whether refactoring the code to loop through all the data one time successfully made the VBA script run faster.
## Analysis and Challenges

### All Stocks Performance Analysis

The purpose of this code is to analyze an entire dataset of stocks prices and volumes at the click of a button. This code also provides the ability to specify the year for analyzing the set of stocks. Currently, Steve's parents are putting all their eggs into one basket by putting all their money in the DQ stock. With this code, Steve and parents would be able to compare DQ to other alternative energy stocks in 2017 and 2018 at the press of a button. 

![All_Stocks_Performance_2017](All_Stocks_Performance_2017.png)

As shown in the image below, DQ had the strongest stock return performance among the stocks set provided by Steve. However, the 2017 performance cannot be used as a proper reference especially when we look at the 2018 All Stocks performanace shown below.

![All_Stocks_Performance_2018](All_Stocks_Performance_2018.png)

According to the graph, the same set of stocks experienced a decline of approximately 76% of the average retuns of the set of stocks. Moreover, DQ experienced the greater decline in return in 2018, thus, showing the significant volatility of this stock. On the other hand, there were two stocks that stuck out to me because of their consistent positive returns across 2017 and 2018: ENPH and RUN. Additionally, when comparing ENPH and RUN to the average return of each stock in 2017-2018, ENPH and RUN have a higher average return than DQ as shown in the graph below. Therefore, my suggestion to Steve's parents would be to purchase at least ENPH and SEDG to include in their alternative energy stocks portfolio because of their average higher return and because it will enable them to diversify their portfolio with stocks from companies that provide different solar energy products.

![Average_Return_2017_And_2018](Average_Return_2017_And_2018.png)

### Refactoring of the Code Analysis
My friend Steve appreciates my help with creating this code that analyzes many stocks at the same time, but Steve would also want my help with making the code more flexible to analyze a wider set of stocks. And since a bigger data set would cause the code to take longer to execute, I need to refactor or change the structure of the code without changing the outputs so to make the code run more efficiently. 

For this analysis, I implemented a refactoring of the code to test whether the refactoring could make the VBA script run faster. For the refactoring, I created an index and created three arrays, tickerVolumes, tickerStartingPrices, and tickerEndingPrices, in addition to the tickers array that was also included in the code of the module. The index was used as a reference across the different arrays used in the loop so that when the loop was running though the data, it would locate the specific ticker to use when collecting data. As a result, the loop would go through the data one time and collect all of the information without having to go through the entire set of rows multiple times. In the code from the module, there was no index used as a reference so the For Loop would go over every sincle row trying to determine whether each row belonged to the ticker in use, therefore, the loop would be constantly reinitiating at the start of each new row. When running the original code for the first time without refactoring, the All Stocks Analysis takes about 0.98 and 0.95 seconds to run for the years 2017 and 2018 respectively.

![Original_2017_Times](Original_2017_Times.png)
![Original_2018_Times](Original_2018_Times.png)

Comparatively, after implementing refactoring to the original code, the All Stocks Analysis takes

![VBA_Challenge_2017](VBA_Challenge_2017.png)
![VBA_Challenge_2018](VBA_Challenge_2018.png)

By using the tickerIndex to specify which ticker to look for, it helped reduced the runtime of the code by approximately an average of 0.79 secs or 82% (First run for each case) for a database that contains 3013 rows of entries and which would make an even greater difference for a bigger database.  
### Challenges and Difficulties Encountered

One of the challenges encountered was making sure that all the new created arrays were referencing to the TickeIndex. ......An easy solution to this is that VBA will indicate where the code stopped running which makes it easier to identify the bug. 

Another challenge is to make the code running times comparable since the first time you run a macro, the elapsed time may be longer than subsequent runs because computer resources need to be allocated to run a macro. Therefore, it is important to ensure that the computer resources have been allocated by running the macros several times and compare the elapsed times to look for consistencies. Once the difference between the elapsed times for the same macro is minimal, then it is safe to compare it to the elapsed times of other macro runs. For instance, when conducting 10 consequent runs of the refactored code for 2017, the line graph below shows the elapsed time results for these runs with one repeated output highlighted in light blue. The biggest difference between the elapsed times lies between the first and second run. Therefore, it is important to not use the first run of the macro for elapsed times comparison purposes since the computer resources have not been allocated yet. 

![2017_Runs_Elapsed_Times](2017_Runs_Elapsed_Times.png)



## Results

