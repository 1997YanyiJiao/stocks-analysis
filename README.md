# stocks-analysis
## Overview of Project: 
The purpose of this project is to help Steve and his parents (his first clients), they need to find out whether it is a good investment to invest in a stock called “DQ” or not, with the help of VBA analysis, we wrote excel macros to analysis not only the DQ tickers but other tickers's return rates too from the initial spreadsheet raw date. and output the results into a new worksheet. 
## Results: 
[Screenshot (5)](https://user-images.githubusercontent.com/100504550/156959674-13a50f3e-de33-4d53-95ab-1d42e1a545b6.png)
![Screenshot (6)](https://user-images.githubusercontent.com/100504550/156959677-dbd1c739-57d3-4147-b749-48e81dc87e4b.png)
(1)First we Create Variable and set a loop to loop over all the rows, with the totalVolume initially set to 0, we add up the totalVolume by the value in the current row by using if..then, also we find the startingPrice and Endingprice for our ticker by using "if..then" too. And then we output the totall Valume into new worksheet and use formula (Endingprice/Startingprice)-1 to find the returns of current Ticker.  
(2)Now for All stock analysis, we use the smiliar technic to get our results. 
## Summary: 
Code with multiple loops works the same as Refactoring code, but Refactoring code runs more fast and it takes less time
