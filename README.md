# Analyzing Stock Data

## Overview of Project

An Excel workbook was previously encoded with VBA macros to analyze a subset of stock data for the years 2017 and 2018 for a client. The client wished to analyze volumes and returns in order to advise their parents investing portfolio.

The existing code was then refactored to allow the client to append the dataset if needed, and still have a runable macro with only the "tickers(x)" array needing to be updated. The refactoring also streamlined the code and allowed it to run faster.  

## Results

### Stock Performance

In the year 2017, all stocks except TERP were recorded to be profitable. DQ, ENPH, FSLR, and SEDG showed incredible profits of >100% returns. The only non-profitable stock, TERP, had a loss of only -7.2%, however.
(Ref: All_stocks_outcome_2017.png in Resources folder)

In the year 2018, most stocks were at a loss. Only ENPH and RUN were profitable (>80% return). Of the stocks that were the most profitable in 2017, stated above, they all showed a large decrease in percent return, indicating that these stocks have a higher risk to invest in.
![Alt text](/../<main>/Resources/All_stocks_outcome_2018.png?raw=true "All Stocks Outcome 2018")
(Ref: All_stocks_outcome_2018.png in Resources folder)

### Run-time

The original script ran the analysis for 2017 in 0.79 seconds (Ref: Module_runtime_2017.png in Resources folder), and the refactored script ran in 0.16 seconds (Ref: VBA_Challenge_2017.png in Resources folder). For 2018, the original script ran in 0.88 seconds (Ref: Module_runtime_2018.png in Resources folder), and the refactored code ran in  0.15 seconds (Ref: VBA_Challenge_2018.png in Resources folder).

By creating a tickerIndex variable and using it to input from and output to various arrays (tickerVolume(12), tickerStartingPrices(12), tickerEndingPrices(12)), the workflow of the code was streamlined and optimized to run more efficiently.

## Summary

### General Advantages and Disadvantages of Refactoring Code

Refactoring code has many advantages. As we learn newer techniques and methods we can continuously refer to and refactor old code to work better. Refactoring allows for a fresh perspective that can go back to see and fix flaws in logic, performance, and memory usage that can be missed while one is in the process of writing from start to finish. 

The disadvantages of refactoring code are few. Refactoring can potentially lead to bugs if one gets lost in someone elses, or even poorly commented code. This can prove to be challenging at first but once troubleshooted, it should still lead to a higher quality of code. The main disadvatage is that this whole process can take quite a bit of time.


### How these Pros and Cons Apply to Refactoring the Original VBA Script

Refactoring the original VBA script allowed it to run much more efficiently (see: run-times above). This may not be important with the current dataset but when larger inputs are being used cutting down time where ever possible is crucial. The reafactoring also allowed the code to be more easily updated if and when the data set is increased. This is a very important impact of the refactoring as it increases the functionality for the client. 

With regards to the disadvantages mentioned above, there were some difficulties initially in refactoring the original code. The difficulties came in trying to replace all the appropriate variables where required but, after taking the time to comb through the code, all bugs were easily fixed. 
