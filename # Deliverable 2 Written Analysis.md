# Deliverable 2 Written Analysis


## Overview of Project: Explain the purpose of this analysis.
The purpose was to make the code run time shorter by making the code more efficient.

The original AllStocksAnalysis code runs 12 times individually. It runs start to finish for each ticker including output values just once for each ticker. 

Whereas with the new code, it calculates values for all tickers and then stores those values in the arrays we created and then writes all values out just once. 

So, writing all rows just once helps to speed up the run time.

One strange glitch however, is that the first time I run the new re-factored code excel file after just opening Excel, it runs .6-.8 seconds, the initial run time speed non-factored. But, I press the Run Analsyis for All Stocks again, and whammo, 0.125 seconds and consistently so. I could not figure out why it does that only the first time once opening the file and running the code the first time. 

## Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

#### Comparison of Stock Performance between years 2017 & 2018:

Stock Performance 2017: Mostly positive return values
[2017 stock return](https://github.com/forrestcasey/stock-analysis/blob/main/2017%20stock%20return.png)

Stock Performance 2018: Mostly negative return values
[2018 stock return](https://github.com/forrestcasey/stock-analysis/blob/main/2018%20stock%20return.png)

2017 stock returns were more successful compared to the year 2018.

#### Run-Time Comparisons:




Original non-factored 2017 runtime:! [2017_Original Run Time AllStocksAnalysis.png](https://github.com/forrestcasey/stock-analysis/blob/main/Resources/2017_Original%20Run%20Time%20AllStocksAnalysis.png) 



Original non-factored 2018 runtime:! [2018_Original Run Time AllStocksAnalysis.png](https://github.com/forrestcasey/stock-analysis/blob/main/Resources/2018_Original%20Run%20Time%20AllStocksAnalysis.png)



Refactored code runtime for 2017: [VBA_Challenge_2017.png](https://github.com/forrestcasey/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)



Refactored code runtime for 2018: [VBA_Challenge_2018.png](https://github.com/forrestcasey/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)


### Summary: 

#### What are the advantages or disadvantages of refactoring code?

Advantage is that if you have thousands, maybe even 5 years worth of data with hundreds of thousands of lines of stock data, having a more efficient run time will almost be mandatory to effciiently run your analysis in an appropriate time frame. It is also easier to understand, read and maintain long term. 

Disadvantage: It takes longer to design code that works more efficiently, having to retest functionality several times. While it can be great to improve code, moving so many things around, you may lose the base structure necessary to run the subroutine. In business, with the many projects you might be working on, you may not have the time to get it perfect, so it will depend I think.


#### How do these pros and cons apply to refactoring the original VBA script?

The advantage to this project with the new refactored code is that it runs faster, and is easier to interpret when looking at the code. The variables are more clearly understood as members of an array.

In addition to running more slowly, the original code didn't refer to a tickerIndex which made understanding the flow of the program more difficult. 





