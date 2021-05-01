# stock-analysis

## Project Overview

This project is an analysis of stock performance to inform Steve’s recommendations to his parents in their new forray into investment in the green industry. The dataset consists of financial data from 12 green energy companies from the years 2017 and 2018. 

## Analysis

The initial analysis was centered on the performance of **Daqo New Energy**, the company in which Steve’s parents invested all their money. However, after developing a VBA subscript to analyze the year-over-year performance of each of the stocks in Steve’s dataset, I adjusted the script to allow for a more efficient analysis, extracting the same information from a more robust dataset, at a faster speed. 

This means that Steve, the end user of the analysis, can easily run complex analyses of more extensive sets of stock market data with the click of a button.

## Results

#### 2017 vs. 2018 Trends
In 2017, eleven of the twelve green industry stocks in our analysis saw positive returns — And many, including Daquo New Energy (DQ), saw returns greater than 100%. 

However, analysis of the 2018 data for the same stocks shows a widespread decline in stock performance: The median 2018 return for our dataset was -12.0%. DQ in particular saw the greatest decline among those with a 2018 return of -62.6%.

![Annual_Return](https://user-images.githubusercontent.com/82285562/116794291-85ff1a00-aa91-11eb-9da4-5dddbdf46402.png)
---

#### Refactoring
In order to increase the program's effieiency and broaden Steve’s opportunity for analysis, refactored the initial code. This would allow Steve to apply the same subscript to analyze different sets of stock data.

1. **Measuring Efficiency**

To determing whether the refactoring was successful in creating a more efficient subroutine, I used VBA's `Timer` function to calculate the time elapsed between the end-user's `inputBox` submission and the analysis output.

2. **Decreasing Iterations**

The initial analysis script included a nested `for` loop, which returned stock KPIs by looping through the entire data sheet 12 times, iterating through the entire data sheet to output results for one ticker on each pass. While looping through 3012 rows of data 12 times is not a substantial computational lift, this method of extraction from our dataset is not ideal if we are to draw from a more extensive dataset.

3. **Using an Index and Conditionals**

In order to calculate the Total Daily Volume and Return for each stock in our set in just one pass through the data sheet, I then created a set of arrays, which would hold the output values for each variable:
```
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```
I also created a `tickerIndex` variable, which I initialized at 0 and referenced in a series of conditional statements:
```
    For j = 2 To RowCount
                        
        'increase volume for current ticker
        If Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
                        
        End If
        
        'check if the current row is the first row with the selected tickerIndex
        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                        
      End If
                        
        'check if the current row is the last of that ticker
        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            tickerIndex = tickerIndex + 1
                            
        End If
                
                    
    Next j
```
#### Conclusion
Looping through the 3012 rows of the data sheet just once, the refactored script ran in 0.210 seconds -- a **71.3% decrease** from the initial version (0.734 seconds).

![VBA_Challenge_2017](https://user-images.githubusercontent.com/82285562/116794299-a038f800-aa91-11eb-9983-2fedf86fbbc1.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/82285562/116794301-a333e880-aa91-11eb-9b3c-3c87d7f16d2e.png)

## Summary
There are a few advantages to refactoring code: It allows the programmer to re-imagine their initial approach to the problem (or re-imagine someone else's) and is an opportunity for creativity. Also, as showin in this report, refactoring can have a substantial impact on code performance when optimizing for efficiency, allowing programs to be more flexible in their potential application.
 
One of the main disadvantages to refactoring existing code, which I found *quickly* that it is easy to miss certain spots when updating variables or otherwise altering your script. Building off of my prior versions of this code taught me that keeping organized is especially important in refactoring. Otherwise it can be difficult to keep track of minor details that change as new edits are applied, which makes debugging tricky.



