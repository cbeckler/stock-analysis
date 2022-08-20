# VBA of Wall Street

## Overview of Project

### Purpose

The purpose of this analysis was to provide the client, Steve, with an Excel workbook capable of automated stock market analysis using Excel macro capability. The analyst had previously provided Steve with a macro capable of analyzing stocks with small scale datasets. Steve requested a refactored version of the code that might be capable of more large scale analysis. The analyst utilized indexing to make the code more efficient.

## Analysis

Analysis of this data was done in MS Excel, primarliy using VBA capability. The Excel document used to process data may be found [here](https://github.com/cbeckler/stock-analysis/blob/main/VBA_Challenge.xlsm).

More comprehensive explanation of methods may be found by subcategory below, following presentation of results:

### Results

Most of the stocks performed well in 2017, as seen in the screenshot below:

![2017 Results](https://github.com/cbeckler/stock-analysis/blob/main/resources/2017%20results.png)

All stocks except TERP had a positive return, indicating growth. However, in 2018 the prospective portfolio took a big hit, with only two of the tweleve stocks (ENPH and RUN) remaining positive, with all others losing value, as seen below:

![2018 Results](https://github.com/cbeckler/stock-analysis/blob/main/resources/2018%20results.png)

Based on the YoY performance, this would not be an ideal stock portfolio to invest in. A larger scale analysis would be recommended to find a selection of stocks with less volatility.

Fortunately, the refactored code would be ideal for an increased sample size analysis. An explanation of the differences in the code follows in the next section.

### Refactored versus Original Code

Both analysis were based off the same 12 object tickers array. The primary difference between the original and refactored code was that in the refactor, an index was added for the iteration. Since the original array was based on string data, this saves considerable processing power when scaled up.

For an example of these changes, we can look at how the original code computes ticker starting price versus how the refactored, indexed code does. 

The original code:

```
For i = 2 To rowEnd
    
     If Cells(i, 1).Value = ticker And Cells(i - 1, 1).Value <> ticker Then
    
        startingPrice = Cells(j, 6).Value
        
    End If
```
This code compares if the string value of the specified cell in each row matches the string value of the ticker, iterating over all the rows. Because string data has one of the largest memory footprints of any text-based data, this can slow down processing considerably in scaled-up analysis.

The refactored code:

```
For i = 2 To RowCount

  If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
      tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    
  End If
 ```
 
The tickerIndex, meanwhile, has assigned a interger value to each string value present in tickers, so it is scanning for numeric matches based on indexing. This makes the code considerably more efficent and less resource intensive. The way the indexing is assigned can be seen when the final analysis is called:
 
 ```
 For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
 Next i
 ```
 Now, instead of running through values "AY", "CSIQ", etc, the code is just running integers 0 through 11. The results of this on processing time can be seen in the next section.

### Runtimes

Runtimes are compared between the analyst's original yearValueAnalysis macro and the refactored AllStocksAnalysisRefactored macro.

2017 original:

![2017 original runtime](https://github.com/cbeckler/stock-analysis/blob/main/resources/2017%20original%20runtime.png)

2017 refactored:

![2017 refactored runtime](https://github.com/cbeckler/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

The difference between six tenths of a second and seven hundreths of a second may seem insignificant since both happen in the blink of an eye. However, this is almost an entire order of magnitude difference! When scaled up across thousands or millions of iterations, this can save a lot of time and processing power. 

The results for 2018 are nearly identical:

2017 original:

![2018 original runtime](https://github.com/cbeckler/stock-analysis/blob/main/resources/2018%20original%20runtime.png)

2017 refactored:

![2018 refactored runtime](https://github.com/cbeckler/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)


## Summary

### Advantages and Disadvantages of Refactoring

The immediate, obvious advantage of refactoring is better performance, which can save both time and money (time, because the program runs faster. Money, because it takes up less computing resources). Refactoring may also improve readability, which when working in team environments is critical for having code that can be understood and run or modified by mulitiple devs. Refactoring may lessen the accumluation of technical debt.

The immediate, obvious disadvantage of refactoring is that it may not work. The solutions you have tried to improve performance may not do anything to increase efficiency, or possibly even make the code run worse. With version control you can at least revert your code to the earlier, better performing version (or in the worst case scenario, non-broken version), but you have still lost the time you used to try to develop the refactor.

The less obvious disadvantage of refactoring is when it works...but still was not worth the time it took to develop it. Often, this is an issue of scaling. 

### How This Applies to This Project

Regarding the disadvantage of if refactoring is worth it mentioned above, the analyst took about an hour to refactor the code, which saves about half a second each time it is run. If the code continues to be run on a dataset of this size, it would need to be run around 7200 times for the amount of time saved to even equal the amount of time spent refactoring it. Refactoring this code would probably only be advisable if it were to be scaled up to a MUCH larger data source.

That said, since the performance improvement was an entire order of magnitude, the refactor was successful, and if this script were to be applied to a large scale or recurring project, it would certainly be worth the hour it took to develop it.

