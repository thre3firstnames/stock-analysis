# Stock Analysis with VBA

## Overview of Project
Summarize the performance of 12 stocks over two different years, refactoring an existing VBA module to perform this task more quickly and efficiently than its predecessor. 

## Results
### Stock Performance 2017 & 2018
![All_Stocks_2017.png](/Resources/All_Stocks_2017.png)
![All_Stocks_2018.png](/Resources/All_Stocks_2018.png)

Based on performance, all stocks included in this dataset grew in 2017, with **TERP** as the exception. **TERP** underperformed during both years, and I cannot recommend investment based on the available data. Conversely, for most of these stocks, 2018 was a poor year except for **ENPH** and **RUN**. 

While **DQ** was the high-performing victor in 2017, its loss in 2018 also makes it an unwise investment. As discovered earlier, Steve should advise his parents to invest in a different stock. My recommendation based on this scope would be **ENPH** as it finished each year strong and outpaced its competitors. If they’re looking to diversify, **RUN** would be a great second choice, as its growth over 2018 suggests an upward trend. 

### Execution Times & Code Analysis
#### Original VBA Execution Times
![Pre_Refactoring_2017.png](/Resources/Pre_Refactoring_2017.png)
![Pre_Refactoring_2018.png](/Resources/Pre_Refactoring_2018.png)

In this incarnation of the code, we used the ```tickers``` array within a nested ```For``` loop and forced the code to look at every line of the dataset each time a new ticker was referenced. This made the program unnecessarily cumbersome, hence the longer execution times pictured above. 

```
    For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
        
        Worksheets(yearValue).Activate
        
        For j = 2 To RowCount
         
            If Cells(j, 1).Value = ticker Then
        
                totalVolume = totalVolume + Cells(j, 8).Value
        
            End If
            
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
                startingPrice = Cells(j, 6).Value
        
            End If
        
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
                endingPrice = Cells(j, 6).Value
        
            End If
        Next j
```

#### Refactored VBA Execution Times
![VBA_Challenge_2017.png](/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018.png](/Resources/VBA_Challenge_2018.png)


*(Using images and examples of your code:
1. compare the stock performance between 2017 and 2018
2. and the execution times of the original script and the refactored script.)*



###Summary
*( In a summary statement, address the following questions.
1. What are the advantages or disadvantages of refactoring code?
2. How do these pros and cons apply to refactoring the original VBA script?)*
