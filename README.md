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

In this incarnation of the code, I used the ```tickers``` array within a nested ```For``` loop and forced the code to look at every line of the dataset each time a new ticker was referenced. This made the program unnecessarily cumbersome, hence the longer execution times pictured above. 

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

In the refactored code, I created an index of ```tickers``` which allowed each ticker to be used in the searching of the data but included code to advance the index every time a different ticker value was reached. Using this method, and by assigning values to each of the output arrays, the program looped over the dataset only once, saving time and making the program much cleaner. 

##### *code comments have been removed for length*
```
    For i = tickerIndex To 11

        tickerVolumes(i) = 0
    
    Next i

    For i = 2 To RowCount

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
       
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
         End If

        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            tickerIndex = tickerIndex + 1
             
        End If

    Next i
    
    For i = 0 To 11
    
     Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i
```

### Summary
The advantages of refactoring code are evidenced by both the final performance times and the cleanliness of the refactored product. There is a much lower chance that a different coder would misunderstand how I came to write the final product based on the information contained within the refactored code. 

Disadvantages include the possibility of breaking a working program, not understanding the original coder’s intentions (if there are missing comments on those lines that pull a tremendous amount of weight or the task’s instructions were unclear) and using the tools contained within VBA in unfamiliar ways. 

While the bones of this refactored product were contained within the original script, it held enough clutter to be difficult to sift through to find the important points. Once it became clear that this script would loop only once through all of the rows in the spreadsheet, the purpose of indexing the ```ticker``` value, and moving the output values outside of the loop made sense. 