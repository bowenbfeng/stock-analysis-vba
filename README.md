# Stock Analysis with VBA

## Overview of Project
VBA is a handy tool when analyzing financial data in Excel. This project analyzes the performance on 12 stocks in 2017 and 2018, with "Volume" and "Return" calculated by code instead of manually. Furthermore, the VBA code is then refactored, which drastically improves run time and efficiency. 

## Results

### Stock Performance
From the comparison, it can be seen that the stock market was generally better in 2017 than in 2018, with almost all the stocks increasing in price. The stock market got worse in 2018, with nearly all the stocks declining in price, except for only two. 

Below is the result from 2017. The highest performing stock was "DQ" with an 199.4% return. The lowest performing stock was "TERP" with a negative 7.2% in return. 

![Old 2017 Result](/Resources/Others/Old_2017_Result.png)

Below is the result from 2018. The highest performing stock was "RUN" with an 84.0% return. The lowest performing stock was "DQ" with a negative 62.6% in return. 

![Old 2018 Result](/Resources/Others/Old_2018_Result.png)

### Original Code Run Time
In this project, we are only dealing with limited rows of data for 12 stocks. It takes more than half a second to go through all the data, which is a lot. Therefore, in the next section, a refactoring process has been done. 

Below is the run time for getting the results for 2017. 

![Old 2017 Run Time](/Resources/Others/Old_VBA_Challenge_2017.png)

Below is the run time for getting the results for 2018. 

![Old 2018 Run Time](/Resources/Others/Old_VBA_Challenge_2018.png)

## Refactoring

### Refactoring Codes
Instead of looping through the data 12 times by ticker, this refactoring code would only go through the data only once. 

Some variables will be defined and arrays to be initialized to start with. 
```
'1a) Create a ticker Index
Dim tickerIndex As Single
tickerIndex = 0

'1b) Create three output arrays
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

''2a) Create a for loop to initialize the tickerVolumes to zero.
For i = 0 To 11

    tickerVolumes(i) = 0

Next i
```

The instructions provided are pretty clear. A tickerIndex has been used in combination with "tickerVolumes(tickerIndex)" to summarize the volumes of the stocks. The tickerIndex has also been in combination with "tickerStartingPrices(tickerIndex)" and "tickerEndingPrices(tickerIndex)" to record the starting price and the ending price of the stocks. 
```
''2b) Loop over all the rows in the spreadsheet.
For i = 2 To RowCount

    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
    If Cells(i - 1, 2).Value <> tickers(tickerIndex) Then

        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

    End If
    'End If

    '3c) check if the current row is the last row with the selected ticker
     'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
    'If  Then
    If Cells(i + 1, 2).Value <> tickers(tickerIndex) Then

        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

        '3d Increase the tickerIndex.
        tickerIndex = tickerIndex + 1

    End If
    'End If

Next i
```

In this way, we are only going through all the rows of data only once instead of 12 times from the previous code.

And then, a loop is used to go through the cells in Sheet "All Stocks Analysis" to paste the data retrieved from the arrays. 
```
'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11

    Worksheets("All Stocks Analysis").Activate
    tickerIndex = i
    Cells(i + 4, 1).Value = tickers(tickerIndex)
    Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
    Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1

Next i
```

### Refactoring Results
Before, it takes more then half a second to go through all the data, which is a lot. After refactoring, it only takes around 0.05 second to go through all the data, which is huge improvement. 

Below is the run time for getting the results for 2017 after refactoring. 

![2017 Run Time](/Resources/VBA_Challenge_2017.png)

Below is the run time for getting the results for 2018 after refactoring. 

![2018 Run Time](/Resources/VBA_Challenge_2018.png)

Below shows the same results from 2017. 

![2017 Result](/Resources/Others/2017_Result.png)

Below shows the same results from 2018. 

![2018 Result](/Resources/Others/2018_Result.png)

## Summary

> Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

### Advantages and Disadvantages of Refactoring
In an ideal case, a code should serve the functions as much generally as possible. If a code is written just to satisfy the needs in one specific case, it would be a bit wasteful. If the code is refactored to be not only good for one specific case but also good for other usage, it would be best practice. 

On the other hand, the disadvantages would be that it can be a bit time consuming to write a refactored code every time, especially when you only have limited time and energy and only want to solve this case in particular. It requires a lot of code to cope with different situations, especially on how to deal with outliers, to refactor a code that can be stably run in various situations. 

### Advantages and Disadvantages of Refactoring in this Case
In this case, the advantage of the refactoring is that it reduces the time to run the code drastically, around one tenth. If in the future, this code is used to analyze not just 12 but thousands of stocks, the run time would be acceptable. 

The disadvantage of this refactor done is that it doesn't cover as many cases as possible, such as if the tickers are not in order, if there are missing value in some rows, if the information provided in the future is in different format, if we want to conduct analysis on different variables in the future, and so on. It would be something to improve in the future. 
