# VBA of Wall Street

## Project Overview
### Background
The project originated with a request for a tool to analyze the performance of stocks in the energy sector. After successfully demonstrating the capabilities of the VBA script previously written, the client requested the ability to analyze all stocks in the market in a similar fashion.

### Purpose
The purpose of the project was to refactor the existing VBA script to analyze data more efficiently, allowing more data to be analyzed in less time. Refactoring was particularly important as the client wished to greatly increase the dataset. It is expected the successfully refactored script will perform the same analysis on a dataset faster than the original script.

## Results
The refactored script resulted in an analysis of the desired data in significantly less time. The analysis was completed in just over 0.164 seconds for the year 2017, a reduction of over 91% from the original script. Likewise, the analysis was completed in just over 0.210 seconds for the year 2018, a reduction of 89%.

![VBA_Challenge_2017Original](https://user-images.githubusercontent.com/96216947/147621204-de5ac995-f806-4da8-8f13-277ba33e2d00.png) ![VBA_Challenge_2017](https://user-images.githubusercontent.com/96216947/147621207-91f48fe1-e71a-402f-b986-9cd245ac1c06.png)

![VBA_Challenge_2018Original](https://user-images.githubusercontent.com/96216947/147621211-9d0d47df-7564-4aa2-99da-887cd8a96806.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/96216947/147621215-20611a1a-27e6-4a7d-ba5b-2545d0368d57.png)

The gains in efficiency were primarily due how the script read the raw data. The original script selected a given ticker and then proceeded to search each line of data for that ticker. This resulted in the entire dataset being read for each ticker and several times to complete the analysis. A sample of the original script is below.

For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0

    For j = 2 To RowCount

    If Cells(j, 1).Value = ticker Then
      totalVolume = totalVolume + Cells(j, 8).Value
    End If

In contrast, the refactored script created several arrays and placed data from each row into those arrays. The script was able to do this by using a variable for each ticker to index the data in each array. As a result, the entire dataset only needed to be read one time for all tickers, drastically reducing the time to process the data. A sample of the refactored script is below.

For j = 2 To RowCount
            If Cells(j, 1).Value = tickers(tickerIndex) Then
          tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        End If
               
         If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value        
         End If

The refactored script also included formatting changes to help visualize the performance of the stocks. Performance of those stocks for 2017 and 2018 are below. As can be seen, the year 2017 was much more fruitful for investors.

![StockPerformance_2017-2018](https://user-images.githubusercontent.com/96216947/147622230-18f87fd7-94e4-4b14-a4f0-3588d882c6d2.JPG)

## Summary
### Advantages of Refactoring


### Disadvantages of Refactoring

There is a detailed statement on the advantages and disadvantages of refactoring code in general
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.
