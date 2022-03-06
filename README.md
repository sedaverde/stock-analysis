# Stock Analysis using VBA and Excel

## Overview of Project

The intent of this project is to automate the process for the analysis of stock information collected for twelve different stocks over the course of a calendar year. Stock information was collected for the years 2017 and 2018. The script was written to automate the calculation of total daily volume and percent return in order to provide the user with data that could guide them to educated investment decisions.

## Purpose

The purpose of this project was to learn to refactor code in order to imrpove its efficiency. The idea is to write code that would require one pass through of the data to obtain all relevant information, instead of the original module code that was written to loop through all the data one time for each of the twelve different stocks.

## Results

### Code Comparison
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
       
       Worksheets("AllStockAnalysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
   ```
The original code has the program going through each ticker one at a time and entering the data in the appropriate cells once it has found the end of the ticker. Changes to the refactored code include:

- The creation of output arrays to hold all data collected until the loop is completed
  ```
  Dim tickerVolumes(12) As Long
  Dim tickerStartingPrices(12) As Single
  Dim tickerEndingPrices(12) As Single
  ```
- The creation of a loop that will increase the tickerIndex value once the end of the current tickerIndex has been reached while only passing through all the data in the file one time
  ```
    For i = 0 To 11
       tickerVolumes(i) = 0
    Next i
    
    For i = 2 To RowCount   
                   
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
         If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

         End If
                               
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
          
         tickerIndex = tickerIndex + 1
            
         End If
         
    Next i
    ```
- Alteration of output code to enter output array values into correct cells
  ```
  For i = 0 To 11
        Worksheets("AllStockAnalysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
  Next i
  ```
### Code Run Time Comparison

Code run time for original and refactored for 2017 data:

![image](https://user-images.githubusercontent.com/90329647/156912266-af52b0c5-ee79-4815-b8f2-af365f491c8d.png)
![image](https://user-images.githubusercontent.com/90329647/156912275-c2f98c25-0562-43f0-bb7b-b82432b3dea5.png)

Code run time for original and refactored for 2018 data:

![image](https://user-images.githubusercontent.com/90329647/156912296-5423389d-424e-4bf4-8453-d8c09256f4dc.png)
![image](https://user-images.githubusercontent.com/90329647/156912315-460e21b8-59e8-40cc-9178-6bd571da6434.png)

In each instance the code runs ~86% faster refactored.

## Summary

The biggest advantage of refactoring code is efficiency. In this case the refactored code required more actual code writing than the original simpler code.
