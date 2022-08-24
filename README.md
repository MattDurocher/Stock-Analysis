# Stock-Analysis
## Overview of Project
### Purpose
The purpose of this project was to find out if stocks were worth investing in or not for Steve. We looked at stocks from the years of 2017 and 2018 using code in VBA that pulled data from multiple dates for the 11 stocks. This data had points from from multiple days for each stock in a given year in order to give us an accurate read on how the stock did year over year. This was then outputted into a table thanks to the VBA code. 
## Results
### Analysis
When looking at the outcomes from the two years, the return in 2017 is far greater than 2018. In 2017, every stock besides TERP posted a postive return but in 2018, only ENPH and RUN posted positive returns. This data was outputted thanks to creating an array at the beginning of the code along with creating a loop so that the code would comb through each and every data point. Dim tickerVolumes(12) As Long, Dim tickerStartingPrices(12) As Single, and Dim tickerEndingPrices(12) As Single all created the three output arrays. After creating a loop that set all the volumes to zero, code was compiled to loop over the whole spread sheet. After that, the results were gathered up and out putted in a table. 

  For i = 0 To 11
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
## Summary
### What are the advantages or disadvantages of refactoring code?
### How do these pros and cons apply to refactoring the original VBA script?
