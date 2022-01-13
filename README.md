# Module_02_Challenge
Module 02 Challenge - Stock Analysis in VBA

## Overview of Project
Module 2 Challenge. This project involves using VBA within Microsoft Excel to analyze data historical data relating to a selection of stocks.

### Purpose
The purpose of this challenge is to provide the end user with the total number of shares traded and the annual percentage return (or loss) for a selection of stocks for a specified year.  

## Results
The overwhelming majority of the stocks included in this analysis performed significantly better in 2017 than in 2018; for 2017, only one stock (TERP) failed to appreciate in value over the course of the year, while in 2018, only two stocks (RUN and ENPH) appreciated in value, with each of the other stocks we analyzed finishing the year well below their per-share price to begin the year.     

### Tables Detailing Annualized Performance
![2017](VBA_Challenge_2017.png)

![2018](VBA_Challenge_2018.png)

## Refactoring Code
As part of this assignment, the code that was originally written earlier in the module was refactored in an attempt to improve performance; the primary change in approach related to the use of addiitional arrays to store values for tickers, tickerVolumes, tickerStartingPrices and tickerEndingPrices.

### Original Code - green_stocks.xslm
    'Create Variables for Starting and Ending Prices
    Dim startingPrice As Single
    Dim endingPrice As Single
    
    'Activate Data Worksheet
    'Worksheets("2018").Activate
    Sheets(yearValue).Activate
    
    'Get Number of Rows to Loop Through
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Loop Through Tickers
    
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
        'Activate Data Worksheet
        'Worksheets("2018").Activate
        Sheets(yearValue).Activate
                
        'Loop Through Rows in Data
        For j = 2 To RowCount

### Refactored Code
Multiple approaches were utilized while attempting to refactor the original code - the primary difference between the two approaches is in the iterator that is utilized to loop through each of the '''tickers''' - Approach A utilizes i as an iterator to control the loop, whereas Approach B uses 

**Attempt A**

'Activate Data Worksheet
    Worksheets(yearValue).Activate
    
    'Get Number of Rows to Loop Through
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Create tickerIndex
    Dim tickerIndex As Integer
    
    'Create Output Arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
      
    'Loop Through Tickers
    
    For i = 0 To 11
       
        ticker = tickers(tickerIndex)
        tickerVolumes(tickerIndex) = 0
        
        'Activate Data Worksheet
        Worksheets(yearValue).Activate
                
        'Loop Through Rows in Data
        For j = 2 To RowCount
        
            'Find Total Volume for Current Ticker
            If Cells(j, 1).Value = ticker Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
                'totalVolume = totalVolume + Cells(j, 8).Value
            End If

**Attempt B**

    'Activate Data Worksheet
    Worksheets(yearValue).Activate
    
    'Get Number of Rows to Loop Through
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Create tickerIndex
    Dim tickerIndex As Integer
    
    'Create Output Arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
      
    'Loop Through Tickers
    For tickerIndex = 0 To 11
    
        ticker = tickers(tickerIndex)
        tickerVolumes(tickerIndex) = 0
        
        'Activate Data Worksheet
         Worksheets(yearValue).Activate
                
        'Loop Through Rows in Data
        For j = 2 To RowCount
        
            'Find Total Volume for Current Ticker
            If Cells(j, 1).Value = ticker Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
               End If
               

### Analysis of Outcomes Based on Goals

### Challenges and Difficulties Encountered

## Results

- What are two conclusions you can draw about the Outcomes based on Launch Date

- What can you conclude about the Outcomes based on Goals?

- What are some limitations of this dataset?

- What are some other possible tables and/or graphs that we could create?
