# stock-analysis

## Overview of Project 
The purpose of this challenge was to compare if refactoring code will make our VBA script run faster, even with a larger dataset. Our friend Steve initially wanted to help his parents' analyze how DQ's stock performed over a year. After we had created a worksheet where Steve could analyze the performance of DQ, he now wants to widen his dataset to analyze the performance of the entire stock market. Since we already created the initial code, we refactor our code to run a larger data set. With the help of refactoring, our code was able to run at a similar rate as the original code. 

## Results 
Running the refactored code on the larger data set we were able to find that the 12 stocks Steve was interested in for his parents performed well in 2017. Only one of the stocks ** TERP (-7.2%)** had a negative for 2017. For 2018, these 10 of the 12 stocks had a negative return. The only stocks to have any growth were **ENPH (+81.9%)** and **RUN (+84.0%)**. 
Our original code overall ran faster than our refactored code. Although, our refactored code was able to analyze a larger data set, while incorporating more functions. The original code’s run time for  **Year 2017** dataset was **0.2900391 seconds**, while the refactored code’s runtime was **0.3583984 seconds ** . Similarly the run time for **Year 2018** data set for the original code was **0.2832031 seconds**, while the refactored code's runtime was **0.3603516 seconds**.  

###### 2017 and 2018 Returns

<img width="265" alt="2017 Returns" src="https://user-images.githubusercontent.com/107595578/177139068-4036fad0-9fef-4352-8e07-3701a3a60be3.png">

<img width="265" alt="2018 Returns" src="https://user-images.githubusercontent.com/107595578/177139078-d3562d6b-a3a0-45a7-9d42-b3b6d30bbfab.png">


###### Original Code Performance vs Refactored Code Preforemance for Year 2017

<img width="265" alt="Original_Code2017" src="https://user-images.githubusercontent.com/107595578/177132659-1d3d570b-0e81-4350-9d93-2cc61aaadb8c.png">

<img width="265" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/107595578/177133079-8e3c7043-f7f7-49fe-9b11-ab5f6f11677a.png">

###### Original Code Performance vs Refactored Code Preforemance for Year 2018

<img width="265" alt="Original_Code2018" src="https://user-images.githubusercontent.com/107595578/177132671-b94ae263-02cd-4210-9de6-e32381fb5014.png">

<img width="265" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/107595578/177133106-0bdaf207-0f08-4b46-8abb-38fac6854d53.png">

###### Original Code
     Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime As Single
   
    yearValue = InputBox("What year would you like to run the analysis on?")
        startTime = Timer
    Worksheets("AllStocksAnalysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
   
    Cells(3, 1).Value = "Tickers"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
   
    Dim tickers(12) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
  
    Dim startingPrice As Single
    Dim endingPrice As Single
   
     Worksheets(yearValue).Activate
   
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
 
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
      
       Worksheets("AllStocksAnalysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i
    endTime = Timer
    MsgBox "This code ran in  " & (endTime - startTime) & "  seconds for the year" & (yearValue)
    End Sub`

###### Refactored Code
     Sub AllStocksAnalysisRefactored()
      Dim startTime As Single
      Dim endTime As Single
    
    'Format the output sheet on All Stocks Analysis worksheet
    yearValue = InputBox("What year would you like to run the analysis on?")
        startTime = Timer
    Worksheets("AllStocksAnalysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Tickers"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String

    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
     Worksheets(yearValue).Activate
   
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0
    
    'Create three output arrays
    Dim tickerVolume(12) As Long
    Dim tickerStartingPrice(12) As Single
    Dim tickerEndingPrice(12) As Single
 
     'Create a for loop to initialize the tickerVolume to zero
    For i = 0 To 11
       tickerVolume(tickerIndex) = 0
    '
       Worksheets(yearValue).Activate
    'Loop over all the rows in the spreadsheet
    For j = 2 To RowCount
           'Increase volume for the current ticker
           If Cells(j, 1).Value = tickers(tickerIndex) Then
                tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(j, 8).Value
            End If
            
           'Check if the current row is the first row with the selected tickerIndex
           'If Then
           If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
                tickerStartingPrice(tickerIndex) = Cells(j, 6).Value
            End If
            
           'check if the current row is the lastrow with the selected ticker
           'If the next row's ticker doesn't match,increase the ticker Index
           'If Then
           If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrice(tickerIndex) = Cells(j, 6).Value
            End If
            
            'Increase the tickerIndex
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
       Next j
       Next i
    
    'Loop through your array to output the Ticker, Total Daily Volume, and Return
    For i = 0 To 11
        Worksheets("AllStocksAnalysis").Activate
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolume(i)
            Cells(4 + i, 3).Value = tickerEndingPrice(i) / tickerStartingPrice(i) - 1
       

     Next i
     'Formatting
    Worksheets("AllStocksAnalysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
    Next i
    endTime = Timer
    MsgBox "This code ran in  " & (endTime - startTime) & "  seconds for the year" & (yearValue)

    End Sub`





## Summary 
###### Advantages and Disadvantages of Refactoring Code 
Refactoring code is essentially using an existing code while revising details to fit new criteria that won’t change the codes overall behavior. There are many advantages of refactoring, two that stand out are saving time, addition to easier workflow. Refactoring saves time by not having to write your code from scratch. It also eases workflow by just altering the code to fit your new data criteria, if that is editing the code, expanding on exiting code. Some disadvantages of refactoring are you have this exiting code that you might have to expand upon if that is setting new conditions, or set new variable that might trigger numerous error codes. If you happen to encounter numerous errors, this could be more time comsuming than writtng a new code. 

###### Advantages and Disadvantages of the Original and Refactored VBA Script
Now comparing the original VBA script to the refactored VBA script, over all the original VBA was easier to follow step by step. I encounter numerous error code while trying to run the refactored code. Refactoring seem that it would be a strong skill set for those who are familar with VBA and coding. Since you are essentially hacking or altering the oringial code. You must be knowledgeable of all the functions. Refactoring does have is advantages point as well, in our case you were able to run one macro to achieve the same output that took the original code several macros to achieve. The original code seemed more straight forward, but the major disadvantage was that you need to go through several macros, which seemed tedious and redundant. 
