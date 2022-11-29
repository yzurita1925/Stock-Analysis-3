![2017](https://user-images.githubusercontent.com/115121417/204408309-daae38c3-c9f8-41e1-8c25-89fe939d86c7.JPG)
![PIC 1](https://user-images.githubusercontent.com/115121417/204408327-0f872101-53e1-45ed-896e-c40141118e63.JPG)
![VBA Challenge 2018](https://user-images.githubusercontent.com/115121417/204408347-e83d580d-ca47-4bce-a1c0-d884ecace030.JPG)
![2017](https://user-images.githubusercontent.com/115121417/204408371-3b8702a5-9caa-4280-b939-b7ec5e97bee0.JPG)
![PIC 1](https://user-images.githubusercontent.com/115121417/204408373-b3b6b20c-4d3a-4690-9ede-4ca90ea047c2.JPG)
![2017](https://user-images.githubusercontent.com/115121417/204408394-62879fc4-b4bd-47d6-b6e9-14084a86b070.JPG)
![PIC 1](https://user-images.githubusercontent.com/115121417/204408395-7571de7b-bed1-4e90-a1fa-931f76b574da.JPG)
![VBA Challenge 2018](https://user-images.githubusercontent.com/115121417/204408396-cbed3a6f-8bd1-4ab9-af03-2b369d32777b.JPG)
# Stock-Analysis-3
Stock Analysis challenge 2
Stock-Analysis with Excel - VBA
Overview of the Project
The main purpose of this project is to refactor a Microsoft excel VBA code to collect stock information for year 2018 and analyze the behavior of the stocks to determine if the stocks are worth investing. 
The stock data provided for the analysis includes 12 different stocks as well as the ticker value, the date of the stock, the opening value closing values, the low and high price and finally the volume of the stocks. 
Results 
As we refactored the code, I input the code to create the input inbox, the ticker array. 
Please see below the order of the code. 
'1a) Create a ticker Index
tickerIndex = 0
'1b) Create three output arrays
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
''2a) Create a for loop to initialize the tickerVolumes to zero.
' If the next row’s ticker doesn’t match, increase the tickerIndex.
For i = 0 To 11
tickerVolumes(i) = 0
tickerStartingPrices(i) = 0
tickerEndingPrices(i) = 0
Next i
''2b) Loop over all the rows in the spreadsheet.
For i = 2 To RowCount
'3a) Increase volume for current ticker
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
'3b) Check if the current row is the first row with the selected tickerIndex.
'If  Then
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
End If
'3c) check if the current row is the last row with the selected ticker
'If  Then
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
End If
'3d Increase the tickerIndex.
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
tickerIndex = tickerIndex + 1
End If
Next i
'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11
Worksheets("All Stocks Analysis").Activate
Cells(4 + i, 1).Value = tickers(i)
Cells(4 + i, 2).Value = tickerVolumes(i)
Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
Next i
'Formatting
Worksheets("All Stocks Analysis").Activate
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
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
End Sub

Summary 
Advantages and Disadvantages of refactoring a code. 
One of the biggest advantages of refactoring a code is the length of the process compared to creating a brand new code, as well as making the code more understandable and the cost effectiveness compare to preparing a brand new code. Finally, is the incremental software development. 
