Attribute VB_Name = "Module1"
Sub YearlySummary(ws As Worksheet)
    
    ' Declare working variables
    Dim Ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    Dim outputRow As Integer
    Dim cellValue As Double
    
    ' Declare the variables for the greatest values
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    ' Initialize the output row
    outputRow = 2

    ' Find the last row of the sheet
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

       
    ' Initialize the "greatest" values
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    ' Set the headers and greatest values
    ws.Cells(1, "K").Value = "Ticker"
    ws.Cells(1, "L").Value = "Yearly Change"
    ws.Cells(1, "M").Value = "Percent Change"
    ws.Cells(1, "N").Value = "Total Volume"
    ws.Cells(1, "Q").Value = "Ticker"
    ws.Cells(1, "R").Value = "Value"
    ws.Cells(2, "P").Value = "Greatest % Increase"
    ws.Cells(3, "P").Value = "Greatest % Decrease"
    ws.Cells(4, "P").Value = "Greatest Total Volume"

    ' Loop through all rows
    For i = 2 To lastRow
    
        ' Check if we have hit a new ticker symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
           ' Set the Ticker, the closing price and calculate the yearly change
            Ticker = ws.Cells(i, 1).Value
            openingPrice = ws.Cells(i, 3).Value
            closingPrice = ws.Cells(i, 6).Value
            yearlyChange = closingPrice - openingPrice
            
                        
            ' Calculate the percentage change and round it to two decimal places
            If openingPrice <> 0 Then
                percentageChange = Round((yearlyChange / openingPrice) * 100, 2)
            Else
                percentageChange = 0
            End If

            ' Calculate the total volume
            totalVolume = Application.WorksheetFunction.SumIf(ws.Range("A:A"), Ticker, ws.Range("G:G"))
            
            ' Print Ticker, yearly change, percentage change and total volume to the K, L, M and N columns
            ws.Cells(outputRow, "K").Value = Ticker
            ws.Cells(outputRow, "L").Value = yearlyChange
            ' Apply conditional formatting based on the yearly change
            If yearlyChange > 0 Then
                ws.Cells(outputRow, "L").FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0").Interior.Color = RGB(0, 255, 0) ' Green
            ElseIf yearlyChange < 0 Then
                ws.Cells(outputRow, "L").FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0").Interior.Color = RGB(255, 0, 0) ' Red
            End If
'
            ' Resume with Coluns M and N
            ws.Cells(outputRow, "M").Value = percentageChange
            ws.Cells(outputRow, "N").Value = totalVolume
            
            ' Move to the next row
            outputRow = outputRow + 1
            
            ' Check if this ticker has the greatest % increase, % decrease, or total volume
            If percentageChange > greatestIncrease Then
                greatestIncrease = percentageChange
                greatestIncreaseTicker = Ticker
            ElseIf percentageChange < greatestDecrease Then
                greatestDecrease = percentageChange
                greatestDecreaseTicker = Ticker
            End If
            
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = Ticker
            End If
            
            ' Set openingPrice for the new ticker
            openingPrice = ws.Cells(i + 1, 3).Value
        End If
        
        
    Next i
    
           
    ' Print the greatest % increase, % decrease, and total volume to the Q and R columns
    ws.Cells(2, "Q").Value = greatestIncreaseTicker
    ws.Cells(2, "R").Value = greatestIncrease
    ws.Cells(3, "Q").Value = greatestDecreaseTicker
    ws.Cells(3, "R").Value = greatestDecrease
    ws.Cells(4, "Q").Value = greatestVolumeTicker
    ws.Cells(4, "R").Value = greatestVolume
        

            
 End Sub
 
 Sub allWorksheets()
 
 ' Initialize variables
 Dim ws As Worksheet
 
 ' Create loop to
    For Each ws In Worksheets
        YearlySummary ws
    Next ws
 
 End Sub
