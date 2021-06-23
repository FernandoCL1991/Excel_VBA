Sub Stock_Analysis()

For Each ws In Worksheets

    ' Naming Headers
    ws.Cells.Range("I1").Value = "Ticker"
    ws.Cells.Range("J1").Value = "Yearly Change"
    ws.Cells.Range("K1").Value = "Percent Change"
    ws.Cells.Range("L1").Value = "Total Stock Volume"
    ws.Cells.Range("O2").Value = "Greatest Percentage Increase"
    ws.Cells.Range("O3").Value = "Greatest Percentage Decrease"
    ws.Cells.Range("O4").Value = "Greatest Total Stock Volume"
    ws.Cells.Range("P1").Value = "Ticker"
    ws.Cells.Range("Q1").Value = "Value"
    ws.Cells.Range("I:Q").EntireColumn.AutoFit
    ' Selecting last row of data
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    ' Setting up variables
    Dim Ticker As String
    Dim tickerTotal As Double
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    ' Setting counter
    tickerTotal = 0
    ' Results Table
    Dim summaryTableRow As Integer
    summaryTableRow = 2
    ' Dummy variable
    Dim openeingPriceRow As Double
    openingPriceRow = 2
    Dim maxValue As Long
    Dim minValue As Long
    Dim maxTotalStock As Long
    Dim maxTicker As String
    Dim minTicker As String
    Dim maxTotalTicker As String
    maxValue = -999999999
    minValue = 999999999
    
    ' Starting For Loop!
    For i = 2 To RowCount
        
        ' If Ticker is different:
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' Calculations for Ticker
            tickerName = ws.Cells(i, 1).Value
            ws.Cells(summaryTableRow, 9).Value = tickerName
            ' Calculations for Ticker Counter
            tickerTotal = tickerTotal + ws.Cells(i, 7)
            ws.Cells(summaryTableRow, 12).Value = tickerTotal
            ' Calculations for Opening Price
            openingPrice = ws.Cells(i + openingPriceRow, 3).Value
            ' Calculations for Closing Price
            closingPrice = ws.Cells(i, 6).Value
            
            ' Calculations for Yearly Change
            yearlyChange = closingPrice - openingPrice
            ws.Cells(summaryTableRow, 10).Value = yearlyChange
            
            ' Color formatting
            If yearlyChange <= 0 Then
                ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 3 ' Color Red
            Else
                ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 4 ' Color Green
            End If
            
            ' Calculations for Percent Change
            If openingPrice <= 0 Then
                percentChange = 0
                ws.Cells(summaryTableRow, 11).Value = percentChange
            Else
                percentChange = yearlyChange / openingPrice
                ws.Cells(summaryTableRow, 11).Value = percentChange
            End If
            
            ' Filling last summary table
            
            ' Greatest Percentage Increase & Ticker
            If maxValue < percentChange Then
                maxValue = percentChange
                maxTicker = tickerName
            End If
            ' Greatest Percentage Decrease & Ticker
            If minValue > percentChange Then
                minValue = percentChange
                minTicker = tickerName
            End If
            ' Greatest Total Stock Volume Ticker & Value
            ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
            volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)
            ws.Range("P4") = ws.Cells(volume_number + 1, 9)


            ' Adding a line of data per iteration
            summaryTableRow = summaryTableRow + 1
            ' Reseting the Ticker Counter as it is a new/ different Ticker
            tickerTotal = 0
                
        ' If Ticker stays the same
        Else
            ' Calculations for Ticker Counter
            tickerTotal = tickerTotal + ws.Cells(i, 7)
        End If

    Next i
    
    ' Printing Greatest % Increase Ticker
    ws.Cells(2, 16).Value = maxTicker
    ' Printing Greatest % Increase Value
    ws.Cells(2, 17).Value = maxValue
    
    ' Printing Greatest % Decrease Ticker
    ws.Cells(3, 16).Value = minTicker
    ' Printing Greatest % Decrease Value
    ws.Cells(3, 17).Value = minValue
    
    Next ws
    
End Sub