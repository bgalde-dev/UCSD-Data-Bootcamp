Sub TickerAnalysis()
    
    ' Created variables for Ticker, Opening Price, Closing Price, Total Stock Volume
    Dim WorksheetName As String
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim tickerMarker As Long
    Dim greatestIncTicker As String
    Dim greatestIncValue As Double
    Dim greatestDecTicker As String
    Dim greatestDecValue As Double
    Dim greatestVolTicker As String
    Dim greatestVolValue As Double
    

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
        
        ' Initialize variables
        ticker = ""
        openPrice = 0
        closePrice = 0
        totalVolume = 0
        tickerMarker = 2
        greatestIncTicker = ""
        greatestIncValue = 0
        greatestDecTicker = ""
        greatestDecValue = 0
        greatestVolTicker = ""
        greatestVolValue = 0
        ' Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Initialize the headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Loop through each row
        For i = 2 To lastRow
            If ticker = "" Then
                ' Set new ticker, totalVolume, and openPrice
                ticker = ws.Cells(i, 1)
                openPrice = ws.Cells(i, 3)
                totalVolume = ws.Cells(i, 7)

            ElseIf ws.Cells(i, 1).Value <> ticker Then
                ' Do calculations
                ws.Cells(tickerMarker, 9).Value = ticker
                ws.Cells(tickerMarker, 10).Value = (closePrice - openPrice)
                ' Make sure you don't have a divide by zero
                If openPrice > 0 Then
                    ws.Cells(tickerMarker, 11).Value = (closePrice - openPrice) / openPrice
                Else
                    ws.Cells(tickerMarker, 11).Value = 0
                End If                
                ws.Cells(tickerMarker, 12).Value = totalVolume
                
                ' Set the greatest increase and decrease
                If ws.Cells(tickerMarker, 11).Value <> 0 And ws.Cells(tickerMarker, 11).Value >= greatestIncValue Then
                    greatestIncValue = ws.Cells(tickerMarker, 11).Value
                    greatestIncTicker = ticker
                ElseIf ws.Cells(tickerMarker, 11).Value < greatestDecValue Then
                    greatestDecValue = ws.Cells(tickerMarker, 11).Value
                    greatestDecTicker = ticker
                End If
                
                ' Set the greatest volume
                If ws.Cells(tickerMarker, 12).Value > greatestVolValue Then
                    greatestVolValue = ws.Cells(tickerMarker, 12).Value
                    greatestVolTicker = ticker
                End If
                               
                ' Set new ticker, totalVolume, and openPrice
                ticker = ws.Cells(i, 1)
                openPrice = ws.Cells(i, 3)
                totalVolume = ws.Cells(i, 7)
                ' Increment tickerMarker
                tickerMarker = tickerMarker + 1
            Else
                ' Update closePrice and totalVolume
                closePrice = ws.Cells(i, 6)
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            End If
            
        Next i
        
        ' Format data
        lastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        Dim dataRange As Range
        'dataRange = Range("J2:J" & lastRow)
        Set dataRange = ws.Range("J2:J" & lastRow)
        dataRange.FormatConditions.Delete
        dataRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0"
        dataRange.FormatConditions(1).Interior.Color = RGB(0, 255, 0)
        dataRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        dataRange.FormatConditions(2).Interior.Color = RGB(255, 0, 0)
        
        ws.Columns("K").NumberFormat = "0.00%"
        
        ' Add bonus assignment
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = greatestIncTicker
        ws.Cells(2, 16).Value = greatestIncValue
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = greatestDecTicker
        ws.Cells(3, 16).Value = greatestDecValue
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = greatestVolTicker
        ws.Cells(4, 16).Value = greatestVolValue
        
        ' Autofit to display data
        ws.Columns("A:P").AutoFit

    Next ws

    MsgBox ("Completed")

End Sub