Sub TotalStockVolumeSub()

For Each ws In Worksheets

'Add headers to new columns
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker Symbol"
    ws.Cells(1, 16).Value = "Value"

'Declare and initialize variables
    Dim TotalVolume As Double
        TotalVolume = 0
    Dim SummaryTableRow As Double
        SummaryTableRow = 2
    Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim Stockticker As String
    Dim OpenPrice As Double
        OpenPrice = ws.Cells(2, 3).Value
    Dim ClosePrice As Double
    Dim MaxP As Double
        MaxP = 0
    Dim MinP As Double
        MinP = 0
    Dim GreatestVolume As Double
        GreatestVolume = 0
    Dim MaxPTicker As String
    Dim MinPTicker As String
    Dim GreatestVolumeTicker As String

'Loop through rows to find the opening/closing prices of each ticker and do the math
    For i = 2 To LastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            Stockticker = ws.Cells(i, 1).Value
            ClosePrice = ws.Cells(i, 6).Value
            TotalVolume = TotalVolume + (ws.Cells(i, 7).Value)
            ws.Cells(SummaryTableRow, 10).Value = (ClosePrice - OpenPrice)
                If (ClosePrice - OpenPrice) = 0 Or OpenPrice = 0 Then
                    ws.Cells(SummaryTableRow, 11).Value = 0
                    Else: ws.Cells(SummaryTableRow, 11).Value = ((ClosePrice - OpenPrice) / OpenPrice)
                End If
            ws.Range("K2:K2000").NumberFormat = "0.00%"
            ws.Cells(SummaryTableRow, 9).Value = Stockticker
            ws.Cells(SummaryTableRow, 12).Value = TotalVolume
            SummaryTableRow = SummaryTableRow + 1
            TotalVolume = 0
            OpenPrice = ws.Cells(i + 1, 3).Value
        Else
            TotalVolume = (TotalVolume) + (ws.Cells(i, 7).Value)
        End If
    Next i
    
'Conditional formatting for Yearly Change Column
    For l = 2 To ws.Cells(Rows.Count, 10).End(xlUp).Row
        If ws.Cells(l, 10) > 0 Then
        ws.Cells(l, 10).Interior.ColorIndex = 4
        Else
        ws.Cells(l, 10).Interior.ColorIndex = 3
        End If
    Next l
    
'Find the greatest/least % changes and greatest volume of all stock tickers
    For m = 2 To ws.Cells(Rows.Count, 11).End(xlUp).Row
        If ws.Cells(m, 11).Value > MaxP Then
        MaxP = ws.Cells(m, 11).Value
        MaxPTicker = ws.Cells(m, 9).Value
        End If
        If ws.Cells(m, 11).Value < MinP Then
        MinP = ws.Cells(m, 11).Value
        MinPTicker = ws.Cells(m, 9).Value
        End If
    Next m
    For n = 2 To ws.Cells(Rows.Count, 12).End(xlUp).Row
        If ws.Cells(n, 12) > GreatestVolume Then
        GreatestVolume = ws.Cells(n, 12).Value
        GreatestVolumeTicker = ws.Cells(n, 9).Value
        End If
    Next n
    ws.Cells(2, 16) = MaxP
    ws.Cells(2, 15) = MaxPTicker
    ws.Cells(3, 16) = MinP
    ws.Cells(3, 15) = MinPTicker
    ws.Cells(4, 15) = GreatestVolumeTicker
    ws.Cells(4, 16) = GreatestVolume
    ws.Range("P2:P3").NumberFormat = "0.00%"

Next ws
End Sub