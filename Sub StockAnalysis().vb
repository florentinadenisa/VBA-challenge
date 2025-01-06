Sub StockAnalysis()
'Define all variables

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String

    ' Loop through each worksheet
    For Each ws In Worksheets
        ws.Activate
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        startRow = 2

        ' Create header for summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Initialize variables for greatest increase/decrease
        greatestIncrease = -999999
        greatestDecrease = 999999
        greatestVolume = 0

        ' Loop through all rows to analyze each stock
        Do While startRow <= lastRow
            ticker = ws.Cells(startRow, 1).Value
            openPrice = ws.Cells(startRow, 3).Value
            totalVolume = 0

            Do While ws.Cells(startRow, 1).Value = ticker And startRow <= lastRow
                closePrice = ws.Cells(startRow, 6).Value
                totalVolume = totalVolume + ws.Cells(startRow, 7).Value
                startRow = startRow + 1
            Loop

            ' Calculate quarterly change and percent change
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentChange = (quarterlyChange / openPrice) * 100
            Else
                percentChange = 0
            End If

            ' Populate summary table
            ws.Cells(summaryRow, 9).Value = ticker
            ws.Cells(summaryRow, 10).Value = quarterlyChange
            ws.Cells(summaryRow, 11).Value = percentChange
            ws.Cells(summaryRow, 12).Value = totalVolume

            ' Conditional formatting for positive/negative change
            If quarterlyChange > 0 Then
                ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
            ElseIf quarterlyChange < 0 Then
                ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
            End If

            ' Track greatest increase, decrease, and volume
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If

            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If

            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If

            summaryRow = summaryRow + 1
        Loop

        ' Output greatest increase, decrease, and volume
        ws.Cells(1, 14).Value = "Greatest % Increase"
        ws.Cells(2, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = greatestIncreaseTicker
        ws.Cells(2, 15).Value = greatestDecreaseTicker
        ws.Cells(3, 15).Value = greatestVolumeTicker
        ws.Cells(1, 16).Value = greatestIncrease & "%"
        ws.Cells(2, 16).Value = greatestDecrease & "%"
        ws.Cells(3, 16).Value = greatestVolume
    Next ws
End Sub

