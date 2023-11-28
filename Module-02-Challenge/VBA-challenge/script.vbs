Sub CycleTicker()

    ' Prevents screen refreshing.
    Application.ScreenUpdating = False

    ' Declare variables
    Dim tickerListRow, greatestInc, greatestDec, GreatestVol as long
    Dim openPrice, closePrice As Double
    Dim MyRange As Range
    Dim tickerData() As String

    ' Cycle through worksheets
    For Each ws In Worksheets

        ' Get the hight of the cells
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        tickerListRow = 1

        'Collect all the data into an array
        reDim tickerData(2 To LastRow, 0 To 3)
        For stockRow = 2 To LastRow
            tickerData(stockRow, 0) = ws.Cells(stockRow, 1).Value 'Ticker
            tickerData(stockRow, 1) = ws.Cells(stockRow, 3).Value 'Open
            tickerData(stockRow, 2) = ws.Cells(stockRow, 6).Value 'Close
            tickerData(stockRow, 3) = ws.Cells(stockRow, 7).Value 'Volume
        Next stockRow

        ' Set up headers (make sure we can auto format)
        ws.Cells(1, 9).Value = "Ticker       "
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ws.Cells(1, 15).Value = "                                          "
        ws.Cells(1, 16).Value = "Ticker    "
        ws.Cells(1, 17).Value = "Value              "

        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"

        ' Cycle through Stocks
        For stockRow = 2 To LastRow

            ' If this row is the same as the current ticker
            If tickerData(stockRow, 0) = ws.Cells(tickerListRow, 9).Value Then
                closePrice = CDbl(tickerData(stockRow, 2))

                ws.Cells(tickerListRow, 10).Value = closePrice - openPrice
                ws.Cells(tickerListRow, 11).Value = FormatPercent((closePrice - openPrice) / openPrice, 2, True)
                ws.Cells(tickerListRow, 12).Value = ws.Cells(tickerListRow, 12).Value + CLngLng(tickerData(stockRow, 3))
            Else
                ' if not then add a ticker row
                tickerListRow = tickerListRow +1

                openPrice = CDbl(tickerData(stockRow, 1))
                closePrice = CDbl(tickerData(stockRow, 2))

                ws.Cells(tickerListRow, 9).Value = tickerData(stockRow, 0)
                ws.Cells(tickerListRow, 10).Value = openPrice - closePrice
                ws.Cells(tickerListRow, 11).Value = 0
                ws.Cells(tickerListRow, 12).Value = CLngLng(tickerData(stockRow, 3))
            End If
            'We are looking for a  change in the data

        Next stockRow

        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

        ' format cells
        Set MyRange = ws.Range("J2:K"+Trim(Str(LastRow)))
        MyRange.FormatConditions.Delete
        MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, _
                Formula1:="=0"
        MyRange.FormatConditions(1).Interior.Color = RGB(0, 255, 0)
        MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                Formula1:="=0"
        MyRange.FormatConditions(2).Interior.Color = RGB(255, 0, 0)

        ' get the Greatest increase, decrease & value and display them
        greatestInc = 2
        greatestDec = 2
        GreatestVol = 2

        For stockRow = 2 to LastRow
            If ws.Cells(stockRow, 11).Value > ws.Cells(greatestInc, 11).Value Then
                greatestInc = stockRow
            End If
            If ws.Cells(stockRow, 11).Value < ws.Cells(greatestDec, 11).Value Then
                greatestDec = stockRow
            End If
            If ws.Cells(stockRow, 12).Value > ws.Cells(GreatestVol, 12).Value Then
                GreatestVol = stockRow
            End If
        Next stockRow

        ws.Cells(2, 16).Value = ws.Cells(greatestInc, 9).Value
        ws.Cells(2, 17).Value = FormatPercent(ws.Cells(greatestInc, 11).Value, 2)

        ws.Cells(3, 16).Value = ws.Cells(greatestDec, 9).Value
        ws.Cells(3, 17).Value = FormatPercent(ws.Cells(greatestDec, 11).Value, 2)

        ws.Cells(4, 16).Value = ws.Cells(GreatestVol, 9).Value
        ws.Cells(4, 17).Value = ws.Cells(GreatestVol, 12).Value

        ws.Rows(1).Columns.AutoFit

    Next ws

    ' Enables screen refreshing.
    Application.ScreenUpdating = True
End Sub