Sub VolumeTester():

Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Dim ticker_line As Integer
ticker_line = 0

Dim ticker_volume As Double
ticker_volume = 0

Cells(1, 10).Value = "Ticker:"
Cells(1, 11).Value = "Total Volume:"


    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker_line = ticker_line + 1
            Cells(ticker_line + 1, 10).Value = Cells(i, 1).Value
            ' add to the ticker volume total
            ticker_volume = ticker_volume + Cells(i, 7).Value
            ' print the brand amount to the summary table
            Cells(ticker_line + 1, 11).Value = ticker_volume
            ' reset the ticker volume total
            ticker_volume = 0
        Else
            ' add to the ticker volume total
            ticker_volume = ticker_volume + Cells(i, 7).Value
        End If
    Next i
    
End Sub


