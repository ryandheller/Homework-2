Attribute VB_Name = "Module2"
Sub totalz()
    tot_vol = 0
    current_summary_row = 2
    Cells(2, 9).Value = "A"
    Cells(1, 9).Value = "Tickers"
    Cells(1, 10).Value = "Total Volume"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    open_vol = Cells(2, 6).Value
    For Row = 2 To 797711
        current_tick = Cells(Row, 1).Value
        next_tick = Cells(Row + 1, 1).Value
        vol = Cells(Row, 7).Value
        If current_tick = next_tick Then
            tot_vol = tot_vol + vol
        Else
            close_vol = Cells(Row, 6).Value
            vol_change = close_vol - open_vol
            tot_vol = tot_vol + vol
            Cells(current_summary_row, 10).Value = tot_vol
            Cells(current_summary_row, 11).Value = vol_change
            per_change = (vol_change) / (open_vol + 0.0001)
            Cells(current_summary_row, 12).Value = per_change
            current_summary_row = current_summary_row + 1
            Cells(current_summary_row, 9).Value = Cells(Row + 1, 1).Value
            tot_vol = 0
            open_vol = Cells(Row + 1, 6).Value
            close_vol = 0
            vol_change = 0
            per_change = 0
        End If
       
    Next Row
End Sub
