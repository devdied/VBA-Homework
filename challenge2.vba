
Sub stock()
Dim ws As Worksheet
For Each ws In Worksheets
With ws
.Range("I1").Value = "Ticker"
.Range("J1").Value = "Yearly Change"
.Range("K1").Value = "Percent Change"
.Range("L1").Value = "Total Stock Volume"
.Range("N2").Value = "Greatest % increase"
.Range("N3").Value = "Greatest % decrease"
.Range("N4").Value = "Greatest total volume"
.Range("O1").Value = "Ticker"
.Range("P1").Value = "Value"

Dim LR, opening, closing, tick_row, i, total, Great_inc, Great_dec, Great_vol As Variant
Dim percent As Double
Great_inc = 0
Great_dec = 0
Great_vol = 0
total = 0
.Cells(2, 9).Value = .Range("A2").Value
tick_row = 2
opening = .Range("C2").Value
LR = .Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To LR
    total = total + .Cells(i, 7).Value
    If .Cells(i, 1).Value <> .Cells(i + 1, 1).Value Then
        .Cells(tick_row + 1, 9) = .Cells(i + 1, 1).Value
        closing = .Cells(i, 6).Value
        .Cells(tick_row, 10).Value = closing - opening
        If .Cells(tick_row, 10).Value > 0 Then
        .Cells(tick_row, 10).Interior.ColorIndex = 4
        Else
        .Cells(tick_row, 10).Interior.ColorIndex = 3
        End If
        percent = (.Cells(tick_row, 10).Value / opening) * 100
        .Cells(tick_row, 11).Value = Str(percent) + "%"
        If percent > Great_inc Then
        Great_inc = percent
        .Range("O2") = .Cells(i, 1).Value
        .Range("P2") = Str(percent) + "%"
        End If
        If percent < Great_dec Then
        Great_dec = percent
        .Range("O3") = .Cells(i, 1).Value
        .Range("P3") = Str(percent) + "%"
        End If
        If total > Great_vol Then
        Great_vol = total
        .Range("O4") = .Cells(i, 1).Value
        .Range("P4") = total
        End If
        .Cells(tick_row, 12).Value = total
        opening = .Cells(i + 1, 3).Value
        total = 0
        tick_row = tick_row + 1
    End If
Next i
End With
Next ws
End Sub

