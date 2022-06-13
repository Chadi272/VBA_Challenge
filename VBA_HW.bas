Attribute VB_Name = "Module3"
Sub VBA_HW()

' Definig variables
Dim tickr As String
Dim rowu As Integer
Dim openp As Double
Dim closep As Double
Dim lrow As Integer
Dim Ticker_Total As Double

' Starting row position
rowu = 2

' Assigning values to the variables
Ticker_Total = 0
lrow = Cells(Rows.Count, 9).End(xlUp).row

' Adding titles
Range("I1") = "Tickers"
Range("J1") = "Yearly Change"
Range("K1") = "Percentage Change"
Range("L1") = "Total Volume"
Range("M1") = "Greatest % increase"
Range("M2") = "Greatest % decrease"
Range("M3") = "Greatest total volume"

' Looping to pick unique ticker values
For i = 2 To 22771
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        tickr = Cells(i, 1).Value
        Range("I" & rowu).Value = tickr
        closep = Cells(i, 6)
        Range("J" & rowu) = closep - openp
        Cells(i, 10).NumberFormat = "0.00"
        Range("K" & rowu) = (closep - openp) / openp
                If Cells(rowu, 11).Value >= 0 Then
                Cells(rowu, 11).Interior.ColorIndex = 4
                Else: Cells(rowu, 11).Interior.ColorIndex = 3
                End If
        rowu = rowu + 1
            ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        openp = Cells(i, 3)
    End If
    Cells(i, 11).NumberFormat = "0.00%"
Next i


' Total Ticker Volume
For i = 2 To lrow
    Cells(i, 12) = Application.WorksheetFunction.SumIf(Range("A2:A22771"), Cells(i, 9), Range("G2:G22771"))
    Cells(i, 12).NumberFormat = "0,0"
Next i

' Greatest % increase
Range("O1") = Application.WorksheetFunction.Max(Range("K2:K" & lrow))
Range("O1").NumberFormat = "0.00%"
For i = 2 To lrow
    If Range("O1") = Cells(i, 11) Then
    Range("N1") = Cells(i, 9)
    End If
Next i

' Greatest % decrease
Range("O2") = Application.WorksheetFunction.Min(Range("K2:K" & lrow))
Range("O2").NumberFormat = "0.00%"
For i = 2 To lrow
    If Range("O2") = Cells(i, 11) Then
    Range("N2") = Cells(i, 9)
    End If
Next i

' Greatest total volume
Range("O3") = Application.WorksheetFunction.Max(Range("L2:L" & lrow))
Range("O3").NumberFormat = "0,0"
For i = 2 To lrow
    If Range("O3") = Cells(i, 12) Then
    Range("N3") = Cells(i, 9)
    End If
Next i

End Sub


