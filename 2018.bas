Attribute VB_Name = "Module3"
Sub Year2018()

' Definig variables
Dim tickr As String
Dim row As Integer
Dim openp As Double
Dim closep As Double
Dim ych As Double
Dim starty As Double
Dim endy As Double
Dim rowy As Integer
Dim rowy1 As Integer
Dim lrow As Integer
Dim STR As Integer
Dim Ticker_Total As Double

' Starting row position
rowu = 2
rowy = 2
rowy1 = 2
STR = 2

' Assigning values to the variables
starty = "20180102"
endy = "20181231"
lrow = Cells(Rows.Count, 9).End(xlUp).row

' Adding titles
Range("I1") = "Tickers"
Range("J1") = "Open Price"
Range("K1") = "Close Price"
Range("L1") = "Yearly Change"
Range("M1") = "Percentage Change"
Range("N1") = "Total Volume"
Range("O1") = "Greatest % increase"
Range("O2") = "Greatest % decrease"
Range("O3") = "Greatest total volume"

' Looping to pick unique ticker values
For i = 2 To 22771
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    tickr = Cells(i, 1).Value
    Range("I" & rowu).Value = tickr
    rowu = rowu + 1
    End If
Next i

' Open price for each ticker
For j = 2 To lrow
    For i = 2 To 22771
    If (Cells(i, 2).Value = starty And Cells(i, 1).Value = Cells(j, 9)) Then
    openp = Cells(i, 3)
    Range("j" & rowy) = openp
    rowy = rowy + 1
End If
Next i
Next j

' Close price for each ticker
For j = 2 To lrow
    For i = 2 To 22771
    If (Cells(i, 2).Value = endy And Cells(i, 1).Value = Cells(j, 9)) Then
    closep = Cells(i, 6)
    Range("K" & rowy1) = closep
    rowy1 = rowy1 + 1
    End If
Next i
Next j
Columns("J:K").Hidden = True

' Calculate Yearly Change
For i = 2 To lrow
Cells(i, 12).Value = Cells(i, 11).Value - Cells(i, 10).Value
Cells(i, 12).NumberFormat = "0.00"
Next i

' Percentage change
For i = 2 To lrow
Cells(i, 13).Value = ((Cells(i, 11).Value - Cells(i, 10).Value) / Cells(i, 10).Value)
Cells(i, 13).NumberFormat = "0.00%"
    If Cells(i, 13).Value >= 0 Then
    Cells(i, 13).Interior.ColorIndex = 4
    Else: Cells(i, 13).Interior.ColorIndex = 3
    End If
Next i

' Total Ticker Volume
For i = 2 To lrow
    Cells(i, 14) = Application.WorksheetFunction.SumIf(Range("A2:A22771"), Cells(i, 9), Range("G2:G22771"))
    Cells(i, 14).NumberFormat = "0,0"
Next i

' Greatest % increase
Range("Q1") = Application.WorksheetFunction.Max(Range("M2:M" & lrow))
Range("Q1").NumberFormat = "0.00%"
For i = 2 To lrow
    If Range("Q1") = Cells(i, 13) Then
    Range("P1") = Cells(i, 9)
    End If
Next i

' Greatest % decrease
Range("Q2") = Application.WorksheetFunction.Min(Range("M2:M" & lrow))
Range("Q2").NumberFormat = "0.00%"
For i = 2 To lrow
    If Range("Q2") = Cells(i, 13) Then
    Range("P2") = Cells(i, 9)
    End If
Next i

' Greatest total volume
Range("Q3") = Application.WorksheetFunction.Max(Range("N2:N" & lrow))
Range("Q3").NumberFormat = "0,0"
For i = 2 To lrow
    If Range("Q3") = Cells(i, 14) Then
    Range("P3") = Cells(i, 9)
    End If
Next i

End Sub


