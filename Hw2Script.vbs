Sub HW2()

Dim year_open As Double
Dim year_close As Double
Dim firstrow As Double
Dim PreviousAmount As Double
Dim percent_change As Double
PreviousAmount = 2
Range("I1") = "Ticket"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Volume"


LastRow = Cells(Rows.Count, 1).End(xlUp).Row
lastrow2 = Cells(Rows.Count, 3).End(xlUp).Row
Row = 2
Sum = 0


For i = 2 To LastRow
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        Sum = Sum + Cells(i, 7).Value
        ' puts name into cell
        Cells(Row, 9).Value = Cells(i, 1).Value
        Cells(Row, 12).Value = Sum
        year_open = Range("C" & PreviousAmount)
        year_close = Range("F" & i)
        yearly_change = year_close - year_open
        Cells(Row, 10).Value = yearly_change
        Cells(Row, 11) = 0
        percent_change = ((year_close - year_open) / year_open) * 100
        percent_change = Round(percent_change, 2)
        Cells(Row, 11) = percent_change
        Row = Row + 1
        Sum = 0
        PreviousAmount = i + 1
    Else
        Sum = Sum + Cells(i, 7).Value


    End If
Next i
  

End Sub