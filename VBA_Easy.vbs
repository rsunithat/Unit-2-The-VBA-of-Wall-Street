Sub stockvol()

totalvol = 0
 
Dim currentrow As Integer
currentrow = 2
 
Dim lastrow As Integer
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
 
For i = 2 To lastrow
  totalvol = totalvol + Cells(i, 7)
 
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
  ticker = Cells(i, 1).Value
 
  Cells(currentrow, 9).Value = ticker
  Cells(currentrow, 10).Value = totalvol
 
  currentrow = currentrow + 1
 
  totalvol = 0
 
  End If
 
 Next i