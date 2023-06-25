Attribute VB_Name = "Module1"
Sub stocks()
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
ws.Activate


Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "percent change"
Cells(1, 12).Value = "Total Stock Volume"
Range("o1").Value = "Greatest Percent increase"
Range("p1").Value = "Greatest percent decrease"
Range("q1").Value = "Greatest total volume"
Range("n2").Value = "Value"
Range("n3").Value = "Ticker"

Dim tablerow As Integer
Dim startnum As Double
Dim endnum As Double
Dim change As Double
Dim volume As Integer



tablerow = 2


lastrow = Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To lastrow
    
    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        Cells(tablerow, 9).Value = Cells(i, 1).Value
        startnum = Cells(i, 3).Value
        volume = 0
    End If
    
    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        volumes = volumes + Cells(i, 7).Value
        
    End If
    
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        endnum = Cells(i, 6).Value
        change = Str(endnum - startnum)
        Cells(tablerow, 10).Value = Str(change)
            If Cells(tablerow, 10).Value > 0 Then
                Cells(tablerow, 10).Interior.ColorIndex = 4
            Else
                If Cells(tablerow, 10).Value < 0 Then
                Cells(tablerow, 10).Interior.ColorIndex = 3
            End If
            End If
        Cells(tablerow, 11).Value = Str(100 * (change / startnum))
        Cells(tablerow, 12).Value = Str(volumes)
        tablerow = tablerow + 1
    End If
    
    Next i
 Range("o2").Value = Cells(2, 11).Value
 Range("o3").Value = Cells(2, 9).Value
  Range("p2").Value = Cells(2, 11).Value
 Range("p3").Value = Cells(2, 9).Value
  Range("q2").Value = Cells(2, 12).Value
 Range("q3").Value = Cells(2, 9).Value
 For i = 3 To 3001
    If Cells(i, 11).Value > Range("o2").Value Then
        Range("O2").Value = Cells(i, 11).Value
        Range("O3").Value = Cells(i, 9).Value
    End If
    If Cells(i, 11).Value < Range("p2").Value Then
        Range("P2").Value = Cells(i, 11).Value
        Range("P3").Value = Cells(i, 9).Value
    End If
    If Cells(i, 12).Value > Range("q3").Value Then
        Range("Q2").Value = Cells(i, 12).Value
        Range("Q3").Value = Cells(i, 9).Value
    End If
    Next i
Range("n5").Value = " "
Next ws

End Sub
