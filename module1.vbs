Attribute VB_Name = "Module1"
Sub stockmarket():
    Dim ws As Worksheet
    Dim LastrowA As Long
    Dim LastrowC As Long
    Dim LastrowF As Long
    Dim LastrowL As Long
    Dim i As Long
    Dim j As Long
    Dim MaxVolume As Double
    
    For Each ws In Worksheets
        
        LastrowA = ws.Cells(Rows.Count, "A").End(xlUp).Row
        Range("I2:I" & LastrowA).Value = Range("A2:A" & LastrowA).Value
        LastrowC = ws.Cells(Rows.Count, "C").End(xlUp).Row
        LastrowF = ws.Cells(Rows.Count, "F").End(xlUp).Row
        LastrowL = ws.Cells(Rows.Count, "L").End(xlUp).Row
        Range("i1") = "Ticker"
        Range("j1") = "Yearly Change"
        Range("K1") = "Percent Change(%)"
        Range("L1") = "Total Stock Volume"
        Range("P1") = "Ticker"
        Range("Q1") = "Value"
        Range("O2") = "Greatest % Increase"
        Range("O3") = "Greatest % decrease"
        Range("O4") = "Greatest Total Volume"
        
        For i = 2 To LastrowC
            Cells(i, "J").Value = Cells(i, "F").Value - Cells(i, "C").Value
            If Cells(i, "j").Value > 0 Then
                Cells(i, "j").Interior.Color = vbGreen
            Else
                Cells(i, "j").Interior.Color = vbRed
            End If
        Next i
        
        For j = 2 To LastrowC
            If Cells(j, "C").Value >= Cells(j, "F").Value Then
                Cells(j, "K").Value = (1 - (Cells(j, "F").Value / Cells(j, "C").Value)) * 100
                
            Else
                Cells(j, "K").Value = (1 - (Cells(j, "C").Value / Cells(j, "F").Value)) * 100
                
            End If
            Cells(j, "K").NumberFormat = "0.00"
        Next j
        
        
        For k = 2 To LastrowF
            Cells(k, "L").Value = Cells(k, "G").Value * Cells(k, "F").Value
        Next k
        
        maxvalue = WorksheetFunction.Max(Range("L2:L" & LastrowL))
        Range("Q4").Value = maxvalue

    Next ws
End Sub

