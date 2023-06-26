
Sub VBAChallenge():

    Dim ws As Worksheet
    For Each ws In Worksheets
    Dim Total As Double
    Dim i As Long
    Dim YearlyChange As Double
    Dim j As Long
    Dim Start As Long
    Dim Rowcount As Long
    Dim Percentchange As Double
    Dim Days As Integer
    Dim Dailychange As Double
    Dim Averagechange As Double
    Dim Ticker As String
    Dim Row As Long

    
    Row = 2
    
    Rowcount = ws.Range("A" & Rows.Count).End(xlUp).Row
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    
    Change = 0
    Start = 2
    
    For i = 2 To Rowcount
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        Total = Total + ws.Cells(i, 7).Value
        YearlyChange = (ws.Cells(i, 6) - ws.Cells(Start, 3))
        Percentchange = YearlyChange / ws.Cells(Start, 3)
        Start = i + 1
        ws.Range("I" & Row).Value = Ticker
        ws.Range("L" & Row).Value = Total
        ws.Range("K" & Row).Value = Percentchange
        ws.Range("K" & Row).NumberFormat = "0.00%"
        ws.Range("J" & Row).Value = YearlyChange
            If ws.Range("J" & Row).Value > 0 Then
            ws.Range("J" & Row).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & Row).Value < 0 Then
            ws.Range("J" & Row).Interior.ColorIndex = 3
            End If
        
        Row = Row + 1
        Total = 0
        Else
        Total = Total + ws.Cells(i, 7).Value
        End If
        Next i
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & Rowcount))
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & Rowcount))
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & Rowcount))
        Maxindex = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & Rowcount)), ws.Range("K2:K" & Rowcount), 0)
        Minindex = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & Rowcount)), ws.Range("K2:K" & Rowcount), 0)
        Maxvolumeindex = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Rowcount)), ws.Range("L2:L" & Rowcount), 0)
        ws.Range("P2").Value = ws.Cells(Maxindex + 1, 9)
        ws.Range("P3").Value = ws.Cells(Minindex + 1, 9)
        ws.Range("P4").Value = ws.Cells(Maxvolumeindex + 1, 9)
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        Next ws

       
End Sub