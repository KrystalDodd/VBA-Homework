# VBA-Homework


#Code for VBA Homework
Sub MultiYearStockData()

Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
        WS.Activate
Dim TickerSymbol As String
Dim YearlyChange As Double
Dim Table_Row As Integer
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim Counter As Long
Dim PercentChange As Double
Dim Column As Double
Dim TotalStockVolume As Long

Cells(1, 9).Value = "Ticker"
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
TableRow = 2
Counter = 2
TotalStockVolume = 0

For i = 2 To Lastrow
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        TickerSymbol = Cells(i, 1).Value
    Range("I" & TableRow).Value = TickerSymbol
    OpenPrice = Cells(Counter, 3).Value
    ClosePrice = Cells(i, 6).Value
    YearlyChange = ClosePrice - OpenPrice
    Range("J" & TableRow).Value = YearlyChange
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
        If PercentChange = ((ClosePrice - OpenPrice) / OpenPrice) Then
        Range("K" & TableRow).Value = PercentChange
            ElseIf ClosePrice = 0 And OpenPrice = 0 Then
                PercentChange = 0
    Range("K" & TableRow).NumberFormat = "0.00%"
    Cells(1, 12).Value = "TotalStockVolume"
    Range("L" & TableRow).Value = TotalStockVolume
    TotalStockVolume = TotalStockVoume + Cells(i, 7).Value
    TableRow = TableRow + 1
    Counter = Counter + 1
            End If
        End If
             
    If Cells(i, 11).Value > 0 Then
        Cells(i, 11).Interior.ColorIndex = 4
            ElseIf Cells(i, 11).Value < 0 Then
                Cells(i, 11).Interior.ColorIndex = 3
                
         End If
      

     Next i
        
  Next WS

End Sub
