Attribute VB_Name = "Module1"
Sub StockEvalutaion():

Dim ws As Worksheet
For Each ws In ThisWorkbook.Sheets

ws.Activate

    Dim Ticker As String
    Dim TableRow As Integer
    Dim StockVolume As Double
    Dim BeginYearDate As Long
    Dim EndYearDate As Long
    Dim BeginYearValue As Double
    Dim EndYearValue As Double
    Dim YearlyChange As Double
    Dim TableValue As Double
    
    BeginYearDate = 99999999
    EndYearDate = 0
    BeginYearValue = 0
    EndYearValue = 0
    TableRow = 2
    StockVolume = 0
    TableValue = 0
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    For i = 2 To LastRow
    
        If Cells(i, 1).Value <> Cells(i + 1, 1) Then
            Cells(TableRow, 9).Value = Cells(i, 1).Value
            StockVolume = StockVolume + Cells(i, 7).Value
            Cells(TableRow, 12).Value = StockVolume
            
            If Cells(i, 2).Value < BeginYearDate Then
                BeginYearDate = Cells(i, 2).Value
                BeginYearValue = Cells(i, 3).Value
            End If
            
            If Cells(i, 2).Value > EndYearDate Then
                EndYearDate = Cells(i, 2).Value
                EndYearValue = Cells(i, 6).Value
            End If
            
            'Cells(TableRow, 13).Value = EndYearValue
            'Cells(TableRow, 14).Value = BeginYearValue
            YearlyChange = EndYearValue - BeginYearValue
            Cells(TableRow, 10).Value = YearlyChange
            
            If BeginYearValue = 0 Then
                Cells(TableRow, 11).Value = False
            Else
                Cells(TableRow, 11).Value = (YearlyChange / BeginYearValue)
                Cells(TableRow, 11).NumberFormat = "0.00%"
            End If
            
            If Cells(TableRow, 10).Value > 0 Then
                Cells(TableRow, 10).Interior.ColorIndex = 4
            ElseIf Cells(TableRow, 10).Value < 0 Then
                Cells(TableRow, 10).Interior.ColorIndex = 3
            End If
            
            StockVolume = 0
            BeginYearDate = 999999999
            EndYearDate = 0
            TableRow = TableRow + 1
        
        ElseIf Cells(i, 1).Value = Cells(i + 1, 1) Then
            StockVolume = StockVolume + Cells(i, 7).Value
            
            If Cells(i, 2).Value < BeginYearDate Then
                BeginYearDate = Cells(i, 2).Value
                BeginYearValue = Cells(i, 3).Value
            End If
            
            If Cells(i, 2).Value > EndYearDate Then
                EndYearDate = Cells(i, 2).Value
                EndYearValue = Cells(i, 6).Value
            End If
        
        End If
    
    Next i
    
    LastRow = Cells(Rows.Count, 9).End(xlUp).Row
    TableValue = 0
    TableRow = 2
    
    For i = 2 To LastRow
            
            If Cells(i, 11).Value > TableValue Then
                TableValue = Cells(i, 11).Value
                TickerValue = Cells(i, 9).Value
                
            End If
        
    Next i
            Cells(TableRow, 17).Value = TableValue
            Cells(TableRow, 17).NumberFormat = "0.00%"
            Cells(TableRow, 16).Value = TickerValue
            TableRow = TableRow + 1
            TableValue = 0
    
    For i = 2 To LastRow
            
            If Cells(i, 11).Value < TableValue Then
                TableValue = Cells(i, 11).Value
                TickerValue = Cells(i, 9).Value
                
            End If
        
    Next i
            Cells(TableRow, 17).Value = TableValue
            Cells(TableRow, 17).NumberFormat = "0.00%"
            Cells(TableRow, 16).Value = TickerValue
            TableRow = TableRow + 1
    
    For i = 2 To LastRow
            
            If Cells(i, 12).Value > TableValue Then
                TableValue = Cells(i, 12).Value
                TickerValue = Cells(i, 9).Value
                
            End If
        
    Next i
            Cells(TableRow, 17).Value = TableValue
            Cells(TableRow, 16).Value = TickerValue


Next ws


End Sub


