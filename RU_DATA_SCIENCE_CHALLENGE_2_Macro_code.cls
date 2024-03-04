VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub TickerSymbol_YearlyChange_PercentChange_StockVolume()
    ' make an array '
   For Each ws In Worksheets
   
    Dim WorksheetName As String
        
    Dim i As Long
    
    Dim j As Long
    
    Dim TickerCounter As Long
    
    Dim LastRowA As Long
    Dim LastRowI As Long
    
    Dim PercentChange As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    
    WorksheetName = ws.Name
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    TickerCounter = 2
    j = 2
    
    LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRowA
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(TickerCounter, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(TickerCounter, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                If ws.Cells(TickerCounter, 10).Value < 0 Then
                ws.Cells(TickerCounter, 10).Interior.ColorIndex = 3
                Else
                ws.Cells(TickerCounter, 10).Interior.ColorIndex = 4
                End If
                
                If ws.Cells(j, 3).Value <> 0 Then
                PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                
                ws.Cells(TickerCounter, 11).Value = Format(PercentChange, "Percent")
                Else
                ws.Cells(TickerCounter, 11).Value = Format(0, "Percent")
                End If
                
             ws.Cells(TickerCounter, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
             TickerCounter = TickerCounter + 1
             j = i + 1
             
             End If
             
        Next i
        
     
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
       
        
        
        GreatestVolume = ws.Cells(2, 12).Value
        GreatestIncrease = ws.Cells(2, 11).Value
        GreatestDecrease = ws.Cells(2, 11).Value
        
            
            For i = 2 To LastRowI
                If ws.Cells(i, 12).Value > GreatVolume Then
                GreatestVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                Else
                GreatestVolume = GreatestVolume
                End If
                If ws.Cells(i, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                Else
                GreatestIncrease = GreatestIncrease
                End If
                If ws.Cells(i, 11).Value < GreatestDecrease Then
                GreatestDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                Else
                GreatestDecrease = GreatestDecrease
                End If
            ws.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatestVolume, "Scientific")
            Next i
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
    Next ws
        
End Sub
            
    
 


