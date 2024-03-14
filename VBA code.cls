VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub MultipleYearStockData()
    Dim ws As Worksheet
    Dim wsName As String
    Dim initialCheck As Long
    Dim endCheck As Long
    Dim TickerCounter As Long
    Dim ColumnLastRow As Long
    Dim PercentageChange As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    
    For Each ws In Worksheets
        TickerCounter = 2
        endCheck = 2
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        
        wsName = ws.Name
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest Percentage Increase"
        ws.Cells(3, 15).Value = "Greatest Percentage Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
  
        ColumnLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        For initialCheck = 2 To ColumnLastRow
            
            If ws.Cells(initialCheck + 1, 1).Value <> ws.Cells(initialCheck, 1).Value Then
              
                ws.Cells(TickerCounter, 9).Value = ws.Cells(initialCheck, 1).Value
                
                ws.Cells(TickerCounter, 10).Value = ws.Cells(initialCheck, 6).Value - ws.Cells(endCheck, 3).Value
                
                If ws.Cells(TickerCounter, 10).Value < 0 Then
                    ws.Cells(TickerCounter, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(TickerCounter, 10).Interior.ColorIndex = 4
                End If
                
                If ws.Cells(endCheck, 3).Value <> 0 Then
                    PercentageChange = ((ws.Cells(initialCheck, 6).Value - ws.Cells(endCheck, 3).Value) / ws.Cells(endCheck, 3).Value)
                    ws.Cells(TickerCounter, 11).Value = Format(PercentageChange, "Percent")
                Else
                    ws.Cells(TickerCounter, 11).Value = Format(0, "Percent")
                End If
                
                ws.Cells(TickerCounter, 12).Value = WorksheetFunction.Sum(ws.Range(ws.Cells(endCheck, 7), ws.Cells(initialCheck, 7)))
                
                'Check for greatest increase
                If PercentageChange > GreatestIncrease Then
                    GreatestIncrease = PercentageChange
                    ws.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
                    ws.Cells(2, 16).Value = ws.Cells(initialCheck, 1).Value
                End If
                
                'Check for greatest decrease
                If PercentageChange < GreatestDecrease Then
                    GreatestDecrease = PercentageChange
                    ws.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
                    ws.Cells(3, 16).Value = ws.Cells(initialCheck, 1).Value
                End If
                
                'Check for greatest total volume
                If ws.Cells(TickerCounter, 12).Value > GreatestVolume Then
                    GreatestVolume = ws.Cells(TickerCounter, 12).Value
                    ws.Cells(4, 17).Value = GreatestVolume
                    ws.Cells(4, 16).Value = ws.Cells(initialCheck, 1).Value
                End If
                
                'Increase TickerCounter by 1
                TickerCounter = TickerCounter + 1
                
                'Set new start row of the ticker block
                endCheck = initialCheck + 1
            End If
        Next initialCheck
    Next ws
End Sub

