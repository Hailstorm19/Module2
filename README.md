# Module2
Module2 Homework
Sub Module2():
    
    'Set initial variable for ticker
    Dim Ticker_Name As String
    
    'set Inital variable for holding the total credit card brand
    Dim Ticker_Total As Double
    Ticker_Total = 0
    
    'Define Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Obtain last row
    Dim LR As Long
    LR = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
    
    'Loop Through all ticker
    For i = 2 To LR
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker_Total = Brand_Total + Cells(i, 3).Value
            'Place in ticker column
            
    
End Sub
