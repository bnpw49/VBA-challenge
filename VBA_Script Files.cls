VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub multiple_yearstock():
    Dim ws As Worksheet
    For Each ws In Worksheets
        Dim total As Double
        Dim j As Integer
    
'Column_labels
        ws.Range("I1").Value = "Ticker"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
       j = 0
        Start = 2
    
    'Loop
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        For i = 2 To RowCount
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                ws.Range("I" & 2 + j).Value = ticker
                volume = ws.Cells(i, 7).Value
                ws.Range("L" & 2 + j).Value = volume + total
            
'Calculate Price Change. Yearly and Percentage
                Change = (ws.Cells(i, 6).Value - ws.Cells(Start, 3).Value)
                percentage_change = Change / ws.Cells(Start, 3).Value
                ws.Range("J" & 2 + j).Value = Change
                ws.Range("J" & 2 + j).NumberFormat = "0.00"
                ws.Range("K" & 2 + j).Value = percentage_change
                ws.Range("K" & 2 + j).NumberFormat = "0.00%"
        
                Select Case Change
                    Case Is > 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select
                
'start of the next ticker
                Start = i + 1
                total = 0
                j = j + 1
        
                Else
        
                volume = ws.Cells(i, 7).Value
                total = total + volume
        
            End If
        
        Next i
        
'find the maximum, minimum % change and max volume and assign them
        greatest_increase = WorksheetFunction.Max(ws.Range("K2:K" & RowCount))
        greatest_decrease = WorksheetFunction.Min(ws.Range("K2:K" & RowCount))
        greatest_volume = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
        
        ws.Range("Q2") = "%" & greatest_increase * 100
        ws.Range("Q3") = "%" & greatest_decrease * 100
        ws.Range("Q4") = greatest_volume

'find ticker symbol
        inc_loc = WorksheetFunction.Match(greatest_increase, ws.Range("K2:K" & RowCount), 0)
        dec_loc = WorksheetFunction.Match(greatest_decrease, ws.Range("K2:K" & RowCount), 0)
        vol_loc = WorksheetFunction.Match(greatest_volume, ws.Range("L2:L" & RowCount), 0)

'assign them to the right cells
        ws.Range("P2") = ws.Cells(inc_loc + 1, 9)
        ws.Range("P3") = ws.Cells(dec_loc + 1, 9)
        ws.Range("P4") = ws.Cells(vol_loc + 1, 9)
        
    Next ws

End Sub
