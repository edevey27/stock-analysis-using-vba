Sub mod2challenge()
'create a list of every ticker in Column H
'create quarterly change
'create percent change
'create total stock volume

'step1 - set my dimensions

Dim total As Double
Dim rowi As Long
Dim columni As Integer
Dim change As Double
Dim start As Long
Dim rowcount As Long
Dim percentchange As Double
Dim days As Integer
Dim dailychange As Single
Dim averagechange As Double
Dim ws As Worksheet


For Each ws In Worksheets
'this will make the code reset for every sheet
columni = 0
total = 0
start = 2
dailychange = 0
change = 0
'this will title in every sheet, ws makes that possible

ws.Range("I1").Value = "TICKER"
ws.Range("J1").Value = "Quarterly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % decrease"
ws.Range("O4").Value = "Greatest TSV"
ws.Range("P1").Value = "TICKER"
ws.Range("Q1").Value = "Result"
'this will now count every row and total it in each sheet

rowcount = ws.Cells(Rows.Count, "A").End(xlUp).Row
'create my index
For rowi = 2 To rowcount

'when the ticker changes it will print results

If ws.Cells(rowi + 1, 1).Value <> ws.Cells(rowi, 1).Value Then
total = total + ws.Cells(rowi, 7).Value

If total = 0 Then
ws.Range("I" & 2 + columni).Value = ws.Cells(rowi, 1).Value
ws.Range("J" & 2 + columni).Value = 0
ws.Range("K" & 2 + columni).Value = "%" & 0
ws.Range("L" & 2 + columni).Value = 0
Else
    If ws.Cells(start, 3) = 0 Then
    
    For find_value = start To rowi
    If ws.Cells(find_value, 3).Value <> 0 Then
    start = find_value
    Exit For
    End If
    
    Next find_value
    End If
    
    change = (ws.Cells(rowi, 6) - ws.Cells(start, 3))
    percentchange = change / ws.Cells(start, 3)
    start = rowi + 1
    ws.Range("I" & 2 + columni) = ws.Cells(rowi, 1).Value
    ws.Range("J" & 2 + columni) = change
    ws.Range("J" & 2 + columni).NumberFormat = "0.00"
    ws.Range("K" & 2 + columni).Value = percentchange
    ws.Range("K" & 2 + columni).NumberFormat = "0.00%"
    ws.Range("L" & 2 + columni).Value = total
    
  Select Case change
  Case Is > 0
    ws.Range("J" & 2 + columni).Interior.ColorIndex = 4
  Case Is < 0
  ws.Range("J" & 2 + columni).Interior.ColorIndex = 3
  Case Else
  ws.Range("J" & 2 + columni).Interior.ColorIndex = 0
  End Select
  
End If
total = 0
change = 0
columni = columni + 1
days = 0
dailychange = 0

Else
'this makes it so if ticker is still the same, then it will add results
    total = total + ws.Cells(rowi, 7).Value
    
End If

Next rowi
'take the max and min
ws.Range("Q2") = "%" & WorsksheetFunction.Max(ws.Range("K2:K" & rowcount)) * 100
ws.Range("Q3") = "%" & WorsksheetFunction.Min(ws.Range("K2:K" & rowcount)) * 100
ws.Range("Q4") = WorsksheetFunction.Max(ws.Range("L2:L" & rowcount))

increase = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)
decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)
volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowcount)), ws.Range("L2:L" & rowcount), 0)

ws.Range("P2") = ws.Cells(increase + 1, 9)
ws.Range("P3") = ws.Cells(decrease + 1, 9)
ws.Range("P4") = ws.Cells(volume + 1, 9)


Next ws
End Sub



