Sub Stock_Market_Analysis():

'Defining dimensions
Dim start As Long 'initial stock value
Dim rowCount As Long 'number of rows with data
Dim percentChange As Single 'percent change of value
Dim ws As Worksheet 'initializing code on each worksheet
Dim total As Double 'total stock volume
Dim i As Long 'i stands for itterator
Dim change As Single '"change" reflects change in Price
Dim j As Integer 'j stands for summary itterator

'For statement to begin running code on each worksheet
For Each ws In Worksheets

'Setting initial values
total = 0
j = 0
start = 2 'Values start at row 2
change = 0

'Setting header row titles
ws.Range("I1").Value = "Ticker Symbol"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker Symbol"
ws.Range("Q1").Value = "Ticker Value"
ws.Range("O2").Value = "Greatest Percent Increase"
ws.Range("O3").Value = "Greatest Percent Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Getting row number of the last row with data
rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

'Starting loop through worksheets
For i = 2 To rowCount

   'If the ticker changes then statement
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
      total = total + ws.Cells(i, 7).Value
      
      If total = 0 Then
        'If statement, print values
        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
        ws.Range("J" & 2 + j).Value = 0
        ws.Range("K" & 2 + j).Value = "%" & 0
        ws.Range("L" & 2 + j).Value = 0
      Else
      If ws.Cells(start, 3) = 0 Then
         For find_value = start To i
           If ws.Cells(find_value, 3).Value <> 0 Then
              start = find_value
         Exit For
        End If
      Next find_value
    End If
      
      'Calculating changes for each worksheet
      change = ws.Cells(i, 6) - ws.Cells(start, 3)
      percentChange = Round((change / ws.Cells(start, 3)) * 100, 2)
      
      'Continuing the next stock ticker
      start = i + 1
      
      ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
      ws.Range("J" & 2 + j).Value = Round(change, 2)
      ws.Range("K" & 2 + j).Value = "%" & percentChange
      ws.Range("L" & 2 + j).Value = total
      
      'Conditional formatting yearly change
      Select Case change
         Case Is > 0
           ws.Range("J" & 2 + j).Interior.ColorIndex = 4
         Case Is < 0
           ws.Range("J" & 2 + j).Interior.ColorIndex = 3
         Case Else
           ws.Range("J" & 2 + j).Interior.ColorIndex = 0
      End Select
     End If
     
    'Resetting variables for new ticker
     total = 0
     change = 0
     j = j + 1
    
    'If ticker stays the same
  Else
     total = total + ws.Cells(i, 7).Value
  End If
Next i

'Calculating ranges for greatest percent increase, greatest percent decrease, and greatest total volume
ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))

Greatest_Percent_Increase = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
Greatest_Percent_Decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
Greatest_Total_Volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)

ws.Range("P2") = ws.Cells(Greatest_Percent_Increase + 1, 9)
ws.Range("P3") = ws.Cells(Greatest_Percent_Decrease + 1, 9)
ws.Range("P4") = ws.Cells(Greatest_Total_Volume + 1, 9)

Next ws
End Sub
