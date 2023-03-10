Attribute VB_Name = "Module1"
Sub Multiple_Stock_Year()


'Define all  variables

Dim Ticker As String
Dim year_open As Double
Dim year_close As Double
Dim Yearly_Change As Double
Dim Total_Stock_Volume As Double
Dim Percent_Change As Double
Dim start As Double

Dim ws As Worksheet

For Each ws In Worksheets

'Assign a column header

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Assign integer for loop to start
start = 2
Previous = 1
Total_Stock_Volume = 0


EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row


For i = 2 To EndRow


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Ticker symbol

Ticker = ws.Cells(i, 1).Value

Previous = Previous + 1


year_open = ws.Cells(Previous, 3).Value
year_close = ws.Cells(i, 6).Value

For j = Previous To i


Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value

Next j


If year_open = 0 Then

Percent_Change = year_close

Else
Yearly_Change = year_close - year_open

Percent_Change = Yearly_Change / year_open

End If

'values in the worksheet summary table

ws.Cells(start, 9).Value = Ticker
ws.Cells(start, 10).Value = Yearly_Change
ws.Cells(start, 11).Value = Percent_Change

'script for percentage format

ws.Cells(start, 11).NumberFormat = "0.00%"
ws.Cells(start, 12).Value = Total_Stock_Volume



start = start + 1

'reset to zero

Total_Stock_Volume = 0
Yearly_Change = 0
Percent_Change = 0

'Move i to previous
Previous = i

End If


Next i

'second summary table


Per_change = ws.Cells(Rows.Count, "K").End(xlUp).Row

'Define variable to start second summary table

Increase = 0
Decrease = 0
Greatest = 0

'find max/min for percentage
For k = 3 To Per_change


last_k = k - 1

'current row for percentage
current_k = ws.Cells(k, 11).Value

'Previous row for percentage
previous_k = ws.Cells(last_k, 11).Value

'greatest total volume row
volume = ws.Cells(k, 12).Value

'Previous greatest volume row
previous_vol = ws.Cells(last_k, 12).Value


'Find increase
If Increase > current_k And Increase > previous_k Then

Increase = Increase


ElseIf current_k > Increase And current_k > previous_k Then

Increase = current_k

'define name for increase percentage
increase_name = ws.Cells(k, 9).Value

ElseIf previous_k > Increase And previous_k > current_k Then

Increase = previous_k

'define name for increase percentage
increase_name = ws.Cells(last_k, 9).Value

End If

'Find the decrease

If Decrease < current_k And Decrease < previous_k Then

'Define decrease

Decrease = Decrease

'Define name for increase percentage

ElseIf current_k < Increase And current_k < previous_k Then

Decrease = current_k


decrease_name = ws.Cells(k, 9).Value

ElseIf previous_k < Increase And previous_k < current_k Then

Decrease = previous_k

decrease_name = ws.Cells(last_k, 9).Value

End If

'Find the greatest volume

If Greatest > volume And Greatest > previous_vol Then

Greatest = Greatest

'name for greatest volume

ElseIf volume > Greatest And volume > previous_vol Then

Greatest = volume

'name for greatest volume
greatest_value = ws.Cells(k, 9).Value

ElseIf previous_vol > Greatest And previous_vol > volume Then

Greatest = previous_vol



End If

Next k

' Assign headers

ws.Range("N1").Value = "Column Name"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker Name"
ws.Range("P1").Value = "Value"

ws.Range("O2").Value = increase_name
ws.Range("O3").Value = decrease_name
ws.Range("O4").Value = greatest_value
ws.Range("P2").Value = Increase
ws.Range("P3").Value = Decrease
ws.Range("P4").Value = Greatest

'script for percent format

ws.Range("P2").NumberFormat = "0.00%"
ws.Range("P3").NumberFormat = "0.00%"





jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row


For j = 2 To jEndRow


If ws.Cells(j, 10) > 0 Then

ws.Cells(j, 10).Interior.ColorIndex = 4

Else

ws.Cells(j, 10).Interior.ColorIndex = 3
End If

Next j


Next ws

End Sub
