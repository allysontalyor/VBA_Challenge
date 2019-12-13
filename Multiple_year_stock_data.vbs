Attribute VB_Name = "Module1"
Sub stock_market()

'instruct the program to loop through all files
For Each ws In Worksheets

'put in header values
ws.Cells(1, 10).Value = "ticker"
ws.Cells(1, 11).Value = "yearly change"
ws.Cells(1, 12).Value = "% change"
ws.Cells(1, 13).Value = "total volume"

'put in labels for bonus problem
ws.Cells(1, 16).Value = "ticker"
ws.Cells(1, 17).Value = "value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'format width of cells
ws.Cells(2, 15).ColumnWidth = 30

'define variable ticker symbol
Dim ticker As String

'define variable yearly change
Dim yearly_change As Double
yearly_change = 0

'define variable percent change
Dim percent_change As Double
percent_change = 0

'define stock total volumes
Dim total_volume As Double
total_volume = 0

'define summary table row
Dim summary_table_row As Integer
summary_table_row = 2

'define last row variable
Dim last_row As Long
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

'define end row and start row
Dim start_row As Long
start_row = 2
Dim end_row As Long

'setup For statement to loop through all
For i = 2 To last_row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'specify what the ticker value should be
        ticker = ws.Cells(i, 1).Value
        
        'add the value to the total_volume
        total_volume = total_volume + ws.Cells(i, 7).Value
        
        'define what the end row is
        end_row = i
        
        'Calculate yearly change
        yearly_change = ws.Cells(end_row, 6).Value - ws.Cells(start_row, 3).Value
        
        'Calculate percent yearly change
        If ws.Cells(start_row, 3).Value = 0 Then
            percent_change = 0
        Else
            
            percent_change = (yearly_change / ws.Cells(start_row, 3).Value) * 100
        End If
        
         
        'assign the ticker value and the total_volume value to the summary chart
        ws.Range("J" & summary_table_row).Value = ticker
        ws.Range("M" & summary_table_row).Value = total_volume
        ws.Range("K" & summary_table_row).Value = yearly_change
        ws.Range("L" & summary_table_row).Value = percent_change
        
        'reset the total volume to 0
        total_volume = 0
        
        'reset end_row and start_row
        start_row = i + 1
        end_row = 0
        
        'add a row the the summary table row value!!
        summary_table_row = summary_table_row + 1
        
        Else
        total_volume = total_volume + ws.Cells(i, 7).Value
        End If
                
Next i

'define last row of summary chart
Dim last_summary_row As Long
last_summary_row = ws.Cells(Rows.Count, 10).End(xlUp).Row

For i = 2 To last_summary_row
        If ws.Cells(i, 11).Value < 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 3
        
        Else
        ws.Cells(i, 11).Interior.ColorIndex = 4
        
        End If
        
Next i

'Bonus

'Determining the maximum percent change
'define maximum change value and set to 0
Dim max_change As Double
max_change = 0

'setup loop to run through the summary row data
For i = 2 To last_summary_row
        
        'if statement comparing each value in the column to the max value
        If ws.Cells(i, 12).Value > max_change Then
        
        'the max_change becomes the new highest value
        max_change = ws.Cells(i, 12).Value
        
        'set the chart cell to become the max change value
        
        
        'put in the ticker value
        ws.Cells(2, 16).Value = ws.Cells(i, 10).Value
        
        Else
        
        ws.Cells(2, 17).Value = max_change
        
        End If
        
Next i

ws.Cells(2, 17).Value = max_change

'Determining maximum total volume
'define maximum total volume value and set to 0
Dim max_volume As Double
max_volume = 0

'setup loop to run through the summary row data
For i = 2 To last_summary_row
        
        'if statement comparing each value in the column to the max volume
        If ws.Cells(i, 13).Value > max_volume Then
        
        'the max_volume becomes the new highest value
        max_volume = ws.Cells(i, 13).Value
        
        
        'set the chart cell to become the max change value
        ws.Cells(4, 17).Value = max_volume
        
        'put in the ticker value
        ws.Cells(4, 16).Value = ws.Cells(i, 10).Value
        
        Else
        
        ws.Cells(4, 17).Value = max_volume
        
        End If
        
Next i

'Find the greatest percent decrease
'define the largest decrease
Dim max_decrease As Double
max_decrease = 0

'Setup loop to go through data in the summary row
For i = 2 To last_summary_row
        If ws.Cells(i, 12).Value < max_decrease Then
        max_decrease = ws.Cells(i, 12).Value
        ws.Cells(3, 17).Value = max_decrease
        ws.Cells(3, 16).Value = ws.Cells(i, 10).Value
        Else
        ws.Cells(3, 17).Value = max_decrease
        End If
Next i

Next
End Sub

