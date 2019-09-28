Attribute VB_Name = "Module1"
Sub stock_analysis_general()
'define variables for final output'

Dim ticker As String
Dim yearlychange As Double
Dim percentchange As Double
Dim totalvolume As Double
Dim yearopen As Double
Dim yearclose As Double

For Each ws In Worksheets

summary_table_row = 2

'label columns'
LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
ws.Cells(1, 9).Value = "ticker"
ws.Cells(1, 10).Value = "yearly_change"
ws.Cells(1, 11).Value = "percent_change"
ws.Cells(1, 12).Value = "total volume"
ws.Cells(1, 13).Value = "Year Close"
ws.Cells(1, 14).Value = "Year Open"

'Pull ticker and year end'
total_volume = 0
For i = 2 To LastRow
If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
' Add to the total volume'
totalvolume = totalvolume + ws.Cells(i, 7).Value
yearclose = ws.Cells(i, "F")

'Print Ticker'

ws.Range("I" & summary_table_row).Value = ws.Cells(i, 1)
ws.Range("L" & summary_table_row).Value = totalvolume
ws.Range("M" & summary_table_row).Value = yearclose
      

      ' Add one to the summary table row
      summary_table_row = summary_table_row + 1
      
      'Reset Total Volume'
      
      totalvolume = 0
      Else: totalvolume = totalvolume + Cells(i, 7).Value
      End If
      
      If ws.Cells(i, 1) = ws.Cells(i + 1, 1) And ws.Cells(i, 1) <> ws.Cells(i - 1, 1) Then
      
      
yearopen = ws.Cells(i, 3).Value
ws.Range("N" & summary_table_row).Value = yearopen

End If
Next i
    
    
For j = 2 To 10000
If ws.Cells(j, "N") > 0 And ws.Cells(j, "M") > 0 Then
ws.Cells(j, "K") = ws.Cells(j, "N") / ws.Cells(j, "M") - 1
ws.Cells(j, "K").NumberFormat = "0.00%"
ws.Cells(j, "J") = ws.Cells(j, "N") - ws.Cells(j, "M")
End If
If ws.Cells(j, "K") > 0 Or ws.Cells(j, "J") > 0 Then
ws.Cells(j, "K").Interior.ColorIndex = 4
ws.Cells(j, "J").Interior.ColorIndex = 4
Else: ws.Cells(j, "K").Interior.ColorIndex = 3
ws.Cells(j, "J").Interior.ColorIndex = 3
End If
If ws.Cells(j, "K") = 0 Then
ws.Cells(j, "K").ClearFormats
ws.Cells(j, "J").ClearFormats
End If
Next j
Next ws

End Sub
