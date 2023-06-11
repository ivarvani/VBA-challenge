Attribute VB_Name = "assigment_module2"
Sub for_looping()

Dim last_row As Long
Dim i As Long
Dim j As Long
Dim total_volume As Double
Dim open_price As Double
Dim difference As Double
Dim close_price As Double
Dim percentage_diff As Double
Dim greatest_total_volume As Double
'-------------------------------------------------------------------------------------------
'looping through worksheets

For Each ws In Worksheets

Dim WorksheetName As String

last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

WorksheetName = ws.Name

ws.Cells(1, 15).Value = "<ticker>"
ws.Cells(1, 16).Value = "<total volume>"
ws.Cells(1, 17).Value = "<Yearly change>"
ws.Cells(1, 18).Value = "<% change>"
ws.Cells(4, 11).Value = "<greatest % increase>"
ws.Cells(5, 11).Value = "<greatset % decrease>"
ws.Cells(6, 11).Value = "<greatest volume>"

 '-----------------------------------------------------------------------------------------

'to get the first column of ticker symbol
'to get the yearly change and the % change column
'j is the row value of where our result will inputed

total_volume = 0
j = 1


For i = 2 To last_row

If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

j = j + 1

ws.Cells(j, 15).Value = ws.Cells(i, 1).Value


open_price = ws.Cells(i, 3).Value


ElseIf ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value Then

close_price = ws.Cells(i, 6).Value

End If
difference = close_price - open_price
ws.Cells(j, 17).Value = difference
percentage_diff = (difference / open_price) * 100
ws.Cells(j, 18).Value = percentage_diff


'--------------------------------------------------------------------------------------
'to get the total volume column

If ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value Then
total_volume = total_volume + ws.Cells(i, 7).Value
ws.Cells(j, 16) = total_volume



ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
total_volume = ws.Cells(i, 7).Value

End If
Next
'-----------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------
'conditional formatting the yearly differece column

lastrow2 = ws.Cells(Rows.Count, "Q").End(xlUp).Row

For i = 2 To lastrow2

If ws.Cells(i, 17).Value > 0 Then
ws.Cells(i, 17).Interior.ColorIndex = 4
Else
ws.Cells(i, 17).Interior.ColorIndex = 3
End If
Next

'finding the highest increase  and decrease in percentage percentage
ws.Cells(4, 13).Value = 0 'declaring the initial value of the highest percentage increase to 0
ws.Cells(5, 13).Value = 0 'declaring the initial value of the highest percentage decrease to 0
greatest_total_volume = 0

For i = 2 To lastrow2
If ws.Cells(i, 18).Value > 0 And ws.Cells(i, 18).Value > ws.Cells(4, 13).Value Then
ws.Cells(4, 13).Value = ws.Cells(i, 18).Value
ws.Cells(4, 12).Value = ws.Cells(i, 15).Value

End If
If ws.Cells(i, 18).Value < 0 And ws.Cells(i, 18).Value < ws.Cells(5, 13) Then
ws.Cells(5, 13).Value = ws.Cells(i, 18).Value
ws.Cells(5, 12).Value = ws.Cells(i, 15).Value
End If
If ws.Cells(i, 16).Value > greatest_total_volume Then
ws.Cells(6, 12).Value = ws.Cells(i, 15).Value
greatest_total_volume = ws.Cells(i, 16).Value
End If
ws.Cells(6, 13).Value = greatest_total_volume




Next



Next ws







End Sub

