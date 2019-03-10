Sub Stock()

'Variable Declarations

Dim ticker As String
Dim volume As Double
Dim i As Double, j As Double
Dim ws As Worksheet
Dim LR As Double
Dim openP As Double
Dim closeP As Double
Dim PERCH As Double



'For Every Sheet
For Each ws In Worksheets

'Last Row of Column A
LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
LL = ws.Cells(Rows.Count, 12).End(xlUp).Row

'Updating Row Header
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"

'Setting initial value
volume = 0
openP = ws.Cells(2, 3).Value
closeP = 0
j = 2

'For loop for the rows

For i = 2 To LR

'If loop to check against next row and update final columns

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
volume = volume + ws.Cells(i, 7).Value
ws.Cells(j, 10).Value = volume
ws.Cells(j, 11).Value = closeP - openP

If openP <> 0 Then
PERCH = ws.Cells(j, 11).Value / openP
ws.Cells(j, 12).Value = PERCH
ws.Cells(j, 12).NumberFormat = "0.00%"
ElseIf openP = 0 Then
ws.Cells(j, 12).Value = "#N/A"


End If
If ws.Cells(j, 11).Value >= 0 Then
ws.Cells(j, 11).Interior.ColorIndex = 4
ElseIf ws.Cells(j, 11).Value < 0 Then
ws.Cells(j, 11).Interior.ColorIndex = 3
End If

openP = ws.Cells(i + 1, 3).Value

j = j + 1
volume = 0

Else

volume = volume + ws.Cells(i, 7).Value
closeP = ws.Cells(i, 6).Value

'Ending all loops

End If
Next i
Next ws

End Sub


