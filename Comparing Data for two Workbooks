# code- Using VBA Script 
# Extract the data from one workbook to another and increase the frequency of the existing data in that workbook

Sub frequency_check()
Dim wrk As Worksheet, wrk1 As Worksheet, wkb As Workbook, wkb1 As Workbook, r1 As Range, r2 As Range
Set wkb = Workbooks("vba")
Set wkb1 = Workbooks("vba1")
Set wrk = wkb.Worksheets("Sheet8")
Set wrk1 = wkb1.Worksheets("Sheet1")
Set r1 = wrk.Range("A1")
Set r2 = wrk1.Range("A1")
wrk1.Activate
For i = Range("b1").Count To Range("b1", Range("b1").End(xlDown)).Count
For j = wrk.Range("b1").Count To (wrk.Range("b1", wrk.Range("b1").End(xlDown)).Count + 1)
Set r1 = wrk.Range("A" & j)
Set r2 = wrk1.Range("A" & i)
If r1 = r2 Then
wrk.Range("B" & j) = WorksheetFunction.sum(wrk.Range("B" & j), wrk1.Range("B" & i))
GoTo line2
ElseIf r1 <> Empty And r1 <> r2 Then
GoTo line1
Else
wrk1.Range("A" & i).Copy
wrk.Range("A1").End(xlDown).Offset(1, 0).PasteSpecial
wrk.Range("B" & j) = wrk1.Range("B" & i)
GoTo line2
End If
line1:
Next j
line2:
Next i
End Sub
