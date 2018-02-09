Attribute VB_Name = "Module1"
Sub insert()
Dim rep As String
For Each c In Worksheets("INSERT SHEET NAME HERE, LIKE Sheet1").Range("INSERT RANGE HERE, LIKE A1:A100").Cells
rep = Left(c, 2) & ":" & Mid(c, 3, 2) & ":" & Mid(c, 5, 2) & ":" & Mid(c, 7, 2) & ":" & Mid(c, 9, 2) & ":" & Right(c, 2)
c.Value = rep
Next c
End Sub