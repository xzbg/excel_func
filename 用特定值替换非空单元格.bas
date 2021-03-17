Attribute VB_Name = "用特定值替换非空单元格"
Sub ReplaceNumbers()
'Update 20141111
    Dim SRg As Range
    Dim Rg As Range
    Dim Str As Variant
    On Error Resume Next
    Set SRg = Application.Selection
    Set SRg = Application.InputBox("select range:", "Kutools for Excel", SRg.Address, , , , , 8)
    If Err <> 0 Then Exit Sub
    Str = Application.InputBox("replace with:", "Kutools for Excel", Str)
    If Str = False Then Exit Sub
    For Each Rg In SRg
        If Rg <> "" Then Rg = Str
    Next
End Sub
