Attribute VB_Name = "Module_1"
Sub SubstName()
    Dim CurrentCell As Range, AllCells As Range, CurrentString As String, CurrentArray As Variant, i As Integer
    Set AllCells = Selection
    For Each CurrentCell In AllCells
        'Convert the string in current cell to array
        CurrentString = CurrentCell.Value
        ReDim CurrentArray(0 To Len(CurrentString) - 1)
        For i = 1 To Len(CurrentString) Step 1
            CurrentArray(i - 1) = Mid(CurrentString, i, 1)
        Next i
        'Convert character
        For i = 1 To UBound(CurrentArray) Step 1
            If CurrentArray(i) <> ";" And CurrentArray(i - 1) <> ";" Then
                CurrentArray(i) = "ĳ"
            End If
        Next i
        'Convert array to string and replace current cell
        CurrentString = Join(CurrentArray, "")
        CurrentCell.Value = CurrentString
        'clear array
        Erase CurrentArray
    Next CurrentCell
End Sub
