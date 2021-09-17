Public Sub linked_checkbox_creator()
    Dim list As New collection
    Dim rng As Range
    Dim cell As Range
    Dim cb As CheckBox
     
    list.Add Application.InputBox(prompt:="Cellák ahová kerüljenek jelölőnégyzetek:", Type:=8)
    
    If Not TypeOf list(1) Is Range Then
        Exit Sub
    End If
    
    Set rng = list(1)
    
    For Each cell In rng
        Set cb = ActiveSheet.CheckBoxes.Add(cell.Left, cell.Top, 20#, 15#)
        cb.Text = ""
        cb.LinkedCell = cell.Address
        cell.Value = False
    Next cell
End Sub
