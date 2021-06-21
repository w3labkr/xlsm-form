' 워크 시트에서 유효성을 제거하려면 없음을 선택하십시오.
' To remove the validation from the worksheet, select NONE.
'
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Skip
    
    Dim KeyCells1 As Range
    Dim KeyCells2 As Range
    Dim KeyCells3 As Range
    Dim KeyCells4 As Range

    ' The variable KeyCells contains the cells that will
    ' cause an alert when they are changed.
    
    '// List1
    Set KeyCells1 = Range("C14:C33")
    If Not Application.Intersect(KeyCells1, Range(Target.Address)) _
           Is Nothing Then
        If Target.Value = "NONE" Then
           Target.Validation.Delete
        End If
    End If

    '// List2
    Set KeyCells2 = Range("D14:D33")
    If Not Application.Intersect(KeyCells2, Range(Target.Address)) _
           Is Nothing Then
        If Target.Value = "NONE" Then
           Target.Validation.Delete
        End If
    End If

    '// List3
    Set KeyCells3 = Range("G14:G33")
    If Not Application.Intersect(KeyCells3, Range(Target.Address)) _
           Is Nothing Then
        If Target.Value = "NONE" Then
           Target.Validation.Delete
        End If
    End If

    '// List4
    Set KeyCells4 = Range("H14:H33")
    If Not Application.Intersect(KeyCells4, Range(Target.Address)) _
           Is Nothing Then
        If Target.Value = "NONE" Then
           Target.Validation.Delete
        End If
    End If

    '// List5
    Set KeyCells4 = Range("H9:J9")
    If Not Application.Intersect(KeyCells4, Range(Target.Address)) _
           Is Nothing Then
        If Target.Value = "NONE" Then
           Target.Validation.Delete
        End If
    End If

    '// List6
    Set KeyCells4 = Range("H10:J10")
    If Not Application.Intersect(KeyCells4, Range(Target.Address)) _
           Is Nothing Then
        If Target.Value = "NONE" Then
           Target.Validation.Delete
        End If
    End If

Done:
    Exit Sub
Skip:
End Sub