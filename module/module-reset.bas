Sub module_reset()
'
' Initialize data

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsPivotData As Worksheet
    Dim wsForm As Worksheet
    Dim msgAnswer As Integer

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("DATA")
    Set wsPivotData = wb.Sheets("PIVOTDATA")
    Set wsForm = wb.Sheets("REPORT")

    '// Confirm Message
    msgAnswer = MsgBox("Are you sure you want to reset the data?", vbYesNo + vbQuestion, "Empty Sheet")
    '// msgAnswer = MsgBox("데이터를 초기화 하시겠습니까?", vbYesNo + vbQuestion, "Empty Sheet")
    
    If msgAnswer = vbNo Then
        Exit Sub
    End If

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False

    '// Clear DATA worksheet
    For K = 1 To wsData.Range("A1").CurrentRegion.Rows.Count
        wsData.Rows(K + 1).ClearContents
    Next

    '// Clear PIVOTDATA worksheet
    For K = 1 To wsPivotData.Range("A1").CurrentRegion.Rows.Count
        wsPivotData.Rows(K + 1).ClearContents
    Next
    
    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    '// Success Message
    MsgBox "Data has been initialized."
    '// MsgBox "데이터가 초기화 되었습니다."
    
    '// Load a new document macro
    module_new
    
End Sub