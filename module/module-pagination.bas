Sub module_prev()
'
' Load previous data

    Dim wb As Workbook
    Dim wsForm As Worksheet
    Dim wsData As Worksheet
    Dim getSerial As Integer

    Set wb = ThisWorkbook
    Set wsForm = wb.Sheets("REPORT")

    '// Serial
    getSerial = wsForm.Range("D9")

    If getSerial = 1 Then
        MsgBox "No previous data."
        '// MsgBox "이전 데이터가 없습니다."
        Exit Sub  
    Else
        wsForm.Range("D9") = getSerial - 1
    End If

    '// Reinit
    module_load_init

    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True


End Sub

Sub module_next()
'
' Load next data

    Dim wb As Workbook
    Dim wsForm As Worksheet
    Dim wsData As Worksheet
    Dim getSerial As Integer
    Dim LastRow As Integer

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("DATA")
    Set wsForm = wb.Sheets("REPORT")

    '// Serial
    getSerial = wsForm.Range("D9")
    LastRow = wsData.Range("a1").CurrentRegion.Rows.Count - 1

    If getSerial >= LastRow Then
        MsgBox "No next data."
        '// MsgBox "다음 데이터가 없습니다."
        Exit Sub  
    Else
        wsForm.Range("D9") = getSerial + 1
    End If

    '// Reinit
    module_load_init

    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub