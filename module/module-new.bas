Sub module_new()
'
' Create a new document

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsForm As Worksheet
    Dim repeatRow As Integer
    Dim jumpedRow As Integer
    Dim msgAnswer As Integer

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("DATA")
    Set wsForm = wb.Sheets("REPORT")

    '// Confirm Message
    msgAnswer = MsgBox("Would you like to create a new document?", vbYesNo + vbQuestion, "Empty Sheet")
    '// msgAnswer = MsgBox("새 문서를 만드시겠습니까?", vbYesNo + vbQuestion, "Empty Sheet")
    
    If msgAnswer = vbNo Then
        Exit Sub
    End If

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False

    '// Static
    With wsForm
        .Range("D9")  = wsData.Range("a1").CurrentRegion.Rows.Count '// Serial
        .Range("V3")  = "=D9"         '// Serial
        .Range("D10") = "=TODAY()"    '// Date
        .Range("V4")  = "=ROUNDUP(MONTH(D10)/3,0)&""Q""" '// Quarter
        .Range("V5")  = "=YEAR(D10)"   '// Year
        .Range("V6")  = "=MONTH(D10)"  '// Month
        .Range("V7")  = "=DAY(D10)"    '// Day
        .Range("H9")  = "" '// StaticItem1
        .Range("H10") = "" '// StaticItem2
        .Range("H11") = "" '// StaticItem3
        .Range("D11") = "" '// StaticItem4
        .Range("M9")  = "" '// StaticItem5
    End With

    '// Static_List1
    With wsForm.Range("H9:J9").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=STATIC_LIST1"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    '// Static_List2
    With wsForm.Range("H10:J10").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=STATIC_LIST2"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    '// Repeat  
    repeatRow = 20
    jumpedRow = 13 '// Title Row Number

    '// No
    For K = 1 To repeatRow
        wsForm.Cells(jumpedRow + K, 2) = K
    Next

    '// Item
    With wsForm
        .Range("C14:M33").ClearContents
        .Range("R14:S33") = ""
    End With

    '// List1
    With wsForm.Range("C14:C33").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=LIST1"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With    

    '// List2
    With wsForm.Range("D14:D33").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=LIST2"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    '// List3
    With wsForm.Range("G14:G33").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=LIST3"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    '// List4
    With wsForm.Range("H14:H33").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=LIST4"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    '// Ref
    For K = 1 To 9
        wsForm.Cells(jumpedRow + K, 21) = "=VALUE($V$3&""0""&ROW()-" & jumpedRow & ")"
    Next
    For K = 10 To 20
        wsForm.Cells(jumpedRow + K, 21) = "=VALUE($V$3&ROW()-" & jumpedRow & ")"
    Next

    '// Vertical Sum
    For K = 1 To repeatRow
        With wsForm
            .Cells(jumpedRow + K, 14) = "=IFERROR(IF(SUM(E" & jumpedRow + K & ":F" & jumpedRow + K & ")>0,SUM(E" & jumpedRow + K & ":F" & jumpedRow + K & "),""""),"""")" '// Item12
            .Cells(jumpedRow + K, 15) = "=IFERROR(IF(SUM(J" & jumpedRow + K & ":L" & jumpedRow + K & ")>0,SUM(J" & jumpedRow + K & ":L" & jumpedRow + K & "),""""),"""")" '// Item13
            .Cells(jumpedRow + K, 16) = "=IFERROR(IF(I" & jumpedRow + K & "*O" & jumpedRow + K & ">0,I" & jumpedRow + K & "*O" & jumpedRow + K & ",""""),"""")"           '// Item14
            .Cells(jumpedRow + K, 17) = "=IFERROR(IF(M" & jumpedRow + K & "*O" & jumpedRow + K & "/3600>0,M" & jumpedRow + K & "*O" & jumpedRow + K & "/3600,""""),"""")" '// Item15
        End With
    Next

    '// Horizontal Sum
    wsForm.Range("E34") = "=IFERROR(IF(SUM(E14:E33)>0,SUM(E14:E33),""""),"""")"     '// Item3
    wsForm.Range("F34") = "=IFERROR(IF(SUM(F14:F33)>0,SUM(F14:F33),""""),"""")"     '// Item4
    wsForm.Range("I34") = "=IFERROR(IF(SUM(I14:I33)>0,SUM(I14:I33),""""),"""")"     '// Item7
    wsForm.Range("J34") = "=IFERROR(IF(SUM(J14:J33)>0,SUM(J14:J33),""""),"""")"     '// Item8
    wsForm.Range("K34") = "=IFERROR(IF(SUM(K14:K33)>0,SUM(K14:K33),""""),"""")"     '// Item9
    wsForm.Range("L34") = "=IFERROR(IF(SUM(L14:L33)>0,SUM(L14:L33),""""),"""")"     '// Item10
    wsForm.Range("M34") = "=IFERROR(IF(SUM(M14:M33)>0,SUM(M14:M33),""""),"""")"     '// Item11
    wsForm.Range("N34") = "=IFERROR(IF(SUM(N14:N33)>0,SUM(N14:N33),""""),"""")"     '// Item12
    wsForm.Range("O34") = "=IFERROR(IF(SUM(O14:O33)>0,SUM(O14:O33),""""),"""")"     '// Item13
    wsForm.Range("P34") = "=IFERROR(IF(SUM(P14:P33)>0,SUM(P14:P33),""""),"""")"     '// Item14
    wsForm.Range("Q34") = "=IFERROR(IF(SUM(Q14:Q33)>0,SUM(Q14:Q33),""""),"""")"     '// Item15
    wsForm.Range("U34") = "=IFERROR(IF(COUNT(U14:U33)>0,COUNT(U14:U33),""""),"""")" '// Ref

    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub