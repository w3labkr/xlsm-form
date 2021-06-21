Sub module_load()
'
' Load data

    Dim wb As Workbook
    Dim wsForm As Worksheet
    Dim msgAnswer As Integer

    Set wb = ThisWorkbook
    Set wsForm = wb.Sheets("REPORT")
    
    '// Confirm Message
    msgAnswer = MsgBox( "Do you want to load data "& wsForm.Range("V3").Value &" ?", vbYesNo + vbQuestion, "Empty Sheet")
    '// msgAnswer = MsgBox(wsForm.Range("V3").Value & "번 데이터를 불러 오시겠습니까 ?", vbYesNo + vbQuestion, "Empty Sheet")
    
    If msgAnswer = vbNo Then
        Exit Sub
    End If
    
    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False

    '// Reset
    With wsForm
        .Range("C14:M23").ClearContents
    End With

    '// Reinit
    module_load_init

    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    '// Success Message
    MsgBox "DATA has been loaded."
    '// MsgBox "데이터가 로딩 되었습니다."
    
End Sub

Sub module_load_init()
'
' Reinit data

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsForm As Worksheet
    Dim searchCol As Integer
    Dim repeatRow As Integer
    Dim jumpedRow As Integer
    Dim totalItem As Integer

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("DATA")
    Set wsForm = wb.Sheets("REPORT")

    '// Static
    With wsForm
        '// .Range("D9")  = "=VLOOKUP(V3,DATA,1,FALSE)"  '// Lock: Serial
        .Range("V3")  = "=D9"  '// Serial
        .Range("D10") = "=VLOOKUP(V3,DATA,2,FALSE)"  '// Date
        .Range("V4")  = "=ROUNDUP(MONTH(D10)/3,0)&""Q""" '// Quarter
        .Range("V5")  = "=YEAR(D10)"   '// Year
        .Range("V6")  = "=MONTH(D10)"  '// Month
        .Range("V7")  = "=DAY(D10)"    '// Day
        .Range("H9")  = "=IF(VLOOKUP(V3,DATA,7,FALSE)="""","""",VLOOKUP(V3,DATA,7,FALSE))"  '// StaticItem1
        .Range("H10") = "=IF(VLOOKUP(V3,DATA,8,FALSE)="""","""",VLOOKUP(V3,DATA,8,FALSE))" '// StaticItem2
        .Range("H11") = "=IF(VLOOKUP(V3,DATA,9,FALSE)="""","""",VLOOKUP(V3,DATA,9,FALSE))" '// StaticItem3
        .Range("D11") = "=IF(VLOOKUP(V3,DATA,10,FALSE)="""","""",VLOOKUP(V3,DATA,10,FALSE))" '// StaticItem4
        .Range("M9")  = "=IF(VLOOKUP(V3,DATA,11,FALSE)="""","""",VLOOKUP(V3,DATA,11,FALSE))" '// StaticItem5
    End With
    
    '// Repeat
    repeatRow = 20
    jumpedRow = 13 '// Title Row Number
    totalItem = 18 '// No(1) + Item(16) + Ref(1)

    '// No
    For K = 1 To repeatRow
        wsForm.Cells(jumpedRow + K, 2) = K
    Next

    '// Item
    For K = 1 To repeatRow
        searchCol = K * totalItem - 6
        With wsForm
            .Cells(jumpedRow + K, 3)  = "=IF(VLOOKUP(V3,DATA," & searchCol + 2 & ",FALSE)="""","""",VLOOKUP(V3,DATA," & searchCol + 2 & ",FALSE))"   '// Item1
            .Cells(jumpedRow + K, 4)  = "=IF(VLOOKUP(V3,DATA," & searchCol + 3 & ",FALSE)="""","""",VLOOKUP(V3,DATA," & searchCol + 3 & ",FALSE))"   '// Item2
            .Cells(jumpedRow + K, 5)  = "=IF(VLOOKUP(V3,DATA," & searchCol + 4 & ",FALSE)="""","""",VLOOKUP(V3,DATA," & searchCol + 4 & ",FALSE))"   '// Item3
            .Cells(jumpedRow + K, 6)  = "=IF(VLOOKUP(V3,DATA," & searchCol + 5 & ",FALSE)="""","""",VLOOKUP(V3,DATA," & searchCol + 5 & ",FALSE))"   '// Item4
            .Cells(jumpedRow + K, 7)  = "=IF(VLOOKUP(V3,DATA," & searchCol + 6 & ",FALSE)="""","""",VLOOKUP(V3,DATA," & searchCol + 6 & ",FALSE))"   '// Item5
            .Cells(jumpedRow + K, 8)  = "=IF(VLOOKUP(V3,DATA," & searchCol + 7 & ",FALSE)="""","""",VLOOKUP(V3,DATA," & searchCol + 7 & ",FALSE))"   '// Item6
            .Cells(jumpedRow + K, 9)  = "=IF(VLOOKUP(V3,DATA," & searchCol + 8 & ",FALSE)="""","""",VLOOKUP(V3,DATA," & searchCol + 8 & ",FALSE))"   '// Item7
            .Cells(jumpedRow + K, 10) = "=IF(VLOOKUP(V3,DATA," & searchCol + 9 & ",FALSE)="""","""",VLOOKUP(V3,DATA," & searchCol + 9 & ",FALSE))"   '// Item8
            .Cells(jumpedRow + K, 11) = "=IF(VLOOKUP(V3,DATA," & searchCol + 10 & ",FALSE)="""","""",VLOOKUP(V3,DATA," & searchCol + 10 & ",FALSE))" '// Item9
            .Cells(jumpedRow + K, 12) = "=IF(VLOOKUP(V3,DATA," & searchCol + 11 & ",FALSE)="""","""",VLOOKUP(V3,DATA," & searchCol + 11 & ",FALSE))" '// Item10
            .Cells(jumpedRow + K, 13) = "=IF(VLOOKUP(V3,DATA," & searchCol + 12 & ",FALSE)="""","""",VLOOKUP(V3,DATA," & searchCol + 12 & ",FALSE))" '// Item11
            .Cells(jumpedRow + K, 18) = "=IF(VLOOKUP(V3,DATA," & searchCol + 17 & ",FALSE)="""","""",VLOOKUP(V3,DATA," & searchCol + 17 & ",FALSE))" '// Item16
        End With
    Next

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

End Sub
