Sub module_save()
'
' Save data

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsForm As Worksheet
    Dim getSerial As Integer
    Dim repeatRow As Integer
    Dim jumpedRow As Integer
    Dim searchRow As Integer
    Dim searchCol As Integer
    Dim totalItem As Integer
    Dim msgAnswer As Integer

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("DATA")
    Set wsForm = wb.Sheets("REPORT")

    '// Confirm Message
    msgAnswer = MsgBox("Are you sure you want to save this data?", vbYesNo + vbQuestion, "Empty Sheet")
    '// msgAnswer = MsgBox("데이터를 저장 하시겠습니까?", vbYesNo + vbQuestion, "Empty Sheet")
    
    If msgAnswer = vbNo Then
        Exit Sub
    End If

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False
    
    '// Static
    getSerial = wsForm.Range("V3")
    searchRow = getSerial + 1

    With wsData
        .Cells(searchRow, 1)  = getSerial           '// Serial
        .Cells(searchRow, 2)  = wsForm.Range("D10") '// Date
        .Cells(searchRow, 3)  = wsForm.Range("V4")  '// Quarter
        .Cells(searchRow, 4)  = wsForm.Range("V5")  '// Year
        .Cells(searchRow, 5)  = wsForm.Range("V6")  '// Month
        .Cells(searchRow, 6)  = wsForm.Range("V7")  '// Day
        .Cells(searchRow, 7)  = wsForm.Range("H9")  '// StaticItem1
        .Cells(searchRow, 8)  = wsForm.Range("H10") '// StaticItem2
        .Cells(searchRow, 9)  = wsForm.Range("H11") '// StaticItem4
        .Cells(searchRow, 10) = wsForm.Range("D11") '// StaticItem3
        .Cells(searchRow, 11) = wsForm.Range("M9")  '// StaticItem5
    End With

    '// Repeat
    repeatRow = 20
    jumpedRow = 13 '// Title Row Number
    totalItem = 18 '// No(1) + Item(16) + Ref(1)

    For K = 1 To repeatRow
        searchCol = K * totalItem - 6
        With wsData
            .Cells(searchRow, searchCol + 0)  = wsForm.Cells(jumpedRow + K, 21) '// Ref
            .Cells(searchRow, searchCol + 1)  = wsForm.Cells(jumpedRow + K, 2)  '// No
            .Cells(searchRow, searchCol + 2)  = wsForm.Cells(jumpedRow + K, 3)  '// Item1
            .Cells(searchRow, searchCol + 3)  = wsForm.Cells(jumpedRow + K, 4)  '// Item2
            .Cells(searchRow, searchCol + 4)  = wsForm.Cells(jumpedRow + K, 5)  '// Item3
            .Cells(searchRow, searchCol + 5)  = wsForm.Cells(jumpedRow + K, 6)  '// Item4
            .Cells(searchRow, searchCol + 6)  = wsForm.Cells(jumpedRow + K, 7)  '// Item5
            .Cells(searchRow, searchCol + 7)  = wsForm.Cells(jumpedRow + K, 8)  '// Item6
            .Cells(searchRow, searchCol + 8)  = wsForm.Cells(jumpedRow + K, 9)  '// Item7
            .Cells(searchRow, searchCol + 9)  = wsForm.Cells(jumpedRow + K, 10) '// Item8
            .Cells(searchRow, searchCol + 10) = wsForm.Cells(jumpedRow + K, 11) '// Item9
            .Cells(searchRow, searchCol + 11) = wsForm.Cells(jumpedRow + K, 12) '// Item10
            .Cells(searchRow, searchCol + 12) = wsForm.Cells(jumpedRow + K, 13) '// Item11
            .Cells(searchRow, searchCol + 13) = wsForm.Cells(jumpedRow + K, 14) '// Item12
            .Cells(searchRow, searchCol + 14) = wsForm.Cells(jumpedRow + K, 15) '// Item13
            .Cells(searchRow, searchCol + 15) = wsForm.Cells(jumpedRow + K, 16) '// Item14
            .Cells(searchRow, searchCol + 16) = wsForm.Cells(jumpedRow + K, 17) '// Item15
            .Cells(searchRow, searchCol + 17) = wsForm.Cells(jumpedRow + K, 18) '// Item16
        End With
    Next
    
    '// Save pivotdata
    module_pivotdata_save

    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    '// Success Message
    MsgBox "Data has been saved."
    '// MsgBox "데이터가 저장 되었습니다.

    '// Load a new document macro
    module_new

End Sub

Sub module_pivotdata_save()
'
' Save pivotdata

    Dim wb As Workbook
    Dim wsForm As Worksheet
    Dim wsPivotData As Worksheet
    Dim getSerial As Integer
    Dim getRefCol As Integer
    Dim searchRef As Integer
    Dim searchRow As Integer
    Dim repeatRow As Integer
    Dim jumpedRow As Integer

    Set wb = ThisWorkbook
    Set wsForm = wb.Sheets("REPORT")
    Set wsPivotData = wb.Sheets("PIVOTDATA")

    '// Static
    getSerial = wsForm.Range("V3")
    getRefCol = 12 '// get Ref Column in PIVOTDATA worksheet

    '// Repeat
    repeatRow = 20
    jumpedRow = 13 '// Title Row Number

    On Error GoTo newLine:  
    ' Insert code that might generate an error here
        '// =MATCH(VLOOKUP(V3,PIVOTDATA,12,FALSE()),PIVOTDATA_REF,0)
        '// =MATCH(VLOOKUP(V3,피벗데이터,12,FALSE()),PIVOTDATA_REF,0)
        searchRef = Application.WorksheetFunction.VLookup(getSerial, _
                    wsPivotData.Range("PIVOTDATA"), getRefCol, False)
        searchRow = Application.Match(searchRef, _
                    wsPivotData.Range("PIVOTDATA_REF"), 0) - 1  
        For K = 1 To repeatRow
            With wsPivotData
                .Cells(searchRow + K, 1)  = getSerial               '// Serial
                .Cells(searchRow + K, 2)  = wsForm.Range("D10")     '// Date
                .Cells(searchRow + K, 3)  = wsForm.Range("V4")      '// Quarter
                .Cells(searchRow + K, 4)  = wsForm.Range("V5")      '// Year
                .Cells(searchRow + K, 5)  = wsForm.Range("V6")      '// Month
                .Cells(searchRow + K, 6)  = wsForm.Range("V7")      '// Day
                .Cells(searchRow + K, 7)  = wsForm.Range("H9")      '// StaticItem1
                .Cells(searchRow + K, 8)  = wsForm.Range("H10")     '// StaticItem2
                .Cells(searchRow + K, 9)  = wsForm.Range("H11")     '// StaticItem3
                .Cells(searchRow + K, 10) = wsForm.Range("D11")     '// StaticItem4
                .Cells(searchRow + K, 11) = wsForm.Range("M9")      '// StaticItem5
                .Cells(searchRow + K, 12) = wsForm.Cells(jumpedRow + K, 21) '// Ref
                .Cells(searchRow + K, 13) = wsForm.Cells(jumpedRow + K, 2)  '// No
                .Cells(searchRow + K, 14) = wsForm.Cells(jumpedRow + K, 3)  '// Item1
                .Cells(searchRow + K, 15) = wsForm.Cells(jumpedRow + K, 4)  '// Item2
                .Cells(searchRow + K, 16) = wsForm.Cells(jumpedRow + K, 5)  '// Item3
                .Cells(searchRow + K, 17) = wsForm.Cells(jumpedRow + K, 6)  '// Item4
                .Cells(searchRow + K, 18) = wsForm.Cells(jumpedRow + K, 7)  '// Item5
                .Cells(searchRow + K, 19) = wsForm.Cells(jumpedRow + K, 8)  '// Item6
                .Cells(searchRow + K, 20) = wsForm.Cells(jumpedRow + K, 9)  '// Item7
                .Cells(searchRow + K, 21) = wsForm.Cells(jumpedRow + K, 10) '// Item8
                .Cells(searchRow + K, 22) = wsForm.Cells(jumpedRow + K, 11) '// Item9
                .Cells(searchRow + K, 23) = wsForm.Cells(jumpedRow + K, 12) '// Item10
                .Cells(searchRow + K, 24) = wsForm.Cells(jumpedRow + K, 13) '// Item11
                .Cells(searchRow + K, 25) = wsForm.Cells(jumpedRow + K, 14) '// Item12
                .Cells(searchRow + K, 26) = wsForm.Cells(jumpedRow + K, 15) '// Item13
                .Cells(searchRow + K, 27) = wsForm.Cells(jumpedRow + K, 16) '// Item14
                .Cells(searchRow + K, 28) = wsForm.Cells(jumpedRow + K, 17) '// Item15
                .Cells(searchRow + K, 29) = wsForm.Cells(jumpedRow + K, 18) '// Item16
            End With
        Next
    Exit Sub  
    newLine:  
    ' Insert code to handle the error here
        searchRow = wsPivotData.Range("A1").CurrentRegion.Rows.Count        
        For K = 1 To repeatRow
            With wsPivotData
                .Cells(searchRow + K, 1)  = getSerial               '// Serial
                .Cells(searchRow + K, 2)  = wsForm.Range("D10")     '// Date
                .Cells(searchRow + K, 3)  = wsForm.Range("V4")      '// Quarter
                .Cells(searchRow + K, 4)  = wsForm.Range("V5")      '// Year
                .Cells(searchRow + K, 5)  = wsForm.Range("V6")      '// Month
                .Cells(searchRow + K, 6)  = wsForm.Range("V7")      '// Day
                .Cells(searchRow + K, 7)  = wsForm.Range("H9")      '// StaticItem1
                .Cells(searchRow + K, 8)  = wsForm.Range("H10")     '// StaticItem2
                .Cells(searchRow + K, 9)  = wsForm.Range("H11")     '// StaticItem3
                .Cells(searchRow + K, 10) = wsForm.Range("D11")     '// StaticItem4
                .Cells(searchRow + K, 11) = wsForm.Range("M9")      '// StaticItem5
                .Cells(searchRow + K, 12) = wsForm.Cells(jumpedRow + K, 21) '// Ref
                .Cells(searchRow + K, 13) = wsForm.Cells(jumpedRow + K, 2)  '// No
                .Cells(searchRow + K, 14) = wsForm.Cells(jumpedRow + K, 3)  '// Item1
                .Cells(searchRow + K, 15) = wsForm.Cells(jumpedRow + K, 4)  '// Item2
                .Cells(searchRow + K, 16) = wsForm.Cells(jumpedRow + K, 5)  '// Item3
                .Cells(searchRow + K, 17) = wsForm.Cells(jumpedRow + K, 6)  '// Item4
                .Cells(searchRow + K, 18) = wsForm.Cells(jumpedRow + K, 7)  '// Item5
                .Cells(searchRow + K, 19) = wsForm.Cells(jumpedRow + K, 8)  '// Item6
                .Cells(searchRow + K, 20) = wsForm.Cells(jumpedRow + K, 9)  '// Item7
                .Cells(searchRow + K, 21) = wsForm.Cells(jumpedRow + K, 10) '// Item8
                .Cells(searchRow + K, 22) = wsForm.Cells(jumpedRow + K, 11) '// Item9
                .Cells(searchRow + K, 23) = wsForm.Cells(jumpedRow + K, 12) '// Item10
                .Cells(searchRow + K, 24) = wsForm.Cells(jumpedRow + K, 13) '// Item11
                .Cells(searchRow + K, 25) = wsForm.Cells(jumpedRow + K, 14) '// Item12
                .Cells(searchRow + K, 26) = wsForm.Cells(jumpedRow + K, 15) '// Item13
                .Cells(searchRow + K, 27) = wsForm.Cells(jumpedRow + K, 16) '// Item14
                .Cells(searchRow + K, 28) = wsForm.Cells(jumpedRow + K, 17) '// Item15
                .Cells(searchRow + K, 29) = wsForm.Cells(jumpedRow + K, 18) '// Item16
            End With
        Next
    Exit Sub  

End Sub