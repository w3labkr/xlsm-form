Sub module_delete()
'
' Delete data

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsForm As Worksheet
    Dim getSerial As Integer
    Dim repeatRow As Integer
    Dim searchRow As Integer
    Dim searchCol As Integer
    Dim totalItem As Integer
    Dim msgAnswer As Integer

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("DATA")
    Set wsForm = wb.Sheets("REPORT")

    '// Confirm Message
    msgAnswer = MsgBox("Are you sure you want to delete this data?", vbYesNo + vbQuestion, "Empty Sheet")
    '// msgAnswer = MsgBox("데이터를 삭제 하시겠습니까?", vbYesNo + vbQuestion, "Empty Sheet")
    
    If msgAnswer = vbNo Then
        Exit Sub
    End If

    '// Prevents screen refreshing.
    Application.Calculation = xlCalculateManual
    Application.ScreenUpdating = False
    
    '// Delete DATA
    getSerial = wsForm.Range("V3")
    searchRow = getSerial + 1

    With wsData
        '// .Cells(searchRow, 1) = getSerial            '// Lock: Serial
        '// .Cells(searchRow, 2) = wsForm.Range("D10")  '// Lock: Date
        '// .Cells(searchRow, 3) = wsForm.Range("V4")   '// Lock: Quarter
        '// .Cells(searchRow, 4) = wsForm.Range("V5")   '// Lock: Year
        '// .Cells(searchRow, 5) = wsForm.Range("V6")   '// Lock: Month
        '// .Cells(searchRow, 6) = wsForm.Range("V7")   '// Lock: Day
        .Cells(searchRow, 7).ClearContents  '// StaticItem1
        .Cells(searchRow, 8).ClearContents  '// StaticItem2
        .Cells(searchRow, 9).ClearContents  '// StaticItem3
        .Cells(searchRow, 10).ClearContents '// StaticItem4
        .Cells(searchRow, 11).ClearContents '// StaticItem5
    End With

    '// Repeat
    repeatRow = 20
    totalItem = 18 '// No(1) + Item(16) + Ref(1)

    For K = 1 To repeatRow
        searchCol = K * totalItem - 6
        With wsData
            '// .Cells(searchRow, searchCol + 0).ClearContents '// Lock: Ref
            '// .Cells(searchRow, searchCol + 1).ClearContents '// Lock: No
            .Cells(searchRow, searchCol + 2).ClearContents     '// Item1
            .Cells(searchRow, searchCol + 3).ClearContents     '// Item2
            .Cells(searchRow, searchCol + 4).ClearContents     '// Item3
            .Cells(searchRow, searchCol + 5).ClearContents     '// Item4
            .Cells(searchRow, searchCol + 6).ClearContents     '// Item5
            .Cells(searchRow, searchCol + 7).ClearContents     '// Item6
            .Cells(searchRow, searchCol + 8).ClearContents     '// Item7
            .Cells(searchRow, searchCol + 9).ClearContents     '// Item8
            .Cells(searchRow, searchCol + 10).ClearContents    '// Item9
            .Cells(searchRow, searchCol + 11).ClearContents    '// Item10
            .Cells(searchRow, searchCol + 12).ClearContents    '// Item11
            .Cells(searchRow, searchCol + 13).ClearContents    '// Item12
            .Cells(searchRow, searchCol + 14).ClearContents    '// Item13
            .Cells(searchRow, searchCol + 15).ClearContents    '// Item14
            .Cells(searchRow, searchCol + 16).ClearContents    '// Item15
            .Cells(searchRow, searchCol + 17).ClearContents    '// Item16
        End With
    Next

    '// Delete PIVOTDATA
    module_pivotdata_delete

    '// Reinit
    module_load_init

    '// Enables screen refreshing.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    '// Success Message
    MsgBox "Data has been deleted."
    '// MsgBox "데이터가 삭제 되었습니다."

    '// Load a new document macro
    module_new

End Sub

Sub module_pivotdata_delete()
'
' Delete pivotdata

    Dim wb As Workbook
    Dim wsForm As Worksheet
    Dim wsPivotData As Worksheet
    Dim getSerial As Integer
    Dim getRefCol As Integer
    Dim searchRef As Integer
    Dim searchRow As Integer
    Dim repeatRow As Integer

    Set wb = ThisWorkbook
    Set wsForm = wb.Sheets("REPORT")
    Set wsPivotData = wb.Sheets("PIVOTDATA")

    '// Static
    getSerial = wsForm.Range("V3")
    getRefCol = 12 '// get Ref Column in PIVOTDATA worksheet

    '// Repeat
    repeatRow = 20

    On Error GoTo newLine:  
    ' Insert code that might generate an error here
        '// =MATCH(VLOOKUP(V2,PIVOTDATA,7,FALSE()),PIVOTDATA_REF,0)
        searchRef = Application.WorksheetFunction.VLookup(getSerial, _
                    wsPivotData.Range("PIVOTDATA"), getRefCol, False)
        searchRow = Application.Match(searchRef, _
                    wsPivotData.Range("PIVOTDATA_REF"), 0) - 1  
        For K = 1 To repeatRow
            With wsPivotData
                '// .Cells(searchRow + K, 1).ClearContents  '// Lock: Serial
                '// .Cells(searchRow + K, 2).ClearContents  '// Lock: Date
                '// .Cells(searchRow + K, 3).ClearContents  '// Lock: Quarter
                '// .Cells(searchRow + K, 4).ClearContents  '// Lock: Year
                '// .Cells(searchRow + K, 5).ClearContents  '// Lock: Month
                '// .Cells(searchRow + K, 6).ClearContents  '// Lock: Day
                .Cells(searchRow + K, 7).ClearContents  '// StaticItem1
                .Cells(searchRow + K, 8).ClearContents  '// StaticItem2
                .Cells(searchRow + K, 9).ClearContents  '// StaticItem3
                .Cells(searchRow + K, 10).ClearContents '// StaticItem4
                .Cells(searchRow + K, 11).ClearContents '// StaticItem5
                '// .Cells(searchRow + K, 12).ClearContents '// Lock: Ref
                '// .Cells(searchRow + K, 13).ClearContents '// Lock: No
                .Cells(searchRow + K, 14).ClearContents '// Item1
                .Cells(searchRow + K, 15).ClearContents '// Item2
                .Cells(searchRow + K, 16).ClearContents '// Item3
                .Cells(searchRow + K, 17).ClearContents '// Item4
                .Cells(searchRow + K, 18).ClearContents '// Item5
                .Cells(searchRow + K, 19).ClearContents '// Item6
                .Cells(searchRow + K, 20).ClearContents '// Item7
                .Cells(searchRow + K, 21).ClearContents '// Item8
                .Cells(searchRow + K, 22).ClearContents '// Item9
                .Cells(searchRow + K, 23).ClearContents '// Item10
                .Cells(searchRow + K, 24).ClearContents '// Item11
                .Cells(searchRow + K, 25).ClearContents '// Item12
                .Cells(searchRow + K, 26).ClearContents '// Item13
                .Cells(searchRow + K, 27).ClearContents '// Item14
                .Cells(searchRow + K, 28).ClearContents '// Item15
                .Cells(searchRow + K, 29).ClearContents '// Item16
            End With
        Next
    Exit Sub  
    newLine:  
    ' Insert code to handle the error here
        searchRow = wsPivotData.Range("A1").CurrentRegion.Rows.Count        
        For K = 1 To repeatRow
            With wsPivotData
                '// .Cells(searchRow + K, 1).ClearContents  '// Lock: Serial
                '// .Cells(searchRow + K, 2).ClearContents  '// Lock: Date
                '// .Cells(searchRow + K, 3).ClearContents  '// Lock: Quarter
                '// .Cells(searchRow + K, 4).ClearContents  '// Lock: Year
                '// .Cells(searchRow + K, 5).ClearContents  '// Lock: Month
                '// .Cells(searchRow + K, 6).ClearContents  '// Lock: Day
                .Cells(searchRow + K, 7).ClearContents  '// StaticItem1
                .Cells(searchRow + K, 8).ClearContents  '// StaticItem2
                .Cells(searchRow + K, 9).ClearContents  '// StaticItem3
                .Cells(searchRow + K, 10).ClearContents '// StaticItem4
                .Cells(searchRow + K, 11).ClearContents '// StaticItem5
                '// .Cells(searchRow + K, 12).ClearContents '// Lock: Ref
                '// .Cells(searchRow + K, 13).ClearContents '// Lock: No
                .Cells(searchRow + K, 14).ClearContents '// Item1
                .Cells(searchRow + K, 15).ClearContents '// Item2
                .Cells(searchRow + K, 16).ClearContents '// Item3
                .Cells(searchRow + K, 17).ClearContents '// Item4
                .Cells(searchRow + K, 18).ClearContents '// Item5
                .Cells(searchRow + K, 19).ClearContents '// Item6
                .Cells(searchRow + K, 20).ClearContents '// Item7
                .Cells(searchRow + K, 21).ClearContents '// Item8
                .Cells(searchRow + K, 22).ClearContents '// Item9
                .Cells(searchRow + K, 23).ClearContents '// Item10
                .Cells(searchRow + K, 24).ClearContents '// Item11
                .Cells(searchRow + K, 25).ClearContents '// Item12
                .Cells(searchRow + K, 26).ClearContents '// Item13
                .Cells(searchRow + K, 27).ClearContents '// Item14
                .Cells(searchRow + K, 28).ClearContents '// Item15
                .Cells(searchRow + K, 29).ClearContents '// Item16
            End With
        Next
    Exit Sub  

End Sub
