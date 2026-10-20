Option Explicit

' 시트 이름
Public Const VIEW_NAME As String = "데이터베이스"
Public Const CONFIG_NAME As String = "제품구성"
Public Const PRODUCT_NAME As String = "제품목록"
Public Const USAGE_NAME As String = "제품사용량"
Public Const LOG_NAME As String = "이월이력"
Public Const ALL_PRODUCT As String = "전체"

' 데이터베이스
Public Const FIRST_VIEW As Long = 4
Private Const MATERIAL_NAME_COL As Long = 1
Public Const CODE_COL As Long = 2
Public Const STOCK_COL As Long = 3
Public Const INCOMING_COL As Long = 4
Public Const USE_COL As Long = 5
Public Const CLOSING_COL As Long = 6
Private Const UNIT_COL As Long = 7
Private Const REPLENISHMENT_STEP_COL As Long = 8
Public Const RECEIPT_COL As Long = 9
Public Const PRINT_LAST_COL As Long = RECEIPT_COL
Public Const TARE_COL As Long = 10
Public Const MEMBER_COL As Long = 11
Public Const SEARCH_COL As Long = 12
Public Const KEY_COL As Long = 13
Private Const VIEW_DATA_RANGE_PREFIX As String = "A" & FIRST_VIEW & ":M"
Private Const VIEW_STOCK_RANGE_PREFIX As String = "C" & FIRST_VIEW & ":C"
Public Const PRODUCT_CELL As String = "P1"
Public Const PRODUCT_RANGE As String = "P1:Q1"
Public Const MODE_CELL As String = "P2"
Public Const MODE_RANGE As String = "P2:Q2"
Public Const QUERY_CELL As String = "P3"
Public Const QUERY_RANGE As String = "P3:Q3"
Public Const PRINT_ALL_CELL As String = "N1"
Private Const CAPACITY_CELL As String = "N2"

' 설정
Private Const CONFIG_PRODUCT_COL As Long = 1
Private Const CONFIG_CODE_COL As Long = 2
Private Const CONFIG_MATERIAL_NAME_COL As Long = 3
Private Const CONFIG_FIRST_CAPACITY_COL As Long = 4
Private Const CONFIG_LAST_CAPACITY_COL As Long = 7
Private Const CONFIG_DATA_RANGE_PREFIX As String = "A1:G"
Private Const CONFIG_FORMULA As String = "=IF(LEN(TRIM(RC2))=0,"""",IFERROR(INDEX(tblInventoryView[품명],MATCH(RC2,tblInventoryView[자재코드],0)),""자재코드 확인""))"

' 제품
Private Const PRODUCT_LIST_NAME_COL As Long = 1
Private Const PRODUCT_DATA_RANGE_PREFIX As String = "A1:B"
Private Const USAGE_KEY_COL As Long = 1
Private Const USAGE_PRODUCT_COL As Long = 2
Private Const USAGE_CODE_COL As Long = 3
Private Const USAGE_QUANTITY_COL As Long = 4
Private Const USAGE_SELECTION_COL As Long = 6
Private Const USAGE_DATA_RANGE_PREFIX As String = "A1:D"
Private Const USAGE_QUANTITY_RANGE_PREFIX As String = "D2:D"
Private Const USAGE_SELECTION_RANGE_PREFIX As String = "F1:F"

' 기록
Private Const LOG_TIMESTAMP_COL As Long = 1
Private Const LOG_CODE_COL As Long = 5

Public Busy As Boolean
Public LastError As String
Private activeProduct As String
Private printAllItems As Boolean
Private savedMaterials As Object
Private savedDrafts As Object
Private savedConfig As Variant
Private savedConfigValues As Variant
Private savedProducts As Variant
Private savedView As Variant
Private savedViewIndex As Object
Private lookupActive As Boolean
Private viewRows As Variant, configRows As Variant, useRows As Variant, productRows As Variant
Private viewLast As Long
Private materialIndex As Object, productIndex As Object, configIndex As Object
Private unionIndex As Object, usageIndex As Object, usageTotals As Object, memberCounts As Object

Private Function NewTextDictionary() As Object
    Set NewTextDictionary = CreateObject("Scripting.Dictionary")
    NewTextDictionary.CompareMode = vbTextCompare
End Function

Private Sub EnsureMacroAccess()
    Dim ws As Worksheet, sheetName As Variant, selectionMode As Long
    For Each sheetName In Array(VIEW_NAME, CONFIG_NAME, PRODUCT_NAME, USAGE_NAME, LOG_NAME)
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        With ws
            If .ProtectContents And Not .ProtectionMode Then
                selectionMode = .EnableSelection
                .Protect UserInterfaceOnly:=True, Contents:=True, DrawingObjects:=.ProtectDrawingObjects, Scenarios:=.ProtectScenarios, _
                    AllowFormattingCells:=.Protection.AllowFormattingCells, AllowFormattingColumns:=.Protection.AllowFormattingColumns, _
                    AllowFormattingRows:=.Protection.AllowFormattingRows, AllowInsertingColumns:=.Protection.AllowInsertingColumns, _
                    AllowInsertingRows:=.Protection.AllowInsertingRows, AllowInsertingHyperlinks:=.Protection.AllowInsertingHyperlinks, _
                    AllowDeletingColumns:=.Protection.AllowDeletingColumns, AllowDeletingRows:=.Protection.AllowDeletingRows, _
                    AllowSorting:=.Protection.AllowSorting, AllowFiltering:=.Protection.AllowFiltering, _
                    AllowUsingPivotTables:=.Protection.AllowUsingPivotTables
                .EnableSelection = selectionMode
            End If
        End With
    Next sheetName
End Sub

Private Sub EndLookup()
    lookupActive = False
End Sub

Private Function ReadColumnArray(ByVal cells As Range, Optional ByVal preserveFormulas As Boolean = False) As Variant
    Dim values(1 To 1, 1 To 1) As Variant
    Dim state As Variant
    If preserveFormulas Then state = cells.formula Else state = cells.Value2
    If cells.CountLarge = 1 Then
        values(1, 1) = state
        ReadColumnArray = values
    Else
        ReadColumnArray = state
    End If
End Function

Private Sub BeginLookup()
    Dim r As Long, code As String, product As String, key As String, ws As Worksheet
    EndLookup
    viewLast = LastViewRow()
    If viewLast < FIRST_VIEW Then viewLast = FIRST_VIEW
    viewRows = ThisWorkbook.Worksheets(VIEW_NAME).Range(VIEW_DATA_RANGE_PREFIX & viewLast).Value2
    With ThisWorkbook.Worksheets(CONFIG_NAME)
        configRows = .Range(CONFIG_DATA_RANGE_PREFIX & LastConfigRow()).Value2
    End With
    Set ws = ThisWorkbook.Worksheets(USAGE_NAME)
    useRows = ws.Range(USAGE_DATA_RANGE_PREFIX & lastUsageRow()).Value2
    Set ws = ThisWorkbook.Worksheets(PRODUCT_NAME)
    productRows = ws.Range(PRODUCT_DATA_RANGE_PREFIX & LastRow(ws)).Value2
    Set materialIndex = NewTextDictionary()
    Set productIndex = NewTextDictionary()
    Set configIndex = NewTextDictionary()
    Set unionIndex = NewTextDictionary()
    Set memberCounts = NewTextDictionary()
    Set usageIndex = NewTextDictionary()
    Set usageTotals = NewTextDictionary()
    For r = 1 To UBound(viewRows, 1)
        code = Trim$(CStr(viewRows(r, CODE_COL)))
        If Len(code) > 0 And Len(Trim$(CStr(viewRows(r, MATERIAL_NAME_COL)))) > 0 Then materialIndex(code) = r + FIRST_VIEW - 1
    Next r
    For r = 2 To UBound(productRows, 1)
        productIndex(Trim$(CStr(productRows(r, PRODUCT_LIST_NAME_COL)))) = r
    Next r
    For r = 2 To UBound(configRows, 1)
        product = Trim$(CStr(configRows(r, CONFIG_PRODUCT_COL)))
        code = Trim$(CStr(configRows(r, CONFIG_CODE_COL)))
        If Len(product) > 0 And Len(code) > 0 Then
            configIndex(PairKey(product, code)) = r
            unionIndex(code) = r
            memberCounts(product) = memberCounts(product) + 1
        End If
    Next r
    For r = 2 To UBound(useRows, 1)
        usageIndex(Trim$(CStr(useRows(r, USAGE_KEY_COL)))) = r
        code = CStr(useRows(r, USAGE_CODE_COL))
        If Not IsError(useRows(r, USAGE_QUANTITY_COL)) Then
            If IsNumeric(useRows(r, USAGE_QUANTITY_COL)) Then usageTotals(code) = usageTotals(code) + CDbl(useRows(r, USAGE_QUANTITY_COL))
        End If
    Next r
    lookupActive = True
End Sub


Private Sub SetViewFormulas(ByVal first As Long, ByVal last As Long, Optional ByVal searchOnly As Boolean = False)
    Dim identities As Variant, r As Long, emptyFirst As Long, emptyRow As Boolean, col As Variant, cleared As Variant
    If last < first Then Exit Sub
    With ThisWorkbook.Worksheets(VIEW_NAME)
        If searchOnly Then
            cleared = Array(SEARCH_COL)
        Else
            .Range("D" & first & ":D" & last).FormulaR1C1 = "=IF(OR(LEN(TRIM(RC1))=0,LEN(TRIM(RC2))=0),"""",IF(ROUND(RC5-RC3,6)<=0,0,IF(RC8>0,ROUND(ROUNDUP(ROUND(RC5-RC3,6)/RC8,0)*RC8,6),""입고량 확인"")))"
            .Range("F" & first & ":F" & last).FormulaR1C1 = "=IF(OR(LEN(TRIM(RC1))=0,LEN(TRIM(RC2))=0),"""",IFERROR(ROUND(RC3+RC4-RC5,6),""입고량 확인""))"
            .Range("I" & first & ":I" & last).FormulaR1C1 = "=IF(OR(LEN(TRIM(RC1))=0,LEN(TRIM(RC2))=0),"""",IFERROR(IF(RC4=0,0,IF(RC7>0,MAX(1,ROUNDUP(ROUND(RC4/RC7,10),0)),""단위 확인"")),""입고량 확인""))"
            cleared = Array(INCOMING_COL, CLOSING_COL, RECEIPT_COL, SEARCH_COL)
        End If
        .Range("L" & first & ":L" & last).FormulaR1C1 = "=RC1&"" ""&RC2"
        identities = .Range("A" & first & ":B" & last).Value2
        For r = 1 To UBound(identities, 1) + 1
            emptyRow = False
            If r <= UBound(identities, 1) Then _
                emptyRow = Len(Trim$(CStr(identities(r, 1)))) = 0 And Len(Trim$(CStr(identities(r, 2)))) = 0
            If emptyRow Then
                If emptyFirst = 0 Then emptyFirst = first + r - 1
            ElseIf emptyFirst > 0 Then
                For Each col In cleared
                    .Range(.cells(emptyFirst, CLng(col)), .cells(first + r - 2, CLng(col))).ClearContents
                Next col
                emptyFirst = 0
            End If
        Next r
    End With
End Sub

Private Function LastRow(ByVal ws As Worksheet, Optional ByVal col As Long = 1) As Long
    LastRow = ws.cells(ws.Rows.count, col).End(xlUp).Row
    If LastRow < 1 Then LastRow = 1
End Function

Private Function PairKey(ByVal product As String, ByVal code As String) As String
    PairKey = Len(product) & ":" & product & ":" & code
End Function

Private Function LastConfigRow() As Long
    Dim ws As Worksheet, col As Long, upper As Long, r As Long, inputs As Variant
    Set ws = ThisWorkbook.Worksheets(CONFIG_NAME)
    LastConfigRow = WorksheetFunction.Max(LastRow(ws, CONFIG_PRODUCT_COL), LastRow(ws, CONFIG_CODE_COL))
    For col = CONFIG_FIRST_CAPACITY_COL To CONFIG_LAST_CAPACITY_COL
        LastConfigRow = WorksheetFunction.Max(LastConfigRow, LastRow(ws, col))
    Next col
    If ws.FilterMode Then
        upper = WorksheetFunction.Min(ws.AutoFilter.Range.Row + ws.AutoFilter.Range.Rows.count - 1, _
                                      ws.UsedRange.Row + ws.UsedRange.Rows.count - 1)
        inputs = ws.Range(CONFIG_DATA_RANGE_PREFIX & upper).Value2
        For r = upper To LastConfigRow + 1 Step -1
            For col = CONFIG_PRODUCT_COL To CONFIG_LAST_CAPACITY_COL
                If col <> CONFIG_MATERIAL_NAME_COL Then
                    If IsError(inputs(r, col)) Then
                        LastConfigRow = r
                        Exit Function
                    ElseIf Len(CStr(inputs(r, col))) > 0 Then
                        LastConfigRow = r
                        Exit Function
                    End If
                End If
            Next col
        Next r
    End If
End Function

Private Function FindRow(ByVal ws As Worksheet, ByVal text As String, Optional ByVal col As Long = 1) As Long
    Dim r As Long
    For r = 2 To LastRow(ws, col)
        If StrComp(Trim$(CStr(ws.cells(r, col).Value2)), Trim$(text), vbTextCompare) = 0 Then
            FindRow = r
            Exit Function
        End If
    Next r
End Function

Private Function LastViewRow() As Long
    Dim ws As Worksheet, col As Variant, r As Long
    If lookupActive Then
        LastViewRow = viewLast
        Exit Function
    End If
    Set ws = ThisWorkbook.Worksheets(VIEW_NAME)
    LastViewRow = FIRST_VIEW - 1
    For Each col In Array(MATERIAL_NAME_COL, STOCK_COL, CODE_COL, UNIT_COL, REPLENISHMENT_STEP_COL, TARE_COL, KEY_COL)
        r = LastRow(ws, CLng(col))
        If r > LastViewRow Then LastViewRow = r
    Next col
End Function

Private Function IsRegisteredMaterial(ByVal r As Long) As Boolean
    If lookupActive And r >= FIRST_VIEW And r <= viewLast Then
        IsRegisteredMaterial = Len(Trim$(CStr(viewRows(r - FIRST_VIEW + 1, MATERIAL_NAME_COL)))) > 0 And Len(Trim$(CStr(viewRows(r - FIRST_VIEW + 1, CODE_COL)))) > 0
    Else
        With ThisWorkbook.Worksheets(VIEW_NAME)
            IsRegisteredMaterial = Len(Trim$(CStr(.cells(r, MATERIAL_NAME_COL).Value2))) > 0 And Len(Trim$(CStr(.cells(r, CODE_COL).Value2))) > 0
        End With
    End If
End Function

Private Function MaterialRow(ByVal code As String) As Long
    Dim ownsLookup As Boolean
    ownsLookup = Not lookupActive
    If ownsLookup Then BeginLookup
    If materialIndex.Exists(Trim$(code)) Then MaterialRow = materialIndex(Trim$(code))
    If ownsLookup Then EndLookup
End Function

Public Function MaterialCount() As Long
    Dim ownsLookup As Boolean
    ownsLookup = Not lookupActive
    If ownsLookup Then BeginLookup
    MaterialCount = materialIndex.count
    If ownsLookup Then EndLookup
End Function

Private Function MaterialValues(ByVal r As Long) As Variant
    With ThisWorkbook.Worksheets(VIEW_NAME)
        MaterialValues = Array(CStr(.cells(r, CODE_COL).Value2), .cells(r, MATERIAL_NAME_COL).Value2, .cells(r, UNIT_COL).Value2, _
                               .cells(r, REPLENISHMENT_STEP_COL).Value2, .cells(r, TARE_COL).Value2, .cells(r, STOCK_COL).Value2)
    End With
End Function

Private Sub WriteMaterial(ByVal r As Long, ByVal values As Variant)
    With ThisWorkbook.Worksheets(VIEW_NAME)
        .cells(r, CODE_COL).Value2 = values(0)
        .cells(r, MATERIAL_NAME_COL).Value2 = values(1)
        .cells(r, UNIT_COL).Value2 = values(2)
        .cells(r, REPLENISHMENT_STEP_COL).Value2 = values(3)
        .cells(r, TARE_COL).Value2 = values(4)
        .cells(r, STOCK_COL).Value2 = values(5)
        .cells(r, KEY_COL).Value2 = values(0)
    End With
End Sub

Private Function lastUsageRow() As Long
    Dim col As Long, ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(USAGE_NAME)
    lastUsageRow = 1
    For col = USAGE_KEY_COL To USAGE_QUANTITY_COL
        lastUsageRow = WorksheetFunction.Max(lastUsageRow, LastRow(ws, col))
    Next col
End Function

Private Function MaterialReferenced(ByVal code As String, Optional ByVal zeroRows As Object = Nothing) As Boolean
    Dim ws As Worksheet, values As Variant, inputs As Variant, r As Long, c As Long
    Dim product As String, storedCode As String, key As String, related As Boolean, seen As Object
    Set ws = ThisWorkbook.Worksheets(CONFIG_NAME)
    values = ws.Range(CONFIG_DATA_RANGE_PREFIX & LastConfigRow()).Value2
    For r = 2 To UBound(values, 1)
        If StrComp(Trim$(CStr(values(r, CONFIG_CODE_COL))), code, vbTextCompare) = 0 Then
            MaterialReferenced = True
            Exit Function
        End If
    Next r
    Set ws = ThisWorkbook.Worksheets(USAGE_NAME)
    values = ws.Range(USAGE_DATA_RANGE_PREFIX & lastUsageRow()).Value2
    inputs = ws.Range(USAGE_DATA_RANGE_PREFIX & lastUsageRow()).formula
    Set seen = NewTextDictionary()
    For r = 2 To UBound(values, 1)
        key = CStr(values(r, USAGE_KEY_COL))
        storedCode = CStr(values(r, USAGE_CODE_COL))
        product = CStr(values(r, USAGE_PRODUCT_COL))
        related = StrComp(Trim$(storedCode), code, vbTextCompare) = 0 Or _
                  (StrComp(Right$(key, Len(code) + 1), ":" & code, vbTextCompare) = 0 And _
                   StrComp(key, PairKey(product, storedCode), vbBinaryCompare) <> 0)
        If related Then
            For c = USAGE_KEY_COL To USAGE_QUANTITY_COL
                If IsError(values(r, c)) Then GoTo Ambiguous
                If VarType(inputs(r, c)) = vbString Then
                    If Left$(CStr(inputs(r, c)), 1) = "=" Then GoTo Ambiguous
                End If
            Next c
            If Len(Trim$(product)) = 0 Or Not productIndex.Exists(product) Then GoTo Ambiguous
            If StrComp(storedCode, code, vbTextCompare) <> 0 Then GoTo Ambiguous
            If StrComp(key, PairKey(product, storedCode), vbBinaryCompare) <> 0 Then GoTo Ambiguous
            If seen.Exists(key) Then GoTo Ambiguous
            seen.Add key, True
            Select Case VarType(values(r, USAGE_QUANTITY_COL))
                Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                Case Else
                    GoTo Ambiguous
            End Select
            If CDbl(values(r, USAGE_QUANTITY_COL)) < 0 Then GoTo Ambiguous
            If CDbl(values(r, USAGE_QUANTITY_COL)) > 0 Then
                MaterialReferenced = True
                Exit Function
            End If
            If Not zeroRows Is Nothing Then zeroRows(CStr(r)) = True
        End If
    Next r
    Exit Function
Ambiguous:
    Err.Raise vbObjectError + 51, , "기록은 확실하지 않아 정리하지 않았습니다: " & code
End Function

Private Function MaterialChangeRows(ByVal deletedCodes As Object, ByVal Target As Range, ByVal physicalDelete As Boolean) As Object
    Dim code As Variant, r As Long, i As Long, oldKey As String, newCode As String
    Dim zeroRows As Object, edits As Object, identified As Object
    Dim unknownCodeEdit As Boolean
    Set zeroRows = NewTextDictionary()
    Set edits = NewTextDictionary()
    Set identified = NewTextDictionary()
    For r = FIRST_VIEW To LastViewRow()
        i = PreviousMaterialRow(r)
        If i > 0 Then
            oldKey = CStr(savedView(i, KEY_COL))
            If savedMaterials.Exists(oldKey) Then
                If identified.Exists(oldKey) Then Err.Raise vbObjectError + 55, , "자재의 이전 식별키가 중복됩니다."
                identified(oldKey) = True
                newCode = Trim$(CStr(viewRows(r - FIRST_VIEW + 1, CODE_COL)))
                If Not IsRegisteredMaterial(r) Then
                    deletedCodes(oldKey) = True
                ElseIf StrComp(oldKey, newCode, vbTextCompare) <> 0 Then
                    edits(oldKey) = True
                End If
            End If
        ElseIf Not physicalDelete And IsRegisteredMaterial(r) Then
            If Not Intersect(Target, ThisWorkbook.Worksheets(VIEW_NAME).cells(r, 1).Resize(1, KEY_COL)) Is Nothing Then unknownCodeEdit = True
        End If
    Next r
    For Each code In savedMaterials.keys
        If Not identified.Exists(code) Then
            If materialIndex.Exists(code) Then
                If MaterialReferenced(CStr(code)) Then Err.Raise vbObjectError + 40, , "참조 중인 자재의 이전 식별키를 확인할 수 없습니다: " & code
            ElseIf unknownCodeEdit Then
                edits(code) = True
            Else
                deletedCodes(code) = True
            End If
        End If
    Next code
    For Each code In edits.keys
        If MaterialReferenced(CStr(code), zeroRows) Then _
            Err.Raise vbObjectError + 40, , "제품 구성에 포함되었거나 사용량이 양수인 자재코드는 바꿀 수 없습니다: " & code
    Next code
    Set MaterialChangeRows = zeroRows
End Function

Private Sub ValidateZeroRemoval(ByVal deletedCodes As Object, ByVal zeroRows As Object)
    Dim ws As Worksheet, inputs As Variant, values As Variant, r As Long, c As Long, upper As Long, formula As String
    Set ws = ThisWorkbook.Worksheets(USAGE_NAME)
    upper = WorksheetFunction.Max(lastUsageRow(), LastRow(ws, USAGE_SELECTION_COL), LastRow(ws, 8))
    inputs = ws.Range("A1:H" & upper).formula
    values = ws.Range("A1:H" & upper).Value2
    For r = 1 To UBound(inputs, 1)
        If r > 1 Then
            If Not IsError(values(r, USAGE_CODE_COL)) Then
                If deletedCodes.Exists(Trim$(CStr(values(r, USAGE_CODE_COL)))) Then GoTo NextRecord
            End If
        End If
        For c = 1 To UBound(inputs, 2)
            If VarType(inputs(r, c)) = vbString Then
                formula = CStr(inputs(r, c))
                If Left$(formula, 1) = "=" Then
                    If (c <> USAGE_QUANTITY_COL And c <> 8) Or Not IsNumeric(Mid$(formula, 2)) Then _
                        Err.Raise vbObjectError + 54, , "다른 수식에서 참조하고 있는 것으로 보입니다. 기록을 정리하지 않았습니다."
                End If
            End If
        Next c
NextRecord:
    Next r
End Sub

Private Sub SaveViewRecovery(Optional ByVal changedRows As Object = Nothing)
    Dim ws As Worksheet, r As Long, key As String
    Set ws = ThisWorkbook.Worksheets(VIEW_NAME)
    savedView = ws.Range(VIEW_DATA_RANGE_PREFIX & LastViewRow()).FormulaR1C1
    Set savedViewIndex = NewTextDictionary()
    For r = 1 To UBound(savedView, 1)
        key = CStr(savedView(r, KEY_COL))
        If Len(key) > 0 Then savedViewIndex(key) = r
    Next r
End Sub

Private Function PreviousMaterialRow(ByVal r As Long) As Long
    Dim ws As Worksheet, key As String, name As String, i As Long, found As Long
    Dim present As Object, keys As Variant
    Set ws = ThisWorkbook.Worksheets(VIEW_NAME)
    If lookupActive Then
        key = CStr(viewRows(r - FIRST_VIEW + 1, KEY_COL))
        name = CStr(viewRows(r - FIRST_VIEW + 1, MATERIAL_NAME_COL))
    Else
        key = CStr(ws.cells(r, KEY_COL).Value2)
        name = CStr(ws.cells(r, MATERIAL_NAME_COL).Value2)
    End If
    If savedViewIndex.Exists(key) Then
        PreviousMaterialRow = savedViewIndex(key)
        Exit Function
    End If
    Set present = NewTextDictionary()
    keys = ReadColumnArray(ws.Range("M" & FIRST_VIEW & ":M" & LastViewRow()))
    For i = 1 To UBound(keys, 1)
        present(CStr(keys(i, 1))) = True
    Next i
    If Len(name) > 0 Then
        For i = 1 To UBound(savedView, 1)
            If CStr(savedView(i, MATERIAL_NAME_COL)) = name And Not present.Exists(CStr(savedView(i, KEY_COL))) Then
                If found > 0 Then Exit Function
                found = i
            End If
        Next i
    End If
    PreviousMaterialRow = found
End Function

Private Sub RestoreCatalogInputs(ByVal Target As Range, ByVal before As Variant, ByVal deleted As Boolean)
    Dim ws As Worksheet, table As ListObject, present As Object, i As Long, r As Long, c As Long, key As String
    Dim cell As Range, oldFill As Boolean, number As Long, description As String
    Set ws = ThisWorkbook.Worksheets(VIEW_NAME)
    Set table = ws.ListObjects(1)
    oldFill = Application.AutoCorrect.AutoFillFormulasInLists
    On Error GoTo Failed
    Application.AutoCorrect.AutoFillFormulasInLists = False
    If deleted Then
        Set present = NewTextDictionary()
        For r = FIRST_VIEW To LastViewRow()
            present(CStr(ws.cells(r, KEY_COL).Value2)) = True
        Next r
        r = WorksheetFunction.Max(FIRST_VIEW, Target.Row)
        For i = 1 To UBound(before, 1)
            key = CStr(before(i, KEY_COL))
            If Len(key) > 0 And Not present.Exists(key) Then
                table.ListRows.Add Position:=WorksheetFunction.Min(r - FIRST_VIEW + 1, table.ListRows.count + 1)
                For c = 1 To KEY_COL
                    ws.cells(r, c).FormulaR1C1 = before(i, c)
                Next c
                r = r + 1
            End If
        Next i
    Else
        For Each cell In Intersect(Target, ws.Range(VIEW_DATA_RANGE_PREFIX & LastViewRow())).cells
            i = PreviousMaterialRow(cell.Row)
            If i = 0 Then i = cell.Row - FIRST_VIEW + 1
            If i <= UBound(before, 1) Then
                cell.FormulaR1C1 = before(i, cell.Column)
            Else
                cell.ClearContents
            End If
        Next cell
    End If
    Application.AutoCorrect.AutoFillFormulasInLists = oldFill
    Exit Sub
Failed:
    number = Err.number
    description = Err.description
    Application.AutoCorrect.AutoFillFormulasInLists = oldFill
    Err.Raise number, , description
End Sub

Private Sub RemoveMaterialRecords(ByVal codes As Object, ByVal zeroRows As Object, ByVal journal As Collection, ByRef backup As Workbook)
    Dim ws As Worksheet, name As Variant, r As Long, width As Long, values As Variant, remove As Boolean, key As Variant
    Dim area As Range, record As Variant, plans As New Collection, plan As Variant, slot As Long
    For Each name In Array(CONFIG_NAME, USAGE_NAME)
        Set ws = ThisWorkbook.Worksheets(CStr(name))
        If name = CONFIG_NAME Then width = 7 Else width = 4
        If name = CONFIG_NAME Then r = LastConfigRow() Else r = lastUsageRow()
        values = ws.cells(1, 1).Resize(r, width).Value2
        For r = UBound(values, 1) To 2 Step -1
            If name = CONFIG_NAME Then
                remove = codes.Exists(Trim$(CStr(values(r, CONFIG_CODE_COL))))
            Else
                remove = codes.Exists(Trim$(CStr(values(r, USAGE_CODE_COL)))) Or zeroRows.Exists(CStr(r))
                For Each key In codes.keys
                    If Not IsError(values(r, USAGE_KEY_COL)) Then
                        If StrComp(Right$(CStr(values(r, USAGE_KEY_COL)), Len(key) + 1), ":" & key, vbTextCompare) = 0 Then
                            If Not remove And CStr(values(r, USAGE_KEY_COL)) <> PairKey(CStr(values(r, USAGE_PRODUCT_COL)), CStr(values(r, USAGE_CODE_COL))) Then _
                                Err.Raise vbObjectError + 55, , "삭제 대상 사용량의 코드와 저장 키가 불일치합니다: " & key
                        End If
                    End If
                Next key
            End If
            If remove Then plans.Add Array(CStr(name), r, width)
        Next r
    Next name
    If plans.count = 0 Then Exit Sub
    Set backup = Application.Workbooks.Add(xlWBATWorksheet)
    backup.Windows(1).Visible = False
    For Each plan In plans
        name = plan(0)
        Set area = ThisWorkbook.Worksheets(CStr(name)).cells(CLng(plan(1)), 1).Resize(1, CLng(plan(2)))
        slot = slot + 1
        area.Copy Destination:=backup.Worksheets(1).cells(slot, 1)
        record = Array(CStr(name), plan(1), plan(2), slot, area.FormulaR1C1, area.Value2)
        area.Delete Shift:=xlUp
        journal.Add record
        ' Cascade fault anchor: after record removal.
    Next plan
End Sub

Private Sub RestoreRemovedRecords(ByVal journal As Collection, ByVal backup As Workbook)
    Dim i As Long, record As Variant, area As Range, source As Range
    For i = journal.count To 1 Step -1
        record = journal(i)
        Set area = ThisWorkbook.Worksheets(CStr(record(0))).cells(CLng(record(1)), 1).Resize(1, CLng(record(2)))
        area.Insert Shift:=xlDown
        Set area = ThisWorkbook.Worksheets(CStr(record(0))).cells(CLng(record(1)), 1).Resize(1, CLng(record(2)))
        Set source = backup.Worksheets(1).cells(CLng(record(3)), 1).Resize(1, CLng(record(2)))
        RestoreRecordFormats area, source
        area.FormulaR1C1 = record(4)
        RestoreNumericText area, record(5), record(4)
    Next i
End Sub

Private Sub RestoreRecordFormats(ByVal area As Range, ByVal source As Range)
    ' UI-only protection permits property writes, but blocks XML import and cross-book PasteAll.
    Dim c As Long, property As Variant, group As Variant, properties As Variant, edge As Long
    Dim current As Object, original As Object
    For c = 1 To area.Columns.count
        For Each property In Array("NumberFormat", "HorizontalAlignment", "VerticalAlignment", "Orientation", "WrapText", "IndentLevel", "ShrinkToFit", "ReadingOrder", "Locked", "FormulaHidden")
            CallByName area.cells(1, c), CStr(property), VbLet, CallByName(source.cells(1, c), CStr(property), VbGet)
        Next property
        For Each group In Array("Font", "Interior")
            Set current = CallByName(area.cells(1, c), CStr(group), VbGet)
            Set original = CallByName(source.cells(1, c), CStr(group), VbGet)
            If group = "Font" Then
                properties = Array("Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Color", "Subscript", "Superscript")
            Else
                properties = Array("Color", "PatternColor", "Pattern")
            End If
            For Each property In properties
                CallByName current, CStr(property), VbLet, CallByName(original, CStr(property), VbGet)
            Next property
        Next group
        For edge = xlDiagonalDown To xlInsideHorizontal
            With area.cells(1, c).Borders(edge)
                .LineStyle = source.cells(1, c).Borders(edge).LineStyle
                If .LineStyle <> xlNone Then
                    .Weight = source.cells(1, c).Borders(edge).Weight
                    .Color = source.cells(1, c).Borders(edge).Color
                End If
            End With
        Next edge
    Next c
End Sub

Private Function ConfigRow(ByVal product As String, ByVal code As String) As Long
    Dim ownsLookup As Boolean, key As String
    ownsLookup = Not lookupActive
    If ownsLookup Then BeginLookup
    If product = ALL_PRODUCT Then
        If unionIndex.Exists(code) Then ConfigRow = unionIndex(code)
    Else
        key = PairKey(product, code)
        If configIndex.Exists(key) Then ConfigRow = configIndex(key)
    End If
    If ownsLookup Then EndLookup
End Function

Private Function UsageRow(ByVal product As String, ByVal code As String, Optional ByVal create As Boolean = False) As Long
    Dim ownsLookup As Boolean, key As String, r As Long, ws As Worksheet
    ownsLookup = Not lookupActive
    If ownsLookup Then BeginLookup
    key = PairKey(product, code)
    If usageIndex.Exists(key) Then r = usageIndex(key)
    If r = 0 And create Then
        Set ws = ThisWorkbook.Worksheets(USAGE_NAME)
        r = LastRow(ws) + 1
        ws.cells(r, USAGE_KEY_COL).Resize(1, 4).Value2 = Array(key, product, code, 0)
        BeginLookup
    End If
    UsageRow = r
    If ownsLookup Then EndLookup
End Function

Private Function ReadNonnegativeQuantity(ByVal value As Variant, Optional ByVal label As String = "수량") As Double
    If IsError(value) Then Err.Raise vbObjectError + 1, , label & "을 확인하세요."
    If Len(CStr(value)) = 0 Then Exit Function
    If Not IsNumeric(value) Then Err.Raise vbObjectError + 2, , label & "에는 0 이상의 숫자를 입력하세요."
    If CDbl(value) < 0 Then Err.Raise vbObjectError + 3, , label & "에는 0 이상의 숫자를 입력하세요."
    ReadNonnegativeQuantity = CDbl(value)
End Function

Public Function ProductUsage(ByVal product As String, ByVal code As String) As Double
    Dim ownsLookup As Boolean, r As Long
    ownsLookup = Not lookupActive
    If ownsLookup Then BeginLookup
    If product = ALL_PRODUCT Then
        If usageTotals.Exists(code) Then ProductUsage = usageTotals(code)
    Else
        r = UsageRow(product, code)
        If r > 0 Then ProductUsage = ReadNonnegativeQuantity(useRows(r, USAGE_QUANTITY_COL))
    End If
    If ownsLookup Then EndLookup
End Function

Public Function ProductCount(ByVal product As String) As Long
    Dim ownsLookup As Boolean
    ownsLookup = Not lookupActive
    If ownsLookup Then BeginLookup
    If product = ALL_PRODUCT Then
        ProductCount = unionIndex.count
    ElseIf memberCounts.Exists(product) Then
        ProductCount = memberCounts(product)
    End If
    If ownsLookup Then EndLookup
End Function

Public Function CurrentProduct() As String
    CurrentProduct = activeProduct
End Function

Public Function errorText() As String
    errorText = LastError
End Function

Private Function CalculateReplenishment(ByVal opening As Double, ByVal used As Double, ByVal lot As Double) As Double
    Dim shortage As Double
    shortage = WorksheetFunction.Round(used - opening, 6)
    If shortage <= 0 Then Exit Function
    If lot <= 0 Then Err.Raise vbObjectError + 4, , "부족분을 계산하려면 입고량을 0보다 크게 입력하세요."
    CalculateReplenishment = WorksheetFunction.Round(WorksheetFunction.RoundUp(shortage / lot, 0) * lot, 6)
End Function

Private Sub ValidateStorage()
    Dim r As Long, field As Variant, code As String, product As String
    Dim seen As Object, key As String, n As Double, savedCode As Variant, col As Long, configInput As Variant
    If Not lookupActive Then BeginLookup
    configInput = ThisWorkbook.Worksheets(CONFIG_NAME).Range(CONFIG_DATA_RANGE_PREFIX & LastConfigRow()).formula
    Set seen = NewTextDictionary()
    For r = 2 To UBound(productRows, 1)
        product = Trim$(CStr(productRows(r, PRODUCT_LIST_NAME_COL)))
        If product = ALL_PRODUCT Then Err.Raise vbObjectError + 35, , "전체는 합산 조회용 이름입니다. 다른 제품명을 입력하세요."
        If Len(product) = 0 Then Err.Raise vbObjectError + 5, , "제품 목록 중간에 빈 제품명이 있습니다. 제품을 지우려면 행 전체를 삭제하세요."
        If seen.Exists(product) Then Err.Raise vbObjectError + 6, , "제품명이 중복됩니다: " & product
        seen.Add product, True
    Next r
    If seen.count = 0 Then Err.Raise vbObjectError + 7, , "제품을 먼저 등록하세요."
    seen.RemoveAll
    If viewLast > FIRST_VIEW + 1999 Then Err.Raise vbObjectError + 8, , "자재는 최대 2,000개까지 지원합니다."
    For r = 1 To UBound(viewRows, 1)
        code = Trim$(CStr(viewRows(r, CODE_COL)))
        If Len(code) > 0 Then
            If seen.Exists(code) Then Err.Raise vbObjectError + 10, , "자재코드가 중복됩니다: " & code
            seen.Add code, True
        End If
        For Each field In Array(STOCK_COL, UNIT_COL, REPLENISHMENT_STEP_COL)
            n = ReadNonnegativeQuantity(viewRows(r, CLng(field)), "자재 수량")
        Next field
        If IsError(viewRows(r, TARE_COL)) Then Err.Raise vbObjectError + 39, , "용기중량 값을 확인하세요."
        If IsNumeric(viewRows(r, TARE_COL)) Then n = ReadNonnegativeQuantity(viewRows(r, TARE_COL), "용기중량")
    Next r
    seen.RemoveAll
    For r = 2 To UBound(configRows, 1)
        product = CStr(configRows(r, CONFIG_PRODUCT_COL))
        code = CStr(configRows(r, CONFIG_CODE_COL))
        For col = CONFIG_FIRST_CAPACITY_COL To CONFIG_LAST_CAPACITY_COL
            n = ReadNonnegativeQuantity(configRows(r, col), "용량별 사용량")
            If VarType(configRows(r, col)) = vbString Then
                If Len(CStr(configRows(r, col))) > 0 Then _
                    Err.Raise vbObjectError + 49, , "용량별 사용량에는 문자 대신 0 이상의 숫자를 입력하세요."
            End If
            If VarType(configInput(r, col)) = vbString Then
                If Left$(CStr(configInput(r, col)), 1) = "=" Then _
                    Err.Raise vbObjectError + 50, , "용량별 사용량에는 수식 대신 숫자를 직접 입력하세요."
            End If
        Next col
        If Len(Trim$(product)) > 0 Then
            If Not productIndex.Exists(Trim$(product)) Then Err.Raise vbObjectError + 11, , "등록되지 않은 제품입니다: " & product
        End If
        If Len(Trim$(code)) > 0 Then
            If MaterialRow(code) = 0 Then Err.Raise vbObjectError + 12, , "등록되지 않은 자재입니다: " & code
        End If
        If Len(Trim$(product)) > 0 And Len(Trim$(code)) > 0 Then
            key = PairKey(Trim$(product), Trim$(code))
            If seen.Exists(key) Then Err.Raise vbObjectError + 13, , "같은 제품에 자재가 중복됩니다: " & code
            seen.Add key, True
        End If
    Next r
    seen.RemoveAll
    For r = 2 To UBound(useRows, 1)
        key = CStr(useRows(r, USAGE_KEY_COL))
        product = CStr(useRows(r, USAGE_PRODUCT_COL))
        code = CStr(useRows(r, USAGE_CODE_COL))
        If Len(Trim$(product)) = 0 Or Len(Trim$(code)) = 0 Then Err.Raise vbObjectError + 52, , "사용량 저장소의 제품·코드를 확인하세요."
        If Not productIndex.Exists(product) Or StrComp(key, PairKey(product, code), vbBinaryCompare) <> 0 Then _
            Err.Raise vbObjectError + 53, , "사용량 저장키가 제품·코드와 일치하지 않습니다: " & product & " / " & code
        If seen.Exists(key) Then Err.Raise vbObjectError + 16, , "사용량 저장키가 중복됩니다."
        seen.Add key, True
        n = ReadNonnegativeQuantity(useRows(r, USAGE_QUANTITY_COL), "금일사용")
        If n > 0 Then
            If MaterialRow(code) = 0 Then Err.Raise vbObjectError + 17, , "사용량이 남은 자재코드가 데이터베이스에 없습니다: " & code
        End If
    Next r
End Sub

Private Sub SaveStorageState(Optional ByVal changedRows As Object = Nothing)
    Dim ws As Worksheet, r As Long, code As String, values As Variant, draft(1 To 1, 1 To 13) As Variant, c As Long
    If changedRows Is Nothing Then
        Set savedMaterials = NewTextDictionary()
        Set savedDrafts = NewTextDictionary()
    End If
    Set ws = ThisWorkbook.Worksheets(VIEW_NAME)
    values = ws.Range(VIEW_DATA_RANGE_PREFIX & LastViewRow()).Value2
    For r = 1 To UBound(values, 1)
        If Not changedRows Is Nothing Then
            If Not changedRows.Exists(CStr(r + FIRST_VIEW - 1)) Then GoTo NextSavedRow
        End If
        code = CStr(values(r, CODE_COL))
        If Len(Trim$(code)) > 0 And Len(Trim$(CStr(values(r, MATERIAL_NAME_COL)))) > 0 Then
            savedMaterials(code) = Array(code, values(r, MATERIAL_NAME_COL), values(r, UNIT_COL), values(r, REPLENISHMENT_STEP_COL), values(r, TARE_COL), values(r, STOCK_COL))
        Else
            For c = 1 To 13
                draft(1, c) = values(r, c)
            Next c
            savedDrafts(CStr(r + FIRST_VIEW - 1)) = draft
        End If
NextSavedRow:
    Next r
    If Not changedRows Is Nothing Then
        SaveViewRecovery changedRows
        Exit Sub
    End If
    SaveViewRecovery
    Set ws = ThisWorkbook.Worksheets(CONFIG_NAME)
    savedConfig = ws.Range(CONFIG_DATA_RANGE_PREFIX & LastConfigRow()).formula
    savedConfigValues = ws.Range(CONFIG_DATA_RANGE_PREFIX & LastConfigRow()).Value2
    Set ws = ThisWorkbook.Worksheets(PRODUCT_NAME)
    savedProducts = ws.Range(PRODUCT_DATA_RANGE_PREFIX & LastRow(ws)).Value2
End Sub

Private Sub RestoreMaterialState()
    Dim ws As Worksheet, seen As Object, r As Long, n As Long, key As String, code As Variant
    If savedMaterials Is Nothing Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(VIEW_NAME)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    n = LastViewRow()
    For r = FIRST_VIEW To n
        key = CStr(ws.cells(r, KEY_COL).Value2)
        If Not savedMaterials.Exists(key) Then
            If savedDrafts.Exists(CStr(r)) Then
                ws.Range("A" & r & ":M" & r).Value2 = savedDrafts(CStr(r))
                GoTo NextRow
            End If
            key = CStr(ws.cells(r, CODE_COL).Value2)
        End If
        If savedMaterials.Exists(key) And Not seen.Exists(key) Then
            WriteMaterial r, savedMaterials(key)
            seen(key) = True
        Else
            ws.Range("A" & r & ":M" & r).ClearContents
        End If
NextRow:
    Next r
    For Each code In savedMaterials.keys
        If Not seen.Exists(CStr(code)) Then
            n = n + 1
            WriteMaterial n, savedMaterials(CStr(code))
        End If
    Next code
End Sub

Private Sub RestoreNumericText(ByVal area As Range, ByVal values As Variant, ByVal formulas As Variant)
    Dim r As Long, c As Long, numberFormat As Variant
    For r = 1 To UBound(values, 1)
        For c = 1 To UBound(values, 2)
            If VarType(values(r, c)) = vbString Then
                If IsNumeric(values(r, c)) And Left$(CStr(formulas(r, c)), 1) <> "=" Then
                    numberFormat = area.cells(r, c).numberFormat
                    area.cells(r, c).numberFormat = "@"
                    area.cells(r, c).Value2 = values(r, c)
                    area.cells(r, c).numberFormat = numberFormat
                End If
            End If
        Next c
    Next r
End Sub

Private Sub RestoreStorageState(ByVal sheetName As String)
    Dim ws As Worksheet, state As Variant, filters As Variant, field As Long, values As Variant, upper As Long
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Select Case sheetName
        Case CONFIG_NAME
            state = savedConfig
            values = savedConfigValues
        Case PRODUCT_NAME
            state = savedProducts
            values = savedProducts
        Case Else
            Exit Sub
    End Select
    If Not IsArray(state) Then Exit Sub
    If sheetName = CONFIG_NAME And ws.AutoFilterMode Then
        filters = CaptureFilters(ws.AutoFilter)
        For field = 1 To ws.AutoFilter.filters.count
            If ws.AutoFilter.filters(field).On Then ws.AutoFilter.Range.AutoFilter field:=field
        Next field
    End If
    upper = WorksheetFunction.Max(UBound(state, 1), ws.UsedRange.Row + ws.UsedRange.Rows.count - 1)
    ws.cells(1, 1).Resize(upper, UBound(state, 2)).ClearContents
    ws.cells(1, 1).Resize(UBound(state, 1), UBound(state, 2)).formula = state
    RestoreNumericText ws.cells(1, 1).Resize(UBound(state, 1), UBound(state, 2)), values, state
    If IsArray(filters) Then RestoreFilters ws.cells(1, 1).Resize(UBound(state, 1), UBound(state, 2)), filters
End Sub

Private Sub UpdateNames(Optional ByVal productsChanged As Boolean = True, Optional ByVal configChanged As Boolean = True, Optional ByVal configTarget As Range = Nothing)
    Dim n As Long, r As Long, options As Worksheet, values() As Variant
    Dim formulas As Range, area As Range, filters As Variant
    If productsChanged Then
        n = LastRow(ThisWorkbook.Worksheets(PRODUCT_NAME))
        ThisWorkbook.Names("ProductNames").RefersTo = "='" & PRODUCT_NAME & "'!$A$2:$A$" & n
        Set options = ThisWorkbook.Worksheets(USAGE_NAME)
        options.Range(USAGE_SELECTION_RANGE_PREFIX & LastRow(options, USAGE_SELECTION_COL)).ClearContents
        ReDim values(1 To n + 1, 1 To 1)
        values(1, 1) = "제품 선택"
        values(2, 1) = ALL_PRODUCT
        For r = 2 To n
            values(r + 1, 1) = productRows(r, PRODUCT_LIST_NAME_COL)
        Next r
        options.Range(USAGE_SELECTION_RANGE_PREFIX & n + 1).Value2 = values
        ThisWorkbook.Names("ProductSelection").RefersTo = "='" & USAGE_NAME & "'!$F$2:$F$" & n + 1
    End If
    n = WorksheetFunction.Max(FIRST_VIEW, LastViewRow())
    ThisWorkbook.Names("MaterialCodes").RefersTo = "='" & VIEW_NAME & "'!$B$" & FIRST_VIEW & ":$B$" & n
    If configChanged Then
        With ThisWorkbook.Worksheets(CONFIG_NAME)
            n = LastConfigRow()
            Set area = .Range(CONFIG_DATA_RANGE_PREFIX & n)
            If .AutoFilterMode Then
                If .AutoFilter.Range.Address <> area.Address Then
                    filters = CaptureFilters(.AutoFilter)
                    area.AutoFilter
                    RestoreFilters area, filters
                End If
            Else
                area.AutoFilter
            End If
            If n >= 2 Then
                Set formulas = .Range("C2:C" & n)
                If Not configTarget Is Nothing Then Set formulas = Intersect(formulas, configTarget.EntireRow)
                If Not formulas Is Nothing Then formulas.FormulaR1C1 = CONFIG_FORMULA
            End If
        End With
    End If
End Sub

Private Function HasValidation(ByVal cell As Range) As Boolean
    Dim kind As Long
    On Error Resume Next
    kind = cell.Validation.Type
    HasValidation = Err.number = 0
    Err.Clear
End Function

Private Sub SetValidation(ByVal cells As Range, ByVal kind As Long, ByVal formula As String, Optional ByVal message As String = "")
    With cells.Validation
        .Delete
        If kind = xlValidateList Then
            .Add Type:=kind, AlertStyle:=xlValidAlertStop, Formula1:=formula
        Else
            .Add Type:=kind, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:=formula
        End If
        .IgnoreBlank = True
        .ShowError = True
        If Len(message) > 0 Then .ErrorMessage = message
    End With
End Sub

Private Sub RestoreBlankRowFormat(ByVal block As Range)
    Dim c As Long, edge As Variant
    block.EntireRow.RowHeight = 17.5
    With block
        .Font.name = "맑은 고딕"
        .Font.Size = 11.5
        .Font.Bold = False
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .ShrinkToFit = False
        .numberFormat = "General"
        .Locked = True
        .Borders.LineStyle = xlNone
    End With
    For c = 1 To KEY_COL
        With block.Columns(c)
            Select Case c
                Case MATERIAL_NAME_COL, CODE_COL
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                    .ShrinkToFit = True
                Case MEMBER_COL
                    .HorizontalAlignment = xlCenter
                Case REPLENISHMENT_STEP_COL
                    .ShrinkToFit = True
            End Select
            Select Case c
                Case MATERIAL_NAME_COL, CODE_COL, STOCK_COL, USE_COL, UNIT_COL, REPLENISHMENT_STEP_COL, TARE_COL, MEMBER_COL
                    .Locked = False
            End Select
            If c = STOCK_COL Or c = USE_COL Then .Interior.Color = RGB(242, 242, 242)
            If c = CODE_COL Then .numberFormat = "@"
            If c = RECEIPT_COL Then .numberFormat = "0"
            For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal)
                If edge <> xlInsideHorizontal Or block.Rows.count > 1 Then
                    If c <= TARE_COL Then
                        With .Borders(edge)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .Color = RGB(128, 128, 128)
                        End With
                    ElseIf edge = xlEdgeBottom Or edge = xlInsideHorizontal Then
                        With .Borders(edge)
                            .LineStyle = xlContinuous
                            .Weight = xlHairline
                            .Color = RGB(217, 217, 217)
                        End With
                    End If
                End If
            Next edge
        End With
    Next c
End Sub

Private Sub RepairInputRules()
    Dim db As Worksheet, cfg As Worksheet, area As Range, i As Long, lastBody As Long, first As Long, col As Variant
    On Error Resume Next
    Set db = ThisWorkbook.Worksheets(VIEW_NAME)
    lastBody = FIRST_VIEW + 1999
    For i = 1 To db.cells.FormatConditions.count
        Set area = Nothing
        Set area = db.cells.FormatConditions(i).AppliesTo
        If Not area Is Nothing Then
            If area.Areas.count = 1 And area.Row = FIRST_VIEW And area.Row + area.Rows.count - 1 < lastBody Then _
                db.cells.FormatConditions(i).ModifyAppliesToRange db.Range(db.cells(FIRST_VIEW, area.Column), db.cells(lastBody, area.Column + area.Columns.count - 1))
        End If
    Next i
    For Each col In Array(STOCK_COL, USE_COL)
        If Not HasValidation(db.cells(lastBody, CLng(col))) Then SetValidation db.Range(db.cells(FIRST_VIEW, CLng(col)), db.cells(lastBody, CLng(col))), xlValidateDecimal, "0"
    Next col
    If Not HasValidation(db.cells(lastBody, MEMBER_COL)) Then _
        SetValidation db.Range(db.cells(FIRST_VIEW, MEMBER_COL), db.cells(lastBody, MEMBER_COL)), xlValidateList, "포함,제외", "목록에서 선택하거나 올바른 값을 입력하세요."
    first = lastBody + 1
    Do While first > FIRST_VIEW
        If Not db.cells(first - 1, MATERIAL_NAME_COL).Locked Then Exit Do
        If db.cells(first - 1, STOCK_COL).Borders(xlEdgeLeft).LineStyle <> xlNone Then Exit Do
        first = first - 1
    Loop
    If first <= lastBody Then RestoreBlankRowFormat db.Range(db.cells(first, 1), db.cells(lastBody, KEY_COL))
    Set cfg = ThisWorkbook.Worksheets(CONFIG_NAME)
    If Not HasValidation(cfg.cells(10001, CONFIG_PRODUCT_COL)) Then _
        SetValidation cfg.Range("A2:A10001"), xlValidateList, "=ProductNames", "목록에서 선택하거나 올바른 값을 입력하세요."
    If Not HasValidation(cfg.cells(10001, CONFIG_CODE_COL)) Then _
        SetValidation cfg.Range("B2:B10001"), xlValidateList, "=MaterialCodes", "목록에서 선택하거나 올바른 값을 입력하세요."
    If Not HasValidation(cfg.cells(10001, CONFIG_LAST_CAPACITY_COL)) Then _
        SetValidation cfg.Range("D2:G10001"), xlValidateDecimal, "0", "목록에서 선택하거나 올바른 값을 입력하세요."
    Err.Clear
End Sub

Private Function StateMissing() As Boolean
    StateMissing = Len(activeProduct) = 0 Or savedMaterials Is Nothing Or Not IsArray(savedView)
End Function

Private Sub InitializeState()
    Dim db As Worksheet, products As Worksheet, r As Long
    BeginLookup
    ValidateStorage
    Set db = ThisWorkbook.Worksheets(VIEW_NAME)
    Set products = ThisWorkbook.Worksheets(PRODUCT_NAME)
    activeProduct = CStr(db.Range(PRODUCT_CELL).Value2)
    If activeProduct <> ALL_PRODUCT Then
        r = FindRow(products, activeProduct)
        If r = 0 Then r = 2
        activeProduct = Trim$(CStr(products.cells(r, PRODUCT_LIST_NAME_COL).Value2))
    End If
    printAllItems = CBool(db.Range(PRINT_ALL_CELL).Value2)
    EnsureMacroAccess
    RepairInputRules
    UpdateNames
    SetViewFormulas FIRST_VIEW, viewLast
    If CStr(db.Range(PRODUCT_CELL).Value2) <> activeProduct Then db.Range(PRODUCT_CELL).Value2 = activeProduct
    RefreshView False
    SaveStorageState
End Sub

Public Sub InitializeIntegrated(Optional ByVal showMessage As Boolean = False)
    Dim oldEvents As Boolean, oldScreen As Boolean
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Busy = True
    InitializeState
    LastError = ""
    GoTo Finish
Failed:
    LastError = Err.description
Finish:
    EndLookup
    Busy = False
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    If showMessage And Len(LastError) > 0 Then _
        MsgBox "저장소에 오류가 있어 입력을 처리할 수 없습니다. 아래 내용을 고치면 다시 사용할 수 있습니다." & vbCrLf & vbCrLf & LastError, vbExclamation, "저장소 확인"
End Sub

Private Sub RecoverState(ByVal ws As Worksheet, ByVal Target As Range)
    Dim inputs As Range, area As Range, pending As New Collection, item As Variant
    If ws.name = VIEW_NAME Then
        Set inputs = Intersect(Target, Union(ws.Columns(USE_COL), ws.Columns(MEMBER_COL)), ws.Range(VIEW_DATA_RANGE_PREFIX & LastViewRow()))
        If Not inputs Is Nothing Then
            For Each area In inputs.Areas
                pending.Add Array(area, area.formula)
            Next area
        End If
    End If
    On Error GoTo Failed
    InitializeState
    For Each item In pending
        item(0).formula = item(1)
    Next item
    Exit Sub
Failed:
    Err.Raise vbObjectError + 56, , "저장소에 오류가 있어 입력을 처리하지 않았습니다. 다음 내용을 고치세요: " & Err.description
End Sub

Private Sub RefreshView(Optional ByVal resetOrder As Boolean = False)
    Dim ws As Worksheet, r As Long, n As Long, code As String, count As Long, oldLast As Long
    Dim uses() As Variant, members() As Variant, keys() As Variant
    Dim table As ListObject, col As Variant, value As Variant, formulas As Variant, stale As Boolean
    Set ws = ThisWorkbook.Worksheets(VIEW_NAME)
    If Not lookupActive Then BeginLookup
    n = viewLast
    count = n - FIRST_VIEW + 1
    Set table = ws.ListObjects(1)
    oldLast = table.Range.Row + table.Range.Rows.count - 1
    If oldLast <> n Then
        If ws.ProtectContents Then Err.Raise vbObjectError + 44, , "자재 추가·삭제는 검토 → 시트 보호 해제 후 진행하세요."
        table.Resize ws.Range("A" & FIRST_VIEW - 1 & ":M" & n)
    End If
    If n > oldLast Then SetViewFormulas oldLast + 1, n
    ReDim uses(1 To count, 1 To 1)
    ReDim members(1 To count, 1 To 1)
    ReDim keys(1 To count, 1 To 1)
    For r = 1 To count
        code = Trim$(CStr(viewRows(r, CODE_COL)))
        If IsRegisteredMaterial(r + FIRST_VIEW - 1) Then
            If code <> CStr(viewRows(r, CODE_COL)) Then ws.cells(r + FIRST_VIEW - 1, CODE_COL).Value2 = code
            uses(r, 1) = ProductUsage(activeProduct, code)
            members(r, 1) = IIf(ConfigRow(activeProduct, code) > 0, "포함", "제외")
            keys(r, 1) = code
        ElseIf Len(Trim$(CStr(viewRows(r, MATERIAL_NAME_COL)))) > 0 Or Len(code) > 0 Then
            members(r, 1) = "미등록"
        End If
    Next r
    If ws.FilterMode Then
        formulas = ws.Range(VIEW_DATA_RANGE_PREFIX & n).formula
        For r = 1 To count
            For Each col In Array(USE_COL, MEMBER_COL, KEY_COL)
                Select Case CLng(col)
                    Case USE_COL
                        value = uses(r, 1)
                    Case MEMBER_COL
                        value = members(r, 1)
                    Case KEY_COL
                        value = keys(r, 1)
                End Select
                If IsError(viewRows(r, CLng(col))) Then
                    stale = True
                Else
                    stale = CStr(viewRows(r, CLng(col))) <> CStr(value)
                End If
                If VarType(formulas(r, CLng(col))) = vbString Then
                    If Left$(CStr(formulas(r, CLng(col))), 1) = "=" Then stale = True
                End If
                If stale Then ws.cells(r + FIRST_VIEW - 1, CLng(col)).Value2 = value
            Next col
        Next r
    Else
        ws.Range("E" & FIRST_VIEW & ":E" & n).Value2 = uses
        ws.Range("K" & FIRST_VIEW & ":K" & n).Value2 = members
        ws.Range("M" & FIRST_VIEW & ":M" & n).Value2 = keys
    End If
    If activeProduct = ALL_PRODUCT Then
        ws.Shapes("txtSummary").TextFrame2.TextRange.text = "전체 제품 합산 · 포함 " & ProductCount(activeProduct) & " / DB " & MaterialCount() & "품목"
        ws.Shapes("btnRoll").Fill.ForeColor.RGB = &H808080
        ws.Shapes("btnApplyCapacity").Fill.ForeColor.RGB = &H808080
    Else
        ws.Shapes("txtSummary").TextFrame2.TextRange.text = "선택 제품 " & ProductCount(activeProduct) & " / DB " & MaterialCount() & "품목"
        ws.Shapes("btnRoll").Fill.ForeColor.RGB = &H6B4217
        ws.Shapes("btnApplyCapacity").Fill.ForeColor.RGB = &H6B4217
    End If
    ws.Calculate
    If ws.FilterMode Then table.AutoFilter.ApplyFilter
    If resetOrder Then SortViewInternal
End Sub

Private Sub FillViewRow(ByVal r As Long, Optional ByVal repairFormulas As Boolean = True)
    Dim ws As Worksheet, code As String
    Set ws = ThisWorkbook.Worksheets(VIEW_NAME)
    code = Trim$(CStr(ws.cells(r, CODE_COL).Value2))
    If IsRegisteredMaterial(r) Then
        ws.cells(r, USE_COL).Value2 = ProductUsage(activeProduct, code)
        ws.cells(r, MEMBER_COL).Value2 = IIf(ConfigRow(activeProduct, code) > 0, "포함", "제외")
        ws.cells(r, KEY_COL).Value2 = code
    Else
        ws.cells(r, USE_COL).ClearContents
        If Len(Trim$(CStr(ws.cells(r, MATERIAL_NAME_COL).Value2))) > 0 Or Len(code) > 0 Then
            ws.cells(r, MEMBER_COL).Value2 = "미등록"
        Else
            ws.cells(r, MEMBER_COL).ClearContents
        End If
        ws.cells(r, KEY_COL).ClearContents
    End If
    If repairFormulas Then SetViewFormulas r, r
End Sub

Private Function CaptureFilters(ByVal filters As AutoFilter) As Variant
    Dim state() As Variant, i As Long
    ReDim state(1 To filters.filters.count, 1 To 4)
    For i = 1 To filters.filters.count
        With filters.filters(i)
            If .On Then
                state(i, 1) = True
                state(i, 2) = .Operator
                On Error Resume Next
                state(i, 3) = .Criteria1
                state(i, 4) = .Criteria2
                On Error GoTo 0
            End If
        End With
    Next i
    CaptureFilters = state
End Function

Private Sub RestoreFilters(ByVal area As Range, ByVal state As Variant)
    Dim i As Long
    For i = 1 To UBound(state, 1)
        If state(i, 1) = True Then
            If IsEmpty(state(i, 4)) Then
                If state(i, 2) = 0 Then
                    area.AutoFilter field:=i, Criteria1:=state(i, 3)
                Else
                    area.AutoFilter field:=i, Criteria1:=state(i, 3), Operator:=state(i, 2)
                End If
            ElseIf IsEmpty(state(i, 3)) Then
                area.AutoFilter field:=i, Operator:=state(i, 2), Criteria2:=state(i, 4)
            Else
                area.AutoFilter field:=i, Criteria1:=state(i, 3), Operator:=state(i, 2), Criteria2:=state(i, 4)
            End If
        End If
    Next i
End Sub

Private Sub SortViewInternal()
    Dim ws As Worksheet, table As ListObject, state As Variant, i As Long
    Dim errorNumber As Long, errorText As String
    Dim positions As Object, codes As Variant, members As Variant, keys() As Variant
    Dim r As Long, n As Long, product As String, code As String, keyed As Boolean
    Set ws = ThisWorkbook.Worksheets(VIEW_NAME)
    Set table = ws.ListObjects(1)
    state = CaptureFilters(table.AutoFilter)
    On Error GoTo Failed
    For i = 1 To table.ListColumns.count
        If state(i, 1) = True Then table.Range.AutoFilter field:=i
    Next i
    If Not lookupActive Then BeginLookup
    product = activeProduct
    If Len(product) = 0 Then product = Trim$(CStr(ws.Range(PRODUCT_CELL).Value2))
    Set positions = NewTextDictionary()
    For r = 2 To UBound(configRows, 1)
        code = Trim$(CStr(configRows(r, CONFIG_CODE_COL)))
        If Len(code) > 0 And Not positions.Exists(code) Then
            If Len(Trim$(CStr(configRows(r, CONFIG_PRODUCT_COL)))) > 0 And (product = ALL_PRODUCT Or _
               StrComp(Trim$(CStr(configRows(r, CONFIG_PRODUCT_COL))), product, vbTextCompare) = 0) Then positions(code) = r
        End If
    Next r
    n = table.ListRows.count
    If n > 0 Then
        codes = ReadColumnArray(table.ListColumns(CODE_COL).DataBodyRange)
        members = ReadColumnArray(table.ListColumns(MEMBER_COL).DataBodyRange)
        ReDim keys(1 To n, 1 To 1)
        For r = 1 To n
            keys(r, 1) = UBound(configRows, 1) + 1
            If Not IsError(codes(r, 1)) And Not IsError(members(r, 1)) Then
                code = Trim$(CStr(codes(r, 1)))
                If CStr(members(r, 1)) = "포함" And positions.Exists(code) Then keys(r, 1) = positions(code)
            End If
        Next r
        keyed = True
        table.ListColumns(SEARCH_COL).DataBodyRange.Value2 = keys
    End If
    With table.Sort
        .SortFields.Clear
        .SortFields.Add key:=table.ListColumns(SEARCH_COL).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add key:=table.ListColumns(MEMBER_COL).Range, SortOn:=xlSortOnValues, Order:=xlDescending
        .SortFields.Add key:=table.ListColumns(MATERIAL_NAME_COL).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add key:=table.ListColumns(CODE_COL).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
        .SortFields.Clear
    End With
    If keyed Then
        keyed = False
        SetViewFormulas FIRST_VIEW, FIRST_VIEW + n - 1, True
        table.ListColumns(SEARCH_COL).DataBodyRange.Calculate
    End If
    RestoreFilters table.Range, state
    Exit Sub
Failed:
    errorNumber = Err.number
    errorText = Err.description
    On Error Resume Next
    If keyed Then SetViewFormulas FIRST_VIEW, FIRST_VIEW + n - 1, True
    RestoreFilters table.Range, state
    On Error GoTo 0
    Err.Raise errorNumber, , errorText
End Sub

Public Sub RestoreDefaultSort()
    Dim oldEvents As Boolean, oldScreen As Boolean
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    EnsureMacroAccess
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Busy = True
    BeginLookup
    SortViewInternal
    LastError = ""
    GoTo Finish
Failed:
    LastError = Err.description
Finish:
    EndLookup
    Busy = False
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
End Sub

Private Sub ClearFiltersInternal()
    Dim ws As Worksheet, table As ListObject, i As Long
    Set ws = ThisWorkbook.Worksheets(VIEW_NAME)
    Set table = ws.ListObjects(1)
    For i = 1 To table.ListColumns.count
        If table.AutoFilter.filters(i).On Then table.Range.AutoFilter field:=i
    Next i
    ws.Range(MODE_CELL).Value2 = "전체"
    ws.Range(QUERY_RANGE).ClearContents
End Sub

Public Sub ResetViewFilters()
    Dim oldEvents As Boolean, oldScreen As Boolean
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    EnsureMacroAccess
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Busy = True
    ClearFiltersInternal
    LastError = ""
    GoTo Finish
Failed:
    LastError = Err.description
Finish:
    EndLookup
    Busy = False
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
End Sub

Private Sub ApplyFiltersInternal()
    Dim ws As Worksheet, table As ListObject, query As String, mode As String, field As Variant
    Set ws = ThisWorkbook.Worksheets(VIEW_NAME)
    Set table = ws.ListObjects(1)
    For Each field In Array(USE_COL, MEMBER_COL, SEARCH_COL)
        If table.AutoFilter.filters(CLng(field)).On Then table.Range.AutoFilter field:=CLng(field)
    Next field
    mode = CStr(ws.Range(MODE_CELL).Value2)
    Select Case mode
        Case "전체"
        Case "선택제품"
            table.Range.AutoFilter field:=MEMBER_COL, Criteria1:="포함"
        Case "금일사용"
            table.Range.AutoFilter field:=USE_COL, Criteria1:=">0"
        Case Else
            Err.Raise vbObjectError + 18, , "조회범위를 확인하세요."
    End Select
    query = Trim$(CStr(ws.Range(QUERY_CELL).Value2))
    If Len(query) > 0 Then
        query = Replace(Replace(Replace(query, "~", "~~"), "*", "~*"), "?", "~?")
        table.Range.AutoFilter field:=SEARCH_COL, Criteria1:="=*" & query & "*"
    End If
End Sub

Private Sub SetMembership(ByVal code As String, ByVal include As Boolean)
    Dim cfg As Worksheet, cr As Long, r As Long
    Set cfg = ThisWorkbook.Worksheets(CONFIG_NAME)
    cr = ConfigRow(activeProduct, code)
    If include Then
        If cr > 0 Then Exit Sub
        r = LastConfigRow() + 1
        cfg.cells(r, CONFIG_PRODUCT_COL).Value2 = activeProduct
        cfg.cells(r, CONFIG_CODE_COL).Value2 = code
        cfg.cells(r, CONFIG_MATERIAL_NAME_COL).FormulaR1C1 = CONFIG_FORMULA
        r = UsageRow(activeProduct, code, True)
    ElseIf cr > 0 Then
        cfg.Rows(cr).Delete
    End If
    BeginLookup
End Sub

Private Function ActiveMembershipChanged() As Boolean
    Dim previousMembers As Object, currentMembers As Object, ws As Worksheet, r As Long, code As String
    Set previousMembers = CreateObject("Scripting.Dictionary")
    previousMembers.CompareMode = vbTextCompare
    If IsArray(savedConfig) Then
        For r = 2 To UBound(savedConfig, 1)
            If Len(Trim$(CStr(savedConfig(r, CONFIG_PRODUCT_COL)))) > 0 And Len(Trim$(CStr(savedConfig(r, CONFIG_CODE_COL)))) > 0 Then
                If activeProduct = ALL_PRODUCT Or StrComp(Trim$(CStr(savedConfig(r, CONFIG_PRODUCT_COL))), activeProduct, vbTextCompare) = 0 Then _
                    previousMembers(Trim$(CStr(savedConfig(r, CONFIG_CODE_COL)))) = True
            End If
        Next r
    End If
    Set currentMembers = CreateObject("Scripting.Dictionary")
    currentMembers.CompareMode = vbTextCompare
    For r = 2 To UBound(configRows, 1)
        If Len(Trim$(CStr(configRows(r, CONFIG_PRODUCT_COL)))) > 0 And Len(Trim$(CStr(configRows(r, CONFIG_CODE_COL)))) > 0 Then
            If activeProduct = ALL_PRODUCT Or StrComp(Trim$(CStr(configRows(r, CONFIG_PRODUCT_COL))), activeProduct, vbTextCompare) = 0 Then
                code = Trim$(CStr(configRows(r, CONFIG_CODE_COL)))
                If Not previousMembers.Exists(code) Then
                    ActiveMembershipChanged = True
                    Exit Function
                End If
                currentMembers(code) = True
            End If
        End If
    Next r
    ActiveMembershipChanged = previousMembers.count <> currentMembers.count
End Function

Private Sub UpdatePrintAreaInternal()
    Dim lastPrint As Long, r As Long
    lastPrint = FIRST_VIEW - 1 + PrintItemCount()
    If activeProduct = ALL_PRODUCT Or printAllItems Then
        lastPrint = FIRST_VIEW - 1
        For r = FIRST_VIEW To LastViewRow()
            If IsRegisteredMaterial(r) Then lastPrint = r
        Next r
    End If
    With ThisWorkbook.Worksheets(VIEW_NAME)
        .PageSetup.PrintArea = .Range(.cells(1, 1), .cells(lastPrint, PRINT_LAST_COL)).Address
    End With
End Sub

Public Function PrintItemCount() As Long
    If activeProduct = ALL_PRODUCT Or printAllItems Then
        PrintItemCount = MaterialCount()
    Else
        PrintItemCount = ProductCount(activeProduct)
    End If
End Function

Private Function MaterialCatalogChanged() As Boolean
    Dim code As Variant
    If savedMaterials Is Nothing Then Exit Function
    If materialIndex.count <> savedMaterials.count Then
        MaterialCatalogChanged = True
        Exit Function
    End If
    For Each code In materialIndex.keys
        If Not savedMaterials.Exists(CStr(code)) Then
            MaterialCatalogChanged = True
            Exit Function
        End If
    Next code
End Function

Public Sub TogglePrintAllItems()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(VIEW_NAME)
    HandleIntegratedChange ws, ws.Range(PRINT_ALL_CELL), True
End Sub

Public Function SetPrintAllItems(ByVal includeAll As Boolean, Optional ByVal showMessage As Boolean = True) As Boolean
    Dim ws As Worksheet, oldEvents As Boolean
    Set ws = ThisWorkbook.Worksheets(VIEW_NAME)
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    ws.Range(PRINT_ALL_CELL).Value2 = includeAll
    Application.EnableEvents = oldEvents
    SetPrintAllItems = HandleIntegratedChange(ws, ws.Range(PRINT_ALL_CELL), showMessage)
    Exit Function
Failed:
    LastError = Err.description
    Application.EnableEvents = oldEvents
End Function

Private Sub AddHistory(ByVal action As String, ByVal batch As String, ByVal product As String, _
    ByVal code As String, ByVal opening As Double, ByVal incomingQty As Double, ByVal used As Double, ByVal closing As Double)
    Dim log As Worksheet, r As Long, idx As Long
    Set log = ThisWorkbook.Worksheets(LOG_NAME)
    r = LastRow(log) + 1
    If r < 6 Then r = 6
    idx = MaterialRow(code)
    log.cells(r, LOG_TIMESTAMP_COL).Resize(1, 10).value = Array(Now, action, batch, product, code, _
        ThisWorkbook.Worksheets(VIEW_NAME).cells(idx, MATERIAL_NAME_COL).Value2, opening, incomingQty, used, closing)
    log.cells(r, LOG_TIMESTAMP_COL).numberFormat = "yyyy-mm-dd hh:mm:ss"
End Sub

Private Sub NormalizeListInputs(ByVal ws As Worksheet, ByVal Target As Range)
    Dim area As Range, cell As Range, text As String, normalized As String, lastCol As Long, upper As Long
    If ws.name = CONFIG_NAME Then lastCol = CONFIG_CODE_COL Else lastCol = PRODUCT_LIST_NAME_COL
    upper = ws.UsedRange.Row + ws.UsedRange.Rows.count - 1
    If upper < 2 Then Exit Sub
    Set area = Intersect(Target, ws.Range(ws.cells(2, 1), ws.cells(upper, lastCol)))
    If area Is Nothing Then Exit Sub
    For Each cell In area.cells
        If Not IsError(cell.Value2) And Not cell.HasFormula Then
            text = CStr(cell.Value2)
            normalized = Trim$(text)
            If ws.name = CONFIG_NAME And Len(normalized) > 0 Then
                If cell.Column = CONFIG_CODE_COL Then
                    If materialIndex.Exists(normalized) Then normalized = Trim$(CStr(viewRows(materialIndex(normalized) - FIRST_VIEW + 1, CODE_COL)))
                ElseIf productIndex.Exists(normalized) Then
                    normalized = Trim$(CStr(productRows(productIndex(normalized), PRODUCT_LIST_NAME_COL)))
                End If
            End If
            If normalized <> text Then
                If IsNumeric(normalized) Then cell.numberFormat = "@"
                cell.Value2 = normalized
            End If
        End If
    Next cell
End Sub

Private Function ConfigRowCount(ByVal values As Variant, ByVal product As String) As Long
    Dim r As Long
    For r = 2 To UBound(values, 1)
        If StrComp(Trim$(CStr(values(r, CONFIG_PRODUCT_COL))), product, vbTextCompare) = 0 Then ConfigRowCount = ConfigRowCount + 1
    Next r
End Function

Private Sub ApplyProductChanges()
    Dim previous As Object, current As Object, removed As Object, added As Object, key As Variant
    Dim cfg As Worksheet, uses As Worksheet, values As Variant, r As Long, count As Long
    Dim name As String, oldName As String, newName As String
    If Not IsArray(savedProducts) Then Exit Sub
    Set previous = NewTextDictionary()
    Set current = NewTextDictionary()
    For r = 2 To UBound(savedProducts, 1)
        name = Trim$(CStr(savedProducts(r, PRODUCT_LIST_NAME_COL)))
        If Len(name) > 0 Then previous(name) = True
    Next r
    For r = 2 To UBound(productRows, 1)
        name = Trim$(CStr(productRows(r, PRODUCT_LIST_NAME_COL)))
        If Len(name) > 0 Then
            If name = ALL_PRODUCT Then Err.Raise vbObjectError + 35, , "전체는 합산 조회용 이름입니다. 다른 제품명을 입력하세요."
            If current.Exists(name) Then Err.Raise vbObjectError + 6, , "제품명이 중복됩니다: " & name
            current(name) = True
        End If
    Next r
    Set removed = NewTextDictionary()
    Set added = NewTextDictionary()
    For Each key In previous.keys
        If Not current.Exists(key) Then removed(key) = True
    Next key
    If removed.count = 0 Then Exit Sub
    For Each key In current.keys
        If Not previous.Exists(key) Then added(key) = True
    Next key
    Set cfg = ThisWorkbook.Worksheets(CONFIG_NAME)
    Set uses = ThisWorkbook.Worksheets(USAGE_NAME)
    If removed.count = 1 And added.count = 1 Then
        oldName = removed.keys()(0)
        newName = added.keys()(0)
        values = cfg.Range(CONFIG_DATA_RANGE_PREFIX & LastConfigRow()).Value2
        For r = 2 To UBound(values, 1)
            If StrComp(Trim$(CStr(values(r, CONFIG_PRODUCT_COL))), oldName, vbTextCompare) = 0 Then cfg.cells(r, CONFIG_PRODUCT_COL).Value2 = newName
        Next r
        values = uses.Range(USAGE_DATA_RANGE_PREFIX & lastUsageRow()).Value2
        For r = 2 To UBound(values, 1)
            If StrComp(Trim$(CStr(values(r, USAGE_PRODUCT_COL))), oldName, vbTextCompare) = 0 Then
                uses.cells(r, USAGE_PRODUCT_COL).Value2 = newName
                uses.cells(r, USAGE_KEY_COL).Value2 = PairKey(newName, CStr(values(r, USAGE_CODE_COL)))
            End If
        Next r
        If StrComp(activeProduct, oldName, vbTextCompare) = 0 Then activeProduct = newName
    ElseIf added.count = 0 Then
        For Each key In removed.keys
            If StrComp(activeProduct, CStr(key), vbTextCompare) = 0 Then _
                Err.Raise vbObjectError + 57, , "선택 중인 제품은 지울 수 없습니다. 다른 제품을 선택한 뒤 지우세요: " & key
            count = ConfigRowCount(cfg.Range(CONFIG_DATA_RANGE_PREFIX & LastConfigRow()).Value2, CStr(key))
            If count > 0 Then Err.Raise vbObjectError + 58, , "제품구성에 남은 " & count & "행을 먼저 정리한 뒤 제품을 지우세요: " & key
            values = uses.Range(USAGE_DATA_RANGE_PREFIX & lastUsageRow()).Value2
            For r = UBound(values, 1) To 2 Step -1
                If StrComp(Trim$(CStr(values(r, USAGE_PRODUCT_COL))), CStr(key), vbTextCompare) = 0 Then
                    If ReadNonnegativeQuantity(values(r, USAGE_QUANTITY_COL), "금일사용") > 0 Then _
                        Err.Raise vbObjectError + 59, , "이월하지 않은 사용량이 있어 제품을 지울 수 없습니다: " & key
                    uses.cells(r, USAGE_KEY_COL).Resize(1, 4).Delete Shift:=xlUp
                End If
            Next r
        Next key
    Else
        Err.Raise vbObjectError + 60, , "제품명은 한 번에 하나씩 바꾸세요."
    End If
End Sub

Public Function HandleIntegratedChange(ByVal ws As Worksheet, ByVal Target As Range, Optional ByVal showMessage As Boolean = True) As Boolean
    Dim oldEvents As Boolean, cell As Range, changed As Range, code As String, idx As Long, cr As Long
    Dim n As Double, oldStock As Double, member As String, oldKey As String, oldValues As Variant
    Dim usageBefore As Variant, usageValuesBefore As Variant, logBefore As Long, previousProduct As String, previousPrintArea As String
    Dim printAreaChanged As Boolean, previousPrintAll As Boolean
    Dim membershipEdited As Boolean, catalogEdited As Boolean, storageSnapshotTaken As Boolean
    Dim oldScreen As Boolean, affectedRows As Object, rowKey As Variant, fullRefresh As Boolean
    Dim configTouched As Boolean, catalogTouched As Boolean, masterTouched As Boolean, repairFormulas As Boolean
    Dim db As Worksheet, uses As Worksheet, log As Worksheet, table As ListObject
    Dim zeroRows As Object, deletedCodes As Object, journal As Collection, catalogTransaction As Boolean
    Dim materialRowsDeleted As Boolean, rollbackError As String, materialBefore As Variant
    Dim oldFill As Boolean, catalogApplied As Boolean, recordBackup As Workbook
    Dim materialsCheckpoint As Object, draftsCheckpoint As Object, viewIndexCheckpoint As Object
    Dim configCheckpoint As Variant, configValuesCheckpoint As Variant, productsCheckpoint As Variant
    Dim recovering As Boolean, recovered As Boolean
    Set journal = New Collection
    Set deletedCodes = NewTextDictionary()
    If Busy Then Exit Function
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldFill = Application.AutoCorrect.AutoFillFormulasInLists
    On Error GoTo Failed
    previousProduct = activeProduct
    previousPrintAll = printAllItems
    Application.EnableEvents = False
    Busy = True
    Application.ScreenUpdating = False
    Application.AutoCorrect.AutoFillFormulasInLists = False
    If ws.name <> VIEW_NAME And ws.name <> CONFIG_NAME And ws.name <> PRODUCT_NAME Then GoTo Succeeded
    EnsureMacroAccess
    If ws.name = VIEW_NAME Then
        If Not Intersect(Target, ws.Range(CAPACITY_CELL)) Is Nothing Then
            If Intersect(Target, ws.Range(CAPACITY_CELL)).CountLarge = Target.CountLarge Then GoTo Succeeded
        End If
        If Not Intersect(Target, Union(ws.Range(MODE_RANGE), ws.Range(QUERY_RANGE))) Is Nothing Then
            ws.Calculate
            ApplyFiltersInternal
            GoTo Succeeded
        End If
    End If
    If StateMissing() Then
        recovering = True
        RecoverState ws, Target
        recovering = False
        recovered = True
        previousProduct = activeProduct
        previousPrintAll = printAllItems
    End If
    If ws.name = VIEW_NAME Then
        If Not Intersect(Target, ws.Range(PRINT_ALL_CELL)) Is Nothing Then
            printAllItems = CBool(ws.Range(PRINT_ALL_CELL).Value2)
            If (previousPrintAll <> printAllItems Or recovered) And activeProduct <> ALL_PRODUCT Then
                BeginLookup
                ValidateStorage
                previousPrintArea = ws.PageSetup.PrintArea
                SortViewInternal
                printAreaChanged = True
                UpdatePrintAreaInternal
                SaveStorageState
            End If
            GoTo Succeeded
        End If
        If Not Intersect(Target, ws.Range(PRODUCT_RANGE)) Is Nothing Then
            If Not recovered And StrComp(CStr(ws.Range(PRODUCT_CELL).Value2), activeProduct, vbTextCompare) = 0 Then GoTo Succeeded
            BeginLookup
            ValidateStorage
            If CStr(ws.Range(PRODUCT_CELL).Value2) = ALL_PRODUCT Then
                activeProduct = ALL_PRODUCT
            Else
                code = Trim$(CStr(ws.Range(PRODUCT_CELL).Value2))
                If Not productIndex.Exists(code) Then Err.Raise vbObjectError + 23, , "등록된 제품을 선택하세요."
                activeProduct = Trim$(CStr(productRows(productIndex(code), PRODUCT_LIST_NAME_COL)))
            End If
            ws.Range(PRODUCT_CELL).Value2 = activeProduct
            previousPrintArea = ws.PageSetup.PrintArea
            ClearFiltersInternal
            RefreshView False
            SortViewInternal
            printAreaChanged = True
            UpdatePrintAreaInternal
            SaveStorageState
            GoTo Succeeded
        End If
    End If
    Set db = ThisWorkbook.Worksheets(VIEW_NAME)
    previousPrintArea = db.PageSetup.PrintArea
    Set uses = ThisWorkbook.Worksheets(USAGE_NAME)
    usageBefore = uses.Range(USAGE_DATA_RANGE_PREFIX & lastUsageRow()).formula
    usageValuesBefore = uses.Range(USAGE_DATA_RANGE_PREFIX & lastUsageRow()).Value2
    Set log = ThisWorkbook.Worksheets(LOG_NAME)
    logBefore = LastRow(log)
    If ws.name = VIEW_NAME And IsArray(savedView) Then
        materialRowsDeleted = ws.ListObjects(1).ListRows.count < UBound(savedView, 1)
        catalogTransaction = materialRowsDeleted Or Not Intersect(Target, Union(ws.Columns(MATERIAL_NAME_COL), ws.Columns(CODE_COL), ws.Columns(KEY_COL))) Is Nothing
        If catalogTransaction Then
            materialBefore = savedView
            Set materialsCheckpoint = savedMaterials
            Set draftsCheckpoint = savedDrafts
            Set viewIndexCheckpoint = savedViewIndex
            configCheckpoint = savedConfig
            configValuesCheckpoint = savedConfigValues
            productsCheckpoint = savedProducts
        End If
    End If
    storageSnapshotTaken = True
    ws.Calculate
    BeginLookup
    If ws.name = VIEW_NAME Then
            Set changed = Intersect(Target, ws.Range(VIEW_DATA_RANGE_PREFIX & ws.Rows.count))
            If Not changed Is Nothing Then
                Set affectedRows = NewTextDictionary()
                catalogTouched = materialRowsDeleted Or Not Intersect(changed, Union(ws.Columns(MATERIAL_NAME_COL), ws.Columns(CODE_COL), ws.Columns(KEY_COL))) Is Nothing
                configTouched = Not materialRowsDeleted And Not Intersect(changed, ws.Columns(MEMBER_COL)) Is Nothing
                masterTouched = Not materialRowsDeleted And Not Intersect(changed, Union(ws.Columns(STOCK_COL), ws.Columns(UNIT_COL), ws.Columns(REPLENISHMENT_STEP_COL), ws.Columns(TARE_COL))) Is Nothing
                repairFormulas = Not materialRowsDeleted And Not Intersect(changed, Union(ws.Columns(INCOMING_COL), ws.Columns(CLOSING_COL), ws.Columns(RECEIPT_COL), ws.Columns(SEARCH_COL))) Is Nothing
                Set table = ws.ListObjects(1)
                fullRefresh = catalogTouched Or configTouched Or changed.Row + changed.Rows.count - 1 > table.Range.Row + table.Range.Rows.count - 1
                If materialRowsDeleted Then
                    ' Native deletion targets shifted cells and a blank tail, not new input.
                    For idx = FIRST_VIEW To LastViewRow()
                        affectedRows(CStr(idx)) = True
                    Next idx
                Else
                For Each cell In changed.cells
                    affectedRows(CStr(cell.Row)) = True
                    oldKey = CStr(ws.cells(cell.Row, KEY_COL).Value2)
                    code = Trim$(CStr(ws.cells(cell.Row, CODE_COL).Value2))
                    If catalogTouched And Not IsRegisteredMaterial(cell.Row) Then
                        If (cell.Column = USE_COL Or cell.Column = MEMBER_COL) And Len(CStr(cell.Value2)) = 0 Then GoTo NextChangedCell
                    End If
                    Select Case cell.Column
                        Case STOCK_COL, UNIT_COL, REPLENISHMENT_STEP_COL
                            n = ReadNonnegativeQuantity(cell.Value2, "자재 수량")
                        Case USE_COL
                            n = ReadNonnegativeQuantity(cell.Value2, "금일사용")
                            If activeProduct = ALL_PRODUCT And (Len(oldKey) > 0 Or Len(CStr(cell.Value2)) > 0) Then _
                                Err.Raise vbObjectError + 36, , "금일사용을 입력할 제품을 먼저 선택하세요. 전체는 합산 조회용입니다."
                            If n > 0 Then
                                If Not IsRegisteredMaterial(cell.Row) Then Err.Raise vbObjectError + 41, , "품명과 자재코드를 먼저 입력하세요."
                            End If
                        Case MEMBER_COL
                            If activeProduct = ALL_PRODUCT Then Err.Raise vbObjectError + 37, , "구성을 변경할 제품을 먼저 선택하세요. 전체는 포함 여부 조회용입니다."
                            If Not IsRegisteredMaterial(cell.Row) Then Err.Raise vbObjectError + 41, , "품명과 자재코드를 먼저 입력하세요."
                            member = CStr(cell.Value2)
                            If member <> "포함" And member <> "제외" Then Err.Raise vbObjectError + 26, , "제품포함은 포함 또는 제외로 선택하세요."
                    End Select
NextChangedCell:
                Next cell
                End If
                If catalogTransaction Then
                    Set zeroRows = MaterialChangeRows(deletedCodes, Target, materialRowsDeleted)
                    If zeroRows.count > 0 Or deletedCodes.count > 0 Then
                        ValidateZeroRemoval deletedCodes, zeroRows
                        RemoveMaterialRecords deletedCodes, zeroRows, journal, recordBackup
                        BeginLookup
                    End If
                End If
                ValidateStorage
                If Not materialRowsDeleted Then
                For Each cell In changed.cells
                    If IsRegisteredMaterial(cell.Row) Then
                        oldKey = CStr(ws.cells(cell.Row, KEY_COL).Value2)
                        code = CStr(ws.cells(cell.Row, CODE_COL).Value2)
                        Select Case cell.Column
                            Case STOCK_COL
                                If savedMaterials.Exists(oldKey) Then
                                    oldValues = savedMaterials(oldKey)
                                    oldStock = ReadNonnegativeQuantity(oldValues(5))
                                    n = ReadNonnegativeQuantity(cell.Value2)
                                    If Abs(n - oldStock) > 0.0000001 Then AddHistory "재고수정", "", activeProduct, code, oldStock, 0, 0, n
                                End If
                            Case USE_COL
                                If activeProduct <> ALL_PRODUCT Then
                                    n = ReadNonnegativeQuantity(cell.Value2)
                                    If n > 0 Or ConfigRow(activeProduct, code) > 0 Or UsageRow(activeProduct, code) > 0 Then
                                        idx = UsageRow(activeProduct, code, True)
                                        uses.cells(idx, USAGE_QUANTITY_COL).Value2 = n
                                    End If
                                End If
                            Case MEMBER_COL
                                SetMembership code, CStr(cell.Value2) = "포함"
                        End Select
                    End If
                Next cell
                End If
                BeginLookup
                If fullRefresh Then ValidateStorage
                If catalogTouched Or configTouched Then UpdateNames False, False
                If fullRefresh Then
                    If catalogTouched Or repairFormulas Then
                        For Each rowKey In affectedRows.keys
                            SetViewFormulas CLng(rowKey), CLng(rowKey)
                        Next rowKey
                    End If
                    catalogApplied = catalogTransaction
                    RefreshView False
                    For Each rowKey In affectedRows.keys
                        FillViewRow CLng(rowKey), False
                    Next rowKey
                Else
                    If affectedRows.count >= 32 And Not repairFormulas Then
                        RefreshView False
                    Else
                        For Each rowKey In affectedRows.keys
                            FillViewRow CLng(rowKey), repairFormulas
                        Next rowKey
                        ws.Calculate
                        If ws.FilterMode Then table.AutoFilter.ApplyFilter
                    End If
                End If
            End If
    ElseIf ws.name = CONFIG_NAME Or ws.name = PRODUCT_NAME Then
        NormalizeListInputs ws, Target
        BeginLookup
        If ws.name = PRODUCT_NAME Then
            ApplyProductChanges
            BeginLookup
        End If
        ValidateStorage
        UpdateNames ws.name = PRODUCT_NAME, ws.name = CONFIG_NAME, Target
        If Len(activeProduct) > 0 And CStr(db.Range(PRODUCT_CELL).Value2) <> activeProduct Then db.Range(PRODUCT_CELL).Value2 = activeProduct
        RefreshView False
    End If
    If configTouched Or deletedCodes.count > 0 Or ws.name = CONFIG_NAME Then membershipEdited = ActiveMembershipChanged()
    If catalogTouched Or fullRefresh Then catalogEdited = MaterialCatalogChanged()
    If membershipEdited And deletedCodes.count = 0 Then
        SortViewInternal
    End If
    If membershipEdited Or _
       ((activeProduct = ALL_PRODUCT Or printAllItems) And catalogEdited) Then
        printAreaChanged = True
        UpdatePrintAreaInternal
    End If
    If ws.name <> VIEW_NAME Or fullRefresh Then
        RepairInputRules
        SaveStorageState
        ThisWorkbook.Worksheets(CONFIG_NAME).Calculate
    ElseIf masterTouched Then
        SaveStorageState affectedRows
    ElseIf Not affectedRows Is Nothing Then
        SaveViewRecovery affectedRows
    End If
Succeeded:
    LastError = ""
    HandleIntegratedChange = True
    GoTo Finish
Failed:
    LastError = Err.description
    Resume Rollback
Rollback:
    On Error GoTo RollbackFailed
    EndLookup
    If recovering Then GoTo ShowFailure
    If Not storageSnapshotTaken Then
        activeProduct = previousProduct
        printAllItems = previousPrintAll
        With ThisWorkbook.Worksheets(VIEW_NAME)
            If Len(activeProduct) > 0 Then .Range(PRODUCT_CELL).Value2 = activeProduct
            .Range(PRINT_ALL_CELL).Value2 = printAllItems
        End With
        BeginLookup
        RefreshView False
    ElseIf catalogTransaction Then
        Set savedMaterials = materialsCheckpoint
        Set savedDrafts = draftsCheckpoint
        Set savedViewIndex = viewIndexCheckpoint
        savedView = materialBefore
        savedConfig = configCheckpoint
        savedConfigValues = configValuesCheckpoint
        savedProducts = productsCheckpoint
        RestoreRemovedRecords journal, recordBackup
        RestoreCatalogInputs Target, materialBefore, materialRowsDeleted
        db.Calculate
        If LastRow(log) > logBefore Then log.Range("A" & logBefore + 1 & ":J" & LastRow(log)).ClearContents
        BeginLookup
        UpdateNames False, False
        If catalogApplied Or journal.count > 0 Then RefreshView False
        db.PageSetup.PrintArea = previousPrintArea
        ThisWorkbook.Worksheets(CONFIG_NAME).Calculate
        SaveStorageState
        If materialRowsDeleted Then LastError = LastError & " / 직접 삭제 오류: 완전 복구를 보장하지 않습니다. 저장 전에 확인하세요."
    ElseIf ws.name = VIEW_NAME Then
        RestoreMaterialState
        RestoreStorageState CONFIG_NAME
        RestoreStorageState PRODUCT_NAME
        With uses
            .Range(USAGE_DATA_RANGE_PREFIX & lastUsageRow()).ClearContents
            .Range("A1").Resize(UBound(usageBefore, 1), UBound(usageBefore, 2)).formula = usageBefore
            RestoreNumericText .Range("A1").Resize(UBound(usageBefore, 1), 4), usageValuesBefore, usageBefore
        End With
        With log
            If LastRow(log) > logBefore Then .Range("A" & logBefore + 1 & ":J" & LastRow(log)).ClearContents
        End With
        activeProduct = previousProduct
        printAllItems = previousPrintAll
        If Len(activeProduct) > 0 Then ws.Range(PRODUCT_CELL).Value2 = activeProduct
        ws.Range(PRINT_ALL_CELL).Value2 = printAllItems
        BeginLookup
        UpdateNames
        SetViewFormulas FIRST_VIEW, viewLast
        RefreshView False
        If Not affectedRows Is Nothing Then
            For Each rowKey In affectedRows.keys
                FillViewRow CLng(rowKey)
            Next rowKey
        End If
    Else
        RestoreStorageState CONFIG_NAME
        RestoreStorageState PRODUCT_NAME
        With uses
            .Range(USAGE_DATA_RANGE_PREFIX & lastUsageRow()).ClearContents
            .Range("A1").Resize(UBound(usageBefore, 1), UBound(usageBefore, 2)).formula = usageBefore
            RestoreNumericText .Range("A1").Resize(UBound(usageBefore, 1), 4), usageValuesBefore, usageBefore
        End With
        activeProduct = previousProduct
        If Len(activeProduct) > 0 Then db.Range(PRODUCT_CELL).Value2 = activeProduct
        BeginLookup
        UpdateNames
        RefreshView False
    End If
    If printAreaChanged Then
        If db Is Nothing Then Set db = ThisWorkbook.Worksheets(VIEW_NAME)
        db.PageSetup.PrintArea = previousPrintArea
    End If
    On Error GoTo 0
    GoTo ShowFailure
RollbackFailed:
    rollbackError = Err.description
    LastError = LastError & " / 복구 실패: " & rollbackError
ShowFailure:
    If showMessage Then MsgBox LastError, vbExclamation, "입력 확인"
Finish:
    On Error Resume Next
    If Not recordBackup Is Nothing Then recordBackup.Close SaveChanges:=False
    On Error GoTo 0
    EndLookup
    Busy = False
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.AutoCorrect.AutoFillFormulasInLists = oldFill
End Function

Public Function SwitchProduct(ByVal product As String, Optional ByVal showMessage As Boolean = True) As Boolean
    Dim ws As Worksheet, oldEvents As Boolean
    Set ws = ThisWorkbook.Worksheets(VIEW_NAME)
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    ws.Range(PRODUCT_CELL).Value2 = product
    Application.EnableEvents = oldEvents
    SwitchProduct = HandleIntegratedChange(ws, ws.Range(PRODUCT_CELL), showMessage)
    Exit Function
Failed:
    LastError = Err.description
    Application.EnableEvents = oldEvents
End Function

Public Sub ApplySelectedCapacityUsage()
    Dim ok As Boolean
    If StateMissing() Then InitializeIntegrated True
    If StateMissing() Then Exit Sub
    ok = ApplyCapacityUsage()
End Sub

Public Function ApplyCapacityUsage(Optional ByVal showMessage As Boolean = True) As Boolean
    Dim db As Worksheet, uses As Worksheet
    Dim oldEvents As Boolean, oldScreen As Boolean, writesStarted As Boolean
    Dim capacityIndex As Double, capacityColumn As Long, r As Long
    Dim itemIndex As Long, itemCount As Long, usageRowNumber As Long
    Dim code As String, key As String, product As String, capacityOptions As Variant
    Dim pendingUsage() As Variant, newUsageRows() As Variant, usageSnapshot As Variant
    Dim lastUsageRow As Long, missingUsageCount As Long, newUsageIndex As Long
    Dim usageChanged As Boolean, hasExistingUsage As Boolean
    If Busy Then Exit Function
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Busy = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set db = ThisWorkbook.Worksheets(VIEW_NAME)
    product = activeProduct
    If product = ALL_PRODUCT Or Len(product) = 0 Then Err.Raise vbObjectError + 45, , "개별 제품을 먼저 선택하세요."
    If StrComp(CStr(db.Range(PRODUCT_CELL).Value2), product, vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 46, , "선택 제품을 다시 확인하세요."
    capacityIndex = ReadNonnegativeQuantity(db.Range(CAPACITY_CELL).Value2, "용량 선택")
    capacityOptions = Array(8000, 12000, 15000, 20000)
    If capacityIndex < 1 Or capacityIndex > UBound(capacityOptions) + 1 Or capacityIndex <> Fix(capacityIndex) Then _
        Err.Raise vbObjectError + 47, , "8000, 12000, 15000, 20000 중 용량을 선택하세요."
    capacityColumn = CONFIG_FIRST_CAPACITY_COL + CLng(capacityIndex) - 1
    EnsureMacroAccess
    BeginLookup
    ValidateStorage
    If Not productIndex.Exists(product) Then Err.Raise vbObjectError + 23, , "등록된 제품을 선택하세요."
    itemCount = ProductCount(product)
    If itemCount = 0 Then Err.Raise vbObjectError + 48, , "선택 제품에 구성 품목이 없습니다."
    Set uses = ThisWorkbook.Worksheets(USAGE_NAME)
    ReDim pendingUsage(1 To itemCount, 1 To 4)
    For r = 2 To UBound(configRows, 1)
        code = Trim$(CStr(configRows(r, CONFIG_CODE_COL)))
        If StrComp(Trim$(CStr(configRows(r, CONFIG_PRODUCT_COL))), product, vbTextCompare) = 0 And Len(code) > 0 Then
            itemIndex = itemIndex + 1
            key = PairKey(product, code)
            usageRowNumber = 0
            If usageIndex.Exists(key) Then usageRowNumber = usageIndex(key)
            pendingUsage(itemIndex, 1) = code
            pendingUsage(itemIndex, 2) = ReadNonnegativeQuantity(configRows(r, capacityColumn), "용량별 사용량")
            pendingUsage(itemIndex, 3) = usageRowNumber
            pendingUsage(itemIndex, 4) = ProductUsage(product, code)
            If usageRowNumber = 0 Then missingUsageCount = missingUsageCount + 1
            If pendingUsage(itemIndex, 4) > 0 Then hasExistingUsage = True
            If pendingUsage(itemIndex, 2) <> pendingUsage(itemIndex, 4) Then usageChanged = True
        End If
    Next r
    If Not usageChanged Then GoTo Succeeded
    If showMessage And hasExistingUsage Then
        If MsgBox(product & " / " & capacityOptions(CLng(capacityIndex) - 1) & " / " & itemCount & "품목" & vbCrLf & _
                  "빈칸은 0으로 입력되며 현재 금일사용량을 교체합니다. 계속할까요?", _
                  vbOKCancel + vbExclamation + vbDefaultButton2, "용량 적용") <> vbOK Then
            LastError = ""
            GoTo Finish
        End If
    End If
    lastUsageRow = LastRow(uses)
    usageSnapshot = uses.Range(USAGE_DATA_RANGE_PREFIX & lastUsageRow).formula
    If missingUsageCount > 0 Then ReDim newUsageRows(1 To missingUsageCount, 1 To 4)
    writesStarted = True
    For itemIndex = 1 To itemCount
        usageRowNumber = CLng(pendingUsage(itemIndex, 3))
        If usageRowNumber = 0 Then
            newUsageIndex = newUsageIndex + 1
            newUsageRows(newUsageIndex, 1) = PairKey(product, CStr(pendingUsage(itemIndex, 1)))
            newUsageRows(newUsageIndex, 2) = product
            newUsageRows(newUsageIndex, 3) = pendingUsage(itemIndex, 1)
            newUsageRows(newUsageIndex, 4) = pendingUsage(itemIndex, 2)
        ElseIf pendingUsage(itemIndex, 2) <> pendingUsage(itemIndex, 4) Then
            uses.cells(usageRowNumber, USAGE_QUANTITY_COL).Value2 = pendingUsage(itemIndex, 2)
        End If
    Next itemIndex
    If missingUsageCount > 0 Then uses.cells(lastUsageRow + 1, USAGE_KEY_COL).Resize(missingUsageCount, 4).Value2 = newUsageRows
    BeginLookup
    RefreshView False
    SaveViewRecovery
Succeeded:
    LastError = ""
    ApplyCapacityUsage = True
    GoTo Finish
Failed:
    LastError = Err.description
    If writesStarted Then
        On Error Resume Next
        EndLookup
        uses.Range(USAGE_DATA_RANGE_PREFIX & LastRow(uses)).ClearContents
        uses.Range("A1").Resize(UBound(usageSnapshot, 1), UBound(usageSnapshot, 2)).formula = usageSnapshot
        BeginLookup
        RefreshView False
        On Error GoTo 0
    End If
    If showMessage Then MsgBox "입력하지 않았습니다. " & LastError, vbExclamation, "용량 입력 확인"
Finish:
    EndLookup
    Busy = False
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
End Function

Public Sub RollOverSelectedProduct()
    If StateMissing() Then InitializeIntegrated True
    If StateMissing() Then Exit Sub
    RollOverProduct activeProduct, True
End Sub

Public Function RollOverProduct(ByVal product As String, Optional ByVal showMessage As Boolean = True) As Boolean
    Dim db As Worksheet, cfg As Worksheet, uses As Worksheet, log As Worksheet
    Dim r As Long, materialRowNumber As Long, usageRowNumber As Long, count As Long, i As Long, lastLog As Long
    Dim code As String, opening As Double, used As Double, receipt As Double
    Dim pendingRollover() As Variant, beforeStock As Variant, beforeUsage As Variant, historyRows() As Variant, stamp As Date
    Dim updatedStock As Variant, updatedUsage As Variant
    Dim oldEvents As Boolean, oldScreen As Boolean, writesStarted As Boolean, batch As String
    If Busy Then Exit Function
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    EnsureMacroAccess
    Set db = ThisWorkbook.Worksheets(VIEW_NAME)
    db.Calculate
    BeginLookup
    ValidateStorage
    If product = ALL_PRODUCT Then Err.Raise vbObjectError + 38, , "이월할 제품을 먼저 선택하세요. 전체에서는 이월하지 않습니다."
    If FindRow(ThisWorkbook.Worksheets(PRODUCT_NAME), product) = 0 Then Err.Raise vbObjectError + 29, , "제품을 확인하세요."
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 30, , "읽기 전용 파일에서는 이월할 수 없습니다."
    Set cfg = ThisWorkbook.Worksheets(CONFIG_NAME)
    Set uses = ThisWorkbook.Worksheets(USAGE_NAME)
    Set log = ThisWorkbook.Worksheets(LOG_NAME)
    ReDim pendingRollover(1 To WorksheetFunction.Max(1, MaterialCount()), 1 To 7)
    For r = 2 To UBound(useRows, 1)
        If StrComp(Trim$(CStr(useRows(r, USAGE_PRODUCT_COL))), product, vbTextCompare) = 0 Then
            code = Trim$(CStr(useRows(r, USAGE_CODE_COL)))
            used = ReadNonnegativeQuantity(useRows(r, USAGE_QUANTITY_COL))
            If used > 0 Then
                materialRowNumber = MaterialRow(code)
                usageRowNumber = r
                opening = ReadNonnegativeQuantity(viewRows(materialRowNumber - FIRST_VIEW + 1, STOCK_COL))
                receipt = CalculateReplenishment(opening, used, ReadNonnegativeQuantity(viewRows(materialRowNumber - FIRST_VIEW + 1, REPLENISHMENT_STEP_COL)))
                count = count + 1
                pendingRollover(count, 1) = materialRowNumber
                pendingRollover(count, 2) = usageRowNumber
                pendingRollover(count, 3) = code
                pendingRollover(count, 4) = opening
                pendingRollover(count, 5) = receipt
                pendingRollover(count, 6) = used
                pendingRollover(count, 7) = WorksheetFunction.Round(opening + receipt - used, 6)
            End If
        End If
    Next r
    If count = 0 Then
        RollOverProduct = True
        If showMessage Then MsgBox "이월할 사용량이 없습니다.", vbInformation, "재고 이월"
        GoTo Finish
    End If
    beforeStock = ReadColumnArray(db.Range(VIEW_STOCK_RANGE_PREFIX & LastViewRow()), True)
    beforeUsage = ReadColumnArray(uses.Range(USAGE_QUANTITY_RANGE_PREFIX & LastRow(uses)), True)
    updatedStock = beforeStock
    updatedUsage = beforeUsage
    lastLog = LastRow(log)
    Busy = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    writesStarted = True
    batch = Format$(Now, "yyyymmdd-hhnnss") & "-" & product & "-" & Format$(Timer * 100, "0")
    ReDim historyRows(1 To count, 1 To 10)
    stamp = Now
    For i = 1 To count
        updatedStock(pendingRollover(i, 1) - FIRST_VIEW + 1, 1) = pendingRollover(i, 7)
        updatedUsage(pendingRollover(i, 2) - 1, 1) = 0
        historyRows(i, 1) = stamp
        historyRows(i, 2) = "제품이월"
        historyRows(i, 3) = batch
        historyRows(i, 4) = product
        historyRows(i, 5) = pendingRollover(i, 3)
        historyRows(i, 6) = viewRows(pendingRollover(i, 1) - FIRST_VIEW + 1, MATERIAL_NAME_COL)
        historyRows(i, 7) = pendingRollover(i, 4)
        historyRows(i, 8) = pendingRollover(i, 5)
        historyRows(i, 9) = pendingRollover(i, 6)
        historyRows(i, 10) = pendingRollover(i, 7)
    Next i
    If db.FilterMode Then
        For i = 1 To count
            db.cells(pendingRollover(i, 1), STOCK_COL).Value2 = pendingRollover(i, 7)
        Next i
    Else
        db.Range(VIEW_STOCK_RANGE_PREFIX & LastViewRow()).formula = updatedStock
    End If
    uses.Range(USAGE_QUANTITY_RANGE_PREFIX & LastRow(uses)).formula = updatedUsage
    log.cells(lastLog + 1, LOG_TIMESTAMP_COL).Resize(count, 10).value = historyRows
    log.cells(lastLog + 1, LOG_TIMESTAMP_COL).Resize(count, 1).numberFormat = "yyyy-mm-dd hh:mm:ss"
    BeginLookup
    RefreshView False
    SaveStorageState
    cfg.Calculate
    db.Calculate
    ThisWorkbook.Save
    RollOverProduct = True
    LastError = ""
    If showMessage Then MsgBox product & "의 사용량을 이월했습니다.", vbInformation, "재고 이월 완료"
    GoTo Finish
Failed:
    LastError = Err.description
    If writesStarted Then
        On Error Resume Next
        If db.FilterMode Then
            For i = 1 To count
                db.cells(pendingRollover(i, 1), STOCK_COL).formula = beforeStock(pendingRollover(i, 1) - FIRST_VIEW + 1, 1)
            Next i
        Else
            db.Range(VIEW_STOCK_RANGE_PREFIX & LastViewRow()).formula = beforeStock
        End If
        uses.Range(USAGE_QUANTITY_RANGE_PREFIX & LastRow(uses)).formula = beforeUsage
        If LastRow(log) > lastLog Then log.Range("A" & lastLog + 1 & ":J" & LastRow(log)).ClearContents
        BeginLookup
        RefreshView False
        SaveStorageState
        On Error GoTo 0
    End If
    If showMessage Then MsgBox "이월하지 않았습니다. " & LastError, vbExclamation, "재고 이월 확인"
Finish:
    EndLookup
    Busy = False
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
End Function


