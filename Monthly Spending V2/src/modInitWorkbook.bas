Attribute VB_Name = "modInitWorkbook"
Option Explicit

Private mWorkbook As Workbook
Private mConnection As ADODB.Connection

Private Sub AddColumnSumFormula(ws As Worksheet, lastRow As Long)

    On Error GoTo eh

    Dim rowNum As Integer
    Dim cellName As String
    Dim formulaValue As String
    
    For rowNum = 2 To lastRow
        cellName = "N" & rowNum
        formulaValue = "=SUM(B" & rowNum & ":M" & rowNum & ")"
        ws.Range(cellName).Formula = formulaValue
        ws.Range(cellName).Font.Bold = True
    Next
    
    Exit Sub
    
eh:
    RaiseError Err.Number, Err.Source, "modInitWorkbook.AddColumnSumFormula", Err.Description, 57
End Sub

Private Sub AddNumberFormat(ws As Worksheet, lastRow As Long)

    On Error GoTo eh
    
    Dim cellName As String
    
    cellName = "B2:N" & lastRow + 2
    ws.Range(cellName).NumberFormat = NUMBER_FORMAT_ACCOUNTING
    
    Exit Sub
    
eh:
    RaiseError Err.Number, Err.Source, "modInitWorkbook.AddNumberFormat", Err.Description, 78
End Sub

Private Sub AddRowSumFormula(ws As Worksheet, lastRow As Long)

    On Error GoTo eh
    
    Dim colls
    Dim i As Integer
    Dim cellName As String
    Dim formulaValue As String
    
    colls = Array("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N")
    
    For i = 0 To UBound(colls)
        cellName = colls(i) & lastRow + 2
        formulaValue = "=SUM(" & colls(i) & "2:" & colls(i) & lastRow & ")"
'        formulaValue = "=SUM(" & colls(i) & lastRow & ":" & colls(i) & "2)"
        ws.Range(cellName).Formula = formulaValue
        ws.Range(cellName).Font.Bold = True
    Next

    Exit Sub
    
eh:
    RaiseError Err.Number, Err.Source, "modInitWorkbook.AddRowSumFormula", Err.Description, 97
End Sub

Private Sub AddWorksheets()

    On Error GoTo eh
    
    Dim worksheetNames As Variant
    worksheetNames = SelectWorksheetsArray
    
    Dim i As Integer
    For i = LBound(worksheetNames) To UBound(worksheetNames)
        mWorkbook.Sheets.Add.Name = worksheetNames(i)
    Next
    
cleanup:
    Set worksheetNames = Nothing
    Exit Sub
    
eh:
    RaiseError Err.Number, Err.Source, "modInitWorkbook.AddWorksheets", Err.Description, 116
    GoTo cleanup
End Sub

Private Sub DeleteWorksheets()
    
    On Error GoTo eh
    
    Application.DisplayAlerts = False 'switching off the alert button
    
    mWorkbook.Sheets("Sheet1").Delete
    mWorkbook.Sheets("Sheet2").Delete
    mWorkbook.Sheets("Sheet3").Delete
    
cleanup:
    Application.DisplayAlerts = True 'switching on the alert button
    Exit Sub
    
eh:
    RaiseError Err.Number, Err.Source, "modInitWorkbook.DeleteWorksheets", Err.Description, 136
    GoTo cleanup
    
End Sub

Private Sub InsertCategoryList()

    On Error GoTo eh
    
    Dim groupNames As Variant
    groupNames = SelectGroupsArray
    
    Dim worksheetNames As Variant
    worksheetNames = SelectWorksheetsArray
    
    Dim i As Integer
    Dim j As Integer
    Dim query As String
    Dim rs As New ADODB.Recordset
    
    For i = LBound(groupNames) To UBound(groupNames)
        query = SelectQuery(groupNames(i))
        rs.Open query, mConnection 'Line 23
        
        'Find the worksheet with the group name in its name
        'Copy this category or subcategory list to that worksheet
        For j = LBound(worksheetNames) To UBound(worksheetNames)
            If InStr(worksheetNames(j), sprintf(" - %1", groupNames(i))) > 0 Then
                mWorkbook.Sheets(worksheetNames(j)).Range("A2").CopyFromRecordset rs
                ResizeColumns mWorkbook.Sheets(worksheetNames(j)), 1
                rs.MoveFirst
            End If
        Next j
        
        rs.Close
    Next i
    
cleanup:
    Set rs = Nothing
    Set groupNames = Nothing
    Set worksheetNames = Nothing
    
    Exit Sub
    
eh:
    RaiseError Err.Number, Err.Source, "modInitWorkbook.InsertCategoryList", Err.Description, 23
    GoTo cleanup
End Sub

' Get the last entry in the Groups list because it will be used as the
' heading for the group list worksheets
Private Sub InsertListWorksheetHeading(wsName As Variant)

    On Error GoTo eh
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(GROUPS_WORKSHEET)

    Dim lastRow As Integer
    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row

    Dim heading As String
    heading = ws.Range("A" & lastRow)

    Set ws = mWorkbook.Sheets(wsName)

    ws.Range("A1:A1").Value = heading
    ws.Range("A1:N1").Font.Bold = True
    
cleanup:

    Set ws = Nothing
    Exit Sub
    
eh:
    RaiseError Err.Number, Err.Source, "modInitWorkbook.InsertListWorksheetHeading", Err.Description, 172
    GoTo cleanup
End Sub

Private Sub InsertWorksheetFormula(wsName As Variant)

    On Error GoTo eh
    
    Dim ws As Worksheet
    Set ws = mWorkbook.Sheets(wsName)
    
    Dim lastRow As Long
    lastRow = ws.Range("A:A").SpecialCells(xlCellTypeLastCell).row
    ws.Cells(lastRow + 2, 1).Value = "Total"
    ws.Cells(lastRow + 2, 1).Font.Bold = True
    
    AddRowSumFormula ws, lastRow
    AddColumnSumFormula ws, lastRow
    AddNumberFormat ws, lastRow
    
cleanup:

    Set ws = Nothing
    Exit Sub

eh:
    RaiseError Err.Number, Err.Source, "modInitWorkbook.InsertWorksheetFormula", Err.Description, 226
    GoTo cleanup
End Sub

Private Sub InsertWorksheetFormulas()

    On Error GoTo eh
    
    Dim worksheetNames As Variant
    worksheetNames = SelectWorksheetsArray

    Dim i As Integer
    For i = LBound(worksheetNames) To UBound(worksheetNames)
        If InStr(worksheetNames(i), "-") > 0 And _
            InStr(worksheetNames(i), "List") = 0 Then
            InsertWorksheetFormula worksheetNames(i)
        End If
    Next

cleanup:
    
    Set worksheetNames = Nothing
    Exit Sub

eh:
    RaiseError Err.Number, Err.Source, "modInitWorkbook.InsertWorksheetFormulas", Err.Description, 252
    GoTo cleanup
End Sub

Private Sub InsertWorksheetHeading(wsName As Variant)

    On Error GoTo eh
    
    Dim ws As Worksheet
    Set ws = mWorkbook.Sheets(wsName)
    
    ws.Range("A1:N1").Value = SelectMonthsHeading
    ws.Range("A1:N1").Font.Bold = True
    
cleanup:
    Set ws = Nothing
    Exit Sub
    
eh:
    RaiseError Err.Number, Err.Source, "modInitWorkbook.InsertWorksheetHeading", Err.Description, 151
    GoTo cleanup
End Sub

Private Sub InsertWorksheetHeadings()

    On Error GoTo eh
    
    Dim worksheetNames As Variant
    worksheetNames = SelectWorksheetsArray

    Dim i As Integer
    For i = LBound(worksheetNames) To UBound(worksheetNames)
        If InStr(worksheetNames(i), "List -") > 0 Then
            InsertListWorksheetHeading worksheetNames(i)
        Else
            InsertWorksheetHeading worksheetNames(i)
        End If
    Next

cleanup:

    Set worksheetNames = Nothing
    Exit Sub

eh:
    RaiseError Err.Number, Err.Source, "modInitWorkbook.InsertWorksheetHeadings", Err.Description, 200
    GoTo cleanup
End Sub

Private Function SelectMonthsHeading() As Variant

    On Error GoTo eh
    
    Dim monthsArray As Variant
    monthsArray = SelectMonthsArray
        
    Dim endsArray As Variant
    endsArray = SelectHeadingEndsArray
    
    monthsArray(LBound(monthsArray)) = endsArray(LBound(endsArray))
    monthsArray(UBound(monthsArray)) = endsArray(UBound(endsArray))
    
    SelectMonthsHeading = monthsArray

    Exit Function
    
eh:
    RaiseError Err.Number, Err.Source, "modInitWorkbook.InsertWorksheetFooters", Err.Description, 277
End Function

Private Function SelectQuery(queryName As Variant)

    On Error GoTo eh
    
    Dim queriesArray As Variant
    queriesArray = SelectQueriesArray
    
    Dim query As String
    Dim i As Integer
    For i = LBound(queriesArray) To UBound(queriesArray)
        If queriesArray(i, LBound(queriesArray)) = queryName Then
            query = queriesArray(UBound(queriesArray), LBound(queriesArray))
            Exit For
        ElseIf queriesArray(i, UBound(queriesArray)) = queryName Then
            query = queriesArray(UBound(queriesArray), UBound(queriesArray))
            Exit For
        Else
            query = ""
        End If
    Next
    
    SelectQuery = query

cleanup:
    
    Set queriesArray = Nothing
    Exit Function
    
eh:
    RaiseError Err.Number, Err.Source, "modInitWorkbook.SelectQuery", Err.Description, 298
    GoTo cleanup
End Function

' Public Methods
Public Sub CreateDestinationFile()

    On Error GoTo eh
    
    Set mConnection = New ADODB.Connection
    mConnection.Open sprintf(CONNECTION_STRING, SourceFilePath)
    
    Set mWorkbook = CurrentWorkbook
    
    AddWorksheets
    InsertWorksheetHeadings
    InsertCategoryList
    InsertWorksheetFormulas
    DeleteWorksheets
    
    mConnection.Close
    Exit Sub
    
eh:
    RaiseError Err.Number, Err.Source, "modInitWorkbook.CreateDestinationFile", Err.Description, 332
End Sub

'******************************************** Code Graveyard ******************************************************************

'Public Function InitDestinationWorkbook(sourceFile As String, destFile As String) As Boolean
'    mSourceFile = sourceFile
'    mDestinationFile = destFile
'
'    Set mConnection = New ADODB.connection
'    mConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & mSourceFile & "';Extended Properties='Excel 12.0;HDR=Yes';"
'
'    Set mWorkbook = Workbooks.Open(mDestinationFile)
'
'    'The workbook has not been initialized
'    If mWorkbook.Sheets.Count = 3 Then
'        AddWorksheets
'        DeleteWorksheets
'        InsertWorksheetHeadings
'        AddCategoryList
'        InsertWorksheetFooters
'    End If
'
'    mConnection.Close
'    mWorkbook.Save
'    mWorkbook.Close
'    Set mWorkbook = Nothing
'
'    InitDestinationWorkbook = True
'
'End Function

'Private Sub AddCategoryList()
'
'    Dim query As String
'    query = "SELECT [Master Category] " & _
'            "FROM [Spending$] " & _
'            "GROUP BY [Master Category]"
'
'    Dim rs As New ADODB.Recordset
'    rs.Open query, mConnection
'
'    mWorkbook.Sheets(CATEGORY_LIST).Range("A2").CopyFromRecordset rs
'    mWorkbook.Sheets(CATEGORY_LIST).Columns(1).AutoFit
'    rs.MoveFirst
'    mWorkbook.Sheets(COMBINED_BY_CATEGORY).Range("A2").CopyFromRecordset rs
'    mWorkbook.Sheets(COMBINED_BY_CATEGORY).Columns(1).AutoFit
'    rs.MoveFirst
'    mWorkbook.Sheets(BILLS_BY_CATEGORY).Range("A2").CopyFromRecordset rs
'    mWorkbook.Sheets(BILLS_BY_CATEGORY).Columns(1).AutoFit
'    rs.MoveFirst
'    mWorkbook.Sheets(JAKE_BY_CATEGORY).Range("A2").CopyFromRecordset rs
'    mWorkbook.Sheets(JAKE_BY_CATEGORY).Columns(1).AutoFit
'    rs.MoveFirst
'    mWorkbook.Sheets(SALLY_BY_CATEGORY).Range("A2").CopyFromRecordset rs
'    mWorkbook.Sheets(SALLY_BY_CATEGORY).Columns(1).AutoFit
'
'    rs.Close
'    Set rs = Nothing
'
'End Sub

'Private Sub InsertWorksheetHeadings()
'
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets(GROUPS_WORKSHSEET)
'
'    Dim lastRow As Integer
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim groupsArray As Variant
'    groupsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'    Dim lbIndex As Integer
'    lbIndex = LBound(groupsArray)
'
'    Dim ubIndex As Integer
'    ubIndex = UBound(groupsArray)
'
'    Set ws = ThisWorkbook.Sheets(WORKSHEETS_WORKSHEET)
'
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim worksheetsArray As Variant
'    worksheetsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'    Dim i As Integer
'    For i = LBound(worksheetsArray) To UBound(worksheetsArray)
'        If InStr(1, worksheetsArray(i), groupsArray(lbIndex), vbTextCompare) > 0 Or _
'            InStr(1, worksheetsArray(i), groupsArray(ubIndex), vbTextCompare) > 0 Then
'            InsertWorksheetHeading worksheetsArray(i)
'        Else
'            InsertListWorksheetHeading worksheetsArray(i)
'        End If
'    Next
'
'    Set ws = Nothing
'    Set groupsArray = Nothing
'
'End Sub

'Private Sub InsertWorksheetHeadings()
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets(GROUPS_WORKSHSEET)
'
'    Dim lastRow As Integer
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim groupsArray As Variant
'    groupsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'    Dim lbIndex As Integer
'    lbIndex = LBound(groupsArray)
'
'    Dim ubIndex As Integer
'    ubIndex = UBound(groupsArray)
'
'    Dim i As Integer
'
'    For i = LBound(mWorksheetNames) To UBound(mWorksheetNames)
'        If InStr(1, mWorksheetNames(i), groupsArray(lbIndex), vbTextCompare) > 0 Or _
'            InStr(1, mWorksheetNames(i), groupsArray(ubIndex), vbTextCompare) > 0 Then
'            InsertWorksheetHeading mWorksheetNames(i)
'        Else
'            InsertListWorksheetHeading mWorksheetNames(i)
'        End If
'    Next
'
'    Set ws = Nothing
'    Set groupsArray = Nothing
'
'End Sub

'Private Sub InsertWorksheetHeadings()
'
'    InsertWorksheetHeading COMBINED_BY_CATEGORY
'    InsertWorksheetHeading BILLS_BY_CATEGORY
'    InsertWorksheetHeading JAKE_BY_CATEGORY
'    InsertWorksheetHeading SALLY_BY_CATEGORY
'    InsertWorksheetHeading COMBINED_BY_SUB_CATEGORY
'    InsertWorksheetHeading BILLS_BY_SUB_CATEGORY
'    InsertWorksheetHeading JAKE_BY_SUB_CATEGORY
'    InsertWorksheetHeading SALLY_BY_SUB_CATEGORY
'    InsertWorksheetHeading CATEGORY_LIST
'    InsertWorksheetHeading SUB_CATEGORY_LIST
'End Sub

'Private Sub InsertWorksheetHeading(wsName As String)
'
'    Dim ws As Worksheet
'    Set ws = mWorkbook.Sheets(wsName)
'
'    Select Case wsName
'        Case COMBINED_BY_CATEGORY, BILLS_BY_CATEGORY, _
'            JAKE_BY_CATEGORY, SALLY_BY_CATEGORY, _
'            COMBINED_BY_SUB_CATEGORY, BILLS_BY_SUB_CATEGORY, _
'            JAKE_BY_SUB_CATEGORY, SALLY_BY_SUB_CATEGORY
'            ws.Range("A1:N1").Value = Array("Category", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Total")
'            ws.Range("A1:N1").Font.Bold = True
'        Case CATEGORY_LIST, SUB_CATEGORY_LIST
'            ws.Range("A1:A1").Value = "Category"
'    End Select
'
'End Sub

'Private Sub AddWorksheets()
'
'    Dim i As Integer
'
'    For i = 0 To UBound(mWorksheetNames)
'        mWorkbook.Sheets.Add.Name = mWorksheetNames(i)
'    Next
'
'End Sub

'Private Sub AddWorksheets()
'
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets(REQUIRED_WORKSHEET)
'
'    Dim lastRow As Integer
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim wsArray As Variant
'    ''https://stackoverflow.com/questions/37689847/creating-an-array-from-a-range-in-vba
'    wsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'    Dim i As Integer
'
'    For i = 1 To UBound(wsArray)
'        mWorkbook.Sheets.Add.Name = wsArray(i)
'    Next
'
'    Set ws = ThisWorkbook.Sheets(GROUPS_WORKHSEET)
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim groupsArray As Variant
'    groupsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'    Set ws = ThisWorkbook.Sheets(ACCOUNTS_WORKSHEET)
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim accountsArray As Variant
'    accountsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'    Dim j As Integer
'
'    For i = 1 To UBound(groupsArray)
'        For j = 1 To UBound(accountsArray)
'            mWorkbook.Sheets.Add.Name = accountsArray(j) & " - " & groupsArray(i)
'        Next j
'    Next i
'
'    Set ws = Nothing
'
'End Sub

'Private Sub AddSubCategoryList()
'
'    Dim query As String
'    query = "SELECT [SubCategory] " & _
'            "FROM [Spending$] " & _
'            "GROUP BY [SubCategory]"
'
'    Dim rs As New ADODB.Recordset
'    rs.Open query, mConnection
'
'    mWorkbook.Sheets(SUB_CATEGORY_LIST).Range("A2").CopyFromRecordset rs
'    mWorkbook.Sheets(SUB_CATEGORY_LIST).Columns(1).AutoFit
'    rs.MoveFirst
'    mWorkbook.Sheets(COMBINED_BY_SUB_CATEGORY).Range("A2").CopyFromRecordset rs
'    mWorkbook.Sheets(COMBINED_BY_SUB_CATEGORY).Columns(1).AutoFit
'    rs.MoveFirst
'    mWorkbook.Sheets(BILLS_BY_SUB_CATEGORY).Range("A2").CopyFromRecordset rs
'    mWorkbook.Sheets(BILLS_BY_SUB_CATEGORY).Columns(1).AutoFit
'    rs.MoveFirst
'    mWorkbook.Sheets(JAKE_BY_SUB_CATEGORY).Range("A2").CopyFromRecordset rs
'    mWorkbook.Sheets(JAKE_BY_SUB_CATEGORY).Columns(1).AutoFit
'    rs.MoveFirst
'    mWorkbook.Sheets(SALLY_BY_SUB_CATEGORY).Range("A2").CopyFromRecordset rs
'    mWorkbook.Sheets(SALLY_BY_SUB_CATEGORY).Columns(1).AutoFit
'
'    rs.Close
'    Set rs = Nothing
'
'End Sub

'Private Sub AddWorksheets()
'
'    mWorkbook.Sheets.Add.Name = TEMP
'    mWorkbook.Sheets.Add.Name = SUB_CATEGORY_LIST
'    mWorkbook.Sheets.Add.Name = CATEGORY_LIST
'    mWorkbook.Sheets.Add.Name = SALLY_BY_SUB_CATEGORY
'    mWorkbook.Sheets.Add.Name = JAKE_BY_SUB_CATEGORY
'    mWorkbook.Sheets.Add.Name = BILLS_BY_SUB_CATEGORY
'    mWorkbook.Sheets.Add.Name = COMBINED_BY_SUB_CATEGORY
'    mWorkbook.Sheets.Add.Name = SALLY_BY_CATEGORY
'    mWorkbook.Sheets.Add.Name = JAKE_BY_CATEGORY
'    mWorkbook.Sheets.Add.Name = BILLS_BY_CATEGORY
'    mWorkbook.Sheets.Add.Name = COMBINED_BY_CATEGORY
'
'End Sub

'Private Sub CollectWorksheetNames()
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets(REQUIRED_WORKSHEET)
'
'    Dim lastRow As Integer
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim wsArray As Variant
'    ''https://stackoverflow.com/questions/37689847/creating-an-array-from-a-range-in-vba
'    wsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'    ReDim mWorksheetNames(UBound(wsArray) - 1)
'
'    Dim i As Integer
'    For i = 0 To UBound(wsArray) - 1
'        mWorksheetNames(i) = wsArray(i + 1) 'wsArray has a 1 based index and mWorksheetNames has a 0 based index
'    Next
'
'    Dim lastIndex As Integer
'    lastIndex = UBound(mWorksheetNames)
'
'    Set ws = ThisWorkbook.Sheets(GROUPS_WORKHSEET)
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim groupsArray As Variant
'    groupsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'    Set ws = ThisWorkbook.Sheets(ACCOUNTS_WORKSHEET)
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim accountsArray As Variant
'    accountsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'    ReDim Preserve mWorksheetNames(lastIndex + (UBound(groupsArray) * UBound(accountsArray)))
'
'    Dim j As Integer
'
'    For i = 1 To UBound(groupsArray)
'        For j = 1 To UBound(accountsArray)
'            mWorksheetNames(lastIndex + j) = accountsArray(j) & " - " & groupsArray(i)
'        Next j
'        lastIndex = lastIndex + UBound(accountsArray)
'    Next i
'
'    Set ws = Nothing
'    Set wsArray = Nothing
'
'End Sub

'Private Sub InsertWorksheetFooter(wsName As String)
'
'    Dim ws As Worksheet
'    Set ws = mWorkbook.Sheets(wsName)
'
'    Dim lastRow As Long
'
'    Select Case wsName
'        Case COMBINED_BY_CATEGORY, BILLS_BY_CATEGORY, _
'            JAKE_BY_CATEGORY, SALLY_BY_CATEGORY, _
'            COMBINED_BY_SUB_CATEGORY, BILLS_BY_SUB_CATEGORY, _
'            JAKE_BY_SUB_CATEGORY, SALLY_BY_SUB_CATEGORY
'
'            lastRow = ws.Range("A:A").SpecialCells(xlCellTypeLastCell).row
'            ws.Cells(lastRow + 2, 1).Value = "Total"
'            ws.Cells(lastRow + 2, 1).Font.Bold = True
'
'            AddRowSumFormula ws, lastRow
'            AddColumnSumFormula ws, lastRow
'            AddNumberFormat ws, lastRow
'    End Select
'
'End Sub

'Private Sub InsertWorksheetFooters()
'
'    InsertWorksheetFooter COMBINED_BY_CATEGORY
'    InsertWorksheetFooter BILLS_BY_CATEGORY
'    InsertWorksheetFooter JAKE_BY_CATEGORY
'    InsertWorksheetFooter SALLY_BY_CATEGORY
'    InsertWorksheetFooter COMBINED_BY_SUB_CATEGORY
'    InsertWorksheetFooter BILLS_BY_SUB_CATEGORY
'    InsertWorksheetFooter JAKE_BY_SUB_CATEGORY
'    InsertWorksheetFooter SALLY_BY_SUB_CATEGORY
'
'End Sub

'Private Function SelectQuery(queryName As Variant)
'
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets(QUERIES_WORKSHEET)
'
'    Dim lastRow As Integer
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim queriesArray As Variant
'    ''https://stackoverflow.com/questions/37689847/creating-an-array-from-a-range-in-vba
'    queriesArray = Application.Transpose(ws.Range("A2", "B" & lastRow))
'
'    Dim query As String
'    Dim i As Integer
'    For i = LBound(queriesArray) To UBound(queriesArray)
'        If queriesArray(i, LBound(queriesArray)) = queryName Then
'            query = queriesArray(UBound(queriesArray), LBound(queriesArray))
'            Exit For
'        ElseIf queriesArray(i, UBound(queriesArray)) = queryName Then
'            query = queriesArray(UBound(queriesArray), UBound(queriesArray))
'            Exit For
'        Else
'            query = ""
'        End If
'    Next
'
'    SelectQuery = query
'
'End Function

'Private Sub AddSubCategoryList()
'
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets(QUERIES_WORKSHEET)
'
'    Dim lastRow As Integer
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim wsArray As Variant
'    ''https://stackoverflow.com/questions/37689847/creating-an-array-from-a-range-in-vba
'    wsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'End Sub

'Private Sub AddCategoryList()
'
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets(GROUPS_WORKSHEET)
'
'    Dim lastRow As Integer
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim groupNames As Variant
'    ''https://stackoverflow.com/questions/37689847/creating-an-array-from-a-range-in-vba
'    groupNames = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'    Set ws = ThisWorkbook.Sheets(WORKSHEETS_WORKSHEET)
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim worksheetNames As Variant
'    worksheetNames = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim query As String
'    Dim rs As New ADODB.Recordset
'
'    For i = LBound(groupNames) To UBound(groupNames)
'        query = SelectQuery(groupNames(i))
'        rs.Open query, mConnection
'
'        'Find the worksheet with the group name in its name
'        'Copy this category or subcategory list to that worksheet
'        For j = LBound(worksheetNames) To UBound(worksheetNames)
'            If InStr(worksheetNames(j), groupNames(i)) > 0 Then
'                mWorkbook.Sheets(worksheetNames(j)).Range("A2").CopyFromRecordset rs
'                rs.MoveFirst
'            End If
'        Next j
'
'        rs.Close
'    Next i
'
'    Set rs = Nothing
'
'End Sub

'Private Sub InsertWorksheetHeadings()
'
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets(WORKSHEETS_WORKSHEET)
'
'    Dim lastRow As Integer
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim worksheetsArray As Variant
'    worksheetsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'    Dim i As Integer
'    For i = LBound(worksheetsArray) To UBound(worksheetsArray)
'        If InStr(worksheetsArray(i), "-") > 0 Then
'            InsertWorksheetHeading worksheetsArray(i)
'        Else
'            InsertListWorksheetHeading worksheetsArray(i)
'        End If
'    Next
'
'    Set ws = Nothing
'    Set worksheetsArray = Nothing
'
'End Sub

'Private Sub InsertWorksheetFooters()
'
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets(WORKSHEETS_WORKSHEET)
'
'    Dim lastRow As Integer
'    lastRow = ws.Range("A2").SpecialCells(xlCellTypeLastCell).row
'
'    Dim worksheetsArray As Variant
'    worksheetsArray = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'    Dim i As Integer
'    For i = LBound(worksheetsArray) To UBound(worksheetsArray)
'        If InStr(worksheetsArray(i), "-") > 0 Then
'            InsertWorksheetFooter worksheetsArray(i)
'        End If
'    Next
'
'    Set ws = Nothing
'    Set worksheetsArray = Nothing
'
'End Sub

'Private Function SelectMonthsHeading() As Variant
'
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets(MONTHS_WORKSHEET)
'
'    Dim lastRow As Integer
'    lastRow = ws.Range("A1").SpecialCells(xlCellTypeLastCell).row
'
'    Dim monthsArray As Variant
'    monthsArray = Application.Transpose(ws.Range("A1", "A" & lastRow + 1))
'
'    Set ws = ThisWorkbook.Sheets(HEADING_ENDS_WORKSHEET)
'    lastRow = ws.Range("A1").SpecialCells(xlCellTypeLastCell).row
'
'    Dim ends As Variant
'    ends = Application.Transpose(ws.Range("A2", "A" & lastRow))
'
'    monthsArray(LBound(monthsArray)) = ends(LBound(ends))
'    monthsArray(UBound(monthsArray)) = ends(UBound(ends))
'
'    SelectMonthsHeading = monthsArray
'
'End Function

'Private Sub InsertWorksheetHeadings()
'
'    On Error GoTo eh
'
'    Dim worksheetNames As Variant
'    worksheetNames = SelectWorksheetsArray
'
'    Dim i As Integer
'    For i = LBound(worksheetNames) To UBound(worksheetNames)
'        If InStr(worksheetNames(i), "-") > 0 Then
'            InsertWorksheetHeading worksheetNames(i)
'        Else
'            InsertListWorksheetHeading worksheetNames(i)
'        End If
'    Next
'
'cleanup:
'
'    Set worksheetNames = Nothing
'    Exit Sub
'
'eh:
'    RaiseError Err.Number, Err.Source, "modInitWorkbook.InsertWorksheetHeadings", Err.Description, 198
'    GoTo cleanup
'End Sub
'

