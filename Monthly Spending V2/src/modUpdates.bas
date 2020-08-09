Attribute VB_Name = "modUpdates"
Option Explicit

Private mWorkbook As Workbook

'Private Methods
Private Function GetEndDate(thisDate As Date) As Date

    On Error GoTo eh
    
    Dim thisMonth As Integer
    thisMonth = month(thisDate)
    
    Dim thisYear As Integer
    thisYear = year(thisDate)
    
    Dim nextMonth As Integer
    Dim nextYear As Integer

    If thisMonth = 12 Then
        nextMonth = 1
        nextYear = thisYear + 1
    Else
        nextMonth = thisMonth + 1
        nextYear = thisYear
    End If
    
    Dim endDate As Date
    endDate = CDate(Str(nextMonth) & "/1/" & Str(nextYear)) - 1
    
    GetEndDate = endDate
    Exit Function
    
eh:
    RaiseError Err.Number, Err.Source, "modUpdates.GetEndDate", Err.Description, 14
End Function

Private Function SelectJoinData(thisFilePath As String, query As String) As ADODB.Recordset

    On Error GoTo eh
    
    Dim conn As New ADODB.Connection
    conn.Open sprintf(CONNECTION_STRING, thisFilePath)
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open query, conn

    Set SelectJoinData = rs
    Exit Function
    
eh:
    RaiseError Err.Number, Err.Source, "modUpdates.SelectJoinData", Err.Description, 46
End Function

Private Function SelectJoinQuery(thisGroup As String) As String

    On Error GoTo eh
    
    Dim queriesArray As Variant
    queriesArray = SelectQueriesArray
    
    Dim query As String
    query = ""
    
    Dim i As Integer
    For i = LBound(queriesArray, 1) To UBound(queriesArray, 2)
        If queriesArray(LBound(queriesArray), i) = sprintf("%1 %2", "Join", thisGroup) Then
            query = queriesArray(UBound(queriesArray), i)
            Exit For
        End If
    Next

cleanup:

    Set queriesArray = Nothing
    
    SelectJoinQuery = query
    Exit Function
        
eh:
    RaiseError Err.Number, Err.Source, "modUpdates.SelectJoinQuery", Err.Description, 63
    GoTo cleanup
End Function

Private Function SelectTempData(thisFilePath As String, query As String) As ADODB.Recordset
    
    On Error GoTo eh
    
    Dim conn As New ADODB.Connection
    conn.Open sprintf(CONNECTION_STRING, thisFilePath)
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open query, conn

    Set SelectTempData = rs
    Exit Function
    
eh:
    RaiseError Err.Number, Err.Source, "modUpdates.SelectTempData", Err.Description, 97
End Function

Private Function SelectTempQuery(thisGroup As String, startDate As Date, endDate As Date) As String
    
    On Error GoTo eh
    
    Dim queriesArray As Variant
    queriesArray = SelectQueriesArray
    
    Dim query As String
    query = ""
    
    Dim i As Integer
    For i = LBound(queriesArray, 1) To UBound(queriesArray, 2)
        If queriesArray(LBound(queriesArray), i) = sprintf("%1 %2", "Temp", thisGroup) Then
            query = queriesArray(UBound(queriesArray), i)
            Exit For
        End If
    Next

cleanup:

    Set queriesArray = Nothing
    
    SelectTempQuery = sprintf(query, CStr(startDate), CStr(endDate))
    Exit Function
    
eh:
    RaiseError Err.Number, Err.Source, "modUpdates.SelectTempQuery", Err.Description, 111
    GoTo cleanup
End Function

Private Sub WriteJoinData(rs As ADODB.Recordset, col As Integer, ws As Worksheet)
    
    On Error GoTo eh
    
    Dim row As Integer
    row = 2
    
    While Not rs.EOF
        ws.Cells(row, col) = rs("Amount")
        row = row + 1
        rs.MoveNext
    Wend
    
    Exit Sub
        
eh:
    RaiseError Err.Number, Err.Source, "modUpdates.WriteJoinData", Err.Description, 144
End Sub

Private Sub WriteTempData(rs As ADODB.Recordset)

    On Error GoTo eh
    
    Dim ws As Worksheet
    Set ws = mWorkbook.Sheets(TEMP)

    ws.Cells.ClearContents
    ws.Range("A1:B1").Value = Array("Category", "Amount")
    ws.Range("A2").CopyFromRecordset rs

cleanup:

    Set ws = Nothing
    Exit Sub
        
eh:
    RaiseError Err.Number, Err.Source, "modUpdates.WriteTempData", Err.Description, 164
    GoTo cleanup
End Sub

'Public Methods

Public Sub UpdateSpending()
    
    Set mWorkbook = CurrentWorkbook
    
    Dim thisYear As Integer
    thisYear = year(Date)
    
    Dim startDate As Date
    startDate = CDate(Str(MonthsIndex) & "/1/" & Str(thisYear))
    
    Dim endDate As Date
    endDate = GetEndDate(startDate)
    
    Dim query As String
    query = SelectTempQuery(GroupName, startDate, endDate)
    
    Dim rs As Recordset
    Set rs = SelectTempData(SourceFilePath, query)
    
    WriteTempData rs
    
    rs.Close
    Set rs = Nothing

    query = SelectJoinQuery(GroupName)

    Set rs = SelectJoinData(DestinationFilePath, query)
    
    Dim coll As Integer
    coll = MonthsIndex + 1 'The month column is one more that the month number e.g. thisMonth = 1 - col = 2
    
    Dim wsName As String
    wsName = AccountName & " - " & GroupName
    
    Dim ws As Worksheet
    Set ws = mWorkbook.Sheets(wsName)
    ws.Activate

    WriteJoinData rs, coll, ws
    
    rs.Close
    Set rs = Nothing

    ResizeColumns ws, coll

    Set ws = Nothing
    
End Sub

'***************** Code Graveyard ****************************************

'Private Sub ResizeColumns(ws As Worksheet)
'
'    Dim i As Integer
'
'    For i = 1 To 14
'        ws.Columns(i).AutoFit
'    Next
'
'End Sub

'Private Function SelectJoinData(query As String) As ADODB.Recordset
'    Dim conn As New ADODB.connection
'    Dim rs As New ADODB.Recordset
'    Dim filePath As String
'
'    filePath = "E:\Documents\Budget\2020\StagingSpending.xlsx"
'    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & filePath & "';Extended Properties='Excel 12.0;HDR=Yes';"
'    rs.Open query, conn
'
'    Set SelectJoinData = rs
'End Function

'Private Function SelectJoinData(query As String) As ADODB.Recordset
'    Dim conn As New ADODB.connection
'    Dim rs As ADODB.Recordset
'
'    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & mWorkbook.FullName & "';Extended Properties='Excel 12.0;HDR=Yes';"
'    Set rs = New ADODB.Recordset
'    rs.Open query, conn
'
'    Set SelectJoinData = rs
'End Function

'Private Sub WriteTempData(rs As ADODB.Recordset)
''    Dim wb As Workbook
'    Dim ws As Worksheet
'
''    Set wb = Workbooks.Open("E:\Documents\Budget\2020\StagingSpending.xlsx")
'    Set ws = mWorkbook.Sheets("Temp")
'
'    ws.Cells.ClearContents
'    ws.Range("A1:B1").value = Array("Category", "Amount")
'    ws.Range("A2").CopyFromRecordset rs
'
'    Set ws = Nothing
'End Sub

'Public Sub CreateReport1(thisReportType As String, thisMonth As Integer, thisFilePath As String)
'    Dim thisYear As Integer
'    Dim query As String
'    Dim ws As Worksheet
'    Dim rs As Recordset
'    Dim col As Integer
'    Dim startDate As Date
'    Dim endDate As Date
'
'    thisYear = year(Date)
'    startDate = CDate(Str(thisMonth) & "/1/" & Str(thisYear))
'    endDate = GetEndDate(startDate)
'    'The month column is one more that the month number e.g. thisMonth = 1 - col = 2
'    col = thisMonth + 1
'    query = SelectTempQuery(thisReportType, startDate, endDate)
'
'    Set rs = SelectTempData(thisFilePath, query)
'
'    WriteTempData rs
'    rs.Close
'    Set rs = Nothing
'
'    'Select combined data
'    query = SelectJoinQuery(thisReportType)
'
'    'Write combined data
'    Set ws = SelectWorkSheet(thisReportType)
'
'    Set rs = SelectJoinData(query)
'    WriteJoinData rs, col, ws
'    rs.Close
'    Set rs = Nothing
'
'End Sub

'Private Function SelectWorkSheet(thisReportType As String) As Worksheet
'
'    Select Case thisReportType
'        Case COMBINED_BY_CATEGORY_TYPE
'            Set SelectWorkSheet = mWorkbook.Sheets(COMBINED_BY_CATEGORY)
'        Case BILLS_BY_CATEGORY_TYPE
'            Set SelectWorkSheet = mWorkbook.Sheets(BILLS_BY_CATEGORY)
'        Case JAKE_BY_CATEGORY_TYPE
'            Set SelectWorkSheet = mWorkbook.Sheets(JAKE_BY_CATEGORY)
'        Case SALLY_BY_CATEGORY_TYPE
'            Set SelectWorkSheet = mWorkbook.Sheets(SALLY_BY_CATEGORY)
'        Case COMBINED_BY_SUB_CATEGORY_TYPE
'            Set SelectWorkSheet = mWorkbook.Sheets(COMBINED_BY_SUB_CATEGORY)
'        Case BILLS_BY_SUB_CATEGORY_TYPE
'            Set SelectWorkSheet = mWorkbook.Sheets(BILLS_BY_SUB_CATEGORY)
'        Case JAKE_BY_SUB_CATEGORY_TYPE
'            Set SelectWorkSheet = mWorkbook.Sheets(JAKE_BY_SUB_CATEGORY)
'        Case SALLY_BY_SUB_CATEGORY_TYPE
'            Set SelectWorkSheet = mWorkbook.Sheets(SALLY_BY_SUB_CATEGORY)
'    End Select
'
'End Function

'Public Sub UpdateSpending(thisReportType As String, thisMonth As Integer, thisSourceFile As String, thisDestinationFile As String)
'
'    Dim thisYear As Integer
'    thisYear = year(Date)
'
'    Dim startDate As Date
'    startDate = CDate(Str(thisMonth) & "/1/" & Str(thisYear))
'
'    Dim endDate As Date
'    endDate = GetEndDate(startDate)
'
'    Dim query As String
'    query = SelectTempQuery(thisReportType, startDate, endDate)
'
'    Dim rs As Recordset
'    Set rs = SelectTempData(thisSourceFile, query)
'
'    If mWorkbook Is Nothing Then
'        Set mWorkbook = Workbooks.Open(thisDestinationFile)
'    End If
'
'    WriteTempData rs
'
'    rs.Close
'    Set rs = Nothing
'
'    'Select combined data
'    query = SelectJoinQuery(thisReportType)
'
'    Set rs = SelectJoinData(thisDestinationFile, query)
'
'    Dim coll As Integer
'    coll = thisMonth + 1 'The month column is one more that the month number e.g. thisMonth = 1 - col = 2
'
'    Dim ws As Worksheet
'    Set ws = SelectWorkSheet(thisReportType)
'
'    WriteJoinData rs, coll, ws
'
'    rs.Close
'    Set rs = Nothing
'
'    ResizeColumns ws, coll
'
'    Set ws = Nothing
'
'End Sub

'Private Sub ResizeColumns(ws As Worksheet, col As Integer)
'
'    ws.Columns(col).AutoFit
'    ws.Columns(TOTALS_COLUMN).AutoFit
'
'End Sub

'Private Function SelectJoinQuery(thisReportType As String) As String
'
'    Dim query As String
'    Dim table As String
'
'    Select Case thisReportType
'        Case COMBINED_BY_CATEGORY_TYPE, BILLS_BY_CATEGORY_TYPE, JAKE_BY_CATEGORY_TYPE, SALLY_BY_CATEGORY_TYPE
'            table = "[Category List$]"
'        Case COMBINED_BY_SUB_CATEGORY_TYPE, BILLS_BY_SUB_CATEGORY_TYPE, JAKE_BY_SUB_CATEGORY_TYPE, SALLY_BY_SUB_CATEGORY_TYPE
'            table = "[Sub Category List$]"
'        End Select
'
'    query = "SELECT cl.[Category], t.[Amount] " & _
'            "FROM " & table & " AS cl " & _
'            "LEFT JOIN [Temp$] AS t " & _
'            "ON cl.[Category] = t.[Category]"
'
'    SelectJoinQuery = query
'
'End Function

'Private Function SelectTempQuery(thisReportType As String, startDate As Date, endDate As Date) As String
'
'    Dim query As String
'
'    Select Case thisReportType
'        Case COMBINED_BY_CATEGORY_TYPE, BILLS_BY_CATEGORY_TYPE, JAKE_BY_CATEGORY_TYPE, SALLY_BY_CATEGORY_TYPE
'            query = "SELECT [Master Category], Sum([Amount]) " & _
'                    "FROM [Spending$] " & _
'                    "WHERE [Date] Between #" & startDate & "# And #" & endDate & "# " & _
'                    "GROUP BY [Master Category]"
'        Case COMBINED_BY_SUB_CATEGORY_TYPE, BILLS_BY_SUB_CATEGORY_TYPE, JAKE_BY_SUB_CATEGORY_TYPE, SALLY_BY_SUB_CATEGORY_TYPE
'            query = "SELECT [SubCategory], Sum([Amount]) " & _
'                    "FROM [Spending$] " & _
'                    "WHERE [Date] Between #" & startDate & "# And #" & endDate & "# " & _
'                    "GROUP BY [SubCategory]"
'    End Select
'
'    SelectTempQuery = query
'
'End Function

'Private Function SelectJoinData(thisFilePath As String, query As String) As ADODB.Recordset
'
'    Dim conn As New ADODB.Connection
'    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & thisFilePath & "';Extended Properties='Excel 12.0;HDR=Yes';"
'
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    rs.Open query, conn
'
'    Set SelectJoinData = rs
'
'End Function

'Private Function SelectTempData(thisFilePath As String, query As String) As ADODB.Recordset
'
'    Dim conn As New ADODB.Connection
'    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & thisFilePath & "';Extended Properties='Excel 12.0;HDR=Yes';"
'
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    rs.Open query, conn
'
'    Set SelectTempData = rs
'End Function

'Public Sub UpdateSpending(thisGroup As String, thisAccount As String, thisMonth As Integer, thisSourceFile As String, thisDestinationFile As String)
'
'    mDestinationFile = thisDestinationFile
'
'    Dim thisYear As Integer
'    thisYear = year(Date)
'
'    Dim startDate As Date
'    startDate = CDate(Str(thisMonth) & "/1/" & Str(thisYear))
'
'    Dim endDate As Date
'    endDate = GetEndDate(startDate)
'
'    Dim query As String
''    query = SelectTempQuery(thisReportType, startDate, endDate)
'    query = SelectTempQuery(thisGroup, startDate, endDate)
'
'    Dim rs As Recordset
'    Set rs = SelectTempData(thisSourceFile, query)
'
'    If mWorkbook Is Nothing Then
'        Set mWorkbook = Workbooks.Open(thisDestinationFile)
'    End If
'
'    WriteTempData rs
'
'    rs.Close
'    Set rs = Nothing
'
'    query = SelectJoinQuery(thisGroup)
'
'    Set rs = SelectJoinData(thisDestinationFile, query)
'
'    Dim coll As Integer
'    coll = thisMonth + 1 'The month column is one more that the month number e.g. thisMonth = 1 - col = 2
'
'    Dim wsName As String
'    wsName = thisAccount & " - " & thisGroup
'
'    Dim ws As Worksheet
'    Set ws = mWorkbook.Sheets(wsName)
'
'    WriteJoinData rs, coll, ws
'
'    rs.Close
'    Set rs = Nothing
'
'    ResizeColumns ws, coll
'
'    Set ws = Nothing
'
'End Sub

'Public Sub ReleaseWorkbook()
'
'    On Error GoTo eh
'
'    If Not mWorkbook Is Nothing Then
'        If FileInUse(mDestinationFile) Then
'            mWorkbook.Close
'        End If
'    End If
'
'cleanup:
'
'    Set mWorkbook = Nothing
'    Exit Sub
'
'eh:
'    RaiseError Err.Number, Err.Source, "modUpdates.ReleaseWorkbook", Err.Description, 184
'    GoTo cleanup
'End Sub

