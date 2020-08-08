Attribute VB_Name = "modReports"
Option Explicit

Private mWorkbook As Workbook
Private mConnection As ADODB.connection

'Private Methods
Private Function GetEndDate(thisDate As Date) As Date

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
    
End Function

Private Sub ResizeColumns(ws As Worksheet, col As Integer)

    ws.Columns(col).AutoFit
    ws.Columns(TOTALS_COLUMN).AutoFit
    
End Sub

Private Function SelectJoinData(thisFilePath As String, query As String) As ADODB.Recordset

    Dim conn As New ADODB.connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & thisFilePath & "';Extended Properties='Excel 12.0;HDR=Yes';"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open query, conn

    Set SelectJoinData = rs
    
End Function

Private Function SelectJoinQuery(thisReportType As String) As String

    Dim query As String
    Dim table As String
    
    Select Case thisReportType
        Case COMBINED_BY_CATEGORY_TYPE, ACCOUNT1_BY_CATEGORY_TYPE, ACCOUNT2_BY_CATEGORY_TYPE, ACCOUNT3_BY_CATEGORY_TYPE
            table = "[Category List$]"
        Case COMBINED_BY_SUB_CATEGORY_TYPE, ACCOUNT1_BY_SUB_CATEGORY_TYPE, ACCOUNT2_BY_SUB_CATEGORY_TYPE, ACCOUNT3_BY_SUB_CATEGORY_TYPE
            table = "[Sub Category List$]"
        End Select
        
    query = "SELECT cl.[Category], t.[Amount] " & _
            "FROM " & table & " AS cl " & _
            "LEFT JOIN [Temp$] AS t " & _
            "ON cl.[Category] = t.[Category]"
                    
    SelectJoinQuery = query
    
End Function

Private Function SelectTempData(thisFilePath As String, query As String) As ADODB.Recordset
    
    Dim conn As New ADODB.connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & thisFilePath & "';Extended Properties='Excel 12.0;HDR=Yes';"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open query, conn

    Set SelectTempData = rs
    
End Function

Private Function SelectTempQuery(thisReportType As String, startDate As Date, endDate As Date) As String
    
    Dim query As String
    
    Select Case thisReportType
        Case COMBINED_BY_CATEGORY_TYPE, ACCOUNT1_BY_CATEGORY_TYPE, ACCOUNT2_BY_CATEGORY_TYPE, ACCOUNT3_BY_CATEGORY_TYPE
            query = "SELECT [Master Category], Sum([Amount]) " & _
                    "FROM [Spending$] " & _
                    "WHERE [Date] Between #" & startDate & "# And #" & endDate & "# " & _
                    "GROUP BY [Master Category]"
        Case COMBINED_BY_SUB_CATEGORY_TYPE, ACCOUNT1_BY_SUB_CATEGORY_TYPE, ACCOUNT2_BY_SUB_CATEGORY_TYPE, ACCOUNT3_BY_SUB_CATEGORY_TYPE
            query = "SELECT [SubCategory], Sum([Amount]) " & _
                    "FROM [Spending$] " & _
                    "WHERE [Date] Between #" & startDate & "# And #" & endDate & "# " & _
                    "GROUP BY [SubCategory]"
    End Select

    SelectTempQuery = query
    
End Function

Private Function SelectWorkSheet(thisReportType As String) As Worksheet

    Select Case thisReportType
        Case COMBINED_BY_CATEGORY_TYPE
            Set SelectWorkSheet = mWorkbook.Sheets(COMBINED_BY_CATEGORY)
        Case ACCOUNT1_BY_CATEGORY_TYPE
            Set SelectWorkSheet = mWorkbook.Sheets(ACCOUNT1_BY_CATEGORY)
        Case ACCOUNT2_BY_CATEGORY_TYPE
            Set SelectWorkSheet = mWorkbook.Sheets(ACCOUNT2_BY_CATEGORY)
        Case ACCOUNT3_BY_CATEGORY_TYPE
            Set SelectWorkSheet = mWorkbook.Sheets(ACCOUNT3_BY_CATEGORY)
        Case COMBINED_BY_SUB_CATEGORY_TYPE
            Set SelectWorkSheet = mWorkbook.Sheets(COMBINED_BY_SUB_CATEGORY)
        Case ACCOUNT1_BY_SUB_CATEGORY_TYPE
            Set SelectWorkSheet = mWorkbook.Sheets(ACCOUNT1_BY_SUB_CATEGORY)
        Case ACCOUNT2_BY_SUB_CATEGORY_TYPE
            Set SelectWorkSheet = mWorkbook.Sheets(ACCOUNT2_BY_SUB_CATEGORY)
        Case ACCOUNT3_BY_SUB_CATEGORY_TYPE
            Set SelectWorkSheet = mWorkbook.Sheets(ACCOUNT3_BY_SUB_CATEGORY)
    End Select
    
End Function

Private Sub WriteJoinData(rs As ADODB.Recordset, col As Integer, ws As Worksheet)
    
    Dim row As Integer
    
    row = 2
    
    While Not rs.EOF
        ws.Cells(row, col) = rs("Amount")
        row = row + 1
        rs.MoveNext
    Wend
    
End Sub

Private Sub WriteTempData(rs As ADODB.Recordset)

    Dim ws As Worksheet
    Set ws = mWorkbook.Sheets(TEMP)

    ws.Cells.ClearContents
    ws.Range("A1:B1").value = Array("Category", "Amount")
    ws.Range("A2").CopyFromRecordset rs

    Set ws = Nothing
    
End Sub

'Public Methods
Public Sub UpdateSpending(thisReportType As String, thisMonth As Integer, thisSourceFile As String, thisDestinationFile As String)

    Dim thisYear As Integer
    thisYear = year(Date)
    
    Dim startDate As Date
    startDate = CDate(Str(thisMonth) & "/1/" & Str(thisYear))
    
    Dim endDate As Date
    endDate = GetEndDate(startDate)
    
    Dim query As String
    query = SelectTempQuery(thisReportType, startDate, endDate)
    
    Dim rs As Recordset
    Set rs = SelectTempData(thisSourceFile, query)

    If mWorkbook Is Nothing Then
        Set mWorkbook = Workbooks.Open(thisDestinationFile)
    End If

    WriteTempData rs
    
    rs.Close
    Set rs = Nothing

    'Select combined data
    query = SelectJoinQuery(thisReportType)

    Set rs = SelectJoinData(thisDestinationFile, query)
    
    Dim coll As Integer
    coll = thisMonth + 1 'The month column is one more that the month number e.g. thisMonth = 1 - col = 2
    
    Dim ws As Worksheet
    Set ws = SelectWorkSheet(thisReportType)

    WriteJoinData rs, coll, ws
    
    rs.Close
    Set rs = Nothing

    ResizeColumns ws, coll

    Set ws = Nothing
    
End Sub
