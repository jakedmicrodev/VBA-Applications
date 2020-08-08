Attribute VB_Name = "modInitWorkbook"
Option Explicit

Private mDestinationFile As String
Private mSourceFile As String
Private mWorkbook As Workbook
Private mConnection As ADODB.connection

Private Sub AddCategoryList()

    Dim query As String
    query = "SELECT [Master Category] " & _
            "FROM [Spending$] " & _
            "GROUP BY [Master Category]"
            
    Dim rs As New ADODB.Recordset
    rs.Open query, mConnection
    
    mWorkbook.Sheets(CATEGORY_LIST).Range("A2").CopyFromRecordset rs
    mWorkbook.Sheets(CATEGORY_LIST).Columns(1).AutoFit
    rs.MoveFirst
    mWorkbook.Sheets(COMBINED_BY_CATEGORY).Range("A2").CopyFromRecordset rs
    mWorkbook.Sheets(COMBINED_BY_CATEGORY).Columns(1).AutoFit
    rs.MoveFirst
    mWorkbook.Sheets(ACCOUNT1_BY_CATEGORY).Range("A2").CopyFromRecordset rs
    mWorkbook.Sheets(ACCOUNT1_BY_CATEGORY).Columns(1).AutoFit
    rs.MoveFirst
    mWorkbook.Sheets(ACCOUNT2_BY_CATEGORY).Range("A2").CopyFromRecordset rs
    mWorkbook.Sheets(ACCOUNT2_BY_CATEGORY).Columns(1).AutoFit
    rs.MoveFirst
    mWorkbook.Sheets(ACCOUNT3_BY_CATEGORY).Range("A2").CopyFromRecordset rs
    mWorkbook.Sheets(ACCOUNT3_BY_CATEGORY).Columns(1).AutoFit
    
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub AddColumnSumFormula(ws As Worksheet, lastRow As Long)

    Dim rowNum As Integer
    Dim cellName As String
    Dim formulaValue As String
    
    For rowNum = 2 To lastRow
        cellName = "N" & rowNum
        formulaValue = "=SUM(B" & rowNum & ":M" & rowNum & ")"
        ws.Range(cellName).Formula = formulaValue
        ws.Range(cellName).Font.Bold = True
    Next
    
End Sub

Private Sub AddNumberFormat(ws As Worksheet, lastRow As Long)
    Dim cellName As String
    
    cellName = "B2:N" & lastRow + 2
    ws.Range(cellName).NumberFormat = NUMBER_FORMAT_ACCOUNTING
End Sub

Private Sub AddRowSumFormula(ws As Worksheet, lastRow As Long)

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
    
End Sub

Private Sub AddSubCategoryList()
    
    Dim query As String
    query = "SELECT [SubCategory] " & _
            "FROM [Spending$] " & _
            "GROUP BY [SubCategory]"
            
    Dim rs As New ADODB.Recordset
    rs.Open query, mConnection
    
    mWorkbook.Sheets(SUB_CATEGORY_LIST).Range("A2").CopyFromRecordset rs
    mWorkbook.Sheets(SUB_CATEGORY_LIST).Columns(1).AutoFit
    rs.MoveFirst
    mWorkbook.Sheets(COMBINED_BY_SUB_CATEGORY).Range("A2").CopyFromRecordset rs
    mWorkbook.Sheets(COMBINED_BY_SUB_CATEGORY).Columns(1).AutoFit
    rs.MoveFirst
    mWorkbook.Sheets(ACCOUNT1_BY_SUB_CATEGORY).Range("A2").CopyFromRecordset rs
    mWorkbook.Sheets(ACCOUNT1_BY_SUB_CATEGORY).Columns(1).AutoFit
    rs.MoveFirst
    mWorkbook.Sheets(ACCOUNT2_BY_SUB_CATEGORY).Range("A2").CopyFromRecordset rs
    mWorkbook.Sheets(ACCOUNT2_BY_SUB_CATEGORY).Columns(1).AutoFit
    rs.MoveFirst
    mWorkbook.Sheets(ACCOUNT3_BY_SUB_CATEGORY).Range("A2").CopyFromRecordset rs
    mWorkbook.Sheets(ACCOUNT3_BY_SUB_CATEGORY).Columns(1).AutoFit
    
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub AddWorksheets()

    mWorkbook.Sheets.Add.Name = TEMP
    mWorkbook.Sheets.Add.Name = SUB_CATEGORY_LIST
    mWorkbook.Sheets.Add.Name = CATEGORY_LIST
    mWorkbook.Sheets.Add.Name = ACCOUNT3_BY_SUB_CATEGORY
    mWorkbook.Sheets.Add.Name = ACCOUNT2_BY_SUB_CATEGORY
    mWorkbook.Sheets.Add.Name = ACCOUNT1_BY_SUB_CATEGORY
    mWorkbook.Sheets.Add.Name = COMBINED_BY_SUB_CATEGORY
    mWorkbook.Sheets.Add.Name = ACCOUNT3_BY_CATEGORY
    mWorkbook.Sheets.Add.Name = ACCOUNT2_BY_CATEGORY
    mWorkbook.Sheets.Add.Name = ACCOUNT1_BY_CATEGORY
    mWorkbook.Sheets.Add.Name = COMBINED_BY_CATEGORY

End Sub

Private Sub DeleteWorksheets()

    mWorkbook.Sheets("Sheet1").Delete
    mWorkbook.Sheets("Sheet2").Delete
    mWorkbook.Sheets("Sheet3").Delete
    
End Sub

Private Sub InsertWorksheetHeading(wsName As String)
    
    Dim ws As Worksheet
    Set ws = mWorkbook.Sheets(wsName)
    
    Select Case wsName
        Case COMBINED_BY_CATEGORY, ACCOUNT1_BY_CATEGORY, _
            ACCOUNT2_BY_CATEGORY, ACCOUNT3_BY_CATEGORY, _
            COMBINED_BY_SUB_CATEGORY, ACCOUNT1_BY_SUB_CATEGORY, _
            ACCOUNT2_BY_SUB_CATEGORY, ACCOUNT3_BY_SUB_CATEGORY
            ws.Range("A1:N1").value = Array("Category", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Total")
            ws.Range("A1:N1").Font.Bold = True
        Case CATEGORY_LIST, SUB_CATEGORY_LIST
            ws.Range("A1:A1").value = "Category"
    End Select
    
End Sub

Private Sub InsertWorksheetHeadings()

    InsertWorksheetHeading COMBINED_BY_CATEGORY
    InsertWorksheetHeading ACCOUNT1_BY_CATEGORY
    InsertWorksheetHeading ACCOUNT2_BY_CATEGORY
    InsertWorksheetHeading ACCOUNT3_BY_CATEGORY
    InsertWorksheetHeading COMBINED_BY_SUB_CATEGORY
    InsertWorksheetHeading ACCOUNT1_BY_SUB_CATEGORY
    InsertWorksheetHeading ACCOUNT2_BY_SUB_CATEGORY
    InsertWorksheetHeading ACCOUNT3_BY_SUB_CATEGORY
    InsertWorksheetHeading CATEGORY_LIST
    InsertWorksheetHeading SUB_CATEGORY_LIST
End Sub

Private Sub InsertWorksheetFooter(wsName As String)

    Dim ws As Worksheet
    Set ws = mWorkbook.Sheets(wsName)
    
    Dim lastRow As Long
    
    Select Case wsName
        Case COMBINED_BY_CATEGORY, ACCOUNT1_BY_CATEGORY, _
            ACCOUNT2_BY_CATEGORY, ACCOUNT3_BY_CATEGORY, _
            COMBINED_BY_SUB_CATEGORY, ACCOUNT1_BY_SUB_CATEGORY, _
            ACCOUNT2_BY_SUB_CATEGORY, ACCOUNT3_BY_SUB_CATEGORY
            
            lastRow = ws.Range("A:A").SpecialCells(xlCellTypeLastCell).row
            ws.Cells(lastRow + 2, 1).value = "Total"
            ws.Cells(lastRow + 2, 1).Font.Bold = True
            
            AddRowSumFormula ws, lastRow
            AddColumnSumFormula ws, lastRow
            AddNumberFormat ws, lastRow
    End Select
    
End Sub

Private Sub InsertWorksheetFooters()

    InsertWorksheetFooter COMBINED_BY_CATEGORY
    InsertWorksheetFooter ACCOUNT1_BY_CATEGORY
    InsertWorksheetFooter ACCOUNT2_BY_CATEGORY
    InsertWorksheetFooter ACCOUNT3_BY_CATEGORY
    InsertWorksheetFooter COMBINED_BY_SUB_CATEGORY
    InsertWorksheetFooter ACCOUNT1_BY_SUB_CATEGORY
    InsertWorksheetFooter ACCOUNT2_BY_SUB_CATEGORY
    InsertWorksheetFooter ACCOUNT3_BY_SUB_CATEGORY
    
End Sub

Public Function InitDestinationWorkbook(sourceFile As String, destFile As String) As Boolean
    mSourceFile = sourceFile
    mDestinationFile = destFile
    
    Set mConnection = New ADODB.connection
    mConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & mSourceFile & "';Extended Properties='Excel 12.0;HDR=Yes';"

    Set mWorkbook = Workbooks.Open(mDestinationFile)
    
    'The workbook has not been initialized
    If mWorkbook.Sheets.Count = 3 Then
        AddWorksheets
        DeleteWorksheets
        InsertWorksheetHeadings
        AddCategoryList
        AddSubCategoryList
        InsertWorksheetFooters
    End If
        
    mWorkbook.Close
    mConnection.Close
    InitDestinationWorkbook = True
    
End Function
