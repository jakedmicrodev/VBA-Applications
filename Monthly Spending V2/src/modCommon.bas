Attribute VB_Name = "modCommon"
Option Explicit

Private mAccountsIndex As Integer
Private mAccountName As String
Private mDestinationFilePath As String
Private mGroupName As String
Private mIsSavedAndClosed As Boolean
Private mMonthsIndex As Integer
Private mSourceFilePath As String
Private mSourceWorkbook As Workbook
Private mWorkbook As Workbook

Enum WorksheetColumns
    Category = 1
    January
    February
    March
    April
    May
    June
    July
    August
    September
    October
    November
    December
    Total
End Enum

'Required worksheets
Public Const ACCOUNTS_WORKSHEET As String = "Accounts"
Public Const GROUPS_WORKSHEET As String = "Groups"
Public Const MONTHS_WORKSHEET As String = "Months"
Public Const WORKSHEETS_WORKSHEET As String = "Worksheets"
Public Const HEADING_ENDS_WORKSHEET As String = "Heading Ends"
Public Const QUERIES_WORKSHEET As String = "Queries"

Public Const CONNECTION_STRING As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='%1';Extended Properties='Excel 12.0;HDR=Yes';"
Public Const DEFAULT_BACKCOLOR As Long = &HE0E0E0
Public Const EXCEL_FILTER As String = "Excel Files (*.xlsx), *.xlsx"
Public Const HIGHLIGHT_BACKCOLOR As Long = 16246743 'RGB(215, 231, 247)
Public Const HIGHLIGHT_BORDERCOLOR As Long = 11235144 'RGB(72, 111, 171)
Public Const MAIN_FORM_CAPTION As String = "Update Monthly Spending - %1"
Public Const NUMBER_FORMAT_ACCOUNTING As String = "_(#,##0.00_);_((#,##0.00);_(""-""??_);_(@_)"
Public Const SAVE_CHANGES_MSG As String = "Do you want to save changes you made to ""%1""?"
Public Const SAVE_CHANGES_TITLE As String = "Update Monthly Spending"
Public Const TEMP As String = "Temp"
Public Const TOTALS_COLUMN As Integer = 14

Public Property Get AccountsIndex() As Integer
    AccountsIndex = mAccountsIndex
End Property

Public Property Let AccountsIndex(Value As Integer)
    mAccountsIndex = Value
End Property

Public Property Get AccountName() As String
    AccountName = mAccountName
End Property

Public Property Let AccountName(Value As String)
    mAccountName = Value
End Property

Public Property Get CurrentWorkbook()
    Set CurrentWorkbook = mWorkbook
End Property

Public Property Set CurrentWorkbook(wb As Workbook)
    Set mWorkbook = wb
End Property

Public Property Get DestinationFilePath() As String
    DestinationFilePath = mDestinationFilePath
End Property

Public Property Let DestinationFilePath(Value As String)
    mDestinationFilePath = Value
End Property

Public Property Get GroupName() As String
    GroupName = mGroupName
End Property

Public Property Let GroupName(Value As String)
    mGroupName = Value
End Property

Public Property Get IsSavedAndClosed() As Boolean
    IsSavedAndClosed = mIsSavedAndClosed
End Property

Public Property Let IsSavedAndClosed(Value As Boolean)
    mIsSavedAndClosed = Value
End Property

Public Property Get MonthsIndex() As Integer
    MonthsIndex = mMonthsIndex
End Property

Public Property Let MonthsIndex(Value As Integer)
    mMonthsIndex = Value
End Property

Public Property Get SourceFilePath() As String
    SourceFilePath = mSourceFilePath
End Property

Public Property Let SourceFilePath(Value As String)
    mSourceFilePath = Value
End Property

Public Property Get SourceWorkbook()
    Set SourceWorkbook = mSourceWorkbook
End Property

Public Property Set SourceWorkbook(wb As Workbook)
    Set mSourceWorkbook = wb
End Property

Public Sub ResizeColumns(ws As Worksheet, col As Integer)

    ws.Columns(col).AutoFit
    ws.Columns(TOTALS_COLUMN).AutoFit

End Sub

Public Function FileInUse(sFileName) As Boolean

'https://stackoverflow.com/questions/9373082/detect-whether-excel-workbook-is-already-open
    On Error Resume Next
    
    Open sFileName For Binary Access Read Lock Read As #1
    Close #1
    FileInUse = IIf(Err.Number > 0, True, False)
    
    On Error GoTo 0
    
End Function


'******************************************** Code Graveyard ******************************************************************

' Event values
'Public Const BY_CATEGORY = 0
'Public Const BY_SUB_CATEGORY = 1

'Public Const PG_BY_CATEGORY As String = "pgByCategory"
'Private mWorkbookIsDirty As Boolean

''Report types
'Public Const BILLS_BY_CATEGORY_TYPE As String = "BillsByCategory"
'Public Const BILLS_BY_SUB_CATEGORY_TYPE As String = "BillsBySubCategory"
'
'Public Const COMBINED_BY_CATEGORY_TYPE As String = "CombinedByCategory"
'Public Const COMBINED_BY_SUB_CATEGORY_TYPE As String = "CombinedBySubCategory"
'
'Public Const JAKE_BY_CATEGORY_TYPE As String = "JakeByCategory"
'Public Const JAKE_BY_SUB_CATEGORY_TYPE As String = "JakeBySubCategory"
'
'Public Const SALLY_BY_CATEGORY_TYPE As String = "SallyByCategory"
'Public Const SALLY_BY_SUB_CATEGORY_TYPE As String = "SallyBySubCategory"

''Worksheet names
'Public Const BILLS_BY_CATEGORY As String = "Bills By Category"
'Public Const BILLS_BY_SUB_CATEGORY As String = "Bills By Sub Category"
'
'Public Const COMBINED_BY_CATEGORY As String = "Combined By Category"
'Public Const COMBINED_BY_SUB_CATEGORY As String = "Combined By Sub Category"
'
'Public Const JAKE_BY_CATEGORY As String = "Jake By Category"
'Public Const JAKE_BY_SUB_CATEGORY As String = "Jake By Sub Category"
'
'Public Const SALLY_BY_CATEGORY As String = "Sally By Category"
'Public Const SALLY_BY_SUB_CATEGORY As String = "Sally By Sub Category"
'
'Public Const CATEGORY_LIST As String = "Category List"
'Public Const SUB_CATEGORY_LIST As String = "Sub Category List"
'
'Public Property Get WorkbookIsDirty() As String
'    WorkbookIsDirty = mWorkbookIsDirty
'End Property
'
'Public Property Let WorkbookIsDirty(Value As String)
'    mWorkbookIsDirty = Value
'End Property

'Public Property Get FormIsInitializing() As Boolean
'    FormIsInitializing = mFormIsInitializing
'End Property
'
'Public Property Let FormIsInitializing(Value As Boolean)
'    mFormIsInitializing = Value
'End Property
'

