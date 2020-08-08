Attribute VB_Name = "modCommon"
Option Explicit

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

'Report types
Public Const ACCOUNT1_BY_CATEGORY_TYPE As String = "Account1ByCategory"
Public Const ACCOUNT1_BY_SUB_CATEGORY_TYPE As String = "Account1BySubCategory"

Public Const COMBINED_BY_CATEGORY_TYPE As String = "CombinedByCategory"
Public Const COMBINED_BY_SUB_CATEGORY_TYPE As String = "CombinedBySubCategory"

Public Const ACCOUNT2_BY_CATEGORY_TYPE As String = "Account2ByCategory"
Public Const ACCOUNT2_BY_SUB_CATEGORY_TYPE As String = "Account2BySubCategory"

Public Const ACCOUNT3_BY_CATEGORY_TYPE As String = "Account3ByCategory"
Public Const ACCOUNT3_BY_SUB_CATEGORY_TYPE As String = "Account3BySubCategory"

'Worksheet names
Public Const ACCOUNT1_BY_CATEGORY As String = "Account1 By Category"
Public Const ACCOUNT1_BY_SUB_CATEGORY As String = "Account1 By Sub Category"

Public Const COMBINED_BY_CATEGORY As String = "Combined By Category"
Public Const COMBINED_BY_SUB_CATEGORY As String = "Combined By Sub Category"

Public Const ACCOUNT2_BY_CATEGORY As String = "Account2 By Category"
Public Const ACCOUNT2_BY_SUB_CATEGORY As String = "Account2 By Sub Category"

Public Const ACCOUNT3_BY_CATEGORY As String = "Account3 By Category"
Public Const ACCOUNT3_BY_SUB_CATEGORY As String = "Account3 By Sub Category"

Public Const CATEGORY_LIST As String = "Category List"
Public Const SUB_CATEGORY_LIST As String = "Sub Category List"

Public Const TEMP As String = "Temp"

Public Const NUMBER_FORMAT_ACCOUNTING As String = "_(#,##0.00_);_((#,##0.00);_(""-""??_);_(@_)"

Public Const TOTALS_COLUMN As Integer = 14

' Event values
Public Const BY_CATEGORY = 0
Public Const BY_SUB_CATEGORY = 1

Public Const PG_BY_CATEGORY As String = "pgByCategory"
Public Const EXCEL_FILTER As String = "Excel Files (*.xlsx), *.xlsx"
