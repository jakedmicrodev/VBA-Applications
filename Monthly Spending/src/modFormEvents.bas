Attribute VB_Name = "modFormEvents"
Option Explicit

Private mDestinationFilePath As String
Private mSourceFilePath As String
Private mMonth As Integer
Private mReportType As String

Private Function SelectReportType(mp As MultiPage) As String

    If mp.SelectedItem.Name = PG_BY_CATEGORY Then
        If mp.Pages(BY_CATEGORY).optCombinedByCategory.value Then
            SelectReportType = COMBINED_BY_CATEGORY_TYPE
        ElseIf mp.Pages(BY_CATEGORY).optAccount1ByCategory.value Then
            SelectReportType = ACCOUNT1_BY_CATEGORY_TYPE
        ElseIf mp.Pages(BY_CATEGORY).optAccount2ByCategory.value Then
            SelectReportType = ACCOUNT2_BY_CATEGORY_TYPE
        ElseIf mp.Pages(BY_CATEGORY).optAccount3ByCategory.value Then
            SelectReportType = ACCOUNT3_BY_CATEGORY_TYPE
        Else
            SelectReportType = ""
        End If
    Else
        If mp.Pages(BY_SUB_CATEGORY).optCombinedBySubCategory.value Then
            SelectReportType = COMBINED_BY_SUB_CATEGORY_TYPE
        ElseIf mp.Pages(BY_SUB_CATEGORY).optAccount1BySubCategory.value Then
            SelectReportType = ACCOUNT1_BY_SUB_CATEGORY_TYPE
        ElseIf mp.Pages(BY_SUB_CATEGORY).optAccount2BySubCategory.value Then
            SelectReportType = ACCOUNT2_BY_SUB_CATEGORY_TYPE
        ElseIf mp.Pages(BY_SUB_CATEGORY).optAccount3BySubCategory.value Then
            SelectReportType = ACCOUNT3_BY_SUB_CATEGORY_TYPE
        Else
            SelectReportType = ""
        End If
    End If
    
End Function

Private Function ValidateFilePath(path As String) As Boolean
    'PURPOSE: Function to determine if a File exists on the user's Computer
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    'RETURNS: True or False
    
    Dim Validation As String
        
    Validation = ""
    
    On Error Resume Next
        If path <> "" Then
            Validation = Dir(path)
        End If
    On Error GoTo 0
    
    If Validation <> "" Then ValidateFilePath = True
End Function

Private Function ValidateMonth(index As Integer) As Boolean
    ValidateMonth = False
    
    If index > 0 Then
        ValidateMonth = True
    End If
End Function

Public Property Get DestinationFilePath() As String
    DestinationFilePath = mDestinationFilePath
End Property

Public Property Get SourceFilePath() As String
    SourceFilePath = mSourceFilePath
End Property

Public Property Get ReportMonth() As String
    ReportMonth = mMonth
End Property

Public Property Get ReportType() As String
    ReportType = mReportType
End Property

Public Sub ClearSelection(mp As MultiPage)
    mp.Pages(BY_CATEGORY).optCombinedByCategory.value = False
    mp.Pages(BY_CATEGORY).optAccount1ByCategory.value = False
    mp.Pages(BY_CATEGORY).optAccount2ByCategory.value = False
    mp.Pages(BY_CATEGORY).optAccount3ByCategory.value = False

    mp.Pages(BY_SUB_CATEGORY).optCombinedBySubCategory.value = False
    mp.Pages(BY_SUB_CATEGORY).optAccount1BySubCategory.value = False
    mp.Pages(BY_SUB_CATEGORY).optAccount2BySubCategory.value = False
    mp.Pages(BY_SUB_CATEGORY).optAccount3BySubCategory.value = False
End Sub

Public Sub CollectSelections(frm As UserForm)
    mSourceFilePath = frm.txtSourceFilePath.Text
    mDestinationFilePath = frm.txtDestinationFilePath.Text
    mMonth = frm.cmbMonths.ListIndex
    mReportType = SelectReportType(frm.mpTabs)
End Sub

Public Sub LoadComboMonths(ByRef cb As ComboBox)
    cb.List = Array("Select", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
End Sub

Public Function SelectDestinationFile() As String
    
    SelectDestinationFile = Application.GetOpenFilename(filefilter:=EXCEL_FILTER, Title:="Open")

End Function

Public Function SelectSourceFile() As String
        
    SelectSourceFile = Application.GetOpenFilename(filefilter:=EXCEL_FILTER, Title:="Open")

End Function

Public Function ValidateFilePaths(frm As UserForm) As String
    ValidateFilePaths = ""
    
    If Not ValidateFilePath(frm.txtSourceFilePath) Then
        ValidateFilePaths = "Select a source file path"
        Exit Function
    End If
    
    If Not ValidateFilePath(frm.txtDestinationFilePath) Then
        ValidateFilePaths = "Select a destination file path"
        Exit Function
    End If
    
End Function

Public Function ValidateSelections(frm As UserForm) As String
    ValidateSelections = ""
    
    If Not ValidateFilePath(frm.txtSourceFilePath) Then
        ValidateSelections = "Select a source file path"
        Exit Function
    End If
    
    If Not ValidateFilePath(frm.txtDestinationFilePath) Then
        ValidateSelections = "Select a destination file path"
        Exit Function
    End If

    If Not ValidateMonth(frm.cmbMonths.ListIndex) Then
        ValidateSelections = "Select a month from the combobox"
        Exit Function
    End If
    
    If Not ValidateTabSelection(frm.mpTabs) Then
        ValidateSelections = "Select a report by category or sub category"
        Exit Function
    End If
End Function

Private Function ValidateTabSelection(mp As MultiPage) As Boolean
    
    If mp.SelectedItem.Name = PG_BY_CATEGORY Then
        ValidateTabSelection = _
        mp.Pages(BY_CATEGORY).optCombinedByCategory.value Or _
        mp.Pages(BY_CATEGORY).optAccount1ByCategory.value Or _
        mp.Pages(BY_CATEGORY).optAccount2ByCategory.value Or _
        mp.Pages(BY_CATEGORY).optAccount3ByCategory.value
    Else
        ValidateTabSelection = _
        mp.Pages(BY_SUB_CATEGORY).optCombinedBySubCategory.value Or _
        mp.Pages(BY_SUB_CATEGORY).optAccount1BySubCategory.value Or _
        mp.Pages(BY_SUB_CATEGORY).optAccount2BySubCategory.value Or _
        mp.Pages(BY_SUB_CATEGORY).optAccount3BySubCategory.value
    End If
End Function

