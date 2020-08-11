Attribute VB_Name = "modFormEvents"
Option Explicit

' Private Members
Private Sub CloseCurrentWorkbook()
    
    On Error Resume Next
    
    Dim wb As Workbook
    Set wb = CurrentWorkbook
    
    wb.Close
    Set wb = Nothing
    Set CurrentWorkbook = Nothing
    
End Sub

Private Sub CreateNewWorkbook()

    On Error GoTo eh
    
    Set CurrentWorkbook = Workbooks.Add
    DestinationFilePath = Application.path & Application.PathSeparator & ActiveWorkbook.FullName

    Exit Sub
    
eh:
    RaiseError Err.Number, Err.Source, "modFormEvents.CreateNewWorkbook", Err.Description, 21
End Sub

Private Sub InitCommonVariables()
    
    AccountsIndex = -1
    AccountName = ""
    Set CurrentWorkbook = Nothing
    DestinationFilePath = ""
    GroupName = ""
    IsBlankWorkbook = True
    IsSavedAndClosed = False
    MonthsIndex = -1
    SourceFilePath = ""
    Set SourceWorkbook = Nothing
    
End Sub

Private Sub LoadCombobox(ByRef cb As ComboBox, worksheetName As String)

    On Error GoTo eh
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(worksheetName)

    Dim lastRow As Integer
    lastRow = ws.Range("A1").SpecialCells(xlCellTypeLastCell).row

    Dim arr As Variant
    ''https://stackoverflow.com/questions/37689847/creating-an-array-from-a-range-in-vba
    arr = Application.Transpose(ws.Range("A1", "A" & lastRow))
    arr(1) = "Select"

    cb.List = arr

cleanup:

    Set ws = Nothing
    Exit Sub
    
eh:
    RaiseError Err.Number, Err.Source, "modFormEvents.LoadCombobox", Err.Description, 22
    GoTo cleanup
End Sub

Private Sub SaveAndCloseWorkbook()
    
    On Error GoTo eh
    
    OnSaveClick
    CloseCurrentWorkbook
    
    Exit Sub
    
eh:
    RaiseError Err.Number, Err.Source, "modFormEvents.SaveAndCloseWorkbook", Err.Description, 432
End Sub

Private Function ValidateFilePath(thisPath As String) As Boolean
    'PURPOSE: Function to determine if a File exists on the user's Computer
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    'RETURNS: True or False
    
    Dim Validation As String
        
    Validation = ""
    
    On Error Resume Next
        If thisPath <> "" Then
            Validation = Dir(thisPath)
        End If
    On Error GoTo 0
    
    If Validation <> "" Then ValidateFilePath = True
End Function

Public Sub EnableFormControls(frm As UserForm)

    On Error GoTo eh
    
    Dim wb As Workbook
    Set wb = CurrentWorkbook

    Dim isEnabled As Boolean
    isEnabled = True
    
    If wb Is Nothing Then
        isEnabled = False
    End If
    
    Set wb = Nothing
    
    frm.lblClose.Enabled = isEnabled
    frm.lblCloseBackground.Enabled = isEnabled
    frm.lblSave.Enabled = isEnabled
    frm.lblSaveBackground.Enabled = isEnabled
    frm.lblSaveAs.Enabled = isEnabled
    frm.lblSaveAsBackground.Enabled = isEnabled
    frm.MultiPage1.Pages(1).Enabled = isEnabled
    
    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.EnableFormControls", 71
End Sub

Public Sub OnAccountsChange(frm As UserForm)
    
    On Error GoTo eh
    
    AccountsIndex = frm.cboAccounts.ListIndex
    AccountName = frm.cboAccounts.Text
    
    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnAccountsChange", 66
End Sub

Public Sub OnCategoryClick(frm As UserForm)

    On Error GoTo eh
    
    GroupName = frm.optCategory.Caption

    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnCategoryClick", 79
End Sub

Public Sub OnCloseClick()

    CloseCurrentWorkbook

End Sub

Public Sub OnCreateDestinationFileClick(frm As UserForm)

    On Error GoTo eh
    
    CreateDestinationFile
    OnSaveAsClick

    Exit Sub
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnCreateDestinationFileClick", 97
End Sub

Public Sub OnFormInitialize(frm As UserForm)
    On Error GoTo eh
    
    InitCommonVariables
    CreateNewWorkbook
    
    frm.optCategory.Value = True
    
    LoadCombobox frm.cboMonths, "Months"
    frm.cboMonths.ListIndex = 0
    
    LoadCombobox frm.cboAccounts, "Accounts"
    frm.cboAccounts.ListIndex = 0
        
    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnFormInitialize", 155
End Sub

Public Sub OnMonthsChange(frm As UserForm)

    On Error GoTo eh
    
    MonthsIndex = frm.cboMonths.ListIndex

    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnMonthsChange", 178
End Sub

Public Sub OnMouseDown(frm As UserForm, Caption As String)

    On Error GoTo eh
    
    Select Case Caption
        Case "Close"
            frm.lblClose.Left = frm.lblClose.Left + 1
            frm.lblClose.Top = frm.lblClose.Top + 1
        Case "Create Destination File"
            frm.lblCreateDestinationFile.Left = frm.lblCreateDestinationFile.Left + 1
            frm.lblCreateDestinationFile.Top = frm.lblCreateDestinationFile.Top + 1
        Case "Exit"
            frm.lblExit.Left = frm.lblExit.Left + 1
            frm.lblExit.Top = frm.lblExit.Top + 1
        Case "New"
            frm.lblNew.Left = frm.lblNew.Left + 1
            frm.lblNew.Top = frm.lblNew.Top + 1
        Case "Open"
            frm.lblOpen.Left = frm.lblOpen.Left + 1
            frm.lblOpen.Top = frm.lblOpen.Top + 1
        Case "Save"
            frm.lblSave.Left = frm.lblSave.Left + 1
            frm.lblSave.Top = frm.lblSave.Top + 1
        Case "Save As"
            frm.lblSaveAs.Left = frm.lblSaveAs.Left + 1
            frm.lblSaveAs.Top = frm.lblSaveAs.Top + 1
        Case "Select"
            frm.lblSelectSourceFile.Left = frm.lblSelectSourceFile.Left + 1
            frm.lblSelectSourceFile.Top = frm.lblSelectSourceFile.Top + 1
        Case "Select New"
            frm.lblSelectNewSourceFile.Left = frm.lblSelectNewSourceFile.Left + 1
            frm.lblSelectNewSourceFile.Top = frm.lblSelectNewSourceFile.Top + 1
        Case "Update"
            frm.lblUpdate.Left = frm.lblUpdate.Left + 1
            frm.lblUpdate.Top = frm.lblUpdate.Top + 1
    End Select
    
    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnMouseDown", 190
End Sub

Public Sub OnMouseMove(frm As UserForm, Caption As String)

    On Error GoTo eh
    
    Select Case Caption
        Case "Close"
            frm.lblClose.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblCloseBackground.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblCloseBackground.BorderColor = HIGHLIGHT_BORDERCOLOR
            frm.lblCloseBackground.BorderStyle = fmBorderStyleSingle
        Case "Create Destination File"
            frm.lblCreateDestinationFile.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblCreateDestinationFile.BorderColor = HIGHLIGHT_BORDERCOLOR
            frm.lblCreateDestinationFile.BorderStyle = fmBorderStyleSingle
        Case "Exit"
            frm.lblExit.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblExitBackground.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblExitBackground.BorderColor = HIGHLIGHT_BORDERCOLOR
            frm.lblExitBackground.BorderStyle = fmBorderStyleSingle
       Case "New"
            frm.lblNew.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblNewBackground.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblNewBackground.BorderColor = HIGHLIGHT_BORDERCOLOR
            frm.lblNewBackground.BorderStyle = fmBorderStyleSingle
       Case "Open"
            frm.lblOpen.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblOpenBackground.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblOpenBackground.BorderColor = HIGHLIGHT_BORDERCOLOR
            frm.lblOpenBackground.BorderStyle = fmBorderStyleSingle
        Case "Save"
            frm.lblSave.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblSaveBackground.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblSaveBackground.BorderColor = HIGHLIGHT_BORDERCOLOR
            frm.lblSaveBackground.BorderStyle = fmBorderStyleSingle
        Case "Save As"
            frm.lblSaveAs.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblSaveAsBackground.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblSaveAsBackground.BorderColor = HIGHLIGHT_BORDERCOLOR
            frm.lblSaveAsBackground.BorderStyle = fmBorderStyleSingle
        Case "Select"
            frm.lblSelectSourceFile.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblSelectSourceFile.BorderColor = HIGHLIGHT_BORDERCOLOR
            frm.lblSelectSourceFile.BorderStyle = fmBorderStyleSingle
        Case "Select New"
            frm.lblSelectNewSourceFile.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblSelectNewSourceFile.BorderColor = HIGHLIGHT_BORDERCOLOR
            frm.lblSelectNewSourceFile.BorderStyle = fmBorderStyleSingle
        Case "Update"
            frm.lblUpdate.BackColor = HIGHLIGHT_BACKCOLOR
            frm.lblUpdate.BorderColor = HIGHLIGHT_BORDERCOLOR
            frm.lblUpdate.BorderStyle = fmBorderStyleSingle
    End Select
    
    ResetLabel frm, Caption
    
    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnMouseMove", 233
End Sub

Public Sub OnMouseUp(frm As UserForm, Caption As String)
    
    On Error GoTo eh
    
    Select Case Caption
        Case "Close"
            frm.lblClose.Left = frm.lblClose.Left - 1
            frm.lblClose.Top = frm.lblClose.Top - 1
        Case "Create Destination File"
            frm.lblCreateDestinationFile.Left = frm.lblCreateDestinationFile.Left - 1
            frm.lblCreateDestinationFile.Top = frm.lblCreateDestinationFile.Top - 1
        Case "Exit"
            frm.lblExit.Left = frm.lblExit.Left - 1
            frm.lblExit.Top = frm.lblExit.Top - 1
        Case "New"
            frm.lblNew.Left = frm.lblNew.Left - 1
            frm.lblNew.Top = frm.lblNew.Top - 1
        Case "Open"
            frm.lblOpen.Left = frm.lblOpen.Left - 1
            frm.lblOpen.Top = frm.lblOpen.Top - 1
        Case "Save"
            frm.lblSave.Left = frm.lblSave.Left - 1
            frm.lblSave.Top = frm.lblSave.Top - 1
        Case "Save As"
            frm.lblSaveAs.Left = frm.lblSaveAs.Left - 1
            frm.lblSaveAs.Top = frm.lblSaveAs.Top - 1
        Case "Select"
            frm.lblSelectSourceFile.Left = frm.lblSelectSourceFile.Left - 1
            frm.lblSelectSourceFile.Top = frm.lblSelectSourceFile.Top - 1
        Case "Select New"
            frm.lblSelectNewSourceFile.Left = frm.lblSelectNewSourceFile.Left - 1
            frm.lblSelectNewSourceFile.Top = frm.lblSelectNewSourceFile.Top - 1
        Case "Update"
            frm.lblUpdate.Left = frm.lblUpdate.Left - 1
            frm.lblUpdate.Top = frm.lblUpdate.Top - 1
    End Select

    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnMouseUp", 294
End Sub

Public Sub OnMultiPageChange(frm As UserForm)

    On Error GoTo eh
    
    frm.fraNewFile.Visible = False
    
    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnMultiPageChange", 337
End Sub

Public Sub OnNewClick(frm As UserForm)

    On Error GoTo eh
    
    If Not IsBlankWorkbook Then
        CloseCurrentWorkbook
        CreateNewWorkbook
    End If
    
    frm.txtSelectNewSourceFile.Text = ""
    frm.lblCreateDestinationFile.Enabled = False
    frm.fraNewFile.Visible = True
    
    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnNewClick", 350
End Sub

Public Sub OnOpenClick(frm As UserForm)
    
    On Error GoTo eh
    
    Dim filePath As String
    filePath = Application.GetOpenFilename(filefilter:=EXCEL_FILTER, Title:="Open")
    
    If filePath <> "False" Then
        DestinationFilePath = filePath
        CloseCurrentWorkbook
        Set CurrentWorkbook = Workbooks.Open(DestinationFilePath)
    End If
    
    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnOpenClick", 354
End Sub

Public Function OnQueryClose() As VbMsgBoxResult
    
    Dim wb As Workbook
    Set wb = CurrentWorkbook
    
    Dim retVal As VbMsgBoxResult
    If Not wb.Saved Then
        retVal = MsgBox(sprintf(SAVE_CHANGES_MSG, wb.Name), vbYesNoCancel, SAVE_CHANGES_TITLE)
    End If
    
    If retVal = vbYes Then
        SaveAndCloseWorkbook
        IsSavedAndClosed = True
    ElseIf retVal = vbNo Then
        wb.Saved = True
        CloseCurrentWorkbook
        IsSavedAndClosed = True
    Else
        'retVal = vbCancel
    End If
    
    Set wb = Nothing
    
    OnQueryClose = retVal
        
End Function

Public Sub OnSourceFilePathChange(frm As UserForm)
    On Error GoTo eh
    
    Dim accountIndex As Integer
    
    For accountIndex = 0 To frm.cboAccounts.ListCount - 1
        If InStr(1, frm.txtSourceFilePath.Text, CStr(frm.cboAccounts.List(accountIndex))) Then
            frm.cboAccounts.ListIndex = accountIndex
            Exit For
        End If
    Next
    
    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnSourceFilePathChange", 367
End Sub

Public Sub OnSaveAsClick()

    On Error GoTo eh
    
    Dim wb As Workbook
    Set wb = CurrentWorkbook
    
    Dim fileName As String
    fileName = Application.GetSaveAsFilename(ActiveWorkbook.FullName, filefilter:=EXCEL_FILTER, Title:="Save As")
    
    If fileName <> "False" Then
        wb.SaveAs fileName
        Set CurrentWorkbook = wb
        DestinationFilePath = fileName
    End If
    
cleanup:

    Set wb = Nothing
    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnSaveAsClick", 464
    GoTo cleanup
End Sub

Public Sub OnSaveClick()

    On Error GoTo eh
    
    Dim wb As Workbook
    Set wb = CurrentWorkbook
    
    wb.Save
    Set wb = Nothing
    
    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnSaveClick", 486
End Sub

Public Sub OnSelectNewSourceFileChange(frm As UserForm)
    
    On Error GoTo eh
    
    frm.lblCreateDestinationFile.Enabled = True
    
    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnSelectNewSourceFileChange", 438
End Sub

Public Sub OnSelectNewSourceFileClick(frm As UserForm)
    
    On Error GoTo eh
    
    Dim filePath As String
    filePath = Application.GetOpenFilename(filefilter:=EXCEL_FILTER, Title:="Open")
    
    If filePath <> "" Then
        frm.txtSelectNewSourceFile.Text = filePath
        frm.txtSourceFilePath.Text = filePath
        SourceFilePath = filePath
    End If

    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnSelectNewSourceFileClick", 450
End Sub

Public Sub OnSelectSourceFileClick(frm As UserForm)
    
    On Error GoTo eh
    
    Dim filePath As String
    filePath = Application.GetOpenFilename(filefilter:=EXCEL_FILTER, Title:="Open")
    
    If filePath <> "" Then
        frm.txtSourceFilePath.Text = filePath
        SourceFilePath = filePath
    End If
    
    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnSelectSourceFileClick", 464
End Sub

Public Sub OnSubCategoryClick(frm As UserForm)

    On Error GoTo eh
    
    GroupName = frm.optSubCategory.Caption
    
    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnSubCategoryClick", 477
End Sub

Public Sub OnUpdateClick()
    
    On Error GoTo eh
    
    Dim message As String
    message = ValidateSelections()
    
    If message <> "" Then
        MsgBox message
        Exit Sub
    End If
    
    UpdateSpending

    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.OnUpdateClick", 497
End Sub

Public Sub ResetLabel(frm As UserForm, Caption As String)

    On Error GoTo eh
    
    If frm.lblCreateDestinationFile.Caption <> Caption Then
        frm.lblCreateDestinationFile.BackColor = DEFAULT_BACKCOLOR
        frm.lblCreateDestinationFile.BorderStyle = fmBorderStyleNone
    End If

    If frm.lblClose.Caption <> Caption Then
        frm.lblClose.BackColor = DEFAULT_BACKCOLOR
        frm.lblCloseBackground.BackColor = DEFAULT_BACKCOLOR
        frm.lblCloseBackground.BorderStyle = fmBorderStyleNone
    End If
    
    If frm.lblExit.Caption <> Caption Then
        frm.lblExit.BackColor = DEFAULT_BACKCOLOR
        frm.lblExitBackground.BackColor = DEFAULT_BACKCOLOR
        frm.lblExitBackground.BorderStyle = fmBorderStyleNone
    End If

    If frm.lblNew.Caption <> Caption Then
        frm.lblNew.BackColor = DEFAULT_BACKCOLOR
        frm.lblNewBackground.BackColor = DEFAULT_BACKCOLOR
        frm.lblNewBackground.BorderStyle = fmBorderStyleNone
    End If

    If frm.lblOpen.Caption <> Caption Then
        frm.lblOpen.BackColor = DEFAULT_BACKCOLOR
        frm.lblOpenBackground.BackColor = DEFAULT_BACKCOLOR
        frm.lblOpenBackground.BorderStyle = fmBorderStyleNone
    End If

    If frm.lblSave.Caption <> Caption Then
        frm.lblSave.BackColor = DEFAULT_BACKCOLOR
        frm.lblSaveBackground.BackColor = DEFAULT_BACKCOLOR
        frm.lblSaveBackground.BorderStyle = fmBorderStyleNone
    End If

    If frm.lblSaveAs.Caption <> Caption Then
        frm.lblSaveAs.BackColor = DEFAULT_BACKCOLOR
        frm.lblSaveAsBackground.BackColor = DEFAULT_BACKCOLOR
        frm.lblSaveAsBackground.BorderStyle = fmBorderStyleNone
    End If

    If frm.lblSelectSourceFile.Tag <> Caption Then
        frm.lblSelectSourceFile.BackColor = DEFAULT_BACKCOLOR
        frm.lblSelectSourceFile.BorderStyle = fmBorderStyleNone
    End If

    If frm.lblSelectNewSourceFile.Tag <> Caption Then
        frm.lblSelectNewSourceFile.BackColor = DEFAULT_BACKCOLOR
        frm.lblSelectNewSourceFile.BorderStyle = fmBorderStyleNone
    End If

    If frm.lblUpdate.Caption <> Caption Then
        frm.lblUpdate.BackColor = DEFAULT_BACKCOLOR
        frm.lblUpdate.BorderStyle = fmBorderStyleNone
    End If

    Exit Sub
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.ResetLabel", 509
End Sub

Public Function SetMainFormCaption() As String

    On Error GoTo eh
    
    Dim wb As Workbook
    Set wb = CurrentWorkbook
    
    Dim workbookName As String
    workbookName = ""
    
    If Not wb Is Nothing Then
        workbookName = wb.Name
    End If
    
    SetMainFormCaption = sprintf(MAIN_FORM_CAPTION, workbookName)
    
    Exit Function
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.SetMainFormCaption", 143
End Function

Public Function ValidateSelections() As String

    On Error GoTo eh
    
    If Not ValidateFilePath(SourceFilePath) Then
        ValidateSelections = "Select a source file path" & vbNewLine
    End If
    
    If Not ValidateFilePath(DestinationFilePath) Then
        ValidateSelections = ValidateSelections & "Create a new workbook or open an existing workbook" & vbNewLine
    End If
    
    If AccountsIndex < 0 Then
        ValidateSelections = ValidateSelections & "Select an account from the Accounts combobox" & vbNewLine
    End If
    
    If MonthsIndex < 0 Then
        ValidateSelections = ValidateSelections & "Select a month from the Months combobox"
    End If
    
    Exit Function
    
eh:
    DisplayError Err.Source, Err.Description, "modFormEvents.ResetLabel", 576
End Function

'*************************************** Code Graveyard ************************************************************

'Private Function SelectFile(dialogTitle As String) As String
'
'    SelectFile = Application.GetOpenFilename(filefilter:=EXCEL_FILTER, Title:=dialogTitle)
'
'End Function

'Public Function SelectDestinationFile() As String
'
'    SelectDestinationFile = SelectFile("Open Destination File")
'
'End Function

'Public Function SelectSourceFile() As String
'
'    SelectSourceFile = SelectFile("Open Source File")
'
'End Function

'Public Sub OnFormInitialize(frm As UserForm)
'
'    Dim wb As Workbook
'    Set wb = Workbooks.Add
'
'    Set CurrentWorkbook = wb
'
'    wb.Close False
'    Debug.Print "Total workbooks after close: " & Workbooks.Count
'
'    frm.optCategory.Value = True
'
'    LoadCombobox frm.cbMonths, "Months"
'    frm.cbMonths.ListIndex = 0
'
'    LoadCombobox frm.cbAccounts, "Accounts"
'    frm.cbAccounts.ListIndex = 0
'
'End Sub

'Public Function ValidateSelections(frm As UserForm) As String
'    ValidateSelections = ""
'
'    If Not ValidateFilePath(frm.txtSourceFilePath) Then
'        ValidateSelections = "Select a source file path"
'        Exit Function
'    End If
'
'    If Not ValidateFilePath(frm.txtDestinationFilePath) Then
'        ValidateSelections = "Select a destination file path"
'        Exit Function
'    End If
'
'    If Not ValidateMonth(frm.cmbMonths.ListIndex) Then
'        ValidateSelections = "Select a month from the combobox"
'        Exit Function
'    End If
'
'    If Not ValidateAccount(frm.cmbAccounts.ListIndex) Then
'        ValidateSelections = "Select an account from the combobox"
'        Exit Function
'    End If
'
'End Function

'Public Sub OnUpdateClick(frm As UserForm)
'
'    On Error GoTo eh
'
'    Dim message As String
'    message = ValidateSelections(frm)
'
'    If message <> "" Then
'        MsgBox message
'        Exit Sub
'    End If
'
'    CollectSelections frm
'    UpdateSpending GroupName, AccountName, MonthsIndex, SourceFilePath, DestinationFilePath
'
'    Exit Sub
'
'eh:
'    DisplayError Err.Source, Err.Description, "modFormEvents.OnUpdateClick", 324
'End Sub


'Public Sub ReleaseWorkbook()
'
'    On Error GoTo eh
'
'    If Not mWorkbook Is Nothing Then
'        If FileInUse(DestinationFilePath) Then
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

'Public Sub OnUpdateClick(frm As UserForm)
'
'    On Error GoTo eh
'
'    Dim message As String
'    message = ValidateSelections(frm)
'
'    If message <> "" Then
'        MsgBox message
'        Exit Sub
'    End If
'
''    CollectSelections frm
'    UpdateSpending
'
'    Exit Sub
'
'eh:
'    DisplayError Err.Source, Err.Description, "modFormEvents.OnUpdateClick", 413
'End Sub

'Public Function ValidateSelections(frm As UserForm) As String
'    ValidateSelections = ""
'
'    If Not ValidateFilePath(frm.txtSourceFilePath) Then
'        ValidateSelections = "Select a source file path"
'        Exit Function
'    End If
'
'    If Not ValidateMonth(frm.cmbMonths.ListIndex) Then
'        ValidateSelections = "Select a month from the combobox"
'        Exit Function
'    End If
'
'    If Not ValidateAccount(frm.cmbAccounts.ListIndex) Then
'        ValidateSelections = "Select an account from the combobox"
'        Exit Function
'    End If
'
'End Function
'
' End Public Methods

'Private Function ValidateAccount(Index As Integer) As Boolean
'
'    ValidateAccount = False
'
'    If Index > 0 Then
'        ValidateAccount = True
'    End If
'
'End Function
'
'Private Function ValidateMonth(Index As Integer) As Boolean
'    ValidateMonth = False
'
'    If Index > 0 Then
'        ValidateMonth = True
'    End If
'End Function
'
' End Private Members

' Public Methods

'Public Sub CollectSelections(frm As UserForm)
'
'    mSourceFilePath = frm.txtSourceFilePath.Text
'    mDestinationFilePath = frm.txtDestinationFilePath.Text
'    mAccountName = frm.cmbAccounts.Text
'    mMonth = frm.cmbMonths.ListIndex
'    mGroupName = IIf(frm.optCategory.Value, frm.optCategory.Tag, frm.optSubCategory.Tag)
'
'End Sub

'Public Function ValidateFilePaths(frm As UserForm) As String
'    ValidateFilePaths = ""
'
'    If Not ValidateFilePath(frm.txtSourceFilePath) Then
'        ValidateFilePaths = "Select a source file path"
'        Exit Function
'    End If
'
'    If Not ValidateFilePath(frm.txtDestinationFilePath) Then
'        ValidateFilePaths = "Select a destination file path"
'        Exit Function
'    End If
'
'End Function

'Private Sub ReleaseWorkbook()
'
'    On Error GoTo eh
'
'    Dim wb As Workbook
'    Set wb = CurrentWorkbook
'
'    If Not wb Is Nothing Then
'        If FileInUse(DestinationFilePath) Then
'            wb.Close
'        End If
'    End If
'
'cleanup:
'
'    Set wb = Nothing
'    Set CurrentWorkbook = Nothing
'    Exit Sub
'
'eh:
'    RaiseError Err.Number, Err.Source, "modUpdates.ReleaseWorkbook", Err.Description, 41
'    GoTo cleanup
'End Sub

'Public Sub OnCloseClick()
'
'    On Error GoTo eh
'
'    ReleaseWorkbook
'
'    Exit Sub
'
'eh:
'    DisplayError Err.Source, Err.Description, "modFormEvents.OnCloseClick", 103
'End Sub

'Public Sub SaveAndCloseWorkbook()
'
'    On Error GoTo eh
'
'    OnSaveClick
'    ReleaseWorkbook
'
'    Exit Sub
'
'eh:
'    DisplayError Err.Source, Err.Description, "modFormEvents.SaveAndCloseWorkbook", 367
'End Sub

'Public Sub ResetLabel(frm As UserForm, Caption As String)
'
'    On Error GoTo eh
'
'    If frm.lblCreateDestinationFile.Caption <> Caption Then
'        frm.lblCreateDestinationFile.BackColor = &H8000000F
'    End If
'
'    If frm.lblClose.Caption <> Caption Then
'        frm.lblClose.BackColor = DEFAULT_BACKCOLOR
'        frm.lblCloseBackground.BackColor = DEFAULT_BACKCOLOR
'        frm.lblCloseBackground.BorderStyle = fmBorderStyleNone
'    End If
'
'    If frm.lblExit.Caption <> Caption Then
'        frm.lblExit.BackColor = DEFAULT_BACKCOLOR
'        frm.lblExitBackground.BackColor = DEFAULT_BACKCOLOR
'        frm.lblExitBackground.BorderStyle = fmBorderStyleNone
'    End If
'
'    If frm.lblNew.Caption <> Caption Then
'        frm.lblNew.BackColor = DEFAULT_BACKCOLOR
'        frm.lblNewBackground.BackColor = DEFAULT_BACKCOLOR
'        frm.lblNewBackground.BorderStyle = fmBorderStyleNone
'    End If
'
'    If frm.lblOpen.Caption <> Caption Then
'        frm.lblOpen.BackColor = DEFAULT_BACKCOLOR
'        frm.lblOpenBackground.BackColor = DEFAULT_BACKCOLOR
'        frm.lblOpenBackground.BorderStyle = fmBorderStyleNone
'    End If
'
'    If frm.lblSave.Caption <> Caption Then
'        frm.lblSave.BackColor = DEFAULT_BACKCOLOR
'        frm.lblSaveBackground.BackColor = DEFAULT_BACKCOLOR
'        frm.lblSaveBackground.BorderStyle = fmBorderStyleNone
'    End If
'
'    If frm.lblSaveAs.Caption <> Caption Then
'        frm.lblSaveAs.BackColor = DEFAULT_BACKCOLOR
'        frm.lblSaveAsBackground.BackColor = DEFAULT_BACKCOLOR
'        frm.lblSaveAsBackground.BorderStyle = fmBorderStyleNone
'    End If
'
'    If frm.lblSelectSourceFile.Caption <> Caption Then
'        frm.lblSelectSourceFile.BackColor = DEFAULT_BACKCOLOR
'    End If
'
'    If frm.lblUpdate.Caption <> Caption Then
'        frm.lblUpdate.BackColor = DEFAULT_BACKCOLOR
'    End If
'
'    Exit Sub
'
'eh:
'    DisplayError Err.Source, Err.Description, "modFormEvents.ResetLabel", 493
'End Sub

'Public Sub OnMouseMove(frm As UserForm, Caption As String)
'
'    On Error GoTo eh
'
'    Select Case Caption
'        Case "Close"
'            frm.lblClose.BackColor = HIGHLIGHT_BACKCOLOR
'            frm.lblCloseBackground.BackColor = HIGHLIGHT_BACKCOLOR
'            frm.lblCloseBackground.BorderColor = HIGHLIGHT_BORDERCOLOR
'            frm.lblCloseBackground.BorderStyle = fmBorderStyleSingle
'            ResetLabel frm, Caption
'        Case "Create Destination File"
'            frm.lblCreateDestinationFile.BackColor = HIGHLIGHT_BACKCOLOR
'        Case "Exit"
'            frm.lblExit.BackColor = HIGHLIGHT_BACKCOLOR
'            frm.lblExitBackground.BackColor = HIGHLIGHT_BACKCOLOR
'            frm.lblExitBackground.BorderColor = HIGHLIGHT_BORDERCOLOR
'            frm.lblExitBackground.BorderStyle = fmBorderStyleSingle
'            ResetLabel frm, Caption
'       Case "New"
'            frm.lblNew.BackColor = HIGHLIGHT_BACKCOLOR
'            frm.lblNewBackground.BackColor = HIGHLIGHT_BACKCOLOR
'            frm.lblNewBackground.BorderColor = HIGHLIGHT_BORDERCOLOR
'            frm.lblNewBackground.BorderStyle = fmBorderStyleSingle
'            ResetLabel frm, Caption
'       Case "Open"
'            frm.lblOpen.BackColor = HIGHLIGHT_BACKCOLOR
'            frm.lblOpenBackground.BackColor = HIGHLIGHT_BACKCOLOR
'            frm.lblOpenBackground.BorderColor = HIGHLIGHT_BORDERCOLOR
'            frm.lblOpenBackground.BorderStyle = fmBorderStyleSingle
'            ResetLabel frm, Caption
'        Case "Save"
'            frm.lblSave.BackColor = HIGHLIGHT_BACKCOLOR
'            frm.lblSaveBackground.BackColor = HIGHLIGHT_BACKCOLOR
'            frm.lblSaveBackground.BorderColor = HIGHLIGHT_BORDERCOLOR
'            frm.lblSaveBackground.BorderStyle = fmBorderStyleSingle
'            ResetLabel frm, Caption
'        Case "Save As"
'            frm.lblSaveAs.BackColor = HIGHLIGHT_BACKCOLOR
'            frm.lblSaveAsBackground.BackColor = HIGHLIGHT_BACKCOLOR
'            frm.lblSaveAsBackground.BorderColor = HIGHLIGHT_BORDERCOLOR
'            frm.lblSaveAsBackground.BorderStyle = fmBorderStyleSingle
'            ResetLabel frm, Caption
'        Case "Select"
'            frm.lblSelectSourceFile.BorderColor = HIGHLIGHT_BORDERCOLOR
'            frm.lblSelectSourceFile.BorderStyle = fmBorderStyleSingle
'        Case "Select New"
'            frm.lblSelectNewSourceFile.BackColor = HIGHLIGHT_BACKCOLOR
'        Case "Update"
'            frm.lblUpdate.BackColor = HIGHLIGHT_BACKCOLOR
'            frm.lblUpdate.BorderColor = HIGHLIGHT_BORDERCOLOR
'            frm.lblUpdate.BorderStyle = fmBorderStyleSingle
'            ResetLabel frm, Caption
'    End Select
'
'    Exit Sub
'
'eh:
'    DisplayError Err.Source, Err.Description, "modFormEvents.OnMouseMove", 211
'End Sub

'Public Function SetMainFormCaption() As String
'
'    On Error GoTo eh
'
'    SetMainFormCaption = sprintf(MAIN_FORM_CAPTION, ActiveWorkbook.Name)
'
'    Exit Function
'
'eh:
'    DisplayError Err.Source, Err.Description, "modFormEvents.SetMainFormCaption", 143
'End Function

'Public Sub OnNewClick(frm As UserForm)
'
'    On Error GoTo eh
'
'    If Not FormIsInitializing Then
'        CloseCurrentWorkbook
'    End If
'
'    frm.txtSelectNewSourceFile.Text = ""
'    frm.lblCreateDestinationFile.Enabled = False
'    frm.fraNewFile.Visible = True
'
'    Exit Sub
'
'eh:
'    DisplayError Err.Source, Err.Description, "modFormEvents.OnNewClick", 350
'End Sub


'Public Sub OnFormInitialize(frm As UserForm)
'
'    On Error GoTo eh
'
'    Set CurrentWorkbook = Workbooks.Add
'    DestinationFilePath = Application.path & Application.PathSeparator & ActiveWorkbook.FullName
'
'    frm.optCategory.Value = True
'
'    LoadCombobox frm.cboMonths, "Months"
'    frm.cboMonths.ListIndex = 0
'
'    LoadCombobox frm.cboAccounts, "Accounts"
'    frm.cboAccounts.ListIndex = 0
'
'    FormIsInitializing = True
'
'    Exit Sub
'
'eh:
'    DisplayError Err.Source, Err.Description, "modFormEvents.OnFormInitialize", 155
'End Sub

'Public Sub OnFormInitialize(frm As UserForm)
'    On Error GoTo eh
'
'    CreateNewWorkbook
'
'    frm.optCategory.Value = True
'
'    LoadCombobox frm.cboMonths, "Months"
'    frm.cboMonths.ListIndex = 0
'
'    LoadCombobox frm.cboAccounts, "Accounts"
'    frm.cboAccounts.ListIndex = 0
'
'    FormIsInitializing = True
'
'    Exit Sub
'
'eh:
'    DisplayError Err.Source, Err.Description, "modFormEvents.OnFormInitialize", 155
'End Sub

'Public Sub OnCreateDestinationFileClick(frm As UserForm)
'
'    On Error GoTo eh
'
'    CreateDestinationFile
'    OnSaveAsClick
'    FormIsInitializing = False
'
'    Exit Sub
'eh:
'    DisplayError Err.Source, Err.Description, "modFormEvents.OnCreateDestinationFileClick", 97
'End Sub

'Public Sub OnExitClick(frm As UserForm)
'
'    On Error GoTo eh
'
'    SaveAndCloseWorkbook
'
'    Unload frm
'
'    Exit Sub
'
'eh:
'    DisplayError Err.Source, Err.Description, "modFormEvents.OnExitClick", 110
'End Sub

'Public Sub OnExitClick()
'
'    On Error GoTo eh
'
'    SaveAndCloseWorkbook
'
'    Exit Sub
'
'eh:
'    DisplayError Err.Source, Err.Description, "modFormEvents.OnExitClick", 110
'End Sub
