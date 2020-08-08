VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Create Monthly Report"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7185
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreateReport_Click()
    Dim message As String
    
    message = ValidateSelections(frmMain)
    
    If message <> "" Then
        MsgBox message
        Exit Sub
    End If
    
    CollectSelections frmMain
    
    UpdateSpending ReportType, ReportMonth, SourceFilePath, DestinationFilePath
    
End Sub

Private Sub cmdInitDestinationFile_Click()
    Dim message As String
    
    message = ValidateFilePaths(frmMain)
    If message <> "" Then
        MsgBox message
        Exit Sub
    End If
    
    cmdInitDestinationFile.Enabled = Not InitDestinationWorkbook(txtSourceFilePath.Text, txtDestinationFilePath.Text)
End Sub

Private Sub cmdSelectDestinationFile_Click()
    frmMain.txtDestinationFilePath = SelectDestinationFile
End Sub

Private Sub cmdSelectSourceFile_Click()
    frmMain.txtSourceFilePath = SelectSourceFile
End Sub

Private Sub mpTabs_Change()
    ClearSelection frmMain.mpTabs
End Sub

Private Sub UserForm_Initialize()
    LoadComboMonths frmMain.cmbMonths
    frmMain.cmbMonths.ListIndex = 0
End Sub

'Range("").SpecialCells(xlCellTypeLastCell).row
