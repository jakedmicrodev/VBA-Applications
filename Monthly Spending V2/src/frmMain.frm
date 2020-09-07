VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Main"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8415
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mExitMouseMove As Boolean

Private Sub cboAccounts_Change()

    OnAccountsChange Me
    
End Sub

Private Sub cboMonths_Change()

    OnMonthsChange Me
    
End Sub

Private Sub fraNewFile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ResetLabel Me, ""

End Sub

Private Sub lblCloseBackground_Click()

    lblClose_Click

End Sub

Private Sub lblCloseBackground_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblClose_MouseDown Button, Shift, X, Y

End Sub

Private Sub lblCloseBackground_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblClose_MouseMove Button, Shift, X, Y

End Sub

Private Sub lblCloseBackground_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblClose_MouseUp Button, Shift, X, Y

End Sub

Private Sub lblClose_Click()

    OnCloseClick Me
    Me.Caption = SetMainFormCaption()
    EnableFormControls Me
    
End Sub

Private Sub lblClose_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = True
    OnMouseDown Me, Me.lblClose.Caption

End Sub

Private Sub lblClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If mExitMouseMove Then Exit Sub
    OnMouseMove Me, Me.lblClose.Caption

End Sub

Private Sub lblClose_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = False
    OnMouseUp Me, Me.lblClose.Caption

End Sub

Private Sub lblCreateDestinationFile_Click()

    OnCreateDestinationFileClick Me
    Me.Caption = SetMainFormCaption()
    
End Sub

Private Sub lblCreateDestinationFile_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    OnMouseDown Me, Me.lblCreateDestinationFile.Caption

End Sub

Private Sub lblCreateDestinationFile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    OnMouseMove Me, Me.lblCreateDestinationFile.Caption

End Sub

Private Sub lblCreateDestinationFile_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    OnMouseUp Me, Me.lblCreateDestinationFile.Caption
    
End Sub

Private Sub lblExitBackground_Click()

    lblExit_Click

End Sub

Private Sub lblExitBackground_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblExit_MouseDown Button, Shift, X, Y

End Sub

Private Sub lblExitBackground_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblExit_MouseMove Button, Shift, X, Y

End Sub

Private Sub lblExitBackground_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblExit_MouseUp Button, Shift, X, Y

End Sub

Private Sub lblExit_Click()

    Unload Me

End Sub

Private Sub lblExit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = True
    OnMouseDown Me, Me.lblExit.Caption

End Sub

Private Sub lblExit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If mExitMouseMove Then Exit Sub
    OnMouseMove Me, Me.lblExit.Caption

End Sub

Private Sub lblExit_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = False
    OnMouseUp Me, Me.lblExit.Caption

End Sub

Private Sub lblNewBackground_Click()

    lblNew_Click

End Sub

Private Sub lblNewBackground_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblNew_MouseDown Button, Shift, X, Y

End Sub

Private Sub lblNewBackground_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblNew_MouseMove Button, Shift, X, Y

End Sub

Private Sub lblNewBackground_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblNew_MouseUp Button, Shift, X, Y

End Sub

Private Sub lblNew_Click()

    OnNewClick Me
    Me.Caption = SetMainFormCaption()
    EnableFormControls Me
    
End Sub

Private Sub lblNew_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = True
    OnMouseDown Me, Me.lblNew.Caption

End Sub

Private Sub lblNew_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If mExitMouseMove Then Exit Sub
    OnMouseMove Me, Me.lblNew.Caption

End Sub

Private Sub lblNew_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = False
    OnMouseUp Me, Me.lblNew.Caption

End Sub

Private Sub lblOpenBackground_Click()

    lblOpen_Click

End Sub

Private Sub lblOpenBackground_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblOpen_MouseDown Button, Shift, X, Y

End Sub

Private Sub lblOpenBackground_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblOpen_MouseMove Button, Shift, X, Y

End Sub

Private Sub lblOpenBackground_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblOpen_MouseUp Button, Shift, X, Y

End Sub

Private Sub lblOpen_Click()

    OnOpenClick Me
    Me.Caption = SetMainFormCaption()
    EnableFormControls Me

End Sub

Private Sub lblOpen_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = True
    OnMouseDown Me, Me.lblOpen.Caption

End Sub

Private Sub lblOpen_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If mExitMouseMove Then Exit Sub
    OnMouseMove Me, Me.lblOpen.Caption

End Sub

Private Sub lblOpen_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = False
    OnMouseUp Me, Me.lblOpen.Caption

End Sub

Private Sub lblSaveBackground_Click()

    lblSave_Click

End Sub

Private Sub lblSaveBackground_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblSave_MouseDown Button, Shift, X, Y

End Sub

Private Sub lblSaveBackground_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblSave_MouseMove Button, Shift, X, Y

End Sub

Private Sub lblSaveBackground_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblSave_MouseUp Button, Shift, X, Y

End Sub

Private Sub lblSave_Click()

    OnSaveClick

End Sub

Private Sub lblSave_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = True
    OnMouseDown Me, Me.lblSave.Caption

End Sub

Private Sub lblSave_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If mExitMouseMove Then Exit Sub
    OnMouseMove Me, Me.lblSave.Caption

End Sub

Private Sub lblSave_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = False
    OnMouseUp Me, Me.lblSave.Caption

End Sub

Private Sub lblSaveAsBackground_Click()

    lblSaveAs_Click

End Sub

Private Sub lblSaveAsBackground_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblSaveAs_MouseDown Button, Shift, X, Y

End Sub

Private Sub lblSaveAsBackground_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblSaveAs_MouseMove Button, Shift, X, Y

End Sub

Private Sub lblSaveAsBackground_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    lblSaveAs_MouseUp Button, Shift, X, Y

End Sub

Private Sub lblSaveAs_Click()

    OnSaveAsClick
    Me.Caption = SetMainFormCaption()

End Sub

Private Sub lblSaveAs_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = True
    OnMouseDown Me, Me.lblSaveAs.Caption

End Sub

Private Sub lblSaveAs_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If mExitMouseMove Then Exit Sub
    OnMouseMove Me, Me.lblSaveAs.Caption

End Sub

Private Sub lblSaveAs_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = False
    OnMouseUp Me, Me.lblSaveAs.Caption

End Sub

Private Sub lblSelectNewSourceFile_Click()

    OnSelectNewSourceFileClick Me
    
End Sub

Private Sub lblSelectNewSourceFile_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    mExitMouseMove = True
    OnMouseDown Me, Me.lblSelectNewSourceFile.Tag

End Sub

Private Sub lblSelectNewSourceFile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If mExitMouseMove Then Exit Sub
    OnMouseMove Me, Me.lblSelectNewSourceFile.Tag

End Sub

Private Sub lblSelectNewSourceFile_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    mExitMouseMove = False
    OnMouseUp Me, Me.lblSelectNewSourceFile.Tag

End Sub

Private Sub lblSelectSourceFile_Click()

    OnSelectSourceFileClick Me

End Sub

Private Sub lblSelectSourceFile_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = True
    OnMouseDown Me, Me.lblSelectSourceFile.Tag

End Sub

Private Sub lblSelectSourceFile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If mExitMouseMove Then Exit Sub
    OnMouseMove Me, "Select"

End Sub

Private Sub lblSelectSourceFile_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = False
    OnMouseUp Me, Me.lblSelectSourceFile.Tag

End Sub

Private Sub lblUpdate_Click()

    OnUpdateClick

End Sub

Private Sub lblUpdate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = True
    OnMouseDown Me, Me.lblUpdate.Caption

End Sub

Private Sub lblUpdate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If mExitMouseMove Then Exit Sub
    OnMouseMove Me, Me.lblUpdate.Caption

End Sub

Private Sub lblUpdate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    mExitMouseMove = False
    OnMouseUp Me, Me.lblUpdate.Caption

End Sub

Private Sub MultiPage1_Change()

    OnMultiPageChange Me
    
End Sub

Private Sub MultiPage1_Click(ByVal Index As Long)

    OnMultiPageChange Me
    
End Sub

Private Sub MultiPage1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ResetLabel Me, ""

End Sub

Private Sub optCategory_Click()

    OnCategoryClick Me
    
End Sub

Private Sub optSubCategory_Click()
    OnSubCategoryClick Me
End Sub

Private Sub txtSelectNewSourceFile_Change()
    
    OnSelectNewSourceFileChange Me
    
End Sub

Private Sub txtSourceFilePath_Change()

    OnSourceFilePathChange Me

End Sub

Private Sub UserForm_Activate()
    
    Me.Caption = SetMainFormCaption()
    
End Sub

Private Sub UserForm_Initialize()

    OnFormInitialize Me

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If Not IsSavedAndClosed Then
        If OnQueryClose() <> vbCancel Then
            Unload Me
        End If
    End If

End Sub
