VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBlackBookSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnExitForm_Click()
    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close acForm, Me.Name, acSaveYes
End Sub
