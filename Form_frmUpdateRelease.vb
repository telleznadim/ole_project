VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmUpdateRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim CommId, AmountToBeReleased As Integer
Dim Quantity As Long
Private Sub Form_Load()
    'txtConfirmationNumber = Forms!frmNewRelease!txtConfirmationNumber
    CommId = Forms!frmNewRelease!listCommodities
    strSQL = "SELECT * FROM tblTempBalanceAR WHERE ARID = " & Int(CommId)
    Set rst = CurrentDb.OpenRecordset(strSQL)
    txtCommodityId = rst!ID
    txtConfirmationNumber = rst!ConfirmationNumber
    CommodityId = rst!CommodityId
    txtCommodity = DLookup("[Description]", "tblCommodity", "ID = " & CommodityId)
    Quantity = DLookup("[Quantity]", "tblCommodity", "ID = " & CommodityId)
    txtAmountReleased = rst!AmountReleased
    AmountToBeReleased = rst!AmountToBeReleased
    txtAmountToBeReleased = AmountToBeReleased
    Set rst = Nothing
End Sub
Private Sub cmdAdd_Click()
    If txtAmountToBeReleased <> AmountToBeReleased Then
        If (Int(txtAmountToBeReleased) + Int(txtAmountReleased) > Quantity) Then
            MsgBox "Total released cannot be higher than Total sold"
        Else
            'strSQL = "SELECT * FROM tblBalanceAR WHERE ID = " & Int(CommId)
            'Set rst = CurrentDb.OpenRecordset(strSQL)
            DoCmd.SetWarnings False 'Deactivate Warnings (Confirmation on Insert)
            DoCmd.RunSQL "UPDATE tblTempBalanceAR SET AmountToBeReleased = " & txtAmountToBeReleased & " WHERE ARID = " & Int(CommId)
            DoCmd.SetWarnings True 'Activate Warnings
            'Set rst = Nothing
            MsgBox "Release Updated"
            DoCmd.Close acForm, "frmUpdateRelease"
        End If
    Else
        MsgBox "No changes made"
    End If
End Sub
Private Sub cmdCancel_Click()
    answer = MsgBox("Are you sure you want to cancel?" & vbCrLf & "Changes will not be saved", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Cancellation")
    If answer = vbYes Then
        DoCmd.Close acForm, "frmUpdateRelease"
    End If
End Sub
