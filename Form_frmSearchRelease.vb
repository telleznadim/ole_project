VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSearchRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Load()
    If CurrentProject.AllForms("frmSearch").IsLoaded Then
        ReleaseId = Forms!frmSearch!listRelease
        txtReleaseId = ReleaseId
        ConfirmationNumber = DLookup("[ConfirmationNumber]", "tblTempReleases", "ID = " & ReleaseId)
    ElseIf CurrentProject.AllForms("frmSalesConfirmation").IsLoaded Then
        ConfirmationNumber = Forms!frmSalesConfirmation!Text63
    End If
    txtConfirmationNumber = ConfirmationNumber
    strSQL = "SELECT ReleaseNumber AS [Release Number], ReleaseDate AS [Release Date], ConfirmationNumber FROM tblBalanceAR WHERE ConfirmationNumber = '" & ConfirmationNumber & "' AND AmountToBeReleased <> 0 GROUP BY ReleaseNumber, ReleaseDate, ConfirmationNumber ORDER BY ReleaseNumber DESC"
    listReleases.RowSource = strSQL
    listReleases.ColumnWidths = "2500;1500;0"
End Sub

Private Sub listReleases_DblClick(Cancel As Integer)
    DoCmd.OpenForm "frmContractRelease"
End Sub
