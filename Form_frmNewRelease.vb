VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmNewRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Load()
    'Contract Release Number should exist before opening this form
    'delete data from tblTempBalanceAR
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM tblTempBalanceAR"
    DoCmd.SetWarnings True
    'test Conf # 21051600010
    txtConfirmationNumber = Forms!frmContractRelease!Text63
    'test Release # RN21051600010
    txtReleaseNumber = Forms!frmContractRelease!Text128
    strSQL = "SELECT COUNT(*) AS Cnt FROM tblCommodity WHERE ConfirmationNumber = '" & txtConfirmationNumber & "'"
    Set rst = CurrentDb.OpenRecordset(strSQL)
    topComms = rst!cnt
    Set rst = Nothing
    
    strSQL = "SELECT AR.ID AS [ARID], AR.ReleaseDate AS [ARRD], AR.ConfirmationNumber, AR.ReleaseNumber, C.Description, C.Quantity, "
    strSQL = strSQL & "(AR.AmountReleased + AR.AmountToBeReleased) AS AmountReleased, 0 AS AmountToBeReleased, AR.CommodityId AS [CommId] "
    strSQL = strSQL & "FROM tblCommodity AS C "
    strSQL = strSQL & "INNER JOIN (SELECT TOP " & topComms & " ID, ReleaseDate, ConfirmationNumber, ReleaseNumber, CommodityId, AmountReleased, AmountToBeReleased "
    strSQL = strSQL & "FROM tblBalanceAR "
    strSQL = strSQL & "WHERE ConfirmationNumber = '" & txtConfirmationNumber & "' "
    strSQL = strSQL & "ORDER BY ID DESC) AS AR "
    strSQL = strSQL & "ON AR.CommodityId = C.ID "
    strSQL = strSQL & "WHERE C.ConfirmationNumber = '" & txtConfirmationNumber & "' "
    strSQL = strSQL & "ORDER BY AR.ID ASC"
    'Debug.Print strSQL
    'send result to temp Table "tblTempBalanceAR"
    Set rst = CurrentDb.OpenRecordset(strSQL)
    Do Until rst.EOF
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO tblTempBalanceAR (ARID, ReleaseDate, ConfirmationNumber, ReleaseNumber, CommodityId, AmountReleased, AmountToBeReleased) VALUES (" & rst!ARID & ", #" & Format(rst!ARRD, "yyyy-mm-dd") & "#, '" & rst!ConfirmationNumber & "', '" & rst!ReleaseNumber & "', " & rst!CommId & ", " & rst!AmountReleased & ", " & rst!AmountToBeReleased & ")"
        DoCmd.SetWarnings True
        rst.MoveNext
    Loop
    Set rst = Nothing
    'Send Inserted rows to listbox
    SQL = "SELECT AR.ARID, C.Description AS [Description], C.Quantity AS [Initial Commodity Sold], AR.AmountReleased AS [Amount released], AR.AmountToBeReleased AS [Amount to be released] "
    SQL = SQL & "FROM tblTempBalanceAR AS AR "
    SQL = SQL & "INNER JOIN tblCommodity AS C "
    SQL = SQL & "ON AR.CommodityId = C.ID"
    Me.listCommodities.RowSource = SQL
    Me.listCommodities.ColumnWidths = "0;5000;1500;1500;1500"
    cmdSave.Enabled = False
End Sub
Private Sub listCommodities_DblClick(Cancel As Integer)
    'MsgBox listCommodities
    DoCmd.OpenForm FormName:="frmUpdateRelease", WindowMode:=acDialog
    'Forms!frmNewRelease!listCommodities.Column(3) = "aas"
    'Me!frmUpdateRelease.SetFocus
    listCommodities.Requery
    cmdSave.Enabled = True
End Sub
Private Sub cmdSave_Click()
    answer = MsgBox("Are you sure you want to save the Release?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Release")
    If answer = vbYes Then
    NewNotes = Forms!frmContractRelease!Notes
    sNotes = Replace(Notes, "'", "''")
    strSQL = "SELECT * FROM tblTempBalanceAR"
        Set rst = CurrentDb.OpenRecordset(strSQL)
        Do Until rst.EOF
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO tblBalanceAR (ReleaseDate, ConfirmationNumber, ReleaseNumber, CommodityId, AmountReleased, AmountToBeReleased, Notes) VALUES (#" & Format(Now(), "yyyy-mm-dd") & "#, '" & rst!ConfirmationNumber & "', '" & txtReleaseNumber & "', " & rst!CommodityId & ", " & rst!AmountReleased & ", " & rst!AmountToBeReleased & ", '" & sNewNotes & "')"
            DoCmd.SetWarnings True
            rst.MoveNext
        Loop
        Set rst = Nothing
        Forms!frmContractRelease!listBalanceAR.Requery
        DoCmd.Close acForm, "frmNewRelease"
    End If
End Sub
Private Sub cmdCancel_Click()
    answer = MsgBox("Are you sure you want to cancel?" & vbCrLf & "Changes will not be saved", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Cancellation")
    If answer = vbYes Then
        DoCmd.Close acForm, "frmNewRelease"
    End If
End Sub
