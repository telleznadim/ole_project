VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CN As String
Dim DR As Integer
Dim DRFrom, DRTo As Date
Public Function createQry() As Variant
    SF = Trim(comboSF.Value)
    ST = Trim(comboST.Value)
    'SF = Replace(SF, "'", "\'")
    'ST = Replace(ST, "'", "\'")
    'DR = "Date range"
    'Debug.Print SF
    'Search Salescons
    SQL = "SELECT O.ConfirmationNumber, O.ConfirmationDate, O.ContractStartDate, O.ContractEndDate, O.Status, SF.AccountName, O.SFAttention, ST.AccountName, O.STAttention"
    SQL = SQL & " FROM (tblOrders O INNER JOIN tblAccounts SF ON SF.ID = O.SFId) INNER JOIN tblAccounts ST ON O.STId = ST.ID WHERE O.Status = 1"
    'If SF <> "" Or ST <> "" Or CN <> "" Or checkDR = True Then
    If SF <> "" Or ST <> "" Or CN <> "" Then
        SQL = SQL & " AND"
        swAND = 0
        
        If SF <> "" Then
            'SQL = SQL & " SF.AccountName = " & Chr(34) & SF & Chr(34)
            SQL = SQL & " SF.AccountName = " & Chr(34) & SF & Chr(34)
            swAND = 1
        End If
        
        If ST <> "" Then
            If swAND = 1 Then
                SQL = SQL & " AND"
            End If
            SQL = SQL & " ST.AccountName = " & Chr(34) & ST & Chr(34)
            swAND = 1
        End If
        
        If CN <> "" Then
            If swAND = 1 Then
                SQL = SQL & " AND"
            End If
            SQL = SQL & " O.ConfirmationNumber LIKE '*" & CN & "*'"
            swAND = 1
        End If
    End If
    SQL = SQL & " ORDER BY O.ID DESC"
    'Debug.Print SQL
    Me.listConfirmation.RowSource = SQL
    
    'Search # of BARs
    SQL = "SELECT *"
    SQL = SQL & " FROM tblTempReleases"
    'If SF <> "" Or ST <> "" Or CN <> "" Or checkDR = True Then
    If SF <> "" Or ST <> "" Or CN <> "" Then
        SQL = SQL & " WHERE"
        swAND = 0
        
        If SF <> "" Then
            SQL = SQL & " SoldFor = " & Chr(34) & SF & Chr(34)
            swAND = 1
        End If
        
        If ST <> "" Then
            If swAND = 1 Then
                SQL = SQL & " AND"
            End If
            SQL = SQL & " SoldTo = " & Chr(34) & ST & Chr(34)
            swAND = 1
        End If
        
        If CN <> "" Then
            If swAND = 1 Then
                SQL = SQL & " AND"
            End If
            SQL = SQL & " ConfirmationNumber LIKE '*" & CN & "*'"
            swAND = 1
        End If
    End If
    SQL = SQL & " ORDER BY ConfirmationNumber DESC"
    'Debug.Print SQL
    Me.listRelease.RowSource = SQL
End Function
Public Function cleanForm() As Variant
    comboST.Value = ""
    comboSF.Value = ""
    textConfirmationNumber = ""
    checkDR = False
    textDRFrom = ""
    textDRFrom.Enabled = False
    textDRTo = ""
    textDRTo.Enabled = False
    CN = ""
    
    createQry
End Function
Private Sub Form_Load()
        If CurrentProject.AllForms("frmSalesConfirmation").IsLoaded = True Then
            DoCmd.Close acForm, "frmSalesConfirmation"
        End If
        If CurrentProject.AllForms("frmContractRelease").IsLoaded = True Then
            DoCmd.Close acForm, "frmContractRelease"
        End If
    If comboSF.ListCount > 0 Then
        For j = comboSF.ListCount - 1 To 0 Step -1
            comboSF.RemoveItem j
        Next j
    End If
    
    If comboST.ListCount > 0 Then
        For j = comboST.ListCount - 1 To 0 Step -1
            comboST.RemoveItem j
        Next j
    End If

    SQL = "SELECT O.ConfirmationNumber, O.ConfirmationDate, O.ContractStartDate, O.ContractEndDate, SF.AccountName, O.SFAttention, ST.AccountName, O.STAttention "
    SQL = SQL & "FROM (tblOrders O INNER JOIN tblAccounts SF ON SF.ID = O.SFId) INNER JOIN tblAccounts ST ON O.STId = ST.ID "
    SQL = SQL & " WHERE Status = 1 ORDER BY O.ID DESC"
    Me.listConfirmation.RowSource = SQL
        
    comboSF.AddItem ""
    strSQL = "SELECT DISTINCTROW tblAccounts.AccountName FROM tblAccounts INNER JOIN tblOrders ON tblAccounts.ID = tblOrders.SFId  WHERE Status = 1 ORDER BY tblOrders.SFId"
    Set rst = CurrentDb.OpenRecordset(strSQL)
    comboSF.Value = ""
    Do Until rst.EOF
        comboSF.AddItem (rst!AccountName)
        rst.MoveNext
    Loop
    Set rst = Nothing
    
    comboST.AddItem ""
    strSQL = "SELECT DISTINCTROW tblAccounts.AccountName FROM tblAccounts INNER JOIN tblOrders ON tblAccounts.ID = tblOrders.STId  WHERE Status = 1 ORDER BY tblOrders.STId"
    Set rst = CurrentDb.OpenRecordset(strSQL)
    comboST.Value = ""
    Do Until rst.EOF
        comboST.AddItem (rst!AccountName)
        rst.MoveNext
    Loop
    Set rst = Nothing
    
    'Clear tblTempReleases
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM tblTempReleases"
    DoCmd.SetWarnings True
    'Search all Salescon with releases
    strSQL = "SELECT [O].ConfirmationNumber AS CN, [O].ConfirmationDate AS CD, [O].ContractStartDate AS CSD, [O].ContractEndDate AS CED, SF.AccountName AS SF, ST.AccountName AS ST FROM (tblOrders AS O INNER JOIN tblAccounts AS SF ON SF.ID=[O].SFId) INNER JOIN tblAccounts AS ST ON [O].STId=ST.ID ORDER BY [O].ID DESC"
    Set rst = CurrentDb.OpenRecordset(strSQL)
    Do Until rst.EOF
        strSQL2 = "SELECT ConfirmationNumber, Count(ReleaseNumber) AS Releases FROM tblBalanceAR WHERE ConfirmationNumber = '" & rst!CN & "' AND AmountToBeReleased <> 0 GROUP BY ConfirmationNumber, ReleaseNumber"
        Set rst2 = CurrentDb.OpenRecordset(strSQL2)
        N = rst2.RecordCount
        If N <> 0 Then
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO tblTempReleases (ConfirmationNumber, ConfirmationDate, ContractStartDate, ContractEndDate, SoldFor, SoldTo, Releases) VALUES ('" & rst!CN & "', #" & Format(rst!CD, "yyyy-mm-dd") & "#, #" & Format(rst!CSD, "yyyy-mm-dd") & "#, #" & Format(rst!CED, "yyyy-mm-dd") & "#, '" & rst!SF & "', '" & rst!ST & "', " & N & ")"
            DoCmd.SetWarnings True
        End If
        Set rst2 = Nothing
        rst.MoveNext
    Loop
    Set rst = Nothing
    
    listConfirmation.Width = Me.Width - 1000
    Me!listConfirmation.Left = ((Me.Width / 2) - (Me!listConfirmation.Width / 2)) - 280
    Me.listConfirmation.ColumnWidths = "1000;1100;1100;1100;0;2000;2000;2000"
    listRelease.Width = Me.Width - 1000
    Me!listRelease.Left = ((Me.Width / 2) - (Me!listRelease.Width / 2)) - 280
    Me.listRelease.ColumnWidths = "0;1500;1100;1100;1100;2000;2000;1100"
    cleanForm
End Sub
Private Sub comboSF_Click()
    createQry
End Sub
Private Sub comboST_Click()
    createQry
End Sub

Private Sub listConfirmation_DblClick(Cancel As Integer)
    'MsgBox listConfirmation
    'DoCmd.OpenForm "frmSalesConfirmation", acViewDesign
    'Forms!frmSalesConfirmation!Text63 = listConfirmation
    DoCmd.OpenForm "frmSalesConfirmation"
End Sub

Private Sub listRelease_DblClick(Cancel As Integer)
    DoCmd.OpenForm "frmSearchRelease"
End Sub

Private Sub textConfirmationNumber_Change()
    CN = textConfirmationNumber.Text
    createQry
End Sub
Private Sub checkDR_Click()
    'If checkDR = True Then
    '    'DRFrom = Format(DateAdd("m", -1, Date), "mm/dd/yyyy")
    '    DRFrom = Format(DateAdd("m", -1, Date))
    '    DRTo = Format(DateAdd("m", 0, Date), "mm/dd/yyyy")
    '    'DRTo = Format(Date, "mm/dd/yyyy")
    '    textDRFrom = DRFrom
    '    textDRTo = DRTo
    '    textDRFrom.Enabled = True
    '    textDRTo.Enabled = True
    '
    'Else
    '    textDRFrom = ""
    '    textDRFrom.Enabled = False
    '    textDRTo = ""
    '    textDRTo.Enabled = False
    'End If
    'createQry
End Sub
Private Sub textDRFrom_Change()
    DRFrom = textDRFrom.Text
    DRTo = textDRTo.Value
    createQry
End Sub
Private Sub textDRTo_Change()
    DRFrom = textDRFrom.Value
    DRTo = textDRTo.Text
    createQry
End Sub
Private Sub cmdReset_Click()
    cleanForm
End Sub
