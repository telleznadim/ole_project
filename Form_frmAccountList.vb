VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnClearSelected_Click()
    Set dbs = CurrentDb
    SqlQuery = "UPDATE tblAccounts SET tblAccounts.IsSelected = 0 WHERE (((tblAccounts.IsSelected)=True))"
    dbs.Execute SqlQuery, dbFailOnError
    Me.Refresh
End Sub

Private Sub btnFilter_Click()
    Dim FilterBy As String
    Dim FilterFor As String
    Dim FilterGood As Boolean
    Dim SqlQuery As String
    
    On Error Resume Next
    
    'Setup FilterGood as false to verify later if the parameters are good
    FilterGood = False
    
    'Setup empty SqlQuery
    SqlQuery = vbNullString
    
    'Test and make sure that both SearchBy and SearchFor have proper values
    'Switch FilterGood to true if all is well
    If Not IsNull(Me.cboFilterBy.Value) _
    And Me.cboFilterBy.Value <> "" _
    And Me.cboFilterBy.Value <> "No Filter" _
    And Not IsNull(Me.txtFilterFor.Value) _
    And Me.txtFilterFor.Value <> "" Then
        FilterBy = Me.cboFilterBy.Value
        FilterFor = Me.txtFilterFor.Value
        FilterGood = True
    End If
    
    'If FilterGood is true then continue with activating the filter
    If FilterGood = True Then
        Select Case FilterBy
            Case "Account Name"
                SqlQuery = "([tblAccounts].[AccountName] Like '*" & FilterFor & "*') AND "
            Case "Primary Contact"
                SqlQuery = "([Lookup_cboPrimaryContact].[FullName] Like '*" & FilterFor & "*') AND "
        End Select
        
        'Test if SqlQuery has been built and enable filter if done.
        If Not IsNull(SqlQuery) And SqlQuery <> vbNullString Then
            SqlQuery = SqlQuery & "True"
            Me.Filter = SqlQuery
            Me.FilterOn = True
        Else
            
        End If
    End If
End Sub

Private Sub btnSelectAllRecords_Click()
    Set dbs = CurrentDb
    SqlQuery = "UPDATE tblAccounts SET tblAccounts.IsSelected = 1"
    dbs.Execute SqlQuery, dbFailOnError
    Me.Refresh
End Sub

Private Sub btnShowFilter_Click()
    If Not IsNull(Me.Filter) And Me.Filter <> vbNullString Then
        MsgBox Prompt:="Filter:" & vbCrLf & Me.Filter, Buttons:=vbInformation, title:="Current Filter Settings"
    Else
        MsgBox Prompt:="No current filters active.", Buttons:=vbInformation, title:="Current Filter Settings"
    End If
End Sub

Private Sub Form_Load()

End Sub
