VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub cmdPrintPreview_Click()
    Dim SQL As String
    Select Case Me.optSelectRpt
        Case 1 'Shipment Schedule
            'DoCmd.OpenReport "rptShipmentScheduele", acViewPreview
            DoCmd.SetWarnings False
            'DoCmd.OpenQuery "qryDeleteShipmentConfirmation"
            'DoCmd.OpenQuery "qryAddOpenConfirmations"
            'DoCmd.OpenQuery "qryAddOpenReleases"
            DoCmd.SetWarnings True
            DoCmd.OpenReport "rptShipmentscheduleNew", acViewPreview
            
        Case 2 'Total Pounds Seller
            'validar fechas de start y end
            If IsNull(FromDate) Or IsNull(ToDate) Then
                MsgBox "You must enter From and To Dates"
            Else
                DoCmd.SetWarnings False
                DoCmd.OpenQuery "qryGetOrders"
                DoCmd.OpenQuery "qrySTPR"
                DoCmd.OpenReport "rptTotalPoundsSeller", acViewPreview
                DoCmd.SetWarnings True
            End If
        Case 3 'total Pounds Buyer
            'validar fechas de start y end
            If IsNull(FromDate) Or IsNull(ToDate) Then
                MsgBox "You must enter From and To Dates"
            Else
                DoCmd.SetWarnings False
                DoCmd.OpenQuery "qryGetOrders"
                DoCmd.OpenQuery "qrySTPR" '
                DoCmd.OpenReport "rptTotalPoundsBuyer", acViewPreview
                DoCmd.SetWarnings True
            End If
            
        Case 4 'Month/Year Total Brokerage Report
            'validar fechas de start y end
            If IsNull(FromDate) Or IsNull(ToDate) Then
                MsgBox "You must enter From and To Dates"
            Else
                DoCmd.OpenReport "rptCommodityTotalPounds", acViewPreview
            End If
            
        Case 5  'Month/Year Total Brokerage Report - Accounting
            DoCmd.OpenReport "rptCommodityTotalPoundsAccounting", acViewPreview
    End Select
End Sub
Private Sub Command288_Click()
    On Error GoTo Err_Command288_Click
        DoCmd.Close
Exit_Command288_Click:
        Exit Sub
    
Err_Command288_Click:
        MsgBox Err.Description
        Resume Exit_Command288_Click
End Sub
Private Sub Option1_GotFocus()
    Me.FromDate.Enabled = False
    Me.ToDate.Enabled = False
    Me.txtSupplier.Enabled = False
    Me.txtBuyer.Enabled = True
    Me.txtContactCommodity.Enabled = False
    Me.txtContactName.Enabled = False
    Me.cmbCountry.Visible = False
    Me.cmbState.Visible = False
End Sub
Private Sub Option2_GotFocus()
    Me.FromDate.Enabled = True
    Me.ToDate.Enabled = True
    Me.txtSupplier.Enabled = True
    Me.txtBuyer.Enabled = True
    Me.txtContactCommodity.Enabled = False
    Me.txtContactName.Enabled = False
    Me.cmbCountry.Visible = False
    Me.cmbState.Visible = False
End Sub
Private Sub Option3_GotFocus()
    Me.FromDate.Enabled = True
    Me.ToDate.Enabled = True
    Me.txtSupplier.Enabled = True
    Me.txtBuyer.Enabled = True
    Me.txtContactCommodity.Enabled = False
    Me.txtContactName.Enabled = False
    Me.cmbCountry.Visible = False
    Me.cmbState.Visible = False
End Sub
Private Sub Option4_GotFocus()
    Me.FromDate.Enabled = True
    Me.ToDate.Enabled = True
    Me.txtSupplier.Enabled = True
    Me.txtBuyer.Enabled = True
    Me.txtContactCommodity.Enabled = False
    Me.txtContactName.Enabled = False
    Me.cmbCountry.Visible = False
    Me.cmbState.Visible = False
End Sub
Private Sub Option5_GotFocus()
    Me.FromDate.Enabled = True
    Me.ToDate.Enabled = True
    Me.txtSupplier.Enabled = True
    Me.txtBuyer.Enabled = True
    Me.txtContactCommodity.Enabled = False
    Me.txtContactName.Enabled = False
    Me.cmbCountry.Visible = False
    Me.cmbState.Visible = False
End Sub
'Private Sub Option10_GotFocus()
'    Me.FromDate.Enabled = True
'    Me.ToDate.Enabled = True
'    Me.txtSupplier.Enabled = True
'    Me.txtBuyer.Enabled = True
'    Me.txtContactCommodity.Enabled = False
'    Me.txtContactName.Enabled = False
'    Me.cmbCountry.Visible = False
'    Me.cmbState.Visible = False
'End Sub
'Private Sub Option2_GotFocus()
'    Me.FromDate.Enabled = True
'    Me.ToDate.Enabled = True
'    Me.txtSupplier.Enabled = False
'    Me.txtBuyer.Enabled = False
'    Me.txtContactCommodity.Enabled = False
'    Me.txtContactName.Enabled = False
'    Me.cmbCountry.Visible = False
'    Me.cmbState.Visible = False
'End Sub
'Private Sub Option8_GotFocus()
'    Me.FromDate.Enabled = True
'    Me.ToDate.Enabled = True
'    Me.txtSupplier.Enabled = True
'    Me.txtBuyer.Enabled = True
'    Me.txtContactCommodity.Enabled = True
'    Me.txtContactName.Enabled = False
'    Me.cmbCountry.RowSource = "SELECT DISTINCT BusinessCountry1 FROM tblContacts ORDER BY BusinessCountry1 ASC;"
'    Me.cmbCountry.Requery
'    Me.cmbCountry.Visible = True
'    Me.cmbState.RowSource = "SELECT DISTINCT BusinessState1 FROM tblContacts ORDER BY BusinessState1 ASC;"
'    Me.cmbState.Requery
'    Me.cmbState.Visible = True
'End Sub
'Private Sub Option3_GotFocus()
'    Me.FromDate.Enabled = True
'    Me.ToDate.Enabled = True
'    Me.txtSupplier.Enabled = True
'    Me.txtBuyer.Enabled = True
'    Me.txtContactCommodity.Enabled = True
'    Me.txtContactName.Enabled = False
'    Me.cmbCountry.RowSource = "SELECT DISTINCT BusinessCountry1 FROM tblContacts ORDER BY BusinessCountry1 ASC;"
'    Me.cmbCountry.Requery
'    Me.cmbCountry.Visible = True
'    Me.cmbState.RowSource = "SELECT DISTINCT BusinessState1 FROM tblContacts ORDER BY BusinessState1 ASC;"
'    Me.cmbState.Requery
'    Me.cmbState.Visible = True
'End Sub
'Private Sub Option4_GotFocus()
'    Me.FromDate.Enabled = False
'    Me.ToDate.Enabled = False
'    Me.txtSupplier.Enabled = False
'    Me.txtBuyer.Enabled = False
'    Me.txtContactCommodity.Enabled = False
'    Me.txtContactName.Enabled = False
'    Me.cmbCountry.Visible = False
'    Me.cmbState.Visible = False
'End Sub
'Private Sub Option5_GotFocus()
'    Me.FromDate.Enabled = False
'    Me.ToDate.Enabled = False
'    Me.txtSupplier.Enabled = True
'    Me.txtBuyer.Enabled = False
'    Me.txtContactCommodity.Enabled = False
'    Me.txtContactName.Enabled = False
'    Me.cmbCountry.Visible = False
'    Me.cmbState.Visible = False
'End Sub
'Private Sub Option7_GotFocus()
'    Me.FromDate.Enabled = False
'    Me.ToDate.Enabled = False
'    Me.txtSupplier.Enabled = False
'    Me.txtBuyer.Enabled = False
'    Me.txtContactCommodity.Enabled = True
'    Me.txtContactName.Enabled = True
'
'    Me.cmbCountry.RowSource = "SELECT DISTINCT BusinessCountry1 FROM tblContacts ORDER BY BusinessCountry1 ASC; "
'    Me.cmbCountry.Requery
'    Me.cmbCountry.Visible = True
'    Me.cmbState.RowSource = "SELECT DISTINCT BusinessState1 FROM tblContacts ORDER BY BusinessState1 ASC; "
'    Me.cmbState.Requery
'    Me.cmbState.Visible = True
'End Sub
'Private Sub Option9_GotFocus()
'    Me.FromDate.Enabled = True
'    Me.ToDate.Enabled = True
'    Me.txtSupplier.Enabled = True
'    Me.txtBuyer.Enabled = True
'    Me.txtContactCommodity.Enabled = False
'    Me.txtContactName.Enabled = False
'    Me.cmbCountry.Visible = False
'    Me.cmbState.Visible = False
'End Sub
'Private Sub Option13_GotFocus()
'    Me.FromDate.Enabled = True
'    Me.ToDate.Enabled = True
'    Me.txtSupplier.Enabled = True
'    Me.txtBuyer.Enabled = True
'    Me.txtContactCommodity.Enabled = False
'    Me.txtContactName.Enabled = False
'    Me.cmbCountry.Visible = False
'    Me.cmbState.Visible = False
'End Sub
