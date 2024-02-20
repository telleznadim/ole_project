VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmUpdateCommodity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim CommodityId, ProductId As Integer
Dim Price As String

Private Sub cmdDelete_Click()
    'CommodityId
    DoCmd.SetWarnings False 'Deactivate Warnings (Confirmation on Insert)
    DoCmd.RunSQL "DELETE * FROM tblTempCommodity WHERE ID = " & CommodityId
    DoCmd.SetWarnings True 'Activate Warnings
    
    DoCmd.Close acForm, "frmUpdateCommodity"
End Sub

Private Sub Form_Load()
    If CurrentProject.AllForms("frmSalesConfirmation").IsLoaded Then
        CommodityId = Forms!frmSalesConfirmation!listCommodity
        txtConfirmation = Forms!frmSalesConfirmation!Text63
               
        With comboUOM
            '.AddItem "#"
            .AddItem "Kilo"
            .AddItem "Ounce"
            .AddItem "Pound"
        End With
        
        With comboPacks
            .AddItem "1-way fiber bins"
            '.AddItem "2,000 lb. Poly Totes"
            .AddItem "Poly Totes"
            .AddItem "Bags"
            .AddItem "Cases"
            .AddItem "Cello bags"
            .AddItem "Drums"
            .AddItem "Pails"
            .AddItem "Pounds"
            .AddItem "Super Sacks"
            .AddItem "Totes"
        End With
        
        ProductId = DLookup("[ProductId]", "tblTempCommodity", "ID = " & CommodityId)
        txtDescription = DLookup("[Description]", "tblTempCommodity", "ID = " & CommodityId)
        txtQuantity = DLookup("[Quantity]", "tblTempCommodity", "ID = " & CommodityId)
        txtNum = DLookup("[Sizes]", "tblTempCommodity", "ID = " & CommodityId)
        UOM = DLookup("[Measurement]", "tblTempCommodity", "ID = " & CommodityId)
        Select Case UOM
            Case "kg"
                comboUOM = "Kilo"
            Case "oz"
                comboUOM = "Ounce"
            Case "lbs"
                comboUOM = "Pound"
        End Select
        txtPrice = DLookup("[PricePerPound]", "tblTempCommodity", "ID = " & CommodityId)
        comboPacks = DLookup("[Pack]", "tblTempCommodity", "ID = " & CommodityId)
        Price = txtPrice
        txtPrice = "$" & Price
    Else
        MsgBox "Invalid Action!"
        DoCmd.Close acForm, "frmUpdateCommodity"
    End If
End Sub
Private Sub cmdAdd_Click()
    Dim flag
    flag = False
    errMsg = "The following field(s) are required" & vbCrLf
    If IsNull(txtDescription) Then
        flag = True
        errMsg = errMsg & "Commodity Description" & vbCrLf
    End If
    If IsNull(txtQuantity) Then
        flag = True
        errMsg = errMsg & "Quantity" & vbCrLf
    Else
        If Not IsNumeric(txtQuantity) Then
            flag = True
            errMsg = errMsg & "Quantity must be a number" & vbCrLf
        Else
            If CDbl(txtQuantity) < 1 Then
                flag = True
                errMsg = errMsg & "Quantity must be at least one" & vbCrLf
            End If
        End If
    End If
    If IsNull(txtNum) Then
        flag = True
        errMsg = errMsg & "Number of Lbs/unit" & vbCrLf
    Else
        If Not IsNumeric(txtNum) Then
            flag = True
            errMsg = errMsg & "Number of Lbs/Unit must be a number" & vbCrLf
        Else
            If CDbl(txtNum) <= 0 Then
                flag = True
                errMsg = errMsg & "Number of Lbs/Unit must be greater than zero" & vbCrLf
            End If
        End If
    End If
    If IsNull(comboPacks) Then
        flag = True
        errMsg = errMsg & "Pack" & vbCrLf
    End If
    If IsNull(txtUOM) Then
        flag = True
        errMsg = errMsg & "Unit of Measure" & vbCrLf
    End If
    If IsNull(Price) Then
        flag = True
        errMsg = errMsg & "Price per Pound" & vbCrLf
    Else
        If Not IsNumeric(Price) Then
            flag = True
            errMsg = errMsg & "Price Per Pound must be a number" & vbCrLf
        Else
            If CDbl(Price) <= 0 Then
                flag = True
                errMsg = errMsg & "Price Per Pound must be greater than zero" & vbCrLf
            End If
        End If
    End If
    'flag = False
    If flag = True Then
        MsgBox errMsg
    Else
        'For i = 11 To 22
        '    Dim cb As ComboBox
        '    Set cb = Me.Controls(i)
        '    If IsNull(cb) = False Then
        '        Description = Description & cb & " "
        '    End If
        'Next i
        'Description = Description & Combo2
        Select Case comboUOM
            Case "Kilo"
                UOM = "kg"
            Case "Ounce"
                UOM = "oz"
            Case "Pound"
                UOM = "lbs"
        End Select
        Description = UCase(txtDescription)
        
        Total = txtQuantity * Price
        strPrice = Replace(CStr(CDbl(Price)), ",", ".")
        'MsgBox strPrice
        Total = Replace(CStr(Total), ",", ".")
        Description = Replace(Description, "'", "''")
        'SQL = "UPDATE tblTempCommodity SET Quantity = " & txtQuantity & ", Sizes = " & txtNum & ", Measurement = " & UOM & ", Pack = " & comboPacks & ", PricePerPound = " & Price & ", ProductId = " & ProductId & ", Description = '" & Description & "', Total = " & Total & " WHERE ID = " & CommodityId
        'Debug.Print SQL
        DoCmd.SetWarnings False 'Deactivate Warnings (Confirmation on Insert)
        DoCmd.RunSQL "UPDATE tblTempCommodity SET Quantity = " & txtQuantity & ", Sizes = " & txtNum & ", Measurement = '" & UOM & "', Pack = '" & comboPacks & "', PricePerPound = " & strPrice & ", ProductId = " & ProductId & ", Description = '" & Trim(Description) & "', Total = " & Total & " WHERE ID = " & CommodityId
        DoCmd.SetWarnings True 'Activate Warnings
        
        'Show Commission (not in use)
        'strSQL = "SELECT Sum(Total) as TotalComm FROM tblTempCommodity"
        'Set rst = CurrentDb.OpenRecordset(strSQL)
        'TotalCommission = "$0.00"
        'If rst.RecordCount <> 0 Then
        '    TotalCommission = "$" & (rst!TotalComm * 0.015)
        'End If
        'Forms!frmSalesConfirmation!Text112 = TotalCommission
        
        DoCmd.Close acForm, "frmUpdateCommodity"
    End If
End Sub
Private Sub cmdCancel_Click()
    answer = MsgBox("Are you sure you want to cancel?" & vbCrLf & "Changes will not be saved", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Cancellation")
    If answer = vbYes Then
        DoCmd.Close acForm, "frmUpdateCommodity"
    End If
End Sub
Private Sub txtNum_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 46) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtPrice_GotFocus()
    txtPrice.Text = Price
End Sub
Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 46) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtPrice_LostFocus()
    Price = txtPrice.Text
    txtPrice.Text = "$" & Price
End Sub

