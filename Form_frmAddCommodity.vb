VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAddCommodity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim ProductId As Integer
Dim Price As String
Dim PriceTest As Double
Private Sub Form_Load()
    txtConfirmation = Forms!frmSalesConfirmation!Text63
    For i = 11 To 22
        Dim cb As ComboBox
        
        Set cb = Me.Controls(i)
        If cb.ListCount > 0 Then
            For j = cb.ListCount - 1 To 0 Step -1
                cb.RemoveItem j
            Next j
        End If
        cb.Visible = False
    Next i
    With comboUOM
        '.AddItem "#"
        .AddItem "Kilo"
        .AddItem "Ounce"
        .AddItem "Pound"
    End With
    
    With comboPacks
        .AddItem "1-way fiber bins"
        .AddItem "2,000 lb. Poly Totes"
        .AddItem "Bags"
        .AddItem "Cases"
        .AddItem "Cello bags"
        .AddItem "Drums"
        .AddItem "Pails"
        .AddItem "Pounds"
        .AddItem "Super Sacks"
        .AddItem "Totes"
    End With
End Sub
Private Sub Clear_SubGroups()
    For i = 11 To 22
        Dim cb As ComboBox
        Set cb = Me.Controls(i)
        If cb.ListCount > 0 Then
            For j = cb.ListCount - 1 To 0 Step -1
                cb.RemoveItem j
            Next j
            cb.Value = ""
            cb.Visible = False
        End If
    Next i
End Sub
Private Sub Write_Description()
    txtDescription = ""
    If Not IsNull(Combo3) And Combo3 <> "-" Then
        txtDescription = txtDescription & Combo3 & " "
    End If
    If Not IsNull(Combo4) And Combo4 <> "-" Then
        txtDescription = txtDescription & Combo4 & " "
    End If
    If Not IsNull(Combo5) And Combo5 <> "-" Then
        txtDescription = txtDescription & Combo5 & " "
    End If
    If Not IsNull(Combo6) And Combo6 <> "-" Then
        txtDescription = txtDescription & Combo6 & " "
    End If
    If Not IsNull(Combo7) And Combo7 <> "-" Then
        txtDescription = txtDescription & Combo7 & " "
    End If
    If Not IsNull(Combo8) And Combo8 <> "-" Then
        txtDescription = txtDescription & Combo8 & " "
    End If
    If Not IsNull(Combo9) And Combo9 <> "-" Then
        txtDescription = txtDescription & Combo9 & " "
    End If
    If Not IsNull(Combo10) And Combo10 <> "-" Then
        txtDescription = txtDescription & Combo10 & " "
    End If
    If Not IsNull(Combo11) And Combo11 <> "-" Then
        txtDescription = txtDescription & Combo11 & " "
    End If
    If Not IsNull(Combo12) And Combo12 <> "-" Then
        txtDescription = txtDescription & Combo12 & " "
    End If
    If Not IsNull(Combo13) And Combo13 <> "-" Then
        txtDescription = txtDescription & Combo13 & " "
    End If
    If Not IsNull(Combo14) And Combo14 <> "-" Then
        txtDescription = txtDescription & Combo14 & " "
    End If
    'If Not IsNull(Combo2) Then
    '    txtDescription = txtDescription & Combo2
    '
    'End If
End Sub
Private Sub cmdAdd_Click()
    Dim flag
    flag = False
    errMsg = "The following field(s) are required" & vbCrLf
    If IsNull(Combo2) Then
        flag = True
        errMsg = errMsg & "Commodity" & vbCrLf
    End If
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
    If IsNull(txtPrice) Then
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
        'strPrice = Replace(CStr(Price), ",", ".")
        'MsgBox strPrice
        PriceTest = Format(Price, "Fixed")
        'PriceTest = Replace(CStr(PriceTest), ",", ".")
        'MsgBox Price
        Total = Replace(CStr(Total), ",", ".")
        Description = Replace(Description, "'", "''")
        'SQL = "INSERT INTO tblTempCommodity (Quantity, NumberOfUnits, UnitOfMeasure, PricePerPound, Description, Total) VALUES (" & txtQuantity & ", " & txtNum & ", " & txtUOM & ", " & Price & ", " & Description & ", " & Total & ")"
        'Debug.Print SQL
        DoCmd.SetWarnings False 'Deactivate Warnings (Confirmation on Insert)
        DoCmd.RunSQL "INSERT INTO tblTempCommodity (Quantity, Sizes, Measurement, Pack, PricePerPound, ProductId, Description, Total) VALUES (" & txtQuantity & ", " & txtNum & ", '" & UOM & "', '" & comboPacks & "', '" & PriceTest & "', " & ProductId & ", '" & Trim(Description) & "', " & Total & ")"
        DoCmd.SetWarnings True 'Activate Warnings
        
        'Show Commission (not in use)
        'strSQL = "SELECT Sum(Total) as TotalComm FROM tblTempCommodity"
        'Set rst = CurrentDb.OpenRecordset(strSQL)
        'TotalCommission = "$0.00"
        'If rst.RecordCount <> 0 Then
        '    TotalCommission = "$" & (rst!TotalComm * 0.015)
        'End If
        'Forms!frmSalesConfirmation!Text112 = TotalCommission
        
        DoCmd.Close acForm, "frmAddCommodity"
    End If
End Sub
Private Sub cmdCancel_Click()
    answer = MsgBox("Are you sure you want to cancel?" & vbCrLf & "Changes will not be saved", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Cancellation")
    If answer = vbYes Then
        DoCmd.Close acForm, "frmAddCommodity"
    End If
End Sub
Private Sub cmdReset_Click()
    answer = MsgBox("Are you sure you want to reset?", vbQuestion + vbYesNo + vbDefaultButton2, "Reset form")
    If answer = vbYes Then
        Combo2.Value = ""
        Clear_SubGroups
        txtQuantity.Value = ""
        txtNum.Value = ""
        txtUOM.Value = ""
        txtPrice.Value = ""
    End If
End Sub
Private Sub Combo2_Change()
    Clear_SubGroups
    Dim rst, rst2 As DAO.Recordset
    If IsNull(Combo2.Text) = False Then
        ProductSG = Combo2.Text
        ProductId = DLookup("[Id]", "tblProducts", "Description = '" & ProductSG & "'")
        strSQL = "SELECT COUNT(*) as cont FROM tblProductsSG2 WHERE IdProduct = " & ProductId & " GROUP BY IdSubGroup"
        Set rst = CurrentDb.OpenRecordset(strSQL)
        N = rst.RecordCount
        For i = 1 To N
            strSQL2 = "SELECT COUNT(*) as cont FROM tblProductsSG2 WHERE IdProduct = " & ProductId & " AND IdSubGroup = " & i & " GROUP BY IdSubGroup"
            Set rst2 = CurrentDb.OpenRecordset(strSQL2)
            M = rst2.RecordCount
            comboName = i + 10
            Set cb = Me.Controls(comboName)
            strSQL3 = "SELECT * FROM tblProductsSG2 WHERE IdProduct = " & ProductId & " AND IdSubGroup = " & i
            Set rst3 = CurrentDb.OpenRecordset(strSQL3)
            O = rst3.RecordCount
            cb.AddItem "-"
            Do Until rst3.EOF
                cb.AddItem rst3!Description
                rst3.MoveNext
            Loop
            cb.Visible = True
            Set cb = Nothing
            rst.MoveNext
        Next i
        Set rst = Nothing
    End If
    Write_Description
End Sub
Private Sub Combo3_Change()
    Write_Description
End Sub
Private Sub Combo4_Change()
    Write_Description
End Sub
Private Sub Combo5_Change()
    Write_Description
End Sub
Private Sub Combo6_Change()
    Write_Description
End Sub
Private Sub Combo7_Change()
    Write_Description
End Sub
Private Sub Combo8_Change()
    Write_Description
End Sub
Private Sub Combo9_Change()
    Write_Description
End Sub
Private Sub Combo10_Change()
    Write_Description
End Sub
Private Sub Combo11_Change()
    Write_Description
End Sub
Private Sub Combo12_Change()
    Write_Description
End Sub
Private Sub Combo13_Change()
    Write_Description
End Sub
Private Sub Combo14_Change()
    Write_Description
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
    txtPrice.Text = "$" & Format(Price, "Fixed")
End Sub
