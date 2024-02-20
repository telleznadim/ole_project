VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Load()
    Combo1.Value = ""
    Combo2.Value = ""
    txtNewProduct = ""
    btnAddProduct.Enabled = False
    txtEditProduct.Enabled = False
    txtEditProduct = ""
    btnUpdateProduct.Enabled = False
    btnDeleteProduct.Enabled = False
    
    'DoCmd.RunSQL "UPDATE tblProducts SET Status = 1"
End Sub
Private Sub Clear_SubGroups()
    'For i = 2 To ?
        'Dim cb As ComboBox
        'Set cb = Me.Controls(i)
        'If cb.ListCount > 0 Then
            For j = Combo2.ListCount - 1 To 0 Step -1
                Combo2.RemoveItem j
            Next j
            Combo2.Value = ""
        'End If
    'Next i
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo1_Change()
    Clear_SubGroups
    Dim rst, rst2 As DAO.Recordset

    Product = Combo1.Text
    If Product <> "" Then
        ProductId = DLookup("[Id]", "tblProducts", "Description = '" & Product & "'")
        strSQL = "SELECT * FROM tblProductsSG1 WHERE IdProduct = " & ProductId
        Set rst = CurrentDb.OpenRecordset(strSQL)
        'Combo2.AddItem "-"
        Do Until rst.EOF
            Combo2.AddItem rst!Description
            rst.MoveNext
        Loop
        Combo2.Enabled = True
        txtEditProduct.Enabled = True
        txtEditProduct = Combo1.Text
        'btnUpdateProduct.Enabled = True
        btnDeleteProduct.Enabled = True
        Set rst = Nothing
    End If
End Sub
Private Sub txtNewProduct_KeyPress(KeyAscii As Integer)
    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub
Private Sub txtNewProduct_KeyUp(KeyCode As Integer, Shift As Integer)
    newProduct = txtNewProduct.Text
    If newProduct <> "" Then
        btnAddProduct.Enabled = True
    Else
        btnAddProduct.Enabled = False
    End If
End Sub
Private Sub btnAddProduct_Click()
    Dim rst As DAO.Recordset
    Product = txtNewProduct.Value
    strSQL = "SELECT * FROM tblProducts WHERE Description = '" & Product & "' and Status = 1"
    Set rst = CurrentDb.OpenRecordset(strSQL)
    N = rst.RecordCount
    If N <> 0 Then
        MsgBox "This product name already exists"
        Exit Sub
    End If
    DoCmd.SetWarnings False 'Deactivate Warnings (Confirmation on Insert)
    DoCmd.RunSQL "INSERT INTO tblProducts (Description, Status) VALUES ('" & Product & "', 1)"
    DoCmd.SetWarnings True 'Activate Warnings
    MsgBox "Product " & Product & " successfully added"
    Combo1.Value = ""
    Combo2.Value = ""
    txtNewProduct = ""
    btnAddProduct.Enabled = False
    txtEditProduct.Enabled = False
    txtEditProduct = ""
    btnUpdateProduct.Enabled = False
    btnDeleteProduct.Enabled = False
    DoCmd.Requery
End Sub
Private Sub txtEditProduct_KeyUp(KeyCode As Integer, Shift As Integer)
    Product = Combo1.Value
    'Debug.Print txtEditProduct.Text & " = " & Product
    If txtEditProduct.Text <> "" Then
        btnUpdateProduct.Enabled = True
        btnDeleteProduct.Enabled = False
    Else
        btnUpdateProduct.Enabled = False
        btnDeleteProduct.Enabled = False
    End If
    If txtEditProduct.Text = Product Then
        btnUpdateProduct.Enabled = False
        btnDeleteProduct.Enabled = True
    Else
        btnUpdateProduct.Enabled = True
        btnDeleteProduct.Enabled = False
    End If
End Sub
Private Sub txtEditProduct_KeyPress(KeyAscii As Integer)
    If KeyAscii > 96 And KeyAscii < 123 Then
      KeyAscii = KeyAscii - 32
    End If
End Sub
Private Sub btnUpdateProduct_Click()
    Product = Combo1.Value
    ProductId = DLookup("[Id]", "tblProducts", "Description = '" & Product & "'")
    answer = MsgBox("Are you sure you want to update " & Product & " with " & txtEditProduct.Value & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Product")
    If answer = vbYes Then
        DoCmd.SetWarnings False 'Deactivate Warnings (Confirmation on Insert)
        DoCmd.RunSQL "UPDATE tblProducts SET Description = '" & txtEditProduct.Value & "' WHERE Id = " & ProductId
        DoCmd.SetWarnings True 'Activate Warnings
        MsgBox "Product " & txtEditProduct.Value & " successfully updated"
        Combo1.Value = ""
        Combo2.Value = ""
        txtNewProduct = ""
        btnAddProduct.Enabled = False
        txtEditProduct.Enabled = False
        txtEditProduct = ""
        btnUpdateProduct.Enabled = False
        btnDeleteProduct.Enabled = False
        DoCmd.Requery
    End If
End Sub
Private Sub btnDeleteProduct_Click()
    Product = Combo1.Value
    answer = MsgBox("Are you sure you want to delete " & Product & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Product")
    If answer = vbYes Then
        DoCmd.SetWarnings False 'Deactivate Warnings (Confirmation on Insert)
        DoCmd.RunSQL "UPDATE tblProducts SET Status = 0 WHERE Description = '" & Product & "'"
        DoCmd.SetWarnings True 'Activate Warnings
        MsgBox "Product " & Product & " successfully deleted"
        Combo1.Value = ""
        Combo2.Value = ""
        txtNewProduct = ""
        btnAddProduct.Enabled = False
        txtEditProduct.Enabled = False
        txtEditProduct = ""
        btnUpdateProduct.Enabled = False
        btnDeleteProduct.Enabled = False
        DoCmd.Requery
    End If
End Sub

