VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSalesConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SISwitch As Integer
Dim SIType As String
Public Function validateForm() As Variant
    Dim result
    Dim flags() As String
    Dim i As Integer
    i = 0
    
    'Confirmation Date
    If IsNull(Confirmation_Date) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "Confirmation Date"
    End If
    
    'Contract Start Date
    If IsNull(Contract_Start_Date) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "Contract Start Date"
    End If
    
    'Contract End Date
    If IsNull(Contract_End_Date) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "Contract End Date"
    End If
        
    'AcctFor
    If IsNull(comboSFName) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "Account Name For"
    End If
            
    'AcctTo
    If IsNull(comboSTName) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "Account Name To"
    End If
    
    'SalesType
    If IsNull(comboSalesType) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "Sales Type"
    End If
        
    'Commodity valiation:
    strSQL = "SELECT * FROM tblTempCommodity"
    Set rst = CurrentDb.OpenRecordset(strSQL)
    If rst.RecordCount = 0 Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "No items in Commodity"
    End If
    Set rst = Nothing
    
    'Shipping, dependen de los radio
    If IsNull(comboSI) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "Shipping Options"
    End If
    If IsNull(comboSIType) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "Shipping Type"
    End If
    If SISwitch = 1 Or SISwitch = 2 Then
        If IsNull(textSIName) Then
            i = i + 1
            ReDim Preserve flags(i)
            flags(i) = "Shipping Name"
        End If
        If IsNull(textSIAddress) Then
            i = i + 1
            ReDim Preserve flags(i)
            flags(i) = "Shipping Physical Address"
        End If
        If IsNull(textSICity) Then
            i = i + 1
            ReDim Preserve flags(i)
            flags(i) = "Shipping Physical City"
        End If
        If IsNull(textSIState) Then
            i = i + 1
            ReDim Preserve flags(i)
            flags(i) = "Shipping Physical State"
        End If
        If IsNull(textSIZipCode) Then
            i = i + 1
            ReDim Preserve flags(i)
            flags(i) = "Shipping Physical Zip Code"
        End If
        If IsNull(textSICountry) Then
            i = i + 1
            ReDim Preserve flags(i)
            flags(i) = "Shipping Physical Country"
        End If
        'If IsNull(textSIAttention) Then
        '   i = i + 1
        '   ReDim Preserve flags(i)
        '   flags(i) = "Shipping Attention"
        'End If
    ElseIf (SISwitch = 3 Or SISwitch = 4) And IsNull(comboSIName) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "Shipping Name"
    End If
    
    'Shipment
    If IsNull(Text90) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "Shipment"
    End If
    If IsNull(Text92) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "Route"
    End If
    If IsNull(Text112) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "Total Commission"
    End If
    If IsNull(Text231) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "FOB"
    End If
    If IsNull(Text229) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "Terms"
    End If
    
    'Signature
    If IsNull(comboSignature) Then
        i = i + 1
        ReDim Preserve flags(i)
        flags(i) = "Signature"
    End If
    
    If i = 0 Then
        validateForm = Null
    Else
        validateForm = flags
    End If
End Function
Public Function lockFields() As Variant
    Confirmation_Date.Locked = True
    Contract_Start_Date.Locked = True
    Contract_End_Date.Locked = True
    idAcctTo.Locked = True
    comboSFName.Locked = True
    textSFAddress.Locked = True
    textSFCity.Locked = True
    textSFState.Locked = True
    textSFZipCode.Locked = True
    textSFCountry.Locked = True
    textSFAttention.Locked = True
    idAcctFor.Locked = True
    comboSTName.Locked = True
    textSTAddress.Locked = True
    textSTCity.Locked = True
    textSTState.Locked = True
    textSTZipCode.Locked = True
    textSTCountry.Locked = True
    textSTAttention.Locked = True
    'comboSIType.Locked = True
    comboSIName.Locked = True
    textSIName.Locked = True
    textSIAddress.Locked = True
    textSICity.Locked = True
    textSIState.Locked = True
    textSIZipCode.Locked = True
    textSICountry.Locked = True
    textSIAttention.Locked = True
    Text231.Locked = True
    comboCommission.Locked = True
    Text227.Locked = True
    Text229.Locked = True
    Text90.Locked = True
    Text92.Locked = True
    textNotes.Locked = True
    comboSignature.Locked = True
    
    Command215.Enabled = False
    btnReleaseEdit.Enabled = True
    btnPrint.Enabled = True
    btnPreview.Enabled = True
    btnPDF.Enabled = True
    btnExitCancel.Caption = "Exit"
End Function
Public Function unlockFields() As Variant
    'Unlock fields to edit
    Confirmation_Date.Locked = True
    Contract_Start_Date.Locked = False
    Contract_End_Date.Locked = False
    idAcctTo.Locked = False
    comboSFName.Locked = False
    textSFAddress.Locked = True
    textSFCity.Locked = True
    textSFState.Locked = True
    textSFZipCode.Locked = True
    textSFCountry.Locked = True
    textSFAttention.Locked = True
    idAcctFor.Locked = False
    comboSTName.Locked = False
    textSTAddress.Locked = True
    textSTCity.Locked = True
    textSTState.Locked = True
    textSTZipCode.Locked = True
    textSTCountry.Locked = True
    textSTAttention.Locked = True
    'comboSIType.Locked = False
    comboSIName.Locked = False
    textSIName.Locked = False
    textSIAddress.Locked = False
    textSICity.Locked = False
    textSIState.Locked = False
    textSIZipCode.Locked = False
    textSICountry.Locked = False
    textSIAttention.Locked = False
    Text231.Locked = False
    comboCommission.Locked = False
    Text227.Locked = False
    Text229.Locked = False
    Text90.Locked = False
    Text92.Locked = False
    textNotes.Locked = False
    comboSignature.Locked = False
    
    Command215.Enabled = True
    'btnReleaseEdit.Enabled = False
    btnPrint.Enabled = False
    btnPreview.Enabled = False
    btnPDF.Enabled = False
    btnSaveExit.Enabled = True
    btnSaveCR.Enabled = True
    btnExitCancel.Caption = "Cancel"
    btnReleaseEdit.Caption = "Delete"
    
End Function
Public Function deleteRelease() As Variant
    answer = MsgBox("Are you sure you want to delete this record?" & vbCrLf & "This cannot be undone", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Cancellation")
    If answer = vbYes Then
        DoCmd.SetWarnings False 'Deactivate Warnings (Confirmation on Update)
        strSQL = "UPDATE tblOrders SET Status = 0 WHERE ConfirmationNumber = '" & Text63 & "'"
        'Debug.Print strSQL
        DoCmd.RunSQL strSQL
        DoCmd.Close acForm, "frmSalesConfirmation"
    End If
End Function
Public Function saveSalesCon() As Variant
    'Set Dates
    TStamp = Now()
    ConfirmationDate = CDate(Confirmation_Date.Value)
    ContractStartDate = CDate(Contract_Start_Date.Value)
    ContractEndDate = CDate(Contract_End_Date.Value)
    
    'If ConfirmationNumber (Text63) already exists, execute an UPDATE instead of INSERT
    'NO, this will have conflicts with Balance AR, when searching is just to show info, not updating at all
    'When hitting the Edit button, check if there's a Contract Release, and if there is at least one, edit cannot be done
    strSQL2 = "SELECT * FROM tblOrders WHERE ConfirmationNumber = '" & Text63 & "'"
    Set rst2 = CurrentDb.OpenRecordset(strSQL2)
    N = rst2.RecordCount
    If N <> 0 Then 'If the Confirmation # already exists
        strSQL = "SELECT COUNT(*) AS cnt FROM tblBalanceAR WHERE ConfirmationNumber = '" & Text63 & "'"
        Set rst = CurrentDb.OpenRecordset(strSQL)
        cntReleases = rst!cnt
        If cntReleases = 1 Then 'If there's only one release (Release 0)
            
            'Update Order
            If (comboSI.Value = "Ship to") And (comboSIType.Value = "New") Then
                SIType = 1
            ElseIf (comboSI.Value = "Customer pickup") And (comboSIType.Value = "New") Then
                SIType = 2
            ElseIf (comboSI.Value = "Ship to") And (comboSIType.Value = "Existing") Then
                SIType = 3
            ElseIf (comboSI.Value = "Customer pickup") And (comboSIType.Value = "Existing") Then
                SIType = 4
            End If
            If SIType = 1 Or SIType = 2 Then
                vSIName = textSIName
            ElseIf SIType = 3 Or SIType = 4 Then
                vSIName = comboSIName
            End If
            
            sShipm = Replace(Text90, "'", "''")
            sRoute = Replace(Text92, "'", "''")
            sNotes = Replace(textNotes, "'", "''")
            DoCmd.SetWarnings False 'Deactivate Warnings (Confirmation on Update)
            'strSQL = "INSERT INTO tblOrders (TStamp, ConfirmationNumber, ConfirmationDate, ContractStartDate, ContractEndDate, SFId, SFAttention, STId, STAttention, SIType, SIName, SIAddress, SICity, SIState, SIZipCode, SICountry, SIAttention, FOB, TotalCommission, BuyerPONo, Terms, Shipment, Route, Notes, Signature, SalesType, Status) "
                    
            strSQL = "UPDATE tblOrders SET " _
            & "ContractStartDate = #" & Format(ContractStartDate, "yyyy-mm-dd") & "#, ContractEndDate = #" _
            & Format(ContractEndDate, "yyyy-mm-dd") & "#, SFId = '" & idAcctFor & "', SFAttention = '" & textSFAttention & "', STId = '" & idAcctTo & "', STAttention = '" & textSTAttention & "', SIType = '" & SIType & "', SIName ='" & vSIName & "', SIAddress = '" _
            & textSIAddress & "', SICity = '" & textSICity & "', SIState = '" & textSIState & "', SIZipCode = '" & textSIZipCode & "', SICountry = '" & textSICountry & "" _
            & "', SIAttention = '" & textSIAttention & "', FOB = '" & Text231 & "', TotalCommission = '" & comboCommission & "', BuyerPONo = '" & Text227 & "', Terms = '" & Text229 & "', Shipment = '" & sShipm & "', Route = '" & sRoute & "', Notes = '" & sNotes & "', SalesType = '" & comboSalesType & "'" _
            & " WHERE ConfirmationNumber = '" & Text63 & "'"
            'Debug.Print strSQL
            DoCmd.RunSQL strSQL
            
            'Update Commodity
            DoCmd.RunSQL "DELETE * FROM tblCommodity WHERE ConfirmationNumber = '" & Text63 & "'"
            strSQL = "SELECT * FROM tblTempCommodity"
            Set rst = CurrentDb.OpenRecordset(strSQL)
            Do Until rst.EOF
                Price = Replace(CStr(rst!PricePerPound), ",", ".")
                Total = Replace(CStr(rst!Total), ",", ".")
                sDescr = Replace(rst!Description, "'", "''")
                DoCmd.SetWarnings False 'Deactivate Warnings (Confirmation on Insert)
                'Debug.Print "INSERT INTO tblCommodity (ConfirmationNumber, Quantity, Sizes, Measurement, Pack, ProductId, Description, PricePerPound, Total) VALUES ('" & Text63.Value & "', " & rst!Quantity & ", " & rst!Sizes & ", '" & rst!Measurement & "', '" & rst!Pack & "', " & rst!ProductId & ", '" & sDescr & "', " & Price & ", " & Total & ")"
                DoCmd.RunSQL "INSERT INTO tblCommodity (ConfirmationNumber, Quantity, Sizes, Measurement, Pack, ProductId, Description, PricePerPound, Total) VALUES ('" & Text63.Value & "', " & rst!Quantity & ", " & rst!Sizes & ", '" & rst!Measurement & "', '" & rst!Pack & "', " & rst!ProductId & ", '" & sDescr & "', " & Price & ", " & Total & ")"
                DoCmd.SetWarnings True 'Activate Warnings
                rst.MoveNext
            Loop
            Set rst = Nothing
            
            'Update BalanceAR 0
            DoCmd.SetWarnings False
            DoCmd.RunSQL "DELETE * FROM tblBalanceAR WHERE ConfirmationNumber = '" & Text63 & "'"
            DoCmd.SetWarnings True
            strSQL = "SELECT * FROM tblCommodity WHERE ConfirmationNumber = '" & Text63.Value & "'"
            Set rst = CurrentDb.OpenRecordset(strSQL)
            Do Until rst.EOF
                DoCmd.SetWarnings False 'Deactivate Warnings (Confirmation on Insert)
                DoCmd.RunSQL "INSERT INTO tblBalanceAR (ReleaseDate, ConfirmationNumber, ReleaseNumber, CommodityId, AmountReleased, AmountToBeReleased) VALUES (#" & Format(TStamp, "yyyy-mm-dd") & "#, '" & Text63.Value & "', '0', " & rst!ID & ", 0, 0)"
                DoCmd.SetWarnings True 'Activate Warnings
                rst.MoveNext
            Loop
            Set rst = Nothing
        End If
    Else
        'Save Order
        If (comboSI.Value = "Ship to") And (comboSIType.Value = "New") Then
            SIType = 1
        ElseIf (comboSI.Value = "Customer pickup") And (comboSIType.Value = "New") Then
            SIType = 2
        ElseIf (comboSI.Value = "Ship to") And (comboSIType.Value = "Existing") Then
            SIType = 3
        ElseIf (comboSI.Value = "Customer pickup") And (comboSIType.Value = "Existing") Then
            SIType = 4
        End If
        If SIType = 1 Or SIType = 2 Then
            vSIName = textSIName
        ElseIf SIType = 3 Or SIType = 4 Then
            vSIName = comboSIName
        End If
        
        sShipm = Replace(Text90, "'", "''")
        sRoute = Replace(Text92, "'", "''")
        sNotes = Replace(textNotes, "'", "''")
        DoCmd.SetWarnings False 'Deactivate Warnings (Confirmation on Insert)
        strSQL = "INSERT INTO tblOrders (TStamp, ConfirmationNumber, ConfirmationDate, ContractStartDate, ContractEndDate, SFId, SFAttention, STId, STAttention, SIType, SIName, SIAddress, SICity, SIState, SIZipCode, SICountry, SIAttention, FOB, TotalCommission, BuyerPONo, Terms, Shipment, Route, Notes, Signature, SalesType, Status) "
        strSQL = strSQL & "VALUES (#" & Format(TStamp, "yyyy-mm-dd") & "#, '" & Text63.Value & "', #" & Format(ConfirmationDate, "yyyy-mm-dd") & "#, #" & Format(ContractStartDate, "yyyy-mm-dd") & "#, #" & Format(ContractEndDate, "yyyy-mm-dd") & "#, '" & idAcctFor & "', '" & textSFAttention & "', '" & idAcctTo & "', '" & textSTAttention & "', '" & SIType & "', '" & vSIName & "', '" & textSIAddress & "', '" & textSICity & "', '" & textSIState & "', '" & textSIZipCode & "', '" & textSICountry & "', '" & textSIAttention & "', '" & Text231 & "', '" & comboCommission & "', '" & Text227 & "', '" & Text229 & "', '" & sShipm & "', '" & sRoute & "', '" & sNotes & "', '" & comboSignature.Value & "', '" & comboSalesType & "', '" & 1 & "')"
        'Debug.Print strSQL
        DoCmd.RunSQL strSQL
        DoCmd.SetWarnings True 'Activate Warnings
        
        'Save Commodity
        strSQL = "SELECT * FROM tblTempCommodity"
        Set rst = CurrentDb.OpenRecordset(strSQL)
        Do Until rst.EOF
            Price = Replace(CStr(rst!PricePerPound), ",", ".")
            Total = Replace(CStr(rst!Total), ",", ".")
            sDescr = Replace(rst!Description, "'", "''")
            DoCmd.SetWarnings False 'Deactivate Warnings (Confirmation on Insert)
            DoCmd.RunSQL "INSERT INTO tblCommodity (ConfirmationNumber, Quantity, Sizes, Measurement, Pack, ProductId, Description, PricePerPound, Total) VALUES ('" & Text63.Value & "', " & rst!Quantity & ", " & rst!Sizes & ", '" & rst!Measurement & "', '" & rst!Pack & "', '" & rst!ProductId & "', '" & sDescr & "', " & Price & ", " & Total & ")"
            'Debug.Print "INSERT INTO tblCommodity (ConfirmationNumber, Quantity, Sizes, Measurement, Description, PricePerPound, Total) VALUES ('" & Text63.Value & "', " & rst!Quantity & ", " & rst!Sizes & ", '" & rst!Measurement & "', '" & rst!Description & "', " & Price & ", " & Total & ")"
            DoCmd.SetWarnings True 'Activate Warnings
            rst.MoveNext
        Loop
        Set rst = Nothing
        
        'Save BalanceAR 0
        strSQL = "SELECT * FROM tblCommodity WHERE ConfirmationNumber = '" & Text63.Value & "'"
        Set rst = CurrentDb.OpenRecordset(strSQL)
        Do Until rst.EOF
            DoCmd.SetWarnings False 'Deactivate Warnings (Confirmation on Insert)
            DoCmd.RunSQL "INSERT INTO tblBalanceAR (ReleaseDate, ConfirmationNumber, ReleaseNumber, CommodityId, AmountReleased, AmountToBeReleased) VALUES (#" & Format(TStamp, "yyyy-mm-dd") & "#, '" & Text63.Value & "', 0, " & rst!ID & ", 0, 0)"
            DoCmd.SetWarnings True 'Activate Warnings
            rst.MoveNext
        Loop
        Set rst = Nothing
    End If
    Set rst2 = Nothing
End Function
Function LineCount(ByRef Str As String) As Long
    Dim cnt As Long
    Dim Lines As Long
    Lines = 0
    LineCount = 0
    If Len(Str) = 0 Then
        Exit Function
    End If
    For cnt = 1 To Len(Str)
        Select Case Mid(Str, cnt, 1)
            Case Chr(13)
                Lines = Lines + 1
            Case "-"
                Lines = Lines + 1
        End Select
    Next cnt
    LineCount = Lines + 1
End Function
Private Sub btnReleaseEdit_Click()
    If btnReleaseEdit.Caption = "Releases" Then
        DoCmd.OpenForm "frmSearchRelease"
    ElseIf btnReleaseEdit.Caption = "Edit" Then
        unlockFields
    Else
        deleteRelease
    End If
End Sub
Private Sub Form_Load()
     Dim thisDate As Date
    btnPrint.Enabled = False
    btnPreview.Enabled = False
    btnPDF.Enabled = False
    'Create confirmation No. consecutive
    thisDate = Date
    'Order_Date = Format(thisDate, "mm/dd/yyyy")
    Order_Date = thisDate
    Dim cdThisDate
    cdThisDate = Format(thisDate, "yymmdd")
    
    Set dbs = CurrentDb
    cdLastOrder = DMax("[ConfirmationNumber]", "tblOrders")
    
    If IsNull(cdLastOrder) Then
        cdConsec = 1
    Else
        cdConsec = Int(Mid(CStr(cdLastOrder), 7, 5)) + 1
        If (Mid(CStr(cdLastOrder), 3, 4) <> Mid(CStr(cdThisDate), 3, 4)) Then 'Never different, check values... Done
            cdConsec = 0
        End If
    End If
    
    cdZeros = 2 - Len(cdConsec)
    
    Dim cdNewConsec
    cdNewConsec = ""
    For i = 1 To cdZeros
        cdNewConsec = cdNewConsec + "0"
    Next
    cdNewConsec = cdNewConsec + CStr(cdConsec)
    
    
    With comboSI
        .AddItem "Ship to"
        .AddItem "Customer pickup"
    End With
    With comboSIType
        .AddItem "New"
        .AddItem "Existing"
    End With
    
    With comboSalesType
        .AddItem "Contract"
        .AddItem "Pending"
        .AddItem "Spot Sale"
    End With
    
    With comboSignature
        .AddItem "DEBBIE ROY"
        .AddItem "TARA PESNELL"
    End With
    
    
    If CurrentProject.AllForms("frmSearch").IsLoaded = True Then 'If this form opens from search
        btnPrint.Enabled = True
        btnPreview.Enabled = True
        btnPDF.Enabled = True
        btnSaveCR.Caption = "Contract Release"
        btnSaveExit.Enabled = False
        CN = Forms!frmSearch!listConfirmation
        DoCmd.Close acForm, "frmSearch"
        Text63 = CN
        
        'Check if the current CN has any releases
        strSQL = "SELECT * FROM tblBalanceAR WHERE ConfirmationNumber = " & Chr(34) & CN & Chr(34) & " AND ReleaseNumber <> " & Chr(34) & "0" & Chr(34)
        Set rst = CurrentDb.OpenRecordset(strSQL)
        If rst.RecordCount <> 0 Then
            btnReleaseEdit.Caption = "Releases"
        Else
            btnReleaseEdit.Caption = "Edit"
        End If
        Set rst = Nothing
        
        'query all data from the CN
        Confirmation_Date = DLookup("[ConfirmationDate]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        Contract_Start_Date = DLookup("[ContractStartDate]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        Contract_End_Date = DLookup("[ContractEndDate]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        
        SFId = DLookup("[SFId]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
            idAcctFor = DLookup("[ID]", "tblAccounts", "ID = " & SFId)
            comboSFName = idAcctFor
            textSFAddress = DLookup("[PhysicalAddress]", "tblAccounts", "ID = " & SFId)
            textSFCity = DLookup("[PhysicalCity]", "tblAccounts", "ID = " & SFId)
            textSFState = DLookup("[PhysicalState]", "tblAccounts", "ID = " & SFId)
            textSFZipCode = DLookup("[PhysicalZip]", "tblAccounts", "ID = " & SFId)
            textSFCountry = DLookup("[PhysicalCountry]", "tblAccounts", "ID = " & SFId)
            textSFAttention = DLookup("[SFAttention]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        
        STId = DLookup("[STId]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
            idAcctTo = DLookup("[ID]", "tblAccounts", "ID = " & STId)
            comboSTName = idAcctTo
            textSTAddress = DLookup("[PhysicalAddress]", "tblAccounts", "ID = " & STId)
            textSTCity = DLookup("[PhysicalCity]", "tblAccounts", "ID = " & STId)
            textSTState = DLookup("[PhysicalState]", "tblAccounts", "ID = " & STId)
            textSTZipCode = DLookup("[PhysicalZip]", "tblAccounts", "ID = " & STId)
            textSTCountry = DLookup("[PhysicalCountry]", "tblAccounts", "ID = " & STId)
            textSTAttention = DLookup("[STAttention]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        SIType = DLookup("[SIType]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
            
            If SIType = 1 Then
                comboSI.Value = "Ship to"
                comboSIType.Value = "New"
            ElseIf SIType = 2 Then
                comboSI.Value = "Customer pickup"
                comboSIType.Value = "New"
            ElseIf SIType = 3 Then
                comboSI.Value = "Ship to"
                comboSIType.Value = "Existing"
            ElseIf SIType = 4 Then
                comboSI.Value = "Customer pickup"
                comboSIType.Value = "Existing"
            End If
            If SIType = 1 Or SIType = 2 Then
                textSIName = DLookup("[SIName]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
                textSIName.Visible = True
                comboSIName.Visible = False
            Else
                comboSIName = DLookup("[SIName]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
                textSIName.Visible = False
                comboSIName.Visible = True
            End If
            textSIAddress = DLookup("[SIAddress]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
            textSICity = DLookup("[SICity]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
            textSIState = DLookup("[SIState]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
            textSIZipCode = DLookup("[SIZipCode]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
            textSICountry = DLookup("[SICountry]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
            textSIAttention = DLookup("[SIAttention]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
            
        Text231 = DLookup("[FOB]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        comboCommission = DLookup("[TotalCommission]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        Text227 = DLookup("[BuyerPONo]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        Text229 = DLookup("[Terms]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        Text90 = DLookup("[Shipment]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        Text92 = DLookup("[Route]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        textNotes = DLookup("[Notes]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        comboSignature = DLookup("[Signature]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        comboSalesType = DLookup("[SalesType]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        
        DoCmd.SetWarnings False
        DoCmd.RunSQL "DELETE * FROM tblTempCommodity"
        strSQL = "SELECT * FROM tblCommodity WHERE ConfirmationNumber = " & Chr(34) & CN & Chr(34)
        Set rst = CurrentDb.OpenRecordset(strSQL)
        'convertir decimales a punto
        Do Until rst.EOF
            Price = Replace(CStr(rst!PricePerPound), ",", ".")
            Total = Replace(CStr(rst!Total), ",", ".")
            sDescr = Replace(rst!Description, "'", "''")
            'Debug.Print "INSERT INTO tblTempCommodity (Quantity, Sizes, Measurement, Pack, Description, PricePerPound, Total) VALUES (" & rst!Quantity & ", " & rst!Sizes & ", '" & rst!Measurement & "', '" & rst!Pack & "', '" & rst!Description & "', " & rst!PricePerPound & ", " & rst!Total & ")"
            DoCmd.RunSQL "INSERT INTO tblTempCommodity (Quantity, Sizes, Measurement, Pack, ProductId, Description, PricePerPound, Total) VALUES (" & rst!Quantity & ", " & rst!Sizes & ", '" & rst!Measurement & "', '" & rst!Pack & "', '" & rst!ProductId & "', '" & sDescr & "', " & Price & ", " & Total & ")"
            rst.MoveNext
        Loop
        Set rst = Nothing
        strSQL = "SELECT Id, Quantity, Sizes, Measurement, Pack, ProductId, Description, '$' & Format(PricePerPound, 'Fixed') AS Price, Total FROM tblTempCommodity"
        listCommodity.RowSource = strSQL
        Me.listCommodity.ColumnWidths = "0;1000;750;1500;1000;0;5000;1350;0"
        'Me!frmCommodity.Requery
        DoCmd.SetWarnings True
                                     
        'Lock all fields
        lockFields
    Else 'If this form does not open from search
        btnSaveExit.Caption = "Save"
        btnSaveCR.Caption = "Save and Contract Release"
        Text63 = cdThisDate & cdNewConsec 'Print confirmation No.
    
        'Verificación143.Value = False
        'radioSI1 = -1
        'radioSI2.Value = 1
        'radioSI3.Value = 1
        
        'Correct size of columns (frmCommodity)
        'Me!frmCommodity.Width = Forms!frmSalesConfirmation.Width
        'Me!frmCommodity.left = ((Forms!frmSalesConfirmation.Width / 2) - (Me!frmCommodity.Width / 2)) + 350
    
        'Me!frmCommodity![Quantity].ColumnWidth = (Me!frmCommodity.Width * 8) / 100
        'Me!frmCommodity![NumberOfUnits].ColumnWidth = (Me!frmCommodity.Width * 7) / 100
        'Me!frmCommodity![UnitOfMeasure].ColumnWidth = (Me!frmCommodity.Width * 13) / 100
        'Me!frmCommodity![Pack].ColumnWidth = (Me!frmCommodity.Width * 13) / 100
        'Me!frmCommodity![Description].ColumnWidth = (Me!frmCommodity.Width * 47) / 100
        'Me!frmCommodity![PricePerPound].ColumnWidth = (Me!frmCommodity.Width * 7) / 100
        'Me!frmCommodity![Total].ColumnWidth = (Me!frmCommodity.Width * 10) / 100
        
        'Clear tblTempCommodity (TRUNCATE) Done! Block temporarily commented
        DoCmd.SetWarnings False
        DoCmd.RunSQL "DELETE * FROM tblTempCommodity"
        'Me!frmCommodity.Requery
        Me.listCommodity.Requery
        DoCmd.SetWarnings True
        
        'Comission (Not in use)
        'strSQL = "SELECT Sum(Total) as TotalComm FROM tblTempCommodity"
        'Set rst = CurrentDb.OpenRecordset(strSQL)
        'TotalCommission = "$0.00"
        'If rst.RecordCount <> 0 Then
        '    TotalCommission = "$" & (rst!TotalComm * 0.015)
        '    Debug.Print "Total: " & rst!TotalComm
        'End If
        'Text112 = TotalCommission
        'Set rst = Nothing
    End If
End Sub
Private Sub btnSaveExit_Click()
    result = validateForm()
    If IsNull(result) Then
        saveSalesCon
        'DoCmd.Close acForm, "frmSalesConfirmation"
        btnPrint.Enabled = True
        btnPreview.Enabled = True
        btnPDF.Enabled = True
    Else
        i = UBound(result)
        Msg = "The following field(s) need to be filled: " & vbNewLine
        For j = 1 To i
            If j <> i Then
                Msg = Msg & result(j) & vbNewLine
            Else
                Msg = Msg & result(j) & "."
            End If
        Next
        MsgBox Msg
    End If
End Sub
Private Sub btnSaveCR_Click()
    result = validateForm()
    If IsNull(result) Then
        saveSalesCon
        btnPrint.Enabled = True
        btnPreview.Enabled = True
        btnPDF.Enabled = True
        DoCmd.OpenForm "frmContractRelease"
    Else
        i = UBound(result)
        Msg = "The following field(s) need to be filled: " & vbNewLine
        For j = 1 To i
            If j <> i Then
                Msg = Msg & result(j) & vbNewLine
            Else
                Msg = Msg & result(j) & "."
            End If
        Next
        MsgBox Msg
    End If
End Sub
Private Sub btnPDF_Click()
    'DoCmd.RunCommand acCmdPrintPreview
    Dim rpt As Report
    Dim lbl As Access.Label
    Const TW As Integer = 567
    If Not IsNull(textNotes) Then
        Lines = LineCount(textNotes)
    Else
        Lines = 0
    End If
    'Debug.Print Lines
    result = validateForm() 'Get response from validateForm function
    If IsNull(result) Then
        DoCmd.OpenReport "rptSalesConfirmation", acViewDesign
        'Header
        Reports!rptSalesConfirmation!labelConfirmationNumber.Caption = Text63
        
        'sConfirmationDate = CStr(Confirmation_Date.Value)
        'cdMonth = Mid(CStr(sConfirmationDate), 4, 2)
        'cdDay = Mid(CStr(sConfirmationDate), 1, 2)
        'cdYear = Mid(CStr(sConfirmationDate), 7, 4)
        'sConfirmationDate = cdMonth & "/" & cdDay & "/" & cdYear
        'ConfirmationDate = sConfirmationDate
        ConfirmationDate = Confirmation_Date.Value
                
        'sContract_Start_Date = CStr(Contract_Start_Date.Value)
        'csdMonth = Mid(CStr(sContract_Start_Date), 4, 2)
        'csdDay = Mid(CStr(sContract_Start_Date), 1, 2)
        'csdYear = Mid(CStr(sContract_Start_Date), 7, 4)
        'sContract_Start_Date = csdMonth & "/" & csdDay & "/" & csdYear
        'ContractStartDate = sContract_Start_Date
        ContractStartDate = Contract_Start_Date.Value
        
        'sContract_End_Date = CStr(Contract_End_Date.Value)
        'cedMonth = Mid(CStr(sContract_End_Date), 4, 2)
        'cedDay = Mid(CStr(sContract_End_Date), 1, 2)
        'cedYear = Mid(CStr(sContract_End_Date), 7, 4)
        'sContract_End_Date = cedMonth & "/" & cedDay & "/" & cedYear
        'ContractEndDate = sContract_End_Date
        ContractEndDate = Contract_End_Date.Value
        
        'Confirmation, Contract Start and End Dates
        Reports!rptSalesConfirmation!labelConfirmationDate.Caption = ConfirmationDate
        Reports!rptSalesConfirmation!labelStartDate.Caption = ContractStartDate
        Reports!rptSalesConfirmation!labelEndDate.Caption = ContractEndDate
        
        'Sold For labelSF
        SFName = DLookup("[AccountName]", "tblAccounts", "ID = " & comboSFName.Value)
        Reports!rptSalesConfirmation!labelSFCompanyName.Caption = SFName
        If IsNull(textSFAddress) Then
           Reports!rptSalesConfirmation!labelSFAddress1.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSFAddress1.Caption = textSFAddress
        End If
        If IsNull(textSFCity) Then
           tempSFCity = ""
        Else
           tempSFCity = textSFCity
        End If
        If IsNull(textSFState) Then
           tempSFState = ""
        Else
           tempSFState = ", " + textSFState
        End If
        If IsNull(textSFZipCode) Then
           tempSFZipCode = ""
        Else
           tempSFZipCode = " " + textSFZipCode
        End If
        Reports!rptSalesConfirmation!labelSFAddress2.Caption = tempSFCity + tempSFState + tempSFZipCode
        If IsNull(textSFCountry) Then
            Reports!rptSalesConfirmation!labelSFCountry.Caption = ""
        Else
            Reports!rptSalesConfirmation!labelSFCountry.Caption = textSFCountry
        End If
        If IsNull(textSFAttention) Then
           Reports!rptSalesConfirmation!labelSFAttention.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSFAttention.Caption = textSFAttention
        End If
        
        'Sold To labelST
        STName = DLookup("[AccountName]", "tblAccounts", "ID = " & comboSTName.Value)
        Reports!rptSalesConfirmation!labelSTCompanyName.Caption = STName
        
        If IsNull(textSTAddress) Then
           Reports!rptSalesConfirmation!labelSTAddress1.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSTAddress1.Caption = textSTAddress
        End If
        If IsNull(textSTCity) Then
           tempSTCity = ""
        Else
           tempSTCity = textSTCity
        End If
        If IsNull(textSTState) Then
           tempSTState = ""
        Else
           tempSTState = ", " + textSTState
        End If
        If IsNull(textSTZipCode) Then
           tempSTZipCode = ""
        Else
           tempSTZipCode = " " + textSTZipCode
        End If
        Reports!rptSalesConfirmation!labelSTAddress2.Caption = tempSTCity + tempSTState + tempSTZipCode
        If IsNull(textSTCountry) Then
           Reports!rptSalesConfirmation!labelSTCountry.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSTCountry.Caption = textSTCountry
        End If
        If IsNull(textSTAttention) Then
           Reports!rptSalesConfirmation!labelSTAttention.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSTAttention.Caption = textSTAttention
        End If
        
        'Commodities
        If CInt(Reports!rptSalesConfirmation!lblCountDesc.Caption) > 0 Then
            For i = 1 To CInt(Reports!rptSalesConfirmation!lblCountDesc.Caption)
                DeleteReportControl "rptSalesConfirmation", "tLabelQuantity" & i
                DeleteReportControl "rptSalesConfirmation", "tLabelCommodity" & i
                DeleteReportControl "rptSalesConfirmation", "tLabelPrice" & i
            Next i
        End If
        strSQL = "SELECT * FROM tblTempCommodity"
        Set rst = CurrentDb.OpenRecordset(strSQL)
        i = 1
        'Top = (8.711 * TW)
        Top = (0.811 * TW)
        H1 = 0.556
        Do Until rst.EOF
            'Quantity text size calculations
            If Len(CDec(rst!Quantity) & " - " & rst!Sizes & " " & rst!Measurement & " " & rst!Pack) > 30 Then
                L1 = Int(Len(rst!Quantity) / 70) + 1 'Lines required (characters mod textbox limit)
            Else
                L1 = 1
            End If
            'Commodity text size calculations
            If Len(rst!Description) > 70 Then
                L2 = Int(Len(rst!Description) / 70) + 1 'Lines required (characters mod textbox limit)
            Else
                L2 = 1
            End If
            'Measurement text size calculations
            If Len(rst!Measurement) > 20 Then
                L3 = Int(Len(rst!Measurement) / 70) + 1 'Lines required (characters mod textbox limit)
            Else
                L3 = 1
            End If
            If L1 >= L2 And L1 >= L3 Then
                L = L1
            ElseIf L2 >= L1 And L2 >= L3 Then
                L = L2
            ElseIf L3 >= L1 And L3 >= L2 Then
                L = L3
            End If
            
            H = (H1 * L)
            Height = (H * TW)
            
            Dim PricePP As Currency
            PricePP = CCur(rst!PricePerPound)
            'MsgBox (PricePP)
            Set lbl = CreateReportControl("rptSalesConfirmation", acLabel, , , CDec(rst!Quantity) & " - " & rst!Sizes & " " & rst!Measurement & " " & rst!Pack, (0.794 * TW), Top, (4.206 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelQuantity" & i
            Set lbl = CreateReportControl("rptSalesConfirmation", acLabel, , , rst!Description, (4.404 * TW), Top, (11.005 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelCommodity" & i
            Set lbl = CreateReportControl("rptSalesConfirmation", acLabel, , , "$" & Format(PricePP, "Fixed") & "/" & rst!Measurement, (17.988 * TW), Top, (3.677 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelPrice" & i
            'cb.AddItem rst!Description
            Top = Top + Height
            i = i + 1
            rst.MoveNext
        Loop
        Reports!rptSalesConfirmation!lblCountDesc.Caption = i - 1
        Set rst = Nothing
        
        If IsNull(Text227) Then
           Reports!rptSalesConfirmation!labelBuyerPONo.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelBuyerPONo.Caption = Text227
        End If
        
        Reports!rptSalesConfirmation!labelTerms.Caption = Text229
        Reports!rptSalesConfirmation!labelFOB.Caption = Text231
        Reports!rptSalesConfirmation!labelShipment.Caption = Text90
        Reports!rptSalesConfirmation!labelRoute.Caption = Text92
        
        'Shipping Info
        Reports!rptSalesConfirmation!lblAttention.Visible = True
        'If comboSIType = "Customer pickup" Then
        '    Reports!rptSalesConfirmation!labelSICompanyName.Caption = "Customer pickup"
            'Reports!rptSalesConfirmation!lblAttention.Visible = False
        'Else
            If comboSI = "Customer pickup" Then
                Reports!rptSalesConfirmation!Label22.Caption = "Customer pickup:"
            Else
                Reports!rptSalesConfirmation!Label22.Caption = "Ship to:"
            End If
            If comboSIType = "New" Then
                Reports!rptSalesConfirmation!labelSICompanyName.Caption = textSIName
                'Reports!rptSalesConfirmation!lblAttention.Visible = True
            Else
                SIName = DLookup("[AccountName]", "tblAccounts", "ID = " & comboSIName)
                Reports!rptSalesConfirmation!labelSICompanyName.Caption = SIName
                'Reports!rptSalesConfirmation!lblAttention.Visible = True
            End If
        'End If
        If IsNull(textSIAddress) Then
           Reports!rptSalesConfirmation!labelSIAddress1.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSIAddress1.Caption = textSIAddress
        End If
        If IsNull(textSICity) Then
           tempSICity = ""
        Else
           tempSICity = textSICity
        End If
        If IsNull(textSIState) Then
           tempSIState = ""
        Else
           tempSIState = ", " + textSIState
        End If
        If IsNull(textSIZipCode) Then
           tempSIZipCode = ""
        Else
           tempSIZipCode = " " + textSIZipCode
        End If
        Reports!rptSalesConfirmation!labelSIAddress2.Caption = tempSICity + tempSIState + tempSIZipCode
        If IsNull(textSICountry) Then
           Reports!rptSalesConfirmation!labelSICountry.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSICountry.Caption = textSICountry
        End If
        If IsNull(textSIAttention) Then
           Reports!rptSalesConfirmation!labelSIAttention.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSIAttention.Caption = textSIAttention
        End If
        
        'Total Commission
        Reports!rptSalesConfirmation!labelTotalCommission.Caption = comboCommission
        
        'Contract Period
        'Dim months(12) As String
        'months(1) = "JANUARY"
        'months(2) = "FEBRUARY"
        'months(3) = "MARCH"
        'months(4) = "APRIL"
        'months(5) = "MAY"
        'months(6) = "JUNE"
        'months(7) = "JULY"
        'months(8) = "AUGUST"
        'months(9) = "SEPTEMBER"
        'months(10) = "OCTOBER"
        'months(11) = "NOVEMBER"
        'months(12) = "DECEMBER"
        'MsgBox months(cedMonth)
        'Reports!rptSalesConfirmation!labelNotes.Caption = "**CONTRACT PERIOD: " & months(csdMonth) & " " & csdYear & " THROUGH " & months(cedMonth) & " " & cedDay & ", " & cedYear & " AT BUYER'S CALL"
       'THIS WAS MODIFIED TO ADJUST NOTES TO RICH TEXT IN THE SALES CONFIRMATION AND SALES CONFIRMATION REPORT
        'If IsNull(textNotes) Then
         '  Reports!rptSalesConfirmation!labelNotes.Caption = ""
        'Else
         '  Reports!rptSalesConfirmation!labelNotes.Caption = textNotes
        'End If
        
        Reports!rptSalesConfirmation!labelSignature.Caption = comboSignature.Value
        '****************************Shipping info!!
        'Show report
        'DoCmd.OpenReport "rptSalesConfirmation", acViewPreview
        'DoCmd.OutputTo acOutputReport, "rptSalesConfirmation", acFormatPDF, , , , , acExportQualityPrint
        DoCmd.OutputTo acOutputReport, "rptSalesConfirmation", acFormatPDF, , , , , acExportQualityPrint
    Else
        i = UBound(result)
        Msg = "The following field(s) need to be filled: " & vbNewLine
        For j = 1 To i
            If j <> i Then
                Msg = Msg & result(j) & vbNewLine
            Else
                Msg = Msg & result(j) & "."
            End If
        Next
        MsgBox Msg
    End If
End Sub
Private Sub btnPreview_Click()
    'DoCmd.RunCommand acCmdPrintPreview
    Dim rpt As Report
    Dim lbl As Access.Label
    Const TW As Integer = 567
    If Not IsNull(textNotes) Then
        Lines = LineCount(textNotes)
    Else
        Lines = 0
    End If
    'Debug.Print Lines
    result = validateForm() 'Get response from validateForm function
    If IsNull(result) Then
        DoCmd.OpenReport "rptSalesConfirmation", acViewDesign
        'Header
        Reports!rptSalesConfirmation!labelConfirmationNumber.Caption = Text63
        
        'sConfirmationDate = CStr(Confirmation_Date.Value)
        'cdMonth = Mid(CStr(sConfirmationDate), 4, 2)
        'cdDay = Mid(CStr(sConfirmationDate), 1, 2)
        'cdYear = Mid(CStr(sConfirmationDate), 7, 4)
        'sConfirmationDate = cdMonth & "/" & cdDay & "/" & cdYear
        'ConfirmationDate = sConfirmationDate
        ConfirmationDate = Confirmation_Date.Value
                
        'sContract_Start_Date = CStr(Contract_Start_Date.Value)
        'csdMonth = Mid(CStr(sContract_Start_Date), 4, 2)
        'csdDay = Mid(CStr(sContract_Start_Date), 1, 2)
        'csdYear = Mid(CStr(sContract_Start_Date), 7, 4)
        'sContract_Start_Date = csdMonth & "/" & csdDay & "/" & csdYear
        'ContractStartDate = sContract_Start_Date
        ContractStartDate = Contract_Start_Date.Value
        
        'sContract_End_Date = CStr(Contract_End_Date.Value)
        'cedMonth = Mid(CStr(sContract_End_Date), 4, 2)
        'cedDay = Mid(CStr(sContract_End_Date), 1, 2)
        'cedYear = Mid(CStr(sContract_End_Date), 7, 4)
        'sContract_End_Date = cedMonth & "/" & cedDay & "/" & cedYear
        'ContractEndDate = sContract_End_Date
        ContractEndDate = Contract_End_Date.Value
        
        'Confirmation, Contract Start and End Dates
        Reports!rptSalesConfirmation!labelConfirmationDate.Caption = ConfirmationDate
        Reports!rptSalesConfirmation!labelStartDate.Caption = ContractStartDate
        Reports!rptSalesConfirmation!labelEndDate.Caption = ContractEndDate
        
        'Sold For labelSF
        SFName = DLookup("[AccountName]", "tblAccounts", "ID = " & comboSFName.Value)
        Reports!rptSalesConfirmation!labelSFCompanyName.Caption = SFName
        If IsNull(textSFAddress) Then
           Reports!rptSalesConfirmation!labelSFAddress1.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSFAddress1.Caption = textSFAddress
        End If
        If IsNull(textSFCity) Then
           tempSFCity = ""
        Else
           tempSFCity = textSFCity
        End If
        If IsNull(textSFState) Then
           tempSFState = ""
        Else
           tempSFState = ", " + textSFState
        End If
        If IsNull(textSFZipCode) Then
           tempSFZipCode = ""
        Else
           tempSFZipCode = " " + textSFZipCode
        End If
        Reports!rptSalesConfirmation!labelSFAddress2.Caption = tempSFCity + tempSFState + tempSFZipCode
        If IsNull(textSFCountry) Then
            Reports!rptSalesConfirmation!labelSFCountry.Caption = ""
        Else
            Reports!rptSalesConfirmation!labelSFCountry.Caption = textSFCountry
        End If
        If IsNull(textSFAttention) Then
           Reports!rptSalesConfirmation!labelSFAttention.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSFAttention.Caption = textSFAttention
        End If
        
        'Sold To labelST
        STName = DLookup("[AccountName]", "tblAccounts", "ID = " & comboSTName.Value)
        Reports!rptSalesConfirmation!labelSTCompanyName.Caption = STName
        
        If IsNull(textSTAddress) Then
           Reports!rptSalesConfirmation!labelSTAddress1.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSTAddress1.Caption = textSTAddress
        End If
        If IsNull(textSTCity) Then
           tempSTCity = ""
        Else
           tempSTCity = textSTCity
        End If
        If IsNull(textSTState) Then
           tempSTState = ""
        Else
           tempSTState = ", " + textSTState
        End If
        If IsNull(textSTZipCode) Then
           tempSTZipCode = ""
        Else
           tempSTZipCode = " " + textSTZipCode
        End If
        Reports!rptSalesConfirmation!labelSTAddress2.Caption = tempSTCity + tempSTState + tempSTZipCode
        If IsNull(textSTCountry) Then
           Reports!rptSalesConfirmation!labelSTCountry.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSTCountry.Caption = textSTCountry
        End If
        If IsNull(textSTAttention) Then
           Reports!rptSalesConfirmation!labelSTAttention.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSTAttention.Caption = textSTAttention
        End If
        
        'Commodities
        If CInt(Reports!rptSalesConfirmation!lblCountDesc.Caption) > 0 Then
            For i = 1 To CInt(Reports!rptSalesConfirmation!lblCountDesc.Caption)
                DeleteReportControl "rptSalesConfirmation", "tLabelQuantity" & i
                DeleteReportControl "rptSalesConfirmation", "tLabelCommodity" & i
                DeleteReportControl "rptSalesConfirmation", "tLabelPrice" & i
            Next i
        End If
        strSQL = "SELECT * FROM tblTempCommodity"
        Set rst = CurrentDb.OpenRecordset(strSQL)
        i = 1
        'Top = (8.711 * TW)
        Top = (0.811 * TW)
        H1 = 0.556
        Do Until rst.EOF
            'Quantity text size calculations
            If Len(CDec(rst!Quantity) & " - " & rst!Sizes & " " & rst!Measurement & " " & rst!Pack) > 30 Then
                L1 = Int(Len(rst!Quantity) / 70) + 1 'Lines required (characters mod textbox limit)
            Else
                L1 = 1
            End If
            'Commodity text size calculations
            If Len(rst!Description) > 70 Then
                L2 = Int(Len(rst!Description) / 70) + 1 'Lines required (characters mod textbox limit)
            Else
                L2 = 1
            End If
            'Measurement text size calculations
            If Len(rst!Measurement) > 20 Then
                L3 = Int(Len(rst!Measurement) / 70) + 1 'Lines required (characters mod textbox limit)
            Else
                L3 = 1
            End If
            If L1 >= L2 And L1 >= L3 Then
                L = L1
            ElseIf L2 >= L1 And L2 >= L3 Then
                L = L2
            ElseIf L3 >= L1 And L3 >= L2 Then
                L = L3
            End If
            
            H = (H1 * L)
            Height = (H * TW)
            
            Dim PricePP As Currency
            PricePP = CCur(rst!PricePerPound)
            'MsgBox (PricePP)
            Set lbl = CreateReportControl("rptSalesConfirmation", acLabel, , , CDec(rst!Quantity) & " - " & rst!Sizes & " " & rst!Measurement & " " & rst!Pack, (0.794 * TW), Top, (4.206 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelQuantity" & i
            Set lbl = CreateReportControl("rptSalesConfirmation", acLabel, , , rst!Description, (4.404 * TW), Top, (11.005 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelCommodity" & i
            Set lbl = CreateReportControl("rptSalesConfirmation", acLabel, , , "$" & Format(PricePP, "Fixed") & "/" & rst!Measurement, (17.988 * TW), Top, (3.677 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelPrice" & i
            'cb.AddItem rst!Description
            Top = Top + Height
            i = i + 1
            rst.MoveNext
        Loop
        Reports!rptSalesConfirmation!lblCountDesc.Caption = i - 1
        Set rst = Nothing
        
        If IsNull(Text227) Then
           Reports!rptSalesConfirmation!labelBuyerPONo.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelBuyerPONo.Caption = Text227
        End If
        
        Reports!rptSalesConfirmation!labelTerms.Caption = Text229
        Reports!rptSalesConfirmation!labelFOB.Caption = Text231
        Reports!rptSalesConfirmation!labelShipment.Caption = Text90
        Reports!rptSalesConfirmation!labelRoute.Caption = Text92
        
        'Shipping Info
        Reports!rptSalesConfirmation!lblAttention.Visible = True
        'If comboSIType = "Customer pickup" Then
        '    Reports!rptSalesConfirmation!labelSICompanyName.Caption = "Customer pickup"
            'Reports!rptSalesConfirmation!lblAttention.Visible = False
        'Else
            If comboSI = "Customer pickup" Then
                Reports!rptSalesConfirmation!Label22.Caption = "Customer pickup:"
            Else
                Reports!rptSalesConfirmation!Label22.Caption = "Ship to:"
            End If
            If comboSIType = "New" Then
                Reports!rptSalesConfirmation!labelSICompanyName.Caption = textSIName
                'Reports!rptSalesConfirmation!lblAttention.Visible = True
            Else
                SIName = DLookup("[AccountName]", "tblAccounts", "ID = " & comboSIName)
                Reports!rptSalesConfirmation!labelSICompanyName.Caption = SIName
                'Reports!rptSalesConfirmation!lblAttention.Visible = True
            End If
        'End If
        If IsNull(textSIAddress) Then
           Reports!rptSalesConfirmation!labelSIAddress1.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSIAddress1.Caption = textSIAddress
        End If
        If IsNull(textSICity) Then
           tempSICity = ""
        Else
           tempSICity = textSICity
        End If
        If IsNull(textSIState) Then
           tempSIState = ""
        Else
           tempSIState = ", " + textSIState
        End If
        If IsNull(textSIZipCode) Then
           tempSIZipCode = ""
        Else
           tempSIZipCode = " " + textSIZipCode
        End If
        Reports!rptSalesConfirmation!labelSIAddress2.Caption = tempSICity + tempSIState + tempSIZipCode
        If IsNull(textSICountry) Then
           Reports!rptSalesConfirmation!labelSICountry.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSICountry.Caption = textSICountry
        End If
        If IsNull(textSIAttention) Then
           Reports!rptSalesConfirmation!labelSIAttention.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSIAttention.Caption = textSIAttention
        End If
        
        'Total Commission
        Reports!rptSalesConfirmation!labelTotalCommission.Caption = comboCommission
        
        'Contract Period
        'Dim months(12) As String
        'months(1) = "JANUARY"
        'months(2) = "FEBRUARY"
        'months(3) = "MARCH"
        'months(4) = "APRIL"
        'months(5) = "MAY"
        'months(6) = "JUNE"
        'months(7) = "JULY"
        'months(8) = "AUGUST"
        'months(9) = "SEPTEMBER"
        'months(10) = "OCTOBER"
        'months(11) = "NOVEMBER"
        'months(12) = "DECEMBER"
        'MsgBox months(cedMonth)
        'Reports!rptSalesConfirmation!labelNotes.Caption = "**CONTRACT PERIOD: " & months(csdMonth) & " " & csdYear & " THROUGH " & months(cedMonth) & " " & cedDay & ", " & cedYear & " AT BUYER'S CALL"
        'THIS WAS MODIFIED TO ADJUST NOTES TO RICH TEXT IN SALES CONF AND REPORT
        'If IsNull(textNotes) Then
        '   Reports!rptSalesConfirmation!labelNotes.Caption = ""
        'Else
        '   Reports!rptSalesConfirmation!labelNotes.Caption = textNotes
        'End If
        
        Reports!rptSalesConfirmation!labelSignature.Caption = comboSignature.Value
        '****************************Shipping info!!
        'Show report
        DoCmd.OpenReport "rptSalesConfirmation", acViewPreview
    Else
        i = UBound(result)
        Msg = "The following field(s) need to be filled: " & vbNewLine
        For j = 1 To i
            If j <> i Then
                Msg = Msg & result(j) & vbNewLine
            Else
                Msg = Msg & result(j) & "."
            End If
        Next
        MsgBox Msg
    End If
End Sub

Private Sub btnPrint_Click()
    'DoCmd.RunCommand acCmdPrintPreview
    Dim rpt As Report
    Dim lbl As Access.Label
    Const TW As Integer = 567
    If Not IsNull(textNotes) Then
        Lines = LineCount(textNotes)
    Else
        Lines = 0
    End If
    'Debug.Print Lines
    result = validateForm() 'Get response from validateForm function
    If IsNull(result) Then
        DoCmd.OpenReport "rptSalesConfirmation", acViewDesign
        'Header
        Reports!rptSalesConfirmation!labelConfirmationNumber.Caption = Text63
        
        'sConfirmationDate = CStr(Confirmation_Date.Value)
        'cdMonth = Mid(CStr(sConfirmationDate), 4, 2)
        'cdDay = Mid(CStr(sConfirmationDate), 1, 2)
        'cdYear = Mid(CStr(sConfirmationDate), 7, 4)
        'sConfirmationDate = cdMonth & "/" & cdDay & "/" & cdYear
        'ConfirmationDate = sConfirmationDate
        ConfirmationDate = Confirmation_Date.Value
                
        'sContract_Start_Date = CStr(Contract_Start_Date.Value)
        'csdMonth = Mid(CStr(sContract_Start_Date), 4, 2)
        'csdDay = Mid(CStr(sContract_Start_Date), 1, 2)
        'csdYear = Mid(CStr(sContract_Start_Date), 7, 4)
        'sContract_Start_Date = csdMonth & "/" & csdDay & "/" & csdYear
        'ContractStartDate = sContract_Start_Date
        ContractStartDate = Contract_Start_Date.Value
        
        'sContract_End_Date = CStr(Contract_End_Date.Value)
        'cedMonth = Mid(CStr(sContract_End_Date), 4, 2)
        'cedDay = Mid(CStr(sContract_End_Date), 1, 2)
        'cedYear = Mid(CStr(sContract_End_Date), 7, 4)
        'sContract_End_Date = cedMonth & "/" & cedDay & "/" & cedYear
        'ContractEndDate = sContract_End_Date
        ContractEndDate = Contract_End_Date.Value
        
        'Confirmation, Contract Start and End Dates
        Reports!rptSalesConfirmation!labelConfirmationDate.Caption = ConfirmationDate
        Reports!rptSalesConfirmation!labelStartDate.Caption = ContractStartDate
        Reports!rptSalesConfirmation!labelEndDate.Caption = ContractEndDate
        
        'Sold For labelSF
        SFName = DLookup("[AccountName]", "tblAccounts", "ID = " & comboSFName.Value)
        Reports!rptSalesConfirmation!labelSFCompanyName.Caption = SFName
        If IsNull(textSFAddress) Then
           Reports!rptSalesConfirmation!labelSFAddress1.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSFAddress1.Caption = textSFAddress
        End If
        If IsNull(textSFCity) Then
           tempSFCity = ""
        Else
           tempSFCity = textSFCity
        End If
        If IsNull(textSFState) Then
           tempSFState = ""
        Else
           tempSFState = ", " + textSFState
        End If
        If IsNull(textSFZipCode) Then
           tempSFZipCode = ""
        Else
           tempSFZipCode = " " + textSFZipCode
        End If
        Reports!rptSalesConfirmation!labelSFAddress2.Caption = tempSFCity + tempSFState + tempSFZipCode
        If IsNull(textSFCountry) Then
            Reports!rptSalesConfirmation!labelSFCountry.Caption = ""
        Else
            Reports!rptSalesConfirmation!labelSFCountry.Caption = textSFCountry
        End If
        If IsNull(textSFAttention) Then
           Reports!rptSalesConfirmation!labelSFAttention.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSFAttention.Caption = textSFAttention
        End If
        
        'Sold To labelST
        STName = DLookup("[AccountName]", "tblAccounts", "ID = " & comboSTName.Value)
        Reports!rptSalesConfirmation!labelSTCompanyName.Caption = STName
        
        If IsNull(textSTAddress) Then
           Reports!rptSalesConfirmation!labelSTAddress1.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSTAddress1.Caption = textSTAddress
        End If
        If IsNull(textSTCity) Then
           tempSTCity = ""
        Else
           tempSTCity = textSTCity
        End If
        If IsNull(textSTState) Then
           tempSTState = ""
        Else
           tempSTState = ", " + textSTState
        End If
        If IsNull(textSTZipCode) Then
           tempSTZipCode = ""
        Else
           tempSTZipCode = " " + textSTZipCode
        End If
        Reports!rptSalesConfirmation!labelSTAddress2.Caption = tempSTCity + tempSTState + tempSTZipCode
        If IsNull(textSTCountry) Then
           Reports!rptSalesConfirmation!labelSTCountry.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSTCountry.Caption = textSTCountry
        End If
        If IsNull(textSTAttention) Then
           Reports!rptSalesConfirmation!labelSTAttention.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSTAttention.Caption = textSTAttention
        End If
        
        'Commodities
        If CInt(Reports!rptSalesConfirmation!lblCountDesc.Caption) > 0 Then
            For i = 1 To CInt(Reports!rptSalesConfirmation!lblCountDesc.Caption)
                DeleteReportControl "rptSalesConfirmation", "tLabelQuantity" & i
                DeleteReportControl "rptSalesConfirmation", "tLabelCommodity" & i
                DeleteReportControl "rptSalesConfirmation", "tLabelPrice" & i
            Next i
        End If
        strSQL = "SELECT * FROM tblTempCommodity"
        Set rst = CurrentDb.OpenRecordset(strSQL)
        i = 1
        'Top = (8.711 * TW)
        Top = (0.811 * TW)
        H1 = 0.556
        Do Until rst.EOF
            'Quantity text size calculations
            If Len(CDec(rst!Quantity) & " - " & rst!Sizes & " " & rst!Measurement & " " & rst!Pack) > 30 Then
                L1 = Int(Len(rst!Quantity) / 70) + 1 'Lines required (characters mod textbox limit)
            Else
                L1 = 1
            End If
            'Commodity text size calculations
            If Len(rst!Description) > 70 Then
                L2 = Int(Len(rst!Description) / 70) + 1 'Lines required (characters mod textbox limit)
            Else
                L2 = 1
            End If
            'Measurement text size calculations
            If Len(rst!Measurement) > 20 Then
                L3 = Int(Len(rst!Measurement) / 70) + 1 'Lines required (characters mod textbox limit)
            Else
                L3 = 1
            End If
            If L1 >= L2 And L1 >= L3 Then
                L = L1
            ElseIf L2 >= L1 And L2 >= L3 Then
                L = L2
            ElseIf L3 >= L1 And L3 >= L2 Then
                L = L3
            End If
            
            H = (H1 * L)
            Height = (H * TW)
            
            Dim PricePP As Currency
            PricePP = CCur(rst!PricePerPound)
            'MsgBox (PricePP)
            Set lbl = CreateReportControl("rptSalesConfirmation", acLabel, , , CDec(rst!Quantity) & " - " & rst!Sizes & " " & rst!Measurement & " " & rst!Pack, (0.794 * TW), Top, (4.206 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelQuantity" & i
            Set lbl = CreateReportControl("rptSalesConfirmation", acLabel, , , rst!Description, (4.404 * TW), Top, (11.005 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelCommodity" & i
            Set lbl = CreateReportControl("rptSalesConfirmation", acLabel, , , "$" & Format(PricePP, "Fixed") & "/" & rst!Measurement, (17.988 * TW), Top, (3.677 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelPrice" & i
            'cb.AddItem rst!Description
            Top = Top + Height
            i = i + 1
            rst.MoveNext
        Loop
        Reports!rptSalesConfirmation!lblCountDesc.Caption = i - 1
        Set rst = Nothing
        
        If IsNull(Text227) Then
           Reports!rptSalesConfirmation!labelBuyerPONo.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelBuyerPONo.Caption = Text227
        End If
        
        Reports!rptSalesConfirmation!labelTerms.Caption = Text229
        Reports!rptSalesConfirmation!labelFOB.Caption = Text231
        Reports!rptSalesConfirmation!labelShipment.Caption = Text90
        Reports!rptSalesConfirmation!labelRoute.Caption = Text92
        
        'Shipping Info
        Reports!rptSalesConfirmation!lblAttention.Visible = True
        'If comboSIType = "Customer pickup" Then
        '    Reports!rptSalesConfirmation!labelSICompanyName.Caption = "Customer pickup"
            'Reports!rptSalesConfirmation!lblAttention.Visible = False
        'Else
            If comboSI = "Customer pickup" Then
                Reports!rptSalesConfirmation!Label22.Caption = "Customer pickup:"
            Else
                Reports!rptSalesConfirmation!Label22.Caption = "Ship to:"
            End If
            If comboSIType = "New" Then
                Reports!rptSalesConfirmation!labelSICompanyName.Caption = textSIName
                'Reports!rptSalesConfirmation!lblAttention.Visible = True
            Else
                SIName = DLookup("[AccountName]", "tblAccounts", "ID = " & comboSIName)
                Reports!rptSalesConfirmation!labelSICompanyName.Caption = SIName
                'Reports!rptSalesConfirmation!lblAttention.Visible = True
            End If
        'End If
        If IsNull(textSIAddress) Then
           Reports!rptSalesConfirmation!labelSIAddress1.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSIAddress1.Caption = textSIAddress
        End If
        If IsNull(textSICity) Then
           tempSICity = ""
        Else
           tempSICity = textSICity
        End If
        If IsNull(textSIState) Then
           tempSIState = ""
        Else
           tempSIState = ", " + textSIState
        End If
        If IsNull(textSIZipCode) Then
           tempSIZipCode = ""
        Else
           tempSIZipCode = " " + textSIZipCode
        End If
        Reports!rptSalesConfirmation!labelSIAddress2.Caption = tempSICity + tempSIState + tempSIZipCode
        If IsNull(textSICountry) Then
           Reports!rptSalesConfirmation!labelSICountry.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSICountry.Caption = textSICountry
        End If
        If IsNull(textSIAttention) Then
           Reports!rptSalesConfirmation!labelSIAttention.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelSIAttention.Caption = textSIAttention
        End If
        
        'Total Commission
        Reports!rptSalesConfirmation!labelTotalCommission.Caption = comboCommission
        
        'Contract Period
        'Dim months(12) As String
        'months(1) = "JANUARY"
        'months(2) = "FEBRUARY"
        'months(3) = "MARCH"
        'months(4) = "APRIL"
        'months(5) = "MAY"
        'months(6) = "JUNE"
        'months(7) = "JULY"
        'months(8) = "AUGUST"
        'months(9) = "SEPTEMBER"
        'months(10) = "OCTOBER"
        'months(11) = "NOVEMBER"
        'months(12) = "DECEMBER"
        'MsgBox months(cedMonth)
        'Reports!rptSalesConfirmation!labelNotes.Caption = "**CONTRACT PERIOD: " & months(csdMonth) & " " & csdYear & " THROUGH " & months(cedMonth) & " " & cedDay & ", " & cedYear & " AT BUYER'S CALL"
        If IsNull(textNotes) Then
           Reports!rptSalesConfirmation!labelNotes.Caption = ""
        Else
           Reports!rptSalesConfirmation!labelNotes.Caption = textNotes
        End If
        
        Reports!rptSalesConfirmation!labelSignature.Caption = comboSignature.Value
        '****************************Shipping info!!
        'Show report
        'DoCmd.OpenReport "rptSalesConfirmation", acViewReport
           Reports!rptSalesConfirmation!labelTotalCommission.Visible = False
           Reports!rptSalesConfirmation!Label27.Visible = False
           Reports!rptSalesConfirmation!labelSFAttention.Visible = False
           Reports!rptSalesConfirmation!Label15.Visible = False
        DoCmd.OpenReport "rptSalesConfirmation", acViewNormal
           Reports!rptSalesConfirmation!labelTotalCommission.Visible = True
           Reports!rptSalesConfirmation!Label27.Visible = True
           Reports!rptSalesConfirmation!labelSTAttention.Visible = False
           'Reports!rptSalesConfirmation!labelSTAttention.Visible = False
           Reports!rptSalesConfirmation!label16.Visible = False
           'Reports!rptSalesConfirmation!labelSTAttention.Visible = False
           Reports!rptSalesConfirmation!Label15.Visible = True
           Reports!rptSalesConfirmation!labelSFAttention.Visible = True
        DoCmd.OpenReport "rptSalesConfirmation", acViewNormal
           
    Else
        i = UBound(result)
        Msg = "The following field(s) need to be filled: " & vbNewLine
        For j = 1 To i
            If j <> i Then
                Msg = Msg & result(j) & vbNewLine
            Else
                Msg = Msg & result(j) & "."
            End If
        Next
        MsgBox Msg
    End If
End Sub

Private Sub Command215_Click()
    DoCmd.OpenForm FormName:="frmAddCommodity", WindowMode:=acDialog
    'Me!frmCommodity.SetFocus
    strSQL = "SELECT Id, Quantity, Sizes, Measurement, Pack, ProductId, Description, '$' & Format(PricePerPound, 'Fixed') AS Price, Total FROM tblTempCommodity"
    listCommodity.RowSource = strSQL
    Me.listCommodity.ColumnWidths = "0;1000;750;1500;1000;0;5000;1350;0"
    listCommodity.Requery
    strSQL = "SELECT Sum(Total) as TotalComm FROM tblTempCommodity"
    Set rst = CurrentDb.OpenRecordset(strSQL)
    TotalCommission = "$0.00"
    If rst.RecordCount <> 0 Then
        TotalCommission = "$" & (rst!TotalComm * 0.015)
    End If
    Text112 = TotalCommission
End Sub
Private Sub btnExitCancel_Click()
    If btnExitCancel.Caption = "Exit" Then
        DoCmd.Close acForm, "frmSalesConfirmation"
    Else
        answer = MsgBox("Are you sure you want to cancel?" & vbCrLf & "Changes will not be saved", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Cancellation")
        If answer = vbYes Then
            lockFields
            'recargar los datos en los campos para eliminar cualquier cambio hecho
        End If
    End If
End Sub
Private Sub comboSFName_Change()
    'AcctFor = Replace(comboSFName.Text, "'", "''")
    AcctFor = comboSFName.Text
    idAcctFor = DLookup("[ID]", "tblAccounts", "AccountName = " & Chr(34) & AcctFor & Chr(34))
    textSFAddress = DLookup("[PhysicalAddress]", "tblAccounts", "AccountName = " & Chr(34) & AcctFor & Chr(34))
    textSFCity = DLookup("[PhysicalCity]", "tblAccounts", "AccountName = " & Chr(34) & AcctFor & Chr(34))
    textSFState = DLookup("[PhysicalState]", "tblAccounts", "AccountName = " & Chr(34) & AcctFor & Chr(34))
    textSFZipCode = DLookup("[PhysicalZip]", "tblAccounts", "AccountName = " & Chr(34) & AcctFor & Chr(34))
    textSFCountry = DLookup("[PhysicalCountry]", "tblAccounts", "AccountName = " & Chr(34) & AcctFor & Chr(34))
    textSFAttention = DLookup("[FullName]", "tblContacts", "Company = " & Chr(34) & AcctFor & Chr(34))
End Sub
Private Sub comboSTName_Change()
    'AcctTo = Replace(comboSTName.Text, "'", "''")
    AcctTo = comboSTName.Text
    idAcctTo = DLookup("[ID]", "tblAccounts", "AccountName = " & Chr(34) & AcctTo & Chr(34))
    textSTAddress = DLookup("[PhysicalAddress]", "tblAccounts", "AccountName = " & Chr(34) & AcctTo & Chr(34))
    textSTCity = DLookup("[PhysicalCity]", "tblAccounts", "AccountName = " & Chr(34) & AcctTo & Chr(34))
    textSTState = DLookup("[PhysicalState]", "tblAccounts", "AccountName = " & Chr(34) & AcctTo & Chr(34))
    textSTZipCode = DLookup("[PhysicalZip]", "tblAccounts", "AccountName = " & Chr(34) & AcctTo & Chr(34))
    textSTCountry = DLookup("[PhysicalCountry]", "tblAccounts", "AccountName = " & Chr(34) & AcctTo & Chr(34))
    textSTAttention = DLookup("[FullName]", "tblContacts", "Company = " & Chr(34) & AcctTo & Chr(34))
End Sub
Private Sub Form_GotFocus()
    'Me!frmCommodity.SetFocus
    DoCmd.Requery
End Sub
Private Sub comboSI_Change()
    If IsNull(comboSIType.Value) = False Then
        SIType = comboSIType.Value
        Select Case SIType
            Case "New"
                If comboSI.Value = "Ship to" Then
                    SISwitch = 1
                Else
                    SISwitch = 2
                End If
                
                comboSIName.Value = ""
                textSIName.Value = ""
                textSIAddress.Value = ""
                textSICity.Value = ""
                textSIState.Value = ""
                textSIZipCode.Value = ""
                textSICountry.Value = ""
                textSIAttention.Value = ""
                comboSIName.Locked = False
                textSIAddress.Locked = False
                textSICity.Locked = False
                textSIState.Locked = False
                textSIZipCode.Locked = False
                textSICountry.Locked = False
                textSIAttention.Locked = False
                comboSIName.Visible = False
                textSIName.Visible = True
            Case "Existing"
                If comboSI.Value = "Ship to" Then
                    SISwitch = 3
                Else
                    SISwitch = 4
                End If
                
                comboSIName.Value = comboSTName.Value
                textSIAddress.Value = textSTAddress.Value
                textSICity.Value = textSTCity.Value
                textSIState.Value = textSTState.Value
                textSIZipCode.Value = textSTZipCode.Value
                textSICountry.Value = textSTCountry.Value
                textSIAttention.Value = textSTAttention.Value
                comboSIName.Locked = False
                textSIAddress.Locked = False
                textSICity.Locked = False
                textSIState.Locked = False
                textSIZipCode.Locked = False
                textSICountry.Locked = False
                textSIAttention.Locked = False
                comboSIName.Visible = True
                textSIName.Visible = False
        End Select
    End If
End Sub
Private Sub comboSIType_Change()
    SIType = comboSIType.Value
    Select Case SIType
        Case "New"
            If comboSI.Value = "Ship to" Then
                SISwitch = 1
            Else
                SISwitch = 2
            End If
            
            comboSIName.Value = ""
            textSIName.Value = ""
            textSIAddress.Value = ""
            textSICity.Value = ""
            textSIState.Value = ""
            textSIZipCode.Value = ""
            textSICountry.Value = ""
            textSIAttention.Value = ""
            comboSIName.Locked = False
            textSIAddress.Locked = False
            textSICity.Locked = False
            textSIState.Locked = False
            textSIZipCode.Locked = False
            textSICountry.Locked = False
            textSIAttention.Locked = False
            comboSIName.Visible = False
            textSIName.Visible = True
        Case "Existing"
            If comboSI.Value = "Ship to" Then
                SISwitch = 3
            Else
                SISwitch = 4
            End If
            
            comboSIName.Value = comboSTName.Value
            textSIAddress.Value = textSTAddress.Value
            textSICity.Value = textSTCity.Value
            textSIState.Value = textSTState.Value
            textSIZipCode.Value = textSTZipCode.Value
            textSICountry.Value = textSTCountry.Value
            textSIAttention.Value = textSTAttention.Value
            comboSIName.Locked = False
            textSIAddress.Locked = False
            textSICity.Locked = False
            textSIState.Locked = False
            textSIZipCode.Locked = False
            textSICountry.Locked = False
            textSIAttention.Locked = False
            comboSIName.Visible = True
            textSIName.Visible = False
    End Select
End Sub
Private Sub comboSIName_Change()
    ShippingAcctName = Replace(comboSIName.Text, "'", "''")
    idShipInfo = DLookup("[ID]", "tblAccounts", "AccountName = '" & ShippingAcctName & "'")
    textSIAddress = DLookup("[PhysicalAddress]", "tblAccounts", "AccountName = '" & ShippingAcctName & "'")
    textSICity = DLookup("[PhysicalCity]", "tblAccounts", "AccountName = '" & ShippingAcctName & "'")
    textSIState = DLookup("[PhysicalState]", "tblAccounts", "AccountName = '" & ShippingAcctName & "'")
    textSIZipCode = DLookup("[PhysicalZip]", "tblAccounts", "AccountName = '" & ShippingAcctName & "'")
    textSICountry = DLookup("[PhysicalCountry]", "tblAccounts", "AccountName = '" & ShippingAcctName & "'")
    textSIAttention = DLookup("[FullName]", "tblContacts", "Company = '" & ShippingAcctName & "'")
End Sub
Private Sub listCommodity_DblClick(Cancel As Integer)
    'MsgBox listCommodities
    DoCmd.OpenForm FormName:="frmUpdateCommodity", WindowMode:=acDialog
    'Forms!frmNewRelease!listCommodities.Column(3) = "aas"
    'Me!frmUpdateRelease.SetFocus
    listCommodity.Requery
    'cmdSave.Enabled = True
End Sub
Private Sub Text92_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii < 97 Or KeyAscii > 122) And (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii <> 34) And (KeyAscii <> 39) And (KeyAscii <> 44) And (KeyAscii <> 8) And (KeyAscii <> 32) Then
        KeyAscii = 0
    End If
End Sub
