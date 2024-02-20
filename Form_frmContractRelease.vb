VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmContractRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public Function lockFields() As Variant
    Text63.Locked = True
    Text128.Locked = True
    Release_Date.Locked = True
    'Contract_Start_Date.Locked = True
    'Contract_End_Date.Locked = True
    textSFName.Locked = True
    textSFAddress.Locked = True
    textSFCity.Locked = True
    textSFState.Locked = True
    textSFZipCode.Locked = True
    textSFCountry.Locked = True
    textSFAttention.Locked = True
    textSTName.Locked = True
    textSTAddress.Locked = True
    textSTCity.Locked = True
    textSTState.Locked = True
    textSTZipCode.Locked = True
    textSTCountry.Locked = True
    textSTAttention.Locked = True
    comboSIType.Locked = True
    textSIName.Locked = True
    textSIAddress.Locked = True
    textSICity.Locked = True
    textSIState.Locked = True
    textSIZipCode.Locked = True
    textSICountry.Locked = True
    textSIAttention.Locked = True
    Text231.Locked = True
    Text183.Locked = True
    Text229.Locked = True
    Text90.Locked = True
    Text92.Locked = True
    'Notes.Locked = True
End Function
Private Sub Form_Load()
    'If SearchRelease or SalesCon triggers this form
    If CurrentProject.AllForms("frmSearchRelease").IsLoaded = True Then 'If frmSearchRelease triggers this form
        Command144.Enabled = False
        cmdNewRelease.Visible = False
        ID = Forms!frmSearchRelease!txtReleaseId
        RN = Forms!frmSearchRelease!listReleases
        CN = Forms!frmSearchRelease!txtConfirmationNumber
        DoCmd.Close acForm, "frmSearchRelease"
        DoCmd.Close acForm, "frmSearch"
        
        'Header
        Text63 = CN
        Text128 = RN
        
        'MsgBox DLookup("[ReleaseDate]", "tblBalanceAR", "Id = " & RN)
        strSQL = "SELECT ReleaseDate, Notes FROM tblBalanceAR WHERE ConfirmationNumber = '" & CN & "' AND ReleaseNumber = '" & RN & "'"
        Set rst = CurrentDb.OpenRecordset(strSQL)
        Release_Date = rst!ReleaseDate
        'sNotes = Replace(Notes, "'", "''")
        Notes = rst!Notes
        Set rst = Nothing
        'Release_Date = DLookup("[ReleaseDate]", "tblBalanceAR", "Id = " & RN)
        'Confirmation_Date = DLookup("[ConfirmationDate]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        'Contract_Start_Date = DLookup("[ContractStartDate]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        'Contract_End_Date = DLookup("[ContractEndDate]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        
        'Get Acct For Info
        idAcFor = DLookup("[SFId]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        textSFName = DLookup("[AccountName]", "tblAccounts", "ID = " & idAcFor)
        textSFAddress = DLookup("[PhysicalAddress]", "tblAccounts", "ID = " & idAcFor)
        textSFCity = DLookup("[PhysicalCity]", "tblAccounts", "ID = " & idAcFor)
        textSFState = DLookup("[PhysicalState]", "tblAccounts", "ID = " & idAcFor)
        textSFZipCode = DLookup("[PhysicalZip]", "tblAccounts", "ID = " & idAcFor)
        textSFCountry = DLookup("[PhysicalCountry]", "tblAccounts", "ID = " & idAcFor)
        textSFAttention = DLookup("[SFAttention]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        
        'Get Acct To Info
        idAcTo = DLookup("[STId]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        textSTName = DLookup("[AccountName]", "tblAccounts", "ID = " & idAcTo)
        textSTAddress = DLookup("[PhysicalAddress]", "tblAccounts", "ID = " & idAcTo)
        textSTCity = DLookup("[PhysicalCity]", "tblAccounts", "ID = " & idAcTo)
        textSTState = DLookup("[PhysicalState]", "tblAccounts", "ID = " & idAcTo)
        textSTZipCode = DLookup("[PhysicalZip]", "tblAccounts", "ID = " & idAcTo)
        textSTCountry = DLookup("[PhysicalCountry]", "tblAccounts", "ID = " & idAcTo)
        textSTAttention = DLookup("[STAttention]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        
        'Get Shipping Info
        SIType = DLookup("[SIType]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        Select Case SIType
            Case 1
                comboSI = "Ship to"
                comboSIType = "New"
            Case 2
                comboSI = "Ship to"
                comboSIType = "Existing"
            Case 3
                comboSI = "Customer pickup"
                comboSIType = "New"
            Case 4
                comboSI = "Customer pickup"
                comboSIType = "Existing"
        End Select
        
        If SIType = 1 Or SIType = 2 Then
            textSIName = DLookup("[SIName]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        Else
            idShipTo = DLookup("[SIName]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
            textSIName = DLookup("[AccountName]", "tblAccounts", "ID = " & idShipTo)
        End If
        textSIAddress = DLookup("[SIAddress]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        textSICity = DLookup("[SICity]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        textSIState = DLookup("[SIState]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        textSIZipCode = DLookup("[SIZipCode]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        textSICountry = DLookup("[SICountry]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        textSIAttention = DLookup("[SIAttention]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        
        
        Text183 = DLookup("[BuyerPONo]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        'If IsNull(Text227) Then
        '   Reports!rptSalesConfirmation!labelBuyerPONo.Caption = ""
        'Else
        '   Reports!rptSalesConfirmation!labelBuyerPONo.Caption = Text227
        'End If
        Text229 = DLookup("[Terms]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        Text231 = DLookup("[FOB]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        Text90 = DLookup("[Shipment]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        Text92 = DLookup("[Route]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        'Notes = DLookup("[Notes]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        
        'Debug.Print "SELECT AR.AmountToBeReleased & C.Measurement AS [Amount To Be Released], C.Description AS [Description], C.PricePerPound & '/' & C.Pack AS [Price], (C.Quantity - (AR.AmountReleased + AR.AmountToBeReleased)) & C.Measurement AS [Balance After release] FROM tblBalanceAR AS AR INNER JOIN tblCommodity AS C ON C.ID = AR.CommodityId WHERE AR.ConfirmationNumber = '" & CN & "' AND AR.ReleaseNumber = '" & RN & "' AND AR.AmountToBeReleased <> 0 ORDER BY AR.ID DESC"
        SQLBAR = "SELECT AR.AmountToBeReleased & '-' & C.Sizes & ' ' & C.Measurement & ' ' & C.Pack AS [Amount To Be Released], C.Description AS [Description], C.PricePerPound & '/' & C.Measurement AS [Price], (C.Quantity - (AR.AmountReleased + AR.AmountToBeReleased)) & C.Measurement AS [Balance After release] FROM tblBalanceAR AS AR INNER JOIN tblCommodity AS C ON C.ID = AR.CommodityId WHERE AR.ConfirmationNumber = '" & CN & "' AND AR.ReleaseNumber = '" & RN & "' AND AR.AmountToBeReleased <> 0 ORDER BY AR.ID DESC"
        listBalanceAR.RowSource = SQLBAR
        Me.listBalanceAR.ColumnWidths = "2350;5000;1200;1200;1200"
        lockFields
        '***********************************************************************************************************
    ElseIf CurrentProject.AllForms("frmSalesConfirmation").IsLoaded = True Then
        'Filling data from prev form
        Command144.Enabled = True
        cmdNewRelease.Visible = True
        Text63 = Forms!frmSalesConfirmation!Text63
        Release_Date = Date
        'Confirmation_Date = Forms!frmSalesConfirmation!Confirmation_Date
        'Contract_Start_Date = Forms!frmSalesConfirmation!Contract_Start_Date
        'Contract_End_Date = Forms!frmSalesConfirmation!Contract_End_Date
        
        'idAcFor = Forms!frmSalesConfirmation!idAcctFor
        'idAcTo = Forms!frmSalesConfirmation!idAcctTo
        
        'Get Acct For Info
        idAcFor = Forms!frmSalesConfirmation!comboSFName
        textSFName = DLookup("[AccountName]", "tblAccounts", "ID = " & idAcFor)
        textSFAddress = DLookup("[PhysicalAddress]", "tblAccounts", "ID = " & idAcFor)
        textSFCity = DLookup("[PhysicalCity]", "tblAccounts", "ID = " & idAcFor)
        textSFState = DLookup("[PhysicalState]", "tblAccounts", "ID = " & idAcFor)
        textSFZipCode = DLookup("[PhysicalZip]", "tblAccounts", "ID = " & idAcFor)
        textSFCountry = DLookup("[PhysicalCountry]", "tblAccounts", "ID = " & idAcFor)
        textSFAttention = Forms!frmSalesConfirmation!textSFAttention
        
        'Get Acct To Info
        idAcTo = Forms!frmSalesConfirmation!comboSTName
        textSTName = DLookup("[AccountName]", "tblAccounts", "ID = " & idAcTo)
        textSTAddress = DLookup("[PhysicalAddress]", "tblAccounts", "ID = " & idAcTo)
        textSTCity = DLookup("[PhysicalCity]", "tblAccounts", "ID = " & idAcTo)
        textSTState = DLookup("[PhysicalState]", "tblAccounts", "ID = " & idAcTo)
        textSTZipCode = DLookup("[PhysicalZip]", "tblAccounts", "ID = " & idAcTo)
        textSTCountry = DLookup("[PhysicalCountry]", "tblAccounts", "ID = " & idAcTo)
        textSTAttention = Forms!frmSalesConfirmation!textSTAttention
        
        'Get Shipping Info
        SIType = DLookup("[SIType]", "tblOrders", "ConfirmationNumber = " & Chr(34) & Text63 & Chr(34))
        Select Case SIType
            Case 1
                comboSI = "Ship to"
                comboSIType = "New"
            Case 2
                comboSI = "Ship to"
                comboSIType = "Existing"
            Case 3
                comboSI = "Customer pickup"
                comboSIType = "New"
            Case 4
                comboSI = "Customer pickup"
                comboSIType = "Existing"
        End Select
        
        If SIType = 1 Or SIType = 2 Then
            textSIName = DLookup("[SIName]", "tblOrders", "ConfirmationNumber = " & Chr(34) & Text63 & Chr(34))
        Else
            idShipTo = DLookup("[SIName]", "tblOrders", "ConfirmationNumber = " & Chr(34) & Text63 & Chr(34))
            textSIName = DLookup("[AccountName]", "tblAccounts", "ID = " & idShipTo)
        End If
        textSIAddress = Forms!frmSalesConfirmation!textSIAddress
        textSICity = Forms!frmSalesConfirmation!textSICity
        textSIState = Forms!frmSalesConfirmation!textSIState
        textSIZipCode = Forms!frmSalesConfirmation!textSIZipCode
        textSICountry = Forms!frmSalesConfirmation!textSICountry
        textSIAttention = Forms!frmSalesConfirmation!textSIAttention
        
        textSIName.Locked = True
        textSIAddress.Locked = True
        textSICity.Locked = True
        textSIState.Locked = True
        textSIZipCode.Locked = True
        textSICountry.Locked = True
        textSIAttention.Locked = True
        
        If IsNull(Forms!frmSalesConfirmation!Text227) Then
           Text183 = ""
        Else
           Text183 = Forms!frmSalesConfirmation!Text227
        End If
        Text183 = Forms!frmSalesConfirmation!Text227
        Text229 = Forms!frmSalesConfirmation!Text229
        Text231 = Forms!frmSalesConfirmation!Text231
        Text90 = Forms!frmSalesConfirmation!Text90
        Text92 = Forms!frmSalesConfirmation!Text92
        Notes = Forms!frmSalesConfirmation!textNotes
        
        strSQL = "SELECT COUNT(*) AS Cnt FROM tblCommodity WHERE ConfirmationNumber = '" & Text63 & "'"
        Set rst = CurrentDb.OpenRecordset(strSQL)
        topComms = rst!cnt
        Set rst = Nothing
        SQLBAR = "SELECT * FROM tblBalanceAR WHERE ConfirmationNumber = '" & Text63 & "' AND AmountReleased <> 0"
        SQLBAR = "SELECT AR.ID AS [ARID], AR.AmountToBeReleased & '-' & C.Sizes & ' ' & C.Measurement & ' ' & C.Pack AS [Amount To Be Released], AR.ConfirmationNumber, AR.ReleaseNumber, C.Description, "
        SQLBAR = SQLBAR & "C.PricePerPound & '/' & C.Measurement AS [Price], (C.Quantity - (AR.AmountReleased + AR.AmountToBeReleased)) & C.Measurement AS [Balance After release], AR.CommodityId AS [CommId] "
        SQLBAR = SQLBAR & "FROM tblCommodity AS C "
        SQLBAR = SQLBAR & "INNER JOIN (SELECT TOP " & topComms & " ID, ReleaseDate, ConfirmationNumber, ReleaseNumber, CommodityId, AmountReleased, AmountToBeReleased "
        SQLBAR = SQLBAR & "FROM tblBalanceAR "
        SQLBAR = SQLBAR & "WHERE ConfirmationNumber = '" & Text63 & "' "
        SQLBAR = SQLBAR & "ORDER BY ID DESC) AS AR "
        SQLBAR = SQLBAR & "ON AR.CommodityId = C.ID "
        SQLBAR = SQLBAR & "WHERE C.ConfirmationNumber = '" & Text63 & "'" ' AND AmountReleased <> 0 "
        SQLBAR = SQLBAR & "ORDER BY AR.ID ASC"
        '****************
        Debug.Print SQLBAR
        listBalanceAR.RowSource = SQLBAR
        Me.listBalanceAR.ColumnWidths = "0;2350;0;0;5000;800;1200"
        'Closes prev form
        'DoCmd.Close acForm, "frmSalesConfirmation"
        lockFields
        Text128.Locked = False
        'Creating Release No
        '************************************************
        'Text183 = DLookup("[BuyerPONo]", "tblOrders", "ConfirmationNumber = " & Chr(34) & CN & Chr(34))
        cdLastRelease = DMax("[ReleaseNumber]", "tblBalanceAR", "ConfirmationNumber = " & Chr(34) & Text63 & Chr(34))
        Text128 = Int(cdLastRelease) + 1
    End If
End Sub
Private Sub btnPrint_Click()
    Dim rpt As Report
    Dim lbl As Access.Label
    Const TW As Integer = 567
    'result = validateForm() 'Get response from validateForm function
    'If IsNull(result) Then
        DoCmd.OpenReport "rptContractRelease", acViewDesign
        'Header
        Reports!rptContractRelease!labelConfirmationNumber.Caption = Text63
        Reports!rptContractRelease!labelReleaseNumber.Caption = Text128
        'sConfirmationDate = CStr(Confirmation_Date.Value)
        'cdMonth = Mid(CStr(sConfirmationDate), 4, 2)
        'cdDay = Mid(CStr(sConfirmationDate), 1, 2)
        'cdYear = Mid(CStr(sConfirmationDate), 7, 4)
        'sConfirmationDate = cdMonth & "/" & cdDay & "/" & cdYear
        'ConfirmationDate = sConfirmationDate
        'ConfirmationDate = Confirmation_Date.Value
                
        'sContract_Start_Date = CStr(Contract_Start_Date.Value)
        'csdMonth = Mid(CStr(sContract_Start_Date), 4, 2)
        'csdDay = Mid(CStr(sContract_Start_Date), 1, 2)
        'csdYear = Mid(CStr(sContract_Start_Date), 7, 4)
        'sContract_Start_Date = csdMonth & "/" & csdDay & "/" & csdYear
        'ContractStartDate = sContract_Start_Date
        'ContractStartDate = Contract_Start_Date.Value
        
        'sContract_End_Date = CStr(Contract_End_Date.Value)
        'cedMonth = Mid(CStr(sContract_End_Date), 4, 2)
        'cedDay = Mid(CStr(sContract_End_Date), 1, 2)
        'cedYear = Mid(CStr(sContract_End_Date), 7, 4)
        'sContract_End_Date = cedMonth & "/" & cedDay & "/" & cedYear
        'ContractEndDate = sContract_End_Date
        'ContractEndDate = Contract_End_Date.Value
        
        ReleaseDate = Release_Date.Value
        
        'Confirmation, Contract Start and End Dates
        Reports!rptContractRelease!labelReleaseDate.Caption = ReleaseDate
        'Reports!rptContractRelease!labelStartDate.Caption = ContractStartDate
        'Reports!rptContractRelease!labelEndDate.Caption = ContractEndDate
        
        'Sold For labelSF
        Reports!rptContractRelease!labelSFCompanyName.Caption = textSFName.Value
        If IsNull(textSFAddress) Then
           Reports!rptContractRelease!labelSFAddress1.Caption = ""
        Else
           Reports!rptContractRelease!labelSFAddress1.Caption = textSFAddress
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
        Reports!rptContractRelease!labelSFAddress2.Caption = tempSFCity + tempSFState + tempSFZipCode
        If IsNull(textSFCountry) Then
            Reports!rptContractRelease!labelSFCountry.Caption = ""
        Else
            Reports!rptContractRelease!labelSFCountry.Caption = textSFCountry
        End If
        If IsNull(textSFAttention) Then
           Reports!rptContractRelease!labelSFAttention.Caption = ""
        Else
           Reports!rptContractRelease!labelSFAttention.Caption = textSFAttention
        End If
        
        'Sold To labelST
        Reports!rptContractRelease!labelSTCompanyName.Caption = textSTName
        
        If IsNull(textSTAddress) Then
           Reports!rptContractRelease!labelSTAddress1.Caption = ""
        Else
           Reports!rptContractRelease!labelSTAddress1.Caption = textSTAddress
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
        Reports!rptContractRelease!labelSTAddress2.Caption = tempSTCity + tempSTState + tempSTZipCode
        If IsNull(textSTCountry) Then
           Reports!rptContractRelease!labelSTCountry.Caption = ""
        Else
           Reports!rptContractRelease!labelSTCountry.Caption = textSTCountry
        End If
        If IsNull(textSTAttention) Then
           Reports!rptContractRelease!labelSTAttention.Caption = ""
        Else
           Reports!rptContractRelease!labelSTAttention.Caption = textSTAttention
        End If
                
        'Releases
        If CInt(Reports!rptContractRelease!lblCountRel.Caption) > 0 Then
            For i = 1 To CInt(Reports!rptContractRelease!lblCountRel.Caption)
                DeleteReportControl "rptContractRelease", "tLabelReleaseDate" & i
                DeleteReportControl "rptContractRelease", "tLabelDescription" & i
                DeleteReportControl "rptContractRelease", "tLabelATBR" & i
                DeleteReportControl "rptContractRelease", "tLabelBAR" & i
            Next i
        End If
        strSQL = "SELECT AR.AmountToBeReleased AS [ToBeReleased], C.Sizes AS [Sizes], C.Pack AS [Pack], C.Description AS [Description], C.Quantity AS [Quantity], C.PricePerPound AS [PPP], AR.AmountReleased AS [AR], AR.AmountToBeReleased AS [ATBR], C.Quantity - (AR.AmountReleased + AR.AmountToBeReleased) AS [BAR], C.Measurement AS Measurement FROM tblBalanceAR AS AR INNER JOIN tblCommodity AS C ON C.ID = AR.CommodityId WHERE AR.ConfirmationNumber = '" & Text63 & "' AND AR.ReleaseNumber = '" & Text128 & "' AND AR.AmountToBeReleased <> 0 ORDER BY AR.ID DESC"
        Set rst = CurrentDb.OpenRecordset(strSQL)
        i = 1
        'Top = (8.711 * TW)
        Top = (0.6 * TW)
        H1 = 0.6
        Do Until rst.EOF
            'Quantity text size calculations
            If Len(CDec(rst!Quantity) & " - " & rst!Sizes & " " & rst!Measurement & " " & rst!Pack) > 15 Then
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
            If L1 >= L2 Then
                L = L1
            ElseIf L2 >= L1 Then
                L = L2
            End If
            H = (H1 * L)
            Height = (H * TW)
            Set lbl = CreateReportControl("rptContractRelease", acLabel, , , rst!ToBeReleased & "-" & rst!Sizes & " " & rst!Measurement & " " & rst!Pack, (0.608 * TW), Top, (3.388 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelReleaseDate" & i
            Set lbl = CreateReportControl("rptContractRelease", acLabel, , , rst!Description, (3.28 * TW), Top, (9.021 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelDescription" & i
            Set lbl = CreateReportControl("rptContractRelease", acLabel, , , "$" & rst!PPP & "/" & rst!Measurement, (15.291 * TW), Top, (4.418 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelATBR" & i
            Set lbl = CreateReportControl("rptContractRelease", acLabel, , , rst!BAR & rst!Measurement, (17.804 * TW), Top, (1.825 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelBAR" & i
            'cb.AddItem rst!Description
            Top = Top + Height
            i = i + 1
            rst.MoveNext
        Loop
        Reports!rptContractRelease!lblCountRel.Caption = i - 1
        Set rst = Nothing
        
        If IsNull(Text183) Then
           Reports!rptContractRelease!labelBuyerPONo.Caption = ""
        Else
           Reports!rptContractRelease!labelBuyerPONo.Caption = Text183
        End If
        Reports!rptContractRelease!labelTerms.Caption = Text229
        Reports!rptContractRelease!labelFOB.Caption = Text231
        Reports!rptContractRelease!labelShipment.Caption = Text90
        Reports!rptContractRelease!labelRoute.Caption = Text92
        
        'Shipping Info
        Reports!rptContractRelease!labelSICompanyName.Caption = ""
        Reports!rptContractRelease!labelSIAddress1.Caption = ""
        Reports!rptContractRelease!labelSIAddress2.Caption = ""
        Reports!rptContractRelease!labelSICountry.Caption = ""
        Reports!rptContractRelease!labelSIAttention.Caption = ""
        
        If comboSI = "Customer pickup" Then
            Reports!rptContractRelease!Label22.Caption = "Customer pickup:"
        Else
            Reports!rptContractRelease!Label22.Caption = "Ship to:"
        End If
        
        If IsNull(textSIName) Then
           Reports!rptContractRelease!labelSICompanyName.Caption = ""
        Else
           Reports!rptContractRelease!labelSICompanyName.Caption = textSIName
        End If
        'Reports!rptContractRelease!labelSICompanyName.Caption = textSIName
        
        If IsNull(textSIAddress) Then
           Reports!rptContractRelease!labelSIAddress1.Caption = ""
        Else
           Reports!rptContractRelease!labelSIAddress1.Caption = textSIAddress
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
        Reports!rptContractRelease!labelSIAddress2.Caption = tempSICity + tempSIState + tempSIZipCode
        If IsNull(textSICountry) Then
           Reports!rptContractRelease!labelSICountry.Caption = ""
        Else
           Reports!rptContractRelease!labelSICountry.Caption = textSICountry
        End If
        If IsNull(textSIAttention) Then
           Reports!rptContractRelease!labelSIAttention.Caption = ""
        Else
           Reports!rptContractRelease!labelSIAttention.Caption = textSIAttention
        End If
        
        'Total Commission
        Reports!rptContractRelease!labelNotes.Caption = Notes
        
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
        'Reports!rptSalesConfirmation!labelContractPeriodInfo.Caption = "**CONTRACT PERIOD: " & months(csdMonth) & " " & csdYear & " THROUGH " & months(cedMonth) & " " & cedDay & ", " & cedYear & " AT BUYER'S CALL"
        'Show report
        DoCmd.OpenReport "rptContractRelease", acViewReport
    'Else
    '    i = UBound(result)
    '    Msg = "The following field(s) need to be filled: " & vbNewLine
    '    For j = 1 To i
    '        If j <> i Then
    '            Msg = Msg & result(j) & vbNewLine
    '        Else
    '            Msg = Msg & result(j) & "."
    '        End If
    '    Next
    '    MsgBox Msg
    'End If
End Sub
Private Sub btnPDF_Click()
    Dim rpt As Report
    Dim lbl As Access.Label
    Const TW As Integer = 567
    'result = validateForm() 'Get response from validateForm function
    'If IsNull(result) Then
        DoCmd.OpenReport "rptContractRelease", acViewDesign
        'Header
        Reports!rptContractRelease!labelConfirmationNumber.Caption = Text63
        Reports!rptContractRelease!labelReleaseNumber.Caption = Text128
        'sConfirmationDate = CStr(Confirmation_Date.Value)
        'cdMonth = Mid(CStr(sConfirmationDate), 4, 2)
        'cdDay = Mid(CStr(sConfirmationDate), 1, 2)
        'cdYear = Mid(CStr(sConfirmationDate), 7, 4)
        'sConfirmationDate = cdMonth & "/" & cdDay & "/" & cdYear
        'ConfirmationDate = sConfirmationDate
        'ConfirmationDate = Confirmation_Date.Value
                
        'sContract_Start_Date = CStr(Contract_Start_Date.Value)
        'csdMonth = Mid(CStr(sContract_Start_Date), 4, 2)
        'csdDay = Mid(CStr(sContract_Start_Date), 1, 2)
        'csdYear = Mid(CStr(sContract_Start_Date), 7, 4)
        'sContract_Start_Date = csdMonth & "/" & csdDay & "/" & csdYear
        'ContractStartDate = sContract_Start_Date
        'ContractStartDate = Contract_Start_Date.Value
        
        'sContract_End_Date = CStr(Contract_End_Date.Value)
        'cedMonth = Mid(CStr(sContract_End_Date), 4, 2)
        'cedDay = Mid(CStr(sContract_End_Date), 1, 2)
        'cedYear = Mid(CStr(sContract_End_Date), 7, 4)
        'sContract_End_Date = cedMonth & "/" & cedDay & "/" & cedYear
        'ContractEndDate = sContract_End_Date
        'ContractEndDate = Contract_End_Date.Value
        
        ReleaseDate = Release_Date.Value
        
        'Confirmation, Contract Start and End Dates
        Reports!rptContractRelease!labelReleaseDate.Caption = ReleaseDate
        'Reports!rptContractRelease!labelStartDate.Caption = ContractStartDate
        'Reports!rptContractRelease!labelEndDate.Caption = ContractEndDate
        
        'Sold For labelSF
        Reports!rptContractRelease!labelSFCompanyName.Caption = textSFName.Value
        If IsNull(textSFAddress) Then
           Reports!rptContractRelease!labelSFAddress1.Caption = ""
        Else
           Reports!rptContractRelease!labelSFAddress1.Caption = textSFAddress
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
        Reports!rptContractRelease!labelSFAddress2.Caption = tempSFCity + tempSFState + tempSFZipCode
        If IsNull(textSFCountry) Then
            Reports!rptContractRelease!labelSFCountry.Caption = ""
        Else
            Reports!rptContractRelease!labelSFCountry.Caption = textSFCountry
        End If
        If IsNull(textSFAttention) Then
           Reports!rptContractRelease!labelSFAttention.Caption = ""
        Else
           Reports!rptContractRelease!labelSFAttention.Caption = textSFAttention
        End If
        
        'Sold To labelST
        Reports!rptContractRelease!labelSTCompanyName.Caption = textSTName
        
        If IsNull(textSTAddress) Then
           Reports!rptContractRelease!labelSTAddress1.Caption = ""
        Else
           Reports!rptContractRelease!labelSTAddress1.Caption = textSTAddress
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
        Reports!rptContractRelease!labelSTAddress2.Caption = tempSTCity + tempSTState + tempSTZipCode
        If IsNull(textSTCountry) Then
           Reports!rptContractRelease!labelSTCountry.Caption = ""
        Else
           Reports!rptContractRelease!labelSTCountry.Caption = textSTCountry
        End If
        If IsNull(textSTAttention) Then
           Reports!rptContractRelease!labelSTAttention.Caption = ""
        Else
           Reports!rptContractRelease!labelSTAttention.Caption = textSTAttention
        End If
                
        'Releases
        If CInt(Reports!rptContractRelease!lblCountRel.Caption) > 0 Then
            For i = 1 To CInt(Reports!rptContractRelease!lblCountRel.Caption)
                DeleteReportControl "rptContractRelease", "tLabelReleaseDate" & i
                DeleteReportControl "rptContractRelease", "tLabelDescription" & i
                DeleteReportControl "rptContractRelease", "tLabelATBR" & i
                DeleteReportControl "rptContractRelease", "tLabelBAR" & i
            Next i
        End If
        strSQL = "SELECT AR.AmountToBeReleased AS [ToBeReleased], C.Sizes AS [Sizes], C.Pack AS [Pack], C.Description AS [Description], C.Quantity AS [Quantity], C.PricePerPound AS [PPP], AR.AmountReleased AS [AR], AR.AmountToBeReleased AS [ATBR], C.Quantity - (AR.AmountReleased + AR.AmountToBeReleased) AS [BAR], C.Measurement AS Measurement FROM tblBalanceAR AS AR INNER JOIN tblCommodity AS C ON C.ID = AR.CommodityId WHERE AR.ConfirmationNumber = '" & Text63 & "' AND AR.ReleaseNumber = '" & Text128 & "' AND AR.AmountToBeReleased <> 0 ORDER BY AR.ID DESC"
        Set rst = CurrentDb.OpenRecordset(strSQL)
        i = 1
        'Top = (8.711 * TW)
        Top = (0.6 * TW)
        H1 = 0.6
        Do Until rst.EOF
            'Quantity text size calculations
            If Len(CDec(rst!Quantity) & " - " & rst!Sizes & " " & rst!Measurement & " " & rst!Pack) > 15 Then
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
            If L1 >= L2 Then
                L = L1
            ElseIf L2 >= L1 Then
                L = L2
            End If
            H = (H1 * L)
            Height = (H * TW)
            Set lbl = CreateReportControl("rptContractRelease", acLabel, , , rst!ToBeReleased & "-" & rst!Sizes & " " & rst!Measurement & " " & rst!Pack, (0.608 * TW), Top, (3.388 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelReleaseDate" & i
            Set lbl = CreateReportControl("rptContractRelease", acLabel, , , rst!Description, (3.28 * TW), Top, (9.021 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelDescription" & i
            Set lbl = CreateReportControl("rptContractRelease", acLabel, , , "$" & rst!PPP & "/" & rst!Measurement, (15.291 * TW), Top, (4.418 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelATBR" & i
            Set lbl = CreateReportControl("rptContractRelease", acLabel, , , rst!BAR & rst!Measurement, (17.804 * TW), Top, (1.825 * TW), Height)
            lbl.TextAlign = 2
            lbl.ForeColor = lngBlack
            lbl.Name = "tLabelBAR" & i
            'cb.AddItem rst!Description
            Top = Top + Height
            i = i + 1
            rst.MoveNext
        Loop
        Reports!rptContractRelease!lblCountRel.Caption = i - 1
        Set rst = Nothing
        
        If IsNull(Text183) Then
           Reports!rptContractRelease!labelBuyerPONo.Caption = ""
        Else
           Reports!rptContractRelease!labelBuyerPONo.Caption = Text183
        End If
        Reports!rptContractRelease!labelTerms.Caption = Text229
        Reports!rptContractRelease!labelFOB.Caption = Text231
        Reports!rptContractRelease!labelShipment.Caption = Text90
        Reports!rptContractRelease!labelRoute.Caption = Text92
        
        'Shipping Info
        Reports!rptContractRelease!labelSICompanyName.Caption = ""
        Reports!rptContractRelease!labelSIAddress1.Caption = ""
        Reports!rptContractRelease!labelSIAddress2.Caption = ""
        Reports!rptContractRelease!labelSICountry.Caption = ""
        Reports!rptContractRelease!labelSIAttention.Caption = ""
        
        If comboSI = "Customer pickup" Then
            Reports!rptContractRelease!Label22.Caption = "Customer pickup:"
        Else
            Reports!rptContractRelease!Label22.Caption = "Ship to:"
        End If
        
        If IsNull(textSIName) Then
           Reports!rptContractRelease!labelSICompanyName.Caption = ""
        Else
           Reports!rptContractRelease!labelSICompanyName.Caption = textSIName
        End If
        'Reports!rptContractRelease!labelSICompanyName.Caption = textSIName
        
        If IsNull(textSIAddress) Then
           Reports!rptContractRelease!labelSIAddress1.Caption = ""
        Else
           Reports!rptContractRelease!labelSIAddress1.Caption = textSIAddress
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
        Reports!rptContractRelease!labelSIAddress2.Caption = tempSICity + tempSIState + tempSIZipCode
        If IsNull(textSICountry) Then
           Reports!rptContractRelease!labelSICountry.Caption = ""
        Else
           Reports!rptContractRelease!labelSICountry.Caption = textSICountry
        End If
        If IsNull(textSIAttention) Then
           Reports!rptContractRelease!labelSIAttention.Caption = ""
        Else
           Reports!rptContractRelease!labelSIAttention.Caption = textSIAttention
        End If
        
        'Total Commission
        Reports!rptContractRelease!labelNotes.Caption = Notes
        
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
        'Reports!rptSalesConfirmation!labelContractPeriodInfo.Caption = "**CONTRACT PERIOD: " & months(csdMonth) & " " & csdYear & " THROUGH " & months(cedMonth) & " " & cedDay & ", " & cedYear & " AT BUYER'S CALL"
        'Show report
        'DoCmd.OpenReport "rptContractRelease", acViewReport
        DoCmd.OutputTo acOutputReport, "rptContractRelease", acFormatPDF, , , , , acExportQualityPrint
    'Else
    '    i = UBound(result)
    '    Msg = "The following field(s) need to be filled: " & vbNewLine
    '    For j = 1 To i
    '        If j <> i Then
    '            Msg = Msg & result(j) & vbNewLine
    '        Else
    '            Msg = Msg & result(j) & "."
    '        End If
    '    Next
    '    MsgBox Msg
    'End If
End Sub
Private Sub btnSaveExit_Click()
    DoCmd.Close acForm, "frmContractRelease"
End Sub
Private Sub cmdNewRelease_Click()
    If IsNull(Text128) Then
        MsgBox "DRBC Release Number cannot be empty"
    Else
        strSQL = "SELECT COUNT(*) AS Cnt FROM tblBalanceAR WHERE ConfirmationNumber = '" & Text63 & "' AND ReleaseNumber = '" & Text128 & "'"
        Set rst = CurrentDb.OpenRecordset(strSQL)
        RNCheck = rst!cnt
        Set rst = Nothing
        If RNCheck <> 0 Then
            MsgBox "DRBC Release Number " & Text128 & " is used, please select another value"
        Else
            DoCmd.OpenForm FormName:="frmNewRelease", WindowMode:=acDialog
            'SQLBAR = "SELECT AR.ARID, AR.ReleaseDate AS [Release Date], C.Description AS [Description], C.Quantity AS [Initial Commodity Sold], AR.AmountReleased AS [Amount released], AR.AmountToBeReleased AS [Amount to be released], (C.Quantity - (AR.AmountReleased + AR.AmountToBeReleased)) AS [Balance After release] "
            'SQLBAR = SQLBAR & "FROM tblTempBalanceAR AS AR "
            'SQLBAR = SQLBAR & "INNER JOIN tblCommodity AS C "
            'SQLBAR = SQLBAR & "ON AR.CommodityId = C.ID"
            strSQL = "SELECT COUNT(*) AS Cnt FROM tblCommodity WHERE ConfirmationNumber = '" & Text63 & "'"
            Set rst = CurrentDb.OpenRecordset(strSQL)
            topComms = rst!cnt
            Set rst = Nothing
        
            SQLBAR = "SELECT * FROM tblBalanceAR WHERE ConfirmationNumber = '" & Text63 & "' AND AmountReleased <> 0"
            SQLBAR = "SELECT AR.ID AS [ARID], AR.ReleaseDate AS [Release Date], AR.ConfirmationNumber, AR.ReleaseNumber, C.Description, "
            SQLBAR = SQLBAR & "AR.AmountToBeReleased & C.Measurement AS AmountToBeReleased, (C.Quantity - (AR.AmountReleased + AR.AmountToBeReleased)) & C.Measurement AS [Balance After release], AR.CommodityId AS [CommId] "
            SQLBAR = SQLBAR & "FROM tblCommodity AS C "
            SQLBAR = SQLBAR & "INNER JOIN (SELECT TOP " & topComms & " ID, ReleaseDate, ConfirmationNumber, ReleaseNumber, CommodityId, AmountReleased, AmountToBeReleased "
            SQLBAR = SQLBAR & "FROM tblBalanceAR "
            SQLBAR = SQLBAR & "WHERE ConfirmationNumber = '" & Text63 & "' "
            SQLBAR = SQLBAR & "ORDER BY ID DESC) AS AR "
            SQLBAR = SQLBAR & "ON AR.CommodityId = C.ID "
            SQLBAR = SQLBAR & "WHERE C.ConfirmationNumber = '" & Text63 & "'" ' AND AmountReleased <> 0 "
            SQLBAR = SQLBAR & "ORDER BY AR.ID ASC"
            listBalanceAR.RowSource = SQLBAR
            Me.listBalanceAR.ColumnWidths = "0;1500;0;0;5000;1200;1200"
            cmdNewRelease.Enabled = False
        End If
    End If
End Sub
