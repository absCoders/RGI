Imports nsoftware.InShip
Imports System.Drawing.Printing
Imports System.IO

Public Class WHFSHIP1

    Private divisionView As DataView
    Private providerView As DataView
    Private shipmethodView As DataView

    Private clsShip As New TAC.WHCSHIP1
    Private SHIP_CNTL_NO As String = String.Empty
    Private MASTER_TRACKING_NO As String = String.Empty

    Private rowWHTSHPC1 As DataRow
    Private tblPayorType As New DataTable
    Private shipPackageDetailList As New List(Of nsoftware.InShip.PackageDetail)
    Private ShippingLabelDirectory As String

    Private WithEvents ultraComboPackage As Infragistics.Win.UltraWinGrid.UltraCombo = New Infragistics.Win.UltraWinGrid.UltraCombo

    Private isInternationalShipment As Boolean = False
    Private Const FedexSmartPost As Int16 = 26
    Private Const tempSHIP_CNTL_NO As String = "@temp@"

    Private printerFound As Boolean = False

    ' When using the CustomContent, the LabelStockType must be either 4 (Stock 4x8) or 5 (Stock 4x9 Leading Doc Tab). 
    ' Also LabelFormatType must be 0 (Common2D) and LabelImageType must 2 (fitEltron), 3 (fitZebra) or 4 (fitUniMark). 
    Private customContent As String = String.Empty
    Private numCallTags As Int16 = 1

    Private Const CarrierMailsLabelsToCustomer As String = "N"
    Private Const CarrierEmailLabelsToCustomer As String = "U"
    Private Const Printlabels As String = "P"
    Private Const EmailLabelsToCustomer As String = "E"
    Private Const UPSOneAttemptToPickup As String = "1"
    Private Const UPSThreeAttemptsToPickup As String = "3"

    Private rowSOTRMAF1 As DataRow = Nothing
    ' 15 Pounds as Ounces
    Private Const PKG_WEIGHT As Int16 = 1 * 16
    Private ReturnLabelsToSendToCustomers As New List(Of String)


    ' Valid values for WHTSHPC1.STATUS
    '   I - Initial Setup before calling request.
    '   P - processed - label printed
    '   C - Cancelled
    '

#Region "ABS Standard Routines"

    Private Sub WHFSHIP1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If ScannerInUse AndAlso Not txtUserSuppliedValue.Focused AndAlso 1 = 2 Then
            txtUserSuppliedValue.Focus()
        End If
    End Sub

    ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim sql As String = String.Empty
        With dst

            Get_PARM("ASTPARM1")
            Get_PARM("SOTPARM1")

            Create_TDA(.Tables.Add, "SOTCARR1", "*")
            Create_TDA(.Tables.Add, "SOTSVIA1", "*")
            Create_TDA(.Tables.Add, "SOTCARR3", "*")
            .Tables("SOTCARR3").Columns.Add("CARRIER_DESC", GetType(System.String))
            .Tables("SOTCARR3").Columns.Add("CARRIER_SHIP_TYPE", GetType(System.String))
            .Tables("SOTCARR3").Columns.Add("PROVIDER_TYPE", GetType(System.String))
            .Tables("SOTCARR3").Columns.Add("CARRIER_DESC_DISP", GetType(System.String))

            Create_TDA(.Tables.Add, "WHTSHPC1", "*")
            Create_TDA(.Tables.Add, "WHTSHPC2", "*", 1)
            .Tables("WHTSHPC2").Columns.Add("POUNDS", GetType(System.Int16))
            .Tables("WHTSHPC2").Columns.Add("OUNCES", GetType(System.Int16))
            .Tables("WHTSHPC2").Columns.Add("SEL", GetType(System.Int16))
            .Tables("WHTSHPC2").Columns("SEL").DefaultValue = 0
            .Tables("WHTSHPC2").Columns.Add("PKG_WEIGHT", GetType(System.Int32), "ISNULL(POUNDS, 0) * 16 + ISNULL(OUNCES, 0)")

            Create_TDA(.Tables.Add, "WHTSHPCC", "*", 1)
            .Tables("WHTSHPCC").Columns.Add("EXTENDED_PRICE", GetType(System.Decimal), "ISNULL(QUANTITY, 0) * ISNULL(UNIT_PRICE, 0)")
            .Tables("WHTSHPCC").Columns.Add("EXTENDED_WEIGHT", GetType(System.Decimal), "ISNULL(QUANTITY, 0) * ISNULL(WEIGHT, 0)")

            Create_TDA(.Tables.Add, "WHTSHPCS", "*", 1)
            .Tables("WHTSHPCS").Columns.Add("SEL", GetType(System.String))

            Create_TDA(.Tables.Add, "SOTRMAFL", "*")

            Create_TDA(.Tables.Add, "SHTSHIPS", "*")
            .Tables("SHTSHIPS").Columns.Add("SEL", GetType(System.String))
            Fill_Records("SHTSHIPS", String.Empty, True, "SELECT '0' SEL, SHTSHIPS.* FROM SHTSHIPS")

            Create_TDA(.Tables.Add, "WHTSDESC", "*")
            Fill_Records("WHTSDESC", String.Empty, True, "SELECT * FROM WHTSDESC")

            Fill_Records("SOTCARR1", "", True, "SELECT * FROM SOTCARR1")
            Fill_Records("SOTSVIA1", "", True, "SELECT * FROM SOTSVIA1")
            Fill_Records("SOTCARR3", "", True, "Select SOTCARR3.*, SOTCARR1.CARRIER_DESC, SOTCARR1.CARRIER_SHIP_TYPE, " _
                         & " SOTCARR1.PROVIDER_TYPE, SOTCARR1.CARRIER_CODE || ' - ' || SOTCARR3.SHIPPER_DIVISION_CODE CARRIER_DESC_DISP" _
                         & " From SOTCARR3, SOTCARR1" _
                         & " Where SOTCARR3.CARRIER_CODE = SOTCARR1.CARRIER_CODE" _
                         & " AND SOTCARR3.SHIPPER_DIVISION_CODE = '" & ASCMAIN1.CLIENT & "'")

            Create_TDA(.Tables.Add, "WHTSHPC4", "*", 1)
            .Tables("WHTSHPC4").Columns.Add("SHIP_VIA_CODE", GetType(System.String))
            .Tables("WHTSHPC4").Columns.Add("ADDON_TOTAL", GetType(System.Decimal))
            .Tables("WHTSHPC4").Columns("ADDON_TOTAL").DefaultValue = 0
            .Tables("WHTSHPC4").Columns.Add("CUSTOMER_BASE_CHARGE", GetType(System.Decimal))
            .Tables("WHTSHPC4").Columns.Add("TOTAL_CHARGE", GetType(System.Decimal), "ISNULL(ADDON_TOTAL, 0) + ISNULL(CUSTOMER_BASE_CHARGE, 0) + ISNULL(SURCHARGE, 0)")


            Create_TDA(.Tables.Add("WHTSHPC5_SF"), "WHTSHPC5", "*")
            Create_TDA(.Tables.Add("WHTSHPC5_ST"), "WHTSHPC5", "*")
            Create_TDA(.Tables.Add("WHTSHPC5_HL"), "WHTSHPC5", "*")

            'ASTATTA2
            Create_TDA(.Tables.Add, "ASTATTA2", "*")
            Create_TDA(.Tables.Add("WHTSHPCP_S"), "WHTSHPCP", "*")
            Create_TDA(.Tables.Add("WHTSHPCP_D"), "WHTSHPCP", "*")

            Create_TDA(.Tables.Add, "WHTPKGM1", "*")
            Fill_Records("WHTPKGM1", String.Empty, True, "SELECT * FROM WHTPKGM1")

            With tblPayorType
                .Columns.Add("P_CODE", GetType(System.String))
                .Columns.Add("P_DESC", GetType(System.String))
            End With

            For Each valuePair As String In New String() {"S:Sender", "R:Recipient", "T:Third Party", "C:Collect"}
                tblPayorType.Rows.Add(New Object() {valuePair.Split(":")(0), valuePair.Split(":")(1)})
            Next

            Create_TDA(.Tables.Add, "ICTWHSE1", "*")
            Fill_Records("ICTWHSE1", "", True, "SELECT * FROM ICTWHSE1")

            Create_Lookup("WHTPKGM1")

        End With

        cmbDutiesPayor.DataSource = tblPayorType
        cmbDutiesPayor.SelectedRow = cmbDutiesPayor.Rows(0)

        cmbShipPayor.DataSource = tblPayorType
        cmbShipPayor.SelectedRow = cmbShipPayor.Rows(0)

        cmbWarehouse.DataSource = dst.Tables("ICTWHSE1")

        divisionView = New DataView(ASCDATA1.SelectDistinct(dst.Tables("SOTCARR3"), New String() {"DIVISION_CODE"}))
        divisionView.Sort = "DIVISION_CODE"
        cmbDivision.DataSource = divisionView
        If cmbDivision.Rows.Count > 0 Then
            cmbDivision.SelectedRow = cmbDivision.Rows(0)
            cmbDivision.Text = cmbDivision.SelectedRow.Cells("DIVISION_CODE").Value
        End If

        cmbDivision_ValueChanged(Nothing, Nothing)

        cmbDivision.DisplayLayout.Bands(0).Columns("DIVISION_CODE").Width = cmbDivision.Width
        cmbProvider.DisplayLayout.Bands(0).Columns("CARRIER_DESC_DISP").Width = cmbProvider.Width
        cmbShipMethod.DisplayLayout.Bands(0).Columns("SHIP_VIA_DESC").Width = cmbShipMethod.Width

        Dim SO_PARM_DEF_PICK_WHSE As String = (ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & String.Empty).ToString.Trim
        If SO_PARM_DEF_PICK_WHSE.Length > 0 Then
            For Each row As UltraWinGrid.UltraGridRow In cmbWarehouse.Rows
                If row.Cells("WHSE_CODE").Value = SO_PARM_DEF_PICK_WHSE Then
                    cmbWarehouse.SelectedRow = row
                    Exit For
                End If
            Next
        End If

        Bind_Controls(grpShipFrom, "WHTSHPC5_SF")
        Bind_Controls(grpShipTo, "WHTSHPC5_ST")
        Bind_Controls(grpHoldAtLocation, "WHTSHPC5_HL")
        Bind_Controls(grpSmartPost, "WHTSHPC1")

        grdWHTSHPC2.DataSource = dst.Tables("WHTSHPC2")

        Create_Summary(grdWHTSHPC2, "POUNDS", "Sum")
        Create_Summary(grdWHTSHPC2, "OUNCES", "Sum")
        Create_Summary(grdWHTSHPC2, "SHIP_PACKAGE_NO", "Count")

        Create_Summary(grdWHTSHPCC, "COMMODITY_DESC", "Count")
        Create_Summary(grdWHTSHPCC, "WEIGHT", "Sum")
        Create_Summary(grdWHTSHPCC, "EXTENDED_PRICE", "Sum")
        Create_Summary(grdWHTSHPCC, "EXTENDED_WEIGHT", "Sum")

        grdWHTSHPC4.DataSource = dst.Tables("WHTSHPC4")
        Create_Summary(grdWHTSHPC4, "SELECTED", "Count")

        LoadSmartPostDropDowns()

        ASCMAIN1.Add_Value_List(grdWHTSHPC2, "SIGNATURE_TYPE", Nothing, New String() {":", "0:Default for requested service", "1:Adult", _
                            "2:Direct", "3:Indirect", "4:Not Required"}, 0)

        ASCMAIN1.Add_Value_List(grdWHTSHPC2, "COD_TYPE", Nothing, New String() {":", "0:Any Check", _
                             "1:Cashier's check or money order", "2:None"}, 0)

        With ultraComboPackage.DisplayLayout.Bands(0)

            ultraComboPackage.Font = grdWHTSHPC2.Font
            ultraComboPackage.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Default
            ultraComboPackage.DropDownStyle = UltraWinGrid.UltraComboStyle.DropDownList

            .Columns.Add("PKG_CODE")
            .Columns("PKG_CODE").Header.Caption = "Code"
            .Columns("PKG_CODE").Width = 75

            .Columns.Add("PKG_DESC")
            .Columns("PKG_DESC").Header.Caption = "Desc"
            .Columns("PKG_DESC").Width = 75

            .Columns.Add("PKG_D")
            .Columns("PKG_D").Header.Caption = "L x W x H"
            .Columns("PKG_D").Width = 200

        End With

        ultraComboPackage.DataSource = ASCDATA1.GetDataTable("SELECT PKG_CODE, PKG_DESC, PKG_L || ' x ' ||  PKG_W || ' x ' || PKG_H PKG_D FROM WHTPKGM1")
        ultraComboPackage.ValueMember = "PKG_CODE"
        ultraComboPackage.DisplayMember = "PKG_DESC"
        grdWHTSHPC2.DisplayLayout.Bands(0).Columns("PKG_CODE").EditorComponent = ultraComboPackage


        grdWHTSHPCC.DataSource = dst.Tables("WHTSHPCC")
        grdSHTSHPCS.DataSource = dst.Tables("SHTSHIPS")

        Dim Commodities As String() = GetCommodityUnits()
        ASCMAIN1.Add_Value_List(grdWHTSHPCC, "MANUFACTURER", "SELECT COUNTRY_CODE, COUNTRY_NAME FROM WHTMANF1")
        ASCMAIN1.Add_Value_List(grdWHTSHPCC, "QUANTITY_UOM", Nothing, Commodities, 0)

        dteShipDate.MinDate = CDate("01/01/1980")
        dteShipDate.MaxDate = DateAdd(DateInterval.Day, 10, DateTime.Now)
        dteShipDate.DateTime = DateTime.Now
        dteShipDate.Value = dteShipDate.DateTime

        'Set Max lengths
        txtShipCountry.MaxLength = dst.Tables("WHTSHPCP_S").Columns("PAYOR_COUNTRY").MaxLength
        txtDutiesCountry.MaxLength = dst.Tables("WHTSHPCP_D").Columns("PAYOR_COUNTRY").MaxLength

        txtShipAccountNo.MaxLength = dst.Tables("WHTSHPCP_S").Columns("PAYOR_ACCT_NO").MaxLength
        txtDutiesAccountNo.MaxLength = dst.Tables("WHTSHPCP_D").Columns("PAYOR_ACCT_NO").MaxLength

        txtShipAccountZip.MaxLength = dst.Tables("WHTSHPCP_S").Columns("PAYOR_ACCT_ZIP").MaxLength
        txtDutiesAccountZip.MaxLength = dst.Tables("WHTSHPCP_D").Columns("PAYOR_ACCT_ZIP").MaxLength


        Dim ZebraPrinters As New List(Of String)
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            Dim defaultprinter = ""
            For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
                If printerName.ToUpper.StartsWith("ZDESIGNER") Or printerName.ToUpper.StartsWith("MONARCH") Or printerName.ToUpper.StartsWith("AVERY") Or printerName.ToUpper.StartsWith("ZEBRA") Then
                    ZebraPrinters.Add(printerName)
                    If printerName = "MONARCH_NJT2" Or defaultprinter = "" Then
                        defaultprinter = printerName
                    End If
                End If
            Next printerName
            If ZebraPrinters.Count >= 1 Then
                cboZebraPrinter.DataSource = ZebraPrinters
                cboZebraPrinter.SelectedItem = defaultprinter
            End If
        End If


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = String.Empty

        Select Case eItemKey

            Case "Request Label", "Get Rates", "Request Return Label"
                isInternationalShipment = False

                If eItemKey = "Request Return Label" Then
                    If cmbProvider.SelectedRow.Cells("PROVIDER_TYPE").Value <> WHCSHIP1.ProviderTypeUPS Then
                        EMsg &= vbCrLf & "Currently Only UPS supports Return Labels"
                        Exit Select
                    End If
                End If

                If (eItemKey = "Request Label" AndAlso Not printerFound) _
                    OrElse (eItemKey = "Request Return Label" AndAlso Not printerFound AndAlso optReturnLabels.Value & String.Empty = Printlabels) Then
                    EMsg &= vbCrLf & "Label printer was not found you may not Request Labels."
                    Exit Select
                End If

                If cmbDivision.SelectedRow Is Nothing Then
                    EMsg &= vbCrLf & "Missing Division"
                    Exit Select
                End If

                If cmbProvider.SelectedRow Is Nothing Then
                    EMsg &= vbCrLf & "Missing Provider"
                    Exit Select
                End If

                If cmbShipMethod.SelectedRow Is Nothing Then
                    EMsg &= vbCrLf & "Missing Shipping Method"
                    Exit Select
                End If

                If eItemKey = "Request Return Label" Then
                    If optReturnLabels.Value & String.Empty = String.Empty Then
                        EMsg &= vbCrLf & "Return Labels require a Label Delivery Method"
                        Exit Select
                    End If

                    Dim lmsg As String = "Do you want to use Label Delivery Method: " & optReturnLabels.CheckedItem.DisplayText
                    If MessageBox.Show(lmsg, "Request Return Label", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If

                    Select Case optReturnLabels.Value
                        Case CarrierEmailLabelsToCustomer, EmailLabelsToCustomer
                            txtToEmail.Text = txtToEmail.Text.Trim
                            If txtToEmail.Left = 0 Then
                                EMsg &= vbCrLf & "The Return Labels delivery method requires a Ship To email address."
                                Exit Select
                            End If

                        Case Printlabels
                            If Not printerFound Then
                                EMsg &= vbCrLf & "Label printer was not found you may not Request Return Labels. Choose a different Label Delivery Method"
                                Exit Select
                            End If
                    End Select
                End If

                txtFromCountry.Text = txtFromCountry.Text.Trim.ToUpper
                txtToCountry.Text = txtToCountry.Text.Trim.ToUpper
                txtToState.Text = txtToState.Text.Trim.ToUpper

                If txtFromCountry.TextLength = 0 Then
                    EMsg &= vbCrLf & "Missing Ship From Country"
                End If

                If txtToCountry.TextLength = 0 Then
                    EMsg &= vbCrLf & "Missing Ship To Country"
                End If

                If grdWHTSHPC2.Rows.Count = 0 AndAlso eItemKey <> "Request Return Label" Then
                    EMsg &= vbCrLf & "At least one package must be entered"
                Else
                    For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWHTSHPC2.Rows
                        ValidatePacking(grdRow)
                    Next
                End If


                If CDate(dteShipDate.DateTime.ToShortDateString) < CDate(DateTime.Now.ToShortDateString) Then
                    EMsg &= vbCrLf & "Ship date may not be less than today."
                End If

                ' Chnage USA to US.
                txtToCountry.Text = txtToCountry.Text.ToUpper.Trim
                If txtToCountry.Text.StartsWith("US") Then txtToCountry.Text = "US"

                txtFromCountry.Text = txtFromCountry.Text.ToUpper.Trim
                If txtFromCountry.Text.StartsWith("US") Then txtFromCountry.Text = "US"

                txtToPhone.Text = txtToPhone.Text.Trim
                If txtToPhone.TextLength = 0 Then
                    EMsg &= vbCrLf & "Missing Ship to Telephone."
                End If

                If EMsg.Length = 0 Then
                    ' Treat PR as international
                    isInternationalShipment = (txtToCountry.Text.ToUpper <> txtFromCountry.Text.ToUpper) OrElse (txtToCountry.Text = "US" AndAlso txtToState.Text = "PR")

                    If isInternationalShipment Then
                        If grdWHTSHPCC.Rows.Count = 0 Then
                            EMsg &= vbCrLf & "International Shipments require Commodities"
                        Else
                            For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWHTSHPCC.Rows
                                ValidateCommodity(grdRow)
                            Next
                        End If

                        If Val(numCustomsValue.Value & String.Empty) <= 0 Then
                            EMsg &= vbCrLf & "International shipments require Customs Value."
                        End If
                    End If
                End If

                ' Fedex Smart Post Specific
                If cmbProvider.SelectedRow.Cells("PROVIDER_TYPE").Value = WHCSHIP1.ProviderTypeFedex _
                    AndAlso cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value = 26 Then
                    If cmbSmartPost.SelectedRow Is Nothing Then
                        EMsg &= vbCrLf & "Smart Post shipping methods require a Smart Post Type."
                    End If
                    If cmbSmartPostPackage.SelectedRow Is Nothing Then
                        EMsg &= vbCrLf & "Smart Post shipping methods require a package Type."
                    End If
                    If cmbProvider.SelectedRow.Cells("FEDEX_HUB_ID").Value & String.Empty = String.Empty Then
                        EMsg &= vbCrLf & "Smart Post shipping methods require a Hub Id."
                    End If
                End If

                ' Look for COD amounts
                If eItemKey = "Request Label" AndAlso cmbShipMethod.SelectedRow.Cells("COD_IND").Value & String.Empty = "1" Then
                    For Each rowWHTSHPC2 As DataRow In dst.Tables("WHTSHPC2").Select("", "", DataViewRowState.CurrentRows)
                        Dim COD_AMOUNT As Decimal = Val(rowWHTSHPC2.Item("COD_AMOUNT") & String.Empty)
                        Dim COD_TYPE As String = Val(rowWHTSHPC2.Item("COD_TYPE") & String.Empty)

                        If (COD_TYPE <> "0" AndAlso COD_TYPE <> "1") OrElse COD_AMOUNT < 0 Then
                            EMsg &= vbCrLf & "COD shipments require all packages have a dollar value and proper COD Type setting"
                            Exit For
                        End If

                    Next
                End If

                If EMsg.Length = 0 Then
                    Select Case cmbProvider.SelectedRow.Cells("PROVIDER_TYPE").Value
                        Case WHCSHIP1.ProviderTypeFedex
                            If Not isInternationalShipment Then
                                clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpress)
                            Else
                                clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpressInternational)
                            End If
                        Case WHCSHIP1.ProviderTypeUPS
                            clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.UPS)
                        Case WHCSHIP1.ProviderTypeUSPS
                            clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.USPS)
                        Case WHCSHIP1.ProviderTypeCanada
                            clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.CanadaPost)
                        Case Else
                            EMsg &= vbCrLf & "Invalid or Missing Provider"
                            Exit Select
                    End Select
                End If

                If EMsg.Length = 0 AndAlso eItemKey = "Request Return Label" Then
                    If rowSOTRMAF1 IsNot Nothing Then
                        Dim RA_NO As String = rowSOTRMAF1.Item("RA_NO")
                        Dim tblSOTRMAFL As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTRMAFL WHERE RA_NO = :PARM1", "SOTRMAFL", "V", New Object() {RA_NO})
                        If tblSOTRMAFL.Rows.Count > 0 Then
                            If MessageBox.Show("The provided RMA has " & tblSOTRMAFL.Rows.Count & " call tags. Do you want to generate more?", "Request Return Label", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If

                If EMsg.Length = 0 Then
                    If eItemKey = "Request Return Label" Then
                        If MessageBox.Show("Do you want to create Return Labels for Shipping Method: " & cmbShipMethod.Text & "?", eItemKey, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    ElseIf MessageBox.Show("Do you want to " & eItemKey & "?", eItemKey, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Clear"
                If MessageBox.Show("Do you want to clear the contents of the screen?", "Clear", _
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Cancel Shipment"
                If dst.Tables("WHTSHPC2").Select("SEL = '1'").Length = 0 Then
                    EMsg &= vbCr & "You must select the shipments to cancel."
                    Exit Select
                End If

                If MessageBox.Show("Do you want to cancel the selected packages?", "Cancel Shipment", _
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                Select Case cmbProvider.SelectedRow.Cells("PROVIDER_TYPE").Value
                    Case WHCSHIP1.ProviderTypeFedex
                        If Not isInternationalShipment Then
                            clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpress)
                        Else
                            clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpressInternational)
                        End If
                    Case WHCSHIP1.ProviderTypeUPS
                        clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.UPS)
                    Case WHCSHIP1.ProviderTypeUSPS
                        clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.USPS)
                    Case WHCSHIP1.ProviderTypeCanada
                        clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.CanadaPost)
                    Case Else
                        EMsg &= vbCrLf & "Invalid or Missing Provider"
                        Exit Select
                End Select

            Case "Reprint Label"
                If Not printerFound AndAlso Not (ASCMAIN1.USER_ID = "edz" AndAlso ASCMAIN1.Running_in_VS) Then
                    EMsg &= vbCrLf & "Label printer was not found you may not Reprint Labels."
                    Exit Select
                End If

                If MessageBox.Show("Do you want to Reprint labels for the shipment displayed on the screen?", "Reprint Label", _
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Request Label"
                EntryMode = "S"
                If RequestShippingLabel() Then
                    Me.Mode_Settings(False)
                End If

            Case "Get Rates"
                RequestRates(tempSHIP_CNTL_NO, False)
                tabChoices.SelectedTab = tabChoices.Tabs("Carrier Rates")

            Case "Clear"
                Mode_Settings(False)

            Case "Cancel Shipment"
                CancelShipment()
                Mode_Settings(False)

            Case "Reprint Label"
                ReprintShippingLabel()
                Mode_Settings(False)

            Case "Request Return Label"

                ReturnLabelsToSendToCustomers.Clear()
                EntryMode = "S"
                Dim numLabels As String = dst.Tables("WHTSHPC2").Rows.Count

                If numLabels <= 0 Then
                    If optUserSuppliedValue.Value = "R" Then
                        Dim userSuppliedValue As String = txtUserSuppliedValue.Text.Trim.ToUpper
                        userSuppliedValue = userSuppliedValue.ToUpper.Trim
                        If userSuppliedValue.Length > 0 Then
                            userSuppliedValue = ASCMAIN1.Format_Field(userSuppliedValue, "RA_NO")
                        End If

                        rowSOTRMAF1 = LookUp("SOTRMAF1", userSuppliedValue)
                        If rowSOTRMAF1 IsNot Nothing Then
                            numLabels = Val(rowSOTRMAF1.Item("RA_CARTONS") & String.Empty)
                        End If
                    End If
                End If

                If numLabels <= 0 Then
                    numLabels = Val(InputBox("How many Return Labels do you want?", "Return Label", "1").Trim)
                    If numLabels <= 0 Then
                        MessageBox.Show("Invalid quantity.", "UPS Return Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    ElseIf numLabels > 20 Then
                        MessageBox.Show("Invalid quantity. Maximum request is 20 labels.", "UPS Return Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                End If

                If MessageBox.Show("Do you want to generate " & numLabels & " Return Labels?", "Return Labels", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                If dst.Tables("WHTSHPC2").Rows.Count = 0 Then
                    For iLoop As Int16 = 1 To numLabels
                        Dim rowWHTSHPC2 As DataRow = dst.Tables("WHTSHPC2").NewRow
                        rowWHTSHPC2.Item("SHIP_CNTL_NO") = dst.Tables("WHTSHPC1").Rows(0).Item("SHIP_CNTL_NO") & String.Empty
                        rowWHTSHPC2.Item("SHIP_PACKAGE_NO") = iLoop
                        'rowWHTSHPC2.Item("BASE_CHARGE") = String.Empty
                        'rowWHTSHPC2.Item("COD_AMOUNT") = String.Empty
                        'rowWHTSHPC2.Item("COD_TYPE") = String.Empty
                        'rowWHTSHPC2.Item("DANG_GOODS_ACCESS") = String.Empty
                        rowWHTSHPC2.Item("DESCRIPTION") = "Floral Product"
                        rowWHTSHPC2.Item("HEIGHT") = "12"
                        'rowWHTSHPC2.Item("INSURED_VALUE") = String.Empty
                        rowWHTSHPC2.Item("LENGTH") = "12"
                        'rowWHTSHPC2.Item("NET_CHARGE") = String.Empty
                        rowWHTSHPC2.Item("PACKAGING_TYPE") = "31"
                        'rowWHTSHPC2.Item("SIGNATURE_TYPE") = String.Empty
                        'rowWHTSHPC2.Item("TOTAL_DISCOUNT") = String.Empty
                        'rowWHTSHPC2.Item("TOTAL_SURCHARGES") = String.Empty
                        'rowWHTSHPC2.Item("TRACKING_NUMBER") = String.Empty
                        rowWHTSHPC2.Item("WEIGHT") = PKG_WEIGHT
                        rowWHTSHPC2.Item("WIDTH") = "12"
                        'rowWHTSHPC2.Item("TRACKING_NO") = String.Empty
                        rowWHTSHPC2.Item("PKG_CODE") = "OTHER"
                        'rowWHTSHPC2.Item("CUST_REF") = String.Empty
                        'rowWHTSHPC2.Item("INV_BOL_NO") = String.Empty
                        'rowWHTSHPC2.Item("INV_NO") = String.Empty
                        'rowWHTSHPC2.Item("PO_ORDER_NO") = String.Empty
                        'rowWHTSHPC2.Item("DEPT_NO") = String.Empty
                        'rowWHTSHPC2.Item("LIST_PRICE") = String.Empty
                        'rowWHTSHPC2.Item("CART_NO") = String.Empty
                        'rowWHTSHPC2.Item("STATUS") = String.Empty
                        'rowWHTSHPC2.Item("CANCEL_OPER") = String.Empty
                        'rowWHTSHPC2.Item("CANCEL_DATE") = String.Empty
                        rowWHTSHPC2.Item("POUNDS") = PKG_WEIGHT / 16
                        rowWHTSHPC2.Item("OUNCES") = 0
                        rowWHTSHPC2.Item("SEL") = "1"

                        dst.Tables("WHTSHPC2").Rows.Add(rowWHTSHPC2)
                    Next

                End If

                For Each rowWHTSHPC2 As DataRow In dst.Tables("WHTSHPC2").Select("SEL = '1'", "SHIP_PACKAGE_NO")
                    If Not RequestReturnLabel(rowWHTSHPC2.Item("SHIP_PACKAGE_NO")) Then
                        ASCMAIN1.Progress("", "")
                        Exit Sub
                    End If
                Next

                Select Case optReturnLabels.Value
                    Case EmailLabelsToCustomer
                        EmailReturnLabels()

                        For Each fileName As String In ReturnLabelsToSendToCustomers
                            Try
                                My.Computer.FileSystem.DeleteFile(fileName)
                            Catch ex As Exception

                            End Try
                        Next

                End Select

                Mode_Settings(False)

                ASCMAIN1.Progress("", "")
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Request Label").Settings.Enabled = DefaultableBoolean.True
                .Groups("Screen Control").Items("Clear").Settings.Enabled = DefaultableBoolean.True
                .Groups("Screen Control").Items("Cancel Shipment").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Reprint Label").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Request Return Label").Settings.Enabled = DefaultableBoolean.True

                If cmbShipMethod.SelectedRow IsNot Nothing AndAlso Val(cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value & String.Empty) = FedexSmartPost Then
                    .Groups("Smart Post").Visible = True
                Else
                    .Groups("Smart Post").Visible = False
                End If
            End With
        End If

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Private Sub Clear_Record()
        MyBase.EnforceConstraints(False)

        For Each tableName As String In New String() {"WHTSHPC1", "WHTSHPC2", "WHTSHPC5_SF", _
                                                      "WHTSHPC5_ST", "WHTSHPC5_HL", "WHTSHPCC", _
                                                      "WHTSHPCP_S", "WHTSHPCP_D", "WHTSHPCS", "WHTSHPC4", _
                                                      "SOTRMAFL", "ASTATTA2"}
            dst.Tables(tableName).Rows.Clear()
        Next

        For Each row As DataRow In dst.Tables("SHTSHIPS").Select
            row.Item("SEL") = "0"
        Next
        dst.Tables("SHTSHIPS").AcceptChanges()

        ' Create Shells for data
        ' Create working header record for WHTSHPC1
        SHIP_CNTL_NO = "XX"

        grdWHTSHPC2.DisplayLayout.Bands(0).Columns("SEL").Hidden = True

        For Each tableName As String In New String() {"WHTSHPC5_SF", "WHTSHPC5_ST", "WHTSHPC5_HL", "WHTSHPCP_S", "WHTSHPCP_D"}
            Dim row As DataRow = dst.Tables(tableName).NewRow
            row.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO

            Select Case tableName
                Case "WHTSHPC5_SF" : row.Item("SHIP_ADDR_TYPE") = "SF"
                Case "WHTSHPC5_ST" : row.Item("SHIP_ADDR_TYPE") = "ST"
                Case "WHTSHPC5_HL" : row.Item("SHIP_ADDR_TYPE") = "HL"
                Case "WHTSHPCP_S" : row.Item("PAYOR_TYPE") = "S"
                Case "WHTSHPCP_D" : row.Item("PAYOR_TYPE") = "D"
            End Select
            dst.Tables(tableName).Rows.Add(row)
        Next

        MyBase.EnforceConstraints(True)

        Show_Filter(grdWHTSHPC2, False)
        grdWHTSHPC2.DisplayLayout.GroupByBox.Hidden = True

        Show_Filter(grdSHTSHPCS, False)
        grdSHTSHPCS.DisplayLayout.GroupByBox.Hidden = True

        chkFromResidential.Checked = False
        chkToResidential.Checked = False

        chkToResidential.Checked = False
        chkToPOBox.Checked = False
        chkFromResidential.Checked = False
        chkSignature.Checked = False

        rowWHTSHPC1 = dst.Tables("WHTSHPC1").NewRow
        rowWHTSHPC1.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO

        ' Keep previously selected options.
        If cmbProvider.SelectedRow IsNot Nothing Then
            cmbDivision_ValueChanged(Nothing, Nothing)
        End If

        If cmbProvider.SelectedRow IsNot Nothing Then
            rowWHTSHPC1.Item("CARRIER_CODE") = cmbProvider.SelectedRow.Cells("CARRIER_CODE").Value
        Else
            rowWHTSHPC1.Item("CARRIER_CODE") = "*"
        End If

        If cmbShipMethod.SelectedRow IsNot Nothing Then
            rowWHTSHPC1.Item("CARRIER_PROD_CODE") = cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value
            rowWHTSHPC1.Item("SHIP_VIA_CODE") = cmbShipMethod.SelectedRow.Cells("SHIP_VIA_CODE").Value
        Else
            rowWHTSHPC1.Item("CARRIER_PROD_CODE") = 0
        End If

        rowWHTSHPC1.Item("CARRIER_ACCOUNT_NO") = String.Empty
        rowWHTSHPC1.Item("STATUS") = "P"
        'rowWHTSHPC1.Item("ERROR_MSG") = String.Empty
        rowWHTSHPC1.Item("SHIP_DATE") = DateTime.Now.ToString("MM/dd/yyyy")
        dteShipDate.Value = DateTime.Now
        rowWHTSHPC1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        rowWHTSHPC1.Item("OPS_YYYYWW") = ASCMAIN1.CYW
        'rowWHTSHPC1.Item("CUST_CODE") = String.Empty
        rowWHTSHPC1.Item("INIT_DATE") = DateTime.Now
        rowWHTSHPC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowWHTSHPC1.Item("LAST_DATE") = DateTime.Now
        rowWHTSHPC1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        'rowWHTSHPC1.Item("MASTER_TRACKING_NO") = String.Empty
        rowWHTSHPC1.Item("CUSTOMS_VALUE") = 0
        dst.Tables("WHTSHPC1").Rows.Add(rowWHTSHPC1)

        Dim rowWHTSHPC5_SF As DataRow
        rowWHTSHPC5_SF = dst.Tables("WHTSHPC5_SF").Rows(0)
        rowWHTSHPC5_SF.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
        rowWHTSHPC5_SF.Item("SHIP_ADDR_TYPE") = "SF"
        cmbWarehouse_ValueChanged(Nothing, Nothing)

        txtUserSuppliedValue.Clear()
        MASTER_TRACKING_NO = String.Empty

        txtShipAccountNo.Clear()
        txtShipAccountZip.Clear()
        txtShipCountry.Clear()
        cmbShipPayor.Value = "S"

        txtDutiesAccountNo.Clear()
        txtDutiesAccountZip.Clear()
        txtDutiesCountry.Clear()
        cmbDutiesPayor.Value = "S"
        rowSOTRMAF1 = Nothing

        txtToEmail.Clear()
        optReturnLabels.CheckedIndex = -1
        ReturnLabelsToSendToCustomers.Clear()

        optReturnLabels.Value = "E"

        If ASCMAIN1.CLIENT = "VAN" Then
            optPrint_Type.Value = "X"
        End If

        SetReadOnly(False)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("")

        MyBase.EnforceConstraints(False)


        MyBase.EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()
        Try
            BeginTrans()
            Select Case EntryMode

                Case "S" ' Shipping label

                    Update_Record_TDA("WHTSHPC1", "SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "'")
                    Update_Record_TDA("WHTSHPC2", "SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "'")

                    Update_Record_TDA("WHTSHPC5_SF", "DELETE FROM WHTSHPC5 WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "' AND SHIP_ADDR_TYPE = 'SF'")
                    Update_Record_TDA("WHTSHPC5_ST", "DELETE FROM WHTSHPC5 WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "' AND SHIP_ADDR_TYPE = 'ST'")
                    Update_Record_TDA("WHTSHPC5_HL", "DELETE FROM WHTSHPC5 WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "' AND SHIP_ADDR_TYPE = 'HL'")

                    Dim row As DataRow = dst.Tables("WHTSHPCP_S").Select("PAYOR_TYPE = 'S'")(0)
                    row.Item("PAYOR_ACCT_NO") = txtShipAccountNo.Text
                    row.Item("PAYOR_COUNTRY") = txtShipCountry.Text
                    row.Item("PAYOR_ACCT_ZIP") = txtShipAccountZip.Text

                    row = dst.Tables("WHTSHPCP_D").Select("PAYOR_TYPE = 'D'")(0)
                    row.Item("PAYOR_ACCT_NO") = txtDutiesAccountNo.Text
                    row.Item("PAYOR_COUNTRY") = txtDutiesCountry.Text
                    row.Item("PAYOR_ACCT_ZIP") = txtDutiesAccountZip.Text

                    Update_Record_TDA("WHTSHPCP_S", "DELETE FROM WHTSHPCP WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "' AND PAYOR_TYPE = 'S'")
                    Update_Record_TDA("WHTSHPCP_D", "DELETE FROM WHTSHPCP WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "' AND PAYOR_TYPE = 'D'")

                    For Each rowWHTSHPCC As DataRow In dst.Tables("WHTSHPCC").Select("", "", DataViewRowState.CurrentRows)
                        If rowWHTSHPCC.Item("MANUFACTURER") & String.Empty = String.Empty Then
                            rowWHTSHPCC.Item("MANUFACTURER") = "US"
                        End If
                    Next
                    Update_Record_TDA("WHTSHPCC", "SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "'")

                    dst.Tables("WHTSHPCS").Rows.Clear()
                    For Each row In dst.Tables("SHTSHIPS").Select("SEL = '1'")
                        dst.Tables("WHTSHPCS").Rows.Add(New Object() {SHIP_CNTL_NO, row.Item("SPCL_SHIP_CODE")})
                    Next
                    Update_Record_TDA("WHTSHPCS", "SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "'")

            End Select

            CommitTrans()

        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Form Controls"

    Private Sub WHFSHIP1_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        SetUpPortsAndPrinters()
    End Sub

    Private Sub cmbDivision_ValueChanged(sender As Object, e As System.EventArgs) Handles cmbDivision.ValueChanged
        providerView = New DataView(dst.Tables("SOTCARR3"))
        providerView.Sort = "CARRIER_DESC_DISP"
        providerView.RowFilter = "DIVISION_CODE = ''"
        cmbProvider.DataSource = providerView
        If cmbDivision.SelectedRow IsNot Nothing Then
            providerView.RowFilter = "DIVISION_CODE = '" & cmbDivision.SelectedRow.Cells("DIVISION_CODE").Value & "'"
            cmbProvider.SelectedRow = cmbProvider.Rows(0)
            cmbProvider.Text = cmbProvider.SelectedRow.Cells("CARRIER_DESC_DISP").Value
        End If
        cmbProvider_ValueChanged(Nothing, Nothing)
    End Sub

    Private Sub cmbProvider_ValueChanged(sender As Object, e As System.EventArgs) Handles cmbProvider.ValueChanged
        shipmethodView = New DataView(dst.Tables("SOTSVIA1"))
        shipmethodView.Sort = "SHIP_VIA_DESC"
        shipmethodView.RowFilter = "SHIP_VIA_CODE = ''"
        cmbShipMethod.DataSource = shipmethodView
        If cmbProvider.SelectedRow IsNot Nothing Then
            shipmethodView.RowFilter = "CARRIER_CODE = '" & cmbProvider.SelectedRow.Cells("CARRIER_CODE").Value & "' AND SHIP_VIA_STATUS = 'A' AND ISNULL(CARRIER_PROD_CODE , '') <> ''"
            If cmbShipMethod.Rows.Count > 0 Then
                cmbShipMethod.SelectedRow = cmbShipMethod.Rows(0)
                cmbShipMethod.Text = cmbShipMethod.SelectedRow.Cells("SHIP_VIA_DESC").Value
            End If
            ASCMAIN1.Add_Value_List(grdWHTSHPC2, "PACKAGING_TYPE", "SELECT PACKAGE_CODE, PACKAGE_DESC FROM SOTCARR4 WHERE CARRIER_CODE = '" & cmbProvider.SelectedRow.Cells("CARRIER_CODE").Value & "'")
            ASCMAIN1.Add_Value_List(grdWHTSHPCC, "QUANTITY_UOM", "SELECT CARRIER_UOM, CARRIER_UOM_DESC FROM SOTCARRU WHERE CARRIER_CODE = '" & cmbProvider.SelectedRow.Cells("CARRIER_CODE").Value & "'")
        Else
            ASCMAIN1.Add_Value_List(grdWHTSHPC2, "PACKAGING_TYPE", "SELECT PACKAGE_CODE, PACKAGE_DESC FROM SOTCARR4 WHERE CARRIER_CODE = ''")
            ASCMAIN1.Add_Value_List(grdWHTSHPCC, "QUANTITY_UOM", "SELECT CARRIER_UOM, CARRIER_UOM_DESC FROM SOTCARRU WHERE CARRIER_CODE = ''")
        End If

    End Sub

    Private Sub cmbShipMethod_ValueChanged(sender As Object, e As System.EventArgs) Handles cmbShipMethod.ValueChanged
        If cmbShipMethod.SelectedRow IsNot Nothing AndAlso dst.Tables("WHTSHPC1").Rows.Count > 0 Then
            dst.Tables("WHTSHPC1").Rows(0).Item("CARRIER_CODE") = cmbShipMethod.SelectedRow.Cells("CARRIER_CODE").Value & String.Empty
        End If

        If cmbShipMethod.SelectedRow IsNot Nothing AndAlso Val(cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value & String.Empty) = FedexSmartPost Then
            UltraExplorerBar1.Groups("Smart Post").Visible = True
        Else
            UltraExplorerBar1.Groups("Smart Post").Visible = False
        End If
    End Sub

    Private Sub grdWHTSHPC2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTSHPC2.BeforeRowUpdate

        EMsg = String.Empty

        ValidatePacking(e.Row)

        Dim rowWHTPKGM1 As DataRow = Nothing
        Dim PKG_CODE As String = e.Row.Cells("PKG_CODE").Value & String.Empty

        If dst.Tables("WHTPKGM1").Select("PKG_CODE = '" & PKG_CODE & "' AND PKG_CODE <> 'OTHER'").Length > 0 Then
            rowWHTPKGM1 = dst.Tables("WHTPKGM1").Select("PKG_CODE = '" & PKG_CODE & "'")(0)
            e.Row.Cells("LENGTH").Value = rowWHTPKGM1.Item("PKG_L")
            e.Row.Cells("WIDTH").Value = rowWHTPKGM1.Item("PKG_W")
            e.Row.Cells("HEIGHT").Value = rowWHTPKGM1.Item("PKG_H")
        End If

        If EMsg.Length > 0 Then
            e.Cancel = True
            MessageBox.Show(EMsg, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

    Private Sub grdWHTSHPCC_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTSHPCC.BeforeRowUpdate
        EMsg = String.Empty

        ValidateCommodity(e.Row)

        If EMsg.Length > 0 Then
            e.Cancel = True
            MessageBox.Show(EMsg, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub grdWHTSHPCC_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdWHTSHPCC.InitializeLayout
        Dim UC As New UltraWinGrid.UltraCombo
        UC.DataSource = dst.Tables("WHTSDESC")
        UC.DisplayLayout.Bands(0).ColHeadersVisible = False
        UC.DisplayLayout.Bands(0).Columns("SHIP_DESC").Width = e.Layout.Bands(0).Columns("COMMODITY_DESC").Width
        UC.Font = grdWHTSHPCC.Font

        e.Layout.Bands(0).Columns("COMMODITY_DESC").EditorComponent = UC
    End Sub

    Private Sub ultraComboPackage_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ultraComboPackage.ValueChanged

        Dim vl As IValueList = Me.ultraComboPackage

        If vl.SelectedItemIndex < 0 Then
            ' NOTHING
        Else
            Dim PKG_CODE As String = vl.GetText(vl.SelectedItemIndex)
            Dim rowWHTPKGM1 As DataRow = LookUp("WHTPKGM1", PKG_CODE)

            If rowWHTPKGM1 IsNot Nothing Then
                grdWHTSHPC2.ActiveRow.Cells("LENGTH").Value = Val(rowWHTPKGM1.Item("PKG_L") & String.Empty)
                grdWHTSHPC2.ActiveRow.Cells("WIDTH").Value = Val(rowWHTPKGM1.Item("PKG_W") & String.Empty)
                grdWHTSHPC2.ActiveRow.Cells("HEIGHT").Value = Val(rowWHTPKGM1.Item("PKG_H") & String.Empty)
            End If
        End If

    End Sub

    Private Sub txtLoadAddress_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtUserSuppliedValue.KeyPress

        If Asc(e.KeyChar) = 13 Then
            rowSOTRMAF1 = Nothing
            Dim rowWHTSHPC5_ST As DataRow = dst.Tables("WHTSHPC5_ST").Rows(0)
            Dim userSuppliedValue As String = txtUserSuppliedValue.Text.Trim.ToUpper
            If userSuppliedValue.Length = 0 Then Exit Sub

            Dim KeyPressEventArgs As New System.Windows.Forms.KeyPressEventArgs(vbCrLf)

            txtUserSuppliedValue.Text = userSuppliedValue
            Dim rowARTCUST1 As DataRow = Nothing

            Select Case optUserSuppliedValue.Value

                Case "P" ' Pick Ticket
                    userSuppliedValue = ASCMAIN1.Format_Field(userSuppliedValue, "PICK_NO")
                    Dim rowSOTPICK1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTPICK1 WHERE PICK_NO = :PARM1", "V", New Object() {userSuppliedValue})
                    If rowSOTPICK1 Is Nothing OrElse rowSOTPICK1.Item("INV_NO") & String.Empty = String.Empty Then
                        MessageBox.Show("Invalid or unprocessed Pick Ticket Number", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                    optUserSuppliedValue.Value = "I"
                    txtUserSuppliedValue.Text = rowSOTPICK1.Item("INV_NO")
                    txtLoadAddress_KeyPress(Nothing, KeyPressEventArgs)
                    optUserSuppliedValue.Value = "P"
                    txtUserSuppliedValue.Text = userSuppliedValue
                    Exit Sub

                Case "W"
                    Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTINVH1 WHERE ORDR_NO_WEB = :PARM1", "V", New Object() {userSuppliedValue})
                    If rowSOTINVH1 Is Nothing Then
                        MessageBox.Show("Invalid Web Order Number", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                    optUserSuppliedValue.Value = "I"
                    txtUserSuppliedValue.Text = rowSOTINVH1.Item("INV_NO")
                    txtLoadAddress_KeyPress(Nothing, KeyPressEventArgs)
                    Exit Sub

                Case "T" ' Tracking No
                    Dim tbl As DataTable = ASCDATA1.GetDataTable("select * from WHTSHPC1 where MASTER_TRACKING_NO = '" & userSuppliedValue & "'")

                    If tbl.Rows.Count = 0 Then
                        ' See if it is part of a Milti package
                        tbl = ASCDATA1.GetDataTable("select * from WHTSHPC2 where TRACKING_NO = '" & userSuppliedValue & "'")
                    End If

                    ' Done for fedex, the scan contains extra data
                    If tbl.Rows.Count = 0 Then
                        ' Scanning in the Tracking Number reads in the additional leading characters.
                        If userSuppliedValue.Length > 15 Then
                            userSuppliedValue = userSuppliedValue.Substring(userSuppliedValue.Length - 15)
                            txtUserSuppliedValue.Text = userSuppliedValue
                        End If
                        tbl = ASCDATA1.GetDataTable("select * from WHTSHPC1 where MASTER_TRACKING_NO = '" & userSuppliedValue & "'")
                    End If

                    If tbl.Rows.Count = 0 Then
                        MessageBox.Show("The provided Tracking No could not be found.", "Load", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    SHIP_CNTL_NO = tbl.Rows(0).Item("SHIP_CNTL_NO") & String.Empty
                    tbl = ASCDATA1.GetDataTable("select * from WHTSHPC1 where SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "'")

                    ' Valid values for WHTSHPC1.STATUS
                    '   I - Initial Setup before calling request.
                    '   P - processed - label printed
                    '   C - Cancelled

                    If tbl.Rows(0).Item("STATUS") & String.Empty = "P" Then
                        UltraExplorerBar1.Groups("Screen Control").Items("Cancel Shipment").Settings.Enabled = DefaultableBoolean.True
                        UltraExplorerBar1.Groups("Screen Control").Items("Reprint Label").Settings.Enabled = DefaultableBoolean.True
                        UltraExplorerBar1.Groups("Screen Control").Items("Request Label").Settings.Enabled = DefaultableBoolean.False
                        grdWHTSHPC2.DisplayLayout.Bands(0).Columns("SEL").Hidden = False
                        grdWHTSHPC2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                    Else
                        UltraExplorerBar1.Groups("Screen Control").Items("Cancel Shipment").Settings.Enabled = DefaultableBoolean.False
                        UltraExplorerBar1.Groups("Screen Control").Items("Reprint Label").Settings.Enabled = DefaultableBoolean.False
                        UltraExplorerBar1.Groups("Screen Control").Items("Request Label").Settings.Enabled = DefaultableBoolean.True
                        grdWHTSHPC2.DisplayLayout.Bands(0).Columns("SEL").Hidden = True
                        grdWHTSHPC2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    End If

                    Fill_Records("WHTSHPC1", SHIP_CNTL_NO)
                    Fill_Records("WHTSHPC2", SHIP_CNTL_NO)
                    Fill_Records("WHTSHPCC", SHIP_CNTL_NO)
                    Fill_Records("WHTSHPCS", SHIP_CNTL_NO)

                    For Each addressType As String In New String() {"SF", "ST", "HL"}
                        Fill_Records("WHTSHPC5_" & addressType, "", True, "SELECT * FROM WHTSHPC5 WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "' AND SHIP_ADDR_TYPE = '" & addressType & "'")
                    Next

                    For Each payorType As String In New String() {"D", "S"}
                        Fill_Records("WHTSHPCP_" & payorType, "", True, "SELECT * FROM WHTSHPCP WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "' AND PAYOR_TYPE = '" & payorType & "'")
                    Next

                    Sort_grdColumns(grdWHTSHPC2, "SHIP_PACKAGE_NO")
                    Sort_grdColumns(grdWHTSHPCC, "COMMODITY_DESC")

                    For Each rowWHTSHPC2 As DataRow In dst.Tables("WHTSHPC2").Select()
                        Dim WEIGHT As Int16 = Val(rowWHTSHPC2.Item("WEIGHT") & String.Empty)
                        Dim POUNDS As Int16 = WEIGHT \ 16
                        Dim OUNCES As Int16 = WEIGHT Mod 16

                        rowWHTSHPC2.Item("POUNDS") = POUNDS
                        rowWHTSHPC2.Item("OUNCES") = OUNCES
                    Next

                    Dim CARRIER_CODE As String = dst.Tables("WHTSHPC1").Rows(0).Item("CARRIER_CODE") & String.Empty
                    For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In cmbProvider.Rows
                        If row.Cells("CARRIER_CODE").Value = CARRIER_CODE Then
                            cmbProvider.SelectedRow = row
                            Exit For
                        End If
                    Next

                    Dim SHIP_VIA_CODE As String = dst.Tables("WHTSHPC1").Rows(0).Item("SHIP_VIA_CODE") & String.Empty
                    For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In cmbShipMethod.Rows
                        If row.Cells("SHIP_VIA_CODE").Value = SHIP_VIA_CODE Then
                            cmbShipMethod.SelectedRow = row
                            Exit For
                        End If
                    Next

                    If Not (ASCMAIN1.USER_SECURITY_CODEs.Contains("SY") OrElse ASCMAIN1.USER_SECURITY_CODEs.Contains("WL")) Then
                        SetReadOnly(True)
                    End If

                Case "C" ' Customer
                    userSuppliedValue = userSuppliedValue.ToUpper.Trim
                    rowARTCUST1 = LookUp("ARTCUST1", userSuppliedValue)
                    If rowARTCUST1 Is Nothing Then
                        MessageBox.Show("Invalid Customer code.", "Error", MessageBoxButtons.OK)
                        Exit Sub
                    End If

                    rowWHTSHPC5_ST.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    rowWHTSHPC5_ST.Item("SHIP_ADDR_TYPE") = "ST"
                    rowWHTSHPC5_ST.Item("SHIP_FIRST_NAME") = rowARTCUST1.Item("CUST_CONTACT")
                    rowWHTSHPC5_ST.Item("SHIP_COMPANY") = rowARTCUST1.Item("CUST_NAME") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_ADDR1") = rowARTCUST1.Item("CUST_ADDR1") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_ADDR2") = rowARTCUST1.Item("CUST_ADDR2") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_CITY") = rowARTCUST1.Item("CUST_CITY") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_STATE") = rowARTCUST1.Item("CUST_STATE") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_ZIP_CODE") = rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_COUNTRY_CODE") = rowARTCUST1.Item("CUST_COUNTRY") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_PHONE") = rowARTCUST1.Item("CUST_PHONE") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_RESIDENTIAL") = "0"
                    rowWHTSHPC5_ST.Item("SHIP_PO_BOX") = "0"
                    rowWHTSHPC5_ST.Item("SHIP_EMAIL") = rowARTCUST1.Item("CUST_EMAIL") & String.Empty
                    If rowWHTSHPC5_ST.Item("SHIP_COUNTRY_CODE") & String.Empty = String.Empty Then rowWHTSHPC5_ST.Item("SHIP_COUNTRY_CODE") = "US"
                    rowWHTSHPC5_ST.AcceptChanges()
                    txtUserSuppliedValue.Text = userSuppliedValue

                Case "V" ' Vendor Address
                    userSuppliedValue = userSuppliedValue.ToUpper.Trim
                    Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", userSuppliedValue)
                    If rowAPTVEND1 Is Nothing Then
                        MessageBox.Show("Invalid Vendor code.", "Error", MessageBoxButtons.OK)
                        Exit Sub
                    End If

                    rowWHTSHPC5_ST.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    rowWHTSHPC5_ST.Item("SHIP_ADDR_TYPE") = "ST"
                    rowWHTSHPC5_ST.Item("SHIP_FIRST_NAME") = rowAPTVEND1.Item("VEND_CONTACT")
                    rowWHTSHPC5_ST.Item("SHIP_COMPANY") = rowAPTVEND1.Item("VEND_NAME") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_ADDR1") = rowAPTVEND1.Item("VEND_ADDR1") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_ADDR2") = rowAPTVEND1.Item("VEND_ADDR2") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_CITY") = rowAPTVEND1.Item("VEND_CITY") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_STATE") = rowAPTVEND1.Item("VEND_STATE") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_ZIP_CODE") = rowAPTVEND1.Item("VEND_ZIP_CODE") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_COUNTRY_CODE") = rowAPTVEND1.Item("VEND_COUNTRY") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_PHONE") = rowAPTVEND1.Item("VEND_PHONE") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_RESIDENTIAL") = "0"
                    rowWHTSHPC5_ST.Item("SHIP_PO_BOX") = "0"
                    rowWHTSHPC5_ST.Item("SHIP_EMAIL") = rowAPTVEND1.Item("VEND_EMAIL") & String.Empty
                    If rowWHTSHPC5_ST.Item("SHIP_COUNTRY_CODE") & String.Empty = String.Empty Then rowWHTSHPC5_ST.Item("SHIP_COUNTRY_CODE") = "US"
                    rowWHTSHPC5_ST.AcceptChanges()
                    txtUserSuppliedValue.Text = userSuppliedValue

                Case "R" ' RMA 
                    userSuppliedValue = userSuppliedValue.ToUpper.Trim
                    If userSuppliedValue.Length > 0 Then
                        userSuppliedValue = ASCMAIN1.Format_Field(userSuppliedValue, "RA_NO")
                    End If

                    rowSOTRMAF1 = LookUp("SOTRMAF1", userSuppliedValue)
                    Dim CUST_CODE As String = String.Empty
                    Dim CUST_STORE_NO As String = String.Empty

                    If rowSOTRMAF1 Is Nothing Then
                        MessageBox.Show("Invalid RMA Number.", "Error", MessageBoxButtons.OK)
                        Exit Sub
                    End If

                    If rowSOTRMAF1.Item("RA_STATUS") & String.Empty <> "O" Then
                        If MessageBox.Show("The provided RMA is not Open. Do you still want to proceed?", "RMA No", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    End If

                    rowWHTSHPC5_ST.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    rowWHTSHPC5_ST.Item("SHIP_ADDR_TYPE") = "ST"

                    CUST_CODE = rowSOTRMAF1.Item("CUST_CODE") & String.Empty
                    CUST_STORE_NO = rowSOTRMAF1.Item("CUST_STORE_NO") & String.Empty

                    rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                    Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO})

                    If rowARTCUST2 IsNot Nothing Then
                        rowWHTSHPC5_ST.Item("SHIP_FIRST_NAME") = rowARTCUST2.Item("CUST_CONTACT")
                        rowWHTSHPC5_ST.Item("SHIP_COMPANY") = rowARTCUST2.Item("CUST_NAME") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_ADDR1") = rowARTCUST2.Item("CUST_ADDR1") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_ADDR2") = rowARTCUST2.Item("CUST_ADDR2") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_CITY") = rowARTCUST2.Item("CUST_CITY") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_STATE") = rowARTCUST2.Item("CUST_STATE") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_ZIP_CODE") = rowARTCUST2.Item("CUST_ZIP_CODE") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_COUNTRY_CODE") = rowARTCUST2.Item("CUST_COUNTRY") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_PHONE") = rowARTCUST2.Item("CUST_PHONE") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_RESIDENTIAL") = "0"
                        rowWHTSHPC5_ST.Item("SHIP_PO_BOX") = "0"

                        rowWHTSHPC5_ST.Item("SHIP_EMAIL") = rowARTCUST2.Item("CUST_EMAIL") & String.Empty

                        If rowARTCUST2.Item("CUST_EMAIL") & String.Empty = String.Empty _
                            AndAlso rowARTCUST1 IsNot Nothing Then
                            rowWHTSHPC5_ST.Item("SHIP_EMAIL") = rowARTCUST1.Item("CUST_EMAIL") & String.Empty
                        End If

                    ElseIf rowARTCUST1 IsNot Nothing Then
                        rowWHTSHPC5_ST.Item("SHIP_FIRST_NAME") = rowARTCUST1.Item("CUST_CONTACT")
                        rowWHTSHPC5_ST.Item("SHIP_COMPANY") = rowARTCUST1.Item("CUST_NAME") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_ADDR1") = rowARTCUST1.Item("CUST_ADDR1") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_ADDR2") = rowARTCUST1.Item("CUST_ADDR2") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_CITY") = rowARTCUST1.Item("CUST_CITY") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_STATE") = rowARTCUST1.Item("CUST_STATE") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_ZIP_CODE") = rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_COUNTRY_CODE") = rowARTCUST1.Item("CUST_COUNTRY") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_PHONE") = rowARTCUST1.Item("CUST_PHONE") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_RESIDENTIAL") = "0"
                        rowWHTSHPC5_ST.Item("SHIP_PO_BOX") = "0"
                        rowWHTSHPC5_ST.Item("SHIP_EMAIL") = rowARTCUST1.Item("CUST_EMAIL") & String.Empty
                    End If

                    If rowSOTRMAF1.Item("RA_EMAIL") & String.Empty <> String.Empty Then
                        rowWHTSHPC5_ST.Item("SHIP_EMAIL") = rowSOTRMAF1.Item("RA_EMAIL") & String.Empty
                    End If

                    If rowWHTSHPC5_ST.Item("SHIP_COUNTRY_CODE") & String.Empty = String.Empty Then rowWHTSHPC5_ST.Item("SHIP_COUNTRY_CODE") = "US"
                    rowWHTSHPC5_ST.AcceptChanges()
                    txtUserSuppliedValue.Text = userSuppliedValue

                Case "I", "S" ' Sales Order 
                    Dim rowSOTINVH1 As DataRow = Nothing
                    Dim SHIP_BOL_NO As String = String.Empty

                    userSuppliedValue = userSuppliedValue.ToUpper.Trim
                    If optUserSuppliedValue.Value = "I" Then
                        userSuppliedValue = ASCMAIN1.Format_Field(userSuppliedValue, "INV_NO")
                        rowSOTINVH1 = LookUp("SOTINVH1", New String() {"I", userSuppliedValue})
                        If rowSOTINVH1 Is Nothing Then
                            MessageBox.Show("Invalid Invoice Number.", "Error", MessageBoxButtons.OK)
                            Exit Sub
                        End If
                        txtUserSuppliedValue.Text = userSuppliedValue
                        userSuppliedValue = rowSOTINVH1.Item("ORDR_NO") & String.Empty
                        SHIP_BOL_NO = rowSOTINVH1.Item("SHIP_BOL_NO") & String.Empty

                        If SHIP_BOL_NO.Length > 0 Then
                            cmbWarehouse.Text = rowSOTINVH1.Item("WHSE_CODE") & String.Empty

                            Dim rowWHTSIP1X As DataRow = ASCDATA1.GetDataRow("SELECT * FROM WHTSHPC1 WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
                            If rowWHTSIP1X IsNot Nothing Then
                                Dim MASTER_TRACKING_NO As String = rowWHTSIP1X.Item("MASTER_TRACKING_NO") & String.Empty
                                If MASTER_TRACKING_NO.Length > 0 Then
                                    optUserSuppliedValue.Value = "T"
                                    txtUserSuppliedValue.Text = MASTER_TRACKING_NO
                                    txtLoadAddress_KeyPress(Nothing, KeyPressEventArgs)
                                    optUserSuppliedValue.Value = "I"
                                    txtUserSuppliedValue.Text = userSuppliedValue
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If

                    userSuppliedValue = ASCMAIN1.Format_Field(userSuppliedValue, "ORDR_NO")
                    Dim rowSOTORDR5 As DataRow = LookUp("SOTORDR5", New String() {userSuppliedValue, "ST"})
                    If rowSOTORDR5 Is Nothing Then
                        rowSOTORDR5 = LookUp("SOTORDR5", New String() {userSuppliedValue, "BT"})
                    End If
                    If rowSOTORDR5 Is Nothing Then
                        MessageBox.Show("Invalid Sales Order.", "Error", MessageBoxButtons.OK)
                        Exit Sub
                    End If

                    rowWHTSHPC5_ST.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    rowWHTSHPC5_ST.Item("SHIP_ADDR_TYPE") = "ST"
                    rowWHTSHPC5_ST.Item("SHIP_FIRST_NAME") = rowSOTORDR5.Item("CUST_CONTACT")
                    rowWHTSHPC5_ST.Item("SHIP_COMPANY") = rowSOTORDR5.Item("CUST_NAME") & String.Empty
                    If (rowSOTORDR5.Item("CUST_ADDR1") & String.Empty).ToString.Trim.Length = 0 Then
                        rowWHTSHPC5_ST.Item("SHIP_ADDR1") = rowSOTORDR5.Item("CUST_ADDR2") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_ADDR2") = rowSOTORDR5.Item("CUST_ADDR3") & String.Empty
                    Else
                        rowWHTSHPC5_ST.Item("SHIP_ADDR1") = rowSOTORDR5.Item("CUST_ADDR1") & String.Empty
                        rowWHTSHPC5_ST.Item("SHIP_ADDR2") = rowSOTORDR5.Item("CUST_ADDR2") & String.Empty
                    End If
                    rowWHTSHPC5_ST.Item("SHIP_CITY") = rowSOTORDR5.Item("CUST_CITY") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_STATE") = rowSOTORDR5.Item("CUST_STATE") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_ZIP_CODE") = rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_COUNTRY_CODE") = rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_PHONE") = rowSOTORDR5.Item("CUST_PHONE") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_EMAIL") = rowSOTORDR5.Item("CUST_EMAIL") & String.Empty
                    rowWHTSHPC5_ST.Item("SHIP_RESIDENTIAL") = "0"
                    rowWHTSHPC5_ST.Item("SHIP_PO_BOX") = "0"
                    If rowWHTSHPC5_ST.Item("SHIP_COUNTRY_CODE") & String.Empty = String.Empty Then rowWHTSHPC5_ST.Item("SHIP_COUNTRY_CODE") = "US"
                    rowWHTSHPC5_ST.AcceptChanges()

                    If optUserSuppliedValue.Value = "S" Then
                        txtUserSuppliedValue.Text = userSuppliedValue
                    End If

                    If rowSOTINVH1 IsNot Nothing AndAlso dst.Tables("WHTSHPCC").Select("", "", DataViewRowState.CurrentRows).Length = 0 Then
                        Dim COMMODITY_LNO As Int16 = 1
                        For Each rowSOTINVH2 As DataRow In ASCDATA1.GetDataTable("SELECT * FROM SOTINVH2 WHERE INV_NO = '" & rowSOTINVH1.Item("INV_NO") & "'").Rows
                            Dim rowWHTSHPCC As DataRow = dst.Tables("WHTSHPCC").NewRow
                            rowWHTSHPCC.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                            rowWHTSHPCC.Item("COMMODITY_LNO") = COMMODITY_LNO
                            COMMODITY_LNO += 1

                            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowSOTINVH2.Item("STYLE_CODE"))
                            rowWHTSHPCC.Item("COMMODITY_DESC") = rowICTSTYL1.Item("STYLE_DESC") & String.Empty
                            rowWHTSHPCC.Item("NUM_PIECES") = Val(rowSOTINVH2.Item("ORDR_QTY_SHIP") & String.Empty)
                            rowWHTSHPCC.Item("MANUFACTURER") = rowICTSTYL1.Item("COUNTRY_CODE") & String.Empty
                            If rowWHTSHPCC.Item("MANUFACTURER") & String.Empty = String.Empty Then
                                rowWHTSHPCC.Item("MANUFACTURER") = "US"
                            End If
                            rowWHTSHPCC.Item("HARMONIZED_CODE") = ""
                            rowWHTSHPCC.Item("WEIGHT") = Val(rowICTSTYL1.Item("STYLE_WEIGHT") & String.Empty)
                            rowWHTSHPCC.Item("QUANTITY") = Val(rowSOTINVH2.Item("ORDR_QTY_SHIP") & String.Empty)
                            rowWHTSHPCC.Item("QUANTITY_UOM") = "EA"
                            rowWHTSHPCC.Item("UNIT_PRICE") = Val(rowSOTINVH2.Item("ORDR_UNIT_PRICE") & String.Empty)
                            dst.Tables("WHTSHPCC").Rows.Add(rowWHTSHPCC)
                        Next

                        Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
                        Dim SHIP_VIA_CODE As String = rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty
                        Dim rowSOTSVIA1 As DataRow = dst.Tables("SOTSVIA1").Rows.Find(SHIP_VIA_CODE)
                        If rowSOTSVIA1 IsNot Nothing Then
                            Dim CARRIER_CODE As String = rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty
                            If CARRIER_CODE.Length > 0 Then
                                For Each rowProvider As UltraWinGrid.UltraGridRow In cmbProvider.Rows
                                    If rowProvider.Cells("CARRIER_CODE").Value = CARRIER_CODE Then
                                        cmbProvider.SelectedRow = rowProvider
                                        cmbProvider_ValueChanged(Nothing, Nothing)

                                        For Each rowShipMethod As UltraWinGrid.UltraGridRow In cmbShipMethod.Rows
                                            If rowShipMethod.Cells("SHIP_VIA_CODE").Value = SHIP_VIA_CODE Then
                                                cmbShipMethod.SelectedRow = rowShipMethod
                                                Exit For
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                        End If
                    End If
            End Select
            txtUserSuppliedValue.Text = userSuppliedValue
        End If

    End Sub

    Private Sub cmbWarehouse_ValueChanged(sender As Object, e As System.EventArgs) Handles cmbWarehouse.ValueChanged
        Dim WHSE_CODE As String = cmbWarehouse.Text

        Dim rowICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(WHSE_CODE)
        If rowICTWHSE1 IsNot Nothing Then
            txtFromCity.Text = rowICTWHSE1.Item("WHSE_CITY") & String.Empty
            txtFromCompany.Text = rowICTWHSE1.Item("WHSE_DESC") & String.Empty
            txtFromCountry.Text = rowICTWHSE1.Item("WHSE_COUNTRY") & String.Empty
            txtFromCountry.Text = txtFromCountry.Text.Trim
            If txtFromCountry.TextLength = 0 Then txtFromCountry.Text = "US"
            txtFromFirstName.Text = rowICTWHSE1.Item("WHSE_CONTACT") & String.Empty
            txtFromPhone.Text = rowICTWHSE1.Item("WHSE_PHONE") & String.Empty
            txtFromState.Text = rowICTWHSE1.Item("WHSE_STATE") & String.Empty
            txtFromStreet.Text = rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty
            txtFromSuite.Text = rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty
            txtFromZip.Text = rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty
            chkFromResidential.Checked = False
        End If
    End Sub

    Private Sub btnFromClear_Click(sender As System.Object, e As System.EventArgs) Handles btnFromClear.Click
        txtFromCity.Clear()
        txtFromCompany.Clear()
        txtFromCountry.Clear()
        txtFromFirstName.Clear()
        txtFromPhone.Clear()
        txtFromState.Clear()
        txtFromStreet.Clear()
        txtFromSuite.Clear()
        txtFromZip.Clear()
        chkFromResidential.Checked = False
    End Sub

    Private Sub btnToClear_Click(sender As System.Object, e As System.EventArgs) Handles btnToClear.Click
        txtToCity.Clear()
        txtToCompany.Clear()
        txtToCountry.Clear()
        txtToFirstName.Clear()
        txtToPhone.Clear()
        txtToState.Clear()
        txtToStreet.Clear()
        txtToSuite.Clear()
        txtToZip.Clear()
        chkToResidential.Checked = False
        chkSignature.Checked = False
        chkToPOBox.Checked = False
    End Sub

    Private Sub btnUseCommodityTotal_Click(sender As System.Object, e As System.EventArgs) Handles btnUseCommodityTotal.Click
        numCustomsValue.Value = Val(dst.Tables("WHTSHPCC").Compute("SUM(EXTENDED_PRICE)", "") & String.Empty)
    End Sub

    Private Sub grdWHTSHPC2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTSHPC2.ClickCellButton
        If e.Cell.Column.Key = "POUNDS" Then
            Dim packageWeight As Decimal = GetScaleWeight()
            e.Cell.Value = Convert.ToInt16(packageWeight)
            e.Cell.Row.Cells("OUNCES").Value = (packageWeight * 16) Mod 16
        End If
    End Sub

#End Region

#Region "Setup Ship Request"

    Private Sub LoadSmartPostDropDowns()

        Dim tblSmartPost As New DataTable
        tblSmartPost.Columns.Add("SMART_POST_TYPE", GetType(System.String))
        tblSmartPost.Columns.Add("SMART_POST_DESC", GetType(System.String))
        tblSmartPost.Rows.Add(New Object() {"0", "Media Mail  (1 to 70 lbs)"})
        tblSmartPost.Rows.Add(New Object() {"1", "Parcel Select (1 to 70 lbs)"})
        tblSmartPost.Rows.Add(New Object() {"2", "Presorted Bound (0.1 to 15 lbs)"})
        tblSmartPost.Rows.Add(New Object() {"3", "Presorted Std (up to 1 lb)"})
        cmbSmartPost.DataSource = tblSmartPost

        Dim tblSmartPostPkg As New DataTable
        tblSmartPostPkg.Columns.Add("SMART_POST_PKG", GetType(System.String))
        tblSmartPostPkg.Columns.Add("SMART_POST_PKG_DESC", GetType(System.String))
        tblSmartPostPkg.Rows.Add(New Object() {"0", "Other"})
        tblSmartPostPkg.Rows.Add(New Object() {"1", "Bag"})
        tblSmartPostPkg.Rows.Add(New Object() {"2", "Barrel"})
        tblSmartPostPkg.Rows.Add(New Object() {"3", "Basket"})
        tblSmartPostPkg.Rows.Add(New Object() {"4", "Box"})
        tblSmartPostPkg.Rows.Add(New Object() {"5", "Bucket"})
        tblSmartPostPkg.Rows.Add(New Object() {"6", "Bundle"})
        tblSmartPostPkg.Rows.Add(New Object() {"7", "Carton"})
        tblSmartPostPkg.Rows.Add(New Object() {"8", "Case"})
        tblSmartPostPkg.Rows.Add(New Object() {"9", "Container"})
        tblSmartPostPkg.Rows.Add(New Object() {"10", "Crate"})
        tblSmartPostPkg.Rows.Add(New Object() {"11", "Cylinder"})
        tblSmartPostPkg.Rows.Add(New Object() {"12", "Drum"})
        tblSmartPostPkg.Rows.Add(New Object() {"13", "Envelope"})
        tblSmartPostPkg.Rows.Add(New Object() {"14", "Hamper"})
        tblSmartPostPkg.Rows.Add(New Object() {"15", "Pail"})
        tblSmartPostPkg.Rows.Add(New Object() {"16", "Pallet"})
        tblSmartPostPkg.Rows.Add(New Object() {"17", "Piece"})
        tblSmartPostPkg.Rows.Add(New Object() {"18", "Reel"})
        tblSmartPostPkg.Rows.Add(New Object() {"19", "Roll"})
        tblSmartPostPkg.Rows.Add(New Object() {"20", "Skid"})
        tblSmartPostPkg.Rows.Add(New Object() {"21", "Tank"})
        tblSmartPostPkg.Rows.Add(New Object() {"22", "Tube"})
        cmbSmartPostPackage.DataSource = tblSmartPostPkg
    End Sub

    Private Function RequestShippingLabel() As Boolean

        Try
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Requesting Shipping Label(s)")

            RequestShippingLabel = False

            If EntryMode = "S" Then
                SHIP_CNTL_NO = ASCMAIN1.Next_Control_No("WHTSHPC1.SHIP_CNTL_NO")
                For Each tableName As String In New String() {"WHTSHPC1", "WHTSHPC2", "WHTSHPC5_SF", "WHTSHPC5_ST", "WHTSHPC5_HL", "WHTSHPCC", "WHTSHPCP_S", "WHTSHPCP_D"}
                    For Each row As DataRow In dst.Tables(tableName).Select("", "", DataViewRowState.CurrentRows)
                        row.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    Next
                Next
            End If

            Dim rowWHTSHPC1 As DataRow = dst.Tables("WHTSHPC1").Rows(0)
            rowWHTSHPC1.Item("CARRIER_PROD_CODE") = cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value
            rowWHTSHPC1.Item("SHIP_VIA_CODE") = cmbShipMethod.SelectedRow.Cells("SHIP_VIA_CODE").Value

            ' Preload any Third party account info 
            txtShipAccountNo.Text = txtShipAccountNo.Text.Trim
            Dim rowSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Rows.Find(rowWHTSHPC1.Item("CARRIER_CODE") & String.Empty)
            If txtShipAccountNo.TextLength = 0 AndAlso rowSOTCARR1.Item("SHIP_ACCT_NO") & String.Empty <> String.Empty Then
                cmbShipPayor.Value = "T" ' Set to Third Party
                txtShipAccountNo.Text = rowSOTCARR1.Item("SHIP_ACCT_NO") & String.Empty
                txtShipAccountZip.Text = rowSOTCARR1.Item("SHIP_3PY_ZIPCODE") & String.Empty
                txtShipCountry.Text = rowSOTCARR1.Item("SHIP_3PY_COUNTRY") & String.Empty
            End If

            If txtDutiesAccountNo.TextLength = 0 AndAlso rowSOTCARR1.Item("SHIP_ACCT_NO") & String.Empty <> String.Empty Then
                cmbDutiesPayor.Value = "T" ' Set to Third Party
                txtDutiesAccountNo.Text = rowSOTCARR1.Item("SHIP_ACCT_NO") & String.Empty
                txtDutiesAccountZip.Text = rowSOTCARR1.Item("SHIP_3PY_ZIPCODE") & String.Empty
                txtDutiesCountry.Text = rowSOTCARR1.Item("SHIP_3PY_COUNTRY") & String.Empty
            End If

            GetSenderInfo()
            GetRecipientInfo()
            BuildPackages()
            GetCredentials()
            GetServiceType()
            GetDropoffType()
            GetShipPayor()
            GetDutiesPayor()
            GetSpecialServices()
            GetHALDetails()
            GetCommodities()
            GetSmartPost()


            If optPrint_Type.Value <> rowSOTCARR1.Item("LABEL_FORMAT") & String.Empty Then
                Dim labelFormatDesc As String = "Unknown"
                For Each vlItem As ValueListItem In optPrint_Type.Items
                    If vlItem.DataValue = rowSOTCARR1.Item("LABEL_FORMAT") & String.Empty Then
                        labelFormatDesc = vlItem.DisplayText
                    End If
                Next
                If ASCMAIN1.CLIENT <> "VAN" Then
                    If MessageBox.Show("Typically you use " & labelFormatDesc & " to print " & rowSOTCARR1.Item("CARRIER_DESC") & " Labels. Do you want to change to " & labelFormatDesc & " labels?", "Labels", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                        optPrint_Type.Value = rowSOTCARR1.Item("LABEL_FORMAT") & String.Empty
                    End If
                End If
            End If

            With clsShip
                '.EzshipLabelImage = nsoftware.InShip.EzshipLabelImageTypes.itEltron
                Select Case optPrint_Type.Value
                    Case "E"
                        .EzshipLabelImage = nsoftware.InShip.EzshipLabelImageTypes.itEltron
                    Case "Z"
                        .EzshipLabelImage = nsoftware.InShip.EzshipLabelImageTypes.itZPL
                    Case "X"
                        .EzshipLabelImage = nsoftware.InShip.EzshipLabelImageTypes.itZebra
                End Select

                .ShippingLabelDirectory = ShippingLabelDirectory
                .ShippingLabelPrefix = SHIP_CNTL_NO
                .ShipDate = dteShipDate.DateTime
                clsShip.TotalCustomsValue = numCustomsValue.Value

                ' Sample valid CustomContent in case needed in the future
                '<CustomContent>
                '<TextEntries>
                '<Position><X>150</X><Y>70</Y></Position>
                '<Format>George is a worm</Format>
                '<ThermalFontId>18</ThermalFontId>
                '</TextEntries>
                '<BarcodeEntries>
                '<Position><X>150</X><Y>150</Y></Position>
                '<Format>123456789102</Format> 
                '<BarHeight>150</BarHeight> 
                '<ThinBarWidth>4</ThinBarWidth> 
                '<BarcodeSymbology>CODE128B</BarcodeSymbology> 
                '</BarcodeEntries>
                '</CustomContent>

                txtFedexContent.Text = txtFedexContent.Text.Trim
                If rowWHTSHPC1.Item("CARRIER_CODE") & String.Empty = "FEDEX" AndAlso txtFedexContent.TextLength > 0 Then
                    .FedexCustomContent = txtFedexContent.Text
                End If

            End With

            clsShip.PackageDetailList = shipPackageDetailList

            Select Case EntryMode
                Case "S"
                    ' Request label
                    If clsShip.RequestLabel() Then

                        For Each shipPackageDetail As nsoftware.InShip.PackageDetail In shipPackageDetailList
                            Dim SHIP_PACKAGE_NO As String = Val(shipPackageDetail.Id)
                            If dst.Tables("WHTSHPC2").Select("SHIP_PACKAGE_NO = " & SHIP_PACKAGE_NO, "").Length > 0 Then
                                Dim rowWHTSHPC2 As DataRow = dst.Tables("WHTSHPC2").Select("SHIP_PACKAGE_NO = " & SHIP_PACKAGE_NO)(0)
                                rowWHTSHPC2.Item("TRACKING_NO") = shipPackageDetail.TrackingNumber & String.Empty
                                rowWHTSHPC2.Item("BASE_CHARGE") = Val(clsShip.ShipmentBaseCharge(SHIP_PACKAGE_NO) & String.Empty)
                                rowWHTSHPC2.Item("NET_CHARGE") = Val(clsShip.ShipmentNetCharge(SHIP_PACKAGE_NO) & String.Empty)
                                rowWHTSHPC2.Item("TOTAL_DISCOUNT") = Val(clsShip.ShipmentDiscountCharge(SHIP_PACKAGE_NO) & String.Empty)
                                rowWHTSHPC2.Item("TOTAL_SURCHARGES") = Val(clsShip.ShipmentSurCharge(SHIP_PACKAGE_NO) & String.Empty)
                                rowWHTSHPC2.Item("LIST_PRICE") = Val(clsShip.ShipmentListCharge(SHIP_PACKAGE_NO) & String.Empty)
                            End If

                            rowWHTSHPC1.Item("ERROR_MSG") = clsShip.LastError & String.Empty
                            If rowWHTSHPC1 IsNot Nothing AndAlso (rowWHTSHPC1.Item("ERROR_MSG") & String.Empty).ToString.Length > 200 Then
                                rowWHTSHPC1.Item("ERROR_MSG") = rowWHTSHPC1("ERROR_MSG").ToString.Substring(0, 200).Trim
                            End If
                            rowWHTSHPC1.Item("MASTER_TRACKING_NO") = clsShip.MasterTrackingNumber & String.Empty

                            If rowWHTSHPC1.Item("MASTER_TRACKING_NO") & String.Empty = String.Empty Then
                                rowWHTSHPC1.Item("MASTER_TRACKING_NO") = shipPackageDetailList(0).TrackingNumber
                            End If

                            RequestShippingLabel = True
                            Update_Record()

                            If shipPackageDetail.ShippingLabel.Length > 0 Then PrintShipingLabel(shipPackageDetail.ShippingLabel)
                            If shipPackageDetail.CODLabel.Length > 0 Then PrintShipingLabel(shipPackageDetail.CODLabel)
                            If shipPackageDetail.ReturnReceipt.Length > 0 Then PrintShipingLabel(shipPackageDetail.ReturnReceipt)

                        Next
                    Else
                        MessageBox.Show("Shipping Label(s) could not be captured. " & clsShip.LastError)
                    End If

                Case "R"
                    clsShip.RequestShipmentRates()

                    Dim msg As String = String.Empty
                    For Each reqSer As ServiceDetail In clsShip.RequestedServicesRates
                        msg &= vbCrLf & reqSer.ServiceType.ToString.PadRight(30, "_") _
                            & (reqSer.ServiceType & " ").ToString.PadRight(5) _
                            & Val(reqSer.AccountNetCharge & String.Empty).ToString("#,##0.00").ToString.PadLeft(8) & reqSer.TransitTime.PadLeft(12)
                    Next

                    MessageBox.Show("Available Delivery Options" & vbCrLf & msg, "Requested Rates", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End Select

        Catch ex As Exception
            MessageBox.Show("Error requesting shipping label: " & ex.Message, "Request Label Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If rowWHTSHPC1 IsNot Nothing Then rowWHTSHPC1.Item("ERROR_MSG") = ex.Message
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try
    End Function

    Private Sub ReprintShippingLabel()

        Try
            Dim CARRIER_CODE As String = cmbProvider.SelectedRow.Cells("CARRIER_CODE").Value & String.Empty
            Dim rowSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)
            Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)
            ShippingLabelDirectory = (rowSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim

            Dim labelFilesFound As Int16 = 0

            If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "edz" Then
                Stop
                Select Case ASCMAIN1.CLIENT
                    Case "RGI"
                        ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "R:\")
                    Case "NYA"
                        ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "N:\")
                    Case "VAN"
                        ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "V:\")
                End Select
            End If

            For Each Label As String In My.Computer.FileSystem.GetFiles(ShippingLabelDirectory, FileIO.SearchOption.SearchTopLevelOnly, SHIP_CNTL_NO & "*.*")
                Dim ShippingLabel As String = String.Empty

                Using sr As New IO.StreamReader(Label)
                    ShippingLabel = sr.ReadToEnd
                    sr.Close()
                End Using

                PrintShipingLabel(ShippingLabel)
                labelFilesFound += 1
            Next

            MessageBox.Show("There were (" & labelFilesFound & ") labels sent to the printer.", "Reprint Label", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub GetSmartPost()
        If cmbProvider.SelectedRow.Cells("PROVIDER_TYPE").Value = WHCSHIP1.ProviderTypeFedex _
        AndAlso Val(cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value & String.Empty) = FedexSmartPost Then
            clsShip.FedexSmartPost.Indicia = cmbSmartPost.SelectedRow.Cells("SMART_POST_TYPE").Value
            clsShip.FedexSmartPost.PhysicalPackaging = cmbSmartPostPackage.SelectedRow.Cells("SMART_POST_PKG").Value
            clsShip.FedexSmartPost.HubId = cmbProvider.SelectedRow.Cells("FEDEX_HUB_ID").Value
        Else
            dst.Tables("WHTSHPC1").Rows(0).Item("SMART_POST_TYPE") = DBNull.Value
            dst.Tables("WHTSHPC1").Rows(0).Item("SMART_POST_PKG") = DBNull.Value
            dst.Tables("WHTSHPC1").Rows(0).Item("SMART_POST_HUB_ID") = DBNull.Value
        End If
    End Sub

    Private Sub GetSenderInfo()

        With clsShip.Sender
            .FirstName = txtFromFirstName.Text.Trim
            .MiddleInitial = ""
            .LastName = "" 'txtFromLastName.Text.Trim
            .Address1 = txtFromStreet.Text.Trim
            .Address2 = txtFromSuite.Text.Trim
            .City = txtFromCity.Text.Trim
            .State = txtFromState.Text.Trim.ToUpper
            .ZipCode = txtFromZip.Text.Trim
            .CountryCode = txtFromCountry.Text.Trim
            .Phone = txtFromPhone.Text

            .Company = ROWs("ASTPARM1").Item("AS_PARM_INST_NAME") & String.Empty
            If .Company.Length = 0 Then
                .Company = txtFromCompany.Text.Trim
            End If

            If .CountryCode = String.Empty OrElse .CountryCode.ToUpper = "USA" Then
                .CountryCode = "US"
            End If

            .IsResidental = chkFromResidential.Checked
            .IsPOBox = False
        End With
    End Sub

    Private Sub GetRecipientInfo()

        With clsShip.Recipient
            .FirstName = txtToFirstName.Text.Trim
            .MiddleInitial = ""
            .LastName = "" 'txtToLastName.Text.Trim
            .Address1 = txtToStreet.Text.Trim
            .Address2 = txtToSuite.Text.Trim
            .City = txtToCity.Text.Trim
            .State = txtToState.Text.Trim
            .ZipCode = txtToZip.Text.Trim
            .CountryCode = txtToCountry.Text.Trim
            .Company = txtToCompany.Text.Trim
            .Phone = txtToPhone.Text

            If .CountryCode = String.Empty OrElse .CountryCode.ToUpper = "USA" Then
                .CountryCode = "US"
            End If

            .IsResidental = chkToResidential.Checked
            .IsPOBox = chkToPOBox.Checked

            If .Company.Length = 0 Then
                .Company = (.FirstName & " " & .LastName).ToString.Trim
            End If

            clsShip.SignatureRequired = chkSignature.Checked
        End With
    End Sub

    Private Sub BuildPackages()

        shipPackageDetailList.Clear()
        Dim idCtr As Int16 = 1
        For Each row As DataRow In dst.Tables("WHTSHPC2").Rows
            Dim shipPackageDetail As New nsoftware.InShip.PackageDetail
            With shipPackageDetail
                .PackagingType = nsoftware.InShip.TPackagingTypes.ptYourPackaging
                .Weight = Convert.ToInt32(row.Item("WEIGHT"))
                .Length = Convert.ToInt32(row.Item("LENGTH"))
                .Width = Convert.ToInt32(row.Item("WIDTH"))
                .Height = Convert.ToInt32(row.Item("HEIGHT"))

                ' Onlt three references per package
                Dim reference As String = String.Empty
                Dim value As String = String.Empty
                Dim iCtr As Int16 = 1
                For Each field As String In New String() {"CR:CUST_REF", "IN:INV_NO", "PO:PO_ORDER_NO", "DN:DEPT_NO"} ' , "BL:INV_BOL_NO"
                    Dim code As String = field.Split(":")(0)
                    Dim fieldName As String = field.Split(":")(1)

                    value = (row.Item(fieldName) & String.Empty).ToString.Trim
                    If value.Length > 0 Then
                        reference &= "; " & code & ":" & value
                        iCtr += 1
                    End If
                    If iCtr > 3 Then Exit For
                Next

                If reference.StartsWith(";") Then
                    reference = reference.Substring(1).Trim
                End If

                .Reference = reference
                .Id = idCtr.ToString.Trim.PadLeft(8, "0")

                If Val(row.Item("INSURED_VALUE") & String.Empty) > 0 Then
                    .InsuredValue = Val(row.Item("INSURED_VALUE") & String.Empty)
                End If

                'COD_TYPE, COD_AMOUNT
                If Val(row.Item("COD_AMOUNT") & String.Empty) > 0 Then
                    .CODAmount = Val(row.Item("COD_AMOUNT") & String.Empty)
                    .CODType = Val(row.Item("COD_TYPE") & String.Empty)
                End If

                idCtr += 1
            End With

            shipPackageDetailList.Add(shipPackageDetail)
        Next

    End Sub

    Private Sub GetCredentials()

        Dim CARRIER_CODE As String = cmbProvider.SelectedRow.Cells("CARRIER_CODE").Value & String.Empty
        Dim CARRIER_ACCOUNT_NO As String = cmbProvider.SelectedRow.Cells("CARRIER_ACCOUNT_NO").Value & String.Empty

        'If dst.Tables("WHTSHPC1").Rows.Count > 0 Then
        '    CARRIER_ACCOUNT_NO = dst.Tables("WHTSHPC1").Rows(0).Item("CARRIER_ACCOUNT_NO") & String.Empty
        'End If

        'If CARRIER_ACCOUNT_NO.Length = 0 Then
        '    CARRIER_ACCOUNT_NO = cmbProvider.SelectedRow.Cells("CARRIER_ACCOUNT_NO").Value & String.Empty
        'End If

        Dim rowSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)
        Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_ID IS NOT NULL AND SHIPPER_PASSWORD IS NOT NULL" & IIf(CARRIER_ACCOUNT_NO.Length > 0, " AND CARRIER_ACCOUNT_NO = '" & CARRIER_ACCOUNT_NO & "'", ""))(0)
        ShippingLabelDirectory = (rowSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim

        Try
            If ASCMAIN1.Running_in_VS Then
                ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "C:\").Replace("R:\", "C:\")
            End If
            If ShippingLabelDirectory.Length > 0 Then
                If Not My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                    My.Computer.FileSystem.CreateDirectory(ShippingLabelDirectory)
                End If
            End If
        Catch ex As Exception
            ShippingLabelDirectory = String.Empty
        End Try

        If ShippingLabelDirectory.Length > 0 AndAlso Not ShippingLabelDirectory.EndsWith("\") Then
            ShippingLabelDirectory = ShippingLabelDirectory & "\"
        End If

        ' Credentials
        clsShip.Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
        clsShip.UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
        clsShip.Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
        clsShip.AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
        clsShip.UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
        clsShip.FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
        clsShip.FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty

        rowWHTSHPC1.Item("CARRIER_ACCOUNT_NO") = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
        rowWHTSHPC1.Item("CARRIER_CODE") = CARRIER_CODE
        clsShip.LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

    End Sub

    Private Sub GetServiceType()
        'Service Type
        If isInternationalShipment Then
            clsShip.RequestedServiceType = Val(cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE_INTL").Value & String.Empty)
        Else
            clsShip.RequestedServiceType = Val(cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value & String.Empty)
        End If
    End Sub

    Private Sub GetDropoffType()
        clsShip.DropOffType = FedexshipintlDropoffTypes.dtRegularPickup
    End Sub

    Private Sub GetShipPayor()
        Select Case cmbShipPayor.Value
            Case "S" : clsShip.Payor = TPayorTypes.ptSender
            Case "R" : clsShip.Payor = TPayorTypes.ptRecipient
            Case "T" : clsShip.Payor = TPayorTypes.ptThirdParty
            Case "C" : clsShip.Payor = TPayorTypes.ptCollect
        End Select
        clsShip.PayorContact.AccountNumber = txtShipAccountNo.Text
        clsShip.PayorContact.CountryCode = txtShipCountry.Text
        clsShip.PayorContact.ZipCode = txtShipAccountZip.Text
    End Sub

    Private Sub GetDutiesPayor()
        Select Case cmbShipPayor.Value
            Case "S" : clsShip.DutiesPayor = TPayorTypes.ptSender
            Case "R" : clsShip.DutiesPayor = TPayorTypes.ptRecipient
            Case "T" : clsShip.DutiesPayor = TPayorTypes.ptThirdParty
            Case "C" : clsShip.DutiesPayor = TPayorTypes.ptCollect
        End Select
        clsShip.DutiesPayorContact.AccountNumber = txtDutiesAccountNo.Text
        clsShip.DutiesPayorContact.CountryCode = txtDutiesCountry.Text
        clsShip.DutiesPayorContact.ZipCode = txtDutiesAccountZip.Text
    End Sub

    Private Sub GetSpecialServices()
        For Each row As DataRow In dst.Tables("SHTSHIPS").Select("SEL = '1'")
            clsShip.ShipmentSpecialServices = clsShip.ShipmentSpecialServices Or Val("&H" & row.Item("SPCL_SHIP_CODE") & "L")
        Next
    End Sub

    Private Sub GetCommodities()
        ' Only used for International
        clsShip.CommodityDetailList.Clear()
        For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWHTSHPCC.Rows
            Dim CommodityDetail As New nsoftware.InShip.CommodityDetail
            CommodityDetail.Description = grdRow.Cells("COMMODITY_DESC").Value & String.Empty
            CommodityDetail.NumberOfPieces = Val(grdRow.Cells("NUM_PIECES").Value & String.Empty)
            CommodityDetail.Quantity = Val(grdRow.Cells("QUANTITY").Value & String.Empty)
            CommodityDetail.QuantityUnit = grdRow.Cells("QUANTITY_UOM").Value & String.Empty
            CommodityDetail.UnitPrice = grdRow.Cells("UNIT_PRICE").Value & String.Empty
            CommodityDetail.Weight = Val(grdRow.Cells("WEIGHT").Value & String.Empty) ' Leave as pounds
            CommodityDetail.Manufacturer = grdRow.Cells("MANUFACTURER").Value & String.Empty
            clsShip.CommodityDetailList.Add(CommodityDetail)
        Next
    End Sub

    Private Sub GetHALDetails()

        With clsShip.HoldAtLocation
            .AccountNumber = ""
            .Address1 = txtHoldAddress1.Text.Trim
            .Address2 = txtHoldAddress2.Text.Trim
            .City = txtHoldCity.Text.Trim
            .Company = txtHoldCompany.Text.Trim
            .CountryCode = txtHoldCountry.Text.Trim
            .eMail = ""
            .FirstName = txtHoldContact.Text.Trim
            .Fax = ""
            .IsPOBox = False
            .IsResidental = False
            .LastName = ""
            .MiddleInitial = ""
            .Phone = txtHoldPhone.Text.Trim
            .State = txtHoldState.Text.Trim
            .ZipCode = txtHoldZip.Text.Trim
        End With
    End Sub

    Private Sub ValidatePacking(ByRef grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow)

        Dim SHIP_PACKAGE_NO As String = (grdRow.Cells("SHIP_PACKAGE_NO").Value & String.Empty).ToString.Trim
        If SHIP_PACKAGE_NO = String.Empty Then
            SHIP_PACKAGE_NO = Val(dst.Tables("WHTSHPC2").Compute("MAX(SHIP_PACKAGE_NO)", "") & String.Empty) + 1
            grdRow.Cells("SHIP_PACKAGE_NO").Value = SHIP_PACKAGE_NO
        End If

        grdRow.Cells("SHIP_CNTL_NO").Value = SHIP_CNTL_NO

        Dim LENGTH As Int16 = Val(grdRow.Cells("LENGTH").Value & String.Empty)
        Dim WIDTH As Int16 = Val(grdRow.Cells("WIDTH").Value & String.Empty)
        Dim HEIGHT As Int16 = Val(grdRow.Cells("HEIGHT").Value & String.Empty)

        If (grdRow.Cells("PACKAGING_TYPE").Value & String.Empty).ToString.Trim.Length = 0 Then
            EMsg &= vbCrLf & "Package type is required"
        ElseIf Val(grdRow.Cells("PACKAGING_TYPE").Value & String.Empty) <> nsoftware.InShip.TPackagingTypes.ptYourPackaging Then
            If LENGTH <= 0 OrElse WIDTH <= 0 OrElse HEIGHT <= 0 Then
                EMsg &= vbCrLf & "Package Length, Width and Height are required and must be greater than 0"
            End If
        End If

        Dim POUNDS As Int16 = Val(grdRow.Cells("POUNDS").Value & String.Empty)
        Dim OUNCES As Int16 = Val(grdRow.Cells("OUNCES").Value & String.Empty)

        If POUNDS <= 0 AndAlso OUNCES <= 0 Then
            EMsg &= vbCrLf & "Pounds and/or Ounces must be greater than 0"
        ElseIf POUNDS < 0 Then
            EMsg &= vbCrLf & "Pounds must be greater equal than 0"
        ElseIf OUNCES < 0 Then
            EMsg &= vbCrLf & "Ounces must be greater equal than 0"
        Else
            grdRow.Cells("WEIGHT").Value = POUNDS * 16 + OUNCES
        End If

        Dim COD_AMOUNT As Decimal = Val(grdRow.Cells("COD_AMOUNT").Value & String.Empty)
        Dim COD_TYPE As String = grdRow.Cells("COD_TYPE").Value & String.Empty
        If COD_AMOUNT < 0 Then
            EMsg &= vbCrLf & "COD Amount must be greater equal than 0"
        ElseIf COD_AMOUNT > 0 Then
            If COD_TYPE <> "0" AndAlso COD_TYPE <> "1" Then
                EMsg &= vbCrLf & "COD Type is required when providing a COD Amount"
            End If
        ElseIf COD_TYPE = "0" AndAlso COD_TYPE = "1" Then
            If COD_AMOUNT = 0 Then
                EMsg &= vbCrLf & "COD Amount is required when providing a COD Type"
            End If
        End If

        If COD_AMOUNT = 0 And COD_TYPE = String.Empty Then
            grdRow.Cells("COD_TYPE").Value = "2" ' None
        End If
    End Sub

    Private Sub ValidateCommodity(ByRef grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow)

        Dim COMMODITY_LNO As String = (grdRow.Cells("COMMODITY_LNO").Value & String.Empty).ToString.Trim
        If COMMODITY_LNO = String.Empty Then
            COMMODITY_LNO = Val(dst.Tables("WHTSHPCC").Compute("MAX(COMMODITY_LNO)", "") & String.Empty) + 1
            grdRow.Cells("COMMODITY_LNO").Value = COMMODITY_LNO
        End If

        grdRow.Cells("SHIP_CNTL_NO").Value = SHIP_CNTL_NO

        Dim COMMODITY_DESC As String = (grdRow.Cells("COMMODITY_DESC").Value & String.Empty).ToString.Trim
        Dim MANUFACTURER As String = (grdRow.Cells("MANUFACTURER").Value & String.Empty).ToString.Trim
        Dim NUM_PIECES As Int16 = Val(grdRow.Cells("NUM_PIECES").Value & String.Empty)
        Dim UOM As String = Val(grdRow.Cells("QUANTITY_UOM").Value & String.Empty)
        Dim UNIT_PRICE As Decimal = Val(grdRow.Cells("UNIT_PRICE").Value & String.Empty)
        Dim QUANTITY As Decimal = Val(grdRow.Cells("QUANTITY").Value & String.Empty)
        Dim WEIGHT As Decimal = Val(grdRow.Cells("WEIGHT").Value & String.Empty)

        ' Required Fields
        If COMMODITY_DESC.Length = 0 Then
            EMsg &= vbCrLf & "Description is required"
        End If

        If MANUFACTURER.Length = 0 Then
            EMsg &= vbCrLf & "Manufacturer is required"
        End If

        If QUANTITY < 0 Then
            EMsg &= vbCrLf & "Quantity must be greater equal 0"
        End If

        If UOM.Length = 0 Then
            EMsg &= vbCrLf & "Unit of Measure is required"
        End If

        If UNIT_PRICE < 0 Then
            EMsg &= vbCrLf & "Unit Price must be greater equal 0"
        End If

        If WEIGHT < 0 Then
            EMsg &= vbCrLf & "Weight must be greater 0"
        End If

    End Sub

    Private Function GetCommodityUnits() As String()
        Return New String() {":", _
              "AR:Carat" _
             , "CG:Centigram" _
             , "CM:Centimeters" _
             , "CM3:Cubic centimeters" _
             , "CFT:Cubic feet" _
             , "M3:Cubic meters" _
             , "DOZ:Dozen" _
             , "DPR:Dozen pair" _
             , "EA:Each" _
             , "GAL:Gallon" _
             , "G:Grams" _
             , "GR:Gross" _
             , "KG:Kilograms" _
             , "LFT:Linear foot" _
             , "LNM:Linear meters" _
             , "LYD:Linear yard" _
             , "LTR:Liters" _
             , "M:Meters" _
             , "MG:Milligram" _
             , "ML:Milliliter" _
             , "NO:Number" _
             , "OZ:Ounces" _
             , "PRS:Pairs" _
             , "PCS:Pieces" _
             , "LB:Pound" _
             , "CM2:Square centimeters" _
             , "SFT:Square feet" _
             , "SQI:Square inches" _
             , "M2:Square meters" _
             , "SYD:Square yards" _
             , "YD:Yard"}

    End Function


    Private Function CancelShipment() As Boolean
        Try
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Requesting Shipment Cancellation")

            GetCredentials()
            GetServiceType()

            Dim rowWHTSHPC1 As DataRow = dst.Tables("WHTSHPC1").Rows(0)
            'Dim multiShipment As Boolean = dst.Tables("WHTSHPC2").Select("ISNULL(TRACKING_NO, '') <> ''").Length > 1
            Dim rowSOTCARR2 As DataRow = ASCDATA1.GetDataRow("select * from sotcarr2 where CARRIER_CODE = :PARM1" _
                                                             & " and CARRIER_PROD_CODE = :PARM2", "VV", _
                                                             New Object() {rowWHTSHPC1.Item("CARRIER_CODE"), rowWHTSHPC1.Item("CARRIER_PROD_CODE")})

            ' cancel one package at a time
            Dim numCancelled As Int16 = 0
            Dim numNotCancelled As Int16 = 0
            With clsShip
                For Each rowWHTSHPC2 As DataRow In dst.Tables("WHTSHPC2").Select("SEL = '1'")
                    Dim TRACKING_NO As String = rowWHTSHPC2.Item("TRACKING_NO") & String.Empty
                    Dim TRACKING_ID_TYPE As Int32 = Val(rowSOTCARR2.Item("TRACKING_ID_TYPE") & String.Empty)

                    ' USPS Pitney Bowes
                    Dim rowSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Rows.Find(rowWHTSHPC1.Item("CARRIER_CODE") & String.Empty) ' LookUp("SOTCARR1", rowWHTSHPC1.Item("CARRIER_CODE") & String.Empty)
                    If rowSOTCARR1 IsNot Nothing Then
                        If rowSOTCARR1.Item("CARRIER_CODE") & String.Empty = "USPS" Then
                            If rowSOTCARR1.Item("USPS_PARTNER") & String.Empty = "3" Then
                                TRACKING_NO = rowWHTSHPC2.Item("TRACKING_NUMBER") & String.Empty
                                TRACKING_ID_TYPE = Val(ASCMAIN1.Next_Control_No("SOTCARR1.USPS"))
                            End If
                        End If
                    End If

                    If .CancelShipment(TRACKING_NO, False, TRACKING_ID_TYPE) Then
                        Try
                            BeginTrans()
                            With rowWHTSHPC2
                                .Item("STATUS") = "C"
                                .Item("CANCEL_OPER") = ASCMAIN1.USER_ID
                                .Item("CANCEL_DATE") = DateTime.Now
                                Update_Record_TDA("WHTSHPC2")
                            End With

                            ASCMAIN1.sql = "Update WHTSHPC1 SET STATUS = (SELECT MAX(STATUS) FROM WHTSHPC2 WHERE WHTSHPC1.SHIP_CNTL_NO = WHTSHPC2.SHIP_CNTL_NO)"
                            ASCMAIN1.sql &= " where WHTSHPC1.SHIP_CNTL_NO = '" & rowWHTSHPC1.Item("SHIP_CNTL_NO") & "'"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                            CommitTrans()
                        Catch ex As Exception
                            Rollback()
                        End Try
                        numCancelled += 1
                    Else
                        numNotCancelled += 1
                        If .LastError.Length > 0 Then
                            MessageBox.Show("Error returned by shipper: " & .LastError, "Cancel Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    End If
                Next
            End With

            Dim zMsg As String = String.Empty
            zMsg &= "There were " & numCancelled & " packages cancelled"
            zMsg &= Environment.NewLine & "There were " & numNotCancelled & " packages NOT cancelled"
            MessageBox.Show(zMsg, "Cancel Package", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Cancel Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False

        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
            Me.Cursor = Cursors.Default
        End Try
    End Function

    Private Sub SetReadOnly(ByVal readOnlyValue As Boolean)
        Set_Read_Only(grpHeader, readOnlyValue)
        Set_Read_Only(grpShipFrom, readOnlyValue)
        Set_Read_Only(grpShipTo, readOnlyValue)
        Set_Read_Only(grpShipPayor, readOnlyValue)
        Set_Read_Only(grpDutiesPayor, readOnlyValue)
        Set_Read_Only(grpHoldAtLocation, readOnlyValue)
        Set_Read_Only(grpCommodity, readOnlyValue)

        'Set_Read_Only(grdWHTSHPC2, readOnlyValue)
        'Set_Read_Only(grdSHTSHIPS, readOnlyValue)
        'Set_Read_Only(grdWHTSHPCC, readOnlyValue)

        If readOnlyValue Then
            grdWHTSHPC2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grdWHTSHPC2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdWHTSHPC2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

            grdSHTSHPCS.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

            grdWHTSHPCC.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grdWHTSHPCC.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdWHTSHPCC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        Else
            grdWHTSHPC2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            grdWHTSHPC2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdWHTSHPC2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

            grdSHTSHPCS.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

            grdWHTSHPCC.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            grdWHTSHPCC.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdWHTSHPCC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        End If

        btnFromClear.Visible = Not readOnlyValue
        btnToClear.Visible = Not readOnlyValue
    End Sub

    Public Function PrintShipingLabel(ByVal LabelData As String) As Boolean

        Try
            If (ASCMAIN1.USER_ID = "edz" OrElse ASCMAIN1.USER_ID = "wjz") AndAlso ASCMAIN1.Running_in_VS Then
                ' Find Zebra printer
                Dim zebraPrinter As String = FindZebraPrinter()

                Dim vLabelPrinter As New ASCPRINT
                Return vLabelPrinter.SendStringToPrinter(zebraPrinter, LabelData)
            End If

            'cboZebraPrinter.Text
            If ASCMAIN1.CLIENT = "VAN" Then
                If txtlabelPrinter.BackColor = Drawing.Color.Green Then
                    ASCMAIN1.LabelPrinterSerialPort.WriteLine(LabelData)
                Else
                    Dim vLabelPrinter As New ASCPRINT
                    Return vLabelPrinter.SendStringToPrinter(cboZebraPrinter.Text, LabelData)
                End If

            Else
                ASCMAIN1.LabelPrinterSerialPort.WriteLine(LabelData)
            End If

        Catch ex As Exception
            MessageBox.Show("Print Shipping Label Error: " & ex.Message)
        End Try

    End Function

    Private Shared Function FindZebraPrinter() As String
        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZEBRA450") Then
                Return printerName
            End If
        Next printerName

        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZEBRA") Then
                Return printerName
            End If
        Next printerName

        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZP450") Then
                Return printerName
            End If
        Next printerName

        Return String.Empty
    End Function

#End Region

#Region "Serial and Com Connections"

    ''' <summary>
    ''' Sets the Printer Settings
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetUpPortsAndPrinters()

        Dim tooltip As New System.Windows.Forms.ToolTip()

        ' Label Printer Port
        Try
            txtlabelPrinter.BackColor = Drawing.Color.Red

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing Then
                txtlabelPrinter.Text = ASCMAIN1.LabelPrinterSerialPort.PortName
                tooltip.SetToolTip(txtlabelPrinter, txtlabelPrinter.Text)
            Else
                Me.txtlabelPrinter.Text = "No Port"
                tooltip.SetToolTip(txtlabelPrinter, txtlabelPrinter.Text)
            End If

            txtlabelPrinter.BackColor = Drawing.Color.Yellow
            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing AndAlso Not ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                ASCMAIN1.LabelPrinterSerialPort.Open()
            End If

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing AndAlso ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                txtlabelPrinter.BackColor = Drawing.Color.Green
                printerFound = True
            End If

            If ASCMAIN1.CLIENT = "VAN" Then
                'IP Printers @ VAN
                printerFound = True
            End If

        Catch ex As Exception
            txtlabelPrinter.BackColor = Drawing.Color.Red
            tooltip.SetToolTip(txtlabelPrinter, txtlabelPrinter.Text)
        End Try

        ' Scale Port
        Try
            tooltip.SetToolTip(txtScale, ASCMAIN1.ScalePort.PortName)
            txtScale.Text = ASCMAIN1.ScalePort.PortName
            txtScale.Appearance.BackColor = Drawing.Color.Green
        Catch ex As Exception
            txtScale.Text = String.Empty
            txtScale.Appearance.BackColor = Drawing.Color.Red
            tooltip.SetToolTip(txtScale, ex.Message)
        End Try

    End Sub

#End Region

#Region "Carrier Rates"

    Private Sub RequestRates(ByVal SHIP_CNTL_NO As String, ByVal autoRateShop As Boolean)

        Try
            Me.Cursor = Cursors.WaitCursor
            Dim rowWHTSHPC4 As DataRow = Nothing
            Dim rowWHTSHPCA As DataRow = Nothing

            Dim CARRIER_SURCHARGE_PERC As Int16 = 0
            Dim FRT_PER_SALES_HOLD As Int16 = 0
            Dim CARRIER_PPA_TYPE As String = "L"
            Dim CARRIER_SURCHARGE_BASE As String = "L"

            If Not IsDate(dteShipDate.Text) Then
                If Not autoRateShop Then
                    MessageBox.Show("You must provide the Ship Date.", "Carrier Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Exit Sub
            End If

            For Each row As DataRow In dst.Tables("WHTSHPC2").Select("", "", DataViewRowState.CurrentRows)
                If Val(row.Item("POUNDS") & String.Empty) + Val(row.Item("OUNCES") & String.Empty) <= 0 Then
                    If Not autoRateShop Then MessageBox.Show("All Cartons must have a weight and dimensions.", "Request Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If Val(row.Item("HEIGHT") & String.Empty) <= 0 OrElse Val(row.Item("LENGTH") & String.Empty) <= 0 OrElse Val(row.Item("WIDTH") & String.Empty) <= 0 Then
                    If Not autoRateShop Then MessageBox.Show("All Cartons must have a weight and dimensions.", "Request Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            Next

            ASCMAIN1.Progress("Request Carrier Rates")
            Dim rList() As WHCSHIP1.RateList

            dst.Tables("WHTSHPC4").Rows.Clear()

            Dim rUPSList(1) As WHCSHIP1.RateList
            Dim rFEDEXList(1) As WHCSHIP1.RateList

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("-", "UPS")
            rUPSList = GetUpsRates()

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("-", "FedEx")
            rFEDEXList = GetFedExRates()

            Me.Cursor = Cursors.WaitCursor

            If rUPSList Is Nothing Then
                ReDim rUPSList(1)
            End If

            If rFEDEXList Is Nothing Then
                ReDim rFEDEXList(1)
            End If

            ReDim rList(rUPSList.Length + rFEDEXList.Length)
            rFEDEXList.CopyTo(rList, rUPSList.Length)
            rUPSList.CopyTo(rList, 0)

            Dim selected As Boolean = False
            Dim CARRIER_CODE As String = String.Empty
            Dim rowSOTCARR1 As DataRow = Nothing

            For iCtr As Int16 = 1 To 2
                Select Case iCtr
                    Case 1
                        rList = rUPSList
                        CARRIER_CODE = "UPS"
                        ASCMAIN1.Progress("-", "UPS")
                    Case 2
                        rList = rFEDEXList
                        CARRIER_CODE = "FEDEX"
                        ASCMAIN1.Progress("-", "FedEx")
                End Select

                If rList IsNot Nothing Then
                    For iLoop As Integer = 0 To rList.Count - 1
                        With rList(iLoop)
                            If .ServiceType Is Nothing OrElse .ServiceType = 0 Then
                                Continue For
                            End If

                            ' Display only those services that are mapped to ship vias
                            If dst.Tables("SOTSVIA1").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND CARRIER_PROD_CODE = '" & .ServiceType & "'").Length = 0 Then
                                Continue For
                            End If

                            'CARRIER_SURCHARGE_PERC
                            rowSOTCARR1 = dst.Tables("SOTCARR1").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)

                            If rowSOTCARR1.Table.Columns.Contains("CARRIER_SURCHARGE_PERC") Then
                                CARRIER_SURCHARGE_PERC = Val(rowSOTCARR1.Item("CARRIER_SURCHARGE_PERC") & String.Empty)
                            End If

                            If rowSOTCARR1.Table.Columns.Contains("FRT_PER_SALES_HOLD") Then
                                FRT_PER_SALES_HOLD = Val(rowSOTCARR1.Item("FRT_PER_SALES_HOLD") & String.Empty)
                            End If

                            If rowSOTCARR1.Table.Columns.Contains("CARRIER_PPA_TYPE") Then
                                CARRIER_PPA_TYPE = rowSOTCARR1.Item("CARRIER_PPA_TYPE") & String.Empty
                                ' If not set then set to List
                                If CARRIER_PPA_TYPE.Length = 0 Then
                                    CARRIER_PPA_TYPE = "L"
                                End If
                            End If

                            If rowSOTCARR1.Table.Columns.Contains("CARRIER_SURCHARGE_BASE") Then
                                CARRIER_SURCHARGE_BASE = rowSOTCARR1.Item("CARRIER_SURCHARGE_BASE") & String.Empty
                                ' If not set then set to List
                                If CARRIER_SURCHARGE_BASE.Length = 0 Then
                                    CARRIER_SURCHARGE_BASE = "L"
                                End If
                            End If

                            rowWHTSHPC4 = dst.Tables("WHTSHPC4").NewRow
                            rowWHTSHPC4.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO

                            Select Case CARRIER_CODE
                                Case "UPS"
                                    rowWHTSHPC4.Item("SERVICE_INDEX") = iLoop + 100
                                Case "FEDEX"
                                    rowWHTSHPC4.Item("SERVICE_INDEX") = iLoop + 200
                            End Select

                            rowWHTSHPC4.Item("SERVICE_TYPE_DESC") = IIf(CARRIER_CODE = "SDC", "Stamps ", "") & .ServiceTypeDescription
                            rowWHTSHPC4.Item("CARRIER_CODE") = CARRIER_CODE '
                            rowWHTSHPC4.Item("SHIP_VIA_CODE") = dst.Tables("SOTSVIA1").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND CARRIER_PROD_CODE = '" & .ServiceType & "'")(0).Item("SHIP_VIA_CODE")

                            If (.AccountNetCharge & String.Empty <> "") Then
                                rowWHTSHPC4.Item("ACCT_NET_CHARGE") = Convert.ToDecimal(.AccountNetCharge)
                            Else
                                rowWHTSHPC4.Item("ACCT_NET_CHARGE") = Convert.ToDecimal(.ListNetCharge)
                            End If

                            rowWHTSHPC4.Item("SERVICE_TYPE") = .ServiceType
                            rowWHTSHPC4.Item("SURCHARGE") = 0
                            rowWHTSHPC4.Item("DELIVERY_TIME") = .DeliveryTime
                            rowWHTSHPC4.Item("LIST_NET_CHARGE") = .ListNetCharge
                            If .TransitTime <> "" Then
                                rowWHTSHPC4.Item("TRANSIT_TIME") = .TransitTime
                            End If

                            rowWHTSHPC4.Item("CARRIER_CODE") = CARRIER_CODE

                            Select Case CARRIER_PPA_TYPE
                                Case "F" ' None
                                    rowWHTSHPC4.Item("CUSTOMER_BASE_CHARGE") = 0
                                Case "N" ' Negotiated
                                    rowWHTSHPC4.Item("CUSTOMER_BASE_CHARGE") = rowWHTSHPC4.Item("ACCT_NET_CHARGE")
                                Case "L" ' List
                                    rowWHTSHPC4.Item("CUSTOMER_BASE_CHARGE") = rowWHTSHPC4.Item("LIST_NET_CHARGE")
                                Case Else
                                    ' If not set then use  List
                                    rowWHTSHPC4.Item("CUSTOMER_BASE_CHARGE") = rowWHTSHPC4.Item("LIST_NET_CHARGE")
                            End Select

                            ' Additional Surcharge based off List
                            If CARRIER_SURCHARGE_PERC > 0 Then
                                Select Case CARRIER_SURCHARGE_BASE
                                    Case "N" ' Negotiated
                                        rowWHTSHPC4.Item("SURCHARGE") = Val(rowWHTSHPC4.Item("ACCT_NET_CHARGE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                                    Case "L" ' List
                                        rowWHTSHPC4.Item("SURCHARGE") = Val(rowWHTSHPC4.Item("LIST_NET_CHARGE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                                    Case Else
                                        ' If not set then use  List
                                        rowWHTSHPC4.Item("SURCHARGE") = Val(rowWHTSHPC4.Item("LIST_NET_CHARGE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                                End Select
                            End If

                            dst.Tables("WHTSHPC4").Rows.Add(rowWHTSHPC4)

                        End With
                    Next
                End If
            Next

            grdWHTSHPC4.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

            Try
                Sort_grdColumns(grdWHTSHPC4, "TOTAL_CHARGE")
            Catch ex As Exception
            End Try

            With grdWHTSHPC4.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With

            grdWHTSHPC4_AfterRowUpdate(Nothing, Nothing)

        Catch ex As Exception
            MessageBox.Show("Get Rates error: " & ex.Message, "Get Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

    End Sub

    Private Function GetUpsRates() As WHCSHIP1.RateList()
        Try

            Dim rList(1) As WHCSHIP1.RateList

            If dst.Tables("SOTCARR1").Select("CARRIER_CODE = 'UPS'").Length = 0 Then
                Return Nothing
            End If

            If dst.Tables("SOTCARR3").Select("CARRIER_CODE = 'UPS'").Length = 0 Then
                Return Nothing
            End If

            Dim rowSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Select("CARRIER_CODE = 'UPS'")(0)
            Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = 'UPS'")(0)

            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", cmbWarehouse.Text)
            If rowICTWHSE1 Is Nothing Then
                Return Nothing
            End If

            Dim upsRates As New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.UPS)

            ' Credentials
            With upsRates
                .Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
                .UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
                .Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
                .AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
                .UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                .FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
                .FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                .LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim
            End With

            upsRates.RequestedServiceType = ServiceTypes.stUnspecified
            upsRates.UPSPickupType = UpsratesPickupTypes.ptDailyPickup
            upsRates.CustomerType = UpsratesCustomerTypes.ccRetail

            upsRates.ShipDate = dteShipDate.DateTime

            For Each rowWHTSHPC2 As DataRow In dst.Tables("WHTSHPC2").Select("", "SHIP_PACKAGE_NO", DataViewRowState.CurrentRows)
                Dim pkgDetail As New nsoftware.InShip.PackageDetail

                pkgDetail.Id = rowWHTSHPC2.Item("SHIP_PACKAGE_NO").ToString.PadLeft(8, "0")

                ' Convert to Ounces
                pkgDetail.Weight = (Val(rowWHTSHPC2.Item("POUNDS") & String.Empty) * 16) + Val(rowWHTSHPC2.Item("OUNCES") & String.Empty)
                If pkgDetail.Weight = "0" Then
                    pkgDetail.Weight = "16.0"
                End If

                pkgDetail.PackagingType = CType(Val(rowWHTSHPC2.Item("PACKAGING_TYPE") & String.Empty), UpsratesPickupTypes)
                pkgDetail.Length = Val(rowWHTSHPC2.Item("Length") & String.Empty)
                pkgDetail.Width = Val(rowWHTSHPC2.Item("Width") & String.Empty)
                pkgDetail.Height = Val(rowWHTSHPC2.Item("Height") & String.Empty)

                ' Can have either Insured or Declared not Both
                If chkInsureShipment.Checked Then
                    pkgDetail.InsuredValue = numInsureValue.Value / dst.Tables("WHTSHPC2").Rows.Count
                Else
                    pkgDetail.InsuredValue = 0 ' numInsureValue.Value * -1 / dst.Tables("SOTCART1").Rows.Count
                End If

                upsRates.PackageDetailList.Add(pkgDetail)
            Next

            With upsRates.Sender
                .Company = (rowICTWHSE1.Item("WHSE_DESC") & String.Empty).ToString.Trim
                .FirstName = (rowICTWHSE1.Item("WHSE_CONTACT") & String.Empty).ToString.Trim
                .MiddleInitial = String.Empty
                .LastName = String.Empty
                .Address1 = (rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty).ToString.Trim
                .Address2 = (rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty).ToString.Trim
                .City = (rowICTWHSE1.Item("WHSE_CITY") & String.Empty).ToString.Trim
                .State = (rowICTWHSE1.Item("WHSE_STATE") & String.Empty).ToString.Trim
                .ZipCode = (rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                .CountryCode = (rowICTWHSE1.Item("WHSE_COUNTRY") & String.Empty).ToString.Trim.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                .Phone = (rowICTWHSE1.Item("WHSE_PHONE") & String.Empty).ToString.Trim
            End With

            With upsRates.Recipient
                .FirstName = txtToFirstName.Text
                .MiddleInitial = ""
                .LastName = ""

                .Address1 = txtToStreet.Text
                .Address2 = txtToSuite.Text
                .City = txtToCity.Text
                .State = txtToState.Text
                .ZipCode = txtToZip.Text
                .CountryCode = txtToCountry.Text.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"

                .Company = txtToCompany.Text
                .Phone = txtToPhone.Text

                If .Phone.Trim.Length = 0 Then
                    .Phone = upsRates.Sender.Phone
                End If

                If .Phone.Trim.Length = 0 Then
                    .Phone = "1234567890"
                End If

                .IsResidental = chkFromResidential.Checked
                .IsPOBox = False ' optAddressType.Value = "P"
            End With

            rList = upsRates.GetUPSRatesList()

            If rList Is Nothing Then
                ReDim rList(1)
            End If

            Return rList

        Catch ex As Exception
            MessageBox.Show("The following error occurred getting UPS Rates: " & ex.Message, "Get UPS Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try

    End Function

    Private Function GetFedExRates() As WHCSHIP1.RateList()
        Try

            Dim rList(1) As WHCSHIP1.RateList

            If dst.Tables("SOTCARR1").Select("CARRIER_CODE = 'FEDEX'").Length = 0 Then
                Return Nothing
            End If

            If dst.Tables("SOTCARR3").Select("CARRIER_CODE = 'FEDEX'").Length = 0 Then
                Return Nothing
            End If

            Dim rowSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Select("CARRIER_CODE = 'FEDEX'")(0)
            Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = 'FEDEX'")(0)

            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", cmbWarehouse.Text)
            If rowICTWHSE1 Is Nothing Then
                Return Nothing
            End If

            Dim fedexRates As New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpress)

            ' Credentials
            With fedexRates
                .Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
                .UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
                .Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
                .AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
                .UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                .FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
                .FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                .LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim
            End With

            fedexRates.RequestedServiceType = ServiceTypes.stUnspecified
            fedexRates.UPSPickupType = UpsratesPickupTypes.ptDailyPickup
            fedexRates.CustomerType = UpsratesCustomerTypes.ccRetail
            fedexRates.ShipDate = dteShipDate.DateTime

            For Each rowWHTSHPC2 As DataRow In dst.Tables("WHTSHPC2").Select("", "SHIP_PACKAGE_NO", DataViewRowState.CurrentRows)
                Dim pkgDetail As New nsoftware.InShip.PackageDetail

                pkgDetail.Id = rowWHTSHPC2.Item("SHIP_PACKAGE_NO").ToString.PadLeft(8, "0")

                ' Convert to Ounces
                pkgDetail.Weight = (Val(rowWHTSHPC2.Item("POUNDS") & String.Empty) * 16) + Val(rowWHTSHPC2.Item("OUNCES") & String.Empty)
                If pkgDetail.Weight = "0" Then
                    pkgDetail.Weight = "16.0"
                End If

                pkgDetail.PackagingType = CType(Val(rowWHTSHPC2.Item("PACKAGING_TYPE") & String.Empty), UpsratesPickupTypes)
                pkgDetail.Length = Val(rowWHTSHPC2.Item("Length") & String.Empty)
                pkgDetail.Width = Val(rowWHTSHPC2.Item("Width") & String.Empty)
                pkgDetail.Height = Val(rowWHTSHPC2.Item("Height") & String.Empty)

                ' Can have either Insured or Declared not Both
                If chkInsureShipment.Checked Then
                    pkgDetail.InsuredValue = numInsureValue.Value / dst.Tables("WHTSHPC2").Rows.Count
                Else
                    pkgDetail.InsuredValue = numInsureValue.Value * -1 / dst.Tables("WHTSHPC2").Rows.Count
                End If

                fedexRates.PackageDetailList.Add(pkgDetail)
            Next

            With fedexRates.Sender
                .Company = (rowICTWHSE1.Item("WHSE_DESC") & String.Empty).ToString.Trim
                .FirstName = (rowICTWHSE1.Item("WHSE_CONTACT") & String.Empty).ToString.Trim
                .MiddleInitial = String.Empty
                .LastName = String.Empty
                .Address1 = (rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty).ToString.Trim
                .Address2 = (rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty).ToString.Trim
                .City = (rowICTWHSE1.Item("WHSE_CITY") & String.Empty).ToString.Trim
                .State = (rowICTWHSE1.Item("WHSE_STATE") & String.Empty).ToString.Trim
                .ZipCode = (rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                .CountryCode = (rowICTWHSE1.Item("WHSE_COUNTRY") & String.Empty).ToString.Trim.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                .Phone = (rowICTWHSE1.Item("WHSE_PHONE") & String.Empty).ToString.Trim
            End With

            With fedexRates.Recipient
                .FirstName = txtToFirstName.Text
                .MiddleInitial = ""
                .LastName = ""

                .Address1 = txtToStreet.Text
                .Address2 = txtToSuite.Text
                .City = txtToCity.Text
                .State = txtToState.Text
                .ZipCode = txtToZip.Text
                .CountryCode = txtToCountry.Text.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"

                .Company = txtToCompany.Text
                .Phone = txtToPhone.Text

                If .Phone.Trim.Length = 0 Then
                    .Phone = fedexRates.Sender.Phone
                End If

                If .Phone.Trim.Length = 0 Then
                    .Phone = "1234567890"
                End If

                .IsResidental = chkFromResidential.Checked
                .IsPOBox = False ' optAddressType.Value = "P"
            End With

            rList = fedexRates.GetFedExRatesList()

            If rList Is Nothing Then
                ReDim rList(1)
            End If

            Return rList

        Catch ex As Exception
            MessageBox.Show("The following error occurred getting FedEx Rates: " & ex.Message, "Get FedEx Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try

    End Function


#End Region

#Region "grdWHTSHPC4"

    Private Sub grdWHTSHPC4_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTSHPC4.DoubleClickRow
        Dim CARRIER_CODE As String = e.Row.Cells("CARRIER_CODE").Value
        Dim CARRIER_PROD_CODE As String = e.Row.Cells("SERVICE_TYPE").Value & String.Empty

        ASCMAIN1.sql = "CARRIER_CODE = '" & CARRIER_CODE & "' AND CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "'"
        If dst.Tables("SOTSVIA1").Select(ASCMAIN1.sql).Length > 0 Then
            Dim rowSOTSVIA1 As DataRow = dst.Tables("SOTSVIA1").Select(ASCMAIN1.sql)(0)
            Dim SHIP_VIA_CODE = rowSOTSVIA1.Item("SHIP_VIA_CODE")

            cmbProvider.Value = CARRIER_CODE
            cmbProvider.PerformAction(UltraWinGrid.UltraComboAction.Dropdown)
            cmbProvider.PerformAction(UltraWinGrid.UltraComboAction.CloseDropdown)

            cmbShipMethod.Value = SHIP_VIA_CODE
            cmbShipMethod.PerformAction(UltraWinGrid.UltraComboAction.Dropdown)
            cmbShipMethod.PerformAction(UltraWinGrid.UltraComboAction.CloseDropdown)

        End If

    End Sub

    Private Sub grdWHTSHPC4_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdWHTSHPC4.InitializeRow

        If e.Row.Cells("TOTAL_CHARGE").Value > e.Row.Cells("LIST_NET_CHARGE").Value + e.Row.Cells("SURCHARGE").Value Then
            e.Row.Cells("TOTAL_CHARGE").Appearance.FontData.Bold = DefaultableBoolean.True
            e.Row.Cells("TOTAL_CHARGE").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("TOTAL_CHARGE").Appearance.FontData.Bold = DefaultableBoolean.False
            e.Row.Cells("TOTAL_CHARGE").Appearance.ForeColor = Drawing.Color.Black
        End If

        Dim CARRIER_CODE As String = e.Row.Cells("CARRIER_CODE").Value
        Dim CARRIER_PROD_CODE As String = e.Row.Cells("SERVICE_TYPE").Value

        If CARRIER_CODE.Length > 0 AndAlso CARRIER_PROD_CODE.Length > 0 Then
            If dst.Tables("SOTSVIA1").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "'")(0).Item("SHIP_VIA_CODE").length > 0 Then
                e.Row.Cells("SHIP_VIA_CODE").Value = dst.Tables("SOTSVIA1").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "'")(0).Item("SHIP_VIA_CODE")
            End If
        End If
    End Sub

    Private Sub grdWHTSHPC4_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdWHTSHPC4.AfterRowUpdate

        Dim SERVICE_INDEX As Int16 = -1
        For Each dRow As DataRow In dst.Tables("WHTSHPC4").Select("SELECTED = '1'", "", DataViewRowState.CurrentRows)
            SERVICE_INDEX = dRow.Item("SERVICE_INDEX")
            Exit For
        Next

        Dim view As New DataView(dst.Tables("WHTSHPCA"))
        view.RowFilter = "SERVICE_INDEX = " & SERVICE_INDEX
        'grdWHTSHPCA.DataSource = view
        'Sort_grdColumns(grdWHTSHPCA, "SELECTED,ADDON_TYPE_DESC")
        'grdWHTSHPCA_AfterRowUpdate(Nothing, Nothing)

        If dst.Tables("WHTSHPC4").Select("SELECTED = '1'").Length > 0 Then
            'SetupPackages(dst.Tables("WHTSHPC4").Select("SELECTED = '1'")(0).Item("SHIP_VIA_CODE") & String.Empty)
        Else
            'SetupPackages(txtShipViaWhtship1.Text)
        End If

    End Sub

    Private Sub grdWHTSHPC4_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdWHTSHPC4.BeforeRowUpdate

        If e.Row.Band.Key <> grdWHTSHPC4.DisplayLayout.Bands(0).Key Then
            Exit Sub
        End If

        If e.Row.Cells("SELECTED").Value = "1" Then
            For Each dRow As DataRow In dst.Tables("WHTSHPC4").Select("SELECTED = '1'", "", DataViewRowState.CurrentRows)
                dRow.Item("SELECTED") = "0"
            Next
        End If

    End Sub

#End Region

#Region "Form Procedures"

    ''' <summary>
    ''' Request weight from scale
    ''' </summary>
    ''' <remarks></remarks>
    Private Function GetScaleWeight() As Decimal

        Try
            Try
                ASCMAIN1.scaleweight = String.Empty
                With ASCMAIN1.ScalePort
                    .DiscardInBuffer()
                    .DiscardOutBuffer()
                    .WriteLine("W" & vbCrLf)
                    System.Threading.Thread.Sleep(1000)
                End With
            Catch ex As Exception
                Return 0
            End Try

            ' Need to set to at least one ounce.
            If Val(ASCMAIN1.scaleweight) > 0 AndAlso Val(ASCMAIN1.scaleweight) < 0.06 Then
                ASCMAIN1.scaleweight = "0.06"
            End If

            Return Val(ASCMAIN1.scaleweight)

        Catch ex As Exception
            MessageBox.Show("Get Scale Weight Error: " & ex.Message, "Get Scale Weight", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' Fires when the weight is requested from the scale
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ProcessScaleData(ByVal scaledata As String)

        If MdiParent.ActiveMdiChild Is Nothing Then Exit Sub
        If MdiParent.ActiveMdiChild.Name <> Me.Name Then Exit Sub

        Try
            'Dim length As Int16 = ASCMAIN1.ScaleSerialPort.BytesToRead
            'If length > 0 Then
            '    Dim numberOfBytesRead As Int16 = 0
            '    Dim readBuffer(length) As Byte
            '    numberOfBytesRead = ASCMAIN1.ScaleSerialPort.Read(readBuffer, 0, length)
            '    MessageBox.Show(readBuffer.ToString)
            '    registeredWeight = Val(readBuffer)
            'End If
        Catch ex As Exception

        End Try
    End Sub

    Private Function RequestReturnLabel(ByVal SHIP_PACKAGE_NO As Int16) As Boolean

        RequestReturnLabel = True

        Try
            Dim CARRIER_CODE As String = cmbProvider.SelectedRow.Cells("CARRIER_CODE").Value
            Dim labelFiles As New List(Of String)

            Dim ReturnLabelRequestDetail As New TAC.SHCUPSCR.ReturnLabelRequestDetail
            With ReturnLabelRequestDetail

                Dim ShippingLabelFile As String = ASCMAIN1.Folders("Temp")
                If Not ShippingLabelFile.EndsWith("\") Then
                    ShippingLabelFile &= "\"
                End If

                .labelEmailAddress = txtToEmail.Text
                Select Case optReturnLabels.Value
                    Case CarrierMailsLabelsToCustomer : .labelDeliveryMethod = SHCUPSCR.LabelDeliveryMethods.CarrierMailsLabelsToCustomer
                    Case CarrierEmailLabelsToCustomer : .labelDeliveryMethod = SHCUPSCR.LabelDeliveryMethods.CarrierEmailLabelsToCustomer
                    Case Printlabels : .labelDeliveryMethod = SHCUPSCR.LabelDeliveryMethods.Printlabels
                    Case EmailLabelsToCustomer : .labelDeliveryMethod = SHCUPSCR.LabelDeliveryMethods.EmailLabelsToCustomer
                    Case UPSOneAttemptToPickup : .labelDeliveryMethod = SHCUPSCR.LabelDeliveryMethods.UPSOneAttemptToPickup
                    Case UPSThreeAttemptsToPickup : .labelDeliveryMethod = SHCUPSCR.LabelDeliveryMethods.UPSThreeAttemptsToPickup
                End Select

                .ShipFastestMethod = False
                .ServiceType = Val(cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value & String.Empty)

                For Each rowWHTSHPC2 As DataRow In dst.Tables("WHTSHPC2").Select("SHIP_PACKAGE_NO = " & SHIP_PACKAGE_NO)
                    Dim ShippingPackage As New TAC.SHCUPSCR.Package
                    With ShippingPackage
                        .AdditionalHandling = False
                        .Description = "Floral Items"
                        .Height = Val(rowWHTSHPC2.Item("HEIGHT") & String.Empty)
                        .InsuredValue = 0
                        .LargePackage = False
                        .Length = Val(rowWHTSHPC2.Item("LENGTH") & String.Empty)

                        If optReturnLabels.Value = EmailLabelsToCustomer Then
                            .ShippingLabelFile = ShippingLabelFile & Guid.NewGuid.ToString.Replace("-", "") & ".gif"
                        Else
                            .ShippingLabelFile = ShippingLabelFile & Guid.NewGuid.ToString.Replace("-", "") & ".txt"
                        End If
                        labelFiles.Add(.ShippingLabelFile)

                        .Weight = Val(rowWHTSHPC2.Item("PKG_WEIGHT") & String.Empty)
                        If .Weight < 16 Then
                            .Weight = 16
                        End If
                        .Width = Val(rowWHTSHPC2.Item("WIDTH") & String.Empty)

                        Select Case optUserSuppliedValue.Value & String.Empty
                            Case "R" ' RMA
                                .Reference = "RMA:" & txtUserSuppliedValue.Text & ";"
                            Case "I"
                                .Reference = "IN:" & txtUserSuppliedValue.Text & ";"
                            Case Else
                                .Reference = "CR:" & optUserSuppliedValue.CheckedItem.DisplayText & " " & txtUserSuppliedValue.Text & ";"
                        End Select

                        If ASCMAIN1.USER_SECURITY_CODEs.Contains("SY") Then
                            .InsuredValue += Val(rowWHTSHPC2.Item("INSURED_VALUE") & String.Empty)
                        End If
                    End With

                    .ShippingPackages.Add(ShippingPackage)
                Next

                GetSenderInfo()
                GetRecipientInfo()
                GetShipPayor()

                With .Recipient
                    .Address1 = clsShip.Recipient.Address1
                    .Address2 = clsShip.Recipient.Address2
                    .Address3 = clsShip.Recipient.Address3
                    .City = clsShip.Recipient.City
                    .Company = clsShip.Recipient.Company
                    .Country = clsShip.Recipient.CountryCode
                    .isPOBox = clsShip.Recipient.IsPOBox
                    .isResidental = clsShip.Recipient.IsResidental
                    .Name = clsShip.Recipient.FirstName
                    .Phone = clsShip.Recipient.Phone
                    .State = clsShip.Recipient.State
                    .ZipCode = txtToZip.Text
                    .eMail = txtToEmail.Text
                End With

                With .Shipper
                    .Address1 = clsShip.Sender.Address1
                    .Address2 = clsShip.Sender.Address2
                    .Address3 = clsShip.Sender.Address3
                    .City = clsShip.Sender.City
                    .Company = clsShip.Sender.Company
                    .Country = clsShip.Sender.CountryCode
                    .isPOBox = clsShip.Sender.IsPOBox
                    .isResidental = clsShip.Sender.IsResidental
                    .Name = clsShip.Sender.FirstName
                    .Phone = clsShip.Sender.Phone
                    .State = txtFromState.Text
                    .ZipCode = clsShip.Sender.ZipCode

                    .eMail = String.Empty
                    Dim rowICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(cmbWarehouse.Text)
                    If rowICTWHSE1 IsNot Nothing Then
                        .eMail = rowICTWHSE1.Item("WHSE_EMAIL") & String.Empty
                    End If
                End With

                With .Payor
                    .PayorType = clsShip.Payor
                    .AccountNumber = clsShip.PayorContact.AccountNumber
                    .AccountZipCode = clsShip.PayorContact.ZipCode
                    .CountryCode = clsShip.PayorContact.CountryCode
                End With

                Dim Sql As String = "CARRIER_CODE = '" & CARRIER_CODE & "' AND DIVISION_CODE = '" & cmbDivision.Text & "'"
                Dim rowSOTCARR3 As DataRow = Nothing
                If dst.Tables("SOTCARR3").Select(Sql).Length > 0 Then
                    rowSOTCARR3 = dst.Tables("SOTCARR3").Select(Sql)(0)
                End If

                Dim rowSOTCARR1 As DataRow = Nothing
                If dst.Tables("SOTCARR1").Select("CARRIER_CODE = '" & CARRIER_CODE & "'").Length > 0 Then
                    rowSOTCARR1 = dst.Tables("SOTCARR1").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)
                End If

                Dim upsCredentials As New TAC.SHCUPSCR.Credentials

                If rowSOTCARR3 IsNot Nothing AndAlso rowSOTCARR1 IsNot Nothing Then
                    With upsCredentials
                        .AccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                        .AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
                        '.LabelImageType = ""
                        .Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
                        .Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
                        .UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
                        .AddressVaildationAccessLicenseNumber = rowSOTCARR3.Item("ADDR_VALID_LICENSE") & String.Empty
                    End With
                Else
                    MessageBox.Show("Cannot determine the " & CARRIER_CODE & " Credentials.", "Return Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Function
                End If

                Dim clsSHCUPSCR As New TAC.SHCUPSCR(upsCredentials)
                Dim response As TAC.SHCUPSCR.Response = clsSHCUPSCR.RequestReturnLabel(ReturnLabelRequestDetail)
                If response Is Nothing Then
                    If clsSHCUPSCR.LastError.Length > 0 Then
                        MessageBox.Show(clsSHCUPSCR.LastError, "Return Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        MessageBox.Show("Error Requesting Return Label", "Return Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                    Return False
                End If

                If optUserSuppliedValue.Value = "R" AndAlso rowSOTRMAF1 IsNot Nothing Then

                    Try
                        BeginTrans()
                        Dim rowSOTRMAFL As DataRow = dst.Tables("SOTRMAFL").NewRow
                        rowSOTRMAFL.Item("RA_NO") = rowSOTRMAF1.Item("RA_NO")
                        rowSOTRMAFL.Item("CARRIER_CODE") = cmbProvider.SelectedRow.Cells("CARRIER_CODE").Text
                        rowSOTRMAFL.Item("TRACKING_NO") = response.TrackingNumber
                        rowSOTRMAFL.Item("SHIP_VIA_CODE") = cmbShipMethod.SelectedRow.Cells("SHIP_VIA_CODE").Text
                        rowSOTRMAFL.Item("DELIVERY_METHOD") = optReturnLabels.Value
                        rowSOTRMAFL.Item("EMAIL_ADDRESS") = txtToEmail.Text
                        rowSOTRMAFL.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        rowSOTRMAFL.Item("INIT_DATE") = DateTime.Now
                        dst.Tables("SOTRMAFL").Rows.Add(rowSOTRMAFL)

                        Update_Record_TDA("SOTRMAFL")

                        ASCMAIN1.sql = "Update SOTRMAF1 SET CALL_TAG_USER = '" & ASCMAIN1.USER_ID & "', CALL_TAG_DATE = SYSDATE WHERE RA_NO = '" & rowSOTRMAF1.Item("RA_NO") & "' AND CALL_TAG_USER IS NULL"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                        CommitTrans()
                    Catch ex As Exception
                        Rollback()
                    End Try

                End If

                ' Need to Print Label 
                Select Case optReturnLabels.Value
                    Case Printlabels
                        If My.Computer.FileSystem.FileExists(response.ShippingLabelFile) Then
                            'Dim vLabelPrinter As ASCPRINT = New ASCPRINT(ASCMAIN1.LabelPrinterSerialPort)
                            Dim label As String = String.Empty

                            Using sr As New StreamReader(response.ShippingLabelFile)
                                label = sr.ReadToEnd
                                sr.Close()
                            End Using

                            'vLabelPrinter.SendStringToPrinter(txtlabelPrinter.Text, label)
                            If ASCMAIN1.CLIENT = "VAN" Then
                                If txtlabelPrinter.BackColor = Drawing.Color.Green Then
                                    ASCMAIN1.LabelPrinterSerialPort.WriteLine(label)
                                Else
                                    Dim vLabelPrinter As New ASCPRINT
                                    Return vLabelPrinter.SendStringToPrinter(cboZebraPrinter.Text, label)
                                End If

                            Else
                                ASCMAIN1.LabelPrinterSerialPort.WriteLine(label)
                            End If

                            My.Computer.FileSystem.DeleteFile(response.ShippingLabelFile)

                        Else
                            RequestReturnLabel = False
                            MessageBox.Show("Cannot determine the Return Label File Location.", "Return Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Function
                        End If

                    Case EmailLabelsToCustomer
                        ReturnLabelsToSendToCustomers.Add(response.ShippingLabelFile)
                End Select


            End With

        Catch ex As Exception
            RequestReturnLabel = False
            MessageBox.Show(ex.Message)
        End Try
    End Function


    Private Function EmailReturnLabels() As Boolean

        Dim attachFileName As String = String.Empty
        Dim customerEmailFound As Boolean = False

        Try

            Dim emailToList As String = txtToEmail.Text
            Dim CUST_XMIT_INV_VIA As String = String.Empty

            ' remove double semi-colons
            While emailToList.Contains(" ")
                emailToList = emailToList.Replace(" ", "")
            End While
            emailToList = emailToList.Replace(",", ";")

            While emailToList.Contains(";;")
                emailToList = emailToList.Replace(";;", ";")
            End While

            ' should be at least 5 characters
            If emailToList.Replace(";", "").Trim.Length < 5 Then
                ' Return False
            End If

            ' Concatentate and process all email addresses
            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            For Each emailAddress As String In (emailToList).ToString.Split(";")
                emailAddress = emailAddress.Trim
                If emailAddress.Length > 5 AndAlso Not EMAIL_ADDRESSs.Keys.Contains(emailAddress) Then
                    EMAIL_ADDRESSs.Add(emailAddress, emailAddress)
                End If
            Next

            Dim ATTACHMENTs As New Dictionary(Of String, String)
            Dim ictr As Int16 = 1
            For Each fileName As String In ReturnLabelsToSendToCustomers
                ATTACHMENTs.Add("RTN LABEL " & ictr, fileName)
                ictr += 1
            Next

            If EMAIL_ADDRESSs.Count = 0 Then
                Return False
            End If

            Dim SUBJECT As String = "Return Labels"
            Dim SEND_NO As String = String.Empty

            ' Need to attach the letter to the sales order when we do no thave an email address.
            If emailToList.Replace(";", "").Trim.Length < 5 OrElse EMAIL_ADDRESSs.Count = 0 Then
                Return False
            End If

            SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
                   (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                    SUBJECT, "RTNLBL", True, False, "", txtToCompany.Text, "Return Label")


        Catch ex As Exception
            Return False
        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

    End Function

#End Region

End Class