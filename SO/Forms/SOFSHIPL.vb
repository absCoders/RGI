Imports Infragistics.Win.UltraWinGrid
Imports DPayments.DShippingSDK

Public Class SOFSHIPL
    ' proceed prereq - after maintenance, or confirmation,need to verify that auth amt on credit check and credit card has not been violated
    Dim expSOTPICK1 As New Dictionary(Of String, String)
    Dim commonCarriersOnly As Boolean = False
    Dim Carrier_Code As String
    Dim Batched_Group As Boolean
    Dim Group_No_Selected As String = ""

    Private Const Printed As String = "Printed"
    Private Const NotPrinted As String = "Not Printed"
    Private Const PartiallyPrinted As String = "Partially Printed"
    Private COMPANY_CODE As String = String.Empty
    Private permitVandaleToPrintUpsConsignee As Integer = 0

#Region "Declarations"
    Dim CUST_CODE As String
    Dim CUST_NAME As String         ' Sold-To Customer Name
    Dim SHIP_BOL_NOs As New List(Of String)
    Dim ORDR_GROUP_NO As String
    Dim ORDR_CUST_PO As String
    Dim rowARTCUST1 As DataRow

    Dim sqlSOTPICK1 As String
    Dim sqlSOTPICK2 As String
    Dim sqlSOTSHIPX As String

    Dim ORDR_SOURCE As String
    Dim SOTSHIP0 As String
    Dim CURR_CODE As String
    Dim CURR_EXCH_RATE As Decimal
    Dim ASW As New Dictionary(Of String, String)

    Dim ORDR_SHIP_DATE As Date
    Dim ORDR_CANCEL_DATE As Date

    Dim SOTSHIPX As String

    Private commonCarrier As Boolean = False
    Private RecreateLabel As Boolean = False

    Private tblTATSTATE As DataTable

    Private WithEvents ultraComboPackage As Infragistics.Win.UltraWinGrid.UltraCombo = New Infragistics.Win.UltraWinGrid.UltraCombo

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("GLTPARM1")
        tblTATSTATE = ASCDATA1.GetDataTable("SELECT * FROM TATSTATE", "TATSTATE") ' WHERE region_code is not null

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            COMPANY_CODE = "VAN"
        ElseIf ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
            COMPANY_CODE = "NYA"
        ElseIf ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            COMPANY_CODE = "RGI"
        Else
            COMPANY_CODE = ASCMAIN1.CLIENT
        End If

        With dst

            ' Create_TDA(.Tables.Add, "EDT855O1", "*")
            ' Create_TDA(.Tables.Add, "EDT855O2", "*")
            'Create_TDA(.Tables.Add, "EDT855O3", "*")
            'Create_TDA(.Tables.Add, "EDT855O5", "*")

            sqlSOTSHIPX = " Select Distinct SOTSHIP1.*, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.ORDR_CUST_PO," & vbCr _
            & " SOTORDR1.CUST_CODE, SOTORDR5.CUST_ADDR_CODE," & vbCr _
            & " SOTORDR5.CUST_NAME, SOTORDR5.CUST_ADDR1, SOTORDR5.CUST_ADDR2, " & vbCr _
            & " SOTORDR5.CUST_CITY, SOTORDR5.CUST_STATE, SOTORDR5.CUST_ZIP_CODE, " & vbCr _
            & " SOTORDR5.CUST_COUNTRY," & vbCr _
            & " SOTORDR5.CUST_CONTACT, SOTORDR5.CUST_PHONE, FDX_ACCT_NO, CUST_STORE_NO" & vbCr _
            & " from SOTORDR1, SOTORDR5, SOTPICK1, SOTSHIP1, ARTCUST2" & vbCr _
            & " where SOTORDR1.ORDR_NO = SOTORDR5.ORDR_NO " & vbCr _
            & " And SOTPICK1.SHIP_BOL_NO  = SOTSHIP1.SHIP_BOL_NO" & vbCr _
            & " And SOTPICK1.ORDR_NO = SOTORDR5.ORDR_NO" & vbCr _
            & " And SOTORDR1.CUST_CODE = ARTCUST2.CUST_CODE" & vbCr _
            & " And SOTORDR5.CUST_ADDR_CODE = ARTCUST2.CUST_ADDR_CODE" & vbCr _
            & " And SOTSHIP1.SHIP_STATUS in ('P','F')" & vbCr _
            & " And SOTORDR5.CUST_ADDR_TYPE = 'ST'" & vbCr
            ASCMAIN1.sql = sqlSOTSHIPX & " and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, False, "", 1)
            dst.Tables("SOTSHIP1").Columns.Add("SELECTED")
            dst.Tables("SOTSHIP1").Columns.Add("STATUS")
            dst.Tables("SOTSHIP1").Columns.Add("ERROR_MSG")
            dst.Tables("SOTSHIP1").Columns("STATUS").DefaultValue = NotPrinted
            dst.Tables("SOTSHIP1").Columns.Add("CUST_ADDR3")

            ASCMAIN1.sql = "SELECT SOTSVIA1.*, SOTCARR1.CARRIER_TYPE" _
                & " FROM SOTSVIA1, SOTCARR1" _
                & " WHERE SOTSVIA1.CARRIER_CODE = SOTCARR1.CARRIER_CODE"
            Create_TDA(.Tables.Add, "SOTSVIA1", "**", 1, False, "", 1)
            Fill_Records("SOTSVIA1", "", False, ASCMAIN1.sql)

            ' Shipping Label
            Create_TDA(.Tables.Add, "WHTSHPC1", "*")
            Create_TDA(.Tables.Add, "WHTSHPC2", "*")
            Create_TDA(.Tables.Add, "WHTSHPC3", "*")
            Create_TDA(.Tables.Add, "WHTSHPC5", "*")
            Create_TDA(.Tables.Add, "WHTSHPCC", "*")
            Create_TDA(.Tables.Add, "WHTSHPCS", "*")
            Create_TDA(.Tables.Add, "WHTSHPCP", "*")

            ' Carrier Tables
            Create_TDA(.Tables.Add, "SOTCARR1", "*")
            Create_TDA(.Tables.Add, "SOTCARR2", "*")
            Create_TDA(.Tables.Add, "SOTCARR3", "*")
            Create_TDA(.Tables.Add, "SOTCARRR", "*")
            Create_TDA(.Tables.Add, "ARTCUSTS", "*")

            Fill_Records("SOTCARR1", "", True, "SELECT * FROM SOTCARR1")
            Fill_Records("SOTCARR2", "", True, "SELECT * FROM SOTCARR2")
            Fill_Records("SOTCARR3", "", True, "Select SOTCARR3.*, SOTCARR1.CARRIER_DESC From SOTCARR3, SOTCARR1 Where SOTCARR3.CARRIER_CODE = SOTCARR1.CARRIER_CODE")
            Fill_Records("SOTCARRR", "", True, "SELECT * FROM SOTCARRR")

            Create_TDA(.Tables.Add, "SOTCART1", "*", , , , , "CART_TRACKING_NO,CART_TOTAL_WGT_ACTUAL,CART_TYPE,PACKAGING_TYPE,PKG_CODE")
            .Tables("SOTCART1").Columns.Add("REFERENCE1", GetType(System.String))
            .Tables("SOTCART1").Columns("REFERENCE1").DefaultValue = String.Empty
            .Tables("SOTCART1").Columns("REFERENCE1").MaxLength = 20
            .Tables("SOTCART1").Columns.Add("REFERENCE2", GetType(System.String))
            .Tables("SOTCART1").Columns("REFERENCE2").DefaultValue = String.Empty
            .Tables("SOTCART1").Columns("REFERENCE2").MaxLength = 20
            .Tables("SOTCART1").Columns.Add("REFERENCE3", GetType(System.String))
            .Tables("SOTCART1").Columns("REFERENCE3").DefaultValue = String.Empty
            .Tables("SOTCART1").Columns("REFERENCE3").MaxLength = 20
            .Tables("SOTCART1").Columns.Add("WIDTH", GetType(System.Int16))
            .Tables("SOTCART1").Columns("WIDTH").DefaultValue = 0
            .Tables("SOTCART1").Columns.Add("LENGTH", GetType(System.Int16))
            .Tables("SOTCART1").Columns("LENGTH").DefaultValue = 0
            .Tables("SOTCART1").Columns.Add("HEIGHT", GetType(System.Int16))
            .Tables("SOTCART1").Columns("HEIGHT").DefaultValue = 0
            .Tables("SOTCART1").Columns.Add("STATUS")
            .Tables("SOTCART1").Columns.Add("ERROR_MSG")

            ASCMAIN1.sql = " Select SOTCART2.CART_NO, Sum(QTY_PACKED * STYLE_WEIGHT) as WEIGHT " _
            & " from SOTCART1, SOTCART2, ICTSTYL1, SOTPICK1,SOTSHIP1 " _
            & "  Where SOTCART2.CART_NO = SOTCART1.CART_NO" _
            & "  AND SOTCART2.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
            & "  And SOTCART1.PICK_NO = SOTPICK1.PICK_NO" _
            & "  And SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" _
            & "  And Rownum < 1" _
            & " Group by SOTCART2.CART_NO"
            Create_TDA(.Tables.Add, "SOTCARTL", "**", 0, False, "", 1)

            'SOTCART2 is for intl shipments
            ASCMAIN1.sql = "Select SOTCART2.*, SOTCART1.PICK_NO, SOTCART2.QTY_PACKED QTY_PACKED_ORIG" & vbCrLf _
            & " from SOTCART2,SOTCART1 where SOTCART1.CART_NO = SOTCART2.CART_NO"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0)

            sqlSOTPICK1 = "Select SOTPICK1.*" & vbCrLf _
                & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_SOURCE" & vbCrLf _
                & ", SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.CUST_BILL_TO_CUST" & vbCrLf _
                & ", SOTORDR1.POST_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_FOB" & vbCrLf _
                & ", SOTORDR1.TERM_CODE, SOTORDR1.SREP_CODE, SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
                & ", SOTSHIP1.BILL_OF_LADING_NO, SOTORDR1.ORDR_INV_COMMENT, SOTORDR1.CUST_FACTOR_IND" & vbCrLf _
                & ", SOTORDR1.CCPA_NO CCPA_NO_ORDR, SOTORDR1.CUST_DC_NO" & vbCrLf _
                & " from SOTPICK1,SOTORDR1,SOTSHIP1 "
            ASCMAIN1.sql = sqlSOTPICK1 & " where ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**")
            dst.Tables("SOTPICK1").Columns.Add("OUR_FREIGHT", GetType(System.Decimal))
            dst.Tables("SOTPICK1").Columns("OUR_FREIGHT").DefaultValue = 0
            'dst.Tables("SOTPICK1").Columns.Add("ERROR_MSG") '.Columns("PICK_AMT_CONF").Expression = IIf(Not tf, "", "ISNULL(PICK_UNIT_PRICE,0) * ISNULL(PICK_QTY_CONF,0)")

            Create_Relation("SOTPICK1", "SOTCART1", "PICK_NO")
            dst.Tables("SOTCART1").Columns.Add("SHIP_BOL_NO", GetType(System.String), "PARENT(SOTPICK1_SOTCART1).SHIP_BOL_NO")
            dst.Tables("SOTCART1").Columns.Add("CUST_STORE_NO", GetType(System.String), "PARENT(SOTPICK1_SOTCART1).CUST_STORE_NO")
            dst.Tables("SOTCART1").Columns.Add("ORDR_CUST_PO", GetType(System.String), "PARENT(SOTPICK1_SOTCART1).ORDR_CUST_PO")

            'SOTPICK2 is for intl shipments
            sqlSOTPICK2 = "Select SOTPICK2.*, " & vbCrLf _
                & " SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_DESC, " & vbCrLf _
                & " SOTORDR2.ORDR_UNIT_PRICE, SOTORDR2.STYLE_CODE_SUB," & vbCrLf _
                & " SOTORDR2.RANGE_STYLE_CODE, SOTORDR2.RANGE_STYLE_LNO, SOTORDR2.QTY_PER_PP" & vbCrLf _
                & " from SOTPICK2,SOTPICK1,SOTORDR2,SOTSHIP1" & vbCrLf
            ASCMAIN1.sql = sqlSOTPICK2 & " where ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "SOTPICK2", "**")

            ASCMAIN1.sql = "SELECT SOTCART1.CART_NO, SOTSHIP1.SHIP_BOL_NO, SOTPICK1.ORDR_NO, SOTPICK1.PICK_NO,
                            SOTCART1.CART_TYPE NOT_ASSIGNED, SOTCART1.CART_TRACKING_NO TRACKING_NO, SOTCART1.CART_TYPE NOT_SCANNED
                            FROM SOTCART1, SOTSHIP1, SOTPICK1
                            WHERE ROWNUM < 1"
            Create_TDA(.Tables.Add, "TRACKING", ASCMAIN1.sql, 0, False, String.Empty, 0)
            With dst.Tables("TRACKING")
                .Columns.Add("STATUS", GetType(System.String))
                .Columns.Add("DATE", GetType(System.String))
                .Columns.Add("TIME", GetType(System.String))
                .Columns.Add("ADDRESS1", GetType(System.String))
                .Columns.Add("ADDRESS2", GetType(System.String))
                .Columns.Add("CITY", GetType(System.String))
                .Columns.Add("STATE", GetType(System.String))
                .Columns.Add("ZIPCODE", GetType(System.String))
                .Columns.Add("COUNTRYCODE", GetType(System.String))
                .Columns.Add("LOCATION", GetType(System.String))
                .Columns.Add("SHIP_TO_ZIPCODE", GetType(System.String))
                .Columns.Add("ZIPCODE_NO_MATCH", GetType(System.String))
                .Columns.Add("CUST_ADDR_CODE", GetType(System.String))
                .Columns.Add("SHIPMENT_CARTONS", GetType(System.Int64))
                .Columns.Add("SHIPMENT_LABELS", GetType(System.Int64))
                .Columns.Add("SHIPMENT_STATUS", GetType(System.String))
            End With

        End With

        With ultraComboPackage.DisplayLayout.Bands(0)
            ultraComboPackage.Font = grdSOTCART1.Font
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
        grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_CODE").EditorComponent = ultraComboPackage


        cbeREF1.DataSource = ASCDATA1.GetDataTable("SELECT REF_CODE, REF_DESC from SOTCARRR WHERE ROWNUM < 1") '  Where COLUMN_NAME in ('ORDR_CUST_PO','CART_NO','PICK_NO')")
        cbeREF1.ValueMember = "REF_CODE"
        cbeREF1.DisplayMember = "REF_DESC"

        cbeREF2.DataSource = ASCDATA1.GetDataTable("SELECT REF_CODE, REF_DESC from SOTCARRR WHERE ROWNUM < 1") ' Where COLUMN_NAME in ('ORDR_CUST_PO','CUST_STORE_NO')")
        cbeREF2.ValueMember = "REF_CODE"
        cbeREF2.DisplayMember = "REF_DESC"

        cbeREF3.DataSource = ASCDATA1.GetDataTable("SELECT REF_CODE, REF_DESC from SOTCARRR WHERE ROWNUM < 1") '  Where COLUMN_NAME in ('CART_NO')")
        cbeREF3.ValueMember = "REF_CODE"
        cbeREF3.DisplayMember = "REF_DESC"

        cbeREF_CODE1.DataSource = ASCDATA1.GetDataTable("SELECT REF_CODE, REF_DESC from SOTCARRR WHERE ROWNUM < 1") '  Where COLUMN_NAME in ('CART_NO')")
        cbeREF_CODE1.ValueMember = "REF_CODE"
        cbeREF_CODE1.DisplayMember = "REF_CODE"

        cbeREF_CODE2.DataSource = ASCDATA1.GetDataTable("SELECT REF_CODE, REF_DESC from SOTCARRR WHERE ROWNUM < 1") '  Where COLUMN_NAME in ('CART_NO')")
        cbeREF_CODE2.ValueMember = "REF_CODE"
        cbeREF_CODE2.DisplayMember = "REF_CODE"

        cbeREF_CODE3.DataSource = ASCDATA1.GetDataTable("SELECT REF_CODE, REF_DESC from SOTCARRR WHERE ROWNUM < 1") '  Where COLUMN_NAME in ('CART_NO')")
        cbeREF_CODE3.ValueMember = "REF_CODE"
        cbeREF_CODE3.DisplayMember = "REF_CODE"


        grdSOTSHIP1.DataSource = dst.Tables("SOTSHIP1")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")

        grdSOTSHIP1.DisplayLayout.UseFixedHeaders = True
        With grdSOTSHIP1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SELECTED"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With
        With grdSOTSHIP1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"SELECTED", "CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY", "CUST_NAME", "CUST_PHONE", "CUST_CONTACT"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With


        grdSOTCART1.DisplayLayout.UseFixedHeaders = True
        With grdSOTCART1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CART_NO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With
        With grdSOTCART1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"CART_FREIGHT", "CART_TOTAL_WGT_ACTUAL", "PKG_CODE", "PACKAGING_TYPE",
                                 "WIDTH", "LENGTH", "HEIGHT", "CART_SEQ", "REFERENCE1"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    'gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    'gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        Create_Summary(grdSOTSHIP1, "STATUS", "Count")
        Create_Summary(grdSOTCART1, "CART_NO", "Count")
        Create_Summary(grdSOTCART1, New String() _
            {"CART_FREIGHT", "CART_TOTAL_UNITS", "CART_TOTAL_WGT_ACTUAL"})

        grdTracking.DataSource = dst.Tables("TRACKING")
        Create_Summary(grdTracking, "CART_NO", "Count")
        grdTracking.DisplayLayout.Bands(0).Columns("SHIPMENT_CARTONS").Hidden = True

        Dim ZebraPrinters As New List(Of String)
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
                If printerName.ToUpper.StartsWith("ZDESIGNER") Or printerName.ToUpper.StartsWith("MONARCH") Or printerName.ToUpper.StartsWith("AVERY") Or printerName.ToUpper.StartsWith("ZEBRA") Then
                    ZebraPrinters.Add(printerName)
                End If
            Next printerName
            If ZebraPrinters.Count >= 1 Then
                cboZebraPrinter.DataSource = ZebraPrinters
            End If
        End If


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select"
                If Batched_Group = False Then
                    Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", Absx1.txtFor("ORDR_GROUP_NO").Text)
                    If rowSOTORDR0 IsNot Nothing Then
                        Absx1.txtFor("CUST_CODE").Text = rowSOTORDR0.Item("CUST_CODE")

                    Else
                        EMsg &= vbCr & "Invalid Group"

                    End If
                End If

                'If EMsg <> "" Then
                '    ASCMAIN1.MultiTask_Release()
                'End If
            Case "Print Labels"
                If dst.Tables("SOTSHIP1").Select("SELECTED = '1'").Count = 0 Then
                    EMsg &= vbCr & "You must Select at Least 1 Shipment to Print"
                End If

                Dim Future_Ship As Date = DateAdd("d", 10, Now)
                If dteSHIP_DATE_SHIPPED.Value = Nothing Then
                    EMsg &= vbCr & "You must Select a Ship Date"
                Else
                    If DateValue(dteSHIP_DATE_SHIPPED.Value) > DateValue(Future_Ship) Then
                        'EMsg &= vbCr & "Ship Date to far in Advance, must be less than 10 days"
                    End If
                End If
                If Absx1.txtFor("SHIP_VIA_CODE").Text = "" Then
                    EMsg &= vbCr & "You must First Select a Ship Via"
                End If
                If Absx1.txtFor("PACKAGE_TYPE").Text = "" Then
                    EMsg &= vbCr & "You must First Select a Package Type"
                End If
                If Absx1.txtFor("PKG_CODE").Text = "" Then
                    EMsg &= vbCr & "You must First Select a Pkg Code"
                End If
                If Absx1.txtFor("BILLING_TYPE").Text = "" Then
                    EMsg &= vbCr & "You must First Select a Billing Type"
                End If

                If Absx1.numFor("CARTONL").Value = 0 Then
                    EMsg &= vbCr & "You must Enter a Carton Length"
                End If
                If Absx1.numFor("CARTONW").Value = 0 Then
                    EMsg &= vbCr & "You must Enter a Carton Width"
                End If
                If Absx1.numFor("CARTONH").Value = 0 Then
                    EMsg &= vbCr & "You must Enter a Carton Height"
                End If

                'If dst.Tables("SOTSHIP1").Select("CUST_ZIP_CODE").Length <> 0 Then
                'End If
                'If txtCUST_ADDR1.TextLength = 0 AndAlso txtCUST_ADDR2.Text.Length = 0 Then
                '    ErrorMessage = "Invalid or missing Ship To Street Address"
                '    Return False
                'ElseIf txtCUST_CITY.TextLength = 0 OrElse txtCUST_STATE.TextLength = 0 OrElse txtCUST_ZIP_CODE.TextLength = 0 Then
                '    ErrorMessage = "Invalid or missing Ship To City, State, Zip Code"
                '    Return False
                'ElseIf txtCUST_COUNTRY.TextLength = 0 Then
                '    Dim STATE_CODE As String = txtCUST_STATE.Text
                '    Dim rowTATSTATE As DataRow = tblTATSTATE.Rows.Find(STATE_CODE)
                '    If rowTATSTATE IsNot Nothing Then
                '        txtCUST_COUNTRY.Text = "US"
                '    Else
                '        ErrorMessage = "Invalid or missing Country Code"
                '        Return False
                '    End If
                'End If

                If EMsg.Length = 0 Then
                    Dim rowSOTCARR1 As DataRow = Nothing
                    Dim SHIP_VIA_CODE As String = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text

                    Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
                    If rowSOTSVIA1 IsNot Nothing Then
                        rowSOTCARR1 = LookUp("SOTCARR1", rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty)
                    End If

                    If rowSOTCARR1 IsNot Nothing Then
                        Dim labelFormatDesc As String = "Unknown"
                        For Each vlItem As ValueListItem In optPrint_Type.Items
                            If vlItem.DataValue = rowSOTCARR1.Item("LABEL_FORMAT") & String.Empty Then
                                labelFormatDesc = vlItem.DisplayText
                                Exit For
                            End If
                        Next
                        'If ASCMAIN1.CLIENT <> "VAN" Then
                        If optPrint_Type.Value <> rowSOTCARR1.Item("LABEL_FORMAT") & String.Empty Then
                            Select Case MessageBox.Show("Typically you use " & labelFormatDesc & " to print " & rowSOTCARR1.Item("CARRIER_DESC") & " Labels. Do you want to change to " & labelFormatDesc & " labels?", "Labels", MessageBoxButtons.YesNoCancel)
                                Case Windows.Forms.DialogResult.Yes
                                    optPrint_Type.Value = rowSOTCARR1.Item("LABEL_FORMAT") & String.Empty

                                Case Windows.Forms.DialogResult.No

                                Case Windows.Forms.DialogResult.Cancel
                                    Exit Sub
                            End Select
                        End If
                        'End If
                    End If
                End If

                If EMsg.Length = 0 AndAlso chkUseStoreAddress.Checked Then
                    If MessageBox.Show("You selected to use the Store Address, is this correct?", "Print Label", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Select"
                Load_Record()
                Mode_Settings(True)

            Case "Print Labels"
                RequestAndPrintShipingLabels()

            Case "Done", "Cancel"
                Mode_Settings(False)

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("Select").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("Print Labels").Settings.Enabled = iScreenMode

                '.Items("Void All Shipments").Settings.Enabled = iScreenMode
                .Items("Void All Shipments").Visible = False

            End With
            .Groups("Label Type").Visible = tf
            .Groups("Print Type").Visible = tf
        End With

        splHeader.Visible = tf
        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then

        Else
            Clear_Record()
        End If
        cboZebraPrinter.Visible = True

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTSHIP1", "SOTPICK1", "SOTPICK2",
             "SOTCART1", "SOTCART2", "SOTORDR1", "SOTORDR2", "SOTSHIP0",
             "WHTSHPC1", "WHTSHPC2", "WHTSHPC3", "WHTSHPC5", "WHTSHPCC", "WHTSHPCS", "WHTSHPCP", "ARTCUSTS", "TRACKING"}
            If dst.Tables.Contains(TABLE_NAME) Then
                dst.Tables(TABLE_NAME).Rows.Clear()
            Else
                'Stop
            End If
        Next

        EnforceConstraints(True)
        grp3rdParty.Visible = True

        dteSHIP_DATE_SHIPPED.Value = Nothing
        txt3PAccountNo.Tag = String.Empty
        numTARE.Value = 0.5

        txt3PAccountNo.Clear()
        txt3pCountry.Clear()
        txt3PZipCode.Clear()

        If ASCMAIN1.CLIENT = "VAN" Then
            optLabel_Type.Value = 0
            optPrint_Type.Value = "X"
        Else
            optLabel_Type.Value = 4
            optPrint_Type.Value = "X"
        End If
        optPrint_Status.Value = "N"
        Batched_Group = False
        dteShip.ReadOnly = True
        dteCancel.ReadOnly = True

        Try
            lblStation_ID.Text = "Station ID: " & System.Environment.GetEnvironmentVariable("USERNAME") & String.Empty
            lblPrinter_Comm.Text = "Printer Comm: " & ASCMAIN1.LabelPrinterSerialPort.PortName
        Catch ex As Exception
            'ASCMAIN1.LabelPrinterSerialPort = Nothing
        End Try


        Absx1.txtFor("CUST_CODE").Clear()
        Absx1.txtFor("ORDR_GROUP_NO").Clear()
        Absx1.txtFor("ORDR_CUST_PO").Clear()
        Absx1.txtFor("SHIP_VIA_CODE").Clear()
        Absx1.txtFor("BILLING_TYPE").Clear()
        Absx1.txtFor("PACKAGE_TYPE").Text = "31"
        Absx1.txtFor("PKG_CODE").Clear()
        Absx1.numFor("CARTONL").Value = 0
        Absx1.numFor("CARTONW").Value = 0
        Absx1.numFor("CARTONH").Value = 0

        txtPRE_1.Clear()
        txtPRE_2.Clear()
        txtPRE_3.Clear()
        txtSUFF_1.Clear()
        txtSUFF_2.Clear()
        txtSUFF_3.Clear()
        cbeREF1.Tag = String.Empty

        txtPRE_3.Enabled = True
        cbeREF3.Enabled = True
        txtSUFF_3.Enabled = True

        cbeREF1.Value = String.Empty
        cbeREF2.Value = String.Empty
        cbeREF3.Value = String.Empty
        cbeREF_CODE1.Value = String.Empty
        cbeREF_CODE2.Value = String.Empty
        cbeREF_CODE3.Value = String.Empty

        txtPkg_Type_Desc.Clear()
        txtBilling_Desc.Clear()
        dteShip.Value = Nothing
        dteCancel.Value = Nothing

        CUST_CODE = String.Empty
        ORDR_GROUP_NO = String.Empty
        ORDR_CUST_PO = String.Empty
        Carrier_Code = String.Empty

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)
        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")


        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)

        Dim rowSOTORDR0 As DataRow
        If Batched_Group = True Then
            Absx1.txtFor("ORDR_CUST_PO").Text = "Batched"
            Absx1.txtFor("ORDR_GROUP_NO").Text = "Batched"
            ORDR_GROUP_NO = Mid(Group_No_Selected, 2)
            rowSOTORDR0 = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR0 WHERE ORDR_GROUP_NO IN (" & ORDR_GROUP_NO & ")")
            dteSHIP_DATE_SHIPPED.Value = rowSOTORDR0.Item("ORDR_SHIP_DATE")
            rowSOTORDR0 = Nothing
        Else
            ORDR_GROUP_NO = "'" & Absx1.txtFor("ORDR_GROUP_NO").Text & "'"
            rowSOTORDR0 = LookUp("SOTORDR0", Absx1.txtFor("ORDR_GROUP_NO").Text)
            dteShip.Value = rowSOTORDR0.Item("ORDR_SHIP_DATE")
            dteCancel.Value = rowSOTORDR0.Item("ORDR_CUST_PO")
            Absx1.txtFor("ORDR_CUST_PO").Text = rowSOTORDR0.Item("ORDR_CUST_PO")
            dteSHIP_DATE_SHIPPED.Value = rowSOTORDR0.Item("ORDR_SHIP_DATE")
        End If

        ASCMAIN1.sql = sqlSOTSHIPX & vbCrLf _
        & " and SOTSHIP1.ORDR_GROUP_NO in (" & ORDR_GROUP_NO & ")"
        Fill_Records("SOTSHIP1", "", True, ASCMAIN1.sql)

        Dim shipViaFound As Boolean = False
        Dim rowSOTSVIA1 As DataRow
        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
            rowSOTSHIP1.Item("SELECTED") = "0"

            ' SEE IF WE HAVE A SHIP VIA FOR FEDEX OR UPS
            If Not shipViaFound AndAlso rowARTCUST1.Item("CUST_CODE") <> "BEDBATH" Then
                rowSOTSVIA1 = dst.Tables("SOTSVIA1").Rows.Find(rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty)
                If rowSOTSVIA1 IsNot Nothing Then
                    If rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty = "UPS" OrElse rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty = "FEDEX" Then
                        shipViaFound = True
                        txtSHIP_VIA_CODE.Text = rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty
                    End If
                End If
            End If
        Next

        If rowARTCUST1.Item("CUST_CODE") = "BEDBATH" Then
            txtSHIP_VIA_CODE.Text = "FEDEX"
            txtBilling_Type.Text = "C"
            txtPackageType.Text = "31"
            txtPackageCode.Text = "48F"
        End If

        txt_ValueChanged(txtBilling_Type, Nothing)
        txt_ValueChanged(txtPackageType, Nothing)
        txt_ValueChanged(txtPackageCode, Nothing)

        ASCMAIN1.sql = sqlSOTPICK1 & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
            & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
            & " and SOTSHIP1.ORDR_GROUP_NO in (" & ORDR_GROUP_NO & ")"
        Fill_Records("SOTPICK1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = sqlSOTPICK2 & vbCrLf _
            & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                    & " and SOTSHIP1.ORDR_GROUP_NO in (" & ORDR_GROUP_NO & ")"
        Fill_Records("SOTPICK2", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTCART2.*, SOTCART1.PICK_NO, SOTCART2.QTY_PACKED QTY_PACKED_ORIG" & vbCrLf _
            & " from SOTCART2,SOTCART1,SOTPICK1,SOTSHIP1" & vbCrLf _
            & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
            & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
            & " and SOTSHIP1.ORDR_GROUP_NO in (" & ORDR_GROUP_NO & ")"
        Fill_Records("SOTCART2", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = " Select SOTCART2.CART_NO, Sum(QTY_PACKED * STYLE_WEIGHT) as WEIGHT " _
        & " from SOTCART1, SOTCART2, ICTSTYL1, SOTPICK1, SOTSHIP1 " _
        & "  Where SOTCART2.CART_NO = SOTCART1.CART_NO" _
        & "  AND SOTCART2.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
        & "  And SOTCART1.PICK_NO = SOTPICK1.PICK_NO" _
        & "  And SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" _
        & " and SOTSHIP1.ORDR_GROUP_NO in (" & ORDR_GROUP_NO & ")" _
        & " and SOTSHIP1.SHIP_STATUS IN ('P', 'F')" _
        & " Group by SOTCART2.CART_NO"
        Fill_Records("SOTCARTL", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTCART1.*" & vbCrLf _
            & " from SOTCART1,SOTPICK1,SOTSHIP1" & vbCrLf _
            & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
            & " and SOTSHIP1.ORDR_GROUP_NO in (" & ORDR_GROUP_NO & ")" & vbCrLf _
            & " and SOTSHIP1.SHIP_STATUS IN ('P', 'F')"
        Fill_Records("SOTCART1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT * FROM WHTSHPC1 WHERE SHIP_CNTL_NO IN" & vbCrLf _
            & " (SELECT MAX(WHTSHPC1.SHIP_CNTL_NO) FROM WHTSHPC1, SOTSHIP1" & vbCrLf _
            & " WHERE WHTSHPC1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
            & " AND SOTSHIP1.ORDR_GROUP_NO in (" & ORDR_GROUP_NO & "))"
        Dim wkRow As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
        If wkRow IsNot Nothing Then
            txtSHIP_VIA_CODE.Text = wkRow.Item("SHIP_VIA_CODE") & String.Empty
            Dim SHIP_BOL_NO As String = wkRow.Item("SHIP_BOL_NO")
            Dim SHIP_CNTL_NO As String = wkRow.Item("SHIP_CNTL_NO")

            ASCMAIN1.sql = "SELECT SOTCART1.* FROM SOTCART1, SOTPICK1" & vbCrLf _
                & " WHERE SOTPICK1.PICK_NO = SOTCART1.PICK_NO" & vbCrLf _
                & " AND SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" & vbCrLf _
                & " AND SOTCART1.PACKAGING_TYPE IS NOT NULL"

            wkRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
            If wkRow IsNot Nothing Then
                Absx1.txtFor("PACKAGE_TYPE").Text = wkRow.Item("PACKAGING_TYPE") & String.Empty
                Absx1.txtFor("PKG_CODE").Text = wkRow.Item("PKG_CODE") & String.Empty
            End If

            ASCMAIN1.sql = "SELECT * FROM WHTSHPC3 WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "'"
            wkRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
            If wkRow IsNot Nothing Then
                txt3PAccountNo.Text = wkRow.Item("ACCOUNT_NO_3PL") & String.Empty
                txt3PZipCode.Text = wkRow.Item("ZIP_CODE_3PL") & String.Empty
                txt3pCountry.Text = wkRow.Item("COUNTRY_3PL") & String.Empty
            End If
        End If

        Dim CART_SEQ As Int16 = 0
        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("", "SHIP_BOL_NO,CART_NO")
            CART_SEQ += 1
            rowSOTCART1.Item("CART_SEQ") = CART_SEQ
            rowSOTCART1.Item("STATUS") = IIf(rowSOTCART1.Item("CART_TRACKING_NO") & "" = "", NotPrinted, Printed)
            rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = Val(dst.Tables("SOTCARTL").Select("CART_NO = '" & rowSOTCART1.Item("CART_NO") & "'")(0).Item("WEIGHT") & "") + numTARE.Value
        Next

        Fill_Records("ARTCUSTS", CUST_CODE)

        If dst.Tables("ARTCUSTS").Rows.Count = 1 Then
            Select Case Carrier_Code
                Case "UPS"
                    cbeREF1.Value = dst.Tables("ARTCUSTS").Rows(0).Item("UPS_REF1") & String.Empty
                    cbeREF2.Value = dst.Tables("ARTCUSTS").Rows(0).Item("UPS_REF2") & String.Empty
                    cbeREF3.Value = dst.Tables("ARTCUSTS").Rows(0).Item("UPS_REF3") & String.Empty

                Case "FEDEX"
                    cbeREF1.Value = dst.Tables("ARTCUSTS").Rows(0).Item("FDX_REF1") & String.Empty
                    cbeREF2.Value = dst.Tables("ARTCUSTS").Rows(0).Item("FDX_REF2") & String.Empty
                    cbeREF3.Value = dst.Tables("ARTCUSTS").Rows(0).Item("FDX_REF3") & String.Empty
            End Select
        End If

        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
            SetPrintedStatus(rowSOTSHIP1.Item("SHIP_BOL_NO"))
            rowSOTSHIP1.Item("SELECTED") = "0"
        Next

        Refresh_Refs(0)
        Sort_grdColumns(grdSOTCART1, "SHIP_BOL_NO,CART_NO")
        Sort_grdColumns(grdSOTSHIP1, "SHIP_BOL_NO")
        Refresh_Printed()
    End Sub

    Sub Refresh_Refs(ByVal ReferenceNumber As Int16)

        Dim REFERENCE1 As String = String.Empty
        Dim REFERENCE2 As String = String.Empty
        Dim REFERENCE3 As String = String.Empty
        Dim REF_CODE As String = String.Empty
        Dim rowSOTCARRR As DataRow = Nothing
        Dim REF_CODE_LBL As String

        If Carrier_Code = "FEDEX" OrElse Carrier_Code = "UPS" Then
            If cbeREF1.Value <> String.Empty AndAlso cbeREF1.Value <> "None" AndAlso cbeREF_CODE1.Value <> Nothing AndAlso (ReferenceNumber = 0 OrElse ReferenceNumber = 1) Then
                REF_CODE = cbeREF1.SelectedItem.DataValue
                REF_CODE_LBL = cbeREF_CODE1.SelectedItem.DataValue & ""
                rowSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {Carrier_Code, REF_CODE})
                If rowSOTCARRR IsNot Nothing Then
                    REFERENCE1 = REF_CODE_LBL.Substring(0, 2) & ":" & rowSOTCARRR.Item("TABLE_NAME") & "." & rowSOTCARRR.Item("COLUMN_NAME")
                End If
            End If

            If cbeREF2.Value <> String.Empty AndAlso cbeREF2.Value <> "None" AndAlso cbeREF_CODE2.Value <> Nothing AndAlso (ReferenceNumber = 0 OrElse ReferenceNumber = 2) Then
                REF_CODE = cbeREF2.SelectedItem.DataValue
                REF_CODE_LBL = cbeREF_CODE2.SelectedItem.DataValue
                rowSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {Carrier_Code, REF_CODE})
                If rowSOTCARRR IsNot Nothing Then
                    REFERENCE2 = REF_CODE_LBL.Substring(0, 2) & ":" & rowSOTCARRR.Item("TABLE_NAME") & "." & rowSOTCARRR.Item("COLUMN_NAME")
                End If
            End If

            If cbeREF3.Value <> String.Empty AndAlso cbeREF3.Value <> "None" AndAlso cbeREF_CODE3.Value <> Nothing AndAlso (ReferenceNumber = 0 OrElse ReferenceNumber = 3) Then
                REF_CODE = cbeREF3.SelectedItem.DataValue
                REF_CODE_LBL = cbeREF_CODE1.SelectedItem.DataValue
                rowSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {Carrier_Code, REF_CODE})
                If rowSOTCARRR IsNot Nothing Then
                    REFERENCE3 = REF_CODE_LBL.Substring(0, 2) & ":" & rowSOTCARRR.Item("TABLE_NAME") & "." & rowSOTCARRR.Item("COLUMN_NAME")
                End If
            End If
        End If

        Dim temp1 As String = String.Empty
        Dim temp2 As String = String.Empty
        Dim temp3 As String = String.Empty

        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("", "SHIP_BOL_NO,CART_NO")

            temp1 = String.Empty
            temp2 = String.Empty
            temp3 = String.Empty

            If REFERENCE1.Length > 0 AndAlso (ReferenceNumber = 0 OrElse ReferenceNumber = 1) Then
                temp1 = GetReferenceValue(txtPRE_1.Text, txtSUFF_1.Text, REFERENCE1.Split(":")(1), rowSOTCART1.Item("CART_NO"))
                If temp1.Length > 0 Then
                    temp1 = REFERENCE1.Split(":")(0).Substring(0, 2) & ":" & temp1
                End If
            End If

            If REFERENCE2.Length > 0 AndAlso (ReferenceNumber = 0 OrElse ReferenceNumber = 2) Then
                temp2 = GetReferenceValue(txtPRE_2.Text, txtSUFF_2.Text, REFERENCE2.Split(":")(1), rowSOTCART1.Item("CART_NO"))
                If temp2.Length > 0 Then
                    temp2 = REFERENCE2.Split(":")(0).Substring(0, 2) & ":" & temp2
                End If
            End If

            If REFERENCE3.Length > 0 AndAlso (ReferenceNumber = 0 OrElse ReferenceNumber = 3) Then
                temp3 = GetReferenceValue(txtPRE_3.Text, txtSUFF_3.Text, REFERENCE3.Split(":")(1), rowSOTCART1.Item("CART_NO"))
                If temp3.Length > 0 Then
                    temp3 = REFERENCE3.Split(":")(0).Substring(0, 2) & ":" & temp3
                End If
            End If

            If Carrier_Code = "FEDEX" AndAlso rowSOTCARRR IsNot Nothing Then
                If temp1 <> "" Then
                    If cbeREF_CODE1.Value Is Nothing Then
                        cbeREF_CODE1.Value = temp1.Substring(0, 2)
                    End If
                    temp1 = Replace(temp1, temp1.Substring(0, 3), Mid(cbeREF_CODE1.Value, 1, 2) & ":")
                End If
                If temp2 <> "" Then
                    If cbeREF_CODE2.Value Is Nothing Then
                        cbeREF_CODE2.Value = temp2.Substring(0, 2)
                    End If
                    temp2 = Replace(temp2, temp2.Substring(0, 3), Mid(cbeREF_CODE2.Value, 1, 2) & ":")
                End If
                If temp3 <> "" Then
                    If cbeREF_CODE3.Value Is Nothing Then
                        cbeREF_CODE3.Value = temp3.Substring(0, 2)
                    End If
                    temp3 = Replace(temp3, temp3.Substring(0, 3), Mid(cbeREF_CODE3.Value, 1, 2) & ":")
                End If



                If temp1.StartsWith("ST:") Then
                    temp1 = temp1.Replace("ST:", "CR:")
                End If

                If temp2.StartsWith("ST:") Then
                    temp2 = temp2.Replace("ST:", "CR:")
                End If

                If temp3.StartsWith("ST:") Then
                    temp3 = temp3.Replace("ST:", "CR:")
                End If


            End If

            If (ReferenceNumber = 0 OrElse ReferenceNumber = 1) Then rowSOTCART1.Item("REFERENCE1") = temp1
            If (ReferenceNumber = 0 OrElse ReferenceNumber = 2) Then rowSOTCART1.Item("REFERENCE2") = temp2
            If (ReferenceNumber = 0 OrElse ReferenceNumber = 3) Then rowSOTCART1.Item("REFERENCE3") = temp3
        Next

    End Sub

    Private Function GetReferenceValue(ByVal Prefix As String, ByVal Suffix As String, ByVal field As String, ByVal CART_NO As String) As String

        Dim TABLE_NAME As String = String.Empty
        Dim COLUMN_NAME As String = String.Empty
        Dim dataRow As DataRow = Nothing
        Dim referenceValue As String = String.Empty

        Try
            TABLE_NAME = field.Split(".")(0)
            COLUMN_NAME = field.Split(".")(1)

            Select Case TABLE_NAME
                Case "SOTCART1"
                    dataRow = dst.Tables("SOTCART1").Rows.Find(CART_NO)

                Case "SOTPICK1"
                    dataRow = dst.Tables("SOTCART1").Rows.Find(CART_NO)
                    If dataRow Is Nothing Then
                        Exit Select
                    End If
                    dataRow = dst.Tables("SOTPICK1").Rows.Find(dataRow.Item("PICK_NO") & String.Empty)

                Case "SOTORDR1"
                    dataRow = dst.Tables("SOTCART1").Rows.Find(CART_NO)
                    If dataRow Is Nothing Then
                        Exit Select
                    End If
                    dataRow = dst.Tables("SOTPICK1").Rows.Find(dataRow.Item("PICK_NO") & String.Empty)
                    If dataRow Is Nothing Then
                        Exit Select
                    End If
                    dataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR1 WHERE ORDR_NO = '" & dataRow.Item("ORDR_NO") & "'")

                Case "SOTSHIP1"
                    dataRow = dst.Tables("SOTCART1").Rows.Find(CART_NO)
                    If dataRow Is Nothing Then
                        Exit Select
                    End If
                    dataRow = dst.Tables("SOTPICK1").Rows.Find(dataRow.Item("PICK_NO") & String.Empty)
                    If dataRow Is Nothing Then
                        Exit Select
                    End If
                    dataRow = dst.Tables("SOTSHIP1").Rows.Find(dataRow.Item("SHIP_BOL_NO") & String.Empty)

                Case String.Empty
                    If COLUMN_NAME = String.Empty Then
                        referenceValue = Prefix & Suffix
                        Return referenceValue.Trim
                    End If

            End Select

            If dataRow Is Nothing Then
                Return String.Empty
            End If

            If dataRow.Item(COLUMN_NAME) & String.Empty = String.Empty Then
                Return String.Empty
            End If

            referenceValue = Prefix & dataRow.Item(COLUMN_NAME) & String.Empty & Suffix
            referenceValue = referenceValue.Trim
            Return referenceValue


        Catch ex As Exception
            Return String.Empty
        End Try

    End Function

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control,
         ByVal COLUMN_NAME As String,
         Optional ByRef sql_where As String = "",
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "SHIP_VIA_CODE"
                sql_where = "CARRIER_CODE  in ('FEDEX','UPS') And CARRIER_PROD_CODE is not null"
            Case "PACKAGE_TYPE"

                Dim SHIP_VIA_CODE As String = Absx1.txtFor("SHIP_VIA_CODE").Text.Trim
                If SHIP_VIA_CODE = "" Then
                    MsgBox("Please Select Ship Via", MsgBoxStyle.OkOnly, "None Selected")
                    Cancel = True
                    Exit Sub
                Else
                    sql_where = " CARRIER_CODE in (Select Distinct CARRIER_CODE from SOTSVIA1 " _
                    & " Where SOTSVIA1.SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "')"
                End If
            Case "ORDR_GROUP_NO"

            Case "BILLING_TYPE"
                Dim SHIP_VIA_CODE As String = Absx1.txtFor("SHIP_VIA_CODE").Text.Trim
                If SHIP_VIA_CODE = "" Then
                    MsgBox("Please Select Ship Via", MsgBoxStyle.OkOnly, "None Selected")
                    Cancel = True
                    Exit Sub
                Else
                    sql_where = " SHIPPER_CODE = '" & IIf(Carrier_Code = "FEDEX", "FDX", "UPS") & "' and SHIP_PARM_TYPE = 'BILLING'"
                End If
        End Select
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "Load", "Edit"

        End Select

        Return return_key
    End Function


#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSHIP1, "BBBPBB", "De-Select All", "Select These", "Un-Select These", "Reprint Labels", "Void Shipment")
        Load_Popup_Menu(grdSOTCART1, "B", "Void Carton")
        Load_Popup_Menu(grdTracking, "SSPB", "Show Filter", "Show GroupBox", "Get Tracking Data")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu OrElse e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdSOTPICK1"
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '  e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdSOTSHIP1"
                    Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value & String.Empty
                    Dim CUST_STORE_NO As String = grd.ActiveRow.Cells("CUST_STORE_NO").Value & String.Empty
                    tlb_pop.Tools("Void Shipment").SharedProps.Caption = $"Void Shipments for {SHIP_BOL_NO}"
                    tlb_pop.Tools("Reprint Labels").SharedProps.Caption = $"Reprint Label(s) for {SHIP_BOL_NO}"

                    If dst.Tables("SOTCART1").Select($"SHIP_BOL_NO = '{SHIP_BOL_NO}' And CUST_STORE_NO = '{CUST_STORE_NO}' and ISNULL(CART_TRACKING_NO, '') <> ''").Length > 0 Then
                        tlb_pop.Tools("Void Shipment").SharedProps.Enabled = True
                        tlb_pop.Tools("Reprint Labels").SharedProps.Enabled = True
                    Else
                        tlb_pop.Tools("Void Shipment").SharedProps.Enabled = False
                        tlb_pop.Tools("Reprint Labels").SharedProps.Enabled = False
                    End If

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.OwningMenu IsNot Nothing Then
            grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        End If

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

            Case "Select All", "De-Select All"
                For Each rowSOTPICK1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                    rowSOTPICK1.Item("SELECTED") = IIf(e.Tool.Key = "Select All", "1", "0")
                Next

            Case "Select These", "Un-Select These"
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTSHIP1.Selected.Rows

                    Dim row As DataRow = dst.Tables("SOTSHIP1").Rows.Find(New Object() {grow.Cells("SHIP_BOL_NO").Value})
                    row.Item("SELECTED") = IIf(e.Tool.Key = "Select These", "1", "0")
                Next
                grdSOTSHIP1.Selected.Rows.Clear()

            Case "Void Shipment"
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value & String.Empty
                Dim CUST_STORE_NO As String = grd.ActiveRow.Cells("CUST_STORE_NO").Value & String.Empty

                If SHIP_BOL_NO.Length = 0 Then
                    MessageBox.Show($"Cannot determine the selected Shipment.", "Void Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If MessageBox.Show($"Do you want to Void shipping label(s) for Shipment: {SHIP_BOL_NO}?", "Void Shipment", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

                If dst.Tables("SOTCART1").Select($"SHIP_BOL_NO = '{SHIP_BOL_NO}' And CUST_STORE_NO = '{CUST_STORE_NO}' and ISNULL(CART_TRACKING_NO, '**') <> '**'").Length > 0 Then
                    VoidShippingLabels(SHIP_BOL_NO, CUST_STORE_NO, String.Empty)
                End If

                For Each rowSOTCART1 As DataRow In dst.Tables($"SOTCART1").Select($"SHIP_BOL_NO = '{SHIP_BOL_NO}' And CUST_STORE_NO = '{CUST_STORE_NO}'", "CART_NO")
                    ASCMAIN1.sql = "Update SOTCART1 Set CART_TRACKING_NO = NULL Where CART_NO = '" & rowSOTCART1.Item("CART_NO") & "'"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                    'multi-po carton table - no harm is missing
                    ASCMAIN1.sql = "Update SOTCARM1 Set CART_TRACKING_NO = NULL Where CART_NO = '" & rowSOTCART1.Item("CART_NO") & "'"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                    dst.Tables("SOTCART1").Select("CART_NO = '" & rowSOTCART1.Item("CART_NO") & "'", "")(0).Item("CART_TRACKING_NO") = ""
                Next

                dst.Tables("SOTSHIP1").Select("SHIP_BOL_NO = '" & grd.ActiveRow.Cells("SHIP_BOL_NO").Value & "'", "")(0).Item("SELECTED") = "0"
                SetPrintedStatus(grd.ActiveRow.Cells("SHIP_BOL_NO").Value)

            Case "Reprint Labels"
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value & String.Empty
                Dim CUST_STORE_NO As String = grd.ActiveRow.Cells("CUST_STORE_NO").Value & String.Empty

                If SHIP_BOL_NO.Length = 0 Then
                    MessageBox.Show($"Cannot determine the selected Shipment.", "Void Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If dst.Tables("SOTCART1").Select($"SHIP_BOL_NO = '{SHIP_BOL_NO}' And CUST_STORE_NO = '{CUST_STORE_NO}'").Length = 0 Then
                    MessageBox.Show("The selected shipment does not have any cartons with tracking numbers.", "Reprint Labels", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                If MessageBox.Show($"Do you want to Reprint shipping label(s) for Shipment: {SHIP_BOL_NO}?", "Void Shipment", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

                RePrintShippingLabels(SHIP_BOL_NO, CUST_STORE_NO)

            Case "Void Carton"
                Dim CART_NO As String = grd.ActiveRow.Cells("CART_NO").Value

                If dst.Tables("SOTCART1").Select($"CART_NO = '{CART_NO}'").Length = 0 Then
                    MessageBox.Show($"Cannot determine the selected Carton.", "Void Carton", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If MessageBox.Show($"Do you want to Void shipping labels for Carton: {CART_NO}?", "Void Carton", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

                Dim CART_TRACKING_NO As String = dst.Tables("SOTCART1").Select($"CART_NO = '{CART_NO}'")(0).Item("CART_TRACKING_NO") & String.Empty
                If CART_TRACKING_NO.Length > 0 Then
                    VoidShippingLabels(String.Empty, String.Empty, CART_NO)
                End If

                ASCMAIN1.sql = "Update SOTCART1 Set CART_TRACKING_NO = NULL Where CART_NO = :PARM1"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {CART_NO})
                ' multi-po carton
                ASCMAIN1.sql = "Update SOTCARM1 Set CART_TRACKING_NO = NULL Where CART_NO = :PARM1"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {CART_NO})
                dst.Tables("SOTCART1").Select("CART_NO = '" & grd.ActiveRow.Cells("CART_NO").Value & "'", "")(0).Item("CART_TRACKING_NO") = ""

        End Select

        If grd Is Nothing Then
            Exit Sub
        Else
            If (grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow) AndAlso grd.Name <> grdTracking.Name Then
                Exit Sub
            End If
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                'Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                'Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")

            Case "Get Tracking Data"
                ' Get all request labels for the selected Order Group. Do not include voided labels

                If MessageBox.Show("Do you want to get tracking data?", "Get Tracking Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                    Exit Sub
                End If

                Dim LAST_STATUS As String = String.Empty
                dst.Tables("TRACKING").Rows.Clear()
                ASCMAIN1.sql = $"SELECT SOTSHIP1.SHIP_BOL_NO SHIP_BOL_NO_SH, WHTSHPC1.*, WHTSHPC2.TRACKING_NO, WHTSHPC2.CART_NO, SOTSHIP1.SHIP_ADDR_CODE
                                    FROM SOTSHIP1, WHTSHPC1, WHTSHPC2
                                    WHERE SOTSHIP1.SHIP_BOL_NO = WHTSHPC1.SHIP_BOL_NO (+)
                                    AND WHTSHPC1.SHIP_CNTL_NO = WHTSHPC2.SHIP_CNTL_NO (+)
                                    AND NVL(WHTSHPC1.STATUS, 'P') IN ('V', 'P')
                                    AND SOTSHIP1.ORDR_GROUP_NO in ({ORDR_GROUP_NO})"

                Dim tblData As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

                If tblData.Rows.Count = 0 Then
                    MessageBox.Show("No Tracking Numbers to process", "Get Tracking Data", MessageBoxButtons.OK)
                    Exit Sub
                End If

                Try

                    Me.Cursor = Cursors.WaitCursor
                    Me.Enabled = False

                    ' Load the cartons. The local copy of SOTCART1 may not have all cartons if caton labels were requested multiple times.
                    ASCMAIN1.sql = $" SELECT SOTCART1.*, SOTPICK1.ORDR_NO 
                                    FROM SOTCART1, SOTPICK1
                                    WHERE SOTCART1.PICK_NO = SOTPICK1.PICK_NO (+)
                                    AND CART_NO IN
                                    (
                                        SELECT WHTSHPC2.CART_NO
                                        FROM SOTSHIP1, WHTSHPC1, WHTSHPC2
                                        WHERE SOTSHIP1.SHIP_BOL_NO = WHTSHPC1.SHIP_BOL_NO (+)
                                        AND WHTSHPC1.SHIP_CNTL_NO = WHTSHPC2.SHIP_CNTL_NO (+)
                                        AND NVL(WHTSHPC1.STATUS, 'P') IN ('V', 'P')
                                        AND SOTSHIP1.ORDR_GROUP_NO in ({ORDR_GROUP_NO})
                                    )"
                    Dim tblSOTCART1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
                    tblSOTCART1.PrimaryKey = New DataColumn() {tblSOTCART1.Columns("cart_no")}

                    ASCMAIN1.sql = "SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM2"
                    Dim tblARTCUST2 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST2", "V", New Object() {HFs("CUST_CODE")})

                    For Each rowData As DataRow In tblData.Select("", "SHIP_BOL_NO")
                        Me.Cursor = Cursors.WaitCursor

                        Dim MASTER_TRACKING_NO As String = rowData.Item("MASTER_TRACKING_NO") & String.Empty
                        Dim TRACKING_NO As String = rowData.Item("TRACKING_NO") & String.Empty
                        Dim SHIP_BOL_NO As String = rowData.Item("SHIP_BOL_NO") & String.Empty

                        If MASTER_TRACKING_NO.Length = 0 AndAlso TRACKING_NO.Length = 0 Then
                            Continue For
                        End If

                        If TRACKING_NO.Length = 0 AndAlso MASTER_TRACKING_NO.Length > 0 Then
                            TRACKING_NO = MASTER_TRACKING_NO
                            rowData.Item("TRACKING_NO") = TRACKING_NO
                        End If

                        ' Need to use the Tracking number to get the tracking information
                        Dim TrackingData As New TAC.WHCSHIP1.TrackingData
                        RequestTrackingInformation(rowData, TrackingData)
                        Me.Cursor = Cursors.WaitCursor

                        Dim rowTRACKING As DataRow = dst.Tables("TRACKING").NewRow
                        rowTRACKING.Item("TRACKING_NO") = TRACKING_NO
                        rowTRACKING.Item("NOT_ASSIGNED") = "0"
                        rowTRACKING.Item("NOT_SCANNED") = "0"
                        rowTRACKING.Item("ZIPCODE_NO_MATCH") = "0"
                        rowTRACKING.Item("SHIPMENT_CARTONS") = Val(dst.Tables("SOTCART1").Compute("COUNT(CART_NO)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}'") & String.Empty)

                        ' Tracking data returned from Shipper
                        rowTRACKING.Item("STATUS") = TrackingData.Status & String.Empty
                        rowTRACKING.Item("Date") = TrackingData.Date & String.Empty
                        rowTRACKING.Item("Time") = TrackingData.Time & String.Empty
                        rowTRACKING.Item("City") = TrackingData.City & String.Empty
                        rowTRACKING.Item("State") = TrackingData.State & String.Empty
                        rowTRACKING.Item("CountryCode") = TrackingData.CountryCode & String.Empty
                        rowTRACKING.Item("Location") = TrackingData.Location & String.Empty
                        rowTRACKING.Item("Address1") = TrackingData.Address1 & String.Empty
                        rowTRACKING.Item("Address2") = TrackingData.Address2 & String.Empty
                        rowTRACKING.Item("Zipcode") = TrackingData.ZipCode & String.Empty
                        rowTRACKING.Item("SHIPMENT_STATUS") = rowData.Item("STATUS")

                        Dim CUST_ADDR_CODE As String = rowData.Item("SHIP_ADDR_CODE") & String.Empty
                        Dim SHIP_TO_ZIPCODE As String = String.Empty
                        If tblARTCUST2.Select($"CUST_ADDR_CODE = '{CUST_ADDR_CODE}'").Length > 0 Then
                            SHIP_TO_ZIPCODE = tblARTCUST2.Select($"CUST_ADDR_CODE = '{CUST_ADDR_CODE}'")(0).Item("CUST_ZIP_CODE") & String.Empty
                        End If
                        rowTRACKING.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                        rowTRACKING.Item("SHIP_TO_ZIPCODE") = SHIP_TO_ZIPCODE

                        ' See if we have this tracking number in the list of cartons.
                        ' Sometimes they request new labels; therefore, for one carton they may have request several labels but the carton will have the last label Tracking No assigned to it
                        If dst.Tables("SOTCART1").Select($"CART_TRACKING_NO = '{TRACKING_NO}'").Length > 0 Then
                            Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").Select($"CART_TRACKING_NO = '{TRACKING_NO}'")(0)

                            rowTRACKING.Item("CART_NO") = rowSOTCART1.Item("CART_NO")
                            rowTRACKING.Item("SHIP_BOL_NO") = rowSOTCART1.Item("SHIP_BOL_NO")
                            rowTRACKING.Item("PICK_NO") = rowSOTCART1.Item("PICK_NO")

                            If dst.Tables("SOTPICK1").Select($"PICK_NO = '{rowSOTCART1.Item("PICK_NO")}'").Length > 0 Then
                                rowTRACKING.Item("ORDR_NO") = dst.Tables("SOTPICK1").Select($"PICK_NO = '{rowSOTCART1.Item("PICK_NO")}'")(0).Item("ORDR_NO")
                            End If
                        Else
                            rowTRACKING.Item("CART_NO") = "Unknown"
                            rowTRACKING.Item("SHIP_BOL_NO") = rowData.Item("SHIP_BOL_NO_SH")
                            rowTRACKING.Item("NOT_ASSIGNED") = "1"

                            Dim rowSOTCART1 As DataRow = tblSOTCART1.Rows.Find(rowData.Item("CART_NO") & String.Empty)
                            If rowSOTCART1 IsNot Nothing Then
                                rowTRACKING.Item("CART_NO") = rowSOTCART1.Item("CART_NO")
                                'rowTRACKING.Item("SHIP_BOL_NO") = rowData.Item("SHIP_BOL_NO")
                                rowTRACKING.Item("PICK_NO") = rowSOTCART1.Item("PICK_NO")
                                rowTRACKING.Item("ORDR_NO") = rowSOTCART1.Item("ORDR_NO")
                            Else
                                ASCMAIN1.sql = $"SELECT * FROM WHTSHPC2 WHERE CART_NO = :PARM1 AND TRACKING_NO = :PARM2"
                                Dim rowWHTSHPC2 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {rowData.Item("CART_NO") & String.Empty, TRACKING_NO})
                                If rowWHTSHPC2 IsNot Nothing Then
                                    rowTRACKING.Item("CART_NO") = rowWHTSHPC2.Item("CART_NO")
                                    If rowTRACKING.Item("CART_NO") & String.Empty = String.Empty Then
                                        rowTRACKING.Item("CART_NO") = "Unknown"
                                    End If
                                End If
                            End If
                        End If
                        dst.Tables("TRACKING").Rows.Add(rowTRACKING)

                        If rowTRACKING.Item("ZIPCODE") & String.Empty <> String.Empty Then
                            If rowTRACKING.Item("ZIPCODE") & String.Empty <> rowTRACKING.Item("SHIP_TO_ZIPCODE") & String.Empty Then
                                rowTRACKING.Item("ZIPCODE_NO_MATCH") = "1"
                            End If
                        End If
                    Next

                    ' Need to see if any Cartons in SOTCART1 are not in the tracking table.
                    For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
                        Dim CART_TRACKING_NO As String = rowSOTCART1.Item("CART_TRACKING_NO") & String.Empty
                        If CART_TRACKING_NO.Length = 0 Then
                            Continue For
                        End If

                        If dst.Tables("TRACKING").Select($"CART_NO = '{rowSOTCART1.Item("CART_NO")}'").Length = 0 Then
                            Dim rowTRACKING As DataRow = dst.Tables("TRACKING").NewRow
                            rowTRACKING.Item("CART_NO") = rowSOTCART1.Item("CART_NO")
                            rowTRACKING.Item("SHIP_BOL_NO") = rowSOTCART1.Item("SHIP_BOL_NO")
                            rowTRACKING.Item("TRACKING_NO") = rowSOTCART1.Item("TRACKING_NO")
                            rowTRACKING.Item("NOT_ASSIGNED") = "1"
                            rowTRACKING.Item("STATUS") = "Not Found"
                            rowTRACKING.Item("PICK_NO") = rowSOTCART1.Item("PICK_NO")

                            If dst.Tables("SOTPICK1").Select($"PICK_NO = '{rowSOTCART1.Item("PICK_NO")}'").Length > 0 Then
                                rowTRACKING.Item("ORDR_NO") = dst.Tables("SOTPICK1").Select($"PICK_NO = '{rowSOTCART1.Item("PICK_NO")}'")(0).Item("ORDR_NO")
                            End If

                            dst.Tables("TRACKING").Rows.Add(rowTRACKING)
                        End If
                    Next

                    ' See if any cartons were not scanned by the Shipper.
                    Select Case tblData.Rows(0).Item("CARRIER_CODE") & String.Empty
                        Case "UPS"
                            LAST_STATUS = "Order Processed: Ready for UPS"
                        Case "FEDEX"
                            LAST_STATUS = "Shipment information sent to FedEx"
                        Case Else
                            LAST_STATUS = "WHO KNOWS"
                    End Select

                    For Each row As DataRow In dst.Tables("TRACKING").Select($"STATUS = '{LAST_STATUS}'")
                        Dim CART_NO As String = row.Item("CART_NO")

                        If dst.Tables("TRACKING").Select($"CART_NO = '{CART_NO}' AND STATUS <> '{LAST_STATUS}'").Length = 0 Then
                            row.Item("NOT_SCANNED") = "1"
                        End If
                    Next

                    Dim tbl As DataTable = ASCDATA1.SelectDistinct(dst.Tables("TRACKING"), "CART_NO")
                    For Each row As DataRow In tbl.Select
                        Dim CART_NO As String = row.Item("CART_NO") & String.Empty
                        Dim SHIPMENT_LABELS As Int16 = dst.Tables("TRACKING").Compute("COUNT(TRACKING_NO)", $"CART_NO = '{CART_NO}'")

                        For Each rowTracking As DataRow In dst.Tables("TRACKING").Select($"CART_NO = '{CART_NO}'")
                            rowTracking.Item("SHIPMENT_LABELS") = SHIPMENT_LABELS
                        Next
                    Next

                    Sort_grdColumns(grdTracking, "CUST_ADDR_CODE,SHIP_BOL_NO,CART_NO")
                    grdTracking.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.VisibleRows, True)

                    MessageBox.Show("Tracking Complete", "Get Tracking Data", MessageBoxButtons.OK)

                Catch ex As Exception
                    MessageBox.Show($"Error Generating Tracking Data: {ex.Message}", "Get Tracking Data", MessageBoxButtons.OK)
                Finally
                    Me.Cursor = Cursors.Default
                    Me.Enabled = True
                End Try

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then

                End If

            Case "ORDR_GROUP_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Absx1.txtFor("ORDR_GROUP_NO").Text = Absx1.txtFor("ORDR_GROUP_NO").Text.PadLeft(10, "0")
                    Click_Command("Select")

                End If

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "CUST_CODE"
                If Not ScreenMode Then
                End If

            Case "ORDR_GROUP_NO"
                MyBase.Absx1.txtFor("ORDR_GROUP_NO").Text = MyBase.Absx1.txtFor("ORDR_GROUP_NO").Text.Trim
                If MyBase.Absx1.txtFor("ORDR_GROUP_NO").TextLength > 0 Then
                    MyBase.Absx1.txtFor("ORDR_GROUP_NO").Text = ASCMAIN1.Format_Field(MyBase.Absx1.txtFor("ORDR_GROUP_NO").Text, "ORDR_GROUP_NO")
                End If

        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"

        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub

    Public Overrides Sub dte_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.dte_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "SHIP_DATE_SHIPPED"

        End Select
    End Sub
#End Region

#Region "grdSOTPICK1"

    Private Sub grdSOTPICK1_BeforeRowActivate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTSHIP1.BeforeRowActivate
        If grdSOTCART1.ActiveRow IsNot Nothing AndAlso grdSOTCART1.ActiveRow.DataChanged Then
            grdSOTCART1.ActiveRow.Update()
        End If
    End Sub




    Private Sub grdSOTSHIP1_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles grdSOTSHIP1.MouseUp
        If grdSOTSHIP1.ActiveCell IsNot Nothing AndAlso grdSOTSHIP1.ActiveCell.Column.Key = "SELECTED" Then
            'grdSOTSHIP1.ActiveRow.Update()
        End If
    End Sub

#End Region

#Region "grdSOTCART1"

    'Private Sub grdSOTCART1_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCART1.AfterCellUpdate

    '    Dim displayBoxAttributes As Boolean = False

    '    For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTCART1.Rows
    '        If row.Cells("PKG_CODE").Text = "OTHER" Then
    '            displayBoxAttributes = True
    '            Exit For
    '        End If
    '    Next

    '    If displayBoxAttributes Then
    '        grdSOTCART1.DisplayLayout.Bands(0).Columns("WIDTH").Hidden = False
    '        grdSOTCART1.DisplayLayout.Bands(0).Columns("LENGTH").Hidden = False
    '        grdSOTCART1.DisplayLayout.Bands(0).Columns("HEIGHT").Hidden = False

    '        grdSOTCART1.DisplayLayout.Bands(0).Columns("WIDTH").CellActivation = UltraWinGrid.Activation.AllowEdit
    '        grdSOTCART1.DisplayLayout.Bands(0).Columns("LENGTH").CellActivation = UltraWinGrid.Activation.AllowEdit
    '        grdSOTCART1.DisplayLayout.Bands(0).Columns("HEIGHT").CellActivation = UltraWinGrid.Activation.AllowEdit
    '    Else
    '        grdSOTCART1.DisplayLayout.Bands(0).Columns("WIDTH").Hidden = True
    '        grdSOTCART1.DisplayLayout.Bands(0).Columns("LENGTH").Hidden = True
    '        grdSOTCART1.DisplayLayout.Bands(0).Columns("HEIGHT").Hidden = True

    '        grdSOTCART1.DisplayLayout.Bands(0).Columns("WIDTH").CellActivation = UltraWinGrid.Activation.NoEdit
    '        grdSOTCART1.DisplayLayout.Bands(0).Columns("LENGTH").CellActivation = UltraWinGrid.Activation.NoEdit
    '        grdSOTCART1.DisplayLayout.Bands(0).Columns("HEIGHT").CellActivation = UltraWinGrid.Activation.NoEdit
    '    End If
    'End Sub


    Private Sub grdSOTCART1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCART1.ClickCellButton
        '    If e.Cell.Column.Key = "CART_TOTAL_WGT_ACTUAL" Then
        '        registeredWeight = 0
        '        RequestWeightFromScale()
        '        e.Cell.Value = registeredWeight
        '    End If
    End Sub

#End Region

#Region "Form Controls"

    Private Sub txtSHIP_VIA_CODE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtSHIP_VIA_CODE.ValueChanged

        If rowARTCUST1 Is Nothing OrElse txtSHIP_VIA_CODE.Text.Trim.Length = 0 Then
            Exit Sub
        Else
            Dim SHIP_VIA_CODE As String = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text.Trim
            Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
            Dim rowSOTCARR1 As DataRow = Nothing
            If rowSOTSVIA1 Is Nothing Then Exit Sub

            If rowSOTSVIA1 IsNot Nothing Then
                rowSOTCARR1 = LookUp("SOTCARR1", rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty)
                If rowSOTCARR1 Is Nothing Then
                    Exit Sub
                End If
            End If

            Carrier_Code = rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty

            Select Case Carrier_Code
                Case "UPS"
                    txtPRE_3.Enabled = False
                    cbeREF3.Enabled = False
                    txtSUFF_3.Enabled = False

                    txtPRE_3.Clear()
                    txtSUFF_3.Clear()
                    txtSUFF_3.Text = String.Empty

                Case Else
                    txtPRE_3.Enabled = True
                    cbeREF3.Enabled = True
                    txtSUFF_3.Enabled = True
            End Select

            If cbeREF1.Tag <> Carrier_Code Then
                cbeREF1.Tag = Carrier_Code

                cbeREF1.Value = "N"
                cbeREF2.Value = "N"
                cbeREF3.Value = "N"

                ' Setup Label Data
                cbeREF1.DataSource = ASCDATA1.GetDataTable("SELECT REF_CODE, REF_DESC from SOTCARRR WHERE CARRIER_CODE = '" & Carrier_Code & "' Union Select 'N', 'None' from dual")
                cbeREF1.ValueMember = "REF_CODE"
                cbeREF1.DisplayMember = "REF_DESC"

                cbeREF2.DataSource = ASCDATA1.GetDataTable("SELECT REF_CODE, REF_DESC from SOTCARRR WHERE CARRIER_CODE = '" & Carrier_Code & "' Union Select 'N', 'None' from dual")
                cbeREF2.ValueMember = "REF_CODE"
                cbeREF2.DisplayMember = "REF_DESC"

                cbeREF3.DataSource = ASCDATA1.GetDataTable("SELECT REF_CODE, REF_DESC from SOTCARRR WHERE CARRIER_CODE = '" & Carrier_Code & "' Union Select 'N', 'None' from dual")
                cbeREF3.ValueMember = "REF_CODE"
                cbeREF3.DisplayMember = "REF_DESC"

                cbeREF_CODE1.DataSource = ASCDATA1.GetDataTable("SELECT REF_CODE, REF_DESC from SOTCARRR WHERE Length(REF_CODE) = 2 and CARRIER_CODE = '" & Carrier_Code & "'")
                cbeREF_CODE1.ValueMember = "REF_CODE"
                cbeREF_CODE1.DisplayMember = "REF_CODE"

                cbeREF_CODE2.DataSource = ASCDATA1.GetDataTable("SELECT REF_CODE, REF_DESC from SOTCARRR WHERE Length(REF_CODE) = 2 and CARRIER_CODE = '" & Carrier_Code & "'")
                cbeREF_CODE2.ValueMember = "REF_CODE"
                cbeREF_CODE2.DisplayMember = "REF_CODE"

                cbeREF_CODE3.DataSource = ASCDATA1.GetDataTable("SELECT REF_CODE, REF_DESC from SOTCARRR WHERE Length(REF_CODE) = 2 and CARRIER_CODE = '" & Carrier_Code & "'")
                cbeREF_CODE3.ValueMember = "REF_CODE"
                cbeREF_CODE3.DisplayMember = "REF_CODE"

                If dst.Tables("ARTCUSTS").Rows.Count = 1 Then
                    Select Case Carrier_Code
                        Case "UPS"
                            cbeREF1.Value = dst.Tables("ARTCUSTS").Rows(0).Item("UPS_REF1") & String.Empty
                            cbeREF2.Value = dst.Tables("ARTCUSTS").Rows(0).Item("UPS_REF2") & String.Empty
                            cbeREF3.Value = dst.Tables("ARTCUSTS").Rows(0).Item("UPS_REF3") & String.Empty

                        Case "FEDEX"
                            cbeREF1.Value = dst.Tables("ARTCUSTS").Rows(0).Item("FDX_REF1") & String.Empty
                            cbeREF2.Value = dst.Tables("ARTCUSTS").Rows(0).Item("FDX_REF2") & String.Empty
                            cbeREF3.Value = dst.Tables("ARTCUSTS").Rows(0).Item("FDX_REF3") & String.Empty
                    End Select
                End If


                Refresh_Refs(0)

            End If

            ' If set to Recipient then do not change.
            If txtBilling_Type.Text <> "R" Then
                If rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" Then
                    If Carrier_Code = "FEDEX" Then
                        txtBilling_Type.Text = "C"
                    Else
                        txtBilling_Type.Text = "2"
                    End If
                ElseIf rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" Then
                    If Carrier_Code = "FEDEX" Then
                        txtBilling_Type.Text = "P"
                    Else
                        txtBilling_Type.Text = "3"
                    End If
                Else
                    If Carrier_Code = "FEDEX" Then
                        txtBilling_Type.Text = "O"
                    Else
                        txtBilling_Type.Text = "1"
                    End If
                End If
            End If

            txt_ValueChanged(txtBilling_Type, Nothing)

            If rowSOTCARR1.Item("CARRIER_TYPE") & String.Empty <> "U" Then
                Exit Sub
            End If

            If txt3PAccountNo.Tag = rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty Then
                Exit Sub
            Else
                txt3PAccountNo.Clear()
                txt3pCountry.Clear()
                txt3PZipCode.Clear()
            End If

            txt3PAccountNo.Tag = rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty

            txt3PAccountNo.Text = txt3PAccountNo.Text.Trim
            txt3pCountry.Text = txt3pCountry.Text.Trim.ToUpper
            txt3PZipCode.Text = txt3PZipCode.Text.Trim

            ' Prepopulate any Account numbers if the user did not provide them
            Select Case rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty
                Case "F"
                    If txt3PAccountNo.TextLength = 0 Then txt3PAccountNo.Text = (rowARTCUST1.Item("FDX_ACCT_NO") & String.Empty).ToString.Trim
                    If txt3pCountry.TextLength = 0 Then txt3pCountry.Text = (rowARTCUST1.Item("CUST_COUNTRY") & String.Empty).ToString.Trim
                    If txt3PZipCode.TextLength = 0 Then txt3PZipCode.Text = (rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty).ToString.Trim
                Case "U"
                    If txt3PAccountNo.TextLength = 0 Then txt3PAccountNo.Text = (rowARTCUST1.Item("UPS_ACCT_NO") & String.Empty).ToString.Trim
                    If txt3pCountry.TextLength = 0 Then txt3pCountry.Text = (rowARTCUST1.Item("CUST_COUNTRY") & String.Empty).ToString.Trim
                    If txt3PZipCode.TextLength = 0 Then txt3PZipCode.Text = (rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty).ToString.Trim
            End Select
        End If

    End Sub

    Private Sub cmdGroup_Click(sender As System.Object, e As System.EventArgs) Handles cmdGroup.Click

        Absx1.txtFor("CUST_CODE").Text = Absx1.txtFor("CUST_CODE").Text.Replace("'", String.Empty).Trim

        If Absx1.txtFor("CUST_CODE").TextLength = 0 Then
            MsgBox("Please Select Customer", MsgBoxStyle.OkOnly, "None Selected")
            Exit Sub
        End If

        Group_No_Selected = String.Empty

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("ORDR_GROUP_NO")
        ASCMAIN1.CodeSelector.SQL = "Select ORDR_GROUP_NO,CUST_CODE,ORDR_DATE,ORDR_SHIP_DATE,ORDR_CANCEL_DATE," _
        & " ORDR_CUST_PO,ORDR_AMT,ORDR_QTY,ORDR_CNT,ORDR_CNT_OPEN," _
        & " ORDR_CNT_PICK, ORDR_CGS_SHIP, ORDR_AMT_DISC, " _
        & " SALES_DIVISION_CODE, CUST_DC_NO, ORDR_DEPT From SOTORDR0" _
        & " Where CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"

        ASCMAIN1.CodeSelector.MultipleSelections = True
        Using F As New ASFCODE1
            F.ShowDialog()
        End Using

        If ASCMAIN1.CodeSelector.Selections <> 0 Then
            For Each ORDR_GROUP_NO As String In ASCMAIN1.CodeSelector.SelectedCodes
                Group_No_Selected &= ",'" & ORDR_GROUP_NO & "'"
            Next
        End If

        If Group_No_Selected.Length <> 0 Then
            If Group_No_Selected.Length = 13 Then
                Batched_Group = False
                Absx1.txtFor("ORDR_GROUP_NO").Text = Mid(Group_No_Selected, 3, 10)
            Else
                Batched_Group = True
                Absx1.txtFor("ORDR_GROUP_NO").Text = "Batched"
            End If
            Click_Command("Select")
        End If

    End Sub

    Private Sub txtPRE_1_ValueChanged(sender As Object, e As System.EventArgs) Handles txtPRE_1.ValueChanged
        Refresh_Refs(1)
    End Sub

    Private Sub txtPRE_2_ValueChanged(sender As Object, e As System.EventArgs) Handles txtPRE_2.ValueChanged
        Refresh_Refs(2)
    End Sub

    Private Sub txtPRE_3_ValueChanged(sender As Object, e As System.EventArgs) Handles txtPRE_3.ValueChanged
        Refresh_Refs(3)
    End Sub

    Private Sub txtSUFF_1_ValueChanged(sender As Object, e As System.EventArgs) Handles txtSUFF_1.ValueChanged
        Refresh_Refs(1)
    End Sub

    Private Sub txtSUFF_2_ValueChanged(sender As Object, e As System.EventArgs) Handles txtSUFF_2.ValueChanged
        Refresh_Refs(2)
    End Sub

    Private Sub txtSUFF_3_ValueChanged(sender As Object, e As System.EventArgs) Handles txtSUFF_3.ValueChanged
        Refresh_Refs(3)
    End Sub

    Private Sub cbeREF1_ValueChanged(sender As Object, e As System.EventArgs) Handles cbeREF1.ValueChanged
        If cbeREF1.Value <> String.Empty Then
            cbeREF_CODE1.Value = cbeREF1.SelectedItem.DataValue
        End If
        Refresh_Refs(1)
    End Sub

    Private Sub cbeREF2_ValueChanged(sender As Object, e As System.EventArgs) Handles cbeREF2.ValueChanged
        If cbeREF2.Value <> String.Empty Then
            cbeREF_CODE2.Value = cbeREF2.SelectedItem.DataValue
        End If
        Refresh_Refs(2)
    End Sub

    Private Sub cbeREF3_ValueChanged(sender As Object, e As System.EventArgs) Handles cbeREF3.ValueChanged
        If cbeREF3.Value <> String.Empty Then
            cbeREF_CODE3.Value = cbeREF3.SelectedItem.DataValue
        End If
        Refresh_Refs(3)
    End Sub

    Private Sub cbeREF_CODE1_ValueChanged(sender As Object, e As EventArgs) Handles cbeREF_CODE1.ValueChanged
        Refresh_Refs(1)
    End Sub

    Private Sub cbeREF_CODE2_ValueChanged(sender As Object, e As EventArgs) Handles cbeREF_CODE2.ValueChanged
        Refresh_Refs(2)
    End Sub

    Private Sub cbeREF_CODE3_ValueChanged(sender As Object, e As EventArgs) Handles cbeREF_CODE3.ValueChanged
        Refresh_Refs(3)
    End Sub

    Private Sub optPrint_Status_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optPrint_Status.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Refresh_Printed()
    End Sub

#End Region

#Region "Form Procedures"

    Private Sub RequestAndPrintShipingLabels()
        Try
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Printing Labels")
            permitVandaleToPrintUpsConsignee = 0

            Dim ErrorMessage As String = String.Empty
            Dim shipLabels As New List(Of String)
            Dim lstErrorMessages As New List(Of String)

            Dim requestedLabels As Int64 = 0
            Dim generatedLabels As Int64 = 0

            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("SELECTED = '1'", "CUST_STORE_NO")
                'System.Threading.Thread.Sleep(2000)
                For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "' And CUST_STORE_NO = '" & rowSOTSHIP1.Item("CUST_STORE_NO") & "'", "CART_NO")
                    shipLabels.Clear()
                    ErrorMessage = String.Empty

                    ASCMAIN1.Progress("-", rowSOTCART1.Item("CART_NO") & String.Empty)

                    requestedLabels += 1
                    RequestShippingLabel(shipLabels, ErrorMessage, False, rowSOTSHIP1.Item("SHIP_BOL_NO"), rowSOTCART1.Item("CART_NO"))

                    Dim labelPrinted As Boolean = False
                    For Each shippingLabel As String In shipLabels
                        If shippingLabel.Trim.Length > 0 Then
                            PrintLabel(shippingLabel)
                            labelPrinted = True
                        End If
                    Next

                    If labelPrinted Then
                        generatedLabels += 1
                    End If

                    If ErrorMessage.Length > 0 Then
                        ErrorMessage = $"Error for Shipment {rowSOTSHIP1.Item("SHIP_BOL_NO")}, Carton {rowSOTCART1.Item("CART_NO")}: {ErrorMessage}"
                        lstErrorMessages.Add(ErrorMessage)
                    End If
                Next
                dst.Tables("SOTSHIP1").Select("SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "'", "")(0).Item("SELECTED") = "0"
            Next

            grdSOTCART1.Refresh()

            If permitVandaleToPrintUpsConsignee = 2 Then
                Exit Sub
            End If

            If lstErrorMessages.Count > 0 Then
                MessageBox.Show($"Labels not printed for the following reason(s): {Environment.NewLine} {String.Join(Environment.NewLine, lstErrorMessages.ToArray)}", "Print Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            If generatedLabels > 0 Then
                BeginTrans()
                Update_Record_TDA("SOTCART1")

                ASCMAIN1.sql = " UPDATE SOTSHIP1 Set BILL_OF_LADING_NO = (Select MIN (CART_TRACKING_NO) FROM SOTCART1,SOTPICK1" _
                & " WHERE SOTCART1.PICK_NO = SOTPICK1.PICK_NO AND SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO)" _
                & " WHERE SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM SOTSHIP1 WHERE ORDR_GROUP_NO in (" & ORDR_GROUP_NO & "))"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", txtSHIP_VIA_CODE.Text)
                If rowSOTSVIA1 IsNot Nothing Then
                    ASCMAIN1.sql = " UPDATE SOTSHIP1 SET SHIP_VIA_CODE = '" & rowSOTSVIA1.Item("CARRIER_CODE") & "" & "'" _
                    & " WHERE SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM SOTSHIP1 WHERE ORDR_GROUP_NO in (" & ORDR_GROUP_NO & "))" _
                    & " AND SHIP_VIA_CODE IS NULL" _
                    & " AND SHIP_STATUS = 'P'"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                End If

                CommitTrans()
                MessageBox.Show($"{generatedLabels} of {requestedLabels} requested Shipping Labels set to printer.", "Print Labels", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show($"Error Requesting and Printing Errors: {ex.Message}", "Print Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Refresh_Printed()
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End Try
    End Sub

    ''' <summary>
    ''' Requests Shipping labels for carriers such as fedex, Ups, USPS
    ''' </summary>
    ''' <param name="ShippingLabels">String list of the labels created. Print these</param>
    ''' <param name="ErrorMessage">Error that occurred when processing request</param>
    ''' <param name="PreScreenForErrorsOnly">Boolen to determine if all requirements are meet. If true only evaluation is done.</param>
    ''' <returns></returns>
    ''' <remarks>Any errors or missing attributes will be returned in the ErrorMessage Parameter</remarks>
    Private Function RequestShippingLabel(ByRef ShippingLabels As List(Of String), ByRef ErrorMessage As String, ByVal PreScreenForErrorsOnly As Boolean, Passed_BOL As String, Passed_Cart_No As String) As Boolean

        Dim createCarrierLabels As Boolean = False
        ErrorMessage = String.Empty

        Dim rowSOTSHIP1 As DataRow = Nothing
        Dim rowSOTSVIA1 As DataRow = Nothing
        Dim rowSOTCARR1 As DataRow = Nothing
        Dim rowSOTPICK1 As DataRow = Nothing

        Dim SHIP_VIA_CODE As String = String.Empty
        Dim SHIP_PACKAGE_NO As Int64 = 0
        Dim pkgId As Int64 = 0

        Try
            rowSOTSHIP1 = dst.Tables("SOTSHIP1").Rows.Find(Passed_BOL)
            SHIP_VIA_CODE = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text ' rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty

            rowSOTSVIA1 = LookUp("SOTSVIA1", SHIP_VIA_CODE)
            If rowSOTSVIA1 IsNot Nothing Then
                rowSOTCARR1 = LookUp("SOTCARR1", rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty)
                If rowSOTCARR1 IsNot Nothing AndAlso rowSOTCARR1.Item("CARRIER_TYPE") = "U" Then
                    createCarrierLabels = True
                End If
            Else
                ErrorMessage = "Invalid or missing Ship Via for shipping label request"
            End If

        Catch ex As Exception
            ErrorMessage = "The following error occurred when evaluating a shipping label request: " & ex.Message
            Return False
        End Try

        ' Returns False since there is nothing to do. False with ErrorMessage indicates as an error occurred.
        If Not createCarrierLabels Then Return False

        RequestShippingLabel = True
        Try

            ' Load and validate Customer
            Dim CUST_CODE As String = MyBase.Absx1.txtFor("CUST_CODE").Text
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
            If rowARTCUST1 Is Nothing Then
                ErrorMessage = "Invalid or missing Customer Code for shipping label request"
                Return False
            Else
                txt3PAccountNo.Text = txt3PAccountNo.Text.Trim
                txt3pCountry.Text = txt3pCountry.Text.Trim.ToUpper
                If txt3pCountry.Text.Length = 0 Then txt3pCountry.Text = "US"
                If txt3pCountry.Text.StartsWith("US") Then txt3pCountry.Text = "US"
                txt3PZipCode.Text = txt3PZipCode.Text.Trim

            End If

            Dim CARRIER_CODE As String = rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty
            Dim CARRIER_PROD_CODE As String = rowSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty

            ' Load and Validate Carrier/Ship Method
            Dim rowSOTCARR2 As DataRow = LookUp("SOTCARR2", New String() {CARRIER_CODE, CARRIER_PROD_CODE})
            If rowSOTCARR2 Is Nothing Then
                ErrorMessage = "Invalid or missing Carrier / Ship Method combination for shipping label request"
                Return False
            End If

            ' Credentials
            Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)
            Dim ShippingLabelDirectory As String = (rowSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim
            Dim PROVIDER_TYPE As String = (rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty).ToString.Trim

            If rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty = String.Empty Then
                ErrorMessage = "Invalid or missing Carrier Account Number for shipping label request"
                Return False
            End If

            Try
                If ASCMAIN1.Running_in_VS Then
                    ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "C:\")
                    ShippingLabelDirectory = ShippingLabelDirectory.Replace("R:\", "C:\")
                End If
                If ShippingLabelDirectory.Length > 0 Then
                    If Not My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                        My.Computer.FileSystem.CreateDirectory(ShippingLabelDirectory)
                    End If
                End If
            Catch ex As Exception
                ErrorMessage = ex.Message
                ShippingLabelDirectory = String.Empty
            End Try

            If ShippingLabelDirectory.Length > 0 AndAlso Not ShippingLabelDirectory.EndsWith("\") Then
                ShippingLabelDirectory = ShippingLabelDirectory & "\"
            End If

            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", rowSOTSHIP1.Item("WHSE_CODE"))
            If rowICTWHSE1 Is Nothing Then
                ErrorMessage = "Invalid or missing Warehouse"
                Return False
            End If

            ' look at ship via settings 

            If rowSOTSVIA1 IsNot Nothing Then
                txt3PAccountNo.Text = txt3PAccountNo.Text.Trim
                txt3pCountry.Text = txt3pCountry.Text.Trim.ToUpper
                txt3PZipCode.Text = txt3PZipCode.Text.Trim.PadLeft(5, "0")
                If txt3pCountry.TextLength = 0 OrElse txt3pCountry.Text.StartsWith("US") Then
                    txt3pCountry.Text = "US"
                End If

                Select Case rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty
                    Case "F" ' Fedex
                        If rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" OrElse txtBilling_Type.Text = "C" OrElse txtBilling_Type.Text = "R" Then
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                'ErrorMessage = "Fedex Collect type Ship Vias require an Account Code, Zip Code and Country Code in the customer master."
                                'Return False
                            End If

                        ElseIf rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" OrElse txtBilling_Type.Text = "P" Then
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                ErrorMessage = "Fedex Third Party Shipments require an Account Code, Zip Code and Country Code."
                                Return False
                            End If
                        End If

                    Case "U" ' Ups
                        If (rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" AndAlso rowSOTSVIA1.Item("REQUIRES_ACCT_NO") & String.Empty <> "1") _
                            OrElse txtBilling_Type.Text = "4" Then
                            If COMPANY_CODE = "VAN" Then
                                If permitVandaleToPrintUpsConsignee <> 1 Then
                                    If MessageBox.Show("You are not permitted to Ship UPS Consignee. Do you want to contiue to Ship UPS Consignee?", "Label", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                        'ErrorMessage = "User Canceled Shipping UPS Consignee."
                                        permitVandaleToPrintUpsConsignee = 2
                                        Return False
                                    End If
                                    permitVandaleToPrintUpsConsignee = 1
                                End If
                            End If
                            'clsShip.Payor = TPayorTypes.ptConsignee
                        ElseIf rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" OrElse txtBilling_Type.Text = "2" Then
                            ' Use the Account Information on Customer Master ARTCUST1
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                ErrorMessage = "UPS Collect Shipments require an Account Code, Zip Code and Country Code."
                                Return False
                            End If
                        ElseIf rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" OrElse txtBilling_Type.Text = "3" Then
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                ErrorMessage = "UPS Third Party Shipments require an Account Code, Zip Code and Country Code."
                                Return False
                            End If
                        End If
                End Select
            End If

            ' If Fedex and Collect must be a Ground delivery
            If rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty = "F" _
                AndAlso CARRIER_PROD_CODE <> "15" _
                AndAlso (rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" OrElse txtBilling_Type.Text = "C") Then
                ErrorMessage = "Fedex Collect shipments must ship ground. Choose Recipient payor type on the 'Header Info' tab for non ground shipments."

                Return False
            End If

            If PreScreenForErrorsOnly Then Return True

            '*******************************************************************************

            Dim isInternationalShipment As Boolean = False
            Dim fedexSmartPost As Int16 = 26

            Dim PICK_NO As String = String.Empty
            Dim ORDR_NO As String = String.Empty
            Dim CUST_STORE_NO As String = String.Empty
            Dim ORDR_NO_WEB As String = String.Empty
            Dim ORDR_CUST_PO As String = String.Empty
            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

            Dim FRT_TERMS As String = rowSOTSHIP1.Item("FRT_TERMS") & ""
            Dim PPA_FREIGHT As Decimal = 0
            Dim OUR_FREIGHT As Decimal = 0

            dst.Tables("WHTSHPC1").Rows.Clear()
            dst.Tables("WHTSHPC2").Rows.Clear()
            dst.Tables("WHTSHPC5").Rows.Clear()
            dst.Tables("WHTSHPCS").Rows.Clear()
            dst.Tables("WHTSHPCC").Rows.Clear()
            dst.Tables("WHTSHPCP").Rows.Clear()

            Dim SHIP_CNTL_NO As String = String.Empty 'ASCMAIN1.Next_Control_No("WHTSHPC1.SHIP_CNTL_NO")
            Dim clsShip As New TAC.WHCSHIP1

            ' Credentials
            clsShip.Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
            clsShip.UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
            clsShip.Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
            clsShip.AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
            clsShip.UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
            clsShip.FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            'clsShip.LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim
            clsShip.LabelStockType = optLabel_Type.Value


            Dim rowWHTSHPC1 As DataRow = Nothing
            Dim rowWHTSHPC2 As DataRow = Nothing
            Dim rowWHTSHPC5 As DataRow = Nothing

            rowWHTSHPC1 = dst.Tables("WHTSHPC1").NewRow
            SHIP_CNTL_NO = ASCMAIN1.Next_Control_No("WHTSHPC1.SHIP_CNTL_NO")
            rowWHTSHPC1.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
            rowWHTSHPC1.Item("CARRIER_CODE") = CARRIER_CODE
            rowWHTSHPC1.Item("CARRIER_PROD_CODE") = CARRIER_PROD_CODE
            rowWHTSHPC1.Item("CARRIER_ACCOUNT_NO") = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
            dst.Tables("WHTSHPC1").Rows.Add(rowWHTSHPC1)

            rowWHTSHPC1.Item("STATUS") = "I"
            rowWHTSHPC1.Item("ERROR_MSG") = String.Empty
            rowWHTSHPC1.Item("SHIP_DATE") = CDate(dteSHIP_DATE_SHIPPED.Value).ToString("MM/dd/yyyy")
            rowWHTSHPC1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowWHTSHPC1.Item("OPS_YYYYWW") = ASCMAIN1.CYW
            rowWHTSHPC1.Item("CUST_CODE") = CUST_CODE
            rowWHTSHPC1.Item("INIT_DATE") = DATETIME_STAMP
            rowWHTSHPC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowWHTSHPC1.Item("LAST_DATE") = DATETIME_STAMP
            rowWHTSHPC1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowWHTSHPC1.Item("MASTER_TRACKING_NO") = String.Empty
            rowWHTSHPC1.Item("CUSTOMS_VALUE") = 0
            rowWHTSHPC1.Item("SHIP_BOL_NO") = SHIP_BOL_NO
            rowWHTSHPC1.Item("SHIP_VIA_CODE") = SHIP_VIA_CODE
            rowWHTSHPC1.Item("INSURED_VALUE") = 0
            rowWHTSHPC1.Item("INSURED_SHIPMENT") = "0"

            ' Sender Information
            With clsShip.Sender
                .Company = (rowICTWHSE1.Item("WHSE_DESC") & String.Empty).ToString.Trim
                .FirstName = (rowICTWHSE1.Item("WHSE_CONTACT") & String.Empty).ToString.Trim
                .MiddleInitial = String.Empty
                .LastName = String.Empty
                .Address1 = (rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty).ToString.Trim
                .Address2 = (rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty).ToString.Trim
                .Address3 = (rowICTWHSE1.Item("WHSE_ADDR3") & String.Empty).ToString.Trim
                .City = (rowICTWHSE1.Item("WHSE_CITY") & String.Empty).ToString.Trim
                .State = (rowICTWHSE1.Item("WHSE_STATE") & String.Empty).ToString.Trim
                .ZipCode = (rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                .CountryCode = (rowICTWHSE1.Item("WHSE_COUNTRY") & String.Empty).ToString.Trim.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 OrElse .CountryCode = "USA" Then .CountryCode = "US"
                .Phone = (rowICTWHSE1.Item("WHSE_PHONE") & String.Empty).ToString.Trim

                rowWHTSHPC5 = dst.Tables("WHTSHPC5").NewRow
                rowWHTSHPC5.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                rowWHTSHPC5.Item("SHIP_ADDR_TYPE") = "SF"
                rowWHTSHPC5.Item("SHIP_FIRST_NAME") = .FirstName
                rowWHTSHPC5.Item("SHIP_LAST_NAME") = .LastName
                rowWHTSHPC5.Item("SHIP_MIDDLE_INIT") = .MiddleInitial
                rowWHTSHPC5.Item("SHIP_PHONE") = .Phone
                rowWHTSHPC5.Item("SHIP_FAX") = .Fax
                rowWHTSHPC5.Item("SHIP_EMAIL") = .eMail
                rowWHTSHPC5.Item("SHIP_COMPANY") = .Company
                rowWHTSHPC5.Item("SHIP_ADDR1") = .Address1
                rowWHTSHPC5.Item("SHIP_ADDR2") = .Address2
                rowWHTSHPC5.Item("SHIP_CITY") = .City
                rowWHTSHPC5.Item("SHIP_STATE") = .State
                rowWHTSHPC5.Item("SHIP_ZIP_CODE") = .ZipCode
                rowWHTSHPC5.Item("SHIP_COUNTRY_CODE") = .CountryCode
                rowWHTSHPC5.Item("SHIP_RESIDENTIAL") = "0"
                rowWHTSHPC5.Item("SHIP_PO_BOX") = "0"
                dst.Tables("WHTSHPC5").Rows.Add(rowWHTSHPC5)

            End With

            ' Recipient
            With clsShip.Recipient
                If ASCMAIN1.USER_ID = "tgcv" Then
                    .Company = (rowICTWHSE1.Item("WHSE_DESC") & String.Empty).ToString.Trim
                    .FirstName = (rowICTWHSE1.Item("WHSE_CONTACT") & String.Empty).ToString.Trim
                    .Address1 = (rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty).ToString.Trim
                    .Address2 = (rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty).ToString.Trim
                    .Address3 = (rowICTWHSE1.Item("WHSE_ADDR3") & String.Empty).ToString.Trim
                    .City = (rowICTWHSE1.Item("WHSE_CITY") & String.Empty).ToString.Trim
                    .State = (rowICTWHSE1.Item("WHSE_STATE") & String.Empty).ToString.Trim
                    .ZipCode = (rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                Else
                    ' .FirstName = IIf(rowSOTSHIP1.Item("CUST_CONTACT") & "".Length > 0, rowSOTSHIP1.Item("CUST_CONTACT"), rowSOTSHIP1.Item("CUST_NAME") & "")

                    .FirstName = rowSOTSHIP1.Item("CUST_CONTACT") & ""

                    .Address1 = rowSOTSHIP1.Item("CUST_ADDR1") & ""
                    .Address2 = rowSOTSHIP1.Item("CUST_ADDR2") & ""
                    .Address3 = rowSOTSHIP1.Item("CUST_ADDR3") & ""
                    .Company = rowSOTSHIP1.Item("CUST_NAME") & ""
                    .City = rowSOTSHIP1.Item("CUST_CITY") & ""
                    .State = rowSOTSHIP1.Item("CUST_STATE") & ""
                    .ZipCode = rowSOTSHIP1.Item("CUST_ZIP_CODE") & ""
                End If

                .MiddleInitial = ""
                .LastName = ""

                .CountryCode = rowSOTSHIP1.Item("CUST_COUNTRY").ToString.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 OrElse .CountryCode = "USA" Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                .Phone = rowSOTSHIP1.Item("CUST_PHONE") & ""

                If .Phone.Trim.Length = 0 Then
                    .Phone = clsShip.Sender.Phone
                End If

                If .Phone.Trim.Length = 0 Then
                    .Phone = "1234567890"
                End If
                .IsResidental = False
                .IsPOBox = False

                rowWHTSHPC5 = dst.Tables("WHTSHPC5").NewRow
                rowWHTSHPC5.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                rowWHTSHPC5.Item("SHIP_ADDR_TYPE") = "ST"
                rowWHTSHPC5.Item("SHIP_FIRST_NAME") = .FirstName
                rowWHTSHPC5.Item("SHIP_LAST_NAME") = .LastName
                rowWHTSHPC5.Item("SHIP_MIDDLE_INIT") = .MiddleInitial
                rowWHTSHPC5.Item("SHIP_PHONE") = .Phone
                rowWHTSHPC5.Item("SHIP_FAX") = .Fax
                rowWHTSHPC5.Item("SHIP_EMAIL") = .eMail
                rowWHTSHPC5.Item("SHIP_COMPANY") = .Company
                rowWHTSHPC5.Item("SHIP_ADDR1") = .Address1
                rowWHTSHPC5.Item("SHIP_ADDR2") = .Address2
                rowWHTSHPC5.Item("SHIP_CITY") = .City
                rowWHTSHPC5.Item("SHIP_STATE") = .State
                rowWHTSHPC5.Item("SHIP_ZIP_CODE") = .ZipCode
                rowWHTSHPC5.Item("SHIP_COUNTRY_CODE") = .CountryCode
                rowWHTSHPC5.Item("SHIP_RESIDENTIAL") = "0"
                rowWHTSHPC5.Item("SHIP_PO_BOX") = "0"
                dst.Tables("WHTSHPC5").Rows.Add(rowWHTSHPC5)

                isInternationalShipment = (.CountryCode <> "US") OrElse (.CountryCode = "US" AndAlso .State = "PR")
            End With

            Select Case PROVIDER_TYPE
                Case WHCSHIP1.ProviderTypeFedex
                    If Not isInternationalShipment Then
                        clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpress
                    Else
                        clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpressInternational
                    End If
                Case WHCSHIP1.ProviderTypeUPS
                    clsShip.Service = WHCSHIP1.ServiceProviders.UPS
                Case WHCSHIP1.ProviderTypeUSPS
                    clsShip.Service = WHCSHIP1.ServiceProviders.USPS
                Case WHCSHIP1.ProviderTypeCanada
                    clsShip.Service = WHCSHIP1.ServiceProviders.CanadaPost
                Case Else
                    Return False
            End Select

            ' Build a package for each Carton for the current Pick Ticket
            ' Change as of 1/21/2013
            ' Some shipments are multi Pick Tickets and some Pick Tickets are combined into 1 carton.
            ' The carton sequence will be used to group pick tickets into one carton and also
            ' be used to identify the sequence the Shipping label will get printed
            ' The user is not permitted to deselect a pick ticket; therefore, no londfer need to use dst.Tables("SOTPICK1").Select("SELECTED = '1'", "PICK_NO")
            clsShip.PackageDetailList.Clear()
            Dim cartSequenceNos As List(Of Int16) = New List(Of Int16)

            ' Commodities for international shipments
            clsShip.TotalCustomsValue = 0
            clsShip.CommodityDetailList.Clear()
            Dim COMMODITY_LNO As Int16 = 1
            Dim itemList As List(Of String) = New List(Of String)

            Dim modifyAddress As Boolean = False

            For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select("PICK_NO = '" & dst.Tables("SOTCART1").Select("CART_NO = '" & Passed_Cart_No & "'")(0).Item("PICK_NO") & "'")
                PICK_NO = rowSOTPICK1.Item("PICK_NO") & String.Empty
                ORDR_NO = rowSOTPICK1.Item("ORDR_NO") & String.Empty
                SHIP_BOL_NO = rowSOTPICK1.Item("SHIP_BOL_NO") & String.Empty
                ORDR_CUST_PO = rowSOTPICK1.Item("ORDR_CUST_PO") & String.Empty
                CUST_STORE_NO = rowSOTPICK1.Item("CUST_STORE_NO") & String.Empty
                PPA_FREIGHT = 0
                OUR_FREIGHT = 0

                If chkUseStoreAddress.Checked AndAlso Not modifyAddress Then
                    modifyAddress = True
                    With clsShip.Recipient
                        Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO})

                        If rowARTCUST2 IsNot Nothing Then
                            .FirstName = rowARTCUST2.Item("CUST_CONTACT") & ""
                            .Address1 = rowARTCUST2.Item("CUST_ADDR1") & ""
                            .Address2 = rowARTCUST2.Item("CUST_ADDR2") & ""
                            .Company = rowARTCUST2.Item("CUST_NAME") & ""
                            .City = rowARTCUST2.Item("CUST_CITY") & ""
                            .State = rowARTCUST2.Item("CUST_STATE") & ""
                            .ZipCode = rowARTCUST2.Item("CUST_ZIP_CODE") & ""
                            .Phone = rowARTCUST2.Item("CUST_PHONE") & ""
                        End If

                    End With
                End If

                ' Get the Invoice Number now so we can put it on the label
                Dim INV_NO As String = ASCMAIN1.Next_Control_No("INV_NO_01")
                rowSOTPICK1.Item("INV_NO") = INV_NO

                ' See if we have cartons setup
                If dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'").Length = 0 Then
                    Continue For
                End If

                ' See if the carton has products
                If dst.Tables("SOTCART2").Select("PICK_NO = '" & PICK_NO & "' AND ISNULL(QTY_PACKED, 0) > 0 ").Length = 0 Then
                    Continue For
                End If

                Dim Carton_Sort As Integer = 0
                For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("CART_NO = '" & Passed_Cart_No & "'", "CART_SEQ")
                    ' This is done to place multi pick tickets into one carton
                    Dim CART_SEQ As Int16 = rowSOTCART1.Item("CART_SEQ")
                    If cartSequenceNos.Contains(CART_SEQ) Then
                        Continue For
                    End If
                    cartSequenceNos.Add(CART_SEQ)

                    Dim PACKAGING_TYPE As String = rowSOTCART1.Item("PACKAGING_TYPE") & String.Empty
                    Dim PKG_CODE As String = rowSOTCART1.Item("PKG_CODE") & String.Empty
                    Dim rowWHTPKGM1 As DataRow = LookUp("WHTPKGM1", PKG_CODE)

                    Carton_Sort += 1
                    pkgId = CART_SEQ ' (Val(StrReverse(StrReverse(rowSOTCART1.Item("CART_NO").ToString).Substring(0, 8))))

                    Dim shipPackageDetail As New PackageDetail
                    With shipPackageDetail
                        .PackagingType = Val(PACKAGING_TYPE)

                        ' This is done to place multi pick tickets into one carton. Need combined weight 
                        .Weight = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "CART_SEQ = " & CART_SEQ) & String.Empty)
                        If .Weight = 0 Then
                            .Weight = 1
                        End If

                        '*************************************
                        '        Convert to Ounces
                        '*************************************
                        .Weight = Convert.ToInt16(.Weight * 16)

                        If rowWHTPKGM1 IsNot Nothing Then
                            If rowWHTPKGM1.Item("PKG_CODE") & String.Empty = "OTHER" Then
                                .Length = Val(rowSOTCART1.Item("LENGTH") & String.Empty)
                                .Width = Val(rowSOTCART1.Item("WIDTH") & String.Empty)
                                .Height = Val(rowSOTCART1.Item("HEIGHT") & String.Empty)
                            Else
                                .Length = Val(rowWHTPKGM1.Item("PKG_L") & String.Empty)
                                .Width = Val(rowWHTPKGM1.Item("PKG_W") & String.Empty)
                                .Height = Val(rowWHTPKGM1.Item("PKG_H") & String.Empty)
                            End If
                        End If

                        Dim reference As String = String.Empty
                        Dim refCount As Int16 = 0

                        Select Case PROVIDER_TYPE
                            Case WHCSHIP1.ProviderTypeFedex
                                ' Fedex allows up to 3 References
                                If (rowSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 0 Then
                                    If Mid(rowSOTCART1.Item("REFERENCE1"), 1, 4) = "PO:-" Then
                                        reference &= ";" & Replace(rowSOTCART1.Item("REFERENCE1") & String.Empty, "PO:-", "PO:")
                                    Else
                                        reference &= ";" & rowSOTCART1.Item("REFERENCE1") & String.Empty
                                    End If

                                    refCount += 1
                                End If

                                If (rowSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 0 Then
                                    If Mid(rowSOTCART1.Item("REFERENCE2"), 1, 4) = "PO:-" Then
                                        reference &= ";" & Replace(rowSOTCART1.Item("REFERENCE2") & String.Empty, "PO:-", "PO:")
                                    Else
                                        reference &= ";" & rowSOTCART1.Item("REFERENCE2") & String.Empty
                                    End If

                                    refCount += 1
                                End If

                                If (rowSOTCART1.Item("REFERENCE3") & String.Empty).ToString.Trim.Length > 0 Then
                                    If Mid(rowSOTCART1.Item("REFERENCE3"), 1, 4) = "PO:-" Then
                                        reference &= ";" & Replace(rowSOTCART1.Item("REFERENCE3") & String.Empty, "PO:-", "PO:")
                                    Else
                                        reference &= ";" & rowSOTCART1.Item("REFERENCE3") & String.Empty
                                    End If
                                    refCount += 1
                                End If

                                If reference.Length > 0 Then
                                    reference = reference.Substring(1).Trim
                                    reference &= ";"
                                End If

                            Case WHCSHIP1.ProviderTypeUPS
                                ' Ups allows up to 2 References
                                If (rowSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 0 Then
                                    reference &= ";" & rowSOTCART1.Item("REFERENCE1") & String.Empty
                                    refCount += 1
                                End If

                                If (rowSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 0 Then
                                    reference &= ";" & rowSOTCART1.Item("REFERENCE2") & String.Empty
                                    refCount += 1
                                End If

                                If reference.Length > 0 Then
                                    reference = reference.Substring(1).Trim
                                    reference &= ";"
                                End If

                        End Select

                        .Reference = reference
                        .Id = pkgId.ToString("D8")

                    End With
                    clsShip.PackageDetailList.Add(shipPackageDetail)

                    rowWHTSHPC2 = dst.Tables("WHTSHPC2").NewRow
                    rowWHTSHPC2.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    rowWHTSHPC2.Item("SHIP_PACKAGE_NO") = pkgId
                    rowWHTSHPC2.Item("HEIGHT") = shipPackageDetail.Height
                    rowWHTSHPC2.Item("INSURED_VALUE") = 0
                    rowWHTSHPC2.Item("LENGTH") = shipPackageDetail.Length
                    rowWHTSHPC2.Item("NET_CHARGE") = 0
                    rowWHTSHPC2.Item("PACKAGING_TYPE") = Val(shipPackageDetail.PackagingType)
                    rowWHTSHPC2.Item("TOTAL_DISCOUNT") = 0
                    rowWHTSHPC2.Item("TOTAL_SURCHARGES") = 0
                    rowWHTSHPC2.Item("TRACKING_NUMBER") = String.Empty
                    rowWHTSHPC2.Item("WEIGHT") = Convert.ToInt16(shipPackageDetail.Weight)
                    rowWHTSHPC2.Item("WIDTH") = shipPackageDetail.Width
                    rowWHTSHPC2.Item("TRACKING_NO") = String.Empty
                    rowWHTSHPC2.Item("CUST_REF") = ORDR_CUST_PO
                    rowWHTSHPC2.Item("INV_BOL_NO") = SHIP_BOL_NO
                    rowWHTSHPC2.Item("CART_NO") = rowSOTCART1.Item("CART_NO") & String.Empty
                    rowWHTSHPC2.Item("INV_NO") = INV_NO
                    rowWHTSHPC2.Item("PO_ORDER_NO") = String.Empty
                    rowWHTSHPC2.Item("DEPT_NO") = (rowSOTPICK1.Item("ORDR_DEPT") & String.Empty).ToString.Trim
                    dst.Tables("WHTSHPC2").Rows.Add(rowWHTSHPC2)
                Next

                If isInternationalShipment Then
                    ' Set the Customs value
                    'clsShip.TotalCustomsValue = Val(rowSOTPICK1.Item("PICK_AMT_CONF") & String.Empty)

                    For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select("PICK_NO = '" & PICK_NO & "'")
                        Dim STYLE_CODE As String = rowSOTCART2.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowSOTCART2.Item("COLOR_CODE")

                        'Dim ITEM_CODE As String = STYLE_CODE & Chr(0) & COLOR_CODE

                        If itemList.Contains(STYLE_CODE) Then Continue For

                        itemList.Add(STYLE_CODE)

                        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        ' Just in case a non item is permitted in the shipment
                        If rowICTSTYL1 Is Nothing Then Continue For

                        Dim CommodityDetail As New CommodityDetail
                        CommodityDetail.Description = rowICTSTYL1.Item("STYLE_DESC") & String.Empty

                        Dim NumberOfPieces As Int16 = Val(dst.Tables("SOTCART2").Compute("SUM(QTY_PACKED)", "STYLE_CODE = '" & STYLE_CODE & "' and PICK_NO = '" & PICK_NO & "'") & String.Empty)

                        CommodityDetail.NumberOfPieces = NumberOfPieces
                        CommodityDetail.Quantity = NumberOfPieces
                        CommodityDetail.QuantityUnit = "EA"

                        Dim pickUnitPrice As Decimal = Val(dst.Tables("SOTPICK2").Compute("MAX(PICK_UNIT_PRICE)", "PICK_NO = '" & PICK_NO & "' AND STYLE_CODE = '" & STYLE_CODE & "'") & String.Empty)
                        CommodityDetail.UnitPrice = pickUnitPrice

                        CommodityDetail.Weight = Val(rowICTSTYL1.Item("STYLE_WEIGHT") & String.Empty) ' Leave as pounds
                        CommodityDetail.Manufacturer = (rowICTSTYL1.Item("COUNTRY_CODE") & String.Empty).ToString.ToUpper.Trim ' "US" '
                        If CommodityDetail.Manufacturer.Length = 0 Then
                            CommodityDetail.Manufacturer = "US"
                        End If
                        clsShip.CommodityDetailList.Add(CommodityDetail)

                        Dim rowWHTSHPCC As DataRow = dst.Tables("WHTSHPCC").NewRow
                        rowWHTSHPCC.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                        rowWHTSHPCC.Item("COMMODITY_LNO") = COMMODITY_LNO
                        COMMODITY_LNO += 1
                        rowWHTSHPCC.Item("COMMODITY_DESC") = CommodityDetail.Description
                        rowWHTSHPCC.Item("NUM_PIECES") = CommodityDetail.NumberOfPieces
                        rowWHTSHPCC.Item("MANUFACTURER") = CommodityDetail.Manufacturer
                        rowWHTSHPCC.Item("HARMONIZED_CODE") = String.Empty
                        rowWHTSHPCC.Item("WEIGHT") = CommodityDetail.Weight
                        rowWHTSHPCC.Item("QUANTITY") = CommodityDetail.Quantity
                        rowWHTSHPCC.Item("QUANTITY_UOM") = CommodityDetail.QuantityUnit
                        rowWHTSHPCC.Item("UNIT_PRICE") = CommodityDetail.UnitPrice
                        dst.Tables("WHTSHPCC").Rows.Add(rowWHTSHPCC)
                    Next
                End If
            Next  ' This is where the For Sotpick1, for sotcart1, for sotcart2 should end 

            ' Shipping Method
            If isInternationalShipment Then
                clsShip.RequestedServiceType = Val(rowSOTSVIA1.Item("CARRIER_PROD_CODE_INTL") & String.Empty)
            Else
                clsShip.RequestedServiceType = Val(rowSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty)
            End If

            If clsShip.RequestedServiceType = fedexSmartPost Then
                clsShip.FedexSmartPost.HubId = rowSOTCARR3.Item("FEDEX_HUB_ID") & String.Empty
            End If

            clsShip.DropOffType = FedexshipintlDropoffTypes.dtRegularPickup


            ' The COLLECT payment type is only supported in FedEx Ground services. The CONSIGNEE type is only supported in UPS service.

            ' For FedEx, when this field is set to a value other than 0 (ptSender), the AccountNumber and 
            ' CountryCode are required to be provided in the request as well. Otherwise, those will default to AccountNumber and CountryCode.

            ' For UPS, when set to ptSender, the AccountNumber is automatically set to AccountNumber. 
            ' When ptRecipient is specified, AccountNumber and ZipCode are required to be provided in the request. 
            ' For return international shipments, this option is invalid for transportation charges. 
            ' And, when ptThirdParty has been specified, the AccountNumber, ZipCode and CountryCode are 
            ' required to be provided in the request. When ptConsignee is specified, it indicates that UPS Consignee Billing 
            ' option is selected, no other fields need to be set. ptConsignee only applies to US/PR and PR/US shipment origins and destination. 

            ' Payor of the Shipmenet
            clsShip.Payor = TPayorTypes.ptSender


            Select Case rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty
                Case "F" ' Fedex
                    If rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" OrElse txtBilling_Type.Text = "C" Then
                        ' Use the Account Information on Customer Master ARTCUST1
                        clsShip.Payor = TPayorTypes.ptCollect
                        If (rowARTCUST1.Item("FDX_ACCT_NO") & String.Empty).ToString.Trim.Length > 0 Then

                            clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                            clsShip.PayorContact.CountryCode = txt3pCountry.Text
                            clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                            If clsShip.PayorContact.CountryCode = String.Empty Then
                                clsShip.PayorContact.CountryCode = "US"
                            End If
                        End If
                    ElseIf rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" OrElse txtBilling_Type.Text = "P" Then
                        clsShip.Payor = TPayorTypes.ptThirdParty
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf txtBilling_Type.Text = "R" Then
                        clsShip.Payor = TPayorTypes.ptRecipient
                        If txt3PAccountNo.Text = "" Then

                            ASCMAIN1.sql = "Select * from ARTCUST2 " _
                            & " where CUST_CODE = '" & rowARTCUST1.Item("CUST_CODE") & "'" _
                            & " And CUST_ADDR_CODE = '" & rowSOTSHIP1.Item("SHIP_ADDR_CODE") & "'" _
                            & " And CUST_ADDR_TYPE = '" & rowSOTSHIP1.Item("SHIP_ADDR_TYPE") & "'"
                            Dim rowARTCUST2 As DataRow = ASCDATA1.GetDataRow
                            If rowARTCUST2 IsNot Nothing Then
                                clsShip.PayorContact.AccountNumber = rowARTCUST2.Item("FDX_ACCT_NO") & ""
                                clsShip.PayorContact.CountryCode = rowARTCUST2.Item("CUST_COUNTRY") & ""
                                clsShip.PayorContact.ZipCode = rowARTCUST2.Item("CUST_ZIP_CODE") & ""
                                If clsShip.PayorContact.CountryCode = String.Empty Then
                                    clsShip.PayorContact.CountryCode = "US"
                                End If
                            End If
                        Else
                            clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                            clsShip.PayorContact.CountryCode = txt3pCountry.Text
                            clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                            If clsShip.PayorContact.CountryCode = String.Empty Then
                                clsShip.PayorContact.CountryCode = "US"
                            End If
                        End If
                    End If

                Case "U" ' Ups
                    If (rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" AndAlso rowSOTSVIA1.Item("REQUIRES_ACCT_NO") & String.Empty <> "1") _
                        OrElse txtBilling_Type.Text = "4" Then
                        clsShip.Payor = TPayorTypes.ptConsignee
                    ElseIf rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" OrElse txtBilling_Type.Text = "2" Then
                        clsShip.Payor = TPayorTypes.ptRecipient
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" OrElse txtBilling_Type.Text = "3" Then
                        clsShip.Payor = TPayorTypes.ptThirdParty
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    End If

            End Select

            Dim rowWHTSHPCP As DataRow

            If clsShip.Payor <> TPayorTypes.ptSender Then
                Dim rowWHTSHPC3 As DataRow = dst.Tables("WHTSHPC3").NewRow
                rowWHTSHPC3("SHIP_CNTL_NO") = SHIP_CNTL_NO
                rowWHTSHPC3("SHIP_BOL_NO") = SHIP_BOL_NO
                rowWHTSHPC3("ACCOUNT_NO_3PL") = txt3PAccountNo.Text
                rowWHTSHPC3("ZIP_CODE_3PL") = txt3PZipCode.Text
                rowWHTSHPC3("COUNTRY_3PL") = txt3pCountry.Text
                dst.Tables("WHTSHPC3").Rows.Add(rowWHTSHPC3)
            End If

            rowWHTSHPCP = dst.Tables("WHTSHPCP").NewRow
            rowWHTSHPCP("SHIP_CNTL_NO") = SHIP_CNTL_NO
            rowWHTSHPCP("PAYOR_TYPE") = "S"
            rowWHTSHPCP("PAYOR_ACCT_NO") = clsShip.PayorContact.AccountNumber & String.Empty
            rowWHTSHPCP("PAYOR_COUNTRY") = clsShip.PayorContact.CountryCode & String.Empty
            dst.Tables("WHTSHPCP").Rows.Add(rowWHTSHPCP)


            ' Payor of the Duties
            clsShip.DutiesPayor = TPayorTypes.ptSender
            If isInternationalShipment Then
                clsShip.DutiesPayor = clsShip.Payor
                clsShip.DutiesPayorContact.AccountNumber = clsShip.PayorContact.AccountNumber
                clsShip.DutiesPayorContact.CountryCode = clsShip.PayorContact.CountryCode
                clsShip.DutiesPayorContact.ZipCode = clsShip.PayorContact.ZipCode
            End If

            rowWHTSHPCP = dst.Tables("WHTSHPCP").NewRow
            rowWHTSHPCP("SHIP_CNTL_NO") = SHIP_CNTL_NO
            rowWHTSHPCP("PAYOR_TYPE") = "D"
            rowWHTSHPCP("PAYOR_ACCT_NO") = clsShip.DutiesPayorContact.AccountNumber & String.Empty
            rowWHTSHPCP("PAYOR_COUNTRY") = clsShip.DutiesPayorContact.CountryCode & String.Empty
            dst.Tables("WHTSHPCP").Rows.Add(rowWHTSHPCP)

            With clsShip
                '.LabelStockType = 4
                '.LabelStockType = "STOCK_4X6.75_LEADING_DOC_TAB"
                Select Case optPrint_Type.Value
                    Case "E"
                        .EzshipLabelImage = EzshipLabelImageTypes.itEltron
                    Case "Z"
                        .EzshipLabelImage = EzshipLabelImageTypes.itZPL
                    Case "X"
                        .EzshipLabelImage = EzshipLabelImageTypes.itZebra
                End Select
                .ShippingLabelDirectory = ShippingLabelDirectory

                ' 09/25/2020
                ' This routine prints one carton label at a time
                ' If we use the Ship BOL No as the file name then its get overwritten if there are multiple cartons.
                ' I need a file for each carton for the Reprint of a generates shipping label.
                If Passed_Cart_No.Length > 0 Then
                    .ShippingLabelPrefix = Passed_Cart_No
                Else
                    .ShippingLabelPrefix = SHIP_CNTL_NO
                End If

                .ShipDate = dteSHIP_DATE_SHIPPED.DateTime.ToString("yyyy-MM-dd")

                If optLabel_Type.Value = "4" Then
                    Dim Custom_Content As String = "<CustomContent>" _
                           & "<TextEntries>" _
                           & "<Position><X>20</X><Y>20</Y></Position>" _
                           & "<Format>PO#:</Format>" _
                           & "<ThermalFontId>18</ThermalFontId>" _
                           & "</TextEntries>" _
                           & "<TextEntries>" _
                           & "<Position><X>75</X><Y>20</Y></Position>" _
                           & "<Format>" & ORDR_CUST_PO & "</Format>" _
                           & "<ThermalFontId>17</ThermalFontId>" _
                           & "</TextEntries>" _
                           & "<TextEntries>" _
                           & "<Position><X>375</X><Y>20</Y></Position>" _
                           & "<Format>Vendor#:</Format>" _
                           & "<ThermalFontId>18</ThermalFontId>" _
                           & "</TextEntries>" _
                           & "<TextEntries>" _
                           & "<Position><X>490</X><Y>20</Y></Position>" _
                           & "<Format>48521</Format>" _
                           & "<ThermalFontId>17</ThermalFontId>" _
                           & "</TextEntries>" _
                           & "<TextEntries>" _
                           & "<Position><X>250</X><Y>130</Y></Position>" _
                           & "<Format>Store#:</Format>" _
                           & "<ThermalFontId>18</ThermalFontId>" _
                           & "</TextEntries>" _
                           & "<TextEntries>" _
                           & "<Position><X>350</X><Y>130</Y></Position>" _
                           & "<Format>" & CUST_STORE_NO & "</Format>" _
                           & "<ThermalFontId>17</ThermalFontId>" _
                           & "</TextEntries>" _
                           & "<TextEntries>" _
                           & "<Position><X>100</X><Y>390</Y></Position>" _
                           & "<Format>SSCC18 - " & Passed_Cart_No & "</Format>" _
                           & "<ThermalFontId>2</ThermalFontId>" _
                           & "</TextEntries>" _
                           & "<BarcodeEntries>" _
                           & "<Position><X>100</X><Y>190</Y></Position>" _
                           & "<Format>" & Passed_Cart_No & "</Format> " _
                           & "<BarHeight>190</BarHeight>" _
                           & "<ThinBarWidth>4</ThinBarWidth>" _
                           & "<BarcodeSymbology>CODE128B</BarcodeSymbology> " _
                           & "</BarcodeEntries>" _
                           & "</CustomContent>"
                    .FedexCustomContent = Custom_Content
                End If

                'If chkSaturday.Checked Then
                '    clsShip.ShipmentSpecialServices = clsShip.ShipmentSpecialServices OrElse &H10000000L
                'End If
            End With

            Try
                BeginTrans()
                Update_Record_TDA("WHTSHPC1")
                Update_Record_TDA("WHTSHPC2")
                Update_Record_TDA("WHTSHPC3")
                Update_Record_TDA("WHTSHPC5")
                Update_Record_TDA("WHTSHPCS")
                Update_Record_TDA("WHTSHPCP")
                Update_Record_TDA("WHTSHPCC")
                CommitTrans()
            Catch ex As Exception
                MsgBox(ex.Message, vbOKOnly, "Update Error")
                ErrorMessage &= " " & ex.Message
                Rollback()
            End Try


            If clsShip.RequestLabel() Then

                dst.Tables("SOTCART1").Select("CART_NO = '" & Passed_Cart_No & "'", "")(0).Item("STATUS") = Printed
                SetPrintedStatus(SHIP_BOL_NO)

                rowWHTSHPC1.Item("ERROR_MSG") = clsShip.LastError & String.Empty
                rowWHTSHPC1.Item("STATUS") = "P"
                If rowWHTSHPC1 IsNot Nothing AndAlso (rowWHTSHPC1.Item("ERROR_MSG") & String.Empty).ToString.Length > 200 Then
                    rowWHTSHPC1.Item("ERROR_MSG") = rowWHTSHPC1("ERROR_MSG").ToString.Substring(0, 200).Trim
                End If
                rowWHTSHPC1.Item("MASTER_TRACKING_NO") = clsShip.MasterTrackingNumber & String.Empty

                For Each shipPackageDetail As PackageDetail In clsShip.PackageDetailList
                    SHIP_PACKAGE_NO = Val(shipPackageDetail.Id)

                    If dst.Tables("WHTSHPC2").Select("SHIP_PACKAGE_NO = " & SHIP_PACKAGE_NO, "").Length > 0 Then
                        rowWHTSHPC2 = dst.Tables("WHTSHPC2").Select("SHIP_PACKAGE_NO = " & SHIP_PACKAGE_NO)(0)
                        rowWHTSHPC2.Item("TRACKING_NO") = shipPackageDetail.TrackingNumber & String.Empty
                        rowWHTSHPC2.Item("BASE_CHARGE") = Val(clsShip.ShipmentBaseCharge(SHIP_PACKAGE_NO) & String.Empty)
                        rowWHTSHPC2.Item("NET_CHARGE") = Val(clsShip.ShipmentNetCharge(SHIP_PACKAGE_NO) & String.Empty)
                        rowWHTSHPC2.Item("TOTAL_DISCOUNT") = Val(clsShip.ShipmentDiscountCharge(SHIP_PACKAGE_NO) & String.Empty)
                        rowWHTSHPC2.Item("TOTAL_SURCHARGES") = Val(clsShip.ShipmentSurCharge(SHIP_PACKAGE_NO) & String.Empty)

                        If clsShip.ShipmentListCharge.ContainsKey(SHIP_PACKAGE_NO) Then
                            rowWHTSHPC2.Item("LIST_PRICE") = Val(clsShip.ShipmentListCharge(SHIP_PACKAGE_NO) & String.Empty)
                        Else
                            rowWHTSHPC2.Item("LIST_PRICE") = rowWHTSHPC2.Item("NET_CHARGE")
                        End If
                        PPA_FREIGHT = Val(rowWHTSHPC2.Item("LIST_PRICE") & String.Empty)
                        OUR_FREIGHT = Val(rowWHTSHPC2.Item("NET_CHARGE") & String.Empty)

                        PICK_NO = String.Empty
                        rowSOTPICK1 = Nothing

                        ' We may have multi pick tickets in a single carton. This stamps them with the same tracking number
                        ' Spread the Customer Freight Cost and Our freight cost across the Pick Tickets
                        Dim numPickTickets As Int16 = dst.Tables("SOTCART1").Select("CART_SEQ = " & SHIP_PACKAGE_NO).Length
                        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("CART_SEQ = " & SHIP_PACKAGE_NO)
                            rowSOTCART1.Item("CART_TRACKING_NO") = shipPackageDetail.TrackingNumber & String.Empty

                            ' Added on 09/25/2020
                            rowSOTCART1.Item("PACKAGING_TYPE") = txtPackageType.Text
                            rowSOTCART1.Item("PKG_CODE") = txtPackageCode.Text
                            rowSOTCART1.Item("PKG_L") = numL.Value
                            rowSOTCART1.Item("PKG_H") = numH.Value
                            rowSOTCART1.Item("PKG_W") = numW.Value

                            'Multi-po if no record, we need to find real carton
                            ASCMAIN1.sql = $"Update SOTCARM1 Set CART_TRACKING_NO = '{shipPackageDetail.TrackingNumber & String.Empty}' 
                                            Where CART_NO in (select distinct CART_NO from SOTCARM2 where ORIG_CART_NO = :PARM1)
                                            and (CART_TRACKING_NO is null or CART_NO = :PARM1)"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {rowSOTCART1("CART_NO")})


                            PICK_NO = rowSOTCART1.Item("PICK_NO") & String.Empty
                            rowSOTPICK1 = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
                            If rowSOTSHIP1.Item("FRT_TERMS") & String.Empty = "PPA" AndAlso rowSOTPICK1("ORDR_SOURCE") & String.Empty <> "W" Then
                                rowSOTPICK1.Item("PICK_FREIGHT") = Val(rowSOTPICK1.Item("PICK_FREIGHT") & String.Empty) + Math.Round(PPA_FREIGHT / numPickTickets, 2)
                            End If
                            rowSOTPICK1.Item("OUR_FREIGHT") = Val(rowSOTPICK1.Item("OUR_FREIGHT") & String.Empty) + Math.Round(OUR_FREIGHT / numPickTickets, 2)
                        Next
                    End If

                    ShippingLabels.Add(shipPackageDetail.ShippingLabel)
                    ShippingLabels.Add(shipPackageDetail.CODLabel)
                    ShippingLabels.Add(shipPackageDetail.ReturnReceipt)
                Next

                Try
                    BeginTrans()
                    Update_Record_TDA("WHTSHPC1")
                    Update_Record_TDA("WHTSHPC2")
                    CommitTrans()
                Catch ex As Exception
                    ErrorMessage &= " " & ex.Message
                    Rollback()
                End Try

            Else
                dst.Tables("SOTSHIP1").Select("SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "'", "")(0).Item("STATUS") = "Error"
                dst.Tables("SOTSHIP1").Select("SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "'", "")(0).Item("ERROR_MSG") = clsShip.LastError

                dst.Tables("SOTCART1").Select("CART_NO = '" & Passed_Cart_No & "'", "")(0).Item("STATUS") = "Error"
                dst.Tables("SOTCART1").Select("CART_NO = '" & Passed_Cart_No & "'", "")(0).Item("ERROR_MSG") = clsShip.LastError

                ErrorMessage &= " " & clsShip.LastError
                RequestShippingLabel = False
            End If

        Catch ex As Exception
            ErrorMessage &= " " & ex.Message
            RequestShippingLabel = False
        End Try

        ErrorMessage = ErrorMessage.Trim

    End Function

    Private Sub RequestTrackingInformation(ByVal rowData As DataRow, ByRef TrackingData As TAC.WHCSHIP1.TrackingData)

        Dim response As String = String.Empty

        Try
            Dim CARRIER_CODE As String = rowData.Item("CARRIER_CODE") & String.Empty
            Dim TRACKING_NO As String = rowData.Item("TRACKING_NO") & String.Empty

            If TRACKING_NO.Length = 0 Then
                TrackingData.Status = "Missing Tracking Number"
            End If

            Dim rowSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)
            Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)

            Dim clsShip As New TAC.WHCSHIP1

            ' Credentials
            clsShip.Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
            clsShip.UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
            clsShip.Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
            clsShip.AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
            clsShip.UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
            clsShip.FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            'clsShip.LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim
            clsShip.LabelStockType = optLabel_Type.Value

            Select Case CARRIER_CODE
                Case "UPS"
                    response = clsShip.UPSTrack(TRACKING_NO)
                Case "FEDEX"
                    response = clsShip.FedExTrack(TRACKING_NO)
            End Select

            TrackingData = clsShip.TrackingInfo

        Catch ex As Exception
            TrackingData.Status = $"Error Requesting Tracking Info: {ex.Message }"
        End Try


    End Sub

    ''' <summary>
    ''' Sends data to the Label Printer
    ''' </summary>
    ''' <param name="LabelData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PrintLabel(ByVal LabelData As String) As Boolean

        Try
            If (ASCMAIN1.USER_ID = "gcv" OrElse ASCMAIN1.USER_ID = "edz") AndAlso ASCMAIN1.Running_in_VS Then
                ' Find Zebra printer

                Dim zebraPrinter As String = FindZebraPrinter()

                Dim vLabelPrinter As New ASCPRINT
                Return vLabelPrinter.SendStringToPrinter(zebraPrinter, LabelData)
            End If
            'cboZebraPrinter.Text
            If ASCMAIN1.CLIENT = "VAN" Then
                If txtLabelPrinter.BackColor = Drawing.Color.Green Then
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
            If printerName.ToUpper.StartsWith("ZP450E") Then
                Return printerName
            End If
        Next printerName

        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZEBRA450") Then
                Return printerName
            End If
        Next printerName

        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZEBRA") Or printerName.ToUpper.StartsWith("ZDESIGN") Or printerName.ToUpper.StartsWith("ZPDESIGN") Then
                Return printerName
            End If
        Next printerName

        Return ""
    End Function

    Private Sub SetPrintedStatus(ByVal SHIP_BOL_NO As String)

        Try
            Dim numPrinted As Int16 = dst.Tables("SOTCART1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "' AND STATUS = 'Printed'").Length
            Dim numCartons As Int16 = dst.Tables("SOTCART1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'").Length
            Dim status As String = NotPrinted

            If numPrinted = numCartons Then
                status = Printed
            ElseIf numPrinted > 0 Then
                status = PartiallyPrinted
            End If

            For Each rowSOTSHIP1x As DataRow In dst.Tables("SOTSHIP1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
                rowSOTSHIP1x.Item("STATUS") = status
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RePrintShippingLabels(ByVal SHIP_BOL_NO As String, ByVal CUST_STORE_NO As String)

        Try

            Dim SHIP_VIA_CODE As String = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text
            Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
            Dim rowSOTCARR1 As DataRow = Nothing

            If rowSOTSVIA1 IsNot Nothing Then
                rowSOTCARR1 = LookUp("SOTCARR1", rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty)
            End If

            If rowSOTCARR1 Is Nothing Then
                MessageBox.Show("Cannot map selected Ship Via to a carrier.", "Reprint Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim ShippingLabelDirectory As String = (rowSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim
            If ShippingLabelDirectory.Length = 0 Then
                MessageBox.Show("The selected Ship Via's Carrier master record does not contain a shipping label directory.", "Reprint Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor

            ' R:\SHIPMENTS\FEDEX
            ' R:\SHIPMENTS\UPS
            If ShippingLabelDirectory.Length > 0 AndAlso Not ShippingLabelDirectory.EndsWith("\") Then
                ShippingLabelDirectory &= "\"
            End If

            Dim printedLabels As Int32 = 0

            For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select($"SHIP_BOL_NO = '{SHIP_BOL_NO}' AND CUST_STORE_NO = '{CUST_STORE_NO}'", "CART_NO")
                Dim CART_NO As String = rowSOTCART1.Item("CART_NO")

                For Each shippingLabelFile As String In My.Computer.FileSystem.GetFiles(ShippingLabelDirectory, FileIO.SearchOption.SearchTopLevelOnly, $"{CART_NO}*.*")
                    Dim ShippingLabel As String = String.Empty
                    Using sw As New System.IO.StreamReader(shippingLabelFile)
                        ShippingLabel = sw.ReadToEnd
                        sw.Close()
                        sw.Dispose()
                    End Using

                    If ShippingLabel.Length > 0 Then
                        PrintLabel(ShippingLabel)
                        printedLabels += 1
                    End If
                Next
            Next

            MessageBox.Show($"{printedLabels} Shipping label(s) sent to printer.", "Reprint Labels", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Error reprinting shipping labels: {ex.Message}", "Reprint Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub VoidShippingLabels(ByVal SHIP_BOL_NO As String, ByVal CUST_STORE_NO As String, ByVal CART_NO As String)
        Try

            Dim SHIP_VIA_CODE As String = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text
            Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
            Dim rowSOTCARR1 As DataRow = Nothing
            Dim numlabelsGenerated As Int32 = 0
            Dim numlabelsVoided As Int32 = 0
            Dim lstErrors As New List(Of String)


            If rowSOTSVIA1 IsNot Nothing Then
                rowSOTCARR1 = LookUp("SOTCARR1", rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty)
            End If

            If rowSOTCARR1 Is Nothing Then
                MessageBox.Show("Cannot map selected Ship Via to a carrier.", "Reprint Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty & "'")(0)
            Dim ShippingLabelDirectory As String = (rowSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim
            Dim PROVIDER_TYPE As String = (rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty).ToString.Trim

            If rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty = String.Empty Then
                MessageBox.Show("Invalid or missing Carrier Account Number for Ship Via Carrier", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor

            Dim clsShip As New TAC.WHCSHIP1

            Select Case rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty
                Case "UPS"
                    clsShip.Service = WHCSHIP1.ServiceProviders.UPS
                Case "FEDEX"
                    clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpress
            End Select

            ' Credentials
            clsShip.Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
            clsShip.UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
            clsShip.Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
            clsShip.AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
            clsShip.UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
            clsShip.FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty

            Dim query As String = String.Empty
            If CART_NO.Length > 0 Then
                query = $"CART_NO = '{CART_NO}'"
            Else
                query = $"SHIP_BOL_NO = '{SHIP_BOL_NO}' And CUST_STORE_NO = '{CUST_STORE_NO}'"
            End If

            For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select(query, "CART_NO")
                Dim CART_TRACKING_NO As String = rowSOTCART1.Item("CART_TRACKING_NO") & String.Empty
                If CART_TRACKING_NO.Length = 0 Then Continue For

                numlabelsGenerated += 1
                Try
                    If clsShip.CancelShipment(CART_TRACKING_NO) Then
                        numlabelsVoided += 1
                        ASCMAIN1.sql = "UPDATE WHTSHPC1 SET STATUS = 'V' WHERE MASTER_TRACKING_NO = :PARM1"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {CART_TRACKING_NO})
                    ElseIf clsShip.LastError.Length > 0 Then
                        lstErrors.Add($"Carton {rowSOTCART1.Item("CART_NO")}: {clsShip.LastError}")
                    End If
                Catch ex As Exception
                    lstErrors.Add($"Carton {rowSOTCART1.Item("CART_NO")}: {ex.Message}")
                End Try
            Next

            If numlabelsGenerated > 0 OrElse numlabelsVoided > 0 Then
                Dim userMessage As String = $"{numlabelsVoided} of {numlabelsGenerated} shipping label(s) were suscessfully voided."
                MessageBox.Show(userMessage, "Void Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Information)

                If lstErrors.Count > 0 Then
                    userMessage = "The following errors occurred:" & Environment.NewLine & String.Join(Environment.NewLine, lstErrors.ToArray)
                    MessageBox.Show(userMessage, "Void Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show($"Error voiding shipping labels: {ex.Message}", "Void Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub Refresh_Printed()
        Dim sql_filter As String = ""
        Dim dvw As DataView = DirectCast(grdSOTSHIP1.DataSource, DataTable).DefaultView
        Select Case optPrint_Status.Value
            Case "A"
                sql_filter = ""
            Case "P"
                sql_filter = "STATUS = '" & Printed & "' or STATUS = '" & PartiallyPrinted & "'"
            Case "N"
                sql_filter = "STATUS <> '" & Printed & "'"
        End Select
        dvw.RowFilter = sql_filter
    End Sub

#End Region

#Region "Overrides"
    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CARTONL", "CARTONW", "CARTONH"
                For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select
                    rowSOTCART1.Item("LENGTH") = Val(Absx1.numFor("CARTONL").Value & "")
                    rowSOTCART1.Item("WIDTH") = Val(Absx1.numFor("CARTONW").Value & "")
                    rowSOTCART1.Item("HEIGHT") = Val(Absx1.numFor("CARTONH").Value & "")
                Next
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "SHIP_VIA_CODE"
                Dim SHIP_VIA_CODE As String = Absx1.txtFor("SHIP_VIA_CODE").Text.Trim

                ASCMAIN1.Add_Value_List(grdSOTCART1, "PACKAGING_TYPE", "SELECT SOTCARR4.PACKAGE_CODE, SOTCARR4.PACKAGE_DESC" _
                                        & " FROM SOTSVIA1, SOTCARR4" _
                                        & " WHERE SOTCARR4.CARRIER_CODE = SOTSVIA1.CARRIER_CODE" _
                                        & " AND SOTSVIA1.SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "'" _
                                        & " ORDER BY PACKAGE_CODE DESC")

                Dim rowSOTSVIA1 As DataRow = dst.Tables("SOTSVIA1").Rows.Find(SHIP_VIA_CODE)
                If rowSOTSVIA1 IsNot Nothing Then
                    commonCarrier = rowSOTSVIA1 IsNot Nothing AndAlso rowSOTSVIA1.Item("CARRIER_TYPE") & String.Empty = "U"

                    Carrier_Code = rowSOTSVIA1.Item("CARRIER_CODE") & ""
                End If

            Case "PACKAGE_TYPE"
                Dim Package_Type As String = Absx1.txtFor("PACKAGE_TYPE").Text
                ASCMAIN1.sql = "Select * from SOTCARR4 where CARRIER_CODE = '" & Carrier_Code & "' " _
                    & " And PACKAGE_CODE = '" & Package_Type & "'"
                Dim row As DataRow = ASCDATA1.GetDataRow
                If row IsNot Nothing Then
                    txtPkg_Type_Desc.Text = row.Item("PACKAGE_DESC") & ""
                Else
                    Package_Type = ""
                    txtPkg_Type_Desc.Text = ""
                End If
                For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select
                    rowSOTCART1.Item("PACKAGING_TYPE") = Package_Type
                Next

            Case "PKG_CODE"
                Dim Pkg_Code As String = Absx1.txtFor("PKG_CODE").Text
                ASCMAIN1.sql = "Select * from WHTPKGM1 Where PKG_CODE = '" & Pkg_Code & "'"
                Dim rowWHTPKGM1 As DataRow = ASCDATA1.GetDataRow
                If rowWHTPKGM1 IsNot Nothing Then
                    Absx1.numFor("CARTONL").Value = Val(rowWHTPKGM1.Item("PKG_L") & "")
                    Absx1.numFor("CARTONW").Value = Val(rowWHTPKGM1.Item("PKG_W") & "")
                    Absx1.numFor("CARTONH").Value = Val(rowWHTPKGM1.Item("PKG_H") & "")

                    For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select
                        rowSOTCART1.Item("PKG_CODE") = Pkg_Code
                    Next
                Else
                    Absx1.numFor("CARTONL").Value = 0
                    Absx1.numFor("CARTONW").Value = 0
                    Absx1.numFor("CARTONH").Value = 0
                End If

            Case "BILLING_TYPE"
                ASCMAIN1.sql = "Select * from SOTSHPP1 where SHIPPER_CODE = '" & IIf(Carrier_Code = "FEDEX", "FDX", "UPS") & "' " _
                    & " And SHIP_PARM_TYPE = 'BILLING'" _
                    & " And SHIP_PARM_CODE = '" & txtBilling_Type.Text & "'"
                Dim row As DataRow = ASCDATA1.GetDataRow
                If row IsNot Nothing Then
                    txtBilling_Desc.Text = row.Item("SHIP_PARM_DESC") & ""
                Else
                    txtBilling_Desc.Text = ""
                End If

        End Select

    End Sub

#End Region

#Region "Serial and Com Connections"

    ' Handles Keyboard wedge
    Private receivingWedgeScan As Boolean = False
    Private strWedgeScan As String = String.Empty
    Private registeredWeight As Decimal = 0

    ''' <summary>
    ''' Form activate - Calls to setup devices
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        SetUpPortsAndPrinters()
    End Sub

    ''' <summary>
    ''' Sets the Printer Settings
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetUpPortsAndPrinters()

        Dim tooltip As New System.Windows.Forms.ToolTip()

        '**************************
        '**    Laser Printer
        '**************************
        Try
            txtLaserPrinter.Text = ASCMAIN1.LaserPrinterIpAddress
            tooltip.SetToolTip(txtLaserPrinter, ASCMAIN1.LaserPrinterIpAddress)
            If ASCMAIN1.LaserPrinterIpAddress.Length = 0 Then
                txtLaserPrinter.Appearance.BackColor = Drawing.Color.Red
            Else
                txtLaserPrinter.Appearance.BackColor = Drawing.Color.Yellow
                If Net.IPAddress.TryParse(ASCMAIN1.LaserPrinterIpAddress, Nothing) Then
                    txtLaserPrinter.Appearance.BackColor = Drawing.Color.Green
                End If
            End If

        Catch ex As Exception
            txtLaserPrinter.Appearance.BackColor = Drawing.Color.Yellow
            tooltip.SetToolTip(txtLaserPrinter, ex.Message)
        End Try


        '**************************
        '**    Label Printer Port
        '**************************        
        Try
            txtLabelPrinter.BackColor = Drawing.Color.Red

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing Then
                txtLabelPrinter.Text = ASCMAIN1.LabelPrinterSerialPort.PortName
                tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
            Else
                Me.txtLabelPrinter.Text = "No Port"
                tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
            End If

            txtLabelPrinter.BackColor = Drawing.Color.Yellow
            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing AndAlso Not ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                If Not ASCMAIN1.Running_in_VS Then ASCMAIN1.LabelPrinterSerialPort.Open()
            End If

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing AndAlso ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                txtLabelPrinter.BackColor = Drawing.Color.Green
            End If

        Catch ex As Exception
            txtLabelPrinter.BackColor = Drawing.Color.Red
            tooltip.SetToolTip(txtLabelPrinter, ex.Message)
        End Try

        '**************************
        '**    Scale Port
        '************************** 
        'ASCMAIN1.ScaleWeightDelegate = AddressOf ProcessScaleData
        Try
            txtScale.BackColor = Drawing.Color.Red

            'If ASCMAIN1.ScaleSerialPort IsNot Nothing Then
            '    txtScale.Text = ASCMAIN1.ScaleSerialPort.PortName
            '    tooltip.SetToolTip(txtScale, txtScale.Text)
            'Else
            '    txtScale.Text = "No Port"
            '    tooltip.SetToolTip(txtScale, txtScale.Text)
            'End If

            'txtScale.BackColor = Drawing.Color.Yellow
            'If ASCMAIN1.ScaleSerialPort IsNot Nothing AndAlso Not ASCMAIN1.ScaleSerialPort.IsOpen Then
            '    If Not ASCMAIN1.Running_in_VS Then ASCMAIN1.ScaleSerialPort.Open()
            'End If

            'If ASCMAIN1.ScaleSerialPort IsNot Nothing AndAlso ASCMAIN1.ScaleSerialPort.IsOpen Then
            '    txtScale.BackColor = Drawing.Color.Green
            'End If

        Catch ex As Exception
            txtScale.BackColor = Drawing.Color.Red
            tooltip.SetToolTip(txtScale, ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' Sends the Scanned Bar Code to the Appropriate Control
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ProcessScannedData(ByVal scannedData As String)

        If MdiParent.ActiveMdiChild Is Nothing Then Exit Sub
        If MdiParent.ActiveMdiChild.Name <> Me.Name Then Exit Sub

        Static dataReceived As String

        dataReceived += scannedData
        If InStr(dataReceived, Chr(13), CompareMethod.Text) = 0 Then
            Exit Sub
        End If

        Dim sender As Object = Nothing
        Dim e As New System.Windows.Forms.KeyEventArgs(Keys.Enter)

        ' Trim Off line feeds
        dataReceived = Replace(dataReceived, Chr(13), String.Empty)
        dataReceived = Replace(dataReceived, Chr(10), String.Empty)
        dataReceived = dataReceived.Trim

        ' Set Sender based on state of the screen

        ProcessEnterKeyStroke(sender, e)
        dataReceived = String.Empty
    End Sub

    ''' <summary>
    ''' Process keyboard 'Enter' key
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ProcessEnterKeyStroke(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

        Select Case Absx1.GetABSColumnName(sender)
            Case "x"

            Case "y"
        End Select
    End Sub

    Private Sub SOFSHIPL_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown

    End Sub

    ' ''' <summary>
    ' ''' Request weight from scale
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Private Sub RequestWeightFromScale()

    '    Try
    '        registeredWeight = 0

    '        If ASCMAIN1.ScaleSerialPort Is Nothing Then Exit Sub

    '        If Not ASCMAIN1.ScaleSerialPort.IsOpen Then
    '            ASCMAIN1.ScaleSerialPort.Open()
    '        End If

    '        ' Request the weight from the scale
    '        Dim encoding As New System.Text.UTF8Encoding()
    '        Dim inBuffer As Byte() = encoding.GetBytes("W")
    '        ASCMAIN1.ScaleSerialPort.Write(inBuffer, 0, inBuffer.Length)

    '    Catch ex As Exception
    '        MessageBox.Show("Scale Weight Error: " & ex.Message, "Scale Weight", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

    ' ''' <summary>
    ' ''' Fires when the weight is requested from the scale
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Private Sub ProcessScaleData(ByVal scaledata As String)

    '    If MdiParent.ActiveMdiChild Is Nothing Then Exit Sub
    '    If MdiParent.ActiveMdiChild.Name <> Me.Name Then Exit Sub

    '    Try
    '        Dim length As Int16 = ASCMAIN1.ScaleSerialPort.BytesToRead
    '        If length > 0 Then
    '            Dim numberOfBytesRead As Int16 = 0
    '            Dim readBuffer(length) As Byte
    '            numberOfBytesRead = ASCMAIN1.ScaleSerialPort.Read(readBuffer, 0, length)
    '            registeredWeight = Val(readBuffer)
    '        End If
    '    Catch ex As Exception

    '    End Try
    'End Sub

#End Region

End Class
