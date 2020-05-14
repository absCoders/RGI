Imports nsoftware.InShip

Public Class SOFSHIP0
    ' proceed prereq - after maintenance, or confirmation,need to verify that auth amt on credit check and credit card has not been violated
    Private expSOTPICK1 As New Dictionary(Of String, String)
    Private commonCarriersOnly As Boolean = False
    Private AddPPAToExistingFreight As Boolean = True

#Region "Declarations"
    Dim CUST_CODE As String
    Dim CUST_NAME As String         ' Sold-To Customer Name
    Dim SHIP_BOL_NOs As New List(Of String)
    Dim ORDR_GROUP_NO As String
    Dim ORDR_CUST_PO As String
    Dim rowARTCUST1 As DataRow
    Dim rowSOTSHIP0 As DataRow
    Dim rowSOTSHIP0_ORIG As DataRow
    Dim clsPrice_Change As Price_Change = Nothing
    Dim sqlSOTPICK1 As String
    Dim sqlSOTPICK2 As String
    Dim sqlSOTSHIPX As String
    Dim edi_customer As Boolean
    Dim edi856_customer As Boolean
    Dim edi_order As Boolean
    Dim ORDR_SOURCE As String
    Dim SOTSHIP0 As String
    Dim CURR_CODE As String
    Dim CURR_EXCH_RATE As Decimal
    Dim GST_TAX As Decimal
    Dim ASW As New Dictionary(Of String, String)
    Dim select_from_3PL_list As Boolean = False
    Dim processing_select_from_3PL_list As Boolean = False
    Dim selectedEDI_BOL_NO As String = String.Empty
    Dim selectedMasterEDI_BOL_NO As String = String.Empty

    Dim MaintenanceMode As Boolean = False

    Dim ORDR_SHIP_DATE As Date
    Dim ORDR_CANCEL_DATE As Date

    Dim dvwSOTORDR5 As DataView
    Dim dvwSOTORDR5_BT As DataView

    Dim SOTSHIPX As String

    Private commonCarrier As Boolean = False
    Private CreditCardProcessor As TAC.TAFCARDF
    Private RecreateLabel As Boolean = False
    Private refreshScreen As Boolean = True
    Private tblTATSTATE As DataTable
    Private EDI_DOC_SEQ_NOs As List(Of String) = New List(Of String)
    Private rowSOTMISC1 As DataRow = Nothing
    'Private clsASCSCRTY As New TAC.ASCSCRTY

    Private Const NonEdiCustomer As String = "3"
    Private Const EdiErrorRecord As String = "4"

    Private WithEvents ultraComboPackage As Infragistics.Win.UltraWinGrid.UltraCombo = New Infragistics.Win.UltraWinGrid.UltraCombo

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("GLTPARM1")
        Check_InquiryMode()

        SOTSHIP0 = ASCMAIN1.Temp_Table("Select SHIP_BOL_NO from SOTSHIP1 where ROWNUM < 1")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIP0 & " Add Primary Key (SHIP_BOL_NO)")

        If MaintenanceMode Then
            ASCMAIN1.sql = "Select SHIP_BOL_NO, '1' SEL, '1' EDI856, '1' SHIP_CART_REQD from SOTSHIP1 where ROWNUM < 1"
            SOTSHIPX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add Primary Key (SHIP_BOL_NO)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add SHIP_CHGREQ_NO VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add CUST_CODE VARCHAR2(10)")
        End If

        tblTATSTATE = ASCDATA1.GetDataTable("SELECT * FROM TATSTATE", "TATSTATE") ' WHERE region_code is not null

        With dst
            If 1 = 2 AndAlso ASCMAIN1.CLIENT = "NYA" Then
                commonCarriersOnly = True
                ' No 3PL and only UPS or Fedex for NYA
                sqlSOTSHIPX = "Select DISTINCT SOTSHIP1.*" _
                     & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE,SOTORDR0.ORDR_CUST_PO" _
                     & " from SOTSHIP1,SOTORDR0,ICTWHSE1,SOTSVIA1,SOTCARR1" _
                     & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" _
                     & " and SOTSHIP1.SHIP_VIA_CODE = SOTSVIA1.SHIP_VIA_CODE" & vbCrLf _
                     & " and SOTSVIA1.CARRIER_CODE = SOTCARR1.CARRIER_CODE" & vbCrLf _
                     & " and SOTCARR1.CARRIER_TYPE = 'U'" & vbCrLf _
                     & " and SOTSHIP1.WHSE_CODE = ICTWHSE1.WHSE_CODE"
            ElseIf ASCMAIN1.CLIENT = "RGI" Then
                sqlSOTSHIPX = "Select DISTINCT SOTSHIP1.*" _
                     & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE,SOTORDR0.ORDR_CUST_PO, ARTCUST1.CUST_NAME" _
                     & " from SOTSHIP1 ,SOTORDR0, ICTWHSE1, ARTCUST1, SOTORDR1" _
                     & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" _
                     & " and SOTSHIP1.WHSE_CODE = ICTWHSE1.WHSE_CODE" _
                     & " and SOTORDR0.CUST_CODE = ARTCUST1.CUST_CODE" _
                     & " and SOTORDR0.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO"
                If Not InquiryMode Then
                    sqlSOTSHIPX &= " and SOTORDR1.ECOM_CODE IS NULL"
                End If

            Else
                'VAN has multiple orders per order group, RGI only will link to SOTORDR1 above
                sqlSOTSHIPX = "Select DISTINCT SOTSHIP1.*" _
                     & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE,SOTORDR0.ORDR_CUST_PO, ARTCUST1.CUST_NAME" _
                     & " from SOTSHIP1 ,SOTORDR0, ICTWHSE1, ARTCUST1" _
                     & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" _
                     & " and SOTSHIP1.WHSE_CODE = ICTWHSE1.WHSE_CODE" _
                     & " and SOTORDR0.CUST_CODE = ARTCUST1.CUST_CODE"
            End If

            If Not InquiryMode Then
                sqlSOTSHIPX &= "   and SOTSHIP1.SHIP_STATUS = 'P'"

                If ASCMAIN1.CLIENT <> "RGI" Then
                    sqlSOTSHIPX &= "   and SOTSHIP1.SHIP_PICK_PRINTED is Not Null"
                End If
            End If

            If ASCMAIN1.CLIENT = "RGI" Then
                If Not InquiryMode Then
                    sqlSOTSHIPX &= " AND ICTWHSE1.WHSE_CODE = 'CG'"
                End If
            End If

            ASCMAIN1.sql = sqlSOTSHIPX
            Create_TDA(.Tables.Add, "SOTSHIPX", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            Create_TDA(.Tables.Add, "SOTORDR2", "*")
            .Tables("SOTORDR2").Columns.Add("ORDR_UNIT_PRICE_ORIG", GetType(System.Decimal))

            Create_TDA(.Tables.Add, "SOTORDC1", "*", 1)
            Create_TDA(.Tables.Add, "SOTORDC2", "*", 1)
            .Tables("SOTORDC2").Columns.Add("PICK_NO", GetType(System.String))

            ASCMAIN1.sql = sqlSOTSHIPX & " and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, True, "", 1)
            ASCMAIN1.sql = sqlSOTSHIPX & " and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTSHIP0", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "SOTINVH1", "*")
            Create_TDA(.Tables.Add, "SOTINVH2", "*")
            Create_TDA(.Tables.Add, "SOTINVH9", "*")
            Create_TDA(.Tables.Add, "SOTINVHM", "*")

            Create_TDA(.Tables.Add, "ARTOPEN1", "*")
            Create_TDA(.Tables.Add, "ICTWHSE1", "*")
            Create_TDA(.Tables.Add, "SOTSHIPB", "*")

            ASCMAIN1.sql = "SELECT SOTSVIA1.*, SOTCARR1.CARRIER_TYPE" _
                & " FROM SOTSVIA1, SOTCARR1" _
                & " WHERE SOTSVIA1.CARRIER_CODE = SOTCARR1.CARRIER_CODE"
            Create_TDA(.Tables.Add, "SOTSVIA1", "**", 1, False, "", 1)
            Fill_Records("SOTSVIA1", "", True, ASCMAIN1.sql)

            ' Credit Card Processing
            Create_TDA(.Tables.Add, "ARTCCPA1", "*")
            Create_TDA(.Tables.Add, "ARTCCPA2", "*")
            Create_TDA(.Tables.Add, "ARTCCPDA", "*")

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

            Fill_Records("SOTCARR1", "", True, "SELECT * FROM SOTCARR1")
            Fill_Records("SOTCARR2", "", True, "SELECT * FROM SOTCARR2")
            Fill_Records("SOTCARR3", "", True, "Select SOTCARR3.*, SOTCARR1.CARRIER_DESC From SOTCARR3, SOTCARR1 Where SOTCARR3.CARRIER_CODE = SOTCARR1.CARRIER_CODE")

            Create_TDA(.Tables.Add, "SOTCART1", "*")
            .Tables("SOTCART1").Columns.Add("INSURANCE", GetType(System.Decimal))
            .Tables("SOTCART1").Columns("INSURANCE").DefaultValue = 0

            .Tables("SOTCART1").Columns.Add("REFERENCE1", GetType(System.String))
            .Tables("SOTCART1").Columns("REFERENCE1").DefaultValue = String.Empty
            .Tables("SOTCART1").Columns("REFERENCE1").MaxLength = 20

            .Tables("SOTCART1").Columns.Add("REFERENCE2", GetType(System.String))
            .Tables("SOTCART1").Columns("REFERENCE2").DefaultValue = String.Empty
            .Tables("SOTCART1").Columns("REFERENCE2").MaxLength = 20

            .Tables("SOTCART1").Columns.Add("WIDTH", GetType(System.Int16))
            .Tables("SOTCART1").Columns("WIDTH").DefaultValue = 0
            .Tables("SOTCART1").Columns.Add("LENGTH", GetType(System.Int16))
            .Tables("SOTCART1").Columns("LENGTH").DefaultValue = 0
            .Tables("SOTCART1").Columns.Add("HEIGHT", GetType(System.Int16))
            .Tables("SOTCART1").Columns("HEIGHT").DefaultValue = 0

            ASCMAIN1.sql = "Select SOTCART2.*, SOTCART1.PICK_NO, SOTCART2.QTY_PACKED QTY_PACKED_ORIG" & vbCrLf _
              & " from SOTCART2,SOTCART1 where SOTCART1.CART_NO = SOTCART2.CART_NO"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0)

            Create_TDA(.Tables.Add, "SOTCART3", "*")

            ASCMAIN1.sql = "Select SOTORDR9.* " & vbCrLf _
                & " from SOTORDR9, SOTORDR1" & vbCrLf _
                & " where SOTORDR9.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR9", "**", 0, True, "V", 2)

            Create_TDA(.Tables.Add, "SOTRNGA1", "*")

            'ASCMAIN1.sql = "Select CUST_CODE, EDI_DOC_NO, EDI_STATUS" _
            '    & " from EDTTRPM1" _
            '    & " where CUST_CODE = :PARM1 AND COMPANY_CODE = '" & ASCMAIN1.CLIENT & "'"

            ASCMAIN1.sql = " Select EDTTRPM1.CUST_CODE, EDTTRPM1.EDI_DOC_NO, EDTTRPM1.EDI_STATUS" _
             & "  from EDTTRPM1, EDTTRPMC" _
             & "  where EDTTRPM1.EDI_TP_ID = EDTTRPMC.EDI_TP_ID" _
             & "  and EDTTRPM1.EDI_OUR_ID = EDTTRPMC.EDI_OUR_ID" _
             & "  and EDTTRPMC.COMPANY_CODE = '" & ASCMAIN1.CLIENT & "'" _
             & "  and EDTTRPM1.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "EDTTRPMC", "**", 0, False, "V", 2)
            '.Tables("EDTTRPM1").PrimaryKey = New DataColumn() {.Tables("EDTTRPM1").Columns("CUST_CODE"), _
            '                                                   .Tables("EDTTRPM1").Columns("EDI_DOC_NO")}

            With .Tables.Add("SOTCARTX")
                .Columns.Add("PICK_NO")
                .Columns.Add("ORDR_NO")
                .Columns.Add("ORDR_LNO", GetType(System.Int64))
                .Columns.Add("PICK_QTY_CONF", GetType(System.Int64), "")
                .Columns.Add("QTY_PACKED", GetType(System.Int64), "")
                .PrimaryKey = New DataColumn() {.Columns("PICK_NO"), .Columns("ORDR_NO"), .Columns("ORDR_LNO")}
            End With

            sqlSOTPICK1 = "Select SOTPICK1.*" & vbCrLf _
                & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_SOURCE" & vbCrLf _
                & ", SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.CUST_BILL_TO_CUST, SOTORDR1.ORDR_TYPE_CODE" & vbCrLf _
                & ", SOTORDR1.POST_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_FOB" & vbCrLf _
                & ", SOTORDR1.TERM_CODE, SOTORDR1.SREP_CODE, SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
                & ", SOTSHIP1.BILL_OF_LADING_NO, SOTORDR1.ORDR_INV_COMMENT, SOTORDR1.CUST_FACTOR_IND, NVL(SOTPICK1.CCPA_NO_AUTH, SOTORDR1.CCPA_NO) CCPA_NO_ORDR" & vbCrLf _
                & " from SOTPICK1,SOTORDR1,SOTSHIP1 "
            ASCMAIN1.sql = sqlSOTPICK1 & " where ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**")
            dst.Tables("SOTPICK1").Columns.Add("SELECTED")
            dst.Tables("SOTPICK1").Columns.Add("OUR_FREIGHT", GetType(System.Decimal))
            dst.Tables("SOTPICK1").Columns("OUR_FREIGHT").DefaultValue = 0
            dst.Tables("SOTPICK1").Columns.Add("INV_MISC_CHG", GetType(System.Decimal))
            dst.Tables("SOTPICK1").Columns("INV_MISC_CHG").DefaultValue = 0


            For Each fieldname As String In New String() {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3", _
                                                          "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY", _
                                                          "CUST_CONTACT", "CUST_PHONE"}
                dst.Tables("SOTPICK1").Columns.Add(fieldname, GetType(System.String))
            Next

            Create_Relation("SOTSHIP1", "SOTPICK1", "SHIP_BOL_NO")

            sqlSOTPICK2 = "Select SOTPICK2.*, " & vbCrLf _
                & " SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_DESC, SOTORDR2.CARTON_PACK_QTY," & vbCrLf _
                & " SOTORDR2.ORDR_UNIT_PRICE, SOTORDR2.STYLE_CODE_SUB," & vbCrLf _
                & " SOTORDR2.RANGE_STYLE_CODE, SOTORDR2.RANGE_STYLE_LNO, SOTORDR2.QTY_PER_PP, ICTSTYL1.CASE_CUBE"

            If ASCMAIN1.CLIENT = "RGI" OrElse ASCMAIN1.CLIENT = "NYA" Then
                sqlSOTPICK2 &= ", SOTORDR2.ORDR_PRICE_SOURCE, SOTORDR2.COMM_RATE " & vbCrLf
            End If

            sqlSOTPICK2 &= " from SOTPICK2, SOTPICK1, SOTORDR2, SOTSHIP1, ICTSTYL1" & vbCrLf
            ASCMAIN1.sql = sqlSOTPICK2 & " where ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "SOTPICK2", "**")
            With .Tables("SOTPICK2").Columns
                .Add("PICK_AMT", GetType(System.Decimal))
                .Add("PICK_AMT_CONF", GetType(System.Decimal))
                .Add("PICK_AMT_CANC", GetType(System.Decimal))
                .Add("PICK_AMT_BACK", GetType(System.Decimal))

                .Add("TOTAL_CARTONS", GetType(System.Decimal), "IIF(ISNULL(CARTON_PACK_QTY,0)=0,0,ISNULL(PICK_QTY_CONF,0) / ISNULL(CARTON_PACK_QTY,0))")
                .Add("TOTAL_CUBE", GetType(System.Decimal), "ISNULL(TOTAL_CARTONS,0) * ISNULL(CASE_CUBE,0)")

            End With

            Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")
            With .Tables("SOTPICK1").Columns
                .Add("PICK_QTY", GetType(System.Int64))
                .Add("PICK_QTY_CONF", GetType(System.Int64))
                .Add("PICK_QTY_CANC", GetType(System.Int64))
                .Add("PICK_QTY_BACK", GetType(System.Int64))
                .Add("PICK_AMT", GetType(System.Decimal))
                .Add("PICK_AMT_CONF", GetType(System.Decimal))
                .Add("PICK_AMT_CANC", GetType(System.Decimal))
                .Add("PICK_AMT_BACK", GetType(System.Decimal))

                .Add("TOTAL_CARTONS", GetType(System.Decimal), "SUM(CHILD.TOTAL_CARTONS)")
                .Add("TOTAL_CUBE", GetType(System.Decimal), "SUM(CHILD.TOTAL_CUBE)")

            End With

            Create_TDA(.Tables.Add, "SOTPICK4", "*", 1)
            dst.Tables("SOTPICK4").Columns.Add("PICK_QTY_USED", GetType(System.Int16))
            dst.Tables("SOTPICK4").Columns("PICK_QTY_USED").DefaultValue = 0

            ASCMAIN1.sql = "Select ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE" & vbCrLf _
                & ", ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP" & vbCrLf _
                & ", ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST" & vbCrLf _
                & ", 'N' CONSUMED, 0.00 CUM_QTY" & vbCrLf _
                & " FROM ICTCOST1" & vbCrLf _
                & " WHERE ROWNUM < 0" & vbCrLf
            Create_TDA(.Tables.Add, "ICTCOST1", "**", 0, False)

            Create_Relation("SOTCART1", "SOTCART2", "CART_NO")
            Create_Relation("SOTCARTX", "SOTPICK2", "PICK_NO,ORDR_NO,ORDR_LNO")
            Create_Relation("SOTCARTX", "SOTCART2", "PICK_NO,ORDR_NO,ORDR_LNO")

            dst.Tables("SOTCART1").Columns.Add("CART_TOTAL_UNITS_CALC", GetType(System.Int64))
            dst.Tables("SOTCART1").Columns.Add("CART_TOTAL_UNITS_ORIG", GetType(System.Int64), "SUM(CHILD.QTY_PACKED_ORIG)")

            Create_Relation("SOTPICK1", "SOTCART1", "PICK_NO")
            dst.Tables("SOTPICK1").Columns.Add("PICK_TOTAL_WGT_CALC", GetType(System.Decimal))
            dst.Tables("SOTPICK1").Columns.Add("PICK_CNT_CARTONS_CALC", GetType(System.Int64))
            dst.Tables("SOTPICK1").Columns.Add("PICK_TOTAL_UNITS_CALC", GetType(System.Int64))

            With .Tables.Add("SOTCONFT")
                .Columns.Add("KEY", GetType(System.Int32))
                .Columns.Add("STATUS")
                .Columns.Add("QTY", GetType(System.Int32))
                .Columns.Add("AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() {.Columns("KEY")}
            End With


            Create_WHT3PLS1()

            Create_TDA(.Tables.Add, "SOTSHIP3", "*")
            Create_TDA(.Tables.Add, "SOTSHIP4", "*")
            Create_TDA(.Tables.Add, "SOTSHIP6", "*")

            Create_TDA(.Tables.Add, "SOTORDR5", "*", 1)
            .Tables("SOTORDR5").Columns("CUST_ADDR_CODE").MaxLength = 10

            Create_TDA(.Tables.Add("SOTORDR5_BT"), "SOTORDR5", "*", 1)
            .Tables("SOTORDR5_BT").Columns("CUST_ADDR_CODE").MaxLength = 10

            Create_TDA(.Tables.Add, "TATEVNT1", "*")
            Create_TDA(.Tables.Add, "SOTORDXR", "*")
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

        grdSOTSHIPX.DataSource = dst.Tables("SOTSHIPX")
        grdSOTCONFT.DataSource = dst.Tables("SOTCONFT")
        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTPICK2.DataSource = dst.Tables("SOTPICK2")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")
        grdSOTCART2.DataSource = dst.Tables("SOTCART2")
        grdSOTSHIP1.DataSource = dst.Tables("SOTSHIP1")
        grdWHT3PLS1.DataSource = dst.Tables("WHT3PLS1")
        grdSOTINVHM.DataSource = dst.Tables("SOTINVHM")

        dvwSOTORDR5 = New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'ST'", "", DataViewRowState.CurrentRows)
        Bind_Controls(grpSHIPTO, "SOTORDR5", dvwSOTORDR5)

        dvwSOTORDR5_BT = New DataView(dst.Tables("SOTORDR5_BT"), "CUST_ADDR_TYPE = 'BT'", "", DataViewRowState.CurrentRows)
        Bind_Controls(grpBillTo, "SOTORDR5_BT", dvwSOTORDR5_BT)


        grdSOTSHIPX.DisplayLayout.UseFixedHeaders = True
        With grdSOTSHIPX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        grdWHT3PLS1.DisplayLayout.UseFixedHeaders = True
        With grdWHT3PLS1.DisplayLayout.Bands(0)
            Dim headerFields As String() = {"SHIP_BOL_NO", "CUST_CODE", "ORDR_CUST_PO"}
            If dst.Tables("WHT3PLS1").Columns.Contains("CUST_NAME") Then
                headerFields = {"SHIP_BOL_NO", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO"}
            End If
            For Each COLUMN_NAME As String In headerFields
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        grdSOTPICK1.DisplayLayout.UseFixedHeaders = True
        With grdSOTPICK1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SELECTED", "PICK_NO", "CUST_STORE_NO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With
        With grdSOTPICK1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ' "SELECTED",
                If Not MaintenanceMode And New String() {"PICK_CNT_CARTONS", "PICK_FREIGHT", "PICK_TOTAL_WGT", "ORDR_INV_COMMENT"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    If New String() {"SELECTED"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.DarkGoldenrod
                    Else
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    End If
                ElseIf New String() {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"PICK_AMT", "PICK_AMT_CONF", "PICK_AMT_CANC", "PICK_AMT_BACK"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    'gcol.Format = "#,##0.00"
                ElseIf New String() {"TOTAL_CARTONS", "TOTAL_CUBE"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Salmon
                    'gcol.Format = "#,##0.00"
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
            If MaintenanceMode Then
                .Columns("PICK_QTY_CONF").Header.Caption = "uRevised"
                .Columns("PICK_AMT_CONF").Header.Caption = "$Revised"
            End If
        End With

        grdSOTPICK2.DisplayLayout.UseFixedHeaders = True
        With grdSOTPICK2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"PICK_LNO", "STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSOTPICK2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If Not MaintenanceMode And New String() {"PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf MaintenanceMode And New String() {"PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If

                If New String() {"TOTAL_CARTONS", "TOTAL_CUBE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Salmon
                End If
            Next

            If MaintenanceMode Then
                .Columns("PICK_QTY_CONF").Header.Caption = "Revised"
                .Columns("PICK_QTY_BACK").Hidden = True
            End If
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
                If New String() {"CART_FREIGHT", "CART_TOTAL_WGT_ACTUAL", "PKG_CODE", "PACKAGING_TYPE", "INSURANCE", _
                                 "WIDTH", "LENGTH", "HEIGHT", "CART_SEQ", "REFERENCE1", "REFERENCE2"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        grdSOTCART2.DisplayLayout.UseFixedHeaders = True
        With grdSOTCART2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CART_LNO", "STYLE_CODE", "COLOR_CODE", "QTY_PACKED"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSOTCART2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"QTY_PACKED"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        grdSOTSHIP1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdSOTSHIP1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdSOTSHIP1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        For i As Integer = 0 To grdSOTSHIP1.DisplayLayout.Bands.Count - 1
            For Each grdcol As Infragistics.Win.UltraWinGrid.UltraGridColumn In grdSOTSHIP1.DisplayLayout.Bands(i).Columns
                grdcol.CellActivation = UltraWinGrid.Activation.NoEdit

                grdcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                grdcol.Header.Appearance.BackColor = Drawing.Color.White

                If New String() {"HANDLING_TYPE", "HANDLING_UNITS"}.Contains(grdcol.Key) Then
                    If Not InquiryMode AndAlso Not MaintenanceMode Then
                        grdcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    End If
                    grdcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                End If
            Next
        Next
        ASCMAIN1.Add_Value_List(grdSOTSHIP1, "HANDLING_TYPE", Nothing, New String() {":", "P:Pallet", "C:Carton", "S:Skid", "T:Totes", "L:Loose", "O:Other"})

        Create_Summary(grdSOTSHIPX, "SHIP_BOL_NO", "Count")
        Create_Summary(grdWHT3PLS1, "SHIP_BOL_NO", "Count")

        Create_Summary(grdSOTPICK2, "PICK_LNO", "Count")
        Create_Summary(grdSOTPICK2, New String() _
            {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK", "TOTAL_CARTONS", "TOTAL_CUBE"})

        Create_Summary(grdSOTCART1, "CART_NO", "Count")
        Create_Summary(grdSOTCART1, New String() _
            {"CART_FREIGHT", "CART_TOTAL_UNITS", "CART_TOTAL_WGT_ACTUAL"})

        Create_Summary(grdSOTINVHM, "INV_MNO", "Count")
        Create_Summary(grdSOTINVHM, New String() {"INV_MISC_CHG"})


        With dst.Tables("SOTCONFT").Rows
            .Add(New Object() {1, "Reld", 0, 0})
            .Add(New Object() {2, IIf(MaintenanceMode, "Revd", "Conf"), 0, 0})
            .Add(New Object() {3, "Canc", 0, 0})
            .Add(New Object() {4, "Back", 0, 0})
            .Add(New Object() {5, "Frt", 0, 0})
            .Add(New Object() {6, "Misc", 0, 0})
        End With

        Sort_grdColumns(grdSOTCONFT, "KEY", True)

        Show_Filter(grdSOTSHIPX, True)
        grdSOTSHIPX.DisplayLayout.GroupByBox.Hidden = False

        ASCMAIN1.Add_Value_List(grdSOTSHIP1, "PICK_STATUS", Nothing, Nothing, 1)


        Show_Filter(grdWHT3PLS1, True)
        grdWHT3PLS1.DisplayLayout.GroupByBox.Hidden = False

        calFrom.Value = Now.Date.AddDays(-30)
        calTo.Value = Now.Date

        ASCMAIN1.Add_Value_List(grdSOTPICK1, "PICK_STATUS", Nothing, New String() {":", "P:In Pick", "F:Shipped", "D:Deleted", "C:Cancelled"})

        Position_txtSTORE()
        '  SplitContainer1.Panel2Collapsed = True

        If MaintenanceMode Then
            tabSelect.Tabs("3PL Shipments").Visible = False
            splHeader.Panel1Collapsed = True
            chkBO.Visible = False
        Else
            lblReason.Visible = False
            txtReason.Visible = False
            lblContact.Visible = False
            txtContact.Visible = False
            lblemail.Visible = False
            txtemail.Visible = False
        End If

        If InquiryMode OrElse Not (ASCMAIN1.CLIENT = "NYA" Or ASCMAIN1.CLIENT = "RGI") Then
            tabSelect.Tabs("3PL Shipments").Visible = False
        End If

        If ASCMAIN1.CLIENT = "RGI" Then
            lblAddressType.Visible = False
            optAddressType.Visible = False
            chkSaturday.Visible = False
            chkSignature.Visible = False
        End If

        UltraExplorerBar1.Groups("Devices").Visible = ASCMAIN1.CLIENT = "NYA"

        If Not InquiryMode Then
            tabSelect.Tabs(0).Text = "Unconfirmed Shipments"
            grdSOTINVHM.DisplayLayout.Bands(0).Columns("INV_TYPE").Hidden = True
            grdSOTINVHM.DisplayLayout.Bands(0).Columns("INV_NO").Hidden = True
        Else
            tabSelect.Tabs(0).Text = optStatus.CheckedItem.DisplayText
            grdSOTINVHM.DisplayLayout.Bands(0).Columns("INV_TYPE").Hidden = False
            grdSOTINVHM.DisplayLayout.Bands(0).Columns("INV_NO").Hidden = False
        End If

        Bind_Controls(grpShippingWindow, "SOTSHIP1")
        splHeaderInfo.Panel2Collapsed = ASCMAIN1.CLIENT = "RGI"

    End Sub

    Sub Check_InquiryMode()
        InquiryMode = (MENU_ITEM_OBJECT = "SOFSHIPI") OrElse (MENU_ITEM_OBJECT = "SOFSHIPR")
        MaintenanceMode = (MENU_ITEM_OBJECT = "SOFSHIPM")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            'Case "Fix Transfers"
            '    Stop

            '    ASCMAIN1.sql = "SELECT XFR_NO FROM ICTIXFR1 " _
            '        & " where WHSE_CODE = '91' AND WHSE_CODE_TO = '02' AND XFR_SOURCE = 'S'"

            '    For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            '        Dim XFR_NO As String = row.Item("XFR_NO")
            '        ASCMAIN1.sql = "Select ICTIXFR2.XFR_NO WHSE_TRAN_NO, ICTIXFR2.XFR_LNO WHSE_TRAN_LNO" _
            '            & ", 'T' WHSE_TRAN_TYPE, ICTIXFR1.WHSE_CODE_TO WHSE_CODE" _
            '            & ", ICTWHSE1.WHSE_LOC_REC LOCATION_CODE, ICTIXFR2.STYLE_CODE, ICTIXFR2.COLOR_CODE" _
            '            & ", ICTIXFR2.XFR_QTY WHSE_TRAN_QTY" _
            '            & " from ICTIXFR1,ICTIXFR2,ICTWHSE1" _
            '            & " where ICTIXFR1.XFR_NO = ICTIXFR2.XFR_NO" _
            '            & "   and ICTWHSE1.WHSE_CODE = ICTIXFR1.WHSE_CODE_TO" _
            '            & "   and ICTIXFR2.XFR_NO = '" & XFR_NO & "'"
            '        WHCMAIN1.Update_WHTLOCBX(Me)
            '    Next

            '    EMsg &= vbCr & "Done"

            Case "Select"

                ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIP0)
                SHIP_BOL_NOs.Clear()
                EDI_DOC_SEQ_NOs.Clear()
                select_from_3PL_list = False

                Dim SHIP_STATUS As String = ""

                If Absx1.txtFor("SHIP_BOL_NO").Text = "" Then
                    EMsg &= vbCr & "You must First Select a Shipment No"
                Else
                    Dim SHIP_BOL_NO As String = Absx1.txtFor("SHIP_BOL_NO").Text
                    Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
                    If rowSOTSHIP1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Shipment No (" & SHIP_BOL_NO & ")"
                    Else
                        SHIP_STATUS = rowSOTSHIP1.Item("SHIP_STATUS")

                        If SHIP_STATUS <> "P" And MENU_ITEM_OBJECT <> "SOFSHIPI" Then

                            If Not processing_select_from_3PL_list Then
                                EMsg = vbCr & "The provided shipment does not have a status of 'In Pick'. You must use Shipment Confirmation Inquiry screen to view this shipment."
                                Exit Select
                            Else
                                EMsg = "The provided shipment does not have a status of 'In Pick'. You must use Shipment Confirmation Inquiry screen to view this shipment." _
                                    & Environment.NewLine & Environment.NewLine _
                                    & "Do you want to disregard this EDI 945?"

                                If MessageBox.Show(EMsg, "Disregard EDI 945", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                                    If selectedEDI_BOL_NO.Length > 0 Then
                                        ASCDATA1.ExecuteSQL("Update EDT945T1 SET EDI_PROCESS_IND = 'D' where EDI_BOL_NO = '" & selectedEDI_BOL_NO & "' and EDI_PROCESS_IND is null")

                                        Fill_Records("WHT3PLS1")
                                        Sort_grdColumns(grdWHT3PLS1, "SHIP_BOL_NO")
                                        If ASCMAIN1.CLIENT = "NYA" Then
                                            Dim SHIP_BOL_NOs As String = String.Empty
                                            ' 5/23/2013 - At some point in time the release of orders to the 3PL stopped populating this field. Taking care of my customer.
                                            For Each row As DataRow In ASCDATA1.SelectDistinct("WHT3PLS1", New String() {"SHIP_BOL_NO"}).Rows
                                                SHIP_BOL_NOs &= ", '" & row.Item("SHIP_BOL_NO") & "'"
                                            Next
                                            If SHIP_BOL_NOs.Length > 0 Then
                                                SHIP_BOL_NOs = SHIP_BOL_NOs.Substring(1)
                                                Try
                                                    ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_PICK_PRINTED = SYSDATE WHERE SHIP_PICK_PRINTED IS NULL " _
                                                        & " AND SHIP_BOL_NO IN (" & SHIP_BOL_NOs & ")"
                                                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                                                Catch ex As Exception

                                                End Try
                                            End If

                                        End If
                                    End If

                                End If
                                processing_select_from_3PL_list = False
                                Exit Sub
                            End If
                        End If

                        ORDR_GROUP_NO = rowSOTSHIP1.Item("ORDR_GROUP_NO")
                        If Not InquiryMode Then If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
                        Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)
                        CUST_CODE = rowSOTORDR0.Item("CUST_CODE")
                        ORDR_CUST_PO = rowSOTORDR0.Item("ORDR_CUST_PO") & ""

                        Dim whse_code As String = rowSOTORDR0.Item("WHSE_CODE") & ""
                        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", whse_code)

                        If rowICTWHSE1 Is Nothing Then
                            EMsg &= vbCr & "Cannot determine warehouse for the selected shipment."
                            Exit Select
                        ElseIf Not InquiryMode Then
                            select_from_3PL_list = rowICTWHSE1.Item("LP_CODE" & String.Empty).ToString.Trim.Length > 0 OrElse processing_select_from_3PL_list
                            If ASCMAIN1.CLIENT = "RGI" Then
                                select_from_3PL_list = False
                            End If
                            optShipmentSelection.Value = "G"
                        End If

                        If ASCMAIN1.CLIENT = "RGI" AndAlso rowICTWHSE1.Item("WHSE_LOCATOR") & String.Empty = "1" And Not InquiryMode Then
                            EMsg &= vbCr & "You are not permitted to process shipments for a Warehouse setup to use a Locator system."
                            Exit Select
                        End If

                        If optShipmentSelection.Value = "S" Then
                            If Not InquiryMode Then If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub
                            If Not SHIP_BOL_NOs.Contains(SHIP_BOL_NO) Then
                                SHIP_BOL_NOs.Add(SHIP_BOL_NO)
                                ASCDATA1.ExecuteSQL("Insert into " & SOTSHIP0 & " (SHIP_BOL_NO) values ('" & SHIP_BOL_NO & "')")
                            End If
                        Else
                            ASCMAIN1.sql = "Select SOTPICK1.SHIP_BOL_NO, SOTPICK1.PICK_NO from SOTSHIP1, SOTPICK1 " _
                                & " where SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" _
                                & " and SOTSHIP1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
                                & " and SOTSHIP1.SHIP_STATUS = '" & SHIP_STATUS & "'" _
                                & " and SOTPICK1.PICK_STATUS = '" & SHIP_STATUS & "'" _
                                & " and SOTSHIP1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "SHIP_BOL_NO, PICK_NO")
                                Dim SHIP_BOL_NO2 As String = row.Item("SHIP_BOL_NO")
                                If select_from_3PL_list Then
                                    If Not SHIP_BOL_NOs.Contains(SHIP_BOL_NO2) Then
                                        'Dim rowWHT3PLS1 As DataRow = dst.Tables("WHT3PLS1").Rows.Find(SHIP_BOL_NO2)
                                        If dst.Tables("WHT3PLS1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO2 & "'") Is Nothing Then
                                            EMsg &= vbCr & "The following Shipment: " & SHIP_BOL_NO2 & " is part of the Order Group for the selected Shipment; however, that shipment does not appear in the list."
                                            Exit For
                                        End If
                                    End If

                                    If ASCMAIN1.CLIENT = "NYA" Then

                                        If dst.Tables("WHT3PLS1").Select("PICK_NO = '" & row.Item("PICK_NO") & "'").Length = 0 Then
                                            EMsg &= vbCr & "Pick Ticket (" & row.Item("PICK_NO") & ") belongs to the selected shipment; however, the EDI 945 is missing"
                                            Exit For
                                        End If

                                        If selectedEDI_BOL_NO.Length > 0 Then
                                            If dst.Tables("WHT3PLS1").Select("PICK_NO = '" & row.Item("PICK_NO") & "' and EDI_BOL_NO = '" & selectedEDI_BOL_NO & "'").Length > 0 Then
                                                EDI_DOC_SEQ_NOs.Add(dst.Tables("WHT3PLS1").Select("PICK_NO = '" & row.Item("PICK_NO") & "' and EDI_BOL_NO = '" & selectedEDI_BOL_NO & "'")(0).Item("EDI_DOC_SEQ_NO"))
                                            ElseIf selectedMasterEDI_BOL_NO.Length > 0 AndAlso _
                                                dst.Tables("WHT3PLS1").Select("PICK_NO = '" & row.Item("PICK_NO") & "' and EDI_MASTER_BOL_NO = '" & selectedMasterEDI_BOL_NO & "'").Length > 0 Then
                                                EDI_DOC_SEQ_NOs.Add(dst.Tables("WHT3PLS1").Select("PICK_NO = '" & row.Item("PICK_NO") & "' and EDI_MASTER_BOL_NO = '" & selectedMasterEDI_BOL_NO & "'")(0).Item("EDI_DOC_SEQ_NO"))
                                            Else
                                                EMsg &= vbCr & "Pick Ticket (" & row.Item("PICK_NO") & ") belongs to the selected shipment; however, it does not have the same BOL No."
                                                Exit For
                                            End If
                                        Else
                                            If dst.Tables("WHT3PLS1").Select("PICK_NO = '" & row.Item("PICK_NO") & "' and ISNULL(EDI_BOL_NO, '') = ''").Length > 0 Then
                                                EDI_DOC_SEQ_NOs.Add(dst.Tables("WHT3PLS1").Select("PICK_NO = '" & row.Item("PICK_NO") & "' and ISNULL(EDI_BOL_NO, '') = ''")(0).Item("EDI_DOC_SEQ_NO"))
                                            Else
                                                Continue For
                                            End If
                                        End If
                                    Else
                                        If dst.Tables("WHT3PLS1").Select("PICK_NO = '" & row.Item("PICK_NO") & "'").Length <> 0 Then
                                            EDI_DOC_SEQ_NOs.Add(dst.Tables("WHT3PLS1").Select("PICK_NO = '" & row.Item("PICK_NO") & "'")(0).Item("EDI_DOC_SEQ_NO"))
                                        End If
                                    End If
                                End If

                                If Not InquiryMode AndAlso Not SHIP_BOL_NOs.Contains(SHIP_BOL_NO2) Then
                                    If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO2) Then
                                        Exit Sub
                                    End If
                                End If

                                If Not SHIP_BOL_NOs.Contains(SHIP_BOL_NO2) Then
                                    Me.SHIP_BOL_NOs.Add(SHIP_BOL_NO2)
                                    ASCDATA1.ExecuteSQL("Insert into " & SOTSHIP0 & " (SHIP_BOL_NO) values ('" & SHIP_BOL_NO2 & "')")
                                End If
                            Next
                        End If
                    End If
                End If

                ' allow user to call up a previously billed shipment - need to look at some other variable

                If EMsg = "" Then
                    For Each SHIP_BOL_NO As String In SHIP_BOL_NOs
                        Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
                        If Not InquiryMode Then
                            If rowSOTSHIP1.Item("SHIP_STATUS") <> SHIP_STATUS Then
                                EMsg &= vbCr & "Shipment Status Changed for Shipment " & SHIP_BOL_NO
                            End If
                            If rowSOTSHIP1.Item("SHIP_BOL_NO_REV") & "" <> "" Then
                                EMsg &= vbCr & "Shipment " & SHIP_BOL_NO & " is a Part of a Shipment/Invoice Reversal"
                            End If
                            If SHIP_STATUS = "P" Then
                                ' Allow 3PL Pick Tickets to not be printed.
                                If rowSOTSHIP1.Item("SHIP_PICK_PRINTED") & "" = "" Then
                                    If ASCMAIN1.CLIENT = "RGI" Then
                                        rowSOTSHIP1.Item("SHIP_PICK_PRINTED") = DateTime.Now
                                        ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_PICK_PRINTED = SYSDATE WHERE SHIP_PICK_PRINTED IS NULL  AND SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "'"
                                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                                    Else
                                        EMsg &= vbCr & "Pick Tickets have not been Printed (yet) for Shipment " & SHIP_BOL_NO
                                    End If
                                End If
                                If ASCMAIN1.CLIENT = "VAN" Then
                                    Dim rowSOTCTLU1 As DataRow = LookUp("SOTCTLU1", "Z")
                                    If rowSOTCTLU1.Item("CTL_UPDATE_REQ") & "" = "D" Then
                                        EMsg &= vbCr & "There Has Been A De-Confirm that has not been updated by the Sales Journal." _
                                            & "Please Run Sales Journal Before Proceeding"
                                    End If
                                End If
                            Else
                                Select Case SHIP_STATUS
                                    Case "F"
                                        If ASCMAIN1.CLIENT = "VAN" Then
                                            Dim rowSOTCTLU1 As DataRow = LookUp("SOTCTLU1", "Z")
                                            If rowSOTCTLU1.Item("CTL_UPDATE_REQ") & "" = "C" Then
                                                MsgBox("There Has Been A Confirm that has not been updated by the Sales Journal." _
                                                       & "Please Run Sales Journal Before Proceeding", _
                                                       MsgBoxStyle.OkOnly, "Sales Journal Update Required First")
                                                Exit Sub
                                            End If
                                        End If

                                        ASCMAIN1.sql = "Select" _
                                            & "  Sum (DECODE(ORDR_YYYYPP_UPDATED,NULL,1,0)) PENDING" _
                                            & ", Sum (DECODE(ORDR_YYYYPP_UPDATED,NULL,0,1)) UPDATED" _
                                            & " FROM SOTINVH1 WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" _
                                            & " and INV_NO_REV is Null"
                                        Dim row As DataRow = ASCDATA1.GetDataRow
                                        If Val(row.Item("UPDATED") & "") = 0 And (rowSOTSHIP1.Item("REGISTER_XNO") & "") = "" Then
                                            If MsgBox(CStr(row.Item("PENDING")) & " Pick Ticket(s) were Confirmed in this Shipment" _
                                                      & vbCrLf & vbCrLf _
                                                      & "Do you want to De-Confirm all Pick Tickets on this Shipment?" _
                                                      & vbCrLf & vbCrLf _
                                                      & "Warning - This feature performs the following: " _
                                                      & vbCrLf _
                                                      & "  1) Deletes Invoice Header & Details" & vbCrLf _
                                                      & "  2) Resets Pick Tickets to 'Unconfirmed'" & vbCrLf _
                                                      & "  3) Resets Shipment to 'Unconfirmed'" _
                                                      & vbCrLf & vbCrLf _
                                                      & "Note: You would not be offered this option if any of the Invoices associated with these Pick Tickets were Updated into the A/R", _
                                                      MsgBoxStyle.YesNo + MsgBoxStyle.Question, _
                                                      "Shipment " & SHIP_BOL_NO & " has been Confirmed as Shipped") = MsgBoxResult.Yes Then
                                                Stop ' SEE WJZ FOR TESTING
                                                De_Confirm(SHIP_BOL_NO)
                                            End If
                                            Exit Sub
                                        Else
                                            If (rowSOTSHIP1.Item("SHIP_810_BATCH_NO") & "" <> "N" _
                                            And rowSOTSHIP1.Item("SHIP_810_BATCH_NO") & "" <> "") _
                                            OrElse MaintenanceMode Then
                                                EMsg &= vbCr & "EDI Shipment " & SHIP_BOL_NO & " has been Confirmed as Shipped & Updated; No further Corrections are Permitted."
                                            Else
                                                ASCMAIN1.sql = "Select Count (*) from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                                                Dim PICKS As Int32 = Val(ASCDATA1.GetDataValue)

                                                If MsgBox("Shipment " & SHIP_BOL_NO & " has been Confirmed as Shipped & Updated" _
                                                          & vbCrLf & vbCrLf _
                                                          & "Do you want to Reverse Invoices Generated for all " & CStr(PICKS) & " Pick Tickets on this Shipment?" _
                                                          & vbCrLf & vbCrLf _
                                                          & "Warning - This feature performs the following: " & vbCrLf _
                                                          & "  1) Creates Negative Invoices" & vbCrLf _
                                                          & "  2) Resets Pick Tickets to 'Unconfirmed'" & vbCrLf _
                                                          & "  3) Resets Shipment to 'Unconfirmed'" & vbCrLf & vbCrLf _
                                                          & "Note: You would not be offered this option if this Shipment had already been Reversed", _
                                                          MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, _
                                                          "Shipment " & SHIP_BOL_NO & " has been Confirmed & Posted") = MsgBoxResult.Yes Then
                                                    If MsgBox("Are You Sure?", MsgBoxStyle.YesNo, _
                                                              "Verification to Reverse Invoices") = MsgBoxResult.Yes Then
                                                        Dim INV_REVERSAL_REASON As String = ""
                                                        Using F As New ASFMSGBF
                                                            INV_REVERSAL_REASON = F.Get_txt_from_User("Please Enter the Reason and then Click OK to Proceed", "Enter the Reason for Reversing")
                                                        End Using
                                                        Stop ' SEE WJZ FOR TESTING
                                                        Reverse_Invoice(SHIP_BOL_NO, INV_REVERSAL_REASON)
                                                    End If
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    Case Else
                                        EMsg &= vbCr & "Shipment " & SHIP_BOL_NO & " is No Longer Open"
                                End Select
                            End If
                        End If
                    Next
                End If

                ' New code allows for this. edz 04/28/2013
                'If EMsg = "" And Not MaintenanceMode And Not InquiryMode Then
                '    Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NOs(0))
                '    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", rowSOTSHIP1.Item("WHSE_CODE"))
                '    If Not select_from_3PL_list Then
                '        '  rowICTWHSE1.Item("LP_CODE") = "" ' TEMP TO TEST ENTRY
                '        If rowICTWHSE1.Item("LP_CODE") & "" <> "" Then
                '            MsgBox("You must select this Shipment from the 3PL Shipments tab" _
                '                   & vbCrLf & " in order to be able to Confirm this Shipment", _
                '                   vbOKOnly, "Selected Shipment is associated with a 3PL Warehouse")
                '            ' TEMPORARY ALLOW UNTIL NS CLEANS UP CANCELLED PTS
                '            'Exit Sub
                '        End If
                '    Else
                '        If rowICTWHSE1.Item("LP_CODE") & "" = "" Then
                '            MsgBox("The Warehouse listed on Shipment" _
                '                   & vbCrLf & " is NOT set up as a 3PL Warehouse", _
                '                   vbOKOnly, "Selected Shipment is NOT associated with a 3PL Warehouse")
                '            Exit Sub
                '        End If
                '    End If
                'End If

                ' evaluate items that need expiration dates
                If EMsg = String.Empty Then
                    Dim SHIP_BOL_NO_list As String = "'" & String.Join("', '", SHIP_BOL_NOs) & "'"

                    ASCMAIN1.sql = " select artcust1.cust_code" _
                        & " from sotordr0, sotship1, artcust1" _
                        & " where sotordr0.ordr_group_no = sotship1.ordr_group_no" _
                        & " and sotship1.SHIP_BOL_NO in (" & SHIP_BOL_NO_list & ")" _
                        & " and artcust1.cust_code = sotordr0.cust_code" _
                        & " and artcust1.EDI_REQUIRES_EXP_DATE = '1'"
                    Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
                    If row IsNot Nothing Then
                        ASCMAIN1.sql = " select distinct(ictstyl1.style_code) style_code" _
                         & " from sotpick1, sotcart1, sotcart2, ictstyl1" _
                         & " where sotpick1.SHIP_BOL_NO in (" & SHIP_BOL_NO_list & ")" _
                         & " and sotpick1.pick_no = sotcart1.pick_no" _
                         & " and sotcart1.cart_no = sotcart2.cart_no" _
                         & " and sotcart2.style_code = ictstyl1.style_code" _
                         & " and ictstyl1.REQUIRES_EXP_DATE = '1'" _
                         & " AND SOTCART2.item_exp_date is null"
                        Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
                        If tbl IsNot Nothing AndAlso tbl.Rows.Count > 0 Then
                            EMsg &= vbCr & "The following styles require Expiration Dates on their cartons: "
                            For Each row In tbl.Select("", "style_code")
                                EMsg &= row.Item("style_code") & ", "
                            Next

                            EMsg = EMsg.Substring(0, EMsg.Length - 2)
                            Exit Select
                        End If
                    End If
                End If

                If EMsg = String.Empty AndAlso ASCMAIN1.CLIENT = "RGI" Then
                    For Each SHIP_BOL_NO As String In SHIP_BOL_NOs
                        ASCMAIN1.sql = "Select SOTPICK1.*, SOTORDR1.CCPA_NO CCPA_NO_ORDER" _
                            & " FROM SOTPICK1, SOTORDR1" _
                            & " WHERE SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" _
                            & " AND SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"

                        Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

                        For Each row As DataRow In tbl.Select("", "ORDR_NO")
                            Dim CCPA_NO As String = row.Item("CCPA_NO_AUTH") & String.Empty
                            If CCPA_NO.Length = 0 Then
                                CCPA_NO = row.Item("CCPA_NO_ORDER") & String.Empty
                            End If

                            If CCPA_NO.Length = 0 AndAlso row.Item("CCPA_NO_STATUS") & String.Empty = "1" Then
                                EMsg &= vbCr & "Sales Order (" & row.Item("ORDR_NO") & ") requires a Credit Card Authorization."
                            End If

                            If CCPA_NO.Length = 0 Then
                                Continue For
                            End If

                            Fill_Records("ARTCCPA1", CCPA_NO)
                            If dst.Tables("ARTCCPA1").Rows.Count = 0 Then
                                EMsg &= vbCr & "Sales Order (" & row.Item("ORDR_NO") & ") Credit Card Authorization cannot be found."
                                Continue For
                            End If

                            If Not IsDate(dst.Tables("ARTCCPA1").Rows(0).Item("CCPA_DATE_AUTH") & String.Empty) Then
                                EMsg &= vbCr & "Sales Order (" & row.Item("ORDR_NO") & ") Credit Card Authorization is missing the Authorization Date. Use Sales Order Inquiry to place an Authorization on this Sales Order."
                                Continue For
                            End If

                            Dim CCPA_DATE_AUTH As Date = dst.Tables("ARTCCPA1").Rows(0).Item("CCPA_DATE_AUTH")
                            Dim ts As TimeSpan
                            ts = DateTime.Now - CCPA_DATE_AUTH
                            If Math.Abs(ts.Days) > 90 Then
                                EMsg &= vbCr & "Sales Order (" & row.Item("ORDR_NO") & ") Credit Card Authorization is greater than 90 days."
                            End If
                        Next
                    Next
                End If


                If EMsg <> "" Then
                    ASCMAIN1.MultiTask_Release()
                End If

            Case "Reverse Shipment"

                If Not InquiryMode AndAlso Not ",edz,wjz,".Contains(ASCMAIN1.USER_ID) Then
                    EMsg &= vbCr & "You are not permitted to Reverse a shipment"
                    Exit Select
                End If

                If grdSOTSHIPX.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "You must select a shipment to Reverse"
                    Exit Select
                End If

                Dim SHIP_BOL_NO As String = grdSOTSHIPX.Selected.Rows(0).Cells("SHIP_BOL_NO").Value & String.Empty
                Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)

                If rowSOTSHIP1.Item("SHIP_STATUS") & String.Empty <> "F" Then
                    EMsg &= vbCr & "Shipment must be finalized to perform a Reverse."
                    Exit Select
                End If

                If InquiryMode AndAlso ",edz,wjz,".Contains(ASCMAIN1.USER_ID) Then

                    ASCMAIN1.sql = "Select" _
                        & "  Sum (DECODE(ORDR_YYYYPP_UPDATED,NULL,1,0)) PENDING" _
                        & ", Sum (DECODE(ORDR_YYYYPP_UPDATED,NULL,0,1)) UPDATED" _
                        & " FROM SOTINVH1 WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" _
                        & " and INV_NO_REV is Null"
                    Dim row As DataRow = ASCDATA1.GetDataRow

                    If Val(row.Item("UPDATED") & "") = 0 And (rowSOTSHIP1.Item("REGISTER_XNO") & "") = "" Then
                        If MsgBox(CStr(row.Item("PENDING") & String.Empty) & " Pick Ticket(s) were Confirmed in this Shipment" _
                                  & vbCrLf & vbCrLf _
                                  & "Do you want to De-Confirm all Pick Tickets on this Shipment?" _
                                  & vbCrLf & vbCrLf _
                                  & "Warning - This feature performs the following: " _
                                  & vbCrLf _
                                  & "  1) Deletes Invoice Header & Details" & vbCrLf _
                                  & "  2) Resets Pick Tickets to 'Unconfirmed'" & vbCrLf _
                                  & "  3) Resets Shipment to 'Unconfirmed'" _
                                  & vbCrLf & vbCrLf _
                                  & "Note: You would not be offered this option if any of the Invoices associated with these Pick Tickets were Updated into the A/R", _
                                  MsgBoxStyle.YesNo + MsgBoxStyle.Question, _
                                  "Shipment " & SHIP_BOL_NO & " has been Confirmed as Shipped") = MsgBoxResult.Yes Then
                            Stop ' SEE WJZ FOR TESTING
                            De_Confirm(SHIP_BOL_NO)
                        End If
                        Exit Sub
                    Else
                        ' Need to allow the user to get around this check
                        ' Done here 8/11/2015
                        Dim ediByPass As Boolean = True
                        If (rowSOTSHIP1.Item("SHIP_810_BATCH_NO") & "" <> "N" _
                            AndAlso rowSOTSHIP1.Item("SHIP_810_BATCH_NO") & "" <> "") _
                            OrElse MaintenanceMode Then

                            If MessageBox.Show("EDI Shipment " & SHIP_BOL_NO & " has been Confirmed as Shipped & Updated and EDI Invoice has been sent. New EDI Inovices will be sent when reprocessing this shipment. Do you still want to do a reversal?", _
                                               "Reverse Shipment", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                ediByPass = False
                            End If

                        End If

                        If ediByPass Then
                            ASCMAIN1.sql = "Select Count (*) from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                            Dim PICKS As Int32 = Val(ASCDATA1.GetDataValue)

                            If MsgBox("Shipment " & SHIP_BOL_NO & " has been Confirmed as Shipped & Updated" _
                                      & vbCrLf & vbCrLf _
                                      & "Do you want to Reverse Invoices Generated for all " & CStr(PICKS) & " Pick Tickets on this Shipment?" _
                                      & vbCrLf & vbCrLf _
                                      & "Warning - This feature performs the following: " & vbCrLf _
                                      & "  1) Creates Negative Invoices" & vbCrLf _
                                      & "  2) Resets Pick Tickets to 'Unconfirmed'" & vbCrLf _
                                      & "  3) Resets Shipment to 'Unconfirmed'" & vbCrLf _
                                      & "  4) Resets Sales Orders to 'In Pick'" & vbCrLf _
                                      & "  5) Adjusts Inventory" & vbCrLf _
                                      & "  6) Creates Negative AR" & vbCrLf _
                                      & "  7) Reverts impact to Sales Summary" & vbCrLf _
                                      & "  8) Replaces Shipped Stock to Shipping Location (if appl)" & vbCrLf & vbCrLf _
                                      & "Note: You would not be offered this option if this Shipment had already been Reversed", _
                                      MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, _
                                      "Shipment " & SHIP_BOL_NO & " has been Confirmed & Posted") = MsgBoxResult.Yes Then

                                If MsgBox("Are You Sure?", MsgBoxStyle.YesNo, _
                                          "Verification to Reverse Invoices") = MsgBoxResult.Yes Then
                                    Dim INV_REVERSAL_REASON As String = ""
                                    Using F As New ASFMSGBF
                                        INV_REVERSAL_REASON = F.Get_txt_from_User("Please Enter the Reason and then Click OK to Proceed", "Enter the Reason for Reversing")
                                    End Using
                                    INV_REVERSAL_REASON = INV_REVERSAL_REASON.Trim
                                    If INV_REVERSAL_REASON.Length = 0 Then
                                        EMsg &= vbCr & "A Reason is required for a reversal."
                                        Exit Select
                                    End If
                                    Reverse_Invoice(SHIP_BOL_NO, INV_REVERSAL_REASON)
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                End If

            Case "Update"


                For Each field As String In New String() {"TERM_CODE", "SHIP_VIA_CODE", "FRT_TERMS", "SREP_CODE", "SREP2_CODE", _
                                                          "REASON_CODE", "SHIP_REF", "BILL_OF_LADING_NO", "ORDR_DEPT", "EDI_LOAD_ID", "BTB_BOL_NO"}
                    Absx1.txtFor(field).Text = Absx1.txtFor(field).Text.Trim
                    For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)
                        If dst.Tables("SOTSHIP1").Columns.Contains(field) Then
                            rowSOTSHIP1.Item(field) = Absx1.txtFor(field).Text
                        End If
                    Next
                Next

                If Not MaintenanceMode AndAlso ASCMAIN1.DBS_COMPANY = "VAN" Then
                    EMsg &= vbCr & "Naughty Girl!!"
                    Exit Select
                End If

                If select_from_3PL_list Then
                    If dst.Tables("SOTPICK1").Select("ISNULL(SELECTED,'0')<>'1'").Length > 0 Then
                        EMsg &= vbCr & "Cannot De-Select Pick Tickets from a 3PL Shipment"
                        EMsg &= vbCr & "- If they did not ship, they must be confirmed as 0 Shipped"
                    End If
                End If

                If dst.Tables("SOTPICK1").Select("ISNULL(SELECTED,'0')='1' and PICK_QTY_CONF<>0").Length = 0 Then
                    EMsg &= vbCr & "Cannot Update when nothing is confirmed as shipped."
                    EMsg &= vbCr & "- Use Cancel Shipment option"
                End If

                If dst.Tables("SOTPICK1").Select("ISNULL(SELECTED,'0')='1' and PICK_QTY_CONF=0 and (PICK_CNT_CARTONS<>0 or PICK_TOTAL_WGT<>0)").Length <> 0 Then
                    EMsg &= vbCr & "Some Pick Tickets have 0 qty confirmed as Shipped"
                    EMsg &= vbCr & "-  but Still have a non-Zero value for cartons or weight"
                End If

                If 1 = 2 AndAlso ASCMAIN1.CLIENT = "RGI" Then
                    For Each rowSOTPICK1 As DataRow In dst.Tables("").Select("SELECTED = '1'", "", DataViewRowState.CurrentRows)
                        Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "'", "", DataViewRowState.CurrentRows)
                            Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                            Dim ORDR_LNO As Int16 = rowSOTPICK1.Item("ORDR_NO")
                            Dim updated As Boolean = False

                            For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("", "", DataViewRowState.CurrentRows)
                                If updated Then Exit For

                                For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select("CART_NO = '" & rowSOTCART1.Item("CART_NO") & "'", "", DataViewRowState.CurrentRows)
                                    If rowSOTCART2.Item("ORDR_NO") & String.Empty = ORDR_NO _
                                        AndAlso Val(rowSOTCART2.Item("ORDR_LNO") & String.Empty) = ORDR_LNO Then

                                        'QTY_PACKED
                                    End If
                                Next
                            Next
                        Next
                    Next
                End If

                ' Although this only matters for edi customers, I think we should enforce the integrity
                ' RGI does not want to validate against cartons amount. Regardless what they ship, they have only one carton
                ' Angela will place the carton qty on the pick ticket header
                Dim rowSOTCARTX_oobal As DataRow() = dst.Tables("SOTCARTX").Select("ISNULL(PICK_QTY_CONF,0) <> ISNULL(QTY_PACKED,0)")
                If rowSOTCARTX_oobal.Length <> 0 AndAlso Not ASCMAIN1.CLIENT = "RGI" Then
                    EMsg &= vbCr & "Pick Ticket Detail Qty Confirmed out of balance with Carton Details"
                    EMsg &= vbCr & " (See Pick Ticket " & rowSOTCARTX_oobal(0).Item("PICK_NO") & " Line " & rowSOTCARTX_oobal(0).Item("ORDR_LNO")
                End If

                If MaintenanceMode Then
                    Dim rowSOTPICK2_higher As DataRow() = dst.Tables("SOTPICK2").Select("ISNULL(PICK_QTY_CONF,0) < 0 or ISNULL(PICK_QTY_CONF,0) > ISNULL(PICK_QTY,0) + ISNULL(PICK_QTY_CANC_REL,0)")
                    If rowSOTPICK2_higher.Length <> 0 Then
                        EMsg &= vbCr & "Pick Ticket Detail Qty cannot be revised upward."
                        EMsg &= vbCr & " (See Pick Ticket " & rowSOTPICK2_higher(0).Item("PICK_NO") & " Line " & rowSOTPICK2_higher(0).Item("PICK_LNO")
                    End If

                    If Format(dteORDR_CANCEL_DATE.Value, "yyyyMMdd") < Format(dteORDR_SHIP_DATE.Value, "yyyyMMdd") Then
                        EMsg &= vbCr & "Cancel Date may not be prior to Ship Date"
                    End If
                    If txtReason.Text = "" OrElse txtContact.Text = "" Then
                        EMsg &= vbCr & "Reason and Contact are Mandatory when making changes to a Shipment"
                    End If

                    For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("ISNULL(PICK_UNIT_PRICE,0) <> ISNULL(ORDR_UNIT_PRICE,0)")
                        Dim rowSOTORDR9 As DataRow = dst.Tables("SOTORDR9").Rows.Find(New Object() {rowSOTPICK2.Item("ORDR_NO"), rowSOTPICK2.Item("RANGE_STYLE_LNO")})
                        If rowSOTORDR9 IsNot Nothing Then
                            If Val(rowSOTORDR9.Item("RANGE_STYLE_PRICE") & "") <> Val(rowSOTPICK2.Item("ORDR_UNIT_PRICE") & "") Then
                                EMsg &= "Range Style Price Mis-Match (See Order " & rowSOTPICK2.Item("ORDR_NO") & " Range Style Ln " & rowSOTPICK2.Item("RANGE_STYLE_LNO") & ")"
                                Exit For
                            End If
                        End If
                    Next
                Else
                    If Absx1.dteFor("SHIP_DATE_SHIPPED").Value & "" = "" _
                        OrElse Absx1.dteFor("INV_DATE").Value & "" = "" Then
                        EMsg &= vbCr & "Date Shipped and Invoice Date are Mandatory"
                    Else
                        If Format(Absx1.dteFor("SHIP_DATE_SHIPPED").Value, "yyyyMMdd") _
                         > Format(Absx1.dteFor("INV_DATE").Value, "yyyyMMdd") Then
                            EMsg &= vbCr & "Invoice Date cannot be Prior to Date Shipped"
                        End If

                        If (Format(Absx1.dteFor("INV_DATE").Value, "yyyyMMdd") < Format(DateTime.Now, "yyyyMMdd")) _
                            AndAlso ASCMAIN1.CLIENT = "NYA" _
                            AndAlso Not select_from_3PL_list Then
                            EMsg &= vbCr & "Invoice Date cannot be Prior to today"
                        Else
                            Dim numdays As Int32 = DateDiff(DateInterval.Day, Absx1.dteFor("INV_DATE").Value, DateTime.Now)
                            If numdays > 4 Then
                                If MessageBox.Show("You choose to invoice this shipment using a date " & numdays & " day(s) ago. Do you want to proceed.", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                    Exit Sub
                                End If
                            End If
                        End If

                        If EMsg.Length = 0 Then
                            If Format(Absx1.dteFor("INV_DATE").Value, "yyyyMM") <> ASCMAIN1.CYM Then
                                If Format(Absx1.dteFor("INV_DATE").Value, "yyyyMM") = ASCMAIN1.Period_Calc(ASCMAIN1.CYM, 1) Then
                                    If MsgBox("You are about to confirm a shippment that will be posted into the Next period.", _
                                               MsgBoxStyle.OkCancel, _
                                               "Invoice Date Confirmation") <> MsgBoxResult.Ok Then
                                        Exit Sub
                                    End If
                                Else
                                    If ASCMAIN1.CLIENT = "NYA" Then
                                        Dim zMsg As String = "Invoice Date is Not in the Current Period. If you continue the shipment invoice date will use today's date."
                                        zMsg &= vbCr & vbCr & "Continue with Update?"
                                        If MessageBox.Show(zMsg, "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                            Exit Sub
                                        End If
                                        Absx1.dteFor("INV_DATE").Value = CDate(DateTime.Now.ToShortDateString)
                                        Application.DoEvents()
                                    Else
                                        EMsg &= vbCr & "Invoice Date Not in Current Period"
                                    End If
                                End If
                            End If
                        End If
                    End If

                    If Absx1.txtFor("TERM_CODE").Text = "" Then
                        EMsg &= vbCr & "Terms Code is Required"
                    Else
                        If LookUp("TATTERM1", Absx1.txtFor("TERM_CODE").Text) Is Nothing Then
                            EMsg &= vbCr & "Invalid Terms Code"
                        End If
                    End If

                    If Absx1.txtFor("SHIP_VIA_CODE").Text = "" Then
                        EMsg &= vbCr & "Ship Via Code is Required"
                    Else
                        If LookUp("SOTSVIA1", Absx1.txtFor("SHIP_VIA_CODE").Text) Is Nothing Then
                            EMsg &= vbCr & "Invalid Ship Via Code"
                        Else
                            Dim CARRIER_MODE As String = cdr.Item("CARRIER_MODE") & String.Empty
                            If CARRIER_MODE <> "U" AndAlso commonCarriersOnly Then
                                EMsg &= vbCr & "This screen is for Common Carriers (Fedex, UPS) only "
                            Else
                                Dim CART_SEQ As Int32 = Val(dst.Tables("SOTCART1").Compute("MAX(CART_SEQ)", "") & String.Empty) + 1
                                For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SELECTED = '1'")
                                    Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")

                                    For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'")
                                        ' Make sure all cartons have a sequence number. This is how the shipping labels will print
                                        If rowSOTCART1.Item("CART_SEQ") & String.Empty = String.Empty Then
                                            rowSOTCART1.Item("CART_SEQ") = CART_SEQ
                                            CART_SEQ += 1
                                        End If

                                        ' As per Debbie she does not have time to enter all carton information
                                        If (commonCarrier OrElse edi_customer) AndAlso ASCMAIN1.CLIENT <> "NYA" Then
                                            If rowSOTCART1.Item("PACKAGING_TYPE") & String.Empty = String.Empty Then
                                                EMsg &= vbCrLf & "Package type is required for all cartons"
                                            ElseIf Val(rowSOTCART1.Item("PACKAGING_TYPE") & String.Empty) = nsoftware.InShip.TPackagingTypes.ptYourPackaging Then
                                                If rowSOTCART1.Item("PKG_CODE") & String.Empty = String.Empty Then
                                                    EMsg &= vbCrLf & "Package code is required for all 'Our Packaging' cartons"
                                                ElseIf rowSOTCART1.Item("PKG_CODE") & String.Empty = "OTHER" Then
                                                    If Val(rowSOTCART1.Item("LENGTH") & String.Empty) <= 0 _
                                                             OrElse Val(rowSOTCART1.Item("WIDTH") & String.Empty) <= 0 _
                                                            OrElse Val(rowSOTCART1.Item("HEIGHT") & String.Empty) <= 0 Then
                                                        EMsg &= vbCr & "package Type Other requires Width, length and height to be set."
                                                    End If
                                                End If
                                            End If
                                        End If

                                        If Val(rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") & String.Empty) <= 0 AndAlso ASCMAIN1.CLIENT <> "NYA" Then
                                            EMsg &= vbCrLf & "Package weight is required for all cartons"
                                        Else
                                            Dim pickWeight As Decimal = Val(rowSOTPICK1.Item("PICK_TOTAL_WGT") & String.Empty)
                                            Dim totalCartWeight As Decimal = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)

                                            If totalCartWeight > pickWeight Then
                                                Select Case ASCMAIN1.CLIENT
                                                    Case "RGI", "NYA"
                                                        ' By Pass validation RGI-Angie, NYA-Debbie
                                                    Case Else
                                                        EMsg &= vbCrLf & "Package weight for a Pick Ticket must be greater equal total cartons weight (" & PICK_NO & ")"
                                                End Select
                                            End If
                                        End If
                                    Next
                                Next
                            End If
                        End If
                    End If

                    Absx1.txtFor("FRT_TERMS").Text = Absx1.txtFor("FRT_TERMS").Text.Trim
                    If Absx1.txtFor("FRT_TERMS").Text = "" Then
                        EMsg &= vbCr & "Frt Terms Code is Required"
                    ElseIf LookUp("ASTCODE1", New String() {"SOTORDR1", "FRT_TERMS", Absx1.txtFor("FRT_TERMS").Text}) Is Nothing Then
                        EMsg &= vbCr & "Invalid Frt Terms Code"
                    Else
                        Select Case Absx1.txtFor("FRT_TERMS").Text
                            Case "COL", "PPD"
                                If dst.Tables("SOTPICK1").Select("ISNULL(PICK_FREIGHT, 0) <> 0 ").Length > 0 Then
                                    EMsg &= vbCr & "Frt Terms Code (" & Absx1.txtFor("FRT_TERMS").Text & ") does not permit freight."
                                End If
                            Case "PPA"
                                ' RGI some times bills freight on a separate invoice. As per Rich 09-06-2013
                                ' and he does not want the Merchandise Invoice to show Freight Prepaid
                                If ASCMAIN1.CLIENT = "RGI" Then
                                    If dst.Tables("SOTPICK1").Select("ISNULL(PICK_FREIGHT, 0) <> 0 ").Length = 0 Then
                                        If MessageBox.Show("Frt Terms Code (" & Absx1.txtFor("FRT_TERMS").Text & ") requires freight." _
                                                         & Environment.NewLine _
                                                         & "Do you want to continue without freight charges?", "Freight", _
                                                          MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                Else
                                    ' 3PLs are exempt from freight on Pick Ticket - NYA gets the freight from UPS/Fedex
                                    If Not select_from_3PL_list AndAlso dst.Tables("SOTPICK1").Select("ISNULL(PICK_FREIGHT, 0) <> 0 ").Length = 0 _
                                         AndAlso Not commonCarrier Then
                                        EMsg &= vbCr & "Frt Terms Code (" & Absx1.txtFor("FRT_TERMS").Text & ") requires freight."
                                    ElseIf Not select_from_3PL_list AndAlso dst.Tables("SOTPICK1").Select("ISNULL(PICK_FREIGHT, 0) <> 0 ").Length > 0 _
                                            AndAlso commonCarrier Then
                                        If MessageBox.Show("The terms are PPA and there is freight on at least one Pick Ticket. Do you want to add the Common Carrier Charges to the Existing Freight?", _
                                                           "PPA Freight", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.No Then
                                            AddPPAToExistingFreight = False
                                        Else
                                            AddPPAToExistingFreight = True
                                        End If
                                    End If

                                End If
                        End Select
                    End If

                    If Absx1.txtFor("SREP_CODE").Text <> "" AndAlso LookUp("SOTSREP1", Absx1.txtFor("SREP_CODE").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Sales Rep Code"
                    End If

                    If Absx1.txtFor("SREP2_CODE").Text <> "" AndAlso LookUp("SOTSREP1", Absx1.txtFor("SREP2_CODE").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Sales Rep2 Code"
                    End If

                    If dst.Tables("SOTPICK2").Select("PICK_QTY_CANC <> 0").Length <> 0 Then
                        If Absx1.txtFor("REASON_CODE").Text = "" Then
                            EMsg &= vbCr & "Reason Code is Required when Cancelling Qty's on any Pick Ticket"
                        Else
                            If LookUp("SOTREAS1", Absx1.txtFor("REASON_CODE").Text) Is Nothing Then
                                EMsg &= vbCr & "Invalid Reason Code"
                            End If
                        End If
                    Else
                        If Absx1.txtFor("REASON_CODE").Text <> "" Then
                            EMsg &= vbCr & "Reason Code should NOT be specified unless Cancelling Qty's"
                        End If
                    End If

                    If dst.Tables("SOTPICK2").Select("PICK_QTY_CONF <> 0").Length <> 0 Then
                        If Absx1.txtFor("SHIP_VIA_CODE").Text = "" _
                            OrElse (Absx1.txtFor("SHIP_REF").Text = "" AndAlso Not ASCMAIN1.CLIENT = "NYA") Then

                            If Absx1.txtFor("SHIP_VIA_CODE").Text = "" Then
                                EMsg &= vbCr & "Ship Via Code is Required"
                            End If

                            If (Absx1.txtFor("SHIP_REF").Text = "" AndAlso Not ASCMAIN1.CLIENT = "NYA") Then
                                EMsg &= vbCr & "Shippers Reference (Pro #) is Required"
                            End If

                        Else
                            Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", Absx1.txtFor("SHIP_VIA_CODE").Text)
                            If rowSOTSVIA1 Is Nothing Then
                                EMsg &= vbCr & "Invalid Ship Via Code"
                            Else
                                If edi_customer Then
                                    If rowSOTSVIA1.Item("SHIP_VIA_SCAC") & "" = "" Then
                                        EMsg &= vbCr & "Selected Shipper Requires SCAC Code For EDI Customers"
                                    End If
                                End If
                            End If
                        End If
                    End If

                    ' If Fedex or UPS then the Tracking Numer becomes the BILL_OF_LADING_NO
                    If Not commonCarrier Then
                        If dst.Tables("SOTPICK2").Select("PICK_QTY_CONF <> 0").Length <> 0 And edi_order Then
                            If dst.Tables("SOTSHIP1").Rows.Count > 1 Then
                                If dst.Tables("SOTPICK1").Select("SELECTED = '1' and BILL_OF_LADING_NO is Null").Length <> 0 Then
                                    EMsg &= vbCr & "BOL No is Mandatory for EDI Orders"
                                End If
                            Else
                                If Absx1.txtFor("BILL_OF_LADING_NO").Text = "" Then
                                    EMsg &= vbCr & "BOL No is Mandatory for EDI Orders"
                                End If
                            End If
                        End If
                    End If

                    Dim sqlw As String = ""

                    'If Absx1.txtFor("FRT_TERMS").Text <> "" Then
                    '    If Absx1.txtFor("FRT_TERMS").Text <> "PPA" Then
                    '        sqlw = "ISNULL(PICK_FREIGHT, 0) <> 0 and SELECTED = '1'"
                    '    Else
                    '        sqlw = "ISNULL(PICK_FREIGHT, 0) = 0 and SELECTED = '1'"
                    '    End If

                    '    If dst.Tables("SOTPICK1").Select(sqlw).Length > 0 Then
                    '        If Absx1.txtFor("FRT_TERMS").Text <> "PPA" Then
                    '            ' If a 3PL then no message. the freight came back in the 945.
                    '            If Not select_from_3PL_list Then
                    '                EMsg &= vbCr & "Freight Terms Code Specified does not permit Non-Zero Freight Amounts"
                    '            End If
                    '        Else
                    '            EMsg &= vbCr & "Freight Terms Code Specified does not permit Zero Freight Amounts"
                    '        End If
                    '    End If
                    'End If

                    If edi856_customer Then
                        sqlw = "PICK_TOTAL_UNITS_CALC <> PICK_QTY_CONF"
                        Dim rows() As DataRow = dst.Tables("SOTPICK1").Select(sqlw)
                        If rows.Length <> 0 Then
                            EMsg = EMsg & vbCr & CStr(rows.Length) & " Pick Ticket(s) not matching Carton Details (See PT#" & rows(0).Item(0) & ")"
                        End If
                    End If

                    Dim PICK_NO_last As String = ""
                    Dim RANGE_STYLE_LNO_last As Int32 = 0
                    Dim PICK_UNIT_PRICE_last As Decimal = 0
                    sqlw = "PICK_QTY_CONF <> 0 and RANGE_STYLE_LNO <> 0"
                    For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select(sqlw, "PICK_NO,RANGE_STYLE_LNO,PICK_UNIT_PRICE")
                        Dim ORDR_NO As String = rowSOTPICK2.Item("ORDR_NO")
                        Dim PICK_NO As String = rowSOTPICK2.Item("PICK_NO")
                        Dim RANGE_STYLE_LNO As Int32 = Val(rowSOTPICK2.Item("RANGE_STYLE_LNO") & "")
                        Dim PICK_UNIT_PRICE As Decimal = Val(rowSOTPICK2.Item("PICK_UNIT_PRICE") & "")
                        Dim rowSOTORDR9 As DataRow = dst.Tables("SOTORDR9").Rows.Find(New Object() {ORDR_NO, RANGE_STYLE_LNO})
                        Dim RANGE_STYLE_QTY_PER_PP As Int64 = Val(rowSOTORDR9.Item("RANGE_STYLE_QTY_PER_PP") & "")
                        If PICK_NO_last = PICK_NO And RANGE_STYLE_LNO_last = RANGE_STYLE_LNO _
                            And System.Math.Abs(PICK_UNIT_PRICE_last - PICK_UNIT_PRICE) > 0.005 Then
                            If RANGE_STYLE_QTY_PER_PP = 0 OrElse _
                                RANGE_STYLE_QTY_PER_PP = 1 Then
                                EMsg &= vbCr & "Range Style Components with Different Prices (Range Style Line No " & CStr(RANGE_STYLE_LNO) & ")"
                            End If
                            Exit For
                        Else
                            PICK_NO_last = PICK_NO
                            RANGE_STYLE_LNO_last = RANGE_STYLE_LNO
                            PICK_UNIT_PRICE_last = PICK_UNIT_PRICE
                        End If
                    Next

                    'Check for Assortments Knocked Out of Balance.

                    Dim LAST_MULT As Decimal = 0
                    RANGE_STYLE_LNO_last = 0
                    sqlw = "PICK_QTY_CONF <> 0 and RANGE_STYLE_LNO <> 0"
                    For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select(sqlw, "ORDR_NO,RANGE_STYLE_LNO")
                        Dim ORDR_NO As String = rowSOTPICK2.Item("ORDR_NO")
                        Dim RANGE_STYLE_LNO As Int32 = Val(rowSOTPICK2.Item("RANGE_STYLE_LNO") & "")
                        Dim rowSOTORDR9 As DataRow = dst.Tables("SOTORDR9").Rows.Find(New Object() {ORDR_NO, RANGE_STYLE_LNO})
                        If Val(rowSOTORDR9.Item("RANGE_STYLE_QTY_PER_PP") & "") > 1 Then
                            Dim PICK_QTY As Int64 = Val(rowSOTPICK2.Item("PICK_QTY") & "")
                            Dim PICK_QTY_CONF As Int64 = Val(rowSOTPICK2.Item("PICK_QTY_CONF") & "")

                            If RANGE_STYLE_LNO_last = 0 OrElse RANGE_STYLE_LNO_last <> RANGE_STYLE_LNO Then
                                LAST_MULT = 0
                                RANGE_STYLE_LNO_last = RANGE_STYLE_LNO
                            End If
                            If LAST_MULT <> 0 Then
                                If PICK_QTY_CONF = 0 Then
                                    'DG SAYS THIS WILL SOLVE ALL OUR PROBLEMS COMPLETELY - WR. 1/4/06
                                    '                        If LAST_MULT <> dynWK.Item("PICK_QTY") Then
                                    '                            EMsg = EMsg & vbCr & "Assortment Pre-Pack Not in Balance"
                                    '                            Exit Do
                                    '                        End If
                                Else
                                    If LAST_MULT <> PICK_QTY / PICK_QTY_CONF Then
                                        EMsg &= vbCr & "Assortment Pre-Pack Not in Balance"
                                        Exit For
                                    End If
                                End If
                            Else
                                If PICK_QTY_CONF = 0 Then
                                    ' DG SAYS THIS WILL SOLVE ALL OUR PROBLEMS COMPLETELY - WR. 1/4/06
                                    '                        LAST_MULT = dynWK.Item("PICK_QTY")
                                Else
                                    LAST_MULT = PICK_QTY / PICK_QTY_CONF
                                End If
                            End If
                        End If
                    Next

                    sqlw = "PICK_QTY_CONF <> 0 and RANGE_STYLE_LNO <> 0 and ISNULL(PICK_UNIT_PRICE,0) <> ISNULL(ORDR_UNIT_PRICE,0)"
                    For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select(sqlw)
                        Dim ORDR_NO As String = rowSOTPICK2.Item("ORDR_NO")
                        Dim RANGE_STYLE_LNO As Int32 = Val(rowSOTPICK2.Item("RANGE_STYLE_LNO") & "")
                        Dim rowSOTORDR9 As DataRow = dst.Tables("SOTORDR9").Rows.Find(New Object() {ORDR_NO, RANGE_STYLE_LNO})
                        If Val(rowSOTORDR9.Item("RANGE_STYLE_QTY_PER_PP") & "") > 1 Then
                            EMsg &= vbCr & "Assortments Prices Have Been Changed.  Not Allowed."
                        End If
                    Next

                    If EMsg = "" Then
                        If (chkFactored.Checked And rowARTCUST1.Item("CUST_FACTOR_IND") & "" <> "1") _
                        OrElse (Not chkFactored.Checked And rowARTCUST1.Item("CUST_FACTOR_IND") & "" = "1") Then
                            If MsgBox("Factor Option is not in synch with Customer Master" & vbCrLf & "Continue Anyway", _
                                       MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If ASCMAIN1.CLIENT = "VAN" AndAlso Not ASCMAIN1.USER_SECURITY_CODEs.Contains("BD") Then
                        For Each row As DataRow In ASCDATA1.SelectDistinct _
                                (dst.Tables("SOTPICK2").Select("PICK_QTY_CONF <> 0"), New String() {"STYLE_CODE", "COLOR_CODE"}).Rows
                            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
                            Dim STYLE_COST As Decimal = Val(rowICTSTYC1.Item("STYLE_COST_FIFO") & "")
                            If STYLE_COST = 0 OrElse STYLE_COST < 0.01 Then
                                If ASCMAIN1.DBS_SERVER = "" OrElse ASCMAIN1.DBS_COMPANY = "TST" Then
                                    ' ALLOW FOR NOW
                                Else
                                    EMsg &= vbCr & "Style " & STYLE_CODE & " has a Valuation Cost of " & Format(STYLE_COST, "##0.00") & "."
                                End If
                            End If
                        Next
                    End If

                    If EMsg.Length = 0 AndAlso dst.Tables("SOTPICK2").Select("ORDR_UNIT_PRICE = 0").Length <> 0 Then
                        If MsgBox("This Shipment Contains Styles That have Zero Prices." _
                             & vbCrLf & "Are You Sure You Want To Update This?", MsgBoxStyle.YesNo, _
                             "Price Check") = MsgBoxResult.No Then
                            EMsg &= vbCr & "Cancelled By User Due To Zero Price."
                        End If
                    End If
                End If

                If EMsg.Length = 0 AndAlso ASCMAIN1.CLIENT = "NYA" Then
                    For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SELECTED = '1' AND ORDR_TYPE_CODE = 'XFR'", "")
                        If Val(rowSOTPICK1.Item("PICK_FREIGHT") & String.Empty) <> 0 Then
                            EMsg &= vbCr & "Pick Ticket (" & rowSOTPICK1.Item("PICK_NO") & ") is a Transfer Sales Order and cannot have freight."
                        End If
                    Next
                End If

                If EMsg.Length = 0 Then
                    If numInsureValue.Value > 200 _
                        AndAlso Not chkInsureShipment.Checked _
                        AndAlso commonCarrier _
                        AndAlso Not select_from_3PL_list _
                        AndAlso ASCMAIN1.CLIENT <> "RGI" Then
                        If MessageBox.Show("Do you want to continue without insuring the shipment.", "Insure", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

                If EMsg.Length = 0 Then
                    Dim packageInsure As Decimal = Val(dst.Tables("SOTCART1").Compute("SUM(INSURANCE)", "") & String.Empty)
                    If chkInsureShipment.Checked Then
                        If packageInsure <= 0 Then
                            EMsg &= vbCr & "You chose to insure the shipment but you did not place Insurance values on the cartons."
                        Else
                            If MessageBox.Show("Do you want to insure the shipment for a total of $" & Format(packageInsure, "#,##0.00"), "Insure", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                Exit Sub
                            End If
                        End If
                    ElseIf packageInsure > 0 Then
                        If MessageBox.Show("You provided insurance for the carton(s); however, you did not check 'Insure Package'. Do you want to continue without applying the insurance?", _
                             "Insure", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

                For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
                    If rowSOTORDR1.Item("CURR_CODE") & String.Empty = String.Empty OrElse
                        Val(rowSOTORDR1.Item("CURR_EXCH_RATE") & String.Empty) <= 0 Then
                        EMsg &= vbCr & "Sales Order (" & rowSOTORDR1.Item("ORDR_NO") & ") has an invalid Currency Code or Currency Rate"
                    End If
                Next

                If EMsg.Length = 0 AndAlso commonCarrier Then
                    RequestShippingLabel(Nothing, EMsg, True)
                End If

                If Not InquiryMode AndAlso EMsg.Length = 0 AndAlso ASCMAIN1.CLIENT = "NYA" AndAlso dst.Tables("SOTSHIP1").Select("WHSE_CODE = '52'").Length > 0 Then
                    txtBTB_BOL_NO.Text = txtBTB_BOL_NO.Text.Trim
                    If txtBTB_BOL_NO.TextLength = 0 Then
                        If MessageBox.Show("This Shipment requires a " & lblFCR.Text & ". Do you want to continue without providing one?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

                ' Clean Up SOTINVHM
                If EMsg.Length = 0 Then
                    dst.Tables("SOTINVHM").AcceptChanges()
                    For Each row As DataRow In dst.Tables("SOTINVHM").Select()
                        row.SetAdded()
                    Next

                    Dim Confirmed As Decimal = Val(dst.Tables("SOTCONFT").Select("KEY = 2")(0).Item("AMT") & String.Empty)
                    Dim Freight As Decimal = Val(dst.Tables("SOTCONFT").Select("KEY = 5")(0).Item("AMT") & String.Empty)
                    Dim MiscCharge As Decimal = Val(dst.Tables("SOTCONFT").Select("KEY = 6")(0).Item("AMT") & String.Empty)

                    If (Confirmed + Freight) + MiscCharge < 0 Then
                        EMsg &= vbCr & "Invoice cannot be negative. Confirmed + Freight + Misc Charges are less than 0."
                    End If
                End If


            Case "Cancel"
                If processing_select_from_3PL_list Then
                Else
                    If MsgBox("Are you sure that you want to Cancel?", _
                          MsgBoxStyle.YesNo, _
                          "Verification to Cancel working with this Record") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If


            Case "Cancel Shipment"

                ' You may use this option only if all of the Pick Tickets are already Cancelled

                If dst.Tables("SOTPICK1").Select("ISNULL(SELECTED,'0')<>'1'").Length <> 0 Then
                    EMsg &= vbCr & "All Pick Tickets must be Selected in order to Cancel the Entire Shipment"
                    EMsg &= vbCr & "You may not cancel some pick tickets (and leave others open) with this option"
                End If

                If dst.Tables("SOTPICK1").Select("PICK_STATUS <> 'C' AND PICK_QTY_CONF <> 0").Length <> 0 Then
                    EMsg &= vbCr & "Cancellation Not Permitted" & vbCrLf & " - Some Pick Tickets on this Shipment are NOT Cancelled"
                    EMsg &= vbCr & vbCr & "Click on Shipment (in Mass Changes) and then use the Cancel button"
                End If

                If EMsg = "" Then
                    If MsgBox("This option will Cancel this Shipment." _
                              & vbCrLf & vbCrLf & "Use this option to Cancel All Pick Tickets on this Shipment" _
                              & vbCrLf & " and also Cancel this Shipment." _
                              & vbCrLf _
                              & vbCrLf & "This option will NOT restore the Order back to an Open state." _
                              & vbCrLf & "This option will NOT cause any EDI documents to transmit." _
                              & vbCrLf & "This option will NOT create Invoices." _
                              & vbCrLf & vbCrLf & "If you want to cancel this shipment so that the orders are re-opened," _
                              & vbCrLf & " then use De-Release." _
                              & vbCrLf & vbCrLf & "Are you sure that you want to Cancel this Shipment?", _
                                  MsgBoxStyle.YesNo, _
                                  "WARNING: This Action is Permanent") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Add Carton"
                If MessageBox.Show("Do you want to add a carton to this shipment?", "Add Carton", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If


            Case "Create Shipping Label"
                If MessageBox.Show("The requested shiment has been finalized. Do you want to get a new shipping label?", _
                    "New label", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                Else
                    RecreateLabel = True
                    Exit Select
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

                If InquiryMode Then
                    EntryMode = "V"
                Else
                    EntryMode = "E"
                End If

                Load_Record()

                ' Special EDI Processing
                If Not InquiryMode AndAlso select_from_3PL_list AndAlso EntryMode = "E" Then
                    If ASCMAIN1.CLIENT = "NYA" Then
                        If Not Load_3PL_Shipment_Details_EDT945T1() Then
                            select_from_3PL_list = False
                            Click_Command("Cancel")
                            Exit Sub
                        End If
                    End If
                End If

                Mode_Settings(True)

            Case "Update"
                refreshScreen = True
                If MaintenanceMode Then
                    Update_Record_Maintenance()
                Else
                    Update_Record()
                End If
                If refreshScreen Then
                    Mode_Settings(False)
                End If

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Done", "Cancel", "Reverse Shipment"
                Mode_Settings(False)

            Case "Cancel Shipment"
                ' grdWHT3PLS1.Tag = ""
                Cancel_Shipment()
                Mode_Settings(False)

            Case "Force PTs to Balance"
                Force_PTs_to_Balance()

            Case "Force Cartons to Balance"
                Force_Cartons_to_Balance()

            Case "Substitute"
                Add_Line(True)

            Case "Add Line"
                Add_Line(False)

            Case "Add Carton"
                AddCarton()

            Case "Create Shipping Label"
                Update_Record()
                Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("Select").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("Cancel Shipment").Settings.Enabled = iScreenMode
                .Items("Add Carton").Settings.Enabled = iScreenMode

                If ",edz,wjz,".Contains(ASCMAIN1.USER_ID) AndAlso InquiryMode Then
                    .Items("Reverse Shipment").Settings.Enabled = not_iScreenMode
                Else
                    .Items("Reverse Shipment").Visible = False
                End If

                .Items("Done").Visible = InquiryMode OrElse (EntryMode = "L" And ScreenMode)
                .Items("Select").Visible = (Not (EntryMode = "L") OrElse Not ScreenMode)
                .Items("Update").Visible = Not InquiryMode And (Not (EntryMode = "L") OrElse Not ScreenMode)
                .Items("Cancel").Visible = Not InquiryMode And (Not (EntryMode = "L") OrElse Not ScreenMode)
                .Items("Cancel Shipment").Visible = (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN") And ScreenMode And MaintenanceMode '  False ' Not InquiryMode And (Not (EntryMode = "L") OrElse Not ScreenMode) And Not select_from_3PL_list And Not MaintenanceMode

                If (ASCMAIN1.CLIENT = "NYA" OrElse ASCMAIN1.CLIENT = "RGI") Then
                    .Items("Add Carton").Visible = Not InquiryMode AndAlso (Not (EntryMode = "L") OrElse Not ScreenMode)
                Else
                    .Items("Add Carton").Visible = False
                End If

                If ASCMAIN1.CLIENT = "NYA" AndAlso InquiryMode Then
                    .Items("Create Shipping Label").Visible = EntryMode = "V" _
                        AndAlso dst.Tables("SOTSHIP1").Rows.Count = 1 _
                        AndAlso dst.Tables("SOTSHIP1").Select("SHIP_STATUS <> 'F'").Length = 0 _
                        AndAlso ASCMAIN1.USER_SECURITY_CODEs.Contains("WS")

                    .Items("Add Carton").Visible = .Items("Create Shipping Label").Visible
                Else
                    .Items("Create Shipping Label").Visible = False
                End If
            End With

            .Groups("Totals").Visible = ScreenMode
            .Groups("Special Operations").Visible = ScreenMode And (EntryMode <> "L") And Not InquiryMode And Not select_from_3PL_list ' And Not MaintenanceMode
            If ASCMAIN1.DBS_COMPANY <> "VAN" Then
                .Groups("Special Operations").Visible = False
            End If
            .Groups("Special Operations").Items("Substitute").Visible = Not MaintenanceMode
            .Groups("Special Operations").Items("Add Line").Visible = Not MaintenanceMode
            '.Groups("Order Header Changes").Visible = ScreenMode And (EntryMode <> "L") And Not InquiryMode And Not select_from_3PL_list And MaintenanceMode
            .Groups("Mass Changes").Visible = ScreenMode And (EntryMode <> "L") And Not InquiryMode And Not select_from_3PL_list AndAlso Not ASCMAIN1.CLIENT = "NYA"

            .Groups("Shipment Status").Visible = Not ScreenMode And InquiryMode
            .Groups("Shipment Selection").Visible = Not ScreenMode

            .Groups("Special Operations").Items("Force PTs to Balance").Visible = ScreenMode And edi_order And edi856_customer
            .Groups("Special Operations").Items("Force Cartons to Balance").Visible = ScreenMode And ((edi_order And edi856_customer) OrElse MaintenanceMode)
        End With

        '  lblStatus.Visible = ScreenMode

        'grdSOTSHIPX.Visible = Not tf
        tabSelect.Visible = Not tf

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.txtFor("ORDR_CUST_PO"), edi_order OrElse Not (EntryMode = "E" OrElse EntryMode = "N"))

        Set_Read_Only_for_ctl(Absx1.txtFor("CUST_CODE"), True)
        Set_Read_Only_for_ctl(Absx1.optFor("SHIP_ADDR_TYPE"), True)
        Set_Read_Only_for_ctl(Absx1.txtFor("SHIP_ADDR_CODE"), True)

        If ScreenMode Then
            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                {grdSOTPICK1, grdSOTPICK2, grdSOTCART1, grdSOTCART2}
                With grd.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowDelete = DefaultableBoolean.False
                    If (EntryMode = "N" OrElse EntryMode = "E") Then
                        .AllowUpdate = DefaultableBoolean.True
                    Else
                        .AllowUpdate = DefaultableBoolean.False
                    End If
                End With
            Next
            Setup_SOTPICK1() ' because allowupdate is toggled based on status of active pick1 record
            If MaintenanceMode Then
                grdSOTCART1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            End If

            Set_Read_Only(splHeader, Not (EntryMode = "N" OrElse EntryMode = "E"))
            If Not InquiryMode And Not MaintenanceMode Then
                If select_from_3PL_list Then
                    Set_Read_Only(splHeader.Panel1, True)
                Else
                    Set_Read_Only(splHeader.Panel1, False)
                End If
            End If

            Set_Read_Only(frmCodes, Not (EntryMode = "E" OrElse EntryMode = "N"))
            Set_Read_Only(frmDates, Not (EntryMode = "E" OrElse EntryMode = "N"))
            Set_Read_Only(grpSHIPTO, edi_order OrElse Not (EntryMode = "E" OrElse EntryMode = "N"))
            Set_Read_Only(grpHeaderInfo, MaintenanceMode _
                          OrElse Not (EntryMode = "E" OrElse EntryMode = "N") _
                          OrElse ASCMAIN1.CLIENT = "RGI")
            Set_Read_Only(grpBillTo, True)
            Set_Read_Only(MyBase.Absx1.txtFor("TERM_CODE"), True)

            Set_Read_Only(grpShippingWindow, Not MaintenanceMode OrElse Not (EntryMode = "E" OrElse EntryMode = "N") OrElse ASCMAIN1.CLIENT = "NYA")

            With grdSOTPICK2.DisplayLayout.Bands(0).Columns("STYLE_CODE")
                If (EntryMode = "N" OrElse EntryMode = "E") And (1 <> 1) Then ' HOW IN THE WORLD IS THIS TO BE PERMITTED?
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                    .Header.Appearance.BackColor2 = Drawing.Color.LightGray 'LightGreen
                End If
            End With

            With grdSOTPICK1.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"PICK_CNT_CARTONS", "PICK_TOTAL_WGT", "PICK_FREIGHT"}
                    .Columns(COLUMN_NAME).Hidden = edi_order AndAlso edi856_customer AndAlso ASCMAIN1.CLIENT <> "NYA" ' As per Debbie
                    If COLUMN_NAME <> "PICK_FREIGHT" Then .Columns(COLUMN_NAME & "_CALC").Hidden = Not (edi_order AndAlso edi856_customer)
                    ' NOTE THAT FRT IS NOT SHOWN IF edi_order And edi856_customer; ASSUMPTION IS THAT THERE WILL BE NO FRT IF EDI
                Next

                With .Columns("BILL_OF_LADING_NO")
                    If Not edi_order OrElse Absx1.optFor("SHIP_ADDR_TYPE").Value = "MK" And Not MaintenanceMode Then
                        .CellActivation = UltraWinGrid.Activation.AllowEdit
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    Else
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    End If
                End With
            End With

            chkBO.Checked = Not MaintenanceMode And Not (edi_customer) And (rowARTCUST1.Item("CUST_ALLOW_BACKORDER") & "" = "1")
            Setup_BO()
        Else
            Clear_Record()
        End If

        Absx1.txtFor("SHIP_ADDR_CODE").Visible = Not ScreenMode OrElse (Absx1.optFor("SHIP_ADDR_TYPE").Value = "DC")
        tabSOTPICK1.Tabs("Cartons").Visible = True '(edi856_customer And edi_order) OrElse (ASCMAIN1.DBS_COMPANY = "NYA" OrElse ASCMAIN1.DBS_SERVER = "NYA") '  And Not MaintenanceMode

        If tabSOTPICK1.Tabs("Cartons").Visible Then
            grdSOTCART2.Parent = splSOTCART1.Panel2
            splSOTPICK2.Panel2Collapsed = True
            grdSOTCART2.DisplayLayout.Bands(0).Columns("CART_NO").Hidden = True
            If Not ASCMAIN1.CLIENT = "NYA" Then
                'splSOTPICK1.Panel2Collapsed = True
            End If
        Else
            grdSOTCART2.Parent = splSOTPICK2.Panel2
            splSOTPICK2.Panel2Collapsed = False
            grdSOTCART2.DisplayLayout.Bands(0).Columns("CART_NO").Hidden = False
            splSOTPICK1.Panel2Collapsed = False
        End If

        Position_txtSTORE()
        ' lblBILL_OF_LADING_NO.Visible = Not (dst.Tables("SOTSHIP1").Rows.Count > 1)
        ' Absx1.txtFor("BILL_OF_LADING_NO").Visible = Not (dst.Tables("SOTSHIP1").Rows.Count > 1)
        If Not InquiryMode Then
            Set_Read_Only_for_ctl(Absx1.txtFor("BILL_OF_LADING_NO"), (dst.Tables("SOTSHIP1").Rows.Count > 1) OrElse select_from_3PL_list)
            'Set_Read_Only_for_ctl(Absx1.txtFor("REASON_CODE"), Not select_from_3PL_list)

            If select_from_3PL_list Then
                If Not (ASCMAIN1.DBS_COMPANY = "NYA" OrElse ASCMAIN1.DBS_SERVER = "NYA") Then
                    grdSOTPICK2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                End If
                grdSOTCART2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            End If
        End If
        grdSOTPICK1.DisplayLayout.Bands(0).Columns("BILL_OF_LADING_NO").Hidden = Not (dst.Tables("SOTSHIP1").Rows.Count > 1)

        If ASCMAIN1.CLIENT = "RGI" Then
            chkFactored.Visible = False
        End If

        If InquiryMode _
                AndAlso EntryMode = "V" _
                AndAlso ASCMAIN1.CLIENT = "NYA" _
                AndAlso dst.Tables("SOTSHIP1").Select("SHIP_STATUS <> 'F'").Length = 0 _
                AndAlso ASCMAIN1.USER_SECURITY_CODEs.Contains("WS") Then
            Set_Read_Only(txtSHIP_VIA_CODE, False)
            Set_Read_Only(txtFRT_TERMS, False)
            grdSOTCART1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdSOTCART1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grdSOTCART1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            Set_Read_Only(optPayor, False)
            Set_Read_Only(txt3PAccountNo, False)
            Set_Read_Only(txt3pCountry, False)
            Set_Read_Only(txt3PZipCode, False)
        End If

        With grdSOTINVHM.DisplayLayout.Override
            If Not InquiryMode AndAlso EntryMode = "E" Then
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            Else
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End If
        End With

        If ASCMAIN1.CLIENT = "RGI" Then
            numInsureValue.ReadOnly = True
        End If

    End Sub

    Sub Clear_Record()

        'Absx1.txtFor("SHIP_BOL_NO").Text = ""
        'Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("PICK_NO").Clear()

        txtReason.Text = String.Empty
        txtContact.Text = String.Empty
        txtemail.Text = String.Empty

        optPayor.Value = "O"
        txt3PAccountNo.Tag = String.Empty
        txt3PAccountNo.Clear()
        txt3pCountry.Clear()
        txt3PZipCode.Clear()

        lblFCR.Visible = False
        txtBTB_BOL_NO.Visible = False

        CUST_CODE = String.Empty
        ORDR_GROUP_NO = String.Empty
        ORDR_CUST_PO = String.Empty

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTPICK1", "SOTPICK2", "SOTPICK4", _
             "SOTINVH1", "SOTINVH2", "SOTINVH9", "SOTINVHM", "ARTOPEN1", _
             "SOTCART1", "SOTCART2", "SOTCART3", "SOTCARTX", _
             "SOTORDR1", "SOTORDR2", "SOTORDR5", "SOTORDR5_BT", "SOTORDXR", _
             "SOTSHIP0", "SOTSHIP1", "SOTSHIP3", "SOTSHIP4", "SOTSHIP6", "SOTRNGA1", _
             "SOTORDC1", "SOTORDC2", _
             "ARTCCPA1", "ARTCCPA2", "ARTCCPDA", _
             "EDT945T1", "EDT945T2", _
             "WHTSHPC1", "WHTSHPC2", "WHTSHPC3", "WHTSHPC5", "WHTSHPCC", "WHTSHPCS", "WHTSHPCP"}
            If dst.Tables.Contains(TABLE_NAME) Then
                dst.Tables(TABLE_NAME).Rows.Clear()
            Else
                'Stop
            End If
        Next

        EnforceConstraints(True)

        select_from_3PL_list = False

        Load_SOTSHIPX()

        ' Vandale loads the Shipments once. It goes out and gets them from another location.
        If ASCMAIN1.CLIENT = "VAN" Then

        Else
            Fill_Records("WHT3PLS1")
            Sort_grdColumns(grdWHT3PLS1, "SHIP_BOL_NO")
            If ASCMAIN1.CLIENT = "NYA" Then
                Dim SHIP_BOL_NOs As String = String.Empty
                ' 5/23/2013 - At some point in time the release of orders to the 3PL stopped populating this field. Taking care of my customer.
                For Each row As DataRow In ASCDATA1.SelectDistinct("WHT3PLS1", New String() {"SHIP_BOL_NO"}).Rows
                    SHIP_BOL_NOs &= ", '" & row.Item("SHIP_BOL_NO") & "'"
                Next
                If SHIP_BOL_NOs.Length > 0 Then
                    SHIP_BOL_NOs = SHIP_BOL_NOs.Substring(1)
                    Try
                        ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_PICK_PRINTED = SYSDATE WHERE SHIP_PICK_PRINTED IS NULL " _
                            & " AND SHIP_BOL_NO IN (" & SHIP_BOL_NOs & ")"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                    Catch ex As Exception

                    End Try
                End If

            End If
        End If

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Permit Price Change"), UltraWinToolbars.StateButtonTool)
        tlb_sbt.Checked = False

        optAddressType.Value = "O"
        chkInsureShipment.Checked = False
        chkInsureShipment_CheckedChanged(Nothing, Nothing)
        chkSaturday.Checked = False
        chkSignature.Checked = False
        optPayor.Value = "O"
        select_from_3PL_list = False
        processing_select_from_3PL_list = False
        selectedEDI_BOL_NO = String.Empty
        selectedMasterEDI_BOL_NO = String.Empty
        RecreateLabel = False
        AddPPAToExistingFreight = True

        MyBase.Absx1.txtFor("BILL_OF_LADING_NO").Enabled = True
        MyBase.Absx1.dteFor("SHIP_DATE_SHIPPED").Enabled = True
        MyBase.Absx1.dteFor("INV_DATE").Enabled = True

        If ASCMAIN1.CLIENT = "RGI" Then
            numInsureValue.ReadOnly = True
            numInsureValue.Appearance.BackColor = Drawing.Color.White
            numInsureValue.Appearance.BackColorDisabled = Drawing.Color.White
        End If

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)
        Dim BOL_NO As String = String.Empty

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        ToggleDataTableExpressions(False)

        If EntryMode = "N" Then
        Else
            rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)

            Fill_Records("EDTTRPMC", CUST_CODE)
            edi_customer = (dst.Tables("EDTTRPMC").Rows.Count <> 0)
            Dim rowEDTTRPMC As DataRow = dst.Tables("EDTTRPMC").Rows.Find(New Object() {CUST_CODE, "856"})
            edi856_customer = rowEDTTRPMC IsNot Nothing AndAlso rowEDTTRPMC.Item("EDI_STATUS") & "" = "P"
            lblASN.Visible = (edi856_customer And edi_order)

            'If MaintenanceMode Then
            Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)
            ORDR_SHIP_DATE = rowSOTORDR0.Item("ORDR_SHIP_DATE")
            ORDR_CANCEL_DATE = rowSOTORDR0.Item("ORDR_CANCEL_DATE")
            dteORDR_SHIP_DATE.Value = ORDR_SHIP_DATE
            dteORDR_CANCEL_DATE.Value = ORDR_CANCEL_DATE
            txtReason.Text = ""
            'End If

            ASCMAIN1.sql = "Select Count (*) from SOTORDR1" _
                & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and ORDR_SOURCE = 'E'"
            edi_order = (Val(ASCDATA1.GetDataValue) <> 0)
            If edi_order Then
                lblSource.Text = "EDI"
            Else
                lblSource.Text = "Manual"
            End If

            ASCMAIN1.sql = "Select Distinct ORDR_SOURCE from SOTORDR1" _
                & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            Dim rowORDR_SOURCE As DataRow = ASCDATA1.GetDataRow
            ORDR_SOURCE = rowORDR_SOURCE.Item("ORDR_SOURCE")

            Dim sqlwhere_SOTSHIP1 As String = "" _
                & "   and SOTSHIP1.SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIP0 & ")" & vbCrLf _
                & IIf(InquiryMode, _
                      "", _
                      "" _
                        & "   and SOTSHIP1.SHIP_STATUS = 'P'" & vbCrLf _
                        & "   and SOTSHIP1.SHIP_PICK_PRINTED is Not Null")

            ASCMAIN1.sql = sqlSOTSHIPX & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTSHIP1", "", True, ASCMAIN1.sql)
            If dst.Tables("SOTSHIP1").Rows.Count <> SHIP_BOL_NOs.Count Then Stop ' NEED AN ABORT LOAD FEATURE IN STDS

            If Not InquiryMode AndAlso Not MaintenanceMode AndAlso Not select_from_3PL_list Then
                If ASCMAIN1.CLIENT = "NYA" Then
                    BOL_NO = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTSHIPB.BOL_NO"), "0" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
                    For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                        rowSOTSHIP1.Item("BILL_OF_LADING_NO") = BOL_NO
                    Next
                    MyBase.Absx1.txtFor("BILL_OF_LADING_NO").Text = BOL_NO
                    MyBase.Absx1.txtFor("BILL_OF_LADING_NO").Appearance.BackColorDisabled = Drawing.Color.White
                    MyBase.Absx1.txtFor("BILL_OF_LADING_NO").Enabled = False
                End If
            End If

            ASCMAIN1.sql = sqlSOTPICK1 & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1
            If Not InquiryMode Then
                ASCMAIN1.sql &= " and SOTPICK1.PICK_STATUS = 'P'"
            End If
            Fill_Records("SOTPICK1", "", True, ASCMAIN1.sql)

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
                rowSOTPICK1.Item("SELECTED") = "1"
            Next

            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO in" & vbCrLf _
               & "(Select Distinct ORDR_NO from SOTPICK1 where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIP0 & "))"
            Fill_Records("SOTORDR1", String.Empty, True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select SOTORDR2.*, SOTORDR2.ORDR_UNIT_PRICE ORDR_UNIT_PRICE_ORIG from SOTORDR2 where ORDR_NO in" & vbCrLf _
              & "(Select Distinct ORDR_NO from SOTPICK1 where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIP0 & "))"
            Fill_Records("SOTORDR2", String.Empty, True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from SOTORDR5 where ORDR_NO in" & vbCrLf _
                & "(Select Distinct ORDR_NO from SOTPICK1 where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIP0 & "))" & vbCrLf _
                & " and CUST_ADDR_TYPE = 'ST'"
            Fill_Records("SOTORDR5", "", True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from SOTORDR5 where ORDR_NO in" & vbCrLf _
                & "(Select Distinct ORDR_NO from SOTPICK1 where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIP0 & "))" & vbCrLf _
                & " and CUST_ADDR_TYPE = 'BT'"
            Fill_Records("SOTORDR5_BT", "", True, ASCMAIN1.sql)

            If dst.Tables("SOTORDR5_BT").Rows.Count = 0 Then
                Dim rowSOTORDR5_BT As DataRow = dst.Tables("SOTORDR5_BT").NewRow
                rowSOTORDR5_BT.Item("ORDR_NO") = "XXX"
                rowSOTORDR5_BT.Item("CUST_ADDR_TYPE") = "BT"
                rowSOTORDR5_BT.Item("CUST_ADDR_CODE") = rowARTCUST1.Item("CUST_CODE")
                rowSOTORDR5_BT.ITEM("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
                rowSOTORDR5_BT.ITEM("CUST_ADDR1") = rowARTCUST1.Item("CUST_ADDR1")
                rowSOTORDR5_BT.ITEM("CUST_ADDR2") = rowARTCUST1.Item("CUST_ADDR2")
                rowSOTORDR5_BT.ITEM("CUST_CITY") = rowARTCUST1.Item("CUST_CITY")
                rowSOTORDR5_BT.ITEM("CUST_STATE") = rowARTCUST1.Item("CUST_STATE")
                rowSOTORDR5_BT.ITEM("CUST_ZIP_CODE") = rowARTCUST1.Item("CUST_ZIP_CODE")
                rowSOTORDR5_BT.ITEM("CUST_COUNTRY") = rowARTCUST1.Item("CUST_COUNTRY")
                rowSOTORDR5_BT.ITEM("CUST_CONTACT") = rowARTCUST1.Item("CUST_CONTACT")
                rowSOTORDR5_BT.ITEM("CUST_PHONE") = rowARTCUST1.Item("CUST_PHONE")
                rowSOTORDR5_BT.ITEM("CUST_EXT") = rowARTCUST1.Item("CUST_EXT")
                rowSOTORDR5_BT.ITEM("CUST_FAX") = rowARTCUST1.Item("CUST_FAX")
                rowSOTORDR5_BT.ITEM("CUST_EMAIL") = rowARTCUST1.Item("CUST_EMAIL")
                rowSOTORDR5_BT.ITEM("CUST_ADDR3") = rowARTCUST1.Item("CUST_ADDR3")
                dst.Tables("SOTORDR5_BT").Rows.Add(rowSOTORDR5_BT)
            End If

            ASCMAIN1.sql = "Select * from SOTINVHM where INV_NO in" & vbCrLf _
                & "(Select Distinct INV_NO from SOTPICK1 where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIP0 & "))"
            Fill_Records("SOTINVHM", String.Empty, True, ASCMAIN1.sql)

            If Not InquiryMode Then
                For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                    For Each COLUMN_NAME As String In New String() _
                        {"FRT_TERMS", "SHIP_VIA_CODE"}
                        If rowSOTSHIP1.Item(COLUMN_NAME) & "" = "" Then rowSOTSHIP1.Item(COLUMN_NAME) = rowARTCUST1.Item(COLUMN_NAME)
                    Next
                    If rowSOTSHIP1.Item("TERM_CODE") & "" = "" Then
                        If rowARTCUST1.Item("TERM_CODE") & "" <> "" Then
                            rowSOTSHIP1.Item("TERM_CODE") = rowARTCUST1.Item("TERM_CODE")
                            For Each rowSOTPICK1 As DataRow In rowSOTSHIP1.GetChildRows("SOTSHIP1_SOTPICK1")
                                rowSOTPICK1.Item("TERM_CODE") = rowARTCUST1.Item("TERM_CODE")
                            Next
                        End If
                    End If
                Next

                Dim rowSOTPICK1_0 As DataRow = dst.Tables("SOTPICK1").Rows(0)
                For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                    For Each COLUMN_NAME As String In New String() _
                        {"ORDR_DEPT", "SREP_CODE", "SREP2_CODE", "TERM_CODE"}
                        If rowSOTSHIP1.Item(COLUMN_NAME) & "" = "" Then rowSOTSHIP1.Item(COLUMN_NAME) = rowSOTPICK1_0.Item(COLUMN_NAME)
                    Next
                Next
            End If

            Dim row As DataRow = dst.Tables("SOTSHIP1").Rows(0)
            rowSOTSHIP0 = dst.Tables("SOTSHIP0").NewRow
            For i As Integer = 0 To dst.Tables("SOTSHIP0").Columns.Count - 1
                rowSOTSHIP0.Item(i) = row.Item(i)
            Next
            dst.Tables("SOTSHIP0").Rows.Add(rowSOTSHIP0)

            rowSOTSHIP0_ORIG = dst.Tables("SOTSHIP0").NewRow
            rowSOTSHIP0_ORIG.ItemArray = rowSOTSHIP0.ItemArray

            chkFactored.Checked = (dst.Tables("SOTPICK1").Select("CUST_FACTOR_IND = '1'").Length > 0)

            ASCMAIN1.sql = sqlSOTPICK2 & vbCrLf _
                & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & sqlwhere_SOTSHIP1
            If Not InquiryMode Then
                ASCMAIN1.sql &= " and SOTPICK1.PICK_STATUS = 'P'"
            End If
            Fill_Records("SOTPICK2", "", True, ASCMAIN1.sql)

            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                ' need to go thru details looking for a price which has fractions of a cent
                Dim extra_decimals As Boolean = False
                For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("")
                    Dim PICK_UNIT_PRICE As Decimal = Val(rowSOTPICK2.Item("PICK_UNIT_PRICE") & "")
                    If PICK_UNIT_PRICE <> Val(Format(PICK_UNIT_PRICE, "#.00")) Then
                        extra_decimals = True
                        Exit For
                    End If
                Next

                With grdSOTPICK2.DisplayLayout.Bands(0).Columns("PICK_UNIT_PRICE")
                    If extra_decimals Then
                        .Format = "#,##0.0000"
                    Else
                        .Format = "#,##0.00"
                    End If
                End With
                With grdSOTPICK2.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE")
                    If extra_decimals Then
                        .Format = "#,##0.0000"
                    Else
                        .Format = "#,##0.00"
                    End If
                End With
            End If


            Fill_Records("SOTORDR9", ORDR_GROUP_NO)

            ASCMAIN1.sql = "Select SOTCART1.*" & vbCrLf _
                & " from SOTCART1,SOTPICK1,SOTSHIP1" & vbCrLf _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTCART1", "", True, ASCMAIN1.sql)

            Dim CART_SEQ As Int32 = 1
            For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
                ' default to packaging code 31 - Our package
                rowSOTCART1.Item("PACKAGING_TYPE") = 31
                rowSOTCART1.Item("CART_SEQ") = CART_SEQ
                CART_SEQ += 1
                If edi_order And edi856_customer Then
                    rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = rowSOTCART1.Item("CART_TOTAL_WGT_CALC")
                ElseIf ASCMAIN1.CLIENT = "RGI" Then
                    ' For RGI set weight as carton type
                    rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = rowSOTCART1.Item("CART_TOTAL_WGT_CALC")
                    If Val(rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") & String.Empty) = 0 Then
                        rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = 1
                    End If
                    rowSOTCART1.Item("PKG_CODE") = "XX"
                End If
            Next

            ASCMAIN1.sql = "Select SOTCART2.*, SOTCART1.PICK_NO, SOTCART2.QTY_PACKED QTY_PACKED_ORIG" & vbCrLf _
                & " from SOTCART2,SOTCART1,SOTPICK1,SOTSHIP1" & vbCrLf _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTCART2", "", True, ASCMAIN1.sql)

            If Not select_from_3PL_list Then
                Load_Record_Ancillary()
            End If
        End If

        Sort_grdColumns(grdSOTINVHM, "INV_NO,INV_MNO")
        MyBase.Populate_Controls_with_Parents("SREP_CODE", MyBase.Absx1.txtFor("SREP_CODE"))
        MyBase.Populate_Controls_with_Parents("SREP2_CODE", MyBase.Absx1.txtFor("SREP2_CODE"))
        MyBase.Populate_Controls_with_Parents("TERM_CODE", MyBase.Absx1.txtFor("TERM_CODE"))

        ' Show Only for NYA and Warehouse 52
        lblFCR.Visible = ASCMAIN1.CLIENT = "NYA" AndAlso dst.Tables("SOTSHIP1").Select("WHSE_CODE = '52'").Length > 0
        txtBTB_BOL_NO.Visible = ASCMAIN1.CLIENT = "NYA" AndAlso dst.Tables("SOTSHIP1").Select("WHSE_CODE = '52'").Length > 0

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("")
    End Sub

    ''' <summary>
    ''' This is here so Non 3PL and #PL can do their work then have this done
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Load_Record_Ancillary()

        If Not InquiryMode Then
            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("")
                rowSOTPICK2.Item("PICK_UNIT_PRICE") = Val(rowSOTPICK2.Item("ORDR_UNIT_PRICE") & "")

                rowSOTPICK2.Item("PICK_QTY_CONF") = Val(rowSOTPICK2.Item("PICK_QTY_CONF") & String.Empty)
                rowSOTPICK2.Item("PICK_QTY_CANC") = Val(rowSOTPICK2.Item("PICK_QTY_CANC") & String.Empty)
                rowSOTPICK2.Item("PICK_QTY_BACK") = Val(rowSOTPICK2.Item("PICK_QTY_BACK") & String.Empty)

                If Not select_from_3PL_list Then
                    rowSOTPICK2.Item("PICK_QTY_CONF") = rowSOTPICK2.Item("PICK_QTY")
                    rowSOTPICK2.Item("PICK_QTY_CANC") = 0
                    rowSOTPICK2.Item("PICK_QTY_BACK") = 0
                Else
                    ' If 3PL the cancel unshipped items from the pick ticket
                    If Val(rowSOTPICK2.Item("PICK_QTY_CONF") & String.Empty) < rowSOTPICK2.Item("PICK_QTY") Then
                        rowSOTPICK2.Item("PICK_QTY_CANC") = rowSOTPICK2.Item("PICK_QTY") - Val(rowSOTPICK2.Item("PICK_QTY_CONF") & String.Empty)
                    End If
                End If
            Next
        End If

        Dim INV_SALES As Decimal = 0
        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("")
            rowSOTPICK2.Item("PICK_UNIT_PRICE") = Val(rowSOTPICK2.Item("ORDR_UNIT_PRICE") & "")
            Dim ORDR_QTY_SHIP As Int64 = Val(rowSOTPICK2.Item("PICK_QTY_CONF") & "")
            Dim PICK_UNIT_PRICE As Decimal = Val(rowSOTPICK2.Item("PICK_UNIT_PRICE") & "")
            INV_SALES = INV_SALES + (ORDR_QTY_SHIP * PICK_UNIT_PRICE)
        Next
        numInsureValue.Value = INV_SALES

        If Not select_from_3PL_list AndAlso Not InquiryMode Then
            dteINV_DATE.DateTime = DateTime.Now
            dteSHIP_DATE_SHIPPED.DateTime = DateTime.Now
        End If

        ToggleDataTableExpressions(True)

        dst.Tables("SOTPICK1").AcceptChanges()
        dst.Tables("SOTPICK2").AcceptChanges()

        Sort_grdColumns(grdSOTPICK1, "PICK_NO")
        Setup_SOTPICK1()

        clsPrice_Change = Nothing

        Select Case rowSOTSHIP0.Item("SHIP_STATUS")
            Case "P"
                lblStatus.Text = "In Pick"
            Case "F"
                lblStatus.Text = "Shipped"
            Case "C"
                lblStatus.Text = "Cancelled"
            Case Else
                lblStatus.Text = "Status Unknown"
        End Select

        If EntryMode = "L" Then
            lblINIT_DATE.Text = "Confirmed by " & rowSOTSHIP0.Item("LAST_OPER") & " on " & Format(rowSOTSHIP0.Item("LAST_DATE"), "MM/dd/yy HH:mm")
        Else
            lblINIT_DATE.Text = "Confirmed by " & ASCMAIN1.USER_ID & " on " & Format(Now, "MM/dd/yy HH:mm")
        End If

        Display_Totals()

        Dim GL_PARM_CURR_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & ""
        CURR_CODE = rowARTCUST1.Item("CURR_CODE") & ""
        If CURR_CODE = "" OrElse CURR_CODE = GL_PARM_CURR_CODE Then
            CURR_CODE = GL_PARM_CURR_CODE
            CURR_EXCH_RATE = 1
            GST_TAX = 0
        Else
            Dim rowTATCURR1 As DataRow = LookUp("TATCURR1", rowARTCUST1.Item("CURR_CODE"))
            CURR_CODE = rowTATCURR1.Item("CURR_CODE")
            CURR_EXCH_RATE = rowTATCURR1.Item("CURR_EXCH_CUR")
            GST_TAX = 0.07
        End If

        dst.Tables("SOTCARTX").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTPICK2"), New String() {"PICK_NO", "ORDR_NO", "ORDR_LNO"}).Rows
            dst.Tables("SOTCARTX").Rows.Add(New Object() {row.Item("PICK_NO"), row.Item("ORDR_NO"), row.Item("ORDR_LNO")})
        Next

        grdSOTPICK1.DisplayLayout.Bands(0).Summaries.Clear()
        If dst.Tables("SOTPICK1").Rows.Count = 1 Then
            splSOTPICK1.SplitterDistance = 80 + grdSOTPICK1.Rows(0).Height * 1
            txtStore.Visible = False
        Else
            'CANT WE JUST CREATE THE SUMMARIES ONCE AND THEN HIDE THEM?
            Create_Summary(grdSOTPICK1, "PICK_NO", "Count")
            Create_Summary(grdSOTPICK1, New String() {"SELECTED", "PICK_CNT_CARTONS", "PICK_TOTAL_WGT", "PICK_FREIGHT", "TOTAL_CARTONS", "TOTAL_CUBE"})
            Create_Summary(grdSOTPICK1, New String() {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"})
            Create_Summary(grdSOTPICK1, New String() {"PICK_AMT", "PICK_AMT_CONF", "PICK_AMT_CANC", "PICK_AMT_BACK"})
            splSOTPICK1.SplitterDistance = 80 + grdSOTPICK1.Rows(0).Height * 4
            txtStore.Visible = True
        End If

        If MaintenanceMode Then
            txtemail.Text = ASCMAIN1.USER_EMAIL
            txtContact.Text = ASCMAIN1.USER_NAME
        End If

        If ASCMAIN1.CLIENT = "NYA" Then
            For Each col As String In New String() {"PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"}
                grdSOTPICK2.DisplayLayout.Bands(0).Columns(col).CellActivation = UltraWinGrid.Activation.NoEdit
            Next
        End If

        Display_Totals()
        EnforceConstraints(True)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        Dim inTransaction As Boolean = False
        Dim shipLabels As List(Of String) = New List(Of String)
        Dim ErrorMessage As String = String.Empty

        Dim rowSOTPICK1 As DataRow = Nothing
        Dim rowSOTORDR1 As DataRow = Nothing

        Try
            refreshScreen = True

            ' Need to see if we needs to get labels from carrier.
            ' This is done first since some Freight Terms require the Customer Absorbs the Freight Cost
            If commonCarrier AndAlso Not select_from_3PL_list Then
                ASCMAIN1.Progress("Generating Shipping label(s)", "")
                If RequestShippingLabel(shipLabels, ErrorMessage, False) Then
                    ' May want to print these later - we will see
                    ' Moved below incase of error create invoice the user does not ship shipment without billing 
                Else
                    ' ErrorMessage should have the error text available
                    ' What to do if an error occurs in the requesting of a label??
                    If ErrorMessage.Length > 0 Then
                        MessageBox.Show("Update Aborted! The following error occurred when processing the shipping label: " & ErrorMessage, "Ship Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        MessageBox.Show("Update Aborted! An error occurred when processing the shipping label.", "Ship Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                    refreshScreen = False
                    Exit Sub
                End If
            End If

            If RecreateLabel Then
                Try
                    For Each row As DataRow In dst.Tables("SOTSHIP1").Rows
                        ASCMAIN1.sql = "UPDATE SOTSHIP1 SET SHIP_VIA_CODE = '" & txtSHIP_VIA_CODE.Text & "', FRT_TERMS = '" & txtFRT_TERMS.Text & "'" _
                             & " WHERE SHIP_BOL_NO = '" & row.Item("SHIP_BOL_NO") & "'"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                    Next

                Catch ex As Exception

                End Try

                For Each shippingLabel As String In shipLabels
                    If shippingLabel.Trim.Length > 0 Then PrintLabel(shippingLabel)
                Next
                Me.Cursor = Cursors.Default
                Exit Sub
            End If

            ASCMAIN1.Progress("Now Updating ...")

            Dim shippingFreight As Decimal = 0

            Dim INV_STAX As Decimal = 0
            Dim INV_MISC_CHG As Decimal = Val(dst.Tables("SOTINVHM").Compute("SUM(INV_MISC_CHG)", "") & String.Empty)
            Dim INV_MISC_CHG_PICK_NO As String = String.Empty
            Dim CreditCardProcessed As Boolean = True
            Dim ship_ref As String = String.Empty

            ' Capture Credit Card Approved $$
            Try
                dst.Tables("TATEVNT1").Rows.Clear()

                dst.Tables("SOTORDC1").Rows.Clear()
                dst.Tables("SOTORDC2").Rows.Clear()

                For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select("SELECTED = '1' AND ISNULL(CCPA_NO_ORDR, '') <> ''", "SHIP_BOL_NO")

                    ASCMAIN1.Progress("Processing Credit Card", "")

                    Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                    Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                    Dim CCPA_NO_ORDR As String = rowSOTPICK1.Item("CCPA_NO_ORDR")

                    If dst.Tables("SOTORDC1").Select("ORDR_NO = '" & ORDR_NO & "'").Length = 0 Then
                        Fill_Records("SOTORDC1", ORDR_NO)
                        Fill_Records("SOTORDC2", ORDR_NO)
                    End If

                    ' This is here for older sales orders
                    If dst.Tables("SOTORDC1").Select("CCPA_NO = '" & CCPA_NO_ORDR & "'").Length = 0 Then
                        Dim rowSOTORDC1 As DataRow = Nothing
                        Dim rowARTCCPA1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCCPA1 WHERE CCPA_NO = :PARM1", "V", CCPA_NO_ORDR)
                        rowSOTORDC1 = dst.Tables("SOTORDC1").NewRow
                        rowSOTORDC1.Item("ORDR_NO") = ORDR_NO
                        rowSOTORDC1.Item("TRANS_NO") = Val(ASCDATA1.GetDataValue("SELECT MAX(TRANS_NO) FROM SOTORDC1 WHERE ORDR_NO = '" & ORDR_NO & "'") & String.Empty) + 1
                        rowSOTORDC1.Item("TRANS_TYPE") = "C"
                        rowSOTORDC1.Item("TRANS_DATE") = rowARTCCPA1.Item("INIT_DATE")
                        rowSOTORDC1.Item("CCPA_NO") = CCPA_NO_ORDR
                        rowSOTORDC1.Item("CCPA_STATUS") = rowARTCCPA1.Item("CCPA_STATUS")
                        rowSOTORDC1.Item("AMOUNT") = rowARTCCPA1.Item("CCPA_AMT")
                        rowSOTORDC1.Item("BALANCE") = rowARTCCPA1.Item("CCPA_AMT")
                        rowSOTORDC1.Item("ACTIVE_IND") = "1"
                        rowSOTORDC1.Item("INIT_OPER") = rowARTCCPA1.Item("INIT_OPER")
                        dst.Tables("SOTORDC1").Rows.Add(rowSOTORDC1)
                        ' SAVE RIGHT AWAY.
                        Update_Record_TDA("SOTORDC1")
                    End If

                    ' Place Misc Charge on random Pick Ticket
                    ' Charge INV_MISC_CHG only once
                    If INV_MISC_CHG_PICK_NO.Length = 0 AndAlso INV_MISC_CHG <> 0 Then
                        rowSOTPICK1.Item("INV_MISC_CHG") = INV_MISC_CHG
                        INV_MISC_CHG_PICK_NO = PICK_NO
                    End If

                    CCPA_NO_ORDR = CCPA_NO_ORDR.Trim
                    If CCPA_NO_ORDR.Length = 0 Then Continue For

                    If rowSOTPICK1.Item("CCPA_NO") & String.Empty <> String.Empty Then
                        ' Credit card against pick ticket already processed. May have been an error in the code
                        Continue For
                    End If

                    Dim chargeAmount As Decimal = Val(rowSOTPICK1.Item("PICK_AMT_CONF") & String.Empty)
                    INV_STAX = 0

                    shippingFreight = Val(rowSOTPICK1.Item("PICK_FREIGHT") & String.Empty) + Val(rowSOTPICK1.Item("ORDR_FOB") & String.Empty) ' + Val(rowSOTPICK1.Item("PPA_FREIGHT") & String.Empty)
                    chargeAmount += shippingFreight + INV_MISC_CHG + INV_STAX

                    ' do we need to add additional funds??
                    For Each row As DataRow In dst.Tables("SOTORDC1").Select("TRANS_TYPE = 'A' AND ACTIVE_IND = '1' AND AMOUNT > 0")
                        chargeAmount += row.Item("AMOUNT")
                        Dim rowSOTORDC2 As DataRow = dst.Tables("SOTORDC2").NewRow
                        rowSOTORDC2.Item("ORDR_NO") = row.Item("ORDR_NO")
                        rowSOTORDC2.Item("TRANS_NO") = row.Item("TRANS_NO")
                        rowSOTORDC2.Item("TRANS_LNO") = Val(dst.Tables("SOTORDC2").Compute("MAX(TRANS_LNO)", "ORDR_NO = '" & row.Item("ORDR_NO") & "' AND TRANS_NO = " & row.Item("TRANS_NO")) & String.Empty) + 1
                        rowSOTORDC2.Item("PICK_NO") = PICK_NO ' temp to convert to Inv No
                        'rowSOTORDC2.Item("INV_DATE") = ""
                        'rowSOTORDC2.Item("INV_AMOUNT") = ""
                        rowSOTORDC2.Item("AMOUNT_APPLIED") = row.Item("AMOUNT")
                        dst.Tables("SOTORDC2").Rows.Add(rowSOTORDC2)
                    Next

                    ' On Account
                    Dim onAccountAmount As Decimal = 0
                    For Each row As DataRow In dst.Tables("SOTORDC1").Select("TRANS_TYPE = 'O' AND ACTIVE_IND = '1' AND BALANCE > 0")
                        If chargeAmount <= 0 Then
                            Exit For
                        End If

                        onAccountAmount = row.Item("BALANCE")
                        If onAccountAmount > chargeAmount Then
                            onAccountAmount = chargeAmount
                        End If

                        Dim rowSOTORDC2 As DataRow = dst.Tables("SOTORDC2").NewRow
                        rowSOTORDC2.Item("ORDR_NO") = row.Item("ORDR_NO")
                        rowSOTORDC2.Item("TRANS_NO") = row.Item("TRANS_NO")
                        rowSOTORDC2.Item("TRANS_LNO") = Val(dst.Tables("SOTORDC2").Compute("MAX(TRANS_LNO)", "ORDR_NO = '" & row.Item("ORDR_NO") & "' AND TRANS_NO = " & row.Item("TRANS_NO")) & String.Empty) + 1
                        rowSOTORDC2.Item("PICK_NO") = PICK_NO ' temp to convert to Inv No
                        'rowSOTORDC2.Item("INV_DATE") = ""
                        'rowSOTORDC2.Item("INV_AMOUNT") = ""
                        rowSOTORDC2.Item("AMOUNT_APPLIED") = onAccountAmount
                        dst.Tables("SOTORDC2").Rows.Add(rowSOTORDC2)
                        chargeAmount -= onAccountAmount
                        row.Item("BALANCE") -= onAccountAmount
                    Next

                    ' Deposits
                    Dim depositAmount As Decimal = 0
                    For Each row As DataRow In dst.Tables("SOTORDC1").Select("TRANS_TYPE = 'D' AND ACTIVE_IND = '1' AND BALANCE > 0")
                        If chargeAmount <= 0 Then
                            Exit For
                        End If

                        depositAmount = row.Item("BALANCE")
                        If depositAmount > chargeAmount Then
                            depositAmount = chargeAmount
                        End If

                        Dim rowSOTORDC2 As DataRow = dst.Tables("SOTORDC2").NewRow
                        rowSOTORDC2.Item("ORDR_NO") = row.Item("ORDR_NO")
                        rowSOTORDC2.Item("TRANS_NO") = row.Item("TRANS_NO")
                        rowSOTORDC2.Item("TRANS_LNO") = Val(dst.Tables("SOTORDC2").Compute("MAX(TRANS_LNO)", "ORDR_NO = '" & row.Item("ORDR_NO") & "' AND TRANS_NO = " & row.Item("TRANS_NO")) & String.Empty) + 1
                        rowSOTORDC2.Item("PICK_NO") = PICK_NO ' temp to convert to Inv No
                        'rowSOTORDC2.Item("INV_DATE") = ""
                        'rowSOTORDC2.Item("INV_AMOUNT") = ""
                        rowSOTORDC2.Item("AMOUNT_APPLIED") = depositAmount
                        dst.Tables("SOTORDC2").Rows.Add(rowSOTORDC2)
                        chargeAmount -= depositAmount
                        row.Item("BALANCE") -= depositAmount
                    Next

                    If chargeAmount > 0 Then
                        Try
                            Dim ResponseText As String = String.Empty
                            Dim rowSOTORDC1 As DataRow = Nothing

                            rowSOTORDC1 = dst.Tables("SOTORDC1").Select("CCPA_NO = '" & CCPA_NO_ORDR & "'")(0)
                            Dim CCPA_NO As String = ProcessCreditCardAuthorization(CCPA_NO_ORDR, chargeAmount, shippingFreight, INV_STAX, ResponseText)
                            CreditCardProcessed = CCPA_NO.Length > 0 AndAlso CreditCardProcessed

                            If CCPA_NO.Length > 0 Then
                                ' if we made a sale then point to the new sale record.
                                If dst.Tables("SOTORDC1").Select("CCPA_NO = '" & CCPA_NO & "'").Length > 0 Then
                                    rowSOTORDC1 = dst.Tables("SOTORDC1").Select("CCPA_NO = '" & CCPA_NO & "'")(0)
                                End If
                                ' This is done to preserve credit card transactions if the code causes an error after this point
                                MyBase.BeginTrans()
                                rowSOTPICK1.Item("CCPA_NO") = CCPA_NO
                                ASCDATA1.ExecuteSQL("Update SOTPICK1 set CCPA_NO = '" & CCPA_NO & "' where PICK_NO = '" & PICK_NO & "'")

                                Dim rowSOTORDC2 As DataRow = dst.Tables("SOTORDC2").NewRow
                                rowSOTORDC2.Item("ORDR_NO") = rowSOTORDC1.Item("ORDR_NO")
                                rowSOTORDC2.Item("TRANS_NO") = rowSOTORDC1.Item("TRANS_NO")
                                rowSOTORDC2.Item("TRANS_LNO") = Val(dst.Tables("SOTORDC2").Compute("MAX(TRANS_LNO)", "ORDR_NO = '" & rowSOTORDC1.Item("ORDR_NO") & "' AND TRANS_NO = " & rowSOTORDC1.Item("TRANS_NO")) & String.Empty) + 1
                                rowSOTORDC2.Item("PICK_NO") = PICK_NO ' temp to convert to Inv No
                                'rowSOTORDC2.Item("INV_DATE") = ""
                                'rowSOTORDC2.Item("INV_AMOUNT") = ""
                                rowSOTORDC2.Item("AMOUNT_APPLIED") = chargeAmount
                                rowSOTORDC1.Item("BALANCE") -= chargeAmount
                                dst.Tables("SOTORDC2").Rows.Add(rowSOTORDC2)

                                ' Record Transaction Number in Order Header. Will be placed in Invoice Header
                                Dim rowARTCCPA1 As DataRow = LookUp("ARTCCPA1", CCPA_NO)
                                If rowARTCCPA1 IsNot Nothing Then
                                    rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
                                    rowSOTORDR1.Item("CC_TRANS_ID") = rowARTCCPA1.Item("TRANS_ID")
                                End If

                                TAC.TACMAIN1.Record_Event("SOTORDR1", _
                                                          rowSOTPICK1.Item("ORDR_NO"), _
                                                          Now, _
                                                          ASCMAIN1.USER_ID, _
                                                          "CCCHG", _
                                                          "Credit card charged: " & Format(chargeAmount, "#,##0.00"))

                                Update_Record_TDA("SOTORDC1")
                                Update_Record_TDA("SOTORDC2")
                                MyBase.CommitTrans()
                            Else
                                MessageBox.Show("Credit Card Could not be captured for the following reason: " & ResponseText, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                CreditCardProcessed = False
                                Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
                                rowTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
                                rowTATEVNT1.Item("TABLE_KEY") = rowSOTPICK1.Item("ORDR_NO")
                                rowTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
                                rowTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                                rowTATEVNT1.Item("EVENT_TYPE") = "CCP"
                                rowTATEVNT1.Item("EVENT_DESC") = "Credit Card Error: " & ResponseText
                                rowTATEVNT1.Item("EVENT_KEY") = ""
                                rowTATEVNT1.Item("FORM_NAME") = "SOFSHIP0"
                                dst.Tables("TATEVNT1").Rows.Add(rowTATEVNT1)
                            End If

                        Catch ex As Exception
                            MyBase.Rollback(ex.Message)
                            CreditCardProcessed = False

                            Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
                            rowTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
                            rowTATEVNT1.Item("TABLE_KEY") = rowSOTPICK1.Item("ORDR_NO")
                            rowTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
                            rowTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                            rowTATEVNT1.Item("EVENT_TYPE") = "CCP"
                            rowTATEVNT1.Item("EVENT_DESC") = "Credit Card Error: " & ex.Message
                            rowTATEVNT1.Item("EVENT_KEY") = ""
                            rowTATEVNT1.Item("FORM_NAME") = "SOFSHIP0"
                            dst.Tables("TATEVNT1").Rows.Add(rowTATEVNT1)

                        End Try
                    End If
                Next
            Catch ex As Exception
                MessageBox.Show("The following error occurred when processing the credit card: " & ex.Message, "Credit Card Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                CreditCardProcessed = False

                Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
                rowTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
                rowTATEVNT1.Item("TABLE_KEY") = rowSOTPICK1.Item("ORDR_NO")
                rowTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
                rowTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowTATEVNT1.Item("EVENT_TYPE") = "CCP"
                rowTATEVNT1.Item("EVENT_DESC") = "Credit Card Error: " & ex.Message
                rowTATEVNT1.Item("EVENT_KEY") = ""
                rowTATEVNT1.Item("FORM_NAME") = "SOFSHIP0"
                dst.Tables("TATEVNT1").Rows.Add(rowTATEVNT1)
            End Try

            If CreditCardProcessed = False Then
                MessageBox.Show("Update aborted! Credit card could not be processed.", "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)

                Dim EVENT_DESC_LEN As Int16 = dst.Tables("TATEVNT1").Columns("EVENT_DESC").MaxLength
                For Each rowTATEVNT1 As DataRow In dst.Tables("TATEVNT1").Select()
                    If (rowTATEVNT1.Item("EVENT_DESC") & String.Empty).ToString.Length > EVENT_DESC_LEN Then
                        rowTATEVNT1.Item("EVENT_DESC") = (rowTATEVNT1.Item("EVENT_DESC") & String.Empty).ToString.Substring(0, EVENT_DESC_LEN)
                    End If
                Next

                Try
                    BeginTrans()
                    Update_Record_TDA("TATEVNT1")
                    CommitTrans()
                Catch ex As Exception
                    Rollback()
                End Try
                Exit Sub
            End If

            ASCMAIN1.Progress("Now Updating ...", "")

            Dim RFIXMSG As Boolean = False
            Dim SOCINVH1 As New TAC.SOCINVH1(dst)
            ' Update the Sales Order records with the Pick Ticket data
            SOCINVH1.ProcessPickTicketsAndUpdateSalesDetails(MyBase.Absx1.dteFor("INV_DATE").Value)

            ' Record event where the Ship via was changed
            For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select("SELECTED = '1'", "", DataViewRowState.CurrentRows)
                Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)

                ' Place Misc Charge on random Pick Ticket
                ' Charge INV_MISC_CHG only once
                If INV_MISC_CHG_PICK_NO.Length = 0 AndAlso INV_MISC_CHG <> 0 Then
                    rowSOTPICK1.Item("INV_MISC_CHG") = INV_MISC_CHG
                    INV_MISC_CHG_PICK_NO = rowSOTPICK1.Item("PICK_NO")
                End If

                If rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty <> MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text Then
                    Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").Rows.Add
                    rowTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
                    rowTATEVNT1.Item("TABLE_KEY") = ORDR_NO
                    rowTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
                    rowTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    rowTATEVNT1.Item("EVENT_TYPE") = "SHPMTC"
                    rowTATEVNT1.Item("EVENT_DESC") = "Ship Via was changed from " _
                        & rowSOTORDR1.Item("SHIP_VIA_CODE") & " to " & MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text
                    rowTATEVNT1.Item("EVENT_KEY") = ""
                    rowTATEVNT1.Item("FORM_NAME") = "SOFSHIP0"
                End If
            Next

            ' Create Invoice Records
            Dim CUST_FACTOR_TRANS_IND As String = String.Empty
            If chkFactored.Checked Then
                CUST_FACTOR_TRANS_IND = "1"
            Else
                CUST_FACTOR_TRANS_IND = "0"
            End If

            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

                ' Log factoring change
                If Val(rowSOTSHIP1.Item("CUST_FACTOR_TRANS_IND") & String.Empty) <> Val(CUST_FACTOR_TRANS_IND) Then
                    Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").Rows.Add
                    rowTATEVNT1.Item("TABLE_NAME") = "SOTSHIP1"
                    rowTATEVNT1.Item("TABLE_KEY") = SHIP_BOL_NO
                    rowTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
                    rowTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    rowTATEVNT1.Item("EVENT_TYPE") = "SHPFAC"
                    rowTATEVNT1.Item("EVENT_DESC") = "Factor Setting was changed from " _
                        & IIf(Val(rowSOTSHIP1.Item("CUST_FACTOR_TRANS_IND") & String.Empty) = 1, "True", "False") & " to " & IIf(Val(CUST_FACTOR_TRANS_IND) = 1, "True", "False")
                    rowTATEVNT1.Item("EVENT_KEY") = ""
                    rowTATEVNT1.Item("FORM_NAME") = "SOFSHIP0"
                End If

                rowSOTSHIP1.Item("CUST_FACTOR_TRANS_IND") = CUST_FACTOR_TRANS_IND
                rowSOTSHIP1.Item("INV_DATE") = CDate(MyBase.Absx1.dteFor("INV_DATE").Value & String.Empty).ToShortDateString
                rowSOTSHIP1.Item("SHIP_DATE_SHIPPED") = CDate(MyBase.Absx1.dteFor("SHIP_DATE_SHIPPED").Value & String.Empty).ToShortDateString
                If Not IsDate(rowSOTSHIP1.Item("SHIPPED_ACTUAL") & String.Empty) Then
                    rowSOTSHIP1.Item("SHIPPED_ACTUAL") = rowSOTSHIP1.Item("SHIP_DATE_SHIPPED")
                End If
                SOCINVH1.CreateInvoices(SHIP_BOL_NO, RFIXMSG)
            Next

            ' Record Invoice information in SOTORDC2 records
            For Each row As DataRow In dst.Tables("SOTORDC2").Select("ISNULL(PICK_NO, '*') <> '*'")
                Dim PICK_NO As String = row.Item("PICK_NO") & String.Empty
                PICK_NO = PICK_NO.TRIM
                If PICK_NO.Length = 0 Then
                    Continue For
                End If
                Dim rowSOTPICK1X As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
                If rowSOTPICK1X IsNot Nothing Then
                    row.Item("INV_NO") = rowSOTPICK1X.Item("INV_NO")
                    Dim INV_NO As String = row.Item("INV_NO") & String.Empty
                    If INV_NO.Length = 0 Then
                        Continue For
                    End If
                    Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").Rows.Find(New Object() {"I", INV_NO})
                    If rowSOTINVH1 Is Nothing Then
                        Continue For
                    End If
                    row.Item("INV_NO") = rowSOTINVH1.Item("INV_NO")
                    row.Item("INV_DATE") = rowSOTINVH1.Item("INV_DATE")
                    row.Item("INV_AMOUNT") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
                End If
            Next

            If INV_MISC_CHG_PICK_NO.Length > 0 Then
                Dim rowSOTPICK1_CHG As DataRow = dst.Tables("SOTPICK1").Rows.Find(INV_MISC_CHG_PICK_NO)
                If rowSOTPICK1_CHG Is Nothing OrElse rowSOTPICK1_CHG.Item("INV_NO") & String.Empty = String.Empty Then
                    Throw New Exception("Unable to locate Pick Ticket Invoice that received Misc Charges.")
                End If

                Dim INV_NO As String = rowSOTPICK1_CHG.Item("INV_NO")
                For Each rowSOTINVHM As DataRow In dst.Tables("SOTINVHM").Select()
                    rowSOTINVHM.Item("INV_NO") = INV_NO
                Next
            End If

            ' Cancel all Pick tickets where nothing was picked. - What happens with In Pick Qtys
            For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select()
                rowSOTPICK1.Item("PICK_STATUS") = "F"
                Dim pickConf As Int32 = Val(dst.Tables("SOTPICK2").Compute("SUM(PICK_QTY_CONF)", "PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'") & String.Empty)
                If pickConf = 0 Then
                    rowSOTPICK1.Item("PICK_STATUS") = "C"
                    rowSOTPICK1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                    rowSOTPICK1.Item("LAST_DATE") = DATETIME_STAMP
                Else
                    If dst.Tables("SOTINVH1").Select("PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'").Length = 0 Then
                        Throw New Exception("Not all Pick tickets created an Invoice")
                    End If
                End If
            Next

            inTransaction = True
            BeginTrans()

            Update_SOTORDR5()

            RecordPriceChanges()
            Update_Record_TDA("TATEVNT1")

            Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
            Dim WHSE_PHYS_STATUS As String = rowICTWHSE1.Item("WHSE_PHYS_STATUS") & ""
            Dim WHSE_LOCATOR As Boolean = rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1"

            ' Calculate and Update Total Cartons & Weight by BOL
            ' Create New BOL's for Pick Tickets excluded from this Shipment

            ' Fetch FIFO costs for all Styles on this Group

            If edi_order And edi856_customer Then
                For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select()
                    rowSOTCART1.Item("CART_TOTAL_UNITS") = rowSOTCART1.Item("CART_TOTAL_UNITS_CALC")
                Next
                For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select("SELECTED = '1'")
                    'rowSOTPICK1.Item("PICK_FREIGHT") = rowSOTPICK1.Item("PICK_FREIGHT_CALC")
                    If ASCMAIN1.CLIENT <> "NYA" Then
                        rowSOTPICK1.Item("PICK_TOTAL_WGT") = rowSOTPICK1.Item("PICK_TOTAL_WGT_CALC")
                    End If
                    rowSOTPICK1.Item("PICK_CNT_CARTONS") = rowSOTPICK1.Item("PICK_CNT_CARTONS_CALC")
                Next
            End If

            ASCMAIN1.Progress("Now Updating ...", "")

            Dim old_new_bols As String = ""
            Dim SHIP_BOL_NO_new As String

            rowSOTSHIP0.Item("INV_DATE") = MyBase.Absx1.dteFor("INV_DATE").Value
            rowSOTSHIP0.Item("SHIP_VIA_CODE") = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text
            rowSOTSHIP0.Item("FRT_TERMS") = MyBase.Absx1.txtFor("FRT_TERMS").Text
            rowSOTSHIP0.Item("BILL_OF_LADING_NO") = MyBase.Absx1.txtFor("BILL_OF_LADING_NO").Text
            rowSOTSHIP0.Item("ORDR_DEPT") = MyBase.Absx1.txtFor("ORDR_DEPT").Text
            rowSOTSHIP0.Item("EDI_LOAD_ID") = MyBase.Absx1.txtFor("EDI_LOAD_ID").Text
            rowSOTSHIP0.Item("BTB_BOL_NO") = MyBase.Absx1.txtFor("BTB_BOL_NO").Text

            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

                rowSOTSHIP1.Item("SHIP_VIA_CODE") = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text
                rowSOTSHIP1.Item("FRT_TERMS") = MyBase.Absx1.txtFor("FRT_TERMS").Text
                rowSOTSHIP1.Item("BILL_OF_LADING_NO") = MyBase.Absx1.txtFor("BILL_OF_LADING_NO").Text
                rowSOTSHIP1.Item("ORDR_DEPT") = MyBase.Absx1.txtFor("ORDR_DEPT").Text
                rowSOTSHIP1.Item("EDI_LOAD_ID") = MyBase.Absx1.txtFor("EDI_LOAD_ID").Text
                rowSOTSHIP1.Item("BTB_BOL_NO") = MyBase.Absx1.txtFor("BTB_BOL_NO").Text
                rowSOTSHIP1.Item("INV_DATE") = MyBase.Absx1.dteFor("INV_DATE").Value
                rowSOTSHIP1.Item("SHIP_REF") = MyBase.Absx1.txtFor("SHIP_REF").Text
                rowSOTSHIP0.Item("SHIP_REF") = MyBase.Absx1.txtFor("SHIP_REF").Text

                If MyBase.Absx1.txtFor("SHIP_REF").Text.Trim.ToString.Length = 0 Then
                    If dst.Tables("WHTSHPC1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'").Length > 0 Then
                        ship_ref = dst.Tables("WHTSHPC1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")(0).Item("MASTER_TRACKING_NO") & String.Empty
                        rowSOTSHIP1.Item("SHIP_REF") = ship_ref
                        rowSOTSHIP0.Item("SHIP_REF") = ship_ref
                    End If
                Else
                    ship_ref = MyBase.Absx1.txtFor("SHIP_REF").Text.Trim
                End If

                ' Update carton Ship Ref with provided ship ref
                If ASCMAIN1.CLIENT = "RGI" AndAlso rowSOTSHIP1.Item("SHIP_REF") & String.Empty <> String.Empty Then
                    For Each rowSOTPICK1x As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
                        Dim PICK_NOx As String = rowSOTPICK1x.Item("PICK_NO") & String.Empty
                        For Each rowSOTCART1x As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NOx & "'")
                            rowSOTCART1x.Item("CART_TRACKING_NO") = rowSOTSHIP1.Item("SHIP_REF") & String.Empty
                        Next
                        'CART_TRACKING_NO
                    Next
                End If


                Dim INV_DATE As Date = rowSOTSHIP1.Item("INV_DATE")
                Dim ORDR_YYYYPP_UPDATED As String = ASCDATA1.GetDataValue("Select MIN(OPS_YYYYPP) from gltparm2 where prd_end_date >= '" & INV_DATE.ToString("dd-MMM-yyyy") & "'") & String.Empty
                If ORDR_YYYYPP_UPDATED.Length = 0 Then
                    ORDR_YYYYPP_UPDATED = ASCMAIN1.CYP
                End If

                Dim sqlw As String = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                Dim sqlw_selected As String = sqlw & " and SELECTED = '1'"
                Dim T As DataTable = dst.Tables("SOTPICK1") ' .Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'").CopyToDataTable
                Dim SHIP_CNT_CARTONS As Int64 = Val(T.Compute("SUM(PICK_CNT_CARTONS)", sqlw_selected) & "")
                Dim SHIP_TOTAL_WGT As Decimal = Val(T.Compute("SUM(PICK_TOTAL_WGT)", sqlw_selected) & "")
                Dim SHIP_TOTAL_FRT As Decimal = Val(T.Compute("SUM(PICK_FREIGHT)", sqlw_selected) & "")
                Dim PICKS_SEL As Int64 = Val(T.Compute("Count(PICK_NO)", sqlw_selected) & "")
                Dim PICKS As Int64 = Val(T.Compute("Count(PICK_NO)", sqlw) & "")

                With rowSOTSHIP1
                    If PICKS_SEL > 0 Then
                        .Item("OPS_YYYYPP") = ORDR_YYYYPP_UPDATED 'ASCMAIN1.CYP
                        .Item("SHIP_CNT_CARTONS") = SHIP_CNT_CARTONS
                        .Item("SHIP_TOTAL_WGT") = SHIP_TOTAL_WGT
                        .Item("SHIP_STATUS") = "F"
                        For Each COLUMN_NAME As String In New String() _
                            {"SHIP_VIA_CODE", "SHIP_DATE_SHIPPED", "INV_DATE", "REASON_CODE", "TERM_CODE", _
                             "SREP_CODE", "SREP2_CODE", "ORDR_DEPT", "EDI_LOAD_ID", "BTB_BOL_NO", "SHIP_REF", "SHIP_MANIFEST_NO", "BILL_OF_LADING_NO", "FRT_TERMS"}
                            .Item(COLUMN_NAME) = rowSOTSHIP0.Item(COLUMN_NAME)

                            If ",SHIP_DATE_SHIPPED,INV_DATE,".Contains("," & COLUMN_NAME & ",") Then
                                If IsDate(.Item(COLUMN_NAME) & String.Empty) Then
                                    .Item(COLUMN_NAME) = CDate(.Item(COLUMN_NAME).ToString).ToShortDateString
                                End If
                            End If
                        Next

                        If PICKS_SEL <> PICKS Then
                            Dim rowSOTSHIP1_P As DataRow = dst.Tables("SOTSHIP1").NewRow
                            With rowSOTSHIP1_P
                                For i As Integer = 0 To dst.Tables("SOTSHIP1").Columns.Count - 1
                                    .Item(i) = rowSOTSHIP1.Item(i)
                                Next i
                                If ASCMAIN1.CLIENT = "VAN" Then
                                    SHIP_BOL_NO_new = ASCMAIN1.Next_Control_No("SHIP_BOL_NO")
                                Else
                                    SHIP_BOL_NO_new = ASCMAIN1.Next_Control_No("SOTSHIP1.SHIP_BOL_NO")
                                End If
                                old_new_bols = old_new_bols & vbCr & SHIP_BOL_NO & " -> " & SHIP_BOL_NO_new
                                .Item("SHIP_BOL_NO") = SHIP_BOL_NO_new
                                .Item("SHIP_CNT_CARTONS") = 0
                                .Item("SHIP_TOTAL_WGT") = 0
                                .Item("SHIP_STATUS") = "P"
                                .Item("OPS_YYYYPP") = ""
                                For Each COLUMN_NAME As String In New String() _
                                    {"SHIP_VIA_CODE", "SHIP_DATE_SHIPPED", "INV_DATE", "REASON_CODE", "TERM_CODE", _
                                     "SREP_CODE", "SREP2_CODE", "ORDR_DEPT", "EDI_LOAD_ID", "BTB_BOL_NO", "SHIP_REF", "SHIP_MANIFEST_NO", "BILL_OF_LADING_NO", "FRT_TERMS"}
                                    .Item(COLUMN_NAME) = rowSOTSHIP0_ORIG.Item(COLUMN_NAME)
                                Next
                            End With
                            dst.Tables("SOTSHIP1").Rows.Add(rowSOTSHIP1_P)
                            sqlw = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and SELECTED <> '1'"
                            For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select(sqlw)
                                rowSOTPICK1.Item("SHIP_BOL_NO") = SHIP_BOL_NO_new
                            Next
                        End If
                    Else
                        .Item("SHIP_CNT_CARTONS") = 0
                        .Item("SHIP_TOTAL_WGT") = 0
                        .Item("SHIP_VIA_CODE") = ""
                        .Item("SHIP_DATE_SHIPPED") = DBNull.Value
                        .Item("FRT_TERMS") = ""
                        .Item("SHIP_REF") = ""
                        .Item("SHIP_MANIFEST_NO") = ""
                        .Item("BILL_OF_LADING_NO") = ""
                    End If
                    .Item("LAST_DATE") = DATETIME_STAMP
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                End With
            Next

            Update_Record_TDA("SOTORDR1")
            Update_Record_TDA("SOTORDR2")
            Update_Record_TDA("SOTORDR5")
            Update_Record_TDA("TATEVNT1")
            Update_Record_TDA("SOTORDXR")
            Update_Record_TDA("SOTORDR9")
            Update_Record_TDA("SOTRNGA1")
            Update_Record_TDA("SOTORDC1")
            Update_Record_TDA("SOTORDC2")

            If RFIXMSG = True Then
                MsgBox("Range Styles Were Fixed On this Shipment. Please Alert ABS", MsgBoxStyle.OkOnly, "Ranges")
            End If

            INIT_LAST("SOTSHIP1", False, , True)
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
                ASCMAIN1.Progress("Updating Shipment " & SHIP_BOL_NO, "")
                Delete_Records(SHIP_BOL_NO, False)
            Next

            ' Copy Work Table Contents to Oracle
            For Each TABLE_NAME As String In New String() {"SOTSHIP1", "SOTPICK1", "SOTPICK2", "SOTCART1", "SOTCART2", "SOTCART3"}
                dst.Tables(TABLE_NAME).AcceptChanges()
                For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")
                    row.SetAdded()
                Next
                Update_Record_TDA(TABLE_NAME)
            Next

            If select_from_3PL_list AndAlso dst.Tables.Contains("EDT945T1") Then
                Update_Record_TDA("EDT945T1")
            End If

            For Each TABLE_NAME As String In New String() {"SOTINVH1", "SOTINVH2", "SOTINVH9", "SOTINVHM", "ARTOPEN1"}
                Update_Record_TDA(TABLE_NAME)
            Next

            For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Rows

                ASCMAIN1.sql = "BEGIN SOPSTAT1('" & rowSOTINVH1.Item("INV_TYPE") & "','" & rowSOTINVH1.Item("INV_NO") & "'); END;"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "BEGIN SOPSTAT2('" & rowSOTINVH1.Item("INV_TYPE") & "','" & rowSOTINVH1.Item("INV_NO") & "'); END;"
                ASCDATA1.ExecuteSQL()

                ' When should this be called - Only whses that use Locationss
                If WHSE_LOCATOR Then
                    TAC.ICCMAIN1.Update_WHTLOCBX("S", rowSOTINVH1.Item("INV_NO"))
                End If

                ASCDATA1.ExecuteSP("ARPCUST6_IC", "VV", _
                   New Object() {rowSOTINVH1.Item("INV_TYPE"), rowSOTINVH1.Item("INV_NO")}, _
                   New String() {"INV_TYPE_IN", "INV_NO_IN"})

            Next

            Dim order_header_updates_required As Boolean = False
            Dim SQLX As String = ""
            For Each COL As String In New String() {"SREP_CODE", "SREP2_CODE", "TERM_CODE", "ORDR_DEPT"}
                If rowSOTSHIP0_ORIG.Item(COL) & "" <> rowSOTSHIP0.Item(COL) & "" Then
                    order_header_updates_required = True
                    Exit For
                End If
            Next

            ' Process each BOL, now that it is in Oracle
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

                If ASCMAIN1.CLIENT = "VAN" Then
                    Dim rowWHT3PLS1 As DataRow = dst.Tables("WHT3PLS1").Rows.Find(SHIP_BOL_NO)
                    If rowWHT3PLS1 IsNot Nothing Then
                        rowWHT3PLS1.Delete()
                    End If
                End If

                'Dependent_Updates(1, SHIP_BOL_NO)
                ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & rowSOTSHIP1.Item("ORDR_GROUP_NO") & "'); END;"
                ASCDATA1.ExecuteSQL()

                If order_header_updates_required Then
                    ASCMAIN1.sql = "Update SOTORDR1 " _
                         & "Set TERM_CODE = :PARM1, ORDR_DEPT = :PARM2" & vbCrLf _
                         & " where ORDR_NO in " & vbCrLf _
                         & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
                         & " where SHIP_BOL_NO = :PARM3)" & vbCrLf
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", _
                                        New Object() {rowSOTSHIP1.Item("TERM_CODE"), _
                                                      rowSOTSHIP1.Item("ORDR_DEPT"), _
                                                      rowSOTSHIP1.Item("SHIP_BOL_NO")})
                End If

                'Kill any possible Store Order Configuration Records from the Shipment.
                'They can and should be re-built by running the report if need be.
                ASCMAIN1.sql = "Delete from SOTCONF2" & vbCrLf _
                    & " where ORDR_NO in " & vbCrLf _
                    & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
                    & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
                ASCDATA1.ExecuteSQL()

                ' 04/19/2018 - removed - calclations are done in the process that creates the invoices
                'If CURR_CODE = "USD" Then
                '    ASCMAIN1.sql = "Update SOTINVH1 Set" & vbCrLf _
                '        & "  CURR_CODE = '" & CURR_CODE & "'" & vbCrLf _
                '        & ", CURR_EXCH_RATE = " & CStr(CURR_EXCH_RATE) & vbCrLf _
                '        & ", INV_SALES_CURR = INV_SALES" & vbCrLf _
                '        & ", INV_FREIGHT_CURR = INV_FREIGHT" & vbCrLf _
                '        & ", INV_MISC_CHG_CURR = INV_MISC_CHG" & vbCrLf _
                '        & ", INV_TOTAL_AMT_CURR = INV_TOTAL_AMOUNT" & vbCrLf _
                '        & ", GST_TAX = 0" & vbCrLf _
                '        & ", GST_TAX_CURR = 0" & vbCrLf _
                '        & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                '    ASCDATA1.ExecuteSQL()

                '    ASCMAIN1.sql = "Update SOTINVH9 Set" & vbCrLf _
                '        & "  RANGE_STYLE_PRICE_CURR = RANGE_STYLE_PRICE" & vbCrLf _
                '        & ", RANGE_STYLE_PP_PRICE_CURR = RANGE_STYLE_PP_PRICE" & vbCrLf _
                '        & " where INV_TYPE = 'I' AND INV_NO IN (" & vbCrLf _
                '        & "   Select INV_NO from SOTINVH1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
                '    ASCDATA1.ExecuteSQL()

                '    ASCMAIN1.sql = "Update SOTINVH2 SET ORDR_UNIT_PRICE_CURR = ORDR_UNIT_PRICE" & vbCrLf _
                '        & " WHERE INV_TYPE = 'I' AND INV_NO IN (" & vbCrLf _
                '        & "   SELECT INV_NO FROM SOTINVH1 WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
                '    ASCDATA1.ExecuteSQL()
                'Else
                '    ASCMAIN1.sql = "" _
                '        & "Begin" & vbCrLf _
                '        & " Declare Cursor C1 is" & vbCrLf _
                '        & "  Select SOTPICK1.INV_NO, SOTPICK2.PICK_LNO INV_LNO, SOTORDR2.ORDR_UNIT_PRICE_CURR" & vbCrLf _
                '        & "    from SOTORDR2, SOTPICK1, SOTPICK2" & vbCrLf _
                '        & "    where SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" & vbCrLf _
                '        & "      and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                '        & "      and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                '        & "      and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                '        & "      and SOTORDR2.ORDR_NO = SOTPICK1.ORDR_NO;" & vbCrLf _
                '        & " Begin" & vbCrLf _
                '        & "  For R1 in C1 Loop" & vbCrLf _
                '        & "    Update SOTINVH2" & vbCrLf _
                '        & "     Set ORDR_UNIT_PRICE_CURR = R1.ORDR_UNIT_PRICE_CURR," & vbCrLf _
                '        & "         ORDR_UNIT_PRICE = R1.ORDR_UNIT_PRICE_CURR * " & CStr(CURR_EXCH_RATE) _
                '        & "    where INV_TYPE = 'I'" & vbCrLf _
                '        & "      and INV_NO = R1.INV_NO" & vbCrLf _
                '        & "      and INV_LNO = R1.INV_LNO;" & vbCrLf _
                '        & "  End Loop;" & vbCrLf _
                '        & " End;" & vbCrLf _
                '        & "End;"
                '    ASCDATA1.ExecuteSQL()

                '    ASCMAIN1.sql = "" _
                '        & "Begin" & vbCrLf _
                '        & " Declare Cursor C1 is" _
                '        & "  Select SOTINVH2.INV_NO" & vbCrLf _
                '        & "  , Sum(ORDR_UNIT_PRICE * ORDR_QTY_SHIP) INV_SALES" & vbCrLf _
                '        & "  , Sum(ORDR_UNIT_PRICE_CURR * ORDR_QTY_SHIP) INV_SALES_CURR" & vbCrLf _
                '        & "  from SOTINVH2" & vbCrLf _
                '        & "  where SOTINVH2.INV_NO in (Select DISTINCT(SOTPICK1.INV_NO)" & vbCrLf _
                '        & "                           from SOTPICK1" & vbCrLf _
                '        & "                           where SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" & vbCrLf _
                '        & "                             and SOTPICK1.PICK_STATUS = 'F')" & vbCrLf _
                '        & "    and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
                '        & "  group by SOTINVH2.INV_NO;" & vbCrLf _
                '        & " Begin" & vbCrLf _
                '        & "  For R1 in C1 Loop" & vbCrLf _
                '        & "   Update SOTINVH1 Set" & vbCrLf _
                '        & "    INV_SALES = R1.INV_SALES," & vbCrLf _
                '        & "    INV_SALES_CURR = R1.INV_SALES_CURR," & vbCrLf _
                '        & "    INV_FREIGHT = 0," & vbCrLf _
                '        & "    INV_FREIGHT_CURR = 0," & vbCrLf _
                '        & "    INV_MISC_CHG = 0," & vbCrLf _
                '        & "    INV_MISC_CHG_CURR = 0" & vbCrLf _
                '        & "   where SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
                '        & "     and SOTINVH1.INV_NO = R1.INV_NO;" & vbCrLf _
                '        & "  End Loop;" & vbCrLf _
                '        & " End;" & vbCrLf _
                '        & "End;"
                '    ASCDATA1.ExecuteSQL()

                '    ASCMAIN1.sql = "Update SOTINVH1" _
                '        & " SET GST_TAX = INV_SALES * 0.070," _
                '        & " GST_TAX_CURR = INV_SALES_CURR * 0.070," _
                '        & " INV_TOTAL_AMOUNT = INV_SALES + (INV_SALES * 0.070)," _
                '        & " INV_TOTAL_AMT_CURR = INV_SALES_CURR + (INV_SALES_CURR * 0.070)," _
                '        & " CURR_CODE = '" & CURR_CODE & "'," _
                '        & " CURR_EXCH_RATE = " & CStr(CURR_EXCH_RATE) _
                '        & " where SOTINVH1.INV_NO IN (SELECT DISTINCT(SOTPICK1.INV_NO)" _
                '        & "                       from SOTPICK1" _
                '        & "                       where SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" _
                '        & "                         and SOTPICK1.PICK_STATUS = 'F')" _
                '        & "   and SOTINVH1.INV_TYPE = 'I'"
                '    ASCDATA1.ExecuteSQL()
                'End If
            Next

            'Update Invoices to remove consolodated records when there is only one invoice.
            'Recv'd from Dave and implimented on 2/16 - W.R.
            ' Only for VAN 9/3/2013 - edz
            If ASCMAIN1.CLIENT = "VAN" Then
                ASCMAIN1.sql = "Update SOTINVH1 SET INV_NO_CONS = NULL where INV_NO_CONS IN" _
                    & " (SELECT INV_NO_CONS FROM" _
                    & " (SELECT INV_NO_CONS, COUNT(*) FROM SOTINVH1 WHERE INV_NO_CONS IS NOT NULL" _
                    & " group by INV_NO_CONS" _
                    & " having COUNT(*) = 1))"
                ASCDATA1.ExecuteSQL()
            End If

            ' Group Record
            ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
            ASCDATA1.ExecuteSQL()

            'Update New Control File To Force SJ&U between Confirm and Deconfirm - WR - 20051024
            ASCMAIN1.sql = "Update SOTCTLU1" _
                & " SET CTL_UPDATE_REQ = 'C'" _
                & " WHERE UPPER(CTL_KEY) = 'Z'"
            ASCDATA1.ExecuteSQL()

            If old_new_bols <> "" Then
                MsgBox(old_new_bols, vbOKOnly, "Unshipped P/T's on the following BOL's have been assigned a New BOL No")
            End If

            For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("ORDR_TYPE_CODE = 'XFR'")
                'Update_Transfer(rowSOTINVH1.Item("INV_NO"))
                Update_Transfer(rowSOTINVH1)
                SOCMAIN1.Create_943_for_Transfer_Order_Receipt(Me, rowSOTINVH1)
            Next

            If ASCMAIN1.CLIENT = "NYA" Then
                CommitTrans()
            Else
                CommitTrans("Shipment Updated - Invoice Number: " & dst.Tables("SOTINVH1").Rows(0).Item("INV_NO"))
            End If

            For Each shippingLabel As String In shipLabels
                If shippingLabel.Trim.Length > 0 Then PrintLabel(shippingLabel)
            Next

            EmailInvoice()

            ' Create Web Invoices
            Try
                ASCMAIN1.Progress("Creating Web Invoice", "")
                For Each row As DataRow In dst.Tables("SOTINVH1").Select("")
                    TAC.SOCMAIN1.CreateWebInvoice(Me, row.Item("INV_TYPE"), row.Item("INV_NO"))
                Next
            Catch ex As Exception

            End Try

        Catch ex As Exception
            If inTransaction Then Rollback()
            MessageBox.Show("Update Error: " & ex.Message, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End Try

    End Sub

    Sub Update_Transfer(rowSOTINVH1 As DataRow) ' (INV_NO As String)

        Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")

        If ASCMAIN1.CLIENT = "VAN" Then
            If ASCMAIN1.DBS_COMPANY = "VAN" Then
                ASCDATA1.ExecuteSP("SOPSHIP1_XFRVAN", "V", New Object() {INV_NO}, New String() {"INV_NO_IN"})
            End If

        Else

            Dim XFR_NO As String = ASCDATA1.ExecuteSF _
                                   ("SOPSHIP1_XFR", New String() {"INV_NO_IN"}, New Object() {INV_NO})

            ' the location relief side of a warehouse transfer should be happening in invoice update

            Dim WHSE_CODE_TO As String = rowSOTINVH1.Item("WHSE_CODE_TO")
            Dim rowICTWHSE1_WHSE_CODE_TO As DataRow = LookUp("ICTWHSE1", WHSE_CODE_TO)
            If rowICTWHSE1_WHSE_CODE_TO.Item("WHSE_LOCATOR") & "" = "1" Then
                ASCMAIN1.sql = "Select ICTIXFR2.XFR_NO WHSE_TRAN_NO, ICTIXFR2.XFR_LNO WHSE_TRAN_LNO" _
                   & ", 'T' WHSE_TRAN_TYPE, ICTIXFR1.WHSE_CODE_TO WHSE_CODE" _
                   & ", ICTWHSE1.WHSE_LOC_REC LOCATION_CODE, ICTIXFR2.STYLE_CODE, ICTIXFR2.COLOR_CODE" _
                   & ", ICTIXFR2.XFR_QTY WHSE_TRAN_QTY" _
                   & " from ICTIXFR1,ICTIXFR2,ICTWHSE1" _
                   & " where ICTIXFR1.XFR_NO = ICTIXFR2.XFR_NO" _
                   & "   and ICTWHSE1.WHSE_CODE = ICTIXFR1.WHSE_CODE_TO" _
                   & "   and ICTIXFR2.XFR_NO = '" & XFR_NO & "'"
                WHCMAIN1.Update_WHTLOCBX(Me)
            End If

        End If
    End Sub

    Sub Update_Record_Maintenance()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIPX)
            'ASCDATA1.ExecuteSQL("Truncate Table " & ASW("SOTCART1_3PL"))
            'ASCDATA1.ExecuteSQL("Truncate Table " & ASW("SOTCART2_3PL"))
        End If

        BeginTrans()

        Dim SHIP_CHGREQ_NOs As New Dictionary(Of String, String)

        Dim date_changed As Boolean = False

        If Format(ORDR_SHIP_DATE, "yyyyMMdd") <> Format(dteORDR_SHIP_DATE.Value, "yyyyMMdd") _
        OrElse Format(ORDR_CANCEL_DATE, "yyyyMMdd") <> Format(dteORDR_CANCEL_DATE.Value, "yyyyMMdd") Then
            date_changed = True
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
                Dim SHIP_CHGREQ_NO As String = ASCMAIN1.Next_Control_No("SOTSHIP3.SHIP_CHGREQ_NO")
                SHIP_CHGREQ_NOs.Add(SHIP_BOL_NO, SHIP_CHGREQ_NO)

                ASCMAIN1.sql = "Update SOTORDR1 " _
                    & "Set ORDR_SHIP_DATE = :PARM1, ORDR_CANCEL_DATE = :PARM2" & vbCrLf _
                    & " where ORDR_NO in " _
                    & " (Select ORDR_NO from SOTPICK1 where SHIP_BOL_NO = :PARM3)" & vbCrLf
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DDV", _
                                    New Object() {dteORDR_SHIP_DATE.Value, _
                                                  dteORDR_CANCEL_DATE.Value, _
                                                  SHIP_BOL_NO})
            Next
        End If

        Dim qty_changed As Boolean = False
        Dim price_changed As Boolean = False
        Dim price_changed_to_Range As Boolean = False

        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select _
            ("", "", DataViewRowState.ModifiedCurrent)
            With rowSOTPICK2
                Dim QTY As Int64 = Val(.Item("PICK_QTY") & "") - Val(.Item("PICK_QTY_CONF") & "")
                If QTY <> 0 Then
                    qty_changed = True
                    .Item("PICK_QTY_CANC_REL") = Val(.Item("PICK_QTY_CANC_REL") & "") + QTY
                    .Item("PICK_QTY") = .Item("PICK_QTY_CONF")
                End If
                .Item("PICK_QTY_CONF") = DBNull.Value
                .Item("PICK_QTY_CANC") = DBNull.Value
                .Item("PICK_QTY_BACK") = DBNull.Value

                Dim SHIP_BOL_NO As String = .GetParentRow("SOTPICK1_SOTPICK2").Item("SHIP_BOL_NO")
                Dim rowSOTSHIP4 As DataRow = dst.Tables("SOTSHIP4").NewRow
                With rowSOTSHIP4
                    Dim SHIP_CHGREQ_NO As String
                    If SHIP_CHGREQ_NOs.ContainsKey(SHIP_BOL_NO) Then
                        SHIP_CHGREQ_NO = SHIP_CHGREQ_NOs(SHIP_BOL_NO)
                    Else
                        SHIP_CHGREQ_NO = ASCMAIN1.Next_Control_No("SOTSHIP3.SHIP_CHGREQ_NO")
                        SHIP_CHGREQ_NOs.Add(SHIP_BOL_NO, SHIP_CHGREQ_NO)
                    End If

                    .Item("SHIP_CHGREQ_NO") = SHIP_CHGREQ_NO
                    .Item("PICK_NO") = rowSOTPICK2.Item("PICK_NO")
                    .Item("PICK_LNO") = rowSOTPICK2.Item("PICK_LNO")
                    .Item("PICK_QTY_OLD") = rowSOTPICK2.Item("PICK_QTY", DataRowVersion.Original)
                    .Item("PICK_QTY_NEW") = rowSOTPICK2.Item("PICK_QTY")
                    .Item("PICK_UNIT_PRICE_OLD") = rowSOTPICK2.Item("PICK_UNIT_PRICE", DataRowVersion.Original)
                    .Item("PICK_UNIT_PRICE_NEW") = rowSOTPICK2.Item("PICK_UNIT_PRICE")

                    If Val(.Item("PICK_UNIT_PRICE_OLD") & "") <> Val(.Item("PICK_UNIT_PRICE_NEW") & "") Then
                        price_changed = True
                        Dim rowSOTORDR9 As DataRow = dst.Tables("SOTORDR9").Rows.Find(New Object() {rowSOTPICK2.Item("ORDR_NO"), rowSOTPICK2.Item("RANGE_STYLE_LNO")})
                        If rowSOTORDR9 IsNot Nothing Then
                            If Val(rowSOTORDR9.Item("RANGE_STYLE_PRICE") & "") <> Val(rowSOTPICK2.Item("ORDR_UNIT_PRICE") & "") Then
                                MsgBox("Problem with Range - Call ABS - SOTORDR9 " & rowSOTORDR9.Item("ORDR_NO") & rowSOTORDR9.Item("RANGE_STYLE_LNO"))
                                Stop
                            Else
                                price_changed_to_Range = True
                                rowSOTORDR9.Item("RANGE_STYLE_PRICE") = Val(.Item("PICK_UNIT_PRICE_NEW") & "")
                                rowSOTORDR9.Item("RANGE_STYLE_PRICE_CURR") = Val(.Item("PICK_UNIT_PRICE_NEW") & "")
                                rowSOTORDR9.Item("RANGE_STYLE_PP_PRICE") = Val(rowSOTORDR9.Item("RANGE_STYLE_QTY") & "") * Val(rowSOTORDR9.Item("RANGE_STYLE_PRICE") & "")
                                rowSOTORDR9.Item("RANGE_STYLE_PP_PRICE_CURR") = Val(rowSOTORDR9.Item("RANGE_STYLE_QTY") & "") * Val(rowSOTORDR9.Item("RANGE_STYLE_PRICE_CURR") & "")
                            End If
                        End If
                    End If

                    dst.Tables("SOTSHIP4").Rows.Add(rowSOTSHIP4)
                End With
            End With
        Next

        For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select _
            ("", "", DataViewRowState.ModifiedCurrent)
            With rowSOTCART2
                Dim QTY As Int64 = Val(.Item("QTY_PACKED") & "") - Val(.Item("QTY_PACKED_ORIG") & "")
                If QTY <> 0 Then
                    qty_changed = True
                    Dim SHIP_BOL_NO As String = dst.Tables("SOTPICK1").Rows.Find(.Item("PICK_NO")).Item("SHIP_BOL_NO")
                    Dim rowSOTSHIP6 As DataRow = dst.Tables("SOTSHIP6").NewRow
                    With rowSOTSHIP6
                        Dim SHIP_CHGREQ_NO As String
                        If SHIP_CHGREQ_NOs.ContainsKey(SHIP_BOL_NO) Then
                            SHIP_CHGREQ_NO = SHIP_CHGREQ_NOs(SHIP_BOL_NO)
                        Else
                            SHIP_CHGREQ_NO = ASCMAIN1.Next_Control_No("SOTSHIP3.SHIP_CHGREQ_NO")
                            SHIP_CHGREQ_NOs.Add(SHIP_BOL_NO, SHIP_CHGREQ_NO)
                        End If

                        .Item("SHIP_CHGREQ_NO") = SHIP_CHGREQ_NO
                        .Item("CART_NO") = rowSOTCART2.Item("CART_NO")
                        .Item("CART_LNO") = rowSOTCART2.Item("CART_LNO")
                        .Item("QTY_PACKED_OLD") = rowSOTCART2.Item("QTY_PACKED", DataRowVersion.Original)
                        .Item("QTY_PACKED_NEW") = rowSOTCART2.Item("QTY_PACKED")
                        dst.Tables("SOTSHIP6").Rows.Add(rowSOTSHIP6)
                    End With
                End If
            End With
        Next



        Update_SOTORDR5()

        'Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        'Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)

        If qty_changed Then
            ' Retract the Qty In Pick for each BOL - Before
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
                Dependent_Updates(-1, SHIP_BOL_NO)
            Next
        End If

        If price_changed Then
            ' CONSIDER CHANGING THE PRICE OF SOTORDR2, SINCE WE ARE CHANGE SOTORDR9
            If price_changed_to_Range Then
                Update_Record_TDA("SOTORDR9")
            End If
        End If


        ' Send changes - only modified rows will be updated, price may be updated even though Qty was not changed
        For Each TABLE_NAME As String In New String() {"SOTPICK2", "SOTCART2"}
            Update_Record_TDA(TABLE_NAME)
        Next

        ' Restore the Qty In Pick for each BOL - After
        If qty_changed Then
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
                Dependent_Updates(-1, SHIP_BOL_NO)
            Next
        End If

        Dim LP_XNO As String = ""

        Dim sqlSHIP_CHG_REQ_NOs As String = ""
        For Each SHIP_BOL_NO As String In SHIP_CHGREQ_NOs.Keys
            Dim SHIP_CHGREQ_NO As String = SHIP_CHGREQ_NOs(SHIP_BOL_NO)
            sqlSHIP_CHG_REQ_NOs &= ",'" & SHIP_CHGREQ_NO & "'"
            Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
            Dim rowSOTSHIP3 As DataRow = dst.Tables("SOTSHIP3").NewRow
            With rowSOTSHIP3
                .Item("SHIP_CHGREQ_NO") = SHIP_CHGREQ_NO
                .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("ORDR_SHIP_DATE_OLD") = ORDR_SHIP_DATE
                .Item("ORDR_CANCEL_DATE_OLD") = ORDR_CANCEL_DATE
                .Item("ORDR_SHIP_DATE_NEW") = dteORDR_SHIP_DATE.Value
                .Item("ORDR_CANCEL_DATE_NEW") = dteORDR_CANCEL_DATE.Value
                .Item("SHIP_CHGREQ_REASON") = txtReason.Text
                .Item("SHIP_CHGREQ_CONTACT") = txtContact.Text
                .Item("SHIP_CHGREQ_EMAIL") = txtemail.Text
                If rowSOTSHIP1.Item("LP_STATUS") & "" = "1" Then
                    .Item("LP_CODE") = rowSOTSHIP1.Item("LP_CODE")
                    .Item("LP_STATUS") = "0"
                    If LP_XNO = "" Then LP_XNO = TAC.WHCMAIN1.Get_LP_XNO(MENU_ITEM_OBJECT, SHIP_CHGREQ_NOs.Count)
                    .Item("LP_XNO") = LP_XNO
                    .Item("LP_STATUS_TS_3PL") = DBNull.Value
                    .Item("LP_STATUS_TS_ERP") = DATETIME_STAMP
                End If
            End With
            dst.Tables("SOTSHIP3").Rows.Add(rowSOTSHIP3)
        Next

        For Each TABLE_NAME As String In New String() _
            {"SOTSHIP3", "SOTSHIP4", "SOTSHIP6"}
            Update_Record_TDA(TABLE_NAME)
        Next

        If sqlSHIP_CHG_REQ_NOs <> "" Then

            ASCMAIN1.sql = "" _
                 & "Begin" & vbCrLf _
                 & " Declare Cursor C1 is " & vbCrLf _
                 & "  Select SOTPICK2.ORDR_NO, SOTPICK2.ORDR_LNO, SOTSHIP4.PICK_UNIT_PRICE_NEW" & vbCrLf _
                 & "  , SOTSHIP4.PICK_QTY_OLD - SOTSHIP4.PICK_QTY_NEW QTY" & vbCrLf _
                 & "   from SOTPICK2,SOTSHIP4" & vbCrLf _
                 & "   where SOTPICK2.PICK_NO = SOTSHIP4.PICK_NO AND SOTPICK2.PICK_LNO = SOTSHIP4.PICK_LNO" & vbCrLf _
                 & "     and SOTSHIP4.SHIP_CHGREQ_NO IN (" & Mid(sqlSHIP_CHG_REQ_NOs, 2) & ");" & vbCrLf _
                 & " Begin" & vbCrLf _
                 & "  For R1 IN C1 Loop" & vbCrLf _
                 & "   Update SOTORDR2" & vbCrLf _
                 & "    Set ORDR_UNIT_PRICE = R1.PICK_UNIT_PRICE_NEW, ORDR_UNIT_PRICE_CURR = R1.PICK_UNIT_PRICE_NEW" & vbCrLf _
                 & "    , ORDR_QTY_PICK = NVL(ORDR_QTY_PICK,0) - R1.QTY, ORDR_QTY_CANC = NVL(ORDR_QTY_CANC,0) + R1.QTY" & vbCrLf _
                 & "    where ORDR_NO = R1.ORDR_NO AND ORDR_LNO = R1.ORDR_LNO;" & vbCrLf _
                 & "  End Loop;" & vbCrLf _
                 & " End;" & vbCrLf _
                 & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        ' Group Record
        ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
        ASCDATA1.ExecuteSQL()

        If ASCMAIN1.CLIENT = "VAN" Then
            If date_changed OrElse qty_changed Then

                ASCDATA1.ExecuteSQL("Delete from " & SOTSHIPX)
                ASCDATA1.ExecuteSQL("Insert into " & SOTSHIPX _
                                    & " Select SOTSHIP1.SHIP_BOL_NO, '0' SEL, '0' EDI856, SOTSHIP1.SHIP_CART_REQD" _
                                    & ", SOTSHIP3.SHIP_CHGREQ_NO, SOTORDR0.CUST_CODE" _
                                    & " from SOTSHIP3,SOTSHIP1,SOTORDR0" _
                                    & " where SOTSHIP3.LP_STATUS = '0'" _
                                    & "   and SOTSHIP3.LP_XNO = '" & LP_XNO & "'" _
                                    & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" _
                                    & "   and SOTSHIP1.SHIP_BOL_NO = SOTSHIP3.SHIP_BOL_NO")
                ASCDATA1.ExecuteSQL("Update " & SOTSHIPX & " SOTSHIPX " _
                    & "Set EDI856 = '1' where CUST_CODE in (Select Distinct CUST_CODE from EDTTRPM1 " _
                    & " where EDI_DOC_NO = '856' and EDI_STATUS = 'P')")

                TAC.WHCMAIN1.Prepare_Carton_Data_3PL(SOTSHIPX, "NJ", ASW)

                ' Stop ' NEED TO GET SOTCART2_3PL POPULATED
                ASCDATA1.ExecuteSQL("Update SOTSHIP3 set LP_STATUS = '1' where LP_STATUS = '0' and LP_XNO = '" & LP_XNO & "'")
            End If
        End If

        CommitTrans("Update Complete")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_SOTORDR5()

        If Me.BindingContext.Contains(dvwSOTORDR5) Then
            ' Without the next 2 lines, data in text boxes in single row datatables (like header tables) will not get written to Oracle
            Dim X As CurrencyManager = Me.BindingContext(dvwSOTORDR5)
            X.EndCurrentEdit()
        End If

        For Each rowSOTORDR5 As DataRow In dst.Tables("SOTORDR5").Select("", "", DataViewRowState.ModifiedCurrent)
            Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
            rowTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
            rowTATEVNT1.Item("TABLE_KEY") = rowSOTORDR5.Item("ORDR_NO")
            rowTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
            rowTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowTATEVNT1.Item("EVENT_TYPE") = "SHPMTC"
            rowTATEVNT1.Item("EVENT_DESC") = "Ship-To Address was Changed"
            rowTATEVNT1.Item("EVENT_KEY") = ""
            rowTATEVNT1.Item("FORM_NAME") = "SOFSHIP0"
            dst.Tables("TATEVNT1").Rows.Add(rowTATEVNT1)
            ' Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", rowTATEVNT1.Item("ORDR_NO"))
            For Each DC As DataColumn In dst.Tables("SOTORDR5").Columns
                Dim COLUMN_NAME As String = DC.ColumnName

                If rowSOTORDR5.Item(COLUMN_NAME) & "" <> rowSOTORDR5.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
                    Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
                    rowSOTORDXR.Item("ORDR_NO") = rowSOTORDR5.Item("ORDR_NO")
                    'rowSOTORDXR.Item("REV_NO") = rowSOTORDR1.Item("ORDR_NO")
                    'rowSOTORDXR.Item("REV_LNO") = 0
                    rowSOTORDXR.Item("INIT_DATE") = DATETIME_STAMP
                    rowSOTORDXR.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    rowSOTORDXR.Item("COLUMN_NAME") = COLUMN_NAME
                    rowSOTORDXR.Item("OLD_VALUE") = rowSOTORDR5.Item(COLUMN_NAME, DataRowVersion.Original)
                    rowSOTORDXR.Item("NEW_VALUE") = rowSOTORDR5.Item(COLUMN_NAME)
                    dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
                End If
            Next
        Next
        Update_Record_TDA("TATEVNT1")
        Update_Record_TDA("SOTORDXR")
        Update_Record_TDA("SOTORDR5")
    End Sub

    Private Sub Create_WHT3PLS1()

        If ASCMAIN1.CLIENT = "VAN" Then
            'For Each TABLE_NAME As String In New String() _
            '    {"SOTSHIP1_3PL", "SOTPICK1_3PL", "SOTPICK2_3PL", "SOTCART1_3PL", "SOTCART2_3PL", "SOTCART3_3PL"}
            '    ASW.Add(TABLE_NAME, ASCMAIN1.Temp_Table("Select * from " & TABLE_NAME & " where ROWNUM <1"))
            'Next

            ASCMAIN1.sql = "Select SOTSHIP1.*" _
                & ", SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, SOTPICK1.PICK_NO" _
                & " from SOTSHIP1, SOTORDR0, SOTPICK1" _
                & " where SOTORDR0.ORDR_GROUP_NO (+) = SOTSHIP1.ORDR_GROUP_NO" _
                & " and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO"
            Create_TDA(dst.Tables.Add, "WHT3PLS1", "**", 0, False, , 1)
            dst.Tables("WHT3PLS1").Columns("CUST_CODE").AllowDBNull = True

        ElseIf ASCMAIN1.CLIENT = "NYA" Or ASCMAIN1.CLIENT = "RGI" Then
            ASCMAIN1.sql = " SELECT SOTSHIP1.SHIP_BOL_NO, EDT945T1.EDI_DOC_SEQ_NO" _
            & " ,EDT945T1.EDI_SHIPMENT_DATE SHIP_DATE_SHIPPED, EDT945T1.EDI_BOL_NO, ARTCUST1.CUST_NAME" _
            & " ,SOTSHIP1.SHIP_VIA_CODE" _
            & " ,SOTSHIP1.SHIP_ADDR_TYPE" _
            & " ,SOTSHIP1.SHIP_ADDR_CODE" _
            & " ,SOTSHIP1.ORDR_GROUP_NO" _
            & " ,SOTSHIP1.SHIP_PICK_PRINTED" _
            & " ,SOTSHIP1.PICK_BATCH_NO" _
            & " ,SOTSHIP1.FRT_TERMS" _
            & " ,SOTSHIP1.WHSE_CODE" _
            & " ,SOTSHIP1.INIT_DATE, SOTSHIP1.INIT_OPER" _
            & " ,SOTSHIP1.LAST_DATE, SOTSHIP1.LAST_OPER" _
            & " ,SOTSHIP1.SREP_CODE" _
            & " ,SOTSHIP1.ORDR_DEPT" _
            & " ,SOTSHIP1.SHIP_DATE_RECEIVED" _
            & " ,SOTSHIP1.SHIP_NOTES" _
            & " ,SOTSHIP1.SREP2_CODE" _
            & " ,SOTSHIP1.BOL_PRINTED" _
            & " ,SOTSHIP1.SHIP_DATE_PACKED" _
            & " ,SOTORDR0.CUST_CODE" _
            & " ,SOTORDR0.ORDR_SHIP_DATE" _
            & " ,SOTORDR0.ORDR_CANCEL_DATE" _
            & " ,SOTORDR0.ORDR_CUST_PO" _
            & " ,SOTPICK1.PICK_NO" _
            & " ,NVL(EDT945T1.EDI_PROCESS_IND, '0') EDI_PROCESS_IND" _
            & ", EDT945T1.EDI_MASTER_BOL_NO" _
            & " FROM EDT945T1, SOTPICK1, SOTSHIP1, SOTORDR0, ICTWHSE1, ARTCUST1" _
            & " WHERE EDT945T1.EDI_PICK_NO = SOTPICK1.PICK_NO" _
            & " AND SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" _
            & " AND SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" _
            & " AND NVL(EDT945T1.EDI_PROCESS_IND, '0') IN ('0', '3', '4') " _
            & " and SOTSHIP1.WHSE_CODE = ICTWHSE1.WHSE_CODE" _
            & " and ICTWHSE1.LP_CODE IS NOT NULL" _
            & " and SOTORDR0.CUST_CODE = ARTCUST1.CUST_CODE"

            Create_TDA(dst.Tables.Add, "WHT3PLS1", "**", 0, False, , 2)
            dst.Tables("WHT3PLS1").Columns("CUST_CODE").AllowDBNull = True

            Create_TDA(dst.Tables.Add, "EDT945T1", "*")
            Create_TDA(dst.Tables.Add, "EDT945T2", "*", 1)
        Else
            ' place holder for table
            ASCMAIN1.sql = " SELECT SOTSHIP1.SHIP_BOL_NO" _
               & " ,SYSDATE SHIP_DATE_SHIPPED" _
               & " ,SOTSHIP1.SHIP_BOL_NO EDI_BOL_NO" _
               & " ,SOTSHIP1.SHIP_VIA_CODE" _
               & " ,SOTSHIP1.SHIP_ADDR_TYPE" _
               & " ,SOTSHIP1.SHIP_ADDR_CODE" _
               & " ,SOTSHIP1.ORDR_GROUP_NO" _
               & " ,SOTSHIP1.SHIP_PICK_PRINTED" _
               & " ,SOTSHIP1.PICK_BATCH_NO" _
               & " ,SOTSHIP1.FRT_TERMS" _
               & " ,SOTSHIP1.WHSE_CODE" _
               & " ,SOTSHIP1.INIT_DATE, SOTSHIP1.INIT_OPER" _
               & " ,SOTSHIP1.LAST_DATE, SOTSHIP1.LAST_OPER" _
               & " ,SOTSHIP1.SREP_CODE" _
               & " ,SOTSHIP1.ORDR_DEPT" _
               & " ,SOTSHIP1.SHIP_DATE_RECEIVED" _
               & " ,SOTSHIP1.SHIP_NOTES" _
               & " ,SOTSHIP1.SREP2_CODE" _
               & " ,SOTSHIP1.BOL_PRINTED" _
               & " ,SOTSHIP1.SHIP_DATE_PACKED" _
               & " ,SOTORDR0.CUST_CODE" _
               & " ,SOTORDR0.ORDR_SHIP_DATE" _
               & " ,SOTORDR0.ORDR_CANCEL_DATE" _
               & " ,SOTORDR0.ORDR_CUST_PO" _
               & " FROM SOTSHIP1, SOTORDR0" _
               & " WHERE SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" _
               & " AND SOTSHIP1.SHIP_BOL_NO = '@@@@@@@@@@@@' "
            Create_TDA(dst.Tables.Add, "WHT3PLS1", "**", 0, False, , 1)
        End If

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "SHIP_BOL_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("ORDR_CUST_PO").Text = "" Then
                    MsgBox("You must enter a Customer Code or a PO No", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""

                If InquiryMode Then
                Else
                    sql_where &= " and SOTRSRV1.RSRV_STATUS = 'O' "
                End If

                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    sql_where &= " and SOTRSRV1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                End If
                If Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                    sql_where &= " and SOTRSRV1.ORDR_CUST_PO = '" & Absx1.txtFor("ORDR_CUST_PO").Text & "'"
                End If

            Case "CUST_ADDR_CODE"
                sql_where = "CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"

            Case "SHIP_VIA_CODE"
                'If ASCMAIN1.CLIENT = "NYA" Then
                '    sql_where = "CARRIER_MODE = 'U'"
                'End If
        End Select
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "Load", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("SHIP_BOL_NO").Text = key
                Click_Command("Load")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTSHIP1"
            E.COLUMN_NAME = "SHIP_BOL_NO"
            E.CODE_VALUE = Absx1.txtFor("SHIP_BOL_NO").Text
            E.DESC_VALUE = "Shipment"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTRSRV1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()

        Check_InquiryMode()

        If InquiryMode Then
            Load_Popup_Menu(grdSOTSHIPX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Print Manifest")
            If ASCMAIN1.USER_SECURITY_CODEs.Contains("WH") Then
                Load_Popup_Menu(grdSOTCART1, "B", "Modify Tracking Number")
            End If
        Else
            Load_Popup_Menu(grdSOTSHIPX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        End If

        Load_Popup_Menu(grdSOTPICK1, "BBBSB", "Select All", "De-Select All", "Propagate Value", "Hide Details", "Sales Order Inquiry")
        Load_Popup_Menu(grdSOTPICK2, "BS", "Style Status Inquiry", "Permit Price Change")

        If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
            Load_Popup_Menu(grdWHT3PLS1, "BBB", "Refresh", "Set Master BOL No", "Clear EDI 945 Record")
        Else
            Load_Popup_Menu(grdWHT3PLS1, "B", "Refresh")
        End If

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
            Case "grdSOTPICK2"
                tlb_pop.Tools("Permit Price Change").SharedProps.Visible = Not InquiryMode _
                    AndAlso (EntryMode = "E" OrElse EntryMode = "N") _
                    AndAlso dst.Tables("SOTORDR1").Select("ISNULL(EDI_DOC_SEQ_NO, '') <> ''").Length = 0 _
                    AndAlso dst.Tables("SOTORDR1").Select("CURR_CODE <> 'USD'").Length = 0
            Case "grdSOTPICK1"
                tlb_pop.Tools("Select All").SharedProps.Visible = Not InquiryMode And (EntryMode = "E" OrElse EntryMode = "N") And Not MaintenanceMode
                tlb_pop.Tools("De-Select All").SharedProps.Visible = Not InquiryMode And (EntryMode = "E" OrElse EntryMode = "N") And Not MaintenanceMode
            Case "grdSOTCART1"
                'tlb_pop.Tools("Modify Tracking Number").SharedProps.Visible = InquiryMode And dst.Tables("SOTCART1").Rows.Count > 0

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '  e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdSOTPICK1"
                    tlb_btn = DirectCast(tlb_pop.Tools("Propagate Value"), UltraWinToolbars.ButtonTool)
                    If Not (EntryMode = "E" OrElse EntryMode = "N") OrElse grd.ActiveCell Is Nothing OrElse Not New String() {"PICK_CNT_CARTONS", "PICK_TOTAL_WGT", "PICK_FREIGHT", "BILL_OF_LADING_NO", "ORDR_INV_COMMENT"}.Contains(grd.ActiveCell.Column.Key) Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        tlb_btn.SharedProps.Visible = True
                        tlb_btn.SharedProps.Caption = "Propagate Values for " & grd.ActiveCell.Column.Header.Caption
                        tlb_btn.Tag = grd.ActiveCell.Column.Key
                    End If
                    For Each ToolKey As String In New String() {"Hide Details"} ' {"Select All", "De-Select All", "Hide Details"}
                        DirectCast(tlb_pop.Tools(ToolKey), UltraWinToolbars.ButtonTool).SharedProps.Visible = (EntryMode = "E" OrElse EntryMode = "N")
                    Next

                Case "grdSOTSHIPX"
                    If tlb_pop.Tools.Contains("Print Manifest") Then
                        tlb_btn = DirectCast(tlb_pop.Tools("Print Manifest"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = optStatus.Value = "C"
                    End If

                Case "grdSOTCART1"
                    'tlb_btn = DirectCast(tlb_pop.Tools("Modify Tracking Number"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = False
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
                For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
                    rowSOTPICK1.Item("SELECTED") = IIf(e.Tool.Key = "Select All", "1", "0")
                Next
                Display_Totals()

            Case "Hide Details"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                splSOTPICK1.Panel2Collapsed = tlb_sbt.Checked

            Case "Refresh"
                ' REFRESH WHAT?

            Case "Permit Price Change"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)

                If dst.Tables("SOTORDR1").Select("ISNULL(EDI_DOC_SEQ_NO, '') <> '' or CURR_CODE <> 'USD'").Length > 0 Then
                    MessageBox.Show("You are not permitted to change prices on EDI orders and non USD orders.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    tlb_sbt.Checked = False
                End If

                With grdSOTPICK2.DisplayLayout.Bands(0).Columns("PICK_UNIT_PRICE")
                    If tlb_sbt.Checked Then
                        .CellActivation = UltraWinGrid.Activation.AllowEdit
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                        .CellAppearance.BackColor = Drawing.Color.Empty
                    Else
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        .CellAppearance.BackColor = Drawing.Color.Beige
                    End If
                End With

            Case "Print Manifest"
                If grdSOTSHIPX.ActiveRow Is Nothing Then
                    MessageBox.Show("Select a Shipment from the list of shipments.")
                    Exit Sub
                End If

                PrintManifest(grdSOTSHIPX.ActiveRow.Cells("SHIP_BOL_NO").Value & String.Empty)

            Case "Modify Tracking Number"
                If grdSOTCART1.Selected.Rows.Count = 0 Then
                    grdSOTCART1.Selected.Rows.Add(grdSOTCART1.ActiveRow)
                End If

                If MessageBox.Show("Do you want to change the Tracking Number for the " & grdSOTCART1.Selected.Rows.Count & " selected rows?", "Modify Tracking Number", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                Dim newTrackingNumber As String = String.Empty
                Using F As New ASFMSGBF
                    newTrackingNumber = F.Get_txt_from_User("New Tracking Number", "New Tracking Number", False, dst.Tables("SOTCART1").Columns("CART_TRACKING_NO").MaxLength, "")
                End Using

                If newTrackingNumber.Length = 0 Then
                    Exit Sub
                End If

                Try
                    BeginTrans()
                    For Each gridRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTCART1.Selected.Rows
                        Dim cart_no As String = gridRow.Cells("CART_NO").Value
                        ASCDATA1.ExecuteSQL("Update SOTCART1 set CART_TRACKING_NO = '" & newTrackingNumber & "' WHERE CART_NO = '" & cart_no & "'")
                        gridRow.Cells("CART_TRACKING_NO").Value = newTrackingNumber
                    Next
                    CommitTrans("Update Successful")

                Catch ex As Exception
                    Rollback(ex.Message)
                End Try

        End Select

        If grd Is Nothing Then
            Exit Sub
        Else
            If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
                Exit Sub
            End If
        End If

        Select Case e.Tool.Key

            Case "Propagate Value"
                Dim COLUMN_NAME As String = e.Tool.Tag
                For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
                    rowSOTPICK1.Item(COLUMN_NAME) = grdSOTPICK1.ActiveCell.Value
                Next

                Display_Totals()

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")

            Case "Clear EDI 945 Record"
                Dim ShipBolNo As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value
                Dim sql As String = String.Empty
                sql = "SELECT * FROM EDT945T1 WHERE EDI_PICK_NO IN"
                sql &= " ("
                sql &= " Select PICK_NO from SOTPICK1 WHERE SHIP_BOL_NO = '" & ShipBolNo & "'"
                sql &= " )"
                sql &= " and NVL(EDI_PROCESS_IND, '0') in ('0', '3', '4')"

                Dim tblWk As DataTable = ASCDATA1.GetDataTable(sql)
                Select Case tblWk.Rows.Count
                    Case 0
                        MessageBox.Show("Cannot locate the EDI 945 record for the selected entry.", "Clear EDI 945 Record", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    Case 1
                        If MessageBox.Show("Do you want to set the EDI 945 record for the Selected Entry to be ignored?" _
                                            , "Clear EDI 945 Record", MessageBoxButtons.YesNo _
                                            , MessageBoxIcon.Question _
                                            , MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If

                    Case Else
                        If MessageBox.Show("There are (" & tblWk.Rows.Count & ") EDI 945 transactions associated with the Selected Entry." _
                                            & " Do you want to set the EDI 945 records to be ignored?" _
                                            , "Clear EDI 945 Record", MessageBoxButtons.YesNo _
                                            , MessageBoxIcon.Question _
                                            , MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                End Select

                Try
                    BeginTrans()

                    sql = "UPDATE EDT945T1 SET EDI_PROCESS_IND = 'X' WHERE EDI_PICK_NO IN"
                    sql &= " ("
                    sql &= " Select PICK_NO from SOTPICK1 WHERE SHIP_BOL_NO = '" & ShipBolNo & "'"
                    sql &= " )"
                    sql &= " and NVL(EDI_PROCESS_IND, '0')  in ('0', '3', '4')"
                    ASCDATA1.ExecuteSQL(sql)

                    CommitTrans("EDI 945 Records set to be ignored.")

                    Fill_Records("WHT3PLS1")
                    Sort_grdColumns(grdWHT3PLS1, "SHIP_BOL_NO")

                Catch ex As Exception
                    Rollback(ex.Message)
                End Try

            Case "Set Master BOL No"
                Dim ShipBolNo As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value
                Dim EDIBolNo As String = dst.Tables("WHT3PLS1").Select("SHIP_BOL_NO = '" & ShipBolNo & "' AND EDI_MASTER_BOL_NO is not null", "EDI_MASTER_BOL_NO")(0).Item("EDI_MASTER_BOL_NO") & String.Empty
                If EDIBolNo.Length = 0 Then
                    EDIBolNo = dst.Tables("WHT3PLS1").Select("SHIP_BOL_NO = '" & ShipBolNo & "'", "EDI_BOL_NO")(0).Item("EDI_BOL_NO") & String.Empty
                End If

                If EDIBolNo.Length = 0 Then
                    MessageBox.Show("Could not reslove the EDI Master Bol NO to use. Process Aborted!", "Set Master Bol NO", MessageBoxButtons.OK)
                    Exit Sub
                End If

                If MessageBox.Show("Do you want to Update all Shipment Records (" & ShipBolNo & ") to have a Master Bol No of " & EDIBolNo & "?", _
                                     "Set Master Bol NO", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If
                SetMasterBOLNo(ShipBolNo, EDIBolNo)

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Load_SOTSHIPX()
                End If

            Case "SHIP_BOL_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("Select")
                End If

            Case "PICK_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Dim PICK_NO As String = Absx1.txtFor("PICK_NO").Text
                    Dim rowSOTPICK1 As DataRow = LookUp("SOTPICK1", PICK_NO)
                    If rowSOTPICK1 Is Nothing Then
                        MsgBox("Invalid Pick Ticket No Specified (" & PICK_NO & ")", MsgBoxStyle.OkOnly, "Cannot Peform Requested Action")
                    Else
                        Absx1.txtFor("SHIP_BOL_NO").Text = rowSOTPICK1.Item("SHIP_BOL_NO")
                        Click_Command("Select")
                    End If
                End If

            Case "CUST_ADDR_CODE"
                e.SuppressKeyPress = True
                ' e.Handled = True

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "CUST_CODE"
                If Not ScreenMode Then
                    Load_SOTSHIPX()
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Load_SOTSHIPX()

            Case "SHIPMENT_NO"
                Click_Command("Select")

            Case "CUST_ADDR_CODE"
                Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {Absx1.txtFor("CUST_CODE").Text, "MK", Absx1.txtFor("CUST_ADDR_CODE").Text})
                If rowARTCUST2 IsNot Nothing Then
                    Dim ORDR_NO As String = grdSOTPICK1.ActiveRow.Cells("ORDR_NO").Value
                    Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Rows.Find(New String() {ORDR_NO, "ST"}) ' .Select(dvwSOTORDR5.RowFilter)
                    For Each COLUMN_NAME As String In New String() _
                        {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", _
                         "CUST_COUNTRY", "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_EMAIL"}

                        If COLUMN_NAME = "CUST_ADDR3" Then
                        Else
                            'rowSOTORDR5.Item(COLUMN_NAME) = rowARTCUST2.Item(COLUMN_NAME)
                            If COLUMN_NAME = "CUST_EXT" OrElse COLUMN_NAME = "CUST_FAX" OrElse COLUMN_NAME = "CUST_EMAIL" Then
                            Else
                                If COLUMN_NAME = "CUST_PHONE" Then
                                    Absx1.medFor("SOTORDR5." & COLUMN_NAME).Value = rowARTCUST2.Item(COLUMN_NAME) & ""
                                Else
                                    Absx1.txtFor("SOTORDR5." & COLUMN_NAME).Text = rowARTCUST2.Item(COLUMN_NAME) & ""
                                End If
                            End If
                        End If
                    Next
                End If

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
                If EntryMode = "E" Then
                    'If Absx1.dteFor("INV_DATE").Value & "" = "" Then
                    Absx1.dteFor("INV_DATE").Value = Absx1.dteFor("SHIP_DATE_SHIPPED").Value
                    'End If
                End If
        End Select
    End Sub
#End Region

#Region "grdSOTPICK1"



    Private Sub grdSOTPICK1_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTPICK1.AfterCellUpdate
        If e.Cell.Column.Key = "PICK_FREIGHT" Then
            Display_Totals()
        End If
    End Sub

    Private Sub grdSOTPICK1_AfterColPosChanged(sender As Object, e As Infragistics.Win.UltraWinGrid.AfterColPosChangedEventArgs) Handles grdSOTPICK1.AfterColPosChanged
        Position_txtSTORE()
    End Sub

    Private Sub grdSOTPICK1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPICK1.AfterRowActivate
        Setup_SOTPICK1()
        Position_txtSTORE()
    End Sub

    Private Sub grdSOTPICK1_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTPICK1.AfterRowsDeleted
        Display_Totals()
    End Sub

    Private Sub grdSOTPICK1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTPICK1.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdSOTPICK1_BeforeRowActivate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTPICK1.BeforeRowActivate
        If grdSOTCART1.ActiveRow IsNot Nothing AndAlso grdSOTCART1.ActiveRow.DataChanged Then
            grdSOTCART1.ActiveRow.Update()
        End If
        If grdSOTPICK2.ActiveRow IsNot Nothing AndAlso grdSOTPICK2.ActiveRow.DataChanged Then
            grdSOTPICK2.ActiveRow.Update()
        End If
    End Sub


    Private Sub grdSOTPICK1_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTPICK1.BeforeRowUpdate
        If e.Row.Cells("PICK_STATUS").Value <> "P" Then
            MsgBox("You are attempting to make changes to a Pick Ticket" _
                    & vbCrLf & " that is Not In-Pick", _
                    MsgBoxStyle.OkOnly, "Changes are not Permitted to Pick Tickets NOT In Pick")
            e.Row.CancelUpdate()
            e.Cancel = True
        End If
    End Sub

    Private Sub grdSOTPICK1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTPICK1.InitializeRow
        If e.Row.Cells("PICK_STATUS").Value = "F" Then
            e.Row.Cells("PICK_STATUS").Appearance.ForeColor = Drawing.Color.Blue
        ElseIf e.Row.Cells("PICK_STATUS").Value <> "P" Then
            e.Row.Cells("PICK_STATUS").Appearance.ForeColor = Drawing.Color.Red
        End If
    End Sub

    Private Sub grdSOTPICK1_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles grdSOTPICK1.MouseUp
        If grdSOTPICK1.ActiveCell IsNot Nothing AndAlso grdSOTPICK1.ActiveCell.Column.Key = "SELECTED" Then
            grdSOTPICK1.ActiveRow.Update()
        End If
    End Sub

    Private Sub grdSOTPICK1_SizeChanged(sender As Object, e As System.EventArgs) Handles grdSOTPICK1.SizeChanged
        Position_txtSTORE()
    End Sub

    Sub Position_txtSTORE()

        If Not ScreenMode Then Exit Sub
        If grdSOTPICK1.Rows(0).Cells("CUST_STORE_NO").GetUIElement() IsNot Nothing Then
            Try
                txtStore.Parent = grdSOTPICK1
                Dim r As System.Drawing.Rectangle = grdSOTPICK1.Rows(0).Cells("CUST_STORE_NO").GetUIElement().ClipRect
                txtStore.Width = grdSOTPICK1.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").Header.SizeResolved.Width
                txtStore.Left = r.Left
                txtStore.Top = grdSOTPICK1.Top
            Catch ex As Exception

            End Try
        End If
    End Sub
#End Region

#Region "grdSOTPICK2"

    Private Sub grdSOTPICK2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTPICK2.InitializeRow
        If Val(e.Row.Cells("PICK_QTY").Value & "") <> Val(e.Row.Cells("PICK_QTY_CONF").Value & "") Then
            e.Row.Cells("PICK_QTY_CONF").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("PICK_QTY_CONF").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If

        If Val(e.Row.Cells("PICK_UNIT_PRICE").Value & "") <> Val(e.Row.Cells("ORDR_UNIT_PRICE").Value & "") Then
            e.Row.Cells("PICK_UNIT_PRICE").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("PICK_UNIT_PRICE").Appearance.ForeColor = Drawing.Color.Empty
        End If

    End Sub

    Private Sub grdSOTPICK2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTPICK2.AfterCellUpdate
        With grdSOTPICK1.ActiveRow
            Select Case e.Cell.Column.Key
                Case "PICK_UNIT_PRICE"

                Case "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"
                    If e.Cell.Tag = "X" Then Exit Sub
                    e.Cell.Tag = "X"
                    Dim PICK_QTY As Int64 = Val(e.Cell.Row.Cells("PICK_QTY").Value & "")
                    Dim PICK_QTY_CONF As Int64 = Val(e.Cell.Row.Cells("PICK_QTY_CONF").Value & "")
                    Dim PICK_QTY_CANC As Int64 = Val(e.Cell.Row.Cells("PICK_QTY_CANC").Value & "")
                    Dim PICK_QTY_BACK As Int64 = Val(e.Cell.Row.Cells("PICK_QTY_BACK").Value & "")

                    If PICK_QTY_CONF < PICK_QTY Then
                        If e.Cell.Column.Key = "PICK_QTY_BACK" Then
                            PICK_QTY_CANC = PICK_QTY - PICK_QTY_CONF - PICK_QTY_BACK
                            If PICK_QTY_CANC < 0 Then
                                PICK_QTY_CANC = 0
                            End If
                            e.Cell.Row.Cells("PICK_QTY_CANC").Value = PICK_QTY_CANC
                        Else
                            If chkBO.Checked Then

                                PICK_QTY_BACK = PICK_QTY - PICK_QTY_CONF - PICK_QTY_CANC

                                If PICK_QTY_BACK < 0 Then
                                    PICK_QTY_BACK = 0
                                End If
                                e.Cell.Row.Cells("PICK_QTY_BACK").Value = PICK_QTY_BACK
                            Else

                                PICK_QTY_CANC = PICK_QTY - PICK_QTY_CONF

                                If PICK_QTY_CANC < 0 Then
                                    PICK_QTY_CANC = 0
                                End If
                                e.Cell.Row.Cells("PICK_QTY_CANC").Value = PICK_QTY_CANC
                            End If
                        End If
                    Else
                        If PICK_QTY_CONF >= PICK_QTY Then
                            e.Cell.Row.Cells("PICK_QTY_CANC").Value = 0
                            e.Cell.Row.Cells("PICK_QTY_BACK").Value = 0
                        End If
                    End If
                    e.Cell.Tag = ""
            End Select
        End With
    End Sub

    Private Sub grdSOTPICK2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPICK2.AfterRowActivate
        Dim STYLE_CODE As String = grdSOTPICK2.ActiveRow.Cells("STYLE_CODE").Value
        Dim COLOR_CODE As String = grdSOTPICK2.ActiveRow.Cells("COLOR_CODE").Value
        optSCB.ValueList.ValueListItems(1).DisplayText = "Sty/Clr " & STYLE_CODE & "/" & COLOR_CODE
        optSCB.ValueList.ValueListItems(1).Tag = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"

        With grdSOTPICK2.ActiveRow
            clsPrice_Change = New Price_Change
            clsPrice_Change.PICK_NO = .Cells("PICK_NO").Value & String.Empty
            clsPrice_Change.PICK_LNO = .Cells("PICK_LNO").Value
            clsPrice_Change.STYLE_CODE = .Cells("STYLE_CODE").Value & String.Empty
            clsPrice_Change.COLOR_CODE = .Cells("COLOR_CODE").Value & String.Empty
            clsPrice_Change.PICK_UNIT_PRICE = Val(.Cells("PICK_UNIT_PRICE").Value & String.Empty)
        End With

        Setup_SOTCART2_from_SOTPICK2()
    End Sub

    Private Sub grdSOTPICK2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTPICK2.AfterRowsDeleted
        Display_Totals()
    End Sub

    Private Sub grdSOTPICK2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTPICK2.AfterRowUpdate
        If clsPrice_Change Is Nothing Then Exit Sub
        If clsPrice_Change.PICK_UNIT_PRICE <> e.Row.Cells("PICK_UNIT_PRICE").Value _
            And clsPrice_Change.PICK_NO = e.Row.Cells("PICK_NO").Value _
            And clsPrice_Change.PICK_LNO = e.Row.Cells("PICK_LNO").Value Then
            Dim sqlw As String = "(PICK_NO <> '" & clsPrice_Change.PICK_NO & "' or PICK_LNO <> " & CStr(clsPrice_Change.PICK_LNO) & ")" _
             & " and PICK_UNIT_PRICE = " & CStr(clsPrice_Change.PICK_UNIT_PRICE) _
             & " and STYLE_CODE = '" & clsPrice_Change.STYLE_CODE & "'" _
             & " and COLOR_CODE = '" & clsPrice_Change.COLOR_CODE & "'"
            Me.Cursor = Cursors.WaitCursor

            '   Update the sales order for the line changed
            Dim ORDR_NO As String = e.Row.Cells("ORDR_NO").Value & String.Empty
            Dim ORDR_LNO As Int16 = Val(e.Row.Cells("ORDR_LNO").Value & String.Empty)
            Dim PICK_UNIT_PRICE As Decimal = Val(e.Row.Cells("PICK_UNIT_PRICE").Value & String.Empty)

            dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_LNO = " & ORDR_LNO)(0).Item("ORDR_UNIT_PRICE") = PICK_UNIT_PRICE
            Dim CURR_EXCH_RATE As Decimal = Val(dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0).Item("CURR_EXCH_RATE") & String.Empty)
            If CURR_EXCH_RATE <= 0 Then
                CURR_EXCH_RATE = 1
            End If
            dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_LNO = " & ORDR_LNO)(0).Item("ORDR_UNIT_PRICE_CURR") = PICK_UNIT_PRICE / CURR_EXCH_RATE

            ASCMAIN1.Progress("Now Changing Price for All Lines with Same Style/Color")
            SOTPICK1_Expressions(True)
            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select(sqlw)
                rowSOTPICK2.Item("PICK_UNIT_PRICE") = PICK_UNIT_PRICE
                ORDR_NO = rowSOTPICK2.Item("ORDR_NO") & String.Empty
                ORDR_LNO = Val(rowSOTPICK2.Item("ORDR_LNO") & String.Empty)

                dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_LNO = " & ORDR_LNO)(0).Item("ORDR_UNIT_PRICE") = PICK_UNIT_PRICE
                CURR_EXCH_RATE = Val(dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0).Item("CURR_EXCH_RATE") & String.Empty)
                If CURR_EXCH_RATE <= 0 Then
                    CURR_EXCH_RATE = 1
                End If
                dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_LNO = " & ORDR_LNO)(0).Item("ORDR_UNIT_PRICE_CURR") = PICK_UNIT_PRICE / CURR_EXCH_RATE
            Next
            SOTPICK1_Expressions(False)

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If
        clsPrice_Change = Nothing

        If MaintenanceMode Then
            Dim PICK_QTY_CONF As Int32 = Val(grdSOTPICK2.ActiveRow.Cells("PICK_QTY_CONF").Value & "")
            Dim PICK_NO As String = grdSOTPICK2.ActiveRow.Cells("PICK_NO").Value
            Dim PICK_LNO As String = Val(grdSOTPICK2.ActiveRow.Cells("PICK_LNO").Value & "")
            Dim sqlw As String = "PICK_NO = '" & PICK_NO & "'  and ORDR_LNO = " & CStr(PICK_LNO)
            Dim QTY_PACKED As Int32 = Val(dst.Tables("SOTCART2").Compute("SUM(QTY_PACKED)", sqlw) & "")
            Dim LINES As Int32 = dst.Tables("SOTCART2").Select(sqlw).Length
            If LINES = 1 And PICK_QTY_CONF <> QTY_PACKED Then
                Dim row As DataRow = dst.Tables("SOTCART2").Select(sqlw)(0)
                row.Item("QTY_PACKED") = PICK_QTY_CONF
            End If
        End If

        Display_Totals()
    End Sub

    Sub SOTPICK1_Expressions(remove_expressions As Boolean)
        If remove_expressions Then
            expSOTPICK1.Clear()
            For Each fCOLUMN_NAME As String In New String() {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK", _
                                                             "PICK_AMT", "PICK_AMT_CONF", "PICK_AMT_CANC", "PICK_AMT_BACK", _
                                                             "PICK_TOTAL_WGT_CALC", "PICK_CNT_CARTONS_CALC", "PICK_TOTAL_UNITS_CALC"}
                expSOTPICK1.Add(fCOLUMN_NAME, dst.Tables("SOTPICK1").Columns(fCOLUMN_NAME).Expression)
                dst.Tables("SOTPICK1").Columns(fCOLUMN_NAME).Expression = ""
            Next
        Else
            For Each fCOLUMN_NAME As String In expSOTPICK1.Keys
                dst.Tables("SOTPICK1").Columns(fCOLUMN_NAME).Expression = expSOTPICK1(fCOLUMN_NAME)
            Next
        End If
    End Sub

    Private Sub grdSOTPICK2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTPICK2.BeforeExitEditMode
        Select Case grdSOTPICK2.ActiveCell.Column.Key
            Case "PICK_UNIT_PRICE"
                If Val(grdSOTPICK2.ActiveCell.Value & "") < 0 Then
                    e.Cancel = True
                End If
            Case "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"
                If Val(grdSOTPICK2.ActiveCell.Value & "") < 0 Then
                    e.Cancel = True
                End If
        End Select
    End Sub

    Private Sub grdSOTPICK2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTPICK2.BeforeRowsDeleted

    End Sub

    Private Sub grdSOTPICK2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTPICK2.BeforeRowUpdate

        If grdSOTPICK1.ActiveRow.Cells("PICK_STATUS").Value <> "P" Then
            MsgBox("You are attempting to make changes to a Pick Ticket" _
                    & vbCrLf & " that is Not In-Pick", _
                    MsgBoxStyle.OkOnly, "Changes are not Permitted to Pick Tickets NOT In Pick")
            e.Row.CancelUpdate()
            e.Cancel = True
            Exit Sub
        End If

        Dim PICK_QTY As Int64 = Val(e.Row.Cells("PICK_QTY").Value)
        Dim PICK_QTY_CONF As Int64 = Val(e.Row.Cells("PICK_QTY_CONF").Value & "")
        Dim PICK_QTY_CANC As Int64 = Val(e.Row.Cells("PICK_QTY_CANC").Value & "")
        Dim PICK_QTY_BACK As Int64 = Val(e.Row.Cells("PICK_QTY_BACK").Value & "")
        Dim PICK_QTY_CANC_REL As Int64 = Val(e.Row.Cells("PICK_QTY_CANC_REL").Value & "")

        If MaintenanceMode Then
            If Val(e.Row.Cells("PICK_QTY_CONF").Value & "") > PICK_QTY + PICK_QTY_CANC_REL Then
                e.Row.Cells("PICK_QTY_CONF").Value = PICK_QTY - PICK_QTY_CANC - PICK_QTY_BACK
            End If
        End If

        If PICK_QTY_CONF < 0 OrElse PICK_QTY_BACK > PICK_QTY OrElse PICK_QTY_BACK < 0 OrElse PICK_QTY_CANC > PICK_QTY OrElse PICK_QTY_CANC < 0 Then
            e.Cancel = True
            Exit Sub
        End If

        If chkBO.Checked Then
            PICK_QTY_BACK = PICK_QTY - PICK_QTY_CONF - PICK_QTY_CANC
            If PICK_QTY_BACK < 0 Then
                PICK_QTY_BACK = 0
            End If
        Else
            PICK_QTY_CANC = PICK_QTY - PICK_QTY_CONF - PICK_QTY_BACK
            If PICK_QTY_CANC < -1 * PICK_QTY_CANC_REL Then
                PICK_QTY_CANC = 0
            End If
        End If
        If PICK_QTY_CONF > PICK_QTY + PICK_QTY_CANC_REL Then
            PICK_QTY_CANC = 0
            PICK_QTY_BACK = 0
        End If
        e.Row.Cells("PICK_QTY_CANC").Value = PICK_QTY_CANC
        e.Row.Cells("PICK_QTY_BACK").Value = PICK_QTY_BACK
    End Sub

    Private Sub grdSOTPICK2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTPICK2.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key

            End Select
        End With

    End Sub
#End Region

#Region "grdSOTCART1"

    Private Sub grdSOTCART1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTCART1.InitializeRow
        If Val(e.Row.Cells("CART_TOTAL_UNITS_CALC").Value & "") <> Val(e.Row.Cells("CART_TOTAL_UNITS_ORIG").Value & "") Then
            e.Row.Cells("CART_TOTAL_UNITS_CALC").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("CART_TOTAL_UNITS_CALC").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Private Sub grdSOTCART1_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCART1.AfterCellUpdate

        Dim displayBoxAttributes As Boolean = False

        For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTCART1.Rows
            If row.Cells("PKG_CODE").Text = "OTHER" Then
                displayBoxAttributes = True
                Exit For
            End If
        Next

        If displayBoxAttributes Then
            grdSOTCART1.DisplayLayout.Bands(0).Columns("WIDTH").Hidden = False
            grdSOTCART1.DisplayLayout.Bands(0).Columns("LENGTH").Hidden = False
            grdSOTCART1.DisplayLayout.Bands(0).Columns("HEIGHT").Hidden = False

            grdSOTCART1.DisplayLayout.Bands(0).Columns("WIDTH").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("LENGTH").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("HEIGHT").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdSOTCART1.DisplayLayout.Bands(0).Columns("WIDTH").Hidden = True
            grdSOTCART1.DisplayLayout.Bands(0).Columns("LENGTH").Hidden = True
            grdSOTCART1.DisplayLayout.Bands(0).Columns("HEIGHT").Hidden = True

            grdSOTCART1.DisplayLayout.Bands(0).Columns("WIDTH").CellActivation = UltraWinGrid.Activation.NoEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("LENGTH").CellActivation = UltraWinGrid.Activation.NoEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("HEIGHT").CellActivation = UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Private Sub grdSOTCART1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCART1.AfterRowActivate
        Setup_SOTCART2_from_SOTCART1()
    End Sub

    Private Sub grdSOTCART1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCART1.AfterRowUpdate
        Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Text

        Dim pickWeight As Decimal = Val(grdSOTPICK1.ActiveRow.Cells("PICK_TOTAL_WGT").Value & String.Empty)
        Dim totalCartWeight As Decimal = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)

        If ASCMAIN1.CLIENT = "NYA" Then
            If totalCartWeight > pickWeight Then
                grdSOTPICK1.ActiveRow.Cells("PICK_TOTAL_WGT").Value = totalCartWeight
                grdSOTPICK1.UpdateData()
            End If
        Else
            grdSOTPICK1.ActiveRow.Cells("PICK_TOTAL_WGT").Value = totalCartWeight
            grdSOTPICK1.UpdateData()
        End If

    End Sub

    Private Sub grdSOTCART1_BeforeRowActivate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCART1.BeforeRowActivate
        If grdSOTCART2.ActiveRow IsNot Nothing AndAlso grdSOTCART2.ActiveRow.DataChanged Then
            grdSOTCART2.ActiveRow.Update()
        End If
    End Sub

    Private Sub grdSOTCART1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCART1.ClickCellButton

        If grdSOTCART1.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.True Then
            Exit Sub
        End If

        If e.Cell.Column.Key = "CART_TOTAL_WGT_ACTUAL" Then
            e.Cell.Value = GetScaleWeight()
        End If
    End Sub

#End Region

#Region "grdSOTCART2"

    Private Sub grdSOTCART2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTCART2.InitializeRow
        If Val(e.Row.Cells("QTY_PACKED").Value & "") <> Val(e.Row.Cells("QTY_PACKED_ORIG").Value & "") Then
            e.Row.Cells("QTY_PACKED").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("QTY_PACKED").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

#End Region

#Region "grdWHT3PLS1"

    Private Sub grdWHT3PLS1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHT3PLS1.InitializeRow

        If ASCMAIN1.DBS_SERVER = "NYA" OrElse ASCMAIN1.DBS_COMPANY = "NYA" Then

            Select Case e.Row.Cells("EDI_PROCESS_IND").Value & ""
                Case NonEdiCustomer
                    e.Row.Cells("SHIP_BOL_NO").Appearance.BackColor = Drawing.Color.Green

                Case EdiErrorRecord
                    e.Row.Cells("SHIP_BOL_NO").Appearance.BackColor = Drawing.Color.Red
            End Select
        End If
    End Sub

    Private Sub grdWHT3PLS1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHT3PLS1.DoubleClickRow

        Dim SHIP_BOL_NO As String = (e.Row.Cells("SHIP_BOL_NO").Value & String.Empty).trim
        If SHIP_BOL_NO.Length = 0 Then Exit Sub

        processing_select_from_3PL_list = True
        selectedEDI_BOL_NO = (e.Row.Cells("EDI_BOL_NO").Value & String.Empty).trim

        If dst.Tables("WHT3PLS1").Columns.Contains("EDI_MASTER_BOL_NO") Then
            selectedMasterEDI_BOL_NO = (e.Row.Cells("EDI_MASTER_BOL_NO").Value & String.Empty).trim
        End If

        select_from_3PL_list = True
        Absx1.txtFor("SHIP_BOL_NO").Text = SHIP_BOL_NO
        Click_Command("Select")
        processing_select_from_3PL_list = False

    End Sub

#End Region

#Region "grdSOTSHIPX"

    Private Sub grdSOTSHIPX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTSHIPX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("SHIP_BOL_NO").Text = e.Row.Cells("SHIP_BOL_NO").Value
            Click_Command("Select")
        End If
    End Sub

#End Region

#Region "grdSOTINVHM"

    Private Sub grdSOTINVHM_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTINVHM.AfterCellUpdate
        With grdSOTINVHM.ActiveRow
            Select Case e.Cell.Column.Key
                Case "MISC_CHG_CODE"
                    Dim MISC_CHG_CODE As String = Validate_MISC_CHG_CODE(.Cells("MISC_CHG_CODE").Value & "")
                    If MISC_CHG_CODE <> "" Then
                        .Cells("MISC_CHG_DESC").Value = rowSOTMISC1.Item("MISC_CHG_DESC") & String.Empty
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSOTINVHM_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTINVHM.AfterRowActivate

        If Trim(grdSOTINVHM.ActiveRow.Cells("MISC_CHG_CODE").Value & "") = "" And _
            (grdSOTINVHM.ActiveCell Is Nothing OrElse _
             (grdSOTINVHM.ActiveCell.Column.Key <> "MISC_CHG_CODE")) _
        Then
            grdSOTINVHM.ActiveCell = grdSOTINVHM.ActiveRow.Cells("MISC_CHG_CODE")
            Exit Sub
        End If

        If grdSOTINVHM.ActiveRow.IsAddRow Then
            If grdSOTINVHM.ActiveRow.Cells("MISC_CHG_CODE").Value & "" = "" Then
                grdSOTINVHM.ActiveCell = grdSOTINVHM.ActiveRow.Cells("MISC_CHG_CODE")
            End If
        Else
            With grdSOTINVHM.DisplayLayout.Bands(0)
                Validate_MISC_CHG_CODE(grdSOTINVHM.ActiveRow.Cells("MISC_CHG_CODE").Value & "")
            End With
        End If
    End Sub

    Private Sub grdSOTINVHM_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTINVHM.AfterRowsDeleted
        Display_Totals()
    End Sub

    Private Sub grdSOTINVHM_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTINVHM.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdSOTINVHM_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTINVHM.BeforeExitEditMode
        If grdSOTINVHM.ActiveCell IsNot Nothing Then
            With grdSOTINVHM.ActiveCell
                Select Case .Column.Key
                    Case "MISC_CHG_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTINVHM_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTINVHM.BeforeRowUpdate

        If Validate_MISC_CHG_CODE(e.Row.Cells("MISC_CHG_CODE").Value & "") = "" Then
            e.Cancel = True
        End If

        If e.Cancel = True Then
            MessageBox.Show("Invalid Charge Code", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim INV_MISC_CHG As Decimal = Val(e.Row.Cells("INV_MISC_CHG").Value & String.Empty)
        If INV_MISC_CHG <= 0 AndAlso ASCMAIN1.CLIENT <> "RGI" Then
            MessageBox.Show("Charge amount must be greater than 0.00", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("INV_TYPE").Value = "I"
            ' This is done since the invoice is not created yet.
            e.Row.Cells("INV_NO").Value = "XXX"
            Dim INV_MNO As Int64 = Val(dst.Tables("SOTINVHM").Compute("MAX(INV_MNO)", "") & "") + 1
            e.Row.Cells("INV_MNO").Value = INV_MNO
        End If
    End Sub

    Private Sub grdSOTINVHM_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTINVHM.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "MISC_CHG_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdSOTINVHM, sql_where)
            End Select
        End With
    End Sub

    Function Validate_MISC_CHG_CODE(MISC_CHG_CODE As String) As String
        rowSOTMISC1 = LookUp("SOTMISC1", MISC_CHG_CODE)
        If rowSOTMISC1 Is Nothing Then
            Return ""
        Else
            Return rowSOTMISC1.Item("MISC_CHG_CODE")
        End If
    End Function

#End Region

#Region "Form Controls"

    Private Sub chkBO_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBO.CheckedChanged
        If chkBO.Checked Then
            If edi_customer Then
                If MsgBox("Override the Rule?", MsgBoxStyle.YesNo, _
                          "EDI Customers do not Allow Back Orders") = MsgBoxResult.No Then
                    chkBO.Checked = False
                End If
            End If
        End If
        Setup_BO()
    End Sub

    Private Sub txtStore_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtStore.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            If txtStore.Text = "" Then
                MsgBox("You Must First Enter a Store No", MsgBoxStyle.OkOnly, "Cannot Locate Pick Ticket for Selected Store")
                Exit Sub
            Else
                txtStore.Text = txtStore.Text.PadLeft(6, "0")
            End If

            grdSOTPICK1.ActiveRow = Nothing
            grdSOTPICK1.Selected.Rows.Clear()
            For Each grow As UltraWinGrid.UltraGridRow In grdSOTPICK1.Rows
                If grow.Cells("CUST_STORE_NO").Value & "" = txtStore.Text Then
                    grdSOTPICK1.ActiveRow = grow
                    grow.Selected = True
                    Exit For
                End If
            Next
            If grdSOTPICK1.ActiveRow Is Nothing Then
                MsgBox("No Pick Ticket Found for Store " & txtStore.Text, MsgBoxStyle.OkOnly, "Cannot Locate Pick Ticket for Selected Store")
            End If
            txtStore.Text = ""
        End If
    End Sub

    Private Sub cmdSHIP_Click(sender As System.Object, e As System.EventArgs) Handles cmdSHIP.Click
        SCB("PICK_QTY_CONF")
    End Sub

    Private Sub cmdCANC_Click(sender As System.Object, e As System.EventArgs) Handles cmdCANC.Click
        SCB("PICK_QTY_CANC")
    End Sub

    Private Sub cmdBACK_Click(sender As System.Object, e As System.EventArgs) Handles cmdBACK.Click
        SCB("PICK_QTY_BACK")
    End Sub

    Private Sub optStatus_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optStatus.ValueChanged
        grpConfirmed.Visible = (optStatus.Value = "C")
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTSHIPX()

        tabSelect.Tabs(0).Text = optStatus.CheckedItem.DisplayText
    End Sub

    Private Sub btnLoadHistory_Click(sender As System.Object, e As System.EventArgs) Handles btnLoadHistory.Click
        Load_SOTSHIPX()
    End Sub

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

            ' If set to Recipient then do not change.
            If optPayor.Value <> "R" Then
                If rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" Then
                    optPayor.Value = "C"
                ElseIf rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" Then
                    optPayor.Value = "P"
                Else
                    optPayor.Value = "O"
                End If
            End If

            If rowSOTCARR1.Item("CARRIER_TYPE") & String.Empty <> "U" Then
                txt3PAccountNo.Clear()
                txt3pCountry.Clear()
                txt3PZipCode.Clear()
                Exit Sub
            End If

            If rowSOTCARR1.Item("SHIP_ACCT_NO") & String.Empty <> String.Empty Then
                ' Prepopulate any Account numbers if the user did not provide them
                Select Case rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty
                    Case "F"
                        If txt3PAccountNo.TextLength = 0 Then txt3PAccountNo.Text = (rowARTCUST1.Item("FDX_ACCT_NO") & String.Empty).ToString.Trim
                        If txt3pCountry.TextLength = 0 Then txt3pCountry.Text = (rowARTCUST1.Item("FDX_3PY_COUNTRY") & String.Empty).ToString.Trim
                        If txt3PZipCode.TextLength = 0 Then txt3PZipCode.Text = (rowARTCUST1.Item("FDX_3PY_ZIPCODE") & String.Empty).ToString.Trim
                    Case "U"
                        If txt3PAccountNo.TextLength = 0 Then txt3PAccountNo.Text = (rowARTCUST1.Item("UPS_ACCT_NO") & String.Empty).ToString.Trim
                        If txt3pCountry.TextLength = 0 Then txt3pCountry.Text = (rowARTCUST1.Item("UPS_3PY_COUNTRY") & String.Empty).ToString.Trim
                        If txt3PZipCode.TextLength = 0 Then txt3PZipCode.Text = (rowARTCUST1.Item("UPS_3PY_ZIPCODE") & String.Empty).ToString.Trim
                End Select

                If txt3PAccountNo.TextLength = 0 Then
                    txt3PAccountNo.Text = rowSOTCARR1.Item("SHIP_ACCT_NO") & String.Empty
                    txt3pCountry.Text = rowSOTCARR1.Item("SHIP_3PY_COUNTRY") & String.Empty
                    txt3PZipCode.Text = rowSOTCARR1.Item("SHIP_3PY_ZIPCODE") & String.Empty
                End If
                txt3PAccountNo.Tag = rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty
                optPayor.Value = "P"
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
                    If txt3pCountry.TextLength = 0 Then txt3pCountry.Text = (rowARTCUST1.Item("FDX_3PY_COUNTRY") & String.Empty).ToString.Trim
                    If txt3PZipCode.TextLength = 0 Then txt3PZipCode.Text = (rowARTCUST1.Item("FDX_3PY_ZIPCODE") & String.Empty).ToString.Trim
                Case "U"
                    If txt3PAccountNo.TextLength = 0 Then txt3PAccountNo.Text = (rowARTCUST1.Item("UPS_ACCT_NO") & String.Empty).ToString.Trim
                    If txt3pCountry.TextLength = 0 Then txt3pCountry.Text = (rowARTCUST1.Item("UPS_3PY_COUNTRY") & String.Empty).ToString.Trim
                    If txt3PZipCode.TextLength = 0 Then txt3PZipCode.Text = (rowARTCUST1.Item("UPS_3PY_ZIPCODE") & String.Empty).ToString.Trim
            End Select
        End If

    End Sub

    Private Sub chkInsureShipment_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkInsureShipment.CheckedChanged
        If chkInsureShipment.Checked Then
            grdSOTCART1.DisplayLayout.Bands(0).Columns("INSURANCE").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("INSURANCE").Hidden = False
        Else
            grdSOTCART1.DisplayLayout.Bands(0).Columns("INSURANCE").CellActivation = UltraWinGrid.Activation.NoEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("INSURANCE").Hidden = True
        End If
    End Sub

#End Region

#Region "Form Procedures"

    Private Sub RecordPriceChanges()
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            Dim ORDR_UNIT_PRICE_orig As Decimal = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE_ORIG") & String.Empty)
            Dim ORDR_UNIT_PRICE As Decimal = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & String.Empty)

            If ORDR_UNIT_PRICE <> ORDR_UNIT_PRICE_orig Then

                Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
                rowTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
                rowTATEVNT1.Item("TABLE_KEY") = rowSOTORDR2.Item("ORDR_NO")
                rowTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
                rowTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowTATEVNT1.Item("EVENT_TYPE") = "UPC"
                rowTATEVNT1.Item("EVENT_DESC") = "Style / Color: " & rowSOTORDR2.Item("STYLE_CODE") & " / " & rowSOTORDR2.Item("COLOR_CODE") _
                    & " price changed from " & ORDR_UNIT_PRICE_orig & " to " & ORDR_UNIT_PRICE
                rowTATEVNT1.Item("EVENT_KEY") = ""
                rowTATEVNT1.Item("FORM_NAME") = "SOFSHIP0"
                dst.Tables("TATEVNT1").Rows.Add(rowTATEVNT1)
            End If
        Next

    End Sub

    Public Function EmailInvoice() As Boolean

        Dim INV_NO As String = String.Empty
        Dim attachFileName As String = String.Empty

        Try
            'Prevents emailing in the TST companies.
            If ASCMAIN1.CLIENT <> ASCMAIN1.DBS_SERVER Then
                Return True
            End If

            If dst.Tables("SOTINVH1").Rows.Count = 0 Then
                Return False
            End If

            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").Rows(0)
            INV_NO = rowSOTINVH1.Item("INV_NO")
            Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = rowSOTINVH1.Item("CUST_STORE_NO")
            Dim SREP_CODE As String = rowSOTINVH1.Item("SREP_CODE") & String.Empty
            Dim salesRepEmail As String = String.Empty
            Dim CUST_XMIT_INV_VIA As String = String.Empty

            Dim CONS_INV As String = "0"
            If rowSOTINVH1.Item("INV_NO_CONS") & String.Empty <> String.Empty Then
                CONS_INV = "1"
            End If

            ' See if the customer receives an acknowledgment
            Dim rowSOTSREP1 As DataRow = LookUp("SOTSREP1", SREP_CODE)
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)

            Select Case ASCMAIN1.CLIENT
                Case "RGI"
                    If rowSOTSREP1 Is Nothing AndAlso rowARTCUST1 Is Nothing Then
                        Return False
                    End If

                Case Else
                    If rowARTCUST1 Is Nothing Then
                        Return False
                    End If
            End Select


            ' See if we have anyone to email to - Only RGI sends a copy of the invoice to the sales rep
            If ASCMAIN1.CLIENT = "RGI" Then
                If rowSOTSREP1 IsNot Nothing AndAlso rowSOTSREP1.Item("SREP_EMAIL") & String.Empty <> String.Empty Then
                    salesRepEmail = rowSOTSREP1.Item("SREP_EMAIL") & String.Empty
                End If
            Else
                salesRepEmail = String.Empty
            End If

            If rowARTCUST1 IsNot Nothing Then
                ' Mail, Email, Both
                CUST_XMIT_INV_VIA = (rowARTCUST1.Item("CUST_XMIT_INV_VIA") & String.Empty).ToString.Trim
                If CUST_XMIT_INV_VIA.Length > 0 AndAlso "EB".Contains(CUST_XMIT_INV_VIA) Then
                    If rowARTCUST1.Item("CUST_INV_EMAIL") & String.Empty <> String.Empty Then
                        salesRepEmail &= ";" & rowARTCUST1.Item("CUST_INV_EMAIL") & String.Empty
                    End If

                    If rowARTCUST1.Item("CUST_INV_CC") & String.Empty <> String.Empty Then
                        salesRepEmail &= ";" & rowARTCUST1.Item("CUST_INV_CC") & String.Empty
                    End If
                End If
            End If

            ' remove double semi-colons
            salesRepEmail = salesRepEmail.Replace(",", ";")
            salesRepEmail = salesRepEmail.Replace(";;", ";")
            salesRepEmail = salesRepEmail.Replace(" ", "")

            ' should be at least 5 characters
            If salesRepEmail.Replace(";", "").Trim.Length < 5 Then
                Return False
            End If

            attachFileName = rowARTCUST1.Item("CUST_NAME") & " " & INV_NO

            For Each invalidChar As String In New String() {"\", "/", ":", "*", "?", "<", ">", "|", "."}
                attachFileName = attachFileName.Replace(invalidChar, "")
            Next
            attachFileName = attachFileName.Replace(" ", "_")

            Dim invNos As String = String.Empty
            For Each row As DataRow In dst.Tables("SOTINVH1").Select("")
                invNos &= ", '" & row.Item("INV_NO") & "'"
            Next
            invNos = invNos.Substring(1).Trim

            ASCMAIN1.Progress("Emailing Invoice", "")

            Dim RPT As String = "SORINVP1"
            If Not REPORTS.ContainsKey(RPT) Then
                REPORTS.Add(RPT, Load_rptClass(RPT))
                REPORTS(RPT).Prepare_dst(False, "")
            End If

            REPORTS(RPT).Fill_Records_RPT(New String() {" and SOTINVH1.INV_NO IN (" & invNos & ")"})

            Dim REPORT_NO As String = String.Empty
            With REPORTS(RPT).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", CONS_INV)
                .CR_params.Add("EXPORT_INFO", "0")

                ' Set the customers Invoice
                Select Case ASCMAIN1.CLIENT
                    Case "RGI"
                        RPT = "SORINVPR"

                End Select

                REPORT_NO = .Generate_Report(RPT, "Invoice", , True, , , "PDF", attachFileName, False)
                .Print_Report_End(True, True)
            End With

            Dim ATTACHMENTs As New Dictionary(Of String, String)
            ATTACHMENTs.Add(attachFileName & ".pdf", ASCMAIN1.Folders("Temp") & attachFileName & ".pdf")

            Dim SUBJECT As String = String.Empty
            SUBJECT = "Sales Invoice (" & INV_NO & ") for customer " & rowARTCUST1.Item("CUST_NAME")

            ' Concatentate and process all email addresses
            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            For Each emailAddress As String In (salesRepEmail).ToString.Split(";")
                emailAddress = emailAddress.Trim
                If emailAddress.Length > 5 AndAlso Not EMAIL_ADDRESSs.Keys.Contains(emailAddress) Then
                    EMAIL_ADDRESSs.Add(emailAddress, emailAddress)
                End If
            Next

            If EMAIL_ADDRESSs.Count = 0 Then
                Return True
            End If

            Dim EMAIL_KEY As String = "INV"
            Select Case ASCMAIN1.CLIENT
                Case "RGI"
                    EMAIL_KEY = "AUTOINV"
            End Select

            Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                   (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                    SUBJECT, EMAIL_KEY, _
                    True, False, CUST_CODE, rowARTCUST1.Item("CUST_NAME"), "Customer")


            ' Mark email Only Invoices as Mailed
            Try
                If CUST_XMIT_INV_VIA = "E" Then
                    For Each rowSOTINVH1 In dst.Tables("SOTINVH1").Rows
                        INV_NO = rowSOTINVH1.Item("INV_NO")
                        ASCDATA1.ExecuteSQL("Update SOTINVH1 Set INV_PRINTED = SYSDATE where INV_NO = '" & INV_NO & "'")
                    Next
                End If
            Catch ex As Exception
                ' nothing 
            End Try

            EmailInvoice = True

        Catch ex As Exception
            EmailInvoice = False
        End Try

    End Function

    Private Sub CreateSOTSHIPB(ByVal rowSOTSHIP1 As DataRow)

        Dim rowSOTSHIPB As DataRow = Nothing

        Dim BILL_OF_LADING_NO As String = (rowSOTSHIP1.Item("BILL_OF_LADING_NO") & String.Empty).ToString.Trim

        If BILL_OF_LADING_NO.Length Then
            rowSOTSHIPB = dst.Tables("SOTSHIPB").NewRow
            rowSOTSHIPB.Item("BOL_NO") = ""
        ElseIf dst.Tables("SOTSHIPB").Select("BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "'").Length = 0 Then
            rowSOTSHIPB = dst.Tables("SOTSHIPB").NewRow
            rowSOTSHIPB.Item("BOL_NO") = ""
        Else
            rowSOTSHIPB = dst.Tables("SOTSHIPB").Select("BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "'")(0)
        End If

        rowSOTSHIPB.Item("CUST_CODE") = rowSOTSHIP1.Item("CUST_CODE") & String.Empty
        rowSOTSHIPB.Item("BOL_DATE") = DateTime.Now.ToShortDateString
        rowSOTSHIPB.Item("FRT_TERMS") = rowSOTSHIP1.Item("FRT_TERMS") & String.Empty
        rowSOTSHIPB.Item("WHSE_CODE") = rowSOTSHIP1.Item("WHSE_CODE") & String.Empty
        rowSOTSHIPB.Item("MASTER_BOL_NO") = String.Empty
        rowSOTSHIPB.Item("MASTER_BOL") = 0
        rowSOTSHIPB.Item("SHIP_VIA_CODE") = rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty
        Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty)
        rowSOTSHIPB.Item("SHIP_VIA_DESC") = rowSOTSVIA1.Item("SHIP_VIA_DESC") & String.Empty
        rowSOTSHIPB.Item("SHIP_VIA_SCAC") = rowSOTSVIA1.Item("SHIP_VIA_SCAC") & String.Empty
        rowSOTSHIPB.Item("SHIP_TO_NAME") = txtCUST_NAME.Text
        rowSOTSHIPB.Item("SHIP_TO_ADDR1") = txtCUST_ADDR1.Text
        rowSOTSHIPB.Item("SHIP_TO_ADDR2") = txtCUST_ADDR2.Text
        rowSOTSHIPB.Item("SHIP_TO_ADDR3") = String.Empty
        rowSOTSHIPB.Item("SHIP_TO_CITY") = txtCUST_CITY.Text
        rowSOTSHIPB.Item("SHIP_TO_STATE") = txtCUST_STATE.Text
        rowSOTSHIPB.Item("SHIP_TO_ZIP_CODE") = txtCUST_ZIP_CODE.Text
        rowSOTSHIPB.Item("SHIP_TO_COUNTRY") = txtCUST_COUNTRY.Text
        rowSOTSHIPB.Item("SHIP_TO_CONTACT") = txtCUST_CONTACT.Text
        rowSOTSHIPB.Item("SHIP_TO_PHONE") = mdtCUST_PHONE.Value
        'rowSOTSHIPB.Item("FRT_3PY_NAME") = String.Empty
        'rowSOTSHIPB.Item("FRT_3PY_ADDR1") = String.Empty
        'rowSOTSHIPB.Item("FRT_3PY_ADDR2") = String.Empty
        'rowSOTSHIPB.Item("FRT_3PY_ADDR3") = String.Empty
        'rowSOTSHIPB.Item("FRT_3PY_CITY") = String.Empty
        'rowSOTSHIPB.Item("FRT_3PY_STATE") = String.Empty
        'rowSOTSHIPB.Item("FRT_3PY_ZIP_CODE") = String.Empty
        'rowSOTSHIPB.Item("FRT_3PY_COUNTRY") = String.Empty
        'rowSOTSHIPB.Item("FRT_3PY_CONTACT") = String.Empty
        'rowSOTSHIPB.Item("FRT_3PY_PHONE") = String.Empty
        'rowSOTSHIPB.Item("BOL_INST") = String.Empty
        rowSOTSHIPB.Item("THIRD_PARTY") = "0"
        rowSOTSHIPB.Item("SHIP_REF") = rowSOTSHIP1.Item("SHIP_REF") & String.Empty
        'rowSOTSHIPB.Item("SHIP_TRAILER_NO") = String.Empty
        'rowSOTSHIPB.Item("SHIP_SEAL_NO") = String.Empty
        rowSOTSHIPB.Item("BOL_STATUS") = "F"

        If rowSOTSHIPB.Item("INIT_OPER") & String.Empty <> String.Empty Then
            rowSOTSHIPB.Item("INIT_DATE") = DATETIME_STAMP
            rowSOTSHIPB.Item("INIT_OPER") = ASCMAIN1.USER_ID
        End If

        rowSOTSHIPB.Item("LAST_DATE") = DATETIME_STAMP
        rowSOTSHIPB.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowSOTSHIPB.Item("SHIPPED_ACTUAL") = rowSOTSHIP1.Item("SHIP_DATE_SHIPPED")
        'rowSOTSHIPB.Item("SHIP_TO_CODE") = String.Empty
        'rowSOTSHIPB.Item("FRT_3PY_CODE") = String.Empty
        rowSOTSHIPB.Item("BOL_PRINTED") = "1"
        'rowSOTSHIPB.Item("SHIP_LOAD_NO") = String.Empty
        'rowSOTSHIPB.Item("SHIP_APPT_NO") = String.Empty
        'rowSOTSHIPB.Item("SCHED_DELIV_DATE") = String.Empty
        'rowSOTSHIPB.Item("SHIP_FREIGHT") = String.Empty
    End Sub

    Private Sub AddCarton()

        If grdSOTPICK1.ActiveRow Is Nothing Then
            MessageBox.Show("You must select a pick ticket.", "Add Carton", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow
        Dim CART_NO As String = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTCART1.CART_NO"), "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
        rowSOTCART1.Item("CART_NO") = CART_NO ' "NEW" & Format(CART_NO_new, "0000000")
        rowSOTCART1.Item("PICK_NO") = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
        dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)

        rowSOTCART1.Item("CART_TOTAL_UNITS") = 0
        rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = 0
        rowSOTCART1.Item("CART_TOTAL_WGT_CALC") = 0
        rowSOTCART1.Item("CART_TRACKING_NO") = String.Empty
        'rowSOTCART1.Item("CART_SEQ") = String.Empty
        'rowSOTCART1.Item("CART_MEMO") = String.Empty
        'rowSOTCART1.Item("CART_TYPE") = String.Empty
        rowSOTCART1.Item("PACKAGING_TYPE") = 31
        rowSOTCART1.Item("PKG_CODE") = String.Empty
        Dim CART_SEQ As Int32 = Val(dst.Tables("SOTCART1").Compute("MAX(CART_SEQ)", "PICK_NO = '" & grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value & "'") & String.Empty) + 1
        rowSOTCART1.Item("CART_SEQ") = CART_SEQ

        grdSOTPICK1.ActiveRow.Cells("PICK_CNT_CARTONS").Value = dst.Tables("SOTCART1").Select("PICK_NO = '" & rowSOTCART1.Item("PICK_NO") & "'").Length
    End Sub

    Sub Load_SOTSHIPX()
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim sqlSOTSHIPX_LP As String = sqlSOTSHIPX

        If InquiryMode Then
            ASCMAIN1.sql = sqlSOTSHIPX_LP _
                & IIf(CUST_CODE = "", "", " and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'")

            Select Case optStatus.Value
                Case "RNP"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'P'"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_PICK_PRINTED is Null"
                    grdSOTSHIPX.Text = "Shipments Released not Printed"
                Case "PNC"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'P'"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_PICK_PRINTED is Not Null"
                    grdSOTSHIPX.Text = "Shipments Printed not Confirmed"
                Case "C"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'F'"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_DATE_SHIPPED >= '" & Format(calFrom.Value, "dd-MMM-yyyy") & "'"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_DATE_SHIPPED <= '" & Format(calTo.Value, "dd-MMM-yyyy") & "'"
                    grdSOTSHIPX.Text = "Shipments Confirmed as Shipped between " & calFrom.Value & " and " & calTo.Value
            End Select

            If CUST_CODE <> "" Then grdSOTSHIPX.Text &= " associated with " & CUST_CODE
            Fill_Records("SOTSHIPX", "", , ASCMAIN1.sql)

            Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)
        Else
            If ASCMAIN1.CLIENT <> "RGI" Then
                sqlSOTSHIPX_LP &= "  and ICTWHSE1.LP_CODE IS NULL"
            End If

            If CUST_CODE = "" Then
                Fill_Records("SOTSHIPX", "", True, sqlSOTSHIPX_LP)
                grdSOTSHIPX.Text = "Unconfirmed Shipments"
                Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)
            Else
                ASCMAIN1.sql = sqlSOTSHIPX_LP _
                    & " and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'"
                Fill_Records("SOTSHIPX", "", , ASCMAIN1.sql)
                grdSOTSHIPX.Text = "Unconfirmed Shipments associated with " & CUST_CODE
                Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)
            End If
        End If

        grdSOTSHIPX.Visible = True
    End Sub

    Sub Display_Totals()
        Dim KEY As Int32 = 0

        Dim PICK_FREIGHT As Decimal = 0
        Dim MISC_CHG As Decimal = 0
        Dim PICK_AMT_CONF As Decimal = 0

        For Each COL As String In New String() _
            {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK", "PICK_FREIGHT", "MISC_CHG"}
            KEY += 1
            Dim rowSOTCONFT As DataRow = dst.Tables("SOTCONFT").Rows.Find(KEY)

            Select Case COL
                Case "PICK_FREIGHT"
                    rowSOTCONFT.Item("QTY") = Val(dst.Tables("SOTPICK1").Compute("COUNT(" & COL & ")", "SELECTED = '1' AND ISNULL(PICK_FREIGHT, 0) > 0") & "")
                    rowSOTCONFT.Item("AMT") = Val(dst.Tables("SOTPICK1").Compute("SUM(" & Replace(COL, "QTY", "AMT") & ")", "SELECTED = '1'") & "")
                    PICK_FREIGHT = rowSOTCONFT.Item("AMT")

                Case "MISC_CHG"
                    rowSOTCONFT.Item("QTY") = dst.Tables("SOTINVHM").Select("", "", DataViewRowState.CurrentRows).Length
                    rowSOTCONFT.Item("AMT") = Val(dst.Tables("SOTINVHM").Compute("SUM(INV_MISC_CHG)", "") & String.Empty)
                    MISC_CHG = rowSOTCONFT.Item("AMT")

                Case Else
                    rowSOTCONFT.Item("QTY") = Val(dst.Tables("SOTPICK1").Compute("SUM(" & COL & ")", "SELECTED = '1'") & "")
                    rowSOTCONFT.Item("AMT") = Val(dst.Tables("SOTPICK1").Compute("SUM(" & Replace(COL, "QTY", "AMT") & ")", "SELECTED = '1'") & "")

                    If KEY = 2 Then
                        PICK_AMT_CONF = rowSOTCONFT.Item("AMT")
                    End If
            End Select
        Next

        If ASCMAIN1.CLIENT = "RGI" Then
            numInsureValue.Value = PICK_FREIGHT + MISC_CHG + PICK_AMT_CONF
        Else
            numInsureValue.Value = MISC_CHG + PICK_AMT_CONF
        End If

    End Sub

    Function Select_Style(ByRef COLOR_CODE As String) As String

        Dim STYLE_CODE As String = ""

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STYLE_CODE")
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
            Using F As New ASFCODE1
                F.ShowDialog()
            End Using
            STYLE_CODE = ASCMAIN1.CodeSelector.SelectedCode
        End If

        If COLOR_CODE <> "" Then
            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
            If rowICTSTYC1 Is Nothing Then
                MsgBox("Color Code '" & COLOR_CODE & "' is not Associated with Style " & STYLE_CODE)
                STYLE_CODE = ""
            End If
        Else
            ASCMAIN1.sql = "Select * from ICTSTYC1 where STYLE_CODE = :PARM1"
            Dim rows() As DataRow = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {STYLE_CODE}).Select
            If rows.Length = 1 Then
                COLOR_CODE = rows(0).Item("COLOR_CODE")
            Else
                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("COLOR_CODE")
                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False
                    ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                    ASCMAIN1.CodeSelector.SQL = "Select * from (" & ASCMAIN1.CodeSelector.SQL & ")" _
                        & " where COLOR_CODE in " _
                        & " (Select COLOR_CODE from ICTSTYL1 where STYLE_CODE = '" & STYLE_CODE & "')"
                    Using F As New ASFCODE1
                        F.ShowDialog()
                    End Using
                    COLOR_CODE = ASCMAIN1.CodeSelector.SelectedCode
                    If COLOR_CODE = "" Then STYLE_CODE = ""
                End If
            End If
        End If

        Return STYLE_CODE
    End Function

    Private Sub Force_Cartons_to_Balance()
        If dst.Tables("SOTCARTX").Select("PICK_QTY_CONF <> QTY_PACKED").Length <> 0 Then
            If MsgBox("Pick Ticket Details Have Been Found To Be Out Of Balance With Carton Details." & _
                   vbCrLf & "This Update Will Change The Cartons To Force Them to be In Balance With " & _
                   vbCrLf & "The Pick Tickets!" & vbCrLf & _
                   vbCrLf & "Are You SURE This is what you want?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then

                For Each rowSOTCARTX As DataRow In dst.Tables("SOTCARTX").Select("PICK_QTY_CONF <> QTY_PACKED")
                    Dim PICK_QTY_CONF As Int64 = rowSOTCARTX.Item("PICK_QTY_CONF")
                    Dim QTY_PACKED As Int64 = rowSOTCARTX.Item("QTY_PACKED")
                    Dim QTY As Int64 = PICK_QTY_CONF - QTY_PACKED
                    Dim sqlw As String = "ORDR_NO = '" & rowSOTCARTX.Item("ORDR_NO") & "' and ORDR_LNO = " & rowSOTCARTX.Item("ORDR_LNO")
                    For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select(sqlw, "CART_NO DESC")
                        If QTY = 0 Then Exit For
                        QTY_PACKED = Val(rowSOTCART2.Item("QTY_PACKED") & "")
                        If QTY > 0 Then
                            rowSOTCART2.Item("QTY_PACKED") = QTY_PACKED + QTY
                            QTY = 0
                        Else
                            If QTY_PACKED > System.Math.Abs(QTY) Then
                                rowSOTCART2.Item("QTY_PACKED") = QTY_PACKED + QTY
                                QTY = 0
                            Else
                                rowSOTCART2.Item("QTY_PACKED") = 0
                                QTY = QTY + QTY_PACKED
                            End If
                        End If
                    Next
                Next
            End If
        End If
    End Sub

    Private Sub Force_PTs_to_Balance()
        If dst.Tables("SOTCARTX").Select("PICK_QTY_CONF <> QTY_PACKED").Length <> 0 Then
            If MsgBox("Pick Ticket Details Have Been Found To Be Out Of Balance With Carton Details." & _
                   vbCrLf & "This Update Will Change The Pick Tickets To Force Them to be In Balance With " & _
                   vbCrLf & "The Cartons!" & vbCrLf & _
                   vbCrLf & "Are You SURE This is what you want?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then

                Dim dt As New DataTable
                For Each DC As DataColumn In dst.Tables("SOTCARTX").Columns
                    dt.Columns.Add(DC.ColumnName, DC.DataType)
                Next

                For Each rowSOTCARTX As DataRow In dst.Tables("SOTCARTX").Select("PICK_QTY_CONF <> QTY_PACKED")
                    Dim PICK_QTY_CONF As Int64 = rowSOTCARTX.Item("PICK_QTY_CONF")
                    Dim QTY_PACKED As Int64 = rowSOTCARTX.Item("QTY_PACKED")
                    Dim QTY As Int64 = PICK_QTY_CONF - QTY_PACKED
                    If QTY < 0 And "I DON'T KNOW WHY OR WHAT WE ARE DOING HERE - DT IS A TEMP TABLE" = "" Then
                        dt.Rows.Add(rowSOTCARTX.ItemArray)
                    Else
                        Dim sqlw As String = "ORDR_NO = '" & rowSOTCARTX.Item("ORDR_NO") & "' and ORDR_LNO = " & rowSOTCARTX.Item("ORDR_LNO")
                        Dim rowSOTPICK2s() As DataRow = dst.Tables("SOTPICK2").Select(sqlw, "PICK_NO DESC")
                        rowSOTPICK2s(0).Item("PICK_QTY_CONF") = QTY_PACKED
                        rowSOTPICK2s(0).Item("PICK_QTY_CANC") = Val(rowSOTPICK2s(0).Item("PICK_QTY_CANC") & "") + (PICK_QTY_CONF - QTY_PACKED)
                    End If
                Next

                If dt.Rows.Count <> 0 Then
                    Using F As New ASFMSGBF
                        F.Show_grd(dt, Me, "The Following Pick Ticket Lines Were Confirmed Higher Than The Original Qty's Released", "")
                    End Using
                End If
            End If
        End If
    End Sub

    Sub Add_Line(substitute As Boolean)
        If edi_customer And Not substitute Then
            MsgBox("This Function is Not Allowed for EDI Customers", MsgBoxStyle.OkOnly, "Cartons need to be Re-Generated")
            Exit Sub
        End If

        Dim ORDR_NO As String = grdSOTPICK2.ActiveRow.Cells("ORDR_NO").Value

        Dim STYLE_CODE_SUB As String = ""
        Dim SUB_QTY As Int64 = 0
        Dim PICK_UNIT_PRICE As Decimal = 0

        If substitute Then
            If grdSOTPICK2.ActiveRow.Cells("STYLE_CODE_SUB").Value & "" <> "" Then
                STYLE_CODE_SUB = grdSOTPICK2.ActiveRow.Cells("STYLE_CODE_SUB").Value
            Else
                STYLE_CODE_SUB = grdSOTPICK2.ActiveRow.Cells("STYLE_CODE").Value
            End If
            SUB_QTY = Val(grdSOTPICK2.ActiveRow.Cells("PICK_QTY_CONF").Value & "")
            PICK_UNIT_PRICE = Val(grdSOTPICK2.ActiveRow.Cells("PICK_UNIT_PRICE").Value & "")
        End If


        Dim COLOR_CODE As String = ""
        Dim STYLE_CODE As String = Select_Style(COLOR_CODE)
        If STYLE_CODE = "" Then
            Exit Sub
        End If

        Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
        Dim ORDR_LNO As Int32 = Val(dst.Tables("SOTORDR2").Compute("MAX(ORDR_LNO)", " ORDR_NO = '" & ORDR_NO & "'") & "") + 1

        If substitute Then
            grdSOTPICK2.ActiveRow.Cells("PICK_QTY_CANC").Value = grdSOTPICK2.ActiveRow.Cells("PICK_QTY_CONF").Value
            grdSOTPICK2.ActiveRow.Cells("PICK_QTY_CONF").Value = 0
            grdSOTPICK2.ActiveRow.Update()
        End If

        Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").NewRow
        With rowSOTORDR2
            .Item("ORDR_NO") = ORDR_NO
            .Item("ORDR_LNO") = ORDR_LNO
            .Item("STYLE_CODE") = STYLE_CODE
            .Item("COLOR_CODE") = COLOR_CODE
            .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
            .Item("INNER_PACK_QTY") = rowICTSTYL1.Item("INNER_PACK_QTY")
            .Item("STYLE_UOM") = rowICTSTYL1.Item("STYLE_UOM")
            .Item("ORDR_UNIT_PRICE") = 0
            .Item("ORDR_QTY") = 0
            .Item("ORDR_STATUS") = rowSOTORDR1.Item("ORDR_STATUS")
            .Item("ORDR_QTY_ORIG") = 0
            If substitute Then
                .Item("RANGE_STYLE_CODE") = grdSOTPICK2.ActiveRow.Cells("RANGE_STYLE_CODE").Value
                .Item("RANGE_STYLE_LNO") = Val(grdSOTPICK2.ActiveRow.Cells("RANGE_STYLE_LNO").Value)
                .Item("ORDR_UNIT_PRICE") = Val(grdSOTPICK2.ActiveRow.Cells("PICK_UNIT_PRICE").Value)
                .Item("STYLE_CODE_SUB") = STYLE_CODE_SUB

                ' The following lines are required for EDI Order Substitutions
                Stop ' WHY ARE WE NOT LOOKING AT LOCAL DATATABLE?
                Dim rowSOTORDR2_O As DataRow = LookUp("SOTORDR2", New String() {ORDR_NO, grdSOTPICK2.ActiveRow.Cells("ORDR_LNO").Value})
                .Item("EDI_DTL_SEQ") = rowSOTORDR2_O.Item("EDI_DTL_SEQ")
                .Item("EDI_DOC_SEQ_NO") = rowSOTORDR2_O.Item("EDI_DOC_SEQ_NO")
                .Item("CUST_SIZE_CODE") = rowSOTORDR2_O.Item("CUST_SIZE_CODE")
                .Item("CUST_SKU") = rowSOTORDR2_O.Item("CUST_SKU")
            End If
        End With
        dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

        Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").NewRow
        With rowSOTPICK2
            .Item("PICK_NO") = PICK_NO()
            .Item("PICK_LNO") = ORDR_LNO
            .Item("ORDR_NO") = ORDR_NO
            .Item("ORDR_LNO") = ORDR_LNO
            .Item("PICK_QTY") = 0
            .Item("PICK_QTY_CONF") = 0
            .Item("PICK_QTY_CANC") = 0
            .Item("PICK_QTY_BACK") = 0
            .Item("PICK_UNIT_PRICE") = 0
            .Item("PICK_QTY_CANC_REL") = 0
            .Item("PICK_QTY_BACK_REL") = 0

            .Item("STYLE_CODE") = STYLE_CODE
            .Item("COLOR_CODE") = COLOR_CODE
            .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
            .Item("ORDR_UNIT_PRICE") = 0
            If substitute Then
                .Item("RANGE_STYLE_CODE") = grdSOTPICK2.ActiveRow.Cells("RANGE_STYLE_CODE").Value
                .Item("RANGE_STYLE_LNO") = grdSOTPICK2.ActiveRow.Cells("RANGE_STYLE_LNO").Value
                .Item("ORDR_UNIT_PRICE") = grdSOTPICK2.ActiveRow.Cells("PICK_UNIT_PRICE").Value
            End If
            If substitute Then
                .Item("STYLE_CODE_SUB") = STYLE_CODE_SUB
                .Item("PICK_QTY_CONF") = SUB_QTY
                .Item("PICK_UNIT_PRICE") = PICK_UNIT_PRICE
            End If
        End With
        dst.Tables("SOTPICK2").Rows.Add(rowSOTPICK2)
    End Sub

    Sub Setup_BO()
        If chkBO.Checked Then
            cmdBACK.Enabled = True
            With grdSOTPICK2.DisplayLayout.Bands(0).Columns("PICK_QTY_BACK")
                .CellActivation = UltraWinGrid.Activation.AllowEdit
                .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .CellAppearance.BackColor = Drawing.Color.Empty
            End With
        Else
            cmdBACK.Enabled = False
            With grdSOTPICK2.DisplayLayout.Bands(0).Columns("PICK_QTY_BACK")
                .CellActivation = UltraWinGrid.Activation.NoEdit
                .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                .CellAppearance.BackColor = Drawing.Color.Beige
            End With
            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_QTY_BACK <> 0")
                rowSOTPICK2.Item("PICK_QTY_CANC") = Val(rowSOTPICK2.Item("PICK_QTY_CANC") & "") + Val(rowSOTPICK2.Item("PICK_QTY_BACK") & "")
                rowSOTPICK2.Item("PICK_QTY_BACK") = 0
            Next
        End If
    End Sub

    Sub Setup_SOTPICK1()
        If grdSOTPICK1.ActiveRow Is Nothing Then
            tabSOTPICK1.Visible = False
        Else
            Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value & String.Empty
            Dim ORDR_NO As String = grdSOTPICK1.ActiveRow.Cells("ORDR_NO").Value & String.Empty
            Dim CUST_STORE_NO As String = grdSOTPICK1.ActiveRow.Cells("CUST_STORE_NO").Value & String.Empty
            Dim dvw As DataView = DirectCast(grdSOTPICK2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PICK_NO = '" & PICK_NO & "'"
            grdSOTPICK2.Text = "Style Details for Pick No " & PICK_NO & ", Store " & CUST_STORE_NO
            optSCB.ValueList.ValueListItems(2).DisplayText = "Pick Ticket " & PICK_NO
            optSCB.ValueList.ValueListItems(2).Tag = "PICK_NO = '" & PICK_NO & "'"

            dvwSOTORDR5.RowFilter = "CUST_ADDR_TYPE = 'ST' and ORDR_NO = '" & ORDR_NO & "'"

            dvw = DirectCast(grdSOTCART1.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PICK_NO = '" & PICK_NO & "'"
            grdSOTCART1.Text = "Cartons for " & PICK_NO
            Setup_SOTCART2_from_SOTCART1()

            tabSOTPICK1.Visible = True

            If (EntryMode = "N" OrElse EntryMode = "E") Then
                If grdSOTPICK1.ActiveRow.Cells("PICK_STATUS").Value <> "P" Then
                    grdSOTPICK2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grdSOTCART1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grdSOTCART2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                Else
                    grdSOTPICK2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdSOTCART1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdSOTCART2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                End If
            End If
        End If
    End Sub

    Sub Setup_SOTCART2_from_SOTCART1()
        If Not tabSOTPICK1.Tabs("Cartons").Visible Then Exit Sub
        If grdSOTCART1.ActiveRow Is Nothing Then
            grdSOTCART2.Visible = False
        Else
            Dim CART_NO As String = grdSOTCART1.ActiveRow.Cells("CART_NO").Value
            Dim dvw As DataView = DirectCast(grdSOTCART2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "CART_NO = '" & CART_NO & "'"
            grdSOTCART2.Text = "Contents of Carton " & CART_NO
            grdSOTCART2.Visible = True
        End If
    End Sub

    Sub Setup_SOTCART2_from_SOTPICK2()
        If tabSOTPICK1.Tabs("Cartons").Visible Then Exit Sub
        If grdSOTPICK2.ActiveRow Is Nothing Then
            grdSOTCART2.Visible = False
        Else
            Dim PICK_NO As String = grdSOTPICK2.ActiveRow.Cells("PICK_NO").Value
            Dim PICK_LNO As Int32 = Val(grdSOTPICK2.ActiveRow.Cells("PICK_LNO").Value & "")
            Dim dvw As DataView = DirectCast(grdSOTCART2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PICK_NO = '" & PICK_NO & "' and ORDR_LNO = " & CStr(PICK_LNO)
            grdSOTCART2.Text = "Cartons containing Styles Indicated on Pick Ticket " & PICK_NO & ", Line " & CStr(PICK_LNO)
            grdSOTCART2.Visible = True
        End If
    End Sub

    Sub De_Confirm(SHIP_BOL_NO As String)

        Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
        If rowSOTSHIP1 Is Nothing Then
            rowSOTSHIP1 = dst.Tables("SOTSHIPX").Rows.Find(SHIP_BOL_NO)
        End If

        If rowSOTSHIP1 Is Nothing Then
            rowSOTSHIP1 = ASCDATA1.GetDataRow("SELECT * FROM SOTSHIP1 WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
        End If

        If rowSOTSHIP1 Is Nothing Then
            MessageBox.Show("Cannot Locate Shipment BOL " & SHIP_BOL_NO, "Deconfirm Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim ORDR_GROUP_NO As String = rowSOTSHIP1.Item("ORDR_GROUP_NO")

        BeginTrans()
        Dim sqlw As String = "(Select INV_NO from SOTINVH1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
        ASCDATA1.ExecuteSQL("Delete from SOTINVH2 where INV_TYPE = 'I' and INV_NO in " & sqlw)
        ASCDATA1.ExecuteSQL("Delete from SOTINVH1 where INV_TYPE = 'I' and INV_NO in " & sqlw)

        ASCMAIN1.sql = "Update SOTPICK1 Set PICK_STATUS = 'P', PICK_SHIPPED = NULL, INV_NO = NULL" _
            & ", LAST_OPER = '" & ASCMAIN1.USER_ID & "',LAST_DATE = SYSDATE" _
            & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update SOTSHIP1 Set SHIP_STATUS = 'P', SHIP_DATE_SHIPPED = NULL, LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "' where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
        ASCDATA1.ExecuteSQL()

        'Update New Control File To Force SJ&U between Confirm and Deconfirm - WR - 20051024
        ASCMAIN1.sql = "Update SOTCTLU1" _
            & " SET CTL_UPDATE_REQ = 'D'" _
            & " WHERE UPPER(CTL_KEY) = 'Z'"
        ASCDATA1.ExecuteSQL()

        CommitTrans("Shipment " & SHIP_BOL_NO & " has been Successfully De-Confirmed")
    End Sub

    Sub Reverse_Invoice(SHIP_BOL_NO As String, INV_REVERSAL_REASON As String)

        If Not InquiryMode AndAlso Not ",edz,wjz,".Contains(ASCMAIN1.USER_ID) Then
            MessageBox.Show("You are not permitted to Reverse a Shipment.", "Reverse", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ASCMAIN1.sql = sqlSOTSHIPX & vbCrLf _
            & " and SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        Fill_Records("SOTSHIP1", "", True, ASCMAIN1.sql)

        Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)

        If rowSOTSHIP1 Is Nothing Then
            MessageBox.Show("Cannot locate shipment number: " & SHIP_BOL_NO, "Reverse", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim ORDR_GROUP_NO As String = rowSOTSHIP1.Item("ORDR_GROUP_NO")
        Dim REGISTER_XNO As String = rowSOTSHIP1.Item("REGISTER_XNO") & String.Empty

        Dim SHIP_BOL_NOs As New List(Of String)

        ASCMAIN1.sql = "Select SOTSHIP1.*" & vbCrLf _
            & " from SOTSHIP1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
            & " and NVL(REGISTER_XNO, '') = '" & REGISTER_XNO & "'" & vbCrLf _
            & " and SHIP_STATUS = 'F'" & vbCrLf _
            & " and SHIP_BOL_NO_REV IS NULL"
        Dim DT As DataTable = ASCDATA1.GetDataTable
        If DT.Rows.Count > 1 Then
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("SHIP_BOL_NO")
            ASCMAIN1.CodeSelector.MultipleSelections = True
            ASCMAIN1.CodeSelector.UseDataFromTable = DT
            ASCMAIN1.CodeSelector.Caption = "Please Select the Shipments to Reverse"
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                For Each SHIP_BOL_NO In ASCMAIN1.CodeSelector.SelectedCodes
                    SHIP_BOL_NOs.Add(SHIP_BOL_NO)
                Next
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If

            If ASCMAIN1.USER_ID = "angela" OrElse ASCMAIN1.USER_ID = "pat" Then
                If MsgBox("There Are Multiple BOLS In This Confirmation." _
                          & vbCrLf & vbCrLf _
                          & "Please Verify That The Reversal Has Gone Through Correctly" _
                          & vbCrLf _
                          & "By looking In The Customer Inquiry Screen After Completion." _
                          & vbCrLf & vbCrLf _
                          & "Are You Ready To Proceed?", MsgBoxStyle.YesNo, "Multiple BOLs") = MsgBoxResult.No Then
                    Exit Sub
                End If
            Else
                MsgBox("There Are Multiple BOLs In This Confirmation." _
                       & vbCrLf & "Please See Lenora To Proceed.", _
                       MsgBoxStyle.OkOnly, "Multiple BOLs")
                Exit Sub
            End If
            'Stop CMDEXECUTE CHECK ASSUMED THAT THERE WOULD BE ONLY 1 BOL
            'See SHIP0_REV_NOTES.txt in the Misc folder for further instructions.
        Else
            SHIP_BOL_NOs.Add(SHIP_BOL_NO)
        End If

        Try
            If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
            For Each SHIP_BOL_NO In SHIP_BOL_NOs
                If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub
            Next

            BeginTrans()

            EnforceConstraints(False)

            For Each SHIP_BOL_NO In SHIP_BOL_NOs
                Reverse_Invoice_1(SHIP_BOL_NO, dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO), INV_REVERSAL_REASON)
            Next

            ' Group Record
            ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
            ASCDATA1.ExecuteSQL()

            'Update New Control File To Force SJ&U between Confirm and Deconfirm - WR - 20051024
            If ASCMAIN1.CLIENT = "VAN" Then
                ASCMAIN1.sql = "Update SOTCTLU1" _
                    & " SET CTL_UPDATE_REQ = 'D'" _
                    & " WHERE UPPER(CTL_KEY) = 'Z'"
                ASCDATA1.ExecuteSQL()
            End If

            CommitTrans("Shipment Successfully Reversed")

        Catch ex As Exception
            Rollback(ex.Message)
        End Try

        Me.Cursor = Cursors.Default
    End Sub

    Sub Reverse_Invoice_1(SHIP_BOL_NO As String, rowSOTSHIP1 As DataRow, INV_REVERSAL_REASON As String)
        Dim SHIP_BOL_NO_new As String
        Dim REGISTER_XNO As String = rowSOTSHIP1.Item("REGISTER_XNO") & String.Empty
        Dim INV_DATE As Date = rowSOTSHIP1.Item("INV_DATE")

        ' put this routine in a new modal form with its own datalayer
        ' the new form will receive a list of SHIP_BOL_NOs
        ' ask for the reason on that form
        ' put a reverse button on that form
        ' dim rowSOTSHIP1 AS DATAROW = DST.TABLES("SOTSHIP1").ROWS.FIND(SHIP_BOL_NO)

        ' WHY REGISTER XNO IN THE WHERE CLAUSE?
        ' ISNT THIS DATA ALL HERE BY NOW?

        If ASCMAIN1.CLIENT = "VAN" Then
            SHIP_BOL_NO_new = ASCMAIN1.Next_Control_No("SHIP_BOL_NO")
        Else
            SHIP_BOL_NO_new = ASCMAIN1.Next_Control_No("SOTSHIP1.SHIP_BOL_NO")
        End If

        rowSOTSHIP1.Item("SHIP_BOL_NO") = SHIP_BOL_NO_new
        rowSOTSHIP1.Item("SHIP_856_BATCH_NO") = "N"
        rowSOTSHIP1.Item("SHIP_810_BATCH_NO") = "N"
        rowSOTSHIP1.Item("REGISTER_XNO") = ""
        rowSOTSHIP1.Item("SHIP_BOL_NO_REV") = SHIP_BOL_NO

        For Each TABLE_NAME As String In New String() {"ARTOPEN1", "SOTPICK1", "SOTPICK2", "SOTINVH1", "SOTINVH2", "SOTINVH9", "SOTINVHM", "SOTORDR1", "SOTORDR2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        ASCMAIN1.sql = "Select * from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        Fill_Records("SOTPICK1", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = sqlSOTPICK2 & vbCrLf _
            & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
            & "   and SOTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        Fill_Records("SOTPICK2", "", True, ASCMAIN1.sql)

        Dim ORDR_LNO As Int32 = 0

        Dim PICK_QTY_CONF As Int32 = 0
        Dim PICK_QTY_CANC As Int32 = 0
        Dim PICK_QTY_BACK As Int32 = 0

        Dim ORDR_QTY_OPEN As Int32 = 0
        Dim ORDR_QTY_SHIP As Int32 = 0
        Dim ORDR_QTY_CANC As Int32 = 0

        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
            rowSOTPICK1.Item("SHIP_BOL_NO") = SHIP_BOL_NO_new
            rowSOTPICK1.Item("PICK_NO_REV") = rowSOTPICK1.Item("PICK_NO")

            Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO") & String.Empty

            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
            Fill_Records("SOTORDR1", String.Empty, False, ASCMAIN1.sql)
            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)

            ASCMAIN1.sql = "Select SOTORDR2.*, SOTORDR2.ORDR_UNIT_PRICE ORDR_UNIT_PRICE_ORIG from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
            Fill_Records("SOTORDR2", String.Empty, False, ASCMAIN1.sql)

            ' For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")
            rowSOTORDR1.Item("ORDR_STATUS") = "P"
            '  Next

            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "'", "ORDR_LNO")
                ORDR_LNO = rowSOTORDR2.Item("ORDR_LNO")
                For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_LNO = " & ORDR_LNO)
                    Dim PICK_QTY As Integer = Val(rowSOTPICK2.Item("PICK_QTY") & String.Empty)
                    PICK_QTY_CONF = Val(rowSOTPICK2.Item("PICK_QTY_CONF") & String.Empty)
                    PICK_QTY_CANC = Val(rowSOTPICK2.Item("PICK_QTY_CANC") & String.Empty)
                    PICK_QTY_BACK = Val(rowSOTPICK2.Item("PICK_QTY_BACK") & String.Empty)

                    If PICK_QTY_CONF = 0 AndAlso PICK_QTY_CANC = 0 AndAlso PICK_QTY_BACK = 0 AndAlso PICK_QTY = 0 Then
                        Continue For
                    End If

                    rowSOTORDR2.Item("ORDR_STATUS") = "P"

                    If PICK_QTY > 0 Then
                        Dim ORDR_QTY_PICK As Int32 = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty)
                        ORDR_QTY_PICK += (PICK_QTY)
                        rowSOTORDR2.Item("ORDR_QTY_PICK") = ORDR_QTY_PICK
                    End If

                    If PICK_QTY_CONF > 0 Then
                        ORDR_QTY_SHIP = Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & String.Empty)
                        ORDR_QTY_SHIP -= PICK_QTY_CONF
                        rowSOTORDR2.Item("ORDR_QTY_SHIP") = ORDR_QTY_SHIP
                    End If

                    If PICK_QTY_CANC > 0 Then
                        ORDR_QTY_CANC = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & String.Empty)
                        ORDR_QTY_CANC -= PICK_QTY_CANC
                        rowSOTORDR2.Item("ORDR_QTY_CANC") = ORDR_QTY_CANC
                    End If

                    ' IF PICK_QTY_BACK EXISTS, YOU NEED TO TAKE IT BACK OUT OF OPEN
                    'ORDR_QTY_OPEN = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & String.Empty)
                Next
            Next

            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'")
                For Each COLUMN_NAME As String In New String() {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK", "PICK_QTY_CANC_REL", "PICK_QTY_BACK_REL"}
                    rowSOTPICK2.Item(COLUMN_NAME) = -1 * Val(rowSOTPICK2.Item(COLUMN_NAME) & "")
                Next
            Next
        Next

        For Each TABLE_NAME As String In New String() {"SOTINVH1", "SOTINVH2", "SOTINVH9", "SOTINVHM"}
            ASCMAIN1.sql = "Select * from " & TABLE_NAME & " where INV_NO in (Select INV_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
            Fill_Records(TABLE_NAME, String.Empty, True, ASCMAIN1.sql)
        Next
        'ASCMAIN1.sql = "SELECT * FROM SOTINVH1 WHERE INV_NO IN (Select INV_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
        'Fill_Records("SOTINVH1", String.Empty, True, ASCMAIN1.sql)

        'ASCMAIN1.sql = "SELECT * FROM SOTINVH2 WHERE INV_NO IN (Select INV_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
        'Fill_Records("SOTINVH2", String.Empty, True, ASCMAIN1.sql)

        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("")
            rowSOTINVH1.Item("SHIP_BOL_NO") = SHIP_BOL_NO_new
            rowSOTINVH1.Item("INV_NO_REV") = rowSOTINVH1.Item("INV_NO")
            rowSOTINVH1.Item("ORDR_DATE_UPDATED") = DATETIME_STAMP

            rowSOTINVH1.Item("REGISTER_XNO") = DBNull.Value
            rowSOTINVH1.Item("REGISTER_DATE") = DBNull.Value
            rowSOTINVH1.Item("REGISTER_IND") = DBNull.Value

            rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
            rowSOTINVH1.Item("INV_810_BATCH_NO") = DBNull.Value
            For Each COLUMN_NAME As String In New String() _
                {"INV_SALES", "INV_SALES_CURR", "INV_COGS", _
                 "INV_FREIGHT", "INV_FREIGHT_CURR", _
                 "INV_MISC_CHG", "INV_MISC_CHG_CURR", _
                 "INV_TOTAL_AMOUNT", "INV_TOTAL_AMT_CURR", "INV_TOTAL_AMOUNT_CURR", _
                 "GST_TAX", "GST_TAX_CURR", _
                 "INV_STAX", "INV_STAX_CURR"}
                rowSOTINVH1.Item(COLUMN_NAME) = -1 * Val(rowSOTINVH1.Item(COLUMN_NAME) & "")
            Next

            Dim SALES_DIVISION_CODE As String = rowSOTINVH1.Item("SALES_DIVISION_CODE")
            Dim PICK_NO As String = rowSOTINVH1.Item("PICK_NO")
            Dim PICK_NO_new As String = ""
            If ASCMAIN1.CLIENT = "VAN" Then
                PICK_NO_new = ASCMAIN1.Next_Control_No("PICK_NO", 10)
            Else
                PICK_NO_new = ASCMAIN1.Next_Control_No("SOTPICK1.PICK_NO", 10)
            End If
            Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")

            Dim INV_NO_new As String = ""
            If ASCMAIN1.CLIENT = "VAN" Then
                INV_NO_new = ASCMAIN1.Next_Control_No("INV_NO_01")
            Else
                INV_NO_new = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
            End If

            rowSOTINVH1.Item("INV_NO") = INV_NO_new
            'Stop ' do we have a relationship with SOTINVH2 to propagate this change down

            For Each rowSOTINVH2 As DataRow In dst.Tables("SOTINVH2").Select("INV_NO = '" & INV_NO & "'")
                rowSOTINVH2.Item("INV_NO") = INV_NO_new
                rowSOTINVH2.Item("ORDR_YYYYPP_UPDATED") = rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED")
                For Each COLUMN_NAME As String In New String() {"ORDR_QTY_SHIP"}
                    rowSOTINVH2.Item(COLUMN_NAME) = -1 * Val(rowSOTINVH2.Item(COLUMN_NAME) & "")
                Next
            Next

            For Each rowSOTINVH9 As DataRow In dst.Tables("SOTINVH9").Select("INV_NO = '" & INV_NO & "'")
                rowSOTINVH9.Item("INV_NO") = INV_NO_new
                For Each COLUMN_NAME As String In New String() {"RANGE_STYLE_QTY_SHIP", "RANGE_STYLE_PP_QTY_SHIP"}
                    rowSOTINVH9.Item(COLUMN_NAME) = -1 * Val(rowSOTINVH9.Item(COLUMN_NAME) & "")
                Next
            Next

            For Each rowSOTINVHM As DataRow In dst.Tables("SOTINVHM").Select("INV_NO = '" & INV_NO & "'")
                rowSOTINVHM.Item("INV_NO") = INV_NO_new
                For Each COLUMN_NAME As String In New String() {"INV_MISC_CHG"}
                    rowSOTINVHM.Item(COLUMN_NAME) = -1 * Val(rowSOTINVHM.Item(COLUMN_NAME) & "")
                Next
            Next

            rowSOTINVH1.Item("PICK_NO") = PICK_NO_new
            rowSOTINVH1.Item("INV_COMMENT") = INV_REVERSAL_REASON


            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
            ' For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("PICK_NO = '" & PICK_NO & "'")
            rowSOTPICK1.Item("PICK_NO") = PICK_NO_new
            rowSOTPICK1.Item("INV_NO") = INV_NO_new
            '  Next

            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "'")
                rowSOTPICK2.Item("PICK_NO") = PICK_NO_new
            Next

            ASCMAIN1.sql = "Update SOTINVH1 set INV_NO_REV_BY = '" & INV_NO_new & "'" _
                & " where INV_TYPE = 'I' AND INV_NO = '" & INV_NO & "'"
            ASCDATA1.ExecuteSQL()

        Next

        Dim X As New TAC.SOCINVH1(dst)

        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("")
            Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
            Dim INV_TYPE As String = rowSOTINVH1.Item("INV_TYPE")
            X.CreateOpenAR(INV_TYPE, INV_NO, 1)
        Next

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTINVH1").Select("INV_NO_CONS is Not Null"), New String() {"INV_NO_CONS", "SALES_DIVISION_CODE"}).Rows
            Dim SALES_DIVISION_CODE As String = row.Item("SALES_DIVISION_CODE")
            Dim INV_NO As String = row.Item("INV_NO_CONS")
            Dim INV_NO_new As String = ""
            If ASCMAIN1.CLIENT = "VAN" Then
                INV_NO_new = ASCMAIN1.Next_Control_No("INV_NO_01")
            Else
                INV_NO_new = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
            End If
            Dim sqlw As String = "SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "' and INV_NO_CONS = '" & INV_NO & "'"
            For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select(sqlw)
                rowSOTINVH1.Item("INV_NO_CONS") = INV_NO_new
            Next
        Next

        INIT_LAST("SOTSHIP1", False, "SHIP_BOL_NO = '" & SHIP_BOL_NO_new & "'")
        INIT_LAST("SOTPICK1", False, "SHIP_BOL_NO = '" & SHIP_BOL_NO_new & "'")

        For Each tableName As String In New String() {"ARTOPEN1", "SOTSHIP1", "SOTPICK1", "SOTPICK2", "SOTINVH1", "SOTINVH2", "SOTINVH9", "SOTINVHM"}
            dst.Tables(tableName).AcceptChanges()
            For Each row As DataRow In dst.Tables(tableName).Select("")
                row.SetAdded()
            Next
        Next

        For Each TABLE_NAME As String In New String() {"ARTOPEN1", "SOTSHIP1", "SOTPICK1", "SOTPICK2", "SOTINVH1", "SOTINVH2", "SOTINVH9", "SOTINVHM", "SOTORDR1", "SOTORDR2"}
            Update_Record_TDA(TABLE_NAME)
        Next

        ' WE ARE INTENTIONALLY SETTING THE ORIGINAL PICK TICKET AND SHIPMENT BACK TO P
        ' WE ACKNOWLEDGE THAT THE ORIGINAL INVOICE STILL POINTS TO THE SAME PICK TICKET
        ' YET THE ORIGINAL PICK TICKET WILL EVENTUALLY POINT BACK TO THE NEW INVOICE WHEN RE-CONFIRMED
        ' THIS IS TO PRESERVE DOCUMENT NUMBERS AD SHIP_BOL_NOS AND OTHER DATA TIED TO THOSE CONTROL NUMBERS
        ' THAT MIGHT HAVE BEEN RECORDED IN REAL TIME IN THE WAREHOUSE
        ' ALSO NOTE THAT THE REVERSAL DOES NOTHING TO CARTON DATA, AND THIS IS INTENTIONAL
        ' BECAUSE CARTON DATA REPRESENTS PHYSICAL LABELS THAT ARE REALLY ON THE BOXES
        ' AND MAY HAVE BEEN TRANSMITTED TO THE RETAILER

        ASCMAIN1.sql = "Update SOTPICK1 set PICK_STATUS = 'P', PICK_SHIPPED = NULL, INV_NO = NULL" _
            & ", LAST_OPER = '" & ASCMAIN1.USER_ID & "',LAST_DATE = SYSDATE" _
            & " WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTPICK2 set PICK_QTY_CONF = NULL, PICK_QTY_CANC = NULL, PICK_QTY_BACK = NULL" _
          & " WHERE PICK_NO IN (SELECT PICK_NO FROM SOTPICK1 WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "UPDATE SOTSHIP1 SET SHIP_STATUS = 'P'" _
            & ", SHIP_DATE_SHIPPED = NULL" _
            & ", LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" _
            & ", REGISTER_XNO = NULL" _
            & ", SHIP_810_BATCH_NO = NULL, EDI_810_CREATED = NULL" _
            & ", SHIP_856_BATCH_NO = NULL, EDI_856_CREATED = NULL" _
            & ", FACTOR_TRANS_BATCH_LAST = NULL, FACTOR_TRANS_LAST_DATE = NULL" _
            & " WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()

        Dim WHSE_CODE As String = rowSOTSHIP1.Item("WHSE_CODE")
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        Dim WHSE_LOCATOR As Boolean = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")

        Dim sql As String = String.Empty
        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("")
            Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
            Dim INV_TYPE As String = rowSOTINVH1.Item("INV_TYPE")

            ' These 2 stored procedures take the Absolute value of Ship Qty. This will not work here. I need to minus the quantities values on the Invoice to be used.
            ASCMAIN1.sql = "BEGIN SOPSTAT1('" & INV_TYPE & "','" & INV_NO & "'); END;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "BEGIN SOPSTAT2('" & INV_TYPE & "','" & INV_NO & "'); END;"
            ASCDATA1.ExecuteSQL()

            ' When should this be called - Only whses that use Locationss
            If WHSE_LOCATOR Then
                TAC.ICCMAIN1.Update_WHTLOCBX("S", INV_NO)
            End If

            ASCDATA1.ExecuteSP("ARPCUST6_IC", "VV", _
               New Object() {INV_TYPE, INV_NO}, _
               New String() {"INV_TYPE_IN", "INV_NO_IN"})

        Next
    End Sub

    Sub Cancel_Shipment()
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()

        For Each SHIP_BOL_NO As String In SHIP_BOL_NOs
            Dependent_Updates(-1, SHIP_BOL_NO)

            If ASCMAIN1.CLIENT = "VAN" Then
                Dim rowWHT3PLS1 As DataRow = dst.Tables("WHT3PLS1").Rows.Find(SHIP_BOL_NO)
                If rowWHT3PLS1 IsNot Nothing Then
                    rowWHT3PLS1.Delete()
                End If
            End If

            ASCMAIN1.sql = "" _
                & "Begin" _
                & " Declare Cursor C1 is" _
                & "  Select SOTPICK2.* from SOTPICK2 " _
                & "   where SOTPICK2.PICK_NO in (Select PICK_NO from SOTPICK1" _
                & "     where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PICK_STATUS = 'P') for Update;" _
                & " Begin" _
                & "  For R1 in C1 Loop" _
                & "   Update SOTORDR2 " _
                & "    Set ORDR_QTY_CANC = NVL(ORDR_QTY_CANC,0) + R1.PICK_QTY" _
                & "      , ORDR_QTY_PICK = NVL(ORDR_QTY_PICK,0) - R1.PICK_QTY" _
                & "    where ORDR_NO = R1.ORDR_NO and ORDR_LNO = R1.ORDR_LNO;" _
                & "   Update SOTPICK2 Set PICK_QTY_CANC = PICK_QTY, PICK_QTY_CONF = 0 where Current of C1;" _
                & "  End Loop;" _
                & " End;" _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update SOTORDR2 Set ORDR_STATUS = 'C' where ORDR_NO in " _
                & " (Select ORDR_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PICK_STATUS = 'P')" _
                & " and ORDR_STATUS = 'P' and ORDR_QTY_OPEN = 0 and ORDR_QTY_PICK = 0 and ORDR_QTY_CANC <> 0"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Update SOTORDR1 " _
                & " Set ORDR_STATUS = 'C' where ORDR_NO in (" _
                & "Select ORDR_NO from (" _
                & "Select ORDR_NO" _
                & ", SUM (DECODE(ORDR_STATUS,'O',1,0)) O" _
                & ", SUM (DECODE(ORDR_STATUS,'P',1,0)) P" _
                & ", SUM (DECODE(ORDR_STATUS,'C',1,0)) C" _
                & ", SUM (DECODE(ORDR_STATUS,'F',1,0)) F" _
                & " from SOTORDR2 where ORDR_NO in " _
                & "(Select ORDR_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PICK_STATUS = 'P')" _
                & " group by ORDR_NO" _
                & ") where O = 0 and P = 0 and F = 0 and C <> 0)"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update SOTPICK1 Set PICK_STATUS = 'C'" _
                & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PICK_STATUS = 'P'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update SOTSHIP1" _
                & " Set SHIP_STATUS = 'C', SHIP_856_BATCH_NO = 'N', SHIP_810_BATCH_NO = 'N'" _
                & " where SHIP_BOL_NO = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", SHIP_BOL_NO)
            ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
            ASCDATA1.ExecuteSQL()
        Next

        CommitTrans("Shipment has been Cancelled")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Records(SHIP_BOL_NO As String, Optional sDependUpds As Boolean = True)

        If sDependUpds Then
            Dependent_Updates(-1, SHIP_BOL_NO)
        End If

        Dim sqlw As String = "where CART_NO in (" _
            & " Select CART_NO from SOTCART1 where PICK_NO in (" _
            & " Select PICK_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'))"
        ASCDATA1.ExecuteSQL("Delete from SOTCART2 " & sqlw)
        ASCDATA1.ExecuteSQL("Delete from SOTCART1 " & sqlw)

        ASCMAIN1.sql = "Delete from SOTPICK2 where PICK_NO in " _
            & " (Select PICK_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from SOTSHIP1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_ICTSTAT2(STYLE_CODE As String, COLOR_CODE As String, WHSE_CODE As String, QTY As Int64)
        ASCDATA1.ExecuteSP("ICPSTAT2", "VVVNNNNNN", _
                           New Object() {STYLE_CODE, COLOR_CODE, WHSE_CODE, _
                                         0, 0, 0, _
                                         0, QTY, 0}, _
                           New String() {"STYLE_CODE_IN", "COLOR_CODE_IN", "WHSE_CODE_IN", _
                                         "WHSE_QTY_ON_HAND_in", "WHSE_QTY_ON_ORDER_in", "WHSE_QTY_TRAN_in", _
                                         "WHSE_QTY_OPEN_in", "WHSE_QTY_PICK_in", "WHSE_QTY_ALLO_in"})
    End Sub

    Sub Dependent_Updates(S As Integer, SHIP_BOL_NO As String)
        ' If ASCMAIN1.Running_in_VS Then Stop
        Dim PICK_QTY As Int64 = 0
        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text

        ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" _
            & ", SUM (SOTPICK2.PICK_QTY) PICK_QTY" _
            & " from SOTORDR2,SOTPICK2,SOTPICK1" _
            & " where SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" _
            & "   and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" _
            & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" _
            & "   and SOTPICK1.PICK_STATUS = 'P'" _
            & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE"

        For Each rowSOTPICK2X As DataRow In ASCDATA1.GetDataTable.Rows
            Dim STYLE_CODE As String = rowSOTPICK2X.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTPICK2X.Item("COLOR_CODE")
            PICK_QTY = Val(rowSOTPICK2X.Item("PICK_QTY") & "")
            If PICK_QTY <> 0 Then
                Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, S * PICK_QTY)
            End If
        Next
    End Sub

    Sub SCB(COLUMN_NAME As String)
        Dim sqlw As String = ""
        If optSCB.Value = "SHIP_BOL_NO" Then
        ElseIf optSCB.Value = "STYLE_CODE" Then
            ' Stop
            sqlw = optSCB.ValueList.ValueListItems(1).Tag
        ElseIf optSCB.Value = "PICK_NO" Then
            sqlw = optSCB.ValueList.ValueListItems(2).Tag
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Changing Pick Ticket Details Indicated")

        SOTPICK1_Expressions(True)
        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select(sqlw)
            rowSOTPICK2.Item("PICK_QTY_CONF") = 0
            rowSOTPICK2.Item("PICK_QTY_CANC") = 0
            rowSOTPICK2.Item("PICK_QTY_BACK") = 0
            rowSOTPICK2.Item(COLUMN_NAME) = rowSOTPICK2.Item("PICK_QTY")
        Next
        SOTPICK1_Expressions(False)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        grdSOTPICK1.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        Display_Totals()
    End Sub

    Function Load_3PL_Shipment_Details_EDT945T1() As Boolean

        Try
            Load_3PL_Shipment_Details_EDT945T1 = False

            Dim EDI_DOC_SEQ_NO As String
            Dim SHIP_BOL_NO As String = String.Empty
            Dim PICK_NO As String = String.Empty
            Dim pickList As New List(Of String)

            If EDI_DOC_SEQ_NOs.Count = 0 Then
                MessageBox.Show("No EDI 945s found.", "Load", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            dst.Tables("EDT945T1").Rows.Clear()
            dst.Tables("EDT945T2").Rows.Clear()
            dst.Tables("SOTPICK4").Rows.Clear()

            For Each EDI_DOC_SEQ_NO In EDI_DOC_SEQ_NOs
                Fill_Records("EDT945T1", String.Empty, False, "Select * from EDT945T1 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                Fill_Records("EDT945T2", String.Empty, False, "Select * from EDT945T2 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
            Next

            If dst.Tables("EDT945T1").Rows.Count = 0 Then
                MessageBox.Show("Error processing EDI: " & "Cannot locate EDI 945 header data.", "Import EDI", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            If dst.Tables("EDT945T2").Rows.Count = 0 Then
                MessageBox.Show("Error processing EDI: " & "Cannot locate EDI 945 details.", "Import EDI", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            ' Delete original Cartons, New ones created using 945 data
            For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select()
                rowSOTCART1.Delete()
            Next

            For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select()
                rowSOTCART2.Delete()
            Next

            Dim EDI_CARRIER_SCAC_CODE As String = String.Empty
            Dim EDI_ROUTING As String = String.Empty
            Dim rowSOTSVIA1 As DataRow = Nothing

            Dim EDI_TOTAL_ORDR_WEIGHT As Double = 0

            ' Need to create new cartons using edi data
            For Each rowEDT945T1 As DataRow In dst.Tables("EDT945T1").Select("")
                rowEDT945T1.Item("EDI_PROCESS_IND") = "1"
                EDI_TOTAL_ORDR_WEIGHT = Val(rowEDT945T1.Item("EDI_TOTAL_ORDR_WEIGHT") & String.Empty)

                'EDI_CARRIER_SCAC_CODE = String.Empty
                EDI_DOC_SEQ_NO = rowEDT945T1.Item("EDI_DOC_SEQ_NO")
                PICK_NO = rowEDT945T1.Item("EDI_PICK_NO")
                EDI_DOC_SEQ_NOs.Remove(EDI_DOC_SEQ_NO)
                Dim EDI_CART_NO As String = String.Empty
                Dim CART_SEQ As Int32 = 0
                Dim CART_LNO As Int16 = 0

                If rowSOTSVIA1 Is Nothing Then
                    If EDI_CARRIER_SCAC_CODE.Length = 0 Then
                        EDI_CARRIER_SCAC_CODE = (rowEDT945T1.Item("EDI_CARRIER_SCAC_CODE") & String.Empty).ToString.Trim
                        EDI_ROUTING = (rowEDT945T1.Item("EDI_ROUTING") & String.Empty).ToString.Trim
                        Dim tbl As DataTable = ABSolution.ASCDATA1.GetDataTable("Select * from SOTSVIA1 where SHIP_VIA_SCAC = '" & EDI_CARRIER_SCAC_CODE & "'")
                        If EDI_CARRIER_SCAC_CODE.Length > 0 AndAlso tbl.Rows.Count = 1 Then
                            rowSOTSVIA1 = tbl.Rows(0)
                        ElseIf EDI_ROUTING.Length > 0 Then
                            tbl = ABSolution.ASCDATA1.GetDataTable("Select * from SOTSVIA1 where SHIP_VIA_SCAC = '" & EDI_CARRIER_SCAC_CODE & "' and SHIP_VIA_CODE_3PL = '" & EDI_ROUTING & "'")
                            If tbl.Rows.Count = 1 Then
                                rowSOTSVIA1 = tbl.Rows(0)
                            Else
                                rowSOTSVIA1 = Nothing
                            End If
                        End If
                    End If
                End If

                If rowSOTSVIA1 Is Nothing Then
                    MessageBox.Show("Error processing EDI: " & "Invalid SCAC Code: " & EDI_CARRIER_SCAC_CODE & " or EDI Routing Code: " & EDI_ROUTING, "Import EDI", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                End If

                Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Select("PICK_NO = '" & PICK_NO & "'")(0)

                ' Update Shipment Header
                Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Select("SHIP_BOL_NO = '" & rowSOTPICK1.Item("SHIP_BOL_NO") & "'")(0)

                If (rowEDT945T1.Item("EDI_BOL_NO") & String.Empty).ToString.Length > rowSOTSHIP1.Table.Columns("BILL_OF_LADING_NO").MaxLength Then
                    rowSOTSHIP1.Item("BILL_OF_LADING_NO") = (rowEDT945T1.Item("EDI_BOL_NO") & String.Empty).ToString.Substring(0, rowSOTSHIP1.Table.Columns("BILL_OF_LADING_NO").MaxLength).Trim
                Else
                    rowSOTSHIP1.Item("BILL_OF_LADING_NO") = rowEDT945T1.Item("EDI_BOL_NO")
                End If

                rowSOTSHIP1.Item("SHIP_REF") = rowEDT945T1.Item("EDI_SHIPPER_ID_NO")
                rowSOTSHIP1.Item("EDI_LOAD_ID") = rowEDT945T1.Item("EDI_LOAD_ID")
                'rowSOTSHIP1.Item("BTB_BOL_NO") = rowEDT945T1.Item("BTB_BOL_NO")

                Dim EDI_FRT_TERMS As String = (rowEDT945T1.Item("EDI_FRT_TERMS") & String.Empty).ToString.Trim
                Dim FRT_TERMS As String = (rowSOTSHIP1.Item("FRT_TERMS") & String.Empty).ToString.Trim

                Dim warehouseFreightTermDesc As String = String.Empty
                Select Case EDI_FRT_TERMS
                    Case "C"
                        warehouseFreightTermDesc = "Collect"
                    Case "P"
                        warehouseFreightTermDesc = "PPA / PPD"
                    Case "T"
                        warehouseFreightTermDesc = "Third party"
                    Case ""

                    Case Else
                        warehouseFreightTermDesc = "Unknown Code: " & EDI_FRT_TERMS
                End Select

                Select Case EDI_FRT_TERMS
                    Case "C"
                        If FRT_TERMS <> "COL" Then
                            MessageBox.Show("Shipment No (" & rowSOTSHIP1.Item("SHIP_BOL_NO") & ") has a change to the Freight Terms. " _
                                           & "Original Terms: " & FRT_TERMS & ", Warehouse Terms: " & warehouseFreightTermDesc, "Freight Terms")
                        End If

                    Case "P"
                        If Not (FRT_TERMS = "PPA" OrElse FRT_TERMS = "PPD") Then
                            MessageBox.Show("Shipment No (" & rowSOTSHIP1.Item("SHIP_BOL_NO") & ") has a change to the Freight Terms. " _
                                           & "Original Terms: " & FRT_TERMS & ", Warehouse Terms: " & warehouseFreightTermDesc, "Freight Terms")
                        End If

                    Case "T"
                        If FRT_TERMS <> "3PY" Then
                            MessageBox.Show("Shipment No (" & rowSOTSHIP1.Item("SHIP_BOL_NO") & ") has a change to the Freight Terms. " _
                                           & "Original Terms: " & FRT_TERMS & ", Warehouse Terms: " & warehouseFreightTermDesc, "Freight Terms")
                        End If

                    Case ""
                        ' Assume if empty, then skip the check

                    Case Else
                        MessageBox.Show("Shipment No (" & rowSOTSHIP1.Item("SHIP_BOL_NO") & ") has a change to the Freight Terms. " _
                                       & "Original Terms: " & FRT_TERMS & ", Unknown Warehouse Terms: " & warehouseFreightTermDesc, "Freight Terms")

                End Select

                ' Need to add freight Pre Paid and Add
                ', SOTORDR1.ORDR_TYPE_CODE - No freight on Order Type Code XFR
                If rowSOTSHIP1.Item("FRT_TERMS") & String.Empty = "PPA" And rowSOTPICK1.Item("ORDR_TYPE_CODE") & String.Empty <> "XFR" Then
                    rowSOTPICK1.Item("PICK_FREIGHT") = Val(rowSOTPICK1.Item("PICK_FREIGHT") & String.Empty) + Val(rowEDT945T1.Item("EDI_FRT_COST") & String.Empty)
                End If

                If rowSOTSVIA1 IsNot Nothing Then
                    rowSOTSHIP1.Item("SHIP_VIA_CODE") = rowSOTSVIA1.Item("SHIP_VIA_CODE")
                End If

                If IsDate(rowEDT945T1.Item("EDI_SHIPMENT_DATE")) Then
                    rowSOTSHIP1.Item("SHIP_DATE_SHIPPED") = CDate(rowEDT945T1.Item("EDI_SHIPMENT_DATE") & String.Empty).ToShortDateString
                    rowSOTSHIP1.Item("INV_DATE") = CDate(rowEDT945T1.Item("EDI_SHIPMENT_DATE") & String.Empty).ToShortDateString
                    rowSOTSHIP1.Item("SHIPPED_ACTUAL") = CDate(rowEDT945T1.Item("EDI_SHIPMENT_DATE") & String.Empty).ToShortDateString
                Else
                    rowSOTSHIP1.Item("SHIP_DATE_SHIPPED") = CDate(DateTime.Now.ToShortDateString)
                    rowSOTSHIP1.Item("INV_DATE") = CDate(DateTime.Now.ToShortDateString)
                    rowSOTSHIP1.Item("SHIPPED_ACTUAL") = CDate(DateTime.Now.ToShortDateString)
                End If

                rowSOTSHIP1.Item("SHIP_TOTAL_WGT") = EDI_TOTAL_ORDR_WEIGHT

                ' Force controls to show updated values
                MyBase.Absx1.dteFor("SHIP_DATE_SHIPPED").Text = rowSOTSHIP1.Item("SHIP_DATE_SHIPPED")
                MyBase.Absx1.dteFor("INV_DATE").Text = rowSOTSHIP1.Item("INV_DATE") & String.Empty
                MyBase.Absx1.txtFor("BILL_OF_LADING_NO").Text = rowSOTSHIP1.Item("BILL_OF_LADING_NO") & String.Empty
                MyBase.Absx1.txtFor("SHIP_REF").Text = rowSOTSHIP1.Item("SHIP_REF") & String.Empty
                If rowSOTSVIA1 IsNot Nothing Then
                    MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text = rowSOTSVIA1.Item("SHIP_VIA_CODE")
                End If

                ' Need to Explode entries from EDT945T2 using SOTPICK4
                Fill_Records("SOTPICK4", PICK_NO)
                Dim errorMesage As String = String.Empty
                If Not ExplodeSotpick4(EDI_DOC_SEQ_NO, errorMesage) Then
                    MessageBox.Show("Error processing EDI: " & errorMesage, "Import EDI", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                End If

                ' Create Cartons
                For Each rowEDT945T2 As DataRow In dst.Tables("EDT945T2").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' AND ISNULL(EDI_SHIP_QTY, 0) > 0", "EDI_CART_NO")

                    ' See if we need to start aqnew carton
                    If EDI_CART_NO <> rowEDT945T2.Item("EDI_CART_NO") Then
                        EDI_CART_NO = rowEDT945T2.Item("EDI_CART_NO")
                        ASCMAIN1.Progress("Cart No: " & EDI_CART_NO)

                        Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow
                        rowSOTCART1.Item("CART_NO") = EDI_CART_NO
                        'rowSOTCART1.Item("CART_FREIGHT") = ""
                        'rowSOTCART1.Item("CART_PACKER") = ""
                        rowSOTCART1.Item("CART_PACKED") = rowSOTSHIP1.Item("SHIP_DATE_SHIPPED")
                        rowSOTCART1.Item("CART_SHIPPED") = rowSOTSHIP1.Item("SHIP_DATE_SHIPPED")
                        rowSOTCART1.Item("PICK_NO") = PICK_NO
                        rowSOTCART1.Item("CART_TOTAL_UNITS") = 0
                        rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = Val(dst.Tables("EDT945T2").Compute("SUM(EDI_CART_WEIGHT)", "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' AND EDI_CART_NO = '" & EDI_CART_NO & "'") & String.Empty)
                        If rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = 0 Then
                            rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = 1
                        End If
                        rowSOTCART1.Item("CART_TOTAL_WGT_CALC") = rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL")

                        rowSOTCART1.Item("CART_TRACKING_NO") = rowEDT945T2.Item("EDI_SHIPPER_ID_NO")
                        CART_SEQ += 1
                        rowSOTCART1.Item("CART_SEQ") = CART_SEQ
                        'rowSOTCART1.Item("CART_MEMO") = ""
                        'rowSOTCART1.Item("CART_TYPE") = ""
                        rowSOTCART1.Item("PACKAGING_TYPE") = "31"
                        rowSOTCART1.Item("PKG_CODE") = "XX"
                        dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)
                        CART_LNO = 0
                    End If

                    Dim sql As String = "SELECT EDT945T2.*, EDT940O6.SLN_LNO, EDT940O6.EDI_SLN_QTY, EDT940O6.ORDR_LNO" & vbCrLf
                    sql &= " FROM EDT940O1, EDT940O6, EDT945T2 " & vbCrLf
                    sql &= " WHERE EDT940O1.COMPANY_CODE = EDT940O6.COMPANY_CODE" & vbCrLf
                    sql &= " AND EDT945T2.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf
                    sql &= " AND EDT940O1.EDI_OUTBOUND_DOC_NO = EDT940O6.EDI_OUTBOUND_DOC_NO" & vbCrLf
                    sql &= " AND EDT940O6.PICK_LNO = EDT945T2.PICK_LNO" & vbCrLf
                    sql &= " AND EDT940O1.PICK_NO = '" & PICK_NO & "'" & vbCrLf
                    sql &= " AND EDT940O6.PICK_LNO = " & Val(rowEDT945T2.Item("PICK_LNO") & String.Empty) & vbCrLf
                    sql &= " AND EDT945T2.EDI_CART_NO = '" & EDI_CART_NO & "'" & vbCrLf
                    sql &= " AND NVL(EDT940O1.ORDR_STATUS_CODE, 'N') <> 'V'"

                    Dim tblWk As DataTable = ABSolution.ASCDATA1.GetDataTable(sql)

                    If tblWk.Rows.Count = 0 Then
                        Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow
                        rowSOTCART2.Item("CART_NO") = EDI_CART_NO
                        CART_LNO += 1
                        rowSOTCART2.Item("CART_LNO") = CART_LNO

                        ' need to get attributes from pick ticket
                        Dim PICK_LNO As Int16 = Val(rowEDT945T2.Item("PICK_LNO") & "")
                        Dim STYLE_CODE As String = rowEDT945T2.Item("STYLE_CODE")

                        Dim PICK_QTY As Int32 = Val(rowEDT945T2.Item("PICK_QTY") & "")
                        Dim EDI_SHIP_QTY As Int32 = Val(rowEDT945T2.Item("EDI_SHIP_QTY") & "")

                        Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "' AND PICK_LNO = " & PICK_LNO)(0)
                        Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Select("ORDR_NO = '" & rowSOTPICK2.Item("ORDR_NO") & "' AND ORDR_LNO = " & rowSOTPICK2.Item("ORDR_LNO"))(0)

                        If Val(rowSOTPICK2.Item("PICK_QTY") & String.Empty) <> PICK_QTY Then
                            Throw New Exception("Invalid Pick Qty for Pick Ticket: " & PICK_NO & ", Line No: " & PICK_LNO)
                        End If

                        If (rowSOTPICK2.Item("STYLE_CODE") & String.Empty <> STYLE_CODE) AndAlso (rowSOTORDR2.Item("CUST_STYLE_CODE") & String.Empty <> STYLE_CODE) Then
                            Throw New Exception("Invalid Style Code for Pick Ticket: " & PICK_NO & ", Line No: " & PICK_LNO & " " _
                                                & "Pick Ticket Style Code: " & rowSOTPICK2.Item("STYLE_CODE") & ", Customer Style Code: " & rowSOTORDR2.Item("CUST_STYLE_CODE") _
                                                & ", EDI Style Code: " & STYLE_CODE)
                        End If

                        rowSOTCART2.Item("ORDR_NO") = rowSOTPICK2.Item("ORDR_NO")
                        rowSOTCART2.Item("ORDR_LNO") = rowSOTPICK2.Item("ORDR_LNO")

                        rowSOTCART2.Item("QTY_PACKED") = Val(rowSOTCART2.Item("QTY_PACKED") & "") + EDI_SHIP_QTY
                        rowSOTPICK2.Item("PICK_QTY_CONF") = Val(rowSOTPICK2.Item("PICK_QTY_CONF") & "") + EDI_SHIP_QTY
                        rowSOTCART2.Item("STYLE_CODE") = STYLE_CODE
                        rowSOTCART2.Item("COLOR_CODE") = rowSOTORDR2.Item("COLOR_CODE")
                        ' Fields added to the table
                        rowSOTCART2.Item("QTY_PACKED_ORIG") = rowSOTCART2.Item("QTY_PACKED")
                        rowSOTCART2.Item("PICK_NO") = PICK_NO
                        dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)
                    Else

                        Dim prePackQty As Int16 = Val(tblWk.Compute("SUM(EDI_SLN_QTY)", "") & String.Empty)

                        'MSG STAYS ON AFTER CLICKING OK
                        'TAKES TOO LONG TO LOAD
                        'ERROR ON EXIT
                        'NO Refresh

                        If prePackQty = 0 Then
                            Throw New Exception("Invalid Pre Pack Qty for Pick Ticket: " & PICK_NO)
                        End If

                        Dim numPrePacks As Int16 = 0

                        For Each row As DataRow In tblWk.Select("")
                            numPrePacks = Val(row.Item("EDI_SHIP_QTY") & String.Empty) / prePackQty

                            ' need to get attributes from pick ticket
                            Dim PICK_LNO As Int16 = Val(row.Item("PICK_LNO") & "")
                            Dim STYLE_CODE As String = rowEDT945T2.Item("STYLE_CODE")
                            If STYLE_CODE = "22-2110FL" Then Stop
                            Dim EDI_SHIP_QTY As Int32 = Val(row.Item("EDI_SLN_QTY") & "") * numPrePacks
                            Dim ORDR_LNO As Int16 = Val(row.Item("ORDR_LNO") & String.Empty)

                            'Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "' AND PICK_LNO = " & PICK_LNO)(0)
                            Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "' AND ORDR_LNO = " & ORDR_LNO)(0)

                            If Not pickList.Contains(PICK_NO & "/" & ORDR_LNO) Then
                                pickList.Add(PICK_NO & "/" & ORDR_LNO)
                                rowSOTPICK2.Item("PICK_QTY_CONF") = 0
                            End If

                            Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow
                            rowSOTCART2.Item("CART_NO") = EDI_CART_NO
                            CART_LNO += 1
                            rowSOTCART2.Item("CART_LNO") = CART_LNO

                            rowSOTCART2.Item("ORDR_NO") = rowSOTPICK2.Item("ORDR_NO")
                            rowSOTCART2.Item("ORDR_LNO") = rowSOTPICK2.Item("ORDR_LNO")

                            Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Select("ORDR_NO = '" & rowSOTCART2.Item("ORDR_NO") & "' AND ORDR_LNO = " & rowSOTCART2.Item("ORDR_LNO"))(0)

                            rowSOTCART2.Item("QTY_PACKED") = Val(rowSOTCART2.Item("QTY_PACKED") & "") + EDI_SHIP_QTY
                            rowSOTPICK2.Item("PICK_QTY_CONF") = Val(rowSOTPICK2.Item("PICK_QTY_CONF") & "") + EDI_SHIP_QTY
                            rowSOTCART2.Item("STYLE_CODE") = STYLE_CODE
                            rowSOTCART2.Item("COLOR_CODE") = rowSOTORDR2.Item("COLOR_CODE")
                            ' Fields added to the table
                            rowSOTCART2.Item("QTY_PACKED_ORIG") = rowSOTCART2.Item("QTY_PACKED")
                            rowSOTCART2.Item("PICK_NO") = PICK_NO
                            dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)
                        Next
                    End If
                Next
            Next

            For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("", "CART_NO", DataViewRowState.CurrentRows)
                Dim CART_NO As String = rowSOTCART1.Item("CART_NO")
                rowSOTCART1.Item("CART_TOTAL_UNITS") = Val(dst.Tables("SOTCART2").Compute("SUM(QTY_PACKED)", "CART_NO = '" & CART_NO & "'") & String.Empty)
            Next

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("", "PICK_NO", DataViewRowState.CurrentRows)
                PICK_NO = rowSOTPICK1.Item("PICK_NO")
                rowSOTPICK1.Item("PICK_TOTAL_WGT") = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)
                rowSOTPICK1.Item("PICK_CNT_CARTONS") = dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'", "", DataViewRowState.CurrentRows).Length
            Next

            If EDI_DOC_SEQ_NOs.Count > 0 Then
                Throw New Exception("Could not load all EDI 945s for the shipments")
            End If

            Load_Record_Ancillary()

            Load_3PL_Shipment_Details_EDT945T1 = True

            MyBase.Absx1.dteFor("SHIP_DATE_SHIPPED").Enabled = False
            MyBase.Absx1.dteFor("INV_DATE").Enabled = False

        Catch ex As Exception
            MessageBox.Show("Error processing EDI: " & ex.Message, "Import EDI", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Load_3PL_Shipment_Details_EDT945T1 = False
        End Try

    End Function

    Private Function ExplodeSotpick4(ByVal EDI_DOC_SEQ_NO As String, ByRef ErrorMesage As String) As Boolean
        Try

            If 1 = 1 Then
                Return True
            End If

            ' Gather Statistics about Pick Ticket Detail and EDT945T2 details for the Selected Style Code
            If dst.Tables("SOTPICK4").Rows.Count = 0 Then
                Return True
            End If

            Dim PICK_LNO_PPK As Int16 = dst.Tables("SOTPICK4").Rows(0).Item("PICK_LNO_PPK")

            For Each row As DataRow In dst.Tables("SOTPICK4").Select()
                row.Item("PICK_QTY_USED") = 0
            Next

            Dim sqlCriteria As String = "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' AND PICK_LNO = " & PICK_LNO_PPK
            Dim totalShipped As Int32 = Val(dst.Tables("EDT945T2").Compute("SUM(EDI_SHIP_QTY)", sqlCriteria) & String.Empty)
            Dim totalDiffQty As Int32 = Val(dst.Tables("EDT945T2").Compute("SUM(EDI_DIFF_QTY)", sqlCriteria) & String.Empty)

            ' Get carton and quantity data, may not need the SUM but you never know if a carton/product combo will appear on more then one line
            ASCMAIN1.sql = " SELECT EDI_CART_NO, SUM(NVL(EDI_SHIP_QTY, 0)) EDI_SHIP_QTY , SUM(NVL(EDI_DIFF_QTY, 0)) EDI_DIFF_QTY"
            ASCMAIN1.sql &= " FROM EDT945T2"
            ASCMAIN1.sql &= " WHERE EDI_DOC_SEQ_NO = :PARM1 AND PICK_LNO = :PARM2"
            ASCMAIN1.sql &= " GROUP BY EDI_CART_NO"
            Dim tblEDT945T2 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VN", New Object() {EDI_DOC_SEQ_NO, PICK_LNO_PPK})
            tblEDT945T2.Columns("EDI_SHIP_QTY").ReadOnly = False

            Dim tblEDT945T2wk As DataTable = dst.Tables("EDT945T2").Clone
            tblEDT945T2wk.Rows.Clear()
            Dim rowEDT945T2 As DataRow = Nothing

            ' Get a copy of each carton from the original import
            Dim rowEDT945T2wk As DataRow = Nothing
            For Each rowEDT945T2 In tblEDT945T2.Rows
                rowEDT945T2wk = tblEDT945T2wk.NewRow
                rowEDT945T2wk.ItemArray = dst.Tables("EDT945T2").Select("EDI_CART_NO = '" & rowEDT945T2.Item("EDI_CART_NO") & "' and PICK_LNO = " & PICK_LNO_PPK)(0).ItemArray
                tblEDT945T2wk.Rows.Add(rowEDT945T2wk)
            Next

            ' We should destory the original cartons and create new cartons
            For Each rowEDT945T2 In dst.Tables("EDT945T2").Select(sqlCriteria)
                rowEDT945T2.Delete()
            Next
            dst.Tables("EDT945T2").AcceptChanges()

            'Now add in the new cartons using the Pick Ticket Quantities
            Dim EDI_DTL_SEQ As Int16 = Val(dst.Tables("EDT945T2").Compute("MAX(EDI_DTL_SEQ)", "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'") & String.Empty)

            ' Loop through the rows and create the cartons
            Dim EDI_CART_NO As String = String.Empty
            Dim EDI_SHIP_QTY As Int32 = 0
            Dim PICK_QTY As Int32 = 0
            Dim amountToApply As Int32 = 0

            ' always fill the highest quantity first - this way if there are shortages we do not 
            ' have multiple lines with a shipped quantity and a shortage.
            For Each rowSOTPICK4 As DataRow In dst.Tables("SOTPICK4").Select("PICK_QTY > 0", "PICK_QTY")
                ' How much product needs to be placed in a cart
                PICK_QTY = rowSOTPICK4.Item("PICK_QTY")

                For Each rowEDT945T2 In tblEDT945T2.Select("EDI_SHIP_QTY > 0", "EDI_CART_NO")
                    EDI_CART_NO = rowEDT945T2.Item("EDI_CART_NO")
                    EDI_SHIP_QTY = Val(rowEDT945T2.Item("EDI_SHIP_QTY") & String.Empty)

                    ' Determine how much of the Pick Ticket Detail we can use to satisify the the Carton Quantity
                    If PICK_QTY >= EDI_SHIP_QTY Then
                        amountToApply = EDI_SHIP_QTY
                    Else
                        amountToApply = PICK_QTY
                    End If

                    ' Create a carton for the current pick ticket detail using this carton
                    rowEDT945T2wk = dst.Tables("EDT945T2").NewRow
                    rowEDT945T2wk.ItemArray = tblEDT945T2wk.Select("EDI_CART_NO = '" & EDI_CART_NO & "'")(0).ItemArray
                    EDI_DTL_SEQ += 1
                    rowEDT945T2wk.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                    rowEDT945T2wk.Item("PICK_LNO") = rowSOTPICK4.Item("PICK_LNO")
                    rowEDT945T2wk.Item("PICK_QTY") = rowSOTPICK4.Item("PICK_QTY")
                    rowEDT945T2wk.Item("EDI_DIFF_QTY") = DBNull.Value
                    rowEDT945T2wk.Item("EDI_SHIP_QTY") = amountToApply
                    dst.Tables("EDT945T2").Rows.Add(rowEDT945T2wk)

                    ' Reduce the EDI_SHIP_QTY by the amount of pieces paced in the new Pick Ticket Detail Carton
                    rowEDT945T2.Item("EDI_SHIP_QTY") -= amountToApply
                    rowSOTPICK4.Item("PICK_QTY_USED") += amountToApply
                    PICK_QTY -= amountToApply

                    ' Once we satisfied this Pick Lines Quantity then get out and process the next Pick Ticket Detail 
                    If PICK_QTY = 0 Then
                        Exit For
                    End If
                Next

                ' This can happen if there are short ships. Got to a Pick detail Line; however no more cartons with quantities to process
                If PICK_QTY > 0 AndAlso tblEDT945T2.Select("EDI_SHIP_QTY > 0").Length = 0 Then
                    amountToApply = PICK_QTY

                    ' Create a carton for the current pick ticket detail using this any carton
                    rowEDT945T2wk = dst.Tables("EDT945T2").NewRow
                    ' Grab any Carton
                    rowEDT945T2wk.ItemArray = tblEDT945T2wk.Select("")(0).ItemArray
                    EDI_DTL_SEQ += 1

                    ' Cart Number used by TSI when product not shipped
                    rowEDT945T2wk.Item("EDI_CART_NO") = "99999999999999999999"
                    rowEDT945T2wk.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                    rowEDT945T2wk.Item("PICK_LNO") = rowSOTPICK4.Item("PICK_LNO")
                    rowEDT945T2wk.Item("PICK_QTY") = rowSOTPICK4.Item("PICK_QTY")
                    rowEDT945T2wk.Item("EDI_SHIP_QTY") = 0
                    rowEDT945T2wk.Item("EDI_DIFF_QTY") = amountToApply
                    dst.Tables("EDT945T2").Rows.Add(rowEDT945T2wk)

                    rowSOTPICK4.Item("PICK_QTY_USED") += amountToApply
                End If
            Next

            ' If they over shipped then apply to the last carton
            Dim overShipped As Int32 = Val(tblEDT945T2.Compute("SUM(EDI_SHIP_QTY)", "EDI_SHIP_QTY > 0") & String.Empty)
            If overShipped > 0 Then
                rowEDT945T2wk.Item("EDI_SHIP_QTY") += overShipped
            End If

            ' Need to validate the quanties in totalShipped and totalDiffQty are the same going out as when they came in
            sqlCriteria = String.Empty
            For Each rowSOTPICK4 As DataRow In dst.Tables("SOTPICK4").Select("PICK_QTY > 0", "PICK_QTY")
                sqlCriteria &= " OR PICK_LNO = " & rowSOTPICK4.Item("PICK_LNO")
            Next

            sqlCriteria = sqlCriteria.Substring(4).Trim
            sqlCriteria = "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' AND (" & sqlCriteria & ")"

            Dim newtotalShipped As Int32 = Val(dst.Tables("EDT945T2").Compute("SUM(EDI_SHIP_QTY)", sqlCriteria) & String.Empty)
            Dim newtotalDiffQty As Int32 = Val(dst.Tables("EDT945T2").Compute("SUM(EDI_DIFF_QTY)", sqlCriteria) & String.Empty)

            If newtotalShipped <> totalShipped Then
                ErrorMesage = "Could not Explode EDI945T2 (" & EDI_DOC_SEQ_NO & ") and resolve Total Quantity Shipped"
                Return False
            End If

            If newtotalDiffQty <> totalDiffQty Then
                ErrorMesage = "Could not Explode EDI945T2 (" & EDI_DOC_SEQ_NO & ") and resolve Total Quantity Cancelled"
                Return False
            End If

            Return True

        Catch ex As Exception
            ErrorMesage = ex.Message
            Return False
        End Try

    End Function

    Sub ToggleDataTableExpressions(ByVal tf As Boolean)

        With dst.Tables("SOTPICK2")
            .Columns("PICK_AMT").Expression = IIf(Not tf, "", "ISNULL(PICK_UNIT_PRICE,0) * ISNULL(PICK_QTY,0)")
            .Columns("PICK_AMT_CONF").Expression = IIf(Not tf, "", "ISNULL(PICK_UNIT_PRICE,0) * ISNULL(PICK_QTY_CONF,0)")
            .Columns("PICK_AMT_CANC").Expression = IIf(Not tf, "", "ISNULL(PICK_UNIT_PRICE,0) * ISNULL(PICK_QTY_CANC,0)")
            .Columns("PICK_AMT_BACK").Expression = IIf(Not tf, "", "ISNULL(PICK_UNIT_PRICE,0) * ISNULL(PICK_QTY_BACK,0)")
        End With

        With dst.Tables("SOTCARTX")
            .Columns("PICK_QTY_CONF").Expression = IIf(Not tf, "", "SUM(CHILD(SOTCARTX_SOTPICK2).PICK_QTY_CONF)")
            .Columns("QTY_PACKED").Expression = IIf(Not tf, "", "SUM(CHILD(SOTCARTX_SOTCART2).QTY_PACKED)")
        End With

        With dst.Tables("SOTCART1")
            .Columns("CART_TOTAL_UNITS_CALC").Expression = IIf(Not tf, "", "SUM(CHILD(SOTCART1_SOTCART2).QTY_PACKED)")
        End With

        With dst.Tables("SOTPICK1")
            .Columns("PICK_TOTAL_WGT_CALC").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTCART1).CART_TOTAL_WGT_ACTUAL)")
            .Columns("PICK_CNT_CARTONS_CALC").Expression = IIf(Not tf, "", "COUNT(CHILD(SOTPICK1_SOTCART1).CART_NO)")
            .Columns("PICK_TOTAL_UNITS_CALC").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTCART1).CART_TOTAL_UNITS_CALC)")

            .Columns("PICK_QTY").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY)")
            .Columns("PICK_QTY_CONF").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY_CONF)")
            .Columns("PICK_QTY_CANC").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY_CANC)")
            .Columns("PICK_QTY_BACK").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY_BACK)")
            .Columns("PICK_AMT").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_AMT)")
            .Columns("PICK_AMT_CONF").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_AMT_CONF)")
            .Columns("PICK_AMT_CANC").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_AMT_CANC)")
            .Columns("PICK_AMT_BACK").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_AMT_BACK)")
        End With

    End Sub

    Private Function PICK_NO() As Object
        Throw New NotImplementedException
    End Function

    ''' <summary>
    ''' Requests Shipping labels for carriers such as fedex, Ups, USPS
    ''' </summary>
    ''' <param name="ShippingLabels">String list of the labels created. Print these</param>
    ''' <param name="ErrorMessage">Error that occurred when processing request</param>
    ''' <param name="PreScreenForErrorsOnly">Boolen to determine if all requirements are meet. If true only evaluation is done.</param>
    ''' <returns></returns>
    ''' <remarks>Any errors or missing attributes will be returned in the ErrorMessage Parameter</remarks>
    Private Function RequestShippingLabel(ByRef ShippingLabels As List(Of String), ByRef ErrorMessage As String, ByVal PreScreenForErrorsOnly As Boolean) As Boolean

        Dim createCarrierLabels As Boolean = False
        ErrorMessage = String.Empty

        Dim rowSOTSHIP1 As DataRow = Nothing
        Dim rowSOTSVIA1 As DataRow = Nothing
        Dim rowSOTCARR1 As DataRow = Nothing
        Dim rowSOTPICK1 As DataRow = Nothing

        Dim SHIP_VIA_CODE As String = String.Empty
        Dim SHIP_PACKAGE_NO As Int64 = 0
        Dim pkgId As Int64 = 0
        Dim UpsPPAFreightRate As Decimal = 0

        ' RGI ships in the warehouse but invoices in the main office in NYC
        ' If a ship via gets set to UPS then do not request label.
        If ASCMAIN1.CLIENT = "RGI" Then
            Return True
        End If

        ' If shipped from 3PL then no labels
        If select_from_3PL_list Then
            Return True
        End If

        Try
            rowSOTSHIP1 = dst.Tables("SOTSHIP1").Rows.Find(MyBase.Absx1.txtFor("SHIP_BOL_NO").Text)
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
                    ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "N:\")
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

            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", MyBase.Absx1.txtFor("WHSE_CODE").Text)
            If rowICTWHSE1 Is Nothing Then
                ErrorMessage = "Invalid or missing Warehouse"
                Return False
            End If

            txtCUST_ADDR1.Text = txtCUST_ADDR1.Text.Trim
            txtCUST_ADDR2.Text = txtCUST_ADDR2.Text.Trim
            txtCUST_CITY.Text = txtCUST_CITY.Text.Trim
            txtCUST_CONTACT.Text = txtCUST_CONTACT.Text.Trim
            txtCUST_COUNTRY.Text = txtCUST_COUNTRY.Text.Trim
            txtCUST_NAME.Text = txtCUST_NAME.Text.Trim
            txtCUST_STATE.Text = txtCUST_STATE.Text.Trim.ToUpper
            txtCUST_ZIP_CODE.Text = txtCUST_ZIP_CODE.Text.Trim

            If txtCUST_ADDR1.TextLength = 0 AndAlso txtCUST_ADDR2.Text.Length = 0 Then
                ErrorMessage = "Invalid or missing Ship To Street Address"
                Return False
            ElseIf (txtCUST_CITY.TextLength = 0 OrElse txtCUST_STATE.TextLength = 0 OrElse txtCUST_ZIP_CODE.TextLength = 0) _
                    AndAlso txtCUST_COUNTRY.Text.StartsWith("US") Then
                ErrorMessage = "Invalid or missing Ship To City, State, Zip Code"
                Return False
            ElseIf txtCUST_COUNTRY.TextLength = 0 Then
                Dim STATE_CODE As String = txtCUST_STATE.Text
                Dim rowTATSTATE As DataRow = tblTATSTATE.Rows.Find(STATE_CODE)
                If rowTATSTATE IsNot Nothing Then
                    txtCUST_COUNTRY.Text = "US"
                Else
                    ErrorMessage = "Invalid or missing Country Code"
                    Return False
                End If
            End If

            ' look at ship via settings 

            If rowSOTSVIA1 IsNot Nothing Then
                If (rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" OrElse rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" _
                        OrElse optPayor.Value = "C" OrElse optPayor.Value = "P") _
                        AndAlso chkInsureShipment.Checked Then
                    ErrorMessage = "Ship Vias / Shipments classified as Collect or Third party do not permit Insurance"
                    Return False
                End If

                txt3PAccountNo.Text = txt3PAccountNo.Text.Trim
                txt3pCountry.Text = txt3pCountry.Text.Trim.ToUpper
                txt3PZipCode.Text = txt3PZipCode.Text.Trim.PadLeft(5, "0")
                If txt3pCountry.TextLength = 0 OrElse txt3pCountry.Text.StartsWith("US") Then
                    txt3pCountry.Text = "US"
                End If

                Select Case rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty
                    Case "F" ' Fedex
                        If rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" OrElse optPayor.Value = "C" OrElse optPayor.Value = "R" Then
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                'ErrorMessage = "Fedex Collect type Ship Vias require an Account Code, Zip Code and Country Code in the customer master."
                                'Return False
                            End If

                        ElseIf rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" OrElse optPayor.Value = "P" Then
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                ErrorMessage = "Fedex Third Party type Ship Vias require an Account Code, Zip Code and Country Code on Header tab."
                                Return False
                            End If
                        End If

                    Case "U" ' Ups
                        If rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" AndAlso rowSOTSVIA1.Item("REQUIRES_ACCT_NO") & String.Empty <> "1" Then
                            'clsShip.Payor = TPayorTypes.ptConsignee
                        ElseIf rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" OrElse optPayor.Value = "C" OrElse optPayor.Value = "R" Then
                            ' Use the Account Information on Customer Master ARTCUST1
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                ErrorMessage = "UPS Collect type Ship Vias require an Account Code, Zip Code and Country Code in the customer master."
                                Return False
                            End If
                        ElseIf rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" OrElse optPayor.Value = "P" Then
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                ErrorMessage = "UPS Third Party type Ship Vias require an Account Code, Zip Code and Country Code on Header tab."
                                Return False
                            End If
                        End If
                End Select
            End If

            ' If Fedex and Collect must be a Ground delivery
            If rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty = "F" _
                AndAlso CARRIER_PROD_CODE <> "15" _
                AndAlso (rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" OrElse optPayor.Value = "C") Then
                ErrorMessage = "Fedex Collect shipments must ship ground. Choose Recipient payor type on the 'Header Info' tab for non ground shipments."
                tabSOTPICK1.SelectedTab = tabSOTPICK1.Tabs("Header Info")
                Return False
            End If

            If PreScreenForErrorsOnly Then Return True

            '*******************************************************************************

            Dim isInternationalShipment As Boolean = False
            Dim fedexSmartPost As Int16 = 26

            Dim PICK_NO As String = String.Empty
            Dim ORDR_NO As String = String.Empty
            Dim ORDR_NO_WEB As String = String.Empty
            Dim ORDR_CUST_PO As String = String.Empty
            Dim SHIP_BOL_NO As String = Absx1.txtFor("SHIP_BOL_NO").Text

            Dim FRT_TERMS As String = Absx1.txtFor("FRT_TERMS").Text
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
            clsShip.LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

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
            rowWHTSHPC1.Item("SHIP_DATE") = CDate(MyBase.Absx1.dteFor("SHIP_DATE_SHIPPED").Value).ToString("MM/dd/yyyy")
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
            rowWHTSHPC1.Item("INSURED_SHIPMENT") = IIf(MyBase.Absx1.chkFor("INSURED_SHIPMENT").Checked, "1", "0")

            ' Sender Information
            With clsShip.Sender
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
                rowWHTSHPC5.Item("SHIP_RESIDENTIAL") = IIf(.IsResidental, "1", "0")
                rowWHTSHPC5.Item("SHIP_PO_BOX") = IIf(.IsPOBox, "1", "0")
                dst.Tables("WHTSHPC5").Rows.Add(rowWHTSHPC5)

            End With

            ' Recipient
            With clsShip.Recipient
                .FirstName = IIf(txtCUST_CONTACT.TextLength > 0, txtCUST_CONTACT.Text, txtCUST_NAME.Text)
                .MiddleInitial = ""
                .LastName = ""

                .Address1 = txtCUST_ADDR1.Text
                .Address2 = txtCUST_ADDR2.Text
                .City = txtCUST_CITY.Text
                .State = txtCUST_STATE.Text
                .ZipCode = txtCUST_ZIP_CODE.Text
                .CountryCode = txtCUST_COUNTRY.Text.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"

                .Company = txtCUST_NAME.Text
                .Phone = mdtCUST_PHONE.Text

                If .Phone.Trim.Length = 0 Then
                    .Phone = clsShip.Sender.Phone
                End If

                If .Phone.Trim.Length = 0 Then
                    .Phone = "1234567890"
                End If

                .IsResidental = optAddressType.Value = "R"
                .IsPOBox = optAddressType.Value = "P"

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
                rowWHTSHPC5.Item("SHIP_RESIDENTIAL") = IIf(.IsResidental, "1", "0")
                rowWHTSHPC5.Item("SHIP_PO_BOX") = IIf(.IsPOBox, "1", "0")
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

            'For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select("SELECTED = '1'", "PICK_NO")

            Dim cartSequenceNos As List(Of Int16) = New List(Of Int16)

            ' Commodities for international shipments
            clsShip.TotalCustomsValue = 0
            clsShip.CommodityDetailList.Clear()
            Dim COMMODITY_LNO As Int16 = 1
            Dim itemList As List(Of String) = New List(Of String)

            For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select("SELECTED = '1'", "PICK_NO")

                PICK_NO = rowSOTPICK1.Item("PICK_NO") & String.Empty
                ORDR_NO = rowSOTPICK1.Item("ORDR_NO") & String.Empty
                SHIP_BOL_NO = rowSOTPICK1.Item("SHIP_BOL_NO") & String.Empty
                ORDR_CUST_PO = rowSOTPICK1.Item("ORDR_CUST_PO") & String.Empty
                PPA_FREIGHT = 0
                OUR_FREIGHT = 0

                ' Get the Invoice Number now so we can put it on the label
                Dim INV_NO As String = String.Empty

                If ASCMAIN1.CLIENT = "VAN" Then
                    INV_NO = ASCMAIN1.Next_Control_No("INV_NO_01")
                Else
                    INV_NO = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
                End If

                rowSOTPICK1.Item("INV_NO") = INV_NO
                rowSOTPICK1.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO

                ' See if we have cartons setup
                If dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'").Length = 0 Then
                    Continue For
                End If

                ' See if the carton has products
                If dst.Tables("SOTCART2").Select("PICK_NO = '" & PICK_NO & "' AND ISNULL(QTY_PACKED, 0) > 0 ").Length = 0 Then
                    Continue For
                End If

                For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'", "CART_SEQ")
                    ' This is done to place multi pick tickets into one carton
                    Dim CART_SEQ As Int32 = rowSOTCART1.Item("CART_SEQ")
                    If cartSequenceNos.Contains(CART_SEQ) Then
                        Continue For
                    End If
                    cartSequenceNos.Add(CART_SEQ)

                    Dim PACKAGING_TYPE As String = rowSOTCART1.Item("PACKAGING_TYPE") & String.Empty
                    Dim PKG_CODE As String = rowSOTCART1.Item("PKG_CODE") & String.Empty
                    Dim rowWHTPKGM1 As DataRow = LookUp("WHTPKGM1", PKG_CODE)
                    pkgId = CART_SEQ ' (Val(StrReverse(StrReverse(rowSOTCART1.Item("CART_NO").ToString).Substring(0, 8))))

                    Dim shipPackageDetail As New nsoftware.InShip.PackageDetail
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
                                If (rowSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & (rowSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If (rowSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & (rowSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If ORDR_CUST_PO.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & ORDR_CUST_PO
                                    refCount += 1
                                End If

                                If INV_NO.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; IN:" & INV_NO
                                    refCount += 1
                                End If

                                If (rowSOTSHIP1.Item("ORDR_DEPT") & String.Empty).ToString.Trim <> String.Empty AndAlso refCount < 3 Then
                                    reference &= "; DN:" & rowSOTSHIP1.Item("ORDR_DEPT") & String.Empty
                                    refCount += 1
                                End If

                                If reference.Length > 0 Then
                                    reference = reference.Substring(1).Trim
                                End If

                            Case WHCSHIP1.ProviderTypeUPS
                                ' Ups allows up to 2 References
                                If (rowSOTPICK1.Item("CUST_STORE_NO") & String.Empty).ToString.Trim <> String.Empty AndAlso refCount < 2 Then
                                    reference &= "; ST:" & rowSOTPICK1.Item("CUST_STORE_NO") & String.Empty
                                    refCount += 1
                                End If

                                If (rowSOTPICK1.Item("ORDR_CUST_PO") & String.Empty).ToString.Trim <> String.Empty AndAlso refCount < 2 Then
                                    reference &= "; PO:" & rowSOTPICK1.Item("ORDR_CUST_PO") & String.Empty
                                    refCount += 1
                                End If

                                If reference.Length > 0 Then
                                    reference = reference.Substring(1).Trim
                                End If

                        End Select

                        .Reference = reference
                        .Id = pkgId.ToString("D8")

                        If chkInsureShipment.Checked Then
                            .InsuredValue = Val(rowSOTCART1.Item("INSURANCE") & String.Empty)
                            rowWHTSHPC1.Item("INSURED_VALUE") += Val(rowSOTCART1.Item("INSURANCE") & String.Empty)
                        End If

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
                    clsShip.TotalCustomsValue = Val(rowSOTPICK1.Item("PICK_AMT_CONF") & String.Empty)

                    For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select("PICK_NO = '" & PICK_NO & "'")
                        Dim STYLE_CODE As String = rowSOTCART2.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowSOTCART2.Item("COLOR_CODE")

                        'Dim ITEM_CODE As String = STYLE_CODE & Chr(0) & COLOR_CODE

                        If itemList.Contains(STYLE_CODE) Then Continue For

                        itemList.Add(STYLE_CODE)

                        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        ' Just in case a non item is permitted in the shipment
                        If rowICTSTYL1 Is Nothing Then Continue For

                        Dim CommodityDetail As New nsoftware.InShip.CommodityDetail
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

            If numInsureValue.Value > clsShip.TotalCustomsValue Then
                clsShip.TotalCustomsValue = numInsureValue.Value
            End If

            If chkSignature.Checked Then
                clsShip.SignatureRequired = True
            Else
                clsShip.SignatureRequired = False
            End If

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
                    If rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" AndAlso optPayor.Value <> "R" Then
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
                    ElseIf rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" Then
                        clsShip.Payor = TPayorTypes.ptThirdParty
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf optPayor.Value = "C" Then
                        clsShip.Payor = TPayorTypes.ptCollect
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf optPayor.Value = "R" Then
                        clsShip.Payor = TPayorTypes.ptRecipient
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf optPayor.Value = "P" Then
                        clsShip.Payor = TPayorTypes.ptThirdParty
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    End If

                Case "U" ' Ups
                    If rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" AndAlso rowSOTSVIA1.Item("REQUIRES_ACCT_NO") & String.Empty <> "1" Then
                        clsShip.Payor = TPayorTypes.ptConsignee
                    ElseIf rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" Then
                        ' Use the Account Information on Customer Master ARTCUST1
                        clsShip.Payor = TPayorTypes.ptThirdParty
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" Then
                        clsShip.Payor = TPayorTypes.ptThirdParty
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf optPayor.Value <> "O" Then
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
                .EzshipLabelImage = nsoftware.InShip.EzshipLabelImageTypes.itEltron
                .ShippingLabelDirectory = ShippingLabelDirectory
                .ShippingLabelPrefix = SHIP_CNTL_NO
                .ShipDate = dteSHIP_DATE_SHIPPED.DateTime.ToString("yyyy-MM-dd")
                If chkSaturday.Checked Then
                    clsShip.ShipmentSpecialServices = clsShip.ShipmentSpecialServices OrElse &H10000000L
                End If
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
                ErrorMessage &= " " & ex.Message
                Rollback()
            End Try

            ' See if we need to get rates for NYA
            UpsPPAFreightRate = 0
            If ASCMAIN1.CLIENT = "NYA" _
                AndAlso Absx1.txtFor("FRT_TERMS").Text = "PPA" _
                AndAlso (PROVIDER_TYPE = WHCSHIP1.ProviderTypeUPS) _
                AndAlso rowSOTCARR1 IsNot Nothing _
                AndAlso rowSOTCARR1.Item("SHIP_ACCT_NO") & String.Empty <> String.Empty _
                AndAlso rowSOTCARR1.Item("SHIP_ACCT_NO") & String.Empty = txt3PAccountNo.Text & String.Empty Then
                UpsPPAFreightRate = clsShip.GetRates

                If UpsPPAFreightRate = 0 Then
                    If clsShip.LastError.Length > 0 Then
                        ErrorMessage = clsShip.LastError
                    Else
                        ErrorMessage = "UPS Get Rates returned $0.00"
                    End If
                    Return False
                End If
            End If

            If clsShip.RequestLabel() Then

                rowWHTSHPC1.Item("ERROR_MSG") = clsShip.LastError & String.Empty
                rowWHTSHPC1.Item("STATUS") = "P"
                If rowWHTSHPC1 IsNot Nothing AndAlso (rowWHTSHPC1.Item("ERROR_MSG") & String.Empty).ToString.Length > 200 Then
                    rowWHTSHPC1.Item("ERROR_MSG") = rowWHTSHPC1("ERROR_MSG").ToString.Substring(0, 200).Trim
                End If
                rowWHTSHPC1.Item("MASTER_TRACKING_NO") = clsShip.MasterTrackingNumber & String.Empty

                ' Spread UpsPPAFreightRate evenly across all the packages.
                UpsPPAFreightRate /= clsShip.PackageDetailList.Count

                For Each shipPackageDetail As nsoftware.InShip.PackageDetail In clsShip.PackageDetailList
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

                        If PPA_FREIGHT = 0 AndAlso UpsPPAFreightRate > 0 Then
                            PPA_FREIGHT = UpsPPAFreightRate
                        End If

                        PICK_NO = String.Empty
                        rowSOTPICK1 = Nothing

                        ' We may have multi pick tickets in a single carton. This stamps them with the same tracking number
                        ' Spread the Customer Freight Cost and Our freight cost across the Pick Tickets
                        Dim numPickTickets As Int16 = dst.Tables("SOTCART1").Select("CART_SEQ = " & SHIP_PACKAGE_NO).Length

                        If OUR_FREIGHT = 0 AndAlso UpsPPAFreightRate > 0 Then
                            OUR_FREIGHT = UpsPPAFreightRate
                        End If

                        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("CART_SEQ = " & SHIP_PACKAGE_NO)
                            rowSOTCART1.Item("CART_TRACKING_NO") = shipPackageDetail.TrackingNumber & String.Empty
                            PICK_NO = rowSOTCART1.Item("PICK_NO") & String.Empty
                            rowSOTPICK1 = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
                            If Absx1.txtFor("FRT_TERMS").Text = "PPA" AndAlso rowSOTPICK1("ORDR_SOURCE") & String.Empty <> "W" AndAlso AddPPAToExistingFreight Then
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
                    If RecreateLabel Then
                        Update_Record_TDA("SOTCART1")
                    End If
                    CommitTrans()
                Catch ex As Exception
                    ErrorMessage &= " " & ex.Message
                    Rollback()
                End Try

            Else
                ErrorMessage &= " " & clsShip.LastError
                RequestShippingLabel = False
            End If

        Catch ex As Exception
            ErrorMessage &= " " & ex.Message
            RequestShippingLabel = False
        End Try

        ErrorMessage = ErrorMessage.Trim

    End Function

    ''' <summary>
    ''' Requests Shipping labels for carriers such as fedex, Ups, USPS
    ''' </summary>
    ''' <param name="ShippingLabels">String list of the labels created. Print these</param>
    ''' <param name="ErrorMessage">Error that occurred when processing request</param>
    ''' <param name="PreScreenForErrorsOnly">Boolen to determin if all requirements are meet. If true only evaluation is done.
    ''' Any errors or missing attributes will be returned in the ErrorMessage Parameter</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RequestShippingLabel_ORIG(ByRef ShippingLabels As List(Of String), ByRef ErrorMessage As String, ByVal PreScreenForErrorsOnly As Boolean) As Boolean

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
            rowSOTSHIP1 = dst.Tables("SOTSHIP1").Rows.Find(MyBase.Absx1.txtFor("SHIP_BOL_NO").Text)
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

        RequestShippingLabel_ORIG = True
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
                    ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "N:\")
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

            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", MyBase.Absx1.txtFor("WHSE_CODE").Text)
            If rowICTWHSE1 Is Nothing Then
                ErrorMessage = "Invalid or missing Warehouse"
                Return False
            End If

            txtCUST_ADDR1.Text = txtCUST_ADDR1.Text.Trim
            txtCUST_ADDR2.Text = txtCUST_ADDR2.Text.Trim
            txtCUST_CITY.Text = txtCUST_CITY.Text.Trim
            txtCUST_CONTACT.Text = txtCUST_CONTACT.Text.Trim
            txtCUST_COUNTRY.Text = txtCUST_COUNTRY.Text.Trim
            txtCUST_NAME.Text = txtCUST_NAME.Text.Trim
            txtCUST_STATE.Text = txtCUST_STATE.Text.Trim.ToUpper
            txtCUST_ZIP_CODE.Text = txtCUST_ZIP_CODE.Text.Trim

            If txtCUST_ADDR1.TextLength = 0 AndAlso txtCUST_ADDR2.Text.Length = 0 Then
                ErrorMessage = "Invalid or missing Ship To Street Address"
                Return False
            ElseIf txtCUST_CITY.TextLength = 0 OrElse txtCUST_STATE.TextLength = 0 OrElse txtCUST_ZIP_CODE.TextLength = 0 Then
                ErrorMessage = "Invalid or missing Ship To City, State, Zip Code"
                Return False
            ElseIf txtCUST_COUNTRY.TextLength = 0 Then
                Dim STATE_CODE As String = txtCUST_STATE.Text
                Dim rowTATSTATE As DataRow = tblTATSTATE.Rows.Find(STATE_CODE)
                If rowTATSTATE IsNot Nothing Then
                    txtCUST_COUNTRY.Text = "US"
                Else
                    ErrorMessage = "Invalid or missing Country Code"
                    Return False
                End If
            End If

            ' look at ship via settings 

            If rowSOTSVIA1 IsNot Nothing Then
                If ",C,P,".Contains(optPayor.Value) AndAlso chkInsureShipment.Checked Then
                    ErrorMessage = "Ship Vias classified as Collect or Third party do not permit Insurance"
                    Return False
                End If

                txt3PAccountNo.Text = txt3PAccountNo.Text.Trim
                txt3pCountry.Text = txt3pCountry.Text.Trim.ToUpper
                txt3PZipCode.Text = txt3PZipCode.Text.Trim.PadLeft(5, "0")
                If txt3pCountry.TextLength = 0 OrElse txt3pCountry.Text.StartsWith("US") Then
                    txt3pCountry.Text = "US"
                End If

                Select Case rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty
                    Case "F" ' Fedex
                        If optPayor.Value = "C" OrElse optPayor.Value = "R" Then
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                'ErrorMessage = "Fedex Collect type Ship Vias require an Account Code, Zip Code and Country Code in the customer master."
                                'Return False
                            End If

                        ElseIf optPayor.Value = "P" Then
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                ErrorMessage = "Fedex Third Party type Ship Vias require an Account Code, Zip Code and Country Code on Header tab."
                                Return False
                            End If
                        End If

                    Case "U" ' Ups
                        If optPayor.Value = "C" OrElse optPayor.Value = "R" Then
                            ' Use the Account Information on Customer Master ARTCUST1
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                ErrorMessage = "UPS Collect type Ship Vias require an Account Code, Zip Code and Country Code on Header tab."
                                Return False
                            End If
                        ElseIf optPayor.Value = "P" Then
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                ErrorMessage = "UPS Third Party type Ship Vias require an Account Code, Zip Code and Country Code on Header tab."
                                Return False
                            End If
                        End If
                End Select
            End If

            ' If Fedex and Collect must be a Ground delivery
            If rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty = "F" _
                AndAlso CARRIER_PROD_CODE <> "15" AndAlso optPayor.Value = "C" Then
                ErrorMessage = "Fedex Collect shipments must ship ground. Choose Recipient payor type on the 'Header Info' tab for non ground shipments."
                tabSOTPICK1.SelectedTab = tabSOTPICK1.Tabs("Header Info")
                Return False
            End If

            If PreScreenForErrorsOnly Then Return True

            '*******************************************************************************

            Dim isInternationalShipment As Boolean = False
            Dim fedexSmartPost As Int16 = 26

            Dim PICK_NO As String = String.Empty
            Dim ORDR_NO As String = String.Empty
            Dim ORDR_CUST_PO As String = String.Empty
            Dim SHIP_BOL_NO As String = Absx1.txtFor("SHIP_BOL_NO").Text

            Dim FRT_TERMS As String = Absx1.txtFor("FRT_TERMS").Text
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
            clsShip.LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

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
            rowWHTSHPC1.Item("SHIP_DATE") = CDate(MyBase.Absx1.dteFor("SHIP_DATE_SHIPPED").Value).ToString("MM/dd/yyyy")
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
            rowWHTSHPC1.Item("INSURED_SHIPMENT") = IIf(MyBase.Absx1.chkFor("INSURED_SHIPMENT").Checked, "1", "0")

            ' Sender Information
            With clsShip.Sender
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
                rowWHTSHPC5.Item("SHIP_RESIDENTIAL") = IIf(.IsResidental, "1", "0")
                rowWHTSHPC5.Item("SHIP_PO_BOX") = IIf(.IsPOBox, "1", "0")
                dst.Tables("WHTSHPC5").Rows.Add(rowWHTSHPC5)

            End With

            ' Recipient
            With clsShip.Recipient
                .FirstName = IIf(txtCUST_CONTACT.TextLength > 0, txtCUST_CONTACT.Text, txtCUST_NAME.Text)
                .MiddleInitial = ""
                .LastName = ""

                .Address1 = txtCUST_ADDR1.Text
                .Address2 = txtCUST_ADDR2.Text
                .City = txtCUST_CITY.Text
                .State = txtCUST_STATE.Text
                .ZipCode = txtCUST_ZIP_CODE.Text
                .CountryCode = txtCUST_COUNTRY.Text.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"

                .Company = txtCUST_NAME.Text
                .Phone = mdtCUST_PHONE.Text

                If .Phone.Trim.Length = 0 Then
                    .Phone = clsShip.Sender.Phone
                End If

                If .Phone.Trim.Length = 0 Then
                    .Phone = "1234567890"
                End If

                .IsResidental = optAddressType.Value = "R"
                .IsPOBox = optAddressType.Value = "P"

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
                rowWHTSHPC5.Item("SHIP_RESIDENTIAL") = IIf(.IsResidental, "1", "0")
                rowWHTSHPC5.Item("SHIP_PO_BOX") = IIf(.IsPOBox, "1", "0")
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

            'For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select("SELECTED = '1'", "PICK_NO")

            Dim cartSequenceNos As List(Of Int16) = New List(Of Int16)

            ' Commodities for international shipments
            clsShip.TotalCustomsValue = 0
            clsShip.CommodityDetailList.Clear()
            Dim COMMODITY_LNO As Int16 = 1
            Dim itemList As List(Of String) = New List(Of String)

            For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select("SELECTED = '1'", "PICK_NO")

                PICK_NO = rowSOTPICK1.Item("PICK_NO") & String.Empty
                ORDR_NO = rowSOTPICK1.Item("ORDR_NO") & String.Empty
                SHIP_BOL_NO = rowSOTPICK1.Item("SHIP_BOL_NO") & String.Empty
                ORDR_CUST_PO = rowSOTPICK1.Item("ORDR_CUST_PO") & String.Empty
                PPA_FREIGHT = 0
                OUR_FREIGHT = 0

                ' Get the Invoice Number now so we can put it on the label
                Dim INV_NO As String = String.Empty

                If ASCMAIN1.CLIENT = "VAN" Then
                    INV_NO = ASCMAIN1.Next_Control_No("INV_NO_01")
                Else
                    INV_NO = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
                End If

                rowSOTPICK1.Item("INV_NO") = INV_NO
                rowSOTPICK1.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO

                ' See if we have cartons setup
                If dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'").Length = 0 Then
                    Continue For
                End If

                ' See if the carton has products
                If dst.Tables("SOTCART2").Select("PICK_NO = '" & PICK_NO & "' AND ISNULL(QTY_PACKED, 0) > 0 ").Length = 0 Then
                    Continue For
                End If

                For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'", "CART_SEQ")
                    ' This is done to place multi pick tickets into one carton
                    Dim CART_SEQ As Int32 = rowSOTCART1.Item("CART_SEQ")
                    If cartSequenceNos.Contains(CART_SEQ) Then
                        Continue For
                    End If
                    cartSequenceNos.Add(CART_SEQ)

                    Dim PACKAGING_TYPE As String = rowSOTCART1.Item("PACKAGING_TYPE") & String.Empty
                    Dim PKG_CODE As String = rowSOTCART1.Item("PKG_CODE") & String.Empty
                    Dim rowWHTPKGM1 As DataRow = LookUp("WHTPKGM1", PKG_CODE)
                    pkgId = CART_SEQ ' (Val(StrReverse(StrReverse(rowSOTCART1.Item("CART_NO").ToString).Substring(0, 8))))

                    Dim shipPackageDetail As New nsoftware.InShip.PackageDetail
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
                                If (rowSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & (rowSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If (rowSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & (rowSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If ORDR_CUST_PO.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & ORDR_CUST_PO
                                    refCount += 1
                                End If

                                If INV_NO.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; IN:" & INV_NO
                                    refCount += 1
                                End If

                                If (rowSOTSHIP1.Item("ORDR_DEPT") & String.Empty).ToString.Trim <> String.Empty AndAlso refCount < 3 Then
                                    reference &= "; DN:" & rowSOTSHIP1.Item("ORDR_DEPT") & String.Empty
                                    refCount += 1
                                End If

                                If reference.Length > 0 Then
                                    reference = reference.Substring(1).Trim
                                End If

                            Case WHCSHIP1.ProviderTypeUPS
                                ' Ups allows up to 2 References
                                If (rowSOTPICK1.Item("CUST_STORE_NO") & String.Empty).ToString.Trim <> String.Empty AndAlso refCount < 2 Then
                                    reference &= "; ST:" & rowSOTPICK1.Item("CUST_STORE_NO") & String.Empty
                                    refCount += 1
                                End If

                                If (rowSOTPICK1.Item("ORDR_CUST_PO") & String.Empty).ToString.Trim <> String.Empty AndAlso refCount < 2 Then
                                    reference &= "; PO:" & rowSOTPICK1.Item("ORDR_CUST_PO") & String.Empty
                                    refCount += 1
                                End If

                                If reference.Length > 0 Then
                                    reference = reference.Substring(1).Trim
                                End If

                        End Select

                        .Reference = reference
                        .Id = pkgId.ToString("D8")

                        If chkInsureShipment.Checked Then
                            .InsuredValue = Val(rowSOTCART1.Item("INSURANCE") & String.Empty)
                            rowWHTSHPC1.Item("INSURED_VALUE") += Val(rowSOTCART1.Item("INSURANCE") & String.Empty)
                        End If

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
                    clsShip.TotalCustomsValue = Val(rowSOTPICK1.Item("PICK_AMT_CONF") & String.Empty)

                    For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select("PICK_NO = '" & PICK_NO & "'")
                        Dim STYLE_CODE As String = rowSOTCART2.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowSOTCART2.Item("COLOR_CODE")

                        'Dim ITEM_CODE As String = STYLE_CODE & Chr(0) & COLOR_CODE

                        If itemList.Contains(STYLE_CODE) Then Continue For

                        itemList.Add(STYLE_CODE)

                        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        ' Just in case a non item is permitted in the shipment
                        If rowICTSTYL1 Is Nothing Then Continue For

                        Dim CommodityDetail As New nsoftware.InShip.CommodityDetail
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

            If numInsureValue.Value > clsShip.TotalCustomsValue Then
                clsShip.TotalCustomsValue = numInsureValue.Value
            End If

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
                    If optPayor.Value = "C" Then
                        ' Use the Account Information on Customer Master ARTCUST1
                        clsShip.Payor = TPayorTypes.ptCollect
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf optPayor.Value = "P" Then
                        clsShip.Payor = TPayorTypes.ptThirdParty
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf optPayor.Value = "R" Then
                        clsShip.Payor = TPayorTypes.ptRecipient
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    End If

                Case "U" ' Ups
                    If optPayor.Value = "C" Then
                        ' Use the Account Information on Customer Master ARTCUST1
                        clsShip.Payor = TPayorTypes.ptThirdParty
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf optPayor.Value = "P" Then
                        clsShip.Payor = TPayorTypes.ptThirdParty
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf optPayor.Value = "R" Then
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
                .EzshipLabelImage = nsoftware.InShip.EzshipLabelImageTypes.itEltron
                .ShippingLabelDirectory = ShippingLabelDirectory
                .ShippingLabelPrefix = SHIP_CNTL_NO
                .ShipDate = dteSHIP_DATE_SHIPPED.DateTime.ToString("yyyy-MM-dd")
                If chkSaturday.Checked Then
                    clsShip.ShipmentSpecialServices = clsShip.ShipmentSpecialServices OrElse &H10000000L
                End If
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
                ErrorMessage &= " " & ex.Message
                Rollback()
            End Try


            If clsShip.RequestLabel() Then

                rowWHTSHPC1.Item("ERROR_MSG") = clsShip.LastError & String.Empty
                rowWHTSHPC1.Item("STATUS") = "P"
                If rowWHTSHPC1 IsNot Nothing AndAlso (rowWHTSHPC1.Item("ERROR_MSG") & String.Empty).ToString.Length > 200 Then
                    rowWHTSHPC1.Item("ERROR_MSG") = rowWHTSHPC1("ERROR_MSG").ToString.Substring(0, 200).Trim
                End If
                rowWHTSHPC1.Item("MASTER_TRACKING_NO") = clsShip.MasterTrackingNumber & String.Empty

                For Each shipPackageDetail As nsoftware.InShip.PackageDetail In clsShip.PackageDetailList
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
                            PICK_NO = rowSOTCART1.Item("PICK_NO") & String.Empty
                            rowSOTPICK1 = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
                            If Absx1.txtFor("FRT_TERMS").Text = "PPA" AndAlso rowSOTPICK1("ORDR_SOURCE") & String.Empty <> "W" Then
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
                    If RecreateLabel Then
                        Update_Record_TDA("SOTCART1")
                    End If
                    CommitTrans()
                Catch ex As Exception
                    ErrorMessage &= " " & ex.Message
                    Rollback()
                End Try

            Else
                ErrorMessage &= " " & clsShip.LastError
                RequestShippingLabel_ORIG = False
            End If

        Catch ex As Exception
            ErrorMessage &= " " & ex.Message
            RequestShippingLabel_ORIG = False
        End Try

        ErrorMessage = ErrorMessage.Trim

    End Function

    ''' <summary>
    ''' Sends data to the Label Printer
    ''' </summary>
    ''' <param name="LabelData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PrintLabel(ByVal LabelData As String) As Boolean

        Try
            If (ASCMAIN1.USER_ID = "edz" OrElse ASCMAIN1.USER_ID = "wjz") AndAlso ASCMAIN1.Running_in_VS Then
                ' Find Zebra printer

                Dim zebraPrinter As String = FindZebraPrinter()

                Dim vLabelPrinter As New ASCPRINT
                Return vLabelPrinter.SendStringToPrinter(zebraPrinter, LabelData)
            End If

            ASCMAIN1.LabelPrinterSerialPort.WriteLine(LabelData)

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
            If printerName.ToUpper.StartsWith("ZP550") Then
                Return printerName
            End If
        Next printerName

        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZP 550") Then
                Return printerName
            End If
        Next printerName

        Return ""
    End Function

    Private Function ProcessCreditCardAuthorization(ByVal AUTH_CCPA_NO As String, _
                                                    ByVal ChargeAmount As Double, _
                                                    ByVal freightAmount As Decimal, _
                                                    ByVal salesTax As Decimal, _
                                                    ByRef ResponseText As String) As String

        Dim sql As String = String.Empty
        Dim ORDR_UNIT_PRICE As Decimal = 0

        AUTH_CCPA_NO = AUTH_CCPA_NO.Trim
        If AUTH_CCPA_NO.Length = 0 Then Return String.Empty

        If ChargeAmount <= 0 Then Return String.Empty
        ChargeAmount = Math.Round(ChargeAmount, 2)

        ASCMAIN1.Progress("Processing Credit Card", String.Empty)

        MyBase.Fill_Records("ARTCCPA1", AUTH_CCPA_NO)
        'clsASCSCRTY.EncryptDecrypt(String.Empty, TAC.ASCSCRTY.Encryption.Decrypt, ASCMAIN1.EncryptionKey, dst.Tables("ARTCCPA1"))

        MyBase.Fill_Records("ARTCCPDA", AUTH_CCPA_NO)
        'clsASCSCRTY.EncryptDecrypt(String.Empty, TAC.ASCSCRTY.Encryption.Decrypt, ASCMAIN1.EncryptionKey, dst.Tables("ARTCCPDA"))

        If dst.Tables("ARTCCPA1").Rows.Count <> 1 Then
            ResponseText = "Cannot locate Credit Card Trans (" & AUTH_CCPA_NO & ")"
            Return String.Empty
        End If

        ' Changed on 11/19/2013
        If dst.Tables("ARTCCPDA").Rows.Count <> 1 Then
            ResponseText = "Cannot locate Credit Card Trans (" & AUTH_CCPA_NO & ") in ARTCCPDA. This data is needed to settle the payment."
            Return String.Empty
        End If

        Dim rowARTCCPA1_AUTH As DataRow = dst.Tables("ARTCCPA1").Rows(0)

        Dim AUTH_RESPONSE_APPROVAL_CODE As String = (rowARTCCPA1_AUTH.Item("RESPONSE_APPROVAL_CODE") & String.Empty).ToString.Trim
        If AUTH_RESPONSE_APPROVAL_CODE.Length = 0 Then
            ResponseText = "Missing Response Approval code. This data is needed to settle the payment."
            Return String.Empty
        End If

        Dim CCPA_NO As String = String.Empty

        Try
            Me.CreditCardProcessor = New TAC.TAFCARDF(Me)

            With Me.CreditCardProcessor
                .ORDR_NO = rowARTCCPA1_AUTH.Item("ORDR_NO") & String.Empty

                .objCCProcessor.TransactionNumber = ASCMAIN1.Next_Control_No("ARTCCPA1.TRANS_NUM")
                .objCCProcessor.TransactionAmount = ChargeAmount
                .objCCProcessor.CustomerCreditCard.CardNumber = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_NO") & String.Empty

                Dim CUST_CREDIT_CARD_EXP_DATE As String = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_EXP_DATE") & String.Empty
                CUST_CREDIT_CARD_EXP_DATE = CUST_CREDIT_CARD_EXP_DATE.PadRight(4, "0")
                .objCCProcessor.CustomerCreditCard.CardExpMonth = CUST_CREDIT_CARD_EXP_DATE.Substring(0, 2)
                .objCCProcessor.CustomerCreditCard.CardExpYear = CUST_CREDIT_CARD_EXP_DATE.Substring(2)
                .objCCProcessor.ValidateCard()

                With .objCCProcessor.Level2Data
                    .Clear()
                    .CardType = CreditCardProcessor.objCCProcessor.CreditCardType

                    Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", rowARTCCPA1_AUTH.Item("ORDR_NO")) ' dst.Tables("SOTORDR1").Rows(0)

                    If dst.Tables("SOTORDR5").Select("CUST_ADDR_TYPE = 'ST'").Length > 0 Then
                        Dim rowSHIPTO As DataRow = dst.Tables("SOTORDR5").Select("CUST_ADDR_TYPE = 'ST'")(0)
                        If rowSHIPTO IsNot Nothing Then
                            .DestinationZip = rowSHIPTO.Item("CUST_ZIP_CODE") & String.Empty
                            .DestinationState = rowSHIPTO.Item("CUST_STATE") & String.Empty
                        End If
                    End If

                    .DiscountAmount = 0
                    .FreightAmount = freightAmount
                    .InvoiceNumber = rowSOTORDR1.Item("ORDR_NO") & String.Empty
                    .OrderDate = rowSOTORDR1.Item("ORDR_DATE") & String.Empty
                    .PurchaseIdentifier = rowSOTORDR1.Item("ORDR_NO") & String.Empty
                    .TaxAmount = salesTax
                    '.MerchantTaxId = CreditCardProcessor.objCCProcessor.MerchantAccount.MerchantTaxID

                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                    If rowICTWHSE1 IsNot Nothing Then
                        .ShipFromZip = rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty
                    End If

                End With

                With .objCCProcessor.Level3Data
                    .Clear()
                    Dim STYLE_CODE As String = String.Empty
                    Dim Quantity As Integer = 0
                    Dim Description As String = String.Empty

                    For Each rowSOTPICK2 As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTPICK2"), New String() {"STYLE_CODE"}).Rows
                        STYLE_CODE = rowSOTPICK2.Item("STYLE_CODE") & String.Empty
                        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        Dim rowSOTPICK2X As DataRow = dst.Tables("SOTPICK2").Select("STYLE_CODE = '" & STYLE_CODE & "'", "PICK_UNIT_PRICE DESC")(0)

                        ORDR_UNIT_PRICE = Val(rowSOTPICK2X.Item("PICK_UNIT_PRICE") & String.Empty)
                        Quantity = Val(dst.Tables("SOTPICK2").Compute("SUM(PICK_QTY_CONF)", "STYLE_CODE = '" & STYLE_CODE & "' AND PICK_UNIT_PRICE = " & ORDR_UNIT_PRICE) & String.Empty)
                        If Quantity <= 0 Then Continue For

                        Dim level3 As New TAC.ARCCCARD.Level3
                        With level3
                            .Description = StrConv(rowICTSTYL1.Item("STYLE_DESC") & String.Empty, VbStrConv.ProperCase)
                            .DiscountAmount = 0
                            .ProductCode = STYLE_CODE
                            .Quantity = Quantity
                            .TaxAmount = 0
                            .TaxType = TAC.ARCCCARD.TaxTypes.StateSalesTax
                            .UnitCost = ORDR_UNIT_PRICE
                            .Units = "each"
                            .Total = .Quantity * .UnitCost
                            .TaxAmount = Math.Round(.Total * .TaxRate / 100, 2, MidpointRounding.AwayFromZero)
                        End With
                        .Add(level3)
                    Next
                End With

                .rowARTCCPA1 = rowARTCCPA1_AUTH
                CCPA_NO = .CC_Capture(ChargeAmount)

                ResponseText = .responseErrorMessage
            End With

        Catch ex As Exception
            MessageBox.Show("The following error occurred processing a credit card: " & ex.Message, "Charge Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            CreditCardProcessor.Dispose()
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

        Return CCPA_NO

    End Function

    Private Sub PrintManifest(ByVal SHIP_BOL_NO As String)

        MyBase.EnforceConstraints(False)
        ASCMAIN1.sql = sqlSOTSHIPX & vbCrLf _
            & " and SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        Fill_Records("SOTSHIP1", "", True, ASCMAIN1.sql)

        If dst.Tables("SOTSHIP1").Rows.Count = 0 Then
            MessageBox.Show("Cannot locate the selected Shipmnet.")
            Exit Sub
        End If

        Dim WHSE_CODE As String = dst.Tables("SOTSHIP1").Rows(0).Item("WHSE_CODE") & String.Empty
        Fill_Records("ICTWHSE1", WHSE_CODE)

        If dst.Tables("ICTWHSE1").Rows.Count = 0 Then
            MessageBox.Show("Cannot locate the Warehouse for the selected Shipmnet.")
            Exit Sub
        End If

        Dim sql As String = "Select * from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        Fill_Records("SOTPICK1", String.Empty, True, sql)

        If dst.Tables("SOTPICK1").Rows.Count = 0 Then
            MessageBox.Show("Cannot locate the Pick Tickets for the selected Shipmnet.")
            Exit Sub
        End If

        Dim BILL_OF_LADING_NO As String = dst.Tables("SOTSHIP1").Rows(0).Item("BILL_OF_LADING_NO") & String.Empty
        BILL_OF_LADING_NO = BILL_OF_LADING_NO.Trim
        If BILL_OF_LADING_NO.Length > 0 Then
            Fill_Records("SOTSHIPB", BILL_OF_LADING_NO)
        End If

        If BILL_OF_LADING_NO.Length > 0 AndAlso dst.Tables("SOTSHIPB").Rows.Count > 0 Then

            ' Load all the shipments and Pick Tickets for this BOL
            sql = "SELECT * FROM SOTSHIP1 WHERE BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "'"
            Fill_Records("SOTSHIP1", "", True, sql)

            sql = "Select * from SOTPICK1 where SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM SOTSHIP1 WHERE BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "')"
            Fill_Records("SOTPICK1", String.Empty, True, sql)
        Else
            ' Create a Temp SOTSHIPB RECORD

            Dim rowSOTSHIPB As DataRow = dst.Tables("SOTSHIPB").NewRow
            Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Select("", "SHIP_BOL_NO")(0)

            'Dim BOL_NO As String = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTSHIPB.BOL_NO"), "0" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))

            ' Create a bogus BOL NO
            Dim BOL_NO As String = "0" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID") & "9" & rowSOTSHIP1.Item("SHIP_BOL_NO").ToString.Substring(2)

            If BILL_OF_LADING_NO.Length > 0 Then
                BOL_NO = BILL_OF_LADING_NO
            End If

            rowSOTSHIPB.Item("BOL_NO") = BOL_NO
            rowSOTSHIPB.Item("CUST_CODE") = rowSOTSHIP1.Item("CUST_CODE")
            rowSOTSHIPB.Item("BOL_DATE") = rowSOTSHIP1.Item("SHIP_DATE_SHIPPED")
            rowSOTSHIPB.Item("FRT_TERMS") = rowSOTSHIP1.Item("FRT_TERMS")
            rowSOTSHIPB.Item("WHSE_CODE") = rowSOTSHIP1.Item("WHSE_CODE")
            rowSOTSHIPB.Item("MASTER_BOL_NO") = String.Empty
            rowSOTSHIPB.Item("MASTER_BOL") = String.Empty
            rowSOTSHIPB.Item("SHIP_VIA_CODE") = rowSOTSHIP1.Item("SHIP_VIA_CODE")

            Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty, True)
            If rowSOTSVIA1 IsNot Nothing Then
                rowSOTSHIPB.Item("SHIP_VIA_DESC") = rowSOTSVIA1.Item("SHIP_VIA_DESC")
                rowSOTSHIPB.Item("SHIP_VIA_SCAC") = rowSOTSVIA1.Item("SHIP_VIA_SCAC")
            End If

            Dim ORDR_NO As String = dst.Tables("SOTPICK1").Rows(0).Item("ORDR_NO") & String.Empty
            Dim rowSOTORDR5 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR5 WHERE ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'ST'")

            If rowSOTORDR5 IsNot Nothing Then
                rowSOTSHIPB.Item("SHIP_TO_NAME") = rowSOTORDR5.Item("CUST_NAME") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_ADDR1") = rowSOTORDR5.Item("CUST_ADDR1") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_ADDR2") = rowSOTORDR5.Item("CUST_ADDR2") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_ADDR3") = String.Empty
                rowSOTSHIPB.Item("SHIP_TO_CITY") = rowSOTORDR5.Item("CUST_CITY") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_STATE") = rowSOTORDR5.Item("CUST_STATE") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_ZIP_CODE") = rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_COUNTRY") = rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_CONTACT") = rowSOTORDR5.Item("CUST_CONTACT") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_PHONE") = rowSOTORDR5.Item("CUST_PHONE") & String.Empty

                For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Rows
                    For Each fieldname As String In New String() {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", _
                                                         "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY", _
                                                         "CUST_CONTACT", "CUST_PHONE"}
                        rowSOTPICK1.Item(fieldname) = rowSOTORDR5.Item(fieldname)
                    Next
                Next
            End If

            rowSOTSHIPB.Item("FRT_3PY_NAME") = String.Empty
            rowSOTSHIPB.Item("FRT_3PY_ADDR1") = String.Empty
            rowSOTSHIPB.Item("FRT_3PY_ADDR2") = String.Empty
            rowSOTSHIPB.Item("FRT_3PY_ADDR3") = String.Empty
            rowSOTSHIPB.Item("FRT_3PY_CITY") = String.Empty
            rowSOTSHIPB.Item("FRT_3PY_STATE") = String.Empty
            rowSOTSHIPB.Item("FRT_3PY_ZIP_CODE") = String.Empty
            rowSOTSHIPB.Item("FRT_3PY_COUNTRY") = String.Empty
            rowSOTSHIPB.Item("FRT_3PY_CONTACT") = String.Empty
            rowSOTSHIPB.Item("FRT_3PY_PHONE") = String.Empty

            rowSOTSHIPB.Item("BOL_INST") = String.Empty
            rowSOTSHIPB.Item("THIRD_PARTY") = String.Empty
            rowSOTSHIPB.Item("SHIP_REF") = rowSOTSHIP1.Item("SHIP_REF")
            rowSOTSHIPB.Item("SHIP_TRAILER_NO") = String.Empty
            rowSOTSHIPB.Item("SHIP_SEAL_NO") = String.Empty
            rowSOTSHIPB.Item("BOL_STATUS") = "F"
            rowSOTSHIPB.Item("INIT_DATE") = DATETIME_STAMP
            rowSOTSHIPB.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowSOTSHIPB.Item("LAST_DATE") = DATETIME_STAMP
            rowSOTSHIPB.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowSOTSHIPB.Item("SHIPPED_ACTUAL") = rowSOTSHIP1.Item("SHIP_DATE_SHIPPED")
            rowSOTSHIPB.Item("SHIP_TO_CODE") = String.Empty
            rowSOTSHIPB.Item("FRT_3PY_CODE") = String.Empty
            rowSOTSHIPB.Item("BOL_PRINTED") = rowSOTSHIP1.Item("SHIP_DATE_SHIPPED")
            rowSOTSHIPB.Item("SHIP_LOAD_NO") = String.Empty
            rowSOTSHIPB.Item("SHIP_APPT_NO") = String.Empty
            rowSOTSHIPB.Item("SCHED_DELIV_DATE") = rowSOTSHIP1.Item("SHIP_DATE_SHIPPED")
            rowSOTSHIPB.Item("SHIP_FREIGHT") = 0

            dst.Tables("SOTSHIPB").Rows.Add(rowSOTSHIPB)

            ' Update SOTSHIP1 values
            For Each rowSOTSHIP1x As DataRow In dst.Tables("SOTSHIP1").Select("")
                rowSOTSHIP1x.Item("BILL_OF_LADING_NO") = BOL_NO
            Next
        End If

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Dim RPT As String = "SORSHIPB"
        Generate_Report(RPT, "Bill of Lading")

        ' Need to make sure PT's have weights
        'SOTCART1.CART_TOTAL_WGT_ACTUAL, SOTPICK1.PICK_TOTAL_WGT
        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("ISNULL(PICK_TOTAL_WGT,0) = 0")
            Dim PICK_TOTAL_WGT As Decimal = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'") & String.Empty)
            rowSOTPICK1.Item("PICK_TOTAL_WGT") = PICK_TOTAL_WGT
        Next

        CR_params.Add("SUBT", "")
        RPT = "SORSHIPM"
        Generate_Report(RPT, "Consolidation Manifest")

        Print_Report_End()

        For Each table As String In New String() {"SOTSHIP1", "ICTWHSE1", "SOTPICK1", "SOTSHIPB"}
            dst.Tables(table).Rows.Clear()
        Next

        MyBase.EnforceConstraints(True)
    End Sub

    Private Sub SetMasterBOLNo(ByVal ShipBolNo As String, ByVal EDIBolNo As String)

        Try
            BeginTrans()
            Dim sql As String = String.Empty
            For Each row As DataRow In dst.Tables("WHT3PLS1").Select("SHIP_BOL_NO = '" & ShipBolNo & "'")
                row.Item("EDI_MASTER_BOL_NO") = EDIBolNo

                sql = "Update EDT945T1 SET EDI_MASTER_BOL_NO = '" & EDIBolNo & "' WHERE EDI_DOC_SEQ_NO = '" & row.Item("EDI_DOC_SEQ_NO") & "'"
                ASCDATA1.ExecuteSQL(sql)
            Next

            CommitTrans("Update Successful.")
        Catch ex As Exception
            Rollback()
            MessageBox.Show(ex.Message, "Set Master Bol No", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Function GetScaleWeight() As Decimal

        Try
            If ASCMAIN1.CLIENT <> "RGI" Then
                Return 0
            End If

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
        '**    Scale Port
        '************************** 
        Try
            tooltip.SetToolTip(txtScale, ASCMAIN1.ScalePort.PortName & ", " & ASCMAIN1.ScalePort.BaudRate & ", " & ASCMAIN1.ScalePort.DataBits & ", " & ASCMAIN1.ScalePort.Parity.ToString & ", " & ASCMAIN1.ScalePort.StopBits)
            txtScale.Text = ASCMAIN1.ScalePort.PortName
            txtScale.Appearance.BackColor = Drawing.Color.Green
        Catch ex As Exception
            txtScale.Text = String.Empty
            txtScale.Appearance.BackColor = Drawing.Color.Red
            tooltip.SetToolTip(txtScale, ex.Message)
        End Try

        '**************************
        '**    Label Printer Port
        '**************************     

        If ASCMAIN1.CLIENT = "VAN" AndAlso ASCMAIN1.USER_ID = "naseema" Then Exit Sub

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
        'Try
        '    txtScale.BackColor = Drawing.Color.Red

        '    If ASCMAIN1.ScaleSerialPort IsNot Nothing Then
        '        txtScale.Text = ASCMAIN1.ScaleSerialPort.PortName
        '        tooltip.SetToolTip(txtScale, txtScale.Text)
        '    Else
        '        txtScale.Text = "No Port"
        '        tooltip.SetToolTip(txtScale, txtScale.Text)
        '    End If

        '    txtScale.BackColor = Drawing.Color.Yellow
        '    If ASCMAIN1.ScaleSerialPort IsNot Nothing AndAlso Not ASCMAIN1.ScaleSerialPort.IsOpen Then
        '        If Not ASCMAIN1.Running_in_VS Then ASCMAIN1.ScaleSerialPort.Open()
        '    End If

        '    If ASCMAIN1.ScaleSerialPort IsNot Nothing AndAlso ASCMAIN1.ScaleSerialPort.IsOpen Then
        '        txtScale.BackColor = Drawing.Color.Green
        '    End If

        'Catch ex As Exception
        '    txtScale.BackColor = Drawing.Color.Red
        '    tooltip.SetToolTip(txtScale, ex.Message)
        'End Try

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

    ''' <summary>
    ''' Request weight from scale
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RequestWeightFromScale()

        Try
            registeredWeight = 0

            'If ASCMAIN1.ScaleSerialPort Is Nothing Then Exit Sub

            'If Not ASCMAIN1.ScaleSerialPort.IsOpen Then
            '    ASCMAIN1.ScaleSerialPort.Open()
            'End If

            '' Request the weight from the scale
            'Dim encoding As New System.Text.UTF8Encoding()
            'Dim inBuffer As Byte() = encoding.GetBytes("W")
            'ASCMAIN1.ScaleSerialPort.Write(inBuffer, 0, inBuffer.Length)

        Catch ex As Exception
            MessageBox.Show("Scale Weight Error: " & ex.Message, "Scale Weight", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

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
            '    registeredWeight = Val(readBuffer)
            'End If
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "Overrides"

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

                ' RGI ships from warehouse, not NYC where the shipments are processed.
                commonCarrier = rowSOTSVIA1 IsNot Nothing AndAlso rowSOTSVIA1.Item("CARRIER_TYPE") & String.Empty = "U" AndAlso ASCMAIN1.CLIENT <> "RGI"

                ' Insurance applies to Common Carriers (UPS, Fedex, ...)
                chkInsureShipment.Visible = commonCarrier
                If Not commonCarrier Then
                    chkInsureShipment.Checked = False
                End If
                chkInsureShipment_CheckedChanged(Nothing, Nothing)

                'grdSOTCART1.DisplayLayout.Bands(0).Columns("PACKAGING_TYPE").Hidden = Not commonCarrier
                'grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_CODE").Hidden = Not commonCarrier
        End Select

    End Sub
#End Region


End Class

Public Class Price_Change
    Public PICK_NO As String
    Public PICK_LNO As Int32
    Public STYLE_CODE As String
    Public COLOR_CODE As String
    Public PICK_UNIT_PRICE As Decimal
End Class