Imports DPayments.DShippingSDK
Imports System.Net
Imports System.IO
Imports Newtonsoft.Json
Imports Infragistics.Win.UltraWinGrid
Imports System.Text.RegularExpressions
Imports System.Net.NetworkInformation
Imports System.Drawing

' 08/08/2025
' This form was created for Vandale when they purchased Skin.
' It is used to process sales orders that are place in Totes on Trucks.
' It also produces th shipping label. Currently supports only UPS and FedEx

' ALTER TABLE SOTORDR1 ADD ORDR_WEB_IND VARCHAR2(1);
' ALTER TABLE SOTORDR1 ADD ORDR_WEB_TAX NUMBER(12,6);
' ALTER TABLE SOTSHIP1 ADD ORDR_WEB_IND VARCHAR2(1);

Public Class SOFSHIPE

    Private PICK_BATCH_NO As String = String.Empty
    Private WHSE_CODE_TRUCK As String = String.Empty
    Private drSOTTRCK1 As DataRow = Nothing

    Private allItemsOnBackOrder As Boolean = False
    Private AutoCancel As Boolean = False
    Private dictAppearances As New Dictionary(Of String, Infragistics.Win.Appearance)

    Private clsShip As New TAC.WHCSHIP1
    Private defaultPACKAGING_TYPE As String = "31"
    Private defaultPKG_CODE As String = "OTHER"
    Private clsTACZPLT1 As TAC.TACZPLT1

    Private sqlSOTPICK1 As String = String.Empty
    Private sqlSOTPICK2 As String = String.Empty


    Private WithEvents ultraComboPackage As Infragistics.Win.UltraWinGrid.UltraCombo = New Infragistics.Win.UltraWinGrid.UltraCombo

    Private Enum ScreenProcessingModes
        DisplayAvailableTrucks
        TruckSelected
        ProcessingSelectedTruckTote
    End Enum

    Private screenProcessingMode As ScreenProcessingModes = ScreenProcessingModes.DisplayAvailableTrucks

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            Get_PARM("ASTPARM1")
            Get_PARM("SOTPARM1")

            ASCMAIN1.sql = $"SELECT SOTTRCK1.TRUCK_NO, SOTTRCK1.TRUCK_TYPE, SOTTRCK1.PICK_BATCH_NO, SOTTRCK1.WHSE_CODE, 
                                SOTPICK0.PICK_BATCH_STATUS, SOTPICK0.INIT_DATE, SOTPICK0.INIT_OPER, SOTTOTE1X.NUM_TOTES
                                FROM SOTTRCK1, SOTPICK0,
                                        (
                                        SELECT SOTTOTE1.TRUCK_NO, COUNT(*) NUM_TOTES
                                        FROM SOTTOTE1, SOTPICK1
                                        WHERE SOTTOTE1.PICK_NO = SOTPICK1.PICK_NO
                                        AND SOTPICK1.PICK_STATUS = 'P'
                                        GROUP BY SOTTOTE1.TRUCK_NO
                                        ) SOTTOTE1X
                                WHERE SOTTRCK1.PICK_BATCH_NO = SOTPICK0.PICK_BATCH_NO
                                AND SOTTRCK1.WHSE_CODE = SOTPICK0.WHSE_CODE
                                AND SOTPICK0.PICK_BATCH_STATUS IN ('K', 'N')
                                AND SOTTRCK1.TRUCK_NO = SOTTOTE1X.TRUCK_NO (+)"
            Create_TDA(.Tables.Add, "SOTTRCK1X", ASCMAIN1.sql, 0, False, String.Empty)

            ASCMAIN1.sql = $"SELECT SOTPICK1.*, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.CUST_STORE_NO, SOTORDR1.CUST_STORE_NAME, SOTORDR1.CUST_DC_NO,
                                SOTORDR1.ORDR_CUST_PO, SOTSVIA1.SHIP_VIA_DESC, SOTTOTE1.TRUCK_NO, '0' SELECTED,
                                SOTPICK2.PICK_QTY, SOTPICK2.PICK_QTY_CONF, SOTPICK2.PICK_QTY_CANC, SOTPICK2.PICK_QTY_BACK
                                FROM SOTPICK1, SOTORDR1, SOTSVIA1, SOTTOTE1, SOTTRCK1,
                                    (SELECT PICK_NO, 
                                        SUM(DECODE(NVL(PICK_QTY, 0), 0, NVL(PICK_QTY_BACK, 0) + NVL(PICK_QTY_CANC, 0), NVL(PICK_QTY, 0))) PICK_QTY,  
                                        SUM(NVL(PICK_QTY_CONF, 0)) PICK_QTY_CONF, 
                                        SUM(NVL(PICK_QTY_CANC, 0)) PICK_QTY_CANC, 
                                        SUM(NVL(PICK_QTY_BACK, 0)) PICK_QTY_BACK 
                                        FROM SOTPICK2
                                        GROUP BY PICK_NO
                                    ) SOTPICK2
                                WHERE SOTPICK1.PICK_NO = SOTPICK2.PICK_NO
                                AND SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO
                                AND SOTPICK1.PICK_STATUS = 'P'
                                AND SOTPICK1.TOTE_NO = SOTTOTE1.TOTE_NO
                                AND SOTPICK1.PICK_NO = SOTTOTE1.PICK_NO 
                                AND SOTTOTE1.TRUCK_NO = SOTTRCK1.TRUCK_NO
                                AND SOTTRCK1.TRUCK_NO = :PARM1
                                AND SOTPICK1.PICK_BATCH_NO = :PARM2
                                AND SOTTRCK1.PICK_BATCH_NO = SOTPICK1.PICK_BATCH_NO
                                AND SOTTRCK1.PICK_BATCH_NO IS NOT NULL 
                                AND SOTORDR1.SHIP_VIA_CODE = SOTSVIA1.SHIP_VIA_CODE (+)"
            Create_TDA(.Tables.Add, "SOTPICK1X", ASCMAIN1.sql, 0, False, "VV", 1)
            .Tables("SOTPICK1X").Columns.Add("TOTAL_SCANNED", GetType(System.Int16), "ISNULL(PICK_QTY_CONF, 0) + ISNULL(PICK_QTY_CANC, 0) + ISNULL(PICK_QTY_BACK, 0)")
            .Tables("SOTPICK1X").Columns.Add("INCOMPLETE", GetType(System.Int16), "IIF(ISNULL(PICK_QTY, 0) <> ISNULL(TOTAL_SCANNED, 0), '1', '0')")
            .Tables("SOTPICK1X").Columns.Add("ALL_ITEMS_BACK", GetType(System.String))

            Create_TDA(.Tables.Add, "SOTPICK0", "*")
            Create_TDA(.Tables.Add, "SOTSHIP1", "*")
            Create_TDA(.Tables.Add, "ARTOPEN1", "*")
            Create_TDA(.Tables.Add, "ARTCUST1", "*")
            Create_TDA(.Tables.Add, "ARTCUST2", "*")
            Create_TDA(.Tables.Add, "ARTCUSTS", "*")
            Create_TDA(.Tables.Add, "SOTTRCK1", "*")

            sqlSOTPICK1 = "SELECT SOTPICK1.*, SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_SOURCE,
                                SOTORDR1.CUST_BILL_TO_CUST, SOTORDR1.ORDR_TYPE_CODE, SOTORDR1.POST_CODE, SOTORDR1.ORDR_FOB,
                                SOTORDR1.TERM_CODE, SOTORDR1.SREP_CODE, SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT, 
                                SOTORDR1.CUST_FACTOR_IND, SOTORDR1.CURR_CODE, SOTORDR1.CURR_EXCH_RATE, SOTORDR1.ORDR_INV_COMMENT
                                FROM SOTPICK1, SOTORDR1
                                WHERE SOTPICK1.PICK_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))
                                AND SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO"

            ASCMAIN1.sql = "SELECT * FROM SOTPICK1 WHERE PICK_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))"
            Create_TDA(.Tables.Add, "SOTPICK1", ASCMAIN1.sql, 0, True, "C")
            With .Tables("SOTPICK1")
                .Columns.Add("CUST_CODE", GetType(String))
                .Columns.Add("CUST_STORE_NO", GetType(String))
                .Columns.Add("ORDR_CUST_PO", GetType(String))
                .Columns.Add("ORDR_SOURCE", GetType(String))
                .Columns.Add("CUST_BILL_TO_CUST", GetType(String))
                .Columns.Add("ORDR_TYPE_CODE", GetType(String))
                .Columns.Add("POST_CODE", GetType(String))
                .Columns.Add("ORDR_FOB", GetType(String))
                .Columns.Add("TERM_CODE", GetType(String))
                .Columns.Add("SREP_CODE", GetType(String))
                .Columns.Add("SREP2_CODE", GetType(String))
                .Columns.Add("ORDR_DEPT", GetType(String))
                .Columns.Add("CUST_FACTOR_IND", GetType(String))
                .Columns.Add("CURR_CODE", GetType(String))
                .Columns.Add("CURR_EXCH_RATE", GetType(Decimal))
                .Columns.Add("ORDR_INV_COMMENT", GetType(String))
                .Columns.Add("TRUCK_NO", GetType(System.String))
            End With

            'dst.Tables("SOTPICK1").PrimaryKey = New DataColumn() {dst.Tables("SOTPICK1").Columns("PICK_NO")}

            ASCMAIN1.sql = "SELECT SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC, SOTORDR2.ORDR_UNIT_PRICE, SOTORDR2.CUST_UPC, SOTORDR2.COLOR_CODE,
                            SOTORDR2.RANGE_STYLE_CODE, SOTORDR2.RANGE_STYLE_LNO, SOTORDR2.CARTON_PACK_QTY, SOTORDR2.ORDR_QTY, SOTORDR2.STYLE_CODE_SUB, SOTORDR2.QTY_PER_PP
                            FROM SOTPICK2, SOTORDR2, ICTCOLR1
                            WHERE SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO
                            AND SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO
                            AND SOTORDR2.COLOR_CODE = ICTCOLR1.COLOR_CODE (+)
                            AND SOTPICK2.PICK_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))"
            Create_TDA(.Tables.Add, "SOTPICK2", ASCMAIN1.sql, 0, True, "C", 2)
            .Tables("SOTPICK2").Columns.Add("PICK_QTY_SCAN", GetType(System.Int32), "ISNULL(PICK_QTY_CONF, 0) + ISNULL(PICK_QTY_CANC, 0) + ISNULL(PICK_QTY_BACK, 0)")
            dst.Tables("SOTPICK2").PrimaryKey = New DataColumn() {dst.Tables("SOTPICK2").Columns("PICK_NO"), dst.Tables("SOTPICK2").Columns("PICK_LNO")}

            ASCMAIN1.sql = "SELECT * FROM SOTORDR1 WHERE SOTORDR1.ORDR_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))"
            Create_TDA(.Tables.Add, "SOTORDR1", ASCMAIN1.sql, 0, True, "C")
            dst.Tables("SOTORDR1").PrimaryKey = New DataColumn() {dst.Tables("SOTORDR1").Columns("ORDR_NO")}

            ASCMAIN1.sql = "SELECT * FROM SOTORDR2 WHERE ORDR_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))"
            Create_TDA(.Tables.Add, "SOTORDR2", ASCMAIN1.sql, 0, True, "C")
            dst.Tables("SOTORDR2").PrimaryKey = New DataColumn() {dst.Tables("SOTORDR2").Columns("ORDR_NO"), dst.Tables("SOTORDR2").Columns("ORDR_LNO")}

            ASCMAIN1.sql = "SELECT * FROM SOTORDR5 WHERE ORDR_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1))) AND CUST_ADDR_TYPE = 'ST'"
            Create_TDA(.Tables.Add, "SOTORDR5", ASCMAIN1.sql, 0, True, "C")
            .Tables("SOTORDR5").Columns.Add("IS_RESIDENTAL", GetType(System.String))
            .Tables("SOTORDR5").Columns("IS_RESIDENTAL").DefaultValue = "1"
            .Tables("SOTORDR5").Columns.Add("IS_PO_BOX", GetType(System.String))
            .Tables("SOTORDR5").Columns("IS_PO_BOX").DefaultValue = "0"
            'dst.Tables("SOTORDR5").PrimaryKey = New DataColumn() {dst.Tables("SOTORDR5").Columns("ORDR_NO"), dst.Tables("SOTORDR5").Columns("CUST_ADDR_TYPE")}

            Create_TDA(.Tables.Add, "TATEVNT1", "*")

            Create_TDA(.Tables.Add, "SOTINVH1", "*")
            Create_TDA(.Tables.Add, "SOTINVH2", "*", 1)

            .Tables("SOTINVH2").Columns.Add("EXT_PRICE", GetType(System.Decimal), "ISNULL(ORDR_UNIT_PRICE, 0) * ISNULL(ORDR_QTY_SHIP, 0)")

            Create_TDA(.Tables.Add, "SOTINVH9", "*")
            Create_TDA(.Tables.Add, "SOTINVHM", "*")
            Create_TDA(.Tables.Add, "SOTRNGA1", "*")

            Create_TDA(.Tables.Add, "ICTWHSE1", "*", -1, False)
            Fill_Records("ICTWHSE1", String.Empty, True, "SELECT * FROM ICTWHSE1")

            Create_TDA(.Tables.Add, "ASTUSER1", "*", -1, False)
            Fill_Records("ASTUSER1", String.Empty, True, "SELECT * FROM ASTUSER1")

            Create_TDA(.Tables.Add, "WHTPKGM1", "*", -1, False)
            Fill_Records("WHTPKGM1", String.Empty, True, "SELECT * FROM WHTPKGM1")

            Create_TDA(.Tables.Add, "SOTSVIA1", "*", -1, False)
            Fill_Records("SOTSVIA1", String.Empty, True, "SELECT * FROM SOTSVIA1 WHERE CARRIER_PROD_CODE IS NOT NULL")

            Create_TDA(.Tables.Add, "TATSTATE", "*", -1, False)
            Fill_Records("TATSTATE", String.Empty, True, "SELECT * FROM TATSTATE")

            Create_TDA(.Tables.Add, "SOTCARR1", "*")
            Fill_Records("SOTCARR1", String.Empty, True, "SELECT * FROM SOTCARR1")

            Create_TDA(.Tables.Add, "SOTCARR3", "*")
            .Tables("SOTCARR3").Columns.Add("CARRIER_REMOTE_HOST_IP", GetType(System.String))
            Fill_Records("SOTCARR3", "", True, "SELECT SOTCARR3.*, SOTCARR1.CARRIER_REMOTE_HOST_IP FROM SOTCARR3, SOTCARR1 WHERE SOTCARR3.CARRIER_CODE = SOTCARR1.CARRIER_CODE (+)")

            Create_TDA(.Tables.Add, "SOTCARR5", "*")
            Fill_Records("SOTCARR5", String.Empty, True, "SELECT * FROM SOTCARR5")

            Create_TDA(.Tables.Add, "SOTCART1", "*")
            .Tables("SOTCART1").Columns.Add("REFERENCE1", GetType(System.String))
            .Tables("SOTCART1").Columns.Add("REFERENCE2", GetType(System.String))
            .Tables("SOTCART1").Columns.Add("REFERENCE3", GetType(System.String))
            .Tables("SOTCART1").Columns.Add("SHIP_BOL_NO", GetType(System.String))

            Create_TDA(.Tables.Add, "SOTCART2", "*")
            .Tables("SOTCART2").Columns.Add("STYLE_DESC", GetType(System.String))

            Create_Relation("SOTCART1", "SOTCART2", "CART_NO")

            ' Shipping Label
            Create_TDA(.Tables.Add, "WHTSHPC1", "*", 1)
            Create_TDA(.Tables.Add, "WHTSHPC2", "*", 1)
            Create_TDA(.Tables.Add, "WHTSHPC3", "*", 1)

            Create_TDA(.Tables.Add, "WHTSHPC4", "*", 1)
            .Tables("WHTSHPC4").Columns.Add("SHIP_VIA_CODE", GetType(System.String))

            Create_TDA(.Tables.Add, "WHTSHPC5", "*", 1)
            Create_TDA(.Tables.Add, "WHTSHPCG", "*", 1)
            Create_TDA(.Tables.Add, "WHTSHPCC", "*", 1)
            Create_TDA(.Tables.Add, "WHTSHPCS", "*", 1)
            Create_TDA(.Tables.Add, "WHTSHPCP", "*", 1)
            Create_TDA(.Tables.Add, "WHTSHPCA", "*", 1)
            .Tables("WHTSHPCA").Columns.Add("ADDON_TYPE_DESC", GetType(System.String))
            .Tables("WHTSHPCA").Columns.Add("PROHIBITED", GetType(System.String))
            .Tables("WHTSHPCA").Columns.Add("REQUIRED", GetType(System.String))

            .Relations.Add("WHTSHPC4_WHTSHPCA",
               New DataColumn() {dst.Tables("WHTSHPC4").Columns("SHIP_CNTL_NO"), dst.Tables("WHTSHPC4").Columns("SERVICE_INDEX")},
               New DataColumn() {dst.Tables("WHTSHPCA").Columns("SHIP_CNTL_NO"), dst.Tables("WHTSHPCA").Columns("SERVICE_INDEX")})

            .Tables("WHTSHPCA").Columns.Add("ADDON_TOTAL", GetType(System.Decimal), "IIF(SELECTED = '1', ISNULL(ADDON_AMOUNT, 0), 0)")
            .Tables("WHTSHPC4").Columns.Add("ADDON_TOTAL", GetType(System.Decimal), "SUM(CHILD(WHTSHPC4_WHTSHPCA).ADDON_TOTAL)")
            .Tables("WHTSHPC4").Columns.Add("CUSTOMER_BASE_CHARGE", GetType(System.Decimal))
            .Tables("WHTSHPC4").Columns.Add("TOTAL_CHARGE", GetType(System.Decimal), "ISNULL(ADDON_TOTAL, 0) + ISNULL(CUSTOMER_BASE_CHARGE, 0) + ISNULL(SURCHARGE, 0)")

        End With

        grdSOTTRCK1X.DataSource = dst.Tables("SOTTRCK1X")
        Create_Summary(grdSOTTRCK1X, "TRUCK_NO", "Count")

        grdSOTPICK1X.DataSource = dst.Tables("SOTPICK1X")
        Create_Summary(grdSOTPICK1X, "PICK_NO", "Count")

        grdSOTPICK2.DataSource = dst.Tables("SOTPICK2")
        Create_Summary(grdSOTPICK2, "PICK_LNO", "Count")
        Create_Summary(grdSOTPICK2, "PICK_QTY", "Sum")
        Create_Summary(grdSOTPICK2, "PICK_QTY_CONF", "Sum")
        Create_Summary(grdSOTPICK2, "PICK_QTY_CANC", "Sum")
        Create_Summary(grdSOTPICK2, "PICK_QTY_BACK", "Sum")

        grdWHTSHPC4.DataSource = dst.Tables("WHTSHPC4")

        grdSOTORDR5.DataSource = dst.Tables("SOTORDR5")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")
        grdSOTCART1.AllowDrop = True
        grdSOTCART2.DataSource = dst.Tables("SOTCART2")

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

        ultraComboPackage.DataSource = ASCDATA1.GetDataTable("SELECT PKG_CODE, PKG_DESC, PKG_L || ' x ' ||  PKG_W || ' x ' || PKG_H PKG_D FROM WHTPKGM1 order by PKG_CODE")
        ultraComboPackage.ValueMember = "PKG_CODE"
        ultraComboPackage.DisplayMember = "PKG_DESC"
        grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_CODE").EditorComponent = ultraComboPackage

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
                If New String() {"CART_TOTAL_WGT_ACTUAL", "PKG_CODE", "PACKAGING_TYPE", "INSURANCE", "PKG_W", "PKG_L", "PKG_H"}.Contains(gcol.Key) Then
                    ' "CART_FREIGHT", "CART_SEQ", "REFERENCE1", "REFERENCE2", "REFERENCE3", "PKG_BOX_UPC"
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        Create_Summary(grdSOTCART1, "CART_NO", "Count")
        Create_Summary(grdSOTCART1, New String() _
            {"CART_FREIGHT", "CART_TOTAL_UNITS", "CART_TOTAL_WGT_ACTUAL"})

        grdSOTTRCK1X.Parent = tab.Parent
        splTotes.Parent = tab.Parent

        SetUpPortsAndPrinters()
        SetupScanner()
        CreateAppearances()
        txtUSER_ID.Text = ASCMAIN1.USER_ID
        clsTACZPLT1 = New TAC.TACZPLT1

        Timer1.Start()
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = String.Empty

        Select Case eItemKey

            Case "Cancel"
                If Not AutoCancel Then
                    If MessageBox.Show($"Do you want to Cancel processing Truck {txtTRUCK_NO.Text}", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Refresh"

            Case "Request Rates"
                If dst.Tables("SOTCART1").Rows.Count = 0 Then
                    EMsg &= vbCr & "At least one carton must be added before shipping rates can be requested."
                End If

            Case "Ship Order"
                If txtSHIP_VIA_CODE.TextLength = 0 Then
                    EMsg &= vbCr & "A shipping method must be selected before the order can be shipped."
                End If

                If dst.Tables("WHTSHPC4").Rows.Count = 0 Then
                    EMsg &= vbCr & "You need to request shipping rates before the order can be shipped."
                End If

                Dim CART_SEQ As Int16 = 1
                dst.Tables("SOTCART1").AcceptChanges()
                dst.Tables("SOTCART2").AcceptChanges()

                For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("", "CART_SEQ")
                    drSOTCART1.SetAdded()
                    drSOTCART1.Item("CART_SEQ") = CART_SEQ
                    CART_SEQ += 1
                    Dim CART_NO As String = drSOTCART1.Item("CART_NO") & String.Empty
                    If dst.Tables("SOTCART2").Select($"CART_NO = '{CART_NO}'").Length = 0 Then
                        EMsg &= vbCr & "All cartons must include at least one product before proceeding."
                        Exit For
                    End If

                    Dim CART_LNO As Int16 = 1
                    For Each drSOTCART2 As DataRow In dst.Tables("SOTCART2").Select($"CART_NO = '{CART_NO}'", "CART_LNO")
                        drSOTCART2.SetAdded()
                        drSOTCART2.Item("CART_LNO") = CART_LNO
                        CART_LNO += 1
                    Next
                Next

                Dim cost As Decimal = 0

                If dst.Tables("WHTSHPC4").Select($"SHIP_VIA_CODE = '{txtSHIP_VIA_CODE.Text}'").Length = 0 Then
                    EMsg &= vbCr & "You must select a Shipping Method that is listed in the Carrier Shipping Rates grid before the order can be shipped. Double-click a shipping method in the grid to select it for this order."
                Else
                    Dim drWHTSHPC4 As DataRow = dst.Tables("WHTSHPC4").Select($"SHIP_VIA_CODE = '{txtSHIP_VIA_CODE.Text}'")(0)
                    cost = Val(drWHTSHPC4.Item("ACCT_NET_CHARGE") & String.Empty)
                End If

                If EMsg.Length > 0 Then
                    Exit Select
                End If

                Dim zMsg As String = $"Would you like to ship this order using the shipping method {txtSHIP_VIA_DESC.Text} which costs ${cost.ToString("#,##0.00")}?"
                If MessageBox.Show(zMsg, "Ship Order", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

                Dim ErrorMessage As String = String.Empty
                ' This is a prescreen for potential issues.
                If Not RequestShippingLabel("", ErrorMessage, True) Then
                    EMsg &= vbCr & ErrorMessage
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

            Case "Select"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Me.Cursor = Cursors.WaitCursor
                screenProcessingMode = ScreenProcessingModes.DisplayAvailableTrucks
                Mode_Settings(False)
                Me.Cursor = Cursors.Default

            Case "Refresh"
                Me.Cursor = Cursors.WaitCursor
                Mode_Settings(False)
                Me.Cursor = Cursors.Default

            Case "Request Rates"
                Dim PICK_NO As String = dst.Tables("SOTPICK1").Select($"TOTE_NO = '{txtTOTE_NO.Text}'")(0).Item("PICK_NO")
                RequestRates(PICK_NO)

            Case "Ship Order"
                Dim ErrorMessage As String = String.Empty

                ' Create Invoice
                Dim PICK_NO As String = dst.Tables("SOTPICK1").Select($"TOTE_NO = '{txtTOTE_NO.Text}'")(0).Item("PICK_NO")
                If CreateSalesOrderInvoice(PICK_NO, txtSHIP_VIA_CODE.Text) Then
                    dst.Tables("SOTPICK1X").Select($"PICK_NO = '{PICK_NO}'")(0).Item("SELECTED") = "1"
                    Dim INV_NO As String = dst.Tables("SOTINVH1").Rows(0).Item("INV_NO")
                    If Not RequestShippingLabel(INV_NO, ErrorMessage, False) Then
                        MessageBox.Show(ErrorMessage, "Generate Shipping Label", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                    Click_Command("Select")
                Else
                    AutoCancel = True
                    Click_Command("Cancel")
                    AutoCancel = False
                End If
        End Select

        ASCMAIN1.Progress("", "")

        Timer1.Start()

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Request Rates").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Ship Order").Settings.Enabled = DefaultableBoolean.False
            End With
        End If

        If ScreenMode Then
            grdSOTTRCK1X.Visible = False
            splTotes.Visible = True
            Sort_grdColumns(grdSOTPICK1X, "TOTE_NO")
            txtTOTE_NO.ReadOnly = False
            txtTRUCK_NO.ReadOnly = True
        Else
            grdSOTTRCK1X.Visible = True
            Sort_grdColumns(grdSOTTRCK1X, "TRUCK_NO")
            splTotes.Visible = False
            txtTRUCK_NO.ReadOnly = False
            txtTOTE_NO.ReadOnly = True
            Clear_Record()
            AutoCancel = False
        End If

        grdSOTPICK1X.DisplayLayout.Bands(0).Columns("TRUCK_NO").Hidden = True

        'Set_Read_Only(grpHeader, ScreenMode)

    End Sub

    Private Sub Clear_Record()

        EnforceConstraints(False)
        For Each tableName As String In New String() {"SOTINVH1", "SOTINVH2", "SOTINVH9", "SOTINVHM", "ARTOPEN1", "SOTRNGA1",
            "WHTSHPC1", "WHTSHPC2", "WHTSHPC3", "WHTSHPC4", "WHTSHPC5",
            "WHTSHPCG", "WHTSHPCC", "WHTSHPCS", "WHTSHPCP", "WHTSHPCA",
            "SOTPICK0", "SOTPICK1", "SOTPICK2", "SOTORDR1", "SOTORDR2", "SOTORDR5", "TATEVNT1",
            "SOTSHIP1", "SOTCART1", "SOTCART2"}
            If dst.Tables.Contains(tableName) Then
                dst.Tables(tableName).Rows.Clear()
            End If
        Next
        EnforceConstraints(True)

        Clear_All_Filters(grdSOTPICK1X)
        Clear_All_Filters(grdSOTPICK2)
        Sort_grdColumns(grdSOTPICK1X, "TOTE_NO")

        txtTRUCK_NO.Clear()
        txtTOTE_NO.Clear()

        PICK_BATCH_NO = String.Empty
        WHSE_CODE_TRUCK = String.Empty

        grdSOTPICK2.Text = String.Empty
        Fill_Records("SOTTRCK1X")
        Sort_grdColumns(grdSOTTRCK1X, "TRUCK_NO")

        txtTRUCK_NO.ReadOnly = False
        txtTOTE_NO.ReadOnly = True

        txtUSER_ID.Text = ASCMAIN1.USER_ID
        txtWHSE_CODE.Clear()
        txtSHIP_VIA_CODE.Clear()

        drSOTTRCK1 = Nothing

        screenProcessingMode = ScreenProcessingModes.DisplayAvailableTrucks

        ASCMAIN1.MultiTask_Release()

    End Sub

    Private Sub Load_Record()
        Save_Header_Fields(grpHeader)

        grdSOTPICK2.Text = String.Empty
        Dim dvw As DataView = DirectCast(grdSOTPICK2.DataSource, DataTable).DefaultView
        dvw.RowFilter = "PICK_NO = '********'"
        grdSOTPICK2.Text = ""

        Dim dvw5 As DataView = DirectCast(grdSOTORDR5.DataSource, DataTable).DefaultView
        dvw5.RowFilter = "ORDR_NO =  '********'"
        grdSOTORDR5.Text = ""

        grdSOTCART1.Text = ""
        grdSOTCART2.Text = ""
        grdWHTSHPC4.Text = ""

        EnforceConstraints(False)
        For Each tableName As String In New String() {"SOTINVH1", "SOTINVH2", "SOTINVH9", "SOTINVHM", "ARTOPEN1", "SOTRNGA1",
            "WHTSHPC1", "WHTSHPC2", "WHTSHPC3", "WHTSHPC4", "WHTSHPC5",
            "WHTSHPCG", "WHTSHPCC", "WHTSHPCS", "WHTSHPCP", "WHTSHPCA", "TATEVNT1",
            "SOTSHIP1", "SOTCART1", "SOTCART2"}
            If dst.Tables.Contains(tableName) Then
                dst.Tables(tableName).Rows.Clear()
            End If
        Next
        EnforceConstraints(True)

        screenProcessingMode = ScreenProcessingModes.TruckSelected
    End Sub

    Private Sub Update_Record()

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTTRCK1X, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTPICK1X, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTPICK2, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTCART1, "B", "Add Carton")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        If Not GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then
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

        Select Case e.SourceControl.Name

            Case "grdSOTORDR0"

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
        Else

            Select Case e.SourceControl.Name
                Case "grdSOTALLOX"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing

        If grdSOTCART1.ActiveRow Is Nothing Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Add Carton"
                Dim PICK_NO As String = dst.Tables("SOTPICK1").Select($"TOTE_NO = '{txtTOTE_NO.Text}'")(0).Item("PICK_NO")
                AddCarton(PICK_NO)
        End Select

        Me.Cursor = Cursors.Default

    End Sub

#End Region

#Region "Form Procedures"

    Private Delegate Sub ScannerDelegate(ByVal ScannedString As String)
    Private scannedDelegate As ScannerDelegate = Nothing

    Private Sub BackOrderedAndCancelledItems(ByVal PICK_NO As String)

        ' All items in Back or Canc need to come out of Pick
        Dim drSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
        Dim WHSE_CODE As String = drSOTPICK1.Item("WHSE_CODE") & String.Empty
        Dim ORDR_NO As String = drSOTPICK1.Item("ORDR_NO") & String.Empty
        Dim drSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
        Dim PARTNER_CODE As String = drSOTORDR1.Item("PARTNER_CODE") & String.Empty

        For Each drSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}'")
            Dim ITEM_CODE As String = drSOTPICK2.Item("ITEM_CODE") & String.Empty
            Dim TakeOutOfPick As Int32 = Val(drSOTPICK2.Item("PICK_QTY_BACK") & String.Empty) + Val(drSOTPICK2.Item("PICK_QTY_CANC") & String.Empty)

            If TakeOutOfPick > 0 Then
                ASCMAIN1.sql = $"Update ICTSTAT2 SET WHSE_QTY_PICK = NVL(WHSE_QTY_PICK, 0) - :PARM1, 
                                    WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN, 0) + :PARM2
                                    WHERE ITEM_CODE = :PARM3 AND WHSE_CODE = :PARM4"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NNVV", New Object() {TakeOutOfPick, Val(drSOTPICK2.Item("PICK_QTY_BACK") & String.Empty), ITEM_CODE, WHSE_CODE})

                ASCMAIN1.sql = $"Update ICTSTAG2 SET WHSE_QTY_PICK = NVL(WHSE_QTY_PICK, 0) - :PARM1,
                                    WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN, 0) + :PARM2
                                    WHERE PARTNER_CODE = :PARM3 AND ITEM_CODE = :PARM4 AND WHSE_CODE = :PARM5"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NNVVV", New Object() {TakeOutOfPick, Val(drSOTPICK2.Item("PICK_QTY_BACK") & String.Empty), PARTNER_CODE, ITEM_CODE, WHSE_CODE})
            End If
        Next

    End Sub

    Private Function ProcessTote(ByVal TOTE_NO As String) As Boolean

        Dim PKG_WT As Decimal = 0

        ' Clear all Shipping and Invoice Tables for the scanned tote
        EnforceConstraints(False)
        For Each tableName As String In New String() {"TATEVNT1",
                                    "SOTINVH1", "SOTINVH2", "SOTINVH9", "SOTINVHM", "ARTOPEN1", "SOTRNGA1",
                                    "WHTSHPC1", "WHTSHPC2", "WHTSHPC3", "WHTSHPC4", "WHTSHPC5",
                                    "WHTSHPCG", "WHTSHPCC", "WHTSHPCS", "WHTSHPCP", "WHTSHPCA",
                                    "SOTSHIP1", "SOTCART1", "SOTCART2"}
            If dst.Tables.Contains(tableName) Then
                dst.Tables(tableName).Rows.Clear()
            End If
        Next

        Try
            Dim drSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Select($"TOTE_NO = '{TOTE_NO}'")(0)
            Dim PICK_NO As String = drSOTPICK1.Item("PICK_NO") & String.Empty
            Dim ORDR_NO As String = drSOTPICK1.Item("ORDR_NO") & String.Empty
            Dim WHSE_CODE As String = drSOTPICK1.Item("ORDR_NO") & String.Empty
            Dim SHIP_VIA_CODE As String = drSOTPICK1.Item("SHIP_VIA_CODE") & String.Empty

            ' Evaluate Customer Address 1 and 2 to see if it is a PO Box
            For Each drSOTORDR5 As DataRow In dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}'")
                Dim CUST_ADDR1 As String = drSOTORDR5.Item("CUST_ADDR1") & String.Empty
                Dim CUST_ADDR2 As String = drSOTORDR5.Item("CUST_ADDR2") & String.Empty
                If IsPOBox(CUST_ADDR1) OrElse IsPOBox(CUST_ADDR2) Then
                    drSOTORDR5.Item("IS_PO_BOX") = "1"
                End If
            Next
            dst.Tables("SOTORDR5").AcceptChanges()

            grdSOTPICK2.Text = String.Empty
            Dim dvw As DataView = DirectCast(grdSOTPICK2.DataSource, DataTable).DefaultView
            grdSOTPICK2.Text = $"Details for Tote: {TOTE_NO}, Pick Ticket: {PICK_NO}"

            dvw.RowFilter = $"PICK_NO = '{PICK_NO}'"
            grdSOTPICK2.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
            Sort_grdColumns(grdSOTPICK2, "PICK_LNO")

            Dim dvw5 As DataView = DirectCast(grdSOTORDR5.DataSource, DataTable).DefaultView
            dvw5.RowFilter = $"ORDR_NO = '{ORDR_NO}'"
            grdSOTORDR5.Text = $"Ship Address for Tote: {TOTE_NO}"

            grdSOTCART1.Text = $"Cartons for Tote: {TOTE_NO}"

            UltraExplorerBar1.Groups("Screen Control").Items("Request Rates").Settings.Enabled = DefaultableBoolean.True
            UltraExplorerBar1.Groups("Screen Control").Items("Ship Order").Settings.Enabled = DefaultableBoolean.True

            txtSHIP_VIA_CODE.Text = SHIP_VIA_CODE

            ASCMAIN1.sql = $"SELECT * FROM SOTCART1 WHERE PICK_NO = '{PICK_NO}'"
            Fill_Records("SOTCART1", String.Empty, True, ASCMAIN1.sql)

            ASCMAIN1.sql = $"SELECT * FROM SOTCART2 WHERE CART_NO IN (SELECT CART_NO FROM SOTCART1 WHERE PICK_NO = '{PICK_NO}')"
            Fill_Records("SOTCART2", String.Empty, True, ASCMAIN1.sql)

            If dst.Tables("SOTCART1").Rows.Count = 0 Then
                Dim CART_NO As String = AddCarton(PICK_NO)

                ' Add all items to the cart.
                Dim CART_LNO As Int16 = 1
                For Each drSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}' AND ISNULL(PICK_QTY_CONF, 0) > 0", "PICK_LNO")
                    Dim drSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow
                    drSOTCART2.Item("CART_NO") = CART_NO
                    drSOTCART2.Item("CART_LNO") = CART_LNO
                    CART_LNO += 1

                    drSOTCART2.Item("ORDR_NO") = drSOTPICK2.Item("ORDR_NO")
                    drSOTCART2.Item("ORDR_LNO") = drSOTPICK2.Item("ORDR_LNO")
                    drSOTCART2.Item("QTY_PACKED") = drSOTPICK2.Item("PICK_QTY_CONF")
                    drSOTCART2.Item("UPC_CODE") = drSOTPICK2.Item("CUST_UPC")
                    'drSOTCART2.Item("SKU_NO") = ""
                    drSOTCART2.Item("STYLE_CODE") = drSOTPICK2.Item("STYLE_CODE")
                    drSOTCART2.Item("COLOR_CODE") = drSOTPICK2.Item("COLOR_CODE")
                    'drSOTCART2.Item("SIZE_DESC") = ""
                    'drSOTCART2.Item("STYLE_PREPACK") = ""
                    'drSOTCART2.Item("ITEM_EXP_DATE") = ""
                    'drSOTCART2.Item("QTY_REL") = ""
                    'drSOTCART2.Item("P2L_INIT") = ""
                    drSOTCART2.Item("STYLE_DESC") = drSOTPICK2.Item("STYLE_DESC")
                    dst.Tables("SOTCART2").Rows.Add(drSOTCART2)
                Next
            End If

            grdSOTCART1.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
            grdSOTCART2.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)

            If UltraExplorerBar1.Groups.Count > 0 Then
                With UltraExplorerBar1
                    .Groups("Screen Control").Items("Cancel").Settings.Enabled = DefaultableBoolean.True
                    .Groups("Screen Control").Items("Refresh").Settings.Enabled = DefaultableBoolean.False
                    .Groups("Screen Control").Items("Request Rates").Settings.Enabled = DefaultableBoolean.True
                    .Groups("Screen Control").Items("Ship Order").Settings.Enabled = DefaultableBoolean.True
                End With
            End If

            screenProcessingMode = ScreenProcessingModes.ProcessingSelectedTruckTote
            txtTOTE_NO.ReadOnly = True

            HighLightSelectedTote()

            Return True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Process Tote", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function

    Private Sub HighLightSelectedTote()

        Try
            For Each row As UltraGridRow In grdSOTPICK1X.Rows
                If row.Cells("TOTE_NO").Value & String.Empty = txtTOTE_NO.Text Then
                    row.Appearance = dictAppearances("Selected_Tote")
                Else
                    row.Appearance = Nothing
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Enum ValidateTruckToteTypes
        Truck
        Tote
    End Enum

    Private Function ValidateTruckTote(ByVal ValidationType As ValidateTruckToteTypes, ByVal InputValue As String) As Boolean

        Try
            Dim drSOTPICK1X As DataRow = Nothing
            Me.Cursor = Cursors.WaitCursor

            allItemsOnBackOrder = False

            Select Case ValidationType
                Case ValidateTruckToteTypes.Truck
                    WHSE_CODE_TRUCK = String.Empty
                    drSOTTRCK1 = Nothing

                    ASCMAIN1.Progress("Loading Truck", InputValue)

                    If InputValue.Length = 0 Then
                        MessageBox.Show("The supplied Truck is invalid.", "Validate Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    drSOTTRCK1 = Fill_Record("SOTTRCK1", InputValue)
                    If dst.Tables("SOTTRCK1").Rows.Count = 0 OrElse drSOTTRCK1 Is Nothing Then
                        MessageBox.Show("The supplied Truck is invalid or does not have any Totes to process", "Validate Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    If drSOTTRCK1.Item("PICK_BATCH_NO") & String.Empty = String.Empty Then
                        MessageBox.Show("The supplied Truck is NOT assigned to a Pick Batch", "Validate Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    PICK_BATCH_NO = drSOTTRCK1.Item("PICK_BATCH_NO") & String.Empty

                    Fill_Records("SOTPICK0", PICK_BATCH_NO)
                    If dst.Tables("SOTPICK0").Rows.Count = 0 Then
                        MessageBox.Show("The supplied Truck is NOT assigned to a Pick Batch that can be found.", "Validate Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    Dim drSOTPICK0 As DataRow = dst.Tables("SOTPICK0").Rows(0)

                    Dim PICK_BATCH_STATUS As String = drSOTPICK0.Item("PICK_BATCH_STATUS") & String.Empty
                    If Not "NK".Contains(PICK_BATCH_STATUS) Then
                        MessageBox.Show($"The supplied Truck's Pick Batch Status must be Picked or In Pack.", "Validate Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    WHSE_CODE_TRUCK = drSOTPICK0.Item("WHSE_CODE") & String.Empty
                    txtWHSE_CODE.Text = WHSE_CODE_TRUCK

                    ASCMAIN1.Progress("Fill Pick Tickets In Pick", "")
                    FillPickTicketsInPick(InputValue)

                    If dst.Tables("SOTPICK1X").Rows.Count = 0 Then
                        MessageBox.Show("The supplied Truck does not have any Totes to process", "Validate Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    Dim numIncomplete As Int32 = dst.Tables("SOTPICK1X").Select("INCOMPLETE = '1'").Length
                    If numIncomplete > 0 Then
                        Dim zMsg As String = "The following Totes are not Complete:" & Environment.NewLine
                        For Each drSOTPICK1_X As DataRow In dst.Tables("SOTPICK1X").Select("INCOMPLETE = '1'")
                            zMsg &= drSOTPICK1_X.Item("TOTE_NO") & Environment.NewLine
                        Next

                        MessageBox.Show(zMsg, "Process Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        dst.Tables("SOTPICK1X").Rows.Clear()
                        Clear_Record()
                        Return False
                    End If

                    Dim PICK_REQ_RES As Int32 = dst.Tables("SOTPICK1X").Select("PICK_REQ_RES = '1'").Length
                    If PICK_REQ_RES > 0 Then
                        Dim zMsg As String = "The following Totes Require Resolution:" & Environment.NewLine
                        For Each drSOTPICK1_X As DataRow In dst.Tables("SOTTRCK1X").Select("PICK_REQ_RES = '1'")
                            zMsg &= drSOTPICK1_X.Item("TOTE_NO") & Environment.NewLine
                        Next

                        MessageBox.Show(zMsg, "Process Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        dst.Tables("SOTTRCK1X").Rows.Clear()
                        Clear_Record()
                        Return False
                    End If

                    If Not ASCMAIN1.Logical_Lock("SOTTRCK1", InputValue) Then
                        dst.Tables("SOTTRCK1X").Rows.Clear()
                        Clear_Record()
                        Return False
                    End If

                    If Not ASCMAIN1.Logical_Lock("SOTPICK0", PICK_BATCH_NO) Then
                        dst.Tables("SOTTRCK1X").Rows.Clear()
                        Clear_Record()
                        Return False
                    End If

                    drSOTPICK1X = dst.Tables("SOTPICK1X").Rows(0)

                    ASCMAIN1.Progress("Load Pick Tickets and Sales Orders", "")
                    Dim lstPICK_NOs As New List(Of String)
                    Dim lstORDR_NOs As New List(Of String)

                    For Each drSOTPICK1_X As DataRow In dst.Tables("SOTPICK1X").Select("")

                        Dim PICK_NO As String = drSOTPICK1_X.Item("PICK_NO") & String.Empty
                        Dim ORDR_NO As String = drSOTPICK1_X.Item("ORDR_NO") & String.Empty
                        Dim TOTE_NO As String = drSOTPICK1_X.Item("TOTE_NO") & String.Empty

                        ASCMAIN1.Progress("Loading Pick Tickets and Sales Orders", TOTE_NO)

                        If Not ASCMAIN1.Logical_Lock("SOTTOTE1", TOTE_NO) Then
                            Clear_Record()
                            Return False
                        End If

                        If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO) Then
                            Clear_Record()
                            Return False
                        End If

                        If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then
                            Clear_Record()
                            Return False
                        End If

                        lstPICK_NOs.Add(PICK_NO)
                        lstORDR_NOs.Add(ORDR_NO)
                    Next

                    ASCMAIN1.Progress("Load Pick Tickets and Sales Orders", "")
                    Fill_Records("SOTPICK1", String.Join(",", lstPICK_NOs.ToArray), True, sqlSOTPICK1)
                    Fill_Records("SOTPICK2", String.Join(",", lstPICK_NOs.ToArray))
                    Fill_Records("SOTORDR1", String.Join(",", lstORDR_NOs.ToArray))
                    Fill_Records("SOTORDR2", String.Join(",", lstORDR_NOs.ToArray))
                    Fill_Records("SOTORDR5", String.Join(",", lstORDR_NOs.ToArray))
                    Fill_Records("ARTCUST2", String.Join(",", lstORDR_NOs.ToArray))

                    Dim numOrdertypes As Int16 = ASCDATA1.SelectDistinct(dst.Tables("SOTORDR1"), New String() {"ORDR_TYPE_CODE"}).Rows.Count
                    If numOrdertypes > 1 Then
                        MessageBox.Show("The supplied Truck's sales orders have more than 1 Order Type Code", "Validate Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Clear_Record()
                        Return False
                    End If

                    Dim CUST_CODE As String = drSOTPICK1X.Item("CUST_CODE") & String.Empty
                    Dim CUST_STORE_NO As String = drSOTPICK1X.Item("CUST_STORE_NO") & String.Empty
                    Dim SHIP_VIA_CODE As String = drSOTPICK1X.Item("SHIP_VIA_CODE") & String.Empty

                    Dim CUST_DC_NO As String = drSOTPICK1X.Item("CUST_DC_NO") & String.Empty
                    ' See if shipping to a DC
                    If CUST_DC_NO.Length > 0 Then
                        CUST_STORE_NO = CUST_DC_NO
                    End If

                    'isCustomTruck = drSOTTRCK1.Item("TRUCK_TYPE") & String.Empty = "X"

                    ' Message if there are pick tickets where all items are back ordered.
                    Dim lstTotes As New List(Of String)
                    For Each drSOTPICK1X In dst.Tables("SOTPICK1X").Select
                        Dim PICK_NO As String = drSOTPICK1X.Item("PICK_NO") & String.Empty
                        Dim TOTE_NO As String = drSOTPICK1X.Item("TOTE_NO") & String.Empty

                        If dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}' AND ISNULL(PICK_QTY_CONF, 0) > 0").Length = 0 Then
                            lstTotes.Add(TOTE_NO)
                            drSOTPICK1X.Item("ALL_ITEMS_BACK") = "1"
                        End If
                    Next

                    If lstTotes.Count > 0 Then
                        MessageBox.Show($"The following totes have all items back ordered. You must scan these totes to close them out.{Environment.NewLine}{String.Join(Environment.NewLine, lstTotes.ToArray)}", "Back Orders", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                    Click_Command("Select")
                    Return True

                Case ValidateTruckToteTypes.Tote

                    ASCMAIN1.Progress("Validate Tote", InputValue)

                    drSOTPICK1X = Nothing
                    If dst.Tables("SOTPICK1X").Select($"TOTE_NO = '{InputValue}'").Length = 1 Then
                        drSOTPICK1X = dst.Tables("SOTPICK1X").Select($"TOTE_NO = '{InputValue}'")(0)
                    Else
                        MessageBox.Show($"Cannot locate Tote: {InputValue}", "Validate Tote", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    If drSOTPICK1X.Item("SELECTED") & String.Empty = "1" Then
                        MessageBox.Show($"Tote: {InputValue} was already processed.", "Validate Tote", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    If drSOTPICK1X.Item("PICK_REQ_RES") & String.Empty = "1" Then
                        MessageBox.Show($"Tote: {InputValue} Requires Resolution.", "Validate Tote", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    If drSOTPICK1X.Item("INCOMPLETE") & String.Empty = "1" Then
                        MessageBox.Show($"Tote: {InputValue} is incomplete.", "Validate Tote", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    Return ProcessTote(InputValue)
            End Select

        Catch ex As Exception
            Me.Cursor = Cursors.Default

            MessageBox.Show($"Validate Error: {ex.Message}", "Validate Tote", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If ValidationType = ValidateTruckToteTypes.Truck Then
                Clear_Record()
            End If
            Return False

        Finally
            ASCMAIN1.Progress("", "")
            dst.Tables("SOTINVH1").Rows.Clear()
            dst.Tables("SOTINVH2").Rows.Clear()
            Me.Cursor = Cursors.Default
            Timer1.Start()
        End Try
    End Function

    Private Sub FillPickTicketsInPick(ByVal TRUCK_NO As String)
        grdSOTPICK2.Text = String.Empty
        Fill_Records("SOTPICK1X", New Object() {TRUCK_NO, PICK_BATCH_NO})
        grdSOTPICK1X.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
        Clear_All_Filters(grdSOTPICK1X)
        Sort_grdColumns(grdSOTPICK1X, "TOTE_NO")
        If dst.Tables("SOTTRCK1X").Rows.Count > 0 Then
            grdSOTPICK1X.ActiveRow = grdSOTPICK1X.Rows(0)
        End If
    End Sub

    Private Function CreateSalesOrderInvoice(ByVal PICK_NO As String, ByVal SelectedShipViaCode As String) As Boolean

        Dim drSOTPICK1 As DataRow = Nothing
        Dim drSOTPICK1X As DataRow = Nothing
        Try
            dst.Tables("SOTINVH1").Rows.Clear()
            dst.Tables("SOTINVH2").Rows.Clear()
            dst.Tables("SOTINVH9").Rows.Clear()
            dst.Tables("SOTINVHM").Rows.Clear()
            dst.Tables("SOTSHIP1").Rows.Clear()

            Dim RFIXMSG As Boolean = False
            drSOTPICK1 = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
            Dim ORDR_NO As String = drSOTPICK1.Item("ORDR_NO")
            Dim drSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
            Dim drSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}' AND CUST_ADDR_TYPE = 'ST'")(0)
            drSOTPICK1X = dst.Tables("SOTPICK1X").Select($"PICK_NO = '{PICK_NO}'")(0)

            ' Currently, the Ecommerce records do not have a shipping record; therefore, we will make one
            Dim SHIP_BOL_NO As String = drSOTPICK1.Item("SHIP_BOL_NO") & String.Empty
            Dim drSOTSHIP1 As DataRow = Nothing

            If SHIP_BOL_NO.Length > 0 Then
                Fill_Records("SOTSHIP1", SHIP_BOL_NO)
            End If

            If dst.Tables("SOTSHIP1").Rows.Count = 0 Then
                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    SHIP_BOL_NO = ASCMAIN1.Next_Control_No("SHIP_BOL_NO")
                Else
                    SHIP_BOL_NO = ASCMAIN1.Next_Control_No("SOTSHIP1.SHIP_BOL_NO")
                End If
                drSOTSHIP1 = dst.Tables("SOTSHIP1").NewRow
                drSOTSHIP1.Item("SHIP_BOL_NO") = SHIP_BOL_NO
                'drSOTSHIP1.ITEM("SHIP_REF") = ""
                drSOTSHIP1.Item("SHIP_ADDR_TYPE") = drSOTORDR5.Item("CUST_ADDR_TYPE")
                drSOTSHIP1.Item("SHIP_ADDR_CODE") = drSOTORDR5.Item("CUST_ADDR_CODE")
                drSOTSHIP1.Item("ORDR_GROUP_NO") = drSOTORDR1.Item("ORDR_GROUP_NO")
                'drSOTSHIP1.ITEM("SHIP_PICK_PRINTED") = ""
                drSOTSHIP1.Item("PICK_BATCH_NO") = drSOTPICK1.Item("PICK_BATCH_NO")
                drSOTSHIP1.Item("SHIP_STATUS") = "P"
                'drSOTSHIP1.ITEM("SHIP_PULL_BY_STYLE") = ""
                'drSOTSHIP1.ITEM("SHIP_856_BATCH_NO") = ""
                drSOTSHIP1.Item("FRT_TERMS") = drSOTORDR1.Item("FRT_TERMS")
                drSOTSHIP1.Item("WHSE_CODE") = txtWHSE_CODE.Text
                'drSOTSHIP1.ITEM("SHIP_MANIFEST_NO") = ""
                'drSOTSHIP1.ITEM("SHIP_810_BATCH_NO") = ""
                drSOTSHIP1.Item("INIT_DATE") = DateTime.Now
                drSOTSHIP1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                drSOTSHIP1.Item("LAST_DATE") = DateTime.Now
                drSOTSHIP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                'drSOTSHIP1.ITEM("BILL_OF_LADING_NO") = ""
                'drSOTSHIP1.ITEM("REGISTER_XNO") = ""
                'drSOTSHIP1.ITEM("REASON_CODE") = ""
                'drSOTSHIP1.ITEM("SHIP_BOL_NO_REV") = ""
                drSOTSHIP1.Item("TERM_CODE") = drSOTORDR1.Item("TERM_CODE")
                drSOTSHIP1.Item("SREP_CODE") = drSOTORDR1.Item("SREP_CODE")
                drSOTSHIP1.Item("ORDR_DEPT") = drSOTORDR1.Item("ORDR_DEPT")
                'drSOTSHIP1.ITEM("SHIP_NOTES") = ""
                'drSOTSHIP1.ITEM("CUST_FACTOR_TRANS_IND") = ""
                'drSOTSHIP1.ITEM("SHIP_SEAL_NO") = ""
                'drSOTSHIP1.ITEM("SHIP_BOL_NO_ORIG") = ""
                'drSOTSHIP1.ITEM("SHIP_BOL_NO_SPLIT") = ""
                drSOTSHIP1.Item("SREP2_CODE") = drSOTORDR1.Item("SREP2_CODE")
                'drSOTSHIP1.ITEM("BOL_PRINTED") = ""
                'drSOTSHIP1.ITEM("SHIP_SPEC_INST") = ""
                'drSOTSHIP1.ITEM("FACTOR_TRANS_BATCH_LAST") = ""
                'drSOTSHIP1.ITEM("FACTOR_TRANS_LAST_OPER") = ""
                'drSOTSHIP1.ITEM("FACTOR_TRANS_LAST_DATE") = ""
                'drSOTSHIP1.ITEM("MASTER_SHIP_BOL_NO") = ""
                'drSOTSHIP1.ITEM("SHIP_940_BATCH_NO") = ""
                'drSOTSHIP1.ITEM("SHIP_753_IND") = ""
                'drSOTSHIP1.ITEM("HANDLING_TYPE") = ""
                'drSOTSHIP1.ITEM("HANDLING_UNITS") = ""
                'drSOTSHIP1.ITEM("GEN_IND") = ""
                'drSOTSHIP1.ITEM("GEN_XNO") = ""
                'drSOTSHIP1.ITEM("GEN_DATE") = ""
                'drSOTSHIP1.ITEM("DOCUMENTKEY") = ""
                'drSOTSHIP1.ITEM("THIRD_PARTY") = ""
                'drSOTSHIP1.ITEM("OPT_LINE1") = ""
                'drSOTSHIP1.ITEM("OPT_LINE2") = ""
                'drSOTSHIP1.ITEM("SHIP_DATE_PACKED") = ""
                'drSOTSHIP1.ITEM("LP_STATUS") = ""
                'drSOTSHIP1.ITEM("LP_XNO") = ""
                'drSOTSHIP1.ITEM("MASTER_BILL_OF_LADING_NO") = ""
                'drSOTSHIP1.ITEM("SHIP_TRAILER_NO") = ""
                'drSOTSHIP1.ITEM("SHIP_LOAD_NO") = ""
                'drSOTSHIP1.ITEM("SHIP_APPT_NO") = ""
                'drSOTSHIP1.ITEM("ORDR_PICK_TYPE") = ""
                'drSOTSHIP1.ITEM("SHIP_856_IND") = ""
                'drSOTSHIP1.ITEM("SHIP_810_IND") = ""
                'drSOTSHIP1.ITEM("LP_XMIT_DATE") = ""
                'drSOTSHIP1.ITEM("LP_CODE") = ""
                drSOTSHIP1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
                'drSOTSHIP1.ITEM("SHIP_CART_REQD") = ""
                'drSOTSHIP1.ITEM("EDI_856_CREATED") = ""
                'drSOTSHIP1.ITEM("EDI_810_CREATED") = ""
                'drSOTSHIP1.ITEM("SHIP_753_BATCH_NO") = ""
                'drSOTSHIP1.ITEM("SHIP_DATE_PLANNED") = ""
                'drSOTSHIP1.ITEM("SHIP_DATE_ROUTED") = ""
                'drSOTSHIP1.ITEM("SHIP_NOTES_3PL") = ""
                'drSOTSHIP1.ITEM("INSURED_VALUE") = ""
                'drSOTSHIP1.ITEM("INSURED_SHIPMENT") = ""
                'drSOTSHIP1.ITEM("EDI_LOAD_ID") = ""
                'drSOTSHIP1.ITEM("SHIP_WAVE_STATUS") = ""
                'drSOTSHIP1.ITEM("WAVE_NO") = ""
                'drSOTSHIP1.ITEM("BTB_BOL_NO") = ""
                dst.Tables("SOTSHIP1").Rows.Add(drSOTSHIP1)
                dst.Tables("SOTPICK1").Select($"PICK_NO = '{PICK_NO}'")(0).Item("SHIP_BOL_NO") = SHIP_BOL_NO
            Else
                drSOTSHIP1 = dst.Tables("SOTSHIP1").Rows(0)
            End If

            drSOTSHIP1.Item("SHIP_DATE_SHIPPED") = DateTime.Now.ToShortDateString
            drSOTSHIP1.Item("INV_DATE") = DateTime.Now.ToShortDateString
            drSOTSHIP1.Item("SHIP_DATE_RECEIVED") = DateTime.Now.ToShortDateString
            drSOTSHIP1.Item("SHIPPED_ACTUAL") = DateTime.Now.ToShortDateString
            drSOTSHIP1.Item("SHIP_TOTAL_WGT") = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", $"PICK_NO = '{PICK_NO}'") & String.Empty)
            drSOTSHIP1.Item("SHIP_CNT_CARTONS") = dst.Tables("SOTCART1").Rows.Count
            drSOTSHIP1.Item("SHIP_VIA_CODE") = SelectedShipViaCode
            drSOTSHIP1.Item("SHIP_STATUS") = "F"

            Dim SOCINVH1 As New TAC.SOCINVH1(dst)

            SOCINVH1.ProcessPickTicketsAndUpdateSalesDetails(DateTime.Now.ToShortDateString)

            ' Currently There are no back_orders for Ecommerce Sales Orders
            For Each drSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}'")
                Dim ORDR_QTY_OPEN As Int16 = Val(drSOTORDR2.Item("ORDR_QTY_OPEN") & String.Empty)
                Dim ORDR_QTY_PICK As Int16 = Val(drSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty)
                Dim ORDR_QTY_SHIP As Int16 = Val(drSOTORDR2.Item("ORDR_QTY_SHIP") & String.Empty)
                Dim ORDR_QTY_CANC As Int16 = Val(drSOTORDR2.Item("ORDR_QTY_CANC") & String.Empty)

                If ORDR_QTY_OPEN > 0 Then
                    drSOTORDR2.Item("ORDR_QTY_OPEN") = 0
                    drSOTORDR2.Item("ORDR_QTY_CANC") = Val(drSOTORDR2.Item("ORDR_QTY_CANC") & String.Empty) + ORDR_QTY_OPEN
                End If
            Next

            ' Record event where the Ship via was changed
            If drSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty <> SelectedShipViaCode Then
                Dim drTATEVNT1 As DataRow = dst.Tables("TATEVNT1").Rows.Add
                drTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
                drTATEVNT1.Item("TABLE_KEY") = ORDR_NO
                drTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
                drTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                drTATEVNT1.Item("EVENT_TYPE") = "SHPMTC"
                drTATEVNT1.Item("EVENT_DESC") = $"Ship Via was changed from {drSOTORDR1.Item("SHIP_VIA_CODE")} to {SelectedShipViaCode}"
                drTATEVNT1.Item("EVENT_KEY") = ""
                drTATEVNT1.Item("FORM_NAME") = "SOFSHIPE"

                drSOTORDR1.Item("SHIP_VIA_CODE") = SelectedShipViaCode
            End If

            If dst.Tables("SOTPICK2").Select("PICK_QTY > 0 AND PICK_QTY_CONF < PICK_QTY", "").Length > 0 Then
                Dim drTATEVNT1 As DataRow = dst.Tables("TATEVNT1").Rows.Add
                drTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
                drTATEVNT1.Item("TABLE_KEY") = ORDR_NO
                drTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
                drTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                drTATEVNT1.Item("EVENT_TYPE") = "SHSHP"
                drTATEVNT1.Item("EVENT_DESC") = "User chose to short ship Ecommerce order."
                drTATEVNT1.Item("EVENT_KEY") = ""
                drTATEVNT1.Item("FORM_NAME") = "SOFSHIPE"
            End If

            Dim CUST_FACTOR_TRANS_IND As String = "0"

            ' Log factoring change
            If Val(drSOTSHIP1.Item("CUST_FACTOR_TRANS_IND") & String.Empty) <> Val(CUST_FACTOR_TRANS_IND) Then
                Dim drTATEVNT1 As DataRow = dst.Tables("TATEVNT1").Rows.Add
                drTATEVNT1.Item("TABLE_NAME") = "SOTSHIP1"
                drTATEVNT1.Item("TABLE_KEY") = SHIP_BOL_NO
                drTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
                drTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                drTATEVNT1.Item("EVENT_TYPE") = "SHPFAC"
                drTATEVNT1.Item("EVENT_DESC") = "Factor Setting was changed from " _
                    & IIf(Val(drSOTSHIP1.Item("CUST_FACTOR_TRANS_IND") & String.Empty) = 1, "True", "False") & " to " & IIf(Val(CUST_FACTOR_TRANS_IND) = 1, "True", "False")
                drTATEVNT1.Item("EVENT_KEY") = ""
                drTATEVNT1.Item("FORM_NAME") = "SOFSHIPE"
            End If

            ' If nothing to Invoice then get out of here. All items are back ordered.
            allItemsOnBackOrder = dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}' AND ISNULL(PICK_QTY_CONF, 0) > 0 ").Length = 0

            If allItemsOnBackOrder Then
                drSOTPICK1.Item("PICK_STATUS") = "C"
            Else
                SOCINVH1.CreateInvoices(SHIP_BOL_NO, RFIXMSG)
                For Each drSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("")
                    drSOTINVH1.Item("INV_PRINTED") = DateTime.Now
                    drSOTINVH1.Item("ORDR_WEB_IND") = "1"
                Next
            End If

            Dim PICK_BATCH_NO As String = drSOTTRCK1.Item("PICK_BATCH_NO") & String.Empty
            drSOTPICK1X.Item("SELECTED") = "1"
            drSOTPICK1.Item("PICK_STATUS") = "F"

            Try
                BeginTrans()
                Update_Record_TDA("SOTORDR1")
                Update_Record_TDA("SOTORDR2")
                Update_Record_TDA("SOTORDR5")

                Update_Record_TDA("SOTPICK1")
                Update_Record_TDA("SOTPICK2")
                Update_Record_TDA("SOTSHIP1")

                Update_Record_TDA("SOTINVH1")
                Update_Record_TDA("SOTINVH2")
                Update_Record_TDA("SOTINVH9")
                Update_Record_TDA("SOTINVHM")

                Update_Record_TDA("ARTOPEN1")
                Update_Record_TDA("TATEVNT1")

                Update_Record_TDA("SOTCART1")
                Update_Record_TDA("SOTCART2")

                ' Needed when Back Orders exist and nothing is getting shipped.
                ' Also cancel a sales order when all items are cancelled.
                ASCMAIN1.sql = $"Begin Declare Cursor C1 Is
                                 Select ORDR_NO
                                    , Sum (NVL(ORDR_QTY_OPEN, 0)) ORDR_QTY_OPEN
                                    , Sum (NVL(ORDR_QTY_PICK, 0)) ORDR_QTY_PICK
                                    , Sum (NVL(ORDR_QTY_SHIP, 0)) ORDR_QTY_SHIP
                                  From SOTORDR2 Where ORDR_NO = :PARM1 group by ORDR_NO;
                                 Begin For R1 In C1 Loop
                                    Update SOTORDR1 Set
                                      ORDR_STATUS = 
                                        CASE WHEN R1.ORDR_QTY_OPEN > 0 THEN 'O'
                                             ELSE CASE WHEN R1.ORDR_QTY_PICK > 0 THEN 'P'
                                                       ELSE CASE WHEN R1.ORDR_QTY_SHIP > 0 THEN 'F'
                                                                 ELSE 'C' END END END
                                    where ORDR_NO = R1.ORDR_NO;
                                 End Loop; End;
                                End;"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {ORDR_NO})


                'CHANGE SOTPICK0.PICK_STATUS FROM P -> K (IN PACK)
                If dst.Tables("SOTPICK1X").Select("SELECTED = '1'").Length = 1 Then
                    ASCMAIN1.sql = $"UPDATE SOTPICK0 SET PICK_BATCH_STATUS = 'K' WHERE PICK_BATCH_STATUS IN ('N', 'K') AND TRUCK_NO = :PARM1 AND PICK_BATCH_NO = :PARM2"
                    Dim numrows As Int16 = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {HFs("TRUCK_NO"), PICK_BATCH_NO})
                    If numrows <> 1 Then
                        Throw New Exception("Unable to properly Update SOTPICK0.PICK_BATCH_STATUS")
                    End If
                End If

                'WHEN PACKING STATION IS "COMPLETE" WITH TRANSFERRING ORDERS IN PICK INTO PACK
                If dst.Tables("SOTPICK1X").Select("SELECTED = '1'").Length = dst.Tables("SOTPICK1X").Rows.Count Then
                    ASCMAIN1.Progress("Finalizing Truck", "SOTPICK0")
                    'CHANGE SOTPICK0.PICK_STATUS FROM K -> F (FINISHED)
                    ASCMAIN1.sql = "UPDATE SOTPICK0 SET PICK_BATCH_STATUS = 'F' WHERE PICK_BATCH_STATUS = 'K' AND PICK_BATCH_NO = :PARM1"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {PICK_BATCH_NO})

                    'CLEAR PICK BATCH FROM TRUCK
                    ASCMAIN1.Progress("Finalizing Truck", "SOTPICK1")
                    ASCMAIN1.sql = "UPDATE SOTTRCK1 SET PICK_BATCH_NO = NULL WHERE PICK_BATCH_NO = :PARM1"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {PICK_BATCH_NO})

                    'CLEAR PICK NO FROM TOTE (ALL TRUCK TYPES)
                    ASCMAIN1.Progress("Finalizing Truck", "SOTTOTE1")
                    'ASCMAIN1.sql = "UPDATE SOTTOTE1 SET PICK_NO = NULL WHERE TRUCK_NO = :PARM1"
                    'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {HFs("TRUCK_NO")})
                    If drSOTTRCK1.Item("TRUCK_TYPE") = "R" Then
                        ASCMAIN1.sql = $"UPDATE SOTTOTE1 SET PICK_NO = NULL, SLOT_NO = NULL, TRUCK_NO = NULL, INIT_OPER = '{ASCMAIN1.USER_ID}' WHERE PICK_NO IN (SELECT PICK_NO FROM SOTPICK1 WHERE PICK_BATCH_NO = :PARM1)"
                    Else
                        ASCMAIN1.sql = "UPDATE SOTTOTE1 SET PICK_NO = NULL WHERE PICK_NO IN (SELECT PICK_NO FROM SOTPICK1 WHERE PICK_BATCH_NO = :PARM1)"
                    End If
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {PICK_BATCH_NO})

                    'IF TRUCK IS A CUSTOM TRUCK, CHANGE TRUCK TYPE FROM X -> R
                    ASCMAIN1.Progress("Finalizing Truck", "SOTTRCK1")
                    If drSOTTRCK1.Item("TRUCK_TYPE") = "X" Then
                        ASCMAIN1.sql = "UPDATE SOTTRCK1 SET TRUCK_TYPE = 'R' WHERE TRUCK_TYPE = 'X' AND TRUCK_NO = :PARM1 AND DC_CODE = :PARM2"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {HFs("TRUCK_NO"), HFs("DC_CODE")})
                    End If

                    'DELETE CUSTOM TOTES THAT BELONG TO THE CUSTOM TRUCK
                    ASCMAIN1.Progress("Finalizing Truck", "SOTTOTE1")
                    ASCMAIN1.sql = "DELETE FROM SOTTOTE1 WHERE TRUCK_NO = :PARM1 AND DC_CODE = :PARM2 AND TOTE_TYPE = 'X'"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {HFs("TRUCK_NO"), HFs("DC_CODE")})
                End If

                CommitTrans()

            Catch ex As Exception
                Rollback(ex.Message)
                drSOTPICK1.Item("PICK_STATUS") = "P"
                drSOTPICK1X.Item("SELECTED") = "0"
                Return False
            End Try

            txtPickNo.Text = PICK_NO
            Return True

        Catch ex As Exception
            drSOTPICK1.Item("PICK_STATUS") = "P"
            drSOTPICK1X.Item("SELECTED") = "0"
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Sends the Scanned Bar Code to the Appropriate Control based on the Current Processing State
    ''' </summary>
    ''' <param name="scannedData"></param>
    ''' <remarks></remarks>
    Private Sub ProcessScannedData(ByVal scannedData As String)

        Static dataReceived As String

        dataReceived += scannedData
        If InStr(dataReceived, Chr(13), CompareMethod.Text) = 0 Then
            Exit Sub
        End If

        Dim sender As Object = Nothing
        Dim e As New System.Windows.Forms.KeyEventArgs(Keys.Enter)

        ' Trim Off line feeds
        dataReceived = Replace(dataReceived, Chr(10), String.Empty)
        dataReceived = Replace(dataReceived, Chr(13), String.Empty)

        If ScreenMode Then
            sender = txtTRUCK_NO
            txtTRUCK_NO.Clear()
            txtTRUCK_NO.Focus()
            txtTRUCK_NO.Text = dataReceived
        Else
            sender = txtTOTE_NO
            txtTOTE_NO.Clear()
            txtTOTE_NO.Focus()
            txtTOTE_NO.Text = dataReceived
        End If

        txt_KeyDown(sender, e)
        dataReceived = String.Empty

    End Sub

    Private Function AddCarton(ByVal PICK_NO As String) As String

        Dim drSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow

        Dim CART_NO_ctl As String = ""
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            CART_NO_ctl = ASCMAIN1.Next_Control_No("CART_NO")
        Else
            CART_NO_ctl = ASCMAIN1.Next_Control_No("SOTCART1.CART_NO")
        End If
        Dim CART_NO As String = TAC.SOCMAIN1.UPC(Me, CART_NO_ctl, "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))

        drSOTCART1.Item("CART_NO") = CART_NO
        'drSOTCART1.Item("CART_FREIGHT") = ""
        drSOTCART1.Item("CART_PACKER") = ASCMAIN1.USER_ID
        drSOTCART1.Item("CART_PACKED") = DateTime.Now
        drSOTCART1.Item("CART_SHIPPED") = DateTime.Now.ToShortDateString
        drSOTCART1.Item("PICK_NO") = PICK_NO
        drSOTCART1.Item("CART_TOTAL_UNITS") = Val(dst.Tables("SOTPICK2").Compute("SUM(PICK_QTY_CONF)", $"PICK_NO = '{PICK_NO}'") & String.Empty)
        'drSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = ""
        'drSOTCART1.Item("CART_TOTAL_WGT_CALC") = ""
        'drSOTCART1.Item("CART_TRACKING_NO") = ""
        drSOTCART1.Item("CART_SEQ") = Val(dst.Tables("SOTCART1").Compute("MAX(CART_SEQ)", $"PICK_NO = '{PICK_NO}'") & String.Empty) + 1
        'drSOTCART1.Item("CART_MEMO") = ""
        'drSOTCART1.Item("CART_TYPE") = ""
        drSOTCART1.Item("PACKAGING_TYPE") = defaultPACKAGING_TYPE
        drSOTCART1.Item("PKG_CODE") = defaultPKG_CODE
        'drSOTCART1.Item("PKG_L") = ""
        'drSOTCART1.Item("PKG_W") = ""
        'drSOTCART1.Item("PKG_H") = ""
        'drSOTCART1.Item("CART_TOTAL_UNITS_REL") = ""
        'drSOTCART1.Item("PALLET_NO") = ""
        'drSOTCART1.Item("MULTIPO_IND") = ""
        dst.Tables("SOTCART1").Rows.Add(drSOTCART1)

        Return CART_NO
    End Function

#End Region

#Region "Carrier Procedures"

    Private Sub RequestRates(ByVal PICK_NO As String)

        Try
            Me.Cursor = Cursors.WaitCursor
            Dim drWHTSHPC4 As DataRow = Nothing
            Dim drWHTSHPCA As DataRow = Nothing

            Dim CARRIER_SURCHARGE_PERC As Int16 = 0
            Dim FRT_PER_SALES_HOLD As Int16 = 0
            Dim CARRIER_PPA_TYPE As String = "L"
            Dim CARRIER_SURCHARGE_BASE As String = "L"
            Dim drSOTCARR1 As DataRow = Nothing

            For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select($"PICK_NO = '{PICK_NO}'", "", DataViewRowState.CurrentRows)
                If Val(drSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") & String.Empty) <= 0 Then
                    MessageBox.Show("All Cartons must have a weight.", "Request Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                ElseIf Val(drSOTCART1.Item("PKG_L") & String.Empty) <= 0 OrElse Val(drSOTCART1.Item("PKG_W") & String.Empty) <= 0 OrElse Val(drSOTCART1.Item("PKG_H") & String.Empty) <= 0 Then
                    MessageBox.Show("All Cartons must have a Length, Width and Height.", "Request Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            Next

            ASCMAIN1.Progress("Request Carrier Rates")

            dst.Tables("WHTSHPC4").Rows.Clear()

            Dim rUPSList(1) As WHCSHIP1.RateList
            Dim rFEDEXList(1) As WHCSHIP1.RateList
            Dim rUPSFreightList(1) As WHCSHIP1.RateList
            Dim rUSPSList(1) As WHCSHIP1.RateList

            'Me.Cursor = Cursors.WaitCursor
            'ASCMAIN1.Progress("-", "USPS")
            'rUSPSList = GetUSPSRates()

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("-", "UPS")
            rUPSList = GetUpsRates(PICK_NO)

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("-", "FedEx")
            rFEDEXList = GetFedExRates(PICK_NO)

            Me.Cursor = Cursors.WaitCursor

            If rUSPSList Is Nothing Then
                ReDim rUSPSList(1)
            End If

            If rUPSList Is Nothing Then
                ReDim rUPSList(1)
            End If

            If rFEDEXList Is Nothing Then
                ReDim rFEDEXList(1)
            End If

            If rUPSFreightList Is Nothing Then
                ReDim rUPSFreightList(1)
            End If

            Dim selected As Boolean = False
            Dim CARRIER_CODE As String = String.Empty

            For iCtr As Int16 = 1 To 4
                Dim freightShipment As String = " and ISNULL(FREIGHT_SHIPMENT, '0') = '0'"
                Dim rList(1) As WHCSHIP1.RateList

                Select Case iCtr
                    Case 1
                        rList = rUPSList
                        CARRIER_CODE = "UPS"
                        ASCMAIN1.Progress("-", "UPS")

                    Case 2
                        rList = rFEDEXList
                        CARRIER_CODE = "FEDEX"
                        ASCMAIN1.Progress("-", "FedEx")

                    Case 3
                        rList = rUPSFreightList
                        CARRIER_CODE = "UPS"
                        ASCMAIN1.Progress("-", "UPS Freight")
                        freightShipment = " and ISNULL(FREIGHT_SHIPMENT, '0') = '1'"

                    Case 4
                        rList = rUSPSList
                        CARRIER_CODE = "USPS"
                        ASCMAIN1.Progress("-", "USPS")

                End Select

                If rList IsNot Nothing Then
                    For iLoop As Integer = 0 To rList.Count - 1
                        With rList(iLoop)
                            If .ServiceType Is Nothing OrElse (.ServiceType = 0 AndAlso .ServiceTypeDescription.Length = 0) Then
                                Continue For
                            End If

                            ' Display only those services that are mapped to ship vias
                            If dst.Tables("SOTSVIA1").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND CARRIER_PROD_CODE = '" & .ServiceType & "' AND SHIP_VIA_STATUS = 'A'" & freightShipment).Length = 0 Then
                                Continue For
                            End If

                            'CARRIER_SURCHARGE_PERC
                            drSOTCARR1 = dst.Tables("SOTCARR1").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)

                            If drSOTCARR1.Table.Columns.Contains("CARRIER_SURCHARGE_PERC") Then
                                CARRIER_SURCHARGE_PERC = Val(drSOTCARR1.Item("CARRIER_SURCHARGE_PERC") & String.Empty)
                            End If

                            If drSOTCARR1.Table.Columns.Contains("FRT_PER_SALES_HOLD") Then
                                FRT_PER_SALES_HOLD = Val(drSOTCARR1.Item("FRT_PER_SALES_HOLD") & String.Empty)
                            End If

                            If drSOTCARR1.Table.Columns.Contains("CARRIER_PPA_TYPE") Then
                                CARRIER_PPA_TYPE = drSOTCARR1.Item("CARRIER_PPA_TYPE") & String.Empty
                                ' If not set then set to List
                                If CARRIER_PPA_TYPE.Length = 0 Then
                                    CARRIER_PPA_TYPE = "L"
                                End If
                            End If

                            If drSOTCARR1.Table.Columns.Contains("CARRIER_SURCHARGE_BASE") Then
                                CARRIER_SURCHARGE_BASE = drSOTCARR1.Item("CARRIER_SURCHARGE_BASE") & String.Empty
                                ' If not set then set to List
                                If CARRIER_SURCHARGE_BASE.Length = 0 Then
                                    CARRIER_SURCHARGE_BASE = "L"
                                End If
                            End If

                            drWHTSHPC4 = dst.Tables("WHTSHPC4").NewRow
                            drWHTSHPC4.Item("SHIP_CNTL_NO") = "*"

                            Select Case CARRIER_CODE
                                Case "UPS"
                                    drWHTSHPC4.Item("SERVICE_INDEX") = iLoop + (100 * iCtr)
                                Case "FEDEX"
                                    drWHTSHPC4.Item("SERVICE_INDEX") = iLoop + 200
                                Case "USPS"
                                    drWHTSHPC4.Item("SERVICE_INDEX") = iLoop + 300
                            End Select

                            drWHTSHPC4.Item("SERVICE_TYPE_DESC") = .ServiceTypeDescription
                            drWHTSHPC4.Item("DISCLAIMER") = .Disclaimer

                            drWHTSHPC4.Item("CARRIER_CODE") = CARRIER_CODE

                            If dst.Tables("SOTSVIA1").Select($"CARRIER_CODE = '{CARRIER_CODE}' AND CARRIER_PROD_CODE = '{ .ServiceType}' AND SHIP_VIA_STATUS = 'A' AND SHIP_VIA_CODE = '{txtSHIP_VIA_CODE.Text}' {freightShipment}").Length > 0 Then
                                drWHTSHPC4.Item("SHIP_VIA_CODE") = dst.Tables("SOTSVIA1").Select($"CARRIER_CODE = '{CARRIER_CODE}' AND CARRIER_PROD_CODE = '{ .ServiceType}' AND SHIP_VIA_STATUS = 'A' AND SHIP_VIA_CODE = '{txtSHIP_VIA_CODE.Text}' {freightShipment}")(0).Item("SHIP_VIA_CODE")
                            Else
                                drWHTSHPC4.Item("SHIP_VIA_CODE") = dst.Tables("SOTSVIA1").Select($"CARRIER_CODE = '{CARRIER_CODE}' AND CARRIER_PROD_CODE = '{ .ServiceType}' AND SHIP_VIA_STATUS = 'A' {freightShipment}")(0).Item("SHIP_VIA_CODE")
                            End If

                            If (.AccountNetCharge & String.Empty <> "") Then
                                drWHTSHPC4.Item("ACCT_NET_CHARGE") = Convert.ToDecimal(.AccountNetCharge)
                            Else
                                drWHTSHPC4.Item("ACCT_NET_CHARGE") = Convert.ToDecimal(.ListNetCharge)
                            End If

                            drWHTSHPC4.Item("SERVICE_TYPE") = .ServiceType
                            drWHTSHPC4.Item("SURCHARGE") = 0
                            drWHTSHPC4.Item("DELIVERY_TIME") = .DeliveryTime
                            drWHTSHPC4.Item("LIST_NET_CHARGE") = .ListNetCharge
                            If .TransitTime <> "" Then
                                drWHTSHPC4.Item("TRANSIT_TIME") = .TransitTime
                            End If

                            drWHTSHPC4.Item("CARRIER_CODE") = CARRIER_CODE

                            Select Case CARRIER_PPA_TYPE
                                Case "F" ' None
                                    drWHTSHPC4.Item("CUSTOMER_BASE_CHARGE") = 0
                                Case "N" ' Negotiated
                                    drWHTSHPC4.Item("CUSTOMER_BASE_CHARGE") = drWHTSHPC4.Item("ACCT_NET_CHARGE")
                                Case "L" ' List
                                    drWHTSHPC4.Item("CUSTOMER_BASE_CHARGE") = drWHTSHPC4.Item("LIST_NET_CHARGE")
                                Case Else
                                    ' If not set then use  List
                                    drWHTSHPC4.Item("CUSTOMER_BASE_CHARGE") = drWHTSHPC4.Item("LIST_NET_CHARGE")
                            End Select

                            ' Additional Surcharge based off List
                            If CARRIER_SURCHARGE_PERC > 0 Then
                                Select Case CARRIER_SURCHARGE_BASE
                                    Case "N" ' Negotiated
                                        drWHTSHPC4.Item("SURCHARGE") = Val(drWHTSHPC4.Item("ACCT_NET_CHARGE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                                    Case "L" ' List
                                        drWHTSHPC4.Item("SURCHARGE") = Val(drWHTSHPC4.Item("LIST_NET_CHARGE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                                    Case Else
                                        ' If not set then use  List
                                        drWHTSHPC4.Item("SURCHARGE") = Val(drWHTSHPC4.Item("LIST_NET_CHARGE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                                End Select
                            End If

                            dst.Tables("WHTSHPC4").Rows.Add(drWHTSHPC4)

                        End With
                    Next
                End If
            Next

            grdWHTSHPC4.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
            grdWHTSHPC4.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

            Try
                Sort_grdColumns(grdWHTSHPC4, "TOTAL_CHARGE")
            Catch ex As Exception
            End Try

            With grdWHTSHPC4.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With

        Catch ex As Exception
            MessageBox.Show("Get Rates Error: " & ex.Message, "Get Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

    End Sub

    Private Function GetUpsRates(ByVal PICK_NO As String) As WHCSHIP1.RateList()
        Try

            Dim rList(1) As WHCSHIP1.RateList

            If dst.Tables("SOTCARR1").Select("CARRIER_CODE = 'UPS'").Length = 0 Then
                Return Nothing
            End If

            If dst.Tables("SOTCARR3").Select("CARRIER_CODE = 'UPS'").Length = 0 Then
                Return Nothing
            End If

            Dim CARRIER_CODE As String = "UPS"
            Dim drSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Select("CARRIER_CODE = 'UPS'")(0)
            Dim drSOTCARR3 As DataRow = Nothing
            Dim CUST_CODE As String = dst.Tables("SOTORDR1").Rows(0).Item("CUST_CODE") & String.Empty
            Dim ORDR_NO As String = dst.Tables("SOTORDR1").Rows(0).Item("ORDR_NO") & String.Empty

            If dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "'").Length > 0 Then
                drSOTCARR3 = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "'")(0)
            ElseIf dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = DIVISION_CODE AND DIVISION_CODE = '" & ASCMAIN1.CLIENT & "'").Length > 0 Then
                drSOTCARR3 = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = DIVISION_CODE AND DIVISION_CODE = '" & ASCMAIN1.CLIENT & "'")(0)
            Else
                Return Nothing
            End If

            Dim drICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(txtWHSE_CODE.Text)
            If drICTWHSE1 Is Nothing Then
                Return Nothing
            End If

            clsShip.Reset()
            clsShip.Service = WHCSHIP1.ServiceProviders.UPS

            ' Credentials
            With clsShip
                .Server = drSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
                .UserId = drSOTCARR3.Item("SHIPPER_ID") & String.Empty
                .Password = drSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
                .AccountNumber = drSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
                .UPSAccessKey = drSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                .FedexMeterNumber = drSOTCARR3.Item("METER_NUMBER") & String.Empty
                .FedexDeveloperKey = drSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                .LabelStockType = (drSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim
            End With

            clsShip.RequestedServiceType = ServiceTypes.stUnspecified
            clsShip.UPSPickupType = UpsratesPickupTypes.ptDailyPickup
            clsShip.CustomerType = UpsratesCustomerTypes.ccRetail
            clsShip.ShipDate = DateTime.Now.ToShortDateString

            Dim listSeqNo As New List(Of Int16)

            For Each drSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select($"PICK_NO = '{PICK_NO}'")
                For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'", "CART_SEQ, CART_TOTAL_WGT_ACTUAL DESC", DataViewRowState.CurrentRows)

                    If listSeqNo.Contains(drSOTCART1.Item("CART_SEQ")) Then
                        Continue For
                    End If

                    listSeqNo.Add(drSOTCART1.Item("CART_SEQ"))

                    Dim pkgDetail As New PackageDetail

                    pkgDetail.Id = StrReverse(StrReverse(drSOTCART1.Item("CART_NO") & String.Empty).Substring(0, 8))
                    pkgDetail.Weight = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "CART_SEQ = " & drSOTCART1.Item("CART_SEQ")) & String.Empty)

                    ' Convert Pounds to Ounces
                    pkgDetail.Weight *= 16

                    If pkgDetail.Weight = "0" Then
                        pkgDetail.Weight = "16.0"
                    End If

                    pkgDetail.PackagingType = CType(Val(drSOTCART1.Item("PACKAGING_TYPE") & String.Empty), UpsratesPickupTypes)
                    pkgDetail.Length = Val(drSOTCART1.Item("PKG_L") & String.Empty)
                    pkgDetail.Width = Val(drSOTCART1.Item("PKG_W") & String.Empty)
                    pkgDetail.Height = Val(drSOTCART1.Item("PKG_H") & String.Empty)

                    pkgDetail.InsuredValue = 0 ' numInsureValue.Value * -1 / dst.Tables("SOTCART1").Rows.Count
                    clsShip.PackageDetailList.Add(pkgDetail)
                Next
            Next

            With clsShip.Sender
                .Company = (drICTWHSE1.Item("WHSE_DESC") & String.Empty).ToString.Trim
                .FirstName = (drICTWHSE1.Item("WHSE_CONTACT") & String.Empty).ToString.Trim
                .MiddleInitial = String.Empty
                .LastName = String.Empty
                .Address1 = (drICTWHSE1.Item("WHSE_ADDR1") & String.Empty).ToString.Trim
                .Address2 = (drICTWHSE1.Item("WHSE_ADDR2") & String.Empty).ToString.Trim
                .City = (drICTWHSE1.Item("WHSE_CITY") & String.Empty).ToString.Trim
                .State = (drICTWHSE1.Item("WHSE_STATE") & String.Empty).ToString.Trim
                .ZipCode = (drICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                .CountryCode = (drICTWHSE1.Item("WHSE_COUNTRY") & String.Empty).ToString.Trim.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                If .CountryCode = "CAN" Then .CountryCode = "CA"
                .Phone = (drICTWHSE1.Item("WHSE_PHONE") & String.Empty).ToString.Trim
            End With

            Dim drSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}'")(0)

            With clsShip.Recipient
                If drSOTORDR5.Item("CUST_CONTACT") & String.Empty <> String.Empty Then
                    .FirstName = drSOTORDR5.Item("CUST_CONTACT") & String.Empty
                Else
                    .FirstName = drSOTORDR5.Item("CUST_NAME") & String.Empty
                End If

                .MiddleInitial = String.Empty
                .LastName = String.Empty

                .Address1 = drSOTORDR5.Item("CUST_ADDR1") & String.Empty
                .Address2 = drSOTORDR5.Item("CUST_ADDR2") & String.Empty
                .City = drSOTORDR5.Item("CUST_CITY") & String.Empty
                .State = drSOTORDR5.Item("CUST_STATE") & String.Empty
                .ZipCode = drSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty
                .CountryCode = (drSOTORDR5.Item("CUST_COUNTRY") & String.Empty).ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                If .CountryCode = "CAN" Then .CountryCode = "CA"

                .Company = .FirstName
                .Phone = drSOTORDR5.Item("CUST_PHONE") & String.Empty

                If .Phone.Trim.Length = 0 Then
                    .Phone = clsShip.Sender.Phone
                End If

                If .Phone.Trim.Length = 0 Then
                    .Phone = "1234567890"
                End If

                .IsResidental = dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}'")(0).Item("IS_RESIDENTAL") & String.Empty = "1"
                .IsPOBox = dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}'")(0).Item("IS_PO_BOX") & String.Empty = "1"
            End With

            clsShip.RatesTotalValue = 0
            clsShip.ShipmentSpecialServices = 0
            clsShip.SignatureRequired = False
            rList = clsShip.GetUPSRatesList()

            If clsShip.LastError.Length > 0 Then
                MessageBox.Show("UPS Error: " & clsShip.LastError, "UPS Rates Error")
            End If

            If rList Is Nothing Then
                ReDim rList(1)
            End If

            Return rList

        Catch ex As Exception
            MessageBox.Show("The following error occurred getting UPS Rates: " & ex.Message, "Get UPS Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try

    End Function

    Private Function GetFedExRates(ByVal PICK_NO As String) As WHCSHIP1.RateList()
        Try

            Dim rList(1) As WHCSHIP1.RateList

            If dst.Tables("SOTCARR1").Select("CARRIER_CODE = 'FEDEX'").Length = 0 Then
                Return Nothing
            End If

            If dst.Tables("SOTCARR3").Select("CARRIER_CODE = 'FEDEX'").Length = 0 Then
                Return Nothing
            End If

            Dim CARRIER_CODE As String = "FEDEX"
            Dim drSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Select("CARRIER_CODE = 'FEDEX'")(0)
            Dim drSOTCARR3 As DataRow = Nothing
            Dim CUST_CODE As String = dst.Tables("SOTORDR1").Rows(0).Item("CUST_CODE") & String.Empty
            Dim ORDR_NO As String = dst.Tables("SOTORDR1").Rows(0).Item("ORDR_NO") & String.Empty

            If dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "' AND CARRIER_PROD_CODE IS NULL").Length > 0 Then
                drSOTCARR3 = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "' AND CARRIER_PROD_CODE IS NULL")(0)
            ElseIf dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = DIVISION_CODE AND DIVISION_CODE = '" & ASCMAIN1.CLIENT & "'").Length > 0 Then
                drSOTCARR3 = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = DIVISION_CODE AND DIVISION_CODE = '" & ASCMAIN1.CLIENT & "'")(0)
            Else
                Return Nothing
            End If

            Dim drICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(txtWHSE_CODE.Text)
            If drICTWHSE1 Is Nothing Then
                Return Nothing
            End If

            clsShip.Reset()
            clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpress

            ' Credentials
            With clsShip
                .Server = drSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
                .UserId = drSOTCARR3.Item("SHIPPER_ID") & String.Empty
                .Password = drSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
                .AccountNumber = drSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
                .UPSAccessKey = drSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                .FedexMeterNumber = drSOTCARR3.Item("METER_NUMBER") & String.Empty
                .FedexDeveloperKey = drSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                .LabelStockType = (drSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim
            End With

            clsShip.RequestedServiceType = ServiceTypes.stUnspecified
            clsShip.UPSPickupType = UpsratesPickupTypes.ptDailyPickup
            clsShip.CustomerType = UpsratesCustomerTypes.ccRetail
            clsShip.ShipDate = DateTime.Now.ToShortDateString

            Dim listSeqNo As New List(Of Int16)

            For Each drSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select($"PICK_NO = '{PICK_NO}'")
                For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'", "CART_SEQ, CART_TOTAL_WGT_ACTUAL DESC", DataViewRowState.CurrentRows)

                    If listSeqNo.Contains(drSOTCART1.Item("CART_SEQ")) Then
                        Continue For
                    End If

                    listSeqNo.Add(drSOTCART1.Item("CART_SEQ"))

                    Dim pkgDetail As New PackageDetail

                    pkgDetail.Id = drSOTCART1.Item("CART_NO").ToString.Substring(2)
                    pkgDetail.Weight = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "CART_SEQ = " & drSOTCART1.Item("CART_SEQ")) & String.Empty)
                    ' Convert to Ounces
                    pkgDetail.Weight *= 16

                    If pkgDetail.Weight = "0" Then
                        pkgDetail.Weight = "16.0"
                    End If

                    pkgDetail.PackagingType = CType(Val(drSOTCART1.Item("PACKAGING_TYPE") & String.Empty), UpsratesPickupTypes)
                    pkgDetail.Length = Val(drSOTCART1.Item("PKG_L") & String.Empty)
                    pkgDetail.Width = Val(drSOTCART1.Item("PKG_W") & String.Empty)
                    pkgDetail.Height = Val(drSOTCART1.Item("PKG_H") & String.Empty)

                    ' Can have either Insured or Declared not Both
                    pkgDetail.InsuredValue = 0
                    clsShip.PackageDetailList.Add(pkgDetail)
                Next
            Next

            With clsShip.Sender
                .Company = (drICTWHSE1.Item("WHSE_DESC") & String.Empty).ToString.Trim
                .FirstName = (drICTWHSE1.Item("WHSE_CONTACT") & String.Empty).ToString.Trim
                .MiddleInitial = String.Empty
                .LastName = String.Empty
                .Address1 = (drICTWHSE1.Item("WHSE_ADDR1") & String.Empty).ToString.Trim
                .Address2 = (drICTWHSE1.Item("WHSE_ADDR2") & String.Empty).ToString.Trim
                .City = (drICTWHSE1.Item("WHSE_CITY") & String.Empty).ToString.Trim
                .State = (drICTWHSE1.Item("WHSE_STATE") & String.Empty).ToString.Trim
                .ZipCode = (drICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                .CountryCode = (drICTWHSE1.Item("WHSE_COUNTRY") & String.Empty).ToString.Trim.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                .Phone = (drICTWHSE1.Item("WHSE_PHONE") & String.Empty).ToString.Trim
            End With

            Dim drSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}'")(0)

            With clsShip.Recipient
                If drSOTORDR5.Item("CUST_CONTACT") & String.Empty <> String.Empty Then
                    .FirstName = drSOTORDR5.Item("CUST_CONTACT") & String.Empty
                Else
                    .FirstName = drSOTORDR5.Item("CUST_NAME") & String.Empty
                End If

                .MiddleInitial = String.Empty
                .LastName = String.Empty

                .Address1 = drSOTORDR5.Item("CUST_ADDR1") & String.Empty
                .Address2 = drSOTORDR5.Item("CUST_ADDR2") & String.Empty
                .City = drSOTORDR5.Item("CUST_CITY") & String.Empty
                .State = drSOTORDR5.Item("CUST_STATE") & String.Empty
                .ZipCode = drSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty
                .CountryCode = (drSOTORDR5.Item("CUST_COUNTRY") & String.Empty).ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                If .CountryCode = "CAN" Then .CountryCode = "CA"

                .Company = .FirstName
                .Phone = drSOTORDR5.Item("CUST_PHONE") & String.Empty

                If .Phone.Trim.Length = 0 Then
                    .Phone = clsShip.Sender.Phone
                End If

                If .Phone.Trim.Length = 0 Then
                    .Phone = "1234567890"
                End If

                .IsResidental = dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}'")(0).Item("IS_RESIDENTAL") & String.Empty = "1"
                .IsPOBox = dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}'")(0).Item("IS_PO_BOX") & String.Empty = "1"
            End With

            clsShip.ShipmentSpecialServices = 0
            clsShip.SignatureRequired = False
            rList = clsShip.GetFedExRatesList()

            If clsShip.LastError.Length > 0 Then
                MessageBox.Show("FedEx Error: " & clsShip.LastError, "FedEx Rates Error")
            End If

            If rList Is Nothing Then
                ReDim rList(1)
            End If

            Return rList

        Catch ex As Exception
            MessageBox.Show("The following error occurred getting FedEx Rates: " & ex.Message, "Get FedEx Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try

    End Function

    Private Function RequestShippingLabel(ByVal INV_NO As String,
                                          ByRef ErrorMessage As String,
                                          ByVal PreScreenForErrorsOnly As Boolean) As Boolean
        ErrorMessage = String.Empty
        Dim ShippingLabels As New List(Of String)
        Dim drSOTPICK1 As DataRow = Nothing
        Dim drSOTORDR1 As DataRow = Nothing
        Dim drSOTSVIA1 As DataRow = Nothing
        Dim drSOTCARR1 As DataRow = Nothing
        Dim drSOTORDR5 As DataRow = Nothing

        Dim SHIP_VIA_CODE As String = txtSHIP_VIA_CODE.Text
        Dim CARRIER_CODE As String = String.Empty
        Dim CARRIER_PROD_CODE As String = String.Empty

        Dim ORDR_NO As String = String.Empty
        Dim CUST_CODE As String = String.Empty
        Dim ORDR_NO_WEB As String = String.Empty
        Dim ORDR_CUST_PO As String = String.Empty
        Dim SHIP_BOL_NO As String = String.Empty

        Dim SHIP_PACKAGE_NO As Int64 = 0
        Dim pkgId As Int64 = 0
        Dim isPitneyBowes As Boolean = False

        Dim CARRIER_SURCHARGE_PERC As Int16 = 0
        Dim CARRIER_SURCHARGE_BASE As String = "L"

        Dim FRT_PER_SALES_HOLD As Int16 = 0
        Dim CARRIER_PPA_TYPE As String = "L"

        RequestShippingLabel = True

        Dim PICK_NO As String = dst.Tables("SOTPICK1").Select($"TOTE_NO = '{txtTOTE_NO.Text}'")(0).Item("PICK_NO")

        Try
            drSOTSVIA1 = dst.Tables("SOTSVIA1").Rows.Find(SHIP_VIA_CODE)
            CARRIER_CODE = drSOTSVIA1.Item("CARRIER_CODE") & String.Empty
            CARRIER_PROD_CODE = drSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty

            drSOTCARR1 = dst.Tables("SOTCARR1").Rows.Find(CARRIER_CODE)
            drSOTPICK1 = dst.Tables("SOTPICK1").Select($"PICK_NO = '{PICK_NO}'")(0)
            SHIP_BOL_NO = drSOTPICK1.Item("SHIP_BOL_NO") & String.Empty

            ORDR_NO = drSOTPICK1.Item("ORDR_NO") & String.Empty
            drSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
            CUST_CODE = drSOTORDR1.Item("CUST_CODE") & String.Empty
            ORDR_NO_WEB = drSOTORDR1.Item("ORDR_NO_WEB") & String.Empty
            ORDR_CUST_PO = drSOTORDR1.Item("ORDR_CUST_PO") & String.Empty

            Fill_Records("ARTCUSTS", CUST_CODE)
            drSOTORDR5 = dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}' AND CUST_ADDR_TYPE = 'ST'")(0)

            ' Logic added 3/18/2017 for Regency
            If drSOTCARR1.Table.Columns.Contains("CARRIER_SURCHARGE_PERC") Then
                CARRIER_SURCHARGE_PERC = Val(drSOTCARR1.Item("CARRIER_SURCHARGE_PERC") & String.Empty)
            End If

            If drSOTCARR1.Table.Columns.Contains("CARRIER_SURCHARGE_BASE") Then
                CARRIER_SURCHARGE_BASE = drSOTCARR1.Item("CARRIER_SURCHARGE_BASE") & String.Empty
                ' If not set then set to List
                If CARRIER_SURCHARGE_BASE.Length = 0 Then
                    CARRIER_SURCHARGE_BASE = "L"
                End If
            End If

            If drSOTCARR1.Table.Columns.Contains("FRT_PER_SALES_HOLD") Then
                FRT_PER_SALES_HOLD = Val(drSOTCARR1.Item("FRT_PER_SALES_HOLD") & String.Empty)
            End If

            If drSOTCARR1.Table.Columns.Contains("CARRIER_PPA_TYPE") Then
                CARRIER_PPA_TYPE = drSOTCARR1.Item("CARRIER_PPA_TYPE") & String.Empty
                ' If not set then set tp list
                If CARRIER_PPA_TYPE.Length = 0 Then
                    CARRIER_PPA_TYPE = "L"
                End If
            End If

            ' Load and Validate Carrier/Ship Method
            Dim drSOTCARR2 As DataRow = LookUp("SOTCARR2", New String() {CARRIER_CODE, CARRIER_PROD_CODE})
            If drSOTCARR2 Is Nothing Then
                ErrorMessage = "Invalid or missing Carrier / Ship Method combination for shipping label request"
                Return False
            End If

            ' Credentials
            Dim drSOTCARR3 As DataRow = Nothing

            ' SHIPPER_DIVISION_CODE holds a customer code,  SHIPPER_ID
            If dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "' AND CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "'").Length > 0 Then
                drSOTCARR3 = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "' AND CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "'")(0)
            ElseIf dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "' AND CARRIER_PROD_CODE IS NULL").Length > 0 Then
                drSOTCARR3 = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "' AND CARRIER_PROD_CODE IS NULL")(0)
            ElseIf dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = DIVISION_CODE AND DIVISION_CODE = '" & ASCMAIN1.CLIENT & "'").Length > 0 Then
                drSOTCARR3 = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = DIVISION_CODE AND DIVISION_CODE = '" & ASCMAIN1.CLIENT & "'")(0)
            End If

            If drSOTCARR3 Is Nothing Then
                ErrorMessage = "Cannot determine the Carrier Account to use for the shipping label request"
                Return False
            End If

            Dim DIVISION_CODE As String = drSOTCARR3.Item("DIVISION_CODE") & String.Empty
            Dim CARRIER_ACCOUNT_NO As String = drSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty

            ' See if we need to use a different From Address.
            Dim drSOTCARR5 As DataRow = Nothing
            ASCMAIN1.sql = "CARRIER_CODE = '" & CARRIER_CODE & "' and DIVISION_CODE = '" & DIVISION_CODE & "' and CARRIER_ACCOUNT_NO = '" & CARRIER_ACCOUNT_NO & "' and CUST_CODE = '" & CUST_CODE & "'"
            If dst.Tables("SOTCARR5").Select(ASCMAIN1.sql).Length > 0 Then
                drSOTCARR5 = dst.Tables("SOTCARR5").Select(ASCMAIN1.sql)(0)
            Else
                ASCMAIN1.sql = "CARRIER_CODE = '" & CARRIER_CODE & "' and DIVISION_CODE = '" & DIVISION_CODE & "' and CARRIER_ACCOUNT_NO = '" & CARRIER_ACCOUNT_NO & "' and CUST_CODE = '" & "*" & "'"
                If dst.Tables("SOTCARR5").Select(ASCMAIN1.sql).Length > 0 Then
                    drSOTCARR5 = dst.Tables("SOTCARR5").Select(ASCMAIN1.sql)(0)
                End If
            End If

            Dim ShippingLabelDirectory As String = (drSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim
            If ASCMAIN1.Running_in_VS Then
                ShippingLabelDirectory = ShippingLabelDirectory.Replace("R:\", "C:\").Replace("S:\", "C:\")
            End If
            Dim PROVIDER_TYPE As String = (drSOTCARR1.Item("PROVIDER_TYPE") & String.Empty).ToString.Trim

            If drSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty = String.Empty Then
                ErrorMessage = "Invalid or missing Carrier Account Number for shipping label request"
                Return False
            End If

            Try
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

            Dim drICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(Absx1.txtFor("WHSE_CODE").Text)
            If drICTWHSE1 Is Nothing Then
                ErrorMessage = "Invalid or missing Warehouse"
                Return False
            End If

            Dim CUST_NAME As String = (drSOTORDR5.Item("CUST_NAME") & String.Empty).ToString.Trim
            Dim CUST_CONTACT As String = (drSOTORDR5.Item("CUST_CONTACT") & String.Empty).ToString.Trim
            Dim CUST_ADDR1 As String = (drSOTORDR5.Item("CUST_ADDR1") & String.Empty).ToString.Trim
            Dim CUST_ADDR2 As String = (drSOTORDR5.Item("CUST_ADDR2") & String.Empty).ToString.Trim
            Dim CUST_CITY As String = (drSOTORDR5.Item("CUST_CITY") & String.Empty).ToString.Trim
            Dim CUST_STATE As String = (drSOTORDR5.Item("CUST_STATE") & String.Empty).ToString.Trim
            Dim CUST_COUNTRY As String = (drSOTORDR5.Item("CUST_COUNTRY") & String.Empty).ToString.Trim
            Dim CUST_ZIP_CODE As String = (drSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty).ToString.Trim
            Dim CUST_PHONE As String = (drSOTORDR5.Item("CUST_PHONE") & String.Empty).ToString.Trim

            If CUST_ADDR1.Length = 0 AndAlso CUST_ADDR2.Length = 0 Then
                ErrorMessage = "Invalid or missing Ship To Street Address"
                Return False
            ElseIf Not CUST_COUNTRY.StartsWith("US") AndAlso (CUST_CITY.Length = 0 OrElse CUST_ZIP_CODE.Length = 0) Then
                ErrorMessage = "Invalid or missing International Ship To City and/or Zip Code"
                Return False
            ElseIf CUST_CITY.Length = 0 OrElse CUST_STATE.Length = 0 OrElse CUST_ZIP_CODE.Length = 0 Then
                ErrorMessage = "Invalid or missing Ship To City, State or Zip Code"
                Return False
            ElseIf CUST_COUNTRY.Length = 0 Then
                Dim drTATSTATE As DataRow = dst.Tables("TATSTATE").Rows.Find(CUST_STATE)
                If drTATSTATE IsNot Nothing Then
                    CUST_COUNTRY = "US"
                Else
                    ErrorMessage = "Invalid or missing Country Code"
                    Return False
                End If
            End If

            ' 02/25/2020 - Evaluate Cartons to make sure carton dimensions are sent to UPS/ FedEx
            For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select($"PICK_NO = '{PICK_NO}'")
                Dim CART_NO As String = drSOTCART1.Item("CART_NO") & String.Empty
                Dim PACKAGING_TYPE As String = drSOTCART1.Item("PACKAGING_TYPE") & String.Empty
                Dim PKG_CODE As String = drSOTCART1.Item("PKG_CODE") & String.Empty

                ' Make sure FedEx and UPS shipments have box dimensions.
                Dim LENGTH As Decimal = Val(drSOTCART1.Item("PKG_L") & String.Empty)
                Dim WIDTH As Decimal = Val(drSOTCART1.Item("PKG_W") & String.Empty)
                Dim HEIGHT As Decimal = Val(drSOTCART1.Item("PKG_H") & String.Empty)

                If LENGTH <= 0 OrElse WIDTH <= 0 OrElse HEIGHT <= 0 Then
                    ErrorMessage &= vbCr & "Carton " & CART_NO & " has invalid dimensions."
                End If

                If dst.Tables("SOTCART2").Select("CART_NO = '" & CART_NO & "'").Length = 0 Then
                    ErrorMessage &= vbCr & "Carton " & CART_NO & " does not have any assigned products."
                End If

                If PACKAGING_TYPE = defaultPACKAGING_TYPE Then
                    If dst.Tables("WHTPKGM1").Rows.Find(PKG_CODE) Is Nothing Then
                        ErrorMessage &= vbCr & "Carton " & CART_NO & " has an invalid package code."
                    End If
                End If
            Next

            If ErrorMessage.Length > 0 Then
                Return True
            End If

            If PreScreenForErrorsOnly Then Return True

            '*******************************************************************************

            Dim isInternationalShipment As Boolean = False
            Dim fedexSmartPost As Int16 = 26

            Dim FRT_TERMS As String = drSOTORDR1.Item("FRT_TERMS") & String.Empty
            Dim PPA_FREIGHT As Decimal = 0
            Dim OUR_FREIGHT As Decimal = 0

            dst.Tables("WHTSHPC1").Rows.Clear()
            dst.Tables("WHTSHPC2").Rows.Clear()
            dst.Tables("WHTSHPC5").Rows.Clear()
            dst.Tables("WHTSHPCG").Rows.Clear()
            dst.Tables("WHTSHPCS").Rows.Clear()
            dst.Tables("WHTSHPCC").Rows.Clear()
            dst.Tables("WHTSHPCP").Rows.Clear()

            Dim SHIP_CNTL_NO As String = String.Empty
            clsShip.Reset()

            ' Credentials
            clsShip.Server = drSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
            clsShip.UserId = drSOTCARR3.Item("SHIPPER_ID") & String.Empty
            clsShip.Password = drSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
            clsShip.AccountNumber = drSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
            clsShip.UPSAccessKey = drSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.FedexMeterNumber = drSOTCARR3.Item("METER_NUMBER") & String.Empty
            clsShip.FedexDeveloperKey = drSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.LabelStockType = (drSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

            Dim drWHTSHPC1 As DataRow = Nothing
            Dim drWHTSHPC2 As DataRow = Nothing
            Dim drWHTSHPC5 As DataRow = Nothing
            Dim drWHTSHPCG As DataRow = Nothing

            drWHTSHPC1 = dst.Tables("WHTSHPC1").NewRow
            SHIP_CNTL_NO = ASCMAIN1.Next_Control_No("WHTSHPC1.SHIP_CNTL_NO")
            drWHTSHPC1.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
            drWHTSHPC1.Item("CARRIER_CODE") = CARRIER_CODE
            drWHTSHPC1.Item("CARRIER_PROD_CODE") = CARRIER_PROD_CODE
            drWHTSHPC1.Item("CARRIER_ACCOUNT_NO") = drSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
            dst.Tables("WHTSHPC1").Rows.Add(drWHTSHPC1)

            drWHTSHPC1.Item("STATUS") = "I"
            drWHTSHPC1.Item("ERROR_MSG") = String.Empty
            drWHTSHPC1.Item("SHIP_DATE") = DateTime.Now.ToString("MM/dd/yyyy")
            drWHTSHPC1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            drWHTSHPC1.Item("OPS_YYYYWW") = ASCMAIN1.CYW
            drWHTSHPC1.Item("CUST_CODE") = CUST_CODE
            drWHTSHPC1.Item("INIT_DATE") = DateTime.Now
            drWHTSHPC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            drWHTSHPC1.Item("LAST_DATE") = DateTime.Now
            drWHTSHPC1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            drWHTSHPC1.Item("MASTER_TRACKING_NO") = String.Empty
            drWHTSHPC1.Item("CUSTOMS_VALUE") = 0
            drWHTSHPC1.Item("SHIP_BOL_NO") = SHIP_BOL_NO
            drWHTSHPC1.Item("SHIP_VIA_CODE") = SHIP_VIA_CODE

            drWHTSHPC1.Item("INSURED_VALUE") = 0
            drWHTSHPC1.Item("INSURED_SHIPMENT") = "0" ' IIf(Absx1.chkFor("INSURED_SHIPMENT").Checked, "1", "0")

            ' Update the Key in these tables
            For Each tableName As String In New String() {"WHTSHPC4", "WHTSHPCA"}
                For Each dr As DataRow In dst.Tables(tableName).Select("")
                    dr.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                Next
            Next

            ' Sender Information
            With clsShip.Sender
                .Company = (drICTWHSE1.Item("WHSE_DESC") & String.Empty).ToString.Trim
                .Phone = (drICTWHSE1.Item("WHSE_PHONE") & String.Empty).ToString.Trim
                .FirstName = (drICTWHSE1.Item("WHSE_CONTACT") & String.Empty).ToString.Trim
                .MiddleInitial = String.Empty
                .LastName = String.Empty
                .Address1 = (drICTWHSE1.Item("WHSE_ADDR1") & String.Empty).ToString.Trim
                .Address2 = (drICTWHSE1.Item("WHSE_ADDR2") & String.Empty).ToString.Trim
                .City = (drICTWHSE1.Item("WHSE_CITY") & String.Empty).ToString.Trim
                .State = (drICTWHSE1.Item("WHSE_STATE") & String.Empty).ToString.Trim
                .ZipCode = (drICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                .CountryCode = (drICTWHSE1.Item("WHSE_COUNTRY") & String.Empty).ToString.Trim.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                If .CountryCode = "CAN" Then .CountryCode = "CA"

                drWHTSHPC5 = dst.Tables("WHTSHPC5").NewRow
                drWHTSHPC5.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                drWHTSHPC5.Item("SHIP_ADDR_TYPE") = "SF"
                drWHTSHPC5.Item("SHIP_FIRST_NAME") = .FirstName
                drWHTSHPC5.Item("SHIP_LAST_NAME") = .LastName
                drWHTSHPC5.Item("SHIP_MIDDLE_INIT") = .MiddleInitial
                drWHTSHPC5.Item("SHIP_PHONE") = .Phone
                drWHTSHPC5.Item("SHIP_FAX") = .Fax
                drWHTSHPC5.Item("SHIP_EMAIL") = .eMail
                drWHTSHPC5.Item("SHIP_COMPANY") = .Company
                drWHTSHPC5.Item("SHIP_ADDR1") = .Address1
                drWHTSHPC5.Item("SHIP_ADDR2") = .Address2
                drWHTSHPC5.Item("SHIP_CITY") = .City
                drWHTSHPC5.Item("SHIP_STATE") = .State
                drWHTSHPC5.Item("SHIP_ZIP_CODE") = .ZipCode
                drWHTSHPC5.Item("SHIP_COUNTRY_CODE") = .CountryCode
                drWHTSHPC5.Item("SHIP_RESIDENTIAL") = IIf(.IsResidental, "1", "0")
                drWHTSHPC5.Item("SHIP_PO_BOX") = IIf(.IsPOBox, "1", "0")
                dst.Tables("WHTSHPC5").Rows.Add(drWHTSHPC5)
            End With

            ' This is an override address that will print as the Ship From address in the upper left hand corner of the shipping label.
            If drSOTCARR5 IsNot Nothing Then
                With clsShip.Account
                    .Company = drSOTCARR5.Item("ACCOUNT_NAME") & String.Empty
                    .Phone = drSOTCARR5.Item("ACCOUNT_PHONE") & String.Empty

                    .FirstName = drSOTCARR5.Item("ACCOUNT_CONTACT") & String.Empty
                    .MiddleInitial = String.Empty
                    .LastName = String.Empty
                    .Address1 = drSOTCARR5.Item("ACCOUNT_ADDR1") & String.Empty
                    .Address2 = drSOTCARR5.Item("ACCOUNT_ADDR2") & String.Empty
                    .Address3 = drSOTCARR5.Item("ACCOUNT_ADDR3") & String.Empty
                    .City = drSOTCARR5.Item("ACCOUNT_CITY") & String.Empty
                    .State = drSOTCARR5.Item("ACCOUNT_STATE") & String.Empty
                    .ZipCode = drSOTCARR5.Item("ACCOUNT_ZIP_CODE") & String.Empty
                    .CountryCode = drSOTCARR5.Item("ACCOUNT_COUNTRY") & String.Empty
                    If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                    If .CountryCode = "USA" Then .CountryCode = "US"
                    If .CountryCode = "CAN" Then .CountryCode = "CA"

                    drWHTSHPC5 = dst.Tables("WHTSHPC5").NewRow
                    drWHTSHPC5.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    drWHTSHPC5.Item("SHIP_ADDR_TYPE") = "AC"
                    drWHTSHPC5.Item("SHIP_FIRST_NAME") = .FirstName
                    drWHTSHPC5.Item("SHIP_LAST_NAME") = .LastName
                    drWHTSHPC5.Item("SHIP_MIDDLE_INIT") = .MiddleInitial
                    drWHTSHPC5.Item("SHIP_PHONE") = .Phone
                    drWHTSHPC5.Item("SHIP_FAX") = .Fax
                    drWHTSHPC5.Item("SHIP_EMAIL") = .eMail
                    drWHTSHPC5.Item("SHIP_COMPANY") = .Company
                    drWHTSHPC5.Item("SHIP_ADDR1") = .Address1
                    drWHTSHPC5.Item("SHIP_ADDR2") = .Address2
                    drWHTSHPC5.Item("SHIP_CITY") = .City
                    drWHTSHPC5.Item("SHIP_STATE") = .State
                    drWHTSHPC5.Item("SHIP_ZIP_CODE") = .ZipCode
                    drWHTSHPC5.Item("SHIP_COUNTRY_CODE") = .CountryCode
                    drWHTSHPC5.Item("SHIP_RESIDENTIAL") = IIf(.IsResidental, "1", "0")
                    drWHTSHPC5.Item("SHIP_PO_BOX") = IIf(.IsPOBox, "1", "0")
                    dst.Tables("WHTSHPC5").Rows.Add(drWHTSHPC5)
                End With
            Else
                With clsShip.Account
                    .Company = String.Empty
                    .Phone = String.Empty

                    .FirstName = String.Empty
                    .MiddleInitial = String.Empty
                    .LastName = String.Empty
                    .Address1 = String.Empty
                    .Address2 = String.Empty
                    .Address3 = String.Empty
                    .City = String.Empty
                    .State = String.Empty
                    .ZipCode = String.Empty
                    .CountryCode = String.Empty
                End With
            End If

            ' Recipient
            With clsShip.Recipient
                .FirstName = IIf(CUST_CONTACT.Length > 0, CUST_CONTACT, CUST_NAME)
                .MiddleInitial = ""
                .LastName = ""

                .Address1 = CUST_ADDR1
                .Address2 = CUST_ADDR2
                .City = CUST_CITY
                .State = CUST_STATE
                .ZipCode = CUST_ZIP_CODE
                .CountryCode = CUST_COUNTRY.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                If .CountryCode = "CAN" Then .CountryCode = "CA"

                .Company = CUST_NAME

                If .Company = .FirstName Then
                    .FirstName = String.Empty
                End If
                'End If

                .Phone = CUST_PHONE

                If .Phone.Trim.Length = 0 Then
                    .Phone = clsShip.Sender.Phone
                End If

                If .Phone.Trim.Length = 0 Then
                    .Phone = "1234567890"
                End If

                ' Force FedEx Ground Home Delivery to residental
                If drSOTSVIA1 IsNot Nothing Then
                    'CARRIER_CODE = 'FEDEX' AND CARRIER_PROD_CODE = '16'
                    If drSOTSVIA1.Item("CARRIER_CODE") & String.Empty = "FEDEX" Then
                        If drSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty = "16" Then
                        End If
                    End If
                End If

                .IsResidental = dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}'")(0).Item("IS_RESIDENTAL") & String.Empty = "1"
                .IsPOBox = dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}'")(0).Item("IS_PO_BOX") & String.Empty = "1"

                drWHTSHPC5 = dst.Tables("WHTSHPC5").NewRow
                drWHTSHPC5.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                drWHTSHPC5.Item("SHIP_ADDR_TYPE") = "ST"
                drWHTSHPC5.Item("SHIP_FIRST_NAME") = .FirstName
                drWHTSHPC5.Item("SHIP_LAST_NAME") = .LastName
                drWHTSHPC5.Item("SHIP_MIDDLE_INIT") = .MiddleInitial
                drWHTSHPC5.Item("SHIP_PHONE") = .Phone
                drWHTSHPC5.Item("SHIP_FAX") = .Fax
                drWHTSHPC5.Item("SHIP_EMAIL") = .eMail
                drWHTSHPC5.Item("SHIP_COMPANY") = .Company
                drWHTSHPC5.Item("SHIP_ADDR1") = .Address1
                drWHTSHPC5.Item("SHIP_ADDR2") = .Address2
                drWHTSHPC5.Item("SHIP_CITY") = .City
                drWHTSHPC5.Item("SHIP_STATE") = .State
                drWHTSHPC5.Item("SHIP_ZIP_CODE") = .ZipCode
                drWHTSHPC5.Item("SHIP_COUNTRY_CODE") = .CountryCode
                drWHTSHPC5.Item("SHIP_RESIDENTIAL") = IIf(.IsResidental, "1", "0")
                drWHTSHPC5.Item("SHIP_PO_BOX") = IIf(.IsPOBox, "1", "0")
                dst.Tables("WHTSHPC5").Rows.Add(drWHTSHPC5)
            End With

            If dst.Tables("ARTCUSTS").Rows.Count = 1 Then
                Dim drARTCUSTS As DataRow = dst.Tables("ARTCUSTS").Rows(0)
                With clsShip.ReturnAddress
                    Select Case drSOTCARR1.Item("PROVIDER_TYPE") & String.Empty
                        Case "F" ' Federal Express
                            If drARTCUSTS.Item("FDX_RTN_SHIP_COMPANY") & String.Empty <> String.Empty Then
                                .Address1 = drARTCUSTS.Item("FDX_RTN_SHIP_ADDR1") & String.Empty
                                .Address2 = drARTCUSTS.Item("FDX_RTN_SHIP_ADDR2") & String.Empty
                                .Address3 = String.Empty
                                .City = drARTCUSTS.Item("FDX_RTN_SHIP_CITY") & String.Empty
                                .Company = drARTCUSTS.Item("FDX_RTN_SHIP_COMPANY") & String.Empty
                                .CountryCode = drARTCUSTS.Item("FDX_RTN_SHIP_COUNTRY_CODE") & String.Empty
                                .eMail = String.Empty
                                .Fax = String.Empty
                                .FirstName = String.Empty
                                .IsPOBox = False
                                .IsResidental = False
                                .LastName = String.Empty
                                .MiddleInitial = String.Empty
                                .Phone = drARTCUSTS.Item("FDX_RTN_SHIP_PHONE") & String.Empty
                                ' This is required
                                If .Phone.Length = 0 Then
                                    .Phone = clsShip.Sender.Phone
                                End If
                                .State = drARTCUSTS.Item("FDX_RTN_SHIP_STATE") & String.Empty
                                .ZipCode = drARTCUSTS.Item("FDX_RTN_SHIP_ZIP_CODE") & String.Empty
                            End If

                        Case "U" ' UPS
                            If drARTCUSTS.Item("UPS_RTN_SHIP_COMPANY") & String.Empty <> String.Empty Then
                                .Address1 = drARTCUSTS.Item("UPS_RTN_SHIP_ADDR1") & String.Empty
                                .Address2 = drARTCUSTS.Item("UPS_RTN_SHIP_ADDR2") & String.Empty
                                .Address3 = String.Empty
                                .City = drARTCUSTS.Item("UPS_RTN_SHIP_CITY") & String.Empty
                                .Company = drARTCUSTS.Item("UPS_RTN_SHIP_COMPANY") & String.Empty
                                .CountryCode = drARTCUSTS.Item("UPS_RTN_SHIP_COUNTRY_CODE") & String.Empty
                                .eMail = String.Empty
                                .Fax = String.Empty
                                .FirstName = String.Empty
                                .IsPOBox = False
                                .IsResidental = False
                                .LastName = String.Empty
                                .MiddleInitial = String.Empty
                                .Phone = drARTCUSTS.Item("UPS_RTN_SHIP_PHONE") & String.Empty
                                ' This is required
                                If .Phone.Length = 0 Then
                                    .Phone = clsShip.Sender.Phone
                                End If
                                .State = drARTCUSTS.Item("UPS_RTN_SHIP_STATE") & String.Empty
                                .ZipCode = drARTCUSTS.Item("UPS_RTN_SHIP_ZIP_CODE") & String.Empty
                            End If
                    End Select
                End With
            End If

            ' US Puerto Rico is considered International
            isInternationalShipment = (clsShip.Recipient.CountryCode <> clsShip.Sender.CountryCode) OrElse (clsShip.Recipient.CountryCode = "US" AndAlso clsShip.Recipient.State = "PR")

            Select Case PROVIDER_TYPE
                Case WHCSHIP1.ProviderTypeFedex
                    If Not isInternationalShipment Then
                        clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpress
                    Else
                        clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpressInternational
                    End If

                Case WHCSHIP1.ProviderTypeUPS
                    If Not isInternationalShipment Then
                        clsShip.Service = WHCSHIP1.ServiceProviders.UPS
                    Else
                        clsShip.Service = WHCSHIP1.ServiceProviders.UPSInternational
                    End If

                Case WHCSHIP1.ProviderTypeUSPS
                    clsShip.Service = WHCSHIP1.ServiceProviders.USPS
                    Select Case drSOTCARR1.Item("USPS_PARTNER") & String.Empty
                        Case "1"
                            clsShip.USPSPostageProvider = WHCSHIP1.USPSPostageProviders.Endicia
                        Case "2"
                            clsShip.USPSPostageProvider = WHCSHIP1.USPSPostageProviders.StampsCom
                        Case "3"
                            clsShip.USPSPostageProvider = WHCSHIP1.USPSPostageProviders.PitneyBowes
                            clsShip.PitneyBowesUniqueTransactionID = "USPS_PB_" & ASCMAIN1.Next_Control_No("SOTCARR1.USPS")
                            clsShip.PitneyBowesInductionPostalCode = (drICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                            isPitneyBowes = True
                        Case Else
                            Return False
                    End Select

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

            ' Set the References
            For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
                Select Case PROVIDER_TYPE
                    Case WHCSHIP1.ProviderTypeFedex
                        drSOTCART1.Item("REFERENCE1") = $"CR:{drSOTORDR1.Item("ORDR_NO_WEB")}"
                        drSOTCART1.Item("REFERENCE2") = $"IN:{INV_NO}"
                        drSOTCART1.Item("REFERENCE3") = ""

                    Case WHCSHIP1.ProviderTypeUPS
                        drSOTCART1.Item("REFERENCE1") = $"TN:{drSOTORDR1.Item("ORDR_NO_WEB")}"
                        drSOTCART1.Item("REFERENCE2") = $"IK:{INV_NO}"
                        drSOTCART1.Item("REFERENCE3") = ""

                End Select
            Next

            Refresh_Refs(CARRIER_CODE)

            For Each drSOTPICK1 In dst.Tables("SOTPICK1").Select($"PICK_NO = '{PICK_NO}'", "PICK_NO")

                PICK_NO = drSOTPICK1.Item("PICK_NO") & String.Empty
                ORDR_NO = drSOTPICK1.Item("ORDR_NO") & String.Empty
                SHIP_BOL_NO = drSOTPICK1.Item("SHIP_BOL_NO") & String.Empty
                ORDR_CUST_PO = drSOTPICK1.Item("ORDR_CUST_PO") & String.Empty
                PPA_FREIGHT = 0
                OUR_FREIGHT = 0
                itemList.Clear()

                '' Get the Invoice Number now so we can put it on the label
                'Dim INV_NO As String = String.Empty

                'If ASCMAIN1.CLIENT = "VAN" Then
                '    INV_NO = ASCMAIN1.Next_Control_No("INV_NO_01")
                'Else
                '    INV_NO = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
                'End If

                drSOTPICK1.Item("INV_NO") = INV_NO
                drSOTPICK1.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO

                For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'", "CART_SEQ, CART_TOTAL_WGT_ACTUAL DESC")
                    ' This is done to place multi pick tickets into one carton
                    Dim CART_SEQ As Int32 = drSOTCART1.Item("CART_SEQ")
                    If cartSequenceNos.Contains(CART_SEQ) Then
                        Continue For
                    End If
                    cartSequenceNos.Add(CART_SEQ)

                    Dim PACKAGING_TYPE As String = drSOTCART1.Item("PACKAGING_TYPE") & String.Empty
                    Dim PKG_CODE As String = drSOTCART1.Item("PKG_CODE") & String.Empty
                    Dim drWHTPKGM1 As DataRow = dst.Tables("WHTPKGM1").Rows.Find(PKG_CODE) ' LookUp("WHTPKGM1", PKG_CODE)
                    pkgId = CART_SEQ ' (Val(StrReverse(StrReverse(rowSOTCART1.Item("CART_NO").ToString).Substring(0, 8))))

                    Dim shipPackageDetail As New PackageDetail
                    With shipPackageDetail
                        .PackagingType = Val(PACKAGING_TYPE)

                        ' This is done to place multi pick tickets into one carton. Need combined weight 
                        If ASCMAIN1.CLIENT = "RGI" Then
                            .Weight = Val(drSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") & String.Empty)
                        Else
                            .Weight = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "CART_SEQ = " & CART_SEQ) & String.Empty)
                        End If
                        If .Weight = 0 Then
                            .Weight = 1
                        End If

                        '*************************************
                        '        Convert to Ounces
                        '*************************************
                        .Weight = Convert.ToInt16(.Weight * 16)
                        ' Take what is in the grid
                        .Length = Val(drSOTCART1.Item("PKG_L") & String.Empty)
                        .Width = Val(drSOTCART1.Item("PKG_W") & String.Empty)
                        .Height = Val(drSOTCART1.Item("PKG_H") & String.Empty)

                        Dim reference As String = String.Empty
                        Dim refCount As Int16 = 0

                        Select Case PROVIDER_TYPE
                            Case WHCSHIP1.ProviderTypeFedex
                                ' Fedex allows up to 3 References

                                If ASCMAIN1.CLIENT = "VAN" Then
                                    If (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 0 Then
                                        reference &= "; " & (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    End If

                                    If (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 0 Then
                                        reference &= "; " & (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    End If

                                    If (drSOTCART1.Item("REFERENCE3") & String.Empty).ToString.Trim.Length > 0 Then
                                        reference &= "; " & (drSOTCART1.Item("REFERENCE3") & String.Empty).ToString
                                    End If

                                    refCount = 5
                                End If

                                If refCount < 3 _
                                    AndAlso (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 2 _
                                    AndAlso (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Substring(2, 1) = ":" Then
                                    reference &= "; " & (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If refCount < 3 _
                                    AndAlso (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 2 _
                                    AndAlso (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Substring(2, 1) = ":" Then
                                    reference &= "; " & (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If refCount < 3 _
                                    AndAlso (drSOTCART1.Item("REFERENCE3") & String.Empty).ToString.Trim.Length > 2 _
                                    AndAlso (drSOTCART1.Item("REFERENCE3") & String.Empty).ToString.Substring(2, 1) = ":" Then
                                    reference &= "; " & (drSOTCART1.Item("REFERENCE3") & String.Empty).ToString
                                    refCount += 1
                                End If

                                ' This is done because some customers want specific information on the label
                                If reference.Length > 0 Then
                                    refCount = 5
                                End If


                                ' Fedex allows up to 3 References
                                If (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If (drSOTCART1.Item("REFERENCE3") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & (drSOTCART1.Item("REFERENCE3") & String.Empty).ToString
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

                                If reference.Length > 0 Then
                                    reference = reference.Substring(1).Trim
                                End If

                            Case WHCSHIP1.ProviderTypeUPS
                                ' Ups allows up to 2 References

                                If ASCMAIN1.CLIENT = "VAN" Then
                                    If (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 0 Then
                                        reference &= "; " & (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    End If

                                    If (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 0 Then
                                        reference &= "; " & (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    End If

                                    refCount = 5
                                End If

                                If refCount < 2 _
                                    AndAlso (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 2 _
                                    AndAlso (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Substring(2, 1) = ":" Then
                                    reference &= "; " & (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If refCount < 2 _
                                    AndAlso (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 2 _
                                    AndAlso (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Substring(2, 1) = ":" Then
                                    reference &= "; " & (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    refCount += 1
                                End If

                                ' This is done because some customers want specific information on the label
                                If reference.Length > 0 Then
                                    refCount = 5
                                End If

                                If (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 2 Then
                                    reference &= "; CR:" & (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    refCount += 1
                                Else
                                    If (drSOTPICK1.Item("CUST_STORE_NO") & String.Empty).ToString.Trim <> String.Empty AndAlso refCount < 2 Then
                                        reference &= "; ST:" & drSOTPICK1.Item("CUST_STORE_NO") & String.Empty
                                        refCount += 1
                                    End If
                                End If

                                If (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 2 Then
                                    reference &= "; CR:" & (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    refCount += 1
                                Else
                                    If (drSOTPICK1.Item("ORDR_CUST_PO") & String.Empty).ToString.Trim <> String.Empty AndAlso refCount < 2 Then
                                        reference &= "; PO:" & drSOTPICK1.Item("ORDR_CUST_PO") & String.Empty
                                        refCount += 1
                                    End If
                                End If

                                If reference.Length > 0 Then
                                    reference = reference.Substring(1).Trim
                                End If

                        End Select

                        If reference.Length > 0 Then
                            If reference.StartsWith(";") Then
                                reference = reference.Substring(1).Trim
                            End If

                            If Not reference.EndsWith(";") Then
                                reference &= ";"
                            End If
                        End If

                        .Reference = reference
                        .Id = pkgId.ToString("D8")
                    End With

                    clsShip.PackageDetailList.Add(shipPackageDetail)

                    drWHTSHPC2 = dst.Tables("WHTSHPC2").NewRow
                    drWHTSHPC2.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    drWHTSHPC2.Item("SHIP_PACKAGE_NO") = pkgId
                    drWHTSHPC2.Item("HEIGHT") = shipPackageDetail.Height
                    drWHTSHPC2.Item("INSURED_VALUE") = 0
                    drWHTSHPC2.Item("LENGTH") = shipPackageDetail.Length
                    drWHTSHPC2.Item("NET_CHARGE") = 0
                    drWHTSHPC2.Item("PACKAGING_TYPE") = Val(shipPackageDetail.PackagingType)
                    drWHTSHPC2.Item("TOTAL_DISCOUNT") = 0
                    drWHTSHPC2.Item("TOTAL_SURCHARGES") = 0
                    drWHTSHPC2.Item("TRACKING_NUMBER") = String.Empty
                    drWHTSHPC2.Item("WEIGHT") = Convert.ToInt16(shipPackageDetail.Weight)
                    drWHTSHPC2.Item("WIDTH") = shipPackageDetail.Width
                    drWHTSHPC2.Item("TRACKING_NO") = String.Empty

                    drWHTSHPC2.Item("CUST_REF") = ORDR_CUST_PO
                    drWHTSHPC2.Item("INV_BOL_NO") = SHIP_BOL_NO
                    drWHTSHPC2.Item("CART_NO") = drSOTCART1.Item("CART_NO") & String.Empty
                    drWHTSHPC2.Item("INV_NO") = INV_NO
                    drWHTSHPC2.Item("PO_ORDER_NO") = String.Empty
                    drWHTSHPC2.Item("DEPT_NO") = (drSOTPICK1.Item("ORDR_DEPT") & String.Empty).ToString.Trim

                    dst.Tables("WHTSHPC2").Rows.Add(drWHTSHPC2)
                Next

                If isInternationalShipment Then
                    ' Set the Customs value
                    clsShip.TotalCustomsValue = Val(drSOTPICK1.Item("PICK_AMT_CONF") & String.Empty)

                    For Each drSOTCART2 As DataRow In dst.Tables("SOTCART2").Select("PICK_NO = '" & PICK_NO & "'")
                        Dim STYLE_CODE As String = drSOTCART2.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = drSOTCART2.Item("COLOR_CODE")

                        If itemList.Contains(STYLE_CODE) Then Continue For

                        itemList.Add(STYLE_CODE)

                        Dim drICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        ' Just in case a non item is permitted in the shipment
                        If drICTSTYL1 Is Nothing Then Continue For

                        Dim CommodityDetail As New CommodityDetail
                        CommodityDetail.Description = drICTSTYL1.Item("STYLE_DESC") & String.Empty

                        Dim NumberOfPieces As Int32 = Val(dst.Tables("SOTCART2").Compute("SUM(QTY_PACKED)", "STYLE_CODE = '" & STYLE_CODE & "' and PICK_NO = '" & PICK_NO & "'") & String.Empty)

                        CommodityDetail.NumberOfPieces = NumberOfPieces
                        CommodityDetail.Quantity = NumberOfPieces
                        CommodityDetail.QuantityUnit = "EA"

                        Dim pickUnitPrice As Decimal = Val(dst.Tables("SOTPICK2").Compute("MAX(PICK_UNIT_PRICE)", "PICK_NO = '" & PICK_NO & "' AND STYLE_CODE = '" & STYLE_CODE & "'") & String.Empty)
                        CommodityDetail.UnitPrice = pickUnitPrice

                        CommodityDetail.Weight = Val(drICTSTYL1.Item("STYLE_WEIGHT") & String.Empty) ' Leave as pounds
                        CommodityDetail.Manufacturer = (drICTSTYL1.Item("COUNTRY_CODE") & String.Empty).ToString.ToUpper.Trim ' "US" '
                        If CommodityDetail.Manufacturer.Length = 0 Then
                            CommodityDetail.Manufacturer = "US"
                        End If

                        clsShip.CommodityDetailList.Add(CommodityDetail)

                        Dim drWHTSHPCC As DataRow = dst.Tables("WHTSHPCC").NewRow
                        drWHTSHPCC.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                        drWHTSHPCC.Item("COMMODITY_LNO") = COMMODITY_LNO
                        COMMODITY_LNO += 1
                        drWHTSHPCC.Item("COMMODITY_DESC") = CommodityDetail.Description
                        drWHTSHPCC.Item("NUM_PIECES") = CommodityDetail.NumberOfPieces
                        drWHTSHPCC.Item("MANUFACTURER") = CommodityDetail.Manufacturer
                        drWHTSHPCC.Item("HARMONIZED_CODE") = String.Empty
                        drWHTSHPCC.Item("WEIGHT") = CommodityDetail.Weight
                        drWHTSHPCC.Item("QUANTITY") = CommodityDetail.Quantity
                        drWHTSHPCC.Item("QUANTITY_UOM") = CommodityDetail.QuantityUnit
                        drWHTSHPCC.Item("UNIT_PRICE") = CommodityDetail.UnitPrice
                        dst.Tables("WHTSHPCC").Rows.Add(drWHTSHPCC)
                    Next
                End If
            Next  ' This is where the For Sotpick1, for sotcart1, for sotcart2 should end 

            clsShip.TotalCustomsValue = 0
            clsShip.SignatureRequired = False

            ' Shipping Method
            If isInternationalShipment Then
                If drSOTSVIA1.Item("CARRIER_PROD_CODE_INTL") & String.Empty <> String.Empty Then
                    clsShip.RequestedServiceType = Val(drSOTSVIA1.Item("CARRIER_PROD_CODE_INTL") & String.Empty)
                Else
                    clsShip.RequestedServiceType = Val(drSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty)
                End If
            Else
                clsShip.RequestedServiceType = Val(drSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty)
            End If

            If clsShip.RequestedServiceType = fedexSmartPost Then
                clsShip.FedexSmartPost.HubId = drSOTCARR3.Item("FEDEX_HUB_ID") & String.Empty
            End If

            clsShip.USPSEndorsement = WHCSHIP1.USPSEndorsements.NoServiceSelected

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

            Dim drWHTSHPCP As DataRow
            drWHTSHPCP = dst.Tables("WHTSHPCP").NewRow
            drWHTSHPCP("SHIP_CNTL_NO") = SHIP_CNTL_NO
            drWHTSHPCP("PAYOR_TYPE") = "S"
            drWHTSHPCP("PAYOR_ACCT_NO") = clsShip.PayorContact.AccountNumber & String.Empty
            drWHTSHPCP("PAYOR_COUNTRY") = clsShip.PayorContact.CountryCode & String.Empty
            dst.Tables("WHTSHPCP").Rows.Add(drWHTSHPCP)

            ' Payor of the Duties
            clsShip.DutiesPayor = TPayorTypes.ptSender
            If isInternationalShipment Then
                clsShip.DutiesPayor = clsShip.Payor
                clsShip.DutiesPayorContact.AccountNumber = clsShip.PayorContact.AccountNumber
                clsShip.DutiesPayorContact.CountryCode = clsShip.PayorContact.CountryCode
                clsShip.DutiesPayorContact.ZipCode = clsShip.PayorContact.ZipCode
            End If

            drWHTSHPCP = dst.Tables("WHTSHPCP").NewRow
            drWHTSHPCP("SHIP_CNTL_NO") = SHIP_CNTL_NO
            drWHTSHPCP("PAYOR_TYPE") = "D"
            drWHTSHPCP("PAYOR_ACCT_NO") = clsShip.DutiesPayorContact.AccountNumber & String.Empty
            drWHTSHPCP("PAYOR_COUNTRY") = clsShip.DutiesPayorContact.CountryCode & String.Empty
            dst.Tables("WHTSHPCP").Rows.Add(drWHTSHPCP)

            With clsShip
                .EzshipLabelImage = EzshipLabelImageTypes.itZPL
                .ShippingLabelDirectory = ShippingLabelDirectory
                .ShippingLabelPrefix = SHIP_CNTL_NO
                .ShipDate = DateTime.Now.ToString("yyyy-MM-dd")
            End With

            Try
                BeginTrans()
                Update_Record_TDA("WHTSHPC1")
                Update_Record_TDA("WHTSHPC2")
                Update_Record_TDA("WHTSHPC3")
                Update_Record_TDA("WHTSHPC4")
                Update_Record_TDA("WHTSHPC5")
                Update_Record_TDA("WHTSHPCG")
                Update_Record_TDA("WHTSHPCA")
                Update_Record_TDA("WHTSHPCS")
                Update_Record_TDA("WHTSHPCP")
                Update_Record_TDA("WHTSHPCC")
                CommitTrans()
            Catch ex As Exception
                ErrorMessage &= " " & ex.Message
                Rollback()
            End Try

            ' Notifications
            Dim CUST_EMAIL As String = drSOTORDR5.Item("CUST_EMAIL") & String.Empty
            CUST_EMAIL = CUST_EMAIL.Trim

            clsShip.ShipmentNotifications.Clear()
            If CUST_EMAIL.Length > 0 AndAlso drSOTCARR1.Item("CARRIER_SEND_NOTIFY") & String.Empty = "1" Then

                Dim notify As New WHCSHIP1.Notifications
                With notify
                    .email = CUST_EMAIL
                    .NotificationFlags = WHCSHIP1.NotifictaionTypes.On_Shipment
                    .Message = "Your Shipment from " & ROWs("ASTPARM1").Item("AS_PARM_INST_NAME") & " was picked up for shipment."
                End With
                clsShip.ShipmentNotifications.Add(notify)

                notify = New WHCSHIP1.Notifications
                With notify
                    .email = CUST_EMAIL
                    .NotificationFlags = WHCSHIP1.NotifictaionTypes.On_Deleivery
                    .Message = "Your Shipment from " & ROWs("ASTPARM1").Item("AS_PARM_INST_NAME") & " was delivered."
                End With
                clsShip.ShipmentNotifications.Add(notify)

                notify = New WHCSHIP1.Notifications
                With notify
                    .email = CUST_EMAIL
                    .NotificationFlags = WHCSHIP1.NotifictaionTypes.On_Exception
                    .Message = "Your Shipment from " & ROWs("ASTPARM1").Item("AS_PARM_INST_NAME") & " has a delivery problem."
                End With
                clsShip.ShipmentNotifications.Add(notify)
            End If

            If Not isInternationalShipment Then
                clsShip.CommodityDetailList.Clear()
            End If

            Select Case ASCMAIN1.CLIENT
                Case "RGI"
                    clsShip.ShipmentDescription = "Artificial Flowers / Home Decorations"
                Case "VAN"
                    clsShip.ShipmentDescription = "Undergarments"
                Case Else
                    clsShip.ShipmentDescription = "Garments"
            End Select

            clsShip.RequestedUPSInternationalForms.ShippersExportDeclarationInfo = New WHCSHIP1.ShippersExportDeclaration
            clsShip.RequestedUPSInternationalForms.ShippersExportDeclaration = False
            clsShip.RequestedUPSInternationalForms.CommercialInvoice = False

            If isInternationalShipment Then
                clsShip.RequestedUPSInternationalForms.ShippersExportDeclaration = True
                With clsShip.RequestedUPSInternationalForms.ShippersExportDeclarationInfo
                    .ImportEntryNumber = String.Empty
                    .InBond = TInBondCodes.ibcNotInBond
                    .LicenseDate = String.Empty
                    .LicenseExceptionCode = TExceptionCodes.ecNLR
                    .LicenseNumber = String.Empty
                    .PointOfOrigin = "US"
                    .ShippersTaxID = String.Empty
                    .TransPortType = String.Empty
                    .ExportingCarrier = CARRIER_CODE
                    .ExportingDate = System.DateTime.Now.ToString("yyyyMMdd")
                End With

                clsShip.RequestedUPSInternationalForms.CommercialInvoice = True
                With clsShip.RequestedUPSInternationalForms.CommercialInvoiceInfo
                    .Comments = String.Empty
                    .CustomersInvoiceNumber = CUST_CODE
                    .FreightCharge = 0
                    .InvoiceDate = System.DateTime.Now
                    .Purpose = CommercialInvoicePurposes.cipSold
                    .ShipperInsurance = 0
                    .Terms = CommercialInvoiceTerms.citCpt
                End With
            End If

            If clsShip.RequestLabel() Then

                drWHTSHPC1.Item("ERROR_MSG") = clsShip.LastError & String.Empty
                drWHTSHPC1.Item("STATUS") = "P"
                If drWHTSHPC1 IsNot Nothing AndAlso (drWHTSHPC1.Item("ERROR_MSG") & String.Empty).ToString.Length > 200 Then
                    drWHTSHPC1.Item("ERROR_MSG") = drWHTSHPC1("ERROR_MSG").ToString.Substring(0, 200).Trim
                End If

                If isPitneyBowes Then
                    drWHTSHPC1.Item("MASTER_TRACKING_NO") = clsShip.MasterTrackingNumber & String.Empty
                Else
                    drWHTSHPC1.Item("MASTER_TRACKING_NO") = clsShip.MasterTrackingNumber & String.Empty
                End If

                ' Update Pro Number if it is blank
                For Each dr As DataRow In dst.Tables("SOTSHIP1").Select("")
                    If dr.Item("SHIP_REF") & String.Empty = String.Empty Then
                        dr.Item("SHIP_REF") = clsShip.MasterTrackingNumber & String.Empty
                    End If
                Next

                For Each shipPackageDetail As PackageDetail In clsShip.PackageDetailList
                    SHIP_PACKAGE_NO = Val(shipPackageDetail.Id)
                    If dst.Tables("WHTSHPC2").Select("SHIP_PACKAGE_NO = " & SHIP_PACKAGE_NO, "").Length > 0 Then
                        drWHTSHPC2 = dst.Tables("WHTSHPC2").Select("SHIP_PACKAGE_NO = " & SHIP_PACKAGE_NO)(0)

                        Dim pitneyBowesshipdata As New TAC.WHCSHIP1.PitneyBowesPackageInformation

                        If isPitneyBowes Then
                            pitneyBowesshipdata = JsonConvert.DeserializeObject(Of TAC.WHCSHIP1.PitneyBowesPackageInformation)(shipPackageDetail.Reference)
                            drWHTSHPC2.Item("TRACKING_NO") = pitneyBowesshipdata.TrackingNumber & String.Empty
                            drWHTSHPC2.Item("TRACKING_NUMBER") = pitneyBowesshipdata.ShipmentID & String.Empty
                        Else
                            drWHTSHPC2.Item("TRACKING_NO") = shipPackageDetail.TrackingNumber & String.Empty
                        End If

                        drWHTSHPC2.Item("BASE_CHARGE") = Val(clsShip.ShipmentBaseCharge(SHIP_PACKAGE_NO) & String.Empty)
                        drWHTSHPC2.Item("NET_CHARGE") = Val(clsShip.ShipmentNetCharge(SHIP_PACKAGE_NO) & String.Empty)
                        drWHTSHPC2.Item("TOTAL_DISCOUNT") = Val(clsShip.ShipmentDiscountCharge(SHIP_PACKAGE_NO) & String.Empty)
                        drWHTSHPC2.Item("TOTAL_SURCHARGES") = Val(clsShip.ShipmentSurCharge(SHIP_PACKAGE_NO) & String.Empty)

                        drWHTSHPC2.Item("LENGTH") = Val(shipPackageDetail.Length & String.Empty)
                        drWHTSHPC2.Item("WIDTH") = Val(shipPackageDetail.Width & String.Empty)
                        drWHTSHPC2.Item("HEIGHT") = Val(shipPackageDetail.Height & String.Empty)

                        If clsShip.ShipmentListCharge.ContainsKey(SHIP_PACKAGE_NO) Then
                            drWHTSHPC2.Item("LIST_PRICE") = Val(clsShip.ShipmentListCharge(SHIP_PACKAGE_NO) & String.Empty)
                        Else
                            drWHTSHPC2.Item("LIST_PRICE") = drWHTSHPC2.Item("NET_CHARGE")
                        End If

                        OUR_FREIGHT = Val(drWHTSHPC2.Item("NET_CHARGE") & String.Empty)

                        ' Logic added 3/18/2017 for Regency
                        Select Case CARRIER_PPA_TYPE
                            Case "F" ' None
                                PPA_FREIGHT = 0

                            Case "L" ' List Rates
                                PPA_FREIGHT = Val(drWHTSHPC2.Item("LIST_PRICE") & String.Empty)

                            Case "N" ' Negioated Rates
                                PPA_FREIGHT = Val(drWHTSHPC2.Item("NET_CHARGE") & String.Empty)

                            Case Else
                                ' If not set then use List Price
                                PPA_FREIGHT = Val(drWHTSHPC2.Item("LIST_PRICE") & String.Empty)
                        End Select

                        If CARRIER_SURCHARGE_PERC > 0 Then
                            Select Case CARRIER_SURCHARGE_BASE
                                Case "N" ' Negotiated
                                    PPA_FREIGHT += Val(drWHTSHPC2.Item("NET_CHARGE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                                Case "L" ' List
                                    PPA_FREIGHT += Val(drWHTSHPC2.Item("LIST_PRICE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                                Case Else
                                    ' If not set then use List
                                    PPA_FREIGHT += Val(drWHTSHPC2.Item("LIST_PRICE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                            End Select
                        End If

                        PICK_NO = String.Empty
                        drSOTPICK1 = Nothing

                        ' We may have multi pick tickets in a single carton. This stamps them with the same tracking number
                        ' Spread the Customer Freight Cost and Our freight cost across the Pick Tickets
                        Dim numPickTickets As Int16 = dst.Tables("SOTCART1").Select("CART_SEQ = " & SHIP_PACKAGE_NO).Length
                        For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("CART_SEQ = " & SHIP_PACKAGE_NO)

                            If isPitneyBowes Then
                                drSOTCART1.Item("CART_TRACKING_NO") = pitneyBowesshipdata.TrackingNumber & String.Empty
                            Else
                                drSOTCART1.Item("CART_TRACKING_NO") = shipPackageDetail.TrackingNumber & String.Empty
                            End If

                            PICK_NO = drSOTCART1.Item("PICK_NO") & String.Empty
                            drSOTPICK1 = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)

                            'If Absx1.txtFor("FRT_TERMS").Text = "PPA" Then
                            '    ' RGI charges freight for all Orders.
                            '    If ASCMAIN1.CLIENT = "RGI" Then
                            '        drSOTPICK1.Item("PICK_FREIGHT") = Val(drSOTPICK1.Item("PICK_FREIGHT") & String.Empty) + Math.Round(PPA_FREIGHT / numPickTickets, 2)
                            '    ElseIf drSOTPICK1("ORDR_SOURCE") & String.Empty <> "W" Then
                            '        drSOTPICK1.Item("PICK_FREIGHT") = Val(drSOTPICK1.Item("PICK_FREIGHT") & String.Empty) + Math.Round(PPA_FREIGHT / numPickTickets, 2)
                            '    End If
                            'End If
                            drSOTPICK1.Item("PICK_FREIGHT") = Math.Round(OUR_FREIGHT / numPickTickets, 2)
                        Next
                        pitneyBowesshipdata = Nothing
                    End If

                    If isPitneyBowes Then
                        Dim file As String = shipPackageDetail.ShippingLabelFile
                        Using sr As New StreamReader(file)
                            ShippingLabels.Add(sr.ReadToEnd)
                            sr.Close()
                            sr.Dispose()
                        End Using
                    Else
                        ShippingLabels.Add(shipPackageDetail.ShippingLabel)
                        ShippingLabels.Add(shipPackageDetail.CODLabel)
                        ShippingLabels.Add(shipPackageDetail.ReturnReceipt)
                    End If
                Next

                'Dim totalLabelCharge As Decimal = Math.Round(Val(dst.Tables("SOTPICK1").Compute("SUM(PICK_FREIGHT)", $"PICK_NO = '{PICK_NO}'") & String.Empty), 2)
                'Dim rateCharge As Decimal = Math.Round(Val(dst.Tables("WHTSHPC4").Select($"CARRIER_CODE = '{CARRIER_CODE}' AND SERVICE_TYPE = '{CARRIER_PROD_CODE}'", "")(0).Item("TOTAL_CHARGE") & String.Empty))

                'If CInt(totalLabelCharge) > CInt(rateCharge) Then
                ' Dim diff As Decimal = Math.Round(totalLabelCharge - rateCharge, 2)
                'Dim userMessage As String = $"The Customer Freight Rate is {rateCharge.ToString("#,##0.00")} and the Label Charge is {totalLabelCharge.ToString("#,##0.00")}. This is a difference of {diff.ToString("#,##0.00")}."
                'userMessage &= Environment.NewLine & Environment.NewLine & "Do you want to continue?"
                'If MessageBox.Show(userMessage, "Freight Difference", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                '    ErrorMessage = "Discrepency in Customer Freight Rate and Label Charge. User cancelled Finalization."
                '    For Each drSOTPICK1x As DataRow In dst.Tables("SOTPICK1").Select("")
                '        drSOTPICK1x.Item("PICK_FREIGHT") = drSOTPICK1x.Item("PICK_FREIGHT_ORIG")
                '        drSOTPICK1x.Item("OUR_FREIGHT") = 0
                '    Next
                '    Return False
                'Else
                'For Each drSOTPICK1x As DataRow In dst.Tables("SOTPICK1").Select("")
                '        Dim drTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
                '        drTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
                '        drTATEVNT1.Item("TABLE_KEY") = drSOTPICK1x.Item("ORDR_NO")
                '        drTATEVNT1.Item("INIT_DATE") = DateTime.Now
                '        drTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                '        drTATEVNT1.Item("EVENT_TYPE") = "LBFRT"
                '        drTATEVNT1.Item("EVENT_DESC") = $"User {ASCMAIN1.USER_ID} choose to ship when the Customer Freight Rate was {rateCharge.ToString("#,##0.00")} and the Label Charge was {totalLabelCharge.ToString("#,##0.00")}"
                '        drTATEVNT1.Item("EVENT_KEY") = ""
                '        drTATEVNT1.Item("FORM_NAME") = "SOFSHIPB"
                '        dst.Tables("TATEVNT1").Rows.Add(drTATEVNT1)
                '    Next
                'End If
                'End If

                Try
                    BeginTrans()
                    Update_Record_TDA("WHTSHPC1")
                    Update_Record_TDA("WHTSHPC2")
                    Update_Record_TDA("SOTPICK1")
                    Update_Record_TDA("SOTSHIP1")
                    Update_Record_TDA("SOTCART1")
                    CommitTrans()
                Catch ex As Exception
                    ErrorMessage &= " " & ex.Message
                    Rollback()
                End Try

                If clsShip.InternationalFormsFile & String.Empty <> String.Empty Then
                    If My.Computer.FileSystem.FileExists(clsShip.InternationalFormsFile) Then
                        ShippingLabels.Add(clsShip.InternationalFormsFile)
                    End If
                End If

            Else
                ErrorMessage &= " " & clsShip.LastError
                RequestShippingLabel = False
            End If

        Catch ex As Exception
            ErrorMessage &= " " & ex.Message
            RequestShippingLabel = False
        End Try

        For Each shippingLabel As String In ShippingLabels
            If shippingLabel.Trim.Length > 0 Then PrintShippingLabels(shippingLabel)
        Next

        ErrorMessage = ErrorMessage.Trim

    End Function

    Sub Refresh_Refs(ByVal CarrierCode As String, Optional CART_NO As String = "")

        Dim REFERENCE1 As String = String.Empty
        Dim REF_CODE_1 As String = String.Empty
        Dim REF1_PREF As String = String.Empty
        Dim REF1_SUFF As String = String.Empty

        Dim REFERENCE2 As String = String.Empty
        Dim REF_CODE_2 As String = String.Empty
        Dim REF2_PREF As String = String.Empty
        Dim REF2_SUFF As String = String.Empty

        Dim REFERENCE3 As String = String.Empty
        Dim REF_CODE_3 As String = String.Empty
        Dim REF3_PREF As String = String.Empty
        Dim REF3_SUFF As String = String.Empty

        Dim drSOTCARRR As DataRow = Nothing

        If Not (CarrierCode = "FEDEX" OrElse CarrierCode = "UPS") Then
            Exit Sub
        End If

        If dst.Tables("ARTCUSTS").Rows.Count = 0 Then
            Exit Sub
        End If

        Dim drARTCUSTS As DataRow = dst.Tables("ARTCUSTS").Rows(0)

        Select Case CarrierCode

            Case "UPS"
                If drARTCUSTS.Item("UPS_REF1") & String.Empty <> String.Empty Then
                    REF_CODE_1 = drARTCUSTS.Item("UPS_REF1") & String.Empty
                    REF1_PREF = drARTCUSTS.Item("UPS_REF1_PREF") & String.Empty
                    REF1_SUFF = drARTCUSTS.Item("UPS_REF1_SUFF") & String.Empty

                    drSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {CarrierCode, REF_CODE_1})
                    If drSOTCARRR IsNot Nothing Then
                        REFERENCE1 = REF_CODE_1.Substring(0, 2) & ":" & drSOTCARRR.Item("TABLE_NAME") & "." & drSOTCARRR.Item("COLUMN_NAME")
                    End If
                End If

                If drARTCUSTS.Item("UPS_REF2") & String.Empty <> String.Empty Then
                    REF_CODE_2 = drARTCUSTS.Item("UPS_REF2") & String.Empty
                    REF2_PREF = drARTCUSTS.Item("UPS_REF2_PREF") & String.Empty
                    REF2_SUFF = drARTCUSTS.Item("UPS_REF2_SUFF") & String.Empty

                    drSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {CarrierCode, REF_CODE_2})
                    If drSOTCARRR IsNot Nothing Then
                        REFERENCE2 = REF_CODE_2.Substring(0, 2) & ":" & drSOTCARRR.Item("TABLE_NAME") & "." & drSOTCARRR.Item("COLUMN_NAME")
                    End If
                End If

                If drARTCUSTS.Item("UPS_REF3") & String.Empty <> String.Empty Then
                    REF_CODE_3 = drARTCUSTS.Item("UPS_REF3") & String.Empty
                    REF3_PREF = drARTCUSTS.Item("UPS_REF3_PREF") & String.Empty
                    REF3_SUFF = drARTCUSTS.Item("UPS_REF3_SUFF") & String.Empty

                    drSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {CarrierCode, REF_CODE_3})
                    If drSOTCARRR IsNot Nothing Then
                        REFERENCE3 = REF_CODE_3.Substring(0, 2) & ":" & drSOTCARRR.Item("TABLE_NAME") & "." & drSOTCARRR.Item("COLUMN_NAME")
                    End If
                End If

            Case "FEDEX"

                If drARTCUSTS.Item("FDX_REF1") & String.Empty <> String.Empty Then
                    REF_CODE_1 = drARTCUSTS.Item("FDX_REF1") & String.Empty
                    REF1_PREF = drARTCUSTS.Item("FDX_REF1_PREF") & String.Empty
                    REF1_SUFF = drARTCUSTS.Item("FDX_REF1_SUFF") & String.Empty

                    drSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {CarrierCode, REF_CODE_1})
                    If drSOTCARRR IsNot Nothing Then
                        REFERENCE1 = REF_CODE_1.Substring(0, 2) & ":" & drSOTCARRR.Item("TABLE_NAME") & "." & drSOTCARRR.Item("COLUMN_NAME")
                    End If
                End If

                If drARTCUSTS.Item("FDX_REF2") & String.Empty <> String.Empty Then
                    REF_CODE_2 = drARTCUSTS.Item("FDX_REF2") & String.Empty
                    REF2_PREF = drARTCUSTS.Item("FDX_REF2_PREF") & String.Empty
                    REF2_SUFF = drARTCUSTS.Item("FDX_REF2_SUFF") & String.Empty

                    drSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {CarrierCode, REF_CODE_2})
                    If drSOTCARRR IsNot Nothing Then
                        REFERENCE2 = REF_CODE_2.Substring(0, 2) & ":" & drSOTCARRR.Item("TABLE_NAME") & "." & drSOTCARRR.Item("COLUMN_NAME")
                    End If
                End If

                If drARTCUSTS.Item("FDX_REF3") & String.Empty <> String.Empty Then
                    REF_CODE_3 = drARTCUSTS.Item("FDX_REF3") & String.Empty
                    REF3_PREF = drARTCUSTS.Item("FDX_REF3_PREF") & String.Empty
                    REF3_SUFF = drARTCUSTS.Item("FDX_REF3_SUFF") & String.Empty

                    drSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {CarrierCode, REF_CODE_3})
                    If drSOTCARRR IsNot Nothing Then
                        REFERENCE3 = REF_CODE_3.Substring(0, 2) & ":" & drSOTCARRR.Item("TABLE_NAME") & "." & drSOTCARRR.Item("COLUMN_NAME")
                    End If
                End If

        End Select

        Dim temp1 As String = String.Empty
        Dim temp2 As String = String.Empty
        Dim temp3 As String = String.Empty

        If CART_NO.Length > 0 Then
            CART_NO = "CART_NO = '" & CART_NO & "'"
        End If

        If Not dst.Tables("SOTCART1").Columns.Contains("SHIP_BOL_NO") Then
            dst.Tables("SOTCART1").Columns.Add("SHIP_BOL_NO", GetType(String))
        End If

        For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select(CART_NO, "SHIP_BOL_NO,CART_NO")

            temp1 = String.Empty
            temp2 = String.Empty
            temp3 = String.Empty

            If REFERENCE1.Length > 0 Then
                temp1 = GetReferenceValue(REF1_PREF, REF1_SUFF, REFERENCE1.Split(":")(1), drSOTCART1.Item("CART_NO"))
                If temp1.Length > 0 Then
                    temp1 = REFERENCE1.Split(":")(0).Substring(0, 2) & ":" & temp1
                End If
            End If

            If REFERENCE2.Length > 0 Then
                temp2 = GetReferenceValue(REF2_PREF, REF2_SUFF, REFERENCE2.Split(":")(1), drSOTCART1.Item("CART_NO"))
                If temp2.Length > 0 Then
                    temp2 = REFERENCE2.Split(":")(0).Substring(0, 2) & ":" & temp2
                End If
            End If

            If REFERENCE3.Length > 0 Then
                temp3 = GetReferenceValue(REF3_PREF, REF3_SUFF, REFERENCE3.Split(":")(1), drSOTCART1.Item("CART_NO"))
                If temp3.Length > 0 Then
                    temp3 = REFERENCE3.Split(":")(0).Substring(0, 2) & ":" & temp3
                End If
            End If

            If CarrierCode = "FEDEX" AndAlso drSOTCARRR IsNot Nothing Then
                If temp1 <> "" Then
                    If REF_CODE_1.Length = 0 Then
                        REF_CODE_1 = temp1.Substring(0, 2)
                    End If
                    temp1 = Replace(temp1, temp1.Substring(0, 3), Mid(REF_CODE_1, 1, 2) & ":")
                End If

                If temp2 <> "" Then
                    If REF_CODE_2.Length = 0 Then
                        REF_CODE_2 = temp2.Substring(0, 2)
                    End If
                    temp2 = Replace(temp2, temp2.Substring(0, 3), Mid(REF_CODE_2, 1, 2) & ":")
                End If

                If temp3 <> "" Then
                    If REF_CODE_3.Length = 0 Then
                        REF_CODE_3 = temp3.Substring(0, 2)
                    End If
                    temp3 = Replace(temp3, temp3.Substring(0, 3), Mid(REF_CODE_3, 1, 2) & ":")
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

            If temp1.Length > 0 Then drSOTCART1.Item("REFERENCE1") = temp1
            If temp2.Length > 0 Then drSOTCART1.Item("REFERENCE2") = temp2
            If temp3.Length > 0 Then drSOTCART1.Item("REFERENCE3") = temp3
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

                Case "EDT850T1"
                    dataRow = dst.Tables("SOTCART1").Rows.Find(CART_NO)
                    If dataRow Is Nothing Then Exit Select

                    dataRow = dst.Tables("SOTPICK1").Rows.Find(dataRow.Item("PICK_NO") & String.Empty)
                    If dataRow Is Nothing Then Exit Select

                    dataRow = dst.Tables("SOTORDR1").Rows.Find(dataRow.Item("ORDR_NO") & String.Empty)
                    If dataRow Is Nothing Then Exit Select

                    dataRow = dst.Tables("EDT850T1").Rows.Find(dataRow.Item("EDI_DOC_SEQ_NO") & String.Empty)

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

    Public Function PrintShippingLabels(ByVal LabelData As String) As Boolean

        Try
            If LabelData.ToUpper.EndsWith(".PDF") Then
                If My.Computer.FileSystem.FileExists(LabelData) Then
                    Dim waitAmount As Int16 = 0
                    Dim requestedIpAddress As String = txtLaserPrinter.Text
                    If LabelData.ToUpper.EndsWith(".ZPL") Then
                        requestedIpAddress = txtLabelPrinter.Text
                    End If
                    Dim requestedStreamPort As String = "9100"

                    If requestedIpAddress.Contains(":") Then
                        requestedStreamPort = requestedIpAddress.Split(":")(1)
                        requestedIpAddress = requestedIpAddress.Split(":")(0)
                    End If

                    Using ipp As New nsoftware.IPWorks.Ipport
                        ipp.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareipportkey")
                        'ipp.Config("SSLEnabledProtocols=" & TAC.TACMAIN1.SSLEnabledProtocols)
                        ipp.Connect(requestedIpAddress, Val(requestedStreamPort))
                        Using binaryReader As New System.IO.BinaryReader(System.IO.File.Open(LabelData, System.IO.FileMode.Open))
                            waitAmount = 0
                            Do
                                'ErrorMessage = "If Not ipp.Connected Then"
                                If Not ipp.Connected Then
                                    System.Threading.Thread.Sleep(1000)
                                    waitAmount += 1
                                End If
                            Loop While Not ipp.Connected And waitAmount < 15

                            'ErrorMessage = "ipp.Send(binaryReader.ReadBytes(binaryReader.BaseStream.Length))"
                            ipp.Send(binaryReader.ReadBytes(binaryReader.BaseStream.Length))
                            System.Threading.Thread.Sleep(1000)
                            binaryReader.Close()
                            binaryReader.Dispose()
                        End Using

                        'ErrorMessage = "ipp.Disconnect()"
                        ipp.Disconnect()
                        ipp.Dispose()
                    End Using
                End If
                Return True
            End If

            If IsIPAddress(txtLabelPrinter.Text) Then
                clsTACZPLT1.SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, LabelData)
            Else
                ASCMAIN1.LabelPrinterSerialPort.WriteLine(LabelData)
            End If


        Catch ex As Exception
            MessageBox.Show("Print Shipping Label Error: " & ex.Message)
        End Try

    End Function

    Function IsIPAddress(value As String) As Boolean
        Dim ip As IPAddress = Nothing
        Return IPAddress.TryParse(value, ip)
    End Function

    Function IsPOBox(address As String) As Boolean
        If String.IsNullOrWhiteSpace(address) Then
            Return False
        End If

        Dim pattern As String = "\b(P\.?\s*O\.?\s*Box|Post\s+Office\s+Box)\b"
        Dim regex As New Regex(pattern, RegexOptions.IgnoreCase)

        Return regex.IsMatch(address)
    End Function

#End Region

#Region "grdWHTSHPC4"

    Private Sub grdWHTSHPC4_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTSHPC4.DoubleClickRow
        Dim CARRIER_CODE As String = e.Row.Cells("CARRIER_CODE").Value
        Dim CARRIER_PROD_CODE As String = e.Row.Cells("SERVICE_TYPE").Value & String.Empty
        Dim SHIP_VIA_CODE As String = e.Row.Cells("SHIP_VIA_CODE").Value & String.Empty

        ASCMAIN1.sql = $"CARRIER_CODE = '{CARRIER_CODE}' AND CARRIER_PROD_CODE = '{CARRIER_PROD_CODE}' AND SHIP_VIA_CODE = '{SHIP_VIA_CODE}' AND SHIP_VIA_STATUS = 'A'"
        If dst.Tables("SOTSVIA1").Select(ASCMAIN1.sql).Length > 0 Then
            txtSHIP_VIA_CODE.Text = SHIP_VIA_CODE
        End If

    End Sub

    Private Sub grdWHTSHPC4_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdWHTSHPC4.InitializeRow

        If e.Row.Cells("DISCLAIMER").Value & String.Empty <> String.Empty Then
            e.Row.Cells("CARRIER_CODE").Appearance.BackColor = Drawing.Color.DarkMagenta
            e.Row.Cells("CARRIER_CODE").Appearance.ForeColor = Drawing.Color.White
        End If

        If e.Row.Cells("TOTAL_CHARGE").Value > e.Row.Cells("LIST_NET_CHARGE").Value + e.Row.Cells("SURCHARGE").Value Then
            e.Row.Cells("TOTAL_CHARGE").Appearance.FontData.Bold = DefaultableBoolean.True
            e.Row.Cells("TOTAL_CHARGE").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("TOTAL_CHARGE").Appearance.FontData.Bold = DefaultableBoolean.False
            e.Row.Cells("TOTAL_CHARGE").Appearance.ForeColor = Drawing.Color.Black
        End If

        Dim CARRIER_CODE As String = e.Row.Cells("CARRIER_CODE").Value & String.Empty
        Dim CARRIER_PROD_CODE As String = e.Row.Cells("SERVICE_TYPE").Value & String.Empty
        Dim SHIP_VIA_CODE As String = e.Row.Cells("SHIP_VIA_CODE").Value & String.Empty

        If CARRIER_CODE.Length > 0 AndAlso CARRIER_PROD_CODE.Length > 0 Then
            ASCMAIN1.sql = "SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "' and CARRIER_CODE = '" & CARRIER_CODE & "' AND CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "' and SHIP_VIA_STATUS = 'A'"
            If dst.Tables("SOTSVIA1").Select(ASCMAIN1.sql).Length = 0 Then
                If dst.Tables("SOTSVIA1").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "' and SHIP_VIA_STATUS = 'A'")(0).Item("SHIP_VIA_CODE").length > 0 Then
                    e.Row.Cells("SHIP_VIA_CODE").Value = dst.Tables("SOTSVIA1").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "' and SHIP_VIA_STATUS = 'A'")(0).Item("SHIP_VIA_CODE")
                    grdWHTSHPC4.UpdateData()
                End If
            End If
        End If
    End Sub

    Private Sub grdWHTSHPC4_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdWHTSHPC4.BeforeRowUpdate

        If e.Row.Band.Key <> grdWHTSHPC4.DisplayLayout.Bands(0).Key Then
            Exit Sub
        End If

        If e.Row.Cells("SELECTED").Value & String.Empty = "1" Then
            For Each dRow As DataRow In dst.Tables("WHTSHPC4").Select("SELECTED = '1'", "", DataViewRowState.CurrentRows)
                dRow.Item("SELECTED") = "0"
            Next
        End If

    End Sub

#End Region

#Region "Overrides"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

            Case "TRUCK_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    txtTRUCK_NO.Text = txtTRUCK_NO.Text.Trim.ToUpper
                    If txtTRUCK_NO.TextLength > 0 Then
                        If Not ValidateTruckTote(ValidateTruckToteTypes.Truck, txtTRUCK_NO.Text) Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "TOTE_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    txtTOTE_NO.Text = txtTOTE_NO.Text.ToUpper.Trim
                    If txtTOTE_NO.TextLength > 0 Then
                        ValidateTruckTote(ValidateTruckToteTypes.Tote, txtTOTE_NO.Text)
                    End If

                    If dst.Tables("SOTPICK1X").Select("ISNULL(SELECTED, '0') = '0'").Length = 0 Then
                        AutoCancel = True
                        ASCMAIN1.Progress("Resetting screen for next Truck", "")
                        Click_Command("Cancel")
                    End If
                End If

        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "SHIP_VIA_CODE"
                Dim SHIP_VIA_CODE As String = txtSHIP_VIA_CODE.Text.Trim

                ASCMAIN1.Add_Value_List(grdSOTCART1, "PACKAGING_TYPE", $"SELECT SOTCARR4.PACKAGE_CODE, SOTCARR4.PACKAGE_DESC
                                FROM SOTSVIA1, SOTCARR4
                                WHERE SOTCARR4.CARRIER_CODE = SOTSVIA1.CARRIER_CODE
                                AND SOTSVIA1.SHIP_VIA_CODE = '{SHIP_VIA_CODE}'
                                ORDER BY PACKAGE_CODE DESC")

            Case Else

        End Select

    End Sub

#End Region

#Region "Devices"

    ''' <summary>
    ''' Sets up and Initializes the Scanner Control
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetupScanner()

        scannedDelegate = AddressOf ProcessScannedData
        Try
            txtComPort.Appearance.BackColor = Drawing.Color.Red
            txtComPort.Clear()

            'If ASCMAIN1.ScannerSerialPort IsNot Nothing Then
            '    txtComPort.Appearance.BackColor = Drawing.Color.Green
            '    txtComPort.Text = ASCMAIN1.ScannerSerialPort.PortName
            'End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub SetUpPortsAndPrinters()

        Dim prtdoc As New System.Drawing.Printing.PrintDocument
        txtLaserPrinter.Text = prtdoc.PrinterSettings.PrinterName
        'ASCMAIN1.InvoicePrinterIpAddress = txtInvoicePrinter.Text

        txtLabelPrinter.Appearance.BackColor = Drawing.Color.LightGreen
        If ASCMAIN1.LabelPrinterIPAddress.Length > 0 Then
            txtLabelPrinter.Text = ASCMAIN1.LabelPrinterIPAddress

            Dim pingSender As New Ping()
            Dim status As Boolean = False
            Try
                Dim reply As PingReply = pingSender.SendPingAsync(ASCMAIN1.LabelPrinterIPAddress, 1000).Result
                status = (reply.Status = IPStatus.Success)
            Catch
                status = False
            End Try

            If Not status Then
                txtLabelPrinter.Appearance.BackColor = Drawing.Color.Red
                txtLabelPrinter.Appearance.ForeColor = Drawing.Color.White
            End If

        ElseIf ASCMAIN1.LabelPrinterSerialPort IsNot Nothing Then
            txtLabelPrinter.Text = "Serial " & ASCMAIN1.LabelPrinterSerialPort.PortName
            Try
                If Not ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                    ASCMAIN1.LabelPrinterSerialPort.Open()
                End If
            Catch ex As Exception
                txtLabelPrinter.Appearance.BackColor = Drawing.Color.Red
                txtLabelPrinter.Appearance.ForeColor = Drawing.Color.White
            End Try
        ElseIf ASCMAIN1.LabelPrinterIPAddress.Length > 0 Then
            txtLabelPrinter.Text = ASCMAIN1.LabelPrinterIPAddress
        ElseIf ASCMAIN1.LabelPrinterName.Length Then
            txtLabelPrinter.Text = ASCMAIN1.LabelPrinterName
        Else
            txtLabelPrinter.Text = "No Port"
            txtLabelPrinter.Appearance.BackColor = Drawing.Color.Red
            txtLabelPrinter.Appearance.ForeColor = Drawing.Color.White
        End If

    End Sub

    Private Sub CreateAppearances()
        dictAppearances.Add("AllItemsOnBackOrder", New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.DarkSlateGray, .ForeColor = Drawing.Color.White})
        dictAppearances.Add("Appearance_Incomplete", New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.Red, .ForeColor = Drawing.Color.White})
        dictAppearances.Add("Selected_Tote", New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.LightGreen, .ForeColor = Drawing.Color.Black})
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdSOTTRCK1X_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTTRCK1X.DoubleClickRow

        If grdSOTTRCK1X.ActiveRow Is Nothing Then
            Exit Sub
        End If

        If grdSOTTRCK1X.ActiveRow.IsFilterRow OrElse grdSOTTRCK1X.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        txtTRUCK_NO.Text = grdSOTTRCK1X.ActiveRow.Cells("TRUCK_NO").Value & String.Empty
        ValidateTruckTote(ValidateTruckToteTypes.Truck, txtTRUCK_NO.Text)
    End Sub

    Private Sub grdSOTPICK1X_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTPICK1X.InitializeRow
        If e.Row.Cells("INCOMPLETE").Value & String.Empty = "1" Then
            e.Row.Appearance = dictAppearances("Appearance_Incomplete")
        ElseIf e.Row.Cells("ALL_ITEMS_BACK").Value & String.Empty = "1" Then
            e.Row.Cells("TOTE_NO").Appearance = dictAppearances("AllItemsOnBackOrder")
        End If

        If e.Row.Cells("TOTE_NO").Value & String.Empty = txtTOTE_NO.Text Then
            e.Row.Cells("TOTE_NO").Appearance = dictAppearances("Selected_Tote")
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Stop()

        If ScreenMode Then
            Select Case screenProcessingMode
                Case ScreenProcessingModes.DisplayAvailableTrucks
                    txtTOTE_NO.Clear()
                Case ScreenProcessingModes.TruckSelected
                    txtTOTE_NO.Clear()
                    txtTOTE_NO.Focus()
                Case ScreenProcessingModes.ProcessingSelectedTruckTote
                    ' Nothing
            End Select
        Else
            Select Case screenProcessingMode
                Case ScreenProcessingModes.DisplayAvailableTrucks
                    txtTRUCK_NO.Clear()
                    txtTRUCK_NO.Focus()
                Case ScreenProcessingModes.TruckSelected
                    ' Nothing
                Case ScreenProcessingModes.ProcessingSelectedTruckTote
                    ' Nothing 
            End Select
        End If

    End Sub

    Private Sub btnLabelPrinter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLabelPrinter.Click
        Try
            If txtLabelPrinter.Text.Trim.Length = 0 Then
                MessageBox.Show("There is no assigned Label Printer.", "Test Print", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            clsTACZPLT1.PrintSampleShippingLabel()
            MessageBox.Show("Test Label sent to Printer " & txtLabelPrinter.Text, "Test Print", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Label Printer Test Print Error {ex.Message}", "Test Print", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnLaserPrinter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLaserPrinter.Click
        Try
            If txtLaserPrinter.Text.Trim.Length = 0 Then
                MessageBox.Show("There is no assigned Invoice Printer.", "Test Print", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            'PrintInvoice("", True)

        Catch ex As Exception
            MessageBox.Show($"Test Invoice Printer Error: {ex.Message}", "Test Print", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyDown(e)

        If e.KeyCode = System.Windows.Forms.Keys.F8 Then
            'Try
            '    With UltraExplorerBar1.Groups("Screen Control").Items("Update")
            '        If .Visible AndAlso .Settings.Enabled = DefaultableBoolean.True Then
            '            Me.Validate()
            '            UltraExplorerBar1.Focus()
            '            Click_Command("Update")
            '            e.Handled = True
            '            Exit Sub
            '        End If
            '    End With

            'Catch ex As Exception

            'End Try
        End If
    End Sub

    Private Sub grdSOTPICK1X_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTPICK1X.DoubleClickRow

        If ASCMAIN1.Running_in_VS Then
            If grdSOTPICK1X.ActiveRow Is Nothing Then
                Exit Sub
            End If

            If grdSOTPICK1X.ActiveRow.IsFilterRow OrElse grdSOTPICK1X.ActiveRow.IsAddRow Then
                Exit Sub
            End If

            txtTOTE_NO.Text = grdSOTPICK1X.ActiveRow.Cells("TOTE_NO").Value & String.Empty
            ValidateTruckTote(ValidateTruckToteTypes.Tote, txtTOTE_NO.Text)
        End If

    End Sub

#End Region

#Region "grdSOTCART1"

    Private Sub grdSOTCART1_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdSOTCART1.AfterRowsDeleted
        ' When a user changes the vales in the Cartons we must clear the Rates
        dst.Tables("WHTSHPC4").Rows.Clear()
    End Sub

    Private Sub grdSOTCART1_BeforeRowsDeleted(sender As Object, e As UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTCART1.BeforeRowsDeleted

        e.DisplayPromptMsg = False

        For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Rows
            If grdRow.Cells("CART_SEQ").Value = 1 Then
                MessageBox.Show("You cannot Delete Carton sequence 1.", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information)
                e.Cancel = True
                Exit Sub
            End If

            Dim CART_NO As String = grdRow.Cells("CART_NO").Value
            If dst.Tables("SOTCART2").Select($"CART_NO = '{CART_NO}'").Length > 0 Then
                MessageBox.Show("You cannot Delete a Carton that contains product.", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information)
                e.Cancel = True
                Exit Sub
            End If
        Next

        If MessageBox.Show("Do you want to Delete " & e.Rows.Count & " Cartons?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            e.Cancel = True
            Exit Sub
        End If
    End Sub

    Private Sub grdSOTCART1_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCART1.BeforeRowUpdate

        Dim PICK_NO As String = e.Row.Cells("PICK_NO").Value & String.Empty
        If PICK_NO.Length > 0 Then
            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
            If rowSOTPICK1 IsNot Nothing AndAlso rowSOTPICK1.Item("PICK_STATUS") & String.Empty = "C" Then
                MessageBox.Show("You are not permitted to modify a carton where the Pick Ticket has a status of Cancelled", "Update", MessageBoxButtons.OK)
                e.Cancel = True
                Exit Sub
            End If
        End If

        If ASCMAIN1.CLIENT = "RGI" Then
            If e.Row.Cells("REFERENCE1").Value & String.Empty = String.Empty Then
                e.Row.Cells("REFERENCE1").Value = Val(e.Row.Cells("PICK_NO").Value & String.Empty)
            End If
            e.Row.Cells("REFERENCE2").Value = e.Row.Cells("CART_NO").Value & String.Empty
        End If

        Dim rowWHTPKGM1 As DataRow = Nothing
        Dim PKG_CODE As String = e.Row.Cells("PKG_CODE").Value & String.Empty
        Dim PACKAGING_TYPE As String = e.Row.Cells("PACKAGING_TYPE").Value & String.Empty

        If PACKAGING_TYPE = defaultPACKAGING_TYPE Then
            If dst.Tables("WHTPKGM1").Rows.Find(PKG_CODE) Is Nothing Then
                MessageBox.Show("Invalid Package Code for Package Type ", "Update", MessageBoxButtons.OK)
                e.Cancel = True
                Exit Sub
            End If
        End If

        If dst.Tables("WHTPKGM1").Select("PKG_CODE = '" & PKG_CODE & "' AND PKG_CODE <> 'OTHER'").Length > 0 Then
            rowWHTPKGM1 = dst.Tables("WHTPKGM1").Select("PKG_CODE = '" & PKG_CODE & "'")(0)
            e.Row.Cells("PKG_L").Value = rowWHTPKGM1.Item("PKG_L")
            e.Row.Cells("PKG_W").Value = rowWHTPKGM1.Item("PKG_W")
            e.Row.Cells("PKG_H").Value = rowWHTPKGM1.Item("PKG_H")
            Exit Sub
        End If

        'If ASCMAIN1.CLIENT = "NYA" Or ASCMAIN1.CLIENT = "VAN" Then
        '    Exit Sub
        '    'do not add logic below unless that logic is related to sorting the dimensions
        'End If

        ' Sort the values by length, width, height
        Dim PKG_L As Decimal = Val(e.Row.Cells("PKG_L").Value & String.Empty)
        Dim PKG_W As Decimal = Val(e.Row.Cells("PKG_W").Value & String.Empty)
        Dim PKG_H As Decimal = Val(e.Row.Cells("PKG_H").Value & String.Empty)

        If PKG_L <= 0 OrElse PKG_W <= 0 OrElse PKG_H < 0 Then
            MessageBox.Show("All dimensions must be greater than 0", "Update", MessageBoxButtons.OK)
            e.Cancel = True
            Exit Sub
        End If

        Dim dimList As New List(Of Decimal)
        dimList.Add(PKG_L)
        dimList.Add(PKG_W)
        dimList.Add(PKG_H)
        dimList.Sort()
        PKG_L = dimList(2)
        PKG_W = dimList(1)
        PKG_H = dimList(0)

        e.Row.Cells("PKG_L").Value = PKG_L
        e.Row.Cells("PKG_W").Value = PKG_W
        e.Row.Cells("PKG_H").Value = PKG_H

    End Sub

    Private Sub grdSOTCART1_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCART1.AfterCellUpdate

        Dim displayBoxAttributes As Boolean = False
        Static UpdatingDimensions As Boolean = False

        If UpdatingDimensions = True Then Exit Sub

        For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTCART1.Rows
            If row.Cells("PKG_CODE").Text = defaultPKG_CODE Then
                displayBoxAttributes = True
                Exit For
            End If
        Next

        UpdatingDimensions = True

        If displayBoxAttributes OrElse ASCMAIN1.CLIENT = "RGI" OrElse ASCMAIN1.CLIENT = "VAN" Then
            grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_W").Hidden = False
            grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_L").Hidden = False
            grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_H").Hidden = False

            grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_W").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_L").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_H").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_W").Hidden = True
            grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_L").Hidden = True
            grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_H").Hidden = True

            grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_W").CellActivation = UltraWinGrid.Activation.NoEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_L").CellActivation = UltraWinGrid.Activation.NoEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_H").CellActivation = UltraWinGrid.Activation.NoEdit
        End If

        If e.Cell.Column.Key = "PKG_CODE" Then
            Try
                If dst.Tables("WHTPKGM1").Columns.Contains("PKG_CODE") Then
                    Dim PKG_CODE As String = e.Cell.Value & String.Empty
                    If PKG_CODE.Length > 0 Then

                        ' See if someone else create the package since opening the screen
                        If dst.Tables("WHTPKGM1").Select("PKG_CODE = '" & PKG_CODE.Replace("'", "''") & "'").Length = 0 Then
                            Fill_Records("WHTPKGM1", String.Empty, True, "SELECT * FROM WHTPKGM1")
                            ultraComboPackage.DataSource = ASCDATA1.GetDataTable("SELECT PKG_CODE, PKG_DESC, PKG_L || ' x ' ||  PKG_W || ' x ' || PKG_H PKG_D FROM WHTPKGM1 order by PKG_CODE")
                        End If

                        ' Now look for it, if it does not exist then let the user create it
                        If dst.Tables("WHTPKGM1").Select("PKG_CODE = '" & PKG_CODE.Replace("'", "''") & "'").Length > 0 Then
                            Dim rowWHTPKGM1 As DataRow = dst.Tables("WHTPKGM1").Select("PKG_CODE = '" & PKG_CODE.Replace("'", "''") & "'")(0)
                            e.Cell.Row.Cells("PACKAGING_TYPE").Value = defaultPACKAGING_TYPE
                            e.Cell.Row.Cells("PKG_CODE").Value = rowWHTPKGM1.Item("PKG_CODE")
                            e.Cell.Row.Cells("PKG_L").Value = rowWHTPKGM1.Item("PKG_L")
                            e.Cell.Row.Cells("PKG_W").Value = rowWHTPKGM1.Item("PKG_W")
                            e.Cell.Row.Cells("PKG_H").Value = rowWHTPKGM1.Item("PKG_H")
                        Else
                            'Using frmWHFPKGM1 As New TAC.WHFPKGM1
                            '    frmWHFPKGM1.PKG_BOX_UPC = PKG_CODE
                            '    frmWHFPKGM1.ShowDialog()
                            '    If frmWHFPKGM1.Updated Then
                            '        Fill_Records("WHTPKGM1", String.Empty, True, "SELECT * FROM WHTPKGM1")
                            '        ultraComboPackage.DataSource = ASCDATA1.GetDataTable("SELECT PKG_CODE, PKG_DESC, PKG_L || ' x ' ||  PKG_W || ' x ' || PKG_H PKG_D FROM WHTPKGM1 order by PKG_CODE")
                            '        If dst.Tables("WHTPKGM1").Select("PKG_CODE = '" & PKG_CODE.Replace("'", "''") & "'").Length > 0 Then
                            '            Dim rowWHTPKGM1 As DataRow = dst.Tables("WHTPKGM1").Select("PKG_BOX_UPC = '" & PKG_CODE.Replace("'", "''") & "'")(0)
                            '            e.Cell.Row.Cells("PACKAGING_TYPE").Value = defaultPACKAGING_TYPE
                            '            e.Cell.Row.Cells("PKG_CODE").Value = rowWHTPKGM1.Item("PKG_CODE")
                            '            e.Cell.Row.Cells("PKG_L").Value = rowWHTPKGM1.Item("PKG_L")
                            '            e.Cell.Row.Cells("PKG_W").Value = rowWHTPKGM1.Item("PKG_W")
                            '            e.Cell.Row.Cells("PKG_H").Value = rowWHTPKGM1.Item("PKG_H")
                            '        End If
                            '    End If
                            'End Using
                        End If
                    End If
                End If
            Catch ex As Exception
            End Try
        End If

        grdSOTCART1.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.VisibleRows, True)
        UpdatingDimensions = False
    End Sub

    Private Sub grdSOTCART1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCART1.AfterRowUpdate
        ' When a user changes the vales in the Cartons we must clear the Rates
        dst.Tables("WHTSHPC4").Rows.Clear()
    End Sub

    Private Sub grdSOTCART1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTCART1.AfterRowActivate
        Dim CART_NO As String = grdSOTCART1.ActiveRow.Cells("CART_NO").Value & String.Empty

        Dim dv As DataView = DirectCast(grdSOTCART2.DataSource, DataTable).DefaultView
        dv.RowFilter = $"CART_NO = '{CART_NO}'"
        dv.Sort = "CART_LNO"

        grdSOTCART2.Text = $"Details for Carton {CART_NO}"
    End Sub

    Private Sub grdSOTPICK1X_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdSOTPICK1X.InitializeLayout

    End Sub

    Private dragRow As UltraGridRow = Nothing

    Private Sub grdSOTCART2_MouseDown(sender As Object, e As MouseEventArgs) Handles grdSOTCART2.MouseDown
        Try
            If e.Button <> MouseButtons.Left Then
                dragRow = Nothing
                Exit Sub
            End If

            ' Get the element under the mouse
            Dim element As UIElement = grdSOTCART2.DisplayLayout.UIElement.ElementFromPoint(New Point(e.X, e.Y))
            If element Is Nothing Then
                dragRow = Nothing
                Return
            End If

            ' Ask the element for its UltraGridRow context
            Dim row As UltraGridRow = TryCast(element.GetContext(GetType(UltraGridRow)), UltraGridRow)

            If row IsNot Nothing AndAlso row.IsDataRow Then
                dragRow = row
            Else
                dragRow = Nothing
            End If

        Catch ex As Exception
            dragRow = Nothing
        End Try
    End Sub

    Private Sub grdSOTCART2_MouseMove(sender As Object, e As MouseEventArgs) Handles grdSOTCART2.MouseMove
        Try
            If e.Button = MouseButtons.Left AndAlso dragRow IsNot Nothing Then
                ' Start the drag operation with the row object
                grdSOTCART2.DoDragDrop(dragRow, DragDropEffects.Copy)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub grdSOTCART1_DragEnter(sender As Object, e As DragEventArgs) Handles grdSOTCART1.DragEnter
        Try
            If e.Data.GetDataPresent(GetType(UltraGridRow)) Then
                e.Effect = DragDropEffects.Copy
            Else
                e.Effect = DragDropEffects.None
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub grdSOTCART1_DragDrop(sender As Object, e As DragEventArgs) Handles grdSOTCART1.DragDrop
        Try
            If e.Data.GetDataPresent(GetType(UltraGridRow)) Then
                Dim sourceRow As UltraGridRow = CType(e.Data.GetData(GetType(UltraGridRow)), UltraGridRow)

                ' Convert screen coords → client coords for grdSOTCART1
                Dim clientPoint As Point = grdSOTCART1.PointToClient(New Point(e.X, e.Y))

                ' Find the UIElement under the cursor
                Dim element As UIElement = grdSOTCART1.DisplayLayout.UIElement.ElementFromPoint(clientPoint)
                If element Is Nothing Then Return

                ' Get the row under the cursor
                Dim targetRow As UltraGridRow = CType(element.GetContext(GetType(UltraGridRow)), UltraGridRow)

                If targetRow IsNot Nothing Then
                    Dim CART_NO_TARGET As String = targetRow.Cells("CART_NO").Value & String.Empty

                    Dim CART_NO_SOURCE As String = sourceRow.Cells("CART_NO").Value & String.Empty
                    Dim CART_NO_LNO_SOURCE As String = sourceRow.Cells("CART_LNO").Value & String.Empty

                    If CART_NO_TARGET = CART_NO_SOURCE Then
                        Exit Sub
                    End If

                    EnforceConstraints(False)
                    Dim drSOTCART2 As DataRow = dst.Tables("SOTCART2").Rows.Find({CART_NO_SOURCE, CART_NO_LNO_SOURCE})
                    If drSOTCART2 IsNot Nothing Then
                        drSOTCART2.Item("CART_LNO") = Val(dst.Tables("SOTCART2").Compute("MAX(CART_LNO)", $"CART_NO = '{CART_NO_TARGET}'") & String.Empty) + 1
                        drSOTCART2.Item("CART_NO") = CART_NO_TARGET
                    End If
                    EnforceConstraints(True)
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

#End Region

End Class