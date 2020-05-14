Imports System.Drawing

Public Class SOFINVH1


    Dim CUST_CODE As String
    Dim CUST_BILL_TO_CUST As String
    Dim SO_ORDER_NO As String
    Dim SO_ORDER_NO_init As String
    Dim BATCH_NO As String
    Dim SHIP_CODE_ORIG As String
    Dim CURR_CODE As String
    Dim CURR_EXCH_RATE As Decimal
    Dim ORDR_TYPE_CODE As String
    Dim use_CUST_PU_DATE As Boolean

    Dim rowARTCUST1_SOLDTO As DataRow
    Dim rowSOTORDR1 As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Get_PARM("ARTPARM1")
        Get_PARM("APTPARM1")
        Get_PARM("SOTPARM1")
        Get_PARM("ICTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            ASCMAIN1.sql = "Select SOTINVH1.*" & vbCrLf _
            & " from SOTINVH1 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" _
            & " and ORDR_INV_TYPE = 'I' and ORDR_INV_REG IS NULL"
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "SOTINVHW", "*")

            ASCMAIN1.sql = "Select SOTORDR1.*" _
            & " from SOTORDR1 where SO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, True, "V", 1)
            Create_TDA(.Tables.Add, "SOTINVH1", "*")

            Create_TDA(.Tables.Add, "SOTORDR2", "*", 1)
            Create_TDA(.Tables.Add, "SOTINVH2", "*")

            Create_TDA(.Tables.Add, "SOTORDR5", "*", 1)
            Create_TDA(.Tables.Add, "SOTINVH5", "*")

            Create_TDA(.Tables.Add, "SOTORDRB", "*", 1)
            Create_TDA(.Tables.Add, "SOTINVHB", "*")

            ASCMAIN1.sql = "Select * from SOTORDR3 where SO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR3", "**", 0, True, "V")
            .Tables("SOTORDR3").Columns.Add("ACK")

            Create_TDA(.Tables.Add, "SOTINVH3", "*")

            Create_Relation("SOTORDR2", "SOTORDR3", "SO_ORDER_NO,SO_ORDER_LNO")
            Create_Relation("SOTORDR2", "SOTORDRB", "SO_ORDER_NO,SO_ORDER_LNO")
            Create_Relation("SOTORDR1", "SOTORDR2", "SO_ORDER_NO")

            With .Tables("SOTORDR2")
                .Columns.Add("SO_LOT_CASES", GetType(System.Int64), "SUM(CHILD(SOTORDR2_SOTORDR3).SO_LOT_CASES)")
                .Columns.Add("SO_LOT_UNITS", GetType(System.Decimal), "SUM(CHILD(SOTORDR2_SOTORDR3).SO_LOT_UNITS)")
                .Columns.Add("LINE_AMOUNT_LB", GetType(System.Decimal), "ORDR_PRICE_GRS * SO_LOT_UNITS")
                .Columns.Add("LINE_AMOUNT_CS", GetType(System.Decimal), "ORDR_PRICE_GRS_CS * SO_LOT_CASES")
                .Columns.Add("LINE_AMOUNT", GetType(System.Decimal), "IIF(PRICE_UM='CS',LINE_AMOUNT_CS,LINE_AMOUNT_LB)")

                .Columns.Add("ORDR_PRICE_NET_CALC", GetType(System.Decimal))
                .Columns.Add("ORDR_PRICE_NET_CALC_NO_BRKR", GetType(System.Decimal))
                .Columns.Add("LINE_AMOUNT_NET_NO_BRKR", GetType(System.Decimal), "ORDR_PRICE_NET_CALC_NO_BRKR * SO_LOT_UNITS")
                .Columns.Add("LINE_AMOUNT_ALLOW", GetType(System.Decimal), "ALLOW_RATE * SO_LOT_UNITS")
            End With

            With .Tables("SOTORDRB")
                .Columns.Add("SO_LOT_CASES", GetType(System.Decimal), "PARENT(SOTORDR2_SOTORDRB).SO_LOT_CASES")
                .Columns.Add("SO_LOT_UNITS", GetType(System.Decimal), "PARENT(SOTORDR2_SOTORDRB).SO_LOT_UNITS")
                .Columns.Add("LINE_AMOUNT", GetType(System.Decimal), "PARENT(SOTORDR2_SOTORDRB).LINE_AMOUNT")
                .Columns.Add("LINE_AMOUNT_NET_NO_BRKR", GetType(System.Decimal), "PARENT(SOTORDR2_SOTORDRB).LINE_AMOUNT_NET_NO_BRKR")
                .Columns.Add("CONT_TP_AMT_CALC", GetType(System.Decimal), _
                               "IIF(CONT_EXCLUDED = '1',0," _
                             & "IIF(CONT_TP_UOM = '$/LB',CONT_TP_RATE * SO_LOT_UNITS, " _
                             & "IIF(CONT_TP_UOM = '$/CS',CONT_TP_RATE * SO_LOT_CASES, " _
                             & "IIF(CONT_TP_UOM = '%',IIF(CONT_TP_TYPE = 'B',LINE_AMOUNT_NET_NO_BRKR,LINE_AMOUNT) * CONT_TP_RATE / 100,0))))")

                .Columns.Add("RATE_CALC", GetType(System.Decimal), "IIF(SO_LOT_UNITS = 0, 0, CONT_TP_AMT_CALC / SO_LOT_UNITS)")
                .Columns.Add("RATE_MAX", GetType(System.Decimal))
                '.Columns.Add("RATE", GetType(System.Decimal), "IIF(SO_LOT_UNITS =0, 0, IIF(ISNULL(RATE_MAX,0)=0 OR ISNULL(RATE_MAX,0)>=RATE_CALC, RATE_CALC, RATE_MAX))")
                .Columns.Add("RATE", GetType(System.Decimal), "IIF(SO_LOT_UNITS =0, 0, IIF(ISNULL(RATE_MAX,0)<>0 AND ISNULL(RATE_MAX,0)<=RATE_CALC, RATE_MAX, RATE_CALC))")

                .Columns.Add("REBATE", GetType(System.Decimal), "IIF(CONT_TP_TYPE = 'R',RATE,0)")
                .Columns.Add("ALLOW_RATE", GetType(System.Decimal), "IIF((CONT_TP_TYPE = 'A' OR CONT_TP_TYPE = 'O'),RATE,0)")
                .Columns.Add("BRKR_RATE", GetType(System.Decimal), "IIF(CONT_TP_TYPE = 'B',RATE,0)")
                .Columns.Add("FUND_RATE", GetType(System.Decimal), "IIF(CONT_TP_TYPE = 'F',RATE,0)")
            End With

            With .Tables("SOTORDR2")
                .Columns.Add("REBATE_B", GetType(System.Decimal), "SUM(CHILD(SOTORDR2_SOTORDRB).REBATE)")
                .Columns.Add("ALLOW_RATE_B", GetType(System.Decimal), "SUM(CHILD(SOTORDR2_SOTORDRB).ALLOW_RATE)")
                .Columns.Add("BRKR_RATE_B", GetType(System.Decimal), "SUM(CHILD(SOTORDR2_SOTORDRB).BRKR_RATE)")
                .Columns.Add("FUND_RATE_B", GetType(System.Decimal), "SUM(CHILD(SOTORDR2_SOTORDRB).FUND_RATE)")

                .Columns.Add("FRT_RATE", GetType(System.Decimal), "PARENT(SOTORDR1_SOTORDR2).FRT_RATE")

                .Columns("ORDR_PRICE_NET_CALC").Expression = "IIF(ISNULL(ORDR_PRICE_GRS,0)=0" _
                             & ",ISNULL(ORDR_PRICE_GRS,0) - ISNULL(REBATE,0) - ISNULL(ALLOW_RATE,0) - ISNULL(BRKR_RATE,0) - ISNULL(FUND_RATE,0)" _
                             & ",ISNULL(ORDR_PRICE_GRS,0) - ISNULL(REBATE,0) - ISNULL(ALLOW_RATE,0) - ISNULL(BRKR_RATE,0) - ISNULL(FUND_RATE,0) - ISNULL(SVC_CHG_RATE,0) - ISNULL(FRT_RATE,0))"
                .Columns("ORDR_PRICE_NET_CALC_NO_BRKR").Expression = "IIF(ISNULL(ORDR_PRICE_GRS,0)=0" _
                             & ",ISNULL(ORDR_PRICE_GRS,0) - ISNULL(REBATE,0) - ISNULL(ALLOW_RATE,0) - ISNULL(FUND_RATE,0)" _
                             & ",ISNULL(ORDR_PRICE_GRS,0) - ISNULL(REBATE,0) - ISNULL(ALLOW_RATE,0) - ISNULL(FUND_RATE,0) - ISNULL(SVC_CHG_RATE,0) - ISNULL(FRT_RATE,0))"
            End With

            With .Tables("SOTORDR3")
                .Columns.Add("PROD_CODE", GetType(System.String), "PARENT(SOTORDR2_SOTORDR3).PROD_CODE")
                .Columns.Add("SIZE_CODE", GetType(System.String), "PARENT(SOTORDR2_SOTORDR3).SIZE_CODE")
                .Columns.Add("LINE_ITEM_DESCR", GetType(System.String), "PARENT(SOTORDR2_SOTORDR3).LINE_ITEM_DESCR")
                ' XXXX
                .Columns.Add("PACK_DESC")
                .Columns.Add("ITEM_CODE", GetType(System.String), "PARENT(SOTORDR2_SOTORDR3).ITEM_CODE")
            End With

            With .Tables("SOTORDR1")
                .Columns.Add("QTY_CASES", GetType(System.Int64), "SUM(CHILD(SOTORDR1_SOTORDR2).SO_LOT_CASES)")
                .Columns.Add("QTY_UNITS", GetType(System.Decimal), "SUM(CHILD(SOTORDR1_SOTORDR2).SO_LOT_UNITS)")
                .Columns.Add("ORDR_AMT_GROSS", GetType(System.Decimal), "SUM(CHILD(SOTORDR1_SOTORDR2).LINE_AMOUNT)")
                .Columns.Add("ORDR_AMT_ALLOW", GetType(System.Decimal), "SUM(CHILD(SOTORDR1_SOTORDR2).LINE_AMOUNT_ALLOW)")
            End With

            ASCMAIN1.sql = "Select ARTCUST1.*" _
            & " from ARTCUST1 where CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "V", 1)

            With .Tables.Add("SOTORDRT")
                .Columns.Add("TOTAL_SEQ", GetType(System.Int32))
                .Columns.Add("TOTAL_DESC", GetType(System.String))
                .Columns.Add("TOTAL_AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() {.Columns("TOTAL_SEQ")}
            End With

            ASCMAIN1.sql = "Select TRUNC(100 * ICTIREC2.REC_UNITS / ICTIREC2.REC_CASES) / 100 CATCH_WEIGHT" & vbCrLf _
            & "  from ICTIREC2, ICTLOTD1" & vbCrLf _
            & " Where ICTIREC2.RECEIPT_NO = ICTLOTD1.RECEIPT_NO" & vbCrLf _
            & "   and ICTIREC2.RECEIPT_LNO = ICTLOTD1.RECEIPT_LNO" & vbCrLf _
            & "   and ICTLOTD1.WHSE_CODE = :PARM1" & vbCrLf _
            & "   and ICTLOTD1.LOT_NO = :PARM2" & vbCrLf _
            & "   and ICTLOTD1.LOT_SEQ_NO = :PARM3"
            Create_TDA(.Tables.Add, "ICTIREC2_PF", "**", 0, False, "VVN")

            ASCMAIN1.sql = "Select ICTIREC1.CA_COMM_PCT, ICTIREC1.CA_VA_COMM_PCT" & vbCrLf _
            & ", ICTIREC1.CA_EXP_RECOVERY, ICTIREC1.CA_NO_OCEAN_FRT_RECOVER, ICTIREC1.CA_EXP_CAP_PCT" & vbCrLf _
            & ", ICTIREC1.CA_ADV_PCT_SP, ICTIREC1.CA_AMT_LB" & vbCrLf _
            & ", ICTIREC2.CA_VALUE_ADD_IND, ICTCOSTZ.OCEAN_FRT" & vbCrLf _
            & ", ICTIREC1.JOINT_VENTURE, ICTIREC1.JOINT_VENTURE_PCT " & vbCrLf _
            & " from ICTIREC2, ICTIREC1, ICTCOSTZ " & vbCrLf _
            & " where ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
            & "   and ICTIREC2.RECEIPT_NO = :PARM1 " & vbCrLf _
            & "   and ICTIREC2.RECEIPT_LNO = :PARM2 " & vbCrLf _
            & "   and ICTCOSTZ.WHSE_CODE = ICTIREC2.WHSE_CODE " & vbCrLf _
            & "   and ICTCOSTZ.LOT_NO = ICTIREC2.LOT_NO " & vbCrLf _
            & "   and ICTCOSTZ.LOT_SEQ_NO = ICTIREC2.LOT_SEQ_NO "
            Create_TDA(.Tables.Add, "ICTIRECX_CA", "**", 0, False, "VN")

            ASCMAIN1.sql = "Select POTCATG1.* from POTCATG1"
            Create_TDA(.Tables.Add, "POTCATG1", "**", 0)
            ASCMAIN1.sql = "Select SOTSVCG1.* from SOTSVCG1"
            Create_TDA(.Tables.Add, "SOTSVCG1", "**", 0)
            ASCMAIN1.sql = "Select ICTPACK1.* from ICTPACK1"
            Create_TDA(.Tables.Add, "ICTPACK1", "**", 0)

            Create_Relation("ICTPACK1", "SOTORDR3", "PACK_CODE")
            .Tables("SOTORDR3").Columns("PACK_DESC").Expression = "PARENT(ICTPACK1_SOTORDR3).PACK_DESC"

            ASCMAIN1.sql = "SELECT NVL(WHSE_XFR_CHG_EXP,0) WHSE_XFR_CHG_EXP" _
            & ", NVL(WHSE_WD_CHG_EXP,0) WHSE_WD_CHG_EXP, VEND_CODE" _
            & " FROM ICTWHSE1 WHERE WHSE_CODE = (SELECT " _
            & "  DECODE(NVL(ICTWHSE1.WHSE_MIN_STORAGE,0),0,SOTSDIV1.DIV_WHSE_CODE,ICTWHSE1.WHSE_CODE)" _
            & "  RATE_WHSE FROM ICTWHSE1,SOTTERR1,SOTSDIV1" _
            & " WHERE SOTTERR1.TERR_CODE (+) = ICTWHSE1.TERR_CODE" _
            & "   AND SOTSDIV1.DIVISION_CODE (+) = SOTTERR1.DIVISION_CODE" _
            & "   AND ICTWHSE1.WHSE_CODE = :PARM1)"
            Create_TDA(.Tables.Add, "ICTWHSE1_RATE", "**", 0, False, "V")

            ASCMAIN1.sql = "Select Distinct REL_ORDER_NO from EDT945T1" _
            & " where SO_ORDER_NO = :PARM1 AND WHSE_CODE = :PARM2" _
            & " and REL_ORDER_NO is NOT NULL"
            Create_TDA(.Tables.Add, "EDT945T1_REL_ORDER_NO", "**", 0, False, "VV")

            ASCMAIN1.sql = "Select Distinct BILL_OF_LADING_NO from EDT945T1" _
            & " where SO_ORDER_NO = :PARM1 AND WHSE_CODE = :PARM2" _
            & " and BILL_OF_LADING_NO is NOT NULL"
            Create_TDA(.Tables.Add, "EDT945T1_BILL_OF_LADING_NO", "**", 0, False, "VV")

            Create_TDA(.Tables.Add, "EDT810O1", "*")
            Create_TDA(.Tables.Add, "EDT810O2", "*")
            Create_TDA(.Tables.Add, "EDT810O3", "*")
            Create_TDA(.Tables.Add, "EDT810O5", "*")
            Create_TDA(.Tables.Add, "EDT810O7", "*")
            Create_TDA(.Tables.Add, "EDT810O8", "*")

            ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.CUST_CITY || ', ' || ARTCUST1.CUST_STATE CUST_CSZ from ARTCUST1 " _
            & " where (ARTCUST1.CUST_CODE = :PARM1" _
            & "    or  ARTCUST1.CUST_CODE = (Select CUST_BILL_TO_CUST from ARTCUST1 where CUST_CODE = :PARM1)" _
            & "    or  ARTCUST1.CUST_CODE in (Select CUST_BILL_TO_CUST from ARTCUST4 where CUST_CODE = :PARM1))"
            Create_TDA(.Tables.Add, "ARTCUSTX", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select CUST_SHIP_TO_CODE, CUST_SHIP_TO_NAME, CUST_SHIP_TO_ADDR1, CUST_SHIP_TO_ADDR2," _
            & " CUST_SHIP_TO_ADDR3, CUST_SHIP_TO_CITY, CUST_SHIP_TO_STATE, CUST_SHIP_TO_ZIP_CODE, " _
            & " CUST_SHIP_TO_COUNTRY, CUST_SHIP_TO_CONTACT, CUST_SHIP_TO_PHONE, CUST_ROUTING_INST from ARTCUST2 " _
            & " where CUST_CODE = :PARM1 and CUST_SHIP_TO_CODE <> 'BT'"
            Create_TDA(.Tables.Add, "ARTCUSTS", "**", 0, False, "V")

            ASCMAIN1.sql = "Select ICTLOTD1.* from ICTLOTD1 " _
            & " where WHSE_CODE = :PARM1 and LOT_NO = :PARM2 and LOT_SEQ_NO = :PARM3"
            Create_TDA(.Tables.Add, "ICTLOTD1", "**", 0, False, "VVN")

            Create_TDA(.Tables.Add, "ARTCUST6", "*")
            Create_TDA(.Tables.Add, "APTINVH1", "*")
            Create_TDA(.Tables.Add, "APTINVH2", "*")

        End With

        Fill_Records("POTCATG1")
        Fill_Records("SOTSVCG1")
        Fill_Records("ICTPACK1")

        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")
        grdSOTORDRT.DataSource = dst.Tables("SOTORDRT")
        grdSOTORDR3.DataSource = dst.Tables("SOTORDR3")

        cmbCUST_BILL_TO_CUST.DataSource = dst.Tables("ARTCUSTX")
        cmbCUST_SHIP_TO_CODE.DataSource = dst.Tables("ARTCUSTS")

        Create_Summary(grdSOTINVHX, "SO_ORDER_NO", "Count")
        Create_Summary(grdSOTINVHX, "ORDR_TOTAL_AMT")

        Create_Summary(grdSOTORDR3, "SO_LOT_CASES")
        Create_Summary(grdSOTORDR3, "SO_LOT_CASES_ORIG")
        Create_Summary(grdSOTORDR3, "SO_LOT_CASES_SHIP")
        Create_Summary(grdSOTORDR3, "SO_LOT_UNITS")

        Set_Read_Only(grpInvoice, True)

        Show_Filter(grdSOTINVHX)

        With grdSOTINVHX.DisplayLayout.Bands("SOTINVHX")
            .Columns("SO_ORDER_NO").Header.Fixed = True
            .Columns("ORDR_TOTAL_AMT").Header.Fixed = True
            .Columns("ORDR_INV_NO").Header.Fixed = True
            .Columns("ORDR_INV_DATE").Header.Fixed = True
            .Columns("ORDR_DIV_CODE").Header.Fixed = True
        End With

        With grdSOTORDR3.DisplayLayout.Bands("SOTORDR3")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                Select gcol.Key
                    Case "SO_LOT_CASES"
                        'gcol.Hidden = (MENU_ITEM_OBJECT = "SOFINVHA")
                        gcol.CellAppearance.BackColor = Drawing.Color.LightGreen
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                        gcol.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
                    Case "SO_LOT_UNITS"
                        gcol.CellAppearance.BackColor = Drawing.Color.Beige
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    Case "SO_LOT_CASES_ORIG"
                        gcol.CellAppearance.BackColor = Drawing.Color.LightYellow
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        'Case "SO_LOT_CASES_SHIP"
                        '    gcol.Hidden = (MENU_ITEM_OBJECT <> "SOFINVHA")
                        '    gcol.CellAppearance.BackColor = Drawing.Color.LightGreen
                        '    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                        '    gcol.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
                    Case "ACK"
                        gcol.Hidden = (MENU_ITEM_OBJECT = "SOFINVHA")
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    Case "INVOICE_APPROVAL_COMMENTS"
                        If MENU_ITEM_OBJECT = "SOFINVHA" Then
                            gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                            gcol.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
                        Else
                            gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        End If
                    Case Else
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End Select
            Next
        End With

        With grdSOTORDRT.DisplayLayout.Bands("SOTORDRT")
            .Columns("TOTAL_DESC").CellAppearance.BackColor = Drawing.Color.LightGray
        End With

        ASCMAIN1.Add_Value_List(grdSOTINVHX, "ORDR_CREDIT_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTINVHX, "ORDR_TYPE_CODE")
        ASCMAIN1.Add_Value_List(grdSOTINVHX, "TRANSFER_TYPE")

        cbeORDR_TYPE_CODE.DataSource = ASCDATA1.GetDataTable("Select T_CODE ORDR_TYPE_CODE, T_DESC ORDR_TYPE_DESC from ASTCODE1 where COLUMN_NAME = 'ORDR_TYPE_CODE' and TABLE_NAME = 'SOTORDR1'")
        cbeDIVISION_CODE.DataSource = ASCDATA1.GetDataTable("Select DIVISION_CODE, DIVISION_NAME from SOTSDIV1 order by DIVISION_CODE")

        Dim DATE_SHIPPED As Date = Now.Date.AddDays(-1)
        If DATE_SHIPPED.DayOfWeek = System.DayOfWeek.Saturday Then
            DATE_SHIPPED = DATE_SHIPPED.AddDays(-1)
        ElseIf DATE_SHIPPED.DayOfWeek = System.DayOfWeek.Sunday Then
            DATE_SHIPPED = DATE_SHIPPED.AddDays(-2)
        End If
        dteDATE_SHIPPED.Value = DATE_SHIPPED

        If MENU_ITEM_OBJECT = "SOFINVHA" Then
            optShow.Value = "R"
        Else
            optShow.Value = "A"
        End If

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select Order"

                If dteDATE_SHIPPED.Value & "" = "" Then
                    EMsg &= vbCr & "No value specified for Invoice Date"
                End If

                SO_ORDER_NO = ""

                If Absx1.txtFor("SO_ORDER_NO").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Valid Order No"
                Else
                    SO_ORDER_NO = Absx1.txtFor("SO_ORDER_NO").Text
                    SO_ORDER_NO_init = SO_ORDER_NO

                    If SO_ORDER_NO.Length < 6 Then
                        SO_ORDER_NO = SO_ORDER_NO.PadLeft(6, "0")
                    End If

                    rowSOTORDR1 = LookUp("SOTORDR1", SO_ORDER_NO)
                    If rowSOTORDR1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Order No (" & SO_ORDER_NO & ")"
                    Else
                        If rowSOTORDR1.Item("SO_STATUS_CODE") & "" <> "O" Then
                            EMsg &= vbCr & "Order is not Open"
                        End If
                        If rowSOTORDR1.Item("ORDR_CREDIT_STATUS") & "" <> "A" Then
                            EMsg &= vbCr & "Order is not Credit Approved"
                        End If
                        If rowSOTORDR1.Item("ORDR_PRINTED") & "" <> "1" Then
                            EMsg &= vbCr & "Order is not Printed"
                        End If
                        If rowSOTORDR1.Item("BILLING_APPR_BY") & "" = "" And MENU_ITEM_OBJECT = "SOFINVH1" Then
                            EMsg &= vbCr & "Order was not Approved for Invoicing"
                        End If
                        If rowSOTORDR1.Item("ORDR_REL") & "" <> "1" Then
                            If EMsg = "" And rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "T" Then
                                If MsgBox("OK to Proceed Anyway", MsgBoxStyle.YesNo, _
                                          "Transfer Order was not Released") = MsgBoxResult.No Then
                                    Exit Sub
                                End If
                            Else
                                EMsg &= vbCr & "Order was not Released"
                            End If
                        End If

                        SHIP_CODE_ORIG = rowSOTORDR1.Item("SHIP_CODE") & ""
                        BATCH_NO = rowSOTORDR1.Item("BATCH_NO") & ""
                        ORDR_TYPE_CODE = rowSOTORDR1.Item("ORDR_TYPE_CODE")

                        'If ORDR_TYPE_CODE <> "S" Then
                        '    EMsg &= vbCr & "Need to Test with ABS"
                        'End If
                        If BATCH_NO <> "" Then
                            'EMsg &= vbCr & "Need to Test with ABS"
                        End If

                        ASCMAIN1.sql = "Select count (*) from SOTORDR3,ICTLOTD2" & vbCrLf _
                        & " where " & IIf(BATCH_NO = "", _
                                          "SO_ORDER_NO = '" & SO_ORDER_NO & "'", _
                                          "SO_ORDER_NO in (Select SO_ORDER_NO from SOTORDR1 " _
                                          & " where BATCH_NO = '" & rowSOTORDR1.Item("BATCH_NO") & "')") & vbCrLf _
                        & "   and SOTORDR3.WHSE_CODE = ICTLOTD2.WHSE_CODE" & vbCrLf _
                        & "   and SOTORDR3.LOT_NO = ICTLOTD2.LOT_NO" & vbCrLf _
                        & "   and SOTORDR3.LOT_SEQ_NO = ICTLOTD2.LOT_SEQ_NO" & vbCrLf _
                        & "   and ICTLOTD2.ON_HOLD_FLAG is Not Null" ' null = May Proceed
                        If Val(ASCDATA1.GetDataValue) <> 0 Then
                            If rowSOTORDR1.Item("ORDR_TYPE_CODE") = "S" Then
                                EMsg &= vbCr & "Order includes lots on FDA Hold"
                            Else
                                If EMsg = "" Then
                                    ' For T=Transfer and P=RTV Only
                                    If MsgBox("Proceed Anyway?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, _
                                              "Order includes lots on FDA Hold") = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If

                        Dim ORDR_DIV_CODE = rowSOTORDR1.Item("ORDR_DIV_CODE") & ""
                        If ORDR_DIV_CODE <> "A" And ORDR_DIV_CODE <> "T" Then
                            EMsg &= vbCr & "Order Division is " & ORDR_DIV_CODE & vbCr & "Please contact ABS"
                        End If

                    End If
                End If

                If EMsg = "" Then
                    If BATCH_NO <> "" Then
                        ASCMAIN1.sql = "Select * from SOTORDR1 where BATCH_NO = '" & BATCH_NO & "'"
                        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                            For Each COLUMN_NAME As String In New String() _
                                {"ORDR_REL", "ORDR_CREDIT_STATUS", "SO_STATUS_CODE", "ORDR_PRINTED"}
                                If row.Item(COLUMN_NAME) & "" <> rowSOTORDR1.Item(COLUMN_NAME) & "" Then
                                    EMsg &= vbCr & "Batch Invoicing Problem with Column " & COLUMN_NAME
                                End If
                            Next

                            If row.Item("CURR_CODE") & "" <> "USD" Then EMsg &= vbCr & "Cannot Batch Invoice in non-USD"

                            If EMsg <> "" Then
                                EMsg &= "- See Order " & row.Item("SO_ORDER_NO")
                                ASCMAIN1.MultiTask_Release()
                                Exit For
                            End If

                            If Not ASCMAIN1.Logical_Lock("SOTORDR1", row.Item("SO_ORDER_NO")) Then Exit Sub
                        Next
                    Else
                        If Not ASCMAIN1.Logical_Lock("SOTORDR1", SO_ORDER_NO) Then Exit Sub
                    End If
                End If

                If EMsg <> "" Then
                    Absx1.txtFor("SO_ORDER_NO").Text = ""
                    Absx1.txtFor("SO_ORDER_NO").Focus()
                Else
                    ASCMAIN1.sql = "Select * from SOTORDR3 " _
                    & " where " & IIf(BATCH_NO = "", _
                                          "SO_ORDER_NO = '" & SO_ORDER_NO & "'", _
                                          "SO_ORDER_NO in (Select SO_ORDER_NO from SOTORDR1 " _
                                          & " where BATCH_NO = '" & rowSOTORDR1.Item("BATCH_NO") & "')") & vbCrLf _
                    & " and PACK_CODE = '" & TAC.TACMAIN1.CATCH_PACK & "'"
                    Dim row As DataRow = ASCDATA1.GetDataRow
                    If row IsNot Nothing Then
                        MsgBox("There are Catch Weight Lots on this Order", MsgBoxStyle.OkOnly, _
                               "Please Remember to Enter the Exact Weights")
                    End If
                End If

            Case "Update"

                'TAC.SOCMAIN1.Load_PAYEE_into_Contracts(Me, CUST_BILL_TO_CUST) ' to refresh CUST_BILL_TO_CUST if it had changed, and to pick up fixes to Brokerage Contracts or Vendor Codes loaded into Broker Records in response to "invalid payee" messages
                'EMsg &= TAC.SOCMAIN1.Check_Vendor(Me)

                If CURR_CODE <> "USD" Then
                    Dim R As Decimal = Get_CURR_EXCH_RATE()
                    If R = 0 Then
                        EMsg &= vbCr & "No Currency Exchange Rate for " & CURR_CODE & " on file for " & Format(Absx1.dteFor("ORDR_INV_DATE").Value, "MM/dd/yyyy")
                    End If
                End If

                Dim SO_PARM_CATCH_WEIGHT_VAR_PCT As Decimal = Val(ROWs("SOTPARM1").Item("SO_PARM_CATCH_WEIGHT_VAR_PCT") & "")

                Dim SHIP_CODE As String = Absx1.txtFor("SHIP_CODE").Text
                Dim FRT_TERMS As String = Absx1.txtFor("FRT_TERMS").Text
                Dim FRT_RATE As Decimal = Val(Absx1.numFor("FRT_RATE").Value & "")

                For Each rowSOTORDR3 As DataRow In dst.Tables("SOTORDR3").Rows
                    Dim SO_LOT_CASES As Int64 = Val(rowSOTORDR3("SO_LOT_CASES") & "")
                    Dim SO_LOT_CASES_ORIG As Int64 = Val(rowSOTORDR3("SO_LOT_CASES_ORIG") & "")
                    Dim SO_LOT_UNITS As Decimal = Val(rowSOTORDR3("SO_LOT_UNITS") & "")

                    Dim WHSE_CODE As String = rowSOTORDR3("WHSE_CODE")
                    Dim LOT_NO As String = rowSOTORDR3("LOT_NO")
                    Dim LOT_SEQ_NO As Int64 = Val(rowSOTORDR3("LOT_SEQ_NO") & "")
                    Dim SO_ORDER_LNO As Int64 = Val(rowSOTORDR3("SO_ORDER_LNO") & "")
                    Dim SO_LOT_LNO As Int64 = Val(rowSOTORDR3("SO_LOT_LNO") & "")
                    Dim PACK_CODE As String = rowSOTORDR3("PACK_CODE")


                    If SO_LOT_CASES <> SO_LOT_CASES_ORIG Then
                        If rowSOTORDR3.Item("ACK") & "" <> "1" Then
                            EMsg &= vbCr & "You Must Acknowledge That you have reviewed the dicrepency on Lot: " & LOT_NO
                            Exit For
                        End If
                    End If

                    If System.Math.Sign(SO_LOT_CASES) _
                    <> System.Math.Sign(SO_LOT_UNITS) Then
                        EMsg &= vbCr & String.Format("Cases/Units Problem in Lot {0} (See Line {1})", New String() {LOT_NO, CStr(SO_ORDER_LNO)})
                        'EMsg &= vbCr & "Cases/Units Problem in Lot " & LOT_NO & " (See Line " & CStr(SO_ORDER_LNO) & ")"
                        Exit For
                    End If

                    If SO_LOT_CASES <> 0 Then
                        If PACK_CODE = TAC.TACMAIN1.CATCH_PACK Then
                            If System.Math.Round(SO_LOT_UNITS, 0) = 0 Then
                                EMsg &= vbCr & "No Units Entered for catch weight Lot " & LOT_NO
                                Exit For
                            Else
                                Dim rowICTIREC2_PF As DataRow = Fill_Record("ICTIREC2_PF", _
                                                                New String() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})
                                Dim CATCH_WEIGHT_received As Decimal = Val(rowICTIREC2_PF.Item("CATCH_WEIGHT") & "")
                                Dim CATCH_WEIGHT_this_sale As Decimal = SO_LOT_UNITS / SO_LOT_CASES
                                If 100 * System.Math.Abs((CATCH_WEIGHT_this_sale - CATCH_WEIGHT_received) / CATCH_WEIGHT_received) > SO_PARM_CATCH_WEIGHT_VAR_PCT Then
                                    If MsgBox(String.Format("Units per Case for Lot {0} ({1})" & vbCrLf _
                                              & " varies with Original Receipt ({2})" & vbCrLf _
                                              & " by more than {3}%" & vbCrLf & vbCrLf _
                                              & "Continue with Update?", _
                                              New String() {LOT_NO, _
                                                            Format$(CATCH_WEIGHT_this_sale, "##0.00"), _
                                                            Format$(CATCH_WEIGHT_received, "###.00"), _
                                                            Format(SO_PARM_CATCH_WEIGHT_VAR_PCT, "###.0")}), _
                                              MsgBoxStyle.YesNo, _
                                              "Catch Weight Sale has Large Variance with Respect to Original Catch Weight") = MsgBoxResult.No Then
                                        EMsg &= vbCr & "Fix Units/Cases Ratio for Lot " & LOT_NO
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next

                Validate_Code("FRT_TERMS")

                If FRT_TERMS = "TIS" And SHIP_CODE <> "TIS" Then
                    EMsg &= vbCr & "Ship Via Code should be TIS if Frt Terms is TIS"
                End If

                If SHIP_CODE_ORIG <> "TIS" And SHIP_CODE = "TIS" Then
                    EMsg &= vbCr & "Cannot Change Ship Via to TIS (Release Order was Generated); Re-Release as a TIS"
                ElseIf SHIP_CODE_ORIG = "TIS" And SHIP_CODE <> "TIS" Then
                    EMsg &= vbCr & "Cannot Change Ship Via from TIS (No Release Order was Generated)"
                End If

                ASCMAIN1.sql = "Select SOTSVIA1.*, APTVEND1.VEND_CODE VEND_CODE_AP" & vbCrLf _
                & " from SOTSVIA1,APTVEND1" & vbCrLf _
                & " where SOTSVIA1.SHIP_CODE = :PARM1" & vbCrLf _
                & " and APTVEND1.VEND_CODE (+) = SOTSVIA1.VEND_CODE"
                Dim rowSOTSVIA1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {SHIP_CODE})
                'Dim rowSOTSVIA1 As DataRow = Fill_Record("SOTSVIA1", SHIP_CODE)
                If rowSOTSVIA1 Is Nothing Then
                    EMsg &= vbCr & "This order does not have a VALID Ship Via Code"
                Else
                    If rowSOTSVIA1.Item("VEND_CODE_AP") & "" = "" And FRT_TERMS = "PPD" Then
                        EMsg &= vbCr & "Ship Via " & SHIP_CODE & " does not have a VALID Vendor Code (Required for PPD)"
                    End If
                End If

                If EMsg = "" Then
                    If ORDR_TYPE_CODE = "T" AndAlso rowSOTORDR1.Item("TRANSFER_TYPE") = "W" Then
                        Dim VEND_CODE_SVC As String = Absx1.txtFor("VEND_CODE_SVC").Text
                        Dim SVC_CHG_CODE As String = Absx1.txtFor("SVC_CHG_CODE").Text
                        Dim SVC_CHG_RATE As Decimal = Val(Absx1.numFor("SVC_CHG_RATE").Value & "")

                        If VEND_CODE_SVC <> "" And SVC_CHG_CODE <> "" And SVC_CHG_RATE <> 0 Then
                        Else
                            Dim msg As String = ""
                            If VEND_CODE_SVC <> "" And SVC_CHG_CODE = "" And SVC_CHG_RATE = 0 Then
                                msg = "You have specified a Service Charge Vendor but no Service Charge."
                            ElseIf VEND_CODE_SVC <> "" And SVC_CHG_CODE <> "" And SVC_CHG_RATE = 0 Then
                                msg = "You have specified a Service Charge Vendor and Code but no Service Charge Rate."
                            ElseIf VEND_CODE_SVC <> "" And SVC_CHG_CODE = "" And SVC_CHG_RATE = 0 Then
                                msg = "You have specified a Service Charge Code but no Service Charge Vendor and Rate."
                            ElseIf SVC_CHG_RATE <> 0 Then
                                If VEND_CODE_SVC = "" Then
                                    EMsg &= vbCr & "You Must Specify a Service Charge Vendor with a Service Charge Rate"
                                End If
                                If SVC_CHG_CODE = "" Then
                                    EMsg &= vbCr & "You Must Specify a Service Charge Code with a Service Charge Rate"
                                End If
                            End If
                            If EMsg = "" And msg <> "" Then
                                If MsgBox(msg & vbCrLf & vbCrLf _
                                         & "OK to proceed with update?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, _
                                         "OK to Proceed?") = MsgBoxResult.No Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    Else
                        If FRT_TERMS <> "TIS" And SHIP_CODE = "TIS" Then
                            If MsgBox("Ok to Proceed with Update", MsgBoxStyle.Question + MsgBoxStyle.YesNo, _
                                      "Frt Terms do not match Ship Via (TIS)") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                        If FRT_TERMS = "PPD" Then
                            If FRT_RATE = 0 Then
                                If ORDR_TYPE_CODE = "T" Then
                                    If MsgBox("Ok to Proceed with Update", MsgBoxStyle.Question + MsgBoxStyle.YesNo, _
                                              "No Freight Rate Specified") = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    EMsg &= vbCr & "You Must Specify a non-zero value for Freight Rate if using PPD"
                                End If
                            End If
                        Else
                            If FRT_RATE <> 0 Then
                                EMsg &= vbCr & "Freight Rate must be Zero if Freight Terms are not PPD"
                            End If
                        End If
                    End If
                End If

                If EMsg = "" Then
                    Dim priceWarningMsg As String = ""
                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ISNULL(ORDR_PRICE_NET_CALC,0) < 0")
                        If ORDR_TYPE_CODE = "T" AndAlso rowSOTORDR1.Item("TRANSFER_TYPE") = "W" Then
                        Else
                            priceWarningMsg &= vbCr & "There are lines on this order with a negative Net Price (See Line " & rowSOTORDR2.Item("SO_ORDER_LNO") & ")"
                        End If
                    Next

                    If priceWarningMsg <> "" Then
                        If MsgBox(priceWarningMsg & vbCrLf & "Ok to Proceed with Update?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, _
"Negative Net Price Warning") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Select Order"
                Load_Record()
                Mode_Settings(True)

            Case "Refresh"
                Load_SOTINVHX()

            Case "Cancel"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)
            Case "Approve"
                Approve_for_Invoicing()
                'Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1


                .Groups("Screen Control").Items("Approve").Visible = (MENU_ITEM_OBJECT = "SOFINVHA")
                .Groups("Screen Control").Items("Refresh").Visible = (MENU_ITEM_OBJECT <> "SOFINVHA")
                .Groups("Screen Control").Items("Update").Visible = (MENU_ITEM_OBJECT <> "SOFINVHA")
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Select Order").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Approve").Settings.Enabled = iScreenMode
                .Groups("Totals").Visible = ScreenMode
                .Groups("Show").Visible = Not ScreenMode And MENU_ITEM_OBJECT <> "SOFINVHA"
                .Groups("Batch Defaults").Visible = Not ScreenMode And MENU_ITEM_OBJECT <> "SOFINVHA"
            End With
        End If


        grpHeader.Visible = ScreenMode
        Absx1.optFor("TRANSFER_TYPE").Visible = ScreenMode AndAlso (ORDR_TYPE_CODE = "T")

        lblORDR_TYPE_CODE.Visible = ScreenMode And ORDR_TYPE_CODE <> "T"
        Absx1.cbeFor("ORDR_TYPE_CODE").Visible = ScreenMode And ORDR_TYPE_CODE <> "T"

        lblORDR_INV_NO.Visible = ScreenMode
        Absx1.txtFor("ORDR_INV_NO").Visible = ScreenMode
        lblORDR_INV_DATE.Visible = ScreenMode
        Absx1.dteFor("ORDR_INV_DATE").Visible = ScreenMode
        lblORDR_DATE_SHIPPED.Visible = ScreenMode
        Absx1.dteFor("ORDR_DATE_SHIPPED").Visible = ScreenMode

        lblORDR_DIV_CODE.Visible = ScreenMode
        cbeDIVISION_CODE.Visible = ScreenMode

        grdSOTINVHX.Visible = Not ScreenMode

        Set_Read_Only(grpHeader, True)
        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpSHIPTO, (MENU_ITEM_OBJECT <> "SOFINVHA"))
        Set_Read_Only_for_ctl(Absx1.optFor("TRANSFER_TYPE"), Not ScreenMode)
        Set_Read_Only(SplitContainer2.Panel1, False)
        Set_Read_Only_for_ctl(Absx1.txtFor("SO_ORDR_NOTES_EXT"), False)
        Set_Read_Only_for_ctl(Absx1.txtFor("CUST_ORDER_NO"), (MENU_ITEM_OBJECT <> "SOFINVHA"))
        Set_Read_Only_for_ctl(Absx1.txtFor("CUST_ORDER_NO_3PY"), (MENU_ITEM_OBJECT <> "SOFINVHA"))
        Set_Read_Only_for_ctl(Absx1.dteFor("ORDR_INV_DATE"), False)
        Set_Read_Only_for_ctl(Absx1.dteFor("ORDR_DATE_SHIPPED"), False)

        UltraTextEditor26.ReadOnly = True ' this is necessary because there is another control with CUST_NAME as the ABSColumnName, and stds & ABSReadOnly does not know how to differentiate

        tabHeader.Tabs("Batch Orders").Visible = (BATCH_NO <> "")
        tabHeader.Tabs("Batch Orders").Text = "Orders in Batch " & BATCH_NO

        If ScreenMode Then
        Else
            Clear_Record()
            Show_Filter(grdSOTINVHX, True)
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() {"SOTINVHX", "SOTORDRT" _
            , "ICTLOTD1", "SOTINVHW", "ARTCUST6", "ARTCUSTX", "ICTWHSE1_RATE" _
            , "SOTORDR1", "SOTORDR2", "SOTORDR3", "SOTORDR5", "SOTORDRB" _
            , "SOTINVH1", "SOTINVH2", "SOTINVH3", "SOTINVH5", "SOTINVHB"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)

        Load_SOTINVHX()
        Absx1.txtFor("SO_ORDER_NO").Focus()

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Selecting Order for Invoicing")

        EnforceConstraints(False)

        If BATCH_NO = "" Then
            Fill_Records("SOTORDR1", SO_ORDER_NO)
            Fill_Records("SOTORDR2", SO_ORDER_NO)
            Fill_Records("SOTORDR3", SO_ORDER_NO)
            Fill_Records("SOTORDR5", SO_ORDER_NO)
            Fill_Records("SOTORDRB", SO_ORDER_NO)
        Else
            Dim sqlB As String = "Select * from SOTORDR1 where SO_ORDER_NO in " _
            & " (Select SO_ORDER_NO from SOTORDR1 where BATCH_NO = '" & BATCH_NO & "')"

            Fill_Records("SOTORDR1", "", True, sqlB)
            Fill_Records("SOTORDR2", "", True, Replace(sqlB, "SOTORDR1", "SOTORDR2", 1, 1))
            Fill_Records("SOTORDR3", "", True, Replace(sqlB, "SOTORDR1", "SOTORDR3", 1, 1))
            Fill_Records("SOTORDR5", "", True, Replace(sqlB, "SOTORDR1", "SOTORDR5", 1, 1))
            Fill_Records("SOTORDRB", "", True, Replace(sqlB, "SOTORDR1", "SOTORDRB", 1, 1))

        End If

        EnforceConstraints(True)

        'If ASCMAIN1.Running_in_VS Then Stop ' IF INVOICING A BATCH, SET CURSOR OF SOTORDR1 TO SO_ORDER_NO 

        'Toggle_Transfer_Fields()

        use_CUST_PU_DATE = False
        ASCMAIN1.sql = "Select SOTTCLS1.MARKET_CODE from SOTTCLS1,ARTCUST1 " _
        & " where SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE and ARTCUST1.CUST_CODE = :PARM1"
        Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", Absx1.txtFor("CUST_CODE").Text & "")
        If row IsNot Nothing AndAlso row.Item("MARKET_CODE") & "" = "WHOLE" Then
            use_CUST_PU_DATE = True
        End If

        ' this routine appears (with slight code variations) in 3 places in this form
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Rows
            rowSOTORDR1.Item("ORDR_DATE_SHIPPED") = dteDATE_SHIPPED.Value
            If use_CUST_PU_DATE And rowSOTORDR1.Item("FRT_TERMS") <> "PPD" _
            And Format(rowSOTORDR1.Item("ORDR_DATE_SHIPPED"), "yyyyMMdd") _
             >= Format(rowSOTORDR1.Item("CUST_PU_DATE"), "yyyyMMdd") Then
                rowSOTORDR1.Item("ORDR_INV_DATE") = rowSOTORDR1.Item("CUST_PU_DATE")
            Else
                rowSOTORDR1.Item("ORDR_INV_DATE") = rowSOTORDR1.Item("ORDR_DATE_SHIPPED")
            End If
        Next

        CURR_CODE = rowSOTORDR1.Item("CURR_CODE")
        CURR_EXCH_RATE = Get_CURR_EXCH_RATE()
        Absx1.numFor("CURR_EXCH_RATE").Value = CURR_EXCH_RATE

        If CURR_CODE <> "USD" Then
            Stop
            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Rows
                rowSOTORDR2.Item("ORDR_PRICE_GRS") = rowSOTORDR2.Item("ORDR_PRICE_GRS_CURR") * CURR_EXCH_RATE
            Next
        End If

        If MENU_ITEM_OBJECT <> "SOFINVHA" Then
            For Each rowSOTORDR3 As DataRow In dst.Tables("SOTORDR3").Rows
                rowSOTORDR3.Item("SO_LOT_CASES") = rowSOTORDR3.Item("SO_LOT_CASES_SHIP")
            Next
        End If


        CUST_CODE = rowSOTORDR1.Item("CUST_CODE")
        CUST_BILL_TO_CUST = rowSOTORDR1.Item("CUST_BILL_TO_CUST")
        rowARTCUST1_SOLDTO = LookUp("ARTCUST1", CUST_CODE)

        'TAC.SOCMAIN1.Load_PAYEE_into_Contracts(Me, CUST_BILL_TO_CUST)
        Setup_SOTORDR3()
        Toggle_Transfer_Fields()
        Display_Totals()
        Load_Bill_Tos()
        Load_Ship_Tos()

        cmbCUST_BILL_TO_CUST.Value = rowSOTORDR1.Item("CUST_BILL_TO_CUST")
        Load_Bill_To_Address()

        cmbCUST_SHIP_TO_CODE.Value = rowSOTORDR1.Item("CUST_SHIP_TO_CODE")

        If BATCH_NO <> "" Then
            For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR1.Rows
                If grow.Cells("SO_ORDER_NO").Text = SO_ORDER_NO_init Then
                    grdSOTORDR1.ActiveRow = grow
                    Exit For
                End If
            Next
        End If
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub


    Public Overrides Function Remote_Control( _
 ByVal command As String, _
 Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Select Order"
                Absx1.txtFor("SO_ORDER_NO").Text = key
                Click_Command("Select Order")
        End Select

        Return return_key
    End Function


    Sub Update_Batch()

        Synch_TABLE_NAME("SOTORDR1")

        Dim CUST_BILL_TO_CUST As String = cmbCUST_BILL_TO_CUST.Value
        Dim SHIP_CODE As String = Absx1.txtFor("SHIP_CODE").Text
        Dim FRT_TERMS As String = Absx1.txtFor("FRT_TERMS").Text
        Dim FRT_RATE As Decimal = Val(Absx1.numFor("FRT_RATE").Value & "")
        Dim CUST_ROUTING_INST As String = Absx1.txtFor("CUST_ROUTING_INST").Text
        Dim ORDR_INV_DATE As String = Absx1.dteFor("ORDR_INV_DATE").Value
        Dim ORDR_DATE_SHIPPED As String = Absx1.dteFor("ORDR_DATE_SHIPPED").Value

        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Rows
            rowSOTORDR1.Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST
            rowSOTORDR1.Item("SHIP_CODE") = SHIP_CODE
            rowSOTORDR1.Item("FRT_TERMS") = FRT_TERMS
            rowSOTORDR1.Item("FRT_RATE") = FRT_RATE
            rowSOTORDR1.Item("CUST_ROUTING_INST") = CUST_ROUTING_INST
            rowSOTORDR1.Item("ORDR_INV_DATE") = ORDR_INV_DATE
            rowSOTORDR1.Item("ORDR_DATE_SHIPPED") = ORDR_DATE_SHIPPED
        Next

    End Sub

    Sub Update_Record()

        ' ANY CHANGES TO THIS METHOD SHOULD BE CHECKED INTO SOFINVH2.REVERSE_INVOICE

        BeginTrans()

        If BATCH_NO <> "" Then
            Update_Batch()
        End If

        ASCDATA1.DeleteRows("SOTORDRB", "ISNULL(SO_LOT_CASES,0) = 0")
        ASCDATA1.DeleteRows("SOTORDR3", "ISNULL(SO_LOT_CASES,0) = 0")
        ASCDATA1.DeleteRows("SOTORDR2", "ISNULL(SO_LOT_CASES,0) = 0")

        Reset_TP_Prices(True)

        For Each rowSOTORDRB As DataRow In dst.Tables("SOTORDRB").Select
            rowSOTORDRB.Item("CONT_TP_AMT") = Val(rowSOTORDRB.Item("RATE") & "") * Val(rowSOTORDRB.Item("SO_LOT_UNITS") & "")
            rowSOTORDRB.Item("CONT_TP_AMT_APPR") = Val(rowSOTORDRB.Item("RATE") & "") * Val(rowSOTORDRB.Item("SO_LOT_UNITS") & "")
            'rowSOTORDRB.Item("CONT_TP_AMT") = rowSOTORDRB.Item("CONT_TP_AMT_CALC")
            'rowSOTORDRB.Item("CONT_TP_AMT_APPR") = rowSOTORDRB.Item("CONT_TP_AMT_CALC")
        Next

        For Each rowSOTORDR1X As DataRow In dst.Tables("SOTORDR1").Rows

            Dim SO_ORDER_NO As String = rowSOTORDR1X.Item("SO_ORDER_NO")

            If rowSOTORDR1X.Item("ORDR_TYPE_CODE") = "T" Then
                If rowSOTORDR1X.Item("TRANSFER_TYPE") = "W" Then
                    rowSOTORDR1X.Item("SHIP_CODE") = "WORK"
                    rowSOTORDR1X.Item("FRT_RATE") = 0
                    If rowSOTORDR1X.Item("SVC_CHG_CODE") & "" <> "" And Val(rowSOTORDR1X.Item("SVC_CHG_RATE") & "") <> 0 Then
                        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("SO_ORDER_NO = '" & SO_ORDER_NO & "'")
                            rowSOTORDR2.Item("SVC_CHG_CODE") = rowSOTORDR1X.Item("SVC_CHG_CODE")
                            rowSOTORDR2.Item("SVC_CHG_RATE") = rowSOTORDR1X.Item("SVC_CHG_RATE")
                        Next
                    End If
                Else
                    rowSOTORDR1X.Item("VEND_CODE_SVC") = ""
                    rowSOTORDR1X.Item("SVC_CHG_CODE") = ""
                    rowSOTORDR1X.Item("SVC_CHG_RATE") = 0
                End If
            End If

            If BATCH_NO <> "" Then
                ASCMAIN1.Progress("Now Invoicing Order:", SO_ORDER_NO)
            End If

            'Dim TERM_CODE As String = rowSOTORDR1X.Item("TERM_CODE") & ""

            'TAC.SOCMAIN1.Calculate_CGS(Me, "I", SO_ORDER_NO, "SOTORDR3")

            Dim QTY_UNITS As Decimal = Val(rowSOTORDR1X.Item("QTY_UNITS") & "")

            Dim ORDR_WD_CHG As Decimal = 0

            If rowSOTORDR1X.Item("CUST_CHARGE_WD") & "" = "1" Then
                Dim WHSE_WD_CHG As Decimal = 0
                Dim WHSE_CODE As String = ""
                Dim rowICTWHSE1 As DataRow = Nothing
                Dim sql As String = "SO_ORDER_NO = '" & SO_ORDER_NO & "' and SO_LOT_CASES <> 0"

                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTORDR3").Select(sql), "WHSE_CODE", "SO_ORDER_LNO").Rows
                    If WHSE_CODE <> row.Item("WHSE_CODE") Then
                        WHSE_CODE = row.Item("WHSE_CODE")
                        rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)
                        If rowSOTORDR1X.Item("SHIP_CODE") & "" = "TIS" Then
                            WHSE_WD_CHG = Val(rowICTWHSE1.Item("WHSE_XFR_CHG") & "")
                        Else
                            WHSE_WD_CHG = Val(rowICTWHSE1.Item("WHSE_WD_CHG") & "")
                        End If
                    End If
                    ORDR_WD_CHG += WHSE_WD_CHG
                Next
            End If

            Dim ORDR_MISC_CHG As Decimal = Val(rowSOTORDR1X.Item("ORDR_MISC_CHG") & "")
            Dim ORDR_AMT_GROSS As Decimal = Val(rowSOTORDR1X.Item("ORDR_AMT_GROSS") & "")
            Dim ORDR_AMT_ALLOW As Decimal = Val(rowSOTORDR1X.Item("ORDR_AMT_ALLOW") & "")

            Dim ORDR_INV_NO As String = ASCMAIN1.Next_Control_No("SOTINVH1" & cbeDIVISION_CODE.Value)
            Mid$(ORDR_INV_NO, 1, 1) = cbeDIVISION_CODE.Value

            rowSOTORDR1X.Item("ORDR_INV_TYPE") = "I"
            rowSOTORDR1X.Item("ORDR_INV_NO") = ORDR_INV_NO
            'rowSOTORDR1.Item("ORDR_INV_DATE") = absx1.txtfor("ORDR_INV_DATE")).Text
            'rowSOTORDR1X.Item("ORDR_INV_DATE") = dteInvoiceDate.Value DID NOT WORK
            rowSOTORDR1X.Item("ORDR_AMT") = ORDR_AMT_GROSS - ORDR_AMT_ALLOW
            rowSOTORDR1X.Item("ORDR_WD_CHG") = ORDR_WD_CHG
            rowSOTORDR1X.Item("ORDR_TOTAL_AMT") = ORDR_AMT_GROSS - ORDR_AMT_ALLOW + ORDR_WD_CHG + ORDR_MISC_CHG
            rowSOTORDR1X.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowSOTORDR1X.Item("SO_STATUS_CODE") = "I"
            rowSOTORDR1X.Item("SO_DATE_INVOICED") = DATETIME_STAMP
            rowSOTORDR1X.Item("ORDR_INVOICED") = "1"

            rowSOTORDR1X.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowSOTORDR1X.Item("LAST_DATE") = DATETIME_STAMP

            If rowSOTORDR1X.Item("FRT_TERMS") & "" = "PPD" Then
                rowSOTORDR1X.Item("PPD_INV_IND") = "0"
            End If

            If rowSOTORDR1X.Item("ORDR_INV_REPLACING") & "" <> "" Then
                ASCMAIN1.sql = "Update SOTINVH1 set ORDR_INV_REPLACED_BY = '" & ORDR_INV_NO & "'" _
                & " where ORDR_INV_NO = '" & rowSOTORDR1X.Item("ORDR_INV_REPLACING") & "'"
                ASCDATA1.ExecuteSQL()
            End If

            rowSOTORDR1X.Item("FRT_AMT_ACCRUED") = Val(Absx1.numFor("FRT_RATE").Value & "") * QTY_UNITS
            Create_SOTINVHW_Accruals(SO_ORDER_NO)
        Next

        Update_Record_TDA("SOTINVHW")

        If BATCH_NO <> "" Then
            ASCMAIN1.Progress("Lot Integrity Check:", BATCH_NO)
            'This check DOES NOT handle orders that are part of a batch whos qty has been modified
            'In fact, we don't need to, Mary said we could prevent users from changing qty's on orders
            'that are part of a batch.

            Dim qtyAvailMsg As String = ""
            Dim n As Integer = 0

            ASCMAIN1.sql = "SELECT * FROM (SELECT WHSE_CODE, LOT_NO, LOT_SEQ_NO " & vbCrLf _
            & ", (QTY_ON_HAND-QTY_COMMITTED-QTY_COMM+QTY_COMM_ORIG) QTY_CHECK FROM" & vbCrLf _
            & "(SELECT ICTLOTD1.WHSE_CODE, ICTLOTD1.LOT_NO, ICTLOTD1.LOT_SEQ_NO, ICTLOTD1.QTY_ON_HAND, ICTLOTD1.QTY_COMMITTED" & vbCrLf _
            & ", SUM(SOTORDR3.SO_LOT_CASES) QTY_COMM, SUM(SOTORDR3.SO_LOT_CASES) QTY_COMM_ORIG" & vbCrLf _
            & " FROM SOTORDR3, ICTLOTD1 WHERE SOTORDR3.WHSE_CODE = ICTLOTD1.WHSE_CODE" & vbCrLf _
            & " AND SOTORDR3.LOT_NO = ICTLOTD1.LOT_NO" & vbCrLf _
            & " AND SOTORDR3.LOT_SEQ_NO = ICTLOTD1.LOT_SEQ_NO" & vbCrLf _
            & " AND SOTORDR3.SO_ORDER_NO IN (SELECT DISTINCT SO_ORDER_NO FROM SOTORDR1 WHERE BATCH_NO = '" & BATCH_NO & "')" & vbCrLf _
            & " GROUP BY ICTLOTD1.WHSE_CODE, ICTLOTD1.LOT_NO, ICTLOTD1.LOT_SEQ_NO, ICTLOTD1.QTY_ON_HAND, ICTLOTD1.QTY_COMMITTED))" & vbCrLf _
            & " WHERE QTY_CHECK < 0"
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim WHSE_CODE As String = row.Item("WHSE_CODE")
                Dim LOT_NO As String = row.Item("LOT_NO")
                Dim LOT_SEQ_NO As Int64 = row.Item("LOT_SEQ_NO")
                Dim QTY_AVAIL As Integer = row.Item("QTY_CHECK")
                n = n + 1
                qtyAvailMsg = qtyAvailMsg & "Whse: " & WHSE_CODE & " Lot No: " & LOT_NO & " Qty Available: " & QTY_AVAIL.ToString & vbCrLf
            Next
            If qtyAvailMsg <> "" Then
                If n > 10 Then
                    MsgBox("Too many to display. Please take a screenshot and email ABS Support.", MsgBoxStyle.OkOnly, "Orders in the batch created Lots with Negative Qty Available")
                Else
                    MsgBox(qtyAvailMsg, MsgBoxStyle.OkOnly, "Orders in the batch created Lots with Negative Qty Available")
                End If
            End If
        End If
        ASCMAIN1.Progress("Invoices Updating...", "")



        '*********************************
        'TAC.SOCMAIN1.Invoice_Update(1, Me)
        '*********************************



        For Each rowSOTORDR1X As DataRow In dst.Tables("SOTORDR1").Rows
            Dim SO_ORDER_NO As String = rowSOTORDR1X.Item("SO_ORDER_NO")
            Dim ORDR_INV_NO As String = rowSOTORDR1X.Item("ORDR_INV_NO")
            If rowSOTORDR1X.Item("ORDR_SOURCE") = "E" Then
                ASCMAIN1.Progress("Now Generating 810 for Invoice:", ORDR_INV_NO)
                TAC.EDCMAIN1.Generate_810(clsASCBASE1, ORDR_INV_NO)
            End If
        Next

        If BATCH_NO = "" Then
            For Each row As DataRow In ASCDATA1.SelectDistinct _
                ("SOTORDR3", New String() {"WHSE_CODE", "LOT_NO", "LOT_SEQ_NO"}).Rows

                Dim WHSE_CODE As String = row.Item("WHSE_CODE")
                Dim LOT_NO As String = row.Item("LOT_NO")
                Dim LOT_SEQ_NO As Int32 = Val(row.Item("LOT_SEQ_NO"))

                Dim rowICTLOTD1 As DataRow = LookUp("ICTLOTD1", New String() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})

                Dim QTY_AVAIL As Int64 = Val(rowICTLOTD1.Item("QTY_ON_HAND") & "") _
                                       - Val(rowICTLOTD1.Item("QTY_COMMITTED") & "")
                If QTY_AVAIL < 0 Then

                    MsgBox("Qty Available is now Negative", _
                           MsgBoxStyle.OkOnly, _
                           "Warning: Whse " & WHSE_CODE & " Lot No " & LOT_NO)
                End If
            Next

        End If

        ASCMAIN1.Progress("", "")
        CommitTrans("Update Complete")

    End Sub

    Sub Approve_for_Invoicing()
        For Each rowSOTORDR3 As DataRow In dst.Tables("SOTORDR3").Rows
            rowSOTORDR3.Item("SO_LOT_CASES_SHIP") = rowSOTORDR3.Item("SO_LOT_CASES")
            rowSOTORDR3.Item("SO_LOT_CASES") = rowSOTORDR3.Item("SO_LOT_CASES_ORIG")
        Next

        BeginTrans()
        Update_Record_TDA("SOTORDR1")
        Update_Record_TDA("SOTORDR3")
        Update_Approval_Fields(SO_ORDER_NO)
        CommitTrans("Update Complete")
    End Sub

    Sub Update_Approval_Fields(ByVal Approved_Order As String)
        ASCMAIN1.sql = "Update SOTORDR1 Set BILLING_APPR_DATE = '" & Format(DATETIME_STAMP, "dd-MMM-yy") & "'" _
        & " , BILLING_APPR_BY = '" & ASCMAIN1.USER_ID & "'" _
        & " Where SO_ORDER_NO = '" & Approved_Order & "'"
        ASCDATA1.ExecuteSQL()

        Mode_Settings(False)
    End Sub

    Sub Update_Approval_Batch(ByVal Batch_No As String)
        ASCMAIN1.sql = "Update SOTORDR1 Set BILLING_APPR_DATE = '" & Format(DATETIME_STAMP, "dd-MMM-yy") & "'" _
        & " , BILLING_APPR_BY = '" & ASCMAIN1.USER_ID & "'" _
        & " Where BATCH_NO = '" & Batch_No & "'"
        ASCDATA1.ExecuteSQL()

        Mode_Settings(False)
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special( _
    ByVal ctl As Control, _
    ByVal COLUMN_NAME As String, _
    Optional ByRef sql_where As String = "", _
    Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "SO_ORDER_NO"
                sql_where = "ORDR_REL = '1'"
                If cbeDIVISION_CODE.Value <> "" Then
                    sql_where &= " and ORDR_DIV_CODE = '" & cbeDIVISION_CODE.Value & "'"
                End If

            Case "BRKR_CODE"
                sql_where = "SREP_TYPE = 'B'"
            Case Else
        End Select

    End Sub


    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTORDR1"
            E.COLUMN_NAME = "SO_ORDER_NO"
            E.CODE_VALUE = Absx1.txtFor("SO_ORDER_NO").Text
            E.DESC_VALUE = "Sales Order"
            E.ATTACHMENT_NOTES = ""
        End If

        Return E
    End Function
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        If MENU_ITEM_OBJECT = "SOFINVHA" Then
            Load_Popup_Menu(grdSOTINVHX, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Customer Inquiry", "Approve", "Approve Batch")
        Else
            Load_Popup_Menu(grdSOTINVHX, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Customer Inquiry")
        End If
        Load_Popup_Menu(grdSOTORDR3, "S", "Show Filter")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
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

        Select Case e.SourceControl.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        End If

        Select Case e.SourceControl.Name
            Case "grdSOTINVHX"
                e.Tool.ToolbarsManager.Tools("Sales Order Inquiry").SharedProps.Visible = True
                If MENU_ITEM_OBJECT = "SOFINVHA" Then

                    If grd.ActiveRow.Cells("BATCH_NO").Text <> "" Then
                        e.Tool.ToolbarsManager.Tools("Approve Batch").SharedProps.Caption = "Approve Batch: " & grd.ActiveRow.Cells("BATCH_NO").Text
                    End If
                    e.Tool.ToolbarsManager.Tools("Approve Batch").SharedProps.Visible = (grd.ActiveRow.Cells("BATCH_NO").Text <> "")
                    e.Tool.ToolbarsManager.Tools("Approve").SharedProps.Caption = "Approve Order: " & grd.ActiveRow.Cells("SO_ORDER_NO").Text
                End If
        End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                Dim SO_ORDER_NO As String = grd.ActiveRow.Cells("SO_ORDER_NO").Text
                Context_Launch("Load", SO_ORDER_NO, e.Tool.Key, "SOFORDRI", "F", "SO")

            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
            Case "Approve"
                Update_Approval_Fields(grd.ActiveRow.Cells("SO_ORDER_NO").Text)
            Case "Approve Batch"
                Update_Approval_Batch(grd.ActiveRow.Cells("BATCH_NO").Text)
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(txtctl)

        Select Case COLUMN_NAME
            Case "SO_ORDER_NO"
                If Absx1.txtFor("SO_ORDER_NO").Text <> "" Then
                    Click_Command("Select Order")
                End If
        End Select
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)

        Select Case COLUMN_NAME

            Case "SO_ORDER_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Select Order", e)
                End If

        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "SHIP_CODE"
                If Not Me.IsLoading Then
                    ' IF THE FRT_TERMS TEXT BOX.TEXT= "" THEN .TEXT = "SOMETHING"
                End If

                If Absx1.txtFor("SHIP_CODE").Text <> "" Then
                    LookUp("SOTSVIA1", Absx1.txtFor("SHIP_CODE").Text)

                    If cdr IsNot Nothing Then
                        Absx1.txtFor("SHIP_VIA").Text = cdr.Item("SHIP_DESC")
                    End If
                End If

            Case "FRT_TERMS"
                If Absx1.txtFor("FRT_TERMS").Text = "TIS" Or Absx1.txtFor("FRT_TERMS").Text = "CPU" Then
                    Absx1.txtFor("SHIP_CODE").Text = Absx1.txtFor(COLUMN_NAME).Text
                End If

                If Absx1.txtFor("FRT_TERMS").Text <> "PPD" Then
                    Absx1.numFor("FRT_RATE").Value = 0
                    Set_Read_Only_for_ctl(Absx1.numFor("FRT_RATE"), True)
                Else
                    Set_Read_Only_for_ctl(Absx1.numFor("FRT_RATE"), False)
                End If

                If Not IsLoading Then
                    If use_CUST_PU_DATE And Absx1.txtFor("FRT_TERMS").Text <> "PPD" _
                    And Format(Absx1.dteFor("ORDR_DATE_SHIPPED").Value, "yyyyMMdd") _
                                         >= Format(Absx1.dteFor("CUST_PU_DATE").Value, "yyyyMMdd") Then
                        Absx1.dteFor("ORDR_INV_DATE").Value = Absx1.dteFor("CUST_PU_DATE").Value
                    Else
                        Absx1.dteFor("ORDR_INV_DATE").Value = Absx1.dteFor("ORDR_DATE_SHIPPED").Value
                    End If

                End If

        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "ITEM_CODE"
            '    If Absx1.txtFor("ITEM_CODE").Tag = "Y" Then
            '        Absx1.txtFor("ITEM_CODE").Tag = ""
            '        'Dim X As CurrencyManager = Me.BindingContext(dst.Tables("SOTORDR2"))
            '        'X.EndCurrentEdit()

            '        Click_Command("Find Lots")
            '    End If
        End Select
    End Sub

    Public Overrides Sub txt_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.txt_Leave(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

            Case "SHIP_CODE"
                If Absx1.txtFor(COLUMN_NAME).Text = "TIS" Or Absx1.txtFor(COLUMN_NAME).Text = "CPU" Then
                    Absx1.txtFor("FRT_TERMS").Text = Absx1.txtFor(COLUMN_NAME).Text
                End If
            Case "FRT_TERMS"
                If Absx1.txtFor(COLUMN_NAME).Text = "TIS" Or Absx1.txtFor(COLUMN_NAME).Text = "CPU" Then
                    Absx1.txtFor("SHIP_CODE").Text = Absx1.txtFor(COLUMN_NAME).Text
                End If

        End Select

    End Sub

    Public Overrides Sub opt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "TRANSFER_TYPE"
                ' Toggle_Transfer_Fields()

        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "FRT_RATE"
                If ScreenMode Then
                    Synch_TABLE_NAME("SOTORDR1")
                    Display_Totals()
                End If
            Case "SVC_CHG_RATE"
                If ScreenMode Then
                    Synch_TABLE_NAME("SOTORDR1")
                    Display_Totals()
                End If
        End Select

    End Sub

    Public Overrides Sub dte_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.dte_ValueChanged(sender, e)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "ORDR_DATE_SHIPPED"
                If Not IsLoading Then
                    If use_CUST_PU_DATE And Absx1.txtFor("FRT_TERMS").Text <> "PPD" _
                    And Format(Absx1.dteFor("ORDR_DATE_SHIPPED").Value, "yyyyMMdd") _
                                         >= Format(Absx1.dteFor("CUST_PU_DATE").Value, "yyyyMMdd") Then
                        Absx1.dteFor("ORDR_INV_DATE").Value = Absx1.dteFor("CUST_PU_DATE").Value
                    Else
                        Absx1.dteFor("ORDR_INV_DATE").Value = Absx1.dteFor("ORDR_DATE_SHIPPED").Value
                    End If

                End If

                If Absx1.dteFor("ORDR_DATE_SHIPPED").Value = Absx1.dteFor("ORDR_INV_DATE").Value Then
                    Absx1.dteFor("ORDR_INV_DATE").Appearance.ForeColor = Drawing.Color.Empty
                Else
                    Absx1.dteFor("ORDR_INV_DATE").Appearance.ForeColor = Drawing.Color.Red
                End If

            Case "ORDR_INV_DATE"

                If Absx1.dteFor("ORDR_DATE_SHIPPED").Value = Absx1.dteFor("ORDR_INV_DATE").Value Then
                    Absx1.dteFor("ORDR_INV_DATE").Appearance.ForeColor = Drawing.Color.Empty
                Else
                    Absx1.dteFor("ORDR_INV_DATE").Appearance.ForeColor = Drawing.Color.Red
                End If

        End Select

    End Sub
#End Region

    Private Sub Load_SOTINVHX()
        If SELECTION_NO = 0 Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")


        Select Case optShow.Value
            Case "N"
                ASCMAIN1.sql = "Select SOTINVH1.*" & vbCrLf _
                & " from SOTINVH1 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" _
                & " and ORDR_INV_TYPE = 'I' and ORDR_INV_REG IS NULL"

            Case "P"
                ASCMAIN1.sql = "Select SOTORDR1.*" & vbCrLf _
                & " from SOTORDR1 where ORDR_REL = '1' "

            Case "A"
                ASCMAIN1.sql = "Select SOTORDR1.*" & vbCrLf _
                & " from SOTORDR1 where ORDR_REL = '1' " _
                & " and BILLING_APPR_DATE is not Null"


            Case "D"
                ASCMAIN1.sql = "Select SOTINVH1.* from SOTINVH1" & vbCrLf _
                & " where ORDR_INV_DATE = '" & Format(dteShow.Value, "dd-MMM-yyyy") & "'"

            Case "J"
                ASCMAIN1.sql = "Select SOTINVH1.* from SOTINVH1" & vbCrLf _
                & " where ORDR_INV_REG_DATE = '" & Format(dteShow.Value, "dd-MMM-yyyy") & "'"
            Case "R"
                ASCMAIN1.sql = "Select SOTORDR1.*" & vbCrLf _
                & " from SOTORDR1 where ORDR_REL = '1' " _
                & " and BILLING_APPR_DATE is Null"
        End Select

        If chkMyInvoicesOnly.Checked Then
            ASCMAIN1.sql &= " and LAST_OPER = '" & ASCMAIN1.USER_ID & "'"

        End If

        Fill_Records("SOTINVHX", "", , ASCMAIN1.sql)

        Setup_grdSOTINVHX()
        Sort_grdColumns(grdSOTINVHX, "SO_ORDER_NO".ToLower)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_grdSOTINVHX()
        Dim dvw As DataView = DirectCast(grdSOTINVHX.DataSource, DataTable).DefaultView
        dvw.RowFilter = ""
        grdSOTINVHX.Text = Replace(optShow.Text, "...", Format(dteShow.Value, "MM/dd/yyyy"))
    End Sub

    Private Sub grdSOTINVHX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTINVHX.DoubleClickRow
        If optShow.Value = "P" Or optShow.Value = "R" Or optShow.Value = "A" Then
            Absx1.txtFor("SO_ORDER_NO").Text = e.Row.Cells("SO_ORDER_NO").Value
            Click_Command("Select Order")
        End If
    End Sub

    Function Get_CURR_EXCH_RATE() As Decimal
        Dim CURR_EXCH_RATE As Decimal = 0
        If CURR_CODE = "USD" Then
            CURR_EXCH_RATE = 1
        Else
            ASCMAIN1.sql = "Select CURR_EXCH_RATE from TATCURR3 " _
            & " where CURR_CODE = :PARM1" _
            & " and CURR_DATE <= :PARM2" _
            & " ORDER BY CURR_DATE DESC"
            Dim TATCURR3 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VD", _
                            New Object() {CURR_CODE, Absx1.dteFor("ORDR_INV_DATE").Value})
            If TATCURR3 IsNot Nothing Then
                CURR_EXCH_RATE = Val(TATCURR3.Item(0) & "")
            End If
        End If
        Return CURR_EXCH_RATE
    End Function

    Sub Create_SOTINVHW_Accruals(ByVal SO_ORDER_NO As String)

        Dim ORDR_DIV_CODE As String = cbeDIVISION_CODE.Value
        Dim WHSE_CODE As String = ""
        Dim COST_CATGY_CODE As String
        Dim rowPOTCATG1 As DataRow
        Dim rowSOTINVHW As DataRow

        ' Service Charges, to be billed by Preferred as Accessorial Misc Service Charge

        WHSE_CODE = ""
        Dim rowEDT945T1_BILL_OF_LADING_NO As DataRow = Nothing

        For Each rowSOTORDR3 As DataRow In dst.Tables("SOTORDR3").Select _
           ("SO_ORDER_NO = '" & SO_ORDER_NO & "'" & " and SO_LOT_UNITS <> 0 and PARENT(SOTORDR2_SOTORDR3).SVC_CHG_RATE <> 0", "WHSE_CODE")

            If WHSE_CODE <> rowSOTORDR3.Item("WHSE_CODE") Then
                WHSE_CODE = rowSOTORDR3.Item("WHSE_CODE")
                rowEDT945T1_BILL_OF_LADING_NO = Fill_Record("EDT945T1_BILL_OF_LADING_NO", New String() {SO_ORDER_NO, WHSE_CODE})
            End If




            Dim rowSOTORDR2 As DataRow = rowSOTORDR3.GetParentRow("SOTORDR2_SOTORDR3")
            Dim SO_LOT_UNITS As Decimal = Val(rowSOTORDR3.Item("SO_LOT_UNITS") & "")
            Dim SVC_CHG_CODE As String = rowSOTORDR2.Item("SVC_CHG_CODE")
            Dim SVC_CHG_RATE As Decimal = Val(rowSOTORDR2.Item("SVC_CHG_RATE") & "")
            Dim rowSOTSVCG1 As DataRow = dst.Tables("SOTSVCG1").Rows.Find(SVC_CHG_CODE)

            rowSOTINVHW = dst.Tables("SOTINVHW").NewRow
            With rowSOTINVHW
                .Item("CTL_NO") = ASCMAIN1.Next_Control_No("SOTINVHW.CTL_NO")
                .Item("SO_ORDER_NO") = SO_ORDER_NO
                .Item("SO_ORDER_LNO") = rowSOTORDR3.Item("SO_ORDER_LNO")
                .Item("SO_LOT_LNO") = rowSOTORDR3.Item("SO_LOT_LNO")
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("LOT_NO") = rowSOTORDR3.Item("LOT_NO")
                .Item("ORDR_INV_DATE") = rowSOTORDR1.Item("ORDR_INV_DATE")
                If rowEDT945T1_BILL_OF_LADING_NO IsNot Nothing Then
                    .Item("BILL_OF_LADING_NO") = rowEDT945T1_BILL_OF_LADING_NO.Item("BILL_OF_LADING_NO")
                End If
                .Item("WD_CHG_EXP") = SO_LOT_UNITS * SVC_CHG_RATE
                .Item("VEND_CODE") = rowSOTORDR1.Item("VEND_CODE_SVC")
                .Item("COST_CATGY_CODE") = rowSOTSVCG1.Item("COST_CATGY_CODE")
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("ORDR_DIV_CODE") = ORDR_DIV_CODE
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("ENTRY_TYPE") = "S"
            End With
            dst.Tables("SOTINVHW").Rows.Add(rowSOTINVHW)
        Next

        ' WD Charge Expense

        WHSE_CODE = ""
        Dim rowICTWHSE1_RATE As DataRow = Nothing
        Dim rowEDT945T1_REL_ORDER_NO As DataRow = Nothing

        For Each rowSOTORDR3 As DataRow In dst.Tables("SOTORDR3").Select _
            ("SO_ORDER_NO = '" & SO_ORDER_NO & "'", "WHSE_CODE")

            If WHSE_CODE <> rowSOTORDR3.Item("WHSE_CODE") Then
                WHSE_CODE = rowSOTORDR3.Item("WHSE_CODE")
                rowICTWHSE1_RATE = Fill_Record("ICTWHSE1_RATE", WHSE_CODE)
                rowEDT945T1_REL_ORDER_NO = Fill_Record("EDT945T1_REL_ORDER_NO", New String() {SO_ORDER_NO, WHSE_CODE})
            End If

            If Val(rowICTWHSE1_RATE.Item("WHSE_XFR_CHG_EXP") & "") <> 0 _
            Or Val(rowICTWHSE1_RATE.Item("WHSE_WD_CHG_EXP") & "") <> 0 Then

                rowSOTINVHW = dst.Tables("SOTINVHW").NewRow
                With rowSOTINVHW
                    .Item("CTL_NO") = ASCMAIN1.Next_Control_No("SOTINVHW.CTL_NO")
                    .Item("SO_ORDER_NO") = SO_ORDER_NO
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("LOT_NO") = rowSOTORDR3.Item("LOT_NO")
                    '.Item("ORDR_INV_DATE") = rowSOTORDR1.Item("ORDR_INV_DATE")
                    .Item("ORDR_INV_DATE") = dteDATE_SHIPPED.Value 'WORKS
                    If rowEDT945T1_REL_ORDER_NO IsNot Nothing Then
                        .Item("REL_ORDER_NO") = rowEDT945T1_REL_ORDER_NO.Item("REL_ORDER_NO")
                    End If
                    If ORDR_TYPE_CODE = "T" Then
                        .Item("WD_CHG_EXP") = rowICTWHSE1_RATE.Item("WHSE_XFR_CHG_EXP")
                    Else
                        .Item("WD_CHG_EXP") = rowICTWHSE1_RATE.Item("WHSE_WD_CHG_EXP")
                    End If
                    .Item("VEND_CODE") = rowICTWHSE1_RATE.Item("VEND_CODE")
                    .Item("COST_CATGY_CODE") = "WD"
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("ORDR_DIV_CODE") = ORDR_DIV_CODE
                    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    .Item("ENTRY_TYPE") = "S"
                End With
                dst.Tables("SOTINVHW").Rows.Add(rowSOTINVHW)
            End If
        Next

        ' COD Charges

        Dim TERM_CODE As String = rowSOTORDR1.Item("TERM_CODE")
        Dim rowTATTERM1 As DataRow = LookUp("TATTERM1", TERM_CODE)

        If rowTATTERM1.Item("TERM_TYPE") & "" = "C" Then
            COST_CATGY_CODE = "COD"
            rowPOTCATG1 = dst.Tables("POTCATG1").Rows.Find(COST_CATGY_CODE)
            Dim DEFAULT_CHARGE_COD As Decimal = Val(rowPOTCATG1.Item("DEFAULT_CHARGE") & "")

            rowSOTINVHW = dst.Tables("SOTINVHW").NewRow
            With rowSOTINVHW
                .Item("CTL_NO") = ASCMAIN1.Next_Control_No("SOTINVHW.CTL_NO")
                .Item("SO_ORDER_NO") = SO_ORDER_NO
                .Item("ORDR_INV_DATE") = rowSOTORDR1.Item("ORDR_INV_DATE")
                .Item("WD_CHG_EXP") = DEFAULT_CHARGE_COD
                .Item("VEND_CODE") = ""
                .Item("COST_CATGY_CODE") = COST_CATGY_CODE
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("ORDR_DIV_CODE") = ORDR_DIV_CODE
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("ENTRY_TYPE") = "S"
            End With
            dst.Tables("SOTINVHW").Rows.Add(rowSOTINVHW)

        End If

        ' Take Weights charges (Catch Weight)

        Dim CASES_CW As Int64 = Val(dst.Tables("SOTORDR3").Compute _
                ("SUM(SO_LOT_CASES)", _
                 "SO_ORDER_NO = '" & SO_ORDER_NO & "' and PACK_CODE = '" & TAC.TACMAIN1.CATCH_PACK & "'") & "")

        If CASES_CW <> 0 Then
            COST_CATGY_CODE = "TAKELB"
            rowPOTCATG1 = dst.Tables("POTCATG1").Rows.Find(COST_CATGY_CODE)
            Dim DEFAULT_CHARGE_CW As Decimal = Val(rowPOTCATG1.Item("DEFAULT_CHARGE") & "")

            rowSOTINVHW = dst.Tables("SOTINVHW").NewRow
            With rowSOTINVHW
                .Item("CTL_NO") = ASCMAIN1.Next_Control_No("SOTINVHW.CTL_NO")
                .Item("SO_ORDER_NO") = SO_ORDER_NO
                .Item("ORDR_INV_DATE") = rowSOTORDR1.Item("ORDR_INV_DATE")
                .Item("WD_CHG_EXP") = CASES_CW * DEFAULT_CHARGE_CW
                .Item("VEND_CODE") = ""
                .Item("COST_CATGY_CODE") = COST_CATGY_CODE
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("ORDR_DIV_CODE") = ORDR_DIV_CODE
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("ENTRY_TYPE") = "S"
            End With
            dst.Tables("SOTINVHW").Rows.Add(rowSOTINVHW)
        End If

    End Sub

    Sub Toggle_Transfer_Fields()
        Dim TRANSFER_TYPE As String = Absx1.optFor("TRANSFER_TYPE").Value

        splSW.Panel1Collapsed = (ORDR_TYPE_CODE = "T" And TRANSFER_TYPE = "W")
        splSW.Panel2Collapsed = Not (ORDR_TYPE_CODE = "T" And TRANSFER_TYPE = "W")
        grpBilling.Visible = (ORDR_TYPE_CODE <> "T")
        Absx1.optFor("TRANSFER_TYPE").Visible = (ORDR_TYPE_CODE = "T")
        Set_Read_Only_for_ctl(Absx1.optFor("TRANSFER_TYPE"), True)

        If ORDR_TYPE_CODE = "T" Then
            If TRANSFER_TYPE = "W" Then
                rowSOTORDR1.Item("SHIP_CODE") = "WORK"
                rowSOTORDR1.Item("FRT_TERMS") = "CPU"
                rowSOTORDR1.Item("FRT_RATE") = 0
            Else
                rowSOTORDR1.Item("VEND_CODE_SVC") = DBNull.Value
                rowSOTORDR1.Item("SVC_CHG_CODE") = DBNull.Value
                rowSOTORDR1.Item("SVC_CHG_RATE") = DBNull.Value
            End If

            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Rows
                rowSOTORDR2.Item("SVC_CHG_CODE") = Absx1.txtFor("SVC_CHG_CODE").Text
                rowSOTORDR2.Item("SVC_CHG_RATE") = Val(Absx1.numFor("SVC_CHG_RATE").Value & "")
            Next
        End If

    End Sub

    Sub Display_Totals()
        If SELECTION_NO = 0 Then Exit Sub

        ' NEED TO HANDLE CURR_CODE AND EXCH_RATE
        Dim GROSS As Decimal = 0
        Dim FRT As Decimal = 0
        Dim REBATES As Decimal = 0
        Dim FUNDS As Decimal = 0
        Dim ALLOWANCES As Decimal = 0
        Dim SVC_CHGS As Decimal = 0
        Dim BRKR_COMMS As Decimal = 0
        Dim CASES As Decimal = 0
        Dim UNITS As Decimal = 0

        Reset_TP_Prices()

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2") _
            .Select("SO_ORDER_NO = '" & SO_ORDER_NO & "' and SO_ORDER_LNO <> 0")
            Dim QTY_CASES As Int64 = Val(rowSOTORDR2.Item("QTY_CASES") & "")
            Dim QTY_UNITS As Decimal = Val(rowSOTORDR2.Item("QTY_UNITS") & "")
            Dim LINE_AMOUNT As Decimal = Val(rowSOTORDR2.Item("LINE_AMOUNT") & "")
            GROSS += LINE_AMOUNT
            FRT += QTY_UNITS * Val(Absx1.numFor("FRT_RATE").Value & "")
            REBATES += QTY_UNITS * Val(rowSOTORDR2.Item("REBATE") & "")
            FUNDS += QTY_UNITS * Val(rowSOTORDR2.Item("FUND_RATE") & "")
            ALLOWANCES += QTY_UNITS * Val(rowSOTORDR2.Item("ALLOW_RATE") & "")
            BRKR_COMMS += QTY_UNITS * Val(rowSOTORDR2.Item("BRKR_RATE") & "")
            CASES += QTY_CASES
            UNITS += QTY_UNITS
        Next

        SVC_CHGS += UNITS * Val(Absx1.numFor("SVC_CHG_RATE").Value & "")

        Dim ORDR_MISC_CHG As Decimal = Val(dst.Tables("SOTORDR5").Compute("SUM(SO_NINV_AMOUNT)", "") & "")
        Dim ORDR_WD_CHG As Decimal = Val(rowSOTORDR1.Item("ORDR_WD_CHG") & "")
        Dim ORDR_TOTAL_AMT As Decimal = GROSS - ALLOWANCES + ORDR_MISC_CHG + ORDR_WD_CHG

        Dim TOTAL_SEQ As Int32 = 0
        With dst.Tables("SOTORDRT")
            .Rows.Clear()
            TOTAL_SEQ += 1 : .Rows.Add(New Object() {TOTAL_SEQ, "Gross", GROSS})
            TOTAL_SEQ += 1 : .Rows.Add(New Object() {TOTAL_SEQ, "Allows", ALLOWANCES})
            TOTAL_SEQ += 1 : .Rows.Add(New Object() {TOTAL_SEQ, "W/D Chg", ORDR_WD_CHG})
            TOTAL_SEQ += 1 : .Rows.Add(New Object() {TOTAL_SEQ, "Misc", ORDR_MISC_CHG})
            TOTAL_SEQ += 1 : .Rows.Add(New Object() {TOTAL_SEQ, "Total", ORDR_TOTAL_AMT})

            TOTAL_SEQ += 1 : .Rows.Add(New Object() {TOTAL_SEQ, "Freight", FRT})
            TOTAL_SEQ += 1 : .Rows.Add(New Object() {TOTAL_SEQ, "Rebates", REBATES})
            TOTAL_SEQ += 1 : .Rows.Add(New Object() {TOTAL_SEQ, "Funds", FUNDS})
            TOTAL_SEQ += 1 : .Rows.Add(New Object() {TOTAL_SEQ, "Svc Chg", SVC_CHGS})
            TOTAL_SEQ += 1 : .Rows.Add(New Object() {TOTAL_SEQ, "Broker", BRKR_COMMS})
        End With
        Sort_grdColumns(grdSOTORDRT, "TOTAL_SEQ", True)

    End Sub

#Region "grdSOTORDR3"

    Private Sub grdSOTORDR3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTORDR3.AfterCellUpdate
        If e.Cell.Column.Key = "SO_LOT_CASES" Then
            If e.Cell.Row.Cells("PACK_CODE").Text <> TAC.TACMAIN1.CATCH_PACK Then
                e.Cell.Row.Cells("SO_LOT_UNITS").Value = _
                Val(e.Cell.Row.Cells("SO_LOT_CASES").Value & "") * _
                Val(e.Cell.Row.Cells("PACK_FACTOR").Value & "")
            End If
        End If
    End Sub

    Private Sub grdSOTORDR3_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTORDR3.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdSOTORDR3_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles grdSOTORDR3.BeforeCellActivate
        If BATCH_NO <> "" Then
            e.Cell.Column.CellActivation = UltraWinGrid.Activation.NoEdit
        End If
        If e.Cell.Column.Key = "SO_LOT_UNITS" Then
            If e.Cell.Row.Cells("PACK_CODE").Text <> TAC.TACMAIN1.CATCH_PACK Then
                e.Cell.Column.CellActivation = UltraWinGrid.Activation.NoEdit
            Else
                e.Cell.Column.CellActivation = UltraWinGrid.Activation.AllowEdit
            End If
        End If
    End Sub

    Private Sub grdSOTORDR3_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTORDR3.BeforeCellUpdate

        If e.Cell.Column.Key = "SO_LOT_UNITS" Then
            If e.Cell.Row.Cells("PACK_CODE").Text <> TAC.TACMAIN1.CATCH_PACK Then
                If grdSOTORDR3.ActiveCell IsNot Nothing AndAlso grdSOTORDR3.ActiveCell.Column.Key = "SO_LOT_CASES" Then
                Else
                    e.Cancel = True
                End If
            End If
        End If

    End Sub

    Private Sub grdSOTORDR3_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDR3.InitializeRow

        If e.Row.Cells("PACK_CODE").Text = TAC.TACMAIN1.CATCH_PACK Then
            e.Row.Cells("PACK_DESC").Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("PACK_DESC").Appearance.BackColor = Drawing.Color.Empty
            e.Row.Cells("SO_LOT_UNITS").Appearance.BackColor = Drawing.Color.Empty
        End If
        If Val(e.Row.Cells("SO_LOT_CASES").Text) <> Val(e.Row.Cells("SO_LOT_CASES_ORIG").Text) Then
            If e.Row.Cells("ACK").Text = "Checked" Then
                e.Row.Cells("ACK").Appearance.BackColor = Drawing.Color.Green
            Else
                e.Row.Cells("ACK").Appearance.BackColor = Drawing.Color.Red
            End If
        End If

    End Sub

    Private Sub grdSOTORDR3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles grdSOTORDR3.KeyPress
        If grdSOTORDR3.ActiveCell Is Nothing Then
            Exit Sub
        End If
    End Sub

#End Region

    Private Sub grdSOTORDRT_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRT.InitializeRow
        If e.Row.Cells("TOTAL_DESC").Value = "Total" Then
            e.Row.Appearance.FontData.Bold = DefaultableBoolean.True
            e.Row.Appearance.BackColor = Drawing.Color.LightGreen
        End If

        If e.Row.Cells("TOTAL_DESC").Value = "Freight" Then
            If Val(e.Row.Cells("TOTAL_AMT").Value & "") <> 0 Then
                e.Row.Cells("TOTAL_AMT").Appearance.BackColor = Drawing.Color.Yellow
            Else
                e.Row.Cells("TOTAL_AMT").Appearance.BackColor = Drawing.Color.Empty
            End If
        End If

        'If e.Row.Cells("TOTAL_DESC").Value = "Cases" Then
        '    e.Row.Cells("TOTAL_AMT"). = Format(Val(e.Row.Cells("TOTAL_AMT").Value), "##0")
        'End If
    End Sub

#Region "Bill-To"

    Sub Load_Bill_Tos()
        Fill_Records("ARTCUSTX", Absx1.txtFor("CUST_CODE").Text)
    End Sub

    Private Sub cmbCUST_BILL_TO_CUST_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCUST_BILL_TO_CUST.ValueChanged
        Load_Bill_To_Address()
    End Sub

    Sub Load_Bill_To_Address()
        If cmbCUST_BILL_TO_CUST.Value = "" Then
            txtCUST_BILL_TO_CUST_NAME.Text = ""
            txtCUST_BILL_TO_CUST_CSZ.Text = ""
        Else
            ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME" _
            & ", ARTCUST1.CUST_CITY || ', ' || ARTCUST1.CUST_STATE CUST_CSZ from ARTCUST1 " _
            & " where ARTCUST1.CUST_CODE = '" & cmbCUST_BILL_TO_CUST.Value & "'"
            Dim row As DataRow = ASCDATA1.GetDataRow
            txtCUST_BILL_TO_CUST_NAME.Text = row.Item("CUST_NAME")
            txtCUST_BILL_TO_CUST_CSZ.Text = row.Item("CUST_CSZ")
        End If
    End Sub
#End Region

#Region "Ship To"

    Sub Load_Ship_Tos()
        Dim rowCUST As DataRow = dst.Tables("SOTORDR1").Rows(0)

        Fill_Records("ARTCUSTS", Absx1.txtFor("CUST_CODE").Text)

        ASCMAIN1.sql = "Select 'SOLDTO' CUST_SHIP_TO_CODE, CUST_NAME CUST_SHIP_TO_NAME, CUST_ADDR1 CUST_SHIP_TO_ADDR1, " & vbCrLf _
        & " CUST_ADDR2 CUST_SHIP_TO_ADDR2, CUST_CITY CUST_SHIP_TO_CITY, CUST_STATE CUST_SHIP_TO_STATE, " & vbCrLf _
        & " CUST_ZIP_CODE CUST_SHIP_TO_ZIP_CODE, " & vbCrLf _
        & " CUST_COUNTRY CUST_SHIP_TO_COUNTRY, CUST_CONTACT CUST_SHIP_TO_CONTACT, CUST_PHONE " & vbCrLf _
        & " CUST_SHIP_TO_PHONE, CUST_ROUTING_INST " & vbCrLf _
        & " from ARTCUST1 where CUST_CODE = '" & rowCUST.Item("CUST_CODE") & "'"
        Dim rowARTCUST1 As DataRow = ASCDATA1.GetDataRow
        If rowARTCUST1 IsNot Nothing Then
            Dim rowARTCUSTS As DataRow = dst.Tables("ARTCUSTS").NewRow
            With rowARTCUSTS
                .Item("CUST_SHIP_TO_CODE") = rowARTCUST1.Item("CUST_SHIP_TO_CODE") & ""
                .Item("CUST_SHIP_TO_NAME") = rowARTCUST1.Item("CUST_SHIP_TO_NAME") & ""
                .Item("CUST_SHIP_TO_ADDR1") = rowARTCUST1.Item("CUST_SHIP_TO_ADDR1") & ""
                .Item("CUST_SHIP_TO_ADDR2") = rowARTCUST1.Item("CUST_SHIP_TO_ADDR2") & ""
                .Item("CUST_SHIP_TO_ADDR3") = String.Empty
                .Item("CUST_SHIP_TO_CITY") = rowARTCUST1.Item("CUST_SHIP_TO_CITY") & ""
                .Item("CUST_SHIP_TO_STATE") = rowARTCUST1.Item("CUST_SHIP_TO_STATE") & ""
                .Item("CUST_SHIP_TO_ZIP_CODE") = rowARTCUST1.Item("CUST_SHIP_TO_ZIP_CODE") & ""
                .Item("CUST_SHIP_TO_COUNTRY") = rowARTCUST1.Item("CUST_SHIP_TO_COUNTRY") & ""
                .Item("CUST_SHIP_TO_CONTACT") = rowARTCUST1.Item("CUST_SHIP_TO_CONTACT") & ""
                .Item("CUST_SHIP_TO_PHONE") = rowARTCUST1.Item("CUST_SHIP_TO_PHONE") & ""
                .Item("CUST_ROUTING_INST") = rowARTCUST1.Item("CUST_ROUTING_INST") & ""
            End With
            dst.Tables("ARTCUSTS").Rows.Add(rowARTCUSTS)
        End If
        dst.Tables("ARTCUSTS").AcceptChanges()
    End Sub

    Private Sub cmbCUST_SHIP_TO_CODE_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCUST_SHIP_TO_CODE.ValueChanged
        If Not ScreenMode Then
            Exit Sub
        End If

        Dim Ship_To As String = Absx1.cmbFor("CUST_SHIP_TO_CODE").Value & ""
        Dim row As DataRow = dst.Tables("ARTCUSTS").Rows.Find(New Object() {Ship_To})
        If row IsNot Nothing Then
            Absx1.txtFor("CUST_SHIP_TO_NAME").Text = row.Item("CUST_SHIP_TO_NAME") & ""
            Absx1.txtFor("CUST_SHIP_TO_ADDR1").Text = row.Item("CUST_SHIP_TO_ADDR1") & ""
            Absx1.txtFor("CUST_SHIP_TO_ADDR2").Text = row.Item("CUST_SHIP_TO_ADDR2") & ""
            Absx1.txtFor("CUST_SHIP_TO_CITY").Text = row.Item("CUST_SHIP_TO_CITY") & ""
            Absx1.txtFor("CUST_SHIP_TO_STATE").Text = row.Item("CUST_SHIP_TO_STATE") & ""
            Absx1.txtFor("CUST_SHIP_TO_ZIP_CODE").Text = row.Item("CUST_SHIP_TO_ZIP_CODE") & ""
            Absx1.txtFor("CUST_SHIP_TO_COUNTRY").Text = row.Item("CUST_SHIP_TO_COUNTRY") & ""
            Absx1.medFor("CUST_SHIP_TO_PHONE").Text = row.Item("CUST_SHIP_TO_PHONE") & ""
            Absx1.txtFor("CUST_SHIP_TO_CONTACT").Text = row.Item("CUST_SHIP_TO_CONTACT") & ""
            Absx1.txtFor("CUST_ROUTING_INST").Text = row.Item("CUST_ROUTING_INST") & ""
        End If

    End Sub
#End Region

    Private Sub grdSOTORDR1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSOTORDR1.AfterRowActivate
        If BATCH_NO <> "" Then
            SO_ORDER_NO = grdSOTORDR1.ActiveRow.Cells("SO_ORDER_NO").Value
        Else
            SO_ORDER_NO = Me.SO_ORDER_NO
        End If
        Setup_SOTORDR3()
        Display_Totals()
    End Sub

    Sub Setup_SOTORDR3()
        If BATCH_NO <> "" And grdSOTORDR1.ActiveRow Is Nothing Then
            grdSOTORDR3.Visible = False
        Else
            Dim SO_ORDER_NO As String
            If BATCH_NO <> "" Then
                SO_ORDER_NO = grdSOTORDR1.ActiveRow.Cells("SO_ORDER_NO").Value
            Else
                SO_ORDER_NO = Me.SO_ORDER_NO
            End If

            Dim dvw As DataView = DirectCast(grdSOTORDR3.DataSource, DataTable).DefaultView
            dvw.RowFilter = "SO_ORDER_NO = '" & SO_ORDER_NO & "'"
            grdSOTORDR3.Visible = True
            grdSOTORDR3.Text = "Lot Details for Order " & SO_ORDER_NO
        End If
    End Sub

    Private Sub optShow_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optShow.ValueChanged
        dteShow.Visible = (optShow.Value = "D" Or optShow.Value = "J")
        Load_SOTINVHX()
    End Sub

    Private Sub dteShow_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles dteShow.AfterCloseUp
        Load_SOTINVHX()
    End Sub

    Private Sub dteShow_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dteShow.KeyDown
        If e.KeyCode = Keys.Enter Then
            Load_SOTINVHX()
        End If
    End Sub

    Private Sub chkMyInvoicesOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMyInvoicesOnly.CheckedChanged
        Load_SOTINVHX()
    End Sub

    Sub Reset_TP_Prices(Optional ByVal All_Orders As Boolean = False)
        'Optional ByVal SO_ORDER_LNO As Integer = -1, 
        Dim sqlw As String = ""
        If Not All_Orders Then
            sqlw = "SO_ORDER_NO = '" & SO_ORDER_NO & "'"
        End If
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(sqlw)
            With rowSOTORDR2
                ' Eliminate Null values for numeric fields - make them 0's
                For Each COLUMN_NAME As String In New String() _
                    {"ORDR_PRICE_GRS", "ORDR_PRICE_GRS_CS", "ORDR_PRICE_GRS_CURR", "SVC_CHG_RATE"}
                    .Item(COLUMN_NAME) = Val(.Item(COLUMN_NAME) & "")
                Next

                .Item("QTY_CASES") = Val(.Item("SO_LOT_CASES") & "")
                .Item("QTY_UNITS") = Val(.Item("SO_LOT_UNITS") & "")
                .Item("REBATE") = Val(.Item("REBATE_B") & "")
                .Item("ALLOW_RATE") = Val(.Item("ALLOW_RATE_B") & "")
                .Item("FUND_RATE") = Val(.Item("FUND_RATE_B") & "")
                .Item("BRKR_RATE") = Val(.Item("BRKR_RATE_B") & "") ' broker must be done last
                .Item("ORDR_PRICE_NET") = Val(.Item("ORDR_PRICE_NET_CALC") & "")
                'If ORDR_TYPE_CODE = "T" AndAlso rowSOTORDR1.Item("TRANSFER_TYPE") = "W" Then
                '    .Item("ORDR_PRICE_NET") = Val(.Item("ORDR_PRICE_NET_CALC") & "") - Val(.Item("SVC_CHG_RATE") & "")
                'Else
                '    .Item("ORDR_PRICE_NET") = Val(.Item("ORDR_PRICE_NET_CALC") & "")
                'End If

            End With
        Next
    End Sub

    Private Sub grdSOTORDR1_BeforeRowActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTORDR1.BeforeRowActivate
        Update_Batch()
    End Sub

    Private Sub grdSOTINVHX_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTINVHX.InitializeLayout

    End Sub

    Private Sub grdSOTORDR3_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTORDR3.InitializeLayout

    End Sub
End Class