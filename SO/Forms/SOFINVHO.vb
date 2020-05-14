Imports System.Drawing
Imports System.Math

Public Class SOFINVHO

    Private rowARTCUST1 As DataRow = Nothing
    Private rowARTCUST2 As DataRow = Nothing
    Private rowSOTORDR1po As DataRow = Nothing
    Private rowSOTORDR1 As DataRow = Nothing
    Private rowICTSTYL1 As DataRow = Nothing
    Private rowARTCUST1_BT As DataRow = Nothing

    Private ORDR_NO As String = String.Empty
    Private CUST_CODE As String = String.Empty
    Private CUST_STORE_NO As String = String.Empty
    Private CUST_BILL_TO_CUST As String = String.Empty
    Private ORDR_CUST_PO As String = String.Empty
    Private SREP_CODE As String = String.Empty
    Private SREP2_CODE As String = String.Empty
    Private COLOR_CODEs As List(Of String) = New List(Of String)
    Private STYLE_CODE_last_entry As String = String.Empty
    Private ORDR_LNOs As New List(Of Int64)
    Private CUST_DC_NO As String = String.Empty

    Private ORDR_GROUP_NO As String = String.Empty
    Private ORDR_GROUP_NOs As List(Of String) = New List(Of String)
    Private XFR_INV_NOs As List(Of String) = New List(Of String)

    Private TOTAL_ORDR_AMT As Double = 0
    Private addingRecord As Boolean = False
    Private sqlSOTPICK1 As String = String.Empty
    Private sqlSOTPICK2 As String = String.Empty
    Private rowSOTMISC1 As DataRow = Nothing
    Private COMPANY_CODE As String = String.Empty

    Private GL_PARM_CURR_CODE As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            COMPANY_CODE = "VAN"
        ElseIf ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
            COMPANY_CODE = "NYA"
        ElseIf ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            COMPANY_CODE = "RGI"
        Else
            COMPANY_CODE = ASCMAIN1.CLIENT
        End If

        ' Used in external class procedures
        Get_PARM("ICTPARM1")
        Get_PARM("WHTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("SOTPARM1")
        Get_PARM("GLTPARM1")

        GL_PARM_CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & String.Empty

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -6, 0, -1)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -6, 0, -1)

        With dst

            Create_TDA(.Tables.Add, "SOTORDR0", "*")
            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            Create_TDA(.Tables.Add, "SOTORDR7", "*")

            ASCMAIN1.sql = "Select SOTORDR2.*, ICTCOLR1.COLOR_DESC, ICTSTYL1.CASE_CUBE from SOTORDR2,ICTCOLR1,ICTSTYL1" _
                & " where SOTORDR2.ORDR_NO = :PARM1" _
                & "   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE" _
                & "   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, True, "V", 2)
            .Tables("SOTORDR2").Columns.Add("TOTAL_CARTONS", GetType(System.Decimal), "IIF(ISNULL(CARTON_PACK_QTY,0)=0,0,ISNULL(ORDR_QTY,0) / ISNULL(CARTON_PACK_QTY,0))")
            .Tables("SOTORDR2").Columns.Add("TOTAL_CUBE", GetType(System.Decimal), "ISNULL(TOTAL_CARTONS,0) * ISNULL(CASE_CUBE,0)")

            With .Tables("SOTORDR2").Columns
                .Add("RANGE_STYLE_QTY_PER_PP", GetType(System.Int64))
                .Add("ORDR_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_OPEN", GetType(System.Decimal), "ISNULL(ORDR_QTY_OPEN,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_ALLO", GetType(System.Decimal), "ISNULL(ORDR_QTY_OPEN,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_PICK", GetType(System.Decimal), "ISNULL(ORDR_QTY_PICK,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_SHIP", GetType(System.Decimal), "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_CANC", GetType(System.Decimal), "ISNULL(ORDR_QTY_CANC,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_COST", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(STYLE_PRICE,0)")
            End With
            .Tables("SOTORDR2").Columns("ORDR_UNIT_PRICE_MANUAL").DefaultValue = "0"

            Create_TDA(.Tables.Add, "SOTORDR3", "*")
            Create_TDA(.Tables.Add, "SOTORDR4", "*")
            Create_TDA(.Tables.Add, "SOTORDR5", "*")
            Create_TDA(.Tables.Add, "SOTORDR9", "*")

            'Create_TDA(.Tables.Add, "SOTPICK1", "*")
            sqlSOTPICK1 = "Select SOTPICK1.*" & vbCrLf _
                 & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_SOURCE" & vbCrLf _
                 & ", SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.CUST_BILL_TO_CUST" & vbCrLf _
                 & ", SOTORDR1.POST_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_FOB" & vbCrLf _
                 & ", SOTORDR1.TERM_CODE, SOTORDR1.SREP_CODE, SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
                 & ", SOTSHIP1.BILL_OF_LADING_NO, SOTORDR1.ORDR_INV_COMMENT, SOTORDR1.CUST_FACTOR_IND, SOTORDR1.CCPA_NO CCPA_NO_ORDR" & vbCrLf _
                 & " from SOTPICK1,SOTORDR1,SOTSHIP1 "
            ASCMAIN1.sql = sqlSOTPICK1 & " where ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**")
            dst.Tables("SOTPICK1").Columns.Add("SELECTED")
            dst.Tables("SOTPICK1").Columns.Add("OUR_FREIGHT", GetType(System.Decimal))
            dst.Tables("SOTPICK1").Columns("OUR_FREIGHT").DefaultValue = 0
            dst.Tables("SOTPICK1").Columns.Add("INV_MISC_CHG", GetType(System.Decimal))
            dst.Tables("SOTPICK1").Columns("INV_MISC_CHG").DefaultValue = 0


            '& ", SOTORDR5.CUST_NAME, SOTORDR5.CUST_ADDR1, SOTORDR5.CUST_ADDR2, SOTORDR5.CUST_ADDR3" & vbCrLf _
            '& ", SOTORDR5.CUST_CITY, SOTORDR5.CUST_STATE, SOTORDR5.CUST_ZIP_CODE, SOTORDR5.CUST_COUNTRY" & vbCrLf _
            '& ", SOTORDR5.CUST_CONTACT, SOTORDR5.CUST_PHONE, SOTORDR1.CCPA_NO CCPA_NO_ORDR" & vbCrLf _

            For Each fieldname As String In New String() {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3", _
                                                          "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY", _
                                                          "CUST_CONTACT", "CUST_PHONE"}
                dst.Tables("SOTPICK1").Columns.Add(fieldname, GetType(System.String))
            Next

            ' Create_TDA(.Tables.Add, "SOTPICK2", "*")
            sqlSOTPICK2 = "Select SOTPICK2.*, " & vbCrLf _
                & " SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_DESC, " & vbCrLf _
                & " SOTORDR2.ORDR_UNIT_PRICE, SOTORDR2.STYLE_CODE_SUB," & vbCrLf _
                & " SOTORDR2.RANGE_STYLE_CODE, SOTORDR2.RANGE_STYLE_LNO, SOTORDR2.QTY_PER_PP" & vbCrLf

            If COMPANY_CODE = "RGI" OrElse COMPANY_CODE = "NYA" Then
                sqlSOTPICK2 &= ", SOTORDR2.ORDR_PRICE_SOURCE, SOTORDR2.COMM_RATE " & vbCrLf
            End If

            sqlSOTPICK2 &= " from SOTPICK2,SOTPICK1,SOTORDR2,SOTSHIP1" & vbCrLf

            ASCMAIN1.sql = sqlSOTPICK2 & " where ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "SOTPICK2", "**")
            With .Tables("SOTPICK2").Columns
                .Add("PICK_AMT", GetType(System.Decimal))
                .Add("PICK_AMT_CONF", GetType(System.Decimal))
                .Add("PICK_AMT_CANC", GetType(System.Decimal))
                .Add("PICK_AMT_BACK", GetType(System.Decimal))
            End With

            Create_TDA(.Tables.Add, "SOTCART1", "*")
            Create_TDA(.Tables.Add, "SOTCART2", "*")

            Create_TDA(.Tables.Add, "SOTSHIP1", "*")
            Create_TDA(.Tables.Add, "SOTSHIP2", "*")
            Create_TDA(.Tables.Add, "SOTSHIP3", "*")
            Create_TDA(.Tables.Add, "SOTSHIP4", "*")
            Create_TDA(.Tables.Add, "SOTSHIP5", "*")
            Create_TDA(.Tables.Add, "SOTSHIP6", "*")

            Create_TDA(.Tables.Add, "ARTOPEN1", "*")
            Create_TDA(.Tables.Add, "SOTINVH1", "*")
            Create_TDA(.Tables.Add, "SOTINVH2", "*")
            Create_TDA(.Tables.Add, "SOTINVHM", "*")
            Create_TDA(.Tables.Add("SOTINVH2_XFR"), "SOTINVH2", "*", "1", False)
            Create_TDA(.Tables.Add("SOTORDR2_XFR"), "SOTORDR2", "*", "1", False)

            Create_TDA(.Tables.Add, "TATEVNT1", "*")
            Create_TDA(.Tables.Add, "ARTCUST1", "*")
            Create_TDA(.Tables.Add, "ARTCUST2", "*")

            ASCMAIN1.sql = "SELECT * FROM SOTINVH1 WHERE ORDR_TYPE_CODE = 'XFR' AND INV_TYPE = 'I' AND ORDR_YYYYPP_UPDATED BETWEEN :PARM1 AND :PARM2"
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "VV", 0)
            .Tables("SOTINVHX").Columns.Add("SEL", GetType(System.String))
            .Tables("SOTINVHX").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = "Select * from ICTCOLR1"
            Create_TDA(.Tables.Add, "ICTCOLR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC" _
                & " from ICTSTYC1,ICTCOLR1 where ICTSTYC1.STYLE_CODE = :PARM1" _
                & "  and ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTCOLRS", "**", 0, False, "V", 1)

            Create_TDA(.Tables.Add, "SOTRSRV1", "*")
            Create_TDA(.Tables.Add, "SOTRSRV2", "*")

            ASCMAIN1.sql = "Select SOTRSRV2.* from SOTRSRV2,SOTRSRV1" & vbCrLf _
                & " where SOTRSRV1.CUST_CODE = :PARM1 " & vbCrLf _
                & "   and SOTRSRV2.STYLE_CODE = :PARM2 " & vbCrLf _
                & "   and SOTRSRV2.COLOR_CODE = :PARM3" & vbCrLf _
                & "   and SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
                & "   and SOTRSRV1.RSRV_STATUS = 'O'" & vbCrLf _
                & "   and SOTRSRV2.RSRV_QTY_OPEN > 0" & vbCrLf
            Create_TDA(.Tables.Add, "SOTRSRVX", "**", 0, False, "VVV", 0)


        End With

        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")
        grdSOTINVH1X.DataSource = dst.Tables("SOTINVHX")

        grdSOTINVH1X.DisplayLayout.UseFixedHeaders = True
        With grdSOTINVH1X.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SEL", "INV_TYPE", "INV_NO", "CUST_CODE", "CUST_STORE_NO", "WHSE_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSOTORDR2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_LNO", "STYLE_CODE_SUB", "STYLE_CODE", "STYLE_DESC", "RANGE_STYLE_CODE", "COLOR_CODE", "COLOR_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With


        With grdSOTORDR2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If Not New String() {"STYLE_CODE", "COLOR_CODE", "ORDR_QTY", "ORDR_UNIT_PRICE", "ORDR_UNIT_PRICE_MANUAL", "RANGE_STYLE_CODE", "CUST_SKU"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"CUST_UPC", "CUST_SKU", "CUST_STYLE_CODE", "CUST_COLOR_CODE", "CUST_SIZE_CODE", "STYLE_RETAIL"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightPink
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() _
                {"ORDR_LNO", "STYLE_CODE", "COLOR_CODE", "STYLE_DESC", "COLOR_DESC", "RANGE_STYLE_CODE", "STYLE_CODE_SUB", _
                 "INNER_PACK_QTY", "CARTON_PACK_QTY", "TOTAL_CARTONS", "CASE_CUBE", "TOTAL_CUBE", "STYLE_UOM", "STYLE_CLASS_CODE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"STYLE_PRICE", "ORDR_UNIT_PRICE", "ORDR_UNIT_PRICE_CALC", "ORDR_UNIT_PRICE_MANUAL", "ORDR_PRICE_SOURCE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Gold
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"PO_COST"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_LNO", "STYLE_DESC", "COLOR_DESC"}
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            Next
        End With

        With grdSOTORDR2.DisplayLayout.Bands(0)

            If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            Else
                .Columns("STYLE_UOM").Hidden = True
                .Columns("STYLE_CLASS_CODE").Hidden = True
                .Columns("STYLE_PRICE").Hidden = True
                .Columns("ORDR_UNIT_PRICE_CALC").Hidden = True
                .Columns("ORDR_UNIT_PRICE_MANUAL").Hidden = True
            End If
        End With

        grdSOTORDR2.DisplayLayout.Bands(0).Columns("PO_COST").Hidden = True

        Create_Summary(grdSOTINVH1X, "INV_NO", "Count")

        Create_Summary(grdSOTORDR2, "ORDR_LNO", "Count")
        Create_Summary(grdSOTORDR2, New String() {"ORDR_QTY", _
                                                  "ORDR_QTY_OPEN", "ORDR_QTY_ALLO", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", _
                                                  "ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC", _
                                                  "TOTAL_CARTONS", "TOTAL_CUBE"})


        With grdSOTINVH1X.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With

        For Each gridCol As Infragistics.Win.UltraWinGrid.UltraGridColumn In grdSOTINVH1X.DisplayLayout.Bands(0).Columns
            If gridCol.Key = "SEL" Then
                gridCol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gridCol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        grdSOTINVHM.DataSource = dst.Tables("SOTINVHM")
        Create_Summary(grdSOTINVHM, "INV_MNO", "Count")
        Create_Summary(grdSOTINVHM, New String() {"INV_MISC_CHG"})
        grdSOTINVHM.DisplayLayout.Bands(0).Columns("INV_TYPE").Hidden = True
        grdSOTINVHM.DisplayLayout.Bands(0).Columns("INV_NO").Hidden = True

        Show_Filter(grdSOTINVH1X, True)
        grdSOTINVH1X.DisplayLayout.GroupByBox.Hidden = False

        ' This no longer used
        chkNoInventory.Checked = False
        chkNoInventory.Visible = False

        splDetails.Panel2Collapsed = True
        MyBase.Absx1.txtFor("ORDR_INV_COMMENT").MaxLength = dst.Tables("SOTORDR1").Columns("ORDR_INV_COMMENT").MaxLength

        TABLE_NAME = "SOTORDR1"
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 IsNot Nothing Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                        CUST_BILL_TO_CUST = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
                        If CUST_BILL_TO_CUST = "" Then
                            CUST_BILL_TO_CUST = CUST_CODE
                        End If

                        If rowARTCUST1.Item("CURR_CODE") <> GL_PARM_CURR_CODE Then
                            EMsg &= vbCr & "This Customer's currency code does not match the defualt currency code " & GL_PARM_CURR_CODE
                        End If

                        ' Customer must have a Sales Rep assigned
                        SREP_CODE = rowARTCUST1.Item("SREP_CODE") & ""
                        Dim rowSOTSREP1 As DataRow = Nothing
                        If SREP_CODE <> "" Then rowSOTSREP1 = LookUp("SOTSREP1", SREP_CODE)
                        If rowSOTSREP1 Is Nothing Then
                            EMsg &= vbCr & "This Customer Has No Sales Rep Assigned"
                        End If
                        SREP2_CODE = rowARTCUST1.Item("SREP2_CODE") & ""

                        ' apostrophe in Cust PO causes ABSolution to crash when lookig to see if it is a duplicate PO entry
                        Absx1.txtFor("ORDR_CUST_PO").Text = Absx1.txtFor("ORDR_CUST_PO").Text.Trim.Replace("'", "")
                        ORDR_CUST_PO = Absx1.txtFor("ORDR_CUST_PO").Text
                        If ORDR_CUST_PO = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                            EMsg &= vbCr & "You Must Provide a Value for Customer PO"
                        End If
                    Else
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                    End If
                End If

                If Absx1.txtFor("CUST_STORE_NO").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Store (Mark-For)"
                Else
                    rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", Absx1.txtFor("CUST_STORE_NO").Text})
                    If rowARTCUST2 IsNot Nothing Then
                        CUST_STORE_NO = Absx1.txtFor("CUST_STORE_NO").Text
                    Else
                        EMsg &= vbCr & "No Record of Customer Store " & Absx1.txtFor("CUST_STORE_NO").Text
                    End If
                End If

                If EMsg = "" Then
                    ' Load Default values in for Selected Fields if we have seen this Customer PO before
                    ASCMAIN1.sql = "Select ORDR_GROUP_NO, CUST_STORE_NO, ORDR_SHIP_DATE, ORDR_CANCEL_DATE" & vbCrLf _
                        & ", ORDR_DATE, ORDR_DEPT, ORDR_SHIP_INSTR, FRT_TERMS, SALES_DIVISION_CODE" & vbCrLf _
                        & " from SOTORDR1 where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                        & " and ORDR_CUST_PO = '" & ORDR_CUST_PO & "'" & vbCrLf _
                        & " order by ORDR_SHIP_DATE DESC"
                    ASCMAIN1.sql = "Select * from (" & ASCMAIN1.sql & ") where ROWNUM < 2"
                    rowSOTORDR1po = ASCDATA1.GetDataRow

                    If rowSOTORDR1po IsNot Nothing Then
                        ASCMAIN1.sql = "Select ORDR_NO, ORDR_DATE from SOTORDR1 " & vbCrLf _
                            & " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                            & "   and CUST_STORE_NO = '" & CUST_STORE_NO & "'" & vbCrLf _
                            & "   and ORDR_CUST_PO = '" & ORDR_CUST_PO & "'" & vbCrLf _
                            & "   and ORDR_STATUS in ('O','P','F')"
                        Dim rowDup As DataRow = ASCDATA1.GetDataRow
                        If rowDup IsNot Nothing Then
                            If MsgBox("Same Customer PO has already been entered for Store " & CUST_STORE_NO _
                                      & vbCrLf & " (See Sales Order " & rowDup.Item("ORDR_NO") & " dated " & Format(rowDup.Item("ORDR_DATE"), "MM/dd/yyyy") & ")" _
                                      & vbCrLf & vbCrLf & "Are You Sure that you want to Proceed?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, _
                                      "Possible Order Duplication") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If

                XFR_INV_NOs.Clear()

                If dst.Tables("SOTINVHX").Select("SEL = '1'").Length > 0 Then
                    Dim WHSE_CODE_TO As String = dst.Tables("SOTINVHX").Select("SEL = '1'")(0).Item("WHSE_CODE_TO") & String.Empty
                    WHSE_CODE_TO = WHSE_CODE_TO.Trim
                    If WHSE_CODE_TO.Length = 0 Then
                        EMsg &= vbCr & "The select invoice does not have a warehouse value."
                        Exit Select
                    End If

                    If EMsg.Length = 0 Then
                        For Each rowSOTINVHX As DataRow In dst.Tables("SOTINVHX").Select("SEL = '1'")
                            XFR_INV_NOs.Add(rowSOTINVHX.Item("INV_NO") & String.Empty)
                            If rowSOTINVHX.Item("CUST_CODE") & String.Empty <> CUST_CODE Then
                                EMsg &= vbCr & "At least one of the selected Invoices is not for customer (" & CUST_CODE & ")"
                                Exit Select
                            End If
                            If rowSOTINVHX.Item("WHSE_CODE_TO") & String.Empty <> WHSE_CODE_TO Then
                                EMsg &= vbCr & "At least one of the selected Invoices is not for Warehouse (" & WHSE_CODE_TO & ")"
                                Exit Select
                            End If
                        Next
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SOTRSRV1", CUST_CODE) Then Exit Sub
                End If

            Case "Update"
                Dim TERM_TYPE As String = String.Empty
                Dim FRT_TERMS As String = Absx1.txtFor("FRT_TERMS").Text
                Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text

                If FRT_TERMS = "" Then
                    EMsg &= vbCr & "Freight Terms are Mandatory"
                Else
                    Dim row As DataRow = LookUp("ASTCODE1", New String() {"SOTORDR1", "FRT_TERMS", FRT_TERMS})
                    If row Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Freight Terms"
                    ElseIf COMPANY_CODE = "RGI" Then
                        If WHSE_CODE = "FE" And FRT_TERMS <> "COL" Then EMsg &= vbCr & "Frt Terms should be COL for FE Orders"
                        If WHSE_CODE = "FD" And FRT_TERMS <> "PPD" Then EMsg &= vbCr & "Frt Terms should be COL for FD Orders"
                        If WHSE_CODE = "SP" And FRT_TERMS <> "PPD" Then EMsg &= vbCr & "Frt Terms should be COL for SP Orders"
                        If WHSE_CODE = "FA" And FRT_TERMS <> "PPA" Then EMsg &= vbCr & "Frt Terms should be PPA for FA Orders"
                    End If

                    Select Case Absx1.txtFor("FRT_TERMS").Text
                        Case "COL", "PPD"
                            If Absx1.numFor("INV_FREIGHT").Value > 0 Then
                                EMsg &= vbCr & "Frt Terms Code (" & Absx1.txtFor("FRT_TERMS").Text & ") does not permit freight."
                            End If
                        Case "PPA"
                            If Absx1.numFor("INV_FREIGHT").Value = 0 Then
                                EMsg &= vbCr & "Frt Terms Code (" & Absx1.txtFor("FRT_TERMS").Text & ") requires freight."
                            End If
                    End Select

                End If

                ' WHAT IS THIS SECTION DOING?
                Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")

                Dim ORDR_AMT As Decimal = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & "")
                If ORDR_AMT < 0 Then
                    EMsg &= vbCr & "Order Amount may not be less than $0.00"
                    Exit Select
                End If

                If Absx1.txtFor("TERM_CODE").Text = "" Then
                    EMsg &= vbCr & "Terms Code is required"
                Else
                    Validate_Code("TERM_CODE")
                    If cdr IsNot Nothing Then
                        TERM_TYPE = cdr.Item("TERM_TYPE") & String.Empty
                    End If
                End If

                If Absx1.chkFor("CUST_FACTOR_IND").Checked Then
                    If TERM_TYPE = "C" Then EMsg &= vbCr & "Cannot Factor with Terms Code " & Absx1.txtFor("TERM_CODE").Text
                    If ORDR_AMT = 0 Then EMsg &= vbCr & "Cannot Factor with $0 Order"
                End If

                If Absx1.txtFor("SREP_CODE").Text = "" Then
                    EMsg &= vbCr & "Sales Rep is required"
                Else
                    Validate_Code("SREP_CODE")
                End If

                Validate_Code("SREP2_CODE", False, True)
                Validate_Code("WHSE_CODE")
                Validate_Code("SHIP_VIA_CODE")

                If Absx1.txtFor("ORDR_CUST_PO").Text = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                    EMsg &= vbCr & "Customer PO is required"
                End If

                If grdSOTORDR2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Items on Order"
                Else
                    If Val(dst.Tables("SOTORDR2").Compute("COUNT(ORDR_LNO)", "ORDR_QTY > 0") & "") = 0 Then
                        EMsg &= vbCr & "No Items on Order with Qty > 0"
                    End If

                    For Each row As DataRow In ASCDATA1.SelectDistinct _
                        (dst.Tables("SOTORDR2").Select("RANGE_STYLE_CODE is Not Null"), "RANGE_STYLE_CODE").Rows
                        Dim RANGE_STYLE_CODE As String = row.Item("RANGE_STYLE_CODE")
                        If Val(dst.Tables("SOTORDR2").Select("RANGE_STYLE_CODE = '" & RANGE_STYLE_CODE & "'").Length) > 1 Then
                            EMsg &= vbCr & "Range Style " & RANGE_STYLE_CODE & " occurs on this Order More than Once"
                        End If
                    Next

                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select _
                        ("ISNULL(RANGE_STYLE_CODE,'') <> '' and RANGE_STYLE_QTY_PER_PP = 1")
                        Dim RANGE_STYLE_CODE As String = rowSOTORDR2.Item("RANGE_STYLE_CODE") & ""
                        Dim ORDR_LNO As Int32 = Val(rowSOTORDR2.Item("ORDR_LNO"))
                        Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO)
                        Dim LINES As Int64 = Val(dst.Tables("SOTORDRR").Compute("COUNT(ORDR_LNO)", sqlw) & "")
                        '   Dim TOTAL As Decimal = Val(dst.Tables("SOTORDRR").Compute("SUM(ORDR_AMT)", sqlw) & "")
                        Dim ORDR_QTY As Int64 = Val(dst.Tables("SOTORDRR").Compute("SUM(ORDR_QTY)", sqlw) & "")
                        If Val(rowSOTORDR2.Item("ORDR_QTY") & "") <> ORDR_QTY _
                        Or LINES = 0 Then
                            EMsg &= vbCr & "Line " & CStr(ORDR_LNO) & ": Range Style Price Out of Balance w/Components"
                        End If
                    Next

                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                        Dim rowSOTORDR3s() As DataRow = rowSOTORDR2.GetChildRows("SOTORDR2_SOTORDR3")
                        If rowSOTORDR3s.Length > 0 Then
                            Dim ORDR_QTY_2 As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY") & "")

                            Dim ORDR_LNO As Int64 = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
                            Dim unique_QTYs_by_Store As New Dictionary(Of Int64, String)
                            For Each row As DataRow In dst.Tables("SOTORDRS").Select("")
                                Dim QTY As Int64 = Val(row.Item("QTY_" & Format(ORDR_LNO, "000")) & "")
                                If Not unique_QTYs_by_Store.ContainsKey(QTY) Then
                                    Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
                                    unique_QTYs_by_Store.Add(QTY, CUST_STORE_NO)
                                End If
                            Next

                            Dim ORDR_QTY_3 As Int64 = 0

                            For Each rowSOTORDR3 As DataRow In rowSOTORDR3s
                                Dim ORDR_QTY As Int64 = Val(rowSOTORDR3.Item("ORDR_QTY") & "")
                                ORDR_QTY_3 += ORDR_QTY

                                For Each QTY As Int64 In unique_QTYs_by_Store.Keys
                                    If QTY Mod ORDR_QTY <> 0 Then
                                        EMsg &= vbCr & "Rounding Problem with Sub-Details on Line " _
                                            & CStr(ORDR_LNO) & " with Store " & unique_QTYs_by_Store(QTY)
                                        Exit For
                                    End If
                                Next
                                If rowSOTORDR2.Item("SIZE_DESC_01") & "" <> "" Then
                                    Dim TOTAL_SIZE_QTY As Int64 = Val(rowSOTORDR3.Item("TOTAL_SIZE_QTY") & "")
                                    If TOTAL_SIZE_QTY <> ORDR_QTY Then
                                        EMsg = EMsg & vbCr & "Size Distribution out of Balance with a Component on Line " & CStr(ORDR_LNO)
                                    End If
                                End If
                            Next
                            If ORDR_QTY_2 <> ORDR_QTY_3 Then
                                EMsg &= vbCr & "Sub-Details out of Balance with Total Amount for Style on Line " & CStr(ORDR_LNO)
                            End If
                        End If
                    Next

                    Dim STYLE_CODEs As String = ""
                    For Each TABLE_NAME As String In New String() {"SOTORDR2", "SOTORDRR"}
                        For Each row As DataRow In ASCDATA1.SelectDistinct _
                                (dst.Tables("SOTORDR2").Select("STYLE_CODE is Not Null"), "STYLE_CODE").Rows
                            Dim STYLE_CODE As String = row.Item("STYLE_CODE") & ""
                            STYLE_CODEs &= ",'" & STYLE_CODE & "'"
                        Next
                    Next
                End If

                If EMsg = "" AndAlso COMPANY_CODE <> "RGI" Then
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                    If rowICTWHSE1.Item("WHSE_TYPE") & "" = "P" Then
                        EMsg &= vbCr & "Invalid Warehouse Code (" & WHSE_CODE & ") for this type of Order"
                    End If
                End If

                Dim INV_FREIGHT As Decimal = Absx1.numFor("INV_FREIGHT").Value
                If INV_FREIGHT < 0 Then
                    EMsg &= vbCr & "Freight Charges may not be less than 0"
                End If

                If EMsg.Length = 0 Then
                    ' remove $0.00 Misc Charges
                    dst.Tables("SOTINVHM").AcceptChanges()
                    For Each row As DataRow In dst.Tables("SOTINVHM").Select("INV_MISC_CHG = 0")
                        row.Delete()
                    Next
                    ' Set all remaining rows to Added
                    dst.Tables("SOTINVHM").AcceptChanges()
                    For Each row As DataRow In dst.Tables("SOTINVHM").Select()
                        row.SetAdded()
                    Next
                End If

                ' Verify the user does not want to effect Inventory
                If EMsg.Length = 0 AndAlso chkNoInventory.Checked Then
                    If MessageBox.Show("You choose Not to Adjust Inventory. Do you want to continue?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

                ' verify negative misc charges
                Dim negativeMiscCharges As Boolean = dst.Tables("SOTINVHM").Select("INV_MISC_CHG < 0 ").Length > 0
                If negativeMiscCharges Then
                    If MessageBox.Show("You provided negative Misc Charges. Do you want to continue?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If


            Case "Cancel"
                If MessageBox.Show("Do you want to cancel entering in an Invoice?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode

                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode

                    .Items("New").Visible = Not InquiryMode
                    .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode
                    .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode
                End With

                .Groups("Totals").Visible = tf
                .Groups("Period Range").Visible = Not tf AndAlso COMPANY_CODE = "NYA"
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        'splHeader.Visible = ScreenMode
        grpHeader.Visible = ScreenMode

        grdSOTINVH1X.Visible = Not tf AndAlso COMPANY_CODE = "NYA"
        splHeader.Visible = tf

        With grdSOTINVHM.DisplayLayout.Override
            If Not InquiryMode AndAlso EntryMode = "N" Then
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            Else
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End If
        End With

        If ScreenMode Then
            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTORDR2}
                With grd.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = DefaultableBoolean.True
                    .AllowUpdate = DefaultableBoolean.True
                End With
            Next
            With grdSOTORDR2.DisplayLayout.Bands(0)
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
            End With

            With grdSOTORDR2.DisplayLayout.Bands(0)
                Dim sample_or_transfer As Boolean = (rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "SAM" Or rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "XFR")

                For Each COLUMN_NAME As String In New String() _
                      {"ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", _
                       "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}
                    .Columns(COLUMN_NAME).Hidden = (EntryMode = "N") Or (COLUMN_NAME.StartsWith("ORDR_AMT") And sample_or_transfer)
                Next
                .Columns("ORDR_RELEASE_AVAIL").Hidden = (EntryMode = "N") Or sample_or_transfer
                .Columns("ORDR_QTY_ALLO").Hidden = (EntryMode = "N") Or sample_or_transfer
                .Columns("ORDR_UNIT_PRICE").Hidden = sample_or_transfer
                .Columns("ORDR_UNIT_PRICE_CALC").Hidden = sample_or_transfer
                .Columns("ORDR_UNIT_PRICE_MANUAL").Hidden = sample_or_transfer
                .Columns("ORDR_PRICE_SOURCE").Hidden = sample_or_transfer

                .Columns("RANGE_STYLE_CODE").Hidden = False Or (ROWs("SOTPARM1").Item("SO_PARM_RANGES") & "" <> "1")
                .Columns("STYLE_CODE_SUB").Hidden = (EntryMode = "N") Or (ROWs("SOTPARM1").Item("SO_PARM_SUB_STYLES") & "" <> "1")
                .Columns("X").Hidden = InquiryMode Or (EntryMode <> "E")

                If rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "XFR" Then
                    .Columns("RANGE_STYLE_CODE").Hidden = True
                    .Columns("STYLE_CODE_SUB").Hidden = True
                    .Columns("X").Hidden = True

                    .Columns("ORDR_AMT").Hidden = True
                    .Columns("ORDR_AMT_OPEN").Hidden = True
                    .Columns("ORDR_AMT_PICK").Hidden = True
                    .Columns("ORDR_AMT_SHIP").Hidden = True
                    .Columns("ORDR_AMT_CANC").Hidden = True
                End If

            End With
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SOTORDR0", "SOTORDR1", "SOTORDR2", "SOTORDR3", "SOTORDR4", "SOTORDR5", "SOTORDR7", "SOTORDR9", _
                                                        "SOTINVHM", "TATEVNT1", "SOTPICK1", "SOTPICK2", "SOTCART1", "SOTCART2", _
                                                        "SOTSHIP1", "SOTSHIP2", "SOTSHIP3", "SOTSHIP4", "SOTSHIP5", "SOTSHIP6", _
                                                        "ARTOPEN1", "SOTINVH1", "SOTINVH2", "ARTCUST1", "ARTCUST2", _
                                                        "ICTCOLR1", "ICTCOLRS", "SOTRSRVX", "SOTRSRV1", "SOTRSRV2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        rowARTCUST1 = Nothing
        rowARTCUST2 = Nothing
        rowSOTORDR1po = Nothing
        rowSOTORDR1 = Nothing
        rowICTSTYL1 = Nothing
        rowARTCUST1_BT = Nothing

        ORDR_NO = String.Empty
        CUST_CODE = String.Empty
        CUST_STORE_NO = String.Empty
        CUST_BILL_TO_CUST = String.Empty
        ORDR_CUST_PO = String.Empty
        SREP_CODE = String.Empty
        SREP2_CODE = String.Empty
        TOTAL_ORDR_AMT = 0
        CUST_DC_NO = String.Empty

        COLOR_CODEs.Clear()
        ORDR_LNOs.Clear()
        STYLE_CODE_last_entry = String.Empty

        ORDR_GROUP_NO = String.Empty
        ORDR_GROUP_NOs.Clear()
        XFR_INV_NOs.Clear()

        Load_SOTINVHX()
        grdSOTINVH1X.Visible = True AndAlso COMPANY_CODE = "NYA"
        splHeader.Visible = False

        ' Not using check box at the moment
        chkNoInventory.Checked = False
        chkNoInventory.Visible = False
        chkNoInventory.Enabled = False

        DisplayTotals()

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Init_Record()

        CUST_CODE = rowSOTORDR1.Item("CUST_CODE")
        CUST_STORE_NO = rowSOTORDR1.Item("CUST_STORE_NO") & ""
        CUST_DC_NO = rowSOTORDR1.Item("CUST_DC_NO") & ""
        ORDR_GROUP_NO = rowSOTORDR1.Item("ORDR_GROUP_NO") & ""
        ORDR_GROUP_NOs.Clear()
        ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)
        CUST_BILL_TO_CUST = rowSOTORDR1.Item("CUST_BILL_TO_CUST") & ""
        If CUST_BILL_TO_CUST <> CUST_CODE Then
            rowARTCUST1_BT = LookUp("ARTCUST1", CUST_BILL_TO_CUST)
        Else
            rowARTCUST1_BT = rowARTCUST1
        End If

        ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = '" & CUST_CODE & "' and CUST_ADDR_TYPE = 'MK'"
        Fill_Records("ARTCUST2", "", True, ASCMAIN1.sql)
        rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "MK", CUST_STORE_NO})

        Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")
        With grdSOTORDR2.DisplayLayout.Bands(0)
            .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
            .Columns("PO_COST").CellActivation = UltraWinGrid.Activation.AllowEdit
            .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            Else
                For Each COLUMN_NAME As String In New String() {"CUST_STYLE_CODE", "CUST_COLOR_CODE", "CUST_SIZE_CODE", "CUST_UPC", "CUST_SKU", "STYLE_RETAIL"}
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                Next
            End If
        End With

        grdSOTORDR2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        Set_Read_Only(splHeader, False)


        Dim rowSOTORDR5 As DataRow = Nothing
        Dim CUST_ADDR_CODEs() As String = {"BY", "MK", "DC"}
        Dim CUST_ADDR_CODE As String = ""
        If EntryMode = "N" Then
            CUST_ADDR_CODEs = {"BY", "BT", "MK", "DC", "ST"}
        End If
        For Each CUST_ADDR_TYPE As String In CUST_ADDR_CODEs
            Dim row As DataRow = Nothing
            If CUST_ADDR_TYPE = "BY" Then
                row = rowARTCUST1
                CUST_ADDR_CODE = CUST_CODE
            ElseIf CUST_ADDR_TYPE = "BT" Then
                row = rowARTCUST1_BT
                CUST_ADDR_CODE = CUST_BILL_TO_CUST
            ElseIf CUST_ADDR_TYPE = "MK" Then
                row = rowARTCUST2
                CUST_ADDR_CODE = CUST_STORE_NO
            ElseIf CUST_ADDR_TYPE = "DC" Then
                row = LookUp("ARTCUST2", New String() {CUST_CODE, "DC", CUST_DC_NO}, True)
                CUST_ADDR_CODE = CUST_DC_NO
            ElseIf CUST_ADDR_TYPE = "ST" Then
                Dim ORDR_ADDR_TYPE_ST As String = rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST")
                CUST_ADDR_CODE = IIf(ORDR_ADDR_TYPE_ST = "DC", CUST_DC_NO, CUST_STORE_NO)
                row = LookUp("ARTCUST2", New String() {CUST_CODE, ORDR_ADDR_TYPE_ST, CUST_ADDR_CODE}, True)
            End If
            rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
            With rowSOTORDR5
                .Item("ORDR_NO") = ORDR_NO
                .Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
                .Item("CUST_ADDR_CODE") = CUST_ADDR_CODE

                If row IsNot Nothing Then
                    For Each COLUMN_NAME As String In New String() _
                        {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", _
                         "CUST_COUNTRY", "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_EMAIL"}
                        Dim COLUMN_NAME_ST As String = COLUMN_NAME
                        .Item(COLUMN_NAME) = row.Item(COLUMN_NAME_ST)
                    Next
                End If

            End With
            dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)
        Next

        If chkCUST_FACTOR_IND.Checked Or rowARTCUST1.Item("CUST_FACTOR_IND") & "" = "1" Then
            chkCUST_FACTOR_IND.Visible = True
        Else
            chkCUST_FACTOR_IND.Visible = False
        End If

        DisplayTotals()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Try

            BeginTrans()
            Me.Cursor = Cursors.WaitCursor

            Dim ALL_ORDERS As New List(Of String)
            ALL_ORDERS.Add(ORDR_NO)

            Dim ORDR_NO_ORIG As String = ORDR_NO

            ASCMAIN1.Progress("Now Updating ...")

            If rowSOTORDR1.Item("ORDR_GROUP_NO") & "" = "" Then
                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    ORDR_GROUP_NO = ORDR_NO
                Else
                    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        ORDR_GROUP_NO = ASCMAIN1.Next_Control_No("ORDR_GROUP_NO")
                    Else
                        ORDR_GROUP_NO = ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO")
                    End If

                End If

                rowSOTORDR1.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                ORDR_GROUP_NOs.Clear()
                ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
            End If


            ASCDATA1.DeleteRows("SOTORDR5", "CUST_ADDR_TYPE <> 'BT' and CUST_ADDR_TYPE <> 'ST'")

            Dim SALES_DIVISION_CODE As String = ""
            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                rowSOTORDR2.Item("ORDR_STATUS") = rowSOTORDR1.Item("ORDR_STATUS")
                If EntryMode = "N" Then
                    rowSOTORDR2.Item("ORDR_QTY_ORIG") = rowSOTORDR2.Item("ORDR_QTY")
                    If SALES_DIVISION_CODE = "" Then
                        Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE") & ""
                        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        SALES_DIVISION_CODE = rowICTSTYL1.Item("SALES_DIVISION_CODE") & ""
                    End If
                End If
            Next

            ' Double-Check SALES_DIVISION_CODE
            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)

            If EntryMode = "N" Then
                If SALES_DIVISION_CODE <> rowSOTORDR1.Item("SALES_DIVISION_CODE") & "" And SALES_DIVISION_CODE <> "" Then
                    rowSOTORDR1.Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
                End If

                rowSOTORDR1.Item("ORDR_ORIG_SHIP_DATE") = rowSOTORDR1.Item("ORDR_SHIP_DATE")
                rowSOTORDR1.Item("ORDR_ORIG_CANCEL_DATE") = rowSOTORDR1.Item("ORDR_CANCEL_DATE")
            End If

            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                If rowICTWHSE1 IsNot Nothing AndAlso rowICTWHSE1.Item("WHSE_CITY") & "" <> "" Then
                    rowSOTORDR1.Item("ORDR_FOB") = rowICTWHSE1.Item("WHSE_CITY") & "," & rowICTWHSE1.Item("WHSE_STATE")
                Else
                    rowSOTORDR1.Item("ORDR_FOB") = ""
                End If
            End If

            ' Remove all Order Details where no qty ordered - SOTORDR2 & SOTORDR3

            ASCDATA1.DeleteRows(dst.Tables("SOTORDR2"), "ISNULL(ORDR_QTY,0) = 0 and ISNULL(ORDR_QTY_OPEN,0) = 0 and ISNULL(ORDR_QTY_SHIP,0) = 0 and ISNULL(ORDR_QTY_PICK,0) = 0 and ISNULL(ORDR_QTY_CANC,0) = 0")

            ' Update all Currency Fields

            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
                rowSOTORDR1.Item("CURR_CODE") = "USD" ' CURR_CODE
                rowSOTORDR1.Item("CURR_EXCH_RATE") = 1 ' CURR_EXCH_RATE
            Next
            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                rowSOTORDR2.Item("ORDR_UNIT_PRICE_CURR") = rowSOTORDR2.Item("ORDR_UNIT_PRICE") / 1 ' CURR_EXCH_RATE
            Next
            For Each rowSOTORDR9 As DataRow In dst.Tables("SOTORDR9").Select("")
                rowSOTORDR9.Item("RANGE_STYLE_PRICE_CURR") = rowSOTORDR9.Item("RANGE_STYLE_PRICE") / 1 ' CURR_EXCH_RATE
                rowSOTORDR9.Item("RANGE_STYLE_PP_PRICE_CURR") = rowSOTORDR9.Item("RANGE_STYLE_PRICE") / 1 ' CURR_EXCH_RATE
            Next

            Record_Event("UPDT", "Sales Order Updated")
            If chkNoInventory.Checked Then
                Record_Event("UPDT", "User choose not to effect inventory (SOFINVHO)")
            End If

            Dim SQLD As String = "ORDR_NO = '" & ORDR_NO & "'"
            INIT_LAST("SOTORDR1", False, , True)
            Update_Record_TDA("SOTORDR1", SQLD)
            Update_Record_TDA("SOTORDR2", SQLD)
            Update_Record_TDA("SOTORDR3", SQLD)
            Update_Record_TDA("SOTORDR4", SQLD)
            Update_Record_TDA("SOTORDR5", SQLD)
            Update_Record_TDA("SOTORDR9", SQLD)
            Update_Record_TDA("TATEVNT1")

            'Create Pick Tickets, Shipments and Cartons
            Dim INV_MISC_CHG As Decimal = Val(dst.Tables("SOTINVHM").Compute("SUM(INV_MISC_CHG)", "") & String.Empty)

            CreatePickTicketsAndShipmentRecords()
            CreateCarton(dst.Tables("SOTPICK1").Rows(0).Item("PICK_NO"))
            Update_Record_TDA("SOTSHIP1")
            Update_Record_TDA("SOTPICK1")
            Update_Record_TDA("SOTPICK2")
            Update_Record_TDA("SOTORDR1")
            Update_Record_TDA("SOTORDR2")
            Update_Record_TDA("SOTCART1")
            Update_Record_TDA("SOTCART2")

            For Each tableName As String In New String() {"SOTPICK1", "SOTPICK2"}
                dst.Tables(tableName).AcceptChanges()
                For Each row As DataRow In dst.Tables(tableName).Select
                    row.SetModified()
                Next
            Next

            ' Update the Sales Order records with the Pick Ticket data
            Dim SOCINVH1 As New TAC.SOCINVH1(dst)
            SOCINVH1.ProcessPickTicketsAndUpdateSalesDetails(CDate(DateTime.Now.ToShortDateString))
            dst.Tables("SOTPICK1").Rows(0).Item("PICK_STATUS") = "F"

            Update_Record_TDA("SOTSHIP1")
            Update_Record_TDA("SOTPICK1")
            Update_Record_TDA("SOTPICK2")
            Update_Record_TDA("SOTORDR1")
            Update_Record_TDA("SOTORDR2")
            Update_Record_TDA("SOTCART1")
            Update_Record_TDA("SOTCART2")

            For Each ORDR_NOx As String In ALL_ORDERS
                Dependent_Updates(1, ORDR_NOx)
            Next
            ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})

            ' Refresh the data so the Invoice class has all the required fields
            ASCMAIN1.sql = sqlSOTPICK1 & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTPICK1.PICK_NO = '" & dst.Tables("SOTPICK1").Rows(0).Item("PICK_NO") & "'"
            Fill_Records("SOTPICK1", "", True, ASCMAIN1.sql)

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
                rowSOTPICK1.Item("SELECTED") = "1"
            Next
            dst.Tables("SOTPICK1").Rows(0).Item("INV_MISC_CHG") = INV_MISC_CHG

            ASCMAIN1.sql = sqlSOTPICK2 & vbCrLf _
                & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTPICK1.PICK_NO = '" & dst.Tables("SOTPICK1").Rows(0).Item("PICK_NO") & "'"
            Fill_Records("SOTPICK2", "", True, ASCMAIN1.sql)

            Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
            Dim WHSE_PHYS_STATUS As String = rowICTWHSE1.Item("WHSE_PHYS_STATUS") & ""
            Dim WHSE_LOCATOR As Boolean = rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1"

            Dim RFIXMSG As Boolean = False
            Dim SHIP_BOL_NO As String = String.Empty

            ' Create Invoice Records
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)
                SHIP_BOL_NO = rowSOTSHIP1.Item("SHIP_BOL_NO")
                rowSOTSHIP1.Item("SHIP_DATE_SHIPPED") = rowSOTSHIP1.Item("INV_DATE")
                SOCINVH1.CreateInvoices(SHIP_BOL_NO, RFIXMSG)
                rowSOTSHIP1.Item("SHIP_STATUS") = "F"
            Next

            Dim INV_NO As String = dst.Tables("SOTPICK1").Rows(0).Item("INV_NO")
            Dim INV_MNO As Int16 = 1
            For Each rowSOTINVHM As DataRow In dst.Tables("SOTINVHM").Select("", "INV_MNO")
                rowSOTINVHM.Item("INV_NO") = INV_NO
                rowSOTINVHM.Item("INV_MNO") = INV_MNO
                INV_MNO += 1
            Next

            For Each TABLE_NAME As String In New String() {"SOTINVH1", "SOTINVH2", "SOTINVHM", "ARTOPEN1", "SOTPICK1", "SOTSHIP1"}
                Update_Record_TDA(TABLE_NAME)
            Next

            For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Rows

                If Not chkNoInventory.Checked Then

                    Dim PICK_NO As String = dst.Tables("SOTPICK1").Select(" INV_NO = '" & rowSOTINVH1.Item("INV_NO") & "'")(0).Item("PICK_NO")

                    ' This code was taken from Sales Order Release - SOROREL1, I added - and SOTPICK1.PICK_NO = '" & PICK_NO & "'"
                    ' This is needed for procedure call SOPSTAT1
                    ' Also Modified whse_qty_open to not be changed
                    '     WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) - NVL(R1.PICK_QTY,0) - NVL(R1.PICK_QTY_CANC_REL,0)" & vbCrLf _

                    ASCMAIN1.sql = "" _
                        & "Begin Declare Cursor C1 is " & vbCrLf _
                        & " Select SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                        & ", Sum (SOTPICK2.PICK_QTY) PICK_QTY" & vbCrLf _
                        & ", Sum (SOTPICK2.PICK_QTY_CANC_REL) PICK_QTY_CANC_REL" & vbCrLf _
                        & " from SOTORDR1, SOTORDR2, SOTPICK1, SOTPICK2 " & vbCrLf _
                        & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                        & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                        & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                        & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                        & "   and SOTPICK1.PICK_NO = '" & PICK_NO & "'" & vbCrLf _
                        & " group by SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE;" & vbCrLf _
                        & " Begin" & vbCrLf _
                        & "  For R1 in C1 Loop" & vbCrLf _
                        & "   Update ICTSTAT2 SET WHSE_QTY_PICK = NVL(WHSE_QTY_PICK,0) + NVL(R1.PICK_QTY,0), " & vbCrLf _
                        & "                        WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) - 0 - 0 " & vbCrLf _
                        & "   where STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
                        & "     and COLOR_CODE = R1.COLOR_CODE" & vbCrLf _
                        & "     and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
                        & "  End Loop;" & vbCrLf _
                        & " End;" & vbCrLf _
                        & "End;"
                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.sql = "BEGIN SOPSTAT1('" & rowSOTINVH1.Item("INV_TYPE") & "','" & rowSOTINVH1.Item("INV_NO") & "'); END;"
                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.sql = "BEGIN SOPSTAT2('" & rowSOTINVH1.Item("INV_TYPE") & "','" & rowSOTINVH1.Item("INV_NO") & "'); END;"
                    ASCDATA1.ExecuteSQL()

                    ' When should this be called - Only whses that use Locationss
                    If WHSE_LOCATOR Then
                        TAC.ICCMAIN1.Update_WHTLOCBX("S", rowSOTINVH1.Item("INV_NO"))
                    End If
                End If

                ASCDATA1.ExecuteSP("ARPCUST6_IC", "VV", _
                   New Object() {rowSOTINVH1.Item("INV_TYPE"), rowSOTINVH1.Item("INV_NO")}, _
                   New String() {"INV_TYPE_IN", "INV_NO_IN"})
            Next

            ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
            ASCDATA1.ExecuteSQL()

            INV_NO = dst.Tables("SOTINVH1").Rows(0).Item("INV_NO") & String.Empty

            CommitTrans("Update Complete. Invoice Created: " & INV_NO)

            Try
                ASCMAIN1.Progress("Creating Web Invoice", "")
                For Each row As DataRow In dst.Tables("SOTINVH1").Select("")
                    TAC.SOCMAIN1.CreateWebInvoice(Me, row.Item("INV_TYPE"), row.Item("INV_NO"))
                Next
            Catch ex As Exception

            Finally
                ASCMAIN1.Progress("", "")
            End Try

        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

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

            Case "View", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("ORDR_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTORDR1"
            E.COLUMN_NAME = "ORDR_NO"
            E.CODE_VALUE = Absx1.txtFor("ORDR_NO").Text
            E.DESC_VALUE = "Sales Order"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTRSRV1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTINVH1X, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Refresh")
        Load_Popup_Menu(grdSOTORDR2, "BBS", "Style Status Inquiry", "Style Master File", "Show UPC/SKU")

    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        Select Case grd.Name
            Case "grdSOTINVHH"
                'tlb_btn = DirectCast(tlb_pop.Tools("Credit Entire Invoice"), UltraWinToolbars.ButtonTool)
                'tlb_btn.SharedProps.Visible = Not InquiryMode
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSOTORDR2"
                    tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                    tlb_sbt = DirectCast(tlb_pop.Tools("Show UPC/SKU"), UltraWinToolbars.StateButtonTool)
                    tlb_sbt.SharedProps.Visible = True
                    tlb_sbt.Tag = "X"
                    tlb_sbt.Checked = Not grdSOTORDR2.DisplayLayout.Bands(0).Columns("CUST_UPC").Hidden
                    tlb_sbt.Tag = ""
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing


        Select Case e.Tool.Key

            Case "Refresh"
                Load_SOTINVHX()

            Case "Show UPC/SKU"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag <> "X" Then
                    Toggle_Customer_Style_Fields(tlb_sbt.Checked)
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Style Master File"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    'Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICTSTYL1")
                    Dim keys As New Dictionary(Of String, Object)
                    keys.Add("STYLE_CODE", STYLE_CODE)
                    ' If ASCMAIN1.Running_in_VS Then Stop ' NOT WORKING
                    Context_Launch("View", keys, e.Tool.Key, "ICTSTYL1")
                End If

        End Select
    End Sub

    Sub Toggle_Customer_Style_Fields(show As Boolean)
        With grdSOTORDR2.DisplayLayout.Bands(0)
            .Columns("CUST_UPC").Hidden = Not show
            .Columns("CUST_SKU").Hidden = Not show
            .Columns("CUST_STYLE_CODE").Hidden = Not show
            .Columns("CUST_COLOR_CODE").Hidden = Not show
            .Columns("CUST_SIZE_CODE").Hidden = Not show
            .Columns("STYLE_RETAIL").Hidden = Not show
        End With
    End Sub

#End Region

#Region "ABSColumn Controls"

#End Region

#Region "grdSOTORDR2"

    Private Sub grdSOTORDR2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTORDR2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim STYLE_CODE As String = Validate_Style(e.Cell.Value & "") ' grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value)
                If STYLE_CODE <> "" Then
                    e.Cell.Row.Cells("STYLE_UOM").Value = rowICTSTYL1.Item("STYLE_UOM") & ""
                    e.Cell.Row.Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                    e.Cell.Row.Cells("INNER_PACK_QTY").Value = rowICTSTYL1.Item("INNER_PACK_QTY")
                    e.Cell.Row.Cells("CARTON_PACK_QTY").Value = rowICTSTYL1.Item("CARTON_PACK_QTY")
                    e.Cell.Row.Cells("RANGE_STYLE_CODE").Value = DBNull.Value
                    e.Cell.Row.Cells("STYLE_CLASS_CODE").Value = rowICTSTYL1.Item("STYLE_CLASS_CODE")
                    e.Cell.Row.Cells("STYLE_PRICE").Value = rowICTSTYL1.Item("STYLE_PRICE")

                    If COLOR_CODEs.Count = 1 Then
                        e.Cell.Row.Cells("COLOR_CODE").Value = COLOR_CODEs(0)
                    End If
                Else

                End If

            Case "COLOR_CODE"
                Dim COLOR_CODE As String = e.Cell.Value & "" ' grdSOTORDR2.ActiveRow.Cells("COLOR_CODE").Value & ""
                If COLOR_CODE <> "" Then
                    If COLOR_CODEs.Contains(COLOR_CODE) Then
                        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                        e.Cell.Row.Cells("COLOR_DESC").Value = rowICTCOLR1.Item("COLOR_DESC")
                    End If
                End If

            Case "RANGE_STYLE_CODE"
                If grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                    grdSOTORDR2.ActiveRow.Cells("STYLE_UOM").Value = "EA"
                    grdSOTORDR2.ActiveCell = grdSOTORDR2.ActiveRow.Cells("STYLE_DESC")
                    grdSOTORDR2.ActiveRow.Update()
                End If

            Case "ORDR_QTY"
                grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_OPEN").Value = grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value

                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    Dim ORDR_PRICE_SOURCE As String = ""
                    Dim ORDR_UNIT_PRICE_CALC As Decimal = 0

                    ORDR_UNIT_PRICE_CALC = TAC.SOCMAIN1.Price_Line(Me, CUST_CODE, rowARTCUST1,
                           grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "", _
                           grdSOTORDR2.ActiveRow.Cells("COLOR_CODE").Value & "",
                           Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & ""), ORDR_PRICE_SOURCE)

                    e.Cell.Row.Cells("ORDR_UNIT_PRICE_CALC").Value = ORDR_UNIT_PRICE_CALC
                    e.Cell.Row.Cells("ORDR_PRICE_SOURCE").Value = ORDR_PRICE_SOURCE

                    If e.Cell.Row.Cells("ORDR_UNIT_PRICE_MANUAL").Value & "" <> "1" Then
                        e.Cell.Row.Cells("ORDR_UNIT_PRICE").Value = ORDR_UNIT_PRICE_CALC
                    End If
                End If

            Case "ORDR_QTY_OPEN"
                grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & "") _
                    - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_SHIP").Value & "") _
                    - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_OPEN").Value & "") _
                    - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_PICK").Value & "")
                If Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Text) < 0 Then
                    grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value = 0
                End If

            Case "ORDR_UNIT_PRICE"
                If Val(grdSOTORDR2.ActiveRow.Cells("ORDR_UNIT_PRICE").Value & "") <> Val(grdSOTORDR2.ActiveRow.Cells("ORDR_UNIT_PRICE_CALC").Value & "") Then
                    grdSOTORDR2.ActiveRow.Cells("ORDR_UNIT_PRICE_MANUAL").Value = "1"
                End If

            Case "ORDR_UNIT_PRICE_MANUAL"

                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    Dim ORDR_PRICE_SOURCE As String = ""
                    Dim ORDR_UNIT_PRICE_CALC As Decimal = 0

                    ORDR_UNIT_PRICE_CALC = TAC.SOCMAIN1.Price_Line(Me, CUST_CODE, rowARTCUST1,
                               grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "", _
                               grdSOTORDR2.ActiveRow.Cells("COLOR_CODE").Value & "",
                               Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & ""), ORDR_PRICE_SOURCE)
                    e.Cell.Row.Cells("ORDR_UNIT_PRICE_CALC").Value = ORDR_UNIT_PRICE_CALC
                    If e.Cell.Row.Cells("ORDR_UNIT_PRICE_MANUAL").Value & "" <> "1" Then
                        e.Cell.Row.Cells("ORDR_UNIT_PRICE").Value = ORDR_UNIT_PRICE_CALC
                    End If
                End If
        End Select
    End Sub

    Private Sub grdSOTORDR2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDR2.AfterRowActivate

        If grdSOTORDR2.ActiveRow Is Nothing OrElse grdSOTORDR2.ActiveRow.IsAddRow Then

        Else
            Dim STYLE_CODE As String = grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & ""
            Dim COLOR_CODE As String = grdSOTORDR2.ActiveRow.Cells("COLOR_CODE").Value & ""
            Dim STYLE_CLASS_CODE As String = grdSOTORDR2.ActiveRow.Cells("STYLE_CLASS_CODE").Value & ""
            Dim CARTON As String = grdSOTORDR2.ActiveRow.Cells("STYLE_CLASS_CODE").Value & ""

        End If

        If Trim(grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "") = "" And _
           Trim(grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "") = "" And _
            (grdSOTORDR2.ActiveCell Is Nothing OrElse _
             (grdSOTORDR2.ActiveCell.Column.Key <> "STYLE_CODE" And _
              grdSOTORDR2.ActiveCell.Column.Key <> "RANGE_STYLE_CODE")) _
        Then
            grdSOTORDR2.ActiveCell = grdSOTORDR2.ActiveRow.Cells("STYLE_CODE")
        End If

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            With grdSOTORDR2.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE_MANUAL")
                If grdSOTORDR2.ActiveRow.IsAddRow Then
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            End With
        End If

        If grdSOTORDR2.ActiveRow.IsAddRow Then
            With grdSOTORDR2.DisplayLayout.Bands(0)
                .Columns("RANGE_STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("STYLE_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            End With

            ''- DOUBLE REMS ARE MINE
            If grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "" = "" _
                And grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "" = "" Then

                If grdSOTORDR2.ActiveCell.Column.Key = "STYLE_CODE" Then
                    'Setup_SubGrid(False, True)
                    '' grdSOTORDR2.ActiveCell = grdSOTORDR2.ActiveRow.Cells("STYLE_CODE")
                Else
                    'Setup_SubGrid(True, True)
                    '' grdSOTORDR2.ActiveCell = grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE")
                End If
            Else
                ' i don't think I ever hit this code - it should be removed - DOUBLE REMS ARE MINE

                ''With grdSOTORDR2.DisplayLayout.Bands(0)
                ''    If grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                ''        .Columns("STYLE_DESC").CellActivation = UltraWinGrid.Activation.AllowEdit
                ''        If RANGE_TYPE <> "A" Then
                ''            .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                ''        Else
                ''            .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                ''        End If
                ''    Else
                ''        .Columns("STYLE_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                ''    End If
                ''End With
            End If

        Else
            If grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "" = "" Then
                'Setup_SubGrid(True, False)
            Else
                'Setup_SubGrid(False, False)
            End If

            With grdSOTORDR2.DisplayLayout.Bands(0)
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("RANGE_STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit

                If grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "" = "" Then
                    Validate_Style(grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "")
                    .Columns("STYLE_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    .Columns("STYLE_DESC").CellActivation = UltraWinGrid.Activation.AllowEdit
                End If

                If Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_PICK").Value & "") <> 0 _
                Or Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_SHIP").Value & "") <> 0 _
                Or Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value & "") <> 0 _
                Then
                    'Or absx1.txtfor("ORDR_SOURCE")).Text = "E" 'was also part of this
                    .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            End With

            'If grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
            '    If Val(CInt(fpPPQTY.Text)) = 1 Or Val(fpPPQTY.Text) = 0 Then
            '        SetRangeType("R")
            '    Else
            '        SetRangeType("A")
            '    End If
            'End If
        End If
    End Sub

    Private Sub grdSOTORDR2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTORDR2.AfterRowsDeleted

        DisplayTotals()

        If grdSOTORDR2.Rows.Count = 0 Then
            Absx1.txtFor("SALES_DIVISION_CODE").Text = ""
        End If

        Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")

    End Sub

    Private Sub grdSOTORDR2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTORDR2.AfterRowUpdate

        DisplayTotals()

        If e.Row.Cells("STYLE_CODE").Value & "" <> "" And Absx1.txtFor("SALES_DIVISION_CODE").Text = "" Then
            Absx1.txtFor("SALES_DIVISION_CODE").Text = rowICTSTYL1.Item("SALES_DIVISION_CODE") & ""
        End If

        ' If e.Row.IsAddRow Then
        ' if we just added a row
        If EntryMode = "N" Or EntryMode = "E" Then
            If e.Row.Cells("ORDR_STATUS").Tag & "" = "Added" Then
                Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")
                grdSOTORDR2.DisplayLayout.Bands(0).AddNew()
                e.Row.Cells("ORDR_STATUS").Tag = DBNull.Value
            End If
        End If
        ' End If

    End Sub

    Private Sub grdSOTORDR2_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTORDR2.BeforeCellUpdate

        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim STYLE_CODE As String = Validate_Style(e.NewValue & "")
                If STYLE_CODE = "" Then
                    e.Cancel = True
                End If

            Case "RANGE_STYLE_CODE"

        End Select

    End Sub

    Private Sub grdSOTORDR2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTORDR2.BeforeExitEditMode
        If grdSOTORDR2.ActiveCell IsNot Nothing Then
            With grdSOTORDR2.ActiveCell
                Select Case .Column.Key
                    Case "STYLE_CODE", "COLOR_CODE", "RANGE_STYLE_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)

                    Case "ORDR_QTY"
                        If .EditorResolved.Value & "" = "" _
                        Or Val(.EditorResolved.Value & "") < 0 _
                        Then
                            .EditorResolved.Value = 0
                        End If
                        If Val(.EditorResolved.Value & "") < 0 Then
                            .EditorResolved.Value = System.Math.Abs(Val(.EditorResolved.Value & ""))
                        End If

                    Case "ORDR_QTY_OPEN"
                        If .EditorResolved.Value & "" = "" _
                        Or Val(.EditorResolved.Value & "") < 0 _
                        Then
                            .EditorResolved.Value = 0
                        End If
                        Dim q As Int64 = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_PICK").Value & "") _
                                       + Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_SHIP").Value & "")
                        '+ Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value & "")
                        If Val(.EditorResolved.Value & "") > Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & "") - q Then
                            .EditorResolved.Value = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & "") - q
                        End If
                End Select
            End With
        End If
    End Sub


    Private Sub grdSOTORDR2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTORDR2.BeforeRowsDeleted

        ORDR_LNOs.Clear()

        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If Not grow.IsAddRow Then
                If Val(grow.Cells("ORDR_QTY_PICK").Value & "") <> 0 _
                Or Val(grow.Cells("ORDR_QTY_SHIP").Value & "") <> 0 _
                Or Val(grow.Cells("ORDR_QTY_CANC").Value & "") <> 0 _
                Then
                    MsgBox("Cannot Delete a Line if it has ever been " & vbCr & "Picked, Shipped Or Cancelled" & vbCr & "Use the Cancel Button (x)")
                    e.Cancel = True
                    Exit Sub
                End If

                If grow.Cells("RSRV_NO").Value & "" <> "" Then
                    MsgBox("Cannot Delete a Line if it has ever been " _
                           & vbCrLf & "Used in a Reservation" & vbCrLf & "Use the Cancel Button (x)")
                    e.Cancel = True
                    Exit Sub
                End If

                ORDR_LNOs.Add(grow.Cells("ORDR_LNO").Value)
            End If
        Next
    End Sub

    Private Sub grdSOTORDR2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTORDR2.BeforeRowUpdate
        Dim iResult As String

        If addingRecord Then Exit Sub

        If e.Row.Cells("RANGE_STYLE_CODE").Value & "" = "" Then
            Validate_Columns("STYLE_CODE", e.Cancel)
            If e.Cancel Then
                MsgBox("Invalid Style")
            End If
            If Not e.Cancel Then
                Validate_Columns("COLOR_CODE", e.Cancel)
                If e.Cancel Then
                    MsgBox("Invalid Color")
                End If
            Else

            End If
            If Not e.Cancel Then
                Validate_Columns("ORDR_QTY", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
                If e.Cancel Then
                    MsgBox("Invalid Qty")
                End If
            End If
        Else
            'If multi_store_is_active Then
            '    MsgBox("Cannot Have Range Styles on a Multi-Store Order", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            '    e.Cancel = True
            '    Exit Sub
            'End If
        End If

        If rowSOTORDR1.Item("ORDR_SOURCE") = "K" Then 'Only Check KeyBrd Orders 11/24/04 W.R.
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                If e.Row.Cells("CUST_UPC").Value & "" <> "" Then
                    iResult = TAC.TACMAIN1.Validate_UPC(e.Row.Cells("CUST_UPC").Value & "", ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
                    If iResult & "" <> "" Then
                        MsgBox(iResult, vbOKOnly, "UPC Error")
                        e.Cancel = True
                    End If
                End If
            End If
        End If

        If Not e.Cancel Then
            If e.Row.IsAddRow Then
                Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & ""
                Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value & ""
                If dst.Tables("SOTORDR2").Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'").Length <> 0 Then
                    'MsgBox("Style/Color " & "" & " already on Order")
                    'ASCMAIN1.Progress("Style/Color " & STYLE_CODE & "/" & COLOR_CODE & " is already on Order")
                    MsgBox("Style/Color " & STYLE_CODE & "/" & COLOR_CODE & " is already on Order", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    e.Cancel = True
                End If
            End If
        End If


        If e.Cancel = True Then
            Exit Sub
        End If


        STYLE_CODE_last_entry = e.Row.Cells("STYLE_CODE").Value & ""

        If e.Row.IsAddRow Then
            e.Row.Cells("ORDR_NO").Value = ORDR_NO
            Dim ORDR_LNO As Int64 = Val(dst.Tables("SOTORDR2").Compute("MAX(ORDR_LNO)", "") & "") + 1
            e.Row.Cells("ORDR_LNO").Value = ORDR_LNO
            e.Row.Cells("ORDR_QTY_ORIG").Value = e.Row.Cells("ORDR_QTY").Value
            e.Row.Cells("ORDR_STATUS").Value = "O"
            e.Row.Cells("ORDR_STATUS").Tag = "Added"
            If e.Row.Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                e.Row.Cells("STYLE_CODE").Value = DBNull.Value
                e.Row.Cells("COLOR_CODE").Value = DBNull.Value
                e.Row.Cells("COLOR_DESC").Value = DBNull.Value
            End If
        End If
    End Sub

    Private Sub grdSOTORDR2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTORDR2.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "X"
                    If Val(.Cells("ORDR_QTY_CANC").Value & "") <> 0 Then
                        If MsgBox("Restore Cancelled Qty of " & .Cells("ORDR_QTY_CANC").Value, _
                                  MsgBoxStyle.YesNo, "Option to Restore Qty Cancelled") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        If .Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                            For Each rowSOTORDRR As DataRow In dst.Tables("SOTORDRR").Select("ORDR_LNO = " & .Cells("ORDR_LNO").Value)
                                rowSOTORDRR.Item("ORDR_QTY_OPEN") = Val(rowSOTORDRR.Item("ORDR_QTY_OPEN") & "") + Val(rowSOTORDRR.Item("ORDR_QTY_CANC") & "")
                                rowSOTORDRR.Item("ORDR_QTY_CANC") = 0
                            Next
                        End If
                        .Cells("ORDR_QTY_OPEN").Value = Val(.Cells("ORDR_QTY_OPEN").Value & "") + Val(.Cells("ORDR_QTY_CANC").Value & "")
                        ' grdSOWORDR2_AfterColUpdate(.Cells("ORDR_QTY_OPEN").position)
                        .Update()
                    Else
                        If MsgBox("Cancel Remaining Qty Open of " & .Cells("ORDR_QTY_OPEN").Value, _
                                  MsgBoxStyle.YesNo, "Option to Restore Qty Cancelled") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        If .Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                            For Each rowSOTORDRR As DataRow In dst.Tables("SOTORDRR").Select("ORDR_LNO = " & .Cells("ORDR_LNO").Value)
                                rowSOTORDRR.Item("ORDR_QTY_OPEN") = 0
                                Dim ORDR_QTY_CANC As Int64 = Val(rowSOTORDRR.Item("ORDR_QTY") & "") _
                                                           - Val(rowSOTORDRR.Item("ORDR_QTY_SHIP") & "") _
                                                           - Val(rowSOTORDRR.Item("ORDR_QTY_PICK") & "")
                                rowSOTORDRR.Item("ORDR_QTY_CANC") = IIf(ORDR_QTY_CANC < 0, 0, ORDR_QTY_CANC)
                            Next
                        End If
                        .Cells("ORDR_QTY_OPEN").Value = "0"
                        ' grdSOWORDR2_AfterColUpdate(.Cells("ORDR_QTY_OPEN").position)
                        grdSOTORDR2.ActiveRow.Update()
                    End If

                Case "STYLE_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdSOTORDR2, sql_where)

                Case "COLOR_CODE"
                    Dim sql_where As String = "COLOR_CODE in ('" & Join(COLOR_CODEs.ToArray, "','") & "')"
                    grdClickCellButton(grdSOTORDR2, sql_where)

                Case "RANGE_STYLE_CODE"
                    If e.Cell.Value & "" = "" Then
                        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("RANGE_STYLE_CODE", "ICTRSTY1")
                        If ASCMAIN1.CodeSelector.SQL <> "" Then
                            ASCMAIN1.CodeSelector.MultipleSelections = False
                            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                            ASCMAIN1.CodeSelector.COLUMN_PREKEYs.Add("CUST_CODE", Absx1.txtFor("CUST_CODE").Text)
                            Using F As New ASFCODE1
                                F.ShowDialog()
                            End Using
                            If ASCMAIN1.CodeSelector.SelectedCode <> "" Then
                                grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value = ASCMAIN1.CodeSelector.SelectedCode
                                Dim ORDR_LNO As Int64 = Val(e.Cell.Row.Cells("ORDR_LNO").Value & "")
                                'GET_RANGE(grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value, ORDR_LNO)
                            End If
                        End If
                    End If

                Case "STYLE_CODE_SUB"
                    If .IsAddRow Then
                        MsgBox("You Must First Complete the Line, then you may Specify a Substitute", _
                               MsgBoxStyle.OkOnly, "Cannot Specify a Substitute on a Line Not already Added to Order")
                        Exit Sub
                    End If
                    Dim STYLE_CODE_SUB As String
                    Dim COLOR_CODE_SUB As String
                    If Val(.Cells("ORDR_QTY_PICK").Value & "") = 0 And _
                       Val(.Cells("ORDR_QTY_SHIP").Value & "") = 0 And _
                       Val(.Cells("ORDR_QTY_CANC").Value & "") = 0 Then

                        COLOR_CODE_SUB = .Cells("COLOR_CODE").Value & ""
                        STYLE_CODE_SUB = Select_Style(COLOR_CODE_SUB)

                        If STYLE_CODE_SUB <> "" Then
                            Dim z As String = .Cells("STYLE_CODE").Value
                            Dim STYLE_CODE As String = Validate_Style(STYLE_CODE_SUB)
                            If STYLE_CODE = "" Then
                                STYLE_CODE = z
                                Validate_Style(z)
                            Else
                                If .Cells("STYLE_CODE_SUB").Value = "" Then
                                    .Cells("STYLE_CODE_SUB").Value = .Cells("STYLE_CODE").Value
                                End If
                                .Cells("STYLE_CODE").Value = STYLE_CODE_SUB
                                .Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                                .Update()
                            End If
                        End If
                    Else
                        ' CANNOT SUB STYLE IF PICKED, SHIPPED OR CANCELLED
                    End If
            End Select
        End With

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
        DisplayTotals()
    End Sub

    Private Sub grdSOTINVHM_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTINVHM.AfterRowUpdate
        DisplayTotals()
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

        Dim INV_MISC_CHG As Double = Val(e.Row.Cells("INV_MISC_CHG").Value & String.Empty)
        If INV_MISC_CHG = 0 Then
            MessageBox.Show("Charge amount must be unequal to $0.00.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

    Private Sub grdSOTINVH1X_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTINVH1X.AfterRowUpdate
        If e.Row.Cells("SEL").Value & String.Empty = "1" Then
            If MyBase.Absx1.txtFor("CUST_CODE").Text = String.Empty Then
                MyBase.Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value & String.Empty
                MyBase.Absx1.txtFor("CUST_STORE_NO").Text = e.Row.Cells("CUST_STORE_NO").Value & String.Empty
            End If
        End If
    End Sub

    Private Sub cmdInvoiceHistory_Click(sender As System.Object, e As System.EventArgs) Handles cmdInvoiceHistory.Click
        Load_SOTINVHX()
    End Sub

    Private Sub txtInvFreight_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtInvFreight.ValueChanged
        DisplayTotals()
    End Sub

#End Region

#Region "Form Procedures"

    Private Sub CreateCarton(ByVal PICK_NO As String)

        Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow
        Dim CART_NO As String = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTCART1.CART_NO"), "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
        rowSOTCART1.Item("CART_NO") = CART_NO ' "NEW" & Format(CART_NO_new, "0000000")
        rowSOTCART1.Item("PICK_NO") = PICK_NO
        dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)

        rowSOTCART1.Item("CART_TOTAL_UNITS") = dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY)", "")
        rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = 1
        rowSOTCART1.Item("CART_TOTAL_WGT_CALC") = 1
        rowSOTCART1.Item("CART_TRACKING_NO") = String.Empty
        'rowSOTCART1.Item("CART_SEQ") = String.Empty
        'rowSOTCART1.Item("CART_MEMO") = String.Empty
        'rowSOTCART1.Item("CART_TYPE") = String.Empty
        rowSOTCART1.Item("PACKAGING_TYPE") = 31
        rowSOTCART1.Item("PKG_CODE") = String.Empty
        Dim CART_SEQ As Int32 = Val(dst.Tables("SOTCART1").Compute("MAX(CART_SEQ)", "PICK_NO = '" & PICK_NO & "'") & String.Empty) + 1
        rowSOTCART1.Item("CART_SEQ") = CART_SEQ


        Dim CART_LNO As Int16 = 1
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")

            Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow
            rowSOTCART2.Item("CART_NO") = CART_NO
            rowSOTCART2.Item("CART_LNO") = CART_LNO
            CART_LNO += 1

            rowSOTCART2.Item("ORDR_NO") = rowSOTORDR2.Item("ORDR_NO")
            rowSOTCART2.Item("ORDR_LNO") = rowSOTORDR2.Item("ORDR_LNO")
            rowSOTCART2.Item("QTY_PACKED") = rowSOTORDR2.Item("ORDR_QTY")
            'rowSOTCART2.Item("SKU_NO") = rowSOTORDR2.Item("SKU_NO")
            rowSOTCART2.Item("STYLE_CODE") = rowSOTORDR2.Item("STYLE_CODE")
            rowSOTCART2.Item("COLOR_CODE") = rowSOTORDR2.Item("COLOR_CODE")

            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {rowSOTORDR2.Item("STYLE_CODE"), rowSOTORDR2.Item("COLOR_CODE")})
            If rowICTSTYC1 IsNot Nothing Then
                rowSOTCART2.Item("UPC_CODE") = rowICTSTYC1.Item("UPC_CODE")
            End If

            'ICTSIZE1.SIZE_CODE SIZE_DESC
            ' ICTSIZE1.NRF_SIZE_CODE (+) = SOTORDR2.CUST_SIZE_CODE
            Dim rowICTSIZE1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ICTSIZE1 WHERE NRF_SIZE_CODE = '" & rowSOTORDR2.Item("CUST_SIZE_CODE") & "'")
            If rowICTSIZE1 IsNot Nothing Then
                rowSOTCART2.Item("SIZE_DESC") = rowICTSIZE1.Item("SIZE_CODE")
            End If

            'rowSOTCART2.Item("STYLE_PREPACK") = rowSOTORDR2.Item("")

            dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)
        Next

    End Sub

    Sub Init_Record()
        Dim preloadSalesOrder As Boolean = False

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            ORDR_NO = ASCMAIN1.Next_Control_No("ORDR_NO")
        Else
            ORDR_NO = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
        End If

        Dim WHSE_CODE As String = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
        If rowARTCUST1.Item("WHSE_CODE") & "" <> "" Then WHSE_CODE = rowARTCUST1.Item("WHSE_CODE")
        If WHSE_CODE = "" Then WHSE_CODE = ""

        If XFR_INV_NOs.Count > 0 Then
            Dim ORDR_NO_XFR As String = dst.Tables("SOTINVHX").Select("ORDR_NO <> '' and SEL = '1'")(0).Item("ORDR_NO")
            Fill_Records("SOTORDR1", ORDR_NO_XFR)
            rowSOTORDR1 = dst.Tables("SOTORDR1")(0)
            rowSOTORDR1.AcceptChanges()
            rowSOTORDR1.SetAdded()
            preloadSalesOrder = True
            rowSOTORDR1.Item("WHSE_CODE") = rowSOTORDR1.Item("WHSE_CODE_TO")
            rowSOTORDR1.Item("WHSE_CODE_TO") = DBNull.Value
            rowSOTORDR1.Item("LAST_OPER") = DBNull.Value
            rowSOTORDR1.Item("LAST_DATE") = DBNull.Value

            If COMPANY_CODE = "RGI" Or COMPANY_CODE = "RGI" Then
                ORDR_GROUP_NO = ORDR_NO
            Else
                If COMPANY_CODE = "VAN" Or COMPANY_CODE = "VAN" Then
                    ORDR_GROUP_NO = ASCMAIN1.Next_Control_No("ORDR_GROUP_NO")
                Else
                    ORDR_GROUP_NO = ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO")
                End If

            End If
            rowSOTORDR1.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
            ORDR_GROUP_NOs.Clear()
            ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)

        Else
            rowSOTORDR1 = dst.Tables("SOTORDR1").NewRow
            rowSOTORDR1.Item("WHSE_CODE") = WHSE_CODE
        End If

        With rowSOTORDR1
            .Item("ORDR_NO") = ORDR_NO
            .Item("CUST_CODE") = CUST_CODE
            .Item("CUST_STORE_NO") = CUST_STORE_NO

            If ORDR_CUST_PO.Length = 0 Then
                ORDR_CUST_PO = "XFR_" & ORDR_NO
            End If
            .Item("ORDR_CUST_PO") = ORDR_CUST_PO
            .Item("ORDR_DATE") = DATETIME_STAMP.Date
            .Item("ORDR_SOURCE") = "K"
            .Item("ORDR_STATUS") = "O"
            rowSOTORDR1.Item("ORDR_TYPE_CODE") = "REG"
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("ORDR_DATE_BOOKED") = DATETIME_STAMP.Date
            .Item("ORDR_YYYYPP_BOOKED") = ASCMAIN1.CYP
            .Item("ORDR_ADDR_TYPE_ST") = "DC"
            .Item("ORDR_ADDR_TYPE_ST") = "DC"

            'ORDR_SHIP_DATE, ORDR_ARRIVAL_DATE, ORDR_SHIP_DATE, ORDR_CANCEL_DATE
            .Item("ORDR_SHIP_DATE") = DateTime.Now.ToShortDateString
            .Item("ORDR_ARRIVAL_DATE") = DateTime.Now.ToShortDateString
            .Item("ORDR_CANCEL_DATE") = DateTime.Now.ToShortDateString

            Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO})
            Dim rowARTCUST3 As DataRow = LookUp("ARTCUST3", New String() {CUST_CODE, "MK", CUST_STORE_NO, "DC"})
            If rowARTCUST3 IsNot Nothing AndAlso rowARTCUST3.Item("CUST_ADDR_CODE2") & "" <> "" Then
                .Item("ORDR_ADDR_TYPE_ST") = "DC"
                .Item("CUST_DC_NO") = rowARTCUST3.Item("CUST_ADDR_CODE2") & ""
            Else
                .Item("ORDR_ADDR_TYPE_ST") = "MK"
            End If

            .Item("ORDR_DATE_RECD") = DATETIME_STAMP.Date

            ' Sold To
            .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & ""
            .Item("SREP_CODE") = SREP_CODE
            .Item("SREP2_CODE") = SREP2_CODE
            .Item("ORDR_PRIORITY") = rowARTCUST1.Item("CUST_PRIORITY_CODE") & ""
            If .Item("ORDR_PRIORITY") & "" = "" Then
                .Item("ORDR_PRIORITY") = "9"
            End If

            .Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST

            .Item("FRT_TERMS") = rowARTCUST1.Item("FRT_TERMS") & ""
            .Item("ORDR_SHIP_INSTR") = rowARTCUST1.Item("CUST_ROUTING_INST") & ""
            .Item("ORDR_INV_COMMENT") = rowARTCUST1.Item("CUST_INV_COMMENT") & ""
            .Item("SHIP_VIA_CODE") = rowARTCUST1.Item("SHIP_VIA_CODE") & ""

            .Item("CURR_CODE") = "USD"
            .Item("CURR_EXCH_RATE") = 1

            ' Bill To
            If CUST_BILL_TO_CUST <> CUST_CODE Then
                rowARTCUST1_BT = LookUp("ARTCUST1", CUST_BILL_TO_CUST)
            Else
                rowARTCUST1_BT = rowARTCUST1
            End If

            .Item("POST_CODE") = rowARTCUST1_BT.Item("POST_CODE") & ""
            .Item("TERM_CODE") = rowARTCUST1_BT.Item("TERM_CODE") & ""
            .Item("CUST_FACTOR_IND") = rowARTCUST1_BT.Item("CUST_FACTOR_IND") & ""

            ' Store
            If rowARTCUST2 IsNot Nothing Then
                .Item("CUST_STORE_NAME") = rowARTCUST2.Item("CUST_NAME") & ""
            End If

        End With

        If Not preloadSalesOrder Then
            dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)
        End If

        If rowSOTORDR1.Item("ORDR_GROUP_NO") & "" = "" Then
            If COMPANY_CODE = "RGI" Or COMPANY_CODE = "RGI" Then
                ORDR_GROUP_NO = ORDR_NO
            Else
                If COMPANY_CODE = "VAN" Or COMPANY_CODE = "VAN" Then
                    ORDR_GROUP_NO = ASCMAIN1.Next_Control_No("ORDR_GROUP_NO")
                Else
                    ORDR_GROUP_NO = ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO")
                End If

            End If
            rowSOTORDR1.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
            ORDR_GROUP_NOs.Clear()
            ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
        End If

        Dim ORDR_LNO As Int16 = 0
        dst.Tables("SOTINVH2_XFR").Rows.Clear()
        dst.Tables("SOTORDR2_XFR").Rows.Clear()
        If XFR_INV_NOs.Count > 0 Then
            For Each INV_NO_XFR As String In XFR_INV_NOs
                Fill_Records("SOTINVH2_XFR", "", False, "SELECT * FROM SOTINVH2 WHERE INV_NO = '" & INV_NO_XFR & "'")
                Fill_Records("SOTINVHM", String.Empty, False, "SELECT * FROM SOTINVHM WHERE INV_NO = '" & INV_NO_XFR & "'")
                Fill_Records("SOTORDR2_XFR", String.Empty, False, "SELECT * FROM SOTORDR2 WHERE ORDR_NO IN (SELECT ORDR_NO FROM SOTINVH1 WHERE INV_NO = '" & INV_NO_XFR & "')")
            Next

            ' Set All Misc Charges to a temp Invoice Number
            Dim INV_MNO As Int16 = 1
            For Each rowSOTINVHM As DataRow In dst.Tables("SOTINVHM").Select("", "INV_TYPE, INV_NO, INV_MNO")
                rowSOTINVHM.Item("INV_NO") = "XXX"
                rowSOTINVHM.Item("INV_MNO") = INV_MNO
                INV_MNO += 1
            Next
            dst.Tables("SOTINVHM").AcceptChanges()
            For Each rowSOTINVHM As DataRow In dst.Tables("SOTINVHM").Select("")
                rowSOTINVHM.SetAdded()
            Next

            grdSOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop

            For Each row As DataRow In ASCDATA1.SelectDistinct("SOTINVH2_XFR", New String() {"STYLE_CODE", "COLOR_CODE"}).Rows
                Dim sql As String = "STYLE_CODE = '" & row.Item("STYLE_CODE") & "'" _
                                    & " AND COLOR_CODE = '" & row.Item("COLOR_CODE") & "'"

                Dim ORDR_QTY_SHIP As Int16 = Val(dst.Tables("SOTINVH2_XFR").Compute("SUM(ORDR_QTY_SHIP)", sql) & String.Empty)
                If ORDR_QTY_SHIP <= 0 Then ORDR_QTY_SHIP = 1

                Dim rowSOTORDR2_XFR As DataRow = dst.Tables("SOTORDR2_XFR").Select(sql)(0)

                addingRecord = True
                grdSOTORDR2.DisplayLayout.Bands(0).AddNew()
                addingRecord = False
                With grdSOTORDR2.ActiveRow
                    .Cells("STYLE_CODE").Value = row.Item("STYLE_CODE") & String.Empty
                    .Cells("COLOR_CODE").Value = row.Item("COLOR_CODE") & String.Empty

                    .Cells("STYLE_RETAIL").Value = rowSOTORDR2_XFR.Item("STYLE_RETAIL")
                    .Cells("CUST_UPC").Value = rowSOTORDR2_XFR.Item("CUST_UPC")
                    .Cells("CUST_SKU").Value = rowSOTORDR2_XFR.Item("CUST_SKU")
                    .Cells("CUST_STYLE_CODE").Value = rowSOTORDR2_XFR.Item("CUST_STYLE_CODE")
                    .Cells("CUST_COLOR_CODE").Value = rowSOTORDR2_XFR.Item("CUST_COLOR_CODE")
                    .Cells("CUST_SIZE_CODE").Value = rowSOTORDR2_XFR.Item("CUST_SIZE_CODE")

                    .Cells("ORDR_QTY").Value = ORDR_QTY_SHIP
                    .Update()
                End With
            Next
        End If

        ' not sure but there are blank records in SOTORDR2.
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ISNULL(ORDR_LNO, 0) = 0", "")
            rowSOTORDR2.Delete()
        Next

        dst.Tables("SOTORDR2").AcceptChanges()

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("", "")
            rowSOTORDR2.SetAdded()
            rowSOTORDR2.Item("ORDR_NO") = ORDR_NO
        Next

        grdSOTORDR2.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
        Sort_grdColumns(grdSOTINVHM, "INV_NO,INV_MNO")

        addingRecord = False
    End Sub

    Sub DisplayTotals()
        Dim ORDR_SALES As Decimal = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & "")
        Dim INV_MISC_CHG As Decimal = Val(dst.Tables("SOTINVHM").Compute("SUM(INV_MISC_CHG)", "") & String.Empty)
        Dim INV_FREIGHT As Decimal = Val(Absx1.numFor("INV_FREIGHT").Value & String.Empty)
        Dim ORDR_COSTS As Decimal = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT_COST)", "") & "")

        Absx1.numFor("ORDR_SALES").Value = ORDR_SALES
        Absx1.numFor("INV_MISC_CHG").Value = INV_MISC_CHG
        Absx1.numFor("SALES_AMOUNT").Value = ORDR_SALES + INV_FREIGHT + INV_MISC_CHG

        Absx1.numFor("ORDR_COSTS").Value = ORDR_COSTS
    End Sub

    Function Validate_Style(STYLE_CODE_z As String) As String
        Dim E As String = ""

        Dim STYLE_CODE As String = ""
        rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE_z)

        If rowICTSTYL1 Is Nothing Then
            E = "Style is Not on File" & vbCrLf
        Else
            If rowICTSTYL1.Item("STYLE_STATUS") & "" <> "A" Then
                ' E = "Style Status is not Active" & vbCrLf
            End If
            If rowICTSTYL1.Item("STYLE_UOM") & "" = "" Then
                E = "Style does not have a valid Unit of Measure" & vbCrLf
            End If
            If rowICTSTYL1.Item("SALES_DIVISION_CODE") & "" = "" Then
                E = "Style does not have a valid Division Code" & vbCrLf
            End If
        End If

        If E = "" Then
            COLOR_CODEs.Clear()
            Fill_Records("ICTCOLRS", STYLE_CODE_z)
            For Each row As DataRow In dst.Tables("ICTCOLRS").Select("")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                COLOR_CODEs.Add(COLOR_CODE)
            Next
        End If

        If E <> "" And grdSOTORDR2.ActiveRow IsNot Nothing AndAlso grdSOTORDR2.ActiveRow.IsAddRow Then
            MsgBox(E, MsgBoxStyle.OkOnly, "Style Code Entered is Invalid because ...")
        Else
            If E = "" Then
                STYLE_CODE = rowICTSTYL1.Item(0)
            End If
        End If
        Return STYLE_CODE
    End Function

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdSOTORDR2.ActiveRow

            If COLUMN_NAME = "STYLE_CODE" OrElse COLUMN_NAME = "COLOR_CODE" Then
                If grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Text = "" AndAlso grdSOTORDR2.ActiveRow.Cells("COLOR_CODE").Text = "" Then
                    Exit Sub
                End If
            End If

            Select Case COLUMN_NAME
                Case "STYLE_CODE"
                    If .Cells("STYLE_CODE").Text <> "" Then
                        Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Value & "")
                        Cancel = (STYLE_CODE = "")
                    End If
                Case "COLOR_CODE"
                    If .Cells("COLOR_CODE").Value & "" <> "" Then
                        If Not COLOR_CODEs.Contains(.Cells("COLOR_CODE").Value & "") Then
                            Cancel = True
                        End If
                    Else
                        Cancel = True
                    End If
                Case "ORDR_QTY"
                    If Trim(.Cells("STYLE_CODE").Value & "") = "" Then
                        Cancel = True
                        Exit Sub
                    End If
                    If Trim(.Cells("ORDR_QTY").Value & "") = "" Then
                        MsgBox("Order Qty Not Specified", vbOKOnly, "Cannot Update Record")
                        Cancel = True
                        grdSOTORDR2.ActiveCell = grdSOTORDR2.ActiveRow.Cells("ORDR_QTY")
                        Exit Sub
                    End If
                    If Val(.Cells("ORDR_QTY").Value & "") < 0 Then
                        MsgBox("Order Qty May Not be Negative", vbOKOnly, "Invalid Order Quantity")
                        Cancel = True
                    End If
            End Select
        End With
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

    Sub Record_Event(EVENT_TYPE As String, EVENT_DESC As String)
        Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
        rowTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
        rowTATEVNT1.Item("TABLE_KEY") = ORDR_NO
        rowTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
        rowTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowTATEVNT1.Item("EVENT_TYPE") = EVENT_TYPE
        rowTATEVNT1.Item("EVENT_DESC") = EVENT_DESC
        dst.Tables("TATEVNT1").Rows.Add(rowTATEVNT1)
    End Sub

    Sub Dependent_Updates(S As Integer, ORDR_NO As String)

        Dim QTY_TO_COMMIT As Int64
        Dim RSRV_QTY_OPEN_OLD As Int64
        Dim RSRV_QTY_OPEN_NEW As Int64

        ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow

        ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        For Each rowSOTORDR2 As DataRow In ASCDATA1.GetDataTable.Rows
            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")

            RSRV_QTY_OPEN_OLD = 0
            RSRV_QTY_OPEN_NEW = 0
            If S = -1 Then
                If rowSOTORDR2.Item("RSRV_NO") & "" <> "" Then ' restore_reservation Then
                    'Only restore this reservation line if it hasn't been substitutioned.  Per Gabe 07/30/02 - WR.
                    Dim row As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, rowSOTORDR2.Item("ORDR_LNO")})
                    If row IsNot Nothing Then  'Added for Angela. 1/24/05.  She was adding styles to range that had pulled from reservation already.
                        If row.Item("STYLE_CODE_SUB") & "" = "" Then
                            Update_SOTRSRVx(rowSOTORDR2, S)
                        End If
                    End If
                End If
            Else
                '  If ASCMAIN1.Running_in_VS Then Stop ' WON'T MULTIPLE RECORDS COME BACK, PERHAPS?
                Dim rowSOTRSRVX As DataRow = Fill_Record("SOTRSRVX", New String() {CUST_CODE, STYLE_CODE, COLOR_CODE})
                '& " order by SOTRSRV1.ORDR_CANCEL_DATE"

                If ASCMAIN1.CLIENT = "NYA" Then rowSOTRSRVX = Nothing
                ' NOTE - THERE IS A BUG WITH SHIPMENTS AND RESERVATIONS - QTY OPEN KEEPS GETTING MANGLED

                If rowSOTRSRVX IsNot Nothing Then
                    rowSOTORDR2.Item("RSRV_NO") = rowSOTRSRVX.Item("RSRV_NO")
                    rowSOTORDR2.Item("RSRV_LNO") = rowSOTRSRVX.Item("RSRV_LNO")

                    Update_SOTRSRVx(rowSOTORDR2, S)
                Else
                    rowSOTORDR2.Item("RSRV_NO") = DBNull.Value
                    rowSOTORDR2.Item("RSRV_LNO") = DBNull.Value
                End If
                Update_Record_TDA("SOTORDR2")
            End If

            QTY_TO_COMMIT = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
            QTY_TO_COMMIT = QTY_TO_COMMIT - S * (RSRV_QTY_OPEN_OLD - RSRV_QTY_OPEN_NEW)
            If QTY_TO_COMMIT <> 0 Then
                STYLE_CODE = rowSOTORDR2.Item("STYLE_CODE")
                COLOR_CODE = rowSOTORDR2.Item("COLOR_CODE")
                TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", S * QTY_TO_COMMIT)
            End If
        Next

    End Sub

    Sub Update_SOTRSRVx(rowSOTORDR2 As DataRow, S As Integer)
        Dim RSRV_NO As String = rowSOTORDR2.Item("RSRV_NO")
        Dim RSRV_LNO As Int64 = Val(rowSOTORDR2.Item("RSRV_LNO") & "")
        Dim rowSOTRSRV2 As DataRow = Fill_Record("SOTRSRV2", New String() {RSRV_NO, RSRV_LNO})
        With rowSOTRSRV2
            Dim RSRV_QTY As Int64 = .Item("RSRV_QTY")
            Dim RSRV_QTY_OPEN As Int64 = Val(.Item("RSRV_QTY_OPEN") & "")
            Dim RSRV_QTY_CANC As Int64 = Val(.Item("RSRV_QTY_CANC") & "")
            Dim RSRV_QTY_USED As Int64 = Val(.Item("RSRV_QTY_USED") & "") _
                          + S * Val(rowSOTORDR2.Item("ORDR_QTY_ORIG") & "")
            Dim RSRV_QTY_OPEN_OLD As Int64 = RSRV_QTY_OPEN
            RSRV_QTY_OPEN = RSRV_QTY - RSRV_QTY_CANC - RSRV_QTY_USED
            If RSRV_QTY_OPEN < 0 Then
                RSRV_QTY_OPEN = 0
            End If
            Dim RSRV_QTY_OPEN_NEW As Int64 = RSRV_QTY_OPEN
            .Item("RSRV_QTY_USED") = RSRV_QTY_USED
            .Item("RSRV_QTY_OPEN") = RSRV_QTY_OPEN
        End With

        Dim rowSOTRSRV1 As DataRow = Fill_Record("SOTRSRV1", RSRV_NO)

        If S = -1 Then
        Else

            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
            Dim rowSOTORDR7 As DataRow = Fill_Record("SOTORDR7", New String() {ORDR_GROUP_NO, STYLE_CODE, COLOR_CODE})

            If rowSOTORDR7 Is Nothing Then
                rowSOTORDR7 = dst.Tables("SOTORDR7").NewRow
                rowSOTORDR7.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                rowSOTORDR7.Item("STYLE_CODE") = STYLE_CODE
                rowSOTORDR7.Item("COLOR_CODE") = COLOR_CODE
                dst.Tables("SOTORDR7").Rows.Add(rowSOTORDR7)
            End If
            If rowSOTRSRV2.Item("RSRV_PRIORITY_DATE") & "" = "" Then
                rowSOTORDR7.Item("ORDR_PRIORITY_DATE") = CDate(rowSOTRSRV1.Item("INIT_DATE")).Date ' DateValue(Format(rowSOTRSRV1.Item("INIT_DATE"), "MM/dd/yyyy"))
            Else
                rowSOTORDR7.Item("ORDR_PRIORITY_DATE") = CDate(rowSOTRSRV2.Item("RSRV_PRIORITY_DATE")).Date '  DateValue(Format$(rowSOTRSRV2.Item("RSRV_PRIORITY_DATE"), "MM/DD/YYYY"))
            End If
            rowSOTORDR7.Item("ORDR_PRIORITY") = rowSOTRSRV2.Item("RSRV_PRIORITY")
            Update_Record_TDA("SOTORDR7", "ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
        End If
        Update_Record_TDA("SOTRSRV2")

        ASCMAIN1.sql = "Select Sum (RSRV_QTY_OPEN) from SOTRSRV2 where RSRV_NO = :PARM1"
        Dim RSRV_QTY_OPEN_total As Int64 = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {RSRV_NO}))

        If RSRV_QTY_OPEN_total = 0 Then
            rowSOTRSRV1.Item("RSRV_STATUS") = "F"
        Else
            rowSOTRSRV1.Item("RSRV_STATUS") = "O"
        End If
        Update_Record_TDA("SOTRSRV1")
    End Sub

    Private Sub CreateShipmentRecord()

        Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
        Dim SHIP_CNT_CARTONS As Int64 = 1
        Dim SHIP_TOTAL_WGT As Decimal = 1

        Dim ORDR_PICK_SEQ As Integer = Val(rowSOTORDR1.Item("ORDR_PICK_SEQ") & "") + 1

        Dim PICK_NO As String = ASCMAIN1.Next_Control_No("SOTPICK1.PICK_NO")
        Dim SHIP_BOL_NO As String = ASCMAIN1.Next_Control_No("SOTSHIP1.SHIP_BOL_NO")

        Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").NewRow
        With rowSOTSHIP1
            .Item("SHIP_BOL_NO") = SHIP_BOL_NO
            .Item("SHIP_DATE_SHIPPED") = DateTime.Now.ToShortDateString
            .Item("SHIP_VIA_CODE") = rowSOTORDR1.Item("SHIP_VIA_CODE")
            .Item("SHIP_REF") = ""
            .Item("SHIP_TOTAL_WGT") = SHIP_TOTAL_WGT
            .Item("SHIP_CNT_CARTONS") = SHIP_CNT_CARTONS
            .Item("SHIP_ADDR_TYPE") = rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST")
            .Item("SHIP_ADDR_CODE") = IIf(rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST") = "MK", rowSOTORDR1.Item("CUST_STORE_NO"), rowSOTORDR1.Item("CUST_DC_NO"))
            .Item("ORDR_GROUP_NO") = rowSOTORDR1.Item("ORDR_GROUP_NO")
            .Item("SHIP_PICK_PRINTED") = DATETIME_STAMP
            .Item("PICK_BATCH_NO") = "000000"
            .Item("SHIP_STATUS") = "P"
            .Item("FRT_TERMS") = rowSOTORDR1.Item("FRT_TERMS")
            .Item("WHSE_CODE") = rowSOTORDR1.Item("WHSE_CODE")
            .Item("INV_DATE") = DateTime.Now.ToShortDateString
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("BILL_OF_LADING_NO") = "" ' rowPOTSHIP1.Item("PO_SHIP_REF_NO")
            .Item("TERM_CODE") = rowSOTORDR1.Item("TERM_CODE")
            .Item("SREP_CODE") = rowSOTORDR1.Item("SREP_CODE")
            .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT")
            .Item("SHIP_NOTES") = "" ' rowPOTSHIP1.Item("PO_NOTES")
            .Item("SHIPPED_ACTUAL") = .Item("SHIP_DATE_SHIPPED")
            .Item("CUST_FACTOR_TRANS_IND") = rowSOTORDR1.Item("CUST_FACTOR_IND")
            .Item("SREP2_CODE") = rowSOTORDR1.Item("SREP2_CODE")
            .Item("SHIP_SPEC_INST") = rowSOTORDR1.Item("ORDR_SHIP_INSTR")
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("SHIP_856_IND") = "0"
            .Item("SHIP_810_IND") = "0"
            .Item("CUST_FACTOR_TRANS_IND") = "O"

            Dim rowEDTSLSP1 = LookUp("EDTSLSP1", Absx1.txtFor("CUST_CODE").Text)
            If rowEDTSLSP1 IsNot Nothing Then
                If rowEDTSLSP1.Item("EDI_ID_856") & "" <> "" Then .Item("SHIP_856_IND") = "1"
                If rowEDTSLSP1.Item("EDI_ID_810") & "" <> "" Then .Item("SHIP_810_IND") = "1"
            End If

            If chkCUST_FACTOR_IND.Checked Then
                .Item("CUST_FACTOR_TRANS_IND") = "1"
            End If
 
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = DATETIME_STAMP
        End With

        dst.Tables("SOTSHIP1").Rows.Add(rowSOTSHIP1)
    End Sub

    Sub CreatePickTicketsAndShipmentRecords()

        Dim freight As Double = MyBase.Absx1.numFor("INV_FREIGHT").Value

        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
            Dim SHIP_CNT_CARTONS As Int64 = 1
            Dim SHIP_TOTAL_WGT As Decimal = 1

            Dim ORDR_PICK_SEQ As Integer = Val(rowSOTORDR1.Item("ORDR_PICK_SEQ") & "") + 1

            Dim PICK_NO As String = ASCMAIN1.Next_Control_No("SOTPICK1.PICK_NO")
            Dim SHIP_BOL_NO As String = ASCMAIN1.Next_Control_No("SOTSHIP1.SHIP_BOL_NO")

            rowSOTORDR1.Item("ORDR_STATUS") = "P"

            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").NewRow
            With rowSOTPICK1
                .Item("SELECTED") = "1"
                .Item("PICK_NO") = PICK_NO
                .Item("ORDR_NO") = ORDR_NO
                .Item("PICK_FREIGHT") = freight
                freight = 0
                .Item("PICK_PICKER") = ASCMAIN1.USER_ID
                .Item("ORDR_PICK_SEQ") = ORDR_PICK_SEQ
                .Item("PICK_STATUS") = "P"
                .Item("PICK_RELEASED") = DATETIME_STAMP
                .Item("PICK_PRINTED") = DATETIME_STAMP
                .Item("PICK_PACKED") = DATETIME_STAMP
                .Item("PICK_SHIPPED") = DATETIME_STAMP
                .Item("PICK_BATCH_NO") = "000000"
                .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                .Item("INV_NO") = String.Empty
                .Item("PICK_CNT_CARTONS") = SHIP_CNT_CARTONS
                .Item("PICK_TOTAL_WGT") = SHIP_TOTAL_WGT
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("ORDR_INV_COMMENT") = rowSOTORDR1.Item("ORDR_INV_COMMENT")
            End With
            dst.Tables("SOTPICK1").Rows.Add(rowSOTPICK1)

            Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").NewRow
            With rowSOTSHIP1
                .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                .Item("SHIP_DATE_SHIPPED") = DateTime.Now.ToShortDateString
                .Item("SHIP_VIA_CODE") = rowSOTORDR1.Item("SHIP_VIA_CODE")
                .Item("SHIP_REF") = ""
                .Item("SHIP_TOTAL_WGT") = SHIP_TOTAL_WGT
                .Item("SHIP_CNT_CARTONS") = SHIP_CNT_CARTONS
                .Item("SHIP_ADDR_TYPE") = rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST")
                .Item("SHIP_ADDR_CODE") = IIf(rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST") = "MK", rowSOTORDR1.Item("CUST_STORE_NO"), rowSOTORDR1.Item("CUST_DC_NO"))
                .Item("ORDR_GROUP_NO") = rowSOTORDR1.Item("ORDR_GROUP_NO")
                .Item("SHIP_PICK_PRINTED") = DATETIME_STAMP
                .Item("PICK_BATCH_NO") = "000000"
                .Item("SHIP_STATUS") = "P"
                .Item("FRT_TERMS") = rowSOTORDR1.Item("FRT_TERMS")
                .Item("WHSE_CODE") = rowSOTORDR1.Item("WHSE_CODE")
                .Item("INV_DATE") = DateTime.Now.ToShortDateString
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("BILL_OF_LADING_NO") = "" ' rowPOTSHIP1.Item("PO_SHIP_REF_NO")
                .Item("TERM_CODE") = rowSOTORDR1.Item("TERM_CODE")
                .Item("SREP_CODE") = rowSOTORDR1.Item("SREP_CODE")
                .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT")
                .Item("SHIP_NOTES") = "" ' rowPOTSHIP1.Item("PO_NOTES")
                .Item("SHIPPED_ACTUAL") = .Item("SHIP_DATE_SHIPPED")
                .Item("CUST_FACTOR_TRANS_IND") = rowSOTORDR1.Item("CUST_FACTOR_IND")
                .Item("SREP2_CODE") = rowSOTORDR1.Item("SREP2_CODE")
                .Item("SHIP_SPEC_INST") = rowSOTORDR1.Item("ORDR_SHIP_INSTR")
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP

                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_OPER") = DATETIME_STAMP
            End With
            dst.Tables("SOTSHIP1").Rows.Add(rowSOTSHIP1)

            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("", "", DataViewRowState.CurrentRows)
                Dim ORDR_LNO As Integer = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
                Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)

                rowSOTORDR2.Item("ORDR_QTY_PICK") = rowSOTORDR2.Item("ORDR_QTY")
                rowSOTORDR2.Item("ORDR_QTY_OPEN") = 0
                rowSOTORDR2.Item("ORDR_STATUS") = "P"
                'rowSOTORDR2.Item("ORDR_QTY_ALLO") = rowSOTORDR2.Item("ORDR_QTY")

                Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").NewRow
                With rowSOTPICK2
                    .Item("PICK_NO") = PICK_NO
                    .Item("PICK_LNO") = ORDR_LNO
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("ORDR_LNO") = ORDR_LNO
                    .Item("PICK_QTY") = rowSOTORDR2.Item("ORDR_QTY")
                    .Item("PICK_QTY_CONF") = rowSOTORDR2.Item("ORDR_QTY")
                    .Item("PICK_QTY_CANC") = 0
                    .Item("PICK_QTY_BACK") = 0
                    .Item("PICK_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                    .Item("PICK_QTY_CANC_REL") = 0
                    .Item("PICK_QTY_BACK_REL") = 0
                End With
                dst.Tables("SOTPICK2").Rows.Add(rowSOTPICK2)
            Next

            rowSOTORDR1.Item("ORDR_PICK_SEQ") = ORDR_PICK_SEQ
        Next
    End Sub

    Private Sub Load_SOTINVHX()
        Dim RYP0 As String = Absx1.cmbFor("RYP0").SelectedRow.Cells("OPS_YYYYPP").Value
        Dim RYP1 As String = Absx1.cmbFor("RYP1").SelectedRow.Cells("OPS_YYYYPP").Value

        ASCMAIN1.sql = "SELECT * FROM SOTINVH1 WHERE ORDR_TYPE_CODE = 'XFR' AND INV_TYPE = 'I' AND INV_TOTAL_AMOUNT = 0 AND ORDR_YYYYPP_UPDATED BETWEEN '" & RYP0 & "' AND '" & RYP1 & "'"
        Fill_Records("SOTINVHX", String.Empty, True, ASCMAIN1.sql)
        Sort_grdColumns(grdSOTINVH1X, "INV_NO".ToLower)
        grdSOTINVH1X.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
    End Sub

#End Region

End Class