Public Class SOFSHIPC

    ' proceed prereq - after maintenance, or confirmation,need to verify that auth amt on credit check and credit card has not been violated
    Dim expSOTPICK1 As New Dictionary(Of String, String)

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
    Dim MaintenanceMode As Boolean = False

    Dim ORDR_SHIP_DATE As Date
    Dim ORDR_CANCEL_DATE As Date

    Dim dvwSOTORDR5 As DataView

    Dim SOTSHIPX As String

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

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

        Get_PARM("SOTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            sqlSOTSHIPX = "Select SOTSHIP1.*" _
                & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE,SOTORDR0.ORDR_CUST_PO" _
                & " from SOTSHIP1,SOTORDR0" _
                & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf
            If Not InquiryMode Then
                sqlSOTSHIPX &= "" _
                    & "   and SOTSHIP1.SHIP_STATUS = 'P'" _
                    & "   and SOTSHIP1.SHIP_PICK_PRINTED is Not Null"
            End If
            ASCMAIN1.sql = sqlSOTSHIPX
            Create_TDA(.Tables.Add, "SOTSHIPX", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "SOTORDR2", "*")

            ASCMAIN1.sql = sqlSOTSHIPX & " and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, True, "", 1)
            ASCMAIN1.sql = sqlSOTSHIPX & " and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTSHIP0", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "SOTINVH1", "*")
            Create_TDA(.Tables.Add, "SOTINVH2", "*")
            Create_TDA(.Tables.Add, "SOTINVH9", "*")

            Create_TDA(.Tables.Add, "SOTCART1", "*")

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

            ASCMAIN1.sql = "Select CUST_CODE, EDI_DOC_NO, EDI_STATUS" _
                & " from EDTTRPM1" _
                & " where CUST_CODE = :PARM1"
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
                & ", SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                & ", SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.CUST_BILL_TO_CUST" & vbCrLf _
                & ", SOTORDR1.POST_CODE, SOTORDR1.WHSE_CODE" & vbCrLf _
                & ", SOTORDR1.TERM_CODE, SOTORDR1.SREP_CODE, SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
                & ", SOTSHIP1.BILL_OF_LADING_NO, SOTORDR1.ORDR_INV_COMMENT, SOTORDR1.CUST_FACTOR_IND" & vbCrLf _
                & " from SOTPICK1,SOTORDR1,SOTSHIP1 "
            ASCMAIN1.sql = sqlSOTPICK1 & " where ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**")
            dst.Tables("SOTPICK1").Columns.Add("SELECTED")

            Create_Relation("SOTSHIP1", "SOTPICK1", "SHIP_BOL_NO")

            sqlSOTPICK2 = "Select SOTPICK2.*, " & vbCrLf _
                & " SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_DESC, " & vbCrLf _
                & " SOTORDR2.ORDR_UNIT_PRICE, SOTORDR2.STYLE_CODE_SUB," & vbCrLf _
                & " SOTORDR2.RANGE_STYLE_CODE, SOTORDR2.RANGE_STYLE_LNO, SOTORDR2.QTY_PER_PP" & vbCrLf _
                & " from SOTPICK2,SOTPICK1,SOTORDR2,SOTSHIP1" & vbCrLf
            ASCMAIN1.sql = sqlSOTPICK2 & " where ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "SOTPICK2", "**")
            With .Tables("SOTPICK2").Columns
                '.Add("PICK_AMT", GetType(System.Decimal), "ISNULL(PICK_UNIT_PRICE,0) * ISNULL(PICK_QTY,0)")
                '.Add("PICK_AMT_CONF", GetType(System.Decimal), "ISNULL(PICK_UNIT_PRICE,0) * ISNULL(PICK_QTY_CONF,0)")
                '.Add("PICK_AMT_CANC", GetType(System.Decimal), "ISNULL(PICK_UNIT_PRICE,0) * ISNULL(PICK_QTY_CANC,0)")
                '.Add("PICK_AMT_BACK", GetType(System.Decimal), "ISNULL(PICK_UNIT_PRICE,0) * ISNULL(PICK_QTY_BACK,0)")
                .Add("PICK_AMT", GetType(System.Decimal))
                .Add("PICK_AMT_CONF", GetType(System.Decimal))
                .Add("PICK_AMT_CANC", GetType(System.Decimal))
                .Add("PICK_AMT_BACK", GetType(System.Decimal))
            End With

            Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")
            'With .Tables("SOTPICK1").Columns
            '    .Add("PICK_QTY", GetType(System.Int64), "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY)")
            '    .Add("PICK_QTY_CONF", GetType(System.Int64), "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY_CONF)")
            '    .Add("PICK_QTY_CANC", GetType(System.Int64), "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY_CANC)")
            '    .Add("PICK_QTY_BACK", GetType(System.Int64), "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY_BACK)")
            '    .Add("PICK_AMT", GetType(System.Int64), "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_AMT)")
            '    .Add("PICK_AMT_CONF", GetType(System.Int64), "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_AMT_CONF)")
            '    .Add("PICK_AMT_CANC", GetType(System.Int64), "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_AMT_CANC)")
            '    .Add("PICK_AMT_BACK", GetType(System.Int64), "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_AMT_BACK)")
            'End With
            With .Tables("SOTPICK1").Columns
                .Add("PICK_QTY", GetType(System.Int64))
                .Add("PICK_QTY_CONF", GetType(System.Int64))
                .Add("PICK_QTY_CANC", GetType(System.Int64))
                .Add("PICK_QTY_BACK", GetType(System.Int64))
                .Add("PICK_AMT", GetType(System.Decimal))
                .Add("PICK_AMT_CONF", GetType(System.Decimal))
                .Add("PICK_AMT_CANC", GetType(System.Decimal))
                .Add("PICK_AMT_BACK", GetType(System.Decimal))
            End With

            ASCMAIN1.sql = "Select ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE" & vbCrLf _
                & ", ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP" & vbCrLf _
                & ", ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST" & vbCrLf _
                & ", 'N' CONSUMED, 0.00 CUM_QTY" & vbCrLf _
                & " FROM ICTCOST1" & vbCrLf _
                & " WHERE ROWNUM < 0" & vbCrLf
            Create_TDA(.Tables.Add, "ICTCOST1", "**", 0, False)

            Create_Relation("SOTCART1", "SOTCART2", "CART_NO")
            'dst.Tables("SOTCART2").Columns.Add("PICK_NO", GetType(System.String), "PARENT(SOTCART1_SOTCART2).PICK_NO")

            'dst.Tables("SOTCART2").Columns.Add("PICK_NO", GetType(System.String))
            Create_Relation("SOTCARTX", "SOTPICK2", "PICK_NO,ORDR_NO,ORDR_LNO")
            'dst.Tables("SOTCARTX").Columns("PICK_QTY_CONF").Expression = "SUM(CHILD(SOTCARTX_SOTPICK2).PICK_QTY_CONF)"
            Create_Relation("SOTCARTX", "SOTCART2", "PICK_NO,ORDR_NO,ORDR_LNO")
            'dst.Tables("SOTCARTX").Columns("QTY_PACKED").Expression = "SUM(CHILD(SOTCARTX_SOTCART2).QTY_PACKED)"

            dst.Tables("SOTCART1").Columns.Add("CART_TOTAL_UNITS_CALC", GetType(System.Int64))
            dst.Tables("SOTCART1").Columns.Add("CART_TOTAL_UNITS_ORIG", GetType(System.Int64), "SUM(CHILD.QTY_PACKED_ORIG)")

            Create_Relation("SOTPICK1", "SOTCART1", "PICK_NO")
            'dst.Tables("SOTPICK1").Columns.Add("PICK_TOTAL_WGT_CALC", GetType(System.Decimal), "SUM(CHILD(SOTPICK1_SOTCART1).CART_TOTAL_WGT_ACTUAL)")
            'dst.Tables("SOTPICK1").Columns.Add("PICK_CNT_CARTONS_CALC", GetType(System.Int64), "COUNT(CHILD(SOTPICK1_SOTCART1).CART_NO)")
            'dst.Tables("SOTPICK1").Columns.Add("PICK_TOTAL_UNITS_CALC", GetType(System.Int64), "SUM(CHILD(SOTPICK1_SOTCART1).CART_TOTAL_UNITS_CALC)")
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

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                For Each TABLE_NAME As String In New String() _
                    {"SOTSHIP1_3PL", "SOTPICK1_3PL", "SOTPICK2_3PL", "SOTCART1_3PL", "SOTCART2_3PL", "SOTCART3_3PL"}
                    ASW.Add(TABLE_NAME, ASCMAIN1.Temp_Table("Select * from " & TABLE_NAME & " where ROWNUM <1"))
                Next

                ASCMAIN1.sql = "Select SOTSHIP1_3PL.*" _
                    & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_CUST_PO" _
                    & " from " & ASW("SOTSHIP1_3PL") & " SOTSHIP1_3PL, SOTORDR0" _
                    & " where SOTORDR0.ORDR_GROUP_NO (+) = SOTSHIP1_3PL.ORDR_GROUP_NO"
                Create_TDA(.Tables.Add, "WHT3PLS1", "**", 0, False, , 1)
                .Tables("WHT3PLS1").Columns("CUST_CODE").AllowDBNull = True
            End If


            Create_TDA(.Tables.Add, "SOTSHIP3", "*")
            Create_TDA(.Tables.Add, "SOTSHIP4", "*")
            Create_TDA(.Tables.Add, "SOTSHIP6", "*")

            Create_TDA(.Tables.Add, "SOTORDR5", "*", 1)
            .Tables("SOTORDR5").Columns("CUST_ADDR_CODE").MaxLength = 10

            Create_TDA(.Tables.Add, "SOTORDRE", "*")
            Create_TDA(.Tables.Add, "SOTORDXR", "*")
        End With

        grdSOTSHIPX.DataSource = dst.Tables("SOTSHIPX")
        grdSOTCONFT.DataSource = dst.Tables("SOTCONFT")
        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTPICK2.DataSource = dst.Tables("SOTPICK2")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")
        grdSOTCART2.DataSource = dst.Tables("SOTCART2")
        grdSOTSHIP1.DataSource = dst.Tables("SOTSHIP1")
        grdWHT3PLS1.DataSource = dst.Tables("WHT3PLS1")

        dvwSOTORDR5 = New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'ST'", "", DataViewRowState.CurrentRows)
        Bind_Controls(grpSHIPTO, "SOTORDR5", dvwSOTORDR5)

        grdSOTSHIPX.DisplayLayout.UseFixedHeaders = True
        With grdSOTSHIPX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "CUST_CODE", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            grdWHT3PLS1.DisplayLayout.UseFixedHeaders = True
            With grdWHT3PLS1.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "CUST_CODE", "ORDR_CUST_PO"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                Next
            End With
        End If


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
                If Not MaintenanceMode And New String() {"SELECTED", "PICK_CNT_CARTONS", "PICK_FREIGHT", "PICK_TOTAL_WGT", "ORDR_INV_COMMENT"}.Contains(gcol.Key) Then
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
                If New String() {"CART_FREIGHT", "CART_TOTAL_WGT_ACTUAL", "CART_TRACKING_NO"}.Contains(gcol.Key) Then
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

        Create_Summary(grdSOTSHIPX, "SHIP_BOL_NO", "Count")
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            Create_Summary(grdWHT3PLS1, "SHIP_BOL_NO", "Count")
        End If

        'Create_Summary(grdSOTPICK1, "PICK_NO", "Count")
        'Create_Summary(grdSOTPICK1, New String() _
        '    {"SELECTED", "PICK_CNT_CARTONS", "PICK_TOTAL_WGT", "PICK_FREIGHT"})

        Create_Summary(grdSOTPICK2, "PICK_LNO", "Count")
        Create_Summary(grdSOTPICK2, New String() _
            {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"})

        Create_Summary(grdSOTCART1, "CART_NO", "Count")
        Create_Summary(grdSOTCART1, New String() _
            {"CART_FREIGHT", "CART_TOTAL_UNITS", "CART_TOTAL_WGT_ACTUAL"})

        With dst.Tables("SOTCONFT").Rows
            .Add(New Object() {1, "Reld", 0, 0})
            .Add(New Object() {2, IIf(MaintenanceMode, "Revd", "Conf"), 0, 0})
            .Add(New Object() {3, "Canc", 0, 0})
            .Add(New Object() {4, "Back", 0, 0})
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
        If InquiryMode Then
            tabSelect.Tabs("3PL Shipments").Visible = False
        End If
    End Sub

    Sub Check_InquiryMode()
        InquiryMode = (MENU_ITEM_OBJECT = "SOFSHIPI")
        MaintenanceMode = (MENU_ITEM_OBJECT = "SOFSHIPM")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select"

                ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIP0)
                SHIP_BOL_NOs.Clear()

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
                        ORDR_GROUP_NO = rowSOTSHIP1.Item("ORDR_GROUP_NO")
                        If Not InquiryMode Then If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
                        Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)
                        CUST_CODE = rowSOTORDR0.Item("CUST_CODE")
                        ORDR_CUST_PO = rowSOTORDR0.Item("ORDR_CUST_PO") & ""

                        'If select_from_3PL_list Then
                        '    If ASCMAIN1.Running_in_VS Then Stop ' GET OTHER SHIPMENTS READY TO GO FROM 3PL IN SAME GROUP
                        'Else

                        If optShipmentSelection.Value = "S" Then
                            If Not InquiryMode Then If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub
                            SHIP_BOL_NOs.Add(SHIP_BOL_NO)
                            ASCDATA1.ExecuteSQL("Insert into " & SOTSHIP0 & " (SHIP_BOL_NO) values ('" & SHIP_BOL_NO & "')")
                        Else
                            ASCMAIN1.sql = "Select SHIP_BOL_NO from SOTSHIP1 " _
                            & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
                            & "   and SHIP_STATUS = '" & SHIP_STATUS & "'"
                            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                                Dim SHIP_BOL_NO2 As String = row.Item("SHIP_BOL_NO")
                                If select_from_3PL_list Then

                                    Dim rowWHT3PLS1 As DataRow = dst.Tables("WHT3PLS1").Rows.Find(SHIP_BOL_NO2)
                                    If rowWHT3PLS1 Is Nothing Then
                                        EMsg &= vbCr & "Cannot also bill Shipment " & SHIP_BOL_NO2 & " from same Group"
                                        Exit For
                                    End If
                                End If
                                If Not InquiryMode Then If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO2) Then Exit Sub
                                Me.SHIP_BOL_NOs.Add(SHIP_BOL_NO2)
                                ASCDATA1.ExecuteSQL("Insert into " & SOTSHIP0 & " (SHIP_BOL_NO) values ('" & SHIP_BOL_NO2 & "')")
                            Next
                        End If
                        'End If
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
                                If rowSOTSHIP1.Item("SHIP_PICK_PRINTED") & "" = "" Then
                                    EMsg &= vbCr & "Pick Tickets have not been Printed (yet) for Shipment " & SHIP_BOL_NO
                                End If
                                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                                    Dim rowSOTCTLU1 As DataRow = LookUp("SOTCTLU1", "Z")
                                    If rowSOTCTLU1.Item("CTL_UPDATE_REQ") & "" = "D" Then
                                        EMsg &= vbCr & "There Has Been A De-Confirm that has not been updated by the Sales Journal." _
                                            & "Please Run Sales Journal Before Proceeding"
                                    End If
                                End If
                            Else
                                Select Case SHIP_STATUS
                                    Case "F"
                                        Dim rowSOTCTLU1 As DataRow = LookUp("SOTCTLU1", "Z")
                                        If rowSOTCTLU1.Item("CTL_UPDATE_REQ") & "" = "C" Then
                                            MsgBox("There Has Been A Confirm that has not been updated by the Sales Journal." _
                                                   & "Please Run Sales Journal Before Proceeding", _
                                                   MsgBoxStyle.OkOnly, "Sales Journal Update Required First")
                                            Exit Sub
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
                                                'Stop ' SEE WJZ FOR TESTING
                                                De_Confirm(SHIP_BOL_NO)
                                            End If
                                            Exit Sub
                                        Else
                                            If (rowSOTSHIP1.Item("SHIP_810_BATCH_NO") & "" <> "N" _
                                            And rowSOTSHIP1.Item("SHIP_810_BATCH_NO") & "" <> "") _
                                            Or MaintenanceMode Then
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
                                                        'Stop ' SEE WJZ FOR TESTING
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


                If EMsg = "" And Not MaintenanceMode And Not InquiryMode Then
                    Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NOs(0))
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", rowSOTSHIP1.Item("WHSE_CODE"))
                    If Not select_from_3PL_list Then
                        '  rowICTWHSE1.Item("LP_CODE") = "" ' TEMP TO TEST ENTRY
                        If rowICTWHSE1.Item("LP_CODE") & "" <> "" Then
                            MsgBox("You must select this Shipment from the 3PL Shipments tab" _
                                   & vbCrLf & " in order to be able to Confirm this Shipment", _
                                   vbOKOnly, "Selected Shipment is associated with a 3PL Warehouse")
                            ' TEMPORARY ALLOW UNTIL NS CLEANS UP CANCELLED PTS
                            'Exit Sub
                        End If
                    Else
                        If rowICTWHSE1.Item("LP_CODE") & "" = "" Then
                            MsgBox("The Warehouse listed on Shipment" _
                                   & vbCrLf & " is NOT set up as a 3PL Warehouse", _
                                   vbOKOnly, "Selected Shipment is NOT associated with a 3PL Warehouse")
                            Exit Sub
                        End If
                    End If
                End If

                If EMsg <> "" Then
                    ASCMAIN1.MultiTask_Release()
                End If

            Case "Update"

                If Not MaintenanceMode And ASCMAIN1.DBS_COMPANY = "VAN" Then
                    EMsg &= vbCr & "naughty girl"
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

                ' although this only matters for edi customers, I think we should enforce the integrity
                Dim rowSOTCARTX_oobal As DataRow() = dst.Tables("SOTCARTX").Select("ISNULL(PICK_QTY_CONF,0) <> ISNULL(QTY_PACKED,0)")
                If rowSOTCARTX_oobal.Length <> 0 Then
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
                    If txtReason.Text = "" Or txtContact.Text = "" Then
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
                        Or Absx1.dteFor("INV_DATE").Value & "" = "" Then
                        EMsg &= vbCr & "Date Shipped and Invoice Date are Mandatory"
                    Else
                        If Format(Absx1.dteFor("SHIP_DATE_SHIPPED").Value, "yyyyMMdd") _
                         > Format(Absx1.dteFor("INV_DATE").Value, "yyyyMMdd") Then
                            EMsg &= vbCr & "Invoice Date cannot be Prior to Date Shipped"
                        End If

                        If Format(Absx1.dteFor("INV_DATE").Value, "yyyyMM") <> ASCMAIN1.CYM Then
                            If Format(Absx1.dteFor("INV_DATE").Value, "yyyyMM") = ASCMAIN1.Period_Calc(ASCMAIN1.CYM, 1) Then
                                If MsgBox("You are about to confirm a shippment" _
                                          & vbCrLf & "that will be posted into the Next period.", _
                                           MsgBoxStyle.OkCancel, _
                                           "Invoice Date Confirmation") <> MsgBoxResult.Ok Then
                                    Exit Sub
                                End If
                            Else
                                EMsg &= vbCr & "Invoice Date Not in Current Period"
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
                        End If
                    End If
                    If Absx1.txtFor("FRT_TERMS").Text = "" Then
                        EMsg &= vbCr & "Frt Terms Code is Required"
                    Else
                        If LookUp("ASTCODE1", New String() {"SOTORDR1", "FRT_TERMS", Absx1.txtFor("FRT_TERMS").Text}) Is Nothing Then
                            EMsg &= vbCr & "Invalid Frt Terms Code"
                        End If
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
                            Or Absx1.txtFor("SHIP_REF").Text = "" Then
                            EMsg &= vbCr & "Ship Via Code and Shippers Reference (Pro #) are Required"
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

                    Dim sqlw As String = ""

                    If Absx1.txtFor("FRT_TERMS").Text <> "" Then
                        If Absx1.txtFor("FRT_TERMS").Text <> "PPA" Then
                            sqlw = "PICK_FREIGHT <> 0 and SELECTED = '1'"
                        Else
                            sqlw = "PICK_FREIGHT = 0 and SELECTED = '1'"
                        End If
                        If dst.Tables("SOTPICK1").Select(sqlw).Length > 0 Then
                            If Absx1.txtFor("FRT_TERMS").Text <> "PPA" Then
                                EMsg &= vbCr & "Freight Terms Code Specified does not permit Non-Zero Freight Amounts"
                            Else
                                EMsg &= vbCr & "Freight Terms Code Specified does not permit Zero Freight Amounts"
                            End If
                        End If
                    End If

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
                            If RANGE_STYLE_QTY_PER_PP = 0 Or _
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

                            If RANGE_STYLE_LNO_last = 0 Or RANGE_STYLE_LNO_last <> RANGE_STYLE_LNO Then
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
                        Or (Not chkFactored.Checked And rowARTCUST1.Item("CUST_FACTOR_IND") & "" = "1") Then
                            If MsgBox("Factor Option is not in synch with Customer Master" & vbCrLf & "Continue Anyway", _
                                       MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If ASCMAIN1.USER_SECURITY_CODEs.Contains("BD") Then
                    Else
                        For Each row As DataRow In ASCDATA1.SelectDistinct _
                            (dst.Tables("SOTPICK2").Select("PICK_QTY_CONF <> 0"), New String() {"STYLE_CODE", "COLOR_CODE"}).Rows
                            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                            'Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                            'Dim STYLE_COST As Decimal = Val(rowICTSTYL1.Item("STYLE_COST") & "")
                            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
                            Dim STYLE_COST As Decimal = Val(rowICTSTYC1.Item("STYLE_COST_FIFO") & "")
                            If STYLE_COST = 0 Or STYLE_COST < 0.01 Then
                                If ASCMAIN1.DBS_SERVER = "" Or ASCMAIN1.DBS_COMPANY = "TST" Then
                                    ' ALLOW FOR NOW
                                Else
                                    EMsg &= vbCr & "Style " & STYLE_CODE & " has a Valuation Cost of " & Format(STYLE_COST, "###.00") & "."

                                End If
                            End If
                        Next
                    End If

                    If dst.Tables("SOTPICK2").Select("ORDR_UNIT_PRICE = 0").Length <> 0 Then
                        If MsgBox("This Shipment Contains Styles That have Zero Prices." _
                             & vbCrLf & "Are You Sure You Want To Update This?", MsgBoxStyle.YesNo, _
                             "Price Check") = MsgBoxResult.No Then
                            EMsg &= vbCr & "Cancelled By User Due To Zero Price."
                        End If
                    End If
                End If

            Case "Cancel"
                If select_from_3PL_list Then
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
                Mode_Settings(True)

            Case "Update"
                ' grdWHT3PLS1.Tag = ""
                If MaintenanceMode Then
                    Update_Record_Maintenance()
                Else
                    Update_Record()
                End If
                Mode_Settings(False)

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Done", "Cancel"
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
                .Items("Done").Visible = InquiryMode Or (EntryMode = "L" And ScreenMode)
                .Items("Select").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
                .Items("Update").Visible = Not InquiryMode And (Not (EntryMode = "L") Or Not ScreenMode)
                .Items("Cancel").Visible = Not InquiryMode And (Not (EntryMode = "L") Or Not ScreenMode)
                .Items("Cancel Shipment").Visible = Not InquiryMode And (Not (EntryMode = "L") Or Not ScreenMode) And Not select_from_3PL_list And Not MaintenanceMode
            End With
            .Groups("Totals").Visible = ScreenMode
            .Groups("Special Operations").Visible = ScreenMode And (EntryMode <> "L") And Not InquiryMode And Not select_from_3PL_list ' And Not MaintenanceMode
            If ASCMAIN1.DBS_COMPANY <> "VAN" Then
                .Groups("Special Operations").Visible = False
            End If
            .Groups("Special Operations").Items("Substitute").Visible = Not MaintenanceMode
            .Groups("Special Operations").Items("Add Line").Visible = Not MaintenanceMode
            '.Groups("Order Header Changes").Visible = ScreenMode And (EntryMode <> "L") And Not InquiryMode And Not select_from_3PL_list And MaintenanceMode
            .Groups("Mass Changes").Visible = ScreenMode And (EntryMode <> "L") And Not InquiryMode And Not select_from_3PL_list

            .Groups("Shipment Status").Visible = Not ScreenMode And InquiryMode
            .Groups("Shipment Selection").Visible = Not ScreenMode

            .Groups("Special Operations").Items("Force PTs to Balance").Visible = ScreenMode And edi_order And edi856_customer
            .Groups("Special Operations").Items("Force Cartons to Balance").Visible = ScreenMode And ((edi_order And edi856_customer) Or MaintenanceMode)
        End With

        '  lblStatus.Visible = ScreenMode

        'grdSOTSHIPX.Visible = Not tf
        tabSelect.Visible = Not tf

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.txtFor("ORDR_CUST_PO"), edi_order Or Not (EntryMode = "E" Or EntryMode = "N"))

        Set_Read_Only_for_ctl(Absx1.txtFor("CUST_CODE"), True)
        Set_Read_Only_for_ctl(Absx1.optFor("SHIP_ADDR_TYPE"), True)
        Set_Read_Only_for_ctl(Absx1.txtFor("SHIP_ADDR_CODE"), True)

        If ScreenMode Then
            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                {grdSOTPICK1, grdSOTPICK2, grdSOTCART1, grdSOTCART2}
                With grd.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowDelete = DefaultableBoolean.False
                    If (EntryMode = "N" Or EntryMode = "E") Then
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

            Set_Read_Only(splHeader, Not (EntryMode = "N" Or EntryMode = "E"))
            If Not InquiryMode And Not MaintenanceMode Then
                If select_from_3PL_list Then
                    Set_Read_Only(splHeader.Panel1, True)
                Else
                    Set_Read_Only(splHeader.Panel1, False)
                End If
            End If

            Set_Read_Only(frmCodes, Not (EntryMode = "E" Or EntryMode = "N"))
            Set_Read_Only(frmDates, Not (EntryMode = "E" Or EntryMode = "N"))
            Set_Read_Only(grpSHIPTO, edi_order Or Not (EntryMode = "E" Or EntryMode = "N"))
            Set_Read_Only(grpHeaderInfo, MaintenanceMode Or Not (EntryMode = "E" Or EntryMode = "N"))
            Set_Read_Only(grpShippingWindow, Not MaintenanceMode Or Not (EntryMode = "E" Or EntryMode = "N"))

            With grdSOTPICK2.DisplayLayout.Bands(0).Columns("STYLE_CODE")
                If (EntryMode = "N" Or EntryMode = "E") And (1 <> 1) Then ' HOW IN THE WORLD IS THIS TO BE PERMITTED?
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                    .Header.Appearance.BackColor2 = Drawing.Color.LightGray 'LightGreen
                End If
            End With

            With grdSOTPICK1.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"PICK_CNT_CARTONS", "PICK_TOTAL_WGT", "PICK_FREIGHT"}
                    .Columns(COLUMN_NAME).Hidden = edi_order And edi856_customer
                    If COLUMN_NAME <> "PICK_FREIGHT" Then .Columns(COLUMN_NAME & "_CALC").Hidden = Not (edi_order And edi856_customer)
                    ' NOTE THAT FRT IS NOT SHOWN IF edi_order And edi856_customer; ASSUMPTION IS THAT THERE WILL BE NO FRT IF EDI
                Next

                With .Columns("BILL_OF_LADING_NO")
                    If Not edi_order Or Absx1.optFor("SHIP_ADDR_TYPE").Value = "MK" And Not MaintenanceMode Then
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

        Absx1.txtFor("SHIP_ADDR_CODE").Visible = Not ScreenMode Or (Absx1.optFor("SHIP_ADDR_TYPE").Value = "DC")
        tabSOTPICK1.Tabs("Cartons").Visible = edi856_customer And edi_order '  And Not MaintenanceMode

        If tabSOTPICK1.Tabs("Cartons").Visible Then
            grdSOTCART2.Parent = splSOTCART1.Panel2
            grdSOTCART2.DisplayLayout.Bands(0).Columns("CART_NO").Hidden = True
            splSOTPICK1.Panel2Collapsed = True
        Else
            grdSOTCART2.Parent = splSOTPICK2.Panel2
            grdSOTCART2.DisplayLayout.Bands(0).Columns("CART_NO").Hidden = False
            splSOTPICK1.Panel2Collapsed = False
        End If

        Position_txtSTORE()
        ' lblBILL_OF_LADING_NO.Visible = Not (dst.Tables("SOTSHIP1").Rows.Count > 1)
        ' Absx1.txtFor("BILL_OF_LADING_NO").Visible = Not (dst.Tables("SOTSHIP1").Rows.Count > 1)
        If Not InquiryMode Then
            Set_Read_Only_for_ctl(Absx1.txtFor("BILL_OF_LADING_NO"), (dst.Tables("SOTSHIP1").Rows.Count > 1) Or select_from_3PL_list)
            'Set_Read_Only_for_ctl(Absx1.txtFor("REASON_CODE"), Not select_from_3PL_list)

            If select_from_3PL_list Then
                grdSOTPICK2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdSOTCART2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            End If
        End If
        grdSOTPICK1.DisplayLayout.Bands(0).Columns("BILL_OF_LADING_NO").Hidden = Not (dst.Tables("SOTSHIP1").Rows.Count > 1)

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            chkFactored.Visible = False
        End If
    End Sub

    Sub Clear_Record()

        'Absx1.txtFor("SHIP_BOL_NO").Text = ""
        'Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("PICK_NO").Text = ""

        txtReason.Text = ""
        txtContact.Text = ""
        txtemail.Text = ""

        CUST_CODE = ""
        ORDR_GROUP_NO = ""
        ORDR_CUST_PO = ""

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTSHIP1", "SOTPICK1", "SOTPICK2", "SOTINVH1", "SOTINVH2", "SOTINVH9", _
             "SOTCART1", "SOTCART2", "SOTCART3", "SOTORDR2", "SOTSHIP0", "SOTCARTX", "SOTRNGA1", _
             "SOTSHIP3", "SOTSHIP4", "SOTSHIP6", "SOTORDR5"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        select_from_3PL_list = False

        Load_SOTSHIPX()
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            If grdWHT3PLS1.Tag <> "X" And Not MaintenanceMode And Not InquiryMode Then Get_Shipments_Data_from_3PL()
        End If

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Permit Price Change"), UltraWinToolbars.StateButtonTool)
        tlb_sbt.Checked = False

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

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

            If MaintenanceMode Then
                Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)
                ORDR_SHIP_DATE = rowSOTORDR0.Item("ORDR_SHIP_DATE")
                ORDR_CANCEL_DATE = rowSOTORDR0.Item("ORDR_CANCEL_DATE")
                dteORDR_SHIP_DATE.Value = ORDR_SHIP_DATE
                dteORDR_CANCEL_DATE.Value = ORDR_CANCEL_DATE
                txtReason.Text = ""
            End If

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

            ASCMAIN1.sql = sqlSOTPICK1 & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTPICK1", "", True, ASCMAIN1.sql)

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
                rowSOTPICK1.Item("SELECTED") = "1"
            Next

            ASCMAIN1.sql = "Select * from SOTORDR5 where ORDR_NO in" & vbCrLf _
                & "(Select Distinct ORDR_NO from SOTPICK1 where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIP0 & "))" & vbCrLf _
                & " and CUST_ADDR_TYPE = 'ST'"
            Fill_Records("SOTORDR5", "", True, ASCMAIN1.sql)

            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                For Each COLUMN_NAME As String In New String() _
                    {"FRT_TERMS", "SHIP_VIA_CODE"}
                    If rowSOTSHIP1.Item(COLUMN_NAME) & "" = "" Then rowSOTSHIP1.Item(COLUMN_NAME) = rowARTCUST1.Item(COLUMN_NAME)
                Next
                If rowARTCUST1.Item("TERM_CODE") & "" <> "" Then
                    rowSOTSHIP1.Item("TERM_CODE") = rowARTCUST1.Item("TERM_CODE")
                    For Each rowSOTPICK1 As DataRow In rowSOTSHIP1.GetChildRows("SOTSHIP1_SOTPICK1")
                        rowSOTPICK1.Item("TERM_CODE") = rowARTCUST1.Item("TERM_CODE")
                    Next
                End If
            Next

            Dim rowSOTPICK1_0 As DataRow = dst.Tables("SOTPICK1").Rows(0)
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                For Each COLUMN_NAME As String In New String() _
                    {"ORDR_DEPT", "SREP_CODE", "SREP2_CODE", "TERM_CODE"}
                    If rowSOTSHIP1.Item(COLUMN_NAME) & "" = "" Then rowSOTSHIP1.Item(COLUMN_NAME) = rowSOTPICK1_0.Item(COLUMN_NAME)
                Next
            Next

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
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTPICK2", "", True, ASCMAIN1.sql)

            Fill_Records("SOTORDR9", ORDR_GROUP_NO)

            ASCMAIN1.sql = "Select SOTCART1.*" & vbCrLf _
                & " from SOTCART1,SOTPICK1,SOTSHIP1" & vbCrLf _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTCART1", "", True, ASCMAIN1.sql)
            If edi_order And edi856_customer Then
                For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
                    rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = rowSOTCART1.Item("CART_TOTAL_WGT_CALC")
                Next
            End If

            ASCMAIN1.sql = "Select SOTCART2.*, SOTCART1.PICK_NO, SOTCART2.QTY_PACKED QTY_PACKED_ORIG" & vbCrLf _
                & " from SOTCART2,SOTCART1,SOTPICK1,SOTSHIP1" & vbCrLf _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTCART2", "", True, ASCMAIN1.sql)

            'For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select("")
            '    rowSOTCART2.Item("QTY_PACKED") = rowSOTCART2.Item("QTY_PACKED")
            'Next

            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("")
                rowSOTPICK2.Item("PICK_UNIT_PRICE") = rowSOTPICK2.Item("ORDR_UNIT_PRICE")
                rowSOTPICK2.Item("PICK_QTY_CONF") = rowSOTPICK2.Item("PICK_QTY")
                rowSOTPICK2.Item("PICK_QTY_CANC") = 0
                rowSOTPICK2.Item("PICK_QTY_BACK") = 0
            Next
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
        If CURR_CODE = "" Or CURR_CODE = GL_PARM_CURR_CODE Then
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

        'grdSOTSHIP1.DisplayLayout.Bands(0).Summaries.Clear()

        grdSOTPICK1.DisplayLayout.Bands(0).Summaries.Clear()
        If dst.Tables("SOTPICK1").Rows.Count = 1 Then
            splSOTPICK1.SplitterDistance = 80 + grdSOTPICK1.Rows(0).Height * 1
            txtStore.Visible = False
        Else
            'CANT WE JUST CREATE THE SUMMARIES ONCE AND THEN HIDE THEM?
            Create_Summary(grdSOTPICK1, "PICK_NO", "Count")
            Create_Summary(grdSOTPICK1, New String() {"SELECTED", "PICK_CNT_CARTONS", "PICK_TOTAL_WGT", "PICK_FREIGHT"})
            Create_Summary(grdSOTPICK1, New String() {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"})
            Create_Summary(grdSOTPICK1, New String() {"PICK_AMT", "PICK_AMT_CONF", "PICK_AMT_CANC", "PICK_AMT_BACK"})
            splSOTPICK1.SplitterDistance = 80 + grdSOTPICK1.Rows(0).Height * 4
            txtStore.Visible = True
        End If

        If MaintenanceMode Then
            txtemail.Text = ASCMAIN1.USER_EMAIL
            txtContact.Text = ASCMAIN1.USER_NAME
        End If

        Display_Totals()
        EnforceConstraints(True)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")
        BeginTrans()

        Update_SOTORDR5()

        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        Dim WHSE_PHYS_STATUS As String = rowICTWHSE1.Item("WHSE_PHYS_STATUS") & ""

        Dim RFIXMSG As Boolean

        ' Calculate and Update Total Cartons & Weight by BOL
        ' Create New BOL's for Pick Tickets excluded from this Shipment

        ' Fetch FIFO costs for all Styles on this Group

        If edi_order And edi856_customer Then
            For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select()
                rowSOTCART1.Item("CART_TOTAL_UNITS") = rowSOTCART1.Item("CART_TOTAL_UNITS_CALC")
            Next
            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SELECTED = '1'")
                'rowSOTPICK1.Item("PICK_FREIGHT") = rowSOTPICK1.Item("PICK_FREIGHT_CALC")
                rowSOTPICK1.Item("PICK_TOTAL_WGT") = rowSOTPICK1.Item("PICK_TOTAL_WGT_CALC")
                rowSOTPICK1.Item("PICK_CNT_CARTONS") = rowSOTPICK1.Item("PICK_CNT_CARTONS_CALC")
            Next
        End If

        ASCMAIN1.Progress("Now Updating ...", "")

        Dim old_new_bols As String = ""
        Dim SHIP_BOL_NO_new As String

        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

            If select_from_3PL_list Then
                If ASCMAIN1.Running_in_VS Then Stop
                Dim AT As String = "@ADSIIS"
                ASCMAIN1.sql = "Update ADS.SOTSHIP1_3PL" & AT & " Set LP_STATUS = '3',LP_STATUS_TS_ERP = SYSDATE where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and LP_STATUS = 'V'"
                Dim I As Integer = ASCDATA1.ExecuteSQL()
                If I <> 1 Then
                    MsgBox("Problem Updating 3PL Database", MsgBoxStyle.OkOnly, "This Confirmation will be Reversed")
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")
                    Rollback()
                    Exit Sub
                End If
            End If

            Dim sqlw As String = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            Dim sqlw_selected As String = sqlw & " and SELECTED = '1'"
            Dim T As DataTable = dst.Tables("SOTPICK1")
            Dim SHIP_CNT_CARTONS As Int64 = Val(T.Compute("SUM(PICK_CNT_CARTONS)", sqlw_selected) & "")
            Dim SHIP_TOTAL_WGT As Decimal = Val(T.Compute("SUM(PICK_TOTAL_WGT)", sqlw_selected) & "")
            Dim SHIP_TOTAL_FRT As Decimal = Val(T.Compute("SUM(PICK_FREIGHT)", sqlw_selected) & "")
            Dim PICKS_SEL As Int64 = Val(T.Compute("Count(PICK_NO)", sqlw_selected) & "")
            Dim PICKS As Int64 = Val(T.Compute("Count(PICK_NO)", sqlw) & "")

            With rowSOTSHIP1
                If PICKS_SEL > 0 Then
                    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    .Item("SHIP_CNT_CARTONS") = SHIP_CNT_CARTONS
                    .Item("SHIP_TOTAL_WGT") = SHIP_TOTAL_WGT
                    .Item("SHIP_STATUS") = "F"
                    For Each COLUMN_NAME As String In New String() _
                        {"SHIP_VIA_CODE", "SHIP_DATE_SHIPPED", "INV_DATE", "REASON_CODE", "TERM_CODE", _
                         "SREP_CODE", "SREP2_CODE", "ORDR_DEPT", "SHIP_REF", "SHIP_MANIFEST_NO", "BILL_OF_LADING_NO", "FRT_TERMS"}
                        .Item(COLUMN_NAME) = rowSOTSHIP0.Item(COLUMN_NAME)
                    Next

                    If PICKS_SEL <> PICKS Then
                        Dim rowSOTSHIP1_P As DataRow = dst.Tables("SOTSHIP1").NewRow
                        With rowSOTSHIP1_P
                            For i As Integer = 0 To dst.Tables("SOTSHIP1").Columns.Count - 1
                                .Item(i) = rowSOTSHIP1.Item(i)
                            Next i
                            SHIP_BOL_NO_new = ASCMAIN1.Next_Control_No("SHIP_BOL_NO")
                            old_new_bols = old_new_bols & vbCr & SHIP_BOL_NO & " -> " & SHIP_BOL_NO_new
                            .Item("SHIP_BOL_NO") = SHIP_BOL_NO_new
                            .Item("SHIP_CNT_CARTONS") = 0
                            .Item("SHIP_TOTAL_WGT") = 0
                            .Item("SHIP_STATUS") = "P"
                            .Item("OPS_YYYYPP") = ""
                            For Each COLUMN_NAME As String In New String() _
                                {"SHIP_VIA_CODE", "SHIP_DATE_SHIPPED", "INV_DATE", "REASON_CODE", "TERM_CODE", _
                                 "SREP_CODE", "SREP2_CODE", "ORDR_DEPT", "SHIP_REF", "SHIP_MANIFEST_NO", "BILL_OF_LADING_NO", "FRT_TERMS"}
                                .Item(COLUMN_NAME) = rowSOTSHIP0_ORIG.Item(COLUMN_NAME)
                            Next
                        End With
                        dst.Tables("SOTSHIP1").Rows.Add(rowSOTSHIP1_P)
                        sqlw = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and SELECTED <> '1'"
                        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select(sqlw)
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
                    'For Each COLUMN_NAME As String In New String() _
                    '    {"SHIP_VIA_CODE", "SHIP_DATE_SHIPPED", "INV_DATE", "REASON_CODE", "TERM_CODE", _
                    '     "SREP_CODE", "SREP2_CODE", "ORDR_DEPT", "SHIP_REF", "SHIP_MANIFEST_NO", "BILL_OF_LADING_NO", "FRT_TERMS", ""}
                    '    .Item(COLUMN_NAME) = rowSOTSHIP0_ORIG.Item(COLUMN_NAME)
                    'Next
                End If
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
            End With
        Next

        ' Create Invoice Records

        RFIXMSG = False
        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SELECTED = '1'", "SHIP_BOL_NO")
            Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
            Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
            Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
            Dim SALES_DIVISION_CODE As String = rowSOTPICK1.Item("SALES_DIVISION_CODE")
            Dim INV_NO As String = ""
            If Val(rowSOTPICK1.Item("PICK_QTY_CONF") & "") <> 0 Then
                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    INV_NO = ASCMAIN1.Next_Control_No("INV_NO_01")
                Else
                    INV_NO = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
                End If
            End If
            Dim INV_NO_CONS As String = ""
            Dim SHIP_BOL_NO_cons As String = ""
            If rowARTCUST1.Item("CUST_CONS_INV") & "" = "1" Then
                If SHIP_BOL_NO_cons <> rowSOTPICK1.Item("SHIP_BOL_NO") Then
                    INV_NO_CONS = INV_NO
                    SHIP_BOL_NO_cons = rowSOTPICK1.Item("SHIP_BOL_NO")
                End If
            End If

            Dim INV_COGS As Decimal = 0
            Dim INV_SALES As Decimal = 0
            If INV_NO <> "" Then
                For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "'")
                    Dim ORDR_QTY_SHIP As Int64 = Val(rowSOTPICK2.Item("PICK_QTY_CONF") & "")
                    Dim STYLE_CODE As String = rowSOTPICK2.Item("STYLE_CODE")
                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                    Dim COLOR_CODE As String = rowSOTPICK2.Item("COLOR_CODE")
                    Dim rowICTSTYC1_FIFO As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
                    Dim ORDR_UNIT_COST As Decimal = Val(rowICTSTYC1_FIFO.Item("STYLE_COST_FIFO") & "")
                    Dim PICK_UNIT_PRICE As Decimal = Val(rowSOTPICK2.Item("PICK_UNIT_PRICE") & "")
                    INV_SALES = INV_SALES + ORDR_QTY_SHIP * PICK_UNIT_PRICE
                    Dim rowSOTINVH2 As DataRow = dst.Tables("SOTINVH2").NewRow
                    With rowSOTINVH2
                        .Item("INV_TYPE") = "I"
                        .Item("INV_NO") = INV_NO
                        .Item("INV_LNO") = rowSOTPICK2.Item("PICK_LNO")
                        .Item("STYLE_CODE") = STYLE_CODE
                        .Item("COLOR_CODE") = COLOR_CODE
                        .Item("ORDR_UNIT_COST") = ORDR_UNIT_COST
                        .Item("ORDR_UNIT_PRICE") = PICK_UNIT_PRICE
                        .Item("ORDR_QTY_SHIP") = ORDR_QTY_SHIP
                        .Item("CUST_CODE") = CUST_CODE
                        .Item("STYLE_CUST_CODE") = rowICTSTYL1.Item("CUST_CODE") & ""
                        'ADDED BY wr ON 1/5/04 BECUASE DAVE WAS STANDING OVER MY SHOULDER WITH A BIG KNIFE.
                        .Item("RANGE_STYLE_LNO") = Val(rowSOTPICK2.Item("RANGE_STYLE_LNO") & "")
                        INV_COGS += (ORDR_QTY_SHIP * ORDR_UNIT_COST)
                    End With
                    dst.Tables("SOTINVH2").Rows.Add(rowSOTINVH2)
                Next


                Fill_Records("SOTORDR9", ORDR_NO)
                For Each rowSOTORDR9 As DataRow In dst.Tables("SOTORDR9").Select()
                    Dim RANGE_STYLE_LNO As Int32 = Val(rowSOTORDR9.Item("RANGE_STYLE_LNO"))
                    Dim sqlw As String = "INV_NO = '" & INV_NO & "' and INV_TYPE = 'I' and RANGE_STYLE_LNO = " & CStr(RANGE_STYLE_LNO)
                    Dim ORDR_QTY_SHIP As Int64 = Val(dst.Tables("SOTINVH2").Compute("SUM(ORDR_QTY_SHIP)", sqlw) & "")
                    Dim rowSOTINVH9 As DataRow = dst.Tables("SOTINVH9").NewRow
                    With rowSOTINVH9
                        .Item("INV_TYPE") = "I"
                        .Item("INV_NO") = INV_NO
                        .Item("RANGE_STYLE_LNO") = rowSOTORDR9.Item("RANGE_STYLE_LNO")
                        .Item("RANGE_STYLE_CODE") = rowSOTORDR9.Item("RANGE_STYLE_CODE")
                        .Item("RANGE_STYLE_QTY_SHIP") = ORDR_QTY_SHIP
                        .Item("RANGE_STYLE_PRICE") = rowSOTORDR9.Item("RANGE_STYLE_PRICE")
                        .Item("RANGE_STYLE_PP_PRICE") = rowSOTORDR9.Item("RANGE_STYLE_PP_PRICE")
                        .Item("RANGE_STYLE_QTY_PER_PP") = rowSOTORDR9.Item("RANGE_STYLE_QTY_PER_PP")

                        Dim RANGE_STYLE_QTY_PER_PP As Int64 = Val(rowSOTORDR9.Item("RANGE_STYLE_QTY_PER_PP") & "")
                        If RANGE_STYLE_QTY_PER_PP = 0 Then 'Chances are the range is bad.  Add Audit trail Here.
                            Dim rowSOTRNGA1 As DataRow = dst.Tables("SOTRNGA1").NewRow
                            rowSOTRNGA1.Item("INV_TYPE") = "I"
                            rowSOTRNGA1.Item("INV_NO") = INV_NO
                            rowSOTRNGA1.Item("RANGE_STYLE_LNO") = .Item("RANGE_STYLE_LNO")
                            rowSOTRNGA1.Item("RANGE_STYLE_CODE") = .Item("RANGE_STYLE_CODE")
                            rowSOTRNGA1.Item("RANGE_STYLE_PP_QTY_SHIP") = .Item("RANGE_STYLE_QTY_SHIP")
                            rowSOTRNGA1.Item("RANGE_STYLE_QTY_PER_PP") = .Item("RANGE_STYLE_QTY_PER_PP")
                            rowSOTRNGA1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                            rowSOTRNGA1.Item("LAST_DATE") = DATETIME_STAMP
                            dst.Tables("SOTRNGA1").Rows.Add(rowSOTRNGA1)

                            .Item("RANGE_STYLE_QTY_PER_PP") = 1
                            .Item("RANGE_STYLE_PP_QTY_SHIP") = .Item("RANGE_STYLE_QTY_SHIP") / 1
                            RFIXMSG = True
                        Else
                            .Item("RANGE_STYLE_PP_QTY_SHIP") = .Item("RANGE_STYLE_QTY_SHIP") / RANGE_STYLE_QTY_PER_PP
                        End If
                    End With
                    dst.Tables("SOTINVH9").Rows.Add(rowSOTINVH9)
                Next

                Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").NewRow
                With rowSOTINVH1
                    .Item("INV_TYPE") = "I"
                    .Item("INV_NO") = INV_NO
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("CUST_STORE_NO") = rowSOTPICK1.Item("CUST_STORE_NO")
                    .Item("ORDR_CUST_PO") = rowSOTPICK1.Item("ORDR_CUST_PO")
                    .Item("ORDR_NO") = rowSOTPICK1.Item("ORDR_NO")
                    .Item("WHSE_CODE") = rowSOTPICK1.Item("WHSE_CODE")
                    .Item("INV_SALES") = INV_SALES
                    .Item("INV_FREIGHT") = rowSOTPICK1.Item("PICK_FREIGHT")
                    .Item("INV_MISC_CHG") = 0
                    .Item("INV_COGS") = INV_COGS
                    .Item("INV_TOTAL_AMOUNT") = INV_SALES + Val(rowSOTPICK1.Item("PICK_FREIGHT") & "")
                    .Item("REASON_CODE") = "SHP"
                    .Item("INV_DATE") = rowSOTSHIP0.Item("INV_DATE")
                    .Item("ORDR_BILL_TO_CUST") = rowSOTPICK1.Item("CUST_BILL_TO_CUST")
                    .Item("POST_CODE") = rowSOTPICK1.Item("POST_CODE")
                    .Item("TERM_CODE") = rowSOTSHIP0.Item("TERM_CODE")
                    .Item("SHIP_BOL_NO") = rowSOTPICK1.Item("SHIP_BOL_NO")
                    .Item("SALES_DIVISION_CODE") = rowSOTPICK1.Item("SALES_DIVISION_CODE")
                    .Item("INV_NO_CONS") = INV_NO_CONS
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("PICK_NO") = PICK_NO
                    .Item("SREP_CODE") = rowSOTSHIP0.Item("SREP_CODE")
                    .Item("SREP2_CODE") = rowSOTSHIP0.Item("SREP2_CODE")
                    .Item("INV_COMMENT") = rowSOTPICK1.Item("ORDR_INV_COMMENT")
                    If edi_customer Then
                        .Item("INV_PRINTED") = DATETIME_STAMP
                    End If
                    .Item("CUST_FACTOR_IND") = IIf(chkFactored.Checked, "1", "0") ' dynARTCUST1.Item("CUST_FACTOR_IND")
                    .Item("CUST_SURCHARGE_IND") = rowARTCUST1.Item("CUST_SURCHARGE_IND")
                    .Item("EDI_RETRANSMIT_IND") = "0"
                    If WHSE_PHYS_STATUS = "1" Then
                        .Item("SHIP_DURING_PHY") = "1"
                    End If
                    .Item("CURR_CODE") = CURR_CODE
                    .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE

                    .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT")
                    .Item("ORDR_TYPE_CODE") = rowSOTORDR1.Item("ORDR_TYPE_CODE")
                    .Item("CUST_BILL_TO_CUST") = rowSOTORDR1.Item("CUST_BILL_TO_CUST")
                    .Item("WHSE_CODE_TO") = rowSOTORDR1.Item("WHSE_CODE_TO")

                End With
                dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1)

                If rowSOTINVH1.Item("ORDR_TYPE_CODE") & "" = "T" Then
                    Update_Transfer(INV_NO)
                End If

                rowSOTPICK1.Item("INV_NO") = INV_NO
                rowSOTPICK1.Item("PICK_SHIPPED") = rowSOTSHIP0.Item("SHIP_DATE_SHIPPED")
                rowSOTPICK1.Item("PICK_STATUS") = "F"
                rowSOTPICK1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowSOTPICK1.Item("LAST_DATE") = DATETIME_STAMP
            Else
                'rowSOTPICK1.Item("PICK_SHIPPED") = rowSOTSHIP0.Item("SHIP_DATE_SHIPPED")
                rowSOTPICK1.Item("PICK_STATUS") = "C"
                rowSOTPICK1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowSOTPICK1.Item("LAST_DATE") = DATETIME_STAMP
            End If
        Next

        If RFIXMSG = True Then
            MsgBox("Range Styles Were Fixed On this Order.  Please Alert ABS", MsgBoxStyle.OkOnly, "Ranges")
        End If

        If CURR_CODE = "" Or Val(CURR_EXCH_RATE) = 0 Then
            Stop
        End If

        'For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("")
        '    rowSOTINVH1.Item("CURR_CODE") = CURR_CODE
        '    rowSOTINVH1.Item("CURR_EXCH_RATE") = CURR_EXCH_RATE
        'Next

        ' Update Database
        ' Delete SOTSHIP1,SOTPICK1,SOTPICK2 for each BOL in Process

        INIT_LAST("SOTSHIP1", False, , True)
        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
            ASCMAIN1.Progress("Updating Shipment " & SHIP_BOL_NO, "")
            Delete_Records(SHIP_BOL_NO)
        Next

        ' Copy Work Table Contents to Oracle

        For Each TABLE_NAME As String In New String() _
            {"SOTSHIP1", "SOTPICK1", "SOTPICK2", "SOTCART1", "SOTCART2", "SOTCART3"}
            dst.Tables(TABLE_NAME).AcceptChanges()
            For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")
                row.SetAdded()
            Next
            Update_Record_TDA(TABLE_NAME)
        Next

        For Each TABLE_NAME As String In New String() _
            {"SOTINVH1", "SOTINVH2", "SOTINVH9", "SOTORDR2"}
            Update_Record_TDA(TABLE_NAME)
        Next

        Update_Record_TDA("SOTRNGA1")

        Dim order_header_updates_required As Boolean = False
        Dim SQLX As String = ""
        For Each COL As String In New String() {"SREP_CODE", "SREP2_CODE", "TERM_CODE", "ORDR_DEPT"}
            If rowSOTSHIP0_ORIG.Item(COL) & "" <> rowSOTSHIP0.Item(COL) & "" Then
                order_header_updates_required = True
                Exit For
            End If
        Next

        ' Process each BOL, now that it is in Oracle
        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                Dim rowWHT3PLS1 As DataRow = dst.Tables("WHT3PLS1").Rows.Find(SHIP_BOL_NO)
                If rowWHT3PLS1 IsNot Nothing Then
                    rowWHT3PLS1.Delete()
                End If
            End If

            Dependent_Updates(1, SHIP_BOL_NO)

            If order_header_updates_required Then
                ASCMAIN1.sql = "Update SOTORDR1 " _
                     & "Set SREP_CODE = :PARM1, SREP2_CODE = :PARM2, TERM_CODE = :PARM3, ORDR_DEPT = :PARM4" & vbCrLf _
                     & " where ORDR_NO in " & vbCrLf _
                     & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
                     & " where SHIP_BOL_NO = :PARM5)" & vbCrLf
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVVV", _
                                    New Object() {rowSOTSHIP1.Item("SREP_CODE"), _
                                                  rowSOTSHIP1.Item("SREP2_CODE"), _
                                                  rowSOTSHIP1.Item("TERM_CODE"), _
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

            If CURR_CODE = "USD" Then
                ASCMAIN1.sql = "Update SOTINVH1 Set" & vbCrLf _
                    & "  CURR_CODE = '" & CURR_CODE & "'" & vbCrLf _
                    & ", CURR_EXCH_RATE = " & CStr(CURR_EXCH_RATE) & vbCrLf _
                    & ", INV_SALES_CURR = INV_SALES" & vbCrLf _
                    & ", INV_FREIGHT_CURR = INV_FREIGHT" & vbCrLf _
                    & ", INV_MISC_CHG_CURR = INV_MISC_CHG" & vbCrLf _
                    & ", INV_TOTAL_AMT_CURR = INV_TOTAL_AMOUNT" & vbCrLf _
                    & ", GST_TAX = 0" & vbCrLf _
                    & ", GST_TAX_CURR = 0" & vbCrLf _
                    & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Update SOTINVH9 Set" & vbCrLf _
                    & "  RANGE_STYLE_PRICE_CURR = RANGE_STYLE_PRICE" & vbCrLf _
                    & ", RANGE_STYLE_PP_PRICE_CURR = RANGE_STYLE_PP_PRICE" & vbCrLf _
                    & " where INV_TYPE = 'I' AND INV_NO IN (" & vbCrLf _
                    & "   Select INV_NO from SOTINVH1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Update SOTINVH2 SET ORDR_UNIT_PRICE_CURR = ORDR_UNIT_PRICE" & vbCrLf _
                    & " WHERE INV_TYPE = 'I' AND INV_NO IN (" & vbCrLf _
                    & "   SELECT INV_NO FROM SOTINVH1 WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
                ASCDATA1.ExecuteSQL()
            Else
                ASCMAIN1.sql = "" _
                    & "Begin" & vbCrLf _
                    & " Declare Cursor C1 is" & vbCrLf _
                    & "  Select SOTPICK1.INV_NO, SOTPICK2.PICK_LNO INV_LNO, SOTORDR2.ORDR_UNIT_PRICE_CURR" & vbCrLf _
                    & "    from SOTORDR2, SOTPICK1, SOTPICK2" & vbCrLf _
                    & "    where SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" & vbCrLf _
                    & "      and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                    & "      and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                    & "      and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                    & "      and SOTORDR2.ORDR_NO = SOTPICK1.ORDR_NO;" & vbCrLf _
                    & " Begin" & vbCrLf _
                    & "  For R1 in C1 Loop" & vbCrLf _
                    & "    Update SOTINVH2" & vbCrLf _
                    & "     Set ORDR_UNIT_PRICE_CURR = R1.ORDR_UNIT_PRICE_CURR," & vbCrLf _
                    & "         ORDR_UNIT_PRICE = R1.ORDR_UNIT_PRICE_CURR * " & CStr(CURR_EXCH_RATE) _
                    & "    where INV_TYPE = 'I'" & vbCrLf _
                    & "      and INV_NO = R1.INV_NO" & vbCrLf _
                    & "      and INV_LNO = R1.INV_LNO;" & vbCrLf _
                    & "  End Loop;" & vbCrLf _
                    & " End;" & vbCrLf _
                    & "End;"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "" _
                    & "Begin" & vbCrLf _
                    & " Declare Cursor C1 is" _
                    & "  Select SOTINVH2.INV_NO" & vbCrLf _
                    & "  , Sum(ORDR_UNIT_PRICE * ORDR_QTY_SHIP) INV_SALES" & vbCrLf _
                    & "  , Sum(ORDR_UNIT_PRICE_CURR * ORDR_QTY_SHIP) INV_SALES_CURR" & vbCrLf _
                    & "  from SOTINVH2" & vbCrLf _
                    & "  where SOTINVH2.INV_NO in (Select DISTINCT(SOTPICK1.INV_NO)" & vbCrLf _
                    & "                           from SOTPICK1" & vbCrLf _
                    & "                           where SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" & vbCrLf _
                    & "                             and SOTPICK1.PICK_STATUS = 'F')" & vbCrLf _
                    & "    and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
                    & "  group by SOTINVH2.INV_NO;" & vbCrLf _
                    & " Begin" & vbCrLf _
                    & "  For R1 in C1 Loop" & vbCrLf _
                    & "   Update SOTINVH1 Set" & vbCrLf _
                    & "    INV_SALES = R1.INV_SALES," & vbCrLf _
                    & "    INV_SALES_CURR = R1.INV_SALES_CURR," & vbCrLf _
                    & "    INV_FREIGHT = 0," & vbCrLf _
                    & "    INV_FREIGHT_CURR = 0," & vbCrLf _
                    & "    INV_MISC_CHG = 0," & vbCrLf _
                    & "    INV_MISC_CHG_CURR = 0" & vbCrLf _
                    & "   where SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
                    & "     and SOTINVH1.INV_NO = R1.INV_NO;" & vbCrLf _
                    & "  End Loop" & vbCrLf _
                    & " End;" & vbCrLf _
                    & "End;"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Update SOTINVH1" _
                    & " SET GST_TAX = INV_SALES * 0.070," _
                    & " GST_TAX_CURR = INV_SALES_CURR * 0.070," _
                    & " INV_TOTAL_AMOUNT = INV_SALES + (INV_SALES * 0.070)," _
                    & " INV_TOTAL_AMT_CURR = INV_SALES_CURR + (INV_SALES_CURR * 0.070)," _
                    & " CURR_CODE = '" & CURR_CODE & "'," _
                    & " CURR_EXCH_RATE = " & CStr(CURR_EXCH_RATE) _
                    & " where SOTINVH1.INV_NO IN (SELECT DISTINCT(SOTPICK1.INV_NO)" _
                    & "                       from SOTPICK1" _
                    & "                       where SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" _
                    & "                         and SOTPICK1.PICK_STATUS = 'F')" _
                    & "   and SOTINVH1.INV_TYPE = 'I'"
                ASCDATA1.ExecuteSQL()
            End If
        Next

        'Update Invoices to remove consolodated records when there is only one invoice.
        'Recv'd from Dave and implimented on 2/16 - W.R.
        ASCMAIN1.sql = "Update SOTINVH1 SET INV_NO_CONS = NULL where INV_NO_CONS IN" _
            & " (SELECT INV_NO_CONS FROM" _
            & " (SELECT INV_NO_CONS, COUNT(*) FROM SOTINVH1 WHERE INV_NO_CONS IS NOT NULL" _
            & " group by INV_NO_CONS" _
            & " having COUNT(*) = 1))"
        ASCDATA1.ExecuteSQL()

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

        CommitTrans("Update Complete")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Transfer(INV_NO As String)

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then

            If ASCMAIN1.DBS_COMPANY = "VAN" Then
                ASCDATA1.ExecuteSP("SOPSHIP1_XFRVAN", "V", New Object() {INV_NO}, New String() {"INV_NO_IN"})
                ASCMAIN1.sql = "Select XFR_NO from ICTIXFR1 where XFR_SOURCE = 'S' and CTL_NO = '" & INV_NO & "'"
                Dim XFR_NO As String = ASCDATA1.GetDataValue
                Dim SQLW As String = " where PO_SHIPMENT_NO =  '" & XFR_NO & "' AND SOURCE_TYPE = 'T'"
                ASCDATA1.ExecuteSQL("INSERT INTO ADS.WHTPORD1@ADSIIS SELECT * FROM WHTPORD1 " & SQLW)
                ASCDATA1.ExecuteSQL("INSERT INTO ADS.WHTPORD2@ADSIIS SELECT * FROM WHTPORD2 " & SQLW)
            End If

        Else

            Dim XFR_NO As String = ASCDATA1.ExecuteSF _
                                   ("SOPSHIP1_XFR", New String() {"INV_NO_IN"}, New Object() {INV_NO})
        End If
    End Sub

    Sub Update_Record_Maintenance()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")


        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIPX)
            ASCDATA1.ExecuteSQL("Truncate Table " & ASW("SOTCART1_3PL"))
            ASCDATA1.ExecuteSQL("Truncate Table " & ASW("SOTCART2_3PL"))
        End If


        BeginTrans()


        Dim SHIP_CHGREQ_NOs As New Dictionary(Of String, String)

        Dim date_changed As Boolean = False

        If Format(ORDR_SHIP_DATE, "yyyyMMdd") <> Format(dteORDR_SHIP_DATE.Value, "yyyyMMdd") _
        Or Format(ORDR_CANCEL_DATE, "yyyyMMdd") <> Format(dteORDR_CANCEL_DATE.Value, "yyyyMMdd") Then
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

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            If date_changed Or qty_changed Then

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

                ASCMAIN1.sql = "Select SOTSHIPX.SHIP_CHGREQ_NO, SOTCART2_3PL.CART_NO, SOTCART2_3PL.CART_LNO" & vbCrLf _
                    & ", SOTCART2_3PL.QTY_PACKED QTY_PACKED_OLD, X.QTY_PACKED QTY_PACKED_NEW" & vbCrLf _
                    & ", X.ITEM_CODE" & vbCrLf _
                    & " from SOTCART2_3PL, " & ASW("SOTCART2_3PL") & " X, SOTPICK1, " & SOTSHIPX & " SOTSHIPX" & vbCrLf _
                    & " where X.CART_NO = SOTCART2_3PL.CART_NO" & vbCrLf _
                    & "   and X.PICK_NO = SOTCART2_3PL.PICK_NO" & vbCrLf _
                    & "   and X.PICK_LNO = SOTCART2_3PL.PICK_LNO" & vbCrLf _
                    & "   and X.ITEM_CODE = SOTCART2_3PL.ITEM_CODE" & vbCrLf _
                    & "   and SOTPICK1.PICK_NO = SOTCART2_3PL.PICK_NO" & vbCrLf _
                    & "   and SOTSHIPX.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                    & "   and NVL(SOTCART2_3PL.QTY_PACKED,0) <> NVL(X.QTY_PACKED,0)"

                ASCMAIN1.sql = "" _
                    & "Begin " & vbCrLf _
                    & " Declare Cursor C1 is " & vbCrLf _
                    & ASCMAIN1.sql & ";" & vbCrLf _
                    & " Begin" & vbCrLf _
                    & "  For R1 in C1 Loop" & vbCrLf _
                    & "  Insert into SOTSHIP6_3PL Values (R1.SHIP_CHGREQ_NO,R1.CART_NO,R1.CART_LNO,R1.QTY_PACKED_OLD,R1.QTY_PACKED_NEW,R1.ITEM_CODE);" & vbCrLf _
                    & "  Update SOTCART2_3PL Set QTY_PACKED = R1.QTY_PACKED_NEW where CART_NO = R1.CART_NO and CART_LNO = R1.CART_LNO;" & vbCrLf _
                    & "  End Loop;" & vbCrLf _
                    & " End;" & vbCrLf _
                    & "End;"
                ASCDATA1.ExecuteSQL()

                'ASCMAIN1.sql = "Insert into SOTSHIP6_3PL " & ASCMAIN1.sql
                'ASCDATA1.ExecuteSQL()

                ' Stop ' NEED TO GET SOTCART2_3PL POPULATED

                ' should drop ADS.SOTSHIP4_3PL@ADSIIS

                Dim sqlw As String = " where SHIP_CHGREQ_NO in " _
                     & "(Select SHIP_CHGREQ_NO from SOTSHIP3" _
                     & " where LP_STATUS = '0' and LP_XNO = '" & LP_XNO & "')"
                ASCDATA1.ExecuteSQL("Insert into ADS.SOTSHIP3_3PL@ADSIIS Select SOTSHIP3.*,SOTSHIP1.WHSE_CODE from SOTSHIP3,SOTSHIP1" & sqlw & " and SOTSHIP1.SHIP_BOL_NO = SOTSHIP3.SHIP_BOL_NO")
                'ASCDATA1.ExecuteSQL("Insert into ADS.SOTSHIP4_3PL@ADSIIS Select SOTSHIP4.* from SOTSHIP4" & sqlw & " and SOTSHIP4.PICK_QTY_NEW <> SOTSHIP4.PICK_QTY_OLD")
                ASCDATA1.ExecuteSQL("Insert into ADS.SOTSHIP6_3PL@ADSIIS Select SOTSHIP6_3PL.* from SOTSHIP6_3PL" & sqlw & " and SOTSHIP6_3PL.QTY_PACKED_NEW <> SOTSHIP6_3PL.QTY_PACKED_OLD")
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
            Dim rowSOTORDRE As DataRow = dst.Tables("SOTORDRE").NewRow
            rowSOTORDRE.Item("ORDR_NO") = rowSOTORDR5.Item("ORDR_NO")
            rowSOTORDRE.Item("INIT_DATE") = DATETIME_STAMP
            rowSOTORDRE.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowSOTORDRE.Item("EVENT_CODE") = "SHPMTC"
            rowSOTORDRE.Item("EVENT_DESC") = "Ship-To Address was Changed"
            rowSOTORDRE.Item("EVENT_KEY") = ""
            dst.Tables("SOTORDRE").Rows.Add(rowSOTORDRE)

            ' Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", rowSOTORDRE.Item("ORDR_NO"))
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
        Update_Record_TDA("SOTORDRE")
        Update_Record_TDA("SOTORDXR")
        Update_Record_TDA("SOTORDR5")
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
        Load_Popup_Menu(grdSOTSHIPX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTPICK1, "BBBSB", "Select All", "De-Select All", "Propagate Value", "Hide Details", "Sales Order Inquiry")
        Load_Popup_Menu(grdSOTPICK2, "BS", "Style Status Inquiry", "Permit Price Change")
        Load_Popup_Menu(grdWHT3PLS1, "B", "Refresh")
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

        Select Case grd.Name
            Case "grdSOTPICK2"
                tlb_pop.Tools("Permit Price Change").SharedProps.Visible = Not InquiryMode And (EntryMode = "E" Or EntryMode = "N")
            Case "grdSOTPICK1"
                tlb_pop.Tools("Select All").SharedProps.Visible = Not InquiryMode And (EntryMode = "E" Or EntryMode = "N") And Not MaintenanceMode
                tlb_pop.Tools("De-Select All").SharedProps.Visible = Not InquiryMode And (EntryMode = "E" Or EntryMode = "N") And Not MaintenanceMode
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
                    If Not (EntryMode = "E" Or EntryMode = "N") Or grd.ActiveCell Is Nothing OrElse Not New String() {"PICK_CNT_CARTONS", "PICK_TOTAL_WGT", "PICK_FREIGHT", "BILL_OF_LADING_NO", "ORDR_INV_COMMENT"}.Contains(grd.ActiveCell.Column.Key) Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        tlb_btn.SharedProps.Visible = True
                        tlb_btn.SharedProps.Caption = "Propagate Values for " & grd.ActiveCell.Column.Header.Caption
                        tlb_btn.Tag = grd.ActiveCell.Column.Key
                    End If
                    For Each ToolKey As String In New String() {"Hide Details"} ' {"Select All", "De-Select All", "Hide Details"}
                        DirectCast(tlb_pop.Tools(ToolKey), UltraWinToolbars.ButtonTool).SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")
                    Next
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
                Get_Shipments_Data_from_3PL()

            Case "Permit Price Change"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
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
                Context_Launch("Load", ORDR_NO, e.Tool.Key, "SOFORDRI")
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
                            If COLUMN_NAME = "CUST_EXT" Or COLUMN_NAME = "CUST_FAX" Or COLUMN_NAME = "CUST_EMAIL" Then
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

    Sub Load_SOTSHIPX()
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        If InquiryMode Then
            ASCMAIN1.sql = sqlSOTSHIPX _
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
            If CUST_CODE = "" Then
                Fill_Records("SOTSHIPX")
                grdSOTSHIPX.Text = "Unconfirmed Shipments"
                Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)
            Else
                ASCMAIN1.sql = sqlSOTSHIPX _
                    & " and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'"
                Fill_Records("SOTSHIPX", "", , ASCMAIN1.sql)
                grdSOTSHIPX.Text = "Unconfirmed Shipments associated with " & CUST_CODE
                Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)
            End If
        End If

        grdSOTSHIPX.Visible = True
    End Sub

    Private Sub grdSOTSHIPX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTSHIPX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("SHIP_BOL_NO").Text = e.Row.Cells("SHIP_BOL_NO").Value
            Click_Command("Select")
        End If
    End Sub

    Sub Display_Totals()
        Dim KEY As Int32 = 0
        For Each COL As String In New String() _
            {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"}
            KEY += 1
            Dim rowSOTCONFT As DataRow = dst.Tables("SOTCONFT").Rows.Find(KEY)
            rowSOTCONFT.Item("QTY") = Val(dst.Tables("SOTPICK1").Compute("SUM(" & COL & ")", "SELECTED = '1'") & "")
            rowSOTCONFT.Item("AMT") = Val(dst.Tables("SOTPICK1").Compute("SUM(" & Replace(COL, "QTY", "AMT") & ")", "SELECTED = '1'") & "")
        Next
    End Sub

#Region "grdSOTPICK1"

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
        Try
            txtStore.Parent = grdSOTPICK1
            Dim r As System.Drawing.Rectangle = grdSOTPICK1.Rows(0).Cells("CUST_STORE_NO").GetUIElement().ClipRect
            txtStore.Width = grdSOTPICK1.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").Header.SizeResolved.Width
            txtStore.Left = r.Left
            txtStore.Top = grdSOTPICK1.Top
        Catch ex As Exception

        End Try


    End Sub
#End Region

#Region "grdSOTPICK2"

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
            clsPrice_Change.PICK_NO = .Cells("PICK_NO").Value
            clsPrice_Change.PICK_LNO = .Cells("PICK_LNO").Value
            clsPrice_Change.STYLE_CODE = .Cells("STYLE_CODE").Value
            clsPrice_Change.COLOR_CODE = .Cells("COLOR_CODE").Value
            clsPrice_Change.PICK_UNIT_PRICE = .Cells("PICK_UNIT_PRICE").Value
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
            ASCMAIN1.Progress("Now Changing Price for All Lines with Same Style/Color")

            SOTPICK1_Expressions(True)
            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select(sqlw)
                rowSOTPICK2.Item("PICK_UNIT_PRICE") = e.Row.Cells("PICK_UNIT_PRICE").Value
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

        If PICK_QTY_CONF < 0 Or PICK_QTY_BACK > PICK_QTY Or PICK_QTY_BACK < 0 Or PICK_QTY_CANC > PICK_QTY Or PICK_QTY_CANC < 0 Then
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

    Private Sub grdSOTCART1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCART1.AfterRowActivate
        Setup_SOTCART2_from_SOTCART1()
    End Sub

    Private Sub grdSOTCART1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCART1.AfterRowUpdate

    End Sub

    Private Sub grdSOTCART1_BeforeRowActivate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCART1.BeforeRowActivate
        If grdSOTCART2.ActiveRow IsNot Nothing AndAlso grdSOTCART2.ActiveRow.DataChanged Then
            grdSOTCART2.ActiveRow.Update()
        End If
    End Sub

#End Region

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
            Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
            Dim ORDR_NO As String = grdSOTPICK1.ActiveRow.Cells("ORDR_NO").Value
            Dim CUST_STORE_NO As String = grdSOTPICK1.ActiveRow.Cells("CUST_STORE_NO").Value
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

            If (EntryMode = "N" Or EntryMode = "E") Then
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

        Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
        Dim ORDR_GROUP_NO As String = rowSOTSHIP1.Item("ORDR_GROUP_NO")
        Dim REGISTER_XNO As String = rowSOTSHIP1.Item("REGISTER_XNO")

        Dim SHIP_BOL_NOs As New List(Of String)

        ASCMAIN1.sql = "Select SOTSHIP1.*" & vbCrLf _
            & " from SOTSHIP1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
            & " and REGISTER_XNO = '" & REGISTER_XNO & "'" & vbCrLf _
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

            If ASCMAIN1.USER_ID = "angela" Or ASCMAIN1.USER_ID = "pat" Then
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

        For Each SHIP_BOL_NO In SHIP_BOL_NOs
            Reverse_Invoice_1(SHIP_BOL_NO, dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO), INV_REVERSAL_REASON)
        Next

        ' Group Record
        ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
        ASCDATA1.ExecuteSQL()

        'Update New Control File To Force SJ&U between Confirm and Deconfirm - WR - 20051024
        ASCMAIN1.sql = "Update SOTCTLU1" _
            & " SET CTL_UPDATE_REQ = 'D'" _
            & " WHERE UPPER(CTL_KEY) = 'Z'"
        ASCDATA1.ExecuteSQL()

        CommitTrans("Record Successfully Updated")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Reverse_Invoice_1(SHIP_BOL_NO As String, rowSOTSHIP1 As DataRow, INV_REVERSAL_REASON As String)
        Dim SHIP_BOL_NO_new As String
        Dim REGISTER_XNO As String = rowSOTSHIP1.Item("REGISTER_XNO")
        Dim INV_DATE As Date = rowSOTSHIP1.Item("INV_DATE")

        Stop ' WHY REGISTER XNO IN THE WHERE CLAUSE?
        Stop ' ISNT THIS DATA ALL HERE BY NOW?
        'SQL = "Select * from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        'Call Ora_to_Acc(Nothing, "SOWPICK1", 0, "X", SQL)
        'SQL = "Select SOTPICK2.*, SOTORDR2.RANGE_STYLE_LNO, SOTORDR2.QTY_PER_PP"
        'SQL = SQL & " from SOTPICK2, SOTORDR2 "
        'SQL = SQL & " where SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO"
        'SQL = SQL & " AND SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO"
        'SQL = SQL & " AND PICK_NO in (Select PICK_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
        'Call Ora_to_Acc(Nothing, "SOWPICK2", 0, "X", SQL)
        'SQL = "Select * from SOTINVH1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and REGISTER_XNO = '" & REGISTER_XNO & "'"
        'Call Ora_to_Acc(Nothing, "SOWINVH1", 0, "X", SQL)
        'SQL = "Select * from SOTINVH2 where INV_TYPE = 'I' and INV_NO in (Select INV_NO from SOTINVH1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and REGISTER_XNO = '" & REGISTER_XNO & "')"
        'Call Ora_to_Acc(Nothing, "SOWINVH2", 0, "X", SQL)

        SHIP_BOL_NO_new = ASCMAIN1.Next_Control_No("SHIP_BOL_NO")
        rowSOTSHIP1.Item("SHIP_BOL_NO") = SHIP_BOL_NO_new
        rowSOTSHIP1.Item("SHIP_856_BATCH_NO") = "N"
        rowSOTSHIP1.Item("SHIP_810_BATCH_NO") = "N"
        rowSOTSHIP1.Item("REGISTER_XNO") = ""
        rowSOTSHIP1.Item("SHIP_BOL_NO_REV") = SHIP_BOL_NO

        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
            rowSOTPICK1.Item("SHIP_BOL_NO") = SHIP_BOL_NO_new
            rowSOTPICK1.Item("PICK_NO_REV") = rowSOTPICK1.Item("PICK_NO")
        Next
        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("")
            For Each COLUMN_NAME As String In New String() {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK", "PICK_QTY_CANC_REL", "PICK_QTY_BACK_REL"}
                rowSOTPICK2.Item(COLUMN_NAME) = -1 * Val(rowSOTPICK2.Item(COLUMN_NAME) & "")
            Next
        Next

        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("")
            rowSOTINVH1.Item("SHIP_BOL_NO") = SHIP_BOL_NO_new
            rowSOTINVH1.Item("INV_NO_REV") = rowSOTINVH1.Item("INV_NO")
            rowSOTINVH1.Item("ORDR_DATE_UPDATED") = DBNull.Value
            rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED") = DBNull.Value
            rowSOTINVH1.Item("INV_810_BATCH_NO") = DBNull.Value
            For Each COLUMN_NAME As String In New String() _
                {"INV_SALES", "INV_COGS", "INV_FREIGHT", "INV_MISC_CHG", "INV_TOTAL_AMOUNT"}
                rowSOTINVH1.Item(COLUMN_NAME) = -1 * Val(rowSOTINVH1.Item(COLUMN_NAME) & "")
            Next

            Dim SALES_DIVISION_CODE As String = rowSOTINVH1.Item("SALES_DIVISION_CODE")
            Dim PICK_NO As String = rowSOTINVH1.Item("PICK_NO")
            Dim PICK_NO_new As String = ASCMAIN1.Next_Control_No("PICK_NO", 10)
            Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
            Dim INV_NO_new As String = ASCMAIN1.Next_Control_No("INV_NO_01")

            rowSOTINVH1.Item("INV_NO") = INV_NO_new
            Stop ' do we have a relationship with SOTINVH2 to propagate this change down

            rowSOTINVH1.Item("PICK_NO") = PICK_NO_new
            rowSOTINVH1.Item("INV_COMMENT") = INV_REVERSAL_REASON

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("PICK_NO = '" & PICK_NO & "'")
                rowSOTPICK1.Item("PICK_NO") = PICK_NO_new
                rowSOTPICK1.Item("INV_NO") = INV_NO_new
            Next
            'For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "'")
            '    rowSOTPICK1.Item("PICK_NO") = PICK_NO_new
            'Next

            ASCMAIN1.sql = "Update SOTINVH1 set INV_NO_REV_BY = '" & INV_NO_new & "'" _
                & " where INV_TYPE = 'I' AND INV_NO = '" & INV_NO & "'"
            ASCDATA1.ExecuteSQL()
        Next
        For Each rowSOTINVH2 As DataRow In dst.Tables("SOTINVH2").Select("")
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY_SHIP"}
                rowSOTINVH2.Item(COLUMN_NAME) = -1 * Val(rowSOTINVH2.Item(COLUMN_NAME) & "")
            Next
        Next

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTINVH1").Select("INV_NO_CONS is Not Null"), "INV_NO_CONS, SALES_DIVISION_CODE").Rows
            Dim SALES_DIVISION_CODE As String = row.Item("SALES_DIVISION_CODE")
            Dim INV_NO As String = row.Item("INV_NO_CONS")
            Dim INV_NO_new As String = ASCMAIN1.Next_Control_No("INV_NO_01")
            Dim sqlw As String = "SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "' and INV_NO_CONS = '" & INV_NO & "'"
            For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select(sqlw)
                rowSOTINVH1.Item("INV_NO_CONS") = INV_NO_new
            Next
        Next

        INIT_LAST("SOTSHIP1", False, "SHIP_BOL_NO = '" & SHIP_BOL_NO_new & "'")
        INIT_LAST("SOTPICK1", False, "SHIP_BOL_NO = '" & SHIP_BOL_NO_new & "'")
        For Each TABLE_NAME As String In New String() {"SOTSHIP1", "SOTPICK1", "SOTPICK2", "SOTINVH1", "SOTINVH2"}
            Update_Record_TDA(TABLE_NAME)
        Next

        ASCMAIN1.sql = "Update SOTPICK1 set PICK_STATUS = 'P', PICK_SHIPPED = NULL, INV_NO = NULL" _
            & ", LAST_OPER = '" & ASCMAIN1.USER_ID & "',LAST_DATE = SYSDATE" _
            & " WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "UPDATE SOTSHIP1 SET SHIP_STATUS = 'P'" _
            & ", SHIP_DATE_SHIPPED = NULL" _
            & ", LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" _
            & ", REGISTER_XNO = NULL" _
            & ", SHIP_810_BATCH_NO = NULL" _
            & ", SHIP_856_BATCH_NO = NULL" _
            & " WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Cancel_Shipment()
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()

        For Each SHIP_BOL_NO As String In SHIP_BOL_NOs
            Dependent_Updates(-1, SHIP_BOL_NO)

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
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

    Sub Delete_Records(SHIP_BOL_NO As String)

        Dependent_Updates(-1, SHIP_BOL_NO)

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

    Private Sub grdSOTSHIP1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTSHIP1.InitializeLayout

    End Sub

    Private Sub optStatus_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optStatus.ValueChanged
        grpConfirmed.Visible = (optStatus.Value = "C")
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTSHIPX()
    End Sub

    Private Sub btnLoadHistory_Click(sender As System.Object, e As System.EventArgs) Handles btnLoadHistory.Click
        Load_SOTSHIPX()
    End Sub

    Sub Get_Shipments_Data_from_3PL()
        ' If ASCMAIN1.Running_in_VS Then Stop
        'If ASCMAIN1.DBS_COMPANY <> "TST" Or ASCMAIN1.DBS_SERVER <> "VAN" Then
        '    'If ASCMAIN1.DBS_COMPANY <> "VAN" Or ASCMAIN1.DBS_SERVER <> "VAN" Then
        '    If ASCMAIN1.Running_in_VS Then
        '        Fill_Records("WHT3PLS1")
        '        Sort_grdColumns(grdWHT3PLS1, "SHIP_BOL_NO")
        '        If ASCMAIN1.Running_in_VS Then Stop
        '    End If
        '    Exit Sub
        'End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Checking 3PL for Shipment Confirmation Data")


        'Dim testing As Boolean = False
        'Stop
        'testing = True

        ' BeginTrans()

        Try

            'If testing Then

            'ASCMAIN1.sql = "Update ADS.SOTSHIP1_3PL  Set LP_STATUS = 'V',LP_STATUS_TS_ERP = SYSDATE where LP_STATUS = '2'"
            'ASCDATA1.ExecuteSQL()
            Dim AT As String = "@ADSIIS"
            Dim sqlSHIP_BOL_NO As String = " where SHIP_BOL_NO in (Select SHIP_BOL_NO from ADS.SOTSHIP1_3PL" & AT & " where LP_STATUS IN ('2','V'))"
            ' Dim sqlPICK_NO As String = " where PICK_NO in (Select PICK_NO from ADS.SOTPICK1_3PL where SHIP_BOL_NO in (Select SHIP_BOL_NO from ADS.SOTSHIP1_3PL  where LP_STATUS = 'V'))"

            ' ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTSHIP1_3PL") & " Select * from ADS.SOTSHIP1_3PL " & sqlSHIP_BOL_NO)
            'ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTPICK1_3PL") & " Select * from ADS.SOTPICK1_3PL " & sqlSHIP_BOL_NO)
            'ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTPICK2_3PL") & " Select * from ADS.SOTPICK2_3PL " & sqlPICK_NO)
            'ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTCART1_3PL") & " Select * from ADS.SOTCART1_3PL " & sqlPICK_NO)
            'ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTCART2_3PL") & " Select * from ADS.SOTCART2_3PL " & sqlPICK_NO)

            'ASCMAIN1.sql = "Update ADS.SOTSHIP1_3PL  Set LP_STATUS = '3',LP_STATUS_TS_ERP = SYSDATE where LP_STATUS = 'V'"
            'ASCDATA1.ExecuteSQL()
            ' Else
            'ASCMAIN1.sql = "Update ADS.SOTSHIP1_3PL@ADSIIS Set LP_STATUS = 'V',LP_STATUS_TS_ERP = SYSDATE where LP_STATUS = '2'"
            'ASCDATA1.ExecuteSQL()

            'Dim sqlSHIP_BOL_NO As String = " where SHIP_BOL_NO in (Select SHIP_BOL_NO from ADS.SOTSHIP1_3PL@ADSIIS where LP_STATUS = 'V')"
            'Dim sqlPICK_NO As String = " where PICK_NO in (Select PICK_NO from ADS.SOTPICK1_3PL@ADSIIS where SHIP_BOL_NO in (Select SHIP_BOL_NO from ADS.SOTSHIP1_3PL@ADSIIS where LP_STATUS = 'V'))"

            If 1 <> 1 Then
                ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTSHIP1_3PL") & " Select * from ADS.SOTSHIP1_3PL" & AT & sqlSHIP_BOL_NO, True)
            End If
            'ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTPICK1_3PL") & " Select * from ADS.SOTPICK1_3PL@ADSIIS" & sqlSHIP_BOL_NO)
            'ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTPICK2_3PL") & " Select * from ADS.SOTPICK2_3PL@ADSIIS" & sqlPICK_NO)
            'ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTCART1_3PL") & " Select * from ADS.SOTCART1_3PL@ADSIIS" & sqlPICK_NO)
            'ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTCART2_3PL") & " Select * from ADS.SOTCART2_3PL@ADSIIS" & sqlPICK_NO)

            'ASCMAIN1.sql = "Update ADS.SOTSHIP1_3PL@ADSIIS Set LP_STATUS = '3',LP_STATUS_TS_ERP = SYSDATE where LP_STATUS = 'V'"
            'ASCDATA1.ExecuteSQL()
            ' End If

            '   CommitTrans()

        Catch ex As Exception

            '   Rollback()

        End Try

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        dst.Tables("WHT3PLS1").Rows.Clear()

        Fill_Records("WHT3PLS1")
        Sort_grdColumns(grdWHT3PLS1, "SHIP_BOL_NO")

        grdWHT3PLS1.Tag = "X"
    End Sub

    Private Sub grdWHT3PLS1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHT3PLS1.DoubleClickRow

        optShipmentSelection.Value = "G"
        Dim SHIP_BOL_NO As String = e.Row.Cells("SHIP_BOL_NO").Value
        Absx1.txtFor("SHIP_BOL_NO").Text = SHIP_BOL_NO
        select_from_3PL_list = True
        Click_Command("Select")

        If ScreenMode Then
            If Not Load_3PL_Shipment_Details() Then
                Click_Command("Cancel")
            End If
        Else
            select_from_3PL_list = False
        End If
    End Sub

    Function Load_3PL_Shipment_Details() As Boolean

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data from 3PL")

        For Each TABLE_NAME As String In New String() _
           {"SOTSHIP1_3PL", "SOTPICK1_3PL", "SOTPICK2_3PL", "SOTCART1_3PL", "SOTCART2_3PL", "SOTCART3_3PL"}
            ASCDATA1.ExecuteSQL("Truncate Table " & ASW(TABLE_NAME))
        Next

        Dim AT As String = "@ADSIIS"
        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")

            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
            ASCMAIN1.sql = "Update ADS.SOTSHIP1_3PL" & AT & " Set LP_STATUS = 'V',LP_STATUS_TS_ERP = SYSDATE where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and LP_STATUS = '2'"
            ASCDATA1.ExecuteSQL()

            Dim sqlSHIP_BOL_NO As String = " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            Dim sqlPICK_NO As String = " where PICK_NO in (Select PICK_NO from ADS.SOTPICK1_3PL" & AT & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
            Dim sqlCART_NO As String = " where CART_NO in (Select CART_NO from ADS.SOTCART1_3PL" & AT & " where PICK_NO in (Select PICK_NO from ADS.SOTPICK1_3PL" & AT & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'))"

            ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTSHIP1_3PL") & " Select * from ADS.SOTSHIP1_3PL" & AT & " " & sqlSHIP_BOL_NO)
            ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTPICK1_3PL") & " Select * from ADS.SOTPICK1_3PL" & AT & " " & sqlSHIP_BOL_NO)
            ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTPICK2_3PL") & " Select * from ADS.SOTPICK2_3PL" & AT & " " & sqlPICK_NO)
            ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTCART1_3PL") & " Select * from ADS.SOTCART1_3PL" & AT & " " & sqlPICK_NO)
            ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTCART2_3PL") & " Select * from ADS.SOTCART2_3PL" & AT & " " & sqlPICK_NO)
            ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTCART3_3PL") & " Select * from ADS.SOTCART3_3PL" & AT & " " & sqlCART_NO)

            Dim row As DataRow = LookUp(ASW("SOTSHIP1_3PL"), SHIP_BOL_NO)

            Dim SHIP_VIA_CODE As String = row.Item("SHIP_VIA_CODE") & ""
            If SHIP_VIA_CODE <> "" Then
                Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
                If rowSOTSVIA1 Is Nothing Then
                    If MsgBox("OK to Add this Ship Via Code to the Ship Via Master?" _
                              & vbCrLf & vbCrLf & "Code: " & SHIP_VIA_CODE _
                              & vbCrLf & "Description: " & row.Item("SHIP_VIA_DESC") _
                              & vbCrLf & "SCAC Code: " & row.Item("SHIP_VIA_SCAC"), _
                              MsgBoxStyle.YesNo, "New Ship Via Encountered (" & SHIP_VIA_CODE & ")") = MsgBoxResult.Yes Then
                        ASCDATA1.ExecuteSQL("Insert into SOTSVIA1 (SHIP_VIA_CODE, SHIP_VIA_DESC, SHIP_VIA_SCAC)" _
                                            & " values (:PARM1,:PARM2,:PARM3)", _
                                            "VVV", New Object() {row.Item("SHIP_VIA_CODE"), _
                                                                 row.Item("SHIP_VIA_DESC"), _
                                                                 row.Item("SHIP_VIA_SCAC")})
                    Else
                        MsgBox("Confirmation of this Shipment will Not be permitted unless a Valid Ship Via Record for " & SHIP_VIA_CODE & " is established", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                        Me.Cursor = Cursors.Default
                        ASCMAIN1.Progress("")
                        Return False
                    End If
                End If
            End If

            row.Item("INV_DATE") = row.Item("SHIP_DATE_SHIPPED")
            For Each COLUMN_NAME As String In New String() _
                {"SHIP_DATE_SHIPPED", "SHIP_REF", "INV_DATE", "SHIP_VIA_CODE", "BILL_OF_LADING_NO", "MASTER_BILL_OF_LADING_NO", "SHIP_TRAILER_NO", "SHIP_LOAD_NO", "SHIP_APPT_NO", "SHIP_SEAL_NO"}
                rowSOTSHIP0.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                rowSOTSHIP1.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
            Next

            If rowSOTSHIP0.Item("SHIP_REF") & "" = "" Then
                If row.Item("SHIP_LOAD_NO") & "" <> "" Then
                    rowSOTSHIP0.Item("SHIP_REF") = row.Item("SHIP_LOAD_NO")
                ElseIf row.Item("SHIP_APPT_NO") & "" <> "" Then
                    rowSOTSHIP0.Item("SHIP_REF") = row.Item("SHIP_APPT_NO")
                ElseIf row.Item("SHIP_TRAILER_NO") & "" <> "" Then
                    rowSOTSHIP0.Item("SHIP_REF") = row.Item("SHIP_TRAILER_NO")
                ElseIf row.Item("SHIP_DATE_SHIPPED") & "" <> "" Then
                    rowSOTSHIP0.Item("SHIP_REF") = row.Item("SHIP_DATE_SHIPPED")
                End If
            End If

            'rowSOTSHIP1.Item("SHIP_DATE_SHIPPED") = row.Item("SHIP_DATE_SHIPPED")
            'rowSOTSHIP1.Item("INV_DATE") = row.Item("SHIP_DATE_SHIPPED")
            'rowSOTSHIP1.Item("SHIP_VIA_CODE") = row.Item("SHIP_VIA_CODE")
            'rowSOTSHIP1.Item("BILL_OF_LADING_NO") = row.Item("BILL_OF_LADING_NO")
            'rowSOTSHIP1.Item("MASTER_BILL_OF_LADING_NO") = row.Item("MASTER_BILL_OF_LADING_NO")
            'rowSOTSHIP1.Item("SHIP_TRAILER_NO") = row.Item("SHIP_TRAILER_NO")
            'rowSOTSHIP1.Item("SHIP_LOAD_NO") = row.Item("SHIP_LOAD_NO")
            'rowSOTSHIP1.Item("SHIP_APPT_NO") = row.Item("SHIP_APPT_NO")
            'rowSOTSHIP1.Item("SHIP_SEAL_NO") = row.Item("SHIP_SEAL_NO")
        Next

        ASCMAIN1.sql = "Select * from " & ASW("SOTSHIP1_3PL")
        For Each rowSOTSHIP1_3PL As DataRow In ASCDATA1.GetDataTable.Rows
            Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find _
                    (New Object() {rowSOTSHIP1_3PL.Item("SHIP_BOL_NO")})
            If rowSOTSHIP1_3PL.Item("LP_STATUS") & "" <> "V" _
            Or rowSOTSHIP1.Item("SHIP_STATUS") & "" <> "P" Then
                MsgBox("This Shipment does not have Status Flags set to values which are expected - please call ABS for 1st time handling")
                Return False
            End If
        Next

        ASCMAIN1.sql = "Select * from " & ASW("SOTPICK1_3PL")
        For Each rowSOTPICK1_3PL As DataRow In ASCDATA1.GetDataTable.Rows
            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find _
                    (New Object() {rowSOTPICK1_3PL.Item("PICK_NO")})
            If Val(rowSOTPICK1_3PL.Item("PICK_FREIGHT") & "") <> 0 _
            Or Val(rowSOTPICK1_3PL.Item("PICK_TOTAL_WGT") & "") <> 0 _
            Or Val(rowSOTPICK1_3PL.Item("PICK_CNT_CARTONS") & "") <> Val(rowSOTPICK1.Item("PICK_CNT_CARTONS") & "") Then
                MsgBox("This Shipment contains Pick Tickets with Freigh, Weight, or Carton Count mis-match - please call ABS for 1st time handling")
                Return False
            End If
            rowSOTPICK1.Item("PICK_FREIGHT") = rowSOTPICK1_3PL.Item("PICK_FREIGHT")
            rowSOTPICK1.Item("PICK_TOTAL_WGT") = rowSOTPICK1_3PL.Item("PICK_TOTAL_WGT")
            rowSOTPICK1.Item("PICK_CNT_CARTONS") = rowSOTPICK1_3PL.Item("PICK_CNT_CARTONS")
        Next


        ' PROBLEM WITH PREPACKS
        ASCMAIN1.sql = "Select * from " & ASW("SOTPICK2_3PL")
        For Each rowSOTPICK2_3PL As DataRow In ASCDATA1.GetDataTable.Rows
            Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").Rows.Find _
                    (New Object() {rowSOTPICK2_3PL.Item("PICK_NO"), rowSOTPICK2_3PL.Item("PICK_LNO")})
            If CStr(rowSOTPICK2_3PL.Item("ITEM_CODE") & "").EndsWith("PPK") Then
                MsgBox("This Shipment contains Pre-Packs - please call ABS for 1st time handling")
                Return False
            End If
            Dim PICK_QTY As Int64 = Val(rowSOTPICK2_3PL.Item("PICK_QTY") & "")
            Dim PICK_QTY_CONF As Int64 = Val(rowSOTPICK2_3PL.Item("PICK_QTY_CONF") & "")
            Dim PICK_QTY_CANC As Int64 = PICK_QTY - PICK_QTY_CONF
            If PICK_QTY_CANC < 0 Then PICK_QTY_CANC = 0
            rowSOTPICK2.Item("PICK_QTY_CONF") = PICK_QTY_CONF
            rowSOTPICK2.Item("PICK_QTY_CANC") = PICK_QTY_CANC
        Next

        If edi856_customer Then
            ASCMAIN1.sql = "Select * from " & ASW("SOTCART2_3PL")
            For Each rowSOTCART2_3PL As DataRow In ASCDATA1.GetDataTable.Rows
                Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").Rows.Find _
                        (New Object() {rowSOTCART2_3PL.Item("CART_NO"), rowSOTCART2_3PL.Item("CART_LNO")})
                If CStr(rowSOTCART2_3PL.Item("ITEM_CODE") & "").EndsWith("PPK") Then
                    MsgBox("This Shipment contains Pre-Packs - please call ABS for 1st time handling")
                    Return False
                End If
                rowSOTCART2.Item("QTY_PACKED") = rowSOTCART2_3PL.Item("QTY_PACKED_ACT")
            Next
        End If

        ' maybe need to record SOTCART3 only for non-856
        ASCMAIN1.sql = "Select * from " & ASW("SOTCART3_3PL")
        For Each rowSOTCART3_3PL As DataRow In ASCDATA1.GetDataTable.Rows
            dst.Tables("SOTCART3").Rows.Add(rowSOTCART3_3PL.ItemArray)
        Next

        Display_Totals()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return True
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

    Private Sub optShipmentSelection_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optShipmentSelection.ValueChanged

    End Sub

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

    Private Sub grdSOTCART1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTCART1.InitializeRow
        If Val(e.Row.Cells("CART_TOTAL_UNITS_CALC").Value & "") <> Val(e.Row.Cells("CART_TOTAL_UNITS_ORIG").Value & "") Then
            e.Row.Cells("CART_TOTAL_UNITS_CALC").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("CART_TOTAL_UNITS_CALC").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Private Sub grdSOTCART2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTCART2.InitializeRow
        If Val(e.Row.Cells("QTY_PACKED").Value & "") <> Val(e.Row.Cells("QTY_PACKED_ORIG").Value & "") Then
            e.Row.Cells("QTY_PACKED").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("QTY_PACKED").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Private Sub grdSOTPICK1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTPICK1.InitializeLayout

    End Sub

    Public Class Price_Change
        Public PICK_NO As String
        Public PICK_LNO As Int32
        Public STYLE_CODE As String
        Public COLOR_CODE As String
        Public PICK_UNIT_PRICE As Decimal
    End Class
End Class
