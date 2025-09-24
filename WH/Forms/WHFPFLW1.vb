Imports Infragistics.Win.UltraWinGrid
Imports System.Text
Imports System.Drawing.Printing

Public Class WHFPFLW1
    Dim SOTPICKP As String
    Dim ShowCLvl As String = "0"

    Dim ORDR_NO As String
    Dim CUST_CODE As String
    Dim PrintCust As Boolean
    Dim MergedPicks As Boolean
    'Refresh after Print
    'Print direct to printer
    'Master BOL? talk to ED

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Build_TempTable()

        With dst
            ASCMAIN1.sql = "SELECT SOTPICKP.* FROM " & SOTPICKP & " SOTPICKP"
            Create_TDA(.Tables.Add, "SOTPICKP", "**", 0, False, "", 1)

            ASCMAIN1.sql = "SELECT SOTORDR1.ORDR_TYPE_CODE, COUNT (*) TOTAL" & vbCrLf _
                & " , SUM (CASE WHEN SOTPICK1.PICK_SHIP_DATE <= TRUNC(SYSDATE) THEN 1 ELSE 0 END) TODAY " & vbCrLf _
                & ", SUM (CASE WHEN (SOTPICK1.PICK_SHIP_DATE IS NULL OR SOTPICK1.PICK_SHIP_DATE > TRUNC(SYSDATE))	" & vbCrLf _
                & " AND SOTORDR1.ORDR_CANCEL_DATE <= TRUNC(SYSDATE +5) THEN 1 ELSE 0 END) NEXT5" & vbCrLf _
                & ", SUM (CASE WHEN (SOTPICK1.PICK_SHIP_DATE IS NULL OR SOTPICK1.PICK_SHIP_DATE > TRUNC(SYSDATE)) " & vbCrLf _
                & " AND SOTORDR1.ORDR_CANCEL_DATE > next_day(trunc(sysdate),'SUN')" & vbCrLf _
                & " AND SOTORDR1.ORDR_CANCEL_DATE <= next_day(trunc(sysdate + 7),'SUN') THEN 1 ELSE 0 END) NEXTWK " & vbCrLf _
                & ", SUM (CASE WHEN (SOTPICK1.PICK_SHIP_DATE IS NULL OR SOTPICK1.PICK_SHIP_DATE > TRUNC(SYSDATE)) " & vbCrLf _
                & " AND SOTORDR1.ORDR_CANCEL_DATE > next_day(trunc(sysdate + 7),'SUN') THEN 1 ELSE 0 END) FUTURE" & vbCrLf _
                & "FROM SOTPICK1,SOTORDR1, " & SOTPICKP & " SOTPICKP " & vbCrLf _
                & "WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "  AND SOTPICK1.PICK_NO = SOTPICKP.PICK_NO" & vbCrLf _
                & "  AND SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                & "	 AND SOTORDR1.ORDR_SOURCE <> 'E'" & vbCrLf _
                & "	 AND SOTORDR1.ORDR_TYPE_CODE IN ('REG','SAM','XFR')" & vbCrLf _
                & "GROUP BY SOTORDR1.ORDR_TYPE_CODE" & vbCrLf _
                & "ORDER BY SOTORDR1.ORDR_TYPE_CODE" & vbCrLf
            Create_TDA(.Tables.Add, "SOTPICKS", "**", 0, False, "", 1)

            ASCMAIN1.sql = "SELECT SOTPICKP.PICK_NO, SOTPICKP.PICK_QTY, SOTPICKP.PICK_AMT, " & vbCrLf _
                & "SOTPICK1.PICK_SHIP_DATE, SOTPICK1.PICK_RELEASED, SOTPICK1.PICK_PRIORITY, SOTPICK1.PICK_COMPLEXITY, " & vbCrLf _
                & "SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, " & vbCrLf _
                & "SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTPICKP.PICK_COUNT, " & vbCrLf _
                & "SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CUST_PO, SOTPICK1.SHIP_BOL_NO,  SOTORDR1.ORDR_GROUP_NO, " & vbCrLf _
                & "NVL(SOTORDR1.CUST_DC_NO, SOTORDR1.CUST_STORE_NO) CUST_DC_NO, SOTORDR1.TERM_CODE " & vbCrLf _
                & "	FROM " & SOTPICKP & " SOTPICKP,SOTPICK1,SOTORDR1" & vbCrLf _
                & " WHERE SOTPICK1.PICK_NO = SOTPICKP.PICK_NO" & vbCrLf _
                & "  AND SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "  AND SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                & "  AND SOTPICK1.PICK_SHIP_DATE <= TRUNC(SYSDATE)" & vbCrLf
            Create_TDA(.Tables.Add, "SOTPICKF", "**", 0, False, "", 1)
            With .Tables("SOTPICKF")
                .Columns.Add("ADDR1")
            End With

            ASCMAIN1.sql = "SELECT SOTPICKP.PICK_NO, SOTPICKP.PICK_QTY, SOTPICKP.PICK_AMT, " & vbCrLf _
                & "SOTPICK1.PICK_SHIP_DATE, SOTPICK1.PICK_RELEASED, nvl(SOTPICK1.PICK_PRIORITY, '3') PICK_PRIORITY, nvl(SOTPICK1.PICK_COMPLEXITY, 'C') PICK_COMPLEXITY, " & vbCrLf _
                & "SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, " & vbCrLf _
                & "SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTPICKP.PICK_COUNT, " & vbCrLf _
                & "SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CUST_PO, SOTPICK1.SHIP_BOL_NO, SOTORDR1.TERM_CODE " & vbCrLf _
                & "	FROM " & SOTPICKP & " SOTPICKP,SOTPICK1,SOTORDR1" & vbCrLf _
                & " WHERE SOTPICK1.PICK_NO = SOTPICKP.PICK_NO" & vbCrLf _
                & "  AND SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                & "  AND SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
            '& "  AND (SOTPICK1.PICK_SHIP_DATE > TRUNC(SYSDATE) or SOTPICK1.PICK_SHIP_DATE IS NULL)" & vbCrLf

            Create_TDA(.Tables.Add, "WHTPICKP", "**", 0, False, "", 1)
            With .Tables("WHTPICKP")
                .Columns.Add("SEL", GetType(System.Int16))
                .Columns("SEL").DefaultValue = 0
            End With

            With .Tables.Add("WHTPICKW")
                .Columns.Add("CUST_CODE")
                .Columns.Add("SEL", GetType(System.Int16))
                .Columns("SEL").DefaultValue = 0
                .PrimaryKey = New DataColumn() {.Columns("CUST_CODE")}
            End With
            Create_Relation("WHTPICKW", "WHTPICKP", "CUST_CODE")
            With .Tables("WHTPICKW").Columns
                .Add("ORDR_CANCEL_DATE", GetType(System.DateTime), "min(child(WHTPICKW_WHTPICKP).ORDR_CANCEL_DATE)")
                .Add("PICK_RELEASED", GetType(System.DateTime), "min(child(WHTPICKW_WHTPICKP).PICK_RELEASED)")
                .Add("CUST_NAME", GetType(System.String), "min(child(WHTPICKW_WHTPICKP).CUST_NAME)")
                .Add("PICK_AMT", GetType(System.Decimal), "sum(child(WHTPICKW_WHTPICKP).PICK_AMT)")
                .Add("PICK_TICKETS", GetType(System.Int16), "count(child(WHTPICKW_WHTPICKP).PICK_NO)")
                .Add("PICK_PRINTED", GetType(System.Int16))
                .Add("CREDIT_CARD", GetType(System.String))
            End With

            ASCMAIN1.sql = "SELECT SOTPICKP.PICK_NO, SOTPICK2.PICK_LNO, SOTORDR2.STYLE_CODE, SOTORDR2.STYLE_DESC, SOTORDR2.COLOR_CODE, " & vbCrLf _
                & "SOTPICK2.PICK_QTY, SOTORDR2.CARTON_PACK_QTY, SOTORDR2.INNER_PACK_QTY " & vbCrLf _
                & " FROM " & SOTPICKP & " SOTPICKP, SOTPICK2, SOTORDR2 " & vbCrLf _
                & " WHERE SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO " & vbCrLf _
                & "  AND SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "  and SOTPICK2.PICK_NO = SOTPICKP.PICK_NO" & vbCrLf _
                & "  and SOTPICK2.PICK_QTY > 0" & vbCrLf
            Create_TDA(.Tables.Add, "SOTPICKD", "**", 0, False, "", 2)
            With .Tables("SOTPICKD")
                .Columns.Add("LOCATION_CODE")
                .Columns.Add("LOCATION_ROUTE_SEQ")
            End With

            ASCMAIN1.sql = "Select SOTPICK1.*" & vbCrLf _
                & ", SOTORDR1.CUST_STORE_NO, ARTCUST2.CUST_NAME CUST_STORE_NAME" & vbCrLf _
                & ", TRIM(SUBSTR(LPAD(SOTORDR1.CUST_STORE_NO,6,' '),3,4)) CUST_STORE_NO4, SOTSREP1.SREP_NAME " & vbCrLf _
                & " from SOTPICK1, SOTORDR1, ARTCUST2, SOTSREP1" & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = :PARM1" & vbCrLf _
                & "   and SOTPICK1.PICK_PRINTED IS NULL " & vbCrLf _
                & "   and ARTCUST2.CUST_CODE (+) = SOTORDR1.CUST_CODE and ARTCUST2.CUST_ADDR_TYPE (+) = 'MK' and ARTCUST2.CUST_ADDR_CODE (+) = SOTORDR1.CUST_STORE_NO AND SOTSREP1.SREP_CODE (+) = SOTORDR1.SREP_CODE "
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)
            .Tables("SOTPICK1").Columns.Add("CART_SERIAL_NO", GetType(System.Int32))
            .Tables("SOTPICK1").Columns.Add("PICK_TOTAL_QTY", GetType(System.Int32))

            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_DESC, ICTSTYL1.CARTON_PACK_QTY, SOTPICK1.SHIP_BOL_NO," & vbCrLf _
               & "  ICTSTYC1.STYLE_BIN, ICTSTYC1.STYLE_BIN as LOCATION_CODE, ICTSTYL1.CASE_CUBE, ICTSTYC1.UPC_CODE, ICTSTYL1.CASE_WEIGHT_GRS, nvl(ICTSTYL1.CARTONS_PER_UNIT, 0) CARTONS_PER_UNIT" & vbCrLf _
               & IIf(ASCMAIN1.CLIENT = "RGI", ", nvl(ICTSTYL1.STYLE_ASST_QTY,0) STYLE_ASST_QTY" & vbCrLf, "") _
               & IIf(ASCMAIN1.CLIENT = "RGI", ", ICTSTYL1.WHSE_MESSAGE" & vbCrLf, "") _
               & " from SOTPICK2, SOTPICK1, SOTORDR2, ICTSTYL1, ICTSTYC1" & vbCrLf _
               & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
               & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
               & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
               & "   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
               & "   and ICTSTYC1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
               & "   and ICTSTYC1.COLOR_CODE = SOTORDR2.COLOR_CODE" & vbCrLf _
               & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
               & "   and SOTPICK1.PICK_PRINTED IS NULL " & vbCrLf _
               & "   and SOTPICK1.ORDR_NO = :PARM1" & vbCrLf _
               & "   and SOTPICK2.PICK_QTY <> 0"  'To avoid picking up records representing a cancellation or backorder generated during Pick Ticket Release
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, True, "V", 2)

            If Not .Tables("SOTPICK2").Columns.Contains("LOCATION_ROUTE_SEQ") Then
                .Tables("SOTPICK2").Columns.Add("LOCATION_ROUTE_SEQ", GetType(System.Int32))
            End If
            .Tables("SOTPICK2").Columns.Add("USL_FLAG", GetType(System.String))

            ASCMAIN1.sql = "Select SOTSHIP1.* , DECODE (SOTSHIP1.SHIP_ADDR_TYPE,'DC',SOTSHIP1.SHIP_BOL_NO,'MK') SHIP_BOL_NO_X from SOTSHIP1 WHERE SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, True, "V", 1, "SHIP_PICK_PRINTED,BILL_OF_LADING_NO")

            ASCMAIN1.sql = "Select SOTORDR1.*, 'MK' AS MARK_FOR, 'ST' AS SHIP_TO from SOTORDR1 WHERE ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "V", 1)
            .Tables("SOTORDR1").Columns("ORDR_SHIP_INSTR").MaxLength = 512

            'ASCMAIN1.sql = "Select SOTORDR2.* from SOTORDR2 WHERE ORDR_NO = :PARM1"
            ASCMAIN1.sql = "Select SOTORDR2.*, ICTCOLR1.COLOR_DESC, ICTCOLR1.COLOR_CODE_LONG, ICTSIZE1.SIZE_CODE SIZE_DESC" & vbCrLf _
                & " from SOTORDR2,SOTORDR1, ICTSIZE1, ICTCOLR1 " & vbCrLf _
                & " where ICTCOLR1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE" & vbCrLf _
                & "   and ICTSIZE1.NRF_SIZE_CODE (+) = SOTORDR2.CUST_SIZE_CODE" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select SOTORDR5.* from SOTORDR5 WHERE ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR5", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select ARTCUSTQ.* from ARTCUSTQ WHERE CUST_CODE = :PARM1 and CUST_ADDR_CODE = :PARM2"
            Create_TDA(.Tables.Add, "ARTCUSTQ", "**", 0, False, "VV", 2)

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 0, False, "", 1)
            Create_TDA(.Tables.Add, "ICTWHSE1", "*", 0, False, "", 1)
            Create_TDA(.Tables.Add, "SOTSVIA1", "*", 0, False, "", 1)

            Create_TDA(.Tables.Add, "SOTSHIPB", "*", 0, True)

        End With

        grdSOTPICKS.DataSource = dst.Tables("SOTPICKS")
        grdSOTPICKF.DataSource = dst.Tables("SOTPICKF")
        grdWHTPICKP.DataSource = dst.Tables("WHTPICKW")
        grdSOTPICKD.DataSource = dst.Tables("SOTPICKD")
        grdSOTPICKY.DataSource = dst.Tables("SOTPICKD")

        'Fill_Records("ARTCUST1")
        Fill_Records("ICTWHSE1")
        Fill_Records("SOTSVIA1")
        Fill_Records("SOTPICKP")

        Create_Summary(grdWHTPICKP, "CUST_CODE", "Count")
        Create_Summary(grdWHTPICKP, New String() {"PICK_AMT", "PICK_TICKETS", "PICK_PRINTED"})

        Create_Summary(grdSOTPICKF, "CUST_CODE", "Count")
        Create_Summary(grdSOTPICKF, New String() {"PICK_AMT"})

        Create_Summary(grdSOTPICKD, "PICK_NO", "Count")
        Create_Summary(grdSOTPICKD, New String() {"PICK_QTY"})

        Create_Summary(grdSOTPICKY, "PICK_NO", "Count")
        Create_Summary(grdSOTPICKY, New String() {"PICK_QTY"})

        With grdSOTPICKS.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = Activation.NoEdit
                gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                If gcol.Key = "min_qty_ecom" Or gcol.Key = "max_qty_ecom" Or gcol.Key = "pct_qty_ecom" Or gcol.Key = "not_inseason" Then
                    gcol.CellActivation = Activation.AllowEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"TODAY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"NEXT5"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Red
                ElseIf New String() {"NEXTWK"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"FUTURE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End If
            Next
        End With

        With grdSOTPICKF.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = Activation.NoEdit
                gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"PICK_SHIP_DATE", "PICK_PRIORITY", "PICK_COMPLEXITY"}.Contains(gcol.Key) Then
                    gcol.CellAppearance.BackColor = Drawing.Color.Orange
                End If
                If New String() {"PICK_AMT", "PICK_RELEASED", "PICK_SHIP_DATE", "PICK_PRIORITY", "PICK_COMPLEXITY", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End If
            Next
        End With

        For band As Integer = 0 To 1
            With grdWHTPICKP.DisplayLayout.Bands(band)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.CellActivation = Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                    If New String() {"PICK_SHIP_DATE", "PICK_PRIORITY", "PICK_COMPLEXITY"}.Contains(gcol.Key) Then
                        gcol.CellActivation = Activation.AllowEdit
                        gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                    End If
                    If gcol.Key = "SEL" Then
                        gcol.CellActivation = Activation.AllowEdit
                        gcol.CellAppearance.BackColor = Drawing.Color.LightGreen
                    End If
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                    If New String() {"PICK_AMT", "PICK_RELEASED", "PICK_SHIP_DATE", "PICK_PRIORITY", "PICK_COMPLEXITY", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    End If
                Next
            End With
        Next

        Show_Filter(grdSOTPICKF)
        Show_Filter(grdWHTPICKP)

        spl.Panel1Collapsed = True
        'splStats.Panel2Collapsed = True

        lblDefaultPrinter.Text = Default_Printer()

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit"

                'If Absx1.txtFor("WHSE_CODE").Text = "" Then
                '    EMsg &= vbCrLf & "You must specify a Warehouse"
                'Else
                '    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                '    If rowICTWHSE1 Is Nothing Then
                '        EMsg &= vbCrLf & "Invalid Value specified for Warehouse"
                '    Else
                '        ASCMAIN1.sql = "Select Count (*) from WHTLOCM1 where LOCATION_USE = 'E' and WHSE_CODE = '" & rowICTWHSE1.Item("WHSE_CODE") & "'"
                '        Dim eCom_Locations_Count As Integer = (ASCDATA1.GetDataValue)
                '        If eCom_Locations_Count = 0 Then
                '            EMsg &= vbCrLf & "Invalid Value specified for Warehouse (no eCommerce Locations found)"
                '        End If
                '    End If
                'End If

                'If EMsg = "" Then
                '    WHSE_CODE = rowICTWHSE1.Item("WHSE_CODE")
                '    '  If Not ASCMAIN1.Logical_Open("WHTPACK1", WHSE_CODE) Then Exit Sub
                'End If

            Case "Update"
                Dim dtTable As DataTable = dst.Tables("WHTPICKP").GetChanges()
                If dtTable Is Nothing Then
                    MsgBox("No records to Update", MsgBoxStyle.Information, "Info")
                End If

            Case "Print"
                Dim PICK_NO As String
                If grdSOTPICKF.ActiveRow IsNot Nothing AndAlso grdSOTPICKF.ActiveRow.IsDataRow Then
                    ORDR_NO = grdSOTPICKF.ActiveRow.Cells("ORDR_NO").Value & ""
                    PICK_NO = grdSOTPICKF.ActiveRow.Cells("PICK_NO").Value & ""
                    CUST_CODE = grdSOTPICKF.ActiveRow.Cells("CUST_CODE").Value & ""
                    If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO, False, False, True, 2) Then
                        EMsg = "Unable to Lock Order '" & ORDR_NO & "'"
                    Else
                        If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO, False, True, True, 2) Then
                            EMsg = "Unable to lock Pick Ticket '" & PICK_NO & "'"
                        End If
                    End If
                    WIP_SHIPMENTS(CUST_CODE)
                    PrintCust = False
                    If dst.Tables("SOTPICKP").Compute("Count(PICK_NO)", "ORDR_NO = '" & ORDR_NO & "' and PICK_NO <> '" & PICK_NO & "'") <> 0 Then
                        If MsgBox("Multiple open tickets found for Order, Select 'Y' to consolidate open tickets", vbYesNoCancel, "Consolidate Ticket") = vbYes Then
                            MergePicks(ORDR_NO)
                        End If
                    End If
                    If dst.Tables("SOTPICKP").Compute("Count(PICK_NO)", "CUST_CODE = '" & CUST_CODE & "' and ORDR_NO <> '" & ORDR_NO & "'") <> 0 Then
                        Dim ORDR_NOs As String = ""
                        Dim Answer As MsgBoxResult = MsgBox("Multiple open shipments found for Customer " & CUST_CODE & vbCrLf & " Yes to Print all, No to Print Single ", vbYesNoCancel, "Multiple Open Picks")
                        If Answer = vbYes Then
                            For Each row As DataRow In dst.Tables("SOTPICKF").Select("CUST_CODE = '" & CUST_CODE & "'")
                                If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO, False, False, True, 2) Then
                                    EMsg = "Unable to Lock Order '" & ORDR_NO & "'"
                                Else
                                    If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO, False, True, True, 2) Then
                                        EMsg = "Unable to lock Pick Ticket '" & PICK_NO & "'"
                                    End If
                                End If
                                If Not ORDR_NOs.Contains(row.Item("ORDR_NO")) Then
                                    If dst.Tables("SOTPICKP").Compute("Count(PICK_NO)", "ORDR_NO = '" & row.Item("ORDR_NO") & "' and PICK_NO <> '" & row.Item("PICK_NO") & "'") <> 0 Then
                                        If MsgBox("Multiple open tickets found for Order: " & row.Item("ORDR_NO") & " , Select 'Y' to consolidate open tickets", vbYesNoCancel, "Consolidate Ticket") = vbYes Then
                                            MergePicks(row.Item("ORDR_NO"))
                                            ORDR_NOs += row.Item("ORDR_NO")
                                        End If
                                    End If
                                End If
                            Next
                            PrintCust = True
                            AssignBOL()
                        ElseIf Answer = vbCancel Then
                            EMsg = "Print Cancelled"
                        End If
                    End If
                Else
                    EMsg = "Nothing to Print"
                End If
            Case "Cancel"
                Dim dtTable As DataTable = dst.Tables("WHTPICKP").GetChanges()
                If (dtTable IsNot Nothing) AndAlso MsgBox("Select 'yes' to lose changes", vbYesNoCancel, "Not Updated") <> vbYes Then
                    EMsg &= vbCrLf & "Changes not updated"
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

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)
                Clear_Record()

            Case "Done", "Cancel"
                Mode_Settings(False)

            Case "Print"
                If PrintCust Then
                    If MergedPicks Then
                        Build_TempTable()
                        Clear_Record()
                        MergedPicks = False
                    End If
                    For Each row As DataRow In dst.Tables("SOTPICKF").Select("CUST_CODE = '" & CUST_CODE & "'")
                        ORDR_NO = row.Item("ORDR_NO")
                        Print_Report(row.Item("PICK_NO"))
                    Next
                    Clear_Record()
                    PrintCust = False
                Else
                    Print_Report(grdSOTPICKF.ActiveRow.Cells("PICK_NO").Value)
                    Clear_Record()
                End If
                ASCMAIN1.MultiTask_Release("", 0, 2)

            Case "Refresh"
                Build_TempTable()
                Clear_Record()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = not_iScreenMode
                    .Items("Refresh").Settings.Enabled = not_iScreenMode
                End With

                .Groups("Update").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tab0.Visible = Not ScreenMode
        'splShipments.Visible = ScreenMode

        With grdWHTPICKP.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            If EntryMode = "E" Then
                .AllowUpdate = DefaultableBoolean.True
            Else
                .AllowUpdate = DefaultableBoolean.False
            End If
        End With

        If EntryMode = "E" Then
            If dtPickDate.Value < Today Then
                dtPickDate.Value = Today.Date
            End If
            btnChngDt.Visible = True
        Else
            btnChngDt.Visible = False
        End If

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Preparing Data ....")

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTPICKS", "SOTPICKF", "SOTPICKD"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Fill_Records("SOTPICKS")
        Fill_Records("SOTPICKF")
        Fill_Records("SOTPICKD")
        Sort_grdColumns(grdSOTPICKF, "PICK_SHIP_DATE, PICK_PRIORITY,ORDR_CANCEL_DATE")

        For Each row As DataRow In dst.Tables("SOTPICKF").Select("")
            row.Item("ADDR1") = ASCDATA1.GetDataValue("select CUST_ADDR1 || ' - ' || CUST_CITY || ' ' || CUST_STATE  from SOTORDR5 WHERE SOTORDR5.ORDR_NO = :PARM1 and SOTORDR5.CUST_ADDR_TYPE = 'ST'", "V", row.Item("ORDR_NO"))
        Next

        If Not grdSOTPICKF.ActiveRow Is Nothing Then
            Dim PICK_NO As String = grdSOTPICKF.ActiveRow.Cells("PICK_NO").Value & ""
            Dim dvw As DataView = DirectCast(grdSOTPICKD.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PICK_NO = '" & PICK_NO & "'"
        End If
        '     Setup_tab0()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Load_Data()

    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing to Update")
        Dim sql As String

        BeginTrans()

        'SOTPICK1.PICK_SHIP_DATE, SOTPICK1.PICK_PRIORITY, SOTPICK1.PICK_COMPLEXITY
        Dim dtTable As DataTable = dst.Tables("WHTPICKP").GetChanges()
        If dtTable IsNot Nothing Then
            For Each row As DataRow In dtTable.Rows
                If Not row.Item("PICK_SHIP_DATE") & "" = "" Then
                    sql = "Update SOTPICK1 SET PICK_SHIP_DATE = :PARM1 " & vbCrLf _
                        & ", PICK_PRIORITY = :PARM2 " & vbCrLf _
                        & ", PICK_COMPLEXITY = :PARM3 " & vbCrLf _
                        & "Where  PICK_NO = :PARM4 "
                    ASCDATA1.ExecuteSQL(sql, "VVVV", New String() {String.Format("{0:dd-MMM-yy}", row.Item("PICK_SHIP_DATE")), row.Item("PICK_PRIORITY"), row.Item("PICK_COMPLEXITY"), row.Item("PICK_NO")})
                Else
                    sql = "Update SOTPICK1 SET PICK_PRIORITY = :PARM1 " & vbCrLf _
                        & ", PICK_COMPLEXITY = :PARM2 " & vbCrLf _
                        & "Where  PICK_NO = :PARM3 "
                    ASCDATA1.ExecuteSQL(sql, "VVV", New String() {row.Item("PICK_PRIORITY"), row.Item("PICK_COMPLEXITY"), row.Item("PICK_NO")})
                End If
            Next
        End If
        ASCMAIN1.MultiTask_Release("", 0, 3)
        CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        Call BeginTrans()
        Stop
        'Call Delete_Records("table")
        Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTPICKS, "S", "Show Filter")
        Load_Popup_Menu(grdSOTPICKF, "S", "Show Filter")
        Load_Popup_Menu(grdWHTPICKP, "SBBBB", "Show Filter", "Select Selected", "De-Select Selected", "Select All", "De-Select All")
        Load_Popup_Menu(grdSOTPICKD, "B", "Style Status Inquiry")
        Load_Popup_Menu(grdSOTPICKY, "B", "Style Status Inquiry")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdSOTPACKX"


            End Select

        End If
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
            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                    If Not Nothing Is grow.ChildBands Then
                        ' Loop throgh each of the child bands.
                        For Each grow2 As UltraWinGrid.UltraGridRow In grow.ChildBands(0).Rows
                            grow2.Cells("SEL").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                            grow2.Update()
                        Next
                    End If
                Next
            Case "Select Selected", "De-Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                    grow.Update()
                    If Not Nothing Is grow.ChildBands Then
                        ' Loop throgh each of the child bands.
                        For Each grow2 As UltraWinGrid.UltraGridRow In grow.ChildBands(0).Rows
                            grow2.Cells("SEL").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                            grow2.Update()
                        Next
                    End If
                Next
            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If
                'Case "Assign BOL"
                '    Me.Cursor = Cursors.WaitCursor
                '    ASCMAIN1.Progress("Checking Shipments")
                '    If grd.Selected.Rows.Count < 2 Then
                '        MsgBox("You must select multiple shipments to the same location to assign a BOL", vbCritical, "Assign BOL")
                '        Exit Select
                '    End If
                '    Dim CUST_CODEs As String = ""
                '    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                '        If CUST_CODEs = "" Then
                '            CUST_CODEs = grow.Cells("CUST_CODE").Value
                '        End If
                '        If CUST_CODEs <> grow.Cells("CUST_CODE").Value Then
                '            MsgBox("Multiple Customers Cannot be combined in a BOL", vbCritical, "Assign BOL")
                '            Exit Select
                '        End If
                '    Next
                '    AssignBOL()
                '    Me.Cursor = Cursors.Default
                '    ASCMAIN1.Progress("")
                'Case "Show Pick Ticket"

                '    Me.Cursor = Cursors.WaitCursor
                '    ASCMAIN1.Progress("Now Printing Pick Ticket")

                '    Dim REPORT_NAME As String = "SORPICKE"

                '    Print_Report_Begin()
                '    Generate_Report(REPORT_NAME, "eCommerce Pick Ticket", "")
                '    Print_Report_End()

                '    Me.Cursor = Cursors.Default
                '    ASCMAIN1.Progress("")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                Click_Command("View")
        End Select
    End Sub
#End Region

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab1()
    End Sub

    Sub Setup_tab1()
        'lblPickFilter.Visible = (tab1.SelectedTab.Key = "Pick Tickets")
        'optPickFilter.Visible = (tab1.SelectedTab.Key = "Pick Tickets")
    End Sub

    Private Sub btnChngDt_Click(sender As Object, e As EventArgs) Handles btnChngDt.Click
        If Not dtPickDate.IsDateValid Then
            MsgBox("Please select a valid date", MsgBoxStyle.Critical, "Invalid Date")
            Exit Sub
        End If
        If dst.Tables("WHTPICKP").Select("SEL = 1").Count > 0 Then
            For Each row As DataRow In dst.Tables("WHTPICKP").Select("SEL = 1")
                row.Item("PICK_SHIP_DATE") = dtPickDate.Value
                row.Item("SEL") = "0"
            Next
            'Load_Data()
        Else
            MsgBox("use CheckBox to select rows to update", vbOKOnly, "No rows Selected")
        End If
    End Sub

    Private Sub tab1_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs)
        Setup_tab1()
    End Sub

    Sub Load_Data()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        EnforceConstraints(False)
        dst.Tables("WHTPICKW").Rows.Clear()
        Fill_Records("WHTPICKP")
        For Each row As DataRow In ASCDATA1.SelectDistinct("WHTPICKP", New String() {"CUST_CODE"}).Select("")
            dst.Tables("WHTPICKW").Rows.Add(row.ItemArray)
        Next

        For Each parentRow As DataRow In dst.Tables("WHTPICKW").Rows
            ' Get all the child rows for the current parent.
            Dim childRows As DataRow() = parentRow.GetChildRows("WHTPICKW_WHTPICKP")

            ' Use LINQ to check if any child row has the specific value.
            Dim hasMatch As Boolean = childRows.Any(Function(childRow) childRow.Field(Of String)("TERM_CODE") = "CRED")

            ' Set the value of the new column for this parent row.
            parentRow("CREDIT_CARD") = IIf(hasMatch, "CC", "")
        Next

        ASCMAIN1.sql = "SELECT SOTORDR1.CUST_CODE, SUM(1) PICK_PRINTED FROM SOTORDR1, SOTSHIP1 " & vbCrLf _
                & " WHERE SOTSHIP1.ORDR_GROUP_NO = SOTORDR1.ORDR_NO " & vbCrLf _
                & " AND SOTSHIP1.SHIP_STATUS = 'P' " & vbCrLf _
                & " AND SOTSHIP1.SHIP_PICK_PRINTED IS NOT NULL " & vbCrLf _
                & " GROUP BY CUST_CODE"

        For Each row1 As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
            For Each row As DataRow In dst.Tables("WHTPICKW").Select($"CUST_CODE = '{row1.Item("CUST_CODE")}'")
                row.Item("PICK_PRINTED") = row1.Item("PICK_PRINTED")
            Next
        Next

        'temporary - it was as of 9/23/2025: Sort_grdColumns(grdWHTPICKP, "ORDR_CANCEL_DATE,pick_amt")
        Sort_grdColumns(grdWHTPICKP, "credit_card,PICK_RELEASED,ORDR_CANCEL_DATE,pick_amt")
        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Function Default_Printer()
        Dim settings As New PrinterSettings
        For Each printer As String In PrinterSettings.InstalledPrinters
            settings.PrinterName = printer
            If settings.IsDefaultPrinter Then
                Return printer
            End If
        Next
        Return String.Empty
    End Function

    Sub Build_TempTable()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Building Work Table ...")

        'allow Von Maur EDI orders

        If SOTPICKP = "" Then
            ASCMAIN1.sql = "SELECT SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO, SOTORDR1.CUST_CODE, " & vbCrLf _
            & "SUM (SOTPICK2.PICK_QTY) PICK_QTY, SUM (SOTPICK2.PICK_QTY * SOTORDR2.ORDR_UNIT_PRICE) PICK_AMT, " & vbCrLf _
            & "SUM (case when SOTPICK2.PICK_QTY = 0 then 0 else 1 end) PICK_COUNT " & vbCrLf _
            & "FROM SOTPICK1,SOTORDR1,SOTPICK2,SOTORDR2 " & vbCrLf _
            & "WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
            & "	AND SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
            & "	AND SOTPICK1.PICK_PICKER IS NULL " & vbCrLf _
            & "	AND (SOTORDR1.ORDR_SOURCE <> 'E' or SOTORDR1.CUST_CODE = '307260')" & vbCrLf _
            & "	AND SOTORDR1.ORDR_TYPE_CODE IN ('REG','SAM','XFR')" & vbCrLf _
            & "	AND SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & "	AND SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "	AND SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & " AND SOTPICK1.PICK_PRINTED IS NULL " & vbCrLf _
            & " AND SOTPICK1.WHSE_CODE in ('MS','NY') " & vbCrLf _
            & "	AND SOTPICK1.PICK_NO not in (" & vbCrLf _
            & "	SELECT PICK_NO from SOTPICK5 " & vbCrLf _
            & "	WHERE SOTPICK5.PICK_NO = SOTPICK1.PICK_NO)" & vbCrLf _
            & "GROUP BY SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO, SOTORDR1.CUST_CODE" & vbCrLf
            SOTPICKP = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        Else
            ASCMAIN1.sql = "TRUNCATE TABLE " & SOTPICKP & ""
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = "INSERT INTO " & SOTPICKP & " " & vbCrLf _
            & "SELECT SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO, SOTORDR1.CUST_CODE, " & vbCrLf _
            & "SUM (SOTPICK2.PICK_QTY) PICK_QTY, SUM (SOTPICK2.PICK_QTY * SOTORDR2.ORDR_UNIT_PRICE) PICK_AMT, " & vbCrLf _
            & "SUM (case when SOTPICK2.PICK_QTY = 0 then 0 else 1 end) PICK_COUNT " & vbCrLf _
            & "FROM SOTPICK1,SOTORDR1,SOTPICK2,SOTORDR2 " & vbCrLf _
            & "WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
            & "	AND SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
            & "	AND SOTPICK1.PICK_PICKER IS NULL " & vbCrLf _
            & "	AND (SOTORDR1.ORDR_SOURCE <> 'E' or SOTORDR1.CUST_CODE = '307260')" & vbCrLf _
            & "	AND SOTORDR1.ORDR_TYPE_CODE IN ('REG','SAM','XFR')" & vbCrLf _
            & "	AND SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & "	AND SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "	AND SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & " AND SOTPICK1.PICK_PRINTED IS NULL " & vbCrLf _
            & " AND SOTPICK1.WHSE_CODE in ('MS','NY') " & vbCrLf _
            & "	AND SOTPICK1.PICK_NO not in (" & vbCrLf _
            & "	SELECT PICK_NO from SOTPICK5 " & vbCrLf _
            & "	WHERE SOTPICK5.PICK_NO = SOTPICK1.PICK_NO)" & vbCrLf _
            & "GROUP BY SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO, SOTORDR1.CUST_CODE" & vbCrLf
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            Fill_Records("SOTPICKP")
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub MergePicks(ORDR_NO As String)

        Dim errflag As Boolean = False
        Dim sql As String = String.Empty
        Dim PICK_NO As String = String.Empty

        Dim ROWS() As DataRow = dst.Tables("SOTPICKP").Select("ORDR_NO = '" & ORDR_NO & "'", "PICK_NO")
        For Each row As DataRow In ROWS
            PICK_NO = row.Item("PICK_NO") & ""
            If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO, False, False, True, 4) Then
                errflag = True
                EMsg = EMsg & vbCrLf & "Unable to lock all Pick tickets for update, Try again"
                Exit For
            End If
        Next

        If Not errflag And Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO, False, False, True, 4) Then
            errflag = True
            EMsg = EMsg & vbCrLf & "Unable to lock all Order for update, Try again"
        End If

        If Not errflag Then
            Try
                BeginTrans()

                sql = "Begin" & vbCrLf _
                & "declare cursor c1 is" & vbCrLf _
                & "select SOTPICK1.SHIP_BOL_NO, SOTPICK1.PICK_NO, SOTPICK1.PICK_SHIP_DATE " & vbCrLf _
                & "from SOTPICK1 " & vbCrLf _
                & "WHERE SOTPICK1.ORDR_NO = :PARM1 " & vbCrLf _
                & "AND SOTPICK1.PICK_NO <> :PARM2 " & vbCrLf _
                & "AND SOTPICK1.PICK_STATUS = 'P' " & vbCrLf _
                & "AND SOTPICK1.PICK_PICKER IS NULL " & vbCrLf _
                & "AND SOTPICK1.PICK_NO not in ( " & vbCrLf _
                & "	SELECT PICK_NO from SOTPICK5 " & vbCrLf _
                & "	WHERE SOTPICK5.PICK_NO = SOTPICK1.PICK_NO);" & vbCrLf _
                & "--AND SOTPICK1.PICK_PRINTED is null;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "  begin" & vbCrLf _
                & "   UPDATE SOTPICK1 " & vbCrLf _
                & "   SET PICK_SHIP_DATE = LEAST(nvl(PICK_SHIP_DATE,TRUNC(SYSDATE)), nvl(R1.PICK_SHIP_DATE,TRUNC(SYSDATE))) " & vbCrLf _
                & "   WHERE SOTPICK1.PICK_NO = :PARM2;" & vbCrLf _
                & "   Merge into SOTPICK2 P" & vbCrLf _
                & "   USING (SELECT * FROM SOTPICK2 WHERE PICK_NO = R1.PICK_NO) R2" & vbCrLf _
                & "   ON (P.PICK_NO = lpad(:PARM2,10,'0') and P.ORDR_NO = R2.ORDR_NO and P.ORDR_LNO = R2.ORDR_LNO)" & vbCrLf _
                & "   WHEN MATCHED THEN" & vbCrLf _
                & "    UPDATE SET P.PICK_QTY = nvl(P.PICK_QTY,0) + nvl(R2.PICK_QTY,0)," & vbCrLf _
                & "           P.PICK_QTY_CONF = nvl(P.PICK_QTY_CONF,0) + nvl(R2.PICK_QTY_CONF,0)" & vbCrLf _
                & "    --WHERE P.PICK_NO = :PARM2" & vbCrLf _
                & "   WHEN NOT MATCHED THEN" & vbCrLf _
                & "        insert (PICK_NO, PICK_LNO, ORDR_NO, ORDR_LNO, PICK_QTY, PICK_QTY_CONF,PICK_QTY_CANC," & vbCrLf _
                & "        PICK_QTY_BACK,PICK_UNIT_PRICE,PICK_QTY_CANC_REL,PICK_QTY_BACK_REL,PICK_856_TD5_IND,PICK_SPLIT," & vbCrLf _
                & "        LOCATION_ROUTE_SEQ,SHORT_REASON_CODE,SHORT_REASON_COMMENT)" & vbCrLf _
                & "        VALUES (lpad(:PARM2,10,'0'), R2.PICK_LNO, R2.ORDR_NO, R2.ORDR_LNO, R2.PICK_QTY, R2.PICK_QTY_CONF,R2.PICK_QTY_CANC," & vbCrLf _
                & "        R2.PICK_QTY_BACK,R2.PICK_UNIT_PRICE,R2.PICK_QTY_CANC_REL,R2.PICK_QTY_BACK_REL,R2.PICK_856_TD5_IND," & vbCrLf _
                & "        R2.PICK_SPLIT,R2.LOCATION_ROUTE_SEQ,R2.SHORT_REASON_CODE,R2.SHORT_REASON_COMMENT);" & vbCrLf _
                & "   UPDATE SOTCART1" & vbCrLf _
                & "   SET SOTCART1.PICK_NO = :PARM2 " & vbCrLf _
                & "   WHERE SOTCART1.PICK_NO = R1.PICK_NO;" & vbCrLf _
                & "   DELETE SOTPICK2 WHERE SOTPICK2.PICK_NO = R1.PICK_NO; " & vbCrLf _
                & "   DELETE SOTPICK1 WHERE SOTPICK1.PICK_NO = R1.PICK_NO;" & vbCrLf _
                & "   DELETE SOTSHIP1 WHERE SOTSHIP1.SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
                & "   end;" & vbCrLf _
                & "  End Loop; " & vbCrLf _
                & " End;" & vbCrLf _
                & "End;" & vbCrLf
                ASCDATA1.ExecuteSQL(sql, "VV", New String() {ORDR_NO, PICK_NO})

                CommitTrans()
                MergedPicks = True

                ASCMAIN1.MultiTask_Release("", 0, 4)

            Catch ex As Exception
                Rollback(ex.Message)
                EMsg = EMsg & vbCrLf & "Unable to Merge Pick tickets, Oracle Error, Call ABS"
            End Try
        End If



    End Sub


    Sub Print_Report(ByVal PICK_NO As String)

        Try

            Fill_Records("SOTORDR1", ORDR_NO)
        Fill_Records("SOTORDR2", ORDR_NO)
        Fill_Records("SOTORDR5", ORDR_NO)
        Fill_Records("SOTPICK1", ORDR_NO)
        Fill_Records("SOTPICK2", ORDR_NO)

        Dim row As DataRow
            Dim SHIP_BOL_NO As String
            Dim USL_LIST As New List(Of String)

            If MergedPicks Then
            row = dst.Tables("SOTPICK1").Select("ORDR_NO = '" & ORDR_NO & "'")(0)
        Else
            row = dst.Tables("SOTPICK1").Select("PICK_NO = '" & PICK_NO & "'")(0)
        End If

        SHIP_BOL_NO = row.Item("SHIP_BOL_NO")
        PICK_NO = row.Item("PICK_NO")
        Fill_Records("SOTSHIP1", SHIP_BOL_NO)

            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1")(0)
            Dim WHSE_CODE = rowSOTORDR1.Item("WHSE_CODE")
            Dim CUST_CODE = rowSOTORDR1.Item("CUST_CODE")

            Fill_Records("ARTCUST1", , , "Select * from ARTCUST1 where CUST_CODE = '" & CUST_CODE & "'")
            Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1")(0)

            'If Not IsNothing(rowARTCUST1) Then
            '    If rowARTCUST1("CUST_SPECIAL_INST") & "" <> "" Then
            '        rowSOTORDR1.Item("ORDR_SHIP_INSTR") = rowARTCUST1("CUST_SPECIAL_INST")
            '    End If
            'End If


            row = dst.Tables("SOTORDR5").Select("CUST_ADDR_TYPE = 'ST'")(0)
            Dim SHIP_TO = row.Item("CUST_ADDR_CODE")

            Fill_Records("ARTCUSTQ", New String() {CUST_CODE, SHIP_TO})

            For Each rowARTCUSTQ As DataRow In dst.Tables("ARTCUSTQ").Select("")
                Dim SpecialInst As String = ""
                If rowARTCUSTQ.Item("RESIDENTIAL_ORDR") & "" = "1" Then
                    SpecialInst &= ", Residential"
                End If
                If rowARTCUSTQ.Item("INSIDE_REQ") & "" = "1" Then
                    SpecialInst &= ", Inside Delivery"
                End If
                If rowARTCUSTQ.Item("GATE_LIFT_REQ") & "" = "1" Then
                    SpecialInst &= ", Lift Gate Req"
                End If
                If rowARTCUSTQ.Item("LIMITED_ACCESS") & "" = "1" Then
                    SpecialInst &= ", Limited Access- " & rowARTCUSTQ.Item("LIMITED_ACCESS_NOTE") & ""
                End If
                If rowARTCUSTQ.Item("IRREGULAR_HOURS") & "" = "1" Then
                    SpecialInst &= ", Hours-" & rowARTCUSTQ.Item("IRREGULAR_HOURS_NOTE") & ""
                End If
                If rowARTCUSTQ.Item("APPOINTMENT_REQUIRED") & "" = "1" Then
                    SpecialInst &= ", Appointment Req- " & rowARTCUSTQ.Item("APPOINTMENT_REQUIRED_NOTE") & ""
                End If
                If rowARTCUSTQ.Item("BROKER") & "" = "1" Then
                    SpecialInst &= ", Broker- " & rowARTCUSTQ.Item("BROKER_NOTE") & ""
                End If
                If SpecialInst <> "" Then
                    rowSOTORDR1.Item("ORDR_SHIP_INSTR") &= SpecialInst.Substring(1)
                End If
            Next
            rowSOTORDR1.Item("ORDR_SHIP_INSTR") = rowSOTORDR1.Item("ORDR_SHIP_INSTR").ToString.Replace(vbCrLf, ", ") & ""

            GetLocs(WHSE_CODE, PICK_NO)

            ASCMAIN1.sql = "Select * from ICTSTAT2 where WHSE_CODE = 'US' and WHSE_QTY_PICK > 0"
            Dim tblICTSTAT2 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("")
                For Each rowSOTPICKD As DataRow In dst.Tables("SOTPICKD").Select("PICK_NO = '" & rowSOTPICK2.Item("PICK_NO") & "' and PICK_LNO = " & rowSOTPICK2.Item("PICK_LNO"))
                    rowSOTPICK2.Item("LOCATION_CODE") = rowSOTPICKD.Item("LOCATION_CODE")
                    rowSOTPICK2.Item("LOCATION_ROUTE_SEQ") = rowSOTPICKD.Item("LOCATION_ROUTE_SEQ")
                    Exit For
                Next
                For Each rowICTSTAT2 As DataRow In tblICTSTAT2.Select($"STYLE_CODE = '{rowSOTPICK2.Item("STYLE_CODE")}' and COLOR_CODE = '{rowSOTPICK2.Item("COLOR_CODE")}'")
                    rowSOTPICK2.Item("USL_FLAG") = "1"
                    If USL_LIST.Count < 5 Then
                        USL_LIST.Add(rowSOTPICK2.Item("STYLE_CODE") & "-" & rowSOTPICK2.Item("COLOR_CODE"))
                    ElseIf USL_LIST.Count = 5 Then
                        USL_LIST.Add("More...")
                    End If
                    Exit For
                Next
            Next


            Print_Report_Begin()
            CR_params.Add("SUBT", IIf(USL_LIST.Count = 0, "", "This pick ticket has USL inventory:" & String.Join(", ", USL_LIST)))
            Dim RPT As String = "WHRPFLW1"

            Generate_Report(RPT, "Pick Ticket", "", "")

            If ASCMAIN1.Running_in_VS Then
                Print_Report_End(False)
            Else
                Dim PRINTER_PORT As String = lblDefaultPrinter.Text
                Print_Report_End(True, , PRINTER_PORT) ' set to true to print without asking
            End If
            UpdatePrintRecord(SHIP_BOL_NO, PICK_NO)

            If MergedPicks Then
                Proceed("Refresh")
                MergedPicks = False
            End If

        Catch ex As Exception
            Rollback(ex.Message)
            EMsg = EMsg & vbCrLf & "Unable to Print Pick ticket"
        End Try

    End Sub

    Sub UpdatePrintRecord(SHIP_BOL_NO As String, PICK_NO As String)
        BeginTrans()

        ASCMAIN1.sql = "Update SOTPICK1 " & vbCrLf _
            & " Set PICK_PRINTED = SYSDATE, PICK_PRINTED_OPER = '" & ASCMAIN1.USER_ID & "'" & vbCrLf _
            & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" & vbCrLf
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY) " _
                 & " Select 'SOTORDR1', ORDR_NO, PICK_PRINTED, PICK_PRINTED_OPER, 'PICKTP','Pick Ticket Print', NULL from SOTPICK1" _
                 & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()

        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
            If rowSOTSHIP1.Item("SHIP_PICK_PRINTED") & "" = "" Then
                rowSOTSHIP1.Item("SHIP_PICK_PRINTED") = DATETIME_STAMP
            End If
        Next

        ASCMAIN1.sql = "Delete " & SOTPICKP & " SOTPICKP " & vbCrLf _
            & " where PICK_NO = '" & PICK_NO & "'" & vbCrLf
        ASCDATA1.ExecuteSQL()

        Update_Record_TDA("SOTSHIP1")

        CommitTrans("")
    End Sub

    Private Sub AssignBOL()
        Dim SHIP_BOL_NOs As String = ""
        Dim ORDR_NOs As String = ""
        Dim Addr As String = ""
        Dim rowSOTSHIPB As DataRow = Nothing
        Dim rowSOTSHIP1 As DataRow = Nothing
        Dim ORD_SHIP As New Dictionary(Of String, String)
        Dim BOL_NO As String = ""
        Dim CUST_ADDR As String = ""
        Dim NO_DUPS As String = ""

        ASCMAIN1.Progress("Checking Shipments")
        dst.Tables("SOTSHIP1").Rows.Clear()

        'For Each grow As UltraGridRow In grdSOTPICKF.Selected.Rows
        For Each row As DataRow In dst.Tables("SOTPICKF").Select("CUST_CODE = '" & CUST_CODE & "'")
            Fill_Records("SOTSHIP1", row.Item("SHIP_BOL_NO") & "", False)
            rowSOTSHIP1 = dst.Tables("SOTSHIP1").Rows.Find(row.Item("SHIP_BOL_NO") & "")
            If IsNothing(rowSOTSHIP1) = False AndAlso rowSOTSHIP1.Item("BILL_OF_LADING_NO") & "" = "" Then
                SHIP_BOL_NOs = SHIP_BOL_NOs & ",'" & row.Item("SHIP_BOL_NO") & "'"
                ORDR_NOs = ORDR_NOs & ",'" & row.Item("ORDR_NO") & "'"
                ORD_SHIP.Add(row.Item("ORDR_NO"), row.Item("SHIP_BOL_NO"))
                If ASCMAIN1.Logical_Lock("SOTSHIP1", row.Item("SHIP_BOL_NO") & "", False, True, True, 4) = False Or
                    ASCMAIN1.Logical_Lock("SOTORDR0", row.Item("ORDR_NO") & "", False, True, True, 4) = False Then
                    MsgBox("Unable to Lock Records, please try again", vbCritical, "Cannot Continue")
                    Return
                End If
            End If
        Next
        If IsNothing(ORD_SHIP) OrElse ORD_SHIP.Count = 0 Then
            MsgBox("No Bill of Ladings created, all previously assigned", vbOKOnly, "No Records")
            Return
        End If

        ASCMAIN1.Progress("Assigning BOL Numbers")
        BeginTrans()

        ASCMAIN1.sql = "SELECT SOTORDR5.*, SOTORDR1.CUST_CODE " & vbCrLf _
            & " FROM SOTORDR5, SOTORDR1 " & vbCrLf _
            & " WHERE SOTORDR5.CUST_ADDR_TYPE = 'ST' " & vbCrLf _
            & " AND SOTORDR1.ORDR_NO = SOTORDR5.ORDR_NO " & vbCrLf _
            & " AND SOTORDR5.ORDR_NO in (" & ORDR_NOs.Substring(1) & ")" & vbCrLf

        Dim TBL As DataTable = ASCDATA1.GetDataTable
        For Each row As DataRow In TBL.Select("", "CUST_ADDR_CODE,CUST_NAME,CUST_ADDR1")

            rowSOTSHIP1 = dst.Tables("SOTSHIP1").Rows.Find(ORD_SHIP(row.Item("ORDR_NO")))
            ASCMAIN1.Progress("-", row.Item("ORDR_NO"))
            'try to fix address discrepancies
            CUST_ADDR = row.Item("CUST_ADDR1").ToString.ToUpper & " "
            CUST_ADDR = CUST_ADDR.Replace(" ST ", " STREET ")
            CUST_ADDR = CUST_ADDR.Replace(" AV ", " AVENUE ")
            CUST_ADDR = CUST_ADDR.Replace(" AVE ", " AVENUE ")
            CUST_ADDR = CUST_ADDR.Replace(" HWY ", " HIGHWAY ")
            CUST_ADDR = CUST_ADDR.Replace(" RT ", " ROUTE ")
            CUST_ADDR = CUST_ADDR.Replace(" RTE ", " ROUTE ")

            If Addr <> row.Item("CUST_ADDR_CODE") & row.Item("CUST_NAME") & CUST_ADDR & rowSOTSHIP1.Item("FRT_TERMS") Then
                Addr = row.Item("CUST_ADDR_CODE") & row.Item("CUST_NAME") & CUST_ADDR & rowSOTSHIP1.Item("FRT_TERMS")

                BOL_NO = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTSHIPB.BOL_NO"), "0" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))

                rowSOTSHIPB = dst.Tables("SOTSHIPB").NewRow
                rowSOTSHIPB.Item("BOL_NO") = BOL_NO

                rowSOTSHIPB.Item("CUST_CODE") = row.Item("CUST_CODE") & String.Empty
                rowSOTSHIPB.Item("BOL_DATE") = DateTime.Now.ToShortDateString
                rowSOTSHIPB.Item("FRT_TERMS") = rowSOTSHIP1.Item("FRT_TERMS") & String.Empty
                rowSOTSHIPB.Item("WHSE_CODE") = rowSOTSHIP1.Item("WHSE_CODE") & String.Empty
                rowSOTSHIPB.Item("MASTER_BOL_NO") = String.Empty
                rowSOTSHIPB.Item("MASTER_BOL") = "0"
                rowSOTSHIPB.Item("SHIP_VIA_CODE") = rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty
                rowSOTSHIPB.Item("EDI_LOAD_ID") = rowSOTSHIP1.Item("EDI_LOAD_ID") & String.Empty
                rowSOTSHIPB.Item("BTB_BOL_NO") = rowSOTSHIP1.Item("BTB_BOL_NO") & String.Empty

                Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty)
                If rowSOTSVIA1 IsNot Nothing Then
                    rowSOTSHIPB.Item("SHIP_VIA_DESC") = rowSOTSVIA1.Item("SHIP_VIA_DESC") & String.Empty
                    rowSOTSHIPB.Item("SHIP_VIA_SCAC") = rowSOTSVIA1.Item("SHIP_VIA_SCAC") & String.Empty
                End If

                rowSOTSHIPB.Item("SHIP_TO_NAME") = row.Item("CUST_NAME") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_ADDR1") = row.Item("CUST_ADDR1") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_ADDR2") = row.Item("CUST_ADDR2") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_ADDR3") = row.Item("CUST_ADDR3") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_CITY") = row.Item("CUST_CITY") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_STATE") = row.Item("CUST_STATE") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_ZIP_CODE") = row.Item("CUST_ZIP_CODE") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_COUNTRY") = row.Item("CUST_COUNTRY") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_CONTACT") = row.Item("CUST_CONTACT") & String.Empty
                rowSOTSHIPB.Item("SHIP_TO_PHONE") = row.Item("CUST_PHONE") & String.Empty

                rowSOTSHIPB.Item("THIRD_PARTY") = "0"
                rowSOTSHIPB.Item("SHIP_REF") = rowSOTSHIP1.Item("SHIP_REF") & String.Empty
                rowSOTSHIPB.Item("BOL_STATUS") = "O"

                If rowSOTSHIPB.Item("INIT_OPER") & String.Empty = String.Empty Then
                    rowSOTSHIPB.Item("INIT_DATE") = DATETIME_STAMP
                    rowSOTSHIPB.Item("INIT_OPER") = ASCMAIN1.USER_ID
                End If

                rowSOTSHIPB.Item("LAST_DATE") = DATETIME_STAMP
                rowSOTSHIPB.Item("LAST_OPER") = ASCMAIN1.USER_ID

                rowSOTSHIPB.Item("FRT_3PY_CODE") = String.Empty

                dst.Tables("SOTSHIPB").Rows.Add(rowSOTSHIPB)
            End If
            rowSOTSHIP1.Item("BILL_OF_LADING_NO") = BOL_NO

        Next

        Update_Record_TDA("SOTSHIP1")
        Update_Record_TDA("SOTSHIPB")

        CommitTrans("Shipments Assigned BOL Num")

        ASCMAIN1.MultiTask_Release("", 0, 4)

    End Sub

    Private Sub grdWHTPICKP_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdWHTPICKP.AfterRowUpdate
        'Dim CART_NO As String = grdWHTPNPS1.ActiveRow.Cells("CART_NO").Value
        'ASCDATA1.ExecuteSQL("Update WHTCART1 Set CART_SEQ = " & grdWHTCART1.ActiveRow.Cells("CART_SEQ").Value & " where CART_NO = '" & CART_NO & "'")
        'Update_Record_TDA("WHTPNPS1")
        If e.Row.Band.Index = 0 Then
            For Each row As DataRow In dst.Tables("WHTPICKP").Select("CUST_CODE = '" & e.Row.Cells("CUST_CODE").Value & "'")
                row.Item("SEL") = e.Row.Cells("SEL").Value
            Next
        End If
    End Sub



    Private Sub grdWHTPICKP_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs) Handles grdWHTPICKP.BeforeCellUpdate
        If e.Cell.Column.Key = "SEL" Or e.Cell.Column.Key = "PICK_SHIP_DATE" Then
            If e.Cell.Value = "1" Then

                If e.Cell.Row.Band.Index = 0 Then
                    For Each row As DataRow In dst.Tables("WHTPICKP").Select("CUST_CODE = '" & e.Cell.Row.Cells("CUST_CODE").Value & "'")
                        Dim PICK_NO As String = row.Item("PICK_NO")
                        If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO, False, False, True, 3) Then
                            MsgBox("Pick Ticket Is Open, wait for close.", vbInformation, "Cannot Update")
                            e.Cancel = True
                        End If
                    Next
                Else
                    Dim PICK_NO As String = grdWHTPICKP.ActiveRow.Cells("PICK_NO").Value
                    If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO, False, False, True, 3) Then
                        MsgBox("Pick Ticket Is Open, wait for close.", vbInformation, "Cannot Update")
                        e.Cancel = True
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub grdSOTPICKF_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPICKF.AfterRowActivate
        SetUp_SOTPICKD(sender)
    End Sub

    Private Sub grdWHTPICKP_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTPICKP.AfterRowActivate
        SetUp_SOTPICKD(sender)
    End Sub

    Sub SetUp_SOTPICKD(grd As UltraWinGrid.UltraGrid)
        If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
            'Exit Sub
        Else
            If grd.ActiveRow.Band.Index = 1 Or grd.Name = "grdSOTPICKF" Then
                Dim PICK_NO As String
                Dim WHSE_CODE As String
                Dim dvw As DataView

                PICK_NO = grd.ActiveRow.Cells("PICK_NO").Value & ""
                WHSE_CODE = grd.ActiveRow.Cells("WHSE_CODE").Value & ""
                GetLocs(WHSE_CODE, PICK_NO)
                If EntryMode = "" Then
                    dvw = DirectCast(grdSOTPICKD.DataSource, DataTable).DefaultView
                Else
                    dvw = DirectCast(grdSOTPICKY.DataSource, DataTable).DefaultView
                End If
                dvw.RowFilter = "PICK_NO = '" & PICK_NO & "'"
            Else
                ' hide grid details

            End If

        End If
    End Sub

    Sub GetLocs(WHSE_CODE As String, PICK_NO As String)

        'Dim ROWS() As DataRow = dst.Tables("SOTPICK2").Select("PICK_NO ='" & PICK_NO & "'", "LOCATION_ROUTE_SEQ")
        'Dim PICK_NO_LINES As Int64 = ROWS.Length
        For Each rowSOTPICKD As DataRow In dst.Tables("SOTPICKD").Select("PICK_NO ='" & PICK_NO & "'")
            Dim STYLE_CODE As String = rowSOTPICKD.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTPICKD.Item("COLOR_CODE")
            Dim PICK_QTY As Int64 = rowSOTPICKD.Item("PICK_QTY")

            Dim LOCATION_CODE As String = ""
            Dim LOCATION_ROUTE_SEQ As Int32 = 0
            TAC.SOCMAIN1.GET_STYLE_COLOR_LOCATIONS(WHSE_CODE, STYLE_CODE, COLOR_CODE, LOCATION_CODE, LOCATION_ROUTE_SEQ, PICK_QTY)

            rowSOTPICKD.Item("LOCATION_CODE") = LOCATION_CODE
            rowSOTPICKD.Item("LOCATION_ROUTE_SEQ") = LOCATION_ROUTE_SEQ
        Next

    End Sub

    Private Sub grdSOTPICKF_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTPICKF.InitializeRow
        If e.Row.IsDataRow Then
            'If e.Row.Cells("LOCATION_USE").Value & "" = "E" Then
            Dim PICK_NO As String
            Dim WHSE_CODE As String

            PICK_NO = e.Row.Cells("PICK_NO").Value & ""
            WHSE_CODE = e.Row.Cells("WHSE_CODE").Value & ""
            If ShowCLvl = "1" Then GetLocs(WHSE_CODE, PICK_NO)

            If dst.Tables("SOTPICKD").Compute("Count(LOCATION_CODE)", "PICK_NO = '" & PICK_NO & "' and LOCATION_CODE like '%-C'") > 0 Then
                e.Row.CellAppearance.BackColor = Drawing.Color.DarkOrange
            End If
        End If

    End Sub


    Private Sub WIP_SHIPMENTS(ByVal CUST_CODE As String)
        'Sql = "Select * SOTSHIP1 "
    End Sub
End Class