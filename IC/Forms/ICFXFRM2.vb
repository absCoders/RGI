Imports System.Linq

Public Class ICFXFRM2
    Dim ICTSTYLX As String
    Dim sqlICTSTYLX As String

    Dim SOTORDRX As String
    Dim sqlSOTORDRX As String

    Dim SOTINVHX As String
    Dim sqlSOTINVHX As String

    Dim TABLES_OXFR() As String = {"SOTORDR1", "SOTORDR2", "SOTORDR5", "SOTPICK0", "SOTPICK1", "SOTPICK2", "SOTSHIP1", "SOTCART1", "SOTCART2"}

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Create_WorkTables(True)
        Get_PARM("SOTPARM1")
        Get_PARM("POTPARM1")
        With dst

            ASCMAIN1.sql = $"Select * from {SOTORDRX} where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2"
            Create_TDA(.Tables.Add("SOTORDRX"), SOTORDRX, "**", 0, False, "VV")

            ASCMAIN1.sql = "Select * from " & ICTSTYLX
            Create_TDA(.Tables.Add("ICTSTYLX"), ICTSTYLX, "**", 0, False)
            .Tables("ICTSTYLX").Columns.Add("QTY2XFR", GetType(System.Int32))
            .Tables("ICTSTYLX").Columns.Add("SUG_XFR", GetType(System.Int32))
            .Tables("ICTSTYLX").Columns.Add("HIDE_STYLE")
            .Tables("ICTSTYLX").Columns.Add("AVA_MS", GetType(System.Int32), "ISNULL(ONHD_MS,0)-ISNULL(PICK_MS,0)-ISNULL(OPEN_MS,0)+ISNULL(ONPO_MS,0)+ISNULL(TRAN_MS,0)")
            .Tables("ICTSTYLX").Columns.Add("AVA_US", GetType(System.Int32), "ISNULL(ONHD_US,0)-ISNULL(PICK_US,0)")
            .Tables("ICTSTYLX").Columns.Add("ATS_MS", GetType(System.Int32), "ISNULL(ONHD_MS,0)-ISNULL(OPEN_MS,0)-ISNULL(PICK_MS,0)")
            '.Tables("ICTSTYLX").Columns.Add("QTY_SHORT", GetType(System.Int32), "IIF(ORDR_QTY_OPEN > 0 AND AVA_MS < ORDR_QTY_OPEN, AVA_MS - ORDR_QTY_OPEN, NULL)")
            .Tables("ICTSTYLX").Columns.Add("TOTAL_UNITS", GetType(System.Int32), "QTY2XFR")
            .Tables("ICTSTYLX").Columns.Add("TOTAL_CASES", GetType(System.Int32), "TOTAL_UNITS / ISNULL(CARTON_PACK_QTY,0)")
            .Tables("ICTSTYLX").Columns.Add("TOTAL_CUBES", GetType(System.Decimal), "TOTAL_CASES * ISNULL(CASE_CUBE,0)")

            ASCMAIN1.sql = $"Select * from {SOTINVHX} where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2"
            Create_TDA(.Tables.Add("SOTINVHX"), SOTINVHX, "**", 0, False, "VV")

            Create_TDA(.Tables.Add, "SOTORDR1", "*", 1)
            Create_TDA(.Tables.Add, "SOTORDR2", "*", 1)
            Create_TDA(.Tables.Add, "SOTORDR5", "*", 1)
            Create_TDA(.Tables.Add, "SOTORDRG", "*")
            Create_TDA(.Tables.Add, "SOTPICK0", "*")
            Create_TDA(.Tables.Add, "SOTPICK1", "*")
            Create_TDA(.Tables.Add, "SOTPICK2", "*")
            Create_TDA(.Tables.Add, "SOTSHIP1", "*")
            Create_TDA(.Tables.Add, "SOTCART1", "*")
            Create_TDA(.Tables.Add, "SOTCART2", "*")

            ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'SOTORDR1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", 0, True, "V")

            ASCMAIN1.sql = "Select POTORDR1.INIT_DATE, POTORDR1.WHSE_CODE, POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & ", POTORDR1.PO_DATE_SHIP_BY PO_DATE_SHIP_BY_REQ, POTORDR2.PO_DATE_SHIP_BY" & vbCrLf _
                & ", POTORDR1.FACTORY_CODE, POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & ", POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_VESSEL" & vbCrLf _
                & ", POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_REF_NO, POTSHIP2.CONTAINER_NO" & vbCrLf _
                & ", POTSHIP2.PO_DATE_RECEIVED" & vbCrLf _
                & ", POTSHIP3.PO_QTY_SHP, POTSHIP3.PO_QTY_REC" & vbCrLf _
                & ", POTORDR1.VEND_CODE, POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
                & ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_ETA + NVL(POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS,0) PO_ARRIVAL_DATE" & vbCrLf _
                & ", POTORDR2.LAST_DATE_SHIP_BY, POTORDR2.LAST_OPER_SHIP_BY" & vbCrLf _
                & ", ICTATOP2.STYLE_ARRIVAL_BUFFER_DAYS, ICTATOP2.STYLE_AT_ONCE_UNTIL, ICTATOP2.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
                & " From POTSHIP1, POTSHIP2, POTSHIP3, POTORDR1, POTORDR2, ICTATOP2 " & vbCrLf _
                & " where ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "V", 0)
            With .Tables("POTORDRX")
                .Columns("PO_SHIPMENT_NO").AllowDBNull = True
                .Columns("PO_SHIPMENT_LNO").AllowDBNull = True
                .Columns("PO_REFERENCE").AllowDBNull = True
                .Columns.Add("PO_ARRIVAL_DATE_PLUS", GetType(System.DateTime))
                '  .Columns("PO_SHIPMENT_NO").AllowDBNull = True
            End With

        End With


        grdICTSTYLX.DataSource = dst.Tables("ICTSTYLX")
        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")
        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")

        Create_Summary(grdICTSTYLX, "STYLE_CODE", "Count")
        Create_Summary(grdICTSTYLX, "QTY2XFR")
        Create_Summary(grdICTSTYLX, "ORDR_QTY_OPEN")
        Create_Summary(grdICTSTYLX, "HIDE_STYLE")
        Create_Summary(grdICTSTYLX, New String() {"TOTAL_UNITS", "TOTAL_CASES", "TOTAL_CUBES"})

        Create_Summary(grdSOTORDRX, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDRX, "ORDR_QTY_OPEN")

        Create_Summary(grdSOTINVHX, "ECOM_CODE", "Count")
        Create_Summary(grdSOTINVHX, "QTY_SHIPPED_TOT")
        Create_Summary(grdSOTINVHX, "AVG_WEEK_SALES")

        Create_Summary(grdPOTORDRX, "INIT_DATE", "Count")
        Create_Summary(grdPOTORDRX, New String() {"PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_ORD", "PO_QTY_OPN"})

        For i As Integer = 1 To 18
            Dim colName As String = $"MTH_{i.ToString("D2")}_QTY"
            Create_Summary(grdSOTINVHX, colName)
        Next


        With grdICTSTYLX.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.True

            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("STYLE_STATUS").Header.Fixed = True
            .Columns("STYLE_DESC").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
            .Columns("COLOR_DESC").Header.Fixed = True
            .Columns("STYLE_COLOR_STATUS").Header.Fixed = True
            .Columns("HIDE_STYLE").Header.Fixed = True
            .Columns("STYLE_CLASS_CODE").Header.Fixed = True
            .Columns("INIT_DATE").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "QTY2XFR" Or gcol.Key = "HIDE_STYLE" Then
                    gcol.CellAppearance.BackColor = Color.PaleGreen
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                If New String() {"STYLE_CODE", "COLOR_CODE", "STYLE_DESC", "COLOR_DESC", "STYLE_STATUS", "STYLE_COLOR_STATUS"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf gcol.Key.EndsWith("_MS") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf gcol.Key.EndsWith("_US") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"CARTON_PACK_QTY", "INNER_PACK_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"ORDR_QTY_OPEN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleTurquoise
                ElseIf New String() {"QTY2XFR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleVioletRed
                ElseIf New String() {"QTY_SHORT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                End If
            Next
        End With

        'grdSOTORDR0.DisplayLayout.UseFixedHeaders = True
        With grdSOTORDRX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME", "ORDR_GROUP_NO", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            .Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"ORDR_DATE_RECD", "ORDR_PRIORITY",
                                 "ORDR_RELEASE_AVAIL_MIN", "ORDR_RELEASE_AVAIL_MAX", "ORDR_REL_SHORT", "ORDR_REL_SHORT_OPER",
                                 "ORDR_REL_ACTION_DATE", "ORDR_REL_ACTION_OPER", "TERM_CODE", "LAST_DATE", "LAST_OPER", "ORDR_SHIP_INSTR", "ORDR_MESSAGE", "EDI_PO_TYPE"}.Contains(gcol.Key) Then
                    gcol.Hidden = Not (ASCMAIN1.CLIENT = "RGI")
                End If

                If New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf gcol.Key = "ORDR_QTY_ALLO_3PL" Or gcol.Key = "ORDR_AMT_ALLO_3PL" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LimeGreen
                ElseIf gcol.Key.StartsWith("ORDR_AMT_ALLO_") Or gcol.Key.StartsWith("PCT_ALLO_") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"ORDR_CNT", "ORDR_CNT_OPEN", "ORDR_CNT_PICK"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleTurquoise
                ElseIf New String() {"ORDR_DATE", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_ORIG_SHIP_DATE", "ORDR_ORIG_CANCEL_DATE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"ORDR_CUST_PO", "CUST_DC_NO", "ORDR_DEPT", "WHSE_CODE", "EDI_MERCH_TYPE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                ElseIf New String() {"CUST_CODE", "CUST_NAME", "ORDR_GROUP_NO", "SALES_DIVISION_CODE", "SREP_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"CUST_CITY", "CUST_STATE", "CUST_COUNTRY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                Else
                    gcol.Header.Appearance.BackColor = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

            Next

        End With

        With grdSOTINVHX.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
            .Columns("ECOM_NAME").Header.Fixed = True
        End With

        With grdPOTORDRX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackColor2 = Color.LightBlue
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"STYLE_ARRIVAL_BUFFER_DAYS", "STYLE_AT_ONCE_UNTIL", "STYLE_AT_ONCE_ACTIVE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                    gcol.Hidden = Not (ASCMAIN1.CLIENT = "RGI")
                End If
            Next
        End With

        If (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") Then
            With grdPOTORDRX.DisplayLayout.Bands(0)
                .Columns("PO_SPEC_ORDR_NO").Hidden = True
                .Columns("FACTORY_CODE").Hidden = True
            End With

        End If

        ASCMAIN1.Add_Value_List(grdICTSTYLX, "STYLE_STATUS")
        ASCMAIN1.Add_Value_List(grdICTSTYLX, "STYLE_COLOR_STATUS")

        'spl.Panel1Collapsed = True

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            'need to change ava_ms back to ava_us after testing
            Case "Update"
                If dst.Tables("ICTSTYLX").Select("QTY2XFR IS NOT NULL AND QTY2XFR > 0").Length = 0 Then
                    EMsg &= "No transfer quantities were entered. Nothing to update."
                End If


                For Each row As DataRow In dst.Tables("ICTSTYLX").Select("QTY2XFR IS NOT NULL AND QTY2XFR > 0")
                    Dim QTY2XFR As Integer = row.Field(Of Integer)("QTY2XFR")
                    Dim AVA_US As Integer = row.Field(Of Integer)("AVA_US")

                    If QTY2XFR > AVA_US Then
                        EMsg &= $"Transfer qty for {row("STYLE_CODE")}-{row("COLOR_CODE")} exceeds available US qty ({QTY2XFR} > {AVA_US}).{vbCrLf}"
                        Continue For
                    End If

                    Dim valid As Boolean = True
                    Dim INNER_PACK_QTY = row("INNER_PACK_QTY")
                    Dim CARTON_PACK_QTY = row("CARTON_PACK_QTY")

                    If Not IsDBNull(INNER_PACK_QTY) AndAlso CInt(INNER_PACK_QTY) > 0 Then
                        If QTY2XFR Mod CInt(INNER_PACK_QTY) <> 0 Then
                            valid = False
                            EMsg &= $"Transfer qty for {row("STYLE_CODE")}-{row("COLOR_CODE")} must be a multiple of INNER_PACK_QTY ({INNER_PACK_QTY}).{vbCrLf}"
                        End If
                    ElseIf Not IsDBNull(CARTON_PACK_QTY) AndAlso CInt(CARTON_PACK_QTY) > 0 Then
                        If QTY2XFR Mod CInt(CARTON_PACK_QTY) <> 0 Then
                            valid = False
                            EMsg &= $"Transfer qty for {row("STYLE_CODE")}-{row("COLOR_CODE")} must be a multiple of CARTON_PACK_QTY ({CARTON_PACK_QTY}).{vbCrLf}"
                        End If
                    End If
                Next

                'Stop ' check that all non-0 qtys are positive, a multiple of case or inner, and also less than qty avail
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

            Case "Auto Populate #ToXfr"
                Auto_Populate()

            Case "Save Progress"
                Save_Progress()

            Case "Load From Save"
                Load_From_Saved()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Visible = (ScreenMode And EntryMode = "E")
                    .Items("Auto Populate #ToXfr").Visible = (ScreenMode And EntryMode = "E")
                    .Items("Save Progress").Visible = (ScreenMode And EntryMode = "E")
                    .Items("Load from Save").Visible = (ScreenMode And EntryMode = "E")
                    .Items("Cancel").Visible = (ScreenMode And EntryMode = "E")
                    .Items("Edit").Visible = (ScreenMode And EntryMode = "L")
                    .Items("Done").Visible = (ScreenMode And EntryMode = "L")
                End With
                With .Groups("Hide/Unhide Styles")
                    .Visible = (ScreenMode And EntryMode = "E")
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        SplitContainer1.Visible = True

        If ScreenMode Then

            With grdICTSTYLX.DisplayLayout.Bands(0).Columns("QTY2XFR")
                If EntryMode = "E" Then
                    .CellAppearance.BackColor = Color.LightGreen
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .CellAppearance.BackColor = Color.Empty
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            End With

            With grdICTSTYLX.DisplayLayout.Bands(0).Columns("HIDE_STYLE")
                If EntryMode = "E" Then
                    .CellAppearance.BackColor = Color.LightGreen
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .CellAppearance.BackColor = Color.Empty
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            End With

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTSTYLX", "SOTORDRX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        SplitContainer1.Visible = False
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data")
        Save_Header_Fields(UltraGroupBox1)

        Create_WorkTables()

        Fill_Records("SOTORDRX")
        Fill_Records("ICTSTYLX")

        Dim styleColorList As New List(Of String)
        For Each row As DataRow In dst.Tables("ICTSTYLX").Rows
            styleColorList.Add($"('{row("STYLE_CODE")}', '{row("COLOR_CODE")}')")
        Next

        Dim inClause As String = String.Join(",", styleColorList.Distinct())

        ASCMAIN1.sql = "
        SELECT h2.STYLE_CODE, h2.COLOR_CODE" & vbCrLf &
        String.Join(vbCrLf, Enumerable.Range(1, 18).Select(Function(i) _
            $" , SUM(CASE WHEN TO_CHAR(h1.INV_DATE, 'YYYYMM') = TO_CHAR(ADD_MONTHS(TRUNC(SYSDATE,'MM'), -{i - 1}), 'YYYYMM') THEN h2.ORDR_QTY_SHIP ELSE 0 END) AS MTH_{i:D2}_QTY")) & vbCrLf &
        "FROM SOTINVH1 h1
         JOIN SOTINVH2 h2 ON h1.INV_TYPE = h2.INV_TYPE AND h1.INV_NO = h2.INV_NO
         JOIN ECTECOM1 e ON h1.CUST_CODE = e.CUST_CODE
        WHERE h1.INV_DATE >= ADD_MONTHS(TRUNC(SYSDATE, 'MM'), -17)
          AND h2.ORDR_QTY_SHIP > 0
          AND (h2.STYLE_CODE, h2.COLOR_CODE) IN (" & inClause & ")
        GROUP BY h2.STYLE_CODE, h2.COLOR_CODE"

        Dim dtMonthlySales As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        For Each row As DataRow In dst.Tables("ICTSTYLX").Rows
            Dim STYLE_CODE As String = row("STYLE_CODE")
            Dim COLOR_CODE As String = row("COLOR_CODE")
            Dim matchedRows = dtMonthlySales.Select($"STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE = '{COLOR_CODE}'")

            If matchedRows.Length = 0 Then
                row("AVG_MONTH_SALES") = 0
                row("AVG_WEEK_SALES") = 0
                Continue For
            End If

            Dim monthlyQtys As New List(Of Decimal)
            For i As Integer = 1 To 18
                Dim qty As Decimal = Val(matchedRows(0)($"MTH_{i:D2}_QTY") & "")
                monthlyQtys.Add(qty)
            Next

            Dim firstNonZero As Integer = monthlyQtys.FindIndex(Function(q) q <> 0)
            If firstNonZero >= 0 Then
                Dim relevantMonths As List(Of Decimal) =
                monthlyQtys.Skip(firstNonZero).ToList()

                Dim totalShipped As Decimal = relevantMonths.Sum()
                Dim monthsCount As Integer = relevantMonths.Count

                row("AVG_MONTH_SALES") = Math.Round(totalShipped / monthsCount, 2)
                row("AVG_WEEK_SALES") = Math.Round(totalShipped / (monthsCount * 4.33D), 2)
            Else
                row("AVG_MONTH_SALES") = 0
                row("AVG_WEEK_SALES") = 0
            End If


            Dim AVG_WEEK_SALES As Decimal = Val(row("AVG_WEEK_SALES") & "")
            Dim ATS_MS As Integer = Val(row("ATS_MS") & "")

            If AVG_WEEK_SALES <= 0 Then
                row("SUG_XFR") = 0
                Continue For
            End If

            Dim targetStock As Integer = CInt(Math.Ceiling(AVG_WEEK_SALES * 4))
            Dim shortage As Integer = targetStock - ATS_MS

            If shortage > 0 Then
                row("SUG_XFR") = shortage
            Else
                row("SUG_XFR") = 0
            End If

            Dim ON_HAND As Decimal = Val(row("ONHD_MS") & "")
            Dim MS_IN_PICK As Decimal = Val(row("PICK_MS") & "")
            Dim MS_OPEN_ORDERS As Decimal = Val(row("OPEN_MS") & "")

            Dim wosNumerator As Decimal = ON_HAND - MS_IN_PICK '- MS_OPEN_ORDERS
            If AVG_WEEK_SALES > 0 Then
                row("WOS_MS") = Math.Round(wosNumerator / AVG_WEEK_SALES, 2)
            Else
                row("WOS_MS") = 0
            End If

        Next


        Sort_grdColumns(grdICTSTYLX, "STYLE_CODE,COLOR_CODE")

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        BeginTrans()
        Dim WHSE_CODE As String = "US"
        Create_Transfer_Order()
        Release_Transfer_Order()

        For Each TABLE_NAME As String In TABLES_OXFR
            Update_Record_TDA(TABLE_NAME)
        Next

        Dim SHIP_BOL_NO As String = dst.Tables("SOTSHIP1").Rows(0).Item("SHIP_BOL_NO")
        Dim ORDR_GROUP_NO As String = dst.Tables("SOTSHIP1").Rows(0).Item("ORDR_GROUP_NO")
        ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})

        For Each rowICTSTYLX As DataRow In dst.Tables("ICTSTYLX").Select("QTY2XFR IS NOT NULL AND QTY2XFR > 0")
            Dim STYLE_CODE As String = rowICTSTYLX.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowICTSTYLX.Item("COLOR_CODE")
            Dim QTY As Int32 = Val(rowICTSTYLX.Item("QTY2XFR"))

            TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", QTY)

            ASCMAIN1.sql = $"Update SOTOXFR1 SET OXFR_STATUS = '1', SHIP_BOL_NO = '{SHIP_BOL_NO}'" & vbCrLf _
                     & "WHERE STYLE_CODE = :PARM1 AND COLOR_CODE = :PARM2 AND OXFR_STATUS = '0'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {STYLE_CODE, COLOR_CODE})
        Next

        ' Incease Qty In PICK by the PICK_QTY for all selected SCs to Transfer

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
            Dim QTY As Int32 = Val(rowSOTORDR2.Item("ORDR_QTY_PICK"))
            TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_PICK", QTY)
        Next

        'Stop
        ' create Sales Order of type XFR - SOTORDR1/2/5 SOTORDR0
        ' release the order - SOTPICK1/2 SOTPICK0, SOTSHIP1, SOTCART1/2

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTSTYLX, "SSB", "Show Filter", "Show GroupBox", "Style Status Inquiry", "Select All", "De-Select All", "Select Selected", "De-Select Selected")
        Load_Popup_Menu(grdSOTORDRX, "SSBB", "Show Filter", "Show GroupBox", "Sales Order Inquiry", "Customer Order Inquiry")

    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing


        Select Case e.SourceControl.Name
            Case "grdICTSTYLX"
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdICTSTYLX"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key

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

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Customer Order Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                If rowARTCUST1 IsNot Nothing Then
                    Context_Launch("Select", CUST_CODE, e.Tool.Key, "SOFCORD1")
                End If

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                    grow.Cells("HIDE_STYLE").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next

            Case "Select Selected", "De-Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("HIDE_STYLE").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                    grow.Update()
                Next

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "CUST_CODE"

        End Select

    End Sub

    Overrides Sub dte_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select

    End Sub

#End Region

    Sub Create_WorkTables(Optional initialize As Boolean = False)

        If initialize Then

            ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO" & vbCrLf _
                & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_OPEN, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME" & vbCrLf _
                & ", SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.SREP_CODE" & vbCrLf _
                & ", SOTORDR1.ORDR_TYPE_CODE, SOTORDR1.ORDR_SOURCE, SOTORDR1.EDI_DOC_SEQ_NO" & vbCrLf _
                & ", SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_DATE_RECD" & vbCrLf _
                & ", SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
                & ", SOTORDR1.ORDR_PRIORITY" & vbCrLf _
                & " from SOTORDR1,SOTORDR2,ARTCUST1" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_STATUS = 'O'" & vbCrLf _
                & "   and SOTORDR1.WHSE_CODE = 'MS'" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                & "" ' "   and (ARTCUST1.CUST_COUNTRY <> 'USA' or ARTCUST1.CUST_COUNTRY <> 'US')"
            '& "   and SOTORDR2.ORDR_QTY_OPEN > 0" & vbCrLf _

            sqlSOTORDRX = ASCMAIN1.sql
            ASCMAIN1.sql = $"{sqlSOTORDRX} and ROWNUM < 1"
            SOTORDRX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL($"Alter Table {SOTORDRX} Add Primary Key (ORDR_NO, ORDR_LNO)")

            ASCMAIN1.sql = "Select C1.STYLE_CODE, C1.COLOR_CODE, X.ORDR_QTY_OPEN" & vbCrLf _
            & " , L1.STYLE_DESC, L1.CARTON_PACK_QTY, L1.INNER_PACK_QTY" & vbCrLf _
            & " , L1.STYLE_STATUS, L1.STYLE_CLASS_CODE, C1.STYLE_COLOR_STATUS, L1.CASE_CUBE, L1.INIT_DATE" & vbCrLf _
            & " , R1.COLOR_DESC" & vbCrLf _
            & " , L1.STYLE_UOM, L1.VEND_CODE" & vbCrLf _
            & " , M.SET_QTY" & vbCrLf _
            & " , XMS.ONHD_MS, XMS.PICK_MS, XMS.OPEN_MS, XMS.ONPO_MS, XMS.TRAN_MS" & vbCrLf _
            & " , XUS.ONHD_US, XUS.PICK_US, XUS.OPEN_US, XUS.ONPO_US, XUS.TRAN_US" & vbCrLf _
            & " , NVL(SALES.AVG_WEEK_SALES,0)    AVG_WEEK_SALES" & vbCrLf _
            & " , NVL(SALES.AVG_MONTH_SALES,0)   AVG_MONTH_SALES" & vbCrLf _
            & " , CASE WHEN NVL(SALES.AVG_WEEK_SALES,0)=0 THEN NULL" & vbCrLf _
            & "        ELSE ROUND((NVL(XMS.ONHD_MS,0)-NVL(XMS.OPEN_MS,0)-NVL(XMS.PICK_MS,0))" & vbCrLf _
            & "                   / SALES.AVG_WEEK_SALES,1) END WOS_MS" & vbCrLf _
            & " , CASE WHEN NVL(SALES.AVG_WEEK_SALES,0)=0 THEN NULL" & vbCrLf _
            & "        ELSE ROUND((NVL(XUS.ONHD_US,0)-NVL(XUS.PICK_US,0))" & vbCrLf _
            & "                   / SALES.AVG_WEEK_SALES,1) END WOS_US" & vbCrLf _
            & " from ICTSTYL1 L1, ICTCOLR1 R1, ICTSTYC1 C1" & vbCrLf _
            & " , (Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & "    , Sum(CASE WHEN ORDR_TYPE_CODE='B2C' THEN ORDR_QTY_OPEN ELSE 0 END) AS ORDR_QTY_OPEN" & vbCrLf _
            & $"   from {SOTORDRX}" & vbCrLf _
            & "  group by STYLE_CODE, COLOR_CODE) X" & vbCrLf _
            & " , (Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & "    , Sum(WHSE_QTY_ON_HAND) ONHD_MS" & vbCrLf _
            & "    , Sum(WHSE_QTY_PICK)      PICK_MS" & vbCrLf _
            & "    , Sum(WHSE_QTY_OPEN)      OPEN_MS" & vbCrLf _
            & "    , Sum(WHSE_QTY_ON_ORDER)  ONPO_MS" & vbCrLf _
            & "    , Sum(WHSE_QTY_TRAN)      TRAN_MS" & vbCrLf _
            & "   from ICTSTAT2 where WHSE_CODE='MS'" & vbCrLf _
            & "  group by STYLE_CODE, COLOR_CODE) XMS" & vbCrLf _
            & " , (Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & "    , Sum(WHSE_QTY_ON_HAND) ONHD_US" & vbCrLf _
            & "    , Sum(WHSE_QTY_PICK)      PICK_US" & vbCrLf _
            & "    , Sum(WHSE_QTY_OPEN)      OPEN_US" & vbCrLf _
            & "    , Sum(WHSE_QTY_ON_ORDER)  ONPO_US" & vbCrLf _
            & "    , Sum(WHSE_QTY_TRAN)      TRAN_US" & vbCrLf _
            & "   from ICTSTAT2 where WHSE_CODE='US'" & vbCrLf _
            & "  group by STYLE_CODE, COLOR_CODE) XUS" & vbCrLf _
            & " , (Select h2.STYLE_CODE, h2.COLOR_CODE" & vbCrLf _
            & "    , ROUND(SUM(h2.ORDR_QTY_SHIP)/78,2) AS AVG_WEEK_SALES" & vbCrLf _
            & "    , ROUND(SUM(h2.ORDR_QTY_SHIP)/18,2) AS AVG_MONTH_SALES" & vbCrLf _
            & "   from SOTINVH1 h1" & vbCrLf _
            & "   join SOTINVH2 h2 on h1.INV_TYPE=h2.INV_TYPE and h1.INV_NO=h2.INV_NO" & vbCrLf _
            & "   join SOTORDR1 o1 on h1.ORDR_NO=o1.ORDR_NO" & vbCrLf _
            & "  where h1.INV_DATE>=ADD_MONTHS(TRUNC(SYSDATE,'MM'),-17)" & vbCrLf _
            & "    and h2.ORDR_QTY_SHIP>0" & vbCrLf _
            & "    and o1.ORDR_TYPE_CODE='B2C'" & vbCrLf _
            & "  group by h2.STYLE_CODE, h2.COLOR_CODE) SALES" & vbCrLf _
            & " , (Select STYLE_CODE, MAX(SET_QTY) AS SET_QTY" & vbCrLf _
            & "   from ECTESTY1" & vbCrLf _
            & "  group by STYLE_CODE) M" & vbCrLf _
            & " where L1.STYLE_CODE=X.STYLE_CODE" & vbCrLf _
            & "   and R1.COLOR_CODE=X.COLOR_CODE" & vbCrLf _
            & "   and C1.STYLE_CODE(+)=X.STYLE_CODE" & vbCrLf _
            & "   and C1.COLOR_CODE(+)=X.COLOR_CODE" & vbCrLf _
            & "   and XMS.STYLE_CODE(+)=X.STYLE_CODE" & vbCrLf _
            & "   and XMS.COLOR_CODE(+)=X.COLOR_CODE" & vbCrLf _
            & "   and XUS.STYLE_CODE(+)=X.STYLE_CODE" & vbCrLf _
            & "   and XUS.COLOR_CODE(+)=X.COLOR_CODE" & vbCrLf _
            & "   and SALES.STYLE_CODE(+)=X.STYLE_CODE" & vbCrLf _
            & "   and SALES.COLOR_CODE(+)=X.COLOR_CODE" & vbCrLf _
            & "   and M.STYLE_CODE(+)=L1.STYLE_CODE" & vbCrLf _
            & "   and X.STYLE_CODE in (Select Distinct STYLE_CODE from ECTESTY1)"



            sqlICTSTYLX = ASCMAIN1.sql
            ASCMAIN1.sql = $"{sqlICTSTYLX} and ROWNUM < 1"
            ICTSTYLX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL($"Alter Table {ICTSTYLX} Add Primary Key (STYLE_CODE, COLOR_CODE)")

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("SELECT e.ECOM_CODE, e.ECOM_NAME")
            sb.AppendLine(", h2.STYLE_CODE, h2.COLOR_CODE")
            sb.AppendLine(", SUM(h2.ORDR_QTY_SHIP) QTY_SHIPPED_TOT")
            sb.AppendLine(", ROUND(SUM(h2.ORDR_QTY_SHIP) / 78, 2) AS AVG_WEEK_SALES")
            sb.AppendLine(", ROUND(SUM(h2.ORDR_QTY_SHIP) / 18, 2) AS AVG_MONTH_SALES")
            sb.AppendLine(", y.SET_QTY")
            sb.AppendLine(", y.SHIP_DROP")

            For i As Integer = 0 To 17
                Dim colAlias As String = $"MTH_{(i + 1).ToString("D2")}_QTY"
                Dim monthExpr As String = $"TO_CHAR(ADD_MONTHS(TRUNC(SYSDATE, 'MM'), -{i}), 'YYYYMM')"
                sb.AppendLine($", ROUND(SUM(CASE WHEN TO_CHAR(h1.INV_DATE, 'YYYYMM') = {monthExpr} THEN h2.ORDR_QTY_SHIP ELSE 0 END), 0) AS {colAlias}")
            Next

            sb.AppendLine("FROM SOTINVH1 h1, SOTINVH2 h2, ECTECOM1 e, ECTESTY1 y")
            sb.AppendLine("WHERE h1.INV_TYPE = h2.INV_TYPE")
            sb.AppendLine("AND h1.INV_NO   = h2.INV_NO")
            sb.AppendLine("AND h1.CUST_CODE = e.CUST_CODE")
            sb.AppendLine("AND h1.INV_DATE >= ADD_MONTHS(TRUNC(SYSDATE, 'MM'), -17)")
            sb.AppendLine("AND h2.ORDR_QTY_SHIP > 0")
            sb.AppendLine("AND y.STYLE_CODE = h2.STYLE_CODE")
            sb.AppendLine("AND y.ECOM_CODE = e.ECOM_CODE")
            sb.AppendLine("GROUP BY e.ECOM_CODE, e.ECOM_NAME, h2.STYLE_CODE, h2.COLOR_CODE, y.SET_QTY, y.SHIP_DROP")

            ASCMAIN1.sql = sb.ToString()

            sqlSOTINVHX = ASCMAIN1.sql
            ASCMAIN1.sql = "SELECT * FROM (" & sqlSOTINVHX & ") WHERE ROWNUM < 1"
            SOTINVHX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL($"ALTER TABLE {SOTINVHX} ADD PRIMARY KEY (ECOM_CODE, STYLE_CODE, COLOR_CODE)")

        Else

            ASCDATA1.ExecuteSQL($"Truncate Table {SOTORDRX}")
            ASCDATA1.ExecuteSQL($"Insert into {SOTORDRX} {sqlSOTORDRX}")

            ASCDATA1.ExecuteSQL($"Truncate Table {ICTSTYLX}")
            ASCDATA1.ExecuteSQL($"Insert into {ICTSTYLX} {sqlICTSTYLX}")

            ASCDATA1.ExecuteSQL($"TRUNCATE TABLE {SOTINVHX}")
            ASCDATA1.ExecuteSQL($"INSERT INTO {SOTINVHX} {sqlSOTINVHX}")

        End If
    End Sub

    Private Sub grdICTSTYLX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTSTYLX.AfterRowActivate
        If grdICTSTYLX.ActiveRow Is Nothing OrElse Not grdICTSTYLX.ActiveRow.IsDataRow Then
            SplitContainer1.Panel2Collapsed = True
        Else
            Setup_SOTORDRX()
            Setup_SOTINVHX()
            Setup_POTORDRX()
            SplitContainer1.Panel2Collapsed = False
        End If
    End Sub
    Sub Setup_SOTORDRX()
        If grdICTSTYLX.ActiveRow Is Nothing OrElse Not grdICTSTYLX.ActiveRow.IsDataRow Then
            grdSOTORDRX.Visible = False
        Else
            Dim STYLE_CODE As String = grdICTSTYLX.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdICTSTYLX.ActiveRow.Cells("COLOR_CODE").Value
            Fill_Records("SOTORDRX", New String() {STYLE_CODE, COLOR_CODE})

            Dim dv As DataView = dst.Tables("SOTORDRX").DefaultView
            dv.RowFilter = "ORDR_TYPE_CODE = 'B2C'"
            grdSOTORDRX.DataSource = dv
            grdSOTORDRX.Visible = True
            grdSOTORDRX.Text = $"Open Ecommerce Orders for SC {STYLE_CODE}-{COLOR_CODE}"

        End If
    End Sub
    Sub Setup_POTORDRX()
        If grdICTSTYLX.ActiveRow Is Nothing OrElse Not grdICTSTYLX.ActiveRow.IsDataRow Then
            grdPOTORDRX.Visible = False
        Else

            Dim STYLE_CODE As String = grdICTSTYLX.ActiveRow.Cells("STYLE_CODE").Text
            Dim COLOR_CODE As String = grdICTSTYLX.ActiveRow.Cells("COLOR_CODE").Text
            Dim DEF_DAYS As String = Val(ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETA_TO_ARR") & "")

            ASCMAIN1.sql =
            "SELECT * FROM (" & vbCrLf &
            "SELECT POTORDR1.INIT_DATE, POTSHIP1.WHSE_CODE, POTSHIP3.PO_ORDER_NO, " &
            "POTORDR1.PO_DATE_SHIP_BY PO_DATE_SHIP_BY_REQ, POTORDR2.PO_DATE_SHIP_BY, " &
            "POTORDR1.FACTORY_CODE, POTSHIP3.PO_ORDER_LNO, POTSHIP2.PO_SHIPMENT_NO, " &
            "POTSHIP2.PO_SHIPMENT_LNO, POTSHIP1.PO_SHIP_VESSEL, POTSHIP1.PO_DATE_SHIPPED, " &
            "POTSHIP1.PO_SHIP_ETA, POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS, " &
            "POTSHIP1.PO_SHIP_REF_NO, POTSHIP2.CONTAINER_NO, POTSHIP2.PO_DATE_RECEIVED, " &
            "POTSHIP3.PO_QTY_SHP, POTSHIP3.PO_QTY_REC, " &
            "POTORDR1.VEND_CODE, POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO, " &
            "POTORDR2.PO_QTY_ORD, 0 PO_QTY_OPN, " &
            "POTSHIP1.PO_SHIP_ETA + NVL(POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS, 0) AS PO_ARRIVAL_DATE, " &
            "POTORDR2.LAST_DATE_SHIP_BY, POTORDR2.LAST_OPER_SHIP_BY, " &
            "ICTATOP2.STYLE_ARRIVAL_BUFFER_DAYS, ICTATOP2.STYLE_AT_ONCE_UNTIL, ICTATOP2.STYLE_AT_ONCE_ACTIVE " &
            "FROM POTSHIP1, POTSHIP2, POTSHIP3, POTORDR1, POTORDR2, ICTATOP2 " &
            "WHERE POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO " &
            "AND POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO " &
            "AND POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO " &
            "AND POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO " &
            "AND POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO " &
            "AND POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO " &
            "AND POTORDR2.STYLE_CODE = '" & STYLE_CODE & "' " &
            "AND POTORDR2.COLOR_CODE = '" & COLOR_CODE & "' " &
            "AND ICTATOP2.STYLE_CODE (+) = '" & STYLE_CODE & "' " &
            "AND ICTATOP2.COLOR_CODE (+) = '" & COLOR_CODE & "' " &
            "AND ICTATOP2.PS_CODE (+) = 'S' " &
            "AND ICTATOP2.PS_NO (+) = POTSHIP3.PO_SHIPMENT_NO " &
            "AND POTSHIP2.PO_SHIP_STATUS = 'O' " &
            ") UNION (" & vbCrLf &
            "SELECT POTORDR1.INIT_DATE, POTORDR1.WHSE_CODE, POTORDR2.PO_ORDER_NO, " &
            "POTORDR1.PO_DATE_SHIP_BY PO_DATE_SHIP_BY_REQ, POTORDR2.PO_DATE_SHIP_BY, " &
            "POTORDR1.FACTORY_CODE, POTORDR2.PO_ORDER_LNO, NULL PO_SHIPMENT_NO, 0 PO_SHIPMENT_LNO, " &
            "DECODE(NVL(POTORDR2.PO_QTY_OPN,0),0,'ClosedPO','OpenPO') PO_SHIP_VESSEL, " &
            "POTORDR2.PO_DATE_SHIP_BY, POTORDR2.PO_DATE_ETA, " &
            Val(ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETA_TO_ARR") & "") & " PO_SHIP_LANDING_LEAD_DAYS, " &
            "NULL PO_SHIP_REF_NO, NULL CONTAINER_NO, NULL PO_DATE_RECEIVED, " &
            "0 PO_QTY_SHP, 0 PO_QTY_REC, " &
            "POTORDR1.VEND_CODE, POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO, " &
            "POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN, " &
            "POTORDR2.PO_DATE_ETA + " & Val(ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETA_TO_ARR") & "") & " PO_ARRIVAL_DATE, " &
            "POTORDR2.LAST_DATE_SHIP_BY, POTORDR2.LAST_OPER_SHIP_BY, " &
            "ICTATOP2.STYLE_ARRIVAL_BUFFER_DAYS, ICTATOP2.STYLE_AT_ONCE_UNTIL, ICTATOP2.STYLE_AT_ONCE_ACTIVE " &
            "FROM POTORDR1, POTORDR2, ICTATOP2 " &
            "WHERE POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO " &
            "AND POTORDR2.STYLE_CODE = '" & STYLE_CODE & "' " &
            "AND POTORDR2.COLOR_CODE = '" & COLOR_CODE & "' " &
            "AND ICTATOP2.STYLE_CODE (+) = '" & STYLE_CODE & "' " &
            "AND ICTATOP2.COLOR_CODE (+) = '" & COLOR_CODE & "' " &
            "AND ICTATOP2.PS_CODE (+) = 'P' " &
            "AND ICTATOP2.PS_NO (+) = POTORDR2.PO_ORDER_NO " &
            "AND POTORDR2.PO_QTY_OPN <> 0)"

            Fill_Records("POTORDRX", "", True, ASCMAIN1.sql)

            grdPOTORDRX.Text = $"Open POs for SC {STYLE_CODE}-{COLOR_CODE}"
        End If
    End Sub

    Sub Setup_SOTINVHX()
        If grdICTSTYLX.ActiveRow Is Nothing Then Exit Sub

        Dim STYLE_CODE As String = grdICTSTYLX.ActiveRow.Cells("STYLE_CODE").Text
        Dim COLOR_CODE As String = grdICTSTYLX.ActiveRow.Cells("COLOR_CODE").Text
        Fill_Records("SOTINVHX", New String() {STYLE_CODE, COLOR_CODE})

        grdSOTINVHX.Text = $"eComm Partner Sales (Last 18 Months) – {STYLE_CODE}-{COLOR_CODE}"

        If grdSOTINVHX.DisplayLayout.Bands.Count > 0 Then
            Dim band = grdSOTINVHX.DisplayLayout.Bands(0)
            Dim baseDate As Date = DateSerial(Year(Today), Month(Today), 1)

            For i As Integer = 0 To 17
                Dim colKey As String = $"MTH_{(i + 1).ToString("D2")}_QTY"
                If band.Columns.Exists(colKey) Then
                    Dim labelDate As Date = DateAdd(DateInterval.Month, -i, baseDate)
                    With band.Columns(colKey)
                        .Header.Caption = labelDate.ToString("MMM yy") & " Shp"
                        .Format = "###0"
                        .Width = 80
                    End With
                End If
            Next
        End If

        For Each row As DataRow In dst.Tables("SOTINVHX").Rows
            Dim monthlyQtys As New List(Of Decimal)

            For i As Integer = 1 To 18
                Dim qty As Decimal = Val(row($"MTH_{i:D2}_QTY") & "")
                monthlyQtys.Add(qty)
            Next

            Dim firstNonZero As Integer = monthlyQtys.FindIndex(Function(q) q <> 0)

            Dim relevantMonths As List(Of Decimal)
            If firstNonZero >= 0 Then
                relevantMonths = monthlyQtys.Skip(firstNonZero).ToList()
            Else
                relevantMonths = New List(Of Decimal)
            End If

            Dim totalShipped As Decimal = relevantMonths.Sum()
            Dim nonzeroMonths As Integer = relevantMonths.Count

            If nonzeroMonths > 0 Then
                row("AVG_MONTH_SALES") = Math.Round(totalShipped / nonzeroMonths, 2)
                row("AVG_WEEK_SALES") = Math.Round(totalShipped / (nonzeroMonths * 4.33D), 2)
            Else
                row("AVG_MONTH_SALES") = DBNull.Value
                row("AVG_WEEK_SALES") = DBNull.Value
            End If
        Next
    End Sub
    Sub Record_Event(EVENT_TYPE As String, EVENT_DESC As String, ORDR_NO As String)
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

        ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow

        ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        For Each rowSOTORDR2 As DataRow In ASCDATA1.GetDataTable.Rows
            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")

            QTY_TO_COMMIT = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
            If QTY_TO_COMMIT <> 0 Then
                STYLE_CODE = rowSOTORDR2.Item("STYLE_CODE")
                COLOR_CODE = rowSOTORDR2.Item("COLOR_CODE")
                TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", S * QTY_TO_COMMIT)
            End If
        Next

    End Sub
    Sub Auto_Populate()
        'Dim ANY_POPULATED As Boolean = False
        'For Each row As DataRow In dst.Tables("ICTSTYLX").Select("QTY_SHORT < 0")
        '    Dim QTY_SHORT As Integer = Math.Abs(Val(row("QTY_SHORT") & ""))
        '    Dim AVA_US As Integer = Val(row("AVA_US") & "")
        '    Dim INNER_PACK_QTY As Integer = Val(row("INNER_PACK_QTY") & "")
        '    Dim CARTON_PACK_QTY As Integer = Val(row("CARTON_PACK_QTY") & "")

        '    If QTY_SHORT > AVA_US Then
        '        row("QTY2XFR") = DBNull.Value
        '        Continue For
        '    End If

        '    Dim QTY2XFR As Integer = QTY_SHORT

        '    If INNER_PACK_QTY > 0 Then
        '        QTY2XFR = Math.Ceiling(QTY_SHORT / INNER_PACK_QTY) * INNER_PACK_QTY
        '    ElseIf CARTON_PACK_QTY > 0 Then
        '        QTY2XFR = Math.Ceiling(QTY_SHORT / CARTON_PACK_QTY) * CARTON_PACK_QTY
        '    End If

        '    If QTY2XFR <= AVA_US Then
        '        row("QTY2XFR") = QTY2XFR
        '        ANY_POPULATED = True
        '    Else
        '        row("QTY2XFR") = DBNull.Value
        '    End If
        'Next

        'If ANY_POPULATED Then
        '    MsgBox("Transfer quantities auto-populated based on shortage and available US inventory.", MsgBoxStyle.Information)
        'Else
        '    MsgBox("No rows qualified for auto-population.", MsgBoxStyle.Exclamation)
        'End If

    End Sub
    Sub Create_Transfer_Order()
        ' SOTORDR1 SOTORDR2 SOTORDR5 SOTORDR0 ICTSTAT2

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Building Order from Records Selected in Transfer Queue", "")


        Dim ORDR_LNO As Integer = 0
        Dim ORDR_NO As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")

        Dim CUST_CODE As String = "180000"
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
        Dim CUST_NAME As String = rowARTCUST1.Item("CUST_NAME")
        Dim CUST_BILL_TO_CUST As String = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
        If CUST_BILL_TO_CUST = "" Then
            CUST_BILL_TO_CUST = CUST_CODE
        End If

        Dim CUST_STORE_NO As String = "000027"
        Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO})
        Dim CUST_STORE_NAME As String = rowARTCUST2.Item("CUST_NAME")

        Dim ORDR_CUST_PO As String = $"XFR {ORDR_NO}"
        Dim ORDR_SHIP_DATE As Date = Now.Date
        Dim ORDR_CANCEL_DATE As Date = Now.Date.AddDays(7)

        Dim rowARTCUST1_BT As DataRow = LookUp("ARTCUST1", CUST_BILL_TO_CUST)

        Dim WHSE_CODE As String = "US"
        For Each rowICTSTYLX As DataRow In dst.Tables("ICTSTYLX").Select("QTY2XFR IS NOT NULL AND QTY2XFR > 0")
            Dim STYLE_CODE As String = rowICTSTYLX.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowICTSTYLX.Item("COLOR_CODE")
            Dim ORDR_QTY_OPEN As Int32 = Val(rowICTSTYLX.Item("QTY2XFR"))

            Dim ORDR_UNIT_PRICE As Decimal = 0
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            Dim STYLE_DESC As String = rowICTSTYLX.Item("STYLE_DESC")

            ORDR_LNO += 1

            Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").NewRow
            With rowSOTORDR2
                .Item("ORDR_NO") = ORDR_NO
                .Item("ORDR_LNO") = ORDR_LNO
                .Item("STYLE_CODE") = STYLE_CODE
                .Item("COLOR_CODE") = COLOR_CODE
                .Item("STYLE_DESC") = STYLE_DESC
                .Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE
                .Item("ORDR_UNIT_PRICE_CURR") = ORDR_UNIT_PRICE
                .Item("ORDR_QTY") = ORDR_QTY_OPEN
                .Item("ORDR_QTY_OPEN") = ORDR_QTY_OPEN
                .Item("ORDR_QTY_ORIG") = ORDR_QTY_OPEN
                .Item("ORDR_QTY_ALLO") = 0
                .Item("INNER_PACK_QTY") = rowICTSTYLX("INNER_PACK_QTY")
                .Item("ORDR_EXTD_COST") = 0
                .Item("STYLE_UOM") = rowICTSTYLX("STYLE_UOM")
                .Item("ORDR_QTY_PICK") = 0
                .Item("ORDR_QTY_SHIP") = 0
                .Item("ORDR_QTY_CANC") = 0
                .Item("ORDR_STATUS") = "O"
                .Item("ORDR_QTY_PRE_ALLO") = 0
                .Item("QTY_PER_PP") = 0
                .Item("CARTON_PACK_QTY") = rowICTSTYLX("CARTON_PACK_QTY")
                .Item("STYLE_PRICE") = 0
                .Item("ORDR_UNIT_PRICE_CALC") = 0
                .Item("ORDR_UNIT_PRICE_MANUAL") = ""
                .Item("STYLE_RETAIL") = 0
                .Item("PO_COST") = 0
                .Item("COMM_RATE") = 0
            End With
            dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)
        Next

        ADD_SOTORDR5(ORDR_NO, CUST_CODE, "BT", CUST_CODE, rowARTCUST1)
        ADD_SOTORDR5(ORDR_NO, CUST_CODE, "ST", CUST_STORE_NO, rowARTCUST2)


        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)

        Dim ORDR_FOB As String = rowICTWHSE1.Item("WHSE_CITY") & "," & rowICTWHSE1.Item("WHSE_STATE")

        Dim ORDR_GROUP_NO As String = ORDR_NO

        Dim rowSOTORDR1 As DataRow = Nothing
        rowSOTORDR1 = dst.Tables("SOTORDR1").NewRow
        With rowSOTORDR1
            .Item("ORDR_NO") = ORDR_NO
            .Item("ORDR_DATE") = DATETIME_STAMP.Date
            .Item("CUST_CODE") = CUST_CODE
            .Item("CUST_NAME") = CUST_NAME
            .Item("CUST_STORE_NO") = CUST_STORE_NO
            .Item("CUST_STORE_NAME") = CUST_STORE_NAME
            .Item("ORDR_FOB") = ORDR_FOB
            .Item("ORDR_CUST_PO") = ORDR_CUST_PO
            .Item("ORDR_SHIP_DATE") = ORDR_SHIP_DATE
            .Item("ORDR_CANCEL_DATE") = ORDR_CANCEL_DATE
            .Item("POST_CODE") = rowARTCUST1_BT.Item("POST_CODE")
            .Item("TERM_CODE") = rowARTCUST1_BT.Item("TERM_CODE")
            .Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
            .Item("SHIP_VIA_CODE") = rowARTCUST1.Item("SHIP_VIA_CODE")

            .Item("SREP2_CODE") = ""
            .Item("WHSE_CODE") = WHSE_CODE
            .Item("WHSE_CODE_TO") = "MS"
            .Item("SALES_DIVISION_CODE") = ""
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("ORDR_DATE_RECD") = DATETIME_STAMP.Date
            .Item("ORDR_SOURCE") = "X"
            .Item("FRT_TERMS") = rowARTCUST1.Item("FRT_TERMS")
            .Item("ORDR_ADDR_TYPE_ST") = "MK"
            .Item("ORDR_DATE_BOOKED") = DATETIME_STAMP.Date
            .Item("ORDR_YYYYPP_BOOKED") = ASCMAIN1.CYP
            .Item("ORDR_STATUS") = "O"
            .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
            .Item("ORDR_HOLD") = "0"
            .Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST
            .Item("CUST_FACTOR_IND") = "0"
            .Item("CURR_CODE") = "USD"
            .Item("CURR_EXCH_RATE") = 1
            .Item("ORDR_ORIG_SHIP_DATE") = ORDR_SHIP_DATE
            .Item("ORDR_ORIG_CANCEL_DATE") = ORDR_CANCEL_DATE
            .Item("ORDR_TYPE_CODE") = "XFR"
        End With
        dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub
    Sub Release_Transfer_Order()
        ' SOTPICK1 SOTPICK2 SOTPICK0 SOTSHIP1 SOTCART1 SOTCAR2 ICTSTAT2

        ' Create Pick Tickets, Shipment BOL, and Cartons
        '   do this for the XFR Sales Order just generated

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Releasing Order", "")

        Dim SHIP_BOL_NO As String = ASCMAIN1.Next_Control_No("SOTSHIP1.SHIP_BOL_NO")

        Dim PICK_RELEASED As Date = DATETIME_STAMP
        Dim PICK_NO_seq As Int32 = 0

        Dim CUST_CODE As String = ""
        Dim ORDR_GROUP_NO As String = ""

        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
            ' should be only 1 order

            CUST_CODE = rowSOTORDR1.Item("CUST_CODE")
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)

            Dim PICK_NO As String = ASCMAIN1.Next_Control_No("SOTPICK1.PICK_NO")

            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
            ORDR_GROUP_NO = ORDR_NO

            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")
            Dim SHIP_VIA_CODE As String = rowSOTORDR1.Item("SHIP_VIA_CODE") & ""
            Dim PICK_SEQ_NO As Integer = Val(rowSOTORDR1.Item("ORDR_PICK_SEQ") & "") + 1
            rowSOTORDR1.Item("ORDR_PICK_SEQ") = PICK_SEQ_NO

            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").NewRow
            PICK_NO_seq += 1

            With rowSOTPICK1
                .Item("PICK_NO") = "E-" & PICK_NO
                .Item("ORDR_NO") = ORDR_NO
                .Item("ORDR_PICK_SEQ") = PICK_SEQ_NO
                .Item("PICK_STATUS") = "P"
                .Item("PICK_RELEASED") = PICK_RELEASED
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                .Item("SHIP_VIA_CODE") = SHIP_VIA_CODE
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("CCPA_NO_STATUS") = "0"

                .Item("SHIP_BOL_NO") = SHIP_BOL_NO
            End With

            dst.Tables("SOTPICK1").Rows.Add(rowSOTPICK1)

            Dim TOTAL_OPEN As Int64 = 0 ' Total Units left OPEN in Order after Release
            Dim TOTAL_PICK As Int64 = 0 ' Total Units in PICK in Order after Release

            Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "' and ORDR_QTY_OPEN <> 0"
            For Each rowSOTORDR2_rel As DataRow In dst.Tables("SOTORDR2").Select(sqlw, "ORDR_LNO")
                Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").NewRow
                With rowSOTPICK2
                    .Item("PICK_NO") = PICK_NO
                    Dim ORDR_LNO As Int32 = Val(rowSOTORDR2_rel.Item("ORDR_LNO"))
                    Dim PICK_LNO As Int32 = ORDR_LNO
                    .Item("PICK_LNO") = PICK_LNO
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("ORDR_LNO") = ORDR_LNO

                    '.Item("STYLE_CODE") = rowSOTORDR2_rel.Item("STYLE_CODE")
                    '.Item("COLOR_CODE") = rowSOTORDR2_rel.Item("COLOR_CODE")
                    .Item("PICK_UNIT_PRICE") = rowSOTORDR2_rel.Item("ORDR_UNIT_PRICE")

                    Dim qCANC As Int64 = 0
                    Dim qBACK As Int64 = 0

                    'Dim qA As Int64 = Val(rowSOTORDR2_rel.Item("ORDR_QTY_ALLO_CUR") & "")
                    Dim qO As Int64 = Val(rowSOTORDR2_rel.Item("ORDR_QTY_OPEN") & "")

                    rowSOTORDR2_rel.Item("ORDR_QTY_PICK") = qO ' Val(rowSOTORDR2_rel.Item("ORDR_QTY_PICK") & "") + qa
                    .Item("PICK_QTY") = qO
                    TOTAL_PICK = TOTAL_PICK + qO

                    rowSOTORDR2_rel.Item("ORDR_QTY_OPEN") = 0
                    'rowSOTORDR2_rel.Item("ORDR_QTY_ALLO_CUR") = 0

                    .Item("PICK_QTY_CANC_REL") = qCANC
                    .Item("PICK_QTY_BACK_REL") = qBACK

                End With

                dst.Tables("SOTPICK2").Rows.Add(rowSOTPICK2)
            Next

            'If TOTAL_OPEN = 0 Then
            rowSOTORDR1.Item("ORDR_STATUS") = "P"
            rowSOTORDR1.Item("ORDR_CUST_PO") = PICK_NO

            'End If


            Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").NewRow
            With rowSOTSHIP1
                .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                .Item("SHIP_VIA_CODE") = rowSOTORDR1.Item("SHIP_VIA_CODE")
                .Item("SHIP_ADDR_TYPE") = "MK" ' ORDR_ADDR_TYPE_ST
                .Item("SHIP_ADDR_CODE") = rowSOTORDR1.Item("CUST_STORE_NO")
                .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                .Item("SHIP_STATUS") = "P"
                .Item("TERM_CODE") = rowSOTORDR1.Item("TERM_CODE")
                .Item("SREP_CODE") = rowSOTORDR1.Item("SREP_CODE")
                .Item("FRT_TERMS") = rowSOTORDR1.Item("FRT_TERMS")
                .Item("WHSE_CODE") = rowSOTORDR1.Item("WHSE_CODE")
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LP_STATUS") = DBNull.Value
                '.Item("ORDR_PICK_TYPE") = ORDR_PICK_TYPE
                '.Item("SHIP_CART_REQD") = SHIP_CART_REQD
                .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT")
            End With
            dst.Tables("SOTSHIP1").Rows.Add(rowSOTSHIP1)



            Dim PICK_BATCH_NO As String = ASCMAIN1.Next_Control_No("SOTPICK0.PICK_BATCH_NO")
            Dim rowSOTPICK0 As DataRow = dst.Tables("SOTPICK0").NewRow
            With rowSOTPICK0
                .Item("PICK_BATCH_NO") = PICK_BATCH_NO
                .Item("PICK_SHPS") = 1
                .Item("PICK_CTNS") = 1
                .Item("PICK_PKTS") = PICK_NO_seq
                .Item("PICK_BATCH_STATUS") = "O"
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("PICK_SHIP_REL_DATE") = Now.Date
            End With
            dst.Tables("SOTPICK0").Rows.Add(rowSOTPICK0)

            rowSOTPICK1.Item("PICK_BATCH_NO") = PICK_BATCH_NO
        Next

        Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")
        dst.Tables("SOTPICK2").Columns.Add("SHIP_BOL_NO", GetType(System.String), "PARENT(SOTPICK1_SOTPICK2).SHIP_BOL_NO")
        dst.Tables("SOTPICK2").Columns.Add("PICK_AMT", GetType(System.Decimal), "ISNULL(PICK_QTY,0)*ISNULL(PICK_UNIT_PRICE,0)")
        dst.Tables("SOTPICK1").Columns.Add("PICK_AMT", GetType(System.Decimal), "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_AMT)")
        dst.Tables("SOTPICK1").Columns.Add("PICK_QTY", GetType(System.Decimal), "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY)")

#Region "Standard Cartonization"

        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
            Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
            Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")

            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)

            Dim CART_LNO_seq As Int64 = 0
            Dim CART_NO_seq As Int32 = 0

            Dim CART_NO As String = New_Carton(PICK_NO, CART_NO_seq) : CART_LNO_seq = 0


            Dim sortby As String = "PICK_LNO"
            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}'", sortby)
                Dim ORDR_LNO As Int32 = Val(rowSOTPICK2.Item("ORDR_LNO") & "")
                Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                'Dim CARTON_PACK_QTY As Int64 = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")


                Dim QTY_TO_PACK As Int32 = Val(rowSOTPICK2.Item("PICK_QTY") & "")

                CART_LNO_seq = CART_LNO_seq + 1
                Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow
                With rowSOTCART2
                    .Item("CART_NO") = CART_NO
                    .Item("CART_LNO") = CART_LNO_seq
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("ORDR_LNO") = ORDR_LNO
                    .Item("STYLE_CODE") = STYLE_CODE
                    .Item("COLOR_CODE") = COLOR_CODE
                    .Item("QTY_PACKED") = QTY_TO_PACK
                    .Item("QTY_REL") = QTY_TO_PACK
                End With
                dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)
            Next

        Next
#End Region

        Create_Relation("SOTCART1", "SOTCART2", "CART_NO")
        'dst.Tables("SOTCART2").Columns.Add("WGT", GetType(System.Decimal), "ISNULL(QTY_PACKED,0) * ISNULL(STYLE_WEIGHT,0)")
        dst.Tables("SOTCART1").Columns.Add("QTY", GetType(System.Int64), "SUM(CHILD(SOTCART1_SOTCART2).QTY_PACKED)")
        'dst.Tables("SOTCART1").Columns.Add("WGT", GetType(System.Int64), "SUM(CHILD(SOTCART1_SOTCART2).WGT)")
        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
            rowSOTCART1.Item("CART_TOTAL_UNITS") = rowSOTCART1.Item("QTY")
            rowSOTCART1.Item("CART_TOTAL_UNITS_REL") = rowSOTCART1.Item("QTY")
            'rowSOTCART1.Item("CART_TOTAL_WGT_CALC") = rowSOTCART1.Item("WGT")
        Next

        Create_Relation("SOTPICK1", "SOTCART1", "PICK_NO")
        dst.Tables("SOTPICK1").Columns.Add("CTNS", GetType(System.Int64), "COUNT(CHILD(SOTPICK1_SOTCART1).CART_NO)")
        dst.Tables("SOTPICK1").Columns.Add("WGT", GetType(System.Int64), "SUM(CHILD(SOTPICK1_SOTCART1).CART_TOTAL_WGT_CALC)")
        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
            rowSOTPICK1.Item("PICK_CNT_CARTONS") = rowSOTPICK1.Item("CTNS")
            rowSOTPICK1.Item("PICK_TOTAL_WGT") = rowSOTPICK1.Item("WGT")
        Next

        Create_Relation("SOTSHIP1", "SOTPICK1", "SHIP_BOL_NO")
        dst.Tables("SOTPICK1").Columns.Add("PICKS_C", GetType(System.Int64), "IIF(PICK_STATUS = 'C',1,0)")
        dst.Tables("SOTPICK1").Columns.Add("PICKS_P", GetType(System.Int64), "IIF(PICK_STATUS = 'P',1,0)")
        dst.Tables("SOTSHIP1").Columns.Add("PICKS_C", GetType(System.Int64), "SUM(CHILD(SOTSHIP1_SOTPICK1).PICKS_C)")
        dst.Tables("SOTSHIP1").Columns.Add("PICKS_P", GetType(System.Int64), "SUM(CHILD(SOTSHIP1_SOTPICK1).PICKS_P)")

        'For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("PICKS_C >0 AND PICKS_P =0")
        '    rowSOTSHIP1.Item("SHIP_STATUS") = "C"
        'Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub
    Sub ADD_SOTORDR5(ORDR_NO As String, CUST_CODE As String, CUST_ADDR_TYPE As String, CUST_ADDR_CODE As String, row As DataRow)
        Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").NewRow
        With rowSOTORDR5
            .Item("ORDR_NO") = ORDR_NO
            .Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
            .Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
            For Each COLUMN_NAME As String In New String() _
                {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY"}
                .Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
            Next
        End With
        dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)
    End Sub
    Function New_Carton(PICK_NO As String, ByRef CART_NO_seq As Int32) As String
        Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow
        CART_NO_seq += 1
        Dim CART_NO_ctl As String = ASCMAIN1.Next_Control_No("SOTCART1.CART_NO")
        Dim CART_NO As String = TAC.SOCMAIN1.UPC(Me, CART_NO_ctl, "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))

        rowSOTCART1.Item("CART_NO") = CART_NO
        rowSOTCART1.Item("PICK_NO") = PICK_NO
        rowSOTCART1.Item("CART_TOTAL_UNITS") = 0
        rowSOTCART1.Item("CART_TOTAL_WGT_CALC") = 0
        dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)
        Return CART_NO
    End Function
    Sub Save_Progress()
        Dim INIT_OPER As String = ASCMAIN1.USER_ID
        Dim INIT_DATE As Date = DATETIME_STAMP

        ASCMAIN1.sql = $"DELETE FROM ICTXFRM2 WHERE INIT_OPER = :PARM1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New String() {INIT_OPER})

        For Each row As DataRow In dst.Tables("ICTSTYLX").Select("QTY2XFR > 0 OR HIDE_STYLE = 1")
            Dim STYLE As String = row("STYLE_CODE") & ""
            Dim COLOR As String = row("COLOR_CODE") & ""
            Dim QTY As Integer = Val(row("QTY2XFR") & "")
            Dim CHK As String = If(row("HIDE_STYLE") & "" = "1", "1", "0")

            ASCMAIN1.sql = $"INSERT INTO ICTXFRM2 (STYLE_CODE, COLOR_CODE, CHECKED, INIT_DATE, INIT_OPER, TRANSFER_QTY)
                         VALUES (:PARM1, :PARM2, :PARM3, :PARM4, :PARM5, :PARM6)"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVVDV", New Object() {STYLE, COLOR, CHK, INIT_DATE, INIT_OPER, QTY})
        Next

        MsgBox("Progress saved.", MsgBoxStyle.Information)
    End Sub
    Sub Load_From_Saved()
        Dim INIT_OPER As String = ASCMAIN1.USER_ID
        Dim TBL As String = "ICTXFRSAVE1"

        ASCMAIN1.sql = $"SELECT STYLE_CODE, COLOR_CODE, TRANSFER_QTY, CHECKED FROM {TBL} WHERE INIT_OPER = '{INIT_OPER}'"
        Dim dtSaved As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        If dtSaved.Rows.Count = 0 Then
            MsgBox("No saved progress found for your user.", MsgBoxStyle.Information)
            Exit Sub
        End If

        Dim updated As Integer = 0
        For Each savedRow As DataRow In dtSaved.Rows
            Dim STYLE As String = savedRow("STYLE_CODE") & ""
            Dim COLOR As String = savedRow("COLOR_CODE") & ""

            Dim foundRows() As DataRow = dst.Tables("ICTSTYLX").Select($"STYLE_CODE = '{STYLE}' AND COLOR_CODE = '{COLOR}'")

            If foundRows.Length > 0 Then
                Dim row As DataRow = foundRows(0)

                row("QTY2XFR") = savedRow("TRANSFER_QTY")
                row("HIDE_STYLE") = (savedRow("CHECKED") & "") = "1"
                updated += 1
            End If
        Next

        MsgBox($"{updated} style-color rows loaded from saved progress.", MsgBoxStyle.Information)
    End Sub

    Private Sub chkHideStyles_CheckedChanged(sender As Object, e As EventArgs) Handles chkHideStyles.CheckedChanged
        Dim dv As DataView = dst.Tables("ICTSTYLX").DefaultView

        If chkHideStyles.Checked Then
            dv.RowFilter = "ISNULL(HIDE_STYLE, 0) = 0"
        Else
            dv.RowFilter = ""
        End If

        grdICTSTYLX.DataSource = dv
    End Sub
    Private Sub grdICTSTYLX_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYLX.AfterCellUpdate

        If e.Cell.Column.Key = "QTY2XFR" Then
            Dim CARTON_PACK_QTY As Integer = 0
            If Not IsDBNull(e.Cell.Row.Cells("CARTON_PACK_QTY").Value) Then
                CARTON_PACK_QTY = CInt(e.Cell.Row.Cells("CARTON_PACK_QTY").Value)
            End If

            If CARTON_PACK_QTY > 0 Then
                Dim entered As Integer
                If Integer.TryParse(e.Cell.Text, entered) Then
                    Dim rounded As Integer = CInt(Math.Ceiling(entered / CARTON_PACK_QTY) * CARTON_PACK_QTY)
                    If rounded <> entered Then
                        e.Cell.Value = rounded
                    End If
                End If
            End If
            Dim QTY2XFR As Integer = If(e.Cell.Value Is Nothing, 0, CInt(e.Cell.Value))
            Dim TOTAL_UNITS = e.Cell.Row.Cells("TOTAL_UNITS").Value
            Dim ATS_US = e.Cell.Row.Cells("AVA_US").Value

            Dim totCell = e.Cell.Row.Cells("TOTAL_UNITS")
            If QTY2XFR > ATS_US Then
                totCell.Appearance.BackColor = Color.LightCoral
            Else
                totCell.Appearance.ResetBackColor()
            End If
        End If



    End Sub
End Class