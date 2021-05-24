Imports Infragistics.Win.UltraWinGrid

Public Class WHFSCSQ1
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'If MENU_ITEM_OBJECT = "SOFOPNPO" Then
        '    InquiryMode = True
        'End If

        ''Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, -12)
        ''Set_cmbYP("RYP1", ASCMAIN1.CYP, -24, 0, 0)
        ''Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        'Check_Form_Options()
        'Dim SQLB As New System.Text.StringBuilder

        'With dst
        '    S.Length = 0
        '    S.AppendLine("SELECT")
        '    S.AppendLine("POTORDR2.PO_ORDER_NO,")
        '    S.AppendLine("POTORDR2.PO_ORDER_LNO,")
        '    S.AppendLine("'Not On Vessel' PO_SHIP_VESSEL,")
        '    S.AppendLine("'OPENPO' PO_SHIPMENT_NO,")
        '    S.AppendLine("0 PO_SHIPMENT_LNO,")
        '    S.AppendLine("POTORDR1.PO_DATE_ORDERED,")
        '    S.AppendLine("POTORDR1.PO_SPEC_ORDR_NO,")
        '    S.AppendLine("POTORDR2.PO_DATE_ETA,")
        '    S.AppendLine("POTORDR2.STYLE_CODE,")
        '    S.AppendLine("POTORDR2.COLOR_CODE,")
        '    S.AppendLine("ICTSTYL1.STYLE_DESC,")
        '    S.AppendLine("DECODE(NVL(ICTSTYL1.CUST_CODE,'STOCK'),'','STOCK',NVL(ICTSTYL1.CUST_CODE,'STOCK')) AS CUST_CODE,")
        '    S.AppendLine("NVL(POTORDR2.PO_QTY_ORD,0) AS PO_QTY_ORD,")
        '    S.AppendLine("NVL(POTORDR2.PO_QTY_SHP,0) AS PO_QTY_SHP,")
        '    S.AppendLine("NVL(POTORDR2.PO_QTY_REC,0) AS PO_QTY_REC,")
        '    S.AppendLine("NVL(POTORDR2.PO_QTY_OPN,0) AS PO_QTY_OPN,")
        '    S.AppendLine("0 SHIP_QTY,")
        '    S.AppendLine("0 SHIP_OPN,")
        '    S.AppendLine("0 SHIP_REC")
        '    S.AppendLine("FROM POTORDR2,ICTSTYL1,POTORDR1")
        '    S.AppendLine("WHERE ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE")
        '    S.AppendLine("AND POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO")
        '    S.AppendLine("AND POTORDR1.PO_STATUS = 'O'")
        '    S.AppendLine("AND POTORDR2.PO_STATUS = 'O'")
        '    S.AppendLine("AND NVL(POTORDR2.PO_QTY_OPN,0) <> 0")
        '    S.AppendLine(" UNION")
        '    S.AppendLine("SELECT")
        '    S.AppendLine("POTORDR2.PO_ORDER_NO,")
        '    S.AppendLine("POTORDR2.PO_ORDER_LNO,")
        '    S.AppendLine("POTSHIP1.PO_SHIP_VESSEL,")
        '    S.AppendLine("POTSHIP3.PO_SHIPMENT_NO,")
        '    S.AppendLine("POTSHIP3.PO_SHIPMENT_LNO,")
        '    S.AppendLine("POTORDR1.PO_DATE_ORDERED,")
        '    S.AppendLine("POTORDR1.PO_SPEC_ORDR_NO,")
        '    S.AppendLine("POTORDR2.PO_DATE_ETA,")
        '    S.AppendLine("POTORDR2.STYLE_CODE,")
        '    S.AppendLine("POTORDR2.COLOR_CODE,")
        '    S.AppendLine("ICTSTYL1.STYLE_DESC,")
        '    S.AppendLine("DECODE(NVL(ICTSTYL1.CUST_CODE,'STOCK'),'','STOCK',NVL(ICTSTYL1.CUST_CODE,'STOCK')) AS CUST_CODE,")
        '    S.AppendLine("0 PO_QTY_ORD,")
        '    S.AppendLine("0 PO_QTY_SHP,")
        '    S.AppendLine("0 PO_QTY_REC,")
        '    S.AppendLine("0 PO_QTY_OPN,")
        '    S.AppendLine("NVL(POTSHIP3.PO_QTY_SHP,0) AS SHIP_QTY,")
        '    S.AppendLine("DECODE (POTSHIP2.PO_SHIP_STATUS,'O',NVL(POTSHIP3.PO_QTY_SHP,0),0) SHIP_OPN,")
        '    S.AppendLine("NVL(POTSHIP3.PO_QTY_REC,0) SHIP_REC")
        '    S.AppendLine("FROM POTORDR2,ICTSTYL1,POTORDR1,POTSHIP2,POTSHIP3,POTSHIP1")
        '    S.AppendLine("WHERE  ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE")
        '    S.AppendLine("AND POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO")
        '    S.AppendLine("AND POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO")
        '    S.AppendLine("AND POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO")
        '    S.AppendLine("AND POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO")
        '    S.AppendLine("AND POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO")
        '    S.AppendLine("AND POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO")
        '    S.AppendLine("AND POTSHIP2.PO_SHIP_STATUS = 'O'")
        '    S.AppendLine("AND NVL(POTSHIP3.PO_QTY_REC,0) = 0")
        '    ASCMAIN1.sql = S.ToString
        '    Create_TDA(.Tables.Add, "SOTOPNPO", "**", 0, False)
        'End With

        'grdSOTOPNPO.DataSource = dst.Tables("SOTOPNPO")

        'Sort_grdColumns(grdSOTOPNPO, "PO_DATE_ETA, STYLE_CODE, COLOR_CODE", False)

        'Create_Summary(grdSOTOPNPO, "PO_QTY_ORD",,, "#,###,###,##0")
        'Create_Summary(grdSOTOPNPO, "PO_QTY_SHP",,, "#,###,###,##0")
        'Create_Summary(grdSOTOPNPO, "PO_QTY_REC",,, "#,###,###,##0")
        'Create_Summary(grdSOTOPNPO, "PO_QTY_OPN",,, "#,###,###,##0")
        'Create_Summary(grdSOTOPNPO, "SHIP_QTY",,, "#,###,###,##0")
        'Create_Summary(grdSOTOPNPO, "SHIP_OPN",,, "#,###,###,##0")
        'Create_Summary(grdSOTOPNPO, "SHIP_REC",,, "#,###,###,##0")

        'TABLE_NAME = "SOTOPNPO"

        'EntryMode = "E"
        ''Call Load_Record()

        'With grdSOTOPNPO.DisplayLayout.Bands(0)
        '    For Each COLUMN_NAME As String In New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN", "SHIP_QTY", "SHIP_OPN", "SHIP_REC"}
        '        .Columns(COLUMN_NAME).Format = "#,###,##0"
        '    Next
        '    For Each COLUMN_NAME As String In New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN"}
        '        With .Columns(COLUMN_NAME)
        '            .Header.Appearance.BackColor2 = Drawing.Color.Gold
        '        End With
        '    Next
        '    For Each COLUMN_NAME As String In New String() {"SHIP_QTY", "SHIP_OPN", "SHIP_REC"}
        '        With .Columns(COLUMN_NAME)
        '            .Header.Appearance.BackColor2 = Drawing.Color.Blue
        '        End With
        '    Next
        '    For Each COLUMN_NAME As String In New String() {"PO_DATE_ORDERED", "PO_DATE_ETA"}
        '        .Columns(COLUMN_NAME).Format = "MM/dd/yyyy"
        '    Next
        'End With


        With dst
            ASCMAIN1.sql = "Select * from WHTSCSEQ"
            Create_TDA(.Tables.Add, "WHTSCSEQ", ASCMAIN1.sql, 0, True, 3)

            ASCMAIN1.sql = "Select WHTLOCM1.LOCATION_CODE, WHTLOCM1.LOCATION_ROUTE_SEQ STYLE_SEQ,ICVLUPC1.UPC_CODE" & vbCrLf _
                        & " ,ICVLUPC1.STYLE_CODE,ICVLUPC1.COLOR_CODE,ICVLUPC1.COLOR_CODE_UPC" & vbCrLf _
                        & " ,ICVLUPC1.SIZE_CODE, ICTSTYC1.STYLE_COLOR_DESC,  WHTP2LM1.CUST_CODE" & vbCrLf _
                        & " From WHTSCSEQ, WHTLOCM1, ICTSTYL1, ICVLUPC1, ICTSTYC1, WHTP2LM1" & vbCrLf _
                        & " Where ICTSTYL1.STYLE_CODE = WHTSCSEQ.STYLE_CODE And" & vbCrLf _
                        & " ICTSTYC1.STYLE_CODE = WHTSCSEQ.STYLE_CODE And" & vbCrLf _
                        & " ICTSTYC1.COLOR_CODE = WHTSCSEQ.COLOR_CODE And" & vbCrLf _
                        & " WHTLOCM1.LOCATION_ROUTE_SEQ = WHTSCSEQ.STYLE_SEQ And" & vbCrLf _
                        & " WHTP2LM1.CUST_CODE = WHTSCSEQ.CUST_CODE And" & vbCrLf _
                        & " WHTP2LM1.P2L_LINE_ID = SUBSTR(WHTLOCM1.LOCATION_CODE, 1, 2) And " & vbCrLf _
                        & " ICVLUPC1.STYLE_CODE = WHTSCSEQ.STYLE_CODE And" & vbCrLf _
                        & " ICVLUPC1.COLOR_CODE = WHTSCSEQ.COLOR_CODE" & vbCrLf
            Create_TDA(.Tables.Add, "WHTSCLAB", ASCMAIN1.sql, 0, False, "V", 2)

        End With

        '   grdWHTSCSEQ.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdWHTSCSEQ.DataSource = dst.Tables("WHTSCSEQ")
        grdWHTSCLAB.DataSource = dst.Tables("WHTSCLAB")



        'Bind_Controls(grpHeader, "SOTSHPWH")

        'ASCMAIN1.Add_Value_List(grdSOTOPNPO, "ORDR_STATUS", , New String() {":", "C:Cancelled", "D:Deleted"})

        Call Mode_Settings(True)

        'SplitContainer2.SplitterDistance = 120

    End Sub

    Sub Check_Inquiry_Mode()
        If InquiryMode Then
        Else
        End If
    End Sub

    Sub Check_Form_Options()
        'With UltraExplorerBar1.Groups("Screen Control")
        '    .Items("New").Visible = (Me.Name = "PMFVIST1")
        'End With
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Load from XLS"


            Case "Cancel"
                Mode_Settings(False)

            Case "Print Labels"
                PRINT_LABELS()
                Mode_Settings(False)



        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "ReLoad"
                Call Mode_Settings(False)
                Clear_Record()
                Load_Record()

                'Me.Close()
            Case "Cancel"
                Call Mode_Settings(False)
                '  UNLOAD Me
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        TabControl1.Visible = Not tf

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("ReLoad").Settings.Enabled = not_iScreenMode
            '     .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Cancel").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Print Labels").Settings.Enabled = not_iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If Not ScreenMode Then
        Else
            Clear_Record()
            Load_Record()
        End If

    End Sub

    Sub Clear_Record()

        'Absx1.txtFor("ORDR_GROUP_NO").Text = ""
        'Absx1.txtFor("ORDR_CUST_PO").Text = ""

        dst.EnforceConstraints = False
        For Each TABLE_NAME As String In New String() _
            {"WHTSCSEQ", "WHTSCLAB"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        'Fill_Records("PMTVIST1")

        dst.EnforceConstraints = True
        'Setup_Summary()
    End Sub

    Sub Load_Record()

        'tab.Visible = ScreenMode

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        'dst.Tables("SOTROYLI").Rows.Clear()

        dst.EnforceConstraints = False
        Fill_Records("WHTSCSEQ")

        ASCMAIN1.sql = "Select WHTLOCM1.LOCATION_CODE, WHTLOCM1.LOCATION_ROUTE_SEQ STYLE_SEQ" & vbCrLf _
                    & ", WHTSCSEQ.STYLE_CODE,WHTSCSEQ.COLOR_CODE,WHTP2LM1.CUST_CODE" & vbCrLf _
                    & " From WHTSCSEQ, WHTP2LM1, WHTLOCM1" & vbCrLf _
                    & " Where WHTP2LM1.P2L_LINE_ID = SUBSTR(WHTLOCM1.LOCATION_CODE, 1, 2)" & vbCrLf _
                    & " And WHTP2LM1.WHSE_CODE = WHTLOCM1.WHSE_CODE" & vbCrLf _
                    & " And WHTSCSEQ.STYLE_SEQ(+) = WHTLOCM1.LOCATION_ROUTE_SEQ" & vbCrLf _
                    & " And WHTSCSEQ.CUST_CODE = WHTP2LM1.CUST_CODE"
        Fill_Records("WHTSCLAB",, True, ASCMAIN1.sql)
        Dim rowWHTSCLAB As DataRow

        ASCMAIN1.sql = "  Select WHTSCSEQ.STYLE_CODE,WHTSCSEQ.COLOR_CODE,ICVLUPC1.UPC_CODE,ICVLUPC1.COLOR_CODE_UPC," & vbCrLf _
                        & " ICVLUPC1.SIZE_CODE, ICTSTYC1.STYLE_COLOR_DESC, whtscseq.cust_code, whtscseq.style_seq" & vbCrLf _
                        & " From WHTSCSEQ, ICTSTYL1, ICVLUPC1, ICTSTYC1, (Select distinct CUST_CODE from WHTP2LM1) WHTP2LM1" & vbCrLf _
                        & " Where ICTSTYL1.STYLE_CODE = WHTSCSEQ.STYLE_CODE And" & vbCrLf _
                        & " ICTSTYC1.STYLE_CODE = WHTSCSEQ.STYLE_CODE And" & vbCrLf _
                        & " ICTSTYC1.COLOR_CODE = WHTSCSEQ.COLOR_CODE And" & vbCrLf _
                        & " ICVLUPC1.STYLE_CODE = WHTSCSEQ.STYLE_CODE And" & vbCrLf _
                        & " ICVLUPC1.COLOR_CODE = WHTSCSEQ.COLOR_CODE" & vbCrLf _
                        & " And whtscseq.cust_code = whtp2lm1.cust_code"

        For Each ROW As DataRow In ASCDATA1.GetDataTable().Select("")
            Dim STYLE_CODE As String = ROW.Item("STYLE_CODE")
            Dim COLOR_CODE As String = ROW.Item("COLOR_CODE")
            For Each rowWHTSCLAB In dst.Tables("WHTSCLAB").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
                rowWHTSCLAB.Item("UPC_CODE") = ROW.Item("UPC_CODE")
                rowWHTSCLAB.Item("COLOR_CODE_UPC") = ROW.Item("COLOR_CODE_UPC")
                rowWHTSCLAB.Item("SIZE_CODE") = ROW.Item("SIZE_CODE")
                rowWHTSCLAB.Item("STYLE_COLOR_DESC") = ROW.Item("STYLE_COLOR_DESC")
            Next
        Next
        dst.Tables("WHTSCLAB").AcceptChanges()


        dst.EnforceConstraints = True

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record()
        Try
            BeginTrans()

            ASCMAIN1.sql = "Delete from WHTSCSEQ"
            ASCDATA1.ExecuteSQL()

            Update_Record_TDA("WHTSCSEQ", "1=1")

            CommitTrans("Update Complete")

        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

    Sub Setup_Summary()
        grdWHTSCSEQ.Update()
        grdWHTSCSEQ.Refresh()
        Me.Cursor = Cursors.Default
    End Sub

    Sub Print_Report()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'CR_params.Add("GROUP_BY", optGroupBy.Value)
        'Generate_Report("PMRVIST1", "Open Site Visit Report")
        'Print_Report_End()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        'Select Case COLUMN_NAME
        '    Case "JOB_NO"
        '        sql_where = "JOB_STATUS = 'O' and SITE_VISITS > 0"
        'End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdWHTSCSEQ, "SS", "Show Filter", "Show GroupBox")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool
        'Dim tlb_btn As UltraWinToolbars.ButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case ""

                    If grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.True Then
                        e.Cancel = True
                    End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub

#End Region

#Region "ABSColumn Controls"
    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

    End Sub

    Private Sub UltraGroupBox1_Click(sender As Object, e As EventArgs) Handles UltraGroupBox1.Click

    End Sub

    Private Sub grdWHTSCSEQ_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdWHTSCSEQ.InitializeLayout

    End Sub


#End Region

#Region "Custom Methods"

#End Region

#Region "Form Controls"

#End Region

    Sub PRINT_LABELS()

        Fill_Records("WHTSCSEQ")

        Dim sqlw As String = ""

        If chkF1.Checked = False Then
            sqlw = sqlw & " or substring(LOCATION_CODE, 1, 2) = 'F1'"
        End If
        If chkF2.Checked = False Then
            sqlw = sqlw & " or substring(LOCATION_CODE,1,2) = 'F2'"
        End If
        If chkF3.Checked = False Then
            sqlw = sqlw & " or substring(LOCATION_CODE,1,2) = 'F3'"
        End If
        If chkF4.Checked = False Then
            sqlw = sqlw & " or substring(LOCATION_CODE,1,2) = 'F4'"
        End If

        If sqlw <> "" Then
            sqlw = Mid(sqlw, 5)
            ASCDATA1.DeleteRows("WHTSCLAB", sqlw)
        End If

        dst.Tables("WHTSCLAB").AcceptChanges()



        '  Synch_TABLE_NAME("WHTSCLAB")
        Print_Report_Begin()
        Dim RPT As String = "WHRSCSEQ"
        Dim FILTER As String = ""

        ''If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        ''    If grdWHTSCLAB.Selected.Rows.Count <> 0 Then
        ''        For Each grow As UltraWinGrid.UltraGridRow In grdWHTSCLAB.Selected.Rows
        ''            '     Dim PO_SHIPMENT_LNO As Integer = Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO"))
        ''            '   FILTER = " or {POTSHIP2.PO_SHIPMENT_LNO} = " & CStr(PO_SHIPMENT_LNO)
        ''        Next
        ''        FILTER = Mid(FILTER, 5)
        ''    End If
        ''End If

        Dim RPT_TITLE As String = ""
        Dim SUBT As String = ""
        Generate_Report(RPT)

        Print_Report_End()

    End Sub


End Class