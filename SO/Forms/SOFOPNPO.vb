Imports Infragistics.Win.UltraWinGrid

Public Class SOFOPNPO
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Dim S As New Text.StringBuilder With {.Length = 0}

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SOFOPNPO" Then
            InquiryMode = True
        End If

        'Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, -12)
        'Set_cmbYP("RYP1", ASCMAIN1.CYP, -24, 0, 0)
        'Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        Check_Form_Options()
        Dim SQLB As New System.Text.StringBuilder

        With dst
            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("POTORDR2.PO_ORDER_NO,")
            S.AppendLine("POTORDR2.PO_ORDER_LNO,")
            S.AppendLine("'Not On Vessel' PO_SHIP_VESSEL,")
            S.AppendLine("'OPENPO' PO_SHIPMENT_NO,")
            S.AppendLine("0 PO_SHIPMENT_LNO,")
            S.AppendLine("POTORDR1.PO_DATE_ORDERED,")
            S.AppendLine("POTORDR1.PO_SPEC_ORDR_NO,")
            S.AppendLine("POTORDR2.PO_DATE_ETA,")
            S.AppendLine("POTORDR2.STYLE_CODE,")
            S.AppendLine("POTORDR2.COLOR_CODE,")
            S.AppendLine("ICTSTYL1.STYLE_DESC,")
            S.AppendLine("DECODE(NVL(ICTSTYL1.CUST_CODE,'STOCK'),'','STOCK',NVL(ICTSTYL1.CUST_CODE,'STOCK')) AS CUST_CODE,")
            S.AppendLine("NVL(POTORDR2.PO_QTY_ORD,0) AS PO_QTY_ORD,")
            S.AppendLine("NVL(POTORDR2.PO_QTY_SHP,0) AS PO_QTY_SHP,")
            S.AppendLine("NVL(POTORDR2.PO_QTY_REC,0) AS PO_QTY_REC,")
            S.AppendLine("NVL(POTORDR2.PO_QTY_OPN,0) AS PO_QTY_OPN,")
            S.AppendLine("0 SHIP_QTY,")
            S.AppendLine("0 SHIP_OPN,")
            S.AppendLine("0 SHIP_REC")
            S.AppendLine("FROM POTORDR2,ICTSTYL1,POTORDR1")
            S.AppendLine("WHERE ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE")
            S.AppendLine("AND POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO")
            S.AppendLine("AND POTORDR1.PO_STATUS = 'O'")
            S.AppendLine("AND POTORDR2.PO_STATUS = 'O'")
            S.AppendLine("AND NVL(POTORDR2.PO_QTY_OPN,0) <> 0")
            S.AppendLine(" UNION")
            S.AppendLine("SELECT")
            S.AppendLine("POTORDR2.PO_ORDER_NO,")
            S.AppendLine("POTORDR2.PO_ORDER_LNO,")
            S.AppendLine("POTSHIP1.PO_SHIP_VESSEL,")
            S.AppendLine("POTSHIP3.PO_SHIPMENT_NO,")
            S.AppendLine("POTSHIP3.PO_SHIPMENT_LNO,")
            S.AppendLine("POTORDR1.PO_DATE_ORDERED,")
            S.AppendLine("POTORDR1.PO_SPEC_ORDR_NO,")
            S.AppendLine("POTORDR2.PO_DATE_ETA,")
            S.AppendLine("POTORDR2.STYLE_CODE,")
            S.AppendLine("POTORDR2.COLOR_CODE,")
            S.AppendLine("ICTSTYL1.STYLE_DESC,")
            S.AppendLine("DECODE(NVL(ICTSTYL1.CUST_CODE,'STOCK'),'','STOCK',NVL(ICTSTYL1.CUST_CODE,'STOCK')) AS CUST_CODE,")
            S.AppendLine("0 PO_QTY_ORD,")
            S.AppendLine("0 PO_QTY_SHP,")
            S.AppendLine("0 PO_QTY_REC,")
            S.AppendLine("0 PO_QTY_OPN,")
            S.AppendLine("NVL(POTSHIP3.PO_QTY_SHP,0) AS SHIP_QTY,")
            S.AppendLine("DECODE (POTSHIP2.PO_SHIP_STATUS,'O',NVL(POTSHIP3.PO_QTY_SHP,0),0) SHIP_OPN,")
            S.AppendLine("NVL(POTSHIP3.PO_QTY_REC,0) SHIP_REC")
            S.AppendLine("FROM POTORDR2,ICTSTYL1,POTORDR1,POTSHIP2,POTSHIP3,POTSHIP1")
            S.AppendLine("WHERE  ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE")
            S.AppendLine("AND POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO")
            S.AppendLine("AND POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO")
            S.AppendLine("AND POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO")
            S.AppendLine("AND POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO")
            S.AppendLine("AND POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO")
            S.AppendLine("AND POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO")
            S.AppendLine("AND POTSHIP2.PO_SHIP_STATUS = 'O'")
            S.AppendLine("AND NVL(POTSHIP3.PO_QTY_REC,0) = 0")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "SOTOPNPO", "**", 0, False)
        End With

        grdSOTOPNPO.DataSource = dst.Tables("SOTOPNPO")

        Sort_grdColumns(grdSOTOPNPO, "PO_DATE_ETA, STYLE_CODE, COLOR_CODE", False)

        Create_Summary(grdSOTOPNPO, "PO_QTY_ORD",,, "#,###,###,##0")
        Create_Summary(grdSOTOPNPO, "PO_QTY_SHP",,, "#,###,###,##0")
        Create_Summary(grdSOTOPNPO, "PO_QTY_REC",,, "#,###,###,##0")
        Create_Summary(grdSOTOPNPO, "PO_QTY_OPN",,, "#,###,###,##0")
        Create_Summary(grdSOTOPNPO, "SHIP_QTY",,, "#,###,###,##0")
        Create_Summary(grdSOTOPNPO, "SHIP_OPN",,, "#,###,###,##0")
        Create_Summary(grdSOTOPNPO, "SHIP_REC",,, "#,###,###,##0")

        TABLE_NAME = "SOTOPNPO"

        EntryMode = "E"
        'Call Load_Record()

        With grdSOTOPNPO.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN", "SHIP_QTY", "SHIP_OPN", "SHIP_REC"}
                .Columns(COLUMN_NAME).Format = "#,###,##0"
            Next
            For Each COLUMN_NAME As String In New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN"}
                With .Columns(COLUMN_NAME)
                    .Header.Appearance.BackColor2 = Drawing.Color.Gold
                End With
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_QTY", "SHIP_OPN", "SHIP_REC"}
                With .Columns(COLUMN_NAME)
                    .Header.Appearance.BackColor2 = Drawing.Color.Blue
                End With
            Next
            For Each COLUMN_NAME As String In New String() {"PO_DATE_ORDERED", "PO_DATE_ETA"}
                .Columns(COLUMN_NAME).Format = "MM/dd/yyyy"
            Next
        End With

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
            Case "Update"
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Done"
                Call Mode_Settings(False)
                'Me.Close()
            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        TabControl1.Visible = Not tf

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            '.Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
            '.Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
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
            {"SOTOPNPO"}
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
        Fill_Records("SOTOPNPO")
        dst.EnforceConstraints = True

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record()
        'BeginTrans()
        'INIT_LAST("PMTVIST1", True, "", True)
        'Update_Record_TDA("PMTVIST1")
        'CommitTrans("Update Complete")
    End Sub

    Sub Setup_Summary()
        grdSOTOPNPO.Update()
        grdSOTOPNPO.Refresh()
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
        Call Load_Popup_Menu(grdSOTOPNPO, "SS", "Show Filter", "Show GroupBox")
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

#End Region

#Region "Custom Methods"

#End Region

#Region "Form Controls"
    Private Sub grdSOTOPNPO_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTOPNPO.DoubleClickRow
        'If e.Row.IsDataRow Then
        'Absx1.txtFor("ORDR_GROUP_NO").Text = ORDR_GROUP_NO
        'Click_Command("View")
        'End If
    End Sub
#End Region

End Class