Imports Infragistics.Win.UltraWinGrid

Public Class SOFSBYST
    Private sql As New Text.StringBuilder With {.Length = 0}
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SOFSHPWI" Then
            InquiryMode = True
        End If

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, (Val(ASCMAIN1.CYP.Substring(4)) + 11) * -1)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -24, 0, (Val(ASCMAIN1.CYP.Substring(4))) * -1)
        'Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        Check_Form_Options()
        Dim SQLB As New System.Text.StringBuilder

        With dst
            SQLB.Length = 0
            SQLB.AppendLine("SELECT")
            SQLB.AppendLine("DECODE(INV.WHSE_CODE,'CANADA','CANADA','USA') SHIP_FROM,")
            SQLB.AppendLine("INV.SALES_DIVISION_CODE,")
            SQLB.AppendLine("DECODE(SOTORDR5.CUST_STATE,'NY','NY',DECODE(SOTORDR5.CUST_STATE,'NJ','NJ','OTHER')) STATE,")
            SQLB.AppendLine("DECODE(SOTORDR5.CUST_CITY, 'NEW YORK', 'NEW YORK','OTHER') CITY,")
            SQLB.AppendLine("COUNT(DISTINCT SOTORDR5.ORDR_NO) ORDR_CNT,")
            SQLB.AppendLine("SUM(INV.INV_SALES_CALC) INV")
            SQLB.AppendLine("FROM SOTORDR1, SOTORDR5,")
            SQLB.AppendLine(" (")
            SQLB.AppendLine("  SELECT")
            SQLB.AppendLine("  S1.ORDR_NO,")
            SQLB.AppendLine("  S1.WHSE_CODE,")
            SQLB.AppendLine("  I1.SALES_DIVISION_CODE,")
            SQLB.AppendLine("  SUM(NVL(S2.ORDR_UNIT_PRICE,0) * NVL(S2.ORDR_QTY_SHIP,0)) INV_SALES_CALC")
            SQLB.AppendLine("  FROM SOTINVH1 S1, SOTINVH2 S2, ICTSTYL1 I1")
            SQLB.AppendLine("  WHERE S1.INV_TYPE = S2.INV_TYPE")
            SQLB.AppendLine("  AND S1.INV_NO = S2.INV_NO")
            SQLB.AppendLine("  AND S2.STYLE_CODE = I1.STYLE_CODE")
            SQLB.AppendLine("  AND S1.INV_TYPE = 'I'")
            SQLB.AppendLine("  AND S1.ORDR_YYYYPP_UPDATED >= :PARM1")
            SQLB.AppendLine("  AND S1.ORDR_YYYYPP_UPDATED <= :PARM2")
            SQLB.AppendLine("  AND NVL(INV_NO_REV,'NULL') = 'NULL'")
            SQLB.AppendLine("  AND NVL(INV_NO_REV_BY,'NULL') = 'NULL'")
            SQLB.AppendLine("  GROUP BY")
            SQLB.AppendLine("  S1.ORDR_NO,")
            SQLB.AppendLine("  S1.WHSE_CODE,")
            SQLB.AppendLine("  I1.SALES_DIVISION_CODE")
            SQLB.AppendLine(") INV")
            SQLB.AppendLine("WHERE SOTORDR5.ORDR_NO = INV.ORDR_NO")
            SQLB.AppendLine("AND SOTORDR1.ORDR_NO = INV.ORDR_NO")
            SQLB.AppendLine("AND SOTORDR5.CUST_ADDR_TYPE= 'ST'")
            SQLB.AppendLine("GROUP BY")
            SQLB.AppendLine("DECODE(INV.WHSE_CODE,'CANADA','CANADA','USA'),")
            SQLB.AppendLine("INV.SALES_DIVISION_CODE,")
            SQLB.AppendLine("DECODE(CUST_STATE,'NY','NY',DECODE(CUST_STATE,'NJ','NJ','OTHER')),")
            SQLB.AppendLine("DECODE(CUST_CITY, 'NEW YORK', 'NEW YORK','OTHER')")
            SQLB.AppendLine("ORDER BY")
            SQLB.AppendLine("DECODE(INV.WHSE_CODE,'CANADA','CANADA','USA'),")
            SQLB.AppendLine("INV.SALES_DIVISION_CODE,")
            SQLB.AppendLine("DECODE(CUST_STATE,'NY','NY',DECODE(CUST_STATE,'NJ','NJ','OTHER')),")
            SQLB.AppendLine("DECODE(CUST_CITY, 'NEW YORK', 'NEW YORK','OTHER')")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTSBYST", "**", 0, False, "VV")
        End With

        grdSOFSBYST.DataSource = dst.Tables("SOTSBYST")

        Sort_grdColumns(grdSOFSBYST, "SHIP_FROM, SALES_DIVISION_CODE, STATE, CITY", False)

        'grdGROUPS.DataSource = dst.Tables("SOTGROUP")

        TABLE_NAME = "SOTSBYST"

        EntryMode = "E"
        'Call Load_Record()

        Create_Summary(grdSOFSBYST, "ORDR_CNT")
        Create_Summary(grdSOFSBYST, "INV")

        With grdSOFSBYST.DisplayLayout.Bands(0)
            .Columns("ORDR_CNT").Format = "###,##0"
            .Columns("INV").Format = "###,###,###,##0.00"
        End With

        'ASCMAIN1.Add_Value_List(grdSOFSHPWA, "ORDR_STATUS", , New String() {":", "C:Cancelled", "D:Deleted", "F:Final", "O:Open", "P:In Pick"})

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
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        'TabControl1.Visible = Not tf

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Done").Settings.Enabled = DefaultableBoolean.True
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        For Each TABLE_NAME As String In New String() _
            {"SOTSBYST"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        dst.EnforceConstraints = True
        'Setup_Summary()
    End Sub

    Sub Load_Record()

        'tab.Visible = ScreenMode

        'Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        'dst.Tables("SOTROYLI").Rows.Clear()

        'dst.EnforceConstraints = False

        'Fill_Records("SOTSBYST", ORDR_GROUP_NO)

        'dst.EnforceConstraints = True

        'Save_Header_Fields(UltraGroupBox1)

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
        Call Load_Popup_Menu(grdSOFSBYST, "SS", "Show Filter", "Show GroupBox")
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

#Region "Form Controls"

    Private Sub btnFETCH_Click(sender As Object, e As EventArgs) Handles btnFETCH.Click
        Me.Cursor = Cursors.WaitCursor

        Dim RYPLEGEND0 As String = Absx1.cmbFor("RYP0", True).Value
        Dim RYP0 As String = Mid(RYPLEGEND0, 1, 4) & Mid(RYPLEGEND0, 6, 2)
        Dim RYPLEGEND1 As String = Absx1.cmbFor("RYP1", True).Value
        Dim RYP1 As String = Mid(RYPLEGEND1, 1, 4) & Mid(RYPLEGEND1, 6, 2)

        Fill_Records("SOTSBYST", New String() {RYP0, RYP1})

        Me.Cursor = Cursors.Default
    End Sub

#End Region

#Region "Custom Methods"

#End Region
End Class