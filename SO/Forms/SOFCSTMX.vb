Imports System.Text

Public Class SOFCSTMX
    Dim sqlSOFCSTMX As String = ""
    Dim SREP_CODE As String = ""
    Dim Remote As New REMOTE(Me)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim BaseYear As Int64 = Now().Year + 1
        Dim BaseMonth As Int64 = Now().Month - 3
        If BaseMonth < 1 Then
            BaseMonth = 4 + BaseMonth
            BaseYear = BaseYear - 1
        End If

        Dim CYP_CALC_BEG As String = BaseYear.ToString & BaseMonth.ToString
        Dim CYP_CALC_END As String = BaseYear.ToString & BaseMonth.ToString

        Set_cmbYP("RYP0", CYP_CALC_BEG, -36, 0, -11)
        Set_cmbYP("RYP1", CYP_CALC_BEG, -36, 0, 0)

        SREP_CODE = Remote.SREP_CODE
        dt1.DateTime = DateSerial(Now.Year, 1, 1)
        dt2.DateTime = DateSerial(Now.Year, 12, 31)

        If MENU_ITEM_OBJECT = "SOFCSTMI" Then
            InquiryMode = True
        End If

        Check_Form_Options()

        With dst
            Dim SQLB As New System.Text.StringBuilder
            SQLB.Length = 0
            SQLB.AppendLine("SELECT")
            SQLB.AppendLine("C1.CUST_CODE,")
            SQLB.AppendLine("C1.CUST_NAME,")
            SQLB.AppendLine("C1.CUST_ADDR1,")
            SQLB.AppendLine("C1.CUST_ADDR2,")
            SQLB.AppendLine("C1.CUST_ADDR3,")
            SQLB.AppendLine("C1.CUST_CITY,")
            SQLB.AppendLine("C1.CUST_STATE,")
            SQLB.AppendLine("C1.CUST_ZIP_CODE,")
            SQLB.AppendLine("C1.CUST_COUNTRY,")
            SQLB.AppendLine("C1.CUST_CONTACT,")
            SQLB.AppendLine("C1.CUST_PHONE,")
            SQLB.AppendLine("C1.CUST_EMAIL,")
            SQLB.AppendLine("C1.SREP_CODE,")
            SQLB.AppendLine("C1.INIT_DATE,")
            SQLB.AppendLine("NVL(C1.CUST_SALES_HOLD,'0') CUST_SALES_HOLD,")
            SQLB.AppendLine("NVL(C1.CUST_CREDIT_HOLD,'0') CUST_CREDIT_HOLD")
            SQLB.AppendLine("FROM ARTCUST1 C1")
            SQLB.AppendLine("WHERE C1.CUST_STATUS = 'A'")
            If ASCMAIN1.DBS_COMPANY <> "RGI" Then
                MAKE_SR_FILTER("C1", SQLB)
            End If
            ASCMAIN1.sql = SQLB.ToString
            sqlSOFCSTMX = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTCSTMX", "**", 0, False)
            With .Tables("SOTCSTMX").Columns
                .Add("BUYER1", GetType(String))
                .Add("BUYER2", GetType(String))
                .Add("BUYER3", GetType(String))
                .Add("BUYER_EMAIL1", GetType(String))
                .Add("BUYER_EMAIL2", GetType(String))
                .Add("BUYER_EMAIL3", GetType(String))
                .Add("ORDRED_TY", GetType(Double))
                .Add("SHIPPED_TY", GetType(Double))
                .Add("CANCELLED_TY", GetType(Double))
                .Add("ORDRED_LY", GetType(Double))
                .Add("SHIPPED_LY", GetType(Double))
                .Add("CANCELLED_LY", GetType(Double))
                .Add("VARIANCE_TY", GetType(System.Decimal), "ISNULL(SHIPPED_TY,0) - ISNULL(ORDRED_TY,0)")
                .Add("VARIANCE_LY", GetType(System.Decimal), "ISNULL(SHIPPED_LY,0) - ISNULL(ORDRED_LY,0)")
                .Add("BOOKED_VAR", GetType(System.Decimal), "ISNULL(ORDRED_TY,0) - ISNULL(ORDRED_LY,0)")
            End With

            ASCMAIN1.sql = MAKE_SQL("SOTCSTMD", True)
            Create_TDA(.Tables.Add, "SOTCSTMD", "**", 0, False)

            ASCMAIN1.sql = MAKE_SQL("SOTCUSTX", True)
            Create_TDA(.Tables.Add, "SOTCUSTX", "**", 0, False)

            Create_TDA(.Tables.Add, "SOTORDR2", "*", 1)
            With .Tables("SOTORDR2").Columns
                .Add("ORDR_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("CASE_CUBE", GetType(System.Double))
                .Add("TCUFT", GetType(System.Double))
                .Add("ORDR_AMT_OPEN", GetType(System.Decimal), "ISNULL(ORDR_QTY_OPEN,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_ALLO", GetType(System.Decimal), "ISNULL(ORDR_QTY_ALLO,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_PICK", GetType(System.Decimal), "ISNULL(ORDR_QTY_PICK,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_SHIP", GetType(System.Decimal), "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_CANC", GetType(System.Decimal), "ISNULL(ORDR_QTY_CANC,0) * ISNULL(ORDR_UNIT_PRICE,0)")
            End With


            SQLB.Length = 0
            SQLB.AppendLine("SELECT *")
            SQLB.AppendLine("FROM ARTCUSTD")
            SQLB.AppendLine("WHERE CONTACT_TYPE = 'B'")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "ARTCUSTD", "*", 0, False)
            Fill_Records("ARTCUSTD")
        End With

        grdSOFCSTMX.DataSource = dst.Tables("SOTCSTMX")
        grdSOFCSTMD.DataSource = dst.Tables("SOTCSTMD")
        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")

        'ASCMAIN1.Add_Value_List(grdSOFCSTMX, "REPORT_TYPE", , New String() {":", "I:Initial", "A:Amended", "S:Subsequent", "R:Revised"})

        Create_Summary(grdSOFCSTMX, "ORDRED_TY", "Sum")
        Create_Summary(grdSOFCSTMX, "SHIPPED_TY", "Sum")
        Create_Summary(grdSOFCSTMX, "CANCELLED_TY", "Sum")
        Create_Summary(grdSOFCSTMX, "ORDRED_LY", "Sum")
        Create_Summary(grdSOFCSTMX, "SHIPPED_LY", "Sum")
        Create_Summary(grdSOFCSTMX, "CANCELLED_LY", "Sum")
        Create_Summary(grdSOFCSTMX, "CUST_CODE", "Count")

        Create_Summary(grdSOFCSTMD, "ORDRED", "Sum")
        Create_Summary(grdSOFCSTMD, "SHIPPED", "Sum")
        Create_Summary(grdSOFCSTMD, "CANCELLED", "Sum")

        Create_Summary(grdSOTORDR2, "ORDR_LNO", "Count")
        Create_Summary(grdSOTORDR2, New String() {"ORDR_QTY", "ORDR_AMT", "TCUFT", "ORDR_QTY_ALLO", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_AMT_OPEN", "ORDR_AMT_ALLO", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"})

        Sort_grdColumns(grdSOFCSTMD, "ORDR_DATE, ORDR_NO".ToLower(), False)

        Sort_grdColumns(grdSOFCSTMX, "CUST_NAME", False)

        With grdSOFCSTMX.DisplayLayout.Bands(0)
            For Each COL_NAME As String In New String() {"CUST_CODE", "CUST_NAME"}
                .Columns(COL_NAME).Header.Fixed = True
            Next
            .Columns("INIT_DATE").Format = "MM/dd/yy"
        End With

        With grdSOTORDR2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY_ALLO", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_QTY"}
                With .Columns(COLUMN_NAME)
                    .Header.Appearance.BackColor2 = Drawing.Color.LightCyan
                End With
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_AMT_OPEN", "ORDR_AMT_ALLO", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC", "ORDR_AMT"}
                With .Columns(COLUMN_NAME)
                    .Header.Appearance.BackColor2 = Drawing.Color.Gold
                End With
            Next
        End With
        SetShowOrderDetails()

        TABLE_NAME = "SOTCSTMX"

        EntryMode = "E"
        'Call Load_Record()
        Call Mode_Settings(True)
    End Sub

    Private Sub MAKE_SR_FILTER(ByVal PREFIX As String, ByRef SQL As StringBuilder)
        Dim NEXT_LINE As String = ""
        If Not Remote.IsUserSuper Then
            NEXT_LINE = String.Format("AND {0}.SREP_CODE = '{1}'", PREFIX, Remote.SREP_CODE)
        End If
        'Special Code For James and Dimple.
        If Remote.SREP_CODE = "MD" Or Remote.SREP_CODE = "JD" Or Remote.SREP_CODE = "JE" Then
            NEXT_LINE = String.Format("AND {0}.SREP_CODE IN ('MD','JD','JE')", PREFIX)
        End If

        If NEXT_LINE.Length > 0 Then
            SQL.AppendLine(NEXT_LINE)
        End If
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
                Me.Close()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            '.Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
            '.Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
        End With

        UltraExplorerBar1.Groups("Dates").Visible = False
        UltraExplorerBar1.Groups("Periods").Visible = True

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        With grdSOFCSTMX.DisplayLayout.Bands(0)
            For Each thisCOL As String In New String() {"ORDRED_TY", "SHIPPED_TY", "CANCELLED_TY"}
                .Columns.Item(thisCOL).Header.Appearance.BackColor = Drawing.Color.Khaki
            Next
            For Each thisCOL As String In New String() {"ORDRED_LY", "SHIPPED_LY", "CANCELLED_LY"}
                .Columns.Item(thisCOL).Header.Appearance.BackColor = Drawing.Color.Bisque
            Next
        End With


        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        'dst.EnforceConstraints = False
        'dst.Tables("PMTVIST1").Rows.Clear()
        'dst.Tables("PMTVISTH").Rows.Clear()

        'Dim dvw As DataView = DirectCast(grdPMTVIST1.DataSource, DataTable).DefaultView
        'dvw.RowStateFilter = DataViewRowState.CurrentRows

        'Fill_Records("PMTVIST1")
        'Process_SVRs()

        'Sort_grdColumns(grdPMTVIST1, "DATE_VISITED".ToLower)
        'Sort_grdColumns(grdPMTVISTH, "DATE_VISITED".ToLower)
        'dst.EnforceConstraints = True
        'Setup_Summary()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Setup_Summary()

        Setup_SOTCSTMX()

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

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        'Select Case COLUMN_NAME
        '    Case "JOB_NO"
        '        sql_where = "JOB_STATUS = 'O' and SITE_VISITS > 0"
        'End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdSOFCSTMX, "SSB", "Show Filter", "Show GroupBox", "Customer Master File")
        Call Load_Popup_Menu(grdSOFCSTMD, "SSB", "Show Filter", "Show GroupBox", "View Order")
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
            Case "Customer Master File"
                If Not IsNothing(grdSOFCSTMX.ActiveRow) Then
                    Dim CUST_CODE As String = grdSOFCSTMX.ActiveRow.Cells.Item("CUST_CODE").Text
                    If CUST_CODE.Length > 0 Then
                        'Context_Launch("Edit", CUST_CODE, "Customer Master File", "SOTCUST1")
                        Context_Launch("Edit", CUST_CODE, e.Tool.Key, "SOTCUST1")
                    Else
                        Context_Launch("Customer Master File", CUST_CODE, "Customer Master File", "SOTCUST1")
                    End If
                End If
            Case "View Order"
                If Not IsNothing(grdSOFCSTMD.ActiveRow) Then
                    Dim ORDR_GROUP_NO As String = grdSOFCSTMD.ActiveRow.Cells.Item("ORDR_GROUP_NO").Text
                    Dim ORDR_NO As String = grdSOFCSTMD.ActiveRow.Cells.Item("ORDR_NO").Text
                    If ORDR_GROUP_NO.Length > 0 Then
                        If ASCMAIN1.DBS_COMPANY = "RGI" Then
                            Context_Launch("View", Column_Values("ORDR_NO", ORDR_GROUP_NO), "Sales Order Entry", "SOFORDRI")
                        End If
                        If ASCMAIN1.DBS_COMPANY = "RGO" Then
                            Context_Launch("Edit", Column_Values("ORDR_GROUP_NO", ORDR_GROUP_NO), "Sales Order Entry", "SOFORDRO")
                        End If


                        'Context_Launch("Edit", ORDR_GROUP_NO, "Sales Order Entry", "SOFORDRO")
                        'Context_Launch("Edit", ORDR_GROUP_NO, "Sales Order Entry", "SOFORDRO")
                    End If
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Project Center"
                Dim JOB_NO As String = grd.ActiveRow.Cells("JOB_NO").Text
                Context_Launch("Edit", Column_Values("JOB_NO", JOB_NO), e.Tool.Key, "PMFJOBM1")
            Case "Show Report"
                Dim FILENAME As String = "C:\Documents and Settings\wjz\Desktop\randfromdrc\RandInvoices\310 West 52nd Street - 30760.pdf"
                Show_Document(FILENAME)

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        'Select Case Absx1.GetABSColumnName(sender)
        'Case "EMPLOYEE_CODE"
        '    If e.KeyCode = Windows.Forms.Keys.Enter Then
        '        Setup_Summary()
        '    End If
        'End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        'Select Case COLUMN_NAME
        '    Case "EMPLOYEE_CODE"
        '        Setup_Summary()
        'End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        'Select Case Absx1.GetABSColumnName(txtctl)
        '    Case "EMPLOYEE_CODE"
        '        Setup_Summary()
        'End Select
    End Sub

#End Region

    Sub Setup_Summary()
        Dim sqlwhere As String = ""
        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        'grdSOFCSTMX.DataSource = Nothing
        'grdSOFCSTMD.DataSource = Nothing
        'grdSOTORDR2.DataSource = Nothing

        dst.Tables("SOTCSTMX").Rows.Clear()
        dst.Tables("SOTCSTMD").Rows.Clear()
        dst.Tables("SOTCUSTX").Rows.Clear()

        dst.EnforceConstraints = False
        Fill_Records("SOTCSTMX")
        'Dim Dates As String() = CalculateDates()
        Dim sqlSOTCSTMD As String = MAKE_SQL("SOTCSTMD", False)
        Fill_Records("SOTCSTMD",,, sqlSOTCSTMD)
        Dim sqlSOTCUSTX As String = MAKE_SQL("SOTCUSTX", False)
        Fill_Records("SOTCUSTX",,, sqlSOTCUSTX)
        Fill_Extra_Fields()

        'grdSOFCSTMX.DataSource = dst.Tables("SOTCSTMX")
        'grdSOFCSTMD.DataSource = dst.Tables("SOTCSTMD")
        'grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")

        'grdSOFCSTMX.Refresh()
        'grdSOFCSTMD.Refresh()
        'grdSOTORDR2.Refresh()

        filterNonActive()


        ASCMAIN1.Progress("")
        grdSOFCSTMX.Update()
        grdSOFCSTMX.Refresh()
        Me.Cursor = Cursors.Default
    End Sub

    Sub Print_Report()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'CR_params.Add("GROUP_BY", optGroupBy.Value)
        'Generate_Report("PMRVIST1", "Open Site Visit Report")
        'Print_Report_End()
    End Sub

    Private Sub Fill_Extra_Fields()
        Dim RecTotal As Int64 = dst.Tables("SOTCSTMX").Rows.Count
        Dim OnRow As Int64 = 0
        Dim PCT As String = ""

        For Each rowSOTCSTMX As DataRow In dst.Tables("SOTCSTMX").Select()
            OnRow += 1
            PCT = Format((OnRow / RecTotal) * 100, "###,###") & " %"
            ASCMAIN1.Progress("Fetching Buyers & Sales", PCT)

            Dim CUST_CODE As String = rowSOTCSTMX.Item("CUST_CODE").ToString()
            Dim BLIST As List(Of String) = GetBuyers(CUST_CODE)
            Dim Filter As String = "CUST_CODE = '" & CUST_CODE & "'"
            Dim rowSOTCUSTX As DataRow = dst.Tables.Item("SOTCUSTX").Select(Filter).FirstOrDefault()

            rowSOTCSTMX.Item("BUYER1") = BLIST.Item(0)
            rowSOTCSTMX.Item("BUYER2") = BLIST.Item(1)
            rowSOTCSTMX.Item("BUYER3") = BLIST.Item(2)
            rowSOTCSTMX.Item("BUYER_EMAIL1") = BLIST.Item(3)
            rowSOTCSTMX.Item("BUYER_EMAIL2") = BLIST.Item(4)
            rowSOTCSTMX.Item("BUYER_EMAIL3") = BLIST.Item(5)
            If Not IsNothing(rowSOTCUSTX) Then
                rowSOTCSTMX.Item("ORDRED_TY") = Val(rowSOTCUSTX.Item("ORDERED_TY").ToString() & "")
                rowSOTCSTMX.Item("SHIPPED_TY") = Val(rowSOTCUSTX.Item("SHIPPED_TY").ToString() & "")
                rowSOTCSTMX.Item("CANCELLED_TY") = Val(rowSOTCUSTX.Item("CANCELLED_TY").ToString() & "")
                rowSOTCSTMX.Item("ORDRED_LY") = Val(rowSOTCUSTX.Item("ORDERED_LY").ToString() & "")
                rowSOTCSTMX.Item("SHIPPED_LY") = Val(rowSOTCUSTX.Item("SHIPPED_LY").ToString() & "")
                rowSOTCSTMX.Item("CANCELLED_LY") = Val(rowSOTCUSTX.Item("CANCELLED_LY").ToString() & "")
            Else
                rowSOTCSTMX.Item("ORDRED_TY") = 0
                rowSOTCSTMX.Item("SHIPPED_TY") = 0
                rowSOTCSTMX.Item("CANCELLED_TY") = 0
                rowSOTCSTMX.Item("ORDRED_LY") = 0
                rowSOTCSTMX.Item("SHIPPED_LY") = 0
                rowSOTCSTMX.Item("CANCELLED_LY") = 0
            End If
        Next
    End Sub

    Sub Setup_SOTCSTMX()
        If grdSOFCSTMX.ActiveRow Is Nothing OrElse (Not grdSOFCSTMX.ActiveRow.IsDataRow Or grdSOFCSTMX.ActiveRow.IsAddRow) Then

        Else
            Dim dvw As DataView = DirectCast(grdSOFCSTMD.DataSource, DataTable).DefaultView
            Dim CUST_CODE As String = grdSOFCSTMX.ActiveRow.Cells("CUST_CODE").Value & ""
            Dim FILTER As String = "CUST_CODE = '" & CUST_CODE & "'"
            If chkHideQuotes.Checked Then
                FILTER = FILTER & " AND ORDR_STATUS <> 'Q'"
            End If
            dvw.RowFilter = FILTER
            grdSOFCSTMD.Text = "Customer " & CUST_CODE
            For Each grow As UltraWinGrid.UltraGridRow In grdSOFCSTMD.Rows
                If chkByPeriods.Checked Then
                    Dim Periods As String() = CalculatePeriods()
                    Dim thisPeriod As String = grow.Cells.Item("ORDR_YYYYPP_BOOKED").Text
                    If thisPeriod >= Periods(0) And thisPeriod <= Periods(1) Then
                        grow.Appearance.BackColor = Drawing.Color.Khaki
                    Else
                        grow.Appearance.BackColor = Drawing.Color.Bisque
                    End If
                Else
                    Dim Dates As String() = CalculateDates()
                    Dim thisDate As DateTime = CDate(grow.Cells.Item("ORDR_DATE").Text)
                    If thisDate >= CDate(Dates(0)) And thisDate <= CDate(Dates(1)) Then
                        grow.Appearance.BackColor = Drawing.Color.Khaki
                    Else
                        grow.Appearance.BackColor = Drawing.Color.Bisque
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub grdSOFCSTMX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOFCSTMX.AfterRowActivate
        Setup_SOTCSTMX()
    End Sub

    Private Function GetBuyers(ByVal CUST_CODE As String) As List(Of String)
        Dim RetVal As New List(Of String)
        Dim RecCnt As Integer = 0
        'Dim sql As New Text.StringBuilder With {.Length = 0}
        'sql.AppendLine("SELECT ARTCUSTD.*")
        'sql.AppendLine("FROM ARTCUSTD")
        'sql.AppendLine("WHERE CUST_CODE = '" & CUST_CODE & "'")
        'sql.AppendLine("AND CONTACT_TYPE = 'B'")
        'Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty, "V")
        Dim Filter As String = String.Format("CUST_CODE = '{0}'", CUST_CODE)
        For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(Filter, "CONTACT_NO")
            RecCnt += 1
            RetVal.Add(rowARTCUSTD.Item("CONTACT_NAME").ToString())
            If RecCnt >= 2 Then
                Exit For
            End If
        Next
        For i As Integer = RecCnt To 2
            RetVal.Add("")
            RecCnt += 1
        Next
        For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(Filter, "CONTACT_NO")
            RecCnt += 1
            RetVal.Add(rowARTCUSTD.Item("CONTACT_EMAIL").ToString())
            If RecCnt >= 5 Then
                Exit For
            End If
        Next
        For i As Integer = RecCnt To 5
            RetVal.Add("")
            RecCnt += 1
        Next
        Return RetVal
    End Function

    Private Function CalculateDates(Optional ForOracle As Boolean = True) As String()
        Dim RetVal As String()
        ReDim RetVal(3)
        Dim FORMATDT As String = "dd-MMM-yyyy"
        If ForOracle Then
            FORMATDT = "dd-MMM-yyyy"
        End If
        RetVal(0) = Format(dt1.DateTime, FORMATDT)
        RetVal(1) = Format(dt2.DateTime, FORMATDT)
        RetVal(2) = Format(DateSerial(dt1.DateTime.Year - 1, dt1.DateTime.Month, dt1.DateTime.Day), FORMATDT)
        RetVal(3) = Format(DateSerial(dt2.DateTime.Year - 1, dt2.DateTime.Month, dt2.DateTime.Day), FORMATDT)
        Return RetVal
    End Function

    Private Function CalculatePeriods() As String()
        Dim RetVal As String()
        ReDim RetVal(3)
        If Absx1.cmbFor("RYP0").Value <> "" Then
            RetVal(0) = Absx1.cmbFor("RYP0").Value
            RetVal(1) = Absx1.cmbFor("RYP1").Value
            RetVal(2) = ASCMAIN1.Period_Calc(Absx1.cmbFor("RYP0").Value, -12)
            RetVal(3) = ASCMAIN1.Period_Calc(Absx1.cmbFor("RYP1").Value, -12)
        End If
        Return RetVal
    End Function

    Private Sub cmdRefresh_Click(sender As Object, e As EventArgs) Handles cmdRefresh.Click
        Setup_Summary()
    End Sub

    Private Sub chkHideZeros_CheckedChanged(sender As Object, e As EventArgs) Handles chkHideZeros.CheckedChanged
        filterNonActive()
    End Sub

    Private Sub filterNonActive()
        Dim Filter As String = ""
        If chkHideZeros.Checked Then
            Filter = "ORDRED_TY <> 0 OR  SHIPPED_TY <> 0 OR CANCELLED_TY <> 0 OR ORDRED_LY <> 0 OR  SHIPPED_LY <> 0 OR CANCELLED_LY <> 0"
        End If
        Dim dvw As DataView = DirectCast(grdSOFCSTMX.DataSource, DataTable).DefaultView
        dvw.RowFilter = String.Format(Filter)
    End Sub

    Sub Setup_SOTORDR2()
        Dim dvw As DataView = DirectCast(grdSOTORDR2.DataSource, DataTable).DefaultView
        Dim ORDR_NO As String = "XXXXXXXXXX"
        If grdSOFCSTMD.ActiveRow Is Nothing OrElse (Not grdSOFCSTMD.ActiveRow.IsDataRow) Then
            Fill_Records("SOTORDR2", ORDR_NO, True)
        Else
            ORDR_NO = grdSOFCSTMD.ActiveRow.Cells("ORDR_NO").Value
            Fill_Records("SOTORDR2", ORDR_NO, True)
        End If
    End Sub

    Private Sub grdSOFCSTMD_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOFCSTMD.AfterRowActivate
        Setup_SOTORDR2()
    End Sub

    Private Sub chkShowOrderDetails_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowOrderDetails.CheckedChanged
        SetShowOrderDetails()
    End Sub
    Private Sub SetShowOrderDetails()
        SplitContainer2.AutoSize = True
        If chkShowOrderDetails.Checked Then
            SplitContainer2.Panel2.Show()
            SplitContainer2.Panel2Collapsed = False
            'SplitContainer2.Height = 395
        Else
            SplitContainer2.Panel2.Hide()
            SplitContainer2.Panel2Collapsed = True
            'SplitContainer2.Height = 395 / 2
        End If
    End Sub

    Private Sub grdSOFCSTMX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSOFCSTMX.InitializeRow
        Dim VARIANCE_TY As String = e.Row.Cells("VARIANCE_TY").Value
        Dim VARIANCE_LY As Int64 = Val(e.Row.Cells("VARIANCE_LY").Value & "")

        If VARIANCE_TY = 0 Then
            e.Row.Cells("VARIANCE_TY").Appearance.BackColor = Drawing.Color.Empty
        Else
            If VARIANCE_TY > 0 Then
                e.Row.Cells("VARIANCE_TY").Appearance.BackColor = Drawing.Color.LightGreen
            Else
                e.Row.Cells("VARIANCE_TY").Appearance.BackColor = Drawing.Color.Tomato
            End If
        End If

        If VARIANCE_LY = 0 Then
            e.Row.Cells("VARIANCE_LY").Appearance.BackColor = Drawing.Color.Empty
        Else
            If VARIANCE_TY > 0 Then
                e.Row.Cells("VARIANCE_LY").Appearance.BackColor = Drawing.Color.LightGreen
            Else
                e.Row.Cells("VARIANCE_LY").Appearance.BackColor = Drawing.Color.Tomato
            End If
        End If

    End Sub

    Private Sub chkByPeriods_CheckedChanged(sender As Object, e As EventArgs) Handles chkByPeriods.CheckedChanged
        If chkByPeriods.Checked Then
            UltraExplorerBar1.Groups("Dates").Visible = False
            UltraExplorerBar1.Groups("Periods").Visible = True
        Else
            UltraExplorerBar1.Groups("Dates").Visible = True
            UltraExplorerBar1.Groups("Periods").Visible = False
        End If
    End Sub

    Private Function MAKE_SQL(ByVal TABLE_NAME As String, FOR_INIT As Boolean) As String
        Dim RETVAL As String = ""
        Dim SQLB As New System.Text.StringBuilder With {.Length = 0}
        Dim Dates As String() = CalculateDates()
        Dim Periods As String() = CalculatePeriods()

        Select Case TABLE_NAME
            Case "SOTCSTMD"
                SQLB.Length = 0
                SQLB.AppendLine("SELECT")
                SQLB.AppendLine("S1.ORDR_NO,")
                SQLB.AppendLine("S1.ORDR_GROUP_NO,")
                SQLB.AppendLine("S1.ORDR_DATE,")
                SQLB.AppendLine("S1.ORDR_DATE_RECD,")
                SQLB.AppendLine("S1.CUST_CODE,")
                SQLB.AppendLine("S1.CUST_NAME,")
                SQLB.AppendLine("S1.CUST_STORE_NAME,")
                SQLB.AppendLine("S1.ORDR_CUST_PO,")
                SQLB.AppendLine("S1.ORDR_SHIP_DATE,")
                SQLB.AppendLine("S1.ORDR_CANCEL_DATE,")
                SQLB.AppendLine("S1.ORDR_STATUS,")
                SQLB.AppendLine("S1.ORDR_YYYYPP_BOOKED,")
                SQLB.AppendLine("SUM(NVL(S2.ORDR_QTY,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS ORDRED,")
                SQLB.AppendLine("SUM(nvl(S2.ORDR_QTY_SHIP,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SHIPPED,")
                SQLB.AppendLine("SUM(nvl(S2.ORDR_QTY_CANC,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS CANCELLED")
                SQLB.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2, ARTCUST1 C1")
                SQLB.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
                SQLB.AppendLine("AND S1.CUST_CODE = C1.CUST_CODE")
                If ASCMAIN1.DBS_COMPANY <> "RGI" Then
                    If Not Remote.IsUserSuper Then
                        MAKE_SR_FILTER("C1", SQLB)
                    End If
                End If
                SQLB.AppendLine("AND (S1.ORDR_STATUS <> 'C' AND S1.ORDR_STATUS <> 'D')")
                If Not FOR_INIT Then
                    If chkByPeriods.Checked Then
                        SQLB.AppendLine(String.Format("AND ((S1.ORDR_YYYYPP_BOOKED >= '{0}' AND S1.ORDR_YYYYPP_BOOKED <= '{1}')", Periods(0), Periods(1)))
                        SQLB.AppendLine(String.Format("OR (S1.ORDR_YYYYPP_BOOKED >= '{0}' AND S1.ORDR_YYYYPP_BOOKED <= '{1}'))", Periods(2), Periods(3)))
                    Else
                        SQLB.AppendLine(String.Format("AND ((S1.ORDR_DATE >= '{0}' AND S1.ORDR_DATE <= '{1}')", Dates(0), Dates(1)))
                        SQLB.AppendLine(String.Format("OR (S1.ORDR_DATE >= '{0}' AND S1.ORDR_DATE <= '{1}'))", Dates(2), Dates(3)))
                    End If
                End If
                SQLB.AppendLine("GROUP BY")
                SQLB.AppendLine("S1.ORDR_NO,")
                SQLB.AppendLine("S1.ORDR_GROUP_NO,")
                SQLB.AppendLine("S1.ORDR_DATE,")
                SQLB.AppendLine("S1.ORDR_DATE_RECD,")
                SQLB.AppendLine("S1.CUST_CODE,")
                SQLB.AppendLine("S1.CUST_NAME,")
                SQLB.AppendLine("S1.CUST_STORE_NAME,")
                SQLB.AppendLine("S1.ORDR_CUST_PO,")
                SQLB.AppendLine("S1.ORDR_SHIP_DATE,")
                SQLB.AppendLine("S1.ORDR_CANCEL_DATE,")
                SQLB.AppendLine("S1.ORDR_STATUS,")
                SQLB.AppendLine("S1.ORDR_YYYYPP_BOOKED")
                RETVAL = SQLB.ToString
            Case "SOTCUSTX"
                SQLB.Length = 0
                SQLB.AppendLine("SELECT CUST_CODE,")
                SQLB.AppendLine("SUM(ORDERED_TY) AS ORDERED_TY,")
                SQLB.AppendLine("SUM(SHIPPED_TY) AS SHIPPED_TY,")
                SQLB.AppendLine("SUM(CANCELLED_TY) AS CANCELLED_TY,")
                SQLB.AppendLine("SUM(ORDERED_LY) AS ORDERED_LY,")
                SQLB.AppendLine("SUM(SHIPPED_LY) AS SHIPPED_LY,")
                SQLB.AppendLine("SUM(CANCELLED_LY) AS CANCELLED_LY")
                SQLB.AppendLine("FROM(")
                SQLB.AppendLine("  SELECT")
                SQLB.AppendLine("  S1.CUST_CODE,")
                SQLB.AppendLine("  SUM(NVL(S2.ORDR_QTY,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS ORDERED_TY,")
                SQLB.AppendLine("  SUM(NVL(S2.ORDR_QTY_SHIP,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SHIPPED_TY,")
                SQLB.AppendLine("  SUM(NVL(S2.ORDR_QTY_CANC,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS CANCELLED_TY,")
                SQLB.AppendLine("  SUM(0) ORDERED_LY,")
                SQLB.AppendLine("  SUM(0) SHIPPED_LY,")
                SQLB.AppendLine("  SUM(0) CANCELLED_LY")
                SQLB.AppendLine("  FROM SOTORDR1 S1, SOTORDR2 S2, ARTCUST1 C1")
                SQLB.AppendLine("  WHERE S1.ORDR_NO = S2.ORDR_NO")
                SQLB.AppendLine("  AND S1.CUST_CODE = C1.CUST_CODE")
                If ASCMAIN1.DBS_COMPANY = "RGO" Then
                    If Not Remote.IsUserSuper Then
                        MAKE_SR_FILTER("C1", SQLB)
                    End If
                End If
                SQLB.AppendLine("  AND S1.ORDR_STATUS <> 'C'")
                If Not FOR_INIT Then
                    If chkByPeriods.Checked Then
                        SQLB.AppendLine(String.Format("  AND S1.ORDR_YYYYPP_BOOKED >= '{0}'", Periods(0)))
                        SQLB.AppendLine(String.Format("  AND S1.ORDR_YYYYPP_BOOKED <= '{0}'", Periods(1)))
                    Else
                        SQLB.AppendLine(String.Format("  AND S1.ORDR_DATE >= '{0}'", Dates(0)))
                        SQLB.AppendLine(String.Format("  AND S1.ORDR_DATE <= '{0}'", Dates(1)))
                    End If
                End If
                SQLB.AppendLine("  GROUP BY S1.CUST_CODE")
                SQLB.AppendLine("  UNION")
                SQLB.AppendLine("  SELECT")
                SQLB.AppendLine("  S1.CUST_CODE,")
                SQLB.AppendLine("  SUM(0) ORDERED_TY,")
                SQLB.AppendLine("  SUM(0) SHIPPED_TY,")
                SQLB.AppendLine("  SUM(0) CANCELLED_TY,")
                SQLB.AppendLine("  SUM(NVL(S2.ORDR_QTY,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS ORDERED_LY,")
                SQLB.AppendLine("  SUM(NVL(S2.ORDR_QTY_SHIP,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SHIPPED_LY,")
                SQLB.AppendLine("  SUM(NVL(S2.ORDR_QTY_CANC,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS CANCELLED_LY")
                SQLB.AppendLine("  FROM SOTORDR1 S1, SOTORDR2 S2, ARTCUST1 C1")
                SQLB.AppendLine("  WHERE S1.ORDR_NO = S2.ORDR_NO")
                SQLB.AppendLine("  AND S1.CUST_CODE = C1.CUST_CODE")
                If ASCMAIN1.DBS_COMPANY = "RGO" Then
                    If Not Remote.IsUserSuper Then
                        MAKE_SR_FILTER("C1", SQLB)
                    End If
                End If
                SQLB.AppendLine("  AND S1.ORDR_STATUS <> 'C'")
                If Not FOR_INIT Then
                    If chkByPeriods.Checked Then
                        SQLB.AppendLine(String.Format("  AND S1.ORDR_YYYYPP_BOOKED >= '{0}'", Periods(2)))
                        SQLB.AppendLine(String.Format("  AND S1.ORDR_YYYYPP_BOOKED <= '{0}'", Periods(3)))
                    Else
                        SQLB.AppendLine(String.Format("  AND S1.ORDR_DATE >= '{0}'", Dates(2)))
                        SQLB.AppendLine(String.Format("  AND S1.ORDR_DATE <= '{0}'", Dates(3)))
                    End If

                End If
                SQLB.AppendLine("  GROUP BY S1.CUST_CODE")
                SQLB.AppendLine(")")
                SQLB.AppendLine("GROUP BY CUST_CODE")
                RETVAL = SQLB.ToString
        End Select

        Return RETVAL
    End Function
End Class