Imports System.ComponentModel
Imports System.Text
Imports Infragistics.Win.UltraWinEditors
Imports Infragistics.Win.UltraWinGrid

Public Class ECFSTYL1
    Private SQL As New StringBuilder With {.Length = 0}
    Private rowECTSTYL1 As DataRow
    Private STYLE_CODE As String = ""
    Private FormLoading As Boolean = True
    Private RecordLoading As Boolean = False
    Private colorsLocked As Boolean = True
    Private partnerCount As Integer = 0
    Private EC_PARM_IMAGES_FOLDER As String = ""
    Private ECTPSTY2_INIT As Boolean = False
    Private DEFAULT_SET_QTY_ORIG As Int64
    Private SEL_ECOM_CODE As String = ""
    Dim PartnerNewCbo As New List(Of String)
    Private WebLinks As New Dictionary(Of String, String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        FormLoading = True

        If MENU_ITEM_OBJECT = "ECTSTYLI" Then
            InquiryMode = True
        End If

        Get_PARM("ECTPARM1")

        initDST()

        setGridDataSources()

        setGridSummaries()

        setGridSorts()

        initFillDST()

        setBinding()

        setGridDefaults()

        setGridValueLists()

        Setup_ECTESTY3()

        'Set_Read_Only(grpTotals, True)

        'grpHeader.Visible = False

        If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
            UltraTabControl2.Tabs("Upsert Styles").Enabled = True
        Else
            UltraTabControl2.Tabs("Upsert Styles").Enabled = False
        End If

        UltraTabControl2.Tabs("Upsert Styles").Enabled = True

        FormLoading = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                'Validate_Code("STYLE_CODE")

                'If Absx1.txtFor("STYLE_CODE").Value & "" = "" Then
                '    EMsg &= vbCr & "Missing Style Code"
                'End If
            Case "View"
                If Absx1.txtFor("STYLE_CODE").Text & "" = "" Then
                    EMsg &= vbCr & "You Must Select A Style To View"
                Else
                    Dim STYLE_CODE_ORIG As String = Absx1.txtFor("STYLE_CODE").Text.ToUpper
                    Dim STYLE_CODE_NEW As String = Absx1.txtFor("STYLE_CODE").Text.ToUpper
                    Dim FOUND_MATCH As Boolean = False
                    If IsValidEcomStyle(STYLE_CODE_NEW) Then
                        STYLE_CODE_ORIG = STYLE_CODE_NEW
                        FOUND_MATCH = True
                    Else
                        STYLE_CODE_NEW = String.Format("MT{0}", STYLE_CODE_NEW)
                        If IsValidEcomStyle(STYLE_CODE_NEW) Then
                            STYLE_CODE_ORIG = STYLE_CODE_NEW
                            FOUND_MATCH = True
                        End If
                    End If
                    If FOUND_MATCH Then
                        If Absx1.txtFor("STYLE_CODE").Text <> STYLE_CODE_ORIG Then
                            Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE_ORIG
                        End If
                    Else
                        EMsg &= vbCr & "Invalid Style Code Selected"
                    End If
                End If
            Case "Update"
                Dim BadPctFound As Boolean = False
                For Each rowECTESTY1 As DataRow In dst.Tables("ECTESTY1").Select()
                    If IsNumeric(rowECTESTY1.Item("ALT_UNIT_PCT") & String.Empty) Then
                        If Val(rowECTESTY1.Item("ALT_UNIT_PCT") & String.Empty) > 100 Or
                            Val(rowECTESTY1.Item("ALT_UNIT_PCT") & String.Empty) < 0 Then
                            BadPctFound = True
                        End If
                    End If
                Next

                If BadPctFound Then
                    EMsg &= vbCr & "Pct Off > 100 or < 0"
                End If

            Case "Cancel"
                'Removed Per Kevin 8/31/18
                'If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                '          "You may have made Changes") = MsgBoxResult.No Then
                '    Exit Sub
                'End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            'Case "New"
            '    EntryMode = "N"
            '    Load_Record()
            '    Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)
                Fill_Records("ECTECOM1_PARTNER")

            Case "Cancel"
                Mode_Settings(False)

            'Case "Done"
            '    Mode_Settings(False)

            Case "Test"
                Dim Emsg As String = ""
                Dim TEST = New TAC.EDC84601(dst.Tables.Item("EDT846O1"), dst.Tables.Item("EDT846O2"), dst.Tables.Item("EDTSYSIH"))
                Dim TestOut As String = TEST.CreateEDI846("OVERSTOCK", Emsg)
                Stop
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    '.Items("New").Settings.Enabled = ScreenMode
                    .Items("View").Settings.Enabled = ScreenMode
                    .Items("Update").Settings.Enabled = Not ScreenMode
                    .Items("Cancel").Settings.Enabled = Not ScreenMode
                    '.Items("Done").Settings.Enabled = iScreenMode
                    If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
                        .Items("Test").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Test").Settings.Enabled = DefaultableBoolean.False
                    End If
                End With
                .Groups("Partners").Visible = Not ScreenMode
                If UltraTabControl2.Tabs("Sales Analysis").Selected Then
                    .Groups("Sales Analysis Options").Visible = Not ScreenMode
                Else
                    .Groups("Sales Analysis Options").Visible = False
                End If

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        'SplitContainer1.Visible = ScreenMode

        Panel1.Visible = Not ScreenMode
        'UltraTabControl1.Visible = Not ScreenMode

        If ScreenMode Then

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        Dim clearTables As New List(Of String)
        clearTables.Add("ASTAUDT1")
        clearTables.Add("ECTECOM2")
        clearTables.Add("ECTECOM2_FILL")
        clearTables.Add("ECTESTY1")
        clearTables.Add("ECTESTY2")
        clearTables.Add("ECTESTY3")
        clearTables.Add("ECTPSTY2")
        clearTables.Add("ICTSTYL1")
        clearTables.Add("ICTSTYLD")
        clearTables.Add("ECTSTYB1")
        clearTables.Add("ECTSTYL1")
        clearTables.Add("SOTCSTY1")

        For Each clearTable As String In clearTables
            dst.Tables(clearTable).Rows.Clear()
        Next

        EnforceConstraints(True)

        Absx1.txtFor("STYLE_CODE").Text = ""
        txtPKG_CODE.Text = ""
        STYLE_CODE = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        RecordLoading = True

        Save_Header_Fields(UltraGroupBox1)

        STYLE_CODE = Absx1.txtFor("STYLE_CODE").Text & ""

        Dim loadTables As New List(Of String)
        loadTables.Add("ECTESTY1")
        loadTables.Add("ECTESTY2")
        loadTables.Add("ECTESTY3")
        loadTables.Add("ECTPSTY2")
        loadTables.Add("ICTSTYL1")
        loadTables.Add("ICTSTYLD")
        loadTables.Add("ECTSTYB1")
        loadTables.Add("ECTSTYL1")
        loadTables.Add("SOTCSTY1")
        loadTables.Add("ECTSALS1")
        grdECTSALS1.Text = String.Format("Sales For Style {0}", STYLE_CODE)
        If EntryMode = "N" Then
            rowECTSTYL1 = dst.Tables("ECTSTYL1").NewRow
            rowECTSTYL1.Item("STYLE_CODE") = STYLE_CODE
            rowECTSTYL1.Item("DEFAULT_SET_QTY") = 1
            DEFAULT_SET_QTY_ORIG = 1
            dst.Tables("ECTSTYL1").Rows.Add(rowECTSTYL1)
        Else
            For Each loadTable As String In loadTables
                Fill_Records(loadTable, STYLE_CODE)
            Next
            If dst.Tables("ECTSTYL1").Rows.Count = 0 Then
                rowECTSTYL1 = dst.Tables("ECTSTYL1").NewRow
                rowECTSTYL1.Item("STYLE_CODE") = STYLE_CODE
                rowECTSTYL1.Item("DEFAULT_SET_QTY") = 1
                dst.Tables("ECTSTYL1").Rows.Add(rowECTSTYL1)
            Else
                If dst.Tables("ECTSTYL1").Rows.Count = 1 Then
                    DEFAULT_SET_QTY_ORIG = Val(dst.Tables("ECTSTYL1").Rows(0).Item("DEFAULT_SET_QTY") & String.Empty)
                End If
            End If
            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine(" FROM ASTAUDT1")
            SQL.AppendLine(" WHERE TABLE_NAME = 'ECTESTY1'")
            SQL.AppendLine(String.Format(" AND KEY_VALUE = '{0}'", STYLE_CODE))
            Fill_Records("ASTAUDT1",, True, SQL.ToString)
        End If


        'fillCUST_STYLE()

        init_ECTECOM2()

        init_Misc()
        init_grdECTPSTY2()
        init_grdECTECOM1_PARTNER()
        init_grdECTSTYB2()
        init_Images()

        DEFAULT_SET_QTY_ORIG = 0

        grdECTPSTY2.Update()
        grdECTPSTY2.Refresh()

        RecordLoading = False
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Dim updateTables As New List(Of String)
        updateTables.Add("ASTAUDT1")
        updateTables.Add("ECTECOM2")
        updateTables.Add("ECTESTY1")
        updateTables.Add("ECTESTY2")
        updateTables.Add("ECTESTY3")
        updateTables.Add("ECTSTYB1")
        updateTables.Add("ECTSTYL1")
        updateTables.Add("SOTCSTY1")

        BeginTrans()

        For Each updateTable As String In updateTables
            Call Update_Record_TDA(updateTable)
        Next

        CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        Dim deleteTables As New List(Of String)
        'deleteTables.Add("")

        BeginTrans()
        For Each deleteTable As String In deleteTables
            'Delete_Records(deleteTable)
        Next

        CommitTrans("Delete Complete")
    End Sub

    Private Sub ICFSTYL1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            If (e.KeyCode = Keys.NumPad1 Or e.KeyCode = Keys.D1) And e.Alt Then
                Call Click_Command("Update", e)
            End If
        End If
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Dim BTNS As String = "SSBB"
        WebLinks.Add("Show Filter", "")
        WebLinks.Add("Show GroupBox", "")
        WebLinks.Add("Show 846 History", "")
        WebLinks.Add("Style Status Inquiry", "")
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT ECOM_CODE, ECOM_STYLE_URL")
        sql.AppendLine("FROM ECTECOM1")
        sql.AppendLine("WHERE NVL(ECOM_STYLE_URL,'NULL') <> 'NULL'")
        Dim tblECTECOM1 As DataTable = ASCDATA1.GetDataTable(sql.ToString())
        For Each rowECTECOM1 As DataRow In tblECTECOM1.Rows
            WebLinks.Add(String.Format("View On {0}", rowECTECOM1.Item("ECOM_CODE").ToString & String.Empty), rowECTECOM1.Item("ECOM_STYLE_URL").ToString & String.Empty)
            BTNS = BTNS & "B"
        Next
        Load_Popup_Menu(grdECTSTYLX, BTNS, WebLinks.Keys.ToArray)
        'Load_Popup_Menu(grdECTSTYLX, "SSB", "Show Filter", "Show GroupBox", "Show 846 History")
        Load_Popup_Menu(grdECTECOM1_PARTNER, "SSBBB", "Show Filter", "Show GroupBox", "Show Only Selected", "Show All", "Audit Trail")
        Load_Popup_Menu(grdECTSTYB2, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdECTSALS1, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdECTSALSX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdECUPSERT, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdECTECOM1_FILTER, "BB", "Select All", "Select None")
        Load_Popup_Menu(grdICTEDI01, "SS", "Show Filter", "Show GroupBox", "Show 846 History")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case = "grdECTECOM1_PARTNER"
                Case = "grdECTSTYLX"
                    tlb_btn = DirectCast(tlb_pop.Tools("Show 846 History"), UltraWinToolbars.ButtonTool)
                    If SEL_ECOM_CODE = "" Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        tlb_btn.SharedProps.Visible = True
                    End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

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
            Case "Show Only Selected"
                Dim PCNT As Integer = dst.Tables.Item("ECTECOM1_PARTNER").Select.Count
                If PCNT > 0 Then
                    Dim ECOM_CODE As String = grd.ActiveRow.Cells.Item("ECOM_CODE").Text & String.Empty
                    If ECOM_CODE.Length > 0 Then
                        For Each grdCol As UltraGridColumn In grdECTPSTY2.DisplayLayout.Bands(0).Columns
                            If grdCol.Header.Tag <> "" Then
                                If grdCol.Header.Tag = ECOM_CODE Then
                                    grdCol.Hidden = False
                                Else
                                    grdCol.Hidden = True
                                End If
                            End If
                        Next
                    End If
                End If
            Case "Show All"
                For Each grdCol As UltraGridColumn In grdECTPSTY2.DisplayLayout.Bands(0).Columns
                    If Not IsNothing(grdCol.Header.Tag) Then
                        grdCol.Hidden = False
                    End If
                Next
            Case "Show 846 History"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim COLOR_CODE As String = grd.ActiveRow.Cells("COLOR_CODE").Text
                Dim ECOM_CODE As String = grd.ActiveRow.Cells("ECOM_CODE").Text
                Dim Frm As New ECF84601(Me, ECOM_CODE, STYLE_CODE, COLOR_CODE)
                Frm.ShowDialog()
            Case "Audit Trail"
                Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text
                Dim ECOM_CODE As String = grd.ActiveRow.Cells("ECOM_CODE").Text
                Dim Frm As New ECFAUDT1(Me, STYLE_CODE, ECOM_CODE)
                Frm.ShowDialog()
            Case "Select All"
                For Each rowECTECOM1_FILTER As DataRow In dst.Tables("ECTECOM1_FILTER").Select()
                    rowECTECOM1_FILTER.Item("SEL") = "1"
                Next
            Case "Select None"
                For Each rowECTECOM1_FILTER As DataRow In dst.Tables("ECTECOM1_FILTER").Select()
                    rowECTECOM1_FILTER.Item("SEL") = "0"
                Next
        End Select

        If e.Tool.Key.ToString.Length >= 7 Then
            If e.Tool.Key.ToString.Substring(0, 7) = "View On" Then
                Dim ECOM_CODE As String = e.Tool.Key.ToString.Replace("View On ", "")
                If WebLinks.ContainsKey("View On " & ECOM_CODE) Then
                    Dim ECOM_STYLE_URL_STRING As String = WebLinks.Item("View On " & ECOM_CODE)
                    Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                    SQLS.AppendLine("SELECT CUST_CODE")
                    SQLS.AppendLine("FROM ECTECOM1")
                    SQLS.AppendLine(String.Format("WHERE ECOM_CODE = '{0}'", ECOM_CODE))
                    ASCMAIN1.sql = SQLS.ToString()
                    Dim CUST_CODE As String = ASCDATA1.GetDataValue
                    SQLS.Length = 0
                    SQLS.AppendLine("SELECT CUST_STYLE_CODE")
                    SQLS.AppendLine("FROM SOTCSTY1")
                    SQLS.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
                    SQLS.AppendLine(String.Format("AND STYLE_CODE = '{0}'", grd.ActiveRow.Cells("STYLE_CODE").Text))
                    SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", grd.ActiveRow.Cells("COLOR_CODE").Text))
                    SQLS.AppendLine("ORDER BY LAST_DATE DESC")
                    ASCMAIN1.sql = SQLS.ToString()
                    Dim CUST_STYLE_CODE As String = ASCDATA1.GetDataValue
                    If CUST_STYLE_CODE.Length = 0 Then
                        MsgBox("No Customer Style Code Setup", vbOKOnly, "Customer Style Code")
                    Else
                        Dim OVERRIDE_CODE As Boolean = False
                        Select Case ECOM_CODE
                            Case "QVC"
                                If IsNumeric(CUST_STYLE_CODE.Substring(CUST_STYLE_CODE.Length - 1, 1)) Then
                                    Dim NEWCAR As String = Val(CUST_STYLE_CODE.Substring(CUST_STYLE_CODE.Length - 1, 1)) - 1
                                    CUST_STYLE_CODE = CUST_STYLE_CODE.Substring(0, CUST_STYLE_CODE.Length - 1) & NEWCAR
                                End If
                            Case "OVERSTOCK"
                                If CUST_STYLE_CODE.EndsWith("-000-000") Then
                                    CUST_STYLE_CODE = CUST_STYLE_CODE.Replace("-000-000", "")
                                End If
                            Case "WALMART"
                                OVERRIDE_CODE = True
                                SQLS.Length = 0
                                SQLS.AppendLine("SELECT MIN(ALT_ITEM_CODE) ALT_ITEM_CODE")
                                SQLS.AppendLine("FROM ECTESTY2")
                                SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", grd.ActiveRow.Cells("STYLE_CODE").Text))
                                SQLS.AppendLine(String.Format("AND COLOR_CODE = 'TTGR'", grd.ActiveRow.Cells("COLOR_CODE").Text))
                                SQLS.AppendLine("AND ECOM_CODE = 'WALMART'")
                                ASCMAIN1.sql = SQLS.ToString()
                                Dim ALT_ITEM_CODE As String = ASCDATA1.GetDataValue
                                If ALT_ITEM_CODE.Length > 0 Then
                                    Dim LURL As String = String.Format(ECOM_STYLE_URL_STRING, ALT_ITEM_CODE)
                                    Dim pInfo As ProcessStartInfo = New ProcessStartInfo(LURL)
                                    Process.Start(pInfo)
                                End If
                        End Select
                        If Not OVERRIDE_CODE Then
                            Dim LURL As String = String.Format(ECOM_STYLE_URL_STRING, CUST_STYLE_CODE)
                            Dim pInfo As ProcessStartInfo = New ProcessStartInfo(LURL)
                            Process.Start(pInfo)
                        End If

                    End If
                Else
                    MsgBox("Bad Record for Customer In Ecom", vbOKOnly, "Problem With Customer")
                End If
            End If
        End If
    End Sub
#End Region

#Region "ABSColumn Controls"
    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "STYLE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode Then
                        Click_Command("View", e)
                    End If
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "STYLE_CODE"
                If Not InquiryMode Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "SOME_COLUMN"
        End Select
    End Sub
#End Region

#Region "Form Controls"

    Private Sub btnAddStyles_Click(sender As Object, e As EventArgs) Handles btnAddStyles.Click
        Dim emsg As String = ""
        Dim needsPricing As Boolean = False
        Dim NeedsType As Boolean = False
        For Each rowECTESTY1_NEW As DataRow In dst.Tables("ECTESTY1_NEW").Select()
            If rowECTESTY1_NEW.Item("SHIP_ECOM").ToString & String.Empty = "" And rowECTESTY1_NEW.Item("SHIP_DROP").ToString & String.Empty = "" Then
                NeedsType = True
            End If
            If rowECTESTY1_NEW.Item("SHIP_ECOM").ToString & String.Empty = "0" And rowECTESTY1_NEW.Item("SHIP_DROP").ToString & String.Empty = "0" Then
                NeedsType = True
            End If
            If Val(rowECTESTY1_NEW.Item("ECOM_UNIT_PRICE").ToString & String.Empty) = 0 Then
                needsPricing = True
            End If
        Next
        If needsPricing Then
            emsg += vbCrLf & "Some Styles Have No Pricing."
        End If
        If NeedsType Then
            emsg += vbCrLf & "Some Styles Have E-Commerce or Drop Ship Designation."
        End If
        If emsg.Length > 0 Then
            MsgBox(emsg.ToString, vbOKOnly, "Please Fix The Following.")
        Else
            For Each rowECTESTY1_NEW As DataRow In dst.Tables("ECTESTY1_NEW").Select()
                Dim newECTESTY1 As DataRow = dst.Tables.Item("ECTESTY1").NewRow
                For Each dc As DataColumn In dst.Tables.Item("ECTESTY1").Columns
                    newECTESTY1.Item(dc.ColumnName) = rowECTESTY1_NEW.Item(dc.ColumnName)
                Next
                dst.Tables.Item("ECTESTY1").Rows.Add(newECTESTY1)
                Dim STYLE_FILTER As String = String.Format("STYLE_CODE = '{0}'", rowECTESTY1_NEW.Item("STYLE_CODE").ToString & String.Empty)
                For Each rowECTESTY2_NEW As DataRow In dst.Tables("ECTESTY2_NEW").Select(STYLE_FILTER)
                    Dim newECTESTY2 As DataRow = dst.Tables.Item("ECTESTY2").NewRow
                    For Each dc As DataColumn In dst.Tables.Item("ECTESTY2").Columns
                        If dc.ColumnName <> "ECOM_PARTNER_SKU" Then
                            newECTESTY2.Item(dc.ColumnName) = rowECTESTY2_NEW.Item(dc.ColumnName)
                        End If
                    Next
                    dst.Tables.Item("ECTESTY2").Rows.Add(newECTESTY2)
                Next
            Next
            dst.Tables.Item("ECTESTY2_NEW").Clear()
            dst.Tables.Item("ECTESTY1_NEW").Clear()
            btnAddStyles.Visible = False
            btnCancelStyles.Visible = False
            Update_Record()
            initFillDST()
        End If
    End Sub

    Private Sub btnCancelStyles_Click(sender As Object, e As EventArgs) Handles btnCancelStyles.Click
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Cancel Adding New Styles?"
        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
        iMSG.AppendLine("This Will Clear The Grid Below")
        iMSG.AppendLine("Reset The Process Of Adding New")
        iMSG.AppendLine("Styles.")
        iMSG.AppendLine("")
        iMSG.AppendLine("Is That What You Want?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then
            dst.Tables.Item("ECTESTY2_NEW").Clear()
            dst.Tables.Item("ECTESTY1_NEW").Clear()
            btnAddStyles.Visible = False
            btnCancelStyles.Visible = False
        End If
    End Sub

    Private Sub btnRefreshX_Click(sender As Object, e As EventArgs) Handles btnRefreshX.Click
        If UltraTabControl2.Tabs("Sales Analysis").Selected Then
            refreshECTSALSX()
        Else
            refreshECTSTYLX()
        End If
    End Sub

    Private Sub btnSelStyles_Click(sender As Object, e As EventArgs) Handles btnSelStyles.Click
        Dim StylesAdded As Boolean = False
        Dim S As New Text.StringBuilder With {.Length = 0}
        S.AppendLine("SELECT STYLE_CODE, STYLE_STATUS, STYLE_DESC, STYLE_PRICE")
        S.AppendLine("FROM ICTSTYL1")
        S.AppendLine("WHERE STYLE_CODE NOT IN")
        S.AppendLine("(")
        S.AppendLine("SELECT STYLE_CODE FROM ECTESTY1")
        S.AppendLine(")")
        With ASCMAIN1.CodeSelector
            .SQL = S.ToString
            .MultipleSelections = True
            .PreviouslySelectedCodes0 = ""
            .Caption = "Pick Styles To Add"
            .TABLE_NAME = ""
            .VIEW_NAME = ""
            .VIEW_DESC = ""
            .COLUMN_NAME = ""
            .COLUMN_PREKEYs = New Dictionary(Of String, String)
            .Custom_sql_where = ""
            .tblASTVIEW1 = New DataTable
        End With
        Dim F As New ASFCODE1
        F.ShowDialog()
        If ASCMAIN1.CodeSelector.Selections <> 0 Then
            For Each dr As DataRow In ASCMAIN1.CodeSelector.SelectedRows
                Dim STYLE_CODE As String = dr.Item("STYLE_CODE").ToString & String.Empty
                Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                SQLS.AppendLine("SELECT DEFAULT_SET_QTY")
                SQLS.AppendLine("FROM ECTSTYL1")
                SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                ASCMAIN1.sql = SQLS.ToString()
                Dim DEFAULT_SET_QTY As Int64 = Val(ASCDATA1.GetDataValue & String.Empty)
                'Dim rowECTSTYL1 As DataRow = LookUp("ECTSTYL1", STYLE_CODE)
                'Dim DEFAULT_SET_QTY As Int64 = Val(rowECTSTYL1.Item("DEFAULT_SET_QTY").ToString & String.Empty)
                If DEFAULT_SET_QTY < 1 Then
                    DEFAULT_SET_QTY = 1
                End If
                Dim filter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
                If dst.Tables.Item("ECTESTY1_NEW").Select(filter).Count = 0 Then
                    StylesAdded = True
                    Dim newECTESTY1_NEW As DataRow = dst.Tables.Item("ECTESTY1_NEW").NewRow
                    newECTESTY1_NEW.Item("STYLE_CODE") = STYLE_CODE
                    newECTESTY1_NEW.Item("ECOM_CODE") = cboPartnerNew.Text
                    newECTESTY1_NEW.Item("ECOM_STYLE_STATUS") = "A"
                    newECTESTY1_NEW.Item("ECOM_UNIT_PRICE") = 0.00
                    newECTESTY1_NEW.Item("ALT_UNIT_PCT") = 0.00
                    newECTESTY1_NEW.Item("ALT_UNIT_PRICE") = 0.00
                    newECTESTY1_NEW.Item("SHIP_ECOM") = "0"
                    newECTESTY1_NEW.Item("SHIP_DROP") = "0"
                    newECTESTY1_NEW.Item("STYLE_DESC") = dr.Item("STYLE_DESC").ToString & String.Empty
                    newECTESTY1_NEW.Item("STYLE_PRICE") = Val(dr.Item("STYLE_PRICE").ToString & String.Empty)
                    newECTESTY1_NEW.Item("SET_QTY") = DEFAULT_SET_QTY
                    dst.Tables.Item("ECTESTY1_NEW").Rows.Add(newECTESTY1_NEW)

                    Dim sql As New Text.StringBuilder With {.Length = 0}
                    sql.AppendLine("SELECT *")
                    sql.AppendLine("FROM ICTSTYC1")
                    sql.AppendLine("WHERE STYLE_CODE = :PARM1")
                    Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty, "V", STYLE_CODE)
                    For Each rowICTSTYC2 As DataRow In tbl.Rows
                        Dim COLOR_CODE As String = rowICTSTYC2.Item("COLOR_CODE").ToString & String.Empty
                        Dim newECTESTY2_NEW As DataRow = dst.Tables.Item("ECTESTY2_NEW").NewRow
                        newECTESTY2_NEW.Item("STYLE_CODE") = STYLE_CODE
                        newECTESTY2_NEW.Item("COLOR_CODE") = COLOR_CODE
                        newECTESTY2_NEW.Item("ECOM_CODE") = cboPartnerNew.Text
                        newECTESTY2_NEW.Item("ECOM_STYLE_COLOR_STATUS") = "A"
                        dst.Tables.Item("ECTESTY2_NEW").Rows.Add(newECTESTY2_NEW)
                    Next
                End If
            Next
        End If
        If StylesAdded Then
            btnAddStyles.Visible = True
            btnCancelStyles.Visible = True
        End If
    End Sub

    Private Sub txtPKG_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtPKG_CODE.ValueChanged
        Dim PKG_CODE As String = txtPKG_CODE.Text
        Dim rowECTSTYL1 As DataRow = dst.Tables.Item("ECTSTYL1").Select.FirstOrDefault
        If Not IsNothing(rowECTSTYL1) Then
            rowECTSTYL1.Item("PKG_CODE") = PKG_CODE
        End If
    End Sub

    Private Sub UltraTabControl2_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl2.SelectedTabChanged
        Select Case UltraTabControl2.SelectedTab.Text
            Case "Existing Styles"
                UltraExplorerBar1.Groups("Sales Analysis Options").Visible = False
                UltraExplorerBar1.Groups("Partners").Visible = True
            Case "Sales Analysis"
                UltraExplorerBar1.Groups("Sales Analysis Options").Visible = True
                UltraExplorerBar1.Groups("Partners").Visible = True
            Case "Add New Styles"
                UltraExplorerBar1.Groups("Sales Analysis Options").Visible = False
                UltraExplorerBar1.Groups("Partners").Visible = False
            Case "Upsert Styles"
                UltraExplorerBar1.Groups("Sales Analysis Options").Visible = False
                UltraExplorerBar1.Groups("Partners").Visible = False
            Case "EDI Inventory"
                UltraExplorerBar1.Groups("Sales Analysis Options").Visible = False
                UltraExplorerBar1.Groups("Partners").Visible = False
            Case Else
                UltraExplorerBar1.Groups("Sales Analysis Options").Visible = False
                UltraExplorerBar1.Groups("Partners").Visible = True
        End Select

        'If UltraTabControl2.Tabs("Sales Analysis").Selected Then
        '    UltraExplorerBar1.Groups("Sales Analysis Options").Visible = True
        'Else
        '    UltraExplorerBar1.Groups("Sales Analysis Options").Visible = False
        'End If
    End Sub
#End Region

#Region "Custom Methods"

    Private Sub fillPartnerSKUs()
        For Each rowECTSTYLX As DataRow In dst.Tables("ECTSTYLX").Select()
            Dim STYLE_CODE As String = rowECTSTYLX.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowECTSTYLX.Item("COLOR_CODE").ToString & String.Empty
            Dim ECOM_CODE As String = rowECTSTYLX.Item("ECOM_CODE").ToString & String.Empty
            Dim CUST_CODE As String = ""
            Dim rowECTECOM1_FILTER As DataRow = dst.Tables.Item("ECTECOM1_FILTER").Select(String.Format("ECOM_CODE = '{0}'", ECOM_CODE)).FirstOrDefault
            If Not IsNothing(rowECTECOM1_FILTER) Then
                CUST_CODE = rowECTECOM1_FILTER.Item("CUST_CODE").ToString & String.Empty
                If CUST_CODE.Length <> 0 Then
                    Fill_Records("SOTCSTY1", STYLE_CODE)
                    'If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                    '    If STYLE_CODE = "MT21078" And COLOR_CODE = "ANWH" Then Stop
                    'End If
                    Dim filter As String = String.Format("CUST_CODE = '{0}' AND STYLE_CODE = '{1}' AND COLOR_CODE = '{2}'", CUST_CODE, STYLE_CODE, COLOR_CODE)
                    Dim rowSOTCSTY1 As DataRow = dst.Tables("SOTCSTY1").Select(filter, "LAST_DATE DESC").FirstOrDefault
                    If Not IsNothing(rowSOTCSTY1) Then
                        rowECTSTYLX.Item("ECOM_PARTNER_SKU") = rowSOTCSTY1.Item("CUST_STYLE_CODE").ToString & String.Empty
                    End If
                End If
            End If
        Next

        grdECTSTYLX.Refresh()
        grdECTSTYLX.Update()
    End Sub

    Private Sub fillEDI846Data(ByVal SEL_ECOM_CODE As String)
        'If Not (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
        SEL_ECOM_CODE = SEL_ECOM_CODE.Replace("'", "")
        Dim rowEDTREFX As DataRow = dst.Tables.Item("EDTXREFX").Select(String.Format("ECOM_CODE = '{0}'", SEL_ECOM_CODE)).FirstOrDefault
        If Not IsNothing(rowEDTREFX) Then
            Dim EDI_SUPPLIER_NO As String = rowEDTREFX.Item("EDI_SUPPLIER_NO") & String.Empty
            Fill_Records("EDT846OX", EDI_SUPPLIER_NO)

            For Each rowECTSTYLX As DataRow In dst.Tables("ECTSTYLX").Select()
                Dim STYLE_CODE As String = rowECTSTYLX.Item("STYLE_CODE").ToString & String.Empty
                Dim COLOR_CODE As String = rowECTSTYLX.Item("COLOR_CODE").ToString & String.Empty
                'Dim ECOM_CODE As String = rowECTSTYLX.Item("ECOM_CODE").ToString & String.Empty

                Dim filter As String = String.Format("EDI_STYLE = '{0}' AND EDI_COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
                Dim rowEDT846OX As DataRow = dst.Tables("EDT846OX").Select(filter).FirstOrDefault
                If Not IsNothing(rowEDT846OX) Then
                    If IsDate(rowEDT846OX.Item("EDI_REPORT_DATE").ToString & String.Empty) Then
                        rowECTSTYLX.Item("EDI_REPORT_DATE") = CDate(rowEDT846OX.Item("EDI_REPORT_DATE").ToString & String.Empty)
                    End If
                    rowECTSTYLX.Item("EDI_AVAIL_QTY") = rowEDT846OX.Item("EDI_AVAIL_QTY").ToString & String.Empty
                    rowECTSTYLX.Item("EDI_STATUS") = rowEDT846OX.Item("EDI_STATUS").ToString & String.Empty
                End If
            Next

            grdECTSTYLX.Refresh()
            grdECTSTYLX.Update()
            'End If
        End If
    End Sub

    Private Sub fillSTAT2Data()
        For Each rowECTSTYLX As DataRow In dst.Tables("ECTSTYLX").Select()
            Dim STYLE_CODE As String = rowECTSTYLX.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowECTSTYLX.Item("COLOR_CODE").ToString & String.Empty
            Fill_Records("ICTSTATC", New String() {STYLE_CODE, COLOR_CODE})
            If dst.Tables.Item("ICTSTATC").Rows.Count = 1 Then
                Dim rowICTSTATC As DataRow = dst.Tables.Item("ICTSTATC").Rows(0)
                rowECTSTYLX.Item("WHSE_QTY_ON_HAND") = Val(rowICTSTATC.Item("WHSE_QTY_ON_HAND").ToString & String.Empty)
                rowECTSTYLX.Item("WHSE_QTY_PICK") = Val(rowICTSTATC.Item("WHSE_QTY_PICK").ToString & String.Empty)
                rowECTSTYLX.Item("OPEN_TO_SELL") = Val(rowICTSTATC.Item("OPEN_TO_SELL").ToString & String.Empty)
                rowECTSTYLX.Item("WHSE_QTY_TRAN") = Val(rowICTSTATC.Item("WHSE_QTY_TRAN").ToString & String.Empty)
                rowECTSTYLX.Item("WHSE_QTY_ON_ORDER") = Val(rowICTSTATC.Item("WHSE_QTY_ON_ORDER").ToString & String.Empty)
                rowECTSTYLX.Item("WHSE_QTY_OPEN") = Val(rowICTSTATC.Item("WHSE_QTY_OPEN").ToString & String.Empty)
                rowECTSTYLX.Item("FUT_AVAIL") = Val(rowICTSTATC.Item("FUT_AVAIL").ToString & String.Empty)
            Else
                rowECTSTYLX.Item("WHSE_QTY_ON_HAND") = 0
                rowECTSTYLX.Item("WHSE_QTY_PICK") = 0
                rowECTSTYLX.Item("OPEN_TO_SELL") = 0
                rowECTSTYLX.Item("WHSE_QTY_TRAN") = 0
                rowECTSTYLX.Item("WHSE_QTY_ON_ORDER") = 0
                rowECTSTYLX.Item("WHSE_QTY_OPEN") = 0
                rowECTSTYLX.Item("FUT_AVAIL") = 0
            End If
        Next
    End Sub

    Private Function getCUST_STYLE_CODE(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal CUST_CODE As String) As String
        Dim RetVal As String = ""
        Fill_Records("SOTCSTY1", STYLE_CODE)
        Dim FILTER_SCE As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}' AND CUST_CODE = '{2}'", STYLE_CODE, COLOR_CODE, CUST_CODE)
        Dim rowSOTCSTY1 As DataRow = dst.Tables.Item("SOTCSTY1").Select(FILTER_SCE, "LAST_DATE DESC").FirstOrDefault
        If Not IsNothing(rowSOTCSTY1) Then
            RetVal = rowSOTCSTY1.Item("CUST_STYLE_CODE").ToString & String.Empty
        End If
        Return RetVal
    End Function

    Private Sub initDST()
        With dst
            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ECTPARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTPARM1", "**", 0, False)

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ECTECOM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTECOM1", "**", 0, False)
            .Tables("ECTECOM1").Columns.Add("SEL", GetType(System.String))

            Create_TDA(.Tables.Add, "ECTECOM1_FILTER", "**", 0, False)
            .Tables("ECTECOM1_FILTER").Columns.Add("SEL", GetType(System.String))

            Create_TDA(.Tables.Add, "ECTECOM1_PARTNER", "**", 0, False)
            With .Tables("ECTECOM1_PARTNER").Columns
                .Add("SEL", GetType(System.String))
                .Add("ECOM_UNIT_PRICE", GetType(System.Double))
                .Add("ALT_UNIT_PCT", GetType(System.Double))
                .Add("ALT_UNIT_PRICE", GetType(System.Double))
                .Add("SET_QTY", GetType(System.Int64))
                .Add("SHIP_ECOM", GetType(System.String))
                .Add("SHIP_DROP", GetType(System.String))
                .Add("ECOM_SET_PRICE", GetType(System.Double), "ISNULL(ECOM_UNIT_PRICE,0) * ISNULL(SET_QTY,0)")
                .Add("ECOM_SET_PRICE_ALT", GetType(System.Double), "ISNULL(ALT_UNIT_PRICE,0) * ISNULL(SET_QTY,0)")
            End With

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ECTECOM2")
            SQL.AppendLine("WHERE ECOM_CODE = :PARM1")
            SQL.AppendLine("AND STYLE_CODE = :PARM2")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTECOM2", "**", 0, True, "VV")
            Create_TDA(.Tables.Add, "ECTECOM2_FILL", "**", 0, False, "VV")

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("ICTSTYL1.STYLE_CODE,")
            SQL.AppendLine("ECTESTY2.COLOR_CODE,")
            'SQL.AppendLine("ECTESTY2.ECOM_PARTNER_SKU,")
            SQL.AppendLine("ICTSTYL1.STYLE_STATUS,")
            SQL.AppendLine("ICTSTYL1.STYLE_DESC,")
            SQL.AppendLine("ICTSTYL1.STYLE_CLASS_CODE,")
            SQL.AppendLine("NVL(ECTESTY1.SET_QTY,0) AS SET_QTY,")
            SQL.AppendLine("ECTESTY1.ECOM_CODE,")
            SQL.AppendLine("ECTESTY1.ECOM_STYLE_STATUS,")
            SQL.AppendLine("ECTESTY1.ECOM_UNIT_PRICE,")
            SQL.AppendLine("ECTESTY1.SHIP_ECOM,")
            SQL.AppendLine("ECTESTY1.SHIP_DROP")
            SQL.AppendLine("FROM ECTESTY1, ICTSTYL1, ECTESTY2")
            SQL.AppendLine("WHERE ECTESTY1.STYLE_CODE = ICTSTYL1.STYLE_CODE")
            SQL.AppendLine("AND ECTESTY1.STYLE_CODE = ECTESTY2.STYLE_CODE")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTSTYLX", "**", 0, False)
            With .Tables("ECTSTYLX").Columns
                .Add("SHORT_DESC", GetType(System.String))
                .Add("LONG_DESC", GetType(System.String))
                .Add("EDI_REPORT_DATE", GetType(System.DateTime))
                .Add("EDI_AVAIL_QTY", GetType(System.Int64))
                .Add("EDI_STATUS", GetType(System.String))
                .Add("ECOM_PARTNER_SKU", GetType(System.String))
                .Add("WHSE_QTY_ON_HAND", GetType(System.Int64))
                .Add("WHSE_QTY_PICK", GetType(System.Int64))
                .Add("OPEN_TO_SELL", GetType(System.Int64))
                .Add("WHSE_QTY_TRAN", GetType(System.Int64))
                .Add("WHSE_QTY_ON_ORDER", GetType(System.Int64))
                .Add("WHSE_QTY_OPEN", GetType(System.Int64))
                .Add("FUT_AVAIL", GetType(System.Int64))
            End With

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ECTESTY1")
            SQL.AppendLine("WHERE STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTESTY1", "**", 0, True, "V")
            Create_TDA(.Tables.Add("ECTESTY1_NEW"), "ECTESTY1", "**", 0, True, "V")
            With .Tables("ECTESTY1_NEW").Columns
                .Add("STYLE_DESC", GetType(System.String))
                .Add("STYLE_PRICE", GetType(System.Decimal))
            End With

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ECTESTY2")
            SQL.AppendLine("WHERE STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTESTY2", "**", 0, True, "V")
            With .Tables("ECTESTY2").Columns
                .Add("ECOM_PARTNER_SKU", GetType(System.String))
            End With
            Create_TDA(.Tables.Add("ECTESTY2_NEW"), "ECTESTY2", "**", 0, True, "V")

            dst.Relations.Add("ECTESTY1N_ECTESTY2N",
                              dst.Tables("ECTESTY1_NEW").Columns("STYLE_CODE"),
                              dst.Tables("ECTESTY2_NEW").Columns("STYLE_CODE"))

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ECTESTY3")
            SQL.AppendLine("WHERE STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTESTY3", "**", 0, True, "V")

            SQL.Length = 0
            SQL.AppendLine("")
            SQL.AppendLine("SELECT")
            SQL.AppendLine("C1.STYLE_CODE,")
            SQL.AppendLine("C1.COLOR_CODE,")
            SQL.AppendLine("C1.STYLE_COLOR_STATUS,")
            SQL.AppendLine("C1.UPC_CODE,")
            SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_ON_HAND,0)) AS WHSE_QTY_ON_HAND,")
            SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0) - NVL(S2.WHSE_QTY_OPEN,0)) AS NETPOS")
            SQL.AppendLine("FROM ICTSTYC1 C1, ICTSTAT2 S2")
            SQL.AppendLine("WHERE C1.STYLE_CODE = S2.STYLE_CODE (+)")
            SQL.AppendLine("AND C1.COLOR_CODE = S2.COLOR_CODE (+)")
            SQL.AppendLine("AND C1.STYLE_CODE = :PARM1")
            SQL.AppendLine("GROUP BY C1.STYLE_CODE,")
            SQL.AppendLine("C1.COLOR_CODE,")
            SQL.AppendLine("C1.STYLE_COLOR_STATUS,")
            SQL.AppendLine("C1.UPC_CODE")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTPSTY2", "**", 0, False, "V")
            ASCMAIN1.sql = "SELECT COUNT(*) FROM ECTECOM1"
            partnerCount = Val(ASCDATA1.GetDataValue)
            For i As Integer = 1 To partnerCount
                With .Tables("ECTPSTY2").Columns
                    .Add("ECOM_SEL_" & i, GetType(System.String))
                    .Add("ECOM_CODE_" & i, GetType(System.String))
                    .Add("ECOM_PCT_" & i, GetType(System.Int64))
                    .Add("ECOM_MIN_" & i, GetType(System.Int64))
                    .Add("ECOM_PRICE_" & i, GetType(System.Double))
                    .Add("ECOM_PARTNER_SKU_" & i, GetType(System.String))
                End With
            Next

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ICTSTYL1")
            SQL.AppendLine("WHERE STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 0, False, "V")

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ICTSTAT2")
            SQL.AppendLine("WHERE STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ICTSTAT2", "**", 0, False, "V")

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ECTSTYB1")
            SQL.AppendLine("WHERE STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTSTYB1", "**", 0, True, "V")

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ECTSTYB2")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTSTYB2", "**", 0, False)
            .Tables("ECTSTYB2").Columns.Add("SEL", GetType(System.String))

            SQL.Length = 0
            SQL.AppendLine("SELECT ICTSTYLD.*, ICTSTYLM.PACK_DESC")
            SQL.AppendLine("FROM ICTSTYLD, ICTSTYLM")
            SQL.AppendLine("WHERE ICTSTYLD.PACK_CODE = ICTSTYLM.PACK_CODE")
            SQL.AppendLine("AND ICTSTYLD.STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ICTSTYLD", "**", 0, True, "V")

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ECTSTYL1")
            SQL.AppendLine("WHERE STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTSTYL1", "**", 0, True, "V")

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM SOTCSTY1")
            SQL.AppendLine("WHERE STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "SOTCSTY1", "**", 0, True, "V")
            'Create_TDA(.Tables.Add, "SOTCSTY1", "**", 0, True)

            Dim EDI_PRE As String = ""
            If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
                EDI_PRE = "GEN."
            End If
            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine(String.Format("FROM {0}EDT846O1", EDI_PRE))
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "EDT846O1", "**", 0, False)

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine(String.Format("FROM {0}EDT846O2", EDI_PRE))
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "EDT846O2", "**", 0, False)

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM EDTSYSIH")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "EDTSYSIH", "**", 0, False)

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ECTUPSRT")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTUPSRT", "**", 0, False)
            Fill_Records("ECTUPSRT")

            'SQL.Length = 0
            'SQL.AppendLine("SELECT")
            'SQL.AppendLine("E2.EDI_STYLE,")
            'SQL.AppendLine("E2.EDI_COLOR_CODE,")
            'SQL.AppendLine("E1.EDI_REPORT_DATE,")
            'SQL.AppendLine("E2.EDI_AVAIL_QTY,")
            'SQL.AppendLine("DECODE(E2.EDI_MAINT_TYPE_CODE,'001','Active','In-Active') AS EDI_STATUS")
            'SQL.AppendLine(String.Format("FROM {0}EDT846O1 E1, {0}EDT846O2 E2", EDI_PRE))
            'SQL.AppendLine("WHERE E1.COMPANY_CODE = E2.COMPANY_CODE")
            'SQL.AppendLine("AND E1.EDI_OUTBOUND_DOC_NO = E2.EDI_OUTBOUND_DOC_NO")
            'SQL.AppendLine("AND E2.EDI_OUTBOUND_DOC_NO IN")
            'SQL.AppendLine("(")
            'SQL.AppendLine(String.Format("	SELECT MAX(I1.EDI_OUTBOUND_DOC_NO) FROM {0}EDT846O1 I1, {0}EDT846O2 I2", EDI_PRE))
            'SQL.AppendLine("	WHERE I1.COMPANY_CODE = I2.COMPANY_CODE")
            'SQL.AppendLine("	AND I1.EDI_OUTBOUND_DOC_NO = I2.EDI_OUTBOUND_DOC_NO")
            'SQL.AppendLine("	AND I1.EDI_SUPPLIER_NO = :PARM1")
            'SQL.AppendLine("	AND I2.EDI_STYLE = :PARM2")
            'SQL.AppendLine("	AND I2. EDI_COLOR_CODE = :PARM3")
            'SQL.AppendLine(")")
            'SQL.AppendLine("AND E2.EDI_STYLE = :PARM2")
            'SQL.AppendLine("AND E2.EDI_COLOR_CODE = :PARM3")
            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("E2.EDI_STYLE,")
            SQL.AppendLine("E2.EDI_COLOR_CODE,")
            SQL.AppendLine("E1.EDI_REPORT_DATE,")
            SQL.AppendLine("E2.EDI_AVAIL_QTY,")
            SQL.AppendLine("DECODE(E2.EDI_MAINT_TYPE_CODE,'001','Active','Inactive') AS EDI_STATUS")
            SQL.AppendLine("FROM GEN.EDT846O1 E1, GEN.EDT846O2 E2,")
            SQL.AppendLine("(")
            SQL.AppendLine("  SELECT I2.EDI_STYLE, I2. EDI_COLOR_CODE, MAX(I1.EDI_OUTBOUND_DOC_NO) AS EDI_OUTBOUND_DOC_NO")
            SQL.AppendLine("  FROM GEN.EDT846O1 I1, GEN.EDT846O2 I2")
            SQL.AppendLine("  WHERE I1.COMPANY_CODE = I2.COMPANY_CODE")
            SQL.AppendLine("  AND I1.EDI_OUTBOUND_DOC_NO = I2.EDI_OUTBOUND_DOC_NO")
            SQL.AppendLine("  AND I1.EDI_SUPPLIER_NO = :PARM1")
            SQL.AppendLine("  GROUP BY I2.EDI_STYLE, I2. EDI_COLOR_CODE")
            SQL.AppendLine(") E3")
            SQL.AppendLine("WHERE E1.COMPANY_CODE = E2.COMPANY_CODE")
            SQL.AppendLine("AND E1.EDI_OUTBOUND_DOC_NO = E2.EDI_OUTBOUND_DOC_NO")
            SQL.AppendLine("AND E2.EDI_STYLE = E3.EDI_STYLE")
            SQL.AppendLine("AND E2.EDI_COLOR_CODE = E3.EDI_COLOR_CODE")
            SQL.AppendLine("AND E2.EDI_OUTBOUND_DOC_NO = E3.EDI_OUTBOUND_DOC_NO")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "EDT846OX", "**", 0, False, "V")

            SQL.Length = 0
            SQL.AppendLine("SELECT C1.ECOM_CODE, X4.EDI_SUPPLIER_NO")
            SQL.AppendLine("FROM EDTXREF4 X4, EDTTRPM1 PM, ECTECOM1 C1")
            SQL.AppendLine("WHERE X4.SENDER_ID_QUAL = PM.EDI_TP_QUAL")
            SQL.AppendLine("AND X4.SENDER_ID = PM.EDI_TP_ID")
            SQL.AppendLine("AND X4.SENDER_ID = C1.EDI_TP_ID")
            SQL.AppendLine("AND PM.EDI_DOC_NO = '846'")
            SQL.AppendLine("AND X4.WHSE_CODE = 'MS'")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "EDTXREFX", "**", 0, False)

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("I1.INV_NO,")
            SQL.AppendLine("I1.ORDR_NO,")
            SQL.AppendLine("I1.CUST_CODE,")
            SQL.AppendLine("C1.CUST_NAME,")
            SQL.AppendLine("I1.ORDR_CUST_PO,")
            SQL.AppendLine("O1.EDI_PO_TYPE,")
            SQL.AppendLine("I1.INV_DATE,")
            SQL.AppendLine("I1.ORDR_YYYYPP_UPDATED,")
            SQL.AppendLine("I2.STYLE_CODE,")
            SQL.AppendLine("I2.COLOR_CODE,")
            SQL.AppendLine("I2.ORDR_UNIT_PRICE,")
            SQL.AppendLine("I2.ORDR_QTY_SHIP,")
            SQL.AppendLine("(I2.ORDR_UNIT_PRICE * I2.ORDR_QTY_SHIP) SHIP_TOTAL")
            SQL.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2, ARTCUST1 C1, SOTORDR1 O1")
            SQL.AppendLine("WHERE I1.INV_NO = I2.INV_NO")
            SQL.AppendLine("AND I1.INV_TYPE = I2.INV_TYPE")
            SQL.AppendLine("AND I1.CUST_CODE = C1.CUST_CODE")
            SQL.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO")
            SQL.AppendLine("AND I2.STYLE_CODE = :PARM1")
            SQL.AppendLine("AND I1.INV_TYPE = 'I'")
            SQL.AppendLine("AND I1.WHSE_CODE = 'MS'")
            SQL.AppendLine("AND I1.CUST_CODE IN (")
            SQL.AppendLine("SELECT DISTINCT CUST_CODE FROM ECTECOM1)")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTSALS1", "**", 0, False, "V")

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("I1.INV_NO,")
            SQL.AppendLine("I1.ORDR_NO,")
            SQL.AppendLine("I1.CUST_CODE,")
            SQL.AppendLine("C1.CUST_NAME,")
            SQL.AppendLine("I1.ORDR_CUST_PO,")
            SQL.AppendLine("O1.EDI_PO_TYPE,")
            SQL.AppendLine("I1.INV_DATE,")
            SQL.AppendLine("I1.ORDR_YYYYPP_UPDATED,")
            SQL.AppendLine("I2.STYLE_CODE,")
            SQL.AppendLine("I2.COLOR_CODE,")
            SQL.AppendLine("(I2.STYLE_CODE || '-' || I2.COLOR_CODE) AS STYLE_COLOR,")
            SQL.AppendLine("I2.ORDR_UNIT_PRICE,")
            SQL.AppendLine("I2.ORDR_QTY_SHIP,")
            SQL.AppendLine("(I2.ORDR_UNIT_PRICE * I2.ORDR_QTY_SHIP) SHIP_TOTAL")
            SQL.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2, ARTCUST1 C1, SOTORDR1 O1")
            SQL.AppendLine("WHERE I1.INV_NO = I2.INV_NO")
            SQL.AppendLine("AND I1.INV_TYPE = I2.INV_TYPE")
            SQL.AppendLine("AND I1.CUST_CODE = C1.CUST_CODE")
            SQL.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO")
            SQL.AppendLine("AND I1.INV_TYPE = 'I'")
            SQL.AppendLine("AND I1.WHSE_CODE = 'MS'")
            SQL.AppendLine("AND I1.CUST_CODE IN (")
            SQL.AppendLine("SELECT DISTINCT CUST_CODE FROM ECTECOM1)")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTSALSX", "**", 0, False)

            'SQL.AppendLine("SELECT")
            'SQL.AppendLine("'X' UPSERT_TYPE,")
            'SQL.AppendLine("E1.ECOM_CODE,")
            'SQL.AppendLine("E1.STYLE_CODE,")
            'SQL.AppendLine("E2.COLOR_CODE,")
            'SQL.AppendLine("E1.ECOM_STYLE_STATUS,")
            'SQL.AppendLine("DECODE(NVL(E1.SHIP_ECOM,0),1,'Y','N') SHIP_ECOM,")
            'SQL.AppendLine("DECODE(NVL(E1.SHIP_DROP,0),1,'Y','N') SHIP_DROP,")
            'SQL.AppendLine("NVL(E1.SET_QTY,1) SET_QTY,")
            'SQL.AppendLine("E1.ECOM_UNIT_PRICE,")
            'SQL.AppendLine("NVL(E1.ALT_UNIT_PCT,0) ALT_UNIT_PCT,")
            'SQL.AppendLine("NVL(E1.ALT_UNIT_PRICE,E1.ECOM_UNIT_PRICE) ALT_UNIT_PRICE,")
            'SQL.AppendLine("NVL(E1.ECOM_MIN_QTY_OVERRIDE,C1.ECOM_MIN_QTY_DEFAULT) ECOM_MIN_QTY_OVERRIDE")
            'SQL.AppendLine("FROM ECTESTY1 E1, ECTESTY2 E2, ECTECOM1 C1")
            'SQL.AppendLine("WHERE E1.STYLE_CODE = E2.STYLE_CODE")
            'SQL.AppendLine("AND E1.ECOM_CODE = E2.ECOM_CODE")
            'SQL.AppendLine("AND E1.ECOM_CODE = C1.ECOM_CODE")
            'SQL.AppendLine("AND E1.ECOM_CODE = :PARM1")
            SQL.Length = 0
            SQL.AppendLine("SELECT EC_PARM_EXT_IMAGE_FOLDER")
            SQL.AppendLine("FROM ECTPARM1")
            SQL.AppendLine("WHERE EC_PARM_KEY = 'Z'")
            ASCMAIN1.sql = SQL.ToString()
            Dim EC_PARM_EXT_IMAGE_FOLDER As String = ASCDATA1.GetDataValue

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("E1.ECOM_CODE, ")
            SQL.AppendLine("E1.STYLE_CODE, ")
            SQL.AppendLine("SM.STYLE_DESC, ")
            SQL.AppendLine("E2.COLOR_CODE,")
            SQL.AppendLine("C1.COLOR_DESC, ")
            SQL.AppendLine("E1.ECOM_STYLE_STATUS,")
            SQL.AppendLine("E2.ECOM_STYLE_COLOR_STATUS,")
            SQL.AppendLine("E1.ECOM_UNIT_PRICE,")
            SQL.AppendLine("E1.SET_QTY,")
            SQL.AppendLine("NVL(E1.SHIP_ECOM,0) SHIP_ECOM,")
            SQL.AppendLine("NVL(E1.SHIP_DROP,0) SHIP_DROP,")
            SQL.AppendLine("E1.ALT_UNIT_PCT,")
            SQL.AppendLine("E1.ALT_UNIT_PRICE,")
            SQL.AppendLine("NVL(E1.ECOM_MIN_QTY_OVERRIDE,M1.ECOM_MIN_QTY_DEFAULT) ECOM_MIN_QTY_OVERRIDE,")
            SQL.AppendLine("NVL(S1.SHORT_DESC,'') SHORT_DESC,")
            SQL.AppendLine("NVL(S1.LONG_DESC,'') LONG_DESC,")
            SQL.AppendLine("NVL(S1.PKG_CODE,'') PKG_CODE,")
            SQL.AppendLine("P1.PKG_DESC,")
            SQL.AppendLine("P1.PKG_WT,")
            SQL.AppendLine("P1.PKG_L,")
            SQL.AppendLine("P1.PKG_W,")
            SQL.AppendLine("P1.PKG_H,")
            SQL.AppendLine("P1.PKG_SEQ,")
            SQL.AppendLine("P1.PKG_CUBE,")
            SQL.AppendLine("P1.PKG_COST,")
            SQL.AppendLine("P1.PKG_CHARGE,")
            SQL.AppendLine("SM.STYLE_MATL_DESC")
            SQL.AppendLine("FROM ECTESTY1 E1, ECTESTY2 E2, ECTSTYL1 S1, ECTECOM1 M1, ICTSTYL1 SM, ICTCOLR1 C1, WHTPKGM1 P1")
            SQL.AppendLine("WHERE SM.STYLE_CODE = E1.STYLE_CODE")
            SQL.AppendLine("AND E2.COLOR_CODE = C1.COLOR_CODE")
            SQL.AppendLine("AND E1.STYLE_CODE = E2.STYLE_CODE")
            SQL.AppendLine("AND E1.STYLE_CODE = S1.STYLE_CODE (+)")
            SQL.AppendLine("AND E1.ECOM_CODE = M1.ECOM_CODE")
            SQL.AppendLine("AND E2.ECOM_CODE = M1.ECOM_CODE")
            SQL.AppendLine("AND S1.PKG_CODE = P1.PKG_CODE (+)")
            SQL.AppendLine("AND M1.ECOM_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECUPSERT", "**", 0, False, "V")
            With .Tables("ECUPSERT").Columns
                .Add("UPSERT_TYPE", GetType(System.String))
                .Add("UPSERT_ERROR", GetType(System.String))
                Dim URL_CALC As String = String.Format("'{0}' + STYLE_CODE + '-' + COLOR_CODE + '.jpg'", EC_PARM_EXT_IMAGE_FOLDER)
                .Add("IMG_URL", GetType(System.String), URL_CALC)
                .Add("CUST_STYLE_CODE", GetType(System.String))
            End With

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("STYLE_CODE,")
            SQL.AppendLine("COLOR_CODE,")
            SQL.AppendLine("NVL(WHSE_QTY_ON_HAND,0) WHSE_QTY_ON_HAND,")
            SQL.AppendLine("NVL(WHSE_QTY_PICK,0) WHSE_QTY_PICK,")
            SQL.AppendLine("(NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0)) OPEN_TO_SELL,")
            SQL.AppendLine("NVL(WHSE_QTY_TRAN,0) WHSE_QTY_TRAN,")
            SQL.AppendLine("NVL(WHSE_QTY_ON_ORDER,0) WHSE_QTY_ON_ORDER,")
            SQL.AppendLine("NVL(WHSE_QTY_OPEN,0) WHSE_QTY_OPEN,")
            SQL.AppendLine("(NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0) + NVL(WHSE_QTY_TRAN,0) + NVL(WHSE_QTY_ON_ORDER,0) - NVL(WHSE_QTY_OPEN,0)) FUT_AVAIL")
            SQL.AppendLine("FROM ICTSTAT2")
            SQL.AppendLine("WHERE WHSE_CODE = 'MS'")
            SQL.AppendLine("AND STYLE_CODE = :PARM1")
            SQL.AppendLine("AND COLOR_CODE = :PARM2")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ICTSTATC", "**", 0, False, "VV")

            SQL.Length = 0
            'SQL.AppendLine("SELECT")
            'SQL.AppendLine("E2.ECOM_CODE,")
            'SQL.AppendLine("E2.STYLE_CODE,")
            'SQL.AppendLine("E2.COLOR_CODE,")
            'SQL.AppendLine("S1.STYLE_DESC,")
            'SQL.AppendLine("E2.ECOM_STYLE_COLOR_STATUS,")
            'SQL.AppendLine("C1.STYLE_COLOR_STATUS")
            'SQL.AppendLine("FROM ECTESTY2 E2, ICTSTYL1 S1, ICTSTYC1 C1")
            'SQL.AppendLine("WHERE E2.STYLE_CODE = S1.STYLE_CODE")
            'SQL.AppendLine("AND  E2.STYLE_CODE = C1.STYLE_CODE")
            'SQL.AppendLine("AND E2.COLOR_CODE = C1.COLOR_CODE")
            'SQL.AppendLine("AND E2.ECOM_CODE = :PARM1")
            SQL.AppendLine("SELECT")
            SQL.AppendLine("E2.ECOM_CODE,")
            SQL.AppendLine("E2.STYLE_CODE,")
            SQL.AppendLine("E2.COLOR_CODE,")
            SQL.AppendLine("S1.STYLE_DESC,")
            SQL.AppendLine("E2.ECOM_STYLE_COLOR_STATUS,")
            SQL.AppendLine("C1.STYLE_COLOR_STATUS,")
            SQL.AppendLine("E1.SET_QTY,")
            SQL.AppendLine("DECODE(NVL(E1.ECOM_MIN_QTY_OVERRIDE,-1),-1,C1.ECOM_MIN_QTY_DEFAULT,E1.ECOM_MIN_QTY_OVERRIDE) AS ECOM_MIN_QTY_OVERRIDE")
            SQL.AppendLine("FROM ECTESTY1 E1, ECTESTY2 E2, ICTSTYL1 S1, ICTSTYC1 C1, ECTECOM1 C1")
            SQL.AppendLine("WHERE E1.STYLE_CODE = E2.STYLE_CODE")
            SQL.AppendLine("AND E1.ECOM_CODE = E2.ECOM_CODE")
            SQL.AppendLine("AND E2.STYLE_CODE = S1.STYLE_CODE")
            SQL.AppendLine("AND  E2.STYLE_CODE = C1.STYLE_CODE")
            SQL.AppendLine("AND E2.COLOR_CODE = C1.COLOR_CODE")
            SQL.AppendLine("AND E1.ECOM_CODE = C1.ECOM_CODE")
            SQL.AppendLine("AND E2.ECOM_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ICTEDI01", "**", 0, False, "V")
            With .Tables("ICTEDI01").Columns
                .Add("EDI_SUPPLIER_NO", GetType(System.String))
                .Add("EDI_REPORT_DATE", GetType(System.String))
                .Add("EDI_AVAIL_QTY", GetType(System.Int64))
                .Add("EDI_STATUS", GetType(System.String))
                .Add("WHSE_QTY_ON_HAND", GetType(System.Int64))
                .Add("WHSE_QTY_PICK", GetType(System.Int64))
                .Add("OPEN_TO_SELL", GetType(System.Int64))
                .Add("WHSE_QTY_TRAN", GetType(System.Int64))
                .Add("WHSE_QTY_ON_ORDER", GetType(System.Int64))
                .Add("WHSE_QTY_OPEN", GetType(System.Int64))
                .Add("FUT_AVAIL", GetType(System.Int64))
                .Add("ECOM_QTY", GetType(System.Int64))
            End With

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ASTAUDT1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add("ASTAUDT1_U"), "ASTAUDT1", "**", 0, True)

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ECTESTY1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add("ECTESTY1_U"), "ECTESTY1", "**", 0, True)

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ECTESTY2")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add("ECTESTY2_U"), "ECTESTY2", "**", 0, True)

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ECTSTYL1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add("ECTSTYL1_U"), "ECTSTYL1", "**", 0, True)

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM SOTCSTY1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add("SOTCSTY1_U"), "SOTCSTY1", "**", 0, True)

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("E1.ECOM_CODE,")
            SQL.AppendLine("S1.CUST_CODE,")
            SQL.AppendLine("S1.STYLE_CODE,")
            SQL.AppendLine("S1.COLOR_CODE,")
            SQL.AppendLine("MAX(S1.CUST_STYLE_CODE) CUST_STYLE_CODE")
            SQL.AppendLine("FROM SOTCSTY1 S1, ECTECOM1 E1")
            SQL.AppendLine("WHERE S1.CUST_CODE = E1.CUST_CODE")
            SQL.AppendLine("GROUP BY")
            SQL.AppendLine("E1.ECOM_CODE,")
            SQL.AppendLine("S1.CUST_CODE,")
            SQL.AppendLine("S1.STYLE_CODE,")
            SQL.AppendLine("S1.COLOR_CODE")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "SOTCSTYX", "**", 0, False)
            Fill_Records("SOTCSTYX")

            'SQL.Length = 0
            'SQL.AppendLine("SELECT *")
            'SQL.AppendLine(" FROM ASTAUDT1")
            'SQL.AppendLine(" WHERE TABLE_NAME = 'ECTESTY1'")
            'SQL.AppendLine(" AND KEY_VALUE = :PARM1")
            'ASCMAIN1.sql = SQL.ToString()
            'Create_TDA(.Tables.Add, "ASTAUDT1", "**", 0, True, "V", 0)
        End With
    End Sub

    Private Sub init_ECTECOM2()
        Dim ECOM_CODE As String = "RGI"
        Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text
        Fill_Records("ECTECOM2", New String() {ECOM_CODE, STYLE_CODE})
        If dst.Tables.Item("ECTECOM2").Rows.Count = 0 Then
            Fill_Records("ECTECOM2_FILL", New String() {ECOM_CODE, ECOM_CODE})
        End If
        If dst.Tables.Item("ECTECOM2_FILL").Rows.Count > 1 Then
            For i As Int64 = 0 To dst.Tables.Item("ECTECOM2_FILL").Rows.Count - 1
                Dim rowECTECOM2 As DataRow = dst.Tables.Item("ECTECOM2").NewRow
                Dim rowECTECOM2_FILL As DataRow = dst.Tables.Item("ECTECOM2_FILL").Rows(i)
                For Each dc As DataColumn In dst.Tables.Item("ECTECOM2_FILL").Columns
                    Select Case dc.ColumnName
                        Case "STYLE_CODE"
                            rowECTECOM2.Item(dc.ColumnName) = STYLE_CODE
                        Case "COLUMN_DESC"
                            Select Case rowECTECOM2_FILL.Item(dc.ColumnName).ToString & String.Empty
                                Case "UPC Code"
                                    rowECTECOM2.Item(dc.ColumnName) = rowECTECOM2_FILL.Item(dc.ColumnName)
                                    rowECTECOM2.Item("DATA_VALUE") = "<<COLOR SPECIFIC>>"
                                Case Else
                                    rowECTECOM2.Item(dc.ColumnName) = rowECTECOM2_FILL.Item(dc.ColumnName)
                            End Select
                        Case Else
                            rowECTECOM2.Item(dc.ColumnName) = rowECTECOM2_FILL.Item(dc.ColumnName)
                    End Select
                Next
                dst.Tables.Item("ECTECOM2").Rows.Add(rowECTECOM2)
            Next
        End If
    End Sub

    Private Sub initFillDST()
        Fill_Records("ECTPARM1")
        Fill_Records("ECTECOM1")
        Fill_Records("ECTECOM1_FILTER")
        Fill_Records("ECTECOM1_PARTNER")
        Fill_Records("ECTSTYB2")
        Fill_Records("EDTXREFX")
        'Fill_Records("EDT846OX")
        'If Not (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then

        'End If

        fillPartnerCbos()

        For Each rowECTECOM1_FILTER As DataRow In dst.Tables("ECTECOM1_FILTER").Select()
            rowECTECOM1_FILTER.Item("SEL") = "1"
        Next
        refreshECTSTYLX(True)
    End Sub

    Private Sub fillPartnerCbos()
        PartnerNewCbo.Clear()
        For Each rowECTECOM1 As DataRow In dst.Tables("ECTECOM1").Select("", "ECOM_CODE")
            PartnerNewCbo.Add(rowECTECOM1.Item("ECOM_CODE").ToString & String.Empty)
        Next
        PartnerNewCbo.Add("ALL")
        cboPartnerNew.DataSource = PartnerNewCbo
        cboPartnerNew.SelectedIndex = 0
        cboPartnerUpsert.DataSource = PartnerNewCbo
        cboPartnerUpsert.SelectedIndex = 0

        cboPartnerEDIInv.DataSource = PartnerNewCbo
        cboPartnerEDIInv.SelectedIndex = 0
    End Sub

    Private Sub init_grdECTECOM1_PARTNER()
        For Each grdCol As UltraGridColumn In grdECTPSTY2.DisplayLayout.Bands(0).Columns
            If grdCol.Header.Tag <> "" Then
                grdCol.Hidden = True
            End If
        Next
        For Each rowECTECOM1_PARTNER As DataRow In dst.Tables("ECTECOM1_PARTNER").Select()
            Dim ECOM_CODE As String = rowECTECOM1_PARTNER.Item("ECOM_CODE").ToString & String.Empty
            Dim STYLE_CODE = Absx1.txtFor("STYLE_CODE").Text
            Dim FILTER1 As String = String.Format("ECOM_CODE = '{0}' AND ECOM_STYLE_STATUS = 'A'", ECOM_CODE)
            Dim FILTER2 As String = String.Format("ECOM_CODE = '{0}' AND ECOM_STYLE_COLOR_STATUS = 'A'", ECOM_CODE)
            Dim Selected As Boolean = False
            Dim rowECTESTY1 As DataRow = dst.Tables.Item("ECTESTY1").Select(FILTER1).FirstOrDefault
            Dim ECOM_UNIT_PRICE As String = ""
            Dim ALT_UNIT_PCT As String = ""
            Dim ALT_UNIT_PRICE As String = ""
            Dim SHIP_ECOM As String = "0"
            Dim SHIP_DROP As String = "0"
            Dim SET_QTY As Int64 = 1
            Dim ECOM_MIN_QTY_DEFAULT As Int64 = Val(rowECTECOM1_PARTNER.Item("ECOM_MIN_QTY_DEFAULT").ToString & String.Empty)
            Dim rowECTSTYL1 As DataRow = dst.Tables.Item("ECTSTYL1").Select.FirstOrDefault
            'If Not IsNothing(rowECTSTYL1) Then
            '    SET_QTY = Val(rowECTSTYL1.Item("DEFAULT_SET_QTY").ToString & String.Empty)
            '    If SET_QTY < 0 Then
            '        SET_QTY = 1
            '    End If
            'End If
            If Not IsNothing(rowECTESTY1) Then
                SET_QTY = Val(rowECTSTYL1.Item("DEFAULT_SET_QTY").ToString & String.Empty)
                If SET_QTY < 0 Then
                    SET_QTY = 1
                End If
                If rowECTESTY1.Item("ECOM_MIN_QTY_OVERRIDE").ToString & String.Empty <> String.Empty Then
                    ECOM_MIN_QTY_DEFAULT = Val(rowECTESTY1.Item("ECOM_MIN_QTY_OVERRIDE").ToString & String.Empty)
                End If
                ECOM_UNIT_PRICE = rowECTESTY1.Item("ECOM_UNIT_PRICE").ToString & String.Empty
                ALT_UNIT_PCT = rowECTESTY1.Item("ALT_UNIT_PCT").ToString & String.Empty
                ALT_UNIT_PRICE = rowECTESTY1.Item("ALT_UNIT_PRICE").ToString & String.Empty
                SHIP_ECOM = rowECTESTY1.Item("SHIP_ECOM").ToString & String.Empty
                SHIP_DROP = rowECTESTY1.Item("SHIP_DROP").ToString & String.Empty
                SET_QTY = Val(rowECTESTY1.Item("SET_QTY").ToString & String.Empty)
            End If
            rowECTECOM1_PARTNER.Item("ECOM_MIN_QTY_DEFAULT") = Val(ECOM_MIN_QTY_DEFAULT)
            rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE") = Val(ECOM_UNIT_PRICE)
            rowECTECOM1_PARTNER.Item("ALT_UNIT_PCT") = Val(ALT_UNIT_PCT)
            rowECTECOM1_PARTNER.Item("ALT_UNIT_PRICE") = Val(ALT_UNIT_PRICE)
            rowECTECOM1_PARTNER.Item("SHIP_ECOM") = SHIP_ECOM
            rowECTECOM1_PARTNER.Item("SHIP_DROP") = SHIP_DROP
            rowECTECOM1_PARTNER.Item("SET_QTY") = SET_QTY
            If dst.Tables.Item("ECTESTY2").Select(FILTER2).Count > 0 Then
                rowECTECOM1_PARTNER.Item("SEL") = "1"
                Selected = True
            Else
                rowECTECOM1_PARTNER.Item("SEL") = "0"
                Selected = False
            End If

            For Each grdCol As UltraGridColumn In grdECTPSTY2.DisplayLayout.Bands(0).Columns
                If grdCol.Header.Tag <> "" Then
                    If grdCol.Header.Tag = ECOM_CODE And Selected Then
                        grdCol.Hidden = False
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub init_grdECTPSTY2()
        Dim PCNT As Integer = 0
        For Each rowECTECOM1_PARTNER As DataRow In dst.Tables("ECTECOM1_PARTNER").Select("", "ECOM_CODE")
            PCNT += 1
            Dim ECOM_CODE As String = rowECTECOM1_PARTNER.Item("ECOM_CODE").ToString & String.Empty
            Dim ECOM_NAME As String = rowECTECOM1_PARTNER.Item("ECOM_NAME").ToString & String.Empty
            With grdECTPSTY2.DisplayLayout.Bands(0)
                .Columns("ECOM_SEL_" & PCNT).Header.Caption = "On " & ECOM_NAME
                .Columns("ECOM_SEL_" & PCNT).Header.Tag = ECOM_CODE
                .Columns("ECOM_CODE_" & PCNT).Header.Caption = ECOM_CODE
                .Columns("ECOM_PCT_" & PCNT).Header.Caption = ECOM_NAME & " Pct"
                .Columns("ECOM_PCT_" & PCNT).Header.Tag = ECOM_CODE
                .Columns("ECOM_MIN_" & PCNT).Header.Caption = ECOM_NAME & " Min"
                .Columns("ECOM_MIN_" & PCNT).Header.Tag = ECOM_CODE
                .Columns("ECOM_PRICE_" & PCNT).Header.Caption = ECOM_NAME & " Price"
                .Columns("ECOM_PRICE_" & PCNT).Header.Tag = ECOM_CODE
                .Columns("ECOM_PARTNER_SKU_" & PCNT).Header.Caption = ECOM_NAME & " SKU"
                .Columns("ECOM_PARTNER_SKU_" & PCNT).Header.Tag = ECOM_CODE
                .Columns("ECOM_SEL_" & PCNT).Style = ColumnStyle.CheckBox
            End With
        Next
    End Sub

    Private Sub init_grdECTSTYB2()
        For Each rowECTSTYB2 As DataRow In dst.Tables("ECTSTYB2").Select()
            Dim fltr As String = String.Format("BULLET_CODE = '{0}'", rowECTSTYB2.Item("BULLET_CODE").ToString & String.Empty)
            If dst.Tables.Item("ECTSTYB1").Select(fltr).Count > 0 Then
                rowECTSTYB2.Item("SEL") = "1"
            Else
                rowECTSTYB2.Item("SEL") = "0"
            End If
        Next
    End Sub

    Private Sub init_Images()
        Dim rowECTESTY2 As DataRow = dst.Tables.Item("ECTESTY2").Select("", "COLOR_CODE").FirstOrDefault
        If Not IsNothing(rowECTESTY2) Then
            Dim COLOR_CODE As String = rowECTESTY2.Item("COLOR_CODE").ToString & String.Empty
            Dim IMG As String = EC_PARM_IMAGES_FOLDER & STYLE_CODE & "-" & COLOR_CODE & ".jpg"
            If IO.File.Exists(IMG) Then
                picSTYLECOLOR.ImageLocation = IMG
            End If
        End If
    End Sub

    Private Sub init_Misc()
        Dim rowECTESTY1 As DataRow = dst.Tables.Item("ECTESTY1").Select.FirstOrDefault

        Dim rowECTSTYL1 As DataRow = dst.Tables.Item("ECTSTYL1").Select.FirstOrDefault
        If Not IsNothing(rowECTSTYL1) Then
            txtPKG_CODE.Text = rowECTSTYL1.Item("PKG_CODE").ToString & String.Empty
        End If

        Dim rowECTPARM1 As DataRow = dst.Tables.Item("ECTPARM1").Select.FirstOrDefault
        If Not IsNothing(rowECTPARM1) Then
            EC_PARM_IMAGES_FOLDER = rowECTPARM1.Item("EC_PARM_IMAGES_FOLDER").ToString & String.Empty
            If EC_PARM_IMAGES_FOLDER.Length > 0 Then
                If Not EC_PARM_IMAGES_FOLDER.EndsWith("\") Then
                    EC_PARM_IMAGES_FOLDER = EC_PARM_IMAGES_FOLDER & "\"
                End If
            End If
        End If

        SQL.Length = 0
        SQL.AppendLine("SELECT IMAGE_DESC")
        SQL.AppendLine("FROM ECTIMAGT")
        SQL.AppendLine("WHERE NVL(IMAGE_DEFAULT,'0') = '1'")
        ASCMAIN1.sql = SQL.ToString()
        Dim IMAGE_DESC As String = ASCDATA1.GetDataValue
        Dim lstECTIMAGT As New List(Of String)
        lstECTIMAGT.Add(IMAGE_DESC)
        cboECTIMAGT.DataSource = lstECTIMAGT
        cboECTIMAGT.SelectedIndex = 0

    End Sub

    Private Function IsValidEcomStyle(ByVal STYLE_CODE As String) As Boolean
        Dim RetVal As Boolean = False
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT COUNT(*) AS REC_CNT")
        SQLS.AppendLine("FROM ECTESTY1")
        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Dim REC_CNT As Int16 = Val(ASCDATA1.GetDataValue)
        If REC_CNT > 0 Then
            RetVal = True
        End If
        Return RetVal
    End Function

    Private Sub refreshECTSALSX(Optional ByVal Overrideload As Boolean = False)
        Me.Cursor = Cursors.WaitCursor
        If Not FormLoading Then
            Dim SQLG As New StringBuilder With {.Length = 0}
            Dim SEL_LIST As New List(Of String)
            For Each rowECTECOM1_FILTER As DataRow In dst.Tables("ECTECOM1_FILTER").Select()
                If rowECTECOM1_FILTER.Item("SEL").ToString & String.Empty = "1" Then
                    If rowECTECOM1_FILTER.Item("ECOM_CODE").ToString & String.Empty <> "" Then
                        SEL_LIST.Add("'" & rowECTECOM1_FILTER.Item("ECOM_CODE").ToString & String.Empty & "',")
                    End If
                End If
            Next

            If SEL_LIST.Count = 0 Then
                MsgBox("You Must Select At Least One Partner.", vbOKOnly, "Selection")
            Else
                Dim START_PERIOD = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, txtDefaultMonths.Value * -1)
                Dim list As String = ""
                For Each l As String In SEL_LIST
                    list += l
                Next
                list = list.Substring(0, list.Length - 1)

                SQL.Length = 0
                SQL.AppendLine("SELECT")
                SQL.AppendLine("I1.INV_NO,")
                SQL.AppendLine("I1.ORDR_NO,")
                SQL.AppendLine("I1.CUST_CODE,")
                SQL.AppendLine("C1.CUST_NAME,")
                SQL.AppendLine("I1.ORDR_CUST_PO,")
                SQL.AppendLine("O1.EDI_PO_TYPE,")
                SQL.AppendLine("I1.INV_DATE,")
                SQL.AppendLine("I1.ORDR_YYYYPP_UPDATED,")
                SQL.AppendLine("I2.STYLE_CODE,")
                SQL.AppendLine("I2.COLOR_CODE,")
                SQL.AppendLine("(I2.STYLE_CODE || '-' || I2.COLOR_CODE) AS STYLE_COLOR,")
                SQL.AppendLine("I2.ORDR_UNIT_PRICE,")
                SQL.AppendLine("I2.ORDR_QTY_SHIP,")
                SQL.AppendLine("(I2.ORDR_UNIT_PRICE * I2.ORDR_QTY_SHIP) SHIP_TOTAL")
                SQL.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2, ARTCUST1 C1, SOTORDR1 O1,")
                SQL.AppendLine("(")
                SQL.AppendLine("SELECT")
                SQL.AppendLine("ECC.CUST_CODE, STYLE_CODE")
                SQL.AppendLine("FROM ECTECOM1 ECC, ECTESTY1 ECS")
                SQL.AppendLine("WHERE ECC.ECOM_CODE = ECS.ECOM_CODE")
                SQL.AppendLine("AND ECS.ECOM_STYLE_STATUS = 'A'")
                SQL.AppendLine(String.Format("AND ECS.ECOM_CODE IN ({0})", list))
                If chkSHIP_ECOM.Checked And chkSHIP_DROP.Checked Then
                    SQL.AppendLine("AND (NVL(SHIP_ECOM,'0') = '1' OR NVL(SHIP_DROP,'0') = '1')")
                Else
                    If chkSHIP_ECOM.Checked Then
                        SQL.AppendLine("AND NVL(SHIP_ECOM,'0') = '1'")
                    End If
                    If chkSHIP_DROP.Checked Then
                        SQL.AppendLine("AND NVL(SHIP_DROP,'0') = '1'")
                    End If
                End If
                SQL.AppendLine(") E1")
                SQL.AppendLine("WHERE I1.CUST_CODE = E1.CUST_CODE")
                SQL.AppendLine("AND I2.STYLE_CODE = E1.STYLE_CODE")
                SQL.AppendLine("AND I1.INV_NO = I2.INV_NO")
                SQL.AppendLine("AND I1.INV_TYPE = I2.INV_TYPE")
                SQL.AppendLine("AND I1.CUST_CODE = C1.CUST_CODE")
                SQL.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO")
                SQL.AppendLine("AND I1.INV_TYPE = 'I'")
                SQL.AppendLine("AND I1.WHSE_CODE IN ('MS','CG')")
                SQL.AppendLine("AND I1.ORDR_TYPE_CODE <> 'XFR'")
                SQL.AppendLine(String.Format("AND I1.ORDR_YYYYPP_UPDATED > '{0}'", START_PERIOD))
                dst.Tables.Item("ECTSALSX").Clear()
                Fill_Records("ECTSALSX",, True, SQL.ToString)
            End If

        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub refreshECTSTYLX(Optional ByVal Overrideload As Boolean = False)
        Me.Cursor = Cursors.WaitCursor
        If Not FormLoading Or Overrideload = True Then
            Dim SQLG As New StringBuilder With {.Length = 0}
            Dim SEL_LIST As New List(Of String)
            For Each rowECTECOM1_FILTER As DataRow In dst.Tables("ECTECOM1_FILTER").Select()
                If rowECTECOM1_FILTER.Item("SEL").ToString & String.Empty = "1" Then
                    SEL_LIST.Add("'" & rowECTECOM1_FILTER.Item("ECOM_CODE").ToString & String.Empty & "',")
                End If
            Next

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("ICTSTYL1.STYLE_CODE,")
            SQL.AppendLine("ECTESTY2.COLOR_CODE,")
            'SQL.AppendLine("ECTESTY2.ECOM_PARTNER_SKU,")
            SQL.AppendLine("ICTSTYL1.STYLE_STATUS,")
            SQL.AppendLine("ICTSTYL1.STYLE_DESC,")
            SQL.AppendLine("ICTSTYL1.STYLE_CLASS_CODE,")
            If SEL_LIST.Count = 1 Then
                SQL.AppendLine("NVL(ECTESTY1.SET_QTY,0) AS SET_QTY,")
                SQL.AppendLine("ECTESTY1.ECOM_CODE,")
                SQL.AppendLine("ECTESTY1.ECOM_STYLE_STATUS,")
                SQL.AppendLine("ECTESTY1.ECOM_UNIT_PRICE,")
                SQL.AppendLine("ECTESTY1.SHIP_ECOM,")
                SQL.AppendLine("ECTESTY1.SHIP_DROP,")
                SQL.AppendLine("ECTSTYL1.SHORT_DESC,")
                SQL.AppendLine("ECTSTYL1.LONG_DESC")

                SQLG.AppendLine("GROUP BY")
                SQLG.AppendLine("ICTSTYL1.STYLE_CODE,")
                SQLG.AppendLine("ECTESTY2.COLOR_CODE,")
                'SQLG.AppendLine("ECTESTY2.ECOM_PARTNER_SKU,")
                SQLG.AppendLine("ICTSTYL1.STYLE_STATUS,")
                SQLG.AppendLine("ICTSTYL1.STYLE_DESC,")
                SQLG.AppendLine("ICTSTYL1.STYLE_CLASS_CODE,")
                SQLG.AppendLine("NVL(ECTESTY1.SET_QTY,0),")
                SQLG.AppendLine("ECTESTY1.ECOM_CODE,")
                SQLG.AppendLine("ECTESTY1.ECOM_STYLE_STATUS,")
                SQLG.AppendLine("ECTESTY1.ECOM_UNIT_PRICE,")
                SQLG.AppendLine("ECTESTY1.SHIP_ECOM,")
                SQLG.AppendLine("ECTESTY1.SHIP_DROP,")
                SQLG.AppendLine("ECTSTYL1.SHORT_DESC,")
                SQLG.AppendLine("ECTSTYL1.LONG_DESC")
            Else
                SQL.AppendLine("0 AS SET_QTY,")
                SQL.AppendLine("'X' AS ECOM_CODE,")
                SQL.AppendLine("'X' AS ECOM_STYLE_STATUS,")
                SQL.AppendLine("0 AS ECOM_UNIT_PRICE,")
                SQL.AppendLine("0 AS SHIP_ECOM,")
                SQL.AppendLine("0 AS SHIP_DROP,")
                SQL.AppendLine("'X' AS SHORT_DESC,")
                SQL.AppendLine("'X' AS LONG_DESC")

                SQLG.AppendLine("GROUP BY")
                SQLG.AppendLine("ICTSTYL1.STYLE_CODE,")
                SQLG.AppendLine("ECTESTY2.COLOR_CODE,")
                'SQLG.AppendLine("ECTESTY2.ECOM_PARTNER_SKU,")
                SQLG.AppendLine("ICTSTYL1.STYLE_STATUS,")
                SQLG.AppendLine("ICTSTYL1.STYLE_DESC,")
                SQLG.AppendLine("ICTSTYL1.STYLE_CLASS_CODE")
            End If
            SQL.AppendLine("FROM ECTESTY1, ICTSTYL1, ECTSTYL1, ECTESTY2")
            SQL.AppendLine("WHERE ECTESTY1.STYLE_CODE = ICTSTYL1.STYLE_CODE")
            SQL.AppendLine("AND ECTESTY1.STYLE_CODE = ECTESTY2.STYLE_CODE")
            SQL.AppendLine("AND ECTESTY1.STYLE_CODE = ECTSTYL1.STYLE_CODE (+)")

            If SEL_LIST.Count > 0 Then
                Dim list As String = ""
                For Each l As String In SEL_LIST
                    list += l
                Next
                list = list.Substring(0, list.Length - 1)
                SQL.AppendLine("AND ECTESTY1.ECOM_CODE = ECTESTY2.ECOM_CODE")
                SQL.AppendLine("AND ECTESTY2.ECOM_CODE In (" & list & ")")
                If SEL_LIST.Count = 1 Then
                    If chkSHIP_ECOM.Checked And chkSHIP_DROP.Checked Then
                        SQL.AppendLine("AND (ECTESTY1.SHIP_DROP = '1' OR ECTESTY1.SHIP_ECOM = '1')")
                    Else
                        If chkSHIP_ECOM.Checked Then
                            SQL.AppendLine("AND ECTESTY1.SHIP_ECOM = '1'")
                        ElseIf chkSHIP_DROP.Checked Then
                            SQL.AppendLine("AND ECTESTY1.SHIP_DROP = '1'")
                        Else
                            SQL.AppendLine("AND (ECTESTY1.SHIP_DROP = 'X' OR ECTESTY1.SHIP_ECOM = 'X')")
                        End If
                    End If
                End If
            Else
                SQL.AppendLine("AND ECTESTY1.ECOM_CODE = 'NULL'")
            End If

            SQL.Append(SQLG)
            'SQL.AppendLine("ORDER BY ICTSTYL1.STYLE_CODE, ECTESTY2.COLOR_CODE, ECTESTY2.ECOM_PARTNER_SKU")
            SQL.AppendLine("ORDER BY ICTSTYL1.STYLE_CODE, ECTESTY2.COLOR_CODE")
            dst.Tables.Item("ECTSTYLX").Clear()
            EnforceConstraints(False)
            Fill_Records("ECTSTYLX",, True, SQL.ToString)
            EnforceConstraints(True)
            fillSTAT2Data()
            If SEL_LIST.Count = 1 Then
                SEL_ECOM_CODE = SEL_LIST(0).Substring(0, SEL_LIST(0).Length - 1)
                fillEDI846Data(SEL_ECOM_CODE)
                fillPartnerSKUs()
                showPartnerCols(True)
            Else
                SEL_ECOM_CODE = ""
                showPartnerCols(False)
            End If
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Function setCUST_STYLE_CODE(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal CUST_CODE As String, ByVal CUST_STYLE_CODE As String) As Boolean
        Dim RetVal As Boolean = False
        Dim SQL As New System.Text.StringBuilder With {.Length = 0}
        SQL.AppendLine("SELECT STYLE_CODE || '-' || COLOR_CODE")
        SQL.AppendLine("FROM SOTCSTY1")
        SQL.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
        SQL.AppendLine(String.Format("AND CUST_STYLE_CODE = '{0}'", CUST_STYLE_CODE))
        ASCMAIN1.sql = SQL.ToString()
        Dim USED_CODE As String = ASCDATA1.GetDataValue
        Dim UpdateCustStyle As Boolean = True
        Dim FILTER_SCE As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}' AND CUST_CODE = '{2}'", STYLE_CODE, COLOR_CODE, CUST_CODE)
        Dim rowSOTCSTY1 As DataRow = dst.Tables.Item("SOTCSTY1").Select(FILTER_SCE).FirstOrDefault
        If USED_CODE.Length > 0 Then
            UpdateCustStyle = False
            MsgBox(String.Format("Customer Style Code {0} Is Already Used On Style {1}", CUST_STYLE_CODE, USED_CODE), vbOKOnly, "No Dupes Allowed")
        Else
            If CUST_STYLE_CODE.Length = 0 Or CUST_CODE.Length = 0 Then
                UpdateCustStyle = False
                MsgBox("You Can Not Remove A Customer Style Code Using This Screen!", vbOKOnly, "No Removals Allowed")
            End If
        End If
        If UpdateCustStyle Then
            If ASCMAIN1.Logical_Lock("SOTCSTY1", CUST_CODE) Then
                If Not IsNothing(rowSOTCSTY1) Then
                    If rowSOTCSTY1.Item("CUST_STYLE_CODE").ToString & String.Empty <> CUST_STYLE_CODE & String.Empty Then
                        'rowSOTCSTY1.Item("CUST_STYLE_CODE") = CUST_STYLE_CODE
                        'This is part of key so it has to be updated directly and refreshed
                        'Boy does this code stink though!!!

                        If Not CUST_STYLE_CODE.Contains("DROP TABLE") Then
                            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                            SQLS.AppendLine("UPDATE SOTCSTY1")
                            SQLS.AppendLine(String.Format("SET CUST_STYLE_CODE = '{0}'", CUST_STYLE_CODE))
                            SQLS.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
                            SQLS.AppendLine(String.Format("AND CUST_STYLE_CODE = '{0}'", rowSOTCSTY1.Item("CUST_STYLE_CODE").ToString))
                            ASCMAIN1.sql = SQLS.ToString
                            ASCDATA1.ExecuteSQL()
                            Fill_Records("SOTCSTY1", STYLE_CODE)
                        End If
                    End If
                Else
                    Dim newSOTCSTY1 As DataRow = dst.Tables.Item("SOTCSTY1").NewRow
                    newSOTCSTY1.Item("CUST_CODE") = CUST_CODE
                    newSOTCSTY1.Item("CUST_STYLE_CODE") = CUST_STYLE_CODE
                    newSOTCSTY1.Item("STYLE_CODE") = STYLE_CODE
                    newSOTCSTY1.Item("COLOR_CODE") = COLOR_CODE
                    newSOTCSTY1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    newSOTCSTY1.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                    newSOTCSTY1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                    newSOTCSTY1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                    dst.Tables.Item("SOTCSTY1").Rows.Add(newSOTCSTY1)
                    Update_Record_TDA("SOTCSTY1")
                    Fill_Records("SOTCSTY1", STYLE_CODE)
                End If
            Else
                RetVal = False
            End If
        End If
        Return RetVal
    End Function

    Private Sub setBinding()
        Bind_Controls(pnlICTSTYL1, "ICTSTYL1")
        Bind_Controls(pnlECTSTYL1, "ECTSTYL1")
    End Sub

    Private Sub setGridDataSources()
        grdECTECOM1_FILTER.DataSource = dst.Tables("ECTECOM1_FILTER")
        grdECTECOM1_PARTNER.DataSource = dst.Tables("ECTECOM1_PARTNER")
        grdECTECOM2.DataSource = dst.Tables("ECTECOM2")
        grdECTECOM2_SEL.DataSource = dst.Tables("ECTECOM1_PARTNER")
        grdECTESTY3.DataSource = dst.Tables("ECTESTY3")
        grdECTESTY3_SEL.DataSource = dst.Tables("ECTECOM1_PARTNER")
        grdECTPSTY2.DataSource = dst.Tables("ECTPSTY2")
        grdICTSTYLD.DataSource = dst.Tables("ICTSTYLD")
        grdECTSTYLX.DataSource = dst.Tables("ECTSTYLX")
        grdECTSTYB2.DataSource = dst.Tables("ECTSTYB2")
        grdECTESTYX_NEW.DataSource = dst.Tables("ECTESTY1_NEW")
        grdECTSALS1.DataSource = dst.Tables("ECTSALS1")
        grdECTSALSX.DataSource = dst.Tables("ECTSALSX")
        grdECUPSERT.DataSource = dst.Tables("ECUPSERT")
        grdICTEDI01.DataSource = dst.Tables("ICTEDI01")
    End Sub

    Private Sub setGridDefaults()
        With grdECTECOM1_FILTER.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("SEL").CellActivation = UltraWinGrid.Activation.AllowEdit
        End With

        With grdECTECOM1_PARTNER.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            Dim EditCols As String() = {"SEL", "ECOM_UNIT_PRICE", "ALT_UNIT_PCT", "ALT_UNIT_PRICE", "SHIP_ECOM", "SHIP_DROP", "SET_QTY", "ECOM_MIN_QTY_DEFAULT"}
            For Each EditCol As String In EditCols
                .Columns(EditCol).CellActivation = UltraWinGrid.Activation.AllowEdit
            Next
            .Columns("ECOM_UNIT_PRICE").Format = "###,##0.00"
            .Columns("ALT_UNIT_PCT").Format = "###,##0.00"
            .Columns("ALT_UNIT_PRICE").Format = "###,##0.00"
            .Columns("SET_QTY").Format = "###,##0"
        End With

        With grdECTPSTY2.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False

            .Columns("COLOR_CODE").Header.Fixed = True
            .Columns("STYLE_COLOR_STATUS").Header.Fixed = True
            .Columns("UPC_CODE").Header.Fixed = True


            Dim EditCols As New Dictionary(Of String, String)
            For p As Integer = 1 To partnerCount
                EditCols.Add("ECOM_SEL_" & p, "")
                EditCols.Add("ECOM_PCT_" & p, "##0")
                EditCols.Add("ECOM_MIN_" & p, "##0")
                EditCols.Add("ECOM_PRICE_" & p, "##0.00")
                EditCols.Add("ECOM_PARTNER_SKU_" & p, "")
            Next

            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                If grdCol.Key.Contains("ECOM_PARTNER_SKU_") Then
                    grdCol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next

            For Each EditCol As KeyValuePair(Of String, String) In EditCols
                .Columns(EditCol.Key).CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns(EditCol.Key).Format = EditCol.Value
            Next

            For i As Integer = 1 To partnerCount
                For Each COLUMN_NAME As String In New String() {"ECOM_SEL_", "ECOM_CODE_", "ECOM_PCT_", "ECOM_MIN_", "ECOM_PRICE_", "ECOM_PARTNER_SKU_"}
                    With .Columns(COLUMN_NAME & i)
                        .Width = 110
                        Select Case i
                            Case 1
                                .Header.Appearance.BackColor2 = Drawing.Color.LightCyan
                            Case 2
                                .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                            Case 3
                                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
                            Case 4
                                .Header.Appearance.BackColor2 = Drawing.Color.LightSkyBlue
                            Case 5
                                .Header.Appearance.BackColor2 = Drawing.Color.LightYellow
                            Case Else
                                .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        End Select

                    End With
                Next
            Next
        End With

        With grdECTSTYB2.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            Dim EditCols As String() = {"SEL"}
            For Each EditCol As String In EditCols
                .Columns(EditCol).CellActivation = UltraWinGrid.Activation.AllowEdit
            Next
        End With

        With grdECTSTYLX.DisplayLayout.Bands("ECTSTYLX")
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
            .Columns("SET_QTY").Format = "###,##0"
            .Columns("EDI_REPORT_DATE").Format = "MM/dd/yy hh:mm"
            .Columns("EDI_AVAIL_QTY").Format = "###,##0"
            For Each COLUMN_NAME As String In New String() {"WHSE_QTY_ON_HAND", "WHSE_QTY_PICK", "OPEN_TO_SELL", "WHSE_QTY_TRAN", "WHSE_QTY_ON_ORDER", "WHSE_QTY_OPEN", "FUT_AVAIL"}
                .Columns(COLUMN_NAME).Format = "###,##0"
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightYellow
            Next
        End With

        With grdECTESTYX_NEW.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            Dim EditCols As String() = {"ECOM_STYLE_STATUS", "ECOM_UNIT_PRICE", "SET_QTY", "SHIP_ECOM", "SHIP_DROP"}
            For Each EditCol As String In EditCols
                .Columns(EditCol).CellActivation = UltraWinGrid.Activation.AllowEdit
            Next
        End With
        With grdECTESTYX_NEW.DisplayLayout.Bands(1)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            Dim EditCols As String() = {"ECOM_STYLE_COLOR_STATUS"}
            For Each EditCol As String In EditCols
                .Columns(EditCol).CellActivation = UltraWinGrid.Activation.AllowEdit
            Next
        End With

        With grdECTSALS1.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.False
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("ORDR_UNIT_PRICE").Format = "###,##0.00"
            .Columns("ORDR_QTY_SHIP").Format = "###,##0"
            .Columns("SHIP_TOTAL").Format = "###,##0.00"
            .Columns("INV_DATE").Format = "MM/dd/yy"
        End With

        With grdECTSALSX.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.False
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("ORDR_UNIT_PRICE").Format = "###,##0.00"
            .Columns("ORDR_QTY_SHIP").Format = "###,##0"
            .Columns("SHIP_TOTAL").Format = "###,##0.00"
            .Columns("INV_DATE").Format = "MM/dd/yy"
        End With

        With grdECUPSERT.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("ECOM_UNIT_PRICE").Format = "###,##0.00"
            .Columns("ALT_UNIT_PCT").Format = "###,##0.0"
            .Columns("ALT_UNIT_PRICE").Format = "###,##0.00"
            .Columns("SET_QTY").Format = "###,##0"
            .Columns("ECOM_MIN_QTY_OVERRIDE").Format = "###,##0"
            .Columns("PKG_L").Format = "###,##0"
            .Columns("PKG_WT").Format = "###,##0"
            .Columns("PKG_H").Format = "###,##0"
            .Columns("PKG_CUBE").Format = "###,##0"
            .Columns("UPSERT_TYPE").Hidden = True
            .Columns("UPSERT_ERROR").Hidden = True
        End With

        With grdICTEDI01.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.False
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("EDI_REPORT_DATE").Format = "MM/dd/yy"

            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True

            .Columns("WHSE_QTY_ON_HAND").Hidden = True
            .Columns("WHSE_QTY_PICK").Hidden = True
            .Columns("WHSE_QTY_TRAN").Hidden = True
            .Columns("WHSE_QTY_ON_ORDER").Hidden = True
            .Columns("WHSE_QTY_OPEN").Hidden = True
        End With
    End Sub

    Private Sub setGridSorts()
        Sort_grdColumns(grdECTSTYLX, "STYLE_CODE", False)
        Sort_grdColumns(grdECTECOM1_PARTNER, "ECOM_CODE", False)
        Sort_grdColumns(grdECTESTY3_SEL, "ECOM_CODE", False)
        Sort_grdColumns(grdECTESTY3, "COLOR_CODE", False)
        Sort_grdColumns(grdECTSTYB2, "BULLET_DESC", False)
        Sort_grdColumns(grdECTECOM2, "COLUMN_POS", False)
        Sort_grdColumns(grdECTSALS1, "INV_DATE", False)
        Sort_grdColumns(grdECTSALSX, "INV_DATE", False)
        Sort_grdColumns(grdECUPSERT, "STYLE_CODE", False)
    End Sub

    Private Sub setGridSummaries()
        Create_Summary(grdECTSTYLX, "STYLE_CODE", "Count")
        Create_Summary(grdECTSALS1, "ORDR_QTY_SHIP")
        Create_Summary(grdECTSALS1, "SHIP_TOTAL")
        Create_Summary(grdECTSALSX, "ORDR_QTY_SHIP")
        Create_Summary(grdECTSALSX, "SHIP_TOTAL")
    End Sub

    Private Sub setGridValueLists()
        ASCMAIN1.Add_Value_List(grdECTESTYX_NEW, "ECOM_STYLE_STATUS", , New String() {":", "A:Active", "X:Inactive"}, 0)
        ASCMAIN1.Add_Value_List(grdECTESTYX_NEW, "ECOM_STYLE_COLOR_STATUS", , New String() {":", "A:Active", "X:Inactive"}, 1)
        ASCMAIN1.Add_Value_List(grdECTSALS1, "EDI_PO_TYPE")
        ASCMAIN1.Add_Value_List(grdECTSALSX, "EDI_PO_TYPE")
        ASCMAIN1.Add_Value_List(grdECUPSERT, "ECOM_STYLE_STATUS", , New String() {":", "A:Active", "X:Inactive"}, 0)
        ASCMAIN1.Add_Value_List(grdECUPSERT, "ECOM_STYLE_COLOR_STATUS", , New String() {":", "A:Active", "X:Inactive"}, 0)
        ASCMAIN1.Add_Value_List(grdECUPSERT, "SHIP_ECOM", , New String() {":", "1:Yes", "0:No"}, 0)
        ASCMAIN1.Add_Value_List(grdECUPSERT, "SHIP_DROP", , New String() {":", "1:Yes", "0:No"}, 0)
        ASCMAIN1.Add_Value_List(grdECUPSERT, "UPSERT_TYPE", , New String() {":", "X:No Change", "U:Update", "A:Add", "E:Error"}, 0)
    End Sub

    Private Sub Setup_ECTESTY3()
        If Not IsNothing(grdECTESTY3_SEL.ActiveRow) Then
            Dim dvw As DataView = DirectCast(grdECTESTY3.DataSource, DataTable).DefaultView
            Dim ECOM_CODE As String = grdECTESTY3_SEL.ActiveRow.Cells("ECOM_CODE").Value & String.Empty
            Dim ECOM_NAME As String = grdECTESTY3_SEL.ActiveRow.Cells("ECOM_NAME").Value & String.Empty
            dvw.RowFilter = String.Format("ECOM_CODE = '{0}'", ECOM_CODE)
            grdECTESTY3.Text = String.Format("Promotions for {0}", ECOM_NAME)
        End If
    End Sub

    Private Sub showPartnerCols(ByVal showCols As Boolean)
        Dim partnerCols As String() = {"SET_QTY", "ECOM_CODE", "ECOM_STYLE_STATUS", "ECOM_UNIT_PRICE", "SHIP_ECOM", "SHIP_DROP", "SHORT_DESC",
            "LONG_DESC", "EDI_REPORT_DATE", "EDI_AVAIL_QTY", "EDI_STATUS", "ECOM_PARTNER_SKU"}
        For Each grdCol As UltraGridColumn In grdECTSTYLX.DisplayLayout.Bands(0).Columns
            If partnerCols.Contains(grdCol.Key) Then
                grdCol.Hidden = Not showCols
            End If
        Next
    End Sub

    Private Sub updateECTESTY2(ByVal STYLE_CODE As String,
                               ByVal COLOR_CODE As String,
                               ByVal ECOM_CODE As String,
                               ByVal ECOM_STYLE_STATUS As String,
                               ByVal ECOM_UNIT_PRICE As Double,
                               ByVal SET_QTY As Integer,
                               ByVal SHIP_ECOM As String,
                               ByVal SHIP_DROP As String)

        Dim FILTERST2 As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}' AND ECOM_CODE = '{2}'", STYLE_CODE, COLOR_CODE, ECOM_CODE)
        Dim rowECTESTY2 As DataRow = dst.Tables("ECTESTY2").Select(FILTERST2).FirstOrDefault
        If Not IsNothing(rowECTESTY2) Then
            rowECTESTY2.Item("ECOM_STYLE_COLOR_STATUS") = ECOM_STYLE_STATUS
        Else
            Dim newECTESTY2 As DataRow = dst.Tables("ECTESTY2").NewRow
            newECTESTY2.Item("STYLE_CODE") = STYLE_CODE
            newECTESTY2.Item("COLOR_CODE") = COLOR_CODE
            newECTESTY2.Item("ECOM_CODE") = ECOM_CODE
            newECTESTY2.Item("ECOM_STYLE_COLOR_STATUS") = ECOM_STYLE_STATUS
            dst.Tables("ECTESTY2").Rows.Add(newECTESTY2)
        End If
    End Sub

    Private Sub updateECTESTY1(ByVal STYLE_CODE As String,
                               ByVal ECOM_CODE As String,
                               ByVal ECOM_STYLE_STATUS As String,
                               ByVal ECOM_UNIT_PRICE As Double,
                               ByVal ALT_UNIT_PCT As Integer,
                               ByVal ALT_UNIT_PRICE As Double,
                               ByVal SET_QTY As Integer,
                               ByVal SHIP_ECOM As String,
                               ByVal SHIP_DROP As String,
                               ByVal ECOM_MIN_QTY_DEFAULT As Integer)

        Dim FILTERST1 As String = String.Format("STYLE_CODE = '{0}' AND ECOM_CODE = '{1}'", STYLE_CODE, ECOM_CODE)
        Dim rowECTESTY1 As DataRow = dst.Tables("ECTESTY1").Select(FILTERST1).FirstOrDefault

        If Not IsNothing(rowECTESTY1) Then
            If rowECTESTY1.Item("ECOM_STYLE_STATUS") & String.Empty <> ECOM_STYLE_STATUS Then
                addAuditRecord(STYLE_CODE, ECOM_CODE, "ECOM_STYLE_STATUS", rowECTESTY1.Item("ECOM_STYLE_STATUS") & String.Empty, ECOM_STYLE_STATUS)
                rowECTESTY1.Item("ECOM_STYLE_STATUS") = ECOM_STYLE_STATUS
            End If
            If rowECTESTY1.Item("ECOM_UNIT_PRICE") & String.Empty <> ECOM_UNIT_PRICE Then
                addAuditRecord(STYLE_CODE, ECOM_CODE, "ECOM_UNIT_PRICE", rowECTESTY1.Item("ECOM_UNIT_PRICE") & String.Empty, ECOM_UNIT_PRICE)
                rowECTESTY1.Item("ECOM_UNIT_PRICE") = ECOM_UNIT_PRICE
            End If
            If Val(rowECTESTY1.Item("ALT_UNIT_PCT") & String.Empty) <> ALT_UNIT_PCT Then
                addAuditRecord(STYLE_CODE, ECOM_CODE, "ALT_UNIT_PCT", rowECTESTY1.Item("ALT_UNIT_PCT") & String.Empty, ALT_UNIT_PCT)
                rowECTESTY1.Item("ALT_UNIT_PCT") = ALT_UNIT_PCT
            End If
            If Val(rowECTESTY1.Item("ALT_UNIT_PRICE") & String.Empty) <> ALT_UNIT_PRICE Then
                addAuditRecord(STYLE_CODE, ECOM_CODE, "ALT_UNIT_PRICE", rowECTESTY1.Item("ALT_UNIT_PRICE") & String.Empty, ALT_UNIT_PRICE)
                rowECTESTY1.Item("ALT_UNIT_PRICE") = ALT_UNIT_PRICE
            End If
            If rowECTESTY1.Item("SHIP_ECOM") & String.Empty <> SHIP_ECOM Then
                addAuditRecord(STYLE_CODE, ECOM_CODE, "SHIP_ECOM", rowECTESTY1.Item("SHIP_ECOM") & String.Empty, SHIP_ECOM)
                rowECTESTY1.Item("SHIP_ECOM") = SHIP_ECOM
            End If
            If rowECTESTY1.Item("SHIP_DROP") & String.Empty <> SHIP_DROP Then
                addAuditRecord(STYLE_CODE, ECOM_CODE, "SHIP_DROP", rowECTESTY1.Item("SHIP_DROP") & String.Empty, SHIP_DROP)
                rowECTESTY1.Item("SHIP_DROP") = SHIP_DROP
            End If
            If Val(rowECTESTY1.Item("SET_QTY") & String.Empty) <> SET_QTY Then
                addAuditRecord(STYLE_CODE, ECOM_CODE, "SET_QTY", rowECTESTY1.Item("SET_QTY") & String.Empty, SET_QTY)
                rowECTESTY1.Item("SET_QTY") = SET_QTY
            End If
            If Val(rowECTESTY1.Item("ECOM_MIN_QTY_OVERRIDE") & String.Empty) <> ECOM_MIN_QTY_DEFAULT Then
                addAuditRecord(STYLE_CODE, ECOM_CODE, "ECOM_MIN_QTY_OVERRIDE", rowECTESTY1.Item("ECOM_MIN_QTY_OVERRIDE") & String.Empty, ECOM_MIN_QTY_DEFAULT)
                rowECTESTY1.Item("ECOM_MIN_QTY_OVERRIDE") = ECOM_MIN_QTY_DEFAULT
            End If
        Else
            Dim newECTESTY1 As DataRow = dst.Tables("ECTESTY1").NewRow
            newECTESTY1.Item("STYLE_CODE") = STYLE_CODE
            newECTESTY1.Item("ECOM_CODE") = ECOM_CODE
            newECTESTY1.Item("ECOM_STYLE_STATUS") = ECOM_STYLE_STATUS
            newECTESTY1.Item("ECOM_UNIT_PRICE") = ECOM_UNIT_PRICE
            newECTESTY1.Item("ALT_UNIT_PCT") = ALT_UNIT_PCT
            newECTESTY1.Item("ALT_UNIT_PRICE") = ALT_UNIT_PRICE
            newECTESTY1.Item("SHIP_ECOM") = SHIP_ECOM
            newECTESTY1.Item("SHIP_DROP") = SHIP_DROP
            newECTESTY1.Item("SET_QTY") = SET_QTY
            newECTESTY1.Item("ECOM_MIN_QTY_OVERRIDE") = ECOM_MIN_QTY_DEFAULT
            dst.Tables("ECTESTY1").Rows.Add(newECTESTY1)
        End If
    End Sub

    Private Sub addAuditRecord(ByVal STYLE_CODE As String,
                               ByVal ECOM_CODE As String,
                               ByVal COL_NAME As String,
                               ByVal OLD_VALUE As String,
                               ByVal NEW_VALUE As String,
                               Optional ByVal TABLE_NAME As String = "ECTESTY1",
                               Optional ByVal ASTAUDT1 As String = "ASTAUDT1")
        Dim rowASTAUDT1 As DataRow = dst.Tables(ASTAUDT1).NewRow
        If NEW_VALUE.Length > 500 Then
            NEW_VALUE = NEW_VALUE.Substring(0, 499)
        End If
        If OLD_VALUE.Length > 500 Then
            OLD_VALUE = OLD_VALUE.Substring(0, 499)
        End If
        With rowASTAUDT1
            .Item("TABLE_NAME") = TABLE_NAME
            .Item("KEY_VALUE") = STYLE_CODE
            .Item("KEY_VALUE2") = ECOM_CODE
            .Item("COLUMN_NAME") = COL_NAME
            .Item("USER_ID") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
            .Item("OLD_VALUE") = OLD_VALUE
            .Item("NEW_VALUE") = NEW_VALUE
            .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
            .Item("SELECTION_NO") = Me.SELECTION_NO
            .Item("XNO") = Me.XNO
        End With
        dst.Tables(ASTAUDT1).Rows.Add(rowASTAUDT1)
    End Sub
#End Region

#Region "grdECTSTYLX"
    Private Sub grdECTSTYLX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdECTSTYLX.DoubleClickRow

        If e.Row.IsDataRow Then
            Absx1.txtFor("STYLE_CODE").Text = e.Row.Cells("STYLE_CODE").Text
            Click_Command("View")
        End If
    End Sub
#End Region

#Region "grdECTECOM1_FILTER"
    Private Sub grdECTECOM1_FILTER_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdECTECOM1_FILTER.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "SEL"
        End Select
    End Sub
#End Region

#Region "grdECTECOM1_PARTNER"
    Private Sub grdECTECOM1_PARTNER_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdECTECOM1_PARTNER.AfterCellUpdate
        Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text & String.Empty
        Dim ECOM_CODE As String = e.Cell.Row.Cells.Item("ECOM_CODE").Text & String.Empty
        Select Case e.Cell.Column.Key
            Case "ECOM_MIN_QTY_DEFAULT"
                Dim RowNum As Integer = e.Cell.Row.Index + 1
                Dim ECOM_MIN_QTY_DEFAULT As Double = Val(e.Cell.Value & String.Empty)
                Dim FILTERP1 As String = String.Format("ECOM_CODE = '{0}'", ECOM_CODE)
                Dim rowECTECOM1_PARTNER As DataRow = dst.Tables.Item("ECTECOM1_PARTNER").Select(FILTERP1).FirstOrDefault
                Dim ECOM_UNIT_PRICE As Double = Val(rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE").ToString & String.Empty)
                Dim SET_QTY As Integer = Val(rowECTECOM1_PARTNER.Item("SET_QTY").ToString & String.Empty)
                Dim SHIP_ECOM As String = rowECTECOM1_PARTNER.Item("SHIP_ECOM").ToString & String.Empty
                Dim SHIP_DROP As String = rowECTECOM1_PARTNER.Item("SHIP_DROP").ToString & String.Empty
                Dim ALT_UNIT_PCT As Double = Val(rowECTECOM1_PARTNER.Item("ALT_UNIT_PCT").ToString & String.Empty)
                Dim ALT_UNIT_PCT_OFF As Double = (100 - ALT_UNIT_PCT) / 100
                Dim ALT_UNIT_PRICE As Double = ECOM_UNIT_PRICE * ALT_UNIT_PCT_OFF
                Dim ECOM_STYLE_STATUS As String = "A"
                If rowECTECOM1_PARTNER.Item("SEL").ToString & String.Empty = "1" Then
                    ECOM_STYLE_STATUS = "A"
                Else
                    ECOM_STYLE_STATUS = "X"
                End If
                updateECTESTY1(STYLE_CODE, ECOM_CODE, ECOM_STYLE_STATUS, ECOM_UNIT_PRICE, ALT_UNIT_PCT, ALT_UNIT_PRICE, SET_QTY, SHIP_ECOM, SHIP_DROP, ECOM_MIN_QTY_DEFAULT)
                For Each rowECTPSTY2 As DataRow In dst.Tables("ECTPSTY2").Select()
                    rowECTPSTY2.Item("ECOM_MIN_" & RowNum) = Val(ECOM_MIN_QTY_DEFAULT)
                Next
                grdECTPSTY2.Update()
                grdECTPSTY2.Refresh()
            Case "ECOM_UNIT_PRICE"
                Dim ECOM_UNIT_PRICE As Double = Val(e.Cell.Value & String.Empty)
                Dim RowNum As Integer = e.Cell.Row.Index + 1
                'Dim FILTER1 As String = String.Format("STYLE_CODE = '{0}' AND ECOM_CODE = '{1}'", STYLE_CODE, ECOM_CODE)
                Dim FILTERP1 As String = String.Format("ECOM_CODE = '{0}'", ECOM_CODE)
                Dim rowECTECOM1_PARTNER As DataRow = dst.Tables.Item("ECTECOM1_PARTNER").Select(FILTERP1).FirstOrDefault
                Dim SET_QTY As Integer = Val(rowECTECOM1_PARTNER.Item("SET_QTY").ToString & String.Empty)
                Dim SHIP_ECOM As String = rowECTECOM1_PARTNER.Item("SHIP_ECOM").ToString & String.Empty
                Dim SHIP_DROP As String = rowECTECOM1_PARTNER.Item("SHIP_DROP").ToString & String.Empty
                Dim ALT_UNIT_PCT As Double = Val(rowECTECOM1_PARTNER.Item("ALT_UNIT_PCT").ToString & String.Empty)
                Dim ALT_UNIT_PRICE As Double = ECOM_UNIT_PRICE * ((100 - ALT_UNIT_PCT) / 100)
                Dim ECOM_MIN_QTY_DEFAULT As Double = Val(rowECTECOM1_PARTNER.Item("ECOM_MIN_QTY_DEFAULT").ToString & String.Empty)
                Dim ECOM_STYLE_STATUS As String = "A"
                If rowECTECOM1_PARTNER.Item("SEL").ToString & String.Empty = "1" Then
                    ECOM_STYLE_STATUS = "A"
                Else
                    ECOM_STYLE_STATUS = "X"
                End If
                If ALT_UNIT_PCT < 0 Or ALT_UNIT_PCT > 100 Then
                    ALT_UNIT_PCT = 0
                    ECOM_UNIT_PRICE = Val(rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE").ToString & String.Empty)
                    ALT_UNIT_PRICE = ECOM_UNIT_PRICE
                End If
                ALT_UNIT_PCT = ALT_UNIT_PCT * 100
                rowECTECOM1_PARTNER.Item("ALT_UNIT_PCT") = ALT_UNIT_PCT
                rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE") = ECOM_UNIT_PRICE
                rowECTECOM1_PARTNER.Item("ALT_UNIT_PRICE") = ALT_UNIT_PRICE
                updateECTESTY1(STYLE_CODE, ECOM_CODE, ECOM_STYLE_STATUS, ECOM_UNIT_PRICE, ALT_UNIT_PCT, ALT_UNIT_PRICE, SET_QTY, SHIP_ECOM, SHIP_DROP, ECOM_MIN_QTY_DEFAULT)

                If Not IsNothing(rowECTECOM1_PARTNER) Then
                    For Each rowECTESTY2_FROM As DataRow In dst.Tables("ECTESTY2").Select()
                        Dim COLOR_CODE As String = rowECTESTY2_FROM.Item("COLOR_CODE").ToString & String.Empty
                        Dim ECOM_STYLE_COLOR_STATUS As String = rowECTESTY2_FROM.Item("ECOM_STYLE_COLOR_STATUS").ToString & String.Empty
                        updateECTESTY2(STYLE_CODE, COLOR_CODE, ECOM_CODE, ECOM_STYLE_COLOR_STATUS, ECOM_UNIT_PRICE, SET_QTY, SHIP_ECOM, SHIP_DROP)
                    Next
                End If
                For Each rowECTPSTY2 As DataRow In dst.Tables("ECTPSTY2").Select()
                    rowECTPSTY2.Item("ECOM_PRICE_" & RowNum) = Val(ECOM_UNIT_PRICE)
                Next
                grdECTPSTY2.Update()
                grdECTPSTY2.Refresh()
            Case "ALT_UNIT_PCT"
                Dim FILTERP1 As String = String.Format("ECOM_CODE = '{0}'", ECOM_CODE)
                Dim rowECTECOM1_PARTNER As DataRow = dst.Tables.Item("ECTECOM1_PARTNER").Select(FILTERP1).FirstOrDefault
                Dim ECOM_UNIT_PRICE As Double = Val(rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE").ToString & String.Empty)
                Dim SET_QTY As Integer = Val(rowECTECOM1_PARTNER.Item("SET_QTY").ToString & String.Empty)
                Dim SHIP_ECOM As String = rowECTECOM1_PARTNER.Item("SHIP_ECOM").ToString & String.Empty
                Dim SHIP_DROP As String = rowECTECOM1_PARTNER.Item("SHIP_DROP").ToString & String.Empty
                Dim ALT_UNIT_PCT As Double = Val(rowECTECOM1_PARTNER.Item("ALT_UNIT_PCT").ToString & String.Empty)
                Dim ALT_UNIT_PCT_OFF As Double = (100 - ALT_UNIT_PCT) / 100
                Dim ALT_UNIT_PRICE As Double = ECOM_UNIT_PRICE * ALT_UNIT_PCT_OFF
                Dim ECOM_MIN_QTY_DEFAULT As Double = Val(rowECTECOM1_PARTNER.Item("ECOM_MIN_QTY_DEFAULT").ToString & String.Empty)
                Dim ECOM_STYLE_STATUS As String = "A"
                If ALT_UNIT_PCT < 0 Or ALT_UNIT_PCT > 100 Then
                    ALT_UNIT_PCT = 0
                    ECOM_UNIT_PRICE = Val(rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE").ToString & String.Empty)
                    ALT_UNIT_PRICE = ECOM_UNIT_PRICE
                End If
                rowECTECOM1_PARTNER.Item("ALT_UNIT_PCT") = ALT_UNIT_PCT
                rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE") = ECOM_UNIT_PRICE
                rowECTECOM1_PARTNER.Item("ALT_UNIT_PRICE") = ALT_UNIT_PRICE
                updateECTESTY1(STYLE_CODE, ECOM_CODE, ECOM_STYLE_STATUS, ECOM_UNIT_PRICE, ALT_UNIT_PCT, ALT_UNIT_PRICE, SET_QTY, SHIP_ECOM, SHIP_DROP, ECOM_MIN_QTY_DEFAULT)
            Case "ALT_UNIT_PRICE"
                Dim FILTERP1 As String = String.Format("ECOM_CODE = '{0}'", ECOM_CODE)
                Dim rowECTECOM1_PARTNER As DataRow = dst.Tables.Item("ECTECOM1_PARTNER").Select(FILTERP1).FirstOrDefault
                Dim ECOM_UNIT_PRICE As Double = Val(rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE").ToString & String.Empty)
                Dim SET_QTY As Integer = Val(rowECTECOM1_PARTNER.Item("SET_QTY").ToString & String.Empty)
                Dim SHIP_ECOM As String = rowECTECOM1_PARTNER.Item("SHIP_ECOM").ToString & String.Empty
                Dim SHIP_DROP As String = rowECTECOM1_PARTNER.Item("SHIP_DROP").ToString & String.Empty
                Dim ALT_UNIT_PRICE As Double = Val(rowECTECOM1_PARTNER.Item("ALT_UNIT_PRICE").ToString & String.Empty)
                Dim ECOM_MIN_QTY_DEFAULT As Double = Val(rowECTECOM1_PARTNER.Item("ECOM_MIN_QTY_DEFAULT").ToString & String.Empty)
                If ECOM_UNIT_PRICE = 0 And ALT_UNIT_PRICE > 0 Then
                    ECOM_UNIT_PRICE = ALT_UNIT_PRICE
                End If
                Dim ALT_UNIT_PCT As Double = 1 - (ALT_UNIT_PRICE / ECOM_UNIT_PRICE)
                If ALT_UNIT_PRICE = 0 And ECOM_UNIT_PRICE = 0 Then
                    ALT_UNIT_PCT = 0
                End If
                If ALT_UNIT_PCT < 0 Or ALT_UNIT_PCT > 100 Then
                    ALT_UNIT_PCT = 0
                    ECOM_UNIT_PRICE = Val(rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE").ToString & String.Empty)
                    ALT_UNIT_PRICE = ECOM_UNIT_PRICE
                End If
                ALT_UNIT_PCT = ALT_UNIT_PCT * 100
                rowECTECOM1_PARTNER.Item("ALT_UNIT_PCT") = ALT_UNIT_PCT
                rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE") = ECOM_UNIT_PRICE
                rowECTECOM1_PARTNER.Item("ALT_UNIT_PRICE") = ALT_UNIT_PRICE
                updateECTESTY1(STYLE_CODE, ECOM_CODE, "A", ECOM_UNIT_PRICE, ALT_UNIT_PCT, ALT_UNIT_PRICE, SET_QTY, SHIP_ECOM, SHIP_DROP, ECOM_MIN_QTY_DEFAULT)
            Case "SEL"
                If e.Cell.Text = "1" Then
                    Dim FILTERP1 As String = String.Format("ECOM_CODE = '{0}'", ECOM_CODE)
                    Dim rowECTECOM1_PARTNER As DataRow = dst.Tables.Item("ECTECOM1_PARTNER").Select(FILTERP1).FirstOrDefault
                    If Not IsNothing(rowECTECOM1_PARTNER) Then
                        Dim ECOM_UNIT_PRICE As Double = Val(rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE").ToString & String.Empty)
                        Dim ALT_UNIT_PCT As Integer = Val(rowECTECOM1_PARTNER.Item("ALT_UNIT_PCT").ToString & String.Empty)
                        Dim ALT_UNIT_PRICE As Double = Val(rowECTECOM1_PARTNER.Item("ALT_UNIT_PRICE").ToString & String.Empty)
                        Dim SET_QTY As Integer = Val(rowECTECOM1_PARTNER.Item("SET_QTY").ToString & String.Empty)
                        Dim SHIP_ECOM As String = rowECTECOM1_PARTNER.Item("SHIP_ECOM").ToString & String.Empty
                        Dim SHIP_DROP As String = rowECTECOM1_PARTNER.Item("SHIP_DROP").ToString & String.Empty
                        Dim ECOM_MIN_QTY_DEFAULT As Double = Val(rowECTECOM1_PARTNER.Item("ECOM_MIN_QTY_DEFAULT").ToString & String.Empty)
                        updateECTESTY1(STYLE_CODE, ECOM_CODE, "A", ECOM_UNIT_PRICE, ALT_UNIT_PCT, ALT_UNIT_PRICE, SET_QTY, SHIP_ECOM, SHIP_DROP, ECOM_MIN_QTY_DEFAULT)

                        Dim sql As New Text.StringBuilder With {.Length = 0}
                        sql.AppendLine("SELECT *")
                        sql.AppendLine("FROM ICTSTYC1")
                        sql.AppendLine("WHERE STYLE_CODE = :PARM1")
                        Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty, "V", STYLE_CODE)
                        For Each rowICTSTYC2 As DataRow In tbl.Rows
                            Dim COLOR_CODE As String = rowICTSTYC2.Item("COLOR_CODE").ToString & String.Empty
                            Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}' AND ECOM_CODE = '{2}'", STYLE_CODE, COLOR_CODE, ECOM_CODE)
                            If dst.Tables.Item("ECTESTY2").Select(Filter).Count = 0 Then
                                Dim newECTESTY2 As DataRow = dst.Tables.Item("ECTESTY2").NewRow
                                newECTESTY2.Item("STYLE_CODE") = STYLE_CODE
                                newECTESTY2.Item("COLOR_CODE") = COLOR_CODE
                                newECTESTY2.Item("ECOM_CODE") = ECOM_CODE
                                newECTESTY2.Item("ECOM_STYLE_COLOR_STATUS") = "A"
                                dst.Tables.Item("ECTESTY2").Rows.Add(newECTESTY2)
                            End If

                        Next

                        For Each rowECTESTY2_FROM As DataRow In dst.Tables("ECTESTY2").Select()
                            Dim COLOR_CODE As String = rowECTESTY2_FROM.Item("COLOR_CODE").ToString & String.Empty
                            updateECTESTY2(STYLE_CODE, COLOR_CODE, ECOM_CODE, "A", ECOM_UNIT_PRICE, SET_QTY, SHIP_ECOM, SHIP_DROP)
                        Next
                    End If
                    For Each grdCol As UltraGridColumn In grdECTPSTY2.DisplayLayout.Bands(0).Columns
                        If Not IsNothing(grdCol.Header.Tag) Then
                            If grdCol.Header.Tag = ECOM_CODE Then
                                grdCol.Hidden = False
                                Select Case grdCol.Key.Substring(0, grdCol.Key.Length - 1)
                                    Case "ECOM_PCT_"
                                        For Each rowECTPSTY2 As DataRow In dst.Tables("ECTPSTY2").Select()
                                            rowECTPSTY2.Item(grdCol.Key) = Val(rowECTECOM1_PARTNER.Item("ECOM_ALLOC_PCT_DEFAULT").ToString & String.Empty)
                                        Next
                                    Case "ECOM_MIN_"
                                        For Each rowECTPSTY2 As DataRow In dst.Tables("ECTPSTY2").Select()
                                            rowECTPSTY2.Item(grdCol.Key) = Val(rowECTECOM1_PARTNER.Item("ECOM_MIN_QTY_DEFAULT").ToString & String.Empty)
                                        Next
                                    Case "ECOM_PRICE_"
                                        For Each rowECTPSTY2 As DataRow In dst.Tables("ECTPSTY2").Select()
                                            rowECTPSTY2.Item(grdCol.Key) = Val(rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE").ToString & String.Empty)
                                        Next
                                End Select
                            End If
                        End If
                    Next
                    grdECTPSTY2.Update()
                    grdECTPSTY2.Refresh()
                Else
                    Dim FILTERP1 As String = String.Format("ECOM_CODE = '{0}'", ECOM_CODE)
                    Dim rowECTECOM1_PARTNER As DataRow = dst.Tables.Item("ECTECOM1_PARTNER").Select(FILTERP1).FirstOrDefault
                    If Not IsNothing(rowECTECOM1_PARTNER) Then
                        Dim FILTERST1 As String = String.Format("STYLE_CODE = '{0}' AND ECOM_CODE = '{1}'", STYLE_CODE, ECOM_CODE)
                        Dim rowECTESTY1 As DataRow = dst.Tables("ECTESTY1").Select(FILTERST1).FirstOrDefault
                        Dim SET_QTY As Integer = Val(rowECTECOM1_PARTNER.Item("SET_QTY").ToString & String.Empty)
                        Dim ECOM_MIN_QTY_DEFAULT As Double = Val(rowECTECOM1_PARTNER.Item("ECOM_MIN_QTY_DEFAULT").ToString & String.Empty)
                        updateECTESTY1(STYLE_CODE, ECOM_CODE, "X", 0, 0, 0, SET_QTY, "", "", ECOM_MIN_QTY_DEFAULT)
                        For Each rowECTESTY2_FROM As DataRow In dst.Tables("ECTESTY2").Select()
                            Dim COLOR_CODE As String = rowECTESTY2_FROM.Item("COLOR_CODE").ToString & String.Empty
                            Dim FILTERST2 As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}' AND ECOM_CODE = '{2}'", STYLE_CODE, COLOR_CODE, ECOM_CODE)
                            updateECTESTY2(STYLE_CODE, COLOR_CODE, ECOM_CODE, "X", 0, SET_QTY, "", "")
                        Next
                    End If
                    For Each grdCol As UltraGridColumn In grdECTPSTY2.DisplayLayout.Bands(0).Columns
                        If Not IsNothing(grdCol.Header.Tag) Then
                            If grdCol.Header.Tag = ECOM_CODE Then
                                grdCol.Hidden = True
                                Select Case grdCol.Key.Substring(0, grdCol.Key.Length - 1)
                                    Case "ECOM_PCT_"
                                        For Each rowECTPSTY2 As DataRow In dst.Tables("ECTPSTY2").Select()
                                            rowECTPSTY2.Item(grdCol.Key) = Null
                                        Next
                                    Case "ECOM_MIN_"
                                        For Each rowECTPSTY2 As DataRow In dst.Tables("ECTPSTY2").Select()
                                            rowECTPSTY2.Item(grdCol.Key) = Null
                                        Next
                                    Case "ECOM_PRICE_"
                                        For Each rowECTPSTY2 As DataRow In dst.Tables("ECTPSTY2").Select()
                                            rowECTPSTY2.Item(grdCol.Key) = Null
                                        Next
                                End Select
                            End If
                        End If
                    Next
                    grdECTPSTY2.Update()
                    grdECTPSTY2.Refresh()
                End If
            Case "SHIP_DROP", "SHIP_ECOM"
                Dim RowNum As Integer = e.Cell.Row.Index + 1
                Dim FILTERP1 As String = String.Format("ECOM_CODE = '{0}'", ECOM_CODE)
                Dim rowECTECOM1_PARTNER As DataRow = dst.Tables.Item("ECTECOM1_PARTNER").Select(FILTERP1).FirstOrDefault
                Dim ECOM_UNIT_PRICE As Double = Val(rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE").ToString & String.Empty)
                Dim ALT_UNIT_PCT As Integer = Val(rowECTECOM1_PARTNER.Item("ALT_UNIT_PCT").ToString & String.Empty)
                Dim ALT_UNIT_PRICE As Double = Val(rowECTECOM1_PARTNER.Item("ALT_UNIT_PRICE").ToString & String.Empty)
                Dim SET_QTY As Integer = Val(rowECTECOM1_PARTNER.Item("SET_QTY").ToString & String.Empty)
                Dim SHIP_ECOM As String = rowECTECOM1_PARTNER.Item("SHIP_ECOM").ToString & String.Empty
                Dim SHIP_DROP As String = rowECTECOM1_PARTNER.Item("SHIP_DROP").ToString & String.Empty
                Dim ECOM_MIN_QTY_DEFAULT As Double = Val(rowECTECOM1_PARTNER.Item("ECOM_MIN_QTY_DEFAULT").ToString & String.Empty)
                Dim ECOM_STYLE_STATUS As String = "A"
                If rowECTECOM1_PARTNER.Item("SEL").ToString & String.Empty = "1" Then
                    ECOM_STYLE_STATUS = "A"
                Else
                    ECOM_STYLE_STATUS = "X"
                End If
                updateECTESTY1(STYLE_CODE, ECOM_CODE, ECOM_STYLE_STATUS, ECOM_UNIT_PRICE, ALT_UNIT_PCT, ALT_UNIT_PRICE, SET_QTY, SHIP_ECOM, SHIP_DROP, ECOM_MIN_QTY_DEFAULT)
                If Not IsNothing(rowECTECOM1_PARTNER) Then
                    For Each rowECTESTY2_FROM As DataRow In dst.Tables("ECTESTY2").Select()
                        Dim COLOR_CODE As String = rowECTESTY2_FROM.Item("COLOR_CODE").ToString & String.Empty
                        Dim ECOM_STYLE_COLOR_STATUS As String = rowECTESTY2_FROM.Item("ECOM_STYLE_COLOR_STATUS").ToString & String.Empty
                        updateECTESTY2(STYLE_CODE, COLOR_CODE, ECOM_CODE, ECOM_STYLE_COLOR_STATUS, ECOM_UNIT_PRICE, SET_QTY, SHIP_ECOM, SHIP_DROP)
                    Next
                End If
                For Each rowECTPSTY2 As DataRow In dst.Tables("ECTPSTY2").Select()
                    rowECTPSTY2.Item("ECOM_PRICE_" & RowNum) = Val(ECOM_UNIT_PRICE)
                Next
                grdECTPSTY2.Update()
                grdECTPSTY2.Refresh()
            Case "SET_QTY"
                Dim RowNum As Integer = e.Cell.Row.Index + 1
                Dim FILTERP1 As String = String.Format("ECOM_CODE = '{0}'", ECOM_CODE)
                Dim rowECTECOM1_PARTNER As DataRow = dst.Tables.Item("ECTECOM1_PARTNER").Select(FILTERP1).FirstOrDefault
                Dim ECOM_UNIT_PRICE As Double = Val(rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE").ToString & String.Empty)
                Dim ALT_UNIT_PCT As Integer = Val(rowECTECOM1_PARTNER.Item("ALT_UNIT_PCT").ToString & String.Empty)
                Dim ALT_UNIT_PRICE As Double = Val(rowECTECOM1_PARTNER.Item("ALT_UNIT_PRICE").ToString & String.Empty)
                Dim SET_QTY As Integer = Val(rowECTECOM1_PARTNER.Item("SET_QTY").ToString & String.Empty)
                Dim SHIP_ECOM As String = rowECTECOM1_PARTNER.Item("SHIP_ECOM").ToString & String.Empty
                Dim SHIP_DROP As String = rowECTECOM1_PARTNER.Item("SHIP_DROP").ToString & String.Empty
                Dim ECOM_MIN_QTY_DEFAULT As Double = Val(rowECTECOM1_PARTNER.Item("ECOM_MIN_QTY_DEFAULT").ToString & String.Empty)
                Dim ECOM_STYLE_STATUS As String = "A"
                If rowECTECOM1_PARTNER.Item("SEL").ToString & String.Empty = "1" Then
                    ECOM_STYLE_STATUS = "A"
                Else
                    ECOM_STYLE_STATUS = "X"
                End If
                updateECTESTY1(STYLE_CODE, ECOM_CODE, ECOM_STYLE_STATUS, ECOM_UNIT_PRICE, ALT_UNIT_PCT, ALT_UNIT_PRICE, SET_QTY, SHIP_ECOM, SHIP_DROP, ECOM_MIN_QTY_DEFAULT)
        End Select
    End Sub

    Private Sub grdECTECOM1_PARTNER_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs) Handles grdECTECOM1_PARTNER.BeforeCellUpdate
        Select Case e.Cell.Column.Key
            Case "SEL"
                Dim ECOM_CODE As String = e.Cell.Row.Cells.Item("ECOM_CODE").Text & String.Empty
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Add/Remove Partner"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                If e.Cell.Text = "0" Then
                    iMSG.AppendLine(String.Format("This Will Remove This Style To {0}", ECOM_CODE))
                    iMSG.AppendLine("Is That What You Want?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult <> MsgBoxResult.Yes Then
                        e.Cancel = True
                    End If
                Else
                    iMSG.AppendLine(String.Format("This Will Add This Style To {0}", ECOM_CODE))
                    iMSG.AppendLine("Is That What You Want?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult <> MsgBoxResult.Yes Then
                        e.Cancel = True
                    End If
                End If
            Case "SET_QTY"
                If e.Cell.Text.Trim = "0" Or e.Cell.Text.Trim = "" Then
                    MsgBox("Set Qty Must Be 1 or Greater", MsgBoxStyle.OkOnly, "Set Qty")
                    e.Cancel = True
                End If
        End Select
    End Sub

    Private Sub grdECTECOM1_PARTNER_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdECTECOM1_PARTNER.ClickCell
        'Don't ask me why I have to do this.  Probably something in the stupid standards I am fighting against but I 
        'don't have the better part of an afternoon to waste figuring it out.
        Select Case e.Cell.Column.Key
            Case "ECOM_UNIT_PRICE"
                e.Cell.Activate()
                grdECTECOM1_PARTNER.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode)
            Case "ALT_UNIT_PCT"
                e.Cell.Activate()
                grdECTECOM1_PARTNER.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode)
            Case "ALT_UNIT_PRICE"
                e.Cell.Activate()
                grdECTECOM1_PARTNER.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode)
            Case "SET_QTY"
                e.Cell.Activate()
                grdECTECOM1_PARTNER.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode)
            Case "ECOM_MIN_QTY_DEFAULT"
                e.Cell.Activate()
                grdECTECOM1_PARTNER.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode)
            Case "SEL"
            Case Else

        End Select
    End Sub

    Private Sub grdECTECOM1_PARTNER_AfterEnterEditMode(sender As Object, e As EventArgs) Handles grdECTECOM1_PARTNER.AfterEnterEditMode
        Select Case grdECTECOM1_PARTNER.ActiveCell.Column.Key
            Case "ECOM_UNIT_PRICE"
                grdECTECOM1_PARTNER.ActiveCell.SelectAll()
            Case "ALT_UNIT_PCT"
                grdECTECOM1_PARTNER.ActiveCell.SelectAll()
            Case "ALT_UNIT_PRICE"
                grdECTECOM1_PARTNER.ActiveCell.SelectAll()
            Case "SET_QTY"
                grdECTECOM1_PARTNER.ActiveCell.SelectAll()
            Case "ECOM_MIN_QTY_DEFAULT"
                grdECTECOM1_PARTNER.ActiveCell.SelectAll()
        End Select
    End Sub

    Private Sub grdECTECOM1_PARTNER_BeforeCellActivate(sender As Object, e As CancelableCellEventArgs) Handles grdECTECOM1_PARTNER.BeforeCellActivate
        If e.Cell.Column.Key <> "SEL" Then
            If e.Cell.Row.Cells.Item("SEL").Value = "0" Then
                e.Cancel = True
            End If
        End If
    End Sub
#End Region

#Region "grdECTESTY3_SEL"
    Private Sub grdECTESTY3_SEL_AfterRowActivate(sender As Object, e As EventArgs) Handles grdECTESTY3_SEL.AfterRowActivate
        Setup_ECTESTY3()
    End Sub
#End Region

#Region "grdECTPSTY2"
    Private Sub grdECTPSTY2_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdECTPSTY2.AfterCellUpdate
        Select Case e.Cell.Column.Key.Substring(0, e.Cell.Column.Key.Length - 1)
            Case "ECOM_SEL_"
                Dim STYLE_CODE As String = e.Cell.Row.Cells.Item("STYLE_CODE").Text.ToString
                Dim COLOR_CODE As String = e.Cell.Row.Cells.Item("COLOR_CODE").Text.ToString
                Dim ECOM_CODE As String = e.Cell.Column.Header.Tag

                Dim FILTERP1 As String = String.Format("ECOM_CODE = '{0}'", ECOM_CODE)
                Dim rowECTECOM1_PARTNER As DataRow = dst.Tables.Item("ECTECOM1_PARTNER").Select(FILTERP1).FirstOrDefault
                If Not IsNothing(rowECTECOM1_PARTNER) Then
                    Dim SET_QTY As Integer = Val(rowECTECOM1_PARTNER.Item("SET_QTY").ToString & String.Empty)
                    Dim SHIP_ECOM As String = rowECTECOM1_PARTNER.Item("SHIP_ECOM").ToString & String.Empty
                    Dim SHIP_DROP As String = rowECTECOM1_PARTNER.Item("SHIP_DROP").ToString & String.Empty
                    Dim ECOM_UNIT_PRICE As Double = Val(rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE").ToString & String.Empty)
                    Dim ECOM_STYLE_STATUS As String = "A"
                    If e.Cell.Value <> "1" Then
                        ECOM_STYLE_STATUS = "X"
                    End If
                    updateECTESTY2(STYLE_CODE, COLOR_CODE, ECOM_CODE, ECOM_STYLE_STATUS, ECOM_UNIT_PRICE, SET_QTY, SHIP_ECOM, SHIP_DROP)
                    'Put something here to mark the header records for the last color out the door.
                End If
            Case "ECOM_PARTNER_SKU_"
                Dim STYLE_CODE As String = e.Cell.Row.Cells.Item("STYLE_CODE").Text.ToString
                Dim COLOR_CODE As String = e.Cell.Row.Cells.Item("COLOR_CODE").Text.ToString
                Dim ECOM_CODE As String = e.Cell.Column.Header.Tag
                Dim FILTERP1 As String = String.Format("ECOM_CODE = '{0}'", ECOM_CODE)
                Dim rowECTECOM1_PARTNER As DataRow = dst.Tables.Item("ECTECOM1_PARTNER").Select(FILTERP1).FirstOrDefault
                Dim CUST_CODE As String = rowECTECOM1_PARTNER.Item("CUST_CODE").ToString
                If CUST_CODE.Length > 0 Then
                    If ECTPSTY2_INIT = False Then
                        Dim Retval As Boolean = setCUST_STYLE_CODE(STYLE_CODE, COLOR_CODE, CUST_CODE, e.Cell.Text)
                    End If
                End If
        End Select
    End Sub

    Private Sub grdECTPSTY2_AfterEnterEditMode(sender As Object, e As EventArgs) Handles grdECTPSTY2.AfterEnterEditMode
        If Not colorsLocked Then
            Select Case grdECTPSTY2.ActiveCell.Column.Key.Substring(0, grdECTPSTY2.ActiveCell.Column.Key.Length - 1)
                Case "ECOM_PCT_", "ECOM_MIN_", "ECOM_PRICE_", "ECOM_PARTNER_SKU_"
                    grdECTPSTY2.ActiveCell.SelectAll()
            End Select
        End If
    End Sub

    Private Sub grdECTPSTY2_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdECTPSTY2.ClickCell
        If e.Cell.Column.Key.Substring(0, e.Cell.Column.Key.Length - 1) = "ECOM_SEL_" Then
            e.Cell.Activate()
            grdECTPSTY2.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode)
        Else
            If Not colorsLocked Then
                Select Case e.Cell.Column.Key.Substring(0, e.Cell.Column.Key.Length - 1)
                    Case "ECOM_PCT_", "ECOM_MIN_", "ECOM_PRICE_", "ECOM_PARTNER_SKU_"
                        e.Cell.Activate()
                        grdECTPSTY2.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode)
                End Select
            Else
                Select Case e.Cell.Column.Key.Substring(0, e.Cell.Column.Key.Length - 1)
                    Case "ECOM_PARTNER_SKU_"
                        e.Cell.Activate()
                        grdECTPSTY2.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode)
                End Select
            End If
        End If
    End Sub

    Private Sub grdECTPSTY2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdECTPSTY2.InitializeRow
        ECTPSTY2_INIT = True
        For Each grdCell As UltraGridCell In e.Row.Cells
            If Not IsNothing(grdCell.Column.Header.Tag) Then
                Dim STYLE_CODE As String = e.Row.Cells.Item("STYLE_CODE").Text.ToString & String.Empty
                Dim COLOR_CODE As String = e.Row.Cells.Item("COLOR_CODE").Text.ToString & String.Empty
                Dim ECOM_CODE As String = grdCell.Column.Header.Tag
                Dim FILTER_SCE As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}' AND ECOM_CODE = '{2}'", STYLE_CODE, COLOR_CODE, ECOM_CODE)

                Dim rowECTECOM1_PARTNER As DataRow = dst.Tables.Item("ECTECOM1_PARTNER").Select(String.Format("ECOM_CODE = '{0}'", ECOM_CODE)).FirstOrDefault
                If Not IsNothing(rowECTECOM1_PARTNER) Then
                    If rowECTECOM1_PARTNER.Item("SEL").ToString & String.Empty = "1" Then
                        Dim ColName As String = grdCell.Column.Key.Substring(0, grdCell.Column.Key.Length - 1)
                        Select Case ColName
                            Case "ECOM_SEL_"
                                Dim rowECTESTY2 As DataRow = dst.Tables.Item("ECTESTY2").Select(FILTER_SCE).FirstOrDefault
                                If Not IsNothing(rowECTESTY2) Then
                                    If rowECTESTY2.Item("ECOM_STYLE_COLOR_STATUS").ToString & String.Empty = "A" Then
                                        grdCell.Value = "1"
                                    Else
                                        grdCell.Value = "0"
                                    End If
                                End If
                            Case "ECOM_PCT_"
                                If grdCell.Text & String.Empty = String.Empty Then
                                    grdCell.Value = Val(rowECTECOM1_PARTNER.Item("ECOM_ALLOC_PCT_DEFAULT").ToString)
                                End If
                            Case "ECOM_MIN_"
                                If grdCell.Text & String.Empty = String.Empty Then
                                    grdCell.Value = Val(rowECTECOM1_PARTNER.Item("ECOM_MIN_QTY_DEFAULT").ToString)
                                End If
                            Case "ECOM_PRICE_"
                                If grdCell.Text & String.Empty = String.Empty Then
                                    grdCell.Value = Val(rowECTECOM1_PARTNER.Item("ECOM_UNIT_PRICE").ToString)
                                End If
                            Case "ECOM_PARTNER_SKU_"
                                If ECOM_CODE.Length > 0 Then
                                    Dim CUST_CODE As String = rowECTECOM1_PARTNER.Item("CUST_CODE").ToString & String.Empty
                                    If CUST_CODE.Length > 0 Then
                                        grdCell.Value = getCUST_STYLE_CODE(STYLE_CODE, COLOR_CODE, CUST_CODE)
                                    End If
                                End If
                        End Select
                    End If
                End If
            End If
        Next
        ECTPSTY2_INIT = False

    End Sub
#End Region

#Region "grdECTESTYX_NEW"
    Private Sub grdECTESTYX_NEW_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdECTESTYX_NEW.ClickCell
        'Don't ask me why I have to do this.  Probably something in the stupid standards I am fighting against but I 
        'don't have the better part of an afternoon to waste figuring it out.
        Select Case e.Cell.Column.Key
            Case "ECOM_UNIT_PRICE", "SET_QTY"
                e.Cell.Activate()
                grdECTESTYX_NEW.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode)
            Case "SEL"
            Case Else

        End Select
    End Sub

    Private Sub grdECTESTYX_NEW_AfterEnterEditMode(sender As Object, e As EventArgs) Handles grdECTESTYX_NEW.AfterEnterEditMode
        Select Case grdECTESTYX_NEW.ActiveCell.Column.Key
            Case "ECOM_UNIT_PRICE", "SET_QTY"
                grdECTESTYX_NEW.ActiveCell.SelectAll()
            Case Else
        End Select
    End Sub
#End Region

#Region "grdECTSTYB2"
    Private Sub grdECTSTYB2_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdECTSTYB2.AfterCellUpdate
        If e.Cell.Column.Key = "SEL" Then
            Dim SEL As String = e.Cell.Value
            Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text & String.Empty
            Dim BULLET_CODE As String = e.Cell.Row.Cells.Item("BULLET_CODE").Text & String.Empty
            Dim fltr As String = String.Format("STYLE_CODE = '{0}' AND BULLET_CODE = '{1}'", STYLE_CODE, BULLET_CODE)
            Dim rowECTSTYB1 As DataRow = dst.Tables.Item("ECTSTYB1").Select(fltr).FirstOrDefault
            If Not IsNothing(rowECTSTYB1) Then
                If SEL <> "1" Then
                    rowECTSTYB1.Delete()
                End If
            Else
                Dim newECTSTYB1 As DataRow = dst.Tables.Item("ECTSTYB1").NewRow
                newECTSTYB1.Item("STYLE_CODE") = STYLE_CODE
                newECTSTYB1.Item("BULLET_CODE") = BULLET_CODE
                dst.Tables.Item("ECTSTYB1").Rows.Add(newECTSTYB1)
            End If

        End If
    End Sub

    Private Sub grdECTECOM1_PARTNER_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdECTECOM1_PARTNER.InitializeLayout

    End Sub

    Private Sub txtDEFAULT_SET_QTY_ValueChanged(sender As Object, e As EventArgs) Handles txtDEFAULT_SET_QTY.ValueChanged

    End Sub

    Private Sub txtDEFAULT_SET_QTY_LostFocus(sender As Object, e As EventArgs) Handles txtDEFAULT_SET_QTY.LostFocus
        If Not RecordLoading Then
            If EntryMode = "V" Then
                If DEFAULT_SET_QTY_ORIG <> txtDEFAULT_SET_QTY.Value Then
                    If Val(txtDEFAULT_SET_QTY.Value & String.Empty) = 0 Then
                        MsgBox("Default Set Qty Must be 1 or Greater", vbOKOnly, "Default Set Qty")
                        txtDEFAULT_SET_QTY.Value = DEFAULT_SET_QTY_ORIG
                    Else
                        Dim iResult As MsgBoxResult
                        Dim iTitle As String = "Default Set Qty"
                        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                        iMSG.AppendLine("Do You Want To Update All Of The PArtners")
                        iMSG.AppendLine(String.Format("With A New Set Qty of {0} ?", txtDEFAULT_SET_QTY.Value))
                        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                        If iResult = MsgBoxResult.Yes Then
                            'ECTECOM1_PARTNER
                            For Each grow As UltraWinGrid.UltraGridRow In grdECTECOM1_PARTNER.Rows
                                grow.Cells.Item("SET_QTY").Value = txtDEFAULT_SET_QTY.Value
                            Next
                            grdECTECOM1_PARTNER.UpdateData()
                            grdECTECOM1_PARTNER.Refresh()
                        End If
                        DEFAULT_SET_QTY_ORIG = txtDEFAULT_SET_QTY.Value
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub btnLoadData_Click(sender As Object, e As EventArgs) Handles btnLoadData.Click
        Dim tableData As New DataTable
        Dim fieldName As String = String.Empty
        Dim str As New StringBuilder With {.Length = 0}
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}

        'Dim ECOM_CODE As String = cboPartnerUpsert.Text

        str.AppendLine("This Will Allow You To Upload")
        str.AppendLine("A File To Add New Or Update Existing")
        str.AppendLine("Styles.")
        str.AppendLine("")
        str.AppendLine("It Should Be In The Same Format As If")
        str.AppendLine("You Has Exported The Grid Below.")
        str.AppendLine("")
        str.AppendLine("Are You Ready?")
        Dim iResult As MsgBoxResult = MsgBox(str.ToString, vbYesNo, "Upsert Styles?")
        Dim fileToImport As String = String.Empty
        If iResult = MsgBoxResult.Yes Then
            dst.Tables.Item("ECUPSERT").Clear()

            Using openFileDialog1 As New OpenFileDialog
                openFileDialog1.Title = "Open File To Upsert"
                openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx"
                openFileDialog1.FilterIndex = 1
                openFileDialog1.RestoreDirectory = True

                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                    fileToImport = openFileDialog1.FileName
                End If

                openFileDialog1.Dispose()
            End Using
            If fileToImport.Length = 0 Then
                Exit Sub
            End If

            ASCMAIN1.Progress("Reading File")
            Me.Cursor = Cursors.WaitCursor

            Using cn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileToImport & ";Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""")
                Using cmd As New System.Data.OleDb.OleDbDataAdapter("select * from [Dont Change The Column Names$]", cn)
                    ' Select the data from Sheet1 of the workbook.
                    cn.Open()
                    cmd.Fill(tableData)
                    cn.Close()
                    cmd.Dispose()
                End Using
                cn.Dispose()
            End Using

            Dim rowIndex As Int64 = 0
            Dim rowDataStart As Int64 = 0
            If tableData.Columns(0).ColumnName = "Partner Code" Then
                'Stop
                For i As Integer = 0 To tableData.Columns.Count - 1
                    Dim COL_DESC As String = tableData.Columns(i).ColumnName & String.Empty
                    Dim findFilter As String = String.Format("COL_DESC = '{0}'", COL_DESC)
                    Dim rowECTUPSRT As DataRow = dst.Tables.Item("ECTUPSRT").Select(findFilter).FirstOrDefault
                    If Not IsNothing(rowECTUPSRT) Then
                        rowECTUPSRT.Item("COL_INDEX") = i
                        tableData.Columns(i).ColumnName = rowECTUPSRT.Item("COL_NAME").ToString & String.Empty
                    End If
                Next
            Else
                For Each rowData As DataRow In tableData.Select()
                    If rowIndex > 15 Then
                        MsgBox("Can't Find The First Row!", vbExclamation, "Invalid Spreadsheet!!")
                        Exit Sub
                    End If
                    If rowData.Item(0).ToString & String.Empty = "Partner Code" Then
                        rowDataStart = rowIndex
                        For i As Integer = 0 To tableData.Columns.Count - 1
                            Dim COL_DESC As String = rowData.Item(i).ToString & String.Empty
                            Dim findFilter As String = String.Format("COL_DESC = '{0}'", COL_DESC)
                            Dim rowECTUPSRT As DataRow = dst.Tables.Item("ECTUPSRT").Select(findFilter).FirstOrDefault
                            If Not IsNothing(rowECTUPSRT) Then
                                rowECTUPSRT.Item("COL_INDEX") = i
                                tableData.Columns(i).ColumnName = rowECTUPSRT.Item("COL_NAME").ToString & String.Empty
                            End If
                        Next
                        Exit For
                    End If
                    rowIndex += 1
                Next
            End If


            rowIndex = 0
            Dim rowsBlank As Integer = 0
            Dim ECOM_MIN_QTY_DEFAULT As Int64 = 4
            For Each rowData As DataRow In tableData.Select()
                If rowIndex >= rowDataStart + 1 And rowsBlank <= 10 Then
                    If Not PartnerNewCbo.Contains(rowData.Item("ECOM_CODE").ToString.Trim & String.Empty) Then
                        rowsBlank += 1
                    Else
                        Dim rowErr As String = ""
                        Dim STYLE_CODE As String = rowData.Item("STYLE_CODE").ToString.Trim & String.Empty
                        Dim COLOR_CODE As String = rowData.Item("COLOR_CODE").ToString.Trim & String.Empty
                        Dim ECOM_CODE As String = (rowData.Item("ECOM_CODE").ToString.Trim & String.Empty).ToUpper
                        Dim rowECUPSERT As DataRow = dst.Tables.Item("ECUPSERT").NewRow
                        Dim rowECTESTY1 As DataRow = LookUp("ECTESTY1", New String() {STYLE_CODE, ECOM_CODE})
                        Dim UPSERT_TYPE As String = ""
                        If IsNothing(rowECTESTY1) Then
                            UPSERT_TYPE = "A"
                        Else
                            UPSERT_TYPE = "U"
                        End If
                        For Each rowECTUPSRT As DataRow In dst.Tables("ECTUPSRT").Select()
                            If Val(rowECTUPSRT.Item("COL_INDEX")) >= 0 Then
                                Select Case rowECTUPSRT.Item("COL_NAME").ToString & String.Empty
                                    Case "ECOM_CODE"
                                        Dim rowECTECOM1 As DataRow = LookUp("ECTECOM1", ECOM_CODE)
                                        If IsNothing(rowECTECOM1) Then
                                            setRowError(rowErr, String.Format("'{0}' Is Not A Valid E-COM Partner Code.", ECOM_CODE))
                                        Else
                                            ECOM_MIN_QTY_DEFAULT = Val(rowECTECOM1.Item("ECOM_MIN_QTY_DEFAULT").ToString & String.Empty)
                                        End If
                                        rowECUPSERT.Item("ECOM_CODE") = ECOM_CODE
                                    Case "STYLE_CODE"
                                        Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
                                        If IsNothing(rowICTSTYC1) Then
                                            setRowError(rowErr, String.Format("'{0}-{1}' Is Not A Valid Style/Color.", STYLE_CODE, COLOR_CODE))
                                        End If
                                        rowECUPSERT.Item("STYLE_CODE") = STYLE_CODE
                                        rowECUPSERT.Item("COLOR_CODE") = COLOR_CODE
                                    Case "COLOR_CODE"
                                        'Validating Both Not in Style Code
                                    Case "STYLE_DESC" 'Does Not Update.
                                        SQLS.Length = 0
                                        SQLS.AppendLine("SELECT STYLE_DESC")
                                        SQLS.AppendLine("FROM ICTSTYL1")
                                        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                                        ASCMAIN1.sql = SQLS.ToString()
                                        rowECUPSERT.Item("STYLE_DESC") = ASCDATA1.GetDataValue
                                    Case "COLOR_DESC" 'Does Not Update.
                                        SQLS.Length = 0
                                        SQLS.AppendLine("SELECT COLOR_DESC")
                                        SQLS.AppendLine("FROM ICTCOLR1")
                                        SQLS.AppendLine(String.Format("WHERE COLOR_CODE = '{0}'", COLOR_CODE))
                                        ASCMAIN1.sql = SQLS.ToString()
                                        rowECUPSERT.Item("COLOR_DESC") = ASCDATA1.GetDataValue
                                    Case "ECOM_STYLE_STATUS"
                                        Dim ECOM_STYLE_STATUS As String = rowData.Item("ECOM_STYLE_STATUS").ToString & String.Empty
                                        If ECOM_STYLE_STATUS.Length > 0 Then
                                            Select Case ECOM_STYLE_STATUS.ToUpper
                                                Case "ACTIVE"
                                                    rowECUPSERT.Item("ECOM_STYLE_STATUS") = "A"
                                                Case "INACTIVE"
                                                    rowECUPSERT.Item("ECOM_STYLE_STATUS") = "X"
                                                Case Else
                                                    setRowError(rowErr, String.Format("'{0}' Is Not A Valid E-COM Style Status.  Must Be Active or Inactive", ECOM_STYLE_STATUS))
                                            End Select
                                        End If
                                    Case "ECOM_STYLE_COLOR_STATUS"
                                        Dim ECOM_STYLE_COLOR_STATUS As String = rowData.Item("ECOM_STYLE_COLOR_STATUS").ToString & String.Empty
                                        If ECOM_STYLE_COLOR_STATUS.Length > 0 Then
                                            Select Case ECOM_STYLE_COLOR_STATUS.ToUpper
                                                Case "ACTIVE"
                                                    rowECUPSERT.Item("ECOM_STYLE_COLOR_STATUS") = "A"
                                                Case "INACTIVE"
                                                    rowECUPSERT.Item("ECOM_STYLE_COLOR_STATUS") = "X"
                                                Case Else
                                                    setRowError(rowErr, String.Format("'{0}' Is Not A Valid E-COM Style Color Status.  Must Be Active or Inactive", ECOM_STYLE_COLOR_STATUS))
                                            End Select
                                        End If
                                    Case "ECOM_UNIT_PRICE"
                                        Dim ECOM_UNIT_PRICE As String = rowData.Item("ECOM_UNIT_PRICE").ToString & String.Empty
                                        If ECOM_UNIT_PRICE.Length > 0 Then
                                            If Not IsNumeric(ECOM_UNIT_PRICE) Then
                                                setRowError(rowErr, String.Format("'{0}' Is Not A Valid Unit Price. Must Be Numeric.", ECOM_UNIT_PRICE))
                                                rowECUPSERT.Item("ECOM_UNIT_PRICE") = ECOM_UNIT_PRICE
                                            Else
                                                If Val(CDbl(ECOM_UNIT_PRICE)) > 10000 Or Val(CDbl(ECOM_UNIT_PRICE)) < 0 Then
                                                    setRowError(rowErr, "Unit Price Must Be Between 0-10000.")
                                                End If
                                                rowECUPSERT.Item("ECOM_UNIT_PRICE") = Val(CDbl(ECOM_UNIT_PRICE))
                                            End If
                                        End If
                                    Case "SET_QTY"
                                        Dim SET_QTY As String = rowData.Item("SET_QTY").ToString & String.Empty
                                        If SET_QTY.Length > 0 Then
                                            If Not IsNumeric(SET_QTY) Then
                                                setRowError(rowErr, String.Format("'{0}' Is Not A Valid Set Qty. Must Be Numeric.", SET_QTY))
                                                rowECUPSERT.Item("SET_QTY") = SET_QTY
                                            Else
                                                If Val(CInt(SET_QTY)) > 100 Or Val(CInt(SET_QTY)) < 0 Then
                                                    setRowError(rowErr, "Set Qty Must Be Between 0-100.")
                                                End If
                                                rowECUPSERT.Item("SET_QTY") = Val(CInt(SET_QTY))
                                            End If
                                        End If
                                    Case "SHIP_ECOM"
                                        Dim SHIP_ECOM As String = rowData.Item("SHIP_ECOM").ToString & String.Empty
                                        If SHIP_ECOM.Length > 0 Then
                                            Select Case SHIP_ECOM.ToUpper
                                                Case "YES"
                                                    rowECUPSERT.Item("SHIP_ECOM") = "1"
                                                Case "NO"
                                                    rowECUPSERT.Item("SHIP_ECOM") = "0"
                                                Case Else
                                                    setRowError(rowErr, String.Format("'{0}' Is Not A Valid E-COM Value.  Must Be Yes or No", SHIP_ECOM))
                                                    rowECUPSERT.Item("SHIP_ECOM") = SHIP_ECOM
                                            End Select
                                        End If
                                    Case "SHIP_DROP"
                                        Dim SHIP_DROP As String = rowData.Item("SHIP_DROP").ToString & String.Empty
                                        If SHIP_DROP.Length > 0 Then
                                            Select Case SHIP_DROP.ToUpper
                                                Case "YES"
                                                    rowECUPSERT.Item("SHIP_DROP") = "1"
                                                Case "NO"
                                                    rowECUPSERT.Item("SHIP_DROP") = "0"
                                                Case Else
                                                    setRowError(rowErr, String.Format("'{0}' Is Not A Valid Drop Value.  Must Be Yes or No", SHIP_DROP))
                                                    rowECUPSERT.Item("SHIP_DROP") = SHIP_DROP
                                            End Select
                                        End If
                                    Case "ALT_UNIT_PCT" 'Ignore this and calculate it below after all errors are past.
                                    Case "ALT_UNIT_PRICE"
                                        Dim ALT_UNIT_PRICE As String = rowData.Item("ALT_UNIT_PRICE").ToString & String.Empty
                                        If ALT_UNIT_PRICE.Length > 0 Then
                                            If Not IsNumeric(ALT_UNIT_PRICE) Then
                                                Dim ECOM_UNIT_PRICE As String = rowData.Item("ECOM_UNIT_PRICE").ToString & String.Empty
                                                If Not IsNumeric(ECOM_UNIT_PRICE) Then
                                                    setRowError(rowErr, String.Format("'{0}' Is Not A Valid B2B Price. Must Be Numeric.", ALT_UNIT_PRICE))
                                                    rowECUPSERT.Item("ALT_UNIT_PRICE") = 0
                                                Else
                                                    ALT_UNIT_PRICE = ECOM_UNIT_PRICE
                                                    If Val(CDbl(ALT_UNIT_PRICE)) > 10000 Or Val(CDbl(ALT_UNIT_PRICE)) < 0 Then
                                                        setRowError(rowErr, "B2B Price Must Be Between 0-10000.")
                                                    End If
                                                    rowECUPSERT.Item("ALT_UNIT_PRICE") = Val(CDbl(ALT_UNIT_PRICE))
                                                End If
                                            Else
                                                If Val(CDbl(ALT_UNIT_PRICE)) > 10000 Or Val(CDbl(ALT_UNIT_PRICE)) < 0 Then
                                                    setRowError(rowErr, "B2B Price Must Be Between 0-10000.")
                                                End If
                                                rowECUPSERT.Item("ALT_UNIT_PRICE") = Val(CDbl(ALT_UNIT_PRICE))
                                            End If
                                        End If
                                    Case "ECOM_MIN_QTY_OVERRIDE"
                                        Dim ECOM_MIN_QTY_OVERRIDE As String = rowData.Item("ECOM_MIN_QTY_OVERRIDE").ToString & String.Empty
                                        If ECOM_MIN_QTY_OVERRIDE.Length > 0 Then
                                            If Not IsNumeric(ECOM_MIN_QTY_OVERRIDE) Then
                                                setRowError(rowErr, String.Format("'{0}' Is Not A Valid Min Qty. Must Be Numeric.", ECOM_MIN_QTY_OVERRIDE))
                                                rowECUPSERT.Item("ECOM_MIN_QTY_OVERRIDE") = ECOM_MIN_QTY_OVERRIDE
                                            Else
                                                If Val(CInt(ECOM_MIN_QTY_OVERRIDE)) > 999 Or Val(CInt(ECOM_MIN_QTY_OVERRIDE)) < 0 Then
                                                    setRowError(rowErr, "Min Qty Must Be Between 0-999.")
                                                End If
                                                rowECUPSERT.Item("ECOM_MIN_QTY_OVERRIDE") = Val(CInt(ECOM_MIN_QTY_OVERRIDE))
                                            End If
                                        End If
                                    Case "SHORT_DESC"
                                        Dim SHORT_DESC As String = rowData.Item("SHORT_DESC").ToString & String.Empty
                                        If SHORT_DESC.Length > 0 Then
                                            If SHORT_DESC.Length > 750 Then
                                                setRowError(rowErr, "The Short Description Can Not Exceed 750 Char")
                                            End If
                                            rowECUPSERT.Item("SHORT_DESC") = SHORT_DESC
                                        End If
                                    Case "LONG_DESC"
                                        Dim LONG_DESC As String = rowData.Item("LONG_DESC").ToString & String.Empty
                                        If LONG_DESC.Length > 0 Then
                                            If LONG_DESC.Length > 4000 Then
                                                setRowError(rowErr, "The Long Description Can Not Exceed 4000 Char")
                                            End If
                                            rowECUPSERT.Item("LONG_DESC") = LONG_DESC
                                        End If
                                    Case "PKG_CODE"
                                        Dim PKG_CODE As String = rowData.Item("PKG_CODE").ToString & String.Empty
                                        If PKG_CODE.Length > 0 Then
                                            Dim rowWHTPKGM1 As DataRow = LookUp("WHTPKGM1", PKG_CODE)
                                            If IsNothing(rowWHTPKGM1) Then
                                                setRowError(rowErr, String.Format("'{0}' Is Not A Valid Package Code.", PKG_CODE))
                                                If PKG_CODE.Length > 6 Then
                                                    PKG_CODE = PKG_CODE.Substring(0, 5)
                                                End If
                                                rowECUPSERT.Item("PKG_CODE") = PKG_CODE
                                            Else
                                                rowECUPSERT.Item("PKG_CODE") = PKG_CODE
                                                rowECUPSERT.Item("PKG_DESC") = rowWHTPKGM1.Item("PKG_DESC").ToString & String.Empty
                                                rowECUPSERT.Item("PKG_WT") = Val(rowWHTPKGM1.Item("PKG_WT").ToString & String.Empty)
                                                rowECUPSERT.Item("PKG_L") = Val(rowWHTPKGM1.Item("PKG_L").ToString & String.Empty)
                                                rowECUPSERT.Item("PKG_W") = Val(rowWHTPKGM1.Item("PKG_W").ToString & String.Empty)
                                                rowECUPSERT.Item("PKG_H") = Val(rowWHTPKGM1.Item("PKG_H").ToString & String.Empty)
                                                'rowECUPSERT.Item("PKG_SEQ") = rowWHTPKGM1.Item("PKG_SEQ").ToString & String.Empty
                                                rowECUPSERT.Item("PKG_CUBE") = Val(rowWHTPKGM1.Item("PKG_CUBE").ToString & String.Empty)
                                                rowECUPSERT.Item("PKG_COST") = Val(rowWHTPKGM1.Item("PKG_COST").ToString & String.Empty)
                                                'rowECUPSERT.Item("PKG_CHARGE") = rowWHTPKGM1.Item("PKG_CHARGE").ToString & String.Empty
                                            End If
                                        End If
                                    Case "STYLE_MATL_DESC"
                                        SQLS.Length = 0
                                        SQLS.AppendLine("SELECT STYLE_MATL_DESC")
                                        SQLS.AppendLine("FROM ICTSTYL1")
                                        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                                        ASCMAIN1.sql = SQLS.ToString()
                                        rowECUPSERT.Item("STYLE_MATL_DESC") = ASCDATA1.GetDataValue
                                    Case "PKG_DESC", "PKG_WT", "PKG_L", "PKG_W", "PKG_H", "PKG_SEQ", "PKG_CUBE", "PKG_COST", "PKG_CHARGE", "UPSERT_ERROR", "UPSERT_TYPE"
                                        'Nothing to do with these just skipping then to not have message.
                                    Case "CUST_STYLE_CODE"
                                        Dim CUST_STYLE_CODE As String = rowData.Item("CUST_STYLE_CODE").ToString & String.Empty
                                        Dim CUST_CODE As String = ""
                                        Dim rowECTECOM1_FILTER As DataRow = dst.Tables.Item("ECTECOM1_FILTER").Select(String.Format("ECOM_CODE = '{0}'", ECOM_CODE)).FirstOrDefault
                                        If IsNothing(rowECTECOM1_FILTER) Then
                                            rowECUPSERT.Item("CUST_STYLE_CODE") = CUST_STYLE_CODE
                                            setRowError(rowErr, "Invalid Customer Code")
                                        Else
                                            CUST_CODE = rowECTECOM1_FILTER.Item("CUST_CODE").ToString & String.Empty
                                            If CUST_STYLE_CODE.Length > 0 Then
                                                If CUST_STYLE_CODE.Length = 0 Or CUST_STYLE_CODE.Length > 20 Then
                                                    rowECUPSERT.Item("CUST_STYLE_CODE") = CUST_STYLE_CODE
                                                    setRowError(rowErr, "Cust Style Code Must Be > 1 and < 20")
                                                Else
                                                    If CHECK_CUST_SKU_DUPE(tableData, CUST_CODE, ECOM_CODE, CUST_STYLE_CODE) Then
                                                        rowECUPSERT.Item("CUST_STYLE_CODE") = CUST_STYLE_CODE
                                                        setRowError(rowErr, "Duplicate Cust Style Code")
                                                    Else
                                                        If CUST_CODE.Length <> 0 Then
                                                            rowECUPSERT.Item("CUST_STYLE_CODE") = CUST_STYLE_CODE
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Case Else
                                        Dim msgE As String = String.Format("Invalid Column Encountered: {0}", rowECTUPSRT.Item("COL_NAME").ToString & String.Empty)
                                        MsgBox(msgE, vbExclamation, "Let Wayne Know About This.")
                                End Select
                            End If
                        Next
                        If rowErr.Length = 0 Then
                            Dim ALT_UNIT_PRICE As Double = Val(rowData.Item("ALT_UNIT_PRICE").ToString.Trim & String.Empty)
                            Dim ECOM_UNIT_PRICE As Double = Val(rowData.Item("ECOM_UNIT_PRICE").ToString.Trim & String.Empty)
                            If Not (ECOM_UNIT_PRICE = 0 And ALT_UNIT_PRICE = 0) Then
                                If ECOM_UNIT_PRICE = ALT_UNIT_PRICE Then
                                    rowECUPSERT.Item("ALT_UNIT_PCT") = 0
                                Else
                                    If ECOM_UNIT_PRICE > 0 And (ECOM_UNIT_PRICE - ALT_UNIT_PRICE) > 0 Then
                                        Dim PCT As Double = ((ECOM_UNIT_PRICE - ALT_UNIT_PRICE) / ECOM_UNIT_PRICE) * 100
                                        If PCT <> 100 Then
                                            rowECUPSERT.Item("ALT_UNIT_PCT") = Math.Round(PCT, 2)
                                        End If
                                    Else
                                        setRowError(rowErr, "Can Not Calculate Pct From Prices.")
                                    End If
                                End If
                            End If
                        End If
                        If rowErr.Length <> 0 Then
                            rowECUPSERT.Item("UPSERT_TYPE") = "E"
                            rowECUPSERT.Item("UPSERT_ERROR") = rowErr
                        Else
                            rowECUPSERT.Item("UPSERT_TYPE") = UPSERT_TYPE
                            rowECUPSERT.Item("UPSERT_ERROR") = ""
                        End If
                        dst.Tables.Item("ECUPSERT").Rows.Add(rowECUPSERT)
                    End If
                End If
                rowIndex += 1
            Next
            grdECUPSERT.DisplayLayout.Bands(0).Columns("UPSERT_TYPE").Hidden = False
            grdECUPSERT.DisplayLayout.Bands(0).Columns("UPSERT_ERROR").Hidden = False
            btnUpsertFile.Visible = True
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If
    End Sub

    Private Function CHECK_CUST_SKU_DUPE(ByRef tableData As DataTable, ByVal CUST_CODE As String, ByVal ECOM_CODE As String, ByVal CUST_STYLE_CODE As String) As Boolean
        If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then

        End If

        Dim RetVal As Boolean = False
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("Select Count(*)")
        SQLS.AppendLine("From SOTCSTY1")
        SQLS.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
        SQLS.AppendLine(String.Format("AND CUST_STYLE_CODE = '{0}'", CUST_STYLE_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Dim REC_CNT As Int16 = Val(ASCDATA1.GetDataValue)
        If REC_CNT > 0 Then
            RetVal = True
        Else
            Dim FoundCnt As Int64 = 0
            For Each rowIMPORT As DataRow In tableData.Select()
                If rowIMPORT.Item("ECOM_CODE").ToString & String.Empty = ECOM_CODE And rowIMPORT.Item("CUST_STYLE_CODE").ToString & String.Empty = CUST_STYLE_CODE Then
                    FoundCnt += 1
                End If
            Next
            If FoundCnt > 1 Then
                RetVal = True
            End If
        End If
        Return RetVal
    End Function

    Private Sub btnUpsertFile_Click(sender As Object, e As EventArgs) Handles btnUpsertFile.Click
        Dim rowFilter As String = "UPSERT_TYPE = 'U' OR UPSERT_TYPE = 'A'"

        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Update Records?"
        Dim iMSG As New StringBuilder With {.Length = 0}
        iMSG.AppendLine("This Will Update / Insert All Of")
        iMSG.AppendLine("The Records Below To The E-Commerce")
        iMSG.AppendLine("Styles.")
        iMSG.AppendLine("")
        iMSG.AppendLine("All Rows With Errors Will Be Ignored.")
        iMSG.AppendLine("")
        iMSG.AppendLine("Are You Ready?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then
            Me.Cursor = Cursors.WaitCursor


            SQL.Length = 0
            SQL.AppendLine("SELECT STYLE_CODE FROM ICTSTYL1")
            ASCMAIN1.sql = SQL.ToString
            Dim TABLE_STYLES As String = ASCMAIN1.Temp_Table

            Dim STYLES_FOUND As Boolean = False
            For Each rowECUPSERT As DataRow In dst.Tables("ECUPSERT").Select(rowFilter)
                Dim SC As String = rowECUPSERT.Item("STYLE_CODE").ToString & String.Empty
                If SC.Length > 0 Then
                    Dim SQLU As New System.Text.StringBuilder With {.Length = 0}
                    SQLU.AppendLine(String.Format("INSERT INTO {0} VALUES ('{1}')", TABLE_STYLES, SC))
                    ASCMAIN1.sql = SQLU.ToString
                    ASCDATA1.ExecuteSQL()
                    STYLES_FOUND = True
                End If
            Next
            If STYLES_FOUND = True Then
                dst.Tables.Item("ASTAUDT1_U").Clear()

                SQL.Length = 0
                SQL.AppendLine(String.Format("SELECT * FROM ECTESTY1 WHERE STYLE_CODE IN (SELECT STYLE_CODE FROM {0})", TABLE_STYLES))
                Fill_Records("ECTESTY1_U",, True, SQL.ToString)

                SQL.Length = 0
                SQL.AppendLine(String.Format("SELECT * FROM ECTESTY2 WHERE STYLE_CODE IN (SELECT STYLE_CODE FROM {0})", TABLE_STYLES))
                Fill_Records("ECTESTY2_U",, True, SQL.ToString)

                SQL.Length = 0
                SQL.AppendLine(String.Format("SELECT * FROM ECTSTYL1 WHERE STYLE_CODE IN (SELECT STYLE_CODE FROM {0})", TABLE_STYLES))
                Fill_Records("ECTSTYL1_U",, True, SQL.ToString)

                BeginTrans()
                Dim BATCH_NO As String = ASCMAIN1.Next_Control_No("ECTESTY1_U.BATCH_NO")
                SQL.Length = 0
                SQL.AppendLine("INSERT INTO ECTESTY1_U ")
                SQL.AppendLine(String.Format("SELECT '{0}' AS BATCH_NO, ECTESTY1.* FROM ECTESTY1", BATCH_NO))
                ASCMAIN1.sql = SQL.ToString
                ASCDATA1.ExecuteSQL()

                SQL.Length = 0
                SQL.AppendLine("INSERT INTO ECTESTY2_U ")
                SQL.AppendLine(String.Format("SELECT '{0}' AS BATCH_NO, ECTESTY2.* FROM ECTESTY2", BATCH_NO))
                ASCMAIN1.sql = SQL.ToString
                ASCDATA1.ExecuteSQL()

                SQL.Length = 0
                SQL.AppendLine("INSERT INTO ECTSTYL1_U ")
                SQL.AppendLine(String.Format("SELECT '{0}' AS BATCH_NO, ECTSTYL1.* FROM ECTSTYL1", BATCH_NO))
                ASCMAIN1.sql = SQL.ToString
                ASCDATA1.ExecuteSQL()


                For Each rowECUPSERT As DataRow In dst.Tables("ECUPSERT").Select(rowFilter)
                    Dim STYLE_CODE_UPS As String = rowECUPSERT.Item("STYLE_CODE")
                    Dim COLOR_CODE_UPS As String = rowECUPSERT.Item("COLOR_CODE")
                    Dim ECOM_CODE As String = rowECUPSERT.Item("ECOM_CODE")

                    Dim fltrECTESTY1_U As String = String.Format("STYLE_CODE = '{0}' AND ECOM_CODE = '{1}'", STYLE_CODE_UPS, ECOM_CODE)
                    Dim rowECTESTY1_U As DataRow = dst.Tables("ECTESTY1_U").Select(fltrECTESTY1_U).FirstOrDefault
                    Dim AddECTESTY1_U As Boolean = False
                    Dim cngECTESTY1_U As Boolean = False
                    If IsNothing(rowECTESTY1_U) Then
                        AddECTESTY1_U = True
                        rowECTESTY1_U = dst.Tables("ECTESTY1_U").NewRow
                        rowECTESTY1_U.Item("STYLE_CODE") = STYLE_CODE_UPS
                        rowECTESTY1_U.Item("ECOM_CODE") = ECOM_CODE
                    End If
                    For Each COL_CODE As String In New String() {"ECOM_CODE", "STYLE_CODE", "ECOM_STYLE_STATUS", "ECOM_UNIT_PRICE", "SET_QTY", "SHIP_ECOM", "SHIP_DROP", "ALT_UNIT_PCT", "ALT_UNIT_PRICE", "ECOM_MIN_QTY_OVERRIDE"}
                        If (rowECUPSERT.Item(COL_CODE).ToString.Trim.Length > 0) And (rowECUPSERT.Item(COL_CODE).ToString.Trim <> rowECTESTY1_U.Item(COL_CODE).ToString.Trim) Then
                            addAuditRecord(STYLE_CODE_UPS, ECOM_CODE, COL_CODE, rowECTESTY1_U.Item(COL_CODE).ToString & String.Empty, rowECUPSERT.Item(COL_CODE).ToString & String.Empty, "ECTESTY1", "ASTAUDT1_U")
                            rowECTESTY1_U.Item(COL_CODE) = rowECUPSERT.Item(COL_CODE)
                            cngECTESTY1_U = True
                        End If
                    Next
                    If cngECTESTY1_U = True Then
                        If AddECTESTY1_U Then
                            dst.Tables("ECTESTY1_U").Rows.Add(rowECTESTY1_U)
                        End If
                    End If

                    Dim fltrECTESTY2_U As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}' AND ECOM_CODE = '{2}'", STYLE_CODE_UPS, COLOR_CODE_UPS, ECOM_CODE)
                    Dim rowECTESTY2_U As DataRow = dst.Tables("ECTESTY2_U").Select(fltrECTESTY2_U).FirstOrDefault
                    Dim AddECTESTY2_U As Boolean = False
                    Dim cngECTESTY2_U As Boolean = False
                    If IsNothing(rowECTESTY2_U) Then
                        AddECTESTY2_U = True
                        rowECTESTY2_U = dst.Tables("ECTESTY2_U").NewRow
                        rowECTESTY2_U.Item("STYLE_CODE") = STYLE_CODE_UPS
                        rowECTESTY2_U.Item("COLOR_CODE") = COLOR_CODE_UPS
                        rowECTESTY2_U.Item("ECOM_CODE") = ECOM_CODE
                    End If
                    For Each COL_CODE As String In New String() {"ECOM_STYLE_COLOR_STATUS"}
                        If (rowECUPSERT.Item(COL_CODE).ToString.Trim.Length > 0) And (rowECUPSERT.Item(COL_CODE).ToString.Trim <> rowECTESTY2_U.Item(COL_CODE).ToString.Trim) Then
                            addAuditRecord(STYLE_CODE_UPS, ECOM_CODE, COL_CODE, rowECTESTY2_U.Item(COL_CODE).ToString & String.Empty, rowECUPSERT.Item(COL_CODE), "ECTESTY2", "ASTAUDT1_U")
                            rowECTESTY2_U.Item(COL_CODE) = rowECUPSERT.Item(COL_CODE)
                            cngECTESTY2_U = True
                        End If
                    Next
                    For Each COL_CODE As String In New String() {"CUST_STYLE_CODE"}
                        If rowECUPSERT.Item("CUST_STYLE_CODE").ToString & String.Empty <> "" Then
                            Dim CUST_CODE As String = ""
                            Dim rowECTECOM1_FILTER As DataRow = dst.Tables.Item("ECTECOM1_FILTER").Select(String.Format("ECOM_CODE = '{0}'", ECOM_CODE)).FirstOrDefault
                            If Not IsNothing(rowECTECOM1_FILTER) Then
                                CUST_CODE = rowECTECOM1_FILTER.Item("CUST_CODE").ToString & String.Empty
                                If CUST_CODE.Length <> 0 Then
                                    Dim newSOTCSTY1_U As DataRow = dst.Tables.Item("SOTCSTY1_U").NewRow
                                    newSOTCSTY1_U.Item("CUST_CODE") = CUST_CODE
                                    newSOTCSTY1_U.Item("CUST_STYLE_CODE") = rowECUPSERT.Item("CUST_STYLE_CODE").ToString & String.Empty
                                    newSOTCSTY1_U.Item("STYLE_CODE") = rowECUPSERT.Item("STYLE_CODE").ToString & String.Empty
                                    newSOTCSTY1_U.Item("COLOR_CODE") = rowECUPSERT.Item("COLOR_CODE").ToString & String.Empty
                                    newSOTCSTY1_U.Item("INIT_OPER") = ASCMAIN1.USER_ID
                                    newSOTCSTY1_U.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                                    newSOTCSTY1_U.Item("LAST_OPER") = ASCMAIN1.USER_ID
                                    newSOTCSTY1_U.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                                    dst.Tables.Item("SOTCSTY1_U").Rows.Add(newSOTCSTY1_U)
                                End If
                            End If
                        End If
                    Next
                    If cngECTESTY2_U = True Then
                        If AddECTESTY2_U Then
                            dst.Tables("ECTESTY2_U").Rows.Add(rowECTESTY2_U)
                        End If
                    End If

                    Dim fltrECTSTYL1_U As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE_UPS)
                    Dim rowECTSTYL1_U As DataRow = dst.Tables("ECTSTYL1_U").Select(fltrECTSTYL1_U).FirstOrDefault
                    Dim AddECTSTYL1_U As Boolean = False
                    Dim cngECTSTYL1_U As Boolean = False
                    If IsNothing(rowECTSTYL1_U) Then
                        AddECTSTYL1_U = True
                        rowECTSTYL1_U = dst.Tables("ECTSTYL1_U").NewRow
                        rowECTSTYL1_U.Item("STYLE_CODE") = STYLE_CODE_UPS
                    End If
                    For Each COL_CODE As String In New String() {"SHORT_DESC", "LONG_DESC", "PKG_CODE"}
                        Select Case COL_CODE
                            Case "SHORT_DESC", "LONG_DESC"
                                Dim OLD_TXT As String = rowECTSTYL1_U.Item(COL_CODE).ToString.Trim
                                Dim NEW_TXT As String = rowECUPSERT.Item(COL_CODE).ToString.Trim
                                If NEW_TXT.Length > 0 Then
                                    If OLD_TXT.Length > 200 Then
                                        OLD_TXT = OLD_TXT.Substring(0, 199)
                                    End If
                                    If NEW_TXT.Length > 200 Then
                                        NEW_TXT = NEW_TXT.Substring(0, 199)
                                    End If

                                    If OLD_TXT <> NEW_TXT Then
                                        addAuditRecord(STYLE_CODE_UPS, ECOM_CODE, COL_CODE, rowECTSTYL1_U.Item(COL_CODE).ToString & String.Empty, rowECUPSERT.Item(COL_CODE).ToString & String.Empty, "ECTSTYL1", "ASTAUDT1_U")
                                        rowECTSTYL1_U.Item(COL_CODE) = rowECUPSERT.Item(COL_CODE)
                                        cngECTSTYL1_U = True
                                    End If
                                End If
                            Case Else
                                Dim OLD_TXT As String = rowECTSTYL1_U.Item(COL_CODE).ToString.Trim
                                Dim NEW_TXT As String = rowECUPSERT.Item(COL_CODE).ToString.Trim
                                If NEW_TXT.Length > 0 Then
                                    If rowECUPSERT.Item(COL_CODE).ToString.Trim <> rowECTSTYL1_U.Item(COL_CODE).ToString.Trim Then
                                        addAuditRecord(STYLE_CODE_UPS, ECOM_CODE, COL_CODE, rowECTSTYL1_U.Item(COL_CODE).ToString & String.Empty, rowECUPSERT.Item(COL_CODE).ToString & String.Empty, "ECTSTYL1", "ASTAUDT1_U")
                                        rowECTSTYL1_U.Item(COL_CODE) = rowECUPSERT.Item(COL_CODE)
                                        cngECTSTYL1_U = True
                                    End If
                                End If
                        End Select
                    Next
                    If cngECTSTYL1_U = True Then
                        If AddECTSTYL1_U Then
                            dst.Tables("ECTSTYL1_U").Rows.Add(rowECTSTYL1_U)
                        End If
                    End If
                Next


                Dim Fmsg As New ASFMSGBF
                With Fmsg
                    .Show_grd(dst.Tables.Item("ASTAUDT1_U"), ASCMAIN1.ActiveForm, "Please Review The Audit Trail Before Updating.")
                    If .user_option = 0 Then
                        Update_Record_TDA("ASTAUDT1_U")
                        Update_Record_TDA("ECTESTY1_U")
                        Update_Record_TDA("ECTESTY2_U")
                        Update_Record_TDA("ECTSTYL1_U")
                        Update_Record_TDA("SOTCSTY1_U")

                        CommitTrans()

                        dst.Tables.Item("ASTAUDT1_U").Clear()
                        dst.Tables.Item("ECTESTY1_U").Clear()
                        dst.Tables.Item("ECTESTY2_U").Clear()
                        dst.Tables.Item("ECTSTYL1_U").Clear()
                        dst.Tables.Item("SOTCSTY1_U").Clear()
                        MsgBox("Upsert Is Finished.", vbExclamation, "Done")
                    Else
                        MsgBox("Cancelled After Audit Trail", vbExclamation, "Aborted")
                        Rollback()
                    End If
                End With
                Me.Cursor = Cursors.Default
            Else
                MsgBox("Can't Find Any Styles", vbExclamation, "Problem Upserting")
                Exit Sub
            End If


        Else
            MsgBox("Chicken!", vbOKOnly, "Aborted")
        End If


    End Sub

    Private Sub AddAuditTrail()
        Throw New NotImplementedException()
    End Sub

    Private Sub setRowError(ByRef rowErr As String, ByVal ThisErr As String)
        Dim Prefix As String = ""
        If rowErr.Length > 0 Then
            Prefix = "|"
        End If
        rowErr = rowErr & String.Format(" {0} {1}", Prefix, ThisErr)
    End Sub

    'Private Function ValidateUpsert(ByVal FIELD_NAME As String, ByVal rowData As DataRow, ByRef rowErr As String) As Boolean
    '    Dim RetVal As Boolean = True
    '    Dim Prefix As String = ""
    '    If rowErr.Length > 0 Then
    '        Prefix = "|"
    '    End If
    '    Select Case FIELD_NAME
    '        'Case "ECOM_CODE"
    '        '    Dim rowECTECOM1 As DataRow = LookUp("ECTECOM1", rowData.Item("ECOM_CODE").ToString & String.Empty)
    '        '    If IsNothing(rowECTECOM1) Then
    '        '        Dim err As String = String.Format("'{0}' Is Not A Valid E-COM Partner Code.", rowData.Item("ECOM_CODE").ToString & String.Empty)
    '        '        rowErr = rowErr & String.Format(" {0} {1}", Prefix, err)
    '        '    End If
    '        'Case "STYLE_COLOR"
    '        '    Dim STYLE_CODE As String = rowData.Item("STYLE_CODE").ToString & String.Empty
    '        '    Dim COLOR_CODE As String = rowData.Item("COLOR_CODE").ToString & String.Empty
    '        '    Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
    '        '    If IsNothing(rowICTSTYC1) Then
    '        '        Dim err As String = String.Format("'{0}-{1}' Is Not A Valid Style/Color.", STYLE_CODE, COLOR_CODE)
    '        '        rowErr = rowErr & String.Format(" {0} {1}", Prefix, err)
    '        '    End If
    '        'Case "SHIP_ECOM"
    '        '    Dim SHIP_ECOM As String = rowData.Item("SHIP_ECOM").ToString & String.Empty
    '        '    If SHIP_ECOM <> "Yes" And SHIP_ECOM <> "No" Then
    '        '        Dim err As String = String.Format("'{0}' Is Not A Valid E-COM Value.  Must Be Yes or No", SHIP_ECOM)
    '        '        rowErr = rowErr & String.Format(" {0} {1}", Prefix, err)
    '        '    End If
    '        'Case "SHIP_DROP"
    '        '    Dim SHIP_DROP As String = rowData.Item("SHIP_DROP").ToString & String.Empty
    '        '    If SHIP_DROP <> "Yes" And SHIP_DROP <> "No" Then
    '        '        Dim err As String = String.Format("'{0}' Is Not A Valid Drop Value.  Must Be Yes or No", SHIP_DROP)
    '        '        rowErr = rowErr & String.Format(" {0} {1}", Prefix, err)
    '        '    End If
    '        'Case "SET_QTY"
    '        '    Dim SET_QTY As String = rowData.Item("SET_QTY").ToString & String.Empty
    '        '    If Not IsNumeric(SET_QTY) Then
    '        '        Dim err As String = String.Format("'{0}' Is Not A Valid Set Qty. Must Be Numeric.", SET_QTY)
    '        '        rowErr = rowErr & String.Format(" {0} {1}", Prefix, err)
    '        '    Else
    '        '        If Val(SET_QTY) > 100 Or Val(SET_QTY) < 0 Then
    '        '            Dim err As String = "Set Qty Must Be Between 0-100."
    '        '            rowErr = rowErr & String.Format(" {0} {1}", Prefix, err)
    '        '        End If
    '        '    End If
    '        'Case "ECOM_UNIT_PRICE"
    '        '    Dim ECOM_UNIT_PRICE As String = rowData.Item("ECOM_UNIT_PRICE").ToString & String.Empty
    '        '    If Not IsNumeric(ECOM_UNIT_PRICE) Then
    '        '        Dim err As String = String.Format("'{0}' Is Not A Valid Unit Price. Must Be Numeric.", ECOM_UNIT_PRICE)
    '        '        rowErr = rowErr & String.Format(" {0} {1}", Prefix, err)
    '        '    Else
    '        '        If Val(ECOM_UNIT_PRICE) > 10000 Or Val(ECOM_UNIT_PRICE) < 0 Then
    '        '            Dim err As String = "Unit Price Must Be Between 0-10000."
    '        '            rowErr = rowErr & String.Format(" {0} {1}", Prefix, err)
    '        '        End If
    '        '    End If
    '        'Case "ALT_UNIT_PRICE"
    '        '    Dim ALT_UNIT_PRICE As String = rowData.Item("ALT_UNIT_PRICE").ToString & String.Empty
    '        '    If Not IsNumeric(ALT_UNIT_PRICE) Then
    '        '        Dim err As String = String.Format("'{0}' Is Not A Valid B2B Price. Must Be Numeric.", ALT_UNIT_PRICE)
    '        '        rowErr = rowErr & String.Format(" {0} {1}", Prefix, err)
    '        '    Else
    '        '        If Val(ALT_UNIT_PRICE) > 10000 Or Val(ALT_UNIT_PRICE) < 0 Then
    '        '            Dim err As String = "B2B Price Must Be Between 0-10000."
    '        '            rowErr = rowErr & String.Format(" {0} {1}", Prefix, err)
    '        '        End If
    '        '    End If
    '        'Case "ECOM_MIN_QTY"
    '        '    Dim ECOM_MIN_QTY_OVERRIDE As String = rowData.Item("ECOM_MIN_QTY").ToString & String.Empty
    '        '    If Not IsNumeric(ECOM_MIN_QTY_OVERRIDE) Then
    '        '        Dim err As String = String.Format("'{0}' Is Not A Valid Min Qty. Must Be Numeric.", ECOM_MIN_QTY_OVERRIDE)
    '        '        rowErr = rowErr & String.Format(" {0} {1}", Prefix, err)
    '        '    Else
    '        '        If Val(ECOM_MIN_QTY_OVERRIDE) > 999 Or Val(ECOM_MIN_QTY_OVERRIDE) < 0 Then
    '        '            Dim err As String = "Min Qty Must Be Between 0-999."
    '        '            rowErr = rowErr & String.Format(" {0} {1}", Prefix, err)
    '        '        End If
    '        '    End If
    '    End Select
    '    Return RetVal
    'End Function

    Private Sub btnEDIInventoryFetch_Click(sender As Object, e As EventArgs) Handles btnEDIInventoryFetch.Click
        If cboPartnerEDIInv.Text <> "" Then
            Me.Cursor = Cursors.WaitCursor
            Fill_Records("ICTEDI01", cboPartnerEDIInv.Text)
            AddICTEDI01_FIELDS(cboPartnerEDIInv.Text)
            grdICTEDI01.UpdateData()
            grdICTEDI01.Refresh()
            FilterEDIVariance()
            Me.Cursor = Cursors.Default
        Else
            MsgBox("You First Must Select A Trading Partner!", vbOKOnly, "Pick One")
        End If
    End Sub

    Private Sub AddICTEDI01_FIELDS(ByVal ECOM_CODE As String)
        ECOM_CODE = ECOM_CODE.Replace("'", "")
        Dim EDI_SUPPLIER_NO As String = dst.Tables.Item("EDTXREFX").Select(String.Format("ECOM_CODE = '{0}'", ECOM_CODE)).FirstOrDefault.Item("EDI_SUPPLIER_NO") & String.Empty
        Fill_Records("EDT846OX", EDI_SUPPLIER_NO)

        For Each rowICTEDI01 As DataRow In dst.Tables("ICTEDI01").Select()
            Dim STYLE_CODE As String = rowICTEDI01.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowICTEDI01.Item("COLOR_CODE").ToString & String.Empty
            Dim SET_QTY As Int64 = Val(rowICTEDI01.Item("SET_QTY").ToString & String.Empty)
            Dim ECOM_MIN_QTY_OVERRIDE As Int64 = Val(rowICTEDI01.Item("ECOM_MIN_QTY_OVERRIDE").ToString & String.Empty)
            'If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
            '    If STYLE_CODE = "MTX50556B" Then
            '        Stop
            '    End If
            'End If

            Dim filter As String = String.Format("EDI_STYLE = '{0}' AND EDI_COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
            Dim rowEDT846OX As DataRow = dst.Tables("EDT846OX").Select(filter).FirstOrDefault
            If Not IsNothing(rowEDT846OX) Then
                If IsDate(rowEDT846OX.Item("EDI_REPORT_DATE").ToString & String.Empty) Then
                    rowICTEDI01.Item("EDI_REPORT_DATE") = CDate(rowEDT846OX.Item("EDI_REPORT_DATE").ToString & String.Empty)
                End If
                rowICTEDI01.Item("EDI_AVAIL_QTY") = rowEDT846OX.Item("EDI_AVAIL_QTY").ToString & String.Empty
                rowICTEDI01.Item("EDI_STATUS") = rowEDT846OX.Item("EDI_STATUS").ToString & String.Empty
            End If

            Fill_Records("ICTSTATC", New String() {STYLE_CODE, COLOR_CODE})
            If dst.Tables.Item("ICTSTATC").Rows.Count = 1 Then
                Dim rowICTSTATC As DataRow = dst.Tables.Item("ICTSTATC").Rows(0)
                If Val(rowICTSTATC.Item("WHSE_QTY_ON_HAND").ToString & String.Empty) = 0 Then
                    rowICTEDI01.Item("WHSE_QTY_ON_HAND") = 0
                Else
                    rowICTEDI01.Item("WHSE_QTY_ON_HAND") = Math.Floor(Val(rowICTSTATC.Item("WHSE_QTY_ON_HAND").ToString & String.Empty) / SET_QTY)
                End If
                If Val(rowICTSTATC.Item("WHSE_QTY_PICK").ToString & String.Empty) = 0 Then
                    rowICTEDI01.Item("WHSE_QTY_PICK") = 0
                Else
                    rowICTEDI01.Item("WHSE_QTY_PICK") = Math.Floor(Val(rowICTSTATC.Item("WHSE_QTY_PICK").ToString & String.Empty) / SET_QTY)
                End If
                If Val(rowICTSTATC.Item("OPEN_TO_SELL").ToString & String.Empty) = 0 Then
                    rowICTEDI01.Item("OPEN_TO_SELL") = 0
                Else
                    rowICTEDI01.Item("OPEN_TO_SELL") = Math.Floor(Val(rowICTSTATC.Item("OPEN_TO_SELL").ToString & String.Empty) / SET_QTY)
                End If
                If Val(rowICTSTATC.Item("WHSE_QTY_TRAN").ToString & String.Empty) = 0 Then
                    rowICTEDI01.Item("WHSE_QTY_TRAN") = 0
                Else
                    rowICTEDI01.Item("WHSE_QTY_TRAN") = Math.Floor(Val(rowICTSTATC.Item("WHSE_QTY_TRAN").ToString & String.Empty) / SET_QTY)
                End If
                If Val(rowICTSTATC.Item("WHSE_QTY_ON_ORDER").ToString & String.Empty) = 0 Then
                    rowICTEDI01.Item("WHSE_QTY_ON_ORDER") = 0
                Else
                    rowICTEDI01.Item("WHSE_QTY_ON_ORDER") = Math.Floor(Val(rowICTSTATC.Item("WHSE_QTY_ON_ORDER").ToString & String.Empty) / SET_QTY)
                End If
                If Val(rowICTSTATC.Item("WHSE_QTY_OPEN").ToString & String.Empty) = 0 Then
                    rowICTEDI01.Item("WHSE_QTY_OPEN") = 0
                Else
                    rowICTEDI01.Item("WHSE_QTY_OPEN") = Math.Floor(Val(rowICTSTATC.Item("WHSE_QTY_OPEN").ToString & String.Empty) / SET_QTY)
                End If
                If Val(rowICTSTATC.Item("FUT_AVAIL").ToString & String.Empty) = 0 Then
                    rowICTEDI01.Item("FUT_AVAIL") = 0
                Else
                    rowICTEDI01.Item("FUT_AVAIL") = Math.Floor(Val(rowICTSTATC.Item("FUT_AVAIL").ToString & String.Empty) / SET_QTY)
                End If
                Dim ECOM_QTY As Int64 = rowICTEDI01.Item("OPEN_TO_SELL") - rowICTEDI01.Item("WHSE_QTY_OPEN")
                If ECOM_QTY < ECOM_MIN_QTY_OVERRIDE Then
                    ECOM_QTY = 0
                End If
                rowICTEDI01.Item("ECOM_QTY") = ECOM_QTY

            Else
                rowICTEDI01.Item("WHSE_QTY_ON_HAND") = 0
                rowICTEDI01.Item("WHSE_QTY_PICK") = 0
                rowICTEDI01.Item("OPEN_TO_SELL") = 0
                rowICTEDI01.Item("WHSE_QTY_TRAN") = 0
                rowICTEDI01.Item("WHSE_QTY_ON_ORDER") = 0
                rowICTEDI01.Item("WHSE_QTY_OPEN") = 0
                rowICTEDI01.Item("FUT_AVAIL") = 0
            End If
        Next
    End Sub

    Private Sub chkShowEDIInventoryCols_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowEDIInventoryCols.CheckedChanged
        If Not FormLoading Then
            Dim ShowCols As Boolean = chkShowEDIInventoryCols.Checked
            With grdICTEDI01.DisplayLayout.Bands(0)
                .Columns("WHSE_QTY_ON_HAND").Hidden = ShowCols
                .Columns("WHSE_QTY_PICK").Hidden = ShowCols
                .Columns("WHSE_QTY_TRAN").Hidden = ShowCols
                .Columns("WHSE_QTY_ON_ORDER").Hidden = ShowCols
                .Columns("WHSE_QTY_OPEN").Hidden = ShowCols
            End With
        End If
    End Sub

    Private Sub chkEDIVariance_CheckedChanged(sender As Object, e As EventArgs) Handles chkEDIVariance.CheckedChanged
        If Not FormLoading Then
            FilterEDIVariance()
        End If
    End Sub

    Private Sub FilterEDIVariance()
        Dim dvw As DataView = DirectCast(grdICTEDI01.DataSource, DataTable).DefaultView
        Dim filter As String = ""
        If chkEDIVariance.Checked Then
            filter = "ECOM_QTY > 0 AND ECOM_QTY <> EDI_AVAIL_QTY"
        End If
        dvw.RowFilter = String.Format(filter)
    End Sub

    Private Sub cboPartnerEDIInv_LostFocus(sender As Object, e As EventArgs) Handles cboPartnerEDIInv.LostFocus

    End Sub

    Private Sub btnRefreshUpsert_Click(sender As Object, e As EventArgs) Handles btnRefreshUpsert.Click
        grdECUPSERT.DisplayLayout.Bands(0).Columns("UPSERT_TYPE").Hidden = True
        grdECUPSERT.DisplayLayout.Bands(0).Columns("UPSERT_ERROR").Hidden = True
        Dim ECOM_CODE As String = cboPartnerUpsert.Text
        If ECOM_CODE.ToUpper = "ALL" Then
            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("E1.ECOM_CODE, ")
            SQL.AppendLine("E1.STYLE_CODE, ")
            SQL.AppendLine("SM.STYLE_DESC, ")
            SQL.AppendLine("E2.COLOR_CODE,")
            SQL.AppendLine("C1.COLOR_DESC, ")
            SQL.AppendLine("E1.ECOM_STYLE_STATUS,")
            SQL.AppendLine("E2.ECOM_STYLE_COLOR_STATUS,")
            SQL.AppendLine("E1.ECOM_UNIT_PRICE,")
            SQL.AppendLine("E1.SET_QTY,")
            SQL.AppendLine("NVL(E1.SHIP_ECOM,0) SHIP_ECOM,")
            SQL.AppendLine("NVL(E1.SHIP_DROP,0) SHIP_DROP,")
            SQL.AppendLine("E1.ALT_UNIT_PCT,")
            SQL.AppendLine("E1.ALT_UNIT_PRICE,")
            SQL.AppendLine("NVL(E1.ECOM_MIN_QTY_OVERRIDE,M1.ECOM_MIN_QTY_DEFAULT) ECOM_MIN_QTY_OVERRIDE,")
            SQL.AppendLine("NVL(S1.SHORT_DESC,'') SHORT_DESC,")
            SQL.AppendLine("NVL(S1.LONG_DESC,'') LONG_DESC,")
            SQL.AppendLine("NVL(S1.PKG_CODE,'') PKG_CODE,")
            SQL.AppendLine("P1.PKG_DESC,")
            SQL.AppendLine("P1.PKG_WT,")
            SQL.AppendLine("P1.PKG_L,")
            SQL.AppendLine("P1.PKG_W,")
            SQL.AppendLine("P1.PKG_H,")
            SQL.AppendLine("P1.PKG_SEQ,")
            SQL.AppendLine("P1.PKG_CUBE,")
            SQL.AppendLine("P1.PKG_COST,")
            SQL.AppendLine("P1.PKG_CHARGE,")
            SQL.AppendLine("SM.STYLE_MATL_DESC")
            SQL.AppendLine("FROM ECTESTY1 E1, ECTESTY2 E2, ECTSTYL1 S1, ECTECOM1 M1, ICTSTYL1 SM, ICTCOLR1 C1, WHTPKGM1 P1")
            SQL.AppendLine("WHERE SM.STYLE_CODE = E1.STYLE_CODE")
            SQL.AppendLine("AND E2.COLOR_CODE = C1.COLOR_CODE")
            SQL.AppendLine("AND E1.STYLE_CODE = E2.STYLE_CODE")
            SQL.AppendLine("AND E1.STYLE_CODE = S1.STYLE_CODE (+)")
            SQL.AppendLine("AND E1.ECOM_CODE = M1.ECOM_CODE")
            SQL.AppendLine("AND E2.ECOM_CODE = M1.ECOM_CODE")
            SQL.AppendLine("AND S1.PKG_CODE = P1.PKG_CODE (+)")
            'SQL.AppendLine("AND M1.ECOM_CODE = :PARM1")
            Fill_Records("ECUPSERT", ECOM_CODE, True, SQL.ToString)
        Else
            Fill_Records("ECUPSERT", ECOM_CODE)
        End If
        For Each rowECUPSERT As DataRow In dst.Tables("ECUPSERT").Select()
            Dim EC As String = rowECUPSERT.Item("ECOM_CODE").ToString & String.Empty
            Dim SC As String = rowECUPSERT.Item("STYLE_CODE").ToString & String.Empty
            Dim CC As String = rowECUPSERT.Item("COLOR_CODE").ToString & String.Empty
            Dim filter As String = String.Format("ECOM_CODE = '{0}' AND STYLE_CODE = '{1}' AND COLOR_CODE = '{2}'", EC, SC, CC)
            Dim rowSOTCSTYX As DataRow = dst.Tables.Item("SOTCSTYX").Select(filter).FirstOrDefault
            If Not IsNothing(rowSOTCSTYX) Then
                rowECUPSERT.Item("CUST_STYLE_CODE") = rowSOTCSTYX.Item("CUST_STYLE_CODE").ToString & String.Empty
            Else
                rowECUPSERT.Item("CUST_STYLE_CODE") = ""
            End If

        Next
    End Sub
#End Region
End Class