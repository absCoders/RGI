Imports System.Text
Imports Infragistics.Win.UltraWinGrid
Imports GemBox.Spreadsheet

Public Class SOFOTRP1
    Dim S As New System.Text.StringBuilder With {.Length = 0}
    Dim OrdersRepriced As Boolean = False
    Dim OrdersFinalized As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        dteOBegining.Value = DateSerial(2019, 12, 1)

        If MENU_ITEM_OBJECT = "SOFOTRPI" Then
            InquiryMode = True
        End If

        Check_Form_Options()

        With dst
            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM SOTORDR1")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False)
            With .Tables("SOTORDR1").Columns
                .Add("ORDR_REPRICED", GetType(String))
                .Add("ORDR_TOTAL_ORIG", GetType(Double))
                .Add("ORDR_TOTAL_NEW", GetType(Double))
            End With

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM SOTORDR2")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, True, ,, "ORDR_UNIT_PRICE, ORDR_UNIT_PRICE_CURR")
            With .Tables("SOTORDR2").Columns
                .Add("LINE_REPRICED", GetType(String))
                .Add("CALC_UNIT_PRICE", GetType(Double))
                .Add("ORDR_UNIT_PRICE_ORIG", GetType(Double))
                .Add("NEW_LIST", GetType(String))
                .Add("LINE_EXT", GetType(Double), "ISNULL(ORDR_UNIT_PRICE,0) * ISNULL(ORDR_QTY,0)")
            End With

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ARTCUST1")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ICTCLAS1")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ICTDISC1")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ICTDISC1", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM SOTTARFX")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "SOTTARFX", "**", 0, True)

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("I1.*,")
            S.AppendLine("I2.STYLE_PRICE AS STYLE_PRICE_NEW,")
            S.AppendLine("I2.REPRICE_STYLE")
            S.AppendLine("FROM ICTSTYL1 I1, ICTTFLST I2")
            S.AppendLine("WHERE I1.STYLE_CODE = I2.STYLE_CODE")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ICTTFLST", "**", 0, True,,, "REPRICE_STYLE")
            With .Tables("ICTTFLST").Columns
                .Add("VARIANCE", GetType(System.Decimal), "ISNULL(STYLE_PRICE,0) - ISNULL(STYLE_PRICE_NEW,0)")
            End With

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("C1.STYLE_CODE,")
            S.AppendLine("C1.COLOR_CODE,")
            S.AppendLine("C1.STYLE_COLOR_STATUS")
            S.AppendLine("FROM ICTSTYC1 C1, ICTTFLST I2")
            S.AppendLine("WHERE C1.STYLE_CODE = I2.STYLE_CODE")
            S.AppendLine("AND C1.STYLE_COLOR_STATUS = 'D'")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ICTSTYCX", "**", 0, False)
            With .Tables("ICTSTYCX").Columns
                .Add("INIT_DATE", GetType(System.DateTime))
            End With

            ' create index I_ASTAUDT1_12 ON ASTAUDT1 (TABLE_NAME,COLUMN_NAME,NEW_VALUE,KEY_VALUE);
            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("KEY_VALUE,")
            S.AppendLine("MAX(INIT_DATE) AS INIT_DATE")
            S.AppendLine("FROM ASTAUDT1")
            S.AppendLine("WHERE TABLE_NAME = 'ICTSTYC1'")
            S.AppendLine("AND COLUMN_NAME = 'STYLE_COLOR_STATUS'")
            S.AppendLine("AND NEW_VALUE = 'D'")
            S.AppendLine("AND KEY_VALUE IN")
            S.AppendLine("(")
            S.AppendLine("  SELECT (C1.STYLE_CODE || ':' || C1.COLOR_CODE) AS STYLE_COLOR")
            S.AppendLine("  FROM ICTSTYC1 C1, ICTTFLST I2")
            S.AppendLine("  WHERE C1.STYLE_CODE = I2.STYLE_CODE")
            S.AppendLine("  AND C1.STYLE_COLOR_STATUS = 'D'")
            S.AppendLine(")")
            S.AppendLine("GROUP BY KEY_VALUE")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ICTSTYCZ", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM SOTTFLST")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "SOTTFLST", "**", 0, True)

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM SOTOTRP1")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "SOTOTRP1", "**", 0, True)

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM SOTOTRP2")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "SOTOTRP2", "**", 0, True)

            ' Used to generate emails
            Create_TDA(.Tables.Add, "SOTSREP1", "*")
            Create_TDA(.Tables.Add, "ARTCUSTD", "*")
            Create_TDA(.Tables.Add, "ASTATTA2", "*")
            Create_TDA(.Tables.Add("ARTCUST1_E"), "ARTCUST1", "*")
            Create_TDA(.Tables.Add("SOTORDR1_E"), "SOTORDR1", "*")
            Create_TDA(.Tables.Add("SOTORDR2_E"), "SOTORDR2", "*")
            Create_TDA(.Tables.Add("SOTOTRP1_E"), "SOTOTRP1", "*")
            .Tables("SOTOTRP1_E").Columns.Add("CUST_CODE", GetType(System.String))
            Create_TDA(.Tables.Add("SOTOTRP2_E"), "SOTOTRP2", "")

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ICTSTYL1")
            S.AppendLine("WHERE STYLE_CODE IN ")
            S.AppendLine("(SELECT STYLE_CODE FROM ICTTFLST)")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 0, True, , , "STYLE_PRICE")

        End With

        Fill_Records("ARTCUST1")
        Fill_Records("ICTTFLST")
        Fill_Records("ICTCLAS1")
        Fill_Records("ICTDISC1")
        Fill_Records("SOTTARFX")
        Fill_Records("ICTSTYCX")
        Fill_Records("ICTSTYCZ")
        Fill_Records("SOTTFLST")
        Fill_Records("ICTSTYL1")

        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")
        grdICTTFLST.DataSource = dst.Tables("ICTTFLST")

        Create_Summary(grdSOTORDR1, New String() {"ORDR_TOTAL_ORIG", "ORDR_TOTAL_NEW"})
        Create_Summary(grdSOTORDR1, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDR2, "ORDR_LNO", "Count")
        Create_Summary(grdSOTORDR2, "LINE_EXT")
        Create_Summary(grdICTTFLST, "STYLE_CODE", "Count")

        Sort_grdColumns(grdSOTORDR1, "ORDR_DATE, ORDR_NO".ToLower(), False)
        Sort_grdColumns(grdSOTORDR2, "ORDR_LNO".ToUpper(), False)
        Sort_grdColumns(grdICTTFLST, "STYLE_CODE".ToUpper(), False)

        ASCMAIN1.Add_Value_List(grdSOTORDR1, "ORDR_REPRICED", , New String() {":", "A:Missing Customer", "N:No Repricing", "R:Repriced", "S:Repriced W/Skipped Increase", "C:Excluded Customer", "O:Excluded Order", "F:Flagged As FD Mixed Order"})
        ASCMAIN1.Add_Value_List(grdSOTORDR2, "LINE_REPRICED", , New String() {":", "N:Net priced", "S:Shipped", "L:Not In List", "F:Not Selected For Repricing", "I:List Price Increase", "U:Calculated Price Higher Than Orig", "R:Repriced", "C:Excluded Customer", "O:Excluded Order", "W:Net Priced Lower", "P:No Calculated Change"})

        TABLE_NAME = "SOTORDR1"

        EntryMode = "E"
        'Call Load_Record()
        Setup_SOTORDR2()

        With grdICTTFLST.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.True
            For i As Integer = 0 To .Bands(0).Columns.Count - 1
                .Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
            Next i
            For Each COLNAME As String In New String() {"REPRICE_STYLE"}
                .Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
            Next
            For Each COLNAME As String In New String() {"REPRICE_STYLE"}
                .Bands(0).Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
            Next
        End With

        spl.Panel1Collapsed = True

        Call Mode_Settings(True)
    End Sub

    Sub Check_Inquiry_Mode()
        If InquiryMode Then
        Else
        End If
    End Sub

    Sub Check_Form_Options()

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Done"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Done"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("You're Done?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult <> MsgBoxResult.Yes Then
                    EMsg &= vbCr & "Cancelled By User"
                End If
            Case "Refresh"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Refresh"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("This Will Refresh All Of")
                iMSG.AppendLine("The Orders To Be Re-Priced.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Are You Ready?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult <> MsgBoxResult.Yes Then
                    EMsg &= vbCr & "Cancelled By User"
                End If
            Case "Re-Price"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Re-Price"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("This Will Re-Price All Of")
                iMSG.AppendLine("The Orders Shown Based On")
                iMSG.AppendLine("The Rules Provided.")
                iMSG.AppendLine("")
                iMSG.AppendLine("You Will Then Be Able To")
                iMSG.AppendLine("Review The Results Before")
                iMSG.AppendLine("Finalizing.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Are You Ready?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult <> MsgBoxResult.Yes Then
                    EMsg &= vbCr & "Cancelled By User"
                End If
            Case "Finalize"
                If (ASCMAIN1.USER_ID = "whr" OrElse ASCMAIN1.USER_ID = "wayne" OrElse ASCMAIN1.USER_ID = "edz") Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Finalize"
                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    iMSG.AppendLine("This Process Can Only Be Run Once")
                    iMSG.AppendLine("And May Take A While.")
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Are You Sure You Are Ready?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult <> MsgBoxResult.Yes Then
                        EMsg &= vbCr & "Cancelled By User"
                    End If
                Else
                    MsgBox("Only Wayne Gets To Hit Finalize.", vbOKOnly, "Sorry")
                    Exit Sub
                End If

            Case "Generate Emails"
                If MessageBox.Show("Do you want to Generate Emails?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
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
            Case "Done"
                Call Update_Record()
                Call Mode_Settings(False)
                Me.Close()
            Case "Re-Price"
                Call RePriceOrders()
                UltraExplorerBar1.Groups("Screen Control").Items("Finalize").Settings.Enabled = DefaultableBoolean.True
                OrdersRepriced = True
            Case "Finalize"
                Call FinalizeData()
                Call Update_Record(True)
                Call Mode_Settings(False)
                OrdersFinalized = True
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("Finalize").Settings.Enabled = DefaultableBoolean.False
                    .Items("Re-Price").Settings.Enabled = DefaultableBoolean.False
                    .Items("Refresh").Settings.Enabled = DefaultableBoolean.False
                End With
                'Me.Close()
            Case "Refresh"
                Call RefreshOrders()

            Case "Generate Emails"
                Call GenerateEmails()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("Select Count(*) from SOTOTRP1")
        ASCMAIN1.sql = SQLS.ToString()
        OrdersFinalized = Val(ASCDATA1.GetDataValue) > 0

        With UltraExplorerBar1
            If Not OrdersFinalized Then
                .Groups("Screen Control").Items("Generate Emails").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Re-Price").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Finalize").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
            Else
                .Groups("Screen Control").Items("Generate Emails").Settings.Enabled = DefaultableBoolean.True
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Re-Price").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Finalize").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
            End If

        End With

        'UltraExplorerBar1.Groups("XXXXX").Visible = False

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        'With grdSOTORDR1.DisplayLayout.Bands(0)
        '    For Each thisCOL As String In New String() {"XXXXXX", "XXXXXX", "XXXXXX"}
        '        .Columns.Item(thisCOL).Header.Appearance.BackColor = Drawing.Color.Khaki
        '    Next
        '    For Each thisCOL As String In New String() {"YYYYYY", "YYYYYY", "YYYYYY"}
        '        .Columns.Item(thisCOL).Header.Appearance.BackColor = Drawing.Color.Bisque
        '    Next
        'End With

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        'dst.EnforceConstraints = False
        'dst.Tables("SOTORDR1").Rows.Clear()
        'Fill_Records("SOTORDR1")
        'dst.EnforceConstraints = True
        'Setup_Summary()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Setup_Summary()

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record(Optional ByVal Final As Boolean = False)
        ASCMAIN1.Progress("Updating Data", "")

        Try
            BeginTrans()
            'INIT_LAST("SOTORDR1", True, "", True)
            If Final Then

                For Each TABLE_NAME As String In New String() {"SOTOTRP1", "SOTOTRP2"}

                    ASCDATA1.ExecuteSQL("DELETE FROM " & TABLE_NAME)

                    If dst.Tables(TABLE_NAME).Rows.Count = 0 Then
                        Continue For
                    End If

                    dst.Tables(TABLE_NAME).AcceptChanges()
                    For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")
                        row.SetAdded()
                    Next

                    Create_BAs(TABLE_NAME, True)
                    Update_BAs(TABLE_NAME, True)
                Next

                If ASCMAIN1.Running_in_VS Then
                    Stop
                End If

                'Once In A Lifetime
                Update_Record_TDA("ICTSTYL1")
                Update_Record_TDA("SOTORDR2")

                For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
                    Dim ORDR_GROUP_NO As String = rowSOTORDR1.Item("ORDR_GROUP_NO").ToString & String.Empty
                    ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
                    ASCDATA1.ExecuteSQL()
                Next
            End If

            Update_Record_TDA("SOTTARFX")
            Update_Record_TDA("ICTTFLST")
            Update_Record_TDA("SOTTFLST")

            CommitTrans("Update Complete")

        Catch ex As Exception
            Rollback("Error Updating data: " & ex.Message)

        Finally
            ASCMAIN1.Progress("", "")
        End Try

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        'Select Case COLUMN_NAME
        '    Case "JOB_NO"
        '        sql_where = "JOB_STATUS = 'O' and SITE_VISITS > 0"
        'End Select
    End Sub

    Sub Setup_Summary()
        Dim sqlwhere As String = ""
        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor

        'dst.Tables("SOTXXXXX").Rows.Clear()

        dst.EnforceConstraints = False
        'Fill_Records("SOTXXXXX")

        'grdSOTXXXXX.DataSource = dst.Tables("SOTXXXXX")

        ASCMAIN1.Progress("")
        'grdSOTXXXXX.Update()
        'grdSOTXXXXX.Refresh()
        Me.Cursor = Cursors.Default
    End Sub

    Sub Print_Report()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'CR_params.Add("GROUP_BY", optGroupBy.Value)
        'Generate_Report("PMRVIST1", "Open Site Visit Report")
        'Print_Report_End()
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdSOTORDR1, "SSBBBB", "Show Filter", "Show GroupBox", "View Order", "Exclude Customer", "Exclude Order", "Remove Exclusion", "Add To FD Mix", "Remove From FD Mix")
        Call Load_Popup_Menu(grdSOTORDR2, "SSBB", "Show Filter", "Show GroupBox", "Add Style To Reprice List", "Style To Clipboard", "Show FD Calc")
        Call Load_Popup_Menu(grdICTTFLST, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)
        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        If OrdersFinalized Then
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

            'Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
            'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
            'Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

            Select Case e.SourceControl.Name
                Case "grdSOTORDR2"
                    If Not IsNothing(grdSOTORDR1.ActiveRow) Then
                        Dim ORDR_REPRICED As String = grdSOTORDR1.ActiveRow.Cells.Item("ORDR_REPRICED").Value.ToString & String.Empty
                        If ORDR_REPRICED = "F" Then
                            tlb_pop.Tools("Show FD Calc").SharedProps.Visible = True
                        Else
                            tlb_pop.Tools("Show FD Calc").SharedProps.Visible = False
                        End If
                    End If
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
            Case "View Order"
                If Not IsNothing(grdSOTORDR1.ActiveRow) Then
                    Dim ORDR_GROUP_NO As String = grdSOTORDR1.ActiveRow.Cells.Item("ORDR_GROUP_NO").Text
                    Dim ORDR_NO As String = grdSOTORDR1.ActiveRow.Cells.Item("ORDR_NO").Text
                    If ORDR_GROUP_NO.Length > 0 Then
                        Context_Launch("View", Column_Values("ORDR_NO", ORDR_GROUP_NO), "Sales Order Entry", "SOFORDRI")
                    End If
                End If
            Case "Exclude Customer"
                If Not IsNothing(grdSOTORDR1.ActiveRow) Then
                    Dim CUST_CODE As String = grdSOTORDR1.ActiveRow.Cells.Item("CUST_CODE").Text
                    Dim ORDR_REPRICED As String = grdSOTORDR1.ActiveRow.Cells.Item("ORDR_REPRICED").Value & String.Empty
                    Select Case ORDR_REPRICED
                        Case "C"
                            MsgBox("Customer Already Excluded", vbOKOnly, "Duplicate")
                        Case "O"
                            MsgBox("Order Already Excluded", vbOKOnly, "Duplicate")
                        Case Else
                            Dim EXCLUDED As Boolean = EXCLUDE_TYPE("A", "ARTCUST1", CUST_CODE)
                            If EXCLUDED Then
                                Dim FILTERC As String = String.Format("CUST_CODE = '{0}'", CUST_CODE)
                                For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select(FILTERC)
                                    rowSOTORDR1.Item("ORDR_REPRICED") = "C"
                                    Dim FILTERD As String = String.Format("ORDR_NO = '{0}'", rowSOTORDR1.Item("ORDR_NO").ToString & String.Empty)
                                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(FILTERD)
                                        rowSOTORDR2.Item("LINE_REPRICED") = "C"
                                    Next
                                Next
                            End If
                    End Select
                End If
            Case "Exclude Order"
                If Not IsNothing(grdSOTORDR1.ActiveRow) Then
                    Dim ORDR_NO As String = grdSOTORDR1.ActiveRow.Cells.Item("ORDR_NO").Text
                    Dim ORDR_REPRICED As String = grdSOTORDR1.ActiveRow.Cells.Item("ORDR_REPRICED").Value & String.Empty
                    Select Case ORDR_REPRICED
                        Case "C"
                            MsgBox("Customer Already Excluded", vbOKOnly, "Duplicate")
                        Case "O"
                            MsgBox("Order Already Excluded", vbOKOnly, "Duplicate")
                        Case Else
                            Dim EXCLUDED As Boolean = EXCLUDE_TYPE("A", "SOTORDR1", ORDR_NO)
                            If EXCLUDED Then
                                grdSOTORDR1.ActiveRow.Cells.Item("ORDR_REPRICED").Value = "O"
                                Dim FILTERD As String = String.Format("ORDR_NO = '{0}'", grdSOTORDR1.ActiveRow.Cells.Item("ORDR_NO").Text.ToString & String.Empty)
                                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(FILTERD)
                                    rowSOTORDR2.Item("LINE_REPRICED") = "O"
                                Next
                            End If
                    End Select
                End If
            Case "Remove Exclusion"
                If Not IsNothing(grdSOTORDR1.ActiveRow) Then
                    Dim ORDR_REPRICED As String = grdSOTORDR1.ActiveRow.Cells.Item("ORDR_REPRICED").Value
                    If ORDR_REPRICED <> "C" And ORDR_REPRICED <> "O" Then
                        MsgBox("This Order Or Cusomer Is Not Excluded!", vbExclamation, "Hmmm")
                    Else
                        Dim EXCL_TYPE As String = ""
                        Dim EXCL_CODE As String = ""
                        Dim FILTER_REM As String = ""
                        If ORDR_REPRICED = "C" Then
                            EXCL_CODE = grdSOTORDR1.ActiveRow.Cells.Item("CUST_CODE").Text
                            EXCL_TYPE = "ARTCUST1"
                            FILTER_REM = String.Format("CUST_CODE = '{0}'", EXCL_CODE)
                        Else
                            EXCL_CODE = grdSOTORDR1.ActiveRow.Cells.Item("ORDR_NO").Text
                            EXCL_TYPE = "SOTORDR1"
                            FILTER_REM = String.Format("ORDR_NO = '{0}'", EXCL_CODE)
                        End If
                        Dim EXCLUDED As Boolean = EXCLUDE_TYPE("D", EXCL_TYPE, EXCL_CODE)
                        If EXCLUDED Then
                            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select(FILTER_REM)
                                rowSOTORDR1.Item("ORDR_REPRICED") = ""
                                rowSOTORDR1.Item("ORDR_TOTAL_NEW") = rowSOTORDR1.Item("ORDR_TOTAL_ORIG")
                                Dim FILTER_REM2 As String = String.Format("ORDR_NO = '{0}'", rowSOTORDR1.Item("ORDR_NO").ToString & String.Empty)
                                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(FILTER_REM2)
                                    rowSOTORDR2.Item("LINE_REPRICED") = ""
                                    rowSOTORDR2.Item("ORDR_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE_ORIG")
                                    rowSOTORDR2.Item("CALC_UNIT_PRICE") = Null
                                Next
                            Next
                            RePriceOrders()
                        End If
                    End If
                End If
            Case "Add Style To Reprice List"
                If Not IsNothing(grdSOTORDR2.ActiveRow) Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Add Style To Reprice List"
                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    Dim STYLE_CODE_SEL As String = grdSOTORDR2.ActiveRow.Cells.Item("STYLE_CODE").Value
                    Dim fltICTTFLST As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE_SEL)
                    Dim rowICTTFLST As DataRow = dst.Tables("ICTTFLST").Select(fltICTTFLST).FirstOrDefault
                    If rowICTTFLST.Item("REPRICE_STYLE").ToString & String.Empty = "1" Then
                        iMSG.AppendLine("This Style Is Already Marked For Repricing")
                        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
                    Else
                        iMSG.AppendLine("Are You Sure You Want To Add")
                        iMSG.AppendLine(String.Format("Style {0} To The Reprice List?", STYLE_CODE_SEL))
                        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                        If iResult = MsgBoxResult.Yes Then
                            rowICTTFLST.Item("REPRICE_STYLE") = "1"
                            AddToRepriceList(STYLE_CODE_SEL)
                        End If
                    End If
                End If
            Case "Add To FD Mix"
                If Not OrdersRepriced Then
                    MsgBox("You Should Re-price All Orders Before Making Changes", vbOKOnly, "Please Reprice")
                Else
                    If Not IsNothing(grdSOTORDR1.ActiveRow) Then
                        Dim ORDR_NO As String = grdSOTORDR1.ActiveRow.Cells.Item("ORDR_NO").Text
                        Dim ORDR_REPRICED As String = grdSOTORDR1.ActiveRow.Cells.Item("ORDR_REPRICED").Value & String.Empty
                        Select Case ORDR_REPRICED
                            Case "F"
                                MsgBox("Order Already In FD Mix List", vbOKOnly, "Duplicate")
                            Case Else
                                Dim ADDED As Boolean = AddRemoveFE(ORDR_NO, True)
                                If ADDED Then
                                    Dim FILTERD As String = String.Format("ORDR_NO = '{0}'", grdSOTORDR1.ActiveRow.Cells.Item("ORDR_NO").Text.ToString & String.Empty)
                                    Dim rowSOTORDR1 As DataRow = dst.Tables.Item("SOTORDR1").Select(FILTERD).FirstOrDefault
                                    If Not IsNothing(rowSOTORDR1) Then
                                        rowSOTORDR1.Item("ORDR_REPRICED") = "F"
                                        rowSOTORDR1.Item("ORDR_TOTAL_NEW") = rowSOTORDR1.Item("ORDR_TOTAL_ORIG")
                                        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(FILTERD)
                                            rowSOTORDR2.Item("LINE_REPRICED") = ""
                                            rowSOTORDR2.Item("ORDR_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE_ORIG")
                                            rowSOTORDR2.Item("CALC_UNIT_PRICE") = Null
                                        Next
                                        RePriceOrders()
                                    End If
                                End If
                        End Select
                    End If
                End If
            Case "Remove From FD Mix"
                If Not OrdersRepriced Then
                    MsgBox("You Should Re-price All Orders Before Making Changes", vbOKOnly, "Please Reprice")
                Else
                    If Not IsNothing(grdSOTORDR1.ActiveRow) Then
                        Dim ORDR_NO As String = grdSOTORDR1.ActiveRow.Cells.Item("ORDR_NO").Text
                        Dim ORDR_REPRICED As String = grdSOTORDR1.ActiveRow.Cells.Item("ORDR_REPRICED").Value & String.Empty
                        Select Case ORDR_REPRICED
                            Case "F"
                                Dim REMOVED As Boolean = AddRemoveFE(ORDR_NO, False)
                                If REMOVED Then
                                    Dim FILTERD As String = String.Format("ORDR_NO = '{0}'", grdSOTORDR1.ActiveRow.Cells.Item("ORDR_NO").Text.ToString & String.Empty)
                                    Dim rowSOTORDR1 As DataRow = dst.Tables.Item("SOTORDR1").Select(FILTERD).FirstOrDefault
                                    If Not IsNothing(rowSOTORDR1) Then
                                        rowSOTORDR1.Item("ORDR_REPRICED") = ""
                                        rowSOTORDR1.Item("ORDR_TOTAL_NEW") = rowSOTORDR1.Item("ORDR_TOTAL_ORIG")
                                        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(FILTERD)
                                            rowSOTORDR2.Item("LINE_REPRICED") = ""
                                            rowSOTORDR2.Item("ORDR_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE_ORIG")
                                            rowSOTORDR2.Item("CALC_UNIT_PRICE") = Null
                                        Next
                                        RePriceOrders()
                                    End If
                                End If
                            Case Else
                                MsgBox("Order Not In FD Mix List", vbOKOnly, "Duplicate")
                        End Select
                    End If
                End If
            Case "Show FD Calc"
                Dim FEFD As New FEFDPrice(Me, grdSOTORDR2.ActiveRow.Cells.Item("STYLE_CODE").Text.ToString & String.Empty, 1, True)
            Case "Style To Clipboard"
                Dim STYLE_COPIED As String = grdSOTORDR2.ActiveRow.Cells.Item("STYLE_CODE").Text.ToString & String.Empty
                My.Computer.Clipboard.SetText(STYLE_COPIED)
                MsgBox(String.Format("{0} Copied To Clipboard.", STYLE_COPIED), vbOKOnly, "Copied")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

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

#Region "Form Controls"
    Private Sub cmdRefresh_Click(sender As Object, e As EventArgs)
        RePriceOrders()
    End Sub

    Private Sub chkSomething_CheckedChanged(sender As Object, e As EventArgs) Handles chkWhsMS.CheckedChanged

    End Sub

    Private Sub grdICTTFLST_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdICTTFLST.AfterCellUpdate
        If e.Cell.Column.Key = "REPRICE_STYLE" Then
            Dim STYLE_CODE_SEL As String = grdICTTFLST.ActiveRow.Cells.Item("STYLE_CODE").Value
            AddToRepriceList(STYLE_CODE_SEL)
        End If
    End Sub

    Private Sub grdSOTORDR1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTORDR1.AfterRowActivate
        Setup_SOTORDR2()
    End Sub

    Private Sub grdSOTORDR2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTORDR2.InitializeRow
        Dim CALC_UNIT_PRICE As Double = Val(e.Row.Cells("CALC_UNIT_PRICE").Text & String.Empty)
        e.Row.ToolTipText = String.Format("Calculated Value: {0}", Format(Val(CALC_UNIT_PRICE), "###,###.00"))
    End Sub

    Private Sub grdSOTORDR2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTORDR2.AfterRowActivate
        Dim CALC_UNIT_PRICE As Double = Val(grdSOTORDR2.ActiveRow.Cells("CALC_UNIT_PRICE").Text & String.Empty)
        grdSOTORDR2.ActiveRow.ToolTipText = String.Format("Calculated Value: {0}", Format(Val(CALC_UNIT_PRICE), "###,###.00"))
    End Sub
#End Region

#Region "Custom Methods"
    Private Function AddRemoveFE(ByVal ORDR_NO As String, ByVal ADD_ORDER As Boolean) As Boolean
        Dim RetVal As Boolean = True
        Dim fltSOTTFLST As String = String.Format("ORDR_NO = '{0}'", ORDR_NO)
        Dim rowSOTTFLST As DataRow = dst.Tables.Item("SOTTFLST").Select(fltSOTTFLST).FirstOrDefault

        If ADD_ORDER Then
            If IsNothing(rowSOTTFLST) Then
                Dim newSOTTFLST As DataRow = dst.Tables.Item("SOTTFLST").NewRow
                newSOTTFLST.Item("ORDR_NO") = ORDR_NO
                newSOTTFLST.Item("FEFD_TYPE") = "FDMIX"
                dst.Tables.Item("SOTTFLST").Rows.Add(newSOTTFLST)
            Else
                RetVal = False
            End If
        Else
            If IsNothing(rowSOTTFLST) Then
                RetVal = False
            Else
                rowSOTTFLST.Delete()
            End If
        End If

        Return RetVal
    End Function

    Private Sub AddToRepriceList(ByVal STYLE_CODE_SEL As String)
        Me.Cursor = Cursors.WaitCursor
        Dim ORDR_NOs As New List(Of String)
        Dim fltSOTORDR2 As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE_SEL)
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(fltSOTORDR2)
            If Not ORDR_NOs.Contains(rowSOTORDR2.Item("ORDR_NO").ToString & String.Empty) Then
                ORDR_NOs.Add(rowSOTORDR2.Item("ORDR_NO").ToString & String.Empty)
            End If
        Next
        For Each ORDR_NO As String In ORDR_NOs
            Dim fltORDRNO As String = String.Format("ORDR_NO = '{0}'", ORDR_NO)
            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select(fltORDRNO)
                rowSOTORDR1.Item("ORDR_REPRICED") = ""
                rowSOTORDR1.Item("ORDR_TOTAL_NEW") = rowSOTORDR1.Item("ORDR_TOTAL_ORIG")
                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(fltORDRNO)
                    rowSOTORDR2.Item("LINE_REPRICED") = ""
                    rowSOTORDR2.Item("CALC_UNIT_PRICE") = Null
                    rowSOTORDR2.Item("ORDR_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE_ORIG")
                Next
            Next
        Next
        RePriceOrders()
        Me.Cursor = Cursors.Default
    End Sub

    Private Function EXCLUDE_TYPE(ByVal AddDel As String, ByVal EXCL_TYPE As String, ByVal EXCL_CODE As String) As Boolean
        Dim RetVal As Boolean = True
        Dim FILTER As String = String.Format("EXCL_TYPE = '{0}' AND EXCL_CODE = '{1}'", EXCL_TYPE, EXCL_CODE)
        Dim rowSOTTARFX As DataRow = dst.Tables.Item("SOTTARFX").Select(FILTER).FirstOrDefault
        Select Case AddDel
            Case "A"
                If IsNothing(rowSOTTARFX) Then
                    Dim newSOTTARFX As DataRow = dst.Tables.Item("SOTTARFX").NewRow
                    newSOTTARFX.Item("EXCL_TYPE") = EXCL_TYPE
                    newSOTTARFX.Item("EXCL_CODE") = EXCL_CODE
                    dst.Tables.Item("SOTTARFX").Rows.Add(newSOTTARFX)
                Else
                    MsgBox("Exclusion Already Exists", vbExclamation, "Duplicate")
                End If
            Case "D"
                If IsNothing(rowSOTTARFX) Then
                    MsgBox("Exclusion Does Not Exists", vbExclamation, "Duplicate")
                Else
                    rowSOTTARFX.Delete()
                End If
            Case Else
                MsgBox("Error In Exclusion", vbExclamation, "Please Let Wayne Know")
                RetVal = False
        End Select
        Return RetVal
    End Function

    Private Sub FILL_EXTRA_FIELDS()
        ASCMAIN1.Progress("Calculating Original Amts", "")
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("", "ORDR_NO")
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO").ToString & String.Empty
            Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE").ToString & String.Empty
            ASCMAIN1.Progress("", ORDR_NO)

            Dim FILTERO As String = String.Format("EXCL_TYPE = 'SOTORDR1' and EXCL_CODE = '{0}'", ORDR_NO)
            Dim rowO As DataRow = dst.Tables.Item("SOTTARFX").Select(FILTERO).FirstOrDefault
            If Not IsNothing(rowO) Then

                rowSOTORDR1.Item("ORDR_REPRICED") = "O"
                Dim FILTERD As String = String.Format("ORDR_NO = '{0}'", rowSOTORDR1.Item("ORDR_NO").ToString & String.Empty)
                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(FILTERD)
                    rowSOTORDR2.Item("LINE_REPRICED") = "O"
                Next
            End If
            Dim FILTERC As String = String.Format("EXCL_TYPE = 'ARTCUST1' and EXCL_CODE = '{0}'", CUST_CODE)
            Dim rowC As DataRow = dst.Tables.Item("SOTTARFX").Select(FILTERC).FirstOrDefault
            If Not IsNothing(rowC) Then
                rowSOTORDR1.Item("ORDR_REPRICED") = "C"
                Dim FILTERD As String = String.Format("ORDR_NO = '{0}'", rowSOTORDR1.Item("ORDR_NO").ToString & String.Empty)
                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(FILTERD)
                    rowSOTORDR2.Item("LINE_REPRICED") = "C"
                Next
            End If

            Dim FILTER2 As String = String.Format("ORDR_NO = '{0}'", ORDR_NO)
            Dim ORDR_TOTAL_ORIG As Double = 0
            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(FILTER2)
                'Need Quilifiers Here to Make Sure we only re-price what we need.
                Dim fltICTTFLST As String = String.Format("STYLE_CODE = '{0}'", rowSOTORDR2.Item("STYLE_CODE").ToString & String.Empty)
                Dim rowICTTFLST As DataRow = dst.Tables.Item("ICTTFLST").Select(fltICTTFLST).FirstOrDefault
                If IsNothing(rowICTTFLST) Then
                    rowSOTORDR2.Item("NEW_LIST") = "N/A"
                Else
                    rowSOTORDR2.Item("NEW_LIST") = Format(Val(rowICTTFLST.Item("STYLE_PRICE_NEW").ToString & String.Empty), "###,##0.00")
                End If
                rowSOTORDR2.Item("ORDR_UNIT_PRICE_ORIG") = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                ORDR_TOTAL_ORIG += (Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE").ToString & String.Empty) * Val(rowSOTORDR2.Item("ORDR_QTY").ToString & String.Empty))
            Next

            Dim ftrSOTTFLST As String = String.Format("ORDR_NO = '{0}'", ORDR_NO)
            Dim rowSOTTFLST As DataRow = dst.Tables.Item("SOTTFLST").Select(ftrSOTTFLST).FirstOrDefault
            If Not IsNothing(rowSOTTFLST) Then
                rowSOTORDR1.Item("ORDR_REPRICED") = "F"
                'Dim FILTERD As String = String.Format("ORDR_NO = '{0}'", rowSOTORDR1.Item("ORDR_NO").ToString & String.Empty)
                'For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(FILTERD)
                '    rowSOTORDR2.Item("LINE_REPRICED") = "C"
                'Next
            End If

            rowSOTORDR1.Item("ORDR_TOTAL_ORIG") = ORDR_TOTAL_ORIG
            rowSOTORDR1.Item("ORDR_TOTAL_NEW") = ORDR_TOTAL_ORIG
        Next
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub FILL_INIT_DATES()
        For Each rowICTSTYCX As DataRow In dst.Tables("ICTSTYCX").Select("", "STYLE_CODE, COLOR_CODE")
            Dim KEY_VALUE As String = String.Format("{0}:{1}", rowICTSTYCX.Item("STYLE_CODE").ToString & String.Empty, rowICTSTYCX.Item("COLOR_CODE").ToString & String.Empty)
            ASCMAIN1.Progress("-", KEY_VALUE)
            Dim fltICTSTYCZ As String = String.Format("KEY_VALUE = '{0}'", KEY_VALUE)
            Dim rowICTSTYCZ As DataRow = dst.Tables.Item("ICTSTYCZ").Select(fltICTSTYCZ).FirstOrDefault
            If Not IsNothing(rowICTSTYCZ) Then
                Dim INIT_DATE As String = rowICTSTYCZ.Item("INIT_DATE").ToString & String.Empty
                If IsDate(INIT_DATE) Then
                    rowICTSTYCX.Item("INIT_DATE") = CDate(INIT_DATE)
                Else
                    rowICTSTYCX.Item("INIT_DATE") = DateSerial(1900, 1, 1)
                End If
            Else
                rowICTSTYCX.Item("INIT_DATE") = DateSerial(1900, 1, 1)
            End If
        Next
    End Sub

    Private Sub RefreshOrders()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Orders", "")

        dst.Tables.Item("SOTORDR1").Clear()
        dst.Tables.Item("SOTORDR2").Clear()

        Dim dBegin As String = Format(CDate(dteOBegining.Value), "dd-MMM-yyyy")

        S.Length = 0
        S.AppendLine("SELECT ORDR_NO")
        S.AppendLine("FROM SOTORDR1")
        S.AppendLine("WHERE ORDR_STATUS = 'O'")
        S.AppendLine(String.Format("AND ORDR_DATE >= '{0}'", dBegin))
        If chkWhsMS.Checked And chkWhsNY.Checked Then
            S.AppendLine("AND (WHSE_CODE = 'MS' OR WHSE_CODE = 'NY')")
        Else
            If chkWhsMS.Checked Then
                S.AppendLine("AND WHSE_CODE = 'MS'")
            End If
            If chkWhsNY.Checked Then
                S.AppendLine("AND WHSE_CODE = 'NY'")
            End If
        End If
        If chkEXCL_EDI.Checked And chkEXCL_WEB.Checked Then
            S.AppendLine("AND (ORDR_SOURCE <> 'E' AND ORDR_SOURCE <> 'W')")
        Else
            If chkEXCL_EDI.Checked Then
                S.AppendLine("AND ORDR_SOURCE <> 'E'")
            End If
            If chkEXCL_WEB.Checked Then
                S.AppendLine("AND ORDR_SOURCE <> 'W'")
            End If
        End If
        ASCMAIN1.sql = S.ToString
        Dim tmpORDR_NO As String = ASCMAIN1.Temp_Table

        S.Length = 0
        S.AppendLine("SELECT *")
        S.AppendLine("FROM SOTORDR1")
        S.AppendLine("WHERE ORDR_NO IN")
        S.AppendLine(String.Format("(SELECT ORDR_NO FROM {0})", tmpORDR_NO))
        Fill_Records("SOTORDR1", , , S.ToString)

        S.Length = 0
        S.AppendLine("SELECT *")
        S.AppendLine("FROM SOTORDR2")
        S.AppendLine("WHERE ORDR_NO IN")
        S.AppendLine(String.Format("(SELECT ORDR_NO FROM {0})", tmpORDR_NO))
        Fill_Records("SOTORDR2", , , S.ToString)

        FILL_EXTRA_FIELDS()

        ASCMAIN1.Progress("Calculating Discontinue Dates", "")
        FILL_INIT_DATES()
        ASCMAIN1.Progress("", "")

        UltraExplorerBar1.Groups("Screen Control").Items("Re-Price").Settings.Enabled = DefaultableBoolean.True

        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub RePriceOrders()
        'This is where all The Fun Happens
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Repricing Orders", "")

        UltraExplorerBar1.Groups("Screen Control").Items("Re-Price").Settings.Enabled = DefaultableBoolean.False

        Dim tmpFltr As String = "ISNULL(ORDR_REPRICED,'') = '' OR ORDR_REPRICED = 'F'"

        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select(tmpFltr, "ORDR_DATE, ORDR_NO")
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO").ToString & String.Empty

            Dim ORDR_DATE As Date = CDate(rowSOTORDR1.Item("ORDR_DATE").ToString & String.Empty)
            ASCMAIN1.Progress("-", ORDR_DATE.ToShortDateString())
            Dim ORDR_REPRICED As String = rowSOTORDR1.Item("ORDR_REPRICED").ToString & String.Empty
            If ORDR_REPRICED = "F" Then
                ' FEFDType = "FE" Or FEFDType = "FD" Or FEFDType = "FEMIX" Or FEFDType = "FDMIX"
                Dim fltSOTTFLST As String = String.Format("ORDR_NO = '{0}'", ORDR_NO)
                Dim rowSOTTFLST As DataRow = dst.Tables.Item("SOTTFLST").Select(fltSOTTFLST).FirstOrDefault
                If Not IsNothing(rowSOTTFLST) Then
                    Dim ValChange As Double = RePriceFEFDOrder(ORDR_NO, rowSOTTFLST.Item("FEFD_TYPE").ToString & String.Empty)
                    'rowSOTORDR1.Item("ORDR_TOTAL_NEW") = Val(rowSOTORDR1.Item("ORDR_TOTAL_NEW").ToString & String.Empty) + ValChange
                    rowSOTORDR1.Item("ORDR_TOTAL_NEW") = ValChange
                End If
            Else
                Dim ValChange As Double = RePriceOrder(ORDR_NO, rowSOTORDR1)
                If ValChange = 0 Then
                    rowSOTORDR1.Item("ORDR_REPRICED") = "N"
                Else
                    If rowSOTORDR1.Item("ORDR_REPRICED").ToString & String.Empty <> "S" Then
                        rowSOTORDR1.Item("ORDR_REPRICED") = "R"
                    End If
                    rowSOTORDR1.Item("ORDR_TOTAL_NEW") = Val(rowSOTORDR1.Item("ORDR_TOTAL_NEW").ToString & String.Empty) + ValChange
                End If
            End If
        Next
        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default
        MsgBox("Orders Re-Priced", vbOKOnly, "Done")
    End Sub

    Private Function RePriceOrder(ByVal ORDR_NO As String, ByRef rowSOTORDR1 As DataRow) As Double
        Dim RetVal As Double
        Dim ORIG_VALUE As Double = 0
        Dim NEW_VALUE As Double = 0

        Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE").ToString & String.Empty
        Dim ORDR_DATE As DateTime = CDate(rowSOTORDR1.Item("ORDR_DATE").ToString & String.Empty)

        Dim rowARTCUST1 As DataRow = dst.Tables.Item("ARTCUST1").Select(String.Format("CUST_CODE = '{0}'", CUST_CODE)).FirstOrDefault
        If IsNothing(rowARTCUST1) Then
            rowSOTORDR1.Item("ORDR_REPRICED") = "A"
            Return 0
        End If

        Dim PINNED_CUST_PRICE_TIER_PVC As String = rowARTCUST1.Item("CUST_PRICE_TIER_PVC") & ""
        If PINNED_CUST_PRICE_TIER_PVC = "" Then
            PINNED_CUST_PRICE_TIER_PVC = "PC"
        End If

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
            Dim Discounts As New List(Of DISCOUNTS)
            Dim ORDR_UNIT_PRICE As Double = 0
            Dim ORDR_UNIT_PRICE_NEW As Double = 0
            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE").ToString & String.Empty

            Dim ftrICTTFLST As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
            Dim rowICTTFLST As DataRow = dst.Tables("ICTTFLST").Select(ftrICTTFLST).FirstOrDefault

            If IsNothing(rowICTTFLST) Then
                rowSOTORDR2.Item("LINE_REPRICED") = "L"
            Else
                Dim ftrICTSTYCX As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
                Dim rowICTSTYCX As DataRow = dst.Tables.Item("ICTSTYCX").Select(ftrICTSTYCX).FirstOrDefault
                Dim STYLE_COLOR_STATUS As String = rowICTTFLST.Item("STYLE_STATUS").ToString & String.Empty
                If STYLE_COLOR_STATUS.Length = 0 Then
                    STYLE_COLOR_STATUS = "A"
                End If

                If Not IsNothing(rowICTSTYCX) Then
                    If rowICTSTYCX.Item("STYLE_COLOR_STATUS").ToString = "D" Then
                        If ORDR_DATE >= CDate(rowICTSTYCX.Item("INIT_DATE").ToString) Then
                            STYLE_COLOR_STATUS = "D"
                        End If
                    End If
                End If

                'Dim STYLE_CLASS_CODE As String = rowICTTFLST.Item("STYLE_CLASS_CODE").ToString & String.Empty
                'Dim NOTFXP As Boolean = False
                'If chkOnlyFXP.Checked Then
                '    If STYLE_CLASS_CODE <> "FALL" And STYLE_CLASS_CODE <> "XMAS" And STYLE_CLASS_CODE <> "PVC" Then
                '        NOTFXP = True
                '    End If
                'End If
                Dim REPRICE_STYLE As String = rowICTTFLST.Item("REPRICE_STYLE").ToString & String.Empty
                If REPRICE_STYLE <> "1" Then
                    rowSOTORDR2.Item("LINE_REPRICED") = "F"
                Else
                    Dim STYLE_PRICE_NEW As Double = Val(rowICTTFLST.Item("STYLE_PRICE_NEW").ToString & String.Empty)
                    Dim STYLE_PRICE_ORIG As Double = Val(rowICTTFLST.Item("STYLE_PRICE").ToString & String.Empty)
                    Dim STYLE_PRICE_ORDR As Double = Val(rowSOTORDR2.Item("STYLE_PRICE").ToString & String.Empty)
                    If STYLE_PRICE_NEW > STYLE_PRICE_ORIG Or STYLE_PRICE_NEW > STYLE_PRICE_ORDR Then
                        rowSOTORDR2.Item("LINE_REPRICED") = "I"
                    Else
                        'If rowSOTORDR2.Item("ORDR_UNIT_PRICE_MANUAL").ToString & String.Empty = "1" Then
                        '    rowSOTORDR2.Item("LINE_REPRICED") = "N"
                        'Else
                        Dim ORDR_QTY As Integer = Val(rowSOTORDR2.Item("ORDR_QTY").ToString & String.Empty)
                        Dim ORDR_QTY_OPEN As Integer = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN").ToString & String.Empty)
                        Dim ORDR_QTY_PICK As Integer = Val(rowSOTORDR2.Item("ORDR_QTY_PICK").ToString & String.Empty)
                        Dim ORDR_QTY_CANC As Integer = Val(rowSOTORDR2.Item("ORDR_QTY_CANC").ToString & String.Empty)

                        Dim ORDR_QTY_SHIP As Integer = Val(rowSOTORDR2.Item("ORDR_QTY_SHIP").ToString & String.Empty)

                        If (ORDR_QTY_SHIP > 0) Or ((ORDR_QTY - ORDR_QTY_CANC) = 0) Then
                            rowSOTORDR2.Item("LINE_REPRICED") = "S"
                        Else
                            ORDR_UNIT_PRICE = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE").ToString & String.Empty)
                            ORIG_VALUE += ORDR_QTY * ORDR_UNIT_PRICE
                            Discounts = SOCMAIN2.Price_Discounts(Me, CUST_CODE, rowARTCUST1, rowSOTORDR2.Item("STYLE_CODE"), True, , , STYLE_PRICE_NEW, True, STYLE_COLOR_STATUS)
                            'Begin
                            Dim LastGoodBreak As Integer = 0
                            'Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", rowSOTORDR2.Item("STYLE_CODE"))).FirstOrDefault
                            'Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowSOTORDR2.Item("STYLE_CODE"))
                            Dim rowICTSTYL1 As DataRow = dst.Tables("ICTTFLST").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE)).FirstOrDefault
                            For i As Integer = 0 To 3
                                If Discounts(i).DISCOUNT_QTY > 0 Then
                                    LastGoodBreak = i
                                End If
                                If ORDR_QTY >= Discounts(i).DISCOUNT_QTY Then
                                    ORDR_UNIT_PRICE_NEW = Discounts(i).DISCOUNT_PRICE
                                    Exit For
                                End If
                            Next
                            If ORDR_UNIT_PRICE_NEW = 0 Then
                                ORDR_UNIT_PRICE_NEW = Discounts(LastGoodBreak).DISCOUNT_PRICE
                            End If
                            If rowICTSTYL1.Item("STYLE_CLASS_CODE").ToString = "PVC" Then
                                If PINNED_CUST_PRICE_TIER_PVC.Length > 0 Then
                                    Select Case PINNED_CUST_PRICE_TIER_PVC
                                        Case "5C"
                                            If ORDR_UNIT_PRICE_NEW > Val(Discounts(1).DISCOUNT_PRICE) Then
                                                If Val(Discounts(1).DISCOUNT_PRICE) < ORDR_UNIT_PRICE_NEW Then
                                                    ORDR_UNIT_PRICE_NEW = Val(Discounts(1).DISCOUNT_PRICE)
                                                End If
                                            End If
                                        Case "FC"
                                            If ORDR_UNIT_PRICE_NEW > Val(Discounts(2).DISCOUNT_PRICE) Then
                                                If Val(Discounts(2).DISCOUNT_PRICE) < ORDR_UNIT_PRICE_NEW Then
                                                    ORDR_UNIT_PRICE_NEW = Val(Discounts(2).DISCOUNT_PRICE)
                                                End If
                                            End If
                                    End Select
                                End If
                            Else
                                Select Case rowARTCUST1.Item("CUST_PRICE_TIER").ToString
                                    Case "HC"
                                        If Discounts(2).DISCOUNT_PRICE < ORDR_UNIT_PRICE_NEW Then
                                            ORDR_UNIT_PRICE_NEW = Discounts(2).DISCOUNT_PRICE
                                        End If
                                    Case "FC"
                                        If Discounts(1).DISCOUNT_PRICE < ORDR_UNIT_PRICE_NEW Then
                                            ORDR_UNIT_PRICE_NEW = Discounts(1).DISCOUNT_PRICE
                                        End If
                                End Select
                            End If
                            If ORDR_UNIT_PRICE_NEW = 0 Then
                                ORDR_UNIT_PRICE_NEW = Discounts(LastGoodBreak).DISCOUNT_PRICE
                            End If
                            'End
                            'ORDR_UNIT_PRICE_NEW = 1
                            If Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE").ToString & String.Empty) < Math.Round(ORDR_UNIT_PRICE_NEW, 2) Then
                                If rowSOTORDR2.Item("ORDR_UNIT_PRICE_MANUAL").ToString & String.Empty = "1" Then
                                    rowSOTORDR2.Item("LINE_REPRICED") = "N"
                                    rowSOTORDR2.Item("CALC_UNIT_PRICE") = Math.Round(ORDR_UNIT_PRICE_NEW, 2)
                                Else
                                    rowSOTORDR2.Item("LINE_REPRICED") = "U"
                                    rowSOTORDR2.Item("CALC_UNIT_PRICE") = Math.Round(ORDR_UNIT_PRICE_NEW, 2)
                                    rowSOTORDR1.Item("ORDR_REPRICED") = "S"
                                End If
                                NEW_VALUE += ORDR_QTY * Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE").ToString & String.Empty)
                            Else
                                If Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE").ToString & String.Empty) = Math.Round(ORDR_UNIT_PRICE_NEW, 2) Then
                                    rowSOTORDR2.Item("LINE_REPRICED") = "P"
                                Else
                                    If rowSOTORDR2.Item("ORDR_UNIT_PRICE_MANUAL").ToString & String.Empty = "1" Then
                                        rowSOTORDR2.Item("LINE_REPRICED") = "W"
                                    Else
                                        rowSOTORDR2.Item("LINE_REPRICED") = "R"
                                    End If
                                End If
                                rowSOTORDR2.Item("CALC_UNIT_PRICE") = Math.Round(ORDR_UNIT_PRICE_NEW, 2)
                                rowSOTORDR2.Item("ORDR_UNIT_PRICE") = Math.Round(ORDR_UNIT_PRICE_NEW, 2)
                                rowSOTORDR2.Item("ORDR_UNIT_PRICE_CURR") = Math.Round(ORDR_UNIT_PRICE_NEW, 2)
                                NEW_VALUE += ORDR_QTY * Math.Round(ORDR_UNIT_PRICE_NEW, 2)
                            End If
                        End If
                        'End If
                    End If
                End If
            End If
        Next
        RetVal = NEW_VALUE - ORIG_VALUE
        Return RetVal
    End Function

    Private Function RePriceFEFDOrder(ByVal ORDR_NO As String, ByVal FEFDType As String) As Double
        Dim RetVal As Double
        Dim ORDR_UNIT_PRICE As Double
        Dim ORDR_UNIT_PRICE_NEW As Double
        Dim ORDR_QTY As Integer
        Dim ORIG_VALUE As Double = 0
        Dim NEW_VALUE As Double = 0

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
            Dim LINE_REPRICED As String = rowSOTORDR2.Item("LINE_REPRICED").ToString & String.Empty
            If LINE_REPRICED.Length = 0 Then
                ORDR_QTY = rowSOTORDR2.Item("ORDR_QTY")
                ORDR_UNIT_PRICE = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                ORIG_VALUE += ORDR_QTY * ORDR_UNIT_PRICE

                Dim FEFD As New FEFDPrice(Me, rowSOTORDR2.Item("STYLE_CODE").ToString & String.Empty, 1)

                If Not IsNothing(FEFD.ErrorMsg) Then
                    ORDR_UNIT_PRICE_NEW = ORDR_UNIT_PRICE
                Else
                    Select Case FEFDType
                        Case "FE"
                            ORDR_UNIT_PRICE_NEW = Math.Round(FEFD.FEPrice, 2)
                        Case "FEMIX"
                            ORDR_UNIT_PRICE_NEW = Math.Round(FEFD.FEMixPrice, 2)
                        Case "FD"
                            ORDR_UNIT_PRICE_NEW = Math.Round(FEFD.FDPrice, 2)
                        Case "FDMIX"
                            ORDR_UNIT_PRICE_NEW = Math.Round(FEFD.FDMixPrice, 2)
                        Case Else
                            ORDR_UNIT_PRICE_NEW = Math.Round(ORDR_UNIT_PRICE, 2)
                    End Select
                End If

                If rowSOTORDR2.Item("ORDR_UNIT_PRICE") > ORDR_UNIT_PRICE_NEW Then
                    rowSOTORDR2.Item("CALC_UNIT_PRICE") = ORDR_UNIT_PRICE_NEW
                    rowSOTORDR2.Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE_NEW
                    rowSOTORDR2.Item("ORDR_UNIT_PRICE_CURR") = ORDR_UNIT_PRICE_NEW
                    rowSOTORDR2.Item("LINE_REPRICED") = "R"
                    'rowSOTORDR2.Item("ORDR_UNIT_PRICE_MANUAL") = "1"
                    NEW_VALUE += ORDR_QTY * ORDR_UNIT_PRICE_NEW
                Else
                    If rowSOTORDR2.Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE_NEW Then
                        rowSOTORDR2.Item("LINE_REPRICED") = "P"
                        rowSOTORDR2.Item("CALC_UNIT_PRICE") = ORDR_UNIT_PRICE_NEW
                    Else
                        rowSOTORDR2.Item("LINE_REPRICED") = "U"
                        rowSOTORDR2.Item("CALC_UNIT_PRICE") = ORDR_UNIT_PRICE_NEW
                    End If
                    NEW_VALUE += ORDR_QTY * Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE").ToString & String.Empty)
                End If
            End If
        Next
        'RetVal = NEW_VALUE - ORIG_VALUE
        RetVal = NEW_VALUE
        Return RetVal
    End Function

    Private Sub FinalizeData()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Saving Order Data", "")
        dst.Tables.Item("SOTOTRP1").Clear()
        dst.Tables.Item("SOTOTRP2").Clear()
        Dim REC_TOT As Int64 = dst.Tables.Item("SOTORDR1").Rows.Count
        Dim REC_NOW As Int64 = 0
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
            REC_NOW += 1
            ASCMAIN1.Progress("-", ((REC_NOW / REC_TOT)).ToString("###,##0 %"))
            Dim newSOTOTRP1 As DataRow = dst.Tables.Item("SOTOTRP1").NewRow
            newSOTOTRP1.Item("ORDR_NO") = rowSOTORDR1.Item("ORDR_NO")
            newSOTOTRP1.Item("ORDR_REPRICED") = rowSOTORDR1.Item("ORDR_REPRICED")
            newSOTOTRP1.Item("ORDR_TOTAL_ORIG") = rowSOTORDR1.Item("ORDR_TOTAL_ORIG")
            newSOTOTRP1.Item("ORDR_TOTAL_NEW") = rowSOTORDR1.Item("ORDR_TOTAL_NEW")
            dst.Tables.Item("SOTOTRP1").Rows.Add(newSOTOTRP1)
            Dim ftrSOTORDR2 As String = String.Format("ORDR_NO = '{0}'", rowSOTORDR1.Item("ORDR_NO").ToString & String.Empty)
            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(ftrSOTORDR2)
                Dim newSOTOTRP2 As DataRow = dst.Tables.Item("SOTOTRP2").NewRow
                newSOTOTRP2.Item("ORDR_NO") = rowSOTORDR2.Item("ORDR_NO")
                newSOTOTRP2.Item("ORDR_LNO") = rowSOTORDR2.Item("ORDR_LNO")
                newSOTOTRP2.Item("LINE_REPRICED") = rowSOTORDR2.Item("LINE_REPRICED")
                newSOTOTRP2.Item("CALC_UNIT_PRICE") = rowSOTORDR2.Item("CALC_UNIT_PRICE")
                newSOTOTRP2.Item("ORDR_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                newSOTOTRP2.Item("ORDR_UNIT_PRICE_ORIG") = rowSOTORDR2.Item("ORDR_UNIT_PRICE_ORIG")
                newSOTOTRP2.Item("ORDR_QTY") = rowSOTORDR2.Item("ORDR_QTY")
                newSOTOTRP2.Item("ORDR_QTY_OPEN") = rowSOTORDR2.Item("ORDR_QTY_OPEN")
                newSOTOTRP2.Item("ORDR_QTY_PICK") = rowSOTORDR2.Item("ORDR_QTY_PICK")
                newSOTOTRP2.Item("ORDR_QTY_SHIP") = rowSOTORDR2.Item("ORDR_QTY_SHIP")
                newSOTOTRP2.Item("ORDR_QTY_CANC") = rowSOTORDR2.Item("ORDR_QTY_CANC")
                dst.Tables.Item("SOTOTRP2").Rows.Add(newSOTOTRP2)
            Next
        Next
        For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select()
            Dim STYLE_CODE As String = rowICTSTYL1.Item("STYLE_CODE").ToString & String.Empty
            Dim ftrICTTFLST As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
            Dim STYLE_PRICE_NEW As Double = Val(dst.Tables.Item("ICTTFLST").Select(ftrICTTFLST).FirstOrDefault.Item("STYLE_PRICE_NEW").ToString & String.Empty)
            rowICTSTYL1.Item("STYLE_PRICE") = STYLE_PRICE_NEW
        Next
        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Setup_SOTORDR2()
        If grdSOTORDR1.ActiveRow Is Nothing OrElse (Not grdSOTORDR1.ActiveRow.IsDataRow Or grdSOTORDR1.ActiveRow.IsAddRow) Then
            'grdSOTORDR2.Visible = False
        Else
            Dim dvw As DataView = DirectCast(grdSOTORDR2.DataSource, DataTable).DefaultView
            Dim ORDR_NO As String = grdSOTORDR1.ActiveRow.Cells("ORDR_NO").Value
            dvw.RowFilter = String.Format("ORDR_NO = '{0}'", ORDR_NO)
            grdSOTORDR2.Text = String.Format("Order Details For Order {0}", ORDR_NO)
            'grpSOTORDR3.Visible = True
        End If
    End Sub

#End Region

#Region "Create Emails"

    Private Sub GenerateEmails()

        Try

            For Each tableName As String In New String() {"SOTSREP1", "ARTCUSTD", "ARTCUST1_E", "SOTORDR1_E", "SOTORDR2_E", "SOTOTRP1_E", "SOTOTRP2_E"}
                dst.Tables(tableName).Clear()
            Next

            Dim s As New System.Text.StringBuilder With {.Length = 0}

            s.Length = 0
            s.AppendLine("Select DISTINCT SOTOTRP1.ORDR_NO, ARTCUST1.CUST_CODE")
            s.AppendLine(" From SOTOTRP1, SOTOTRP2, ARTCUST1, SOTORDR1")
            s.AppendLine(" Where SOTOTRP1.ORDR_NO = SOTOTRP2.ORDR_NO")
            s.AppendLine(" And SOTOTRP1. ORDR_REPRICED in ('R', 'S', 'F') ")
            s.AppendLine(" And SOTOTRP2. LINE_REPRICED in ('R', 'W')")
            s.AppendLine(" and SOTOTRP1.ORDR_NO = SOTORDR1.ORDR_NO")
            s.AppendLine(" and SOTORDR1.CUST_CODE = ARTCUST1.CUST_CODE")
            s.AppendLine(" and NVL(SOTOTRP1.EMAIL_STATUS, '1') <> 'F'")

            Dim wkTable As String = ASCMAIN1.Temp_Table(s.ToString)

            s.Length = 0
            s.AppendLine("Select * from SOTORDR1 WHERE ORDR_NO IN")
            s.AppendLine("( ")
            s.AppendLine("SELECT ORDR_NO FROM ")
            s.AppendLine(wkTable)
            s.AppendLine(" )")
            Fill_Records("SOTORDR1_E", String.Empty, True, s.ToString)

            If dst.Tables("SOTORDR1_E").Rows.Count = 0 Then
                MessageBox.Show("There are no emails to generate.", "Generate Emails", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            ASCMAIN1.Progress("Generating Emails", String.Empty)
            Me.Cursor = Cursors.WaitCursor

            s.Length = 0
            s.AppendLine("Select * from SOTORDR2 WHERE ORDR_NO IN")
            s.AppendLine("( ")
            s.AppendLine("SELECT ORDR_NO FROM ")
            s.AppendLine(wkTable)
            s.AppendLine(" )")
            Fill_Records("SOTORDR2_E", String.Empty, True, s.ToString)

            s.Length = 0
            s.AppendLine("Select * from SOTOTRP1 WHERE ORDR_NO IN")
            s.AppendLine("( ")
            s.AppendLine("SELECT ORDR_NO FROM ")
            s.AppendLine(wkTable)
            s.AppendLine(" )")
            Fill_Records("SOTOTRP1_E", String.Empty, True, s.ToString)

            s.Length = 0
            s.AppendLine("Select * from SOTOTRP2 WHERE ORDR_NO IN")
            s.AppendLine("( ")
            s.AppendLine("SELECT ORDR_NO FROM ")
            s.AppendLine(wkTable)
            s.AppendLine(" )")
            Fill_Records("SOTOTRP2_E", String.Empty, True, s.ToString)

            s.Length = 0
            s.AppendLine("Select * from ARTCUST1 WHERE CUST_CODE IN")
            s.AppendLine("( ")
            s.AppendLine("SELECT CUST_CODE FROM ")
            s.AppendLine(wkTable)
            s.AppendLine(" )")
            Fill_Records("ARTCUST1_E", String.Empty, True, s.ToString)

            s.Length = 0
            s.AppendLine("Select * from ARTCUSTD WHERE CUST_CODE IN")
            s.AppendLine("( ")
            s.AppendLine("SELECT CUST_CODE FROM ")
            s.AppendLine(wkTable)
            s.AppendLine(" )")
            Fill_Records("ARTCUSTD", String.Empty, True, s.ToString)

            Fill_Records("SOTSREP1", String.Empty, True, "Select * from SOTSREP1")

            For Each rowARTCUST1 As DataRow In dst.Tables("ARTCUST1_E").Select("", "CUST_CODE")
                Dim CUST_CODE As String = rowARTCUST1.Item("CUST_CODE")
                ASCMAIN1.Progress("-", CUST_CODE)

                Dim fileAttachment As String = String.Empty
                If EmailCustomer(CUST_CODE, fileAttachment) Then
                    For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1_E").Select("CUST_CODE = '" & CUST_CODE & "'")
                        Dim sql As String = "Update SOTOTRP1 set EMAIL_STATUS = 'F' where ORDR_NO = :PARM1"
                        ASCDATA1.ExecuteSQL(sql, "V", rowSOTORDR1.Item("ORDR_NO"))

                        If fileAttachment.Length > 0 Then
                            dst.Tables("ASTATTA2").Rows.Clear()
                            ENTITY.TABLE_NAME = "SOTORDR1"
                            ENTITY.COLUMN_NAME = "ORDR_NO"
                            ENTITY.CODE_VALUE = rowSOTORDR1.Item("ORDR_NO") & String.Empty

                            MyBase.Attach_File(fileAttachment, "Emailed Tarrif Changes")

                            TAC.TACMAIN1.Record_Event("SOTORDR1", _
                                  rowSOTORDR1.Item("ORDR_NO"), _
                                  Now, _
                                  ASCMAIN1.USER_ID, _
                                  "TARRIF", _
                                  "Order Tariff Repricing Letter generated on: " & DateTime.Now)
                        End If
                    Next
                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Error generating emails: " & ex.Message, "Generate Emails", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Function EmailCustomer(ByVal CUST_CODE As String, ByRef fileAttachment As String) As Boolean

        Dim attachFileName As String = String.Empty
        Dim customerEmailFound As Boolean = False

        Try

            If ASCMAIN1.CLIENT <> "RGI" Then
                Return False
            End If

            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1_E").Select("CUST_CODE = '" & CUST_CODE & "'", "")(0)
            Dim SREP_CODE As String = String.Empty
            SREP_CODE = rowSOTORDR1.Item("SREP_CODE") & String.Empty

            Dim emailToList As String = String.Empty

            ' See if the customer receives an acknowledgment
            Dim rowSOTSREP1 As DataRow = dst.Tables("SOTSREP1").Rows.Find(SREP_CODE)
            Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)

            If rowARTCUST1 Is Nothing Then
                Return False
            End If

            If rowSOTSREP1 IsNot Nothing AndAlso rowSOTSREP1.Item("SREP_EMAIL") & String.Empty <> String.Empty Then
                emailToList = rowSOTSREP1.Item("SREP_EMAIL") & String.Empty
                customerEmailFound = True
            End If

            If (rowARTCUST1.Item("CUST_EMAIL") & String.Empty).ToString.Trim.Length > 0 Then
                emailToList &= ";" & (rowARTCUST1.Item("CUST_EMAIL") & String.Empty).ToString.Trim
                customerEmailFound = True
            End If

            S.Length = 0
            S.AppendLine("CUST_CODE = '" & CUST_CODE & "'")
            If dst.Tables("ARTCUSTD").Select(S.ToString, "").Length > 0 Then
                S.AppendLine("  and CONTACT_PRIMARY = '1'")
                If dst.Tables("ARTCUSTD").Select(S.ToString).Length > 0 Then
                    customerEmailFound = True
                    For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(S.ToString)
                        emailToList &= ";" & (rowARTCUSTD.Item("CONTACT_EMAIL") & String.Empty).ToString.Trim
                    Next
                Else
                    emailToList &= ";" & (dst.Tables("ARTCUSTD").Rows(0).Item("CONTACT_EMAIL") & String.Empty).ToString.Trim
                    customerEmailFound = True
                End If
            End If

            ' remove double semi-colons
            While emailToList.Contains(" ")
                emailToList = emailToList.Replace(" ", "")
            End While
            emailToList = emailToList.Replace(",", ";")

            While emailToList.Contains(";;")
                emailToList = emailToList.Replace(";;", ";")
            End While

            ' should be at least 5 characters
            If emailToList.Replace(";", "").Trim.Length < 5 Then
                ' Return False
            End If

            attachFileName = CUST_CODE & "_" & rowARTCUST1.Item("CUST_NAME")

            For Each invalidChar As String In New String() {"\", "/", ":", "*", "?", "<", ">", "|", ".", ",", "'"}
                attachFileName = attachFileName.Replace(invalidChar, "")
            Next
            attachFileName = attachFileName.Replace(" ", "_")

            If ASCMAIN1.Running_in_VS Then
                Stop
                emailToList = "ewz@absolution.com;rich@regency-rib.com"
            End If

            ' Concatentate and process all email addresses
            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            For Each emailAddress As String In (emailToList).ToString.Split(";")
                emailAddress = emailAddress.Trim
                If emailAddress.Length > 5 AndAlso Not EMAIL_ADDRESSs.Keys.Contains(emailAddress) Then
                    EMAIL_ADDRESSs.Add(emailAddress, emailAddress)
                End If
            Next

            ' Build the Excel file
            SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey)

            Dim numberStyle As New CellStyle
            numberStyle.HorizontalAlignment = HorizontalAlignmentStyle.Right
            numberStyle.VerticalAlignment = VerticalAlignmentStyle.Bottom
            numberStyle.NumberFormat = "#,##0"

            Dim decimalStyle As New CellStyle
            decimalStyle.HorizontalAlignment = HorizontalAlignmentStyle.Right
            decimalStyle.VerticalAlignment = VerticalAlignmentStyle.Bottom
            decimalStyle.NumberFormat = "#,##0.00"

            Dim decimalStyleRedForeground As New CellStyle
            decimalStyleRedForeground.HorizontalAlignment = HorizontalAlignmentStyle.Right
            decimalStyleRedForeground.VerticalAlignment = VerticalAlignmentStyle.Bottom
            decimalStyleRedForeground.Font.Color = Drawing.Color.Red
            decimalStyleRedForeground.NumberFormat = "#,##0.00"

            Dim textStyle As New CellStyle
            textStyle.HorizontalAlignment = HorizontalAlignmentStyle.Left
            textStyle.VerticalAlignment = VerticalAlignmentStyle.Bottom

            Dim workbook As New ExcelFile()
            Dim worksheet As ExcelWorksheet = workbook.Worksheets.Add("Sales Orders")

            Dim pixelLength As Int32 = 400

            ' Header Infotmation
            worksheet.Columns(0).Width = 12 * pixelLength
            worksheet.Columns(0).Style = textStyle
            ' Style
            worksheet.Columns(1).Width = 8 * pixelLength
            worksheet.Columns(1).Style = textStyle
            ' Color
            worksheet.Columns(2).Width = 8 * pixelLength
            worksheet.Columns(2).Style = textStyle
            'Description
            worksheet.Columns(3).Width = 40 * pixelLength
            worksheet.Columns(3).Style = textStyle
            ' Quantity
            worksheet.Columns(4).Width = 6 * pixelLength
            worksheet.Columns(4).Style = numberStyle
            'Orig Price
            worksheet.Columns(5).Width = 10 * pixelLength
            worksheet.Columns(5).Style = decimalStyle
            ' New price
            worksheet.Columns(6).Width = 10 * pixelLength
            worksheet.Columns(6).Style = decimalStyle

            Dim sheetRow As Int16 = 0

            worksheet.Cells(sheetRow, 0).Value = CUST_CODE & " - " & rowARTCUST1.Item("CUST_NAME") & String.Empty
            sheetRow += 1

            worksheet.Cells(sheetRow, 0).Value = rowARTCUST1.Item("CUST_ADDR1") & String.Empty
            sheetRow += 1

            If rowARTCUST1.Item("CUST_ADDR2") & String.Empty <> String.Empty Then
                worksheet.Cells(sheetRow, 0).Value = rowARTCUST1.Item("CUST_ADDR2") & String.Empty
                sheetRow += 1
            End If

            worksheet.Cells(sheetRow, 0).Value = rowARTCUST1.Item("CUST_CITY") & ", " & rowARTCUST1.Item("CUST_STATE") & "  " & rowARTCUST1.Item("CUST_ZIP_CODE")
            sheetRow += 2

            worksheet.Cells(sheetRow, 0).Value = "The following sales orders contain items with reduced prices because of Reduced Tariffs."
            worksheet.Cells(sheetRow, 0).Style.Font.Color = Drawing.Color.Red
            'worksheet.Cells(sheetRow, 0).Style.Font.Weight = ExcelFont.BoldWeight
            sheetRow += 2

            For Each rowSOTORDR1 In dst.Tables("SOTORDR1_E").Select("CUST_CODE = '" & CUST_CODE & "'", "ORDR_NO,ORDR_SHIP_DATE,CUST_STORE_NO")
                Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
                Dim SHIP_DATE As String = rowSOTORDR1.Item("ORDR_SHIP_DATE")
                Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO")
                Dim CUST_STORE_NAME As String = rowSOTORDR1.Item("CUST_STORE_NAME")

                Dim rowSOTOTRP1 As DataRow = dst.Tables("SOTOTRP1_E").Rows.Find(ORDR_NO)

                ' Sales Order Header
                worksheet.Cells(sheetRow, 0).Value = "Order No."
                worksheet.Cells(sheetRow, 1).Value = ORDR_NO
                worksheet.Cells(sheetRow, 1).Style = textStyle
                For ictr As Int16 = 0 To 6
                    worksheet.Cells(sheetRow, ictr).Style.FillPattern.SetPattern(FillPatternStyle.Solid, Drawing.Color.Yellow, Drawing.Color.Yellow)
                Next

                sheetRow += 1

                worksheet.Cells(sheetRow, 0).Value = "Ship Date"
                worksheet.Cells(sheetRow, 1).Value = SHIP_DATE
                worksheet.Cells(sheetRow, 1).Style = textStyle
                sheetRow += 1

                worksheet.Cells(sheetRow, 0).Value = "Ship To"
                worksheet.Cells(sheetRow, 1).Value = CUST_STORE_NO & " - " & CUST_STORE_NAME
                worksheet.Cells(sheetRow, 1).Style = textStyle
                sheetRow += 1


                If rowSOTOTRP1 IsNot Nothing Then
                    worksheet.Cells(sheetRow, 0).Value = "Orig Order Total"
                    worksheet.Cells(sheetRow, 1).Value = Val(rowSOTOTRP1.Item("ORDR_TOTAL_ORIG") & String.Empty).ToString("#,##0.00")
                    worksheet.Cells(sheetRow, 1).Style = decimalStyle
                    sheetRow += 1

                    worksheet.Cells(sheetRow, 0).Value = "New Order Total"
                    worksheet.Cells(sheetRow, 1).Value = Val(rowSOTOTRP1.Item("ORDR_TOTAL_NEW") & String.Empty).ToString("#,##0.00")
                    worksheet.Cells(sheetRow, 1).Style = decimalStyleRedForeground
                    sheetRow += 1
                End If

                ' Sales Order Details
                sheetRow += 1

                worksheet.Cells(sheetRow, 1).Value = "Style"
                worksheet.Cells(sheetRow, 2).Value = "Color"
                worksheet.Cells(sheetRow, 3).Value = "Description"
                worksheet.Cells(sheetRow, 4).Value = "Quantity"
                worksheet.Cells(sheetRow, 5).Value = "Orig Unit Price"
                worksheet.Cells(sheetRow, 6).Value = "New Unit Price"
                sheetRow += 1

                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2_E").Select("ORDR_NO = '" & ORDR_NO & "' AND ISNULL(ORDR_QTY_OPEN, 0) + ISNULL(ORDR_QTY_PICK, 0) > 0", "ORDR_NO, ORDR_LNO")
                    worksheet.Cells(sheetRow, 1).Value = rowSOTORDR2.Item("STYLE_CODE")
                    worksheet.Cells(sheetRow, 2).Value = rowSOTORDR2.Item("COLOR_CODE")
                    worksheet.Cells(sheetRow, 3).Value = rowSOTORDR2.Item("STYLE_DESC")
                    worksheet.Cells(sheetRow, 4).Value = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & String.Empty) + Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty)
                    worksheet.Cells(sheetRow, 4).Style = numberStyle

                    Dim rowSOTOTRP2 As DataRow = dst.Tables("SOTOTRP2_E").Rows.Find(New Object() {ORDR_NO, rowSOTORDR2.Item("ORDR_LNO")})
                    If rowSOTOTRP2 IsNot Nothing Then
                        worksheet.Cells(sheetRow, 5).Value = Val(rowSOTOTRP2.Item("ORDR_UNIT_PRICE_ORIG") & String.Empty)
                        worksheet.Cells(sheetRow, 5).Style = decimalStyle
                        worksheet.Cells(sheetRow, 6).Value = Val(rowSOTOTRP2.Item("ORDR_UNIT_PRICE") & String.Empty)
                        worksheet.Cells(sheetRow, 6).Style = decimalStyle

                        If Val(rowSOTOTRP2.Item("ORDR_UNIT_PRICE_ORIG") & String.Empty) > Val(rowSOTOTRP2.Item("ORDR_UNIT_PRICE") & String.Empty) Then
                            worksheet.Cells(sheetRow, 6).Style = decimalStyleRedForeground
                        End If

                    End If
                    sheetRow += 1
                Next

                sheetRow += 2
            Next

            fileAttachment = ASCMAIN1.Folders("Archive")
            If Not fileAttachment.EndsWith("\") Then fileAttachment &= "\"

            fileAttachment &= attachFileName & ".xls"
            workbook.Save(fileAttachment)
            workbook = Nothing

            Dim SUBJECT As String = String.Empty
            Dim SEND_NO As String = String.Empty

            ' Need to attach the letter to the sales order when we do no have an email address.
            If emailToList.Replace(";", "").Trim.Length < 5 OrElse EMAIL_ADDRESSs.Count = 0 Then
                Return False
            End If

            Dim dictATTACHMENTs As New Dictionary(Of String, String)
            dictATTACHMENTs.Add(fileAttachment, fileAttachment)

            SUBJECT = "Regency International Tariff Repricing"
            Dim body As String = "Thank you very much for your Fall/Holiday order." & Environment.NewLine
            body &= "We have adjusted the line item costing of your order(s) to reflect NO tariff pricing.  Certain items still carry a tariff or are not from China, and remain unchanged." & Environment.NewLine
            body &= "Please see your new order(s) copy." & Environment.NewLine
            body &= "With our sincerest hope for a strong 4th quarter and a comfortable resolution to this pandemic."

            SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
                  (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, dictATTACHMENTs, _
                    SUBJECT, "ORDRTARF", True, False, CUST_CODE, CUST_CODE, "Customer", body)

            Return (SEND_NO & String.Empty).Length > 0

        Catch ex As Exception
            Return False
        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

    End Function

#End Region

End Class