
Public Class WBFSDISC
    Dim InquiryOnly As Boolean = False
    Dim SQL As New Text.StringBuilder With {.Length = 0}

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        With dst

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM WBTSDISH")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTSDISH", "**", 0, True)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM WBTSDIS1")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTSDIS1", "**", 0, True)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM WBTSDIS2")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTSDIS2", "**", 0, True)

        End With

        grdWBTSDIS1.DataSource = dst.Tables("WBTSDIS1")
        grdWBTSDIS2.DataSource = dst.Tables("WBTSDIS2")

        Create_Summary(grdWBTSDIS1, "STYLE_CODE", "Count", "", "###,##0")

        'Sort_grdColumns(grdSOTORDRX, "ORDR_DATE, ORDR_GROUP_NO, ORDR_NO".ToLower(), False)

        'grdSOTRUSSE.DisplayLayout.UseFixedHeaders = True
        'With grdSOTRUSSE.DisplayLayout.Bands(0)
        '    For Each COLUMN_NAME As String In New String() {"ORDR_NO", "ORDR_LNO", "STYLE_CODE", "COLOR_CODE", "ORDR_QTY", "ORDR_UNIT_PRICE"}
        '        .Columns(COLUMN_NAME).Header.Fixed = True
        '    Next
        'End With

        tab.Visible = False
        'grdSOTORDRX.Parent = tab.Parent

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit"
                Dim x As Boolean = False
                If x = False Then
                    EMsg &= vbCr & "Some Kind Of Error."
                End If
                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ICTSTYL1", Absx1.txtFor("STYLE_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Cancel"

            Case "Update"

            Case "Done"
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
            Case "Edit"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)
            Case "Update"
                Call Update_Record()
                Call Mode_Settings(False)
            Case "Cancel", "Done"
                Call Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Done").Settings.Enabled = DefaultableBoolean.True
                .Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = DefaultableBoolean.True
            End With
        End If
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            UltraExplorerBar1.Groups("Style Image").Visible = True
        Else
            UltraExplorerBar1.Groups("Style Image").Visible = False
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        'grdSOTORDRX.Visible = Not tf
        'With grdSOTRUSSE.DisplayLayout.Override
        '    .AllowAddNew = UltraWinGrid.AllowAddNew.No
        '    .AllowDelete = DefaultableBoolean.False
        '    .AllowUpdate = DefaultableBoolean.True
        'End With
        'For i As Integer = 0 To grdSOTRUSSE.DisplayLayout.Bands(0).Columns.Count - 1
        '    grdSOTRUSSE.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        'Next i
        'For Each COLNAME As String In New String() {"NEW_QTY", "NEW_UPC", "NEW_SKU", "NEW_COLOR_CODE", "NEW_ORDR_UNIT_PRICE"}
        '    grdSOTRUSSE.DisplayLayout.Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
        'Next
        'For Each COLNAME As String In New String() {"NEW_QTY", "NEW_UPC", "NEW_SKU", "NEW_COLOR_CODE", "NEW_ORDR_UNIT_PRICE"}
        '    grdSOTRUSSE.DisplayLayout.Bands(0).Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        'Next

        'If Not ScreenMode Then
        '    RefreshSOTORDRX()
        'End If
    End Sub

    Sub Clear_Record()
        'dst.Tables("ICTSTYL1").Rows.Clear()
    End Sub

    Sub Load_Record()
        Call Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        'Call Fill_Records("ARTCUST1", Absx1.txtFor("CUST_CODE").Text, True)
        'Call Fill_Records("ARTCUST2", Absx1.txtFor("CUST_CODE").Text, True)

        EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
        'Call BeginTrans()
        'For Each TABLE_NAME As String In {"SOTORDR1", "SOTORDR2", "SOTORDR5"}
        '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where ORDR_NO = '" & ORDR_NO & "'"
        '    ASCDATA1.ExecuteSQL()
        '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & "_L where ORDR_NO = '" & ORDR_NO & "'"
        '    ASCDATA1.ExecuteSQL()
        'Next
        'Call CommitTrans("Order / Quote Deleted")
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        'TODO: Remove this stop before going live.
        Stop
        Call CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
        'Print_Report_Begin()
        'Generate_Report("SORORDRO")
        'Print_Report_End()
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdSOTORDRX, "SSB", "Show Filter", "Show GroupBox")
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

        Select Case e.SourceControl.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        End If

        Select Case e.SourceControl.Name
            'Case "grdSOTORDR1"
            '    If Not InquiryOnly Then
            '        e.Tool.ToolbarsManager.Tools("Edit Ship To").SharedProps.Visible = True
            '    End If
        End Select
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
            'Case "Edit Ship To"
            '    If Not InquiryOnly Then
            '        MsgBox("Edit Ship To Feature Coming Soon", MsgBoxStyle.Exclamation, "Waiting For Feature")
            '    End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("New", e)
                End If
            Case "PYMT_BATCH_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Edit", e)
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
                'FillStyle()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PYMT_BATCH_NO"
                Call Click_Command("Edit")
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        If IsLoading Then
            Exit Sub
        Else
            MyBase.txt_ValueChanged(sender, e)
        End If
    End Sub

    Private Sub btnLoadData_Click(sender As Object, e As EventArgs) Handles btnLoadData.Click
        If txtIMPORT_TABLE.Text.Length = 0 Then
            MsgBox("You Must Provide A Valid Import Table!", vbOKOnly, "Import")
        Else
            Me.Cursor = Cursors.WaitCursor
            dst.Tables.Item("WBTSDISH").Clear()
            dst.Tables.Item("WBTSDIS1").Clear()
            dst.Tables.Item("WBTSDIS2").Clear()
            Dim DISC_BATCH As String = ASCMAIN1.Next_Control_No("WBTSDISH.DISC_BATCH")
            Dim rowWBTSDISH As DataRow = dst.Tables.Item("WBTSDISH").NewRow
            rowWBTSDISH.Item("DISC_BATCH") = DISC_BATCH
            txtDISC_BATCH.Text = DISC_BATCH
            rowWBTSDISH.Item("DISC_DATE") = CDate(Now().ToShortDateString)
            rowWBTSDISH.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
            txtSESSION_NO.Text = ASCMAIN1.SESSION_NO
            If chkDNR.Checked Then
                rowWBTSDISH.Item("STATUS_NEW") = "N"
            Else
                rowWBTSDISH.Item("STATUS_NEW") = "D"
            End If
            SQL.Length = 0
            SQL.AppendLine("SELECT COUNT(*) FROM ICTSTYL1 WHERE STYLE_STATUS = 'A'")
            ASCMAIN1.sql = SQL.ToString()
            rowWBTSDISH.Item("A_COUNT_S") = Val(ASCDATA1.GetDataValue)

            SQL.Length = 0
            SQL.AppendLine("SELECT COUNT(*) FROM ICTSTYL1 WHERE STYLE_STATUS = 'N'")
            ASCMAIN1.sql = SQL.ToString()
            rowWBTSDISH.Item("N_COUNT_S") = Val(ASCDATA1.GetDataValue)

            SQL.Length = 0
            SQL.AppendLine("SELECT COUNT(*) FROM ICTSTYL1 WHERE STYLE_STATUS = 'D'")
            ASCMAIN1.sql = SQL.ToString()
            rowWBTSDISH.Item("D_COUNT_S") = Val(ASCDATA1.GetDataValue)

            SQL.Length = 0
            SQL.AppendLine("SELECT COUNT(*) FROM ICTSTYC1 WHERE STYLE_COLOR_STATUS = 'A'")
            ASCMAIN1.sql = SQL.ToString()
            rowWBTSDISH.Item("A_COUNT_C") = Val(ASCDATA1.GetDataValue)

            SQL.Length = 0
            SQL.AppendLine("SELECT COUNT(*) FROM ICTSTYC1 WHERE STYLE_COLOR_STATUS = 'N'")
            ASCMAIN1.sql = SQL.ToString()
            rowWBTSDISH.Item("N_COUNT_C") = Val(ASCDATA1.GetDataValue)

            SQL.Length = 0
            SQL.AppendLine("SELECT COUNT(*) FROM ICTSTYC1 WHERE STYLE_COLOR_STATUS = 'D'")
            ASCMAIN1.sql = SQL.ToString()
            rowWBTSDISH.Item("D_COUNT_C") = Val(ASCDATA1.GetDataValue)

            rowWBTSDISH.Item("XNO") = ASCMAIN1.ActiveForm.XNO
            txtXNO.Text = ASCMAIN1.ActiveForm.XNO
            dst.Tables.Item("WBTSDISH").Rows.Add(rowWBTSDISH)

            SQL.Length = 0
            SQL.AppendLine(String.Format("SELECT DISTINCT STYLE_CODE, COLOR_CODE FROM {0}", txtIMPORT_TABLE.Text))
            Dim tblIMPORT As DataTable = ASCDATA1.GetDataTable(SQL.ToString(), String.Empty)
            For Each rowIMPORT As DataRow In tblIMPORT.Rows
                Dim STYLE_CODE As String = rowIMPORT.Item("STYLE_CODE").ToString & String.Empty
                Dim COLOR_CODE As String = rowIMPORT.Item("COLOR_CODE").ToString & String.Empty
                If dst.Tables.Item("WBTSDIS1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE)).Count = 0 Then
                    Dim rowWBTSDIS1 As DataRow = dst.Tables.Item("WBTSDIS1").NewRow
                    rowWBTSDIS1.Item("DISC_BATCH") = DISC_BATCH
                    rowWBTSDIS1.Item("STYLE_CODE") = STYLE_CODE
                    SQL.Length = 0
                    SQL.AppendLine(String.Format("SELECT STYLE_STATUS FROM ICTSTYL1 WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                    ASCMAIN1.sql = SQL.ToString()
                    rowWBTSDIS1.Item("STATUS_ORIG") = ASCDATA1.GetDataValue
                    rowWBTSDIS1.Item("STATUS_NEW") = "X"
                    dst.Tables.Item("WBTSDIS1").Rows.Add(rowWBTSDIS1)
                End If
                Dim rowWBTSDIS2 As DataRow = dst.Tables.Item("WBTSDIS2").NewRow
                rowWBTSDIS2.Item("DISC_BATCH") = DISC_BATCH
                rowWBTSDIS2.Item("STYLE_CODE") = STYLE_CODE
                rowWBTSDIS2.Item("COLOR_CODE") = COLOR_CODE
                SQL.Length = 0
                SQL.AppendLine("SELECT STYLE_COLOR_STATUS")
                SQL.AppendLine("FROM ICTSTYC1")
                SQL.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                SQL.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
                ASCMAIN1.sql = SQL.ToString()
                rowWBTSDIS2.Item("STATUS_ORIG") = ASCDATA1.GetDataValue
                If chkDNR.Checked Then
                    rowWBTSDIS2.Item("STATUS_NEW") = "N"
                Else
                    rowWBTSDIS2.Item("STATUS_NEW") = "D"
                End If
                dst.Tables.Item("WBTSDIS2").Rows.Add(rowWBTSDIS2)

                UltraExplorerBar1.Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.True
            Next
            MsgBox("Import Complete!", vbOKOnly, "Done")
        End If
        Me.Cursor = Cursors.Default
    End Sub

#End Region

#Region "Custom Methods"

#End Region

End Class