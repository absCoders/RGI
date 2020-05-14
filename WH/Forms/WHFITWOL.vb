
Public Class WHFITWOL
    Dim InquiryOnly As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        With dst

            SQLs.Length = 0
            SQLs.AppendLine("SELECT")
            SQLs.AppendLine("S1.RANGE_STYLE_CODE AS STYLE_CODE,")
            SQLs.AppendLine("S1.RANGE_STYLE_DESC AS STYLE_DESC,")
            SQLs.AppendLine("G1.GTIN_UPC_CODE,")
            SQLs.AppendLine("G1.GTIN_CODE,")
            SQLs.AppendLine("36 GTIN_QTY")
            SQLs.AppendLine("FROM ICVLUPC1 L1, ICTRSTY1 S1, ICTGTINT G1")
            SQLs.AppendLine("WHERE L1.STYLE_CODE = S1.RANGE_STYLE_CODE")
            SQLs.AppendLine("AND L1.UPC_CODE = G1.GTIN_UPC_CODE")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WHTITWOX", "**", 0, False, "V", 0)
            ASCMAIN1.sql = SQLs.ToString()
            SQLs.AppendLine("AND L1.STYLE_CODE = :PARM1")
            Create_TDA(.Tables.Add, "WHFITWOL", "**", 0, False, "V", 0)
        End With

        grdWHTITWOX.DataSource = dst.Tables("WHTITWOX")

        'Create_Summary(grdWHTITWOX, "NEW_QTY", "Sum", "", "###,##0")

        'Sort_grdColumns(grdWHTITWOX, "ORDR_DATE, ORDR_GROUP_NO, ORDR_NO".ToLower(), False)

        tab.Visible = False
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
            Case "Print"
                If grdWHTITWOX.Selected.Rows.Count <> 1 Then
                    EMsg &= vbCr & "You Must Select One And Only One Style To Create PDF For."
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
            Case "Edit"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)
            Case "Update"
                Call Update_Record()
                Call Mode_Settings(False)
            Case "Cancel", "Done"
                Call Mode_Settings(False)
            Case "Print"
                Call FillSelectedStyle()
                Call Print_Record()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Edit").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Done").Visible = Not ScreenMode

                .Groups("Screen Control").Items("Update").Visible = ScreenMode
                .Groups("Screen Control").Items("Cancel").Visible = ScreenMode

            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        'grdSOTORDRX.Visible = Not tf
        With grdWHTITWOX.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With
        For i As Integer = 0 To grdWHTITWOX.DisplayLayout.Bands(0).Columns.Count - 1
            grdWHTITWOX.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i
        For Each COLNAME As String In New String() {"GTIN_QTY"}
            grdWHTITWOX.DisplayLayout.Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
            grdWHTITWOX.DisplayLayout.Bands(0).Columns(COLNAME).Format = "###,##0"
        Next
        For Each COLNAME As String In New String() {"GTIN_QTY"}
            grdWHTITWOX.DisplayLayout.Bands(0).Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        Next
    End Sub

    Sub Clear_Record()
        dst.Tables("ICTSTYL1").Rows.Clear()
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

    Private Sub Print_Record()
        Print_Report_Begin()
        'Dim FILENAME As String = ASCMAIN1.Folders("Archive") & "Cletters\" & RPTNO & ".pdf"
        'Generate_Report("PMRINVP1", "Invoices to be emailed", , , "PDF", RPTNO, False)
        Dim STYLE_CODE As String = grdWHTITWOX.Selected.Rows(0).Cells("STYLE_CODE").Text & "_36"

        Dim FILENAME_temp As String = ASCMAIN1.Folders("Temp") & STYLE_CODE & ".pdf"
        'Dim FILENAME As String = "C:\KMART\" & STYLE_CODE & ".pdf"
        If IO.File.Exists(FILENAME_temp) Then
            IO.File.Delete(FILENAME_temp)
        End If
        'If IO.File.Exists(FILENAME) Then
        '    IO.File.Delete(FILENAME)
        'End If
        Dim rptName As String = Generate_Report("WHRITWOL", "", "", "", "PDF", STYLE_CODE, False)
        Show_Document(FILENAME_temp)
        'My.Computer.FileSystem.CopyFile(FILENAME_temp, FILENAME, True)

        'Generate_Report("WHRITWOL", "", "", "", "PDF")
        Print_Report_End()
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

#End Region

#Region "Form Controls"
    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        Dim defaultQty As Int64
        If IsNumeric(txtQTY.Text) Then
            defaultQty = Val(txtQTY.Text)
        Else
            MsgBox("Default Qty Is Not Numeric", vbExclamation, "Qty")
            Exit Sub
        End If

        dst.Tables.Item("WHTITWOX").Clear()
        Dim sql As String = makeFillSQL("WHTITWOX")
        If sql.Length = 0 Then
            Exit Sub
        End If

        Dim tblSELECTIONS = ASCDATA1.GetDataTable(sql, String.Empty)
        For Each rowSELECTIONS As DataRow In tblSELECTIONS.Rows
            Dim newWHTITWOX As DataRow = dst.Tables.Item("WHTITWOX").NewRow
            newWHTITWOX.Item("STYLE_CODE") = rowSELECTIONS.Item("STYLE_CODE").ToString
            newWHTITWOX.Item("STYLE_DESC") = rowSELECTIONS.Item("STYLE_DESC").ToString
            newWHTITWOX.Item("GTIN_UPC_CODE") = rowSELECTIONS.Item("GTIN_UPC_CODE").ToString
            newWHTITWOX.Item("GTIN_CODE") = rowSELECTIONS.Item("GTIN_CODE").ToString
            newWHTITWOX.Item("GTIN_QTY") = defaultQty
            dst.Tables.Item("WHTITWOX").Rows.Add(newWHTITWOX)
        Next

    End Sub

    Private Sub btnSelectList_Click(sender As Object, e As EventArgs) Handles btnSelectList.Click
        Dim S As New Text.StringBuilder With {.Length = 0}

        Dim defaultQty As Int64
        If IsNumeric(txtQTY.Text) Then
            defaultQty = Val(txtQTY.Text)
        Else
            MsgBox("Default Qty Is Not Numeric", vbExclamation, "Qty")
            Exit Sub
        End If

        If txtCUST_CODE.Text.Length = 0 Then
            MsgBox("You Must First Select A Customer", vbOKOnly, "Missing Customer")
        Else
            If chkRanges.Checked Then
                S.AppendLine("SELECT")
                S.AppendLine("S1.RANGE_STYLE_CODE AS STYLE_CODE,")
                S.AppendLine("S1.RANGE_STYLE_DESC AS STYLE_DESC,")
                S.AppendLine("G1.GTIN_UPC_CODE,")
                S.AppendLine("G1.GTIN_CODE")
                S.AppendLine("FROM ICVLUPC1 L1, ICTRSTY1 S1, ICTGTINT G1")
                S.AppendLine("WHERE L1.STYLE_CODE = S1.RANGE_STYLE_CODE")
                S.AppendLine("AND L1.UPC_CODE = G1.GTIN_UPC_CODE")
                S.AppendLine(String.Format("AND S1.CUST_CODE = '{0}'", txtCUST_CODE.Text))
            Else
                S.AppendLine("SELECT")
                S.AppendLine("S1.STYLE_CODE AS STYLE_CODE,")
                S.AppendLine("S1.STYLE_DESC AS STYLE_DESC,")
                S.AppendLine("G1.GTIN_UPC_CODE,")
                S.AppendLine("G1.GTIN_CODE")
                S.AppendLine("FROM ICVLUPC1 L1, ICTSTYL1 S1, ICTGTINT G1")
                S.AppendLine("WHERE L1.STYLE_CODE = S1.STYLE_CODE")
                S.AppendLine("AND L1.UPC_CODE = G1.GTIN_UPC_CODE")
                S.AppendLine(String.Format("AND S1.CUST_CODE = '{0}'", txtCUST_CODE.Text))
            End If
        End If
        With ASCMAIN1.CodeSelector
            .SQL = S.ToString
            .MultipleSelections = True
            .PreviouslySelectedCodes0 = ""
            .Caption = "Select Styles"
            .TABLE_NAME = ""
            .VIEW_NAME = ""
            .VIEW_DESC = ""
            .COLUMN_NAME = ""
            .COLUMN_PREKEYs = New Dictionary(Of String, String)
        End With
        Dim F As New ASFCODE1
        F.ShowDialog()
        If ASCMAIN1.CodeSelector.Selections <> 0 Then
            For Each rowSELECTIONS As DataRow In ASCMAIN1.CodeSelector.SelectedRows
                Dim newWHTITWOX As DataRow = dst.Tables.Item("WHTITWOX").NewRow
                newWHTITWOX.Item("STYLE_CODE") = rowSELECTIONS.Item("STYLE_CODE").ToString
                newWHTITWOX.Item("STYLE_DESC") = rowSELECTIONS.Item("STYLE_DESC").ToString
                newWHTITWOX.Item("GTIN_UPC_CODE") = rowSELECTIONS.Item("GTIN_UPC_CODE").ToString
                newWHTITWOX.Item("GTIN_CODE") = rowSELECTIONS.Item("GTIN_CODE").ToString
                newWHTITWOX.Item("GTIN_QTY") = defaultQty
                dst.Tables.Item("WHTITWOX").Rows.Add(newWHTITWOX)
            Next
        End If
    End Sub

    Private Sub chkRanges_CheckedChanged(sender As Object, e As EventArgs) Handles chkRanges.CheckedChanged
        If chkRanges.Checked Then
            lblStyleBegin.Text = "Beginning Range"
            lblStyleEnd.Text = "Ending Range"
            txtRangeStyleBegin.Visible = True
            txtRangeStyleEnd.Visible = True
            txtStyleBegin.Visible = False
            txtStyleEnd.Visible = False
        Else
            lblStyleBegin.Text = "Beginning Style"
            lblStyleEnd.Text = "Ending Style"
            txtRangeStyleBegin.Visible = False
            txtRangeStyleEnd.Visible = False
            txtStyleBegin.Visible = True
            txtStyleEnd.Visible = True
        End If
    End Sub

    Private Sub txtStyleBegin_ValueChanged(sender As Object, e As EventArgs) Handles txtStyleBegin.ValueChanged

    End Sub
#End Region

#Region "Custom Methods"
    Private Sub FillSelectedStyle()
        If grdWHTITWOX.Selected.Rows.Count = 1 Then
            Dim STYLE_CODE As String = grdWHTITWOX.Selected.Rows(0).Cells("STYLE_CODE").Text
            dst.Tables.Item("WHFITWOL").Clear()
            Dim sql As String = makeFillSQL("WHFITWOL")
            If sql.Length > 0 Then
                For i As Int64 = 1 To 6
                    Fill_Records("WHFITWOL", STYLE_CODE, False, sql)
                Next
            End If
        Else
            MsgBox("You Have To Select A Row", vbCancel, "Selection")
        End If

    End Sub

    Private Function makeFillSQL(ByVal TABLENAME As String) As String
        Dim RetVal As String = ""
        Dim custCode As String = txtCUST_CODE.Text
        Dim begStyle As String = ""
        Dim endStyle As String = ""
        If chkRanges.Checked Then
            begStyle = txtRangeStyleBegin.Text
            endStyle = txtRangeStyleEnd.Text
        Else
            begStyle = txtStyleBegin.Text
            endStyle = txtStyleEnd.Text
        End If
        If begStyle.Length = 0 Or begStyle.Length = 0 Then
            MsgBox("You Must Supply A Beginning And Ending Style/Range.", vbExclamation, "Style/Range Selection")
        End If
        Dim GTIN_QTY As Int64 = 0
        If grdWHTITWOX.Selected.Rows.Count = 1 Then
            GTIN_QTY = Val(grdWHTITWOX.Selected.Rows(0).Cells("GTIN_QTY").Text)
        End If
        Dim QTY_SQL As String = String.Format("{0} GTIN_QTY,", GTIN_QTY)

        Dim sql As New Text.StringBuilder With {.Length = 0}
        If chkRanges.Checked Then
            sql.AppendLine("SELECT")
            sql.AppendLine("S1.RANGE_STYLE_CODE AS STYLE_CODE,")
            sql.AppendLine("S1.RANGE_STYLE_DESC AS STYLE_DESC,")
            sql.AppendLine("G1.GTIN_UPC_CODE,")
            If TABLENAME = "WHFITWOL" Then
                sql.AppendLine(QTY_SQL)
            End If
            sql.AppendLine("G1.GTIN_CODE")
            sql.AppendLine("FROM ICVLUPC1 L1, ICTRSTY1 S1, ICTGTINT G1")
            sql.AppendLine("WHERE L1.STYLE_CODE = S1.RANGE_STYLE_CODE")
            sql.AppendLine("AND L1.UPC_CODE = G1.GTIN_UPC_CODE")
            If TABLENAME = "WHTITWOX" Then
                sql.AppendLine(String.Format("AND S1.RANGE_STYLE_CODE >= '{0}'", begStyle))
                sql.AppendLine(String.Format("AND S1.RANGE_STYLE_CODE <= '{0}'", endStyle))
                If custCode.Length > 0 Then
                    sql.AppendLine(String.Format("AND S1.CUST_CODE = '{0}'", custCode))
                End If
            Else
                Dim STYLE_CODE As String = grdWHTITWOX.Selected.Rows(0).Cells("STYLE_CODE").Text
                sql.AppendLine(String.Format("AND S1.RANGE_STYLE_CODE = '{0}'", STYLE_CODE))
            End If
        Else
            sql.AppendLine("SELECT")
            sql.AppendLine("S1.STYLE_CODE AS STYLE_CODE,")
            sql.AppendLine("S1.STYLE_DESC AS STYLE_DESC,")
            sql.AppendLine("G1.GTIN_UPC_CODE,")
            If TABLENAME = "WHFITWOL" Then
                sql.AppendLine(QTY_SQL)
            End If
            sql.AppendLine("G1.GTIN_CODE")
            sql.AppendLine("FROM ICVLUPC1 L1, ICTSTYL1 S1, ICTGTINT G1")
            sql.AppendLine("WHERE L1.STYLE_CODE = S1.STYLE_CODE")
            sql.AppendLine("AND L1.UPC_CODE = G1.GTIN_UPC_CODE")
            If TABLENAME = "WHTITWOX" Then
                sql.AppendLine(String.Format("AND S1.STYLE_CODE >= '{0}'", begStyle))
                sql.AppendLine(String.Format("AND S1.STYLE_CODE <= '{0}'", endStyle))
                If custCode.Length > 0 Then
                    sql.AppendLine(String.Format("AND S1.CUST_CODE = '{0}'", custCode))
                End If
            Else
                Dim STYLE_CODE As String = grdWHTITWOX.Selected.Rows(0).Cells("STYLE_CODE").Text
                sql.AppendLine(String.Format("AND S1.STYLE_CODE = '{0}'", STYLE_CODE))
            End If
        End If
        RetVal = sql.ToString
        Return RetVal
    End Function
#End Region

End Class