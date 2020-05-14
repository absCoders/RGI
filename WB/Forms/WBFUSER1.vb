
Public Class WBFUSER1
    Dim InquiryOnly As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        With dst
            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM RGTUSER1")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "RGTUSERX", "**", 0, False)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM RGTUSER1 WHERE USER_ID = :PARM1")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "RGTUSER1", "**", 0, True, "V", 1)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM RGTUSER2 WHERE USER_ID = :PARM1")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "RGTUSER2", "**", 0, True, "V", 2)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM RGTUSERT WHERE USER_ID = :PARM1")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "RGTUSERT", "**", 0, True, "V", 1)

        End With

        grdRGTUSERX.DataSource = dst.Tables("RGTUSERX")

        'Create_Summary(grdSOTRUSSE, "NEW_QTY", "Sum", "", "###,##0")

        Sort_grdColumns(grdRGTUSERX, "USER_ID", False)

        'grdSOTRUSSE.DisplayLayout.UseFixedHeaders = True
        'With grdSOTRUSSE.DisplayLayout.Bands(0)
        '    For Each COLUMN_NAME As String In New String() {"ORDR_NO", "ORDR_LNO", "STYLE_CODE", "COLOR_CODE", "ORDR_QTY", "ORDR_UNIT_PRICE"}
        '        .Columns(COLUMN_NAME).Header.Fixed = True
        '    Next
        'End With

        tab.Visible = False
        grdRGTUSERX.Parent = tab.Parent

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("USER_ID").Text.Length = 0 Then
                    EMsg &= vbCr & "User ID May Not Be Blank"
                Else
                    Dim rowRGTUSER1 As DataRow = LookUp("RGTUSER1", Absx1.txtFor("USER_ID").Text)
                    If Not IsNothing(rowRGTUSER1) Then
                        EMsg &= vbCr & String.Format("User ID {0} Already Exists", Absx1.txtFor("USER_ID").Text)
                    End If
                    Dim UpperCount As Integer = 0
                    Dim NumCount As Integer = 0
                    For Each CharU As String In Absx1.txtFor("USER_ID").Text
                        If Char.IsUpper(CharU) Then
                            UpperCount += 1
                        End If
                        If IsNumeric(CharU) Then
                            NumCount += 1
                        End If
                    Next
                    If UpperCount > 0 Then
                        EMsg &= vbCr & "User ID Can Not Have Any Upper Case Letters"
                    End If
                    If NumCount > 0 Then
                        EMsg &= vbCr & "User ID Can Not Have Any Numbers"
                    End If
                End If
            Case "Edit"
                'Dim x As Boolean = False
                'If x = False Then
                '    EMsg &= vbCr & "Some Kind Of Error."
                'End If
                'If EMsg = "" Then
                '    If Not ASCMAIN1.Logical_Lock("ICTSTYL1", Absx1.txtFor("STYLE_CODE").Text) Then
                '        Exit Sub
                '    End If
                'End If

            Case "Cancel"

            Case "Update"
                Dim UpperCount As Integer = 0
                Dim NumCount As Integer = 0
                Dim CharCnt As Integer = 0
                For Each CharU As String In Absx1.txtFor("USER_PASSWORD").Text
                    CharCnt += 1
                    If Char.IsUpper(CharU) Then
                        UpperCount += 1
                    End If
                    If IsNumeric(CharU) Then
                        NumCount += 1
                    End If
                Next
                If UpperCount = 0 Then
                    EMsg &= vbCr & "Password Must Have At least One Capital Letter"
                End If
                If NumCount = 0 Then
                    EMsg &= vbCr & "Password Must Have At least One Number"
                End If
                If CharCnt <= 3 Then
                    EMsg &= vbCr & "Password Must Be At Least 3 Characters"
                End If

                If txtSREPCODE.Text.Length = 0 Then
                    EMsg &= vbCr & "Sales Rep Code Can Not Be Blank"
                Else
                    Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
                    SQLS.AppendLine(String.Format("Select Count(*) from SOTSREP1 where SREP_CODE = '{0}'", txtSREPCODE.Text))
                    ASCMAIN1.sql = SQLS.ToString()
                    Dim SCNT As Int16 = Val(ASCDATA1.GetDataValue)
                    If SCNT = 0 Then
                        EMsg &= vbCr & "Invalid Sales Rep Code Entered"
                    End If
                End If

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
            Case "New"
                EntryMode = "N"
                Call Load_Record()
                Call Mode_Settings(True)
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
                .Groups("Screen Control").Items("New").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Edit").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Done").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Update").Visible = ScreenMode
                .Groups("Screen Control").Items("Cancel").Visible = ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdRGTUSERX.Visible = Not tf

        With grdRGTUSERX.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

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

        If Not ScreenMode Then
            RefreshRGTUSERX()
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()
        dst.Tables("RGTUSER1").Rows.Clear()
        dst.Tables("RGTUSER2").Rows.Clear()
        dst.Tables("RGTUSERT").Rows.Clear()
        txtSREPCODE.Text = ""
    End Sub

    Sub Load_Record()
        Call Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Call Fill_Records("RGTUSER1", Absx1.txtFor("USER_ID").Text, True)
        Call Fill_Records("RGTUSER2", Absx1.txtFor("USER_ID").Text, True)
        Call Fill_Records("RGTUSERT", Absx1.txtFor("USER_ID").Text, True)

        EnforceConstraints(True)


        If EntryMode = "N" Then
            Dim rowRGTUSER1 As DataRow = dst.Tables("RGTUSER1").NewRow
            rowRGTUSER1.Item("USER_ID") = Absx1.txtFor("USER_ID").Text
            rowRGTUSER1.Item("USER_STATUS") = "A"
            rowRGTUSER1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowRGTUSER1.Item("INIT_DATE") = Now()
            dst.Tables("RGTUSER1").Rows.Add(rowRGTUSER1)

            Dim rowRGTUSER2 As DataRow = dst.Tables("RGTUSER2").NewRow
            rowRGTUSER2.Item("USER_ID") = Absx1.txtFor("USER_ID").Text
            rowRGTUSER2.Item("SECURITY_CODE") = "SL"
            dst.Tables("RGTUSER2").Rows.Add(rowRGTUSER2)

            Dim rowRGTUSERT As DataRow = dst.Tables("RGTUSERT").NewRow
            rowRGTUSERT.Item("USER_ID") = Absx1.txtFor("USER_ID").Text
            rowRGTUSERT.Item("SREP_CODE") = ""
            dst.Tables("RGTUSERT").Rows.Add(rowRGTUSERT)

            chkIsActive.Checked = True
        Else
            dst.AcceptChanges()
        End If
        Bind_Controls(panUSERS, "RGTUSER1")

        SetSrepCode()

        SetIsActive()

        SetIsSuperUser()

    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
        'Not entertaining this.
    End Sub

    Sub Update_Record()
        Call BeginTrans()

        For Each rowRGTUSER1 As DataRow In dst.Tables("RGTUSER1").Select()
            rowRGTUSER1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowRGTUSER1.Item("LAST_DATE") = Now()
        Next

        Update_Record_TDA("RGTUSER1")
        Update_Record_TDA("RGTUSER2")
        Update_Record_TDA("RGTUSERT")
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
#End Region

#Region "Custom Methods"

#End Region

    Private Sub RefreshRGTUSERX()
        Fill_Records("RGTUSERX")
    End Sub

    Private Sub grdRGTUSERX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdRGTUSERX.DoubleClickRow
        If Not IsDBNull(e.Row.Cells("USER_ID").Value) Then
            Absx1.txtFor("USER_ID").Text = e.Row.Cells("USER_ID").Value
            Click_Command("Edit")
        End If
    End Sub

    Private Sub SetSrepCode()
        Dim Filter As String = String.Format("USER_ID = '{0}'", Absx1.txtFor("USER_ID").Text)
        If dst.Tables("RGTUSERT").Select(Filter).Count = 1 Then
            txtSREPCODE.Text = dst.Tables("RGTUSERT").Select(Filter).FirstOrDefault.Item("SREP_CODE").ToString
        End If
    End Sub

    Private Sub SetIsActive()
        If dst.Tables("RGTUSER1").Rows.Count = 1 Then
            If dst.Tables("RGTUSER1").Rows(0).Item("USER_STATUS").ToString = "A" Then
                chkIsActive.Checked = True
            Else
                chkIsActive.Checked = False
            End If
        Else
            chkIsActive.Checked = False
        End If
    End Sub

    Private Sub SetIsSuperUser()
        Dim Filter As String = String.Format("USER_ID = '{0}' AND SECURITY_CODE = 'X6'", Absx1.txtFor("USER_ID").Text)
        If dst.Tables("RGTUSER2").Select(Filter).Count = 1 Then
            chkIsSuper.Checked = True
        Else
            chkIsSuper.Checked = False
        End If
    End Sub

    Private Sub txtSREPCODE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtSREPCODE.ValueChanged
        If txtSREPCODE.Text <> "" Then
            Dim Filter As String = String.Format("USER_ID = '{0}'", Absx1.txtFor("USER_ID").Text)
            If dst.Tables("RGTUSERT").Select(Filter).Count = 1 Then
                Dim rowRGTUSERT As DataRow = dst.Tables("RGTUSERT").Select(Filter).FirstOrDefault
                rowRGTUSERT.Item("SREP_CODE") = txtSREPCODE.Text
            Else
                Dim newRGTUSERT As DataRow = dst.Tables("RGTUSERT").NewRow
                newRGTUSERT.Item("USER_ID") = Absx1.txtFor("USER_ID").Text
                newRGTUSERT.Item("SREP_CODE") = txtSREPCODE.Text
                dst.Tables("RGTUSERT").Rows.Add(newRGTUSERT)
            End If
        End If
    End Sub

    Private Sub chkIsActive_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkIsActive.CheckedChanged
        If dst.Tables("RGTUSER1").Rows.Count = 1 Then
            If chkIsActive.Checked = True Then
                dst.Tables("RGTUSER1").Rows(0).Item("USER_STATUS") = "A"
            Else
                dst.Tables("RGTUSER1").Rows(0).Item("USER_STATUS") = "I"
            End If
        End If
    End Sub

    Private Sub chkIsSuper_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkIsSuper.CheckedChanged
        Dim Filter As String = String.Format("USER_ID = '{0}' AND SECURITY_CODE = 'X6'", Absx1.txtFor("USER_ID").Text)
        If chkIsSuper.Checked = True Then
            If dst.Tables("RGTUSER2").Select(Filter).Count = 0 Then
                Dim newRGTUSER2 As DataRow = dst.Tables("RGTUSER2").NewRow
                newRGTUSER2.Item("USER_ID") = Absx1.txtFor("USER_ID").Text
                newRGTUSER2.Item("SECURITY_CODE") = "X6"
                dst.Tables("RGTUSER2").Rows.Add(newRGTUSER2)
            End If
        Else
            If dst.Tables("RGTUSER2").Select(Filter).Count = 1 Then
                Dim rowRGTUSER2 As DataRow = dst.Tables("RGTUSER2").Select(Filter).FirstOrDefault
                rowRGTUSER2.Delete()
            End If
        End If
    End Sub
End Class