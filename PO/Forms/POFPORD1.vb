Public Class POFPORD1

    Dim PO_BATCH_NO As String
    Dim CLASS_CODE As String
    Dim rowPOTPORD1 As DataRow
    Dim rowICTCLAS1 As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("POTPARM1")

        With dst
            ASCMAIN1.sql = "Select POTPORD1.*" & vbCrLf _
            & " from POTPORD1 " & vbCrLf
            Create_TDA(.Tables.Add, "POTPORDX", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select POTPORD1.*" & vbCrLf _
            & " from POTPORD1 " & vbCrLf _
            & " where POTPORD1.PO_BATCH_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTPORD1", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select POTPORD2.*" & vbCrLf _
            & " from POTPORD2 " & vbCrLf _
            & " where POTPORD2.PO_BATCH_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTPORD2", "**", 0, True, "V", 1)
            With .Tables("POTPORD2").Columns
                .Add("UNITS", GetType(System.Decimal), "2-1")
            End With
        End With

        grdPOTPORDX.DataSource = dst.Tables("POTPORDX")
        grdPOTPORD2.DataSource = dst.Tables("POTPORD2")

        Create_Summary(grdPOTPORDX, "PO_BATCH_NO", "Count")

        Create_Summary(grdPOTPORD2, "STYLE_CODE", "Count")
        'Create_Summary(grdPOTPORD2, New String() {"PO_QTY_SHP", "UNITS"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("CLASS_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Class"
                Else
                    rowICTCLAS1 = LookUp("ICTCLAS1", Absx1.txtFor("CLASS_CODE").Text)
                    If rowICTCLAS1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Class " & Absx1.txtFor("CLASS_CODE").Text
                    End If
                End If

                If EMsg = "" Then
                    
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("POTPORD1", CLASS_CODE) Then Exit Sub
                End If

            Case "Edit", "Load"

                CLASS_CODE = ""
                PO_BATCH_NO = ""

                If Absx1.txtFor("PO_BATCH_NO").Text = "" Then
                    EMsg &= vbCr & "No  Batch No Specified"
                Else
                    PO_BATCH_NO = Absx1.txtFor("PO_BATCH_NO").Text
                    rowPOTPORD1 = LookUp("POTPORD1", PO_BATCH_NO)
                    If rowPOTPORD1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Batch No " & PO_BATCH_NO
                    Else
                        CLASS_CODE = rowPOTPORD1.Item("CLASS_CODE")
                        If rowPOTPORD1.Item("BATCH_STATUS") & "" <> "O" And eItemKey = "Edit" Then
                            Select Case rowPOTPORD1.Item("BATCH_STATUS")
                                Case "C"
                                    EMsg &= vbCr & "Batch No " & PO_BATCH_NO & " has been Cancelled"
                                Case "D"
                                    EMsg &= vbCr & "Batch No " & PO_BATCH_NO & " has been Deleted"
                                Case Else ' such as "F"
                                    EMsg &= vbCr & "Batch No " & PO_BATCH_NO & " is No Longer Open"
                            End Select
                        End If
                    End If
                End If

                If EMsg = "" And EntryMode = "E" Then
                    If Not ASCMAIN1.Logical_Lock("POTPORD1", PO_BATCH_NO) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("POTPORD1", CLASS_CODE) Then Exit Sub
                End If

            Case "Update"
                'If Absx1.dteFor("ORDR_SHIP_DATE").Value & "" = "" _
                '    Or Absx1.dteFor("ORDR_CANCEL_DATE").Value & "" = "" Then
                '    EMsg &= vbCr & "Ship Date and Cancel Date are Mandatory"
                'Else
                '    If Format(Absx1.dteFor("ORDR_SHIP_DATE").Value, "yyyyMMdd") _
                '     > Format(Absx1.dteFor("ORDR_CANCEL_DATE").Value, "yyyyMMdd") Then
                '        EMsg &= vbCr & "Cancel Date cannot be Prior to Ship Date"
                '    End If
                'End If

                If grdPOTPORD2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Styles on Batch"
                Else
                    If Val(dst.Tables("POTPORD2").Compute("COUNT(STYLE_CODE)", "PO_QTY > 0") & "") = 0 Then
                        EMsg &= vbCr & "No Styles on Batch with PO Qty >0"
                    End If
                End If

            Case "Delete"
                If EMsg = "" Then
                    If MsgBox("Do you want to Mark this Batch as Deleted", _
                              MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub


    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    If (EntryMode = "L" And ScreenMode) Then
                        If rowPOTPORD1.Item("BATCH_STATUS") & "" = "O" Then
                            .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                        Else
                            .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                        End If
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If
                    .Items("Update").Settings.Enabled = iScreenMode

                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode

                    .Items("Cancel Balance").Settings.Enabled = iScreenMode

                    .Items("Done").Visible = (EntryMode = "L" And ScreenMode)
                    .Items("Print").Visible = ScreenMode
                    .Items("Update").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
                    .Items("Delete").Visible = (EntryMode = "E")
                    .Items("Cancel").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
                End With
                .Groups("Sales History").Visible = ScreenMode
                .Groups("Style Filters").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        lblStatus.Visible = ScreenMode
        grdPOTPORDX.Visible = Not ScreenMode
        splPOTPORDA.Visible = ScreenMode

        If ScreenMode Then
            With grdPOTPORD2.DisplayLayout.Override
                If EntryMode = "L" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.False

                    'With grdPOTPORD2.DisplayLayout.Bands(0)
                    '    If EntryMode <> "E" Then
                    '        .Columns("X").Hidden = True
                    '    Else
                    '        .Columns("X").Hidden = False
                    '    End If
                    'End With
                End If

            End With
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"POTPORD1", "POTPORD2", "SOTSLSC1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Load_POTPORDX()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        If EntryMode = "N" Then
            PO_BATCH_NO = ASCMAIN1.Next_Control_No("POTPORD1.PO_BATCH_NO")

            rowPOTPORD1 = dst.Tables("POTPORD1").NewRow
            With rowPOTPORD1
                .Item("PO_BATCH_NO") = PO_BATCH_NO
                .Item("CLASS_CODE") = CLASS_CODE
                .Item("BATCH_STATUS") = "O"
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
            End With
            dst.Tables("POTPORD1").Rows.Add(rowPOTPORD1)

        Else
            rowPOTPORD1 = Fill_Record("POTPORD1", PO_BATCH_NO)
        End If

        CLASS_CODE = rowPOTPORD1.Item("CLASS_CODE")
        rowICTCLAS1 = Fill_Record("ICTCLAS1", CLASS_CODE)

        Fill_Records("POTPORD2", PO_BATCH_NO)
        Sort_grdColumns(grdPOTPORD2, "STYLE_CODE,COLOR_CODE")

        If EntryMode = "N" Then
            lblStatus.Text = "New Batch"
        Else
            Select Case rowPOTPORD1.Item("BATCH_STATUS")
                Case "O"
                    lblStatus.Text = "Open"
                Case "C"
                    lblStatus.Text = "Closed"
                Case "D"
                    lblStatus.Text = "Deleted"
            End Select
        End If

        With grdPOTPORD2.DisplayLayout.Bands(0)
            If (EntryMode = "E" Or EntryMode = "N") Then
                .Columns("PO_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                'If EntryMode = "E" Then
                '    .Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                'Else
                '    .Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                'End If
            Else
                .Columns("PO_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                '.Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        If (EntryMode = "N" Or EntryMode = "E") Then ' InquiryMode uses EntryMode = "V"
            With grdPOTPORD2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.False
            End With
            'grdPOTPORD2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            With grdSOTRSRV2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
            'grdPOTPORD2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        End If

        'Display_Totals()
        EnforceConstraints(True)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        CommitTrans("Delete")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Stop
        CommitTrans("Update Complete")
    End Sub

    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Report")

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Generate_Report("PORWREC2")
        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "PO_BATCH_NO"
                'sql_where = "STATUS = '0'"
        End Select

    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "PO_SHIPMENT_NO"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Call Click_Command("Load", e)
            '    End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "PO_SHIPMENT_NO"
            '    Call Click_Command("View")
        End Select
    End Sub

#End Region

    Private Sub grdPOTPORDX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTPORDX.AfterRowActivate

    End Sub

    Private Sub grdPOTPORDX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTPORDX.DoubleClickRow
        'If grdPOTPORDX.ActiveRow IsNot Nothing Then
        '    Absx1.txtFor("PO_SHIPMENT_NO").Text = grdPOTPORDX.ActiveRow.Cells("PO_SHIPMENT_NO").Text
        '    Click_Command("View")
        'End If
    End Sub

    Sub Setup_SOTSLSC1()
        If grdPOTPORD2.ActiveRow Is Nothing OrElse Not grdPOTPORD2.ActiveRow.IsDataRow Then
            grdSOTSLSC1.Visible = False
        Else
            Dim STYLE_CODE As String = grdPOTPORD2.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = Val(grdPOTPORD2.ActiveRow.Cells("COLOR_CODE").Value & "")

            Fill_Records("SOTSLSC1", New Object() {STYLE_CODE, COLOR_CODE})
            Sort_grdColumns(grdSOTSLSC1, "WHSE_CODE,CUST_CODE")

            grdSOTSLSC1.Text = "Style " & STYLE_CODE & " Color " & COLOR_CODE & "; Sales Summary"
            grdSOTSLSC1.Visible = True
        End If
    End Sub

    Sub Load_POTPORDX()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        Fill_Records("POTPORDX")
        Sort_grdColumns(grdPOTPORDX, "PO_BATCH_NO".ToLower)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdPOTPORDX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTPORDX.InitializeRow
        If e.Row.Cells("BATCH_STATUS").Value & "" <> "O" Then
            e.Row.CellAppearance.BackColor = Drawing.Color.LightGray
        End If
    End Sub

    Private Function grdSOTRSRV2() As Object
        Throw New NotImplementedException
    End Function

End Class