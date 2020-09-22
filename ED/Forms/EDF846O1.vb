Public Class EDF846O1

    Private wkTable As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "EDTSYSIH", "*")
            Create_TDA(.Tables.Add, "EDT846O1", "*")
            .Tables("EDT846O1").Columns.Add("EDI_TP_ID", GetType(System.String))
            Create_TDA(.Tables.Add, "EDT846O2", "*", 2)

            Create_TDA(.Tables.Add, "ECTECOM1", "*")
            Fill_Records("ECTECOM1", String.Empty, True, "SELECT * FROM ECTECOM1")

            Create_TDA(.Tables.Add, "EDTTRPM1", "*")
            Fill_Records("EDTTRPM1", String.Empty, True, "SELECT * FROM EDTTRPM1 where EDI_DOC_NO = '846'")

        End With

        grdEDT846O1.DataSource = dst.Tables("EDT846O1")
        grdEDT846O1.DisplayLayout.Bands(0).Columns("EDI_TP_ID").Header.VisiblePosition = 1
        grdEDT846O2.DataSource = dst.Tables("EDT846O2")

        Create_Summary(grdEDT846O1, "EDI_OUTBOUND_DOC_NO", "Count")
        Create_Summary(grdEDT846O2, "EDI_DOC_LNO", "Count")

        ASCMAIN1.Add_Value_List(grdEDT846O1, "EDI_TP_ID", "SELECT EDI_TP_ID, ECOM_CODE FROM ECTECOM1")

        dteSearchE.MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)
        dteSearchE.MinDate = DateAdd(DateInterval.Day, -60, DateTime.Now)
        dteSearchE.DateTime = DateTime.Now

        dteSearchS.MaxDate = dteSearchE.MaxDate
        dteSearchS.MinDate = dteSearchE.MinDate
        dteSearchS.DateTime = dteSearchE.DateTime

        wkTable = ASCMAIN1.Temp_Table("Select EDI_OUTBOUND_DOC_NO FROM EDTSYSIH WHERE ROWNUM < 1")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Generate"
                Validate_Code("ECOM_CODE")
                If EMsg.Length = 0 Then
                    If MessageBox.Show($"Do you want to generate an EDI 846 for Ecom Code: {Absx1.txtFor("ECOM_CODE").Text}?") = DialogResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Update"
                If dst.Tables("EDTSYSIH").Rows.Count = 0 OrElse dst.Tables("EDT846O1").Rows.Count = 0 OrElse dst.Tables("EDT846O2").Rows.Count = 0 Then
                    EMsg &= vbCr & "There is not EDI 846 to Update."
                Else
                    If MessageBox.Show($"Do you want to Update the EDI 846 for Ecom Code: {Absx1.txtFor("ECOM_CODE").Text}?") = DialogResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Cancel"
                If MessageBox.Show($"Do you want to Cancel the EDI 846 for Ecom Code: {Absx1.txtFor("ECOM_CODE").Text}?") = DialogResult.No Then
                    Exit Sub
                End If

            Case "Refresh"
                If MessageBox.Show($"Do you want to Refresh the EDI 846?") = DialogResult.No Then
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

            Case "Generate"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Refresh"
                Refresh_Entries()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Generate").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = not_iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() {"EDTSYSIH", "EDT846O1", "EDT846O2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        Absx1.txtFor("ECOM_CODE").Clear()

        Refresh_Entries()

        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Generating EDI 846")

        Try
            EnforceConstraints(False)

            Dim ECOM_CODE As String = Absx1.txtFor("ECOM_CODE").Text

            For Each TABLE_NAME As String In New String() {"EDTSYSIH", "EDT846O1", "EDT846O2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next

            Dim Emsg As String = String.Empty
            Dim clsEDC84601 As New TAC.EDC84601(dst.Tables.Item("EDT846O1"), dst.Tables.Item("EDT846O2"), dst.Tables.Item("EDTSYSIH"))
            Dim EDI_OUTBOUND_DOC_NO As String = clsEDC84601.CreateEDI846(ECOM_CODE, Emsg)

            If Emsg.Length > 0 Then
                MessageBox.Show($"Message returned by Generation of EDI 846: {Emsg}", "Generate EDI 846", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            EnforceConstraints(True)

            'grdEDT846O1.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
            'grdEDT846O2.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

        Catch ex As Exception
            MessageBox.Show($"Error Generating EDI 846 {ex.Message }", "Generate EDI 846", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Clear_Record()
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End Try
    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()

            Update_Record_TDA("EDTSYSIH")
            Update_Record_TDA("EDT846O1")
            Update_Record_TDA("EDT846O2")

            CommitTrans("EDI 846 Updated")

        Catch ex As Exception
            Rollback($"Error Updating EDI 846: {ex.Message }")
        End Try

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(
    ByVal ctl As Control,
    ByVal COLUMN_NAME As String,
    Optional ByRef sql_where As String = "",
    Optional ByRef cancel As Boolean = False)
        Select Case COLUMN_NAME
            'Case "SEASON_CODE"
            '    If Absx1.optFor("STMT_TYPE").CheckedIndex <> -1 Then
            '        sql_where = "STMT_TYPE = '" & Absx1.optFor("STMT_TYPE").Value & "'"
            '    End If
        End Select
    End Sub

    Private Sub Refresh_Entries()

        Dim ECOM_CODE As String = Absx1.txtFor("ECOM_CODE").Text

        For Each TABLE_NAME As String In New String() {"EDTSYSIH", "EDT846O1", "EDT846O2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        Try
            ASCDATA1.ExecuteSQL($"DELETE FROM {wkTable}")
            ASCMAIN1.sql = $"SELECT EDI_OUTBOUND_DOC_NO 
                                FROM EDTSYSIH 
                                where TRUNC(INIT_DATE) BETWEEN '{dteSearchS.DateTime.ToString("dd-MMM-yyyy")}' AND '{dteSearchE.DateTime.ToString("dd-MMM-yyyy")}' 
                                and EDI_APPLICATION_ID = 'IB'"

            If ECOM_CODE.Length > 0 Then
                Dim rowECTECOM1 As DataRow = dst.Tables("ECTECOM1").Rows.Find(ECOM_CODE)
                Dim EDI_OUR_ID As String = String.Empty

                Dim EDI_TP_ID As String = String.Empty
                Dim EDI_TP_QUAL As String = String.Empty

                If rowECTECOM1 IsNot Nothing Then
                    EDI_TP_ID = rowECTECOM1.Item("EDI_TP_ID") & String.Empty
                    EDI_TP_QUAL = rowECTECOM1.Item("EDI_TP_QUAL") & String.Empty

                    Dim Sql As String = $"EDI_TP_QUAL = '{EDI_TP_QUAL}' and EDI_TP_ID = '{EDI_TP_ID}'"

                    If dst.Tables("EDTTRPM1").Select(Sql).Length > 0 Then
                        Dim rowEDTTRPM1 As DataRow = dst.Tables("EDTTRPM1").Select(Sql)(0)
                        EDI_OUR_ID = rowEDTTRPM1.Item("EDI_OUR_ID") & String.Empty
                    End If

                End If

                If EDI_OUR_ID.Length > 0 AndAlso EDI_TP_ID.Length > 0 Then
                    ASCMAIN1.sql &= $" and EDI_OUR_ID = '{EDI_OUR_ID}' and EDI_TP_ID = '{EDI_TP_ID}' "
                End If
            End If

            ASCDATA1.ExecuteSQL($"INSERT INTO {wkTable}  {ASCMAIN1.sql}")

            Fill_Records("EDT846O1", String.Empty, True, $"SELECT EDTSYSIH.EDI_TP_ID, EDT846O1.* FROM EDT846O1, EDTSYSIH WHERE EDT846O1.EDI_OUTBOUND_DOC_NO IN (SELECT EDI_OUTBOUND_DOC_NO FROM {wkTable}) AND EDT846O1.EDI_OUTBOUND_DOC_NO = EDTSYSIH.EDI_OUTBOUND_DOC_NO (+)")

        Catch ex As Exception
            MessageBox.Show($"Error getting EDI 846s: {ex.Message }", "Load EDI 846", MessageBoxButtons.OK, MessageBoxIcon.Error)
            For Each TABLE_NAME As String In New String() {"EDTSYSIH", "EDT846O1", "EDT846O2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
        End Try

    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdEDT846O1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdEDT846O2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")
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

        Select Case e.SourceControl.Name
            'Case "grdPOTORDRR"
            '    If EntryMode = "V" Then e.Cancel = True

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdPOTORDR3"
                '    tlb_sbt = DirectCast(tlb.Tools("Show Cartons"), UltraWinToolbars.StateButtonTool)
                '    e.Tool.SharedProps.Visible = tlb_sbt.Checked
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            'Case "Style Multi-Color"
            '    Using F As New TAC.ICFSTYCX
            '        F.STYLE_CODE = ""
            '        F.Price_Caption = "Cost" & IIf(ssdDZGRD.Value = 1, "", "/Dz")
            '        F.ShowDialog()
            '        If F.STYLE_CODE <> "" Then
            '            Add_Colors(F.STYLE_CODE, F.dst.Tables("ICTCOLRM"), F.PRICE)
            '        End If
            '    End Using

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("EDI_STYLE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

        End Select
    End Sub

#End Region

    Private Sub grdEDT846O1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdEDT846O1.AfterRowActivate

        If grdEDT846O1.ActiveRow Is Nothing Then
            Exit Sub
        End If

        Dim EDI_OUTBOUND_DOC_NO As String = grdEDT846O1.ActiveRow.Cells("EDI_OUTBOUND_DOC_NO").Value & String.Empty

        ' LOAD THEM ONLY IF THEY AERE NEEDED. TAKES TO LONG
        If dst.Tables("EDT846O2").Select($"EDI_OUTBOUND_DOC_NO = '{EDI_OUTBOUND_DOC_NO}'").Length = 0 Then
            Me.Cursor = Cursors.WaitCursor
            Fill_Records("EDT846O2", New Object() {ASCMAIN1.CLIENT, EDI_OUTBOUND_DOC_NO}, False, String.Empty)
            Me.Cursor = Cursors.Default
        End If

        Dim dView As New DataView(dst.Tables("EDT846O2"))
        dView.RowFilter = $"EDI_OUTBOUND_DOC_NO = '{EDI_OUTBOUND_DOC_NO}'"
        grdEDT846O2.DataSource = dView

    End Sub


End Class