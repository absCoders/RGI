Public Class EDF846T1

    Dim EDI_DOC_SEQ_NO As String
    Dim rowEDT846T1 As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")

        ASCMAIN1.sql = "Update EDT846T1 Set EDI_PROCESS_IND = '0', EDI_TP_ID = TRIM(EDI_TP_ID) where EDI_PROCESS_IND is Null"
        ASCDATA1.ExecuteSQL()

        With dst
            ASCMAIN1.sql = "Select * from EDT846T1 where EDI_TP_QUAL = :PARM1 and EDI_TP_ID = :PARM2 and EDI_PROCESS_IND = '0'"
            Create_TDA(.Tables.Add, "EDT846T1", "**", 0, False, "VV", 1)

            'ASCMAIN1.sql = "Select EDT846T2.*, ICTSTAT2.WHSE_QTY_ON_HAND" _
            '    & " from EDT846T2,ICTSTAT2 where EDT846T2.EDI_DOC_SEQ_NO = :PARM1" _
            '    & " and ICTSTAT2.STYLE_CODE (+) = EDT846T2.STYLE_CODE" _
            '    & " and ICTSTAT2.COLOR_CODE (+) = 'AST'"
            ASCMAIN1.sql = "Select EDT846T2.* from EDT846T2 where EDT846T2.EDI_DOC_SEQ_NO = :PARM1"
            Create_TDA(.Tables.Add, "EDT846T2", "**", 0, False, "V", 0)
            .Tables("EDT846T2").Columns.Add("WHSE_CODE")
            .Tables("EDT846T2").Columns.Add("WHSE_QTY_ON_HAND", GetType(System.Int64))
            .Tables("EDT846T2").Columns.Add("VARIANCE", GetType(System.Int64), "ISNULL(EDI_ON_HAND_QTY,0)-ISNULL(WHSE_QTY_ON_HAND,0)")
        End With

        grdEDT846T1.DataSource = dst.Tables("EDT846T1")
        grdEDT846T2.DataSource = dst.Tables("EDT846T2")

        Create_Summary(grdEDT846T1, "EDI_DOC_SEQ_NO", "Count")

        Create_Summary(grdEDT846T2, "EDI_DTL_SEQ", "Count")
        Create_Summary(grdEDT846T2, New String() {"EDI_ON_HAND_QTY", "EDI_ALLOCATED_QTY", "WHSE_QTY_ON_HAND", "VARIANCE"})

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                If grdEDT846T1.ActiveRow IsNot Nothing Then
                    EDI_DOC_SEQ_NO = grdEDT846T1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value
                End If
                rowEDT846T1 = LookUp("EDT846T1", EDI_DOC_SEQ_NO)
                If rowEDT846T1 Is Nothing Then
                    Exit Sub
                End If


                'If EMsg = "" Then
                '    If Not ASCMAIN1.Logical_Lock(TABLE_NAME, Absx1.txtFor("SEASON_CODE").Text) Then
                '        Exit Sub
                '    End If
                'End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdEDT846T1.Visible = Not ScreenMode
        '     grdEDT846T2.Visible = ScreenMode
        SplitContainer2.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        For Each TABLE_NAME As String In New String() {"EDT846T1", "EDT846T2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        Absx1.txtFor("LP_CODE").Text = ""
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Inventories")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            'rowICTPLIN1.Item("STMT_TYPE") = HFs("STMT_TYPE")
            'rowICTPLIN1.Item("STMT_DESC") = HFs("STMT_DESC")
        Else

        End If

        Dim WHSE_CODE As String = ""
        Dim EDI_ADDR_CODE As String = ""

        Fill_Records("EDT846T2", EDI_DOC_SEQ_NO)
        EDI_ADDR_CODE = rowEDT846T1.Item("EDI_ADDR_CODE")
        WHSE_CODE = Get_WHSE_CODE(EDI_ADDR_CODE)
        For Each row As DataRow In dst.Tables("EDT846T2").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
            row.Item("WHSE_CODE") = WHSE_CODE
        Next

        ' THE FOLLOWING LINES AND THE HARD CODE BELOW FOR 93 91 ARE CHGS MADE TO LAST MINUTE MODS TO CONSOLIDATE ALL WHSES INTO 93/91
        
        If grdEDT846T1.Selected.Rows.Count <> 0 Then
            For Each grow As UltraWinGrid.UltraGridRow In grdEDT846T1.Selected.Rows
                Dim EDS As String = grow.Cells("EDI_DOC_SEQ_NO").Value
                If EDS <> EDI_DOC_SEQ_NO Then
                    Fill_Records("EDT846T2", EDS, False)
                    WHSE_CODE = Get_WHSE_CODE(grow.Cells("EDI_ADDR_CODE").Value)
                    For Each row As DataRow In dst.Tables("EDT846T2").Select("EDI_DOC_SEQ_NO = '" & EDS & "'")
                        row.Item("WHSE_CODE") = WHSE_CODE
                    Next
                End If
            Next
        End If

        Dim sqlw1 As String = "EDI_DOC_SEQ_NO <> '" & EDI_DOC_SEQ_NO & "'"
        For Each row As DataRow In dst.Tables("EDT846T2").Select(sqlw1)
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            WHSE_CODE = row.Item("WHSE_CODE")
            Dim sqlw2 As String = " and STYLE_CODE = '" & STYLE_CODE & "' and WHSE_CODE = '" & WHSE_CODE & "'"
            Dim rowEDT846T2s() As DataRow = dst.Tables("EDT846T2").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & sqlw2)
            If rowEDT846T2s.Length = 0 Then
                Dim rowEDT846T2 As DataRow = dst.Tables("EDT846T2").NewRow
                rowEDT846T2.ItemArray = row.ItemArray
                rowEDT846T2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO

                dst.Tables("EDT846T2").Rows.Add(rowEDT846T2)
            Else
                rowEDT846T2s(0).Item("EDI_ON_HAND_QTY") = Val(rowEDT846T2s(0).Item("EDI_ON_HAND_QTY") & "") + Val(row.Item("EDI_ON_HAND_QTY") & "")
            End If
        Next
        dst.Tables("EDT846T2").AcceptChanges()
        ASCDATA1.DeleteRows(dst.Tables("EDT846T2"), "EDI_DOC_SEQ_NO <> '" & EDI_DOC_SEQ_NO & "'")
        dst.Tables("EDT846T2").AcceptChanges()

        ASCMAIN1.sql = "Select WHSE_CODE, STYLE_CODE, COLOR_CODE from ICTSTAT2 " _
            & " where WHSE_QTY_ON_HAND <> 0 and WHSE_CODE in ('95','91','96')"
        ASCMAIN1.sql = "Select ICTSTAT2.*,ICTSTYL1.STYLE_DESC from ICTSTAT2,ICTSTYL1" _
            & " where ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE and (ICTSTAT2.WHSE_CODE, ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE) " _
            & " in (" & ASCMAIN1.sql & ")"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            WHSE_CODE = row.Item("WHSE_CODE")
            ' If WHSE_CODE = "96" Then WHSE_CODE = "95"
            Dim sqlw2 As String = "STYLE_CODE = '" & STYLE_CODE & "' and WHSE_CODE = '" & WHSE_CODE & "'"
            Dim rowEDT846T2s() As DataRow = dst.Tables("EDT846T2").Select(sqlw2)
            If rowEDT846T2s.Length = 0 Then
                Dim rowEDT846T2 As DataRow = dst.Tables("EDT846T2").NewRow
                rowEDT846T2.Item("WHSE_CODE") = row.Item("WHSE_CODE")
                rowEDT846T2.Item("STYLE_CODE") = row.Item("STYLE_CODE")
                rowEDT846T2.Item("STYLE_DESC") = row.Item("STYLE_DESC")
                rowEDT846T2.Item("WHSE_QTY_ON_HAND") = row.Item("WHSE_QTY_ON_HAND")
                rowEDT846T2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                rowEDT846T2.Item("EDI_DTL_SEQ") = 0
                dst.Tables("EDT846T2").Rows.Add(rowEDT846T2)
            Else
                rowEDT846T2s(0).Item("WHSE_QTY_ON_HAND") = Val(rowEDT846T2s(0).Item("WHSE_QTY_ON_HAND") & "") + Val(row.Item("WHSE_QTY_ON_HAND") & "")
            End If
        Next

        If EntryMode = "N" Then
        Else

        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")


    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special( _
    ByVal ctl As Control, _
    ByVal COLUMN_NAME As String, _
    Optional ByRef sql_where As String = "", _
    Optional ByRef cancel As Boolean = False)
        Select Case COLUMN_NAME
            'Case "SEASON_CODE"
            '    If Absx1.optFor("STMT_TYPE").CheckedIndex <> -1 Then
            '        sql_where = "STMT_TYPE = '" & Absx1.optFor("STMT_TYPE").Value & "'"
            '    End If
        End Select
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(tvw, "BBB", "Insert Above", "Insert Below", "Insert Within")
        Load_Popup_Menu(grdEDT846T2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")
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

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "POE")

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "PO Shipment Inquiry"
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Text
                Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPI", "F", "POE")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "LP_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Me.UltraGroupBox1.Select() ' to force txt_Leave event to fire, for formatting
                    Load_EDT846T1()
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "LP_CODE"
                Load_EDT846T1()
        End Select
    End Sub


    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

        With Absx1.txtFor(COLUMN_NAME)
            Select Case COLUMN_NAME

                Case "LP_CODE"
                    Load_EDT846T1()

            End Select

        End With
    End Sub

#End Region

    Sub Load_EDT846T1()
        Dim LP_CODE As String = Absx1.txtFor("LP_CODE").Text
        Dim rowWHTTPLP1 As DataRow = LookUp("WHTTPLP1", LP_CODE)

        Dim EDI_TP_QUAL As String = rowWHTTPLP1.Item("EDI_TP_QUAL")
        Dim EDI_TP_ID As String = rowWHTTPLP1.Item("EDI_TP_ID")

        Fill_Records("EDT846T1", New String() {EDI_TP_QUAL, EDI_TP_ID})
        Sort_grdColumns(grdEDT846T1, "EDI_DOC_SEQ_NO".ToLower)
    End Sub

    Private Sub grdEDT846T1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdEDT846T1.DoubleClickRow
        If e.Row.IsDataRow Then
            EDI_DOC_SEQ_NO = e.Row.Cells("EDI_DOC_SEQ_NO").Value & ""
            Click_Command("View")
        End If
    End Sub

    Function Get_WHSE_CODE(EDI_ADDR_CODE As String) As String
        Dim WHSE_CODE As String = ""
        Select Case EDI_ADDR_CODE
            Case "NYA"
                WHSE_CODE = "91"
            Case "NYAW", "NYWM", "NYDG", "NYWB"
                WHSE_CODE = "95"
            Case Else
                Stop
        End Select

        Return WHSE_CODE
    End Function
End Class