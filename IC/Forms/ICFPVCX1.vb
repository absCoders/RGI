Public Class ICFPVCX1
    Private sql As New System.Text.StringBuilder With {.Length = 0}
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")

        With dst
            sql.Length = 0
            sql.AppendLine("Select *")
            ASCMAIN1.sql = sql.ToString
            Create_TDA(.Tables.Add, "ICFPVCX1", "**", 0, False, "", 0)
        End With

        grdICTPVCX1.DataSource = dst.Tables("ICFPVCX1")

        Create_Summary(grdICTPVCX1, "STYLE_CODE", "Count")

        spl.Panel1Collapsed = True
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"

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
                .Groups("Screen Control").Visible = False
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdICTPVCX1.Visible = True

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        For Each TABLE_NAME As String In New String() {"ICFPVCX1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            'rowICTPLIN1.Item("STMT_TYPE") = HFs("STMT_TYPE")
            'rowICTPLIN1.Item("STMT_DESC") = HFs("STMT_DESC")
        Else

        End If

        If EntryMode = "N" Then
        Else

        End If

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

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTPVCX1, "SS", "Show Filter", "Show GroupBox")
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
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = Lookup("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        'Select Case Absx1.GetABSColumnName(sender)
        '    Case "LP_CODE"
        '        If e.KeyCode = Windows.Forms.Keys.Enter Then
        '            Me.UltraGroupBox1.Select() ' to force txt_Leave event to fire, for formatting
        '            Load_ICTSTYC1()
        '        End If
        'End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        'Select Case Absx1.GetABSColumnName(txtctl)
        '    Case "LP_CODE"
        '        Load_ICTSTYC1()
        'End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

        'With Absx1.txtFor(COLUMN_NAME)
        '    Select Case COLUMN_NAME

        '        Case "LP_CODE"
        '            Load_ICTSTYC1()

        '    End Select

        'End With
    End Sub

#End Region

End Class