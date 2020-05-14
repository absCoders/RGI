Public Class ICTCLAS1

    Dim sqlICTCLAS2 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select ICTCLAS2.* " _
            & " from ICTCLAS2" _
            & " where ICTCLAS2.STYLE_CLASS_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTCLAS2", "**", 0, True, "V", 3)

        End With

        grdICTCLAS2.DataSource = dst.Tables("ICTCLAS2")

        Create_Summary(grdICTCLAS2, "CUST_COUNT", "Count")
        grdICTCLAS2.Visible = (ASCMAIN1.CLIENT = "RGI")

        lblDISC_CODE.Visible = (ASCMAIN1.CLIENT = "RGI")
        txtDISC_CODE.Visible = (ASCMAIN1.CLIENT = "RGI")
        txtDISC_DESC.Visible = (ASCMAIN1.CLIENT = "RGI")

        grpStyleCodeGeneration.Visible = (ASCMAIN1.CLIENT = "NYA")

        chkSTYLE_CLASS_RELEASE_ATONCE.Visible = (ASCMAIN1.CLIENT = "RGI")
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTCLAS2, "SS", "Show Filter", "Show GroupBox")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

    End Sub
#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
            Case "Edit"

            Case "Update"

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = "STYLE_CLASS_CODE = '" & Absx1.txtFor("STYLE_CLASS_CODE").Text & "'"
        Update_Record_TDA("ICTCLAS2", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("ICTCLAS2", New String() {Absx1.txtFor("STYLE_CLASS_CODE").Text})
        Sort_grdColumns(grdICTCLAS2, "CUST_COUNT")
        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("ICTCLAS2").Rows.Clear()
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdICTCLAS2.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTCLAS2}
            With grd.DisplayLayout.Override
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        Next
    End Sub

#End Region

#Region "grdICTCLAS2"
    Private Sub grdICTCLAS2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTCLAS2.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("STYLE_CLASS_CODE").Value = Absx1.txtFor("STYLE_CLASS_CODE").Text
        End If
    End Sub
#End Region

    Private Sub UltraTextEditor1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles UltraTextEditor1.ValueChanged

    End Sub

    Private Sub chkSTYLE_CLASS_RELEASE_ATONCE_CheckedChanged(sender As Object, e As EventArgs) Handles chkSTYLE_CLASS_RELEASE_ATONCE.CheckedChanged

    End Sub
End Class