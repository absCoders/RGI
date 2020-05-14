Imports System.Net.Mail
Imports Microsoft.Office.Interop.Outlook

Public Class ICTFACT1

    Dim sqlICTFACT2 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("APTPARM1")

        With dst
            ASCMAIN1.sql = "Select ICTFACT2.*, ICTBODY2.SUB_BODY_DESC" & vbCrLf _
                & " from ICTFACT2,ICTBODY2" & vbCrLf _
                & " where ICTFACT2.FACTORY_CODE = :PARM1" & vbCrLf _
                & "   and ICTBODY2.SUB_BODY_CODE = ICTFACT2.SUB_BODY_CODE"
            Create_TDA(.Tables.Add, "ICTFACT2", "**", 0, True, "V", 2)

            Create_TDA(.Tables.Add, "APTVEND1", "*")
        End With

        grdICTFACT2.DataSource = dst.Tables("ICTFACT2")

        With grdICTFACT2.DisplayLayout.Bands(0)
            .Columns("SUB_BODY_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
        End With

        If ASCMAIN1.CLIENT = "RGI" Then
            grdICTFACT2.Visible = False
        End If
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTFACT2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
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

        Select Case grd.Name

            Case "grdICTFACT2"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Select Case e.Tool.Key
            Case "Add Codes"
                If grd.Name = "grdICTFACT2" Then
                    Add_Codes(grdICTFACT2, "ICTBODY2", "SUB_BODY_CODE", "Sub-Body Codes")
                End If
        End Select

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
        Dim FACTORY_CODE As String = Absx1.txtFor("FACTORY_CODE").Text
        Dim sqlDelete = "FACTORY_CODE = '" & FACTORY_CODE & "'"
        Update_Record_TDA("ICTFACT2", sqlDelete)

        'If ASCMAIN1.CLIENT = "VAN" Then
        '    Update_Record_TDA("APTVEND1")
        'End If
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("ICTFACT2", New String() {Absx1.txtFor("FACTORY_CODE").Text})
        Sort_grdColumns(grdICTFACT2, "SUB_BODY_CODE")
        EnforceConstraints(True)

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"ICTFACT2", "APTVEND1"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdICTFACT2.Enabled = tf

        If ASCMAIN1.CLIENT = "VAN" Then
            cmdCreateVendor.Visible = False
            If EntryMode = "Edit" Then
                If Absx1.txtFor("VEND_CODE").Text = "" Then
                    cmdCreateVendor.Visible = True
                End If
            End If
        End If
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTFACT2}
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

#Region "grdICTFACT2"

    Private Sub grdICTFACT2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTFACT2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "SUB_BODY_CODE"
                grdCodeDesc(grdICTFACT2, "ICTBODY2", "SUB_BODY_CODE", "SUB_BODY_DESC")
        End Select
    End Sub

    Private Sub grdICTFACT2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTFACT2.AfterRowActivate

        With grdICTFACT2.DisplayLayout.Bands(0).Columns("SUB_BODY_CODE")
            If grdICTFACT2.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdICTFACT2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdICTFACT2.AfterRowsDeleted

    End Sub

    Private Sub grdICTFACT2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTFACT2.AfterRowUpdate

    End Sub

    Private Sub grdICTFACT2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdICTFACT2.BeforeRowsDeleted
        'For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        '    Dim FACTORY_CODE As String = grow.Cells("FACTORY_CODE").Value
        '    Dim SUB_BODY_CODE As String = grow.Cells("SUB_BODY_CODE").Value
        '    Dim rowICTFACT2 As DataRow = dst.Tables("ICTFACT2").Rows.Find(New String() {FACTORY_CODE, SUB_BODY_CODE})
        '    If Not rowICTFACT2.RowState = DataRowState.Added Then
        '        MsgBox("Cannot Delete Previously Updated Colors", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        '        e.Cancel = True
        '    End If
        'Next
    End Sub

    Private Sub grdICTFACT2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTFACT2.BeforeRowUpdate

        Dim row As DataRow = LookUp("ICTBODY2", e.Row.Cells("SUB_BODY_CODE").Text)
        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("FACTORY_CODE").Value = Absx1.txtFor("FACTORY_CODE").Text
            'e.Row.Cells("ORDR_FORM_LNO").Value = Val(dst.Tables("ICTFACT2").Compute("MAX(ORDR_FORM_LNO)", "") & "") + 10
        End If

    End Sub

    Private Sub grdICTFACT2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTFACT2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "SUB_BODY_CODE"
                Dim sql_where As String = Get_List_of_Codes("ICTBODY2.SUB_BODY_CODE not in", "ICTFACT2", "SUB_BODY_CODE")
                grdClickCellButton(grdICTFACT2, sql_where, True)
        End Select
    End Sub

#End Region

    Private Sub cmdCreateVendor_Click(sender As Object, e As EventArgs) Handles cmdCreateVendor.Click
        If Absx1.txtFor("VEND_CODE").Text <> "" Then
            MsgBox("A Vendor has Already been Assigned to this Factory", MsgBoxStyle.OkOnly, "Cannot Create New Vendor")
            Exit Sub
        End If

        Dim FACTORY_CODE As String = Absx1.txtFor("FACTORY_CODE").Text
        Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", FACTORY_CODE)
        If rowAPTVEND1 IsNot Nothing Then
            MsgBox("A Vendor has Already been Created with Vendor Code " & FACTORY_CODE, MsgBoxStyle.OkOnly, "Cannot Create New Vendor")
            Exit Sub
        End If

        If MsgBox("Are you sure that you want to create a Vendor Record for Factory " & FACTORY_CODE, MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
            Exit Sub
        End If

        dst.Tables("APTVEND1").Rows.Clear()
        rowAPTVEND1 = dst.Tables("APTVEND1").NewRow
        rowAPTVEND1.Item("VEND_CODE") = FACTORY_CODE
        rowAPTVEND1.Item("VEND_NAME") = Absx1.txtFor("FACTORY_DESC").Text
        rowAPTVEND1.Item("VEND_STATUS") = "A"
        rowAPTVEND1.Item("VEND_TYPE") = "S"
        rowAPTVEND1.Item("TERM_CODE") = "01"
        rowAPTVEND1.Item("POST_CODE") = "TRADE"
        rowAPTVEND1.Item("ACCT_CODE") = "4410"
        rowAPTVEND1.Item("VEND_SEP_CHECKS") = "0"
        rowAPTVEND1.Item("BANK_CODE") = ROWs("APTPARM1").Item("AP_PARM_BANK_CODE")
        rowAPTVEND1.Item("VEND_ON_HOLD") = "0"
        rowAPTVEND1.Item("VEND_PYMT_METHOD_FIXED") = "0"
        rowAPTVEND1.Item("VEND_PYMT_METHOD") = "WIRE"
        rowAPTVEND1.Item("VEND_STOP_PURCHASE") = "0"
        rowAPTVEND1.Item("VEND_ALWAYS_TAKE_DISC") = "0"

        rowAPTVEND1.Item("VEND_DUE_FROM_INV_DATE") = "0"
        rowAPTVEND1.Item("LABEL_RESP_CODE_NOT") = "0"

        rowAPTVEND1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowAPTVEND1.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
        rowAPTVEND1.Item("LAST_OPER") = rowAPTVEND1.Item("INIT_OPER")
        rowAPTVEND1.Item("LAST_DATE") = rowAPTVEND1.Item("INIT_DATE")

        dst.Tables("APTVEND1").Rows.Add(rowAPTVEND1)

        Update_Record_TDA("APTVEND1")
        cmdCreateVendor.Visible = False

        MsgBox("Vendor " & FACTORY_CODE & " has been Created", MsgBoxStyle.OkOnly, "Verification")

        Absx1.txtFor("VEND_CODE").Text = FACTORY_CODE
        ' Absx1.txtFor("VEND_NAME").Text = Absx1.txtFor("FACTORY_DESC").Text
    End Sub

End Class