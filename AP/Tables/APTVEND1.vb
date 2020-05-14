Public Class APTVEND1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select APTVEND9.*, GLTACCT1.ACCT_DESC from APTVEND9,GLTACCT1 where GLTACCT1.ACCT_CODE = APTVEND9.ACCT_CODE and APTVEND9.VEND_CODE = :PARM1"
            Create_TDA(.Tables.Add, "APTVEND9", "**", 0, True, "V", 5)

            If ASCMAIN1.CLIENT = "RGI" Then
                ASCMAIN1.sql = "Select ICTLSTCV.*, ICTLSTC1.LIST_CALC_DESC from ICTLSTCV,ICTLSTC1 where ICTLSTC1.LIST_CALC_CODE = ICTLSTCV.LIST_CALC_CODE and ICTLSTCV.VEND_CODE = :PARM1"
                Create_TDA(.Tables.Add, "ICTLSTCV", "**", 0, True, "V", 2)
                .Tables("ICTLSTCV").Columns.Add("SEL")
                .Tables("ICTLSTCV").Columns("SEL").DefaultValue = "0"

                Create_TDA(.Tables.Add, "ICTLSTC1", "*", 0, False)
                Fill_Records("ICTLSTC1")
            End If

        End With

        grdAPTVEND9.DataSource = dst.Tables("APTVEND9")

        If ASCMAIN1.CLIENT = "RGI" Then
            grdICTLSTCV.DataSource = dst.Tables("ICTLSTCV")
        Else
            grdICTLSTCV.Visible = False
        End If


        Get_PARM("GLTPARM1")

        Set_SEGS(grdAPTVEND9, "APTVEND9")
        Create_Summary(grdAPTVEND9, "DIST_AMT")

        If ASCMAIN1.CLIENT = "RGI" Then
            Create_Summary(grdICTLSTCV, "SEL")
            Create_Summary(grdICTLSTCV, "LIST_CALC_DESC", "Count")
        End If


        If ASCMAIN1.DBS_SERVER = "EXP" Or ASCMAIN1.DBS_COMPANY = "EXP" Then
            UltraTabControl1.Tabs("Purchasing Information").Visible = False
        End If

    End Sub
    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "APTVEND1"
            E.COLUMN_NAME = "VEND_CODE"
            E.CODE_VALUE = Absx1.txtFor("VEND_CODE").Text
            E.DESC_VALUE = "Vendor"
            E.ATTACHMENT_NOTES = ""
            'E.RESTRICTIONS = "D"
            'E.READ_ONLY = True
        End If

        Return E
    End Function

    'Public Overrides Function Audit_Context() As Audit_Entity

    '    Dim E As New Audit_Entity
    '    If ScreenMode Then
    '        E.TABLE_NAME = "APTVEND1"
    '        E.KEY_VALUE = Absx1.txtFor("VEND_CODE").Text
    '        E.KEY_DESC = "Vendor"
    '    End If
    '    Return E
    'End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "APTVEND1"
        E.TABLE_KEY_CAPTION = "Vendor"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("VEND_CODE").Text
            E.TABLE_KEY_DESC = Absx1.txtFor("VEND_CODE").Text & " " & Absx1.txtFor("VEND_NAME").Text
            E.TABLE_KEY_locked = ScreenMode
        Else
            E.enabled = False
            E.read_only = True
            E.TABLE_KEY_locked = False
            E.TABLE_KEY = ""
        End If

        Return E
    End Function

#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sql As String = ""

        sql = "Delete from APTVEND9 where VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "'"
        Update_Record_TDA("APTVEND9", sql)

        If ASCMAIN1.CLIENT = "RGI" Then
            ASCDATA1.DeleteRows("ICTLSTCV", "ISNULL(SEL,'0')<>'1'")
            sql = "Delete from ICTLSTCV where VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "'"
            Update_Record_TDA("ICTLSTCV", sql)
        End If


    End Sub

    Overrides Sub Show_Record_Special()
        Dim txtctl As UltraWinEditors.UltraTextEditor
        txtctl = Absx1.txtFor("VEND_CODE")
        Clear_Record_Special()
        Load_Report_Form(txtctl.Text)
    End Sub

    Sub Load_Report_Form(ByVal VEND_CODE As String)

        EnforceConstraints(False)

        Fill_Records("APTVEND9", VEND_CODE)
        For Each r As DataRow In dst.Tables("APTVEND9").Rows
            'If r.Item("COLUMN_CAPTION") & "" = "" Then
            '    Dim rowASTDSQLK As DataRow = dst.Tables("ASTDSQLK").Rows.Find(r.Item("COLUMN_NAME"))
            '    If Not rowASTDSQLK Is Nothing Then
            '        r.Item("COLUMN_CAPTION") = rowASTDSQLK.Item("COLUMN_CAPTION")
            '    End If
            'End If
        Next

        If ASCMAIN1.CLIENT = "RGI" Then
            Fill_Records("ICTLSTCV", VEND_CODE)
            For Each row As DataRow In dst.Tables("ICTLSTCV").Select("")
                row.Item("SEL") = "1"
            Next
            For Each row As DataRow In dst.Tables("ICTLSTC1").Select("")
                Dim LIST_CALC_CODE As String = row.Item("LIST_CALC_CODE")
                Dim LIST_CALC_DESC As String = row.Item("LIST_CALC_DESC")
                If dst.Tables("ICTLSTCV").Rows.Find(New String() {LIST_CALC_CODE, VEND_CODE}) Is Nothing Then
                    dst.Tables("ICTLSTCV").Rows.Add(New String() {LIST_CALC_CODE, VEND_CODE, LIST_CALC_DESC, "0"})
                End If
            Next
            Sort_grdColumns(grdICTLSTCV, "LIST_CALC_DESC")
        End If

        EnforceConstraints(True)

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("APTVEND9").Rows.Clear()
            If ASCMAIN1.CLIENT = "RGI" Then
                dst.Tables("ICTLSTCV").Rows.Clear()
            End If

            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdAPTVEND9.Enabled = tf

        If ASCMAIN1.CLIENT = "RGI" Then
            grdICTLSTCV.Enabled = tf
        End If

        If tf And Not ASCMAIN1.USER_SECURITY_CODEs.Contains("AP") Then
            Set_Read_Only(UltraTabControl1.Tabs("Name && Address").TabPage, True)
            Set_Read_Only(UltraTabControl1.Tabs("Codes && Other Info").TabPage, True)
            Set_Read_Only(UltraTabControl1.Tabs("Payment Information").TabPage, True)
            Set_Read_Only_for_ctl(Absx1.optFor("VEND_STATUS"), True)
            Set_Read_Only_for_ctl(Absx1.txtFor("VEND_NAME"), True)
            UltraTabControl1.SelectedTab = UltraTabControl1.Tabs("Purchasing Information")
        End If

    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"
                If Absx1.chkFor("PO_EMAIL_CONFIRM").Checked And Absx1.optFor("PO_XMIT_VIA").Value = "M" Then
                    EMsg &= EMsg & "email confirmation not necessary when Transmitting PO's via email"
                End If
                'Stop
                'If grdAPTVEND9.ActiveRow.IsAddRow Then
                '    EMsg &= vbCr & "Data Remaining in Addrow of GL Distribution Template"
                'End If

                ' Added by edz on 01/23/2008 as per Maria
                Dim VEND_TAX_ID As String = Absx1.txtFor("VEND_TAX_ID").Text.Trim
                Dim VEND_1099_BOX As Int32 = Val(Absx1.numFor("VEND_1099_BOX").Value & "")
                Dim VEND_TAX_ID_TYPE As String = Absx1.optFor("VEND_TAX_ID_TYPE").Value

                Select Case VEND_TAX_ID.Length
                    Case 0
                        ' NOTHING
                        If VEND_1099_BOX > 0 Then
                            EMsg &= EMsg & "Tax ID in 1099 Reporting section must be 9 numeric values when providing a value in the 1099 Box."
                        End If

                    Case 9
                        Dim temp As Long = 0
                        Long.TryParse(VEND_TAX_ID, temp)
                        If temp = 0 Then
                            EMsg &= EMsg & "Tax ID in 1099 Reporting section must be 9 numeric values."
                        Else
                            If (VEND_1099_BOX < 1 Or VEND_1099_BOX > 14 Or VEND_1099_BOX = 11 Or VEND_1099_BOX = 12) Then
                                EMsg &= EMsg & "Box values in the 1099 Reporting section must be 1 - 10, 13, 14."
                            ElseIf VEND_TAX_ID_TYPE Is Nothing OrElse VEND_TAX_ID_TYPE.Length = 0 Then
                                EMsg &= EMsg & "When providing 1099 Reporting information you must select the type."
                            End If
                        End If

                    Case Else
                        EMsg &= EMsg & "Tax ID in 1099 Reporting section must be 9 numeric values."

                End Select

        End Select

    End Sub
#End Region

#Region "grdAPTVEND9"

    Private Sub grdAPTVEND9_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTVEND9.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "ACCT_CODE"
                Dim ACCT_CODE As String = e.Cell.Value & ""

                grdCodeDesc(grdAPTVEND9, "GLTACCT1", "ACCT_CODE", "ACCT_DESC")
                For i As Integer = 2 To 4
                    If e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Text = "" Then
                        e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Value = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                    End If
                Next
        End Select
    End Sub

    Private Sub grdAPTVEND9_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTVEND9.AfterExitEditMode
        With grdAPTVEND9
            Select Case .ActiveCell.Column.Key
                Case "ACCT_CODE"
                    Dim ACCT_CODE As String = .ActiveCell.Text
                    If ACCT_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(ACCT_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With

    End Sub

    Private Sub grdAPTVEND9_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTVEND9.AfterRowActivate
        With grdAPTVEND9
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .ActiveCell = grdAPTVEND9.ActiveRow.Cells("ACCT_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                '.DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                ' why cant we edit the acct code?
            End If
        End With
    End Sub

    Private Sub grdAPTVEND9_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdAPTVEND9.BeforeRowUpdate
        With grdAPTVEND9
            If e.Row.Cells("ACCT_CODE").Text = "" Then
                e.Cancel = True
            Else
                Call LookUp("GLTACCT1", e.Row.Cells("ACCT_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Acct Code (" & e.Row.Cells("ACCT_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            Dim COLUMN_NAME As String
            For i As Integer = 2 To 4
                COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
                If Not e.Row.Cells(COLUMN_NAME).Column.Hidden Then
                    If e.Row.Cells(COLUMN_NAME).Text = "" Then
                        e.Cancel = True
                    Else
                        Call LookUp("GLTSEGM1", New String() {CStr(i), e.Row.Cells(COLUMN_NAME).Text})
                        If cdr Is Nothing Then
                            MsgBox("Invalid Value entered for " & e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " (" & e.Row.Cells(COLUMN_NAME).Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                            e.Cancel = True
                        End If
                    End If
                End If
            Next

            If Not e.Cancel Then
                If e.Row.Cells("VEND_CODE").Text = "" Then
                    .ActiveRow.Cells("VEND_CODE").Value = Absx1.CtlFor("VEND_CODE").Text
                End If
            End If
        End With

    End Sub

    Private Sub grdAPTVEND9_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTVEND9.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdAPTVEND9, sql_where, sql_where <> "")
    End Sub

#End Region

End Class