Public Class ASTUSER1
    Dim SECURITY_CODEs As New List(Of String)

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select USER_ID, USER_NAME, USER_STATUS from ASTUSER1"
            '.Tables.Add(ASCDATA1.GetDataTable("", "ASTUSERX"))
            Create_TDA(.Tables.Add, "ASTUSERX", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SECURITY_CODE from ASTSECM1 order by SECURITY_CODE"

            For Each row As DataRow In ASCDATA1.GetDataTable("", "ASTSECM1").Rows
                Dim SECURITY_CODE As String = row.Item("SECURITY_CODE")
                Dim dc As New DataColumn
                dc.ColumnName = SECURITY_CODE
                dc.DataType = GetType(System.Boolean)
                'dc.MaxLength = 1
                dc.DefaultValue = False
                dst.Tables("ASTUSERX").Columns.Add(dc)
                SECURITY_CODEs.Add(SECURITY_CODE)
            Next


            grdASTUSERX.DataSource = dst.Tables("ASTUSERX")

            For Each SECURITY_CODE As String In SECURITY_CODEs
                'grdASTUSERX.DisplayLayout.Bands(0).Columns(SECURITY_CODE).Hidden = False
                With grdASTUSERX.DisplayLayout.Bands(0).Columns(SECURITY_CODE)

                    .Width = 50
                    .Style = UltraWinGrid.ColumnStyle.CheckBox
                    .CellAppearance.TextHAlign = HAlign.Center
                    .Header.Appearance.TextHAlign = HAlign.Center

                End With
            Next

            Call Fill_ASTUSERX()

            ASCMAIN1.sql = "SELECT ASTUSER2.*, ASTSECM1.SECURITY_DESC, '1' SEL " _
            & " FROM ASTUSER2,ASTSECM1 where ASTSECM1.SECURITY_CODE = ASTUSER2.SECURITY_CODE"

            ASCMAIN1.sql = "SELECT ASTUSER1.USER_ID " _
            & ", DECODE(ASTUSER2.USER_ID,NULL,'0','1') SEL " _
            & ", ASTSECM1.SECURITY_CODE, ASTSECM1.SECURITY_DESC " _
            & " FROM ASTSECM1, ASTUSER2, ASTUSER1 " _
            & " WHERE ASTUSER2.USER_ID (+) = :PARM1 " _
            & " AND ASTUSER2.SECURITY_CODE (+) = ASTSECM1.SECURITY_CODE " _
            & " AND ASTUSER1.USER_ID = :PARM2"
            Create_TDA(.Tables.Add, "ASTUSER2", "**", 0, True, "VV", -1)

            ASCMAIN1.sql = "SELECT TRUNC(INIT_DATE) STAT_DATE, COUNT (*) SELECTIONS " _
            & " FROM ASTOPST1 where USER_ID = :PARM1 GROUP BY TRUNC(INIT_DATE)"
            Create_TDA(.Tables.Add, "ASTOPST0", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "SELECT TRUNC(INIT_DATE) STAT_DATE, ASTOPST1.SESSION_NO " _
            & " , ASTOPST1.INIT_DATE, ASTOPST1.LAST_DATE " _
            & " FROM ASTOPST1 " _
            & " where ASTOPST1.USER_ID = :PARM1 and SELECTION_NO = 0"
            Create_TDA(.Tables.Add, "ASTOPST1", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "SELECT TRUNC(INIT_DATE) STAT_DATE, ASTOPST1.* " _
            & " , ASTMENU1.MENU_ITEM_DESC " _
            & " FROM ASTOPST1,ASTMENU1 " _
            & " where ASTOPST1.MENU_ID = ASTMENU1.MENU_ID (+) " _
            & "   and ASTOPST1.MENU_ITEM_TYPE = ASTMENU1.MENU_ITEM_TYPE (+) " _
            & "   and ASTOPST1.MENU_ITEM_OBJECT = ASTMENU1.MENU_ITEM_OBJECT (+) " _
            & "   and ASTOPST1.USER_ID = :PARM1 and SELECTION_NO <> 0"
            Create_TDA(.Tables.Add, "ASTOPST2", "**", 0, False, "V", 0)

            .Relations.Add("ASTOPST1", _
            New DataColumn() {.Tables("ASTOPST0").Columns("STAT_DATE")}, _
            New DataColumn() {.Tables("ASTOPST1").Columns("STAT_DATE")})

            .Relations.Add("ASTOPST2", _
            New DataColumn() {.Tables("ASTOPST1").Columns("STAT_DATE"), .Tables("ASTOPST1").Columns("SESSION_NO")}, _
            New DataColumn() {.Tables("ASTOPST2").Columns("STAT_DATE"), .Tables("ASTOPST2").Columns("SESSION_NO")})
        End With

        grdASTUSER2.DataSource = dst.Tables("ASTUSER2")

        grdASTOPST1.DataSource = dst.Tables("ASTOPST0")

        grdASTOPST1.DisplayLayout.Bands("ASTOPST0").SortedColumns.Clear()
        grdASTOPST1.DisplayLayout.Bands("ASTOPST0").SortedColumns.Add(grdASTOPST1.DisplayLayout.Bands("ASTOPST0").Columns("STAT_DATE"), True)
        grdASTOPST1.DisplayLayout.Bands("ASTOPST1").SortedColumns.Clear()
        grdASTOPST1.DisplayLayout.Bands("ASTOPST1").SortedColumns.Add(grdASTOPST1.DisplayLayout.Bands("ASTOPST1").Columns("SESSION_NO"), False)
        grdASTOPST1.DisplayLayout.Bands("ASTOPST2").SortedColumns.Clear()
        grdASTOPST1.DisplayLayout.Bands("ASTOPST2").SortedColumns.Add(grdASTOPST1.DisplayLayout.Bands("ASTOPST2").Columns("INIT_DATE"), False)

        'grdASTUSER2.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
    End Sub

    Private Sub Form_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        'grdASTUSERX.Left = lblUSER_PASSWORD.Left
        'grdASTUSERX.Top = lblUSER_PASSWORD.Top
        grdASTUSERX.Dock = DockStyle.Bottom


        'For Each SECURITY_CODE As String In SECURITY_CODEs
        '    'grdASTUSERX.DisplayLayout.Bands(0).Columns(SECURITY_CODE).Hidden = False
        '    With grdASTUSERX.DisplayLayout.Bands(0).Columns(SECURITY_CODE)

        '        .Width = 50
        '        .Style = UltraWinGrid.ColumnStyle.CheckBox
        '        .CellAppearance.TextHAlign = HAlign.Center
        '        .Header.Appearance.TextHAlign = HAlign.Center

        '    End With
        'Next


    End Sub


#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()
        'Update_Record_TDA("ASTUSER2")
        Dim sql As String = "Delete from ASTUSER2 where USER_ID = '" & Absx1.txtFor("USER_ID").Text & "'"
        ASCDATA1.ExecuteSQL(sql)
        dst.Tables("ASTUSER2").AcceptChanges()
        For Each row As DataRow In dst.Tables("ASTUSER2").Rows
            If row.Item("SEL") = "1" Then
                row.SetAdded()
            End If
        Next
        Update_Record_TDA("ASTUSER2")
    End Sub

    Overrides Sub Show_Record_Special()

        dst.EnforceConstraints = False
        Call Fill_Records("ASTOPST0", Absx1.txtFor("USER_ID").Text)
        Call Fill_Records("ASTOPST1", Absx1.txtFor("USER_ID").Text)
        Call Fill_Records("ASTOPST2", Absx1.txtFor("USER_ID").Text)

        Call Fill_Records("ASTUSER2", New String() {Absx1.txtFor("USER_ID").Text, Absx1.txtFor("USER_ID").Text})
        If EntryMode = "New" Then
            ASCMAIN1.sql = "Select * from ASTSECM1 order by SECURITY_CODE"
            For Each row As DataRow In ASCDATA1.GetDataTable("", "ASTSECM1").Rows
                Dim rowASTUSER2 As DataRow = dst.Tables("ASTUSER2").NewRow
                rowASTUSER2.Item("USER_ID") = Absx1.txtFor("USER_ID").Text
                rowASTUSER2.Item("SEL") = "0"
                rowASTUSER2.Item("SECURITY_CODE") = row.Item("SECURITY_CODE")
                rowASTUSER2.Item("SECURITY_DESC") = row.Item("SECURITY_DESC")
                dst.Tables("ASTUSER2").Rows.Add(rowASTUSER2)
            Next
        End If

        'dst.EnforceConstraints = True

        'Call Clear_Record_Special()
        'Call Load_Report_Form(txtctl.Text)
    End Sub

    Sub Load_Report_Form(ByVal FORM_NAME As String)

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            'dst.EnforceConstraints = False
            'dst.Tables("ASTOPST0").Rows.Clear()
            'dst.Tables("ASTOPST1").Rows.Clear()
            'dst.Tables("ASTOPST2").Rows.Clear()
            'dst.EnforceConstraints = True
            Call Fill_ASTUSERX()
        End If
    End Sub

    Sub Fill_ASTUSERX()
        Call Fill_Records("ASTUSERX")

        ASCMAIN1.sql = "Select * from ASTUSER2"
        For Each row As DataRow In ASCDATA1.GetDataTable("", "ASTUSER2").Rows
            Dim SECURITY_CODE As String = row.Item("SECURITY_CODE")
            Dim USER_ID As String = row.Item("USER_ID")
            Dim rowASTUSERX As DataRow = dst.Tables("ASTUSERX").Rows.Find(USER_ID)
            If SECURITY_CODEs.Contains(SECURITY_CODE) Then
                If Not rowASTUSERX Is Nothing Then
                    rowASTUSERX.Item(SECURITY_CODE) = True
                End If
            End If
        Next
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdASTUSERX.Visible = Not tf
        UltraTabControl1.Visible = tf
    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"

        End Select

    End Sub
#End Region

    Private Sub grdASTUSERX_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdASTUSERX.DoubleClickCell
        Absx1.txtFor("USER_ID").Text = grdASTUSERX.ActiveCell.Row.Cells("USER_ID").Text
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        Call Click_Command("Edit")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
End Class