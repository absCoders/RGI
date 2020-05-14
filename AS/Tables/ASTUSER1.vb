Public Class ASTUSER1
    Dim SECURITY_CODEs As New List(Of String)
    Dim rowASTPARMP As DataRow
    Dim USER_PASSWORD_orig As String

    ' LIST IF USERS IN ASTUSERX - WE PROB NEED SIMILAR FOR GROUPS

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select USER_ID, USER_NAME, USER_STATUS from ASTUSER1"
            Create_TDA(.Tables.Add, "ASTUSERX", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SECURITY_CODE from ASTSECM1 order by SECURITY_CODE"

            For Each row As DataRow In ASCDATA1.GetDataTable("", "ASTSECM1").Rows
                Dim SECURITY_CODE As String = row.Item("SECURITY_CODE")
                Dim dc As New DataColumn
                dc.ColumnName = SECURITY_CODE
                dc.DataType = GetType(System.Boolean)
                dc.DefaultValue = False
                dst.Tables("ASTUSERX").Columns.Add(dc)
                SECURITY_CODEs.Add(SECURITY_CODE)
            Next

            grdASTUSERX.DataSource = dst.Tables("ASTUSERX")

            For Each SECURITY_CODE As String In SECURITY_CODEs
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

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "SELECT ASTUSER1.USER_ID " _
                & ", CASE ISNULL(ASTUSER2.SECURITY_CODE,'') WHEN '' THEN '0' ELSE '1' END SEL " _
                & ", ASTSECM1.SECURITY_CODE, ASTSECM1.SECURITY_DESC " _
                & " FROM ASTUSER1, ASTSECM1 " _
                & " LEFT OUTER JOIN ASTUSER2 ON ASTUSER2.USER_ID = @PARM1" _
                & " AND ASTUSER2.SECURITY_CODE = ASTSECM1.SECURITY_CODE" _
                & " WHERE ASTUSER1.USER_ID = @PARM2"
            Else
                ASCMAIN1.sql = "SELECT ASTUSER1.USER_ID " _
                & ", DECODE(ASTUSER2.USER_ID,NULL,'0','1') SEL " _
                & ", ASTSECM1.SECURITY_CODE, ASTSECM1.SECURITY_DESC " _
                & " FROM ASTSECM1, ASTUSER2, ASTUSER1 " _
                & " WHERE ASTUSER2.USER_ID (+) = :PARM1 " _
                & " AND ASTUSER2.SECURITY_CODE (+) = ASTSECM1.SECURITY_CODE " _
                & " AND ASTUSER1.USER_ID = :PARM2"
            End If
            Create_TDA(.Tables.Add, "ASTUSER2", "**", 0, True, "VV", -1)



            ASCMAIN1.sql = "SELECT ASTUSER3.*, ASTUSER1.USER_NAME USER_GROUP_NAME, '1' SEL " _
            & " FROM ASTUSER3,ASTUSER1 where ASTUSER3.USER_GROUP_ID = ASTUSER1.USER_ID and ASTUSER1.USER_STATUS = 'G'"

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "SELECT ASTUSER1.USER_ID " _
                & ", CASE ISNULL(ASTUSER3.USER_GROUP_ID,'') WHEN '' THEN '0' ELSE '1' END SEL " _
                & ", ASTUSERG.USER_ID USER_GROUP_ID, ASTUSERG.USER_NAME USER_GROUP_NAME " _
                & " FROM ASTUSER1, ASTUSER1 ASTUSERG " _
                & " LEFT OUTER JOIN ASTUSER3 ON ASTUSER3.USER_ID = @PARM1" _
                & " AND ASTUSER3.USER_GROUP_ID = ASTUSERG.USER_ID" _
                & " WHERE ASTUSER1.USER_ID = @PARM2"
            Else
                ASCMAIN1.sql = "SELECT ASTUSER1.USER_ID " _
                & ", DECODE(ASTUSER3.USER_ID,NULL,'0','1') SEL " _
                & ", ASTUSERG.USER_ID USER_GROUP_ID, ASTUSERG.USER_NAME USER_GROUP_NAME" _
                & " FROM ASTUSER1 ASTUSERG, ASTUSER3, ASTUSER1 " _
                & " WHERE ASTUSER3.USER_ID (+) = :PARM1 " _
                & " AND ASTUSER3.USER_GROUP_ID (+) = ASTUSERG.USER_ID " _
                & " AND ASTUSERG.USER_STATUS = 'G'" _
                & " AND ASTUSER1.USER_STATUS <> 'G'" _
                & " AND ASTUSER1.USER_ID = :PARM2"
            End If
            Create_TDA(.Tables.Add, "ASTUSER3", "**", 0, True, "VV", -1)




            ASCMAIN1.sql = "SELECT ASTUSER4.*, GLTCOMP1.COMPANY_NAME, '1' SEL " _
            & " FROM ASTUSER4,GLTCOMP1 where GLTCOMP1.COMPANY_CODE = ASTUSER4.COMPANY_CODE"

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "SELECT ASTUSER1.USER_ID " _
                & ", CASE ISNULL(ASTUSER4.USER_ID,'') WHEN '' THEN '0' ELSE '1' END SEL " _
                & ", GLTCOMP1.COMPANY_CODE, GLTCOMP1.COMPANY_NAME " _
                & " FROM ASTUSER1, GLTCOMP1 " _
                & " LEFT OUTER JOIN ASTUSER4 ON ASTUSER4.USER_ID = @PARM1" _
                & " AND ASTUSER4.COMPANY_CODE = GLTCOMP1.COMPANY_CODE" _
                & " WHERE ASTUSER1.USER_ID = @PARM2"
            Else
                ASCMAIN1.sql = "SELECT ASTUSER1.USER_ID " _
                & ", DECODE(ASTUSER4.USER_ID,NULL,'0','1') SEL " _
                & ", GLTCOMP1.COMPANY_CODE, GLTCOMP1.COMPANY_NAME " _
                & " FROM GLTCOMP1, ASTUSER4, ASTUSER1 " _
                & " WHERE ASTUSER4.USER_ID (+) = :PARM1 " _
                & " AND ASTUSER4.COMPANY_CODE (+) = GLTCOMP1.COMPANY_CODE " _
                & " AND ASTUSER1.USER_ID = :PARM2"
            End If
            Create_TDA(.Tables.Add, "ASTUSER4", "**", 0, True, "VV", -1)


            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "SELECT CONVERT(DATE,INIT_DATE) STAT_DATE " _
                & " FROM ASTOPST1 where ASTOPST1.USER_ID = @PARM1 AND ASTOPST1.INIT_DATE >= @PARM2 AND ASTOPST1.INIT_DATE -1 <= @PARM3 GROUP BY CONVERT(DATE,INIT_DATE)"
            Else
                ASCMAIN1.sql = "SELECT DISTINCT TRUNC(INIT_DATE) STAT_DATE" _
                & " FROM ASTOPST1 where USER_ID = :PARM1 AND ASTOPST1.INIT_DATE >= :PARM2 AND ASTOPST1.INIT_DATE -1 <= :PARM3 GROUP BY TRUNC(INIT_DATE)" '  and SELECTION_NO = 0 
            End If
            Create_TDA(.Tables.Add, "ASTOPST0", "**", 0, False, "VDD", 1)

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "SELECT CONVERT(DATE,INIT_DATE) STAT_DATE, ASTOPST1.SESSION_NO " _
                & " , ASTOPST1.INIT_DATE, ASTOPST1.LAST_DATE " _
                & " FROM ASTOPST1 " _
                & " where ASTOPST1.USER_ID = @PARM1 AND ASTOPST1.INIT_DATE >= @PARM2 AND ASTOPST1.INIT_DATE -1 <= @PARM3" '  and SELECTION_NO = 0
            Else
                ASCMAIN1.sql = "SELECT TRUNC(INIT_DATE) STAT_DATE, ASTOPST1.SESSION_NO " _
                & " , MIN (ASTOPST1.INIT_DATE) INIT_DATE, MAX (ASTOPST1.LAST_DATE) LAST_DATE " _
                & " FROM ASTOPST1 " _
                & " where ASTOPST1.USER_ID = :PARM1 AND ASTOPST1.INIT_DATE >= :PARM2 AND ASTOPST1.INIT_DATE -1 <= :PARM3 GROUP BY TRUNC(INIT_DATE), ASTOPST1.SESSION_NO" '  and SELECTION_NO = 0
            End If
            Create_TDA(.Tables.Add, "ASTOPST1", "**", 0, False, "VDD", 0)

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "SELECT CONVERT(DATE,INIT_DATE) STAT_DATE, ASTOPST1.* " _
                & " , ASTMENU1.MENU_ITEM_DESC " _
                & " FROM ASTOPST1,ASTMENU1 " _
                & " where ASTOPST1.MENU_ID = ASTMENU1.MENU_ID (+) " _
                & "   and ASTOPST1.MENU_ITEM_TYPE = ASTMENU1.MENU_ITEM_TYPE (+) " _
                & "   and ASTOPST1.MENU_ITEM_OBJECT = ASTMENU1.MENU_ITEM_OBJECT (+) " _
                & "   and ASTOPST1.USER_ID = @PARM1 AND ASTOPST1.INIT_DATE >= @PARM2 AND ASTOPST1.INIT_DATE -1 <= @PARM3 and SELECTION_NO <> 0"
            Else
                ASCMAIN1.sql = "SELECT TRUNC(INIT_DATE) STAT_DATE, ASTOPST1.* " _
                & " , ASTMENU1.MENU_ITEM_DESC " _
                & " FROM ASTOPST1,ASTMENU1 " _
                & " where ASTOPST1.MENU_ID = ASTMENU1.MENU_ID (+) " _
                & "   and ASTOPST1.MENU_ITEM_TYPE = ASTMENU1.MENU_ITEM_TYPE (+) " _
                & "   and ASTOPST1.MENU_ITEM_OBJECT = ASTMENU1.MENU_ITEM_OBJECT (+) " _
                & "   and ASTOPST1.USER_ID = :PARM1 AND ASTOPST1.INIT_DATE >= :PARM2 AND ASTOPST1.INIT_DATE -1 <= :PARM3 and SELECTION_NO <> 0"
            End If
            Create_TDA(.Tables.Add, "ASTOPSTX", "**", 0, False, "VDD", 0)

            Create_Relation("ASTOPST0", "ASTOPST1", "STAT_DATE")
            Create_Relation("ASTOPST1", "ASTOPSTX", "STAT_DATE,SESSION_NO")

            '.Relations.Add("ASTOPST1", _
            'New DataColumn() {.Tables("ASTOPST0").Columns("STAT_DATE")}, _
            'New DataColumn() {.Tables("ASTOPST1").Columns("STAT_DATE")})

            '.Relations.Add("ASTOPSTX", _
            'New DataColumn() {.Tables("ASTOPST1").Columns("SESSION_NO")}, _
            'New DataColumn() {.Tables("ASTOPSTX").Columns("SESSION_NO")})

            .Tables("ASTOPST0").Columns.Add("SESSIONS", GetType(System.Int64), "COUNT(CHILD.SESSION_NO)")
            .Tables("ASTOPST1").Columns.Add("SELECTIONS", GetType(System.Int64), "COUNT(CHILD.SELECTION_NO)")
            .Tables("ASTOPST0").Columns.Add("SELECTIONS", GetType(System.Int64), "SUM(CHILD.SELECTIONS)")


            ASCMAIN1.sql = "SELECT X.*, Y.MENU_ITEM_DESC from" _
                & " (Select MENU_ITEM_OBJECT, MIN (MENU_ITEM_DESC) MENU_ITEM_DESC" _
                & " from ASTMENU1 group by MENU_ITEM_OBJECT) Y," _
                & " (Select MENU_ITEM_OBJECT, COUNT (*) RUNS, MIN (INIT_DATE) INIT_DATE, MAX (INIT_DATE) LAST_DATE" _
                & ", MIN (MENU_ID) MENU_ID1, MAX (MENU_ID) MENU_ID2" _
                & " from ASTOPST1" _
                & " where USER_ID = :PARM1 and INIT_DATE > :PARM2 and INIT_DATE < :PARM2" _
                & " group by MENU_ITEM_OBJECT) X where Y.MENU_ITEM_OBJECT (+) = X.MENU_ITEM_OBJECT" _
                & " and X.MENU_ITEM_OBJECT is Not Null"
            Create_TDA(.Tables.Add, "ASTOPSTF", "**", 0, False, "VDD", 1)
            .Tables("ASTOPSTF").Columns("RUNS").DataType = GetType(System.Int64)
        End With

        grdASTUSER2.DataSource = dst.Tables("ASTUSER2")
        grdASTUSER3.DataSource = dst.Tables("ASTUSER3")
        grdASTUSER4.DataSource = dst.Tables("ASTUSER4")

        grdASTOPST1.DataSource = dst.Tables("ASTOPST0")
        grdASTOPSTF.DataSource = dst.Tables("ASTOPSTF")

        With grdASTOPST1.DisplayLayout
            .Bands(0).SortedColumns.Clear()
            .Bands(0).SortedColumns.Add(.Bands(0).Columns("STAT_DATE"), True)
            .Bands(1).SortedColumns.Clear()
            .Bands(1).SortedColumns.Add(.Bands(1).Columns("SESSION_NO"), False)
            .Bands(2).SortedColumns.Clear()
            .Bands(2).SortedColumns.Add(.Bands(2).Columns("INIT_DATE"), False)

        End With
        grdASTOPST1.DisplayLayout.Override.HeaderPlacement = UltraWinGrid.HeaderPlacement.OncePerRowIsland

        rowASTPARMP = ASCDATA1.GetDataRow("Select * from ASTPARMP where AS_PARM_KEY = 'Z'")

        With grdASTUSERX.DisplayLayout.Bands(0)
            .Columns("USER_ID").Header.Fixed = True
            .Columns("USER_NAME").Header.Fixed = True
            .Columns("USER_STATUS").Header.Fixed = True
        End With

        grdASTUSER2.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.Default

        dte1.Value = Now.Date.AddDays(-90)
        dte2.Value = Now.Date

        ReParent_Tabs(tabASTUSER1)
    End Sub

    Private Sub Form_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If SELECTION_NO = 0 Then Exit Sub
    End Sub


#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdASTUSERX, "SS", "Show Filter", "Show GroupBox")
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

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sql As String = ""

        sql = "Delete from ASTUSER2 where USER_ID = '" & Absx1.txtFor("USER_ID").Text & "'"
        ASCDATA1.ExecuteSQL(sql)
        dst.Tables("ASTUSER2").AcceptChanges()
        For Each row As DataRow In dst.Tables("ASTUSER2").Rows
            If row.Item("SEL") = "1" Then
                row.SetAdded()
            End If
        Next
        Update_Record_TDA("ASTUSER2")

        sql = "Delete from ASTUSER3 where USER_ID = '" & Absx1.txtFor("USER_ID").Text & "'"
        ASCDATA1.ExecuteSQL(sql)
        dst.Tables("ASTUSER3").AcceptChanges()
        If Absx1.optFor("USER_STATUS").Value <> "G" Then
            For Each row As DataRow In dst.Tables("ASTUSER3").Rows
                If row.Item("SEL") = "1" Then
                    row.SetAdded()
                End If
            Next
        End If
        Update_Record_TDA("ASTUSER3")

        sql = "Delete from ASTUSER4 where USER_ID = '" & Absx1.txtFor("USER_ID").Text & "'"
        ASCDATA1.ExecuteSQL(sql)
        dst.Tables("ASTUSER4").AcceptChanges()
        For Each row As DataRow In dst.Tables("ASTUSER4").Rows
            If row.Item("SEL") = "1" Then
                row.SetAdded()
            End If
        Next
        Update_Record_TDA("ASTUSER4")


        If rowASFBASE1.Item("USER_PASSWORD") & "" <> USER_PASSWORD_orig Then
            Dim USER_PASSWORD As String = rowASFBASE1.Item("USER_PASSWORD") & ""
            If rowASTPARMP.Item("AS_PARM_PWD_ENCRYPTED") & "" = "1" Then
                Dim MD5 As New ASCSCMD5
                USER_PASSWORD = MD5.DigestStrToHexStr(USER_PASSWORD)
                rowASFBASE1.Item("USER_PASSWORD") = USER_PASSWORD
                MD5 = Nothing
            End If
            rowASFBASE1.Item("USER_PASSWORD_LAST_DATE") = DATETIME_STAMP
        End If

    End Sub

    Overrides Sub Show_Record_Special()

        EnforceConstraints(False)
        Get_ASTOPST1()

        Fill_Records("ASTUSER2", New String() {Absx1.txtFor("USER_ID").Text, Absx1.txtFor("USER_ID").Text})
        Sort_grdColumns(grdASTUSER2, "SECURITY_CODE")
        If EntryMode = "New" Then
            ASCMAIN1.sql = "Select * from ASTSECM1 order by SECURITY_CODE"
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim rowASTUSER2 As DataRow = dst.Tables("ASTUSER2").NewRow
                rowASTUSER2.Item("USER_ID") = Absx1.txtFor("USER_ID").Text
                rowASTUSER2.Item("SEL") = "0"
                rowASTUSER2.Item("SECURITY_CODE") = row.Item("SECURITY_CODE")
                rowASTUSER2.Item("SECURITY_DESC") = row.Item("SECURITY_DESC")
                dst.Tables("ASTUSER2").Rows.Add(rowASTUSER2)
            Next
        End If

        Fill_Records("ASTUSER3", New String() {Absx1.txtFor("USER_ID").Text, Absx1.txtFor("USER_ID").Text})
        Sort_grdColumns(grdASTUSER3, "USER_GROUP_ID")
        If EntryMode = "New" Then
            ASCMAIN1.sql = "Select * from ASTUSER1 where USER_STATUS = 'G' order by USER_ID"
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim rowASTUSER3 As DataRow = dst.Tables("ASTUSER3").NewRow
                rowASTUSER3.Item("USER_ID") = Absx1.txtFor("USER_ID").Text
                rowASTUSER3.Item("SEL") = "0"
                rowASTUSER3.Item("USER_GROUP_ID") = row.Item("USER_ID")
                rowASTUSER3.Item("USER_GROUP_NAME") = row.Item("USER_NAME")
                dst.Tables("ASTUSER3").Rows.Add(rowASTUSER3)
            Next
        End If

        Fill_Records("ASTUSER4", New String() {Absx1.txtFor("USER_ID").Text, Absx1.txtFor("USER_ID").Text})
        Sort_grdColumns(grdASTUSER4, "COMPANY_CODE")
        If EntryMode = "New" Then
            ASCMAIN1.sql = "Select * from GLTCOMP1 order by COMPANY_CODE"
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim rowASTUSER4 As DataRow = dst.Tables("ASTUSER4").NewRow
                rowASTUSER4.Item("USER_ID") = Absx1.txtFor("USER_ID").Text
                rowASTUSER4.Item("SEL") = "0"
                rowASTUSER4.Item("COMPANY_CODE") = row.Item("COMPANY_CODE")
                rowASTUSER4.Item("COMPANY_NAME") = row.Item("COMPANY_NAME")
                dst.Tables("ASTUSER4").Rows.Add(rowASTUSER4)
            Next
        End If

        If dst.Tables("ASTUSER4").Select("SEL='1'").Length = 0 Then
            chkAllCompanies.Checked = True
        Else
            chkAllCompanies.Checked = False
        End If

        USER_PASSWORD_orig = Absx1.txtFor("USER_PASSWORD").Text

        EnforceConstraints(True)
    End Sub

    Sub Load_Report_Form(ByVal FORM_NAME As String)

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            'dst.EnforceConstraints = False
            'dst.Tables("ASTOPST0").Rows.Clear()
            'dst.Tables("ASTOPST1").Rows.Clear()
            'dst.Tables("ASTOPSTX").Rows.Clear()
            'dst.EnforceConstraints = True
            Fill_ASTUSERX()
        End If
    End Sub

    Sub Fill_ASTUSERX()
        Fill_Records("ASTUSERX")
        Sort_grdColumns(grdASTUSERX, "USER_ID")

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
                Dim USER_ID As String = Absx1.txtFor("USER_ID").Text

                If USER_ID.Length > 0 Then
                    If USER_ID <> USER_ID.ToLower Then
                        EMsg &= vbCr & "User ID should use lowercase letters only"
                    Else
                        For i As Int16 = 1 To USER_ID.Length
                            Dim z As String = USER_ID.Substring(i - 1, 1)
                            If z < "a" Or z > "z" Then
                                If InStr("0123456789.", z) = 0 Then
                                    EMsg &= vbCr & "User ID should use lowercase letters and numbers only"
                                End If
                            End If
                        Next
                    End If
                End If

            Case "Update"
                If Absx1.optFor("USER_STATUS").Value & "" = "G" Then
                Else
                    Dim password_error_checks As String = _
                    ASCMAIN1.Validate_User_Password( _
                    False, _
                    Absx1.txtFor("USER_ID").Text, _
                    Absx1.txtFor("USER_PASSWORD").Text, _
                    rowASTPARMP)

                    If Not chkAllCompanies.Checked Then
                        If dst.Tables("ASTUSER4").Select("SEL='1'").Length = 0 Then
                            EMsg &= vbCr & "You must select at least 1 company if not granting All Companies"
                        End If
                    End If


                    If password_error_checks <> "" Then
                        EMsg &= vbCr & "Password Errors:" & vbCr & vbTab & Replace(password_error_checks, vbCr, vbCr & vbTab)
                    End If
                End If

                If EMsg = "" Then
                    If ASCMAIN1.CLIENT = "VAN" Then
                        Dim USER_INIT As String = Absx1.txtFor("USER_INIT").Text
                        If USER_INIT = "" Then
                            Dim USER_NAMEs() As String = Absx1.txtFor("USER_NAME").Text.Split(" ")
                            If USER_NAMEs.Length = 1 Then
                                USER_INIT = USER_NAMEs(0).Substring(0, 1)
                            Else
                                USER_INIT = USER_NAMEs(0).Substring(0, 1) & USER_NAMEs(1).Substring(0, 1)
                            End If
                            rowASFBASE1.Item("USER_INIT") = USER_INIT
                            'Absx1.txtFor("USER_INIT").Text = USER_INIT
                        End If

                    End If
                End If
        End Select

    End Sub
#End Region

    Private Sub grdASTUSERX_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdASTUSERX.DoubleClickCell
        If grdASTUSERX.ActiveCell Is Nothing Then
            Exit Sub
        End If
        Absx1.txtFor("USER_ID").Text = grdASTUSERX.ActiveCell.Row.Cells("USER_ID").Text
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        Call Click_Command("Edit")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cmdFetch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFetch.Click
        EnforceConstraints(False)
        Get_ASTOPST1()
        EnforceConstraints(True)
    End Sub

    Sub Get_ASTOPST1()

        Dim caption As String = String.Format _
        ("Operator Statistics for {0} for the {1} days from {2} to {3}", _
         Absx1.txtFor("USER_ID").Text, _
         CStr(1 + DateDiff("d", dte1.DateTime, dte2.DateTime)), _
         Format(dte1.Value, "MM/dd/yyyy"), _
         Format(dte2.Value, "MM/dd/yyyy"))

        grdASTOPST1.Text = caption

        caption = String.Format _
            ("Function Statistics for {0} for the {1} days from {2} to {3}", _
             Absx1.txtFor("USER_ID").Text, _
             CStr(1 + DateDiff("d", dte1.DateTime, dte2.DateTime)), _
             Format(dte1.Value, "MM/dd/yyyy"), _
             Format(dte2.Value, "MM/dd/yyyy"))

        grdASTOPSTF.Text = caption

        Fill_Records("ASTOPST0", New Object() {Absx1.txtFor("USER_ID").Text, dte1.Value, dte2.Value})
        Fill_Records("ASTOPST1", New Object() {Absx1.txtFor("USER_ID").Text, dte1.Value, dte2.Value})
        Fill_Records("ASTOPSTX", New Object() {Absx1.txtFor("USER_ID").Text, dte1.Value, dte2.Value})
        Fill_Records("ASTOPSTF", New Object() {Absx1.txtFor("USER_ID").Text, dte1.Value, dte2.Value})

        Sort_grdColumns(grdASTOPST1, "STAT_DATE".ToLower)
        Sort_grdColumns(grdASTOPSTF, "MENU_ITEM_OBJECT")

    End Sub

    Private Sub chkAllCompanies_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAllCompanies.CheckedChanged
        grdASTUSER4.Enabled = Not chkAllCompanies.Checked
        For Each row As DataRow In dst.Tables("ASTUSER4").Rows
            row.Item("SEL") = "0"
        Next
    End Sub
End Class