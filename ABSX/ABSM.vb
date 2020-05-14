Imports System.Math

Public Class ABSM

#Region "Declarations"

    Protected SEQs As Integer       ' Number of Columns Selected to Sort
    Protected FORM_NAME As String   ' Form Object Name for Report (Me.Name)

    Protected SET_ID As String      ' Set used during this Execution
    Protected SET_DESC As String    ' Description of SET_ID

    Protected LIST_CODE As String = ""  ' Working variable for List of Codes
    Protected LIST_DESC As String = ""  ' Corresponding Description

    Protected Page0 As New ArrayList    ' Page 0 descriptive items
    Protected sql_SELECT_cols As String     ' PB Columns used in the Select List
    Protected sql_GROUP_BY_cols As String   ' PB Columns used in the Group By
    Protected sql_WHERE As String           ' PB where clause
    Protected sql_TABLE_NAMEs As String     ' PB Tables
    Protected sql_JOIN As String            ' PB Join
    Protected sql_TABLE_NAME As String      ' PB Primary Table for Data Source

    Protected COLUMN_NAMEs As New ArrayList         ' PB Column Names
    Protected COLUMN_CAPTIONs As New ArrayList      ' PB Column Captions
    Protected GROUP_ALL_OTHERSs As New ArrayList    ' Whether others should be grouped
    Dim PAGE_BREAKs As String                       ' PB Level Page Breaks
    Protected xErrMsg As String     ' if <> "" (error); exit Main Process in Sub Proceed
    Protected COLUMN_NAME_first As String = ""  ' Forced First Column of a PB Report
    Protected COLUMN_NAME_last As String = ""   ' Forced Last Column of a PB Report
    Protected COLUMN_NAMEs_appended As String = ""
    Protected COLUMN_NAME_RECAP_ROW_NO As String = ""
    Protected tblASTDSQLH As DataTable
    Protected tblASTDSQLS As DataTable
    Protected PB_Report As Boolean = False
    Protected tblASTRECAP As DataTable
    Protected Recap_Report As Boolean = False
    Protected ASTSRPT1 As String
    Protected ASTSRPT1_sum_columns As String
    Protected ASTSRPT1_sql_sum As String
    Protected COLUMN_NAME_sum As New Dictionary(Of String, String)

    Protected COLUMN_NAME_by_Lvl() As String
    Protected COLUMN_CAPTION_by_Lvl() As String
    Protected G_by_Lvl() As Integer
    Protected COLUMN_NAME_sum_first As String
    Protected DATA_TYPEs() As String

    Protected sql As String

    Protected tblASTGROUP As DataTable
    Protected tblASTSRPT0 As DataTable
    Protected tblASTSRPT1 As DataTable

    Protected adaASTROPT1 As OracleDataAdapter
    Protected tblASTROPT1 As New DataTable
    Protected tblASTDSQLA As DataTable
    Protected tblASTDSQLB As New DataTable
    Protected tblASTDSQLC As New DataTable
    Protected tblASTDSQLD As New DataTable
    Protected tblASTDSQLE As New DataTable
    Protected tblASTDSQLF As New DataTable
    Protected tblASTDSQLJ As New DataTable
#End Region

#Region "Initialization"

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        FORM_NAME = Me.Name

    End Sub

    'Private Sub Form_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
    '    If Me.Name <> "ASFSRPTM" Then
    '        Call Setup_grdSetup()
    '        Call Initialize_Form()
    '        Call Save_Settings("0000000000")
    '    End If
    'End Sub

    Sub Initialize_Form()
        Dim sql As String = "Select * from ASTROPT1 where FORM_NAME = '" & FORM_NAME & "'"
        adaASTROPT1 = ASCDATA1.GetDataAdapter(tblASTROPT1, "ASTROPT1", sql, True, -1, False, 0)
        grdASTROPT1.DataSource = tblASTROPT1

        Call Show_Settings()
        'Call Mode_Settings(False)

    End Sub
#End Region

#Region "grdSetup"

    Sub Clear_grdSetup(ByVal Clear_All As Boolean)
        grdSetup.UpdateData()
        grdSetup.ActiveRow = Nothing
        For Each dr As DataRow In DirectCast(grdSetup.DataSource, DataTable).Rows
            dr.Item("SEQUENCE") = DBNull.Value
            dr.Item("PAGE_BREAK") = "0"
            If Clear_All Then
                dr.Item("EXCLUDE") = "0"
                dr.Item("GROUP_ALL_OTHERS") = "0"
                dr.Item("CODE_VALUES") = ""
            End If
        Next
        SEQs = 0
        Call Re_SEQ()

    End Sub

    Sub Setup_grdSetup()
        Create_tblASTDSQLA()

        Call Get_PARM("GLTPARM1")

        Dim COLUMN_CAPTION As String = ""
        For Each dr As DataRow In ASCDATA1.GetDataTable("Select ASTDSQLA.COLUMN_NAME, NVL(ASTDSQLA.COLUMN_CAPTION,ASTDSQLK.COLUMN_CAPTION) COLUMN_CAPTION, ASTDSQLA.SORTABLE, ASTDSQLA.COLUMN_LAST from ASTDSQLA,ASTDSQLK WHERE ASTDSQLK.COLUMN_NAME (+) = ASTDSQLA.COLUMN_NAME and ASTDSQLA.FORM_NAME = '" & FORM_NAME & "' ORDER BY NVL(ASTDSQLA.COLUMN_CAPTION,ASTDSQLK.COLUMN_CAPTION)").Rows
            If dr.Item("COLUMN_NAME") = "SEG2_CODE" And ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "" = "" _
            Or dr.Item("COLUMN_NAME") = "SEG3_CODE" And ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "" = "" _
            Or dr.Item("COLUMN_NAME") = "SEG4_CODE" And ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "" = "" _
            Then
                ' SKIP IT
            Else
                COLUMN_CAPTION = dr.Item("COLUMN_CAPTION") & ""
                If dr.Item("COLUMN_NAME") = "SEG2_CODE" Then
                    COLUMN_CAPTION = ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC")
                ElseIf dr.Item("COLUMN_NAME") = "SEG3_CODE" Then
                    COLUMN_CAPTION = ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC")
                ElseIf dr.Item("COLUMN_NAME") = "SEG4_CODE" Then
                    COLUMN_CAPTION = ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC")
                End If
                If dr.Item("SORTABLE") & "" = "1" Or dr.Item("COLUMN_LAST") & "" = "1" Then
                    PB_Report = True
                End If
                If dr.Item("COLUMN_LAST") & "" = "1" Then
                    COLUMN_NAME_last = dr.Item("COLUMN_NAME")
                    dr.Item("SORTABLE") = "0"
                    'PB_Report = True ?
                End If
                Call Add_Row(COLUMN_CAPTION, dr.Item("COLUMN_NAME") & "", dr.Item("SORTABLE") & "")
            End If

        Next dr

        grdSetup.DataSource = tblASTDSQLA
        grdSetup.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChangeOrLostFocus

        SEQs = 0
        Call Re_SEQ()
        'grdSetup.DisplayLayout.Bands(0).SortedColumns.Add(grdSetup.DisplayLayout.Bands(0).Columns("COLUMN_CAPTION"), False)

        ' GET TO THE TOP
        If grdSetup.Rows.Count > 0 Then
            grdSetup.ActiveRow = grdSetup.Rows(0)
        End If

        grdSetup.UpdateData()

        grdSetup.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSetup.DisplayLayout.Bands(0).SortedColumns.Add("COLUMN_CAPTION", False)
    End Sub

    Sub Create_tblASTDSQLA()
        tblASTDSQLA = New DataTable
        tblASTDSQLA.Columns.Add("COLUMN_NAME")
        tblASTDSQLA.Columns.Add("COLUMN_CAPTION")
        tblASTDSQLA.Columns.Add("CODE_VALUES")
        tblASTDSQLA.Columns.Add("EXCLUDE")
        tblASTDSQLA.Columns.Add("SEQUENCE", GetType(System.Int16))
        tblASTDSQLA.Columns.Add("PAGE_BREAK")
        tblASTDSQLA.Columns.Add("SORTABLE")
        tblASTDSQLA.Columns.Add("GROUP_ALL_OTHERS")
        tblASTDSQLA.Columns.Add("COLUMN_LAST")
        tblASTDSQLA.PrimaryKey = New DataColumn() {tblASTDSQLA.Columns("COLUMN_NAME")}
    End Sub

    Sub Add_Row(ByVal COLUMN_CAPTION As String, ByVal COLUMN_NAME As String, ByVal SORTABLE As String)
        Dim dr As DataRow
        dr = tblASTDSQLA.NewRow
        dr.Item("COLUMN_NAME") = COLUMN_NAME
        dr.Item("COLUMN_CAPTION") = COLUMN_CAPTION
        dr.Item("EXCLUDE") = "0"
        dr.Item("PAGE_BREAK") = "0"
        dr.Item("SORTABLE") = SORTABLE
        dr.Item("GROUP_ALL_OTHERS") = "0"
        tblASTDSQLA.Rows.Add(dr)
    End Sub

    Sub Re_SEQ( _
    Optional ByVal COLUMN_NAME As String = "", _
    Optional ByVal add_to_sort As Boolean = False)

        'grdSetup.Update 
        grdSetup.UpdateData()

        Dim tbl As DataTable = DirectCast(grdSetup.DataSource, DataTable)
        Dim row As DataRow

        If COLUMN_NAME <> "" Then
            row = tbl.Rows.Find(COLUMN_NAME)
            If add_to_sort Then
                row.Item("SEQUENCE") = 9
            Else
                row.Item("SEQUENCE") = Null
                row.Item("PAGE_BREAK") = "0"
            End If
        End If

        If COLUMN_NAME_last <> "" Then
            row = tbl.Rows.Find(COLUMN_NAME_last)
            row.Item("SEQUENCE") = Null
            row.Item("PAGE_BREAK") = "0"
        End If

        SEQs = 0
        For Each dr As DataRow In tbl.Select _
            ("SEQUENCE IS NOT NULL OR SEQUENCE <> ''", "SEQUENCE")
            SEQs = SEQs + 1
            dr.Item("SEQUENCE") = SEQs
        Next

        If COLUMN_NAME_last <> "" Then
            row = tbl.Rows.Find(COLUMN_NAME_last)
            SEQs = SEQs + 1
            row.Item("SEQUENCE") = SEQs
        End If

    End Sub

    Sub Rebuild_Values()
        Dim CODE_VALUES As String = ""
        For Each gr As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Rows
            CODE_VALUES = CODE_VALUES & "," & gr.Cells(0).Text
        Next
        CODE_VALUES = Mid$(CODE_VALUES, 2)
        grdSetup.ActiveRow.Cells("CODE_VALUES").Value = CODE_VALUES
        If CODE_VALUES = "" Then
            grdSetup.ActiveRow.Cells("EXCLUDE").Value = "0"
        End If
        If CODE_VALUES = "" Or grdSetup.ActiveRow.Cells("SEQUENCE").Value & "" = "" Then
            grdSetup.ActiveRow.Cells("GROUP_ALL_OTHERS").Value = "0"
        End If
        grdSetup.UpdateData()

        Dim z As String = ASCMAIN1.CodeSelector.VIEW_DESC
        If grd.Rows.Count <> 0 Then
            z = z & " (" & CStr(grd.Rows.Count) & ")"
        End If
        grd.Text = z

        cmdAll.Visible = (CODE_VALUES <> "")
    End Sub

    Private Sub grdSetup_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSetup.AfterCellUpdate
        'Try
        '    grdSetup.UpdateData()
        'Catch ex As Exception
        'End Try
    End Sub

    Private Sub grdSetup_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSetup.AfterRowActivate
        Call Show_grd()
    End Sub

    Private Sub grdSetup_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSetup.AfterRowUpdate
        Call Show_grd()
    End Sub

    Private Sub grdSetup_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSetup.BeforeRowUpdate

        'If Val(grdSetup.ActiveRow.Cells("SEQUENCE").Value & "") = 0 Then
        '    grdSetup.ActiveRow.Cells("PAGE_BREAK").Value = "0"
        'End If
        If Val(e.Row.Cells("SEQUENCE").Value & "") = 0 Then
            grdSetup.DisplayLayout.Rows(e.Row.Index).Cells("PAGE_BREAK").Value = "0"
        End If

        Dim COLUMN_NAME As String = e.Row.Cells("COLUMN_NAME").Text ' grdSetup.ActiveRow.Cells("COLUMN_NAME").Text
        Dim sql As String = ASCMAIN1.CodeSelector.Get_SQL(COLUMN_NAME)
        If sql <> "" Then
            Dim CODE_VALUES_new As String = ""
            Dim CODE_VALUES As String = e.Row.Cells("CODE_VALUES").Text
            Dim KEY_EXPRESSION As String = ASCMAIN1.CodeSelector.grdColumns(0).Item("COLUMN_NAME")

            If CODE_VALUES <> "" Then
                Dim CODE_VALUES_old As String = ""
                For Each txt As String In Split(Replace(CODE_VALUES, "'", ""), ",")
                    CODE_VALUES_old = CODE_VALUES_old & ",'" & ASCMAIN1.Format_Field(txt, COLUMN_NAME, , True) & "'"
                Next
                CODE_VALUES_old = Mid$(CODE_VALUES_old, 2)
                Dim where_or_and As String = " where "
                If ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("WHERE_CLAUSE") & "" <> "" Then
                    where_or_and = " and "
                End If

                For Each dr As DataRow In ASCDATA1.GetDataTable(sql & where_or_and & KEY_EXPRESSION & " IN (" & CODE_VALUES_old & ")").Rows
                    CODE_VALUES_new = CODE_VALUES_new & "," & dr.Item(0)
                Next
            End If

            CODE_VALUES_new = Mid(CODE_VALUES_new, 2)
            If CODE_VALUES_new <> CODE_VALUES Then
                cmdAll.Visible = (CODE_VALUES <> "")

                grdSetup.DisplayLayout.Rows(e.Row.Index).Cells("CODE_VALUES").Value = CODE_VALUES_new '  .ActiveRow.Cells("CODE_VALUES").Value = CODE_VALUES_new
                Call Show_grd()
            End If
        End If

    End Sub

    Private Sub grdSetup_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSetup.ClickCellButton
        If e.Cell.Column.Key = "COLUMN_CAPTION" Then
            If e.Cell.Row.Cells("SORTABLE").Text = "1" Then
                If e.Cell.Row.Cells("SEQUENCE").Text <> "" Then
                    Call Re_SEQ(e.Cell.Row.Cells("COLUMN_NAME").Text, False)
                Else
                    Call Re_SEQ(e.Cell.Row.Cells("COLUMN_NAME").Text, True)
                End If
            End If
        ElseIf e.Cell.Column.Key = "CODE_VALUES" Then
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(grdSetup.ActiveRow.Cells("COLUMN_NAME").Text)
            If ASCMAIN1.CodeSelector.SQL <> "" Then
                ASCMAIN1.CodeSelector.MultipleSelections = True
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = Replace(grdSetup.ActiveRow.Cells("CODE_VALUES").Text & "", ",", Chr(0))
                Dim F As New ASFCODE1
                F.ShowDialog()
                F.Dispose()
                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                    grdSetup.ActiveRow.Cells("CODE_VALUES").Value = Mid$(Replace(ASCMAIN1.CodeSelector.SelectedCodes0, Chr(0), ","), 2)
                    grdSetup.UpdateData()
                    Call Show_grd()
                End If
            End If
        End If
    End Sub

    Private Sub grdSetup_DoubleClickHeader(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickHeaderEventArgs) Handles grdSetup.DoubleClickHeader
        Call Clear_grdSetup(False)
    End Sub

    Private Sub grdSetup_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSetup.InitializeRow
        If e.Row.Cells("SORTABLE").Text <> "1" Then
            e.Row.Cells("COLUMN_CAPTION").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit
        End If
    End Sub

    Private Sub grdSetup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grdSetup.KeyDown
        If e.KeyValue = Windows.Forms.Keys.Delete Then
            If grdSetup.ActiveCell.Column.Key = "SEQUENCE" Then
                If grdSetup.ActiveCell.Text <> "" Then
                    'grdSetup.ActiveCell.Value = DBNull.Value
                    'grdSetup.UpdateData()
                    Call Re_SEQ(grdSetup.ActiveRow.Cells("COLUMN_NAME").Text, False)
                End If
            End If
        End If

        If e.KeyValue = Windows.Forms.Keys.Enter Then
            If grdSetup.ActiveCell.Column.Key = "CODE_VALUES" Then
                grdSetup.Update()
            End If
        End If
    End Sub

    Private Sub grdSetup_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles grdSetup.KeyPress
        If grdSetup.ActiveCell IsNot Nothing Then
            If grdSetup.ActiveCell.Column.Key = "SEQUENCE" And grdSetup.ActiveRow.Cells("SORTABLE").Text = "1" Then
                Dim COLUMN_NAME As String = grdSetup.ActiveRow.Cells("COLUMN_NAME").Text
                Dim SEQcur As Integer = Val(grdSetup.ActiveCell.Text)
                Dim SEQnew As Integer = Val(e.KeyChar)
                If SEQnew < 1 Or SEQnew = SEQcur Or (SEQcur = 0 And SEQnew > SEQs + 1) Or (SEQcur <> 0 And SEQnew > SEQs) Then
                    Exit Sub
                End If

                grdSetup.ActiveCell.Value = SEQnew
                grdSetup.UpdateData()

                Dim i As Integer
                Dim z As String
                If SEQnew < SEQcur Or SEQcur = 0 Then
                    z = ">"
                    i = SEQnew
                Else
                    z = "<"
                    i = 0
                End If
                For Each dr As DataRow In DirectCast(grdSetup.DataSource, DataTable).Select("SEQUENCE " & z & "= " & CStr(SEQnew), "SEQUENCE")
                    If dr.Item("COLUMN_NAME") <> COLUMN_NAME Then
                        i = i + 1
                        dr.Item("SEQUENCE") = i
                    End If
                Next

                If SEQcur = 0 Then
                    SEQs = SEQs + 1
                End If
            End If
        End If
    End Sub

    Private Sub grdSetup_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSetup.Leave
        grdSetup.UpdateData()
    End Sub
#End Region

#Region "grd"
    Sub Show_grd()
        LIST_CODE = ""
        LIST_DESC = ""

        Dim sql As String = ASCMAIN1.CodeSelector.Get_SQL(grdSetup.ActiveRow.Cells("COLUMN_NAME").Text)
        If sql = "" Then
            grd.Visible = False
            cmdAll.Visible = False
            SplitContainer4.Panel1.Hide()
            grpCodeLists.Visible = False
        Else
            Dim CODE_VALUES As String = grdSetup.ActiveRow.Cells("CODE_VALUES").Text
            Dim KEY_EXPRESSION As String = ASCMAIN1.CodeSelector.grdColumns(0).Item("COLUMN_NAME")
            Dim sqlx As String = " where "
            If InStr(sql, " where ") <> 0 Then
                sqlx = " and "
            End If
            If CODE_VALUES <> "" Then
                sql = sql & sqlx & KEY_EXPRESSION & " IN ('" & Replace(Replace(CODE_VALUES, "'", ""), ",", "','") & "')"
            Else
                sql = sql & sqlx & "ROWNUM < 1"
            End If
            grd.DataSource = Nothing
            grd.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
            grd.DataSource = ASCDATA1.GetDataTable(sql)
            Dim z As String = ASCMAIN1.CodeSelector.VIEW_DESC
            If grd.Rows.Count <> 0 Then
                z = z & " (" & CStr(grd.Rows.Count) & ")"
            End If
            grd.Text = z

            For i As Integer = 0 To ASCMAIN1.CodeSelector.grdColumns.Count - 1 ' grd.DisplayLayout.Bands(0).Columns.Count - 1
                grd.DisplayLayout.Bands(0).Columns(i).Header.Caption = ASCMAIN1.CodeSelector.grdColumns(i).Item("COLUMN_CAPTION")
                If Val(ASCMAIN1.CodeSelector.grdColumns(i).Item("COLUMN_WIDTH") & "") <> 0 Then
                    grd.DisplayLayout.Bands(0).Columns(i).Width = ASCMAIN1.CodeSelector.grdColumns(i).Item("COLUMN_WIDTH")
                End If
            Next i
            grd.Visible = True
            cmdAll.Visible = (grd.Rows.Count <> 0)

            SplitContainer4.Panel1.Show()
            grpCodeLists.Visible = True

            txtList.Text = ""
            chkListShareable.Checked = False
            chkListModifiable.Checked = False
        End If
    End Sub

    Private Sub grd_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grd.AfterRowsDeleted
        Call Rebuild_Values()
    End Sub
#End Region

#Region "grdASTROPT1"

    Private Sub grdASTROPT1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTROPT1.AfterRowActivate
        With grdASTROPT1.DisplayLayout.Bands(0)
            If grdASTROPT1.ActiveRow.Cells("INIT_OPER").Text <> ASCMAIN1.USER_ID Then
                grdASTROPT1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdASTROPT1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                '.Columns("SET_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                '.Columns("SET_YP_REL").CellActivation = UltraWinGrid.Activation.NoEdit
                '.Columns("SET_ALLOW_OTHERS").CellActivation = UltraWinGrid.Activation.NoEdit
            Else
                grdASTROPT1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                grdASTROPT1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                '.Columns("SET_DESC").CellActivation = UltraWinGrid.Activation.AllowEdit
                '.Columns("SET_YP_REL").CellActivation = UltraWinGrid.Activation.AllowEdit
                '.Columns("SET_ALLOW_OTHERS").CellActivation = UltraWinGrid.Activation.AllowEdit
            End If
        End With
    End Sub

    Private Sub grdASTROPT1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTROPT1.AfterRowsDeleted
        adaASTROPT1.Update(tblASTROPT1)
        For J As Integer = 1 To ASCMAIN1.grdRows.Count
            Call Delete_Saved_Setting(ASCMAIN1.grdRows(J))
        Next
    End Sub

    Sub Delete_Saved_Setting(ByVal SET_ID As String)
        Dim Sql As String = "Delete from ASTROPT2 where SET_ID = '" & SET_ID & "'"
        Dim i As Integer = ASCDATA1.ExecuteSQL(Sql)
    End Sub

    Private Sub grdASTROPT1_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdASTROPT1.AfterRowUpdate
        adaASTROPT1.Update(tblASTROPT1)
        If grdASTROPT1.ActiveRow.Cells("SET_ID").Text = SET_ID Then
            txtDescription.Text = grdASTROPT1.ActiveRow.Cells("SET_DESC").Text
        End If
    End Sub

    Private Sub grdASTROPT1_BeforeEnterEditMode(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles grdASTROPT1.BeforeEnterEditMode
        'If chkExecutionHistory.Checked Or (grdASTROPT1.ActiveRow.Cells("INIT_OPER").Text <> ASCMAIN1.USER_ID And Not grdASTROPT1.ActiveRow.IsAddRow) Then
        '    e.Cancel = True
        'End If
    End Sub

    Private Sub grdASTROPT1_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdASTROPT1.BeforeRowsDeleted
        'If chkExecutionHistory.Checked Then
        '    e.Cancel = True
        'End If

        'ASCMAIN1.grdRows.Clear()
        'For Each DR As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Rows
        '    ASCMAIN1.grdRows.Add(DR.Cells("SET_ID").Text)
        '    If DR.Cells("INIT_OPER").Value <> ASCMAIN1.USER_ID Then
        '        MsgBox("You cannot delete or modify records that you did not create" & vbCr & vbCr & "(" & DR.Cells("SET_DESC").Value & " was created by " & DR.Cells("INIT_OPER").Value & ")", MsgBoxStyle.OkOnly, "Cannot Peform Requested Action")
        '        e.Cancel = True
        '        ASCMAIN1.grdRows.Clear()
        '        Exit For
        '    End If
        'Next
    End Sub

    Private Sub grdASTROPT1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTROPT1.ClickCellButton
        SET_ID = grdASTROPT1.ActiveRow.Cells("SET_ID").Text
        Call Retrieve_Settings()
    End Sub

    Private Sub grdASTROPT1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grdASTROPT1.KeyDown
        If e.KeyValue = Windows.Forms.Keys.Enter Then
            grdASTROPT1.UpdateData()
        End If
    End Sub

    Private Sub grdASTROPT1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTROPT1.Leave
        grdASTROPT1.UpdateData()
    End Sub
#End Region

#Region "Dynamic SQL - Generation of SQL"

    Sub Get_SQL( _
    ByVal DATA_SOURCE As String, _
    Optional ByVal TABLE_NAME_temp As String = "")

        Call ASCMAIN1.Track("Extracts from Data Sources", DATA_SOURCE)

        Dim TABLE_NAME As String
        Dim COLUMN_NAME As String
        Dim rowASTDSQLC As DataRow

        Dim sql_SELECT_col As String
        Dim sql_Select_col_count As Integer = 0

        DATA_SOURCE = IIf(DATA_SOURCE = "", "*", DATA_SOURCE)

        If PB_Report Then
            Dim rowASTDSQLB As DataRow = tblASTDSQLB.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE})
            sql_TABLE_NAME = rowASTDSQLB.Item("TABLE_NAME")

            sql_SELECT_cols = ""
            sql_GROUP_BY_cols = ""
            sql_WHERE = ""
            sql_TABLE_NAMEs = ""
            sql_JOIN = ""

            ' Forced Joins - SHOULDN'T WE BE LOOKING AT J FOR THIS?

            For Each rowASTDSQLJ As DataRow In tblASTDSQLJ.Select("FORM_NAME = '" & FORM_NAME & "' and DATA_SOURCE = '" & DATA_SOURCE & "' and ALWAYS_JOIN = '1'")
                TABLE_NAME = rowASTDSQLJ.Item("TABLE_NAME")
                Call Get_SQL_Join_Criteria(TABLE_NAME, DATA_SOURCE)
            Next

            ' Sort

            For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE")
                sql_SELECT_col = ""
                'sql_GROUP_BY_col = ""

                COLUMN_NAME = rowASTDSQLA.Item("COLUMN_NAME")

                rowASTDSQLC = tblASTDSQLC.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE, COLUMN_NAME})


                If rowASTDSQLC Is Nothing Then
                    rowASTDSQLC = tblASTDSQLC.Rows.Find(New Object() {FORM_NAME, "*", COLUMN_NAME})
                End If

                TABLE_NAME = rowASTDSQLC.Item("TABLE_NAME") & ""
                If TABLE_NAME <> "" Then
                    If TABLE_NAME <> sql_TABLE_NAME Then
                        Call Get_SQL_Join_Criteria(TABLE_NAME, DATA_SOURCE)
                    End If
                Else
                    TABLE_NAME = sql_TABLE_NAME
                End If

                'If rowASTDSQLC.Item("EXPRESSION_IND") & "" = "1" Then
                If rowASTDSQLC.Item("COLUMN_EXPRESSION") & "" <> "" Then
                    sql_SELECT_col = rowASTDSQLC.Item("COLUMN_EXPRESSION") & ""
                    'sql_GROUP_BY_col = rowASTDSQLC.Item("COLUMN_EXPRESSION_Y") & ""
                Else
                    sql_SELECT_col = TABLE_NAME & "." & COLUMN_NAME
                End If

                If rowASTDSQLA.Item("GROUP_ALL_OTHERS") & "" = "1" And rowASTDSQLA.Item("CODE_VALUES") & "" <> "" Then
                    sql_SELECT_col = "Case When " & sql_SELECT_col & " in (" & Replace(rowASTDSQLA.Item("CODE_VALUES"), ",", "','") & "') Then " & sql_SELECT_col & " else '*' End"
                End If

                ASCMAIN1.TACMAIN1.Get_Column_Expression_Exceptions(FORM_NAME, DATA_SOURCE, COLUMN_NAME, sql_SELECT_col) ' , sql_GROUP_BY_col)

                sql_SELECT_cols = sql_SELECT_cols & ", " & sql_SELECT_col & " AS " & COLUMN_NAME
                'If sql_GROUP_BY_col = "" Then
                '    sql_GROUP_BY_col = sql_SELECT_col
                'End If
                'sql_GROUP_BY_cols = sql_GROUP_BY_cols & ", " & sql_GROUP_BY_col
                sql_GROUP_BY_cols = sql_GROUP_BY_cols & ", " & sql_SELECT_col
                sql_Select_col_count = sql_Select_col_count + 1
            Next

            If COLUMN_NAMEs.Count > sql_Select_col_count Then
                sql_SELECT_col = COLUMN_NAMEs(COLUMN_NAMEs.Count - 1)
                sql_SELECT_cols = sql_SELECT_cols & ", " & sql_SELECT_col
                sql_GROUP_BY_cols = sql_GROUP_BY_cols & ", " & sql_SELECT_col
            End If

            sql_SELECT_cols = Mid$(sql_SELECT_cols, 3)
            sql_GROUP_BY_cols = Mid$(sql_GROUP_BY_cols, 3)
        End If

        ' Filter

        For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select("CODE_VALUES is Not Null AND CODE_VALUES <> ''")
            COLUMN_NAME = rowASTDSQLA.Item("COLUMN_NAME")

            rowASTDSQLC = tblASTDSQLC.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE, COLUMN_NAME})
            If rowASTDSQLC Is Nothing Then
                rowASTDSQLC = tblASTDSQLC.Rows.Find(New Object() {FORM_NAME, "*", COLUMN_NAME})
            End If

            TABLE_NAME = rowASTDSQLC.Item("TABLE_NAME") & ""
            If TABLE_NAME <> "" Then
                If TABLE_NAME <> sql_TABLE_NAME Then
                    Call Get_SQL_Join_Criteria(TABLE_NAME, DATA_SOURCE)
                End If
            Else
                TABLE_NAME = sql_TABLE_NAME
            End If

            If rowASTDSQLC.Item("JOIN_SPECIAL") & "" = "1" Then
                sql_SELECT_col = GetSpecialSelectedJoin(COLUMN_NAME, DATA_SOURCE)
            ElseIf rowASTDSQLC.Item("COLUMN_EXPRESSION") & "" <> "" Then
                sql_SELECT_col = rowASTDSQLC.Item("COLUMN_EXPRESSION") & ""
            Else
                sql_SELECT_col = TABLE_NAME & "." & COLUMN_NAME
            End If

            Dim in_or_equal As String
            Dim not_in_or_not_equal As String

            Dim CODE_VALUES_sql As String = "'" & Replace(rowASTDSQLA.Item("CODE_VALUES"), ",", "','") & "'"
            If InStr(CODE_VALUES_sql, ",") = 0 Then
                in_or_equal = "="
                not_in_or_not_equal = "<>"
            Else
                in_or_equal = "IN"
                not_in_or_not_equal = "NOT IN"
            End If

            If rowASTDSQLA.Item("EXCLUDE") & "" = "1" Then
                sql_WHERE = sql_WHERE & " AND (" & sql_SELECT_col & " IS NULL OR " & sql_SELECT_col & " " & not_in_or_not_equal & " (" & CODE_VALUES_sql & "))"
            Else
                sql_WHERE = sql_WHERE & " AND " & sql_SELECT_col & " " & in_or_equal & " (" & CODE_VALUES_sql & ")"
            End If
        Next

        'tblASTDSQLC = Nothing

        If PB_Report Then
            ' Pad sql_SELECT_cols for unused Group By's

            If COLUMN_NAMEs.Count < 9 Then
                For i As Integer = COLUMN_NAMEs.Count + 1 To 9
                    sql_SELECT_cols = sql_SELECT_cols & ", 'x' as G" & CStr(i)
                Next
                If COLUMN_NAMEs.Count = 0 Then
                    sql_SELECT_cols = Mid(sql_SELECT_cols, 3)
                End If
            End If

            ' Create Report Work File

            If ASTSRPT1 = "" Then
                COLUMN_NAMEs_appended = ""
                Dim sql_sum As String = ""
                Dim sql_sum_group_by As String = ""
                Dim sql As String = ""
                For i As Integer = 1 To 9
                    sql &= ",ASTSRPT1.G" & CStr(i)
                    sql_sum &= ",G" & CStr(i)
                Next
                sql = "Select " & Mid(sql, 2)
                sql_sum_group_by = Mid(sql_sum, 2)
                sql_sum = "Select " & sql_sum_group_by
                Dim ZTBL As String = ""
                If tblASTDSQLH.Rows.Count <> 0 Then
                    For Each ROW As DataRow In tblASTDSQLH.Select("", "COLUMN_SEQ")
                        Dim TABLE_NAME_appended_column As String = ROW.Item("TABLE_NAME") & ""
                        If TABLE_NAME_appended_column = "" Then
                            If TABLE_NAME_temp <> "" Then
                                TABLE_NAME_appended_column = TABLE_NAME_temp
                            Else
                                Dim rowASTDSQLB As DataRow = tblASTDSQLB.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE})
                                TABLE_NAME_appended_column = rowASTDSQLB("TABLE_NAME") & ""
                            End If
                        End If
                        sql = sql & "," & TABLE_NAME_appended_column & "." & ROW.Item("COLUMN_NAME") & " " & ROW.Item("COLUMN_ALIAS")
                        If InStr(ZTBL, "," & TABLE_NAME_appended_column) = 0 Then
                            ZTBL &= "," & TABLE_NAME_appended_column
                        End If
                        Dim COLUMN_ALIAS As String = ROW.Item("COLUMN_ALIAS") & ""
                        If COLUMN_ALIAS = "" Then
                            COLUMN_ALIAS = ROW.Item("COLUMN_NAME")
                        End If
                        sql_sum &= "," & COLUMN_ALIAS
                        COLUMN_NAMEs_appended &= "," & COLUMN_ALIAS
                        sql_sum_group_by &= "," & COLUMN_ALIAS
                    Next
                End If
                If tblASTRECAP.Rows.Count <> 0 Then
                    COLUMN_NAME_RECAP_ROW_NO = ", ASTSRPT1_RECAP_ROW_NO"
                    sql = sql & ", TATWORK1.W_INT " & Mid(COLUMN_NAME_RECAP_ROW_NO, 3)
                    If InStr(ZTBL, "," & "TATWORK1") = 0 Then
                        ZTBL &= "," & "TATWORK1"
                    End If
                    sql_sum &= COLUMN_NAME_RECAP_ROW_NO
                    sql_sum_group_by &= COLUMN_NAME_RECAP_ROW_NO

                    dst.Tables.Add(tblASTRECAP.Copy)
                End If
                ASTSRPT1_sum_columns = ""
                For Each KEY As String In COLUMN_NAME_sum.Keys
                    Select Case COLUMN_NAME_sum(KEY)
                        Case "QTY"
                            sql = sql & ",ASTSRPT1.W_QTY " & KEY
                        Case "AMT"
                            sql = sql & ",ASTSRPT1.W_AMT " & KEY
                        Case "DEC"
                            sql = sql & ",ASTSRPT1.W_DEC " & KEY
                        Case Else
                            MsgBox("Invalid Data Type")
                            Stop
                    End Select
                    sql_sum = sql_sum & ",SUM(" & KEY & ") " & KEY
                    ASTSRPT1_sum_columns &= ",SUM(" & KEY & ") " & KEY
                Next

                sql = sql & " from ASTSRPT1" & ZTBL & " where ROWNUM < 1"
                ASTSRPT1 = ASCMAIN1.Temp_Table(sql)
                sql_sum = sql_sum & " from " & ASTSRPT1 & " group by " & sql_sum_group_by
                ASTSRPT1_sql_sum = sql_sum
            End If
        End If

    End Sub

    Private Sub Get_SQL_Join_Criteria(ByVal TABLE_NAME As String, ByVal DATA_SOURCE As String)

        If InStr(sql_TABLE_NAMEs, "," & TABLE_NAME) <> 0 Then
            Exit Sub
        Else
            sql_TABLE_NAMEs = sql_TABLE_NAMEs & "," & TABLE_NAME
        End If

        Dim COLUMN_NAME As String
        Dim sql As String

        Dim rowASTDSQLJ As DataRow = tblASTDSQLJ.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE, TABLE_NAME})
        If rowASTDSQLJ Is Nothing Then
            rowASTDSQLJ = tblASTDSQLJ.Rows.Find(New Object() {FORM_NAME, "*", TABLE_NAME})
        End If

        Dim JOIN_TYPE As String = ""
        If rowASTDSQLJ IsNot Nothing AndAlso rowASTDSQLJ.Item("OUTER_JOIN") & "" = "1" Then
            JOIN_TYPE = "(+)"
        End If

        Dim drsASTDSQLD() As DataRow
        sql = "FORM_NAME = '" & FORM_NAME & "' and TABLE_NAME = '" & TABLE_NAME _
            & "' and DATA_SOURCE = '" & DATA_SOURCE & "'"
        drsASTDSQLD = tblASTDSQLD.Select(sql)
        If drsASTDSQLD.Length = 0 Then
            sql = "FORM_NAME = '" & FORM_NAME & "' and TABLE_NAME = '" & TABLE_NAME _
                & "' and DATA_SOURCE = '" & "*" & "'"
            drsASTDSQLD = tblASTDSQLD.Select(sql)
        End If

        If drsASTDSQLD.Length = 0 Then
            Dim tbl As DataTable = ASCDATA1.GetDataTable("*", TABLE_NAME, -1, False)
            ReDim drsASTDSQLD(tbl.PrimaryKey.Length - 1)
            For i As Integer = 0 To tbl.PrimaryKey.Length - 1
                Dim row As DataRow = tblASTDSQLD.NewRow
                Dim dc As DataColumn = tbl.PrimaryKey(i)
                row.Item("FORM_NAME") = FORM_NAME
                row.Item("DATA_SOURCE") = DATA_SOURCE
                row.Item("TABLE_NAME") = TABLE_NAME
                row.Item("COLUMN_NAME") = dc.ColumnName
                drsASTDSQLD(i) = row
            Next
        End If

        Dim rowASTDSQLD As DataRow
        For i As Integer = 0 To UBound(drsASTDSQLD)
            rowASTDSQLD = drsASTDSQLD(i)

            If rowASTDSQLD.Item("TABLE_NAME_JOIN") & "" <> "" Then
                Call Get_SQL_Join_Criteria(rowASTDSQLD.Item("TABLE_NAME_JOIN"), DATA_SOURCE)
            End If

            COLUMN_NAME = rowASTDSQLD.Item("COLUMN_NAME")
            Dim TABLE_NAME_JOIN As String = rowASTDSQLD.Item("TABLE_NAME_JOIN") & ""
            If TABLE_NAME_JOIN = "" Then
                TABLE_NAME_JOIN = sql_TABLE_NAME
            End If

            sql = TABLE_NAME & "." & COLUMN_NAME & JOIN_TYPE & " = "

            'sql = JOIN_TYPE & TABLE_NAME & " ON " & COLUMN_NAME & " = "
            ''If rowASTDSQLD.Item("EXPRESSION_IND") & "" = "1" Then
            '' sql = sql & rowASTDSQLD.Item("COLUMN_NAME_JOIN")
            '' Else
            Dim COLUMN_NAME_JOIN As String = rowASTDSQLD.Item("COLUMN_NAME_JOIN") & ""
            If COLUMN_NAME_JOIN = "" Then
                COLUMN_NAME_JOIN = COLUMN_NAME
            End If
            sql = sql & TABLE_NAME_JOIN & "." & COLUMN_NAME_JOIN
            'End If

            sql_JOIN = sql_JOIN & " AND " & sql
        Next

    End Sub

    Private Sub Get_SQL_Join_Criteria_Special(ByVal FORM_NAME As String, ByVal TABLE_NAME As String, ByVal DATA_SOURCE As String)

        'If TABLE_NAME = "" Then Exit Sub

        'jz = ""
        '' Special Conditions
        'Select Case FORM_NAME
        '    Case "SOFWHOD1"
        '        Select Case TABLE_NAME
        '            Case "ARTCUST1"
        '                If DATA_SOURCE = "A" Then
        '                    jz = "ARTCUST1.CUST_CODE (+) = Y2.CUST_CODE"
        '                Else
        '                    jz = "ARTCUST1.CUST_CODE (+) = X.CUST_CODE"
        '                End If

        '            Case "ICTITEM1"
        '                If DATA_SOURCE = "A" Then
        '                    jz = "ICTITEM1.ITEM_CODE (+) = Y1.ITEM_CODE"
        '                Else
        '                    jz = "ICTITEM1.ITEM_CODE (+) = X.ITEM_CODE"
        '                End If

        '        End Select

        'End Select

        'If jz <> "" And InStr(1, sqljoin, jz) = 0 Then
        '    sqljoin = sqljoin & " AND " & jz
        '    If InStr(sql_TABLE_NAMEs, "," & TABLE_NAME) = 0 Then 'Exit Sub
        '        sql_TABLE_NAMEs = sql_TABLE_NAMEs & "," & TABLE_NAME
        '    End If
        'End If

    End Sub

    Private Function GetSpecialSelectedJoin(ByVal COLUMN_NAME As String, ByVal DATA_SOURCE As String) As String
        Dim z As String = ""

        'Select Case FORM_NAME
        '    Case "SOFSLSF1"
        '        Select Case COLUMN_NAME
        '            Case "ITEM_BRAND_CODE"
        '        End Select
        '    Case "SOFWHOD1"
        '        Select Case COLUMN_NAME
        '            Case "CUST_CODE"
        '                If DATA_SOURCE = "A" Then
        '                    z = "Y2.CUST_CODE"
        '                Else
        '                    z = "X.CUST_CODE"
        '                End If

        '            Case "ITEM_CODE"
        '                If DATA_SOURCE = "C" Or DATA_SOURCE = "B" Then
        '                    z = ""
        '                ElseIf DATA_SOURCE = "A" Then
        '                    z = "Y1.ITEM_CODE"
        '                Else
        '                    z = "X.ITEM_CODE"
        '                End If
        '        End Select
        'End Select

        GetSpecialSelectedJoin = z

    End Function
#End Region

#Region "Lists"

    Private Sub cmdListRetrieve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdListRetrieve.Click
        COLUMN_NAME = grdSetup.ActiveRow.Cells("COLUMN_NAME").Text
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("LIST_CODE")
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.SQL &= " where COLUMN_NAME = '" & COLUMN_NAME & "'"
            ASCMAIN1.CodeSelector.SQL &= " and (INIT_OPER = '" & ASCMAIN1.USER_ID & "' or NVL(LIST_SHAREABLE,'0') = '1')"
            ASCMAIN1.CodeSelector.MultipleSelections = False
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections = 1 Then
                LIST_CODE = ASCMAIN1.CodeSelector.SelectedRows(0).Item("LIST_CODE")
                Dim i As Integer
                If grdSetup.ActiveRow.Cells("CODE_VALUES").Text <> "" Then
                    Dim frmASFMSGBF As New ASFMSGBF
                    i = frmASFMSGBF.Get_opt_from_User("Load this List of Codes", New String() {"By Replacing the Existing List of Codes", "By Appending to the Existing List of Codes"}, 0, "Retrieve Code List Option")
                    frmASFMSGBF.Dispose()
                End If

                Call Load_Code_List(i = 0)
            End If
        End If
    End Sub

    Sub Load_Code_List(ByVal replace_codes As Boolean)
        Dim tblASTLIST1 As DataTable = ASCDATA1.GetDataTable("Select * from ASTLIST1 where COLUMN_NAME = '" & COLUMN_NAME & "' and LIST_CODE = '" & LIST_CODE & "'")
        Dim tblASTLIST2 As DataTable = ASCDATA1.GetDataTable("Select COLUMN_VALUE from ASTLIST2 where COLUMN_NAME = '" & COLUMN_NAME & "' and LIST_CODE = '" & LIST_CODE & "' order by COLUMN_VALUE")

        Dim CODE_VALUES As String
        If replace_codes Then
            CODE_VALUES = ""
        Else
            CODE_VALUES = grdSetup.ActiveRow.Cells("CODE_VALUES").Value
        End If

        For Each dr As DataRow In tblASTLIST2.Rows
            If Not InStr("," & CODE_VALUES & ",", "," & dr.Item("COLUMN_VALUE") & ",") Then
                CODE_VALUES &= "," & dr.Item("COLUMN_VALUE")
            End If
        Next
        grdSetup.ActiveRow.Cells("CODE_VALUES").Value = Mid(CODE_VALUES, 2)
        grdSetup.UpdateData()


        Call Show_grd()

        LIST_CODE = tblASTLIST1.Rows(0).Item("LIST_CODE")
        LIST_DESC = tblASTLIST1.Rows(0).Item("LIST_DESC")
        txtList.Text = LIST_DESC
        chkListShareable.Checked = tblASTLIST1.Rows(0).Item("LIST_SHAREABLE")
        chkListModifiable.Checked = tblASTLIST1.Rows(0).Item("LIST_MODIFIABLE")
        chkListShareable.Enabled = (tblASTLIST1.Rows(0).Item("INIT_OPER") = ASCMAIN1.USER_ID)
        chkListModifiable.Enabled = (tblASTLIST1.Rows(0).Item("INIT_OPER") = ASCMAIN1.USER_ID)
    End Sub

    Private Sub cmdListSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdListSaveAs.Click
        If Trim(txtList.Text) = "" Then
            MsgBox("You Must Enter a List Description", MsgBoxStyle.OkOnly, "Cannot Save List")
            Exit Sub
        End If

        COLUMN_NAME = grdSetup.ActiveRow.Cells("COLUMN_NAME").Text
        Dim CODE_VALUES As String = grdSetup.ActiveRow.Cells("CODE_VALUES").Value

        If CODE_VALUES = "" Then
            MsgBox("No Code Values in the List")
            'Stop
            'ABS.UI.MessageBox.Show("No Code Values in the List", ABS.UI.Types.MessageBoxButton.OKOnly, "Cannot Save List")
            Exit Sub
        End If

        Dim i As Integer = 0
        If LIST_CODE = "" Then
            LIST_CODE = ASCMAIN1.Next_Control_No("ASTLIST1.LIST_CODE")
        Else
            Dim frmASFMSGBF As New ASFMSGBF
            If Not chkListModifiable.Enabled And Not chkListModifiable.Checked Then
                If txtList.Text = LIST_DESC Then
                    MsgBox("You must change the Description of this List")
                    'Stop
                    'ABS.UI.MessageBox.Show("You must change the Description of this List" & vbCr & " in order to Save it (as one of your own Lists)", ABS.UI.Types.MessageBoxButton.OKOnly, "Cannot Save List")
                    Exit Sub
                End If
                i = 0
            Else
                i = frmASFMSGBF.Get_opt_from_User("Save this List of Codes", New String() {"As a New List", "By Replacing Existing List"}, 0, "Save Code List Option")
            End If
            If i = -1 Then
                Exit Sub
            ElseIf i = 0 Then
                LIST_CODE = ASCMAIN1.Next_Control_No("ASTLIST1.LIST_CODE")
            Else
                ASCMAIN1.sql = "Delete from ASTLIST2 where COLUMN_NAME = '" & COLUMN_NAME & "' and LIST_CODE = '" & LIST_CODE & "'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            End If
        End If

        Dim tblASTLIST1 As New DataTable
        ASCMAIN1.sql = "Select * from ASTLIST1 where COLUMN_NAME = '" & COLUMN_NAME & "' and LIST_CODE = '" & LIST_CODE & "'"
        Using adaASTLIST1 As OracleDataAdapter = _
            ASCDATA1.GetDataAdapter(tblASTLIST1, "ASTLIST1", "", True)
            If i = 1 Then
                tblASTLIST1.Rows(0).Item("LIST_DESC") = txtList.Text
                tblASTLIST1.Rows(0).Item("LIST_SHAREABLE") = CStr(Abs(Val(chkListShareable.Checked)))
                tblASTLIST1.Rows(0).Item("LIST_MODIFIABLE") = CStr(Abs(Val(chkListModifiable.Checked)))
                tblASTLIST1.Rows(0).Item("LAST_OPER") = ASCMAIN1.USER_ID
                tblASTLIST1.Rows(0).Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
            Else
                Dim rowASTLIST1 As DataRow = tblASTLIST1.NewRow
                rowASTLIST1.Item("COLUMN_NAME") = COLUMN_NAME
                rowASTLIST1.Item("LIST_CODE") = LIST_CODE
                rowASTLIST1.Item("LIST_DESC") = txtList.Text
                rowASTLIST1.Item("LIST_SHAREABLE") = CStr(Abs(Val(chkListShareable.Checked)))
                rowASTLIST1.Item("LIST_MODIFIABLE") = CStr(Abs(Val(chkListModifiable.Checked)))
                rowASTLIST1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowASTLIST1.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                rowASTLIST1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowASTLIST1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                tblASTLIST1.Rows.Add(rowASTLIST1)
                adaASTLIST1.Update(tblASTLIST1)
            End If
        End Using

        Dim tblASTLIST2 As New DataTable
        Using adaASTLIST2 As OracleDataAdapter = _
            ASCDATA1.GetDataAdapter(tblASTLIST2, "ASTLIST2", "*", True, -1, False)
            For Each CODE_VALUE As String In Split(CODE_VALUES, ",")
                Dim rowASTLIST2 As DataRow = tblASTLIST2.NewRow
                rowASTLIST2.Item("COLUMN_NAME") = COLUMN_NAME
                rowASTLIST2.Item("LIST_CODE") = LIST_CODE
                rowASTLIST2.Item("COLUMN_VALUE") = CODE_VALUE
                tblASTLIST2.Rows.Add(rowASTLIST2)
                adaASTLIST2.Update(tblASTLIST2)
            Next
        End Using


        MsgBox("Code List '" & txtList.Text & "' has been Saved", MsgBoxStyle.OkOnly, "Success")
        Call Load_Code_List(True)

    End Sub
#End Region

#Region "Settings"

    Private Sub chkMySettingsOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMySettingsOnly.CheckedChanged
        If chkMySettingsOnly.CheckState Then
            grdASTROPT1.Rows.ColumnFilters("INIT_OPER").FilterConditions.Add(Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Match, ASCMAIN1.USER_ID)
        Else
            grdASTROPT1.Rows.ColumnFilters("INIT_OPER").FilterConditions.Clear()
            grdASTROPT1.Rows.Refresh(Infragistics.Win.UltraWinGrid.RefreshRow.ReloadData)
        End If
    End Sub

    Private Sub cmdSaveSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveSettings.Click

        If SET_ID <> "" Then
            Dim rowASTROPT1 As DataRow = tblASTROPT1.Rows.Find(New Object() {FORM_NAME, SET_ID})
            If rowASTROPT1.Item("INIT_OPER") <> ASCMAIN1.USER_ID Then
                SET_ID = ""
            Else

                Select Case MsgBox("Update the Current Setting (Y) or Create a New One (N)?", MsgBoxStyle.YesNoCancel, "Save Setting Option")
                    Case MsgBoxResult.Yes
                    Case MsgBoxResult.No
                        SET_ID = ""
                    Case MsgBoxResult.Cancel
                        Exit Sub
                End Select
            End If
        End If

        Call Save_Settings(SET_ID)
        grdASTROPT1.ActiveRow = grdASTROPT1.Rows.GetRowWithListIndex(tblASTROPT1.Rows.IndexOf(tblASTROPT1.Rows.Find(New Object() {FORM_NAME, SET_ID})))
    End Sub

    Sub Retrieve_Settings()

        Dim SET_CTL_NAME As String
        Dim SET_CTL_TYPE As String
        Dim SET_CTL_TAG As String
        Dim SET_CTL_DATA As String

        Dim rowASTROPT1 As DataRow = tblASTROPT1.Rows.Find(New Object() {FORM_NAME, SET_ID})

        Dim sql As String

        If SET_ID = "0000000000" Then
            txtDescription.Text = ""
        Else
            txtDescription.Text = grdASTROPT1.ActiveRow.Cells("SET_DESC").Text
        End If

        Call Clear_grdSetup(True)
        tblASTRECAP.Rows.Clear()

        sql = "Select * from ASTROPT2 where FORM_NAME = '" & FORM_NAME & "'"
        sql = sql & " and SET_ID = '" & SET_ID & "'"
        sql = sql & " and XNO is Null"
        For Each rowASTROPT2 As DataRow In ASCDATA1.GetDataTable(sql).Select("", "SET_CTL_TAG")
            SET_CTL_NAME = rowASTROPT2.Item("SET_CTL_NAME") & ""
            SET_CTL_TYPE = rowASTROPT2.Item("SET_CTL_TYPE") & ""
            SET_CTL_TAG = rowASTROPT2.Item("SET_CTL_TAG") & ""
            SET_CTL_DATA = rowASTROPT2.Item("SET_CTL_DATA") & ""

            Dim gDR As DataRow

            If SET_CTL_NAME = "grdSetup" Then
                If SET_CTL_TAG = "" Then
                    Dim GRDCOLS() As String = Split(SET_CTL_DATA, vbTab)
                    gDR = DirectCast(grdSetup.DataSource, DataTable).Rows.Find(GRDCOLS(0))
                    If gDR IsNot Nothing Then
                        If Val(GRDCOLS(1) & "") <> 0 Then
                            gDR.Item("SEQUENCE") = Val(GRDCOLS(1) & "")
                        End If
                        gDR.Item("PAGE_BREAK") = GRDCOLS(2)
                        gDR.Item("EXCLUDE") = GRDCOLS(3)
                        gDR.Item("GROUP_ALL_OTHERS") = GRDCOLS(4)
                    End If
                Else
                    Dim COLUMN_NAME As String = SET_CTL_TAG
                    gDR = DirectCast(grdSetup.DataSource, DataTable).Rows.Find(COLUMN_NAME)
                    If gDR.Item("CODE_VALUES") & "" = "" Then
                        gDR.Item("CODE_VALUES") = SET_CTL_DATA
                    Else
                        gDR.Item("CODE_VALUES") &= "," & SET_CTL_DATA
                    End If
                End If

            ElseIf SET_CTL_NAME = "grdASTRECAP" Then
                tblASTRECAP.Rows.Add(Split(SET_CTL_DATA, vbTab))
            Else
                Dim C As Control = Absx1.CtlFor(SET_CTL_TAG, True)
                If C IsNot Nothing Then
                    Select Case SET_CTL_TYPE
                        Case "UltraCheckEditor"
                            Absx1.chkFor(SET_CTL_TAG).Checked = (SET_CTL_DATA = "True")
                        Case "UltraOptionSet"
                            Absx1.optFor(SET_CTL_TAG).Value = SET_CTL_DATA
                        Case "ABSCheckBox"
                            DirectCast(Absx1.CtlFor(SET_CTL_TAG), ABSCS.ABSCheckBox).ABSChecked = SET_CTL_DATA
                        Case "UltraCombo"
                            Dim cmbctl As UltraWinGrid.UltraCombo = DirectCast(Absx1.CtlFor(SET_CTL_TAG), UltraWinGrid.UltraCombo)
                            cmbctl.Text = SET_CTL_DATA
                            If SET_CTL_TAG = "RYP" Or SET_CTL_TAG = "RYP0" Or SET_CTL_TAG = "RYP1" Then
                                If SET_ID <> "0000000000" AndAlso rowASTROPT1.Item("SET_YP_REL") & "" = "1" And rowASTROPT1.Item("SET_YP_BASE") & "" <> "" Then
                                    Dim RYP As String = Mid(SET_CTL_DATA, 1, 4) & Mid(SET_CTL_DATA, 6, 2)
                                    Dim NP As Integer = ASCMAIN1.Period_Diff(rowASTROPT1.Item("SET_YP_BASE") & "", RYP)
                                    cmbctl.Text = Mid(ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, NP)), 1, 16)
                                End If
                            End If
                            If SET_CTL_TAG = "RYW" Or SET_CTL_TAG = "RYW0" Or SET_CTL_TAG = "RYW1" Then
                                If SET_ID <> "0000000000" AndAlso rowASTROPT1.Item("SET_YP_REL") & "" = "1" And rowASTROPT1.Item("SET_YW_BASE") & "" <> "" Then
                                    Dim RYW As String = Mid(SET_CTL_DATA, 1, 4) & Mid(SET_CTL_DATA, 6, 2)
                                    Dim NW As Integer = ASCMAIN1.Week_Diff(rowASTROPT1.Item("SET_YW_BASE") & "", RYW)
                                    cmbctl.Text = Mid(ASCMAIN1.Get_Legend_Wk(ASCMAIN1.Week_Calc(ASCMAIN1.CYW, NW)), 1, 17)
                                End If
                            End If
                        Case Else

                            Absx1.CtlFor(SET_CTL_TAG).Text = SET_CTL_DATA
                    End Select

                End If
            End If
        Next

        If grdSetup.Rows.Count > 0 Then
            grdSetup.ActiveRow = grdSetup.Rows(0)
        End If
        If grdASTRECAP.Rows.Count > 0 Then
            grdSetup.ActiveRow = grdSetup.Rows(0)
        End If

    End Sub

    Sub Save_Settings(ByRef SET_ID As String, Optional ByVal XNO As String = "")

        Call BeginTrans()

        Dim rowASTROPT1 As DataRow

        If XNO <> "" Then
            ' It is ok that ASTROPT1 does not get recorded 
            ' (although ASTROPT2 does get recorded) here.
            ' When XNO <> "", ASTOPST1 may serve as a "header" for ASTROPT2, 
            ' and in fact does, when we view Execution History
        Else
            Dim LAST_DATE As Date = Now + ASCMAIN1.NowTSD
            If SET_ID = "" Then
                rowASTROPT1 = tblASTROPT1.NewRow()
                SET_ID = ASCMAIN1.Next_Control_No("ASTROPT1.SET_ID")
                rowASTROPT1.Item("FORM_NAME") = FORM_NAME
                rowASTROPT1.Item("SET_ID") = SET_ID
                rowASTROPT1.Item("SET_YP_BASE") = ASCMAIN1.CYP
                rowASTROPT1.Item("SET_YP_REL") = "1"
                rowASTROPT1.Item("SET_ALLOW_OTHERS") = "0"
                rowASTROPT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowASTROPT1.Item("INIT_DATE") = LAST_DATE
                tblASTROPT1.Rows.Add(rowASTROPT1)
            Else
                rowASTROPT1 = tblASTROPT1.Rows.Find(New Object() {FORM_NAME, SET_ID})
                ASCMAIN1.sql = "Delete from ASTROPT2 " _
                    & " where FORM_NAME = '" & FORM_NAME & "'" _
                    & " and SET_ID = '" & SET_ID & "'" _
                    & " and XNO is Null"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            End If

            If SET_ID = "0000000000" Then
                'rowASTROPT1.Item("SET_DESC") = "{Defaults}"
            Else

                If txtDescription.Text = "" Then
                    rowASTROPT1.Item("SET_DESC") = "{Enter a Description for these Settings}"
                Else
                    rowASTROPT1.Item("SET_DESC") = txtDescription.Text
                End If
                rowASTROPT1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowASTROPT1.Item("LAST_DATE") = LAST_DATE
                adaASTROPT1.Update(tblASTROPT1)
            End If

        End If

        Dim rowASTROPT2 As DataRow
        Dim tblASTROPT2 As New DataTable
        Dim adaASTROPT2 As OracleDataAdapter = _
            ASCDATA1.GetDataAdapter(tblASTROPT2, "ASTROPT2", "*", True, 0, False)

        Call Save_Settings_ctls(UltraTabPageControl2, FORM_NAME, SET_ID, XNO, tblASTROPT2)

        For Each gr As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSetup.Rows
            rowASTROPT2 = tblASTROPT2.NewRow()
            With rowASTROPT2
                .Item("FORM_NAME") = FORM_NAME
                .Item("SET_ID") = SET_ID
                .Item("SET_CTL_NAME") = grdSetup.Name
                .Item("SET_CTL_TYPE") = grdSetup.GetType.Name
                .Item("SET_CTL_DATA") = gr.Cells("COLUMN_NAME").Text _
                            & vbTab & gr.Cells("SEQUENCE").Text _
                            & vbTab & gr.Cells("PAGE_BREAK").Text _
                            & vbTab & gr.Cells("EXCLUDE").Text _
                            & vbTab & gr.Cells("GROUP_ALL_OTHERS").Text
                .Item("SET_CTL_TAG") = ""
                .Item("XNO") = XNO
            End With
            tblASTROPT2.Rows.Add(rowASTROPT2)

            If gr.Cells("CODE_VALUES").Text <> "" Then
                Dim CODE_VALUES() As String = Split(gr.Cells("CODE_VALUES").Text, ",")
                For Each CODE_VALUE As String In CODE_VALUES
                    rowASTROPT2 = tblASTROPT2.NewRow()
                    With rowASTROPT2
                        .Item("FORM_NAME") = FORM_NAME
                        .Item("SET_ID") = SET_ID
                        .Item("SET_CTL_NAME") = grdSetup.Name
                        .Item("SET_CTL_TYPE") = grdSetup.GetType.Name
                        .Item("SET_CTL_DATA") = CODE_VALUE
                        .Item("SET_CTL_TAG") = gr.Cells("COLUMN_NAME").Text
                        .Item("XNO") = XNO
                    End With
                    tblASTROPT2.Rows.Add(rowASTROPT2)
                Next
            End If
        Next


        For Each gr As Infragistics.Win.UltraWinGrid.UltraGridRow In grdASTRECAP.Rows
            rowASTROPT2 = tblASTROPT2.NewRow()
            With rowASTROPT2
                .Item("FORM_NAME") = FORM_NAME
                .Item("SET_ID") = SET_ID
                .Item("SET_CTL_NAME") = grdASTRECAP.Name
                .Item("SET_CTL_TYPE") = grdASTRECAP.GetType.Name
                Dim ASTRECAP_row As String = ""
                For i As Integer = 0 To grdASTRECAP.DisplayLayout.Bands(0).Columns.Count - 1
                    ASTRECAP_row &= vbTab & gr.Cells(i).Value
                Next
                .Item("SET_CTL_DATA") = Mid(ASTRECAP_row, 2)
                .Item("SET_CTL_TAG") = ""
                .Item("XNO") = XNO
            End With
            tblASTROPT2.Rows.Add(rowASTROPT2)
        Next

        adaASTROPT2.Update(tblASTROPT2)

        tblASTROPT2.Dispose()
        adaASTROPT2.Dispose()

        Call CommitTrans()

        If XNO = "" And SET_ID <> "0000000000" Then
            MsgBox("Settings have been Saved", MsgBoxStyle.OkOnly, "Verification")
        End If

    End Sub

    Sub Save_Settings_ctls( _
    ByRef cc As System.Windows.Forms.Control, _
    ByRef FORM_NAME As String, _
    ByRef SET_ID As String, _
    ByRef XNO As String, _
    ByRef tblASTROPT2 As DataTable)

        Dim rowASTROPT2 As DataRow
        For Each ctl As System.Windows.Forms.Control In cc.Controls
            If ctl.Controls.Count > 0 Then
                Call Save_Settings_ctls(ctl, FORM_NAME, SET_ID, XNO, tblASTROPT2)
            End If
            Dim ABSCOLUMN_NAME As String = Absx1.GetABSColumnName(ctl)
            If ABSCOLUMN_NAME <> "" Then
                rowASTROPT2 = tblASTROPT2.NewRow()
                With rowASTROPT2
                    .Item("FORM_NAME") = FORM_NAME
                    .Item("SET_ID") = SET_ID
                    .Item("SET_CTL_NAME") = ctl.Name
                    .Item("SET_CTL_TYPE") = ctl.GetType.Name
                    Select Case ctl.GetType.Name
                        Case "UltraCheckEditor"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, Infragistics.Win.UltraWinEditors.UltraCheckEditor).Checked
                        Case "UltraOptionSet"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, Infragistics.Win.UltraWinEditors.UltraOptionSet).Value
                        Case "ABSCheckBox"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, ABSCS.ABSCheckBox).ABSChecked
                        Case Else
                            .Item("SET_CTL_DATA") = ctl.Text
                    End Select
                    .Item("SET_CTL_TAG") = ABSCOLUMN_NAME
                    .Item("XNO") = XNO
                End With

                tblASTROPT2.Rows.Add(rowASTROPT2)
            End If
        Next
    End Sub

    Sub Show_Settings()
        adaASTROPT1.Fill(tblASTROPT1)
        grdASTROPT1.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdASTROPT1.DisplayLayout.Bands(0).SortedColumns.Add(grdASTROPT1.DisplayLayout.Bands(0).Columns("SET_DESC"), False)
        'grdASTROPT1.DisplayLayout.Bands(0).SortedColumns.Add(grdASTROPT1.DisplayLayout.Bands(0).Columns("INIT_DATE"), True)
    End Sub
#End Region

#Region "Supporting Routines"

    Function SQL_in(ByVal COLUMN_NAME As String, _
    Optional ByVal DB_COLUMN_NAME As String = "") As String

        Dim CODE_VALUES = SQLA(COLUMN_NAME, "CODE_VALUES", True)
        Dim sql As String = ""

        If CODE_VALUES <> "" Then
            Dim single_code_value As Boolean = (InStr(CODE_VALUES, "','") = 0)

            If single_code_value Then
                sql = sql & IIf(SQLA(COLUMN_NAME, "EXCLUDE") = "1", " <> ", " = ") & CODE_VALUES
            Else
                sql = sql & IIf(SQLA(COLUMN_NAME, "EXCLUDE") = "1", " NOT", "") & " in (" & CODE_VALUES & ")"
            End If

            If DB_COLUMN_NAME <> "" Then
                sql = " and " & DB_COLUMN_NAME & sql
            Else
                sql = " and " & COLUMN_NAME & sql
            End If
        End If

        Return sql
    End Function

    Function SQLA( _
    ByVal PB_COLUMN_NAME As String, _
    Optional ByVal COLUMN_NAME As String = "CODE_VALUES", _
    Optional ByVal SQL_List As Boolean = False) As String
        Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find(PB_COLUMN_NAME)
        If rowASTDSQLA Is Nothing Then
            SQLA = ""
        Else
            SQLA = rowASTDSQLA.Item(COLUMN_NAME) & ""
            If SQL_List And SQLA <> "" Then
                SQLA = "'" & Replace(SQLA, ",", "','") & "'"
            End If
        End If
        Return SQLA
    End Function

    Function SQLA_filter( _
    ByVal PB_COLUMN_NAME As String, _
    Optional ByVal DB_TABLE_NAME As String = "", _
    Optional ByVal DB_COLUMN_NAME As String = "") As String

        If DB_TABLE_NAME <> "" Then
            If DB_COLUMN_NAME = "" Then
                DB_COLUMN_NAME = DB_TABLE_NAME & "." & PB_COLUMN_NAME
            Else
                DB_COLUMN_NAME = DB_TABLE_NAME & "." & DB_COLUMN_NAME
            End If
        End If

        If DB_COLUMN_NAME = "" Then
            DB_COLUMN_NAME = PB_COLUMN_NAME
        End If

        Dim z As String
        z = SQLA(PB_COLUMN_NAME, "CODE_VALUES", True)
        If z <> "" Then
            SQLA_filter = " AND " & DB_COLUMN_NAME & IIf(SQLA(PB_COLUMN_NAME, "EXCLUDE") = "1", " NOT", "") & " IN (" & z & ")" & vbCr
        Else
            SQLA_filter = ""
        End If
        Return SQLA_filter
    End Function

    Function Get_Filter(ByVal COLUMN_NAME As String, ByVal SQL_ELEMENT_TO_COMPARE_TO As String) As String
        Dim sqlw As String = ""
        If SQLA(COLUMN_NAME, "CODE_VALUES") <> "" Then
            sqlw = " and " & SQL_ELEMENT_TO_COMPARE_TO & " " & IIf(SQLA(COLUMN_NAME, "EXCLUDE") = "1", "Not ", "") & "in (" & SQLA(COLUMN_NAME, "CODE_VALUES", True) & ")"
        End If
        Return sqlw

    End Function

#End Region

    Private Sub cmdAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAll.Click
        grdSetup.ActiveRow.Cells("CODE_VALUES").Value = ""
        DirectCast(grd.DataSource, DataTable).Rows.Clear()
        cmdAll.Visible = False
    End Sub

End Class
