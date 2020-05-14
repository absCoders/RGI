Imports Infragistics.Win

Public Class WBTPAGE1

    Private PAGE_CODE As String = String.Empty

    Private Sub WBTPAGE1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "ICTSTYLW", "*", 2, True, String.Empty, 2)
            With dst.Tables("ICTSTYLW")
                .Columns.Add("STYLE_DESC", GetType(System.String))
                .Columns.Add("STYLE_STATUS", GetType(System.String))
            End With
        End With

        grdICTSTYLW.DataSource = dst.Tables("ICTSTYLW")
        Create_Summary(grdICTSTYLW, "STYLE_CODE", "Count")

        With grdICTSTYLW.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Header.Appearance.ForeColor = System.Drawing.Color.White
            .Columns("STYLE_CODE").Header.Appearance.BackColor2 = System.Drawing.Color.Blue
            .Columns("STYLE_CODE").Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        End With
    End Sub

    Public Overrides Sub Proceed_Update_Special_Pre()
        MyBase.Proceed_Update_Special_Pre()

        Dim sql As String = String.Empty
        Dim rowWBTPAGE1 As DataRow = dst.Tables("WBTPAGE1").Select("PAGE_CODE = '" & MyBase.Absx1.txtFor("PAGE_CODE").Text & "'")(0)

        If rowWBTPAGE1.RowState <> DataRowState.Added AndAlso (rowWBTPAGE1.Item("PAGE_STATUS", DataRowVersion.Current) <> rowWBTPAGE1.Item("PAGE_STATUS", DataRowVersion.Original)) Then
            sql = "UPDATE WBTSTYL1 SET WEB_IND = '1'"
            sql &= " WHERE STYLE_CODE IN "
            sql &= " (SELECT STYLE_CODE FROM ICTSTYLW WHERE PAGE_CODE = '" & MyBase.Absx1.txtFor("PAGE_CODE").Text & "')"
            ASCDATA1.ExecuteSQL(sql)
        End If

        For Each rowICTSTYLW As DataRow In dst.Tables("ICTSTYLW").Select("", "", DataViewRowState.Deleted)
            sql = "UPDATE WBTSTYL1 SET WEB_IND = '1'"
            sql &= " WHERE STYLE_CODE = '" & rowICTSTYLW.Item("STYLE_CODE", DataRowVersion.Original) & "'"
            ASCDATA1.ExecuteSQL(sql)
        Next

        For Each rowICTSTYLW As DataRow In dst.Tables("ICTSTYLW").Select("", "", DataViewRowState.ModifiedCurrent)
            sql = "UPDATE WBTSTYL1 SET WEB_IND = '1'"
            sql &= " WHERE STYLE_CODE = '" & rowICTSTYLW.Item("STYLE_CODE", DataRowVersion.Original) & "'"
            ASCDATA1.ExecuteSQL(sql)
        Next

        For Each rowICTSTYLW As DataRow In dst.Tables("ICTSTYLW").Select("", "", DataViewRowState.Added)
            sql = "UPDATE WBTSTYL1 SET WEB_IND = '1'"
            sql &= " WHERE STYLE_CODE = '" & rowICTSTYLW.Item("STYLE_CODE") & "'"
            ASCDATA1.ExecuteSQL(sql)
        Next

        For Each rowICTSTYLW As DataRow In dst.Tables("ICTSTYLW").Select("", "", DataViewRowState.CurrentRows)
            sql = "UPDATE WBTSTYL1 SET WEB_IND = '1'"
            sql &= " WHERE STYLE_CODE = '" & rowICTSTYLW.Item("STYLE_CODE") & "'"
            ASCDATA1.ExecuteSQL(sql)
        Next


        Update_Record_TDA("ICTSTYLW", "DELETE FROM ICTSTYLW WHERE PAGE_CODE = '" & MyBase.Absx1.txtFor("PAGE_CODE").Text & "'")

    End Sub

    Overrides Sub Show_Record_Special()

        PAGE_CODE = MyBase.Absx1.txtFor("PAGE_CODE").Text.Trim

        Dim sql As String = String.Empty

        MyBase.EnforceConstraints(False)

        Dim SB As New Text.StringBuilder
        SB.Length = 0
        SB.AppendLine("SELECT ICTSTYLW.*, ICTSTYL1.STYLE_DESC, WBTSTYL1.STYLE_STATUS")
        SB.AppendLine("FROM ICTSTYLW, WBTSTYL1, ICTSTYL1")
        SB.AppendLine("WHERE ICTSTYLW.STYLE_CODE = WBTSTYL1.STYLE_CODE")
        SB.AppendLine("AND WBTSTYL1.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        SB.AppendLine("AND ICTSTYLW.PAGE_CODE = '" & PAGE_CODE & "'")
        Call Fill_Records("ICTSTYLW", String.Empty, True, SB.ToString)

        grdICTSTYLW.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdICTSTYLW.DisplayLayout.Bands(0).SortedColumns.Add("STYLE_CODE", False)

    End Sub

    Overrides Sub Clear_Record_Special()

        If ScreenMode Then
            MyBase.EnforceConstraints(False)
            dst.Tables("ICTSTYLW").Rows.Clear()
            MyBase.EnforceConstraints(True)

            PAGE_CODE = String.Empty
        End If

    End Sub

    Private Sub grdICTSTYLW_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTSTYLW.BeforeRowUpdate
        e.Row.Cells("PAGE_CODE").Value = PAGE_CODE

        Dim STYLE_CODE As String = (e.Row.Cells("STYLE_CODE").Value & String.Empty).ToString.ToUpper.Trim
        Dim styleInList As Boolean = dst.Tables("ICTSTYLW").Select("STYLE_CODE = '" & STYLE_CODE & "'").Length > 0
        e.Row.Cells("STYLE_CODE").Value = STYLE_CODE

        Dim rowWBTSTYL1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM WBTSTYL1 WHERE STYLE_CODE = :PARM1", "V", New String() {STYLE_CODE})

        If rowWBTSTYL1 Is Nothing Then
            e.Cancel = True
            MessageBox.Show("Invalid or missing Style Code.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If e.Row.IsAddRow AndAlso styleInList Then
                e.Cancel = True
                MessageBox.Show("Style already belongs to the Page.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf Not e.Row.IsAddRow AndAlso styleInList AndAlso _
                e.Row.Cells("STYLE_CODE").Value & String.Empty <> e.Row.Cells("STYLE_CODE").OriginalValue & String.Empty Then
                e.Cancel = True
                MessageBox.Show("Style already belongs to the Page.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowWBTSTYL1.Item("STYLE_CODE") & String.Empty)
                e.Row.Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC") & String.Empty
                e.Row.Cells("STYLE_STATUS").Value = rowICTSTYL1.Item("STYLE_STATUS") & String.Empty
            End If
        End If
    End Sub

    Private Sub grdICTSTYLW_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYLW.ClickCellButton
        If grdICTSTYLW.ActiveRow Is Nothing Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        Dim sql_where As String = ""
        Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
        Dim VIEW_NAME As String = String.Empty

        Select Case grd.ActiveCell.Column.Key
            Case "STYLE_CODE"
                sql_where = "STYLE_STATUS = 'A'"
                VIEW_NAME = "STYLE_CODE.WBTSTYL1"

        End Select

        Call grdClickCellButton(grdICTSTYLW, sql_where, True)
    End Sub

End Class