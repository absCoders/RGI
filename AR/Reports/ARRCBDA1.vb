Public Class ARRCBDA1

    ' pgm does not take into account DISC & WOFF from ARTPYMT3, nor does it take into account GL write-offs

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -48, 0, -11)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -48, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        ' Prepare Work Tables

        ASCMAIN1.Progress("Work Tables")

        Dim ARTCBDA1 As String = ""
        TAC.ARCMAIN1.Create_ARTCBDA1(ARTCBDA1, RYP0, RYP1)

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""

        ASCMAIN1.Progress("Analysis")
        MyBase.Get_SQL("*")

        sql_Data = "" _
            & ", SUM (BEG_B) BEG_B" & vbCrLf _
            & ", SUM (NEW_B) NEW_B" & vbCrLf _
            & ", SUM (APP_B) APP_B" & vbCrLf _
            & ", SUM (END_B) END_B" & vbCrLf _
            & ", SUM (BEG_C) BEG_C" & vbCrLf _
            & ", SUM (NEW_C) NEW_C" & vbCrLf _
            & ", SUM (APP_C) APP_C" & vbCrLf _
            & ", SUM (END_C) END_C" & vbCrLf _
            & ", SUM (NEW_X) NEW_X" & vbCrLf

        sql_Cols = "" _
            & ",BEG_B,NEW_B,APP_B,END_B,BEG_C,NEW_C,APP_C,END_C,NEW_X"

        sql_filter = ""

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & "" & vbCrLf & sql_Data _
            & " from " & ARTCBDA1 & " ARTCBDA1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()
    End Sub

    Public Overrides Sub Print_Report()
        Dim i As Integer
        If RYP0 = RYP1 Then
            SUBT = RYPLEGEND0
        Else
            SUBT = RYPLEGEND0 & " thru " & RYPLEGEND1
        End If

        Generate_Report(RPT, , SUBT)

        Dim tbl As DataTable = dst.Tables("ASTSRPT1").Copy

        For Each row As DataRow In tbl.Rows
            For i = 1 To COLUMN_NAMEs.Count
                row.Item(i - 1) = Split(row.Item(i - 1), ":")(1)
            Next i
        Next

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        grdASTEXPT1.DataSource = tbl

        grdASTEXPT1.Text = "Chargeback & Deduction Analysis"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Dim SortColumns As String = ""
        For grdcol As Integer = 1 To 9
            If grdcol <= COLUMN_NAMEs.Count Then
                Set_DX_Column(grdASTEXPT1, String.Format("G{0}", grdcol), COLUMN_CAPTIONs(grdcol - 1), , , , System.Drawing.Color.LightGoldenrodYellow)
                SortColumns = SortColumns & "," & String.Format("G{0}", grdcol)
            Else
                Set_DX_Column(grdASTEXPT1, String.Format("G{0}", grdcol), "", 0)
            End If
        Next grdcol

        For Each col As DataRow In tblASTDSQLS.Rows
            Set_DX_Column(grdASTEXPT1, col.Item(2), col.Item(5), col.Item(6), , , System.Drawing.Color.Gold)
        Next

        Sort_grdColumns(grdASTEXPT1, Mid(SortColumns, 2))

    End Sub

End Class