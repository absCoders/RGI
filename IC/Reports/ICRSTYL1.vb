Public Class ICRSTYL1
    Dim ICTSTYL1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        Prepare_Work_File()

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        ASCMAIN1.sql = "Select ICTSTYL1.* " _
            & " from ICTSTYL1"

        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 1))

        Dim SOURCE_TABLE_NAME As String = "ICTSTYL1"
        MyBase.Get_SQL("*", SOURCE_TABLE_NAME)

        sql = "Select " & sql_SELECT_cols & vbCrLf _
        & ", " & SOURCE_TABLE_NAME & ".STYLE_CODE" _
        & " from " & ICTSTYL1 & " " & SOURCE_TABLE_NAME & " " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf

        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")
    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = ""
        Generate_Report(RPT, , SUBT)
    End Sub

    Sub Prepare_Work_File()
        ICTSTYL1 = "ICTSTYL1"
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            'If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
            '    EMsg &= vbCr & "You must Specify at least 1 Sort Field"
            'End If
        End If
    End Sub
End Class