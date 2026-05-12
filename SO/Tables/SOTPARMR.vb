Public Class SOTPARMR
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            ASCMAIN1.sql = "Select *" _
                & " from SOTPARM2 " _
                & " WHERE SO_PARM_KEY = :PARM1"
            Create_TDA(.Tables.Add, "SOTPARM2", "**", 0, True, "V", 1)
        End With

    End Sub

    Overrides Sub Show_Record_Special()
        Fill_Records("SOTPARM2", "Z")
        If dst.Tables.Item("SOTPARMR").Rows.Count = 1 And dst.Tables.Item("SOTPARM2").Rows.Count = 1 Then
            For Each UCOL As String In New String() {"SO_PARM_CONCOST", "SO_PARM_DUTY", "SO_PARM_INLANDFRT", "SO_PARM_OCEANFRTCONS", "SO_PARM_OCEANFRT"}
                dst.Tables.Item("SOTPARMR").Rows(0).Item(UCOL) = dst.Tables.Item("SOTPARM2").Rows(0).Item(UCOL).ToString & String.Empty
            Next
            Absx1.numFor("SO_PARM_FEFACT").Value = Val(dst.Tables.Item("SOTPARM2").Rows(0).Item("SO_PARM_FEFACT").ToString & String.Empty)
        End If
    End Sub
    Overrides Sub Proceed_Update_Special_Pre()
        If dst.Tables.Item("SOTPARMR").Rows.Count = 1 And dst.Tables.Item("SOTPARM2").Rows.Count = 1 Then
            For Each UCOL As String In New String() {"SO_PARM_CONCOST", "SO_PARM_DUTY", "SO_PARM_INLANDFRT", "SO_PARM_OCEANFRTCONS", "SO_PARM_OCEANFRT"}
                dst.Tables.Item("SOTPARM2").Rows(0).Item(UCOL) = dst.Tables.Item("SOTPARMR").Rows(0).Item(UCOL).ToString & String.Empty
            Next
            dst.Tables.Item("SOTPARM2").Rows(0).Item("SO_PARM_FEFACT") = Val(Absx1.numFor("SO_PARM_FEFACT").Value & String.Empty)
        End If

        Update_Record_TDA("SOTPARM2")
    End Sub
End Class