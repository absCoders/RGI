Public Class SORRTRNS

#Region "Declarations"

    Dim SOTINVH2 As String

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""
        Dim sql_filter As String = ""
        Dim sqlX As String = ""
        Dim sqlYP As String = ""

        ' Extracts from Data Sources

        ASCMAIN1.Progress("Returns")
        MyBase.Get_SQL("*")


        sqlYP = "SOTRTRN1.OPS_YYYYPP"

        sql_Data = "" _
            & ", SUM (NVL(SOTRTRN2.RTRN_QTY,0)) QTY" & vbCrLf _
            & ", SUM (NVL(SOTRTRN2.RTRN_QTY,0) * NVL(SOTRTRN2.RTRN_PRICE,0)) AMT" & vbCrLf

        sql_Cols = "" _
            & ",QTY,AMT"

        sql_filter = "" _
            & " and " & sqlYP & " between '" & RYP0 & "' AND '" & RYP1 & "'" & vbCrLf

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & "" & vbCrLf & sql_Data _
            & " from SOTRTRN2" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()

    End Sub

    Public Overrides Sub Print_Report()
        ' CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "1", "0"))
        SUBT = ""
        If RYP0 = RYP1 Then
            SUBT = "Returns Posted in " & RYPLEGEND0
        Else
            SUBT = "Returns Posted in " & RYPLEGEND0 & " thru " & RYPLEGEND1
        End If
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            'If Absx1.cmbFor("RYP").Value & "" = "" Then
            '    EMsg &= vbCr & "You must Specify a Report Period"
            'End If
        End If
    End Sub

    Overrides Sub Verify_Special_Pre(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
                EMsg &= vbCr & "You must Specify at least 1 Sort Field"
            End If
        End If
    End Sub
End Class