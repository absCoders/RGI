Imports System.Text

Public Class SORGPAR1
    Dim S As New StringBuilder With {.Length = 0}

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        RWU = "N"
    End Sub

    Protected Overrides Sub Build_Workfile()


        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        SUBT = ""
        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        ASCMAIN1.Progress("Now Preparing Dataset")
        Get_SQL("*")
        Dim sql_TABLE_NAMEs_orig As String = sql_TABLE_NAMEs
        Dim sql_JOIN_orig As String = sql_JOIN

        Dim sql_filter2 As String = ""

        '-- Shit you may need here --
        'sql_SELECT_cols, sql_TABLE_NAMEs, sql_WHERE, sql_JOIN, sql_filter, sql_filter2
        S.Length = 0
        S.AppendLine("Select " & sql_SELECT_cols)
        S.AppendLine(sql_WHERE)
        S.AppendLine(sql_filter2)
        S.AppendLine("GROUP BY")
        If sql_GROUP_BY_cols.Length > 0 Then
            S.AppendLine(sql_GROUP_BY_cols & ",")
        End If
        ASCMAIN1.sql = S.ToString()

        With dst
            'Dim SOTSUPP1 As String = ASCMAIN1.Temp_Table("Select * from SOTSUPP1")
            'ASCMAIN1.sql = "Select * from " & SOTSUPP1
            'Create_TDA(.Tables.Add, "SOTSUPP1", "**", 0, False)

            'S.Length = 0
            'S.AppendLine("SELECT")
            'ASCMAIN1.sql = S.ToString()
            'Create_TDA(.Tables.Add, "SOTNETPS", "**", 0, False)
            'Fill_Records("SOTNETPS")
        End With

        Prepare_dst(True, sql_filter)
    End Sub

    Public Overrides Sub Build_Report_File_Pre_Ora2ADO(TT As String)
        MyBase.Build_Report_File_Pre_Ora2ADO(TT)
    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()

        'For i As Integer = 1 To 3
        '    Dim iFormat As String = Format(i, "00")
        '    dst.Tables.Item("ASTSRPT1").Columns.Add("ALLOC_DATE_" & iFormat, GetType(System.DateTime))
        '    dst.Tables.Item("ASTSRPT1").Columns.Add("ALLOC_QTY_" & iFormat, GetType(System.Double))
        'Next

        For Each rowASTGROUP As DataRow In dst.Tables("ASTGROUP").Select("GROUP_KEY = 'Customer:' and (GROUP_CODE = '' or GROUP_CODE = 'STOCK')")
            rowASTGROUP.Item("GROUP_CODE") = "STOCK"
            rowASTGROUP.Item("GROUP_DESC") = "Stock Item"
        Next

        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select("", "STYLE_CODE, COLOR_CODE")
            Dim STYLE_CODE As String = rowASTSRPT1.Item("STYLE_CODE").ToString()
            Dim COLOR_CODE As String = rowASTSRPT1.Item("COLOR_CODE").ToString()
        Next

    End Sub

    Public Overrides Sub Print_Report()
        'CalculatePeriods()

        'CR_params.Add("SUBT", txtDescription.Text & SUBT)
        RPT = "SORGPAR1"

        CR_params.Add("SUBT", txtDescription.Text & SUBT)

        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            'If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length > 4 Then
            '    EMsg &= vbCr & "Maximum number of Sort Fields for this report is 4"
            'End If
        End If
    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        With dst

        End With

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        If parms.Length > 0 Then
        End If

        EnforceConstraints(False)
        'Fill_Records("ASTSRPT1")
        EnforceConstraints(True)
    End Sub
End Class