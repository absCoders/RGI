Public Class SORSCOMR

    Dim SOTINVH1 As String
    Dim SOTINVH2 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0), -60, 0, 0)

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            chkROYALTY.Visible = True
            chkROYALTY.Checked = True
        Else
            chkROYALTY.Visible = False
            chkROYALTY.Checked = False
        End If
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()
        RWU = "R"

        If ASCMAIN1.EOM <> "1" Then
            RWU = "N"
        End If



        ASCMAIN1.sql = "Select * from SOTINVH1 where ORDR_YYYYPP_UPDATED = '" & RYP & "'"

        ASCMAIN1.sql &= SQL_in("CUST_CODE", "SOTINVH1.CUST_CODE")
        ASCMAIN1.sql &= SQL_in("SREP_CODE", "SOTINVH1.SREP_CODE")
        ASCMAIN1.sql &= SQL_in("WHSE_CODE", "SOTINVH1.WHSE_CODE")

        SOTINVH1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH1 & " Add Primary Key (INV_TYPE, INV_NO)")

        ASCMAIN1.sql = "Select SOTINVH2.*" _
            & " from SOTINVH2, " & SOTINVH1 & " SOTINVH1" & vbCrLf _
            & " where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH2.INV_NO = SOTINVH1.INV_NO and NVL(ORDR_QTY_SHIP,0) <> 0 "
        SOTINVH2 = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = "Select * from " & SOTINVH1
        Create_TDA(dst.Tables.Add("SOTINVH1"), SOTINVH1, "**", 0, False, "", 2)
        With dst.Tables("SOTINVH1").Columns
            .Add("HAS_ROYATLY", GetType(System.String))
        End With
        Fill_Records("SOTINVH1")
        'dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTINVH1", 2))

        ASCMAIN1.sql = "Select * from " & SOTINVH2
        Create_TDA(dst.Tables.Add("SOTINVH2"), SOTINVH2, "**", 0, False, "", 3)
        With dst.Tables("SOTINVH2").Columns
            .Add("IS_ROYATLY", GetType(System.String))
            .Add("ROYALTY_PCT", GetType(System.Decimal))
        End With
        Fill_Records("SOTINVH2")
        ' dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTINVH2", 3))
        If chkROYALTY.Checked Then
            FILL_ROYALTY()
        End If

        ASCMAIN1.sql = "Select SREP_CODE, SREP_NAME from SOTSREP1" & vbCrLf _
            & " where SREP_CODE in (Select Distinct SREP_CODE from " & SOTINVH1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSREP1", 1))

        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME from ARTCUST1" & vbCrLf _
            & " where CUST_CODE in (Select Distinct CUST_CODE from " & SOTINVH1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST1", 1))


        ASCMAIN1.sql = "Select WHSE_CODE, WHSE_DESC from ICTWHSE1" & vbCrLf _
        & " where WHSE_CODE in (Select Distinct WHSE_CODE from " & SOTINVH1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTWHSE1", 1))

        ASCMAIN1.sql = "Select STYLE_CODE, STYLE_DESC from ICTSTYL1" & vbCrLf _
            & " where STYLE_CODE in (Select Distinct STYLE_CODE from " & SOTINVH2 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 1))

        'ASCMAIN1.Progress("Now Calculating FIFO Costs")
        'TAC.ICCMAIN1.Calculate_FIFO_Cost_OH(Me, RYP)
        'ASCMAIN1.Progress("")
    End Sub

    Private Sub FILL_ROYALTY()
        Dim INV_NO_LIST As New List(Of String)
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("")
        sql.AppendLine("SELECT")
        sql.AppendLine("I2.INV_NO,")
        sql.AppendLine("I2.STYLE_CODE,")
        sql.AppendLine("I2.COLOR_CODE,")
        sql.AppendLine("MIN(R2.ROYALTY_PCT) AS ROYALTY_PCT")
        sql.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2, ICTSTYL1 S1, SOTORDR1 O1, ICTROYL2 R2")
        sql.AppendLine("WHERE I1.INV_NO = I2.INV_NO")
        sql.AppendLine("AND I2.STYLE_CODE = S1.STYLE_CODE")
        sql.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO")
        sql.AppendLine("AND NVL(S1.ROYALTY_CODE,'NULL') <> 'NULL'")
        sql.AppendLine("AND S1.ROYALTY_CODE = R2.ROYALTY_CODE")
        sql.AppendLine("AND (R2.ROYALTY_BEGIN <= O1.ORDR_DATE AND R2.ROYALTY_END >= O1.ORDR_DATE)")
        sql.AppendLine("GROUP BY")
        sql.AppendLine("I2.INV_NO,")
        sql.AppendLine("I2.STYLE_CODE,")
        sql.AppendLine("I2.COLOR_CODE")
        Dim tblROYALTY As DataTable = ASCDATA1.GetDataTable(sql.ToString())
        For Each rowSOTINVH2 As DataRow In dst.Tables("SOTINVH2").Select()
            Dim INV_NO As String = rowSOTINVH2.Item("INV_NO").ToString & String.Empty
            Dim STYLE_CODE As String = rowSOTINVH2.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowSOTINVH2.Item("COLOR_CODE").ToString & String.Empty
            Dim FLTR As String = $"INV_NO = '{INV_NO}' AND STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE ='{COLOR_CODE}'"
            Dim rowROYALTY As DataRow = tblROYALTY.Select(FLTR).FirstOrDefault
            If Not IsNothing(rowROYALTY) Then
                rowSOTINVH2.Item("IS_ROYATLY") = "1"
                rowSOTINVH2.Item("ROYALTY_PCT") = rowROYALTY.Item("ROYALTY_PCT").ToString & String.Empty
                If Not INV_NO_LIST.Contains(INV_NO) Then
                    INV_NO_LIST.Add(INV_NO)
                End If
            Else
                rowSOTINVH2.Item("IS_ROYATLY") = "0"
                rowSOTINVH2.Item("ROYALTY_PCT") = 0
            End If
        Next
        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select()
            If INV_NO_LIST.Contains(rowSOTINVH1.Item("INV_NO").ToString & String.Empty) Then
                rowSOTINVH1.Item("HAS_ROYATLY") = "1"
            Else
                rowSOTINVH1.Item("HAS_ROYATLY") = "0"
            End If
        Next
    End Sub

    Public Overrides Sub Print_Report()

        ' CR_params.Add("SORT_CUST", IIf(tblASTDSQLA.Rows.Find("CUST_CODE").Item("SEQUENCE") & "" = "", "0", "1"))
        ' CR_params.Add("SORT_SREP", IIf(tblASTDSQLA.Rows.Find("SREP_CODE").Item("SEQUENCE") & "" = "", "0", "1"))
        ' CR_params.Add("SORT_WHSE", IIf(tblASTDSQLA.Rows.Find("WHSE_CODE").Item("SEQUENCE") & "" = "", "0", "1"))

        CR_params.Add("SUBT", "")
        Generate_Report(RPT, , SUBT)
        If chkROYALTY.Checked Then
            CR_params.Add("SUBT", "Deductions For Royalty")
            Generate_Report("SORSCOML", , SUBT)
        End If
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub
End Class