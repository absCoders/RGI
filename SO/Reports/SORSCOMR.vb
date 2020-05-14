Public Class SORSCOMR

    Dim SOTINVH1 As String
    Dim SOTINVH2 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0), -60, 0, 0)

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
        Fill_Records("SOTINVH1")
        'dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTINVH1", 2))

        ASCMAIN1.sql = "Select * from " & SOTINVH2
        Create_TDA(dst.Tables.Add("SOTINVH2"), SOTINVH2, "**", 0, False, "", 3)
        Fill_Records("SOTINVH2")
        ' dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTINVH2", 3))


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

    Public Overrides Sub Print_Report()

        ' CR_params.Add("SORT_CUST", IIf(tblASTDSQLA.Rows.Find("CUST_CODE").Item("SEQUENCE") & "" = "", "0", "1"))
        ' CR_params.Add("SORT_SREP", IIf(tblASTDSQLA.Rows.Find("SREP_CODE").Item("SEQUENCE") & "" = "", "0", "1"))
        ' CR_params.Add("SORT_WHSE", IIf(tblASTDSQLA.Rows.Find("WHSE_CODE").Item("SEQUENCE") & "" = "", "0", "1"))

        CR_params.Add("SUBT", "")
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub
End Class