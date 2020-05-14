Public Class SORRTRA1

    Dim SOTRTRN1 As String
    Dim SOTRTRN2 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0), -60, 0, 0)

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        SOTRTRN2 = ASCMAIN1.Temp_Table("Select SOTRTRN2.* from SOTRTRN2 where OPS_YYYYPP = '" & RYP & "'")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTRTRN2 & " Add Primary Key (RTRN_NO, RTRN_LNO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTRTRN2 & "_1 ON " & SOTRTRN2 & " (STYLE_CODE, COLOR_CODE)")

        Dim sqlMTD As String = "ORDR_YYYYPP_UPDATED = '" & RYP & "' "
        Dim sqlYTD As String = "ORDR_YYYYPP_UPDATED >= '" & Mid(RYP, 1, 4) & "01' AND ORDR_YYYYPP_UPDATED <= '" & RYP & "'"
        Dim sqlQTY As String = " THEN NVL(ORDR_QTY_SHIP,0) ELSE 0 END"
        Dim sqlAMT As String = " THEN NVL(ORDR_QTY_SHIP,0) * NVL(ORDR_UNIT_PRICE,0) ELSE 0 END"

        ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE" _
            & ", SUM (CASE WHEN INV_TYPE = 'I' AND " & sqlMTD & sqlQTY & ") MTD_SLS_QTY" & vbCrLf _
            & ", SUM (CASE WHEN INV_TYPE = 'I' AND " & sqlMTD & sqlAMT & ") MTD_SLS_AMT" & vbCrLf _
            & ", SUM (CASE WHEN INV_TYPE = 'C' AND " & sqlMTD & sqlQTY & ") MTD_RTN_QTY" & vbCrLf _
            & ", SUM (CASE WHEN INV_TYPE = 'C' AND " & sqlMTD & sqlAMT & ") MTD_RTN_AMT" & vbCrLf _
            & ", SUM (CASE WHEN INV_TYPE = 'I' AND " & sqlYTD & sqlQTY & ") YTD_SLS_QTY" & vbCrLf _
            & ", SUM (CASE WHEN INV_TYPE = 'I' AND " & sqlYTD & sqlAMT & ") YTD_SLS_AMT" & vbCrLf _
            & ", SUM (CASE WHEN INV_TYPE = 'C' AND" & sqlYTD & sqlQTY & ") YTD_RTN_QTY" & vbCrLf _
            & ", SUM (CASE WHEN INV_TYPE = 'C' AND " & sqlYTD & sqlAMT & ") YTD_RTN_AMT" & vbCrLf _
            & " from SOTINVH2 where ORDR_YYYYPP_UPDATED >= '" & Mid(RYP, 1, 4) & "01' and ORDR_YYYYPP_UPDATED <= '" & RYP & "'" _
            & " and (STYLE_CODE, COLOR_CODE) in (Select Distinct STYLE_CODE, COLOR_CODE from " & SOTRTRN2 & " SOTRTRN2)"

        Create_TDA(dst.Tables.Add, "SOTRTRA1", "**", 0, False, "", 3)
        Fill_Records("SOTRTRA1")

    End Sub

    Public Overrides Sub Print_Report()
        SUBT = RYPLEGEND

        RPT = "SORRTRA1"
        RPT_TITLE = "Style Returns Analysis"
        CR_params.Add("SUBT", "")
        Generate_Report(RPT, , SUBT)

        'RPT = "SORRTRA2"
        'RPT_TITLE = "Customer Returns Analysis"
        'CR_params.Add("SUBT", "")
        'Generate_Report(RPT, RPT_TITLE, SUBT)

        'RPT = "SORRTRA3"
        'RPT_TITLE = "Vendor Returns Analysis"
        'CR_params.Add("SUBT", "")
        'Generate_Report(RPT, RPT_TITLE, SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub
End Class