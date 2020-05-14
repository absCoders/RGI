Imports System.Math

Public Class APR1099E
    Dim Report_Subt As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("APTPARM1")

        ' Range_Events(grpCHK_DATE_RANGE)
        optSHOW.Value = "W"
        Absx1.txtFor("TIN").Text = ROWs("APTPARM1").Item("AP_PARM_1099_TAX_ID") & ""
        Absx1.txtFor("TCC").Text = "" ' "19K37"
        Absx1.numFor("CUTOFF").Value = Val(ROWs("APTPARM1").Item("AP_PARM_1099_LIMIT") & "")

        Dim YYYY = Now.Date.AddMonths(-6).Year
        Absx1.dteFor("CHK_DATE_F").Value = "01/01/" & YYYY
        Absx1.dteFor("CHK_DATE_L").Value = "12/31/" & YYYY
    End Sub

    Protected Overrides Sub Build_Workfile()
        With dst
            Dim SQLX As String = ""
            SQLX = " FROM APTCHCK1, APTCHCK2, APTINVH1, APTVEND1" _
            & "  WHERE APTCHCK1.CHECK_STATUS = 'I'" _
            & " AND APTCHCK1.CHECK_DATE >= '" & Format(Absx1.dteFor("CHK_DATE_F").Value, "dd-MMM-yyyy") & "'" _
            & " AND APTCHCK1.CHECK_DATE <= '" & Format(Absx1.dteFor("CHK_DATE_L").Value, "dd-MMM-yyyy") & "'" _
            & IIf(optSHOW.Value = "O", " AND APTVEND1.VEND_TAX_ID IS NULL", "") _
            & IIf(optSHOW.Value = "W", " AND APTVEND1.VEND_TAX_ID IS NOT NULL", "") _
            & IIf(chkDTL.Checked = False, "    AND APTINVH1.INV_1099_AMT <> 0", "") _
            & "    AND APTCHCK1.BANK_CODE = APTCHCK2.BANK_CODE" _
            & "    AND APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM" _
            & "    AND APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" _
            & "    AND APTVEND1.VEND_CODE = APTCHCK1.VEND_CODE_AP"


            ASCMAIN1.Progress("Compiling Check data", "")

            ASCMAIN1.sql = "SELECT DISTINCT APTCHCK1.* " & SQLX
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTCHCK1", 2))

            ASCMAIN1.sql = "SELECT DISTINCT APTCHCK2.* " & SQLX
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTCHCK2", 3))

            ASCMAIN1.Progress("Evaluating Invoice data", "")
            ASCMAIN1.sql = "SELECT APTINVH1.*, DECODE(APTINVH1.INV_AMT,0,0, " _
            & " APTINVH1.INV_1099_AMT * APTCHCK2.INV_AMT_APPLIED / APTINVH1.INV_AMT) PMT_1099 " & SQLX
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTINVH1", 1))

            ASCMAIN1.Progress("Evaluating Vendor data", "")
            ASCMAIN1.sql = "SELECT DISTINCT APTVEND1.* " & SQLX
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTVEND1", 1))


            ASCMAIN1.Progress("Merging Check, Invoice and Vendor data", "")
            ASCMAIN1.sql = " SELECT APTCHCK1.VEND_CODE_AP, " _
            & " SUM(DECODE(APTINVH1.INV_AMT,0,0,APTINVH1.INV_1099_AMT * APTCHCK2.INV_AMT_APPLIED / APTINVH1.INV_AMT)) AS PMT_1099, " _
            & " '1' AS PRINT_IND" & SQLX _
            & " GROUP BY APTCHCK1.VEND_CODE_AP"
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APT1099V", 1))
            .Tables("APT1099V").Columns("PRINT_IND").ReadOnly = False

            Dim CUTOFF As Decimal = Val(Absx1.numFor("CUTOFF").Value & "")
            For Each rowAPT1900V As DataRow In dst.Tables("APT1099V").Select("PMT_1099 < " & CStr(CUTOFF))
                rowAPT1900V.Item("PRINT_IND") = "0"
            Next

            ASCMAIN1.sql = "Select * from APTPARM1 Where AP_PARM_KEY = 'Z'"
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTPARM1", 1))
        End With

        Check_if_Empty("APTCHCK1")

    End Sub

    Public Overrides Sub Print_Report()
        Report_Subt = "1099 Details for Payments Made from " & Format$(Absx1.dteFor("CHK_DATE_F").Value, "dd-MMM-yyyy") _
        & " to " & Format$(Absx1.dteFor("CHK_DATE_L").Value, "dd-MMM-yyyy") _
        & IIf(Absx1.numFor("CUTOFF").Value > 0, " over " & Format$(Absx1.numFor("CUTOFF").Value, "$##,###.00"), "")
        Generate_Report(RPT, , Report_Subt)

        '' 1099 Form
        Report_Subt = ""
        RPT = "APR1099F"
        Generate_Report(RPT, "1099 Form", Report_Subt)

        '' Payment Review
        Report_Subt = "Summary of Payments Made from " & Format$(Absx1.dteFor("CHK_DATE_F").Value, "dd-MMM-yyyy") _
        & " to " & Format$(Absx1.dteFor("CHK_DATE_L").Value, "dd-MMM-yyyy") _
        & IIf(Absx1.numFor("CUTOFF").Value > 0, " over " & Format$(Absx1.numFor("CUTOFF").Value, "$##,###.00"), "")
        RPT = "APR1099G"
        Generate_Report(RPT, "Payment Review", Report_Subt)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            'If Absx1.cmbFor("RYP").Text = "" Then
            '    EMsg &= "You Must Select a Period"
            'End If
        End If

    End Sub
End Class