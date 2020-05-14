Public Class APRCHKR1

    Dim APTCHKR1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("GLTPARM1")
        Call Get_PARM("APTPARM1")
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        APTCHKR1 = TAC.APCMAIN1.Prepare_Check_Register(Me, dst, True)

        'If dst Is Nothing Then
        '    Stop
        'End If
        Check_if_Empty("APTCHKR1")
    End Sub

    Public Overrides Sub Print_Report()
        If ASCMAIN1.CLIENT = "VAN" Then
            RPT = "APRCHKR5"
        End If
        Generate_Report(RPT)
        Print_GL()

        If ASCMAIN1.CLIENT = "VAN" Then
            RPT = "APRCHKR3"
            'RPT_TITLE = "Check Register"
            'SUBT = "Summary"
            'CR_params.Add("OPTDTL", Absx1.optFor("OPTDTL").Value)
            Generate_Report(RPT, RPT_TITLE, SUBT)

            RPT = "APRCHKR4"
            RPT_TITLE = "Check Register"
            SUBT = "LC Payments"
            'CR_params.Add("OPTDTL", Absx1.optFor("OPTDTL").Value)
            Generate_Report(RPT, RPT_TITLE, SUBT)

        End If


        If ASCMAIN1.CLIENT = "VAN" Then
            Prepare_Data_Extracts()
        End If

    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        grdASTEXPT1.DataSource = dst.Tables("APTCHKR1")

        grdASTEXPT1.Text = "Check Register"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")

        Set_DX_Column(grdASTEXPT1, "BANK_CODE", "Bank", 60,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "CHECK_STATUS", "Status", 60,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "CHECK_NUM", "Check No", 60,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "CHECK_DATE", "Date", 80, "MM/dd/yy",, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "CHECK_AMT", "Check Amt", 100, "#,##0.00", "Sum", System.Drawing.Color.Orange)

        Set_DX_Column(grdASTEXPT1, "PYMT_METHOD", "Method", 60,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "VEND_CODE", "Vendor", 60,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "VEND_NAME", "Vendor Name", 100,,, System.Drawing.Color.Gold)

        Sort_grdColumns(grdASTEXPT1, "BANK_CODE,CHECK_STATUS,CHECK_NUM")

    End Sub


    Overrides Sub Update_Record()

        Dim sql As String

        sql = "Update APTCHCK1 " _
        & " Set REGISTER_IND = '1'" _
        & ", REGISTER_XNO = '" & XNO & "'" _
        & " where (BANK_CODE, CHECK_NUM) in " _
        & "(Select BANK_CODE, CHECK_NUM from " & APTCHKR1 _
        & " where RECORD_TYPE = 'I')"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Update APTCHCK1 " _
        & " Set REGISTER_IND_F = '1'" _
        & ", REGISTER_XNO_F = '" & XNO & "'" _
        & " where (BANK_CODE, CHECK_NUM) in " _
        & "(Select BANK_CODE, CHECK_NUM from " & APTCHKR1 _
        & " where RECORD_TYPE = 'V')"
        ASCDATA1.ExecuteSQL(sql)

        GL_Update()

        If ASCMAIN1.DBS_SERVER = "EXP" Or ASCMAIN1.DBS_COMPANY = "EXP" Then
            For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("GLTINTF1"), New String() {"JOURNAL_NO"}).Select("")
                Dim JOURNAL_NO As String = row.Item(0)
                ASCMAIN1.sql = "Insert into GLTDETL1_OBX Select GLTDETL1.*, NULL DATETIME_STAMP, 'APCD' JOURNAL_TYPE from GLTDETL1 where JOURNAL_NO = '" & JOURNAL_NO & "'"
                ASCDATA1.ExecuteSQL()
            Next
        End If

    End Sub
End Class