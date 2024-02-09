Public Class GLRADTL1
    Dim sqlr As String = ""
    Dim sqlg As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Call Get_PARM("GLTPARM1")
        'Set_cmbYP("RYP0", ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") & "", -60, 12, 0)
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 60, -11)
        Set_cmbYP_Child("RYP1", 12, "RYP0")
        Absx1.chkFor("CHKSEG2").Text = ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & ""
        Absx1.chkFor("CHKSEG3").Text = ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & ""
        Absx1.chkFor("CHKSEG4").Text = ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & ""

    End Sub

    Protected Overrides Sub Build_Workfile()

        ' Prepare Working Variables

        sqlr = ""
        sqlg = ""

        Dim sqlw As String = " where GLTDETL1.OPS_YYYYPP >= '" & RYP0 & "' and GLTDETL1.OPS_YYYYPP <= '" & RYP1 & "'"
        sqlw &= MyBase.Get_Filter("ACCT_CODE", "GLTDETL1.ACCT_CODE")
        sqlw &= MyBase.Get_Filter("SEG2_CODE", "GLTDETL1.SEG2_CODE")
        sqlw &= MyBase.Get_Filter("SEG3_CODE", "GLTDETL1.SEG3_CODE")
        sqlw &= MyBase.Get_Filter("SEG4_CODE", "GLTDETL1.SEG4_CODE")
        sqlw &= MyBase.Get_Filter("JOURNAL_TYPE", "GLTJRNL1.JOURNAL_TYPE")
        sqlw &= MyBase.Get_Filter("ACCT_TYPE", "GLTACCT1.ACCT_TYPE")
        sqlw &= MyBase.Get_Filter("ACCT_CLASS_CODE", "GLTACCT1.ACCT_CLASS_CODE")
        sql = "Select GLTDETL1.*, GLTJRNL1.JOURNAL_TYPE, GLTACCT1.ACCT_TYPE, GLTACCT1.ACCT_CLASS_CODE" _
            & " from GLTDETL1,GLTJRNL1,GLTACCT1" & sqlw _
            & " and GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO" _
            & " and GLTACCT1.ACCT_CODE = GLTDETL1.ACCT_CODE"
        Dim TT_GLTDETL1 As String = ASCMAIN1.Temp_Table(sql)

        Dim TT As String = GL_Prep(Mid(RYP0, 1, 4), Mid(RYP1, 1, 4))

        For Each COLUMN_NAME As String In New String() {"SEG2_CODE", "SEG3_CODE", "SEG4_CODE"}
            If Absx1.chkFor("CHK" & Mid(COLUMN_NAME, 1, 4)).Checked Then
                Update_Tables(TT_GLTDETL1, COLUMN_NAME)
                sqlg = sqlg & ", '*'"
            Else
                sqlr = sqlr & ", " & COLUMN_NAME
                sqlg = sqlg & ", " & COLUMN_NAME
            End If
        Next
       
        sql = "Select TT.ACCT_CODE " & sqlr & vbCrLf _
            & ", Sum (Decode(TT.ACCT_YEAR,'" & Mid$(RYP0, 1, 4) & "',NVL(TT.ACCT_BEG_BAL,0),0)"
        If Val(Mid$(RYP0, 5, 2)) > 1 Then
            For i As Integer = 1 To Val(Mid$(RYP0, 5, 2)) - 1
                sql &= " + Decode(TT.ACCT_YEAR,'" & Mid$(RYP0, 1, 4) & "',NVL(TT.ACCT_ACT_P" & Format$(i, "00") & ",0),0)"
            Next i
        End If
        sql &= ") BEG_BAL from " & TT & " TT, GLTACCT1" & vbCrLf _
            & " where TT.ACCT_YEAR in ('" & Mid$(RYP0, 1, 4) & "','" & Mid$(RYP1, 1, 4) & "')" & vbCrLf _
            & " and GLTACCT1.ACCT_CODE = TT.ACCT_CODE"
        sql &= MyBase.Get_Filter("ACCT_CODE", "TT.ACCT_CODE")
        sql &= MyBase.Get_Filter("SEG2_CODE", "TT.SEG2_CODE")
        sql &= MyBase.Get_Filter("SEG3_CODE", "TT.SEG3_CODE")
        sql &= MyBase.Get_Filter("SEG4_CODE", "TT.SEG4_CODE")
        sql &= MyBase.Get_Filter("ACCT_TYPE", "TT.ACCT_TYPE")
        sql &= MyBase.Get_Filter("ACCT_CLASS_CODE", "GLTACCT1.ACCT_CLASS_CODE")
        sql = sql & " group by TT.ACCT_CODE" & sqlg
        Dim TT_GLTACCTB As String = ASCMAIN1.Temp_Table(sql)

        sql = "Select * from " & TT_GLTACCTB
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTACCTB", 4))

        If Absx1.chkFor("CHKDTL").Checked Then
        Else
            sql = "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE from " & TT_GLTACCTB & " where BEG_BAL <> 0" _
                & " minus " _
                & " Select Distinct ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE from " & TT_GLTDETL1
            sql = "Insert into " & TT_GLTDETL1 & "(OPS_YYYYPP, JOURNAL_NO, JOURNAL_LNO, ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE) " _
                & " Select '" & RYP0 & "' OPS_YYYYPP, '.' JOURNAL_NO, ROWNUM JOURNAL_LNO, ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE from (" & sql & ")"
            ASCDATA1.ExecuteSQL(sql)
        End If

        sql = "Select T.*, GLTACCT1.ACCT_DESC, GLTJRNL1.JOURNAL_DESC, GLTJRNL1.INIT_OPER, GLTJRNL1.INIT_DATE from " & TT_GLTDETL1 & " T, GLTACCT1, GLTJRNL1 where GLTACCT1.ACCT_CODE = T.ACCT_CODE and GLTJRNL1.JOURNAL_NO = T.JOURNAL_NO"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTDETL1", 3))

        ASCMAIN1.sql = "" _
            & "SELECT 'V' DETL_CVX_TYPE, APTVEND1.VEND_CODE DETL_CVX_NO " _
            & ", APTVEND1.VEND_NAME DETL_CVX_NAME" _
            & " from APTVEND1 where VEND_CODE in " _
            & " (Select Distinct DETL_CVX_NO from " & TT_GLTDETL1 _
            & " where DETL_CVX_TYPE = 'V')" _
            & " UNION " _
            & "SELECT 'C' DETL_CVX_TYPE, ARTCUST1.CUST_CODE DETL_CVX_NO " _
            & ", ARTCUST1.CUST_NAME DETL_CVX_NAME" _
            & " from ARTCUST1 where CUST_CODE in " _
            & " (Select Distinct DETL_CVX_NO from " & TT_GLTDETL1 _
            & " where DETL_CVX_TYPE = 'C')"
        Dim x As String = ASCMAIN1.sql

        dst.Tables.Add(ASCDATA1.GetDataTable("", "GLTDETLX", 2))

        sql = "Select * from GLTACCT1 where ACCT_CODE in (Select Distinct ACCT_CODE from " & TT_GLTDETL1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTACCT1", 1))

        sql = "Select * from GLTPARM2 where OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTPARM2", 1))

        sql = "Select * from GLTJRNL1 where JOURNAL_NO in (Select Distinct JOURNAL_NO from " & TT_GLTDETL1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTJRNL1", 1))


    End Sub

    Public Overrides Sub Print_Report()

        Dim SUBT As String = "Activity from " & RYPLEGEND0 & " thru " & RYPLEGEND1
        If Absx1.chkFor("CHKDTL").Checked Then
            SUBT = SUBT & " - (Accounts with Activity Only)"
        End If
        CR_params.Add("SEG2_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "")
        CR_params.Add("SEG3_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "")
        CR_params.Add("SEG4_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "")

        CR_params.Add("RYPLEGEND0", RYPLEGEND0)
        CR_params.Add("RYPLEGEND1", RYPLEGEND1)

        CR_params.Add("SEL_JRNLS", IIf(SQLA("JOURNAL_TYPE", "CODE_VALUES") = "", "0", "1"))
        CR_params.Add("SHOW_DETL_DESC", IIf(Absx1.chkFor("SHOW_DETL_DESC").Checked, "1", "0"))
        CR_params.Add("SHOW_CVX_NAME", IIf(Absx1.chkFor("SHOW_CVX_NAME").Checked, "1", "0"))

        Generate_Report(RPT, , SUBT)

        If ASCMAIN1.CLIENT = "RGI" Then
            Prepare_Data_Extracts()
        End If

    End Sub

    Sub Update_Tables(ByVal TT_GLTDETL1 As String, ByVal COLUMN_NAME As String)
        'For Each rowGLTDETL1 As DataRow In dst.Tables("GLTDETL1").Select()
        '    rowGLTDETL1.Item(COLUMN_NAME) = "*"
        'Next
        ASCDATA1.ExecuteSQL("Update " & TT_GLTDETL1 & " Set " & COLUMN_NAME & " = '*'")
        sqlr = sqlr & ", '*' " & COLUMN_NAME
    End Sub
    Sub Prepare_Data_Extracts()


        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        grdASTEXPT1.DataSource = dst.Tables("GLTDETL1")

        grdASTEXPT1.Text = "Account Detail Report"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        '  UltraTabControl1.Tabs("Data Grids").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")

        Set_DX_Column(grdASTEXPT1, "OPS_YYYYPP", "OPS_YYYYPP", 90,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "JOURNAL_NO", "Journal No", 90,,, System.Drawing.Color.Gold)
        '   Set_DX_Column(grdASTEXPT1, "JOURNAL_LNO", "Journal Lno", 60, "######0", , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "ACCT_CODE", "Accont", 60,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "SEG2_CODE", "Seg2", 60,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "SEG3_CODE", "Seg3", 60,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "SEG4_CODE", "Seg4", 60,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "DETL_CTL_DATE", "DETL_CTL_DATE", 90, "MM/dd/yy",, System.Drawing.Color.Gold)

        Set_DX_Column(grdASTEXPT1, "DETL_CTL_NO", "Detail Ctl No", 90,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "DETL_CTL_LNO", "Ctl LNo", 60, "######0", , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "DETL_EXE_NO", "Detail Exe No", 90,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "DETL_POSTING_AMT", "Detail Posting Amt", 100, "#,##0.00", "Sum", System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT1, "DETL_DESC", "Description", 100,,, System.Drawing.Color.Gold)

        Set_DX_Column(grdASTEXPT1, "DETL_EXP_CTL_NO", "DETL_EXP_CTL_NO", 100,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "DETL_CVX_NO", "DETL_CVX_NO", 100,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "DETL_CVX_REF_DATE", "DETL_CVX_REF_DATE", 90, "MM/dd/yy",, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "DETL_CVX_REF_NO", "DETL_CVX_REF_NO", 100,,, System.Drawing.Color.Gold)

        Set_DX_Column(grdASTEXPT1, "DETL_CVX_REF_LNO", " DETL_CVX_REF_LNO", 60, "######0", , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "DETL_CTL_TYPE", "DETL_CTL_TYPE", 60,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "DETL_CVX_TYPE", "DETL_CVX_TYPE", 60,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "JOURNAL_TYPE", "JOURNAL_TYPE", 60,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "ACCT_TYPE", "ACCT_TYPE", 60,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "ACCT_CLASS_CODE", "ACCT_CLASS_CODE", 60,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "ACCT_DESC", "Acct Desc", 100,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "JOURNAL_DESC", "Journal Desc", 100,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "INIT_OPER", "Init Oper", 100,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "INIT_DATE", "Init Date", 90, "MM/dd/yy",, System.Drawing.Color.Gold)

        Sort_grdColumns(grdASTEXPT1, "ACCT_CODE,SEG2_CODE,OPS_YYYYPP,JOURNAL_NO")


    End Sub
End Class