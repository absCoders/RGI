Imports System.Math
Imports System.Drawing
Public Class GLRADTL1
    Dim sqlr As String = ""
    Dim sqlg As String = ""
    Dim grdASTEXPT2 As New UltraWinGrid.UltraGrid

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Call Get_PARM("GLTPARM1")
        'Set_cmbYP("RYP0", ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") & "", -60, 12, 0)
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 60, -11)
        Set_cmbYP_Child("RYP1", 12, "RYP0")
        Absx1.chkFor("CHKSEG2").Text = ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & ""
        Absx1.chkFor("CHKSEG3").Text = ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & ""
        Absx1.chkFor("CHKSEG4").Text = ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & ""
        If ASCMAIN1.CLIENT = "RGI" Then
            Absx1.chkFor("SHOW_OPSJ").Visible = True
        Else
            Absx1.chkFor("SHOW_OPSJ").Visible = False
        End If


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

        ' If TT_GLTDETL1 = "" Then
        ASCDATA1.ExecuteSQL("Create Index I_" & TT_GLTDETL1 & "_1 on " & TT_GLTDETL1 & " (DETL_EXE_NO)")
        '  End If


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


        ' Need Journal Number in below

        If Absx1.chkFor("SHOW_OPSJ").Checked Then
            ''sql = "Select  SOTINVH2.INV_TYPE,SOTINVH2.INV_NO,SOTINVH2.INV_LNO,SOTINVH2.STYLE_CODE,SOTINVH2.COLOR_CODE,SOTINVH2.ORDR_UNIT_PRICE,ORDR_QTY_SHIP," _
            ''& "SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE INV_EXT, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.WHSE_CODE," _
            ''& " SOTINVH1.REASON_CODE, SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.TERM_CODE, SOTINVH1.SREP_CODE, INV_COMMENT, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.CUST_BILL_TO_CUST, SOTINVH1.REGISTER_XNO, SOTINVH1.REGISTER_DATE " _
            ''& " From SOTINVH2, SOTINVH1 Where SOTINVH2.INV_NO In (" _
            ''& "Select  INV_NO from sotinvh1 where register_xno In (Select Distinct DETL_EXE_NO from " & TT_GLTDETL1 & " WHERE jOURNAL_TYPE = 'OPSJ')" _
            ''& ")" _
            ''& " And SOTINVH1.INV_NO = SOTINVH2.INV_NO"
            sql = $"Select aa.JOURNAL_NO,SOTINVH2.INV_TYPE,SOTINVH2.INV_NO,SOTINVH2.INV_LNO,SOTINVH2.STYLE_CODE,SOTINVH2.COLOR_CODE,SOTINVH2.ORDR_UNIT_PRICE,
                ORDR_QTY_SHIP, SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE INV_EXT, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO,
                SOTINVH1.ORDR_CUST_PO, SOTINVH1.WHSE_CODE, SOTINVH1.REASON_CODE, SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.TERM_CODE,
                SOTINVH1.SREP_CODE, INV_COMMENT, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.CUST_BILL_TO_CUST, SOTINVH1.REGISTER_XNO,
                SOTINVH1.REGISTER_DATE, SOTINVH1.INV_FREIGHT, SOTINVH1.INV_MISC_CHG,''  MISC_FRT From SOTINVH2, SOTINVH1, (SELECT DISTINCT DETL_EXE_NO, JOURNAL_NO FROM {TT_GLTDETL1} WHERE JOURNAL_TYPE =  'OPSJ') AA
                            WHERE SOTINVH1.INV_NO = SOTINVH2.INV_NO
                And SOTINVH1.register_xno = AA.DETL_EXE_NO 
                UNION
                SELECT aa.JOURNAL_NO,SOTINVH1.INV_TYPE,SOTINVH1.INV_NO,0 INV_LNO,'' STYLE_CODE,'' COLOR_CODE, 0 ORDR_UNIT_PRICE,0 ORDR_QTY_SHIP,
                0 INV_EXT, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.WHSE_CODE, SOTINVH1.REASON_CODE,
                SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.TERM_CODE, SOTINVH1.SREP_CODE, INV_COMMENT, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.CUST_BILL_TO_CUST,
                SOTINVH1.REGISTER_XNO,SOTINVH1.REGISTER_DATE, SOTINVH1.INV_FREIGHT,SOTINVH1.INV_MISC_CHG,'Header'  MISC_FRT
                FROM SOTINVH1 ,(SELECT DISTINCT DETL_EXE_NO, JOURNAL_NO FROM  {TT_GLTDETL1}  WHERE JOURNAL_TYPE =  'OPSJ') AA
                 WHERE SOTINVH1.register_xno = AA.DETL_EXE_NO
                AND (SOTINVH1.INV_FREIGHT <> 0 OR SOTINVH1.INV_MISC_CHG <> 0) AND
                 (SOTINVH1.REASON_CODE IS NULL OR SOTINVH1.REASON_CODE <> 'SHP')"
            dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTOPSJ1", 0))
        End If






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
        '  CR_params.Add("SHOW_OPSJ", IIf(Absx1.chkFor("SHOW_OPSJ").Checked, "1", "0"))

        Generate_Report(RPT, , SUBT)

        If ASCMAIN1.CLIENT = "RGI" Then
            Prepare_Data_Extracts()
            'SHOW_OPSJ
            If Absx1.chkFor("SHOW_OPSJ").Checked Then
                Prepare_Data_Extracts1()
            End If

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
        Set_DX_Column(grdASTEXPT1, "ACCT_CODE", "Account", 60,,, System.Drawing.Color.Gold)
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

    Sub Prepare_Data_Extracts1()

        ' grdASTEXPT2 As New UltraWinGrid.UltraGrid
        grdASTEXPT2.Name = "grdASTEXPT2"
        If Not GRDs.ContainsKey("ASTEXPT2") Then
            tabDataExports.Tabs.Add()

            GRDs.Add(Mid(grdASTEXPT2.Name, 4), grdASTEXPT2)
            Add_Handlers_grd(grdASTEXPT2)

            grdASTEXPT2.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.OutlookGroupBy

            grdASTEXPT2.Parent = tabDataExports.Tabs(1).TabPage
            grdASTEXPT2.Text = "OPSJ Journal Details"

            grdASTEXPT2.Dock = System.Windows.Forms.DockStyle.Fill
            tabDataExports.Tabs(1).Text = grdASTEXPT2.Text

            grdASTEXPT2.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
            grdASTEXPT2.DisplayLayout.Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement

            tabDataExports.Tabs(1).Text = grdASTEXPT2.Text
            grdASTEXPT2.DisplayLayout.Override.AllowGroupBy = DefaultableBoolean.True
            grdASTEXPT2.DisplayLayout.GroupByBox.Hidden = False
            grdASTEXPT2.DisplayLayout.MaxColScrollRegions = 1
            grdASTEXPT2.DisplayLayout.MaxRowScrollRegions = 1


        End If

        grdASTEXPT2.DataSource = dst.Tables("GLTOPSJ1")
        ASCMAIN1.grdInitializeLayout(grdASTEXPT2)


        Set_DX_Column(grdASTEXPT2, "")
        Dim Cs As New List(Of String)
        Dim G As Integer = 0
        For Each COLUMN_NAME As String In COLUMN_NAMEs
            Cs.Add(COLUMN_NAME)
            G += 1
            Set_DX_Column(grdASTEXPT2, "G" & CStr(G), COLUMN_CAPTIONs(G - 1), 100, , , Color.Gold)
            grdASTEXPT2.DisplayLayout.Bands(0).Columns("G" & CStr(G)).Header.Fixed = True
        Next


        Set_DX_Column(grdASTEXPT2, "JOURNAL_NO", "Journal No", 90,,, System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT2, "INV_TYPE", "Inv Type", 50,,, System.Drawing.Color.Pink)
        Set_DX_Column(grdASTEXPT2, "INV_NO", "Inv No", 90,,, System.Drawing.Color.Pink)
        Set_DX_Column(grdASTEXPT2, "INV_LNO", "Inv LNo", 90,,, System.Drawing.Color.Pink)
        Set_DX_Column(grdASTEXPT2, "STYLE_CODE", "Style", 90,,, System.Drawing.Color.Pink)
        Set_DX_Column(grdASTEXPT2, "COLOR_CODE", "Color", 90,,, System.Drawing.Color.Pink)
        Set_DX_Column(grdASTEXPT2, "ORDR_UNIT_PRICE", "Price", 90, "###,##0.00", , Color.Pink)
        Set_DX_Column(grdASTEXPT2, "ORDR_QTY_SHIP", "Qty Shipped", 90, "#,###,##0", , Color.Pink)
        Set_DX_Column(grdASTEXPT2, "INV_EXT", "Inv Ext", 120, "##,###,##0.00", , Color.Pink)
        Set_DX_Column(grdASTEXPT2, "REGISTER_XNO", "Register", 120,,, System.Drawing.Color.Pink)
        Set_DX_Column(grdASTEXPT2, "REGISTER_DATE", "Register", 90, "MM/dd/yy",, System.Drawing.Color.Pink)


        Set_DX_Column(grdASTEXPT2, "CUST_CODE", "Cust Cd", 90,,, System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT2, "CUST_STORE_NO", "Cust Store", 90,,, System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT2, "ORDR_CUST_PO", "Cust PO", 90,,, System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT2, "WHSE_CODE", "Whse", 50,,, System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT2, "REASON_CODE", "Reason", 50,,, System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT2, "INV_DATE", "Inv Date", 90, "MM/dd/yy",, System.Drawing.Color.Orange)

        ' Set_DX_Column(grdASTEXPT2, "INV_CODE", "Register", 50,,, System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT2, "POST_CODE", "Post Cd", 50,,, System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT2, "TERM_CODE", "Term Cd", 50,,, System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT2, "SREP_CODE", "Srep", 50,,, System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT2, "INV_COMMENT", "Inv Comment", 120,,, System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT2, "ORDR_TYPE_CODE", "Order Type", 90,,, System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT2, "CUST_BILL_TO_CUST", "Bill to Cust", 90,,, System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT2, "INV_FREIGHT", "Freight", 120, "##,###,##0.00", , Color.Orange)
        Set_DX_Column(grdASTEXPT2, "INV_MISC_CHG", "Misc Charge", 120, "##,###,##0.00", , Color.Orange)
        Set_DX_Column(grdASTEXPT2, "MISC_FRT", "Header Charge", 50,,, System.Drawing.Color.Orange)





        Create_Summary(grdASTEXPT2, "REGISTER_XNO", "Count")
        Create_Summary(grdASTEXPT2, New String() {"ORDR_QTY_SHIP", "INV_EXT"})

        Sort_grdColumns(grdASTEXPT2, "JOURNAL_NO,REGISTER_XNO,MISC_FRT,INV_TYPE,INV_NO,INV_date")



    End Sub




End Class