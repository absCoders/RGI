Public Class GLRJRNL1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Call Get_PARM("GLTPARM1")
        Set_cmbYP("RYP0", ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") & "", -60, 12, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

    End Sub

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGLTDETLA, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "PO Inquiry"
            '    Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
            '    Context_Launch("Load", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "PO")

            'Case "Vendor Invoice Inquiry"
            '    Dim VOUCHER_NO As String = grd.ActiveRow.Cells("VOUCHER_NO").Text
            '    If VOUCHER_NO <> "" Then
            '        Context_Launch("View", VOUCHER_NO, e.Tool.Key, "APTINVHI")
            '    End If
        End Select
    End Sub

#End Region
    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""
        sqlw &= SQLA_filter("JOURNAL_NO", "GLTJRNL1")
        sqlw &= SQLA_filter("JOURNAL_TYPE", "GLTJRNL1")

        If SQLA("JOURNAL_NO") <> "" Then
            sqlw = sqlw & SQL_in("JOURNAL_NO", "GLTJRNL1.JOURNAL_NO")
        Else
            sqlw = sqlw & " and GLTJRNL1.OPS_YYYYPP >= '" & RYP0 & "'"
            sqlw = sqlw & " and GLTJRNL1.OPS_YYYYPP <= '" & RYP1 & "'"
        End If

        ASCMAIN1.sql = "Select * from GLTJRNL1 " & ASCMAIN1.SQL_Add_WHERE(sqlw)
        dst.Tables.Add(ASCDATA1.GetDataTable("", "GLTJRNL1", 1))

        'dst.Tables("GLTJRNL1").Rows.Clear()

        ASCMAIN1.sql = "Select GLTDETL1.* from GLTDETL1,GLTJRNL1 " _
        & " where GLTDETL1.JOURNAL_NO = GLTJRNL1.JOURNAL_NO " & sqlw
        dst.Tables.Add(ASCDATA1.GetDataTable("", "GLTDETL1", 3))

        ASCMAIN1.sql = "SELECT Distinct GLTDETL1.DETL_CVX_TYPE, GLTDETL1.DETL_CVX_NO " _
        & ", APTVEND1.VEND_NAME DETL_CVX_NAME" _
        & " from GLTDETL1,GLTJRNL1,APTVEND1" _
        & " where GLTDETL1.JOURNAL_NO = GLTJRNL1.JOURNAL_NO " & sqlw _
        & "   and GLTDETL1.DETL_CVX_TYPE = 'V' " _
        & "   and APTVEND1.VEND_CODE = GLTDETL1.DETL_CVX_NO" _
        & " UNION " _
        & "SELECT Distinct GLTDETL1.DETL_CVX_TYPE, GLTDETL1.DETL_CVX_NO " _
        & ", ARTCUST1.CUST_NAME DETL_CVX_NAME" _
        & " from GLTDETL1,GLTJRNL1,ARTCUST1" _
        & " where GLTDETL1.JOURNAL_NO = GLTJRNL1.JOURNAL_NO " & sqlw _
        & "   and GLTDETL1.DETL_CVX_TYPE = 'C' " _
        & "   and ARTCUST1.CUST_CODE = GLTDETL1.DETL_CVX_NO"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "GLTDETLX", 2))

        ASCMAIN1.sql = "Select GLTDETL1.OPS_YYYYPP,GLTDETL1.JOURNAL_NO,GLTJRNL1.JOURNAL_DESC,GLTDETL1.JOURNAL_LNO" & vbCrLf _
        & ",GLTDETL1.ACCT_CODE,GLTACCT1.ACCT_DESC,GLTDETL1.SEG2_CODE,GLTDETL1.SEG3_CODE,GLTDETL1.SEG4_CODE" & vbCrLf _
        & ",GLTDETL1.DETL_CTL_DATE,GLTDETL1.DETL_POSTING_AMT,GLTDETL1.DETL_DESC,GLTJRNL1.INIT_OPER,GLTJRNL1.INIT_DATE,GLTJRNL1.JOURNAL_TYPE" & vbCrLf _
        & " from GLTDETL1,GLTJRNL1,GLTACCT1" & vbCrLf _
        & " where GLTACCT1.ACCT_CODE = GLTDETL1.ACCT_CODE" & vbCrLf _
        & "   and GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO" & vbCrLf _
        & "   and GLTDETL1.OPS_YYYYPP >= '" & RYP0 & "'" & vbCrLf _
        & "   and GLTDETL1.OPS_YYYYPP <= '" & RYP1 & "'" & vbCrLf _
        & sqlw
        dst.Tables.Add(ASCDATA1.GetDataTable("", "GLTDETLA", 0))
        grdGLTDETLA.DataSource = dst.Tables("GLTDETLA")
        Sort_grdColumns(grdGLTDETLA, "OPS_YYYYPP,JOURNAL_NO,JOURNAL_LNO")
        grdGLTDETLA.Text = "GL Details posted between " & RYPLEGEND0 & " and " & RYPLEGEND1 ' & IIf(sqlw = "", "", "; Selected Journals Only")
        If grdGLTDETLA.DisplayLayout.Bands(0).Summaries.Count = 0 Then
            Create_Summary(grdGLTDETLA, "OPS_YYYYPP", "Count")
            Create_Summary(grdGLTDETLA, "DETL_POSTING_AMT")
            Set_SEGS(grdGLTDETLA, "GLTDETLA")

        End If

        Call Get_WKCodes("GLTDETL1", "ACCT_CODE", "GLTACCT1", "*")
        Call Get_WKCodes("GLTDETL1", "OPS_YYYYPP", "GLTPARM2", "*")

        dst.Tables.Add(ASCDATA1.GetDataTable("*", "GLTSEGM1"))

        Call Prepare_GL_Account_Activity_Recaps("GLTDETL1")
    End Sub

    Public Overrides Sub Print_Report()

        CR_params.Add("SEG2_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "")
        CR_params.Add("SEG3_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "")
        CR_params.Add("SEG4_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "")
        CR_params.Add("SHOW_JRNL_COMMENTS", IIf(Absx1.chkFor("SHOW_JRNL_COMMENTS").Checked, "1", "0"))
        CR_params.Add("SHOW_DETL_DESC", IIf(Absx1.chkFor("SHOW_DETL_DESC").Checked, "1", "0"))
        CR_params.Add("SHOW_CVX_NAME", IIf(Absx1.chkFor("SHOW_CVX_NAME").Checked, "1", "0"))
        CR_params.Add("PAGE_BREAK", IIf(Absx1.chkFor("PAGE_BREAK").Checked, "1", "0"))
        CR_params.Add("ACCT_RECAPS", ROWs("GLTPARM1").Item("GL_PARM_ACCT_RECAPS") & "")
        Generate_Report(RPT)

        grdGLTDETLA.Visible = True
    End Sub

End Class