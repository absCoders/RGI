Public Class GLFTBALC

    Dim sqlGLT4 As String
    Dim sqlGLT5 As String

    Dim WHSE_CODEs As New List(Of String)
    Dim SALES_DIVISION_CODEs As New List(Of String)
    Dim CUST_CODEs As New List(Of String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")

        With dst
            ASCMAIN1.sql = "Select ORDR_YYYYPP_UPDATED YYYYPP, CUST_CODE CUST, WHSE_CODE WHSE, SALES_DIVISION_CODE DIV, CURR_CODE CURR" & vbCrLf _
                & ", COUNT (*) INVS" & vbCrLf _
                & ", SUM (INV_SALES) SLS, SUM (INV_FREIGHT) FRT, SUM (INV_MISC_CHG) CHG, SUM (INV_STAX) STX, SUM (INV_TOTAL_AMOUNT) TOT" & vbCrLf _
                & ", SUM (INV_SALES_CURR) SLSCAD, SUM (INV_FREIGHT_CURR) FRTCAD, SUM (INV_MISC_CHG_CURR) CHGCAD, SUM (INV_STAX_CURR) STXCAD, SUM (INV_TOTAL_AMOUNT_CURR) TOTCAD" & vbCrLf _
                & ", MIN (INV_NO) MININV, MAX (INV_NO) MAXINV" & vbCrLf _
                & " from SOTINVH1" & vbCrLf _
                & " where CUST_CODE IN (Select CUST_CODE from ARTCUST1 where SEG4_CODE = '001')" & vbCrLf _
                & "   and ORDR_YYYYPP_UPDATED > '201811'" & vbCrLf _
                & " or WHSE_CODE IN (Select WHSE_CODE from ICTWHSE1 where SEG4_CODE = '001')" & vbCrLf _
                & " or SALES_DIVISION_CODE IN (Select SALES_DIVISION_CODE from SOTSDIV1 where SEG4_CODE = '001')" & vbCrLf _
                & " group by ORDR_YYYYPP_UPDATED, CUST_CODE, WHSE_CODE, SALES_DIVISION_CODE, CURR_CODE" & vbCrLf _
                & " order by ORDR_YYYYPP_UPDATED, CUST_CODE, WHSE_CODE, SALES_DIVISION_CODE, CURR_CODE"
            Create_TDA(.Tables.Add, "GLT1", "**", 0, False)

            ASCMAIN1.sql = "Select ICTIREC1.OPS_YYYYPP, ICTIREC1.WHSE_CODE, ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
                & ", COUNT (*) LINES, SUM (ICTIREC2.QTY_REC) QTY_REC, SUM (ICTIREC2.QTY_REC * ICTIREC2.PO_COST) AMT_REC" & vbCrLf _
                & " from ICTIREC1,ICTIREC2,ICTSTYL1" & vbCrLf _
                & " where ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = ICTIREC2.STYLE_CODE" & vbCrLf _
                & "   and ICTIREC1.OPS_YYYYPP >= '201801'" & vbCrLf _
                & "   and (ICTIREC1.WHSE_CODE IN (Select WHSE_CODE FROM ICTWHSE1 where SEG4_CODE = '001')" & vbCrLf _
                & "     or ICTSTYL1.SALES_DIVISION_CODE IN (Select SALES_DIVISION_CODE FROM SOTSDIV1 WHERE SEG4_CODE = '001'))" & vbCrLf _
                & " group by ICTIREC1.OPS_YYYYPP, ICTIREC1.WHSE_CODE, ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
                & " order by ICTIREC1.OPS_YYYYPP, ICTIREC1.WHSE_CODE, ICTSTYL1.SALES_DIVISION_CODE"
            Create_TDA(.Tables.Add, "GLT2", "**", 0, False)

            ASCMAIN1.sql = "SELECT OPS_YYYYPP, WHSE_CODE, SALES_DIVISION_CODE, SUM (QTY_BEG) QTY_BEG, SUM (QTY_END) QTY_END FROM (" & vbCrLf _
                & "Select ICTSTAT1.OPS_YYYYPP, ICTSTAT1.WHSE_CODE, ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
                & ", SUM (ICTSTAT1.WHSE_QTY_BEG) QTY_BEG, 0 QTY_END" & vbCrLf _
                & " FROM ICTSTAT1,ICTSTYL1" & vbCrLf _
                & " WHERE ICTSTYL1.STYLE_CODE = ICTSTAT1.STYLE_CODE" & vbCrLf _
                & "   AND ICTSTAT1.OPS_YYYYPP >= '201801'" & vbCrLf _
                & "   AND (ICTSTAT1.WHSE_CODE IN (SELECT WHSE_CODE FROM ICTWHSE1 WHERE SEG4_CODE = '001')" & vbCrLf _
                & "     OR ICTSTYL1.SALES_DIVISION_CODE IN (SELECT SALES_DIVISION_CODE FROM SOTSDIV1 WHERE SEG4_CODE = '001'))" & vbCrLf _
                & " group by  ICTSTAT1.OPS_YYYYPP, ICTSTAT1.WHSE_CODE, ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
                & " union " _
                & "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, ICTSTAT2.WHSE_CODE, ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
                & ", 0 QTY_BEG, SUM (ICTSTAT2.WHSE_QTY_ON_HAND) QTY_END" _
                & " FROM ICTSTAT2,ICTSTYL1" _
                & " WHERE ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE" & vbCrLf _
                & "   AND (ICTSTAT2.WHSE_CODE IN (SELECT WHSE_CODE FROM ICTWHSE1 WHERE SEG4_CODE = '001')" & vbCrLf _
                & "     OR ICTSTYL1.SALES_DIVISION_CODE IN (SELECT SALES_DIVISION_CODE FROM SOTSDIV1 WHERE SEG4_CODE = '001'))" & vbCrLf _
                & " group by ICTSTAT2.WHSE_CODE, ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
                & " union " _
                & "Select ICTSTAT5.OPS_YYYYPP, ICTSTAT5.WHSE_CODE, ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
                & ", 0 QTY_BEG, SUM (ICTSTAT5.WHSE_QTY_ON_HAND) QTY_END" & vbCrLf _
                & " FROM ICTSTAT5,ICTSTYL1" _
                & " WHERE ICTSTYL1.STYLE_CODE = ICTSTAT5.STYLE_CODE" & vbCrLf _
                & "   AND ICTSTAT5.OPS_YYYYPP >= '201801'" & vbCrLf _
                & "   AND (ICTSTAT5.WHSE_CODE IN (SELECT WHSE_CODE FROM ICTWHSE1 WHERE SEG4_CODE = '001')" & vbCrLf _
                & "     OR ICTSTYL1.SALES_DIVISION_CODE IN (SELECT SALES_DIVISION_CODE FROM SOTSDIV1 WHERE SEG4_CODE = '001'))" & vbCrLf _
                & " group by ICTSTAT5.OPS_YYYYPP, ICTSTAT5.WHSE_CODE, ICTSTYL1.SALES_DIVISION_CODE" _
                & ") group by OPS_YYYYPP, WHSE_CODE, SALES_DIVISION_CODE" & vbCrLf _
                & " having (SUM (QTY_BEG) <> 0 OR SUM (QTY_END) <> 0)" & vbCrLf _
                & " order by OPS_YYYYPP, WHSE_CODE, SALES_DIVISION_CODE"
            Create_TDA(.Tables.Add, "GLT3", "**", 0, False)


            ASCMAIN1.sql = "Select APTINVH1.OPS_YYYYPP, APTINVH1.VEND_CODE, APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.VOUCHER_NO, APTINVH1.INV_AMT_VEND, ICTIREC1.WHSE_CODE, ICTSTYL1.SALES_DIVISION_CODE, COUNT (*) RECS" & vbCrLf _
                & ", SUM (APTINVH5.INV_QTY) INV_QTY" & vbCrLf _
                & ", SUM (APTINVH5.INV_QTY * APTINVH5.INV_COST) INV_AMT" & vbCrLf _
                & ", SUM (APTINVH5.VAR_QTY) VAR_QTY" & vbCrLf _
                & ", SUM (APTINVH5.VAR_AMT) VAR_AMT" & vbCrLf _
                & " from APTINVH5,APTINVH1,POTSHIP2,ICTIREC1,ICTIREC2,ICTSTYL1" & vbCrLf _
                & " where APTINVH5.VOUCHER_NO = APTINVH1.VOUCHER_NO" & vbCrLf _
                & "   and APTINVH1.OPS_YYYYPP >= '201801'" & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_NO = APTINVH5.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_LNO = APTINVH5.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and ICTIREC1.RECEIPT_NO = APTINVH5.RECEIPT_NO" & vbCrLf _
                & "   and ICTIREC2.RECEIPT_NO = APTINVH5.RECEIPT_NO" & vbCrLf _
                & "   and ICTIREC2.RECEIPT_LNO = APTINVH5.RECEIPT_LNO" _
                & "   and ICTSTYL1.STYLE_CODE = ICTIREC2.STYLE_CODE" & vbCrLf _
                & "   and (ICTIREC1.WHSE_CODE IN (Select WHSE_CODE FROM ICTWHSE1 WHERE SEG4_CODE = '001')" & vbCrLf _
                & "     or ICTSTYL1.SALES_DIVISION_CODE IN (Select SALES_DIVISION_CODE FROM SOTSDIV1 WHERE SEG4_CODE = '001'))" & vbCrLf _
                & " group by APTINVH1.OPS_YYYYPP, APTINVH1.VEND_CODE, APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.VOUCHER_NO, APTINVH1.INV_AMT_VEND, ICTIREC1.WHSE_CODE, ICTSTYL1.SALES_DIVISION_CODE"

            sqlGLT4 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "GLT4", "**", 0, False)

            ASCMAIN1.sql = "Select * from APTINVH2 where VOUCHER_NO in (Select Distinct VOUCHER_NO from (" & sqlGLT4 & "))"
            Create_TDA(.Tables.Add, "GLT4A", "**", 0, False)

            Create_Relation("GLT4", "GLT4A", "VOUCHER_NO")


            ASCMAIN1.sql = "SELECT APTINVH1.OPS_YYYYPP, APTINVH1.VEND_CODE, APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.VOUCHER_NO, APTINVH1.INV_AMT_VEND" & vbCrLf _
                & ", POTSHIP1.WHSE_CODE, '?' SALES_DIVISION_CODE, COUNT (*) RECS" & vbCrLf _
                & ", SUM (POTLCST1.COST_ACT) COST_ACT" & vbCrLf _
                & ", SUM (POTLCST1.COST_ACC) COST_ACC" & vbCrLf _
                & ", SUM (APTINVH7.TOTAL_INV) TOTAL_INV" & vbCrLf _
                & " FROM APTINVH7,APTINVH1,POTSHIP2,POTSHIP1,POTLCST1" & vbCrLf _
                & "WHERE APTINVH7.VOUCHER_NO = APTINVH1.VOUCHER_NO" & vbCrLf _
                & "AND APTINVH1.OPS_YYYYPP >= '201801'" & vbCrLf _
                & "AND POTSHIP1.PO_SHIPMENT_NO = POTLCST1.PO_SHIPMENT_NO" & vbCrLf _
                & "AND POTSHIP2.PO_SHIPMENT_NO (+) = POTLCST1.PO_SHIPMENT_NO" & vbCrLf _
                & "AND POTSHIP2.PO_SHIPMENT_LNO (+) = POTLCST1.PO_SHIPMENT_LNO" & vbCrLf _
                & "AND POTLCST1.CTL_NO = APTINVH7.CTL_NO " & vbCrLf _
                & "   AND (POTSHIP1.WHSE_CODE IN (SELECT WHSE_CODE FROM ICTWHSE1 WHERE SEG4_CODE = '001') )" & vbCrLf _
                & "GROUP BY APTINVH1.OPS_YYYYPP, APTINVH1.VEND_CODE, APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.VOUCHER_NO, APTINVH1.INV_AMT_VEND" & vbCrLf _
                & ", POTSHIP1.WHSE_CODE"

            sqlGLT5 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "GLT5", "**", 0, False)

            ASCMAIN1.sql = "Select * from APTINVH2 where VOUCHER_NO in (Select Distinct VOUCHER_NO from (" & sqlGLT5 & "))"
            Create_TDA(.Tables.Add, "GLT5A", "**", 0, False)

            Create_Relation("GLT5", "GLT5A", "VOUCHER_NO")


            ' these next 2 sqls - need to be changed to either include or to exclude the intercompany pymt bank

            ASCMAIN1.sql = "Select APTCHCK2.*, APTCHCK1.CHECK_STATUS" & vbCrLf _
                & " from APTCHCK1,APTCHCK2" & vbCrLf _
                & " where APTCHCK1.BANK_CODE = APTCHCK2.BANK_CODE" & vbCrLf _
                & "   and APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM" & vbCrLf _
                & "   and APTCHCK1.OPS_YYYYPP >= '201801'" & vbCrLf _
                & "   and APTCHCK2.VOUCHER_NO in (Select Distinct VOUCHER_NO from (" & sqlGLT4 & ") union Select Distinct VOUCHER_NO from (" & sqlGLT5 & "))"
            Create_TDA(.Tables.Add, "GLT8", "**", 0, False)

            ASCMAIN1.sql = "Select APTCHCK2.*, APTCHCK1.CHECK_STATUS" & vbCrLf _
                & " from APTCHCK1,APTCHCK2" & vbCrLf _
                & " where APTCHCK1.BANK_CODE = APTCHCK2.BANK_CODE" & vbCrLf _
                & "   and APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM" & vbCrLf _
                & "   and APTCHCK1.OPS_YYYYPP >= '201801'" & vbCrLf _
                & "   and APTCHCK2.VOUCHER_NO in (Select Distinct VOUCHER_NO from (" & sqlGLT4 & ") union Select Distinct VOUCHER_NO from (" & sqlGLT5 & "))"
            Create_TDA(.Tables.Add, "GLT9", "**", 0, False)

        End With

        ASCMAIN1.sql = "Select WHSE_CODE from ICTWHSE1 where SEG4_CODE = '001'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            WHSE_CODEs.Add(row.Item(0))
        Next
        ASCMAIN1.sql = "Select SALES_DIVISION_CODE from SOTSDIV1 where SEG4_CODE = '001'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            SALES_DIVISION_CODEs.Add(row.Item(0))
        Next
        ASCMAIN1.sql = "Select CUST_CODE from ARTCUST1 where SEG4_CODE = '001'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            CUST_CODEs.Add(row.Item(0))
        Next

        grd1.DataSource = dst.Tables("GLT1")
        Create_Summary(grd1, New String() {"SLS", "FRT", "CHG", "STX", "TOT", "SLSCAD", "FRTCAD", "CHGCAD", "STXCAD", "TOTCAD"})

        grd2.DataSource = dst.Tables("GLT2")
        Create_Summary(grd2, New String() {"LINES", "QTY_REC", "AMT_REC"})

        grd3.DataSource = dst.Tables("GLT3")
        Create_Summary(grd3, New String() {"QTY_BEG", "QTY_END"})

        grd4.DataSource = dst.Tables("GLT4")
        Create_Summary(grd4, New String() {"INV_AMT_VEND", "INV_QTY", "INV_AMT", "VAR_QTY", "VAR_AMT"})

        grd5.DataSource = dst.Tables("GLT5")
        Create_Summary(grd5, New String() {"COST_ACT", "COST_ACC", "TOTAL_INV"})

        grd8.DataSource = dst.Tables("GLT8")
        'Create_Summary(grd8, New String() {"COST_ACT", "COST_ACC", "TOTAL_INV"})

        grd9.DataSource = dst.Tables("GLT9")
        'Create_Summary(grd9, New String() {"COST_ACT", "COST_ACC", "TOTAL_INV"})


        spl.Panel1Collapsed = True
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        UltraTabControl1.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"GLT1", "GLT2", "GLT3", "GLT4", "GLT4A", "GLT5", "GLT5A"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("OPS_YYYYPP").Text = ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP")
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Account Summary Data")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Dim ACCT_YEAR As String = Mid$(HFs("OPS_YYYYPP"), 1, 4)
        Dim P As Integer = Val(Mid$(HFs("OPS_YYYYPP"), 5, 2))

        EnforceConstraints(False)

        Fill_Records("GLT1")
        Fill_Records("GLT2")
        Fill_Records("GLT3")

        Fill_Records("GLT4")
        Fill_Records("GLT4A")

        Fill_Records("GLT5")
        Fill_Records("GLT5A")

        Fill_Records("GLT8")
        Fill_Records("GLT9")

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub
#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grd1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grd4, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Voucher Inquiry")
        Load_Popup_Menu(grd5, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Voucher Inquiry")
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
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case "grdX"


            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case ""


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Account Inquiry"
            '    If grd.ActiveRow.Band.Index = 1 Then
            '        Dim ACCT_CODE As String = grd.ActiveRow.Cells("ACCT_CODE").Value
            '        Dim rowGLTACCT1 As DataRow = LookUp("GLTACCT1", ACCT_CODE)
            '        If rowGLTACCT1 IsNot Nothing Then
            '            Context_Launch("Load", ACCT_CODE, e.Tool.Key, "GLFACTI1")
            '        End If
            '    End If


            Case "Voucher Inquiry"
                Dim VOUCHER_NO As String = grd.ActiveRow.Cells("VOUCHER_NO").Value
                Dim rowAPTINVH1 As DataRow = LookUp("APTINVH1", VOUCHER_NO)
                If rowAPTINVH1 IsNot Nothing Then
                    Context_Launch("Load", VOUCHER_NO, e.Tool.Key, "APFINVHI")
                End If

        End Select
    End Sub

#End Region

    Private Sub grd1_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grd1.InitializeLayout

    End Sub

    Private Sub grd2_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grd2.InitializeLayout

    End Sub

    Private Sub grd3_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grd3.InitializeLayout

    End Sub

    Private Sub grd4_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grd4.InitializeLayout

    End Sub

    Private Sub grd5_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grd5.InitializeLayout

    End Sub

    Private Sub grd1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grd1.InitializeRow
        If e.Row.IsDataRow Then

            If Not WHSE_CODEs.Contains(e.Row.Cells("WHSE").Value & "") Then
                e.Row.Cells("WHSE").Appearance.ForeColor = Drawing.Color.Red
            End If

            If Not SALES_DIVISION_CODEs.Contains(e.Row.Cells("DIV").Value & "") Then
                e.Row.Cells("DIV").Appearance.ForeColor = Drawing.Color.Red
            End If

            If Not CUST_CODEs.Contains(e.Row.Cells("CUST").Value & "") Then
                e.Row.Cells("CUST").Appearance.ForeColor = Drawing.Color.Red
            End If

            If e.Row.Cells("CURR").Value & "" <> "CAD" Then
                e.Row.Cells("CURR").Appearance.ForeColor = Drawing.Color.Red
            End If
        End If
    End Sub

    Private Sub grd2_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grd2.InitializeRow

    End Sub

    Private Sub grd3_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grd3.InitializeRow

    End Sub

    Private Sub grd4_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grd4.InitializeRow

    End Sub

    Private Sub grd5_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grd5.InitializeRow

    End Sub
End Class