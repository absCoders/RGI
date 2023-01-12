Imports System.Text
Imports System.Math
Imports System.Drawing


Public Class PORACCR1
    Dim POTACCR1 As String
    Dim RYP_Legend As String = ""
    Dim JOURNAL_LNO As Integer = 0
    Dim NYP As String = ""
    Dim grdASTEXPT2 As New UltraWinGrid.UltraGrid

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("GLTPARM1")
        Get_PARM("POTPARM1")

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        RYP_Legend = Absx1.cmbFor("RYP0").Value
        RYP = Mid(RYP_Legend, 1, 4) & Mid(RYP_Legend, 6, 2)

        NYP = ASCMAIN1.Period_Calc(RYP, 1)

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)
        Dim PRD_END_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")

        If RYP = ASCMAIN1.CYP Then

            If ASCMAIN1.EOM <> "1" Then
                If MsgBox("Month End Initialization is not in Progress, G/L Update disabled" & vbCrLf & vbCrLf & "Proceed with Report Only", MsgBoxStyle.OkOnly, "Verification") Then
                    'Else
                    '    Exit Sub
                End If

                RWU = "N"
            End If

            ' if it is 10/03 and anna has not closed sep
            ' and sm enters a shipment - that shipment probably does not belong in sept
            ' so get the prd end date for sep and add a filter to the sql below where PO_DATE_SHIPPED <= LAST DATE OF CYP
            Dim LAST_DATE_CYP As String = Format(PRD_END_DATE, "dd-MMM-yyyy")


            ASCMAIN1.sql = $"
            SELECT '{ASCMAIN1.CYP}' OPS_YYYYPP, X.* FROM (
            SELECT 'S' STATUS, POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO, MIN (POTORDR1.VEND_CODE) VEND_CODE
            , POTSHIP1.PO_DATE_SHIPPED, POTSHIP2.PO_DATE_RECEIVED, POTSHIP2.TRAN_NO RECEIPT_NO
            , SUM (POTSHIP3.PO_QTY_SHP) QTY
            , SUM (POTSHIP3.PO_QTY_SHP * (POTSHIP3.PO_COST_VCOST + POTSHIP3.PO_COST_MATLS + POTSHIP3.PO_COST_OTHER)) AMT_FIRST
            , SUM (POTSHIP3.PO_QTY_SHP * POTSHIP3.PO_COST_LANDED) AMT_LAND
            , POTSHIP2.BOL_NO, POTSHIP2.COMM_INV_NO, POTSHIP2.CONTAINER_NO
            FROM POTSHIP2,POTSHIP3,POTORDR1,POTSHIP1
            WHERE POTSHIP2.ACCRUAL_STATUS = '0' AND POTSHIP2.TRAN_NO IS NULL
               AND POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO
               AND POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO
               AND POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO
               AND POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO
               AND POTSHIP1.PO_DATE_SHIPPED <= '" & LAST_DATE_CYP & "'
            GROUP BY POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO, POTSHIP2.BOL_NO, POTSHIP2.COMM_INV_NO, POTSHIP2.CONTAINER_NO
            , POTSHIP1.PO_DATE_SHIPPED, POTSHIP2.PO_DATE_RECEIVED, POTSHIP2.TRAN_NO
            UNION
            SELECT 'R' STATUS, POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO, MIN (POTORDR1.VEND_CODE) VEND_CODE
            , POTSHIP1.PO_DATE_SHIPPED, POTSHIP2.PO_DATE_RECEIVED, POTSHIP2.TRAN_NO RECEIPT_NO
            , SUM (ICTIREC2.QTY_REC) QTY
            , SUM (ICTIREC2.QTY_REC * (POTSHIP3.PO_COST_VCOST + POTSHIP3.PO_COST_MATLS + POTSHIP3.PO_COST_OTHER)) AMT_FIRST
            , SUM (ICTIREC2.QTY_REC * POTSHIP3.PO_COST_LANDED) AMT_LAND
            , POTSHIP2.BOL_NO, POTSHIP2.COMM_INV_NO, POTSHIP2.CONTAINER_NO
            FROM POTSHIP2,POTSHIP3,POTORDR1,POTSHIP1,ICTIREC1,ICTIREC2
            WHERE  ICTIREC1.ACCRUAL_STATUS = '0'
               AND POTSHIP2.PO_SHIPMENT_NO = ICTIREC1.PO_SHIPMENT_NO
               AND POTSHIP2.PO_SHIPMENT_LNO = ICTIREC1.PO_SHIPMENT_LNO
               AND ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO
               AND ICTIREC2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO
               AND ICTIREC2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO
               AND ICTIREC2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO
               AND ICTIREC2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO
               AND POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO
               AND POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO
                AND POTSHIP1.PO_DATE_SHIPPED <= '" & LAST_DATE_CYP & "'

            GROUP BY POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO, POTSHIP2.BOL_NO, POTSHIP2.COMM_INV_NO, POTSHIP2.CONTAINER_NO
            , POTSHIP1.PO_DATE_SHIPPED, POTSHIP2.PO_DATE_RECEIVED, POTSHIP2.TRAN_NO 
            ) X"

        Else
            RWU = "N"
            ASCMAIN1.sql = $"Select * from POTACCR1 where OPS_YYYYPP = '{RYP}'"
        End If

        POTACCR1 = ASCMAIN1.Temp_Table


        Dim sqlw As String = ""
        sqlw &= SQL_in("VEND_CODE", "POTACCR1.VEND_CODE")
        If sqlw <> "" Then RWU = "N"

        With dst
            ASCMAIN1.sql = $"Select POTACCR1.* from {POTACCR1} POTACCR1 " & ASCMAIN1.SQL_Add_WHERE(sqlw)
            Create_TDA(.Tables.Add, "POTACCR1", "**", , False)

            Create_TDA(.Tables.Add, "GLTINTF1", "*")
            ASCMAIN1.sql = $"Select APTVEND1.VEND_CODE, APTVEND1.VEND_NAME from APTVEND1 where APTVEND1.VEND_CODE in (Select Distinct VEND_CODE from {POTACCR1} POTACCR1 )"
            Create_TDA(.Tables.Add, "APTVEND1", "**", 0, False, "", 1)

            ASCMAIN1.sql = $"SELECT PO_SHIPMENT_NO,STATUS,SUM(QTY) QTY,SUM(AMT_FIRST) AMT_FIRST,SUM(AMT_LAND) AMT_LAND FROM (Select POTACCR1.* from {POTACCR1} POTACCR1 " & ASCMAIN1.SQL_Add_WHERE(sqlw) & ") GROUP BY PO_SHIPMENT_NO,STATUS "
            Create_TDA(.Tables.Add, "POTACCRS", "**", , False)


        End With

        Fill_Records("POTACCR1")
        Fill_Records("APTVEND1")
        Fill_Records("POTACCRS")

        GL_Interface()

        Check_if_Empty("POTACCR1")
    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT)
        Print_GL()
        If ASCMAIN1.CLIENT = "VAN" Then
            Prepare_Data_Extracts()
            Prepare_Data_Extracts1()
        End If

    End Sub

    Overrides Sub Update_Record()

        Dim sql As String

        sql = $"Insert into POTACCR1 Select * from {POTACCR1}"
        ASCDATA1.ExecuteSQL(sql)

        GL_Update()

    End Sub

    Sub GL_Interface()

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_TYPE As String = "POAC"
        JOURNAL_LNO = 0

        Dim PO_ACCR_AMT_FIRST_S As Decimal = Val(dst.Tables("POTACCR1").Compute("SUM(AMT_FIRST)", "STATUS = 'S'") & "")
        Dim PO_ACCR_AMT_LAND_S As Decimal = Val(dst.Tables("POTACCR1").Compute("SUM(AMT_LAND)", "STATUS = 'S'") & "")

        Dim PO_ACCR_AMT_FIRST_R As Decimal = Val(dst.Tables("POTACCR1").Compute("SUM(AMT_FIRST)", "STATUS = 'R'") & "")
        Dim PO_ACCR_AMT_LAND_R As Decimal = Val(dst.Tables("POTACCR1").Compute("SUM(AMT_LAND)", "STATUS = 'R'") & "")

        Write_GLTINTF1(JOURNAL_NO, JOURNAL_TYPE, ROWs("POTPARM1").Item("PO_PARM_ACCT_ACCR_SHP"), PO_ACCR_AMT_LAND_S, "")
        Write_GLTINTF1(JOURNAL_NO, JOURNAL_TYPE, ROWs("POTPARM1").Item("PO_PARM_ACCT_ACCR_REC"), PO_ACCR_AMT_LAND_R, "")

        Write_GLTINTF1(JOURNAL_NO, JOURNAL_TYPE, ROWs("POTPARM1").Item("PO_PARM_ACCT_ACCR_LIA"), -1 * (PO_ACCR_AMT_LAND_R + PO_ACCR_AMT_LAND_S), "")

    End Sub

    Sub Write_GLTINTF1(JOURNAL_NO As String, JOURNAL_TYPE As String, ACCT_CODE As String, DETL_POSTING_AMT As Decimal, Optional VEND_CODE As String = "")

        Dim rowGLTINTF1 As DataRow
        JOURNAL_LNO += 1

        For I As Integer = 0 To 1
            rowGLTINTF1 = dst.Tables("GLTINTF1").NewRow
            With rowGLTINTF1
                If I = 0 Then
                    .Item("OPS_YYYYPP") = RYP
                Else
                    .Item("OPS_YYYYPP") = NYP
                End If

                .Item("JOURNAL_NO") = JOURNAL_NO
                .Item("JOURNAL_LNO") = JOURNAL_LNO
                .Item("ACCT_CODE") = ACCT_CODE
                .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                .Item("DETL_CTL_DATE") = Format(DATETIME_STAMP, "MM/dd/yyyy")
                .Item("DETL_POSTING_AMT") = DETL_POSTING_AMT
                .Item("DETL_EXE_NO") = XNO
                .Item("DETL_CTL_NO") = DBNull.Value
                .Item("DETL_CTL_LNO") = DBNull.Value
                .Item("DETL_CVX_NO") = VEND_CODE
                .Item("DETL_CVX_REF_DATE") = DBNull.Value
                .Item("DETL_CVX_REF_NO") = DBNull.Value
                .Item("DETL_DESC") = DBNull.Value
                .Item("DETL_CVX_TYPE") = "V"
                .Item("JOURNAL_TYPE") = JOURNAL_TYPE
            End With
            dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)

            DETL_POSTING_AMT = -1 * DETL_POSTING_AMT
        Next

    End Sub
    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        '  grdASTEXPT1.DataSource = dst.Tables("ASTSRPT1")
        grdASTEXPT1.DataSource = dst.Tables("POTACCR1")

        grdASTEXPT1.Text = MENU_ITEM_DESC

        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        Dim Cs As New List(Of String)
        Dim G As Integer = 0
        For Each COLUMN_NAME As String In COLUMN_NAMEs
            Cs.Add(COLUMN_NAME)
            G += 1
            Set_DX_Column(grdASTEXPT1, "G" & CStr(G), COLUMN_CAPTIONs(G - 1), 100, , , Color.Gold)
            grdASTEXPT1.DisplayLayout.Bands(0).Columns("G" & CStr(G)).Header.Fixed = True
        Next
        Set_DX_Column(grdASTEXPT1, "VEND_CODE", "Vendor", 100, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "COMM_INV_NO", "Invoice No", 120, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "BOL_NO", "Bill of Lading", 120, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "CONTAINER_NO", "Container", 120, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "PO_SHIPMENT_NO", "Shipment No", 90, , , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "PO_SHIPMENT_LNO", "Ship Lno", 60, , , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "PO_DATE_SHIPPED", "Shipped Dt", 95, "MM/dd/yy", , Color.Pink)
        Set_DX_Column(grdASTEXPT1, "PO_DATE_RECEIVED", "Received Dt", 95, "MM/dd/yy", , Color.Pink)
        Set_DX_Column(grdASTEXPT1, "RECEIPT_NO", "Receipt No", 90)
        Set_DX_Column(grdASTEXPT1, "STATUS", "Status S/R", 50)
        Set_DX_Column(grdASTEXPT1, "OPS_YYYYPP", "Period", 60)
        Set_DX_Column(grdASTEXPT1, "QTY", "Units", 90, "#,###,##0", , Color.Pink)
        Set_DX_Column(grdASTEXPT1, "AMT_FIRST", "Accrued PO First", 120, "##,###,##0.00", , Color.Pink)
        Set_DX_Column(grdASTEXPT1, "AMT_LAND", "Accrued PO Land", 120, "##,###,##0.00", , Color.Pink)


        Create_Summary(grdASTEXPT1, "VEND_CODE", "Count")
        Create_Summary(grdASTEXPT1, New String() {"QTY", "AMT_FIRST", "AMT_LAND"})



        '  grdASTEXPT1.DisplayLayout.Bands(0).Columns("STYLE_CODE").Header.Fixed = True

        Sort_grdColumns(grdASTEXPT1, "VEND_CODE,PO_DATE_SHIPPED")






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
            grdASTEXPT2.Text = "Shipments Summary"

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

        grdASTEXPT2.DataSource = dst.Tables("POTACCRS")
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
        Set_DX_Column(grdASTEXPT2, "PO_SHIPMENT_NO", "Shipment No", 90, , , Color.Orange)
        Set_DX_Column(grdASTEXPT2, "STATUS", "Status S/R", 50)
        Set_DX_Column(grdASTEXPT2, "QTY", "Units", 90, "#,###,##0", , Color.Pink)
        Set_DX_Column(grdASTEXPT2, "AMT_FIRST", "Accrued PO First", 120, "##,###,##0.00", , Color.Pink)
        Set_DX_Column(grdASTEXPT2, "AMT_LAND", "Accrued PO Land", 120, "##,###,##0.00", , Color.Pink)

        Create_Summary(grdASTEXPT2, "PO_SHIPMENT_NO", "Count")
        Create_Summary(grdASTEXPT2, New String() {"QTY", "AMT_FIRST", "AMT_LAND"})

        Sort_grdColumns(grdASTEXPT2, "PO_SHIPMENT_NO")



    End Sub


End Class