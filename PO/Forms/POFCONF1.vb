Public Class POFCONF1

    Dim POTCONF1 As String 'TABLE_NAME
    Dim sqlPOTCONF1 As String

    Dim POTPPRM1 As String
    Dim sqlPOTPPRM1 As String
    Dim rowPOTPPRM1 As DataRow
    Dim POTPPRM1_CODE As String



#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' HORIZON_DATE = ""
        Dim HORIZON_DATE As Date = Now.AddDays(-730)
        ' HORIZON_DATE = CDate(HORIZON_DATE.AddDays(-730))
        Dim horizon As String = Format(HORIZON_DATE, "dd-MMM-yyyy")

        'Dim z As String = Format(Absx1.dteFor("PO_DATE_L").Value, "dd-MMM-yyyy")

        Dim DANAC As String = ""

        With dst
            ASCMAIN1.sql = "select T2.po_order_no,   T1.VEND_CODE,   t1.WHSE_CODE, T1.PORT_CODE_ORIG, T1.PO_DATE_ORDERED, T2.PO_DATE_SHIP_BY, T1.PO_DATE_CANCEL " & vbCrLf _
            & " , Case When NVL(T2.PO_CONF_DATE,'') IS NULL THEN T2.PO_CONF_DATE ELSE T2.PO_DATE_SHIP_BY END AS VEND_ETD_DATE " & vbCrLf _
            & " , CASE WHEN NVL(T6.PO_SHIP_ETA,'') IS NULL THEN T2.PO_DATE_ETA ELSE T6.PO_SHIP_ETA END AS PO_DATE_ETA, T2.PO_BOOK_BY_DATE " & vbCrLf _
            & " , T2.PO_ON_BOARD_DATE,T6.PO_DATE_SHIPPED ACT_SHIP_DATE, T1.CUST_CODE, T3.CUST_NAME, T1.PO_CARTON_MARKS" & vbCrLf _
            & " , CASE WHEN NVL(T1.PO_DATE_CANCELLED,'') IS NULL THEN T1.PO_STATUS ELSE 'X' END AS PO_STATUS, T2.VEND_CARGO_READY_DATE , t2.PO_ORIG_DATE_SHIP_BY" & vbCrLf _
            & " , case WHEN NVL(T6.PO_DATE_SHIPPED,'') IS NULL THEN '0'  ELSE to_char(T6.PO_DATE_SHIPPED - T2.PO_DATE_SHIP_BY) END AS SHIP_DAYS_ACT " & vbCrLf _
            & " , case WHEN NVL(T6.PO_DATE_SHIPPED,'') IS NULL THEN '0'  ELSE to_char(T6.PO_DATE_SHIPPED - T2.PO_ORIG_DATE_SHIP_BY  ) END AS SHIP_DAYS_ORIG " & vbCrLf _
            & " ,SUM(T2.PO_QTY_ORD * T2.PO_COST_VCOST) ORIG_COST, SUM(T2.PO_QTY_OPN * T2.PO_COST_VCOST) OPEN_COST" & vbCrLf _
            & " ,SUM (TRUNC(T2.PO_QTY_ORD  * NVL(T4.CASE_CUBE,0) / DECODE(NVL(T4.CARTON_PACK_QTY,0),0,1,NVL(T4.CARTON_PACK_QTY,0)) * 100) / 100) PO_CUBE_ORD " & vbCrLf _
            & " ,SUM (TRUNC(T2.PO_QTY_OPN  * NVL(T4.CASE_CUBE,0) / DECODE(NVL(T4.CARTON_PACK_QTY,0),0,1,NVL(T4.CARTON_PACK_QTY,0)) * 100) / 100) PO_CUBE_OPN " & vbCrLf _
            & " ,SUM(T2.PO_QTY_ORD) ORDER_QTY, SUM(T2.PO_QTY_OPN)  OPEN_QTY " & vbCrLf _
            & "FROM potordr2 t2, potordr1 t1, ARTCUST1 T3, ICTSTYL1 T4, POTSHIP3 T5, POTSHIP1 T6 " & vbCrLf _
            & "where(t2.po_order_no = t1.po_order_no)" & vbCrLf _
            & "and T1.CUST_CODE = T3.CUST_CODE (+)" & vbCrLf _
            & "and T2.STYLE_CODE = T4.STYLE_CODE " & vbCrLf _
            & "and T2.po_order_no = T5.po_order_no (+) " & vbCrLf _
            & "and T2.po_order_Lno = T5.po_order_Lno (+) " & vbCrLf _
            & " and T5.PO_SHIPMENT_NO = T6.PO_SHIPMENT_NO(+) " & vbCrLf _
            & " and T1.PO_DATE_ORDERED > '" & horizon & "'" & vbCrLf _
            & "GROUP BY T2.po_order_no, T1.VEND_CODE, t1.WHSE_CODE, T1.PORT_CODE_ORIG, T1.PO_DATE_ORDERED, case WHEN NVL(T2.PO_CONF_DATE,'') IS NULL THEN T2.PO_CONF_DATE ELSE T2.PO_DATE_SHIP_BY END ,T1.PO_DATE_CANCEL, T2.PO_DATE_SHIP_BY, CASE WHEN NVL(T6.PO_SHIP_ETA,'') IS NULL THEN T2.PO_DATE_ETA ELSE T6.PO_SHIP_ETA END, T2.PO_BOOK_BY_DATE, T2.PO_ON_BOARD_DATE" & vbCrLf _
            & ",T6.PO_DATE_SHIPPED, T1.PO_DATE_SHIP_BY, T1.CUST_CODE, T3.CUST_NAME, T1.PO_CARTON_MARKS,CASE WHEN NVL(T1.PO_DATE_CANCELLED,'') IS NULL THEN T1.PO_STATUS ELSE 'X' END, T2.VEND_CARGO_READY_DATE , t2.PO_ORIG_DATE_SHIP_BY" & vbCrLf _
            & ", case WHEN NVL(T6.PO_DATE_SHIPPED,'') IS NULL THEN '0'  ELSE to_char(T2.PO_DATE_SHIP_BY - T6.PO_DATE_SHIPPED) END,  case WHEN NVL(T6.PO_DATE_SHIPPED,'') IS NULL THEN '0'  ELSE to_char(T2.PO_ORIG_DATE_SHIP_BY - T6.PO_DATE_SHIPPED) END " & vbCrLf _
            & "ORDER BY T2.PO_ORDER_NO " & vbCrLf



            ' & "and t1.po_status = 'O'" & vbCrLf _
            sqlPOTCONF1 = ASCMAIN1.sql

            POTCONF1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from " & POTCONF1
            Create_TDA(.Tables.Add("POTCONF1"), POTCONF1, "**", 0, True)

            '.Tables("POTCONF1").Columns.Add("SHIP_DAYS_ACT", GetType(System.Decimal))
            '.Tables("POTCONF1").Columns.Add("SHIP_DAYS_ORIG", GetType(System.Decimal))


            ASCMAIN1.sql = "Select * FROM POTPPRM1 WHERE POTPPRM1_CODE = 'Z' " & vbCrLf
            sqlPOTPPRM1 = ASCMAIN1.sql

            POTPPRM1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from " & POTPPRM1
            Create_TDA(.Tables.Add("POTPPRM1"), POTPPRM1, "**", 0, True)


        End With


        grdPOTCONF1.DataSource = dst.Tables("POTCONF1")

        Create_Summary(grdPOTCONF1, "PO_ORDER_NO", "Count")

        With grdPOTCONF1.DisplayLayout.Bands(0)
            '.Columns("PO_ORDER_NO").Header.Fixed = True
            '.Columns("VEND_CODE").Header.Fixed = True
            '.Columns("WHSE_CODE").Header.Fixed = True
            .Columns("PORT_CODE_ORIG").Header.Fixed = True
            With .Columns("PO_ORDER_NO")
                .Header.Fixed = True
                .Width = 140
                .Header.VisiblePosition = 1
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With
            With .Columns("VEND_CODE")
                .Header.Fixed = True
                .Width = 140
                .Header.VisiblePosition = 2
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With
            With .Columns("WHSE_CODE")
                .Header.Fixed = True
                .Width = 70
                .Header.VisiblePosition = 3
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With
            With .Columns("PORT_CODE_ORIG")
                .Header.Fixed = True
                .Width = 70
                .Header.VisiblePosition = 4
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("PO_STATUS")
                .Header.Fixed = True
                .Width = 110
                .Header.VisiblePosition = 5
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With


            'With .Columns("SHIP_DAYS_ACT")
            '    .Header.Fixed = True
            '    .Header.Caption = "Ship Days Act"
            '    .Width = 60
            '    .Header.VisiblePosition = 25
            '    .Header.Appearance.BackColor = Drawing.Color.White
            '    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            '    .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            'End With

            'With .Columns("SHIP_DAYS_ORIG")
            '    .Header.Caption = "Ship Days Orig"
            '    .Header.Fixed = True
            '    .Width = 26
            '    .Header.VisiblePosition = 31
            '    .Header.Appearance.BackColor = Drawing.Color.White
            '    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            '    .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            'End With

            '.Columns("STYLE_STATUS").Header.Fixed = True
            '.Columns("STYLE_COLOR_STATUS").Header.Fixed = True
            ' .Columns("COLOR_DESC").Header.Fixed = True
            'For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
            '    gcol.Header.Appearance.BackColor = Drawing.Color.White
            '    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            '    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
            '    If gcol.Key = "NEW_PO_COST" Then
            '        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            '    End If
            'Next
        End With

        ' ASCMAIN1.Add_Value_List(grdPOTCONF1, "STYLE_STATUS")

        spl.Panel1Collapsed = True

        'MyBase.Absx1.dteFor("DTE1").DateTime = DateTime.Now
        Dim rowPOTPPRM1 As DataRow = LookUp("POTPPRM1", "Z")
        MyBase.Absx1.txtFor("BOOK_RPT_UPDATE_DATE").Text = rowPOTPPRM1.Item("BOOK_RPT_UPDATE_DATE") & ""
        MyBase.Absx1.txtFor("BOOK_RPT_UPDATE_OPER").Text = rowPOTPPRM1.Item("BOOK_RPT_UPDATE_OPER") & ""

        MyBase.Absx1.txtFor("CENT_IMP_UPDATE_DATE").Text = rowPOTPPRM1.Item("CENT_IMP_UPDATE_DATE") & ""
        MyBase.Absx1.txtFor("CENT_IMP_UPDATE_OPER").Text = rowPOTPPRM1.Item("CENT_IMP_UPDATE_OPER") & ""

        MyBase.Absx1.txtFor("CENT_IMP_EXECUTE_DATE").Text = rowPOTPPRM1.Item("CENT_IMP_EXECUTE_DATE") & ""
        MyBase.Absx1.txtFor("CENT_IMP_EXECUTE_OPER").Text = rowPOTPPRM1.Item("CENT_IMP_EXECUTE_OPER") & ""

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

 
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load Report"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)


            Case "Done"
                Mode_Settings(False)

            Case "Update"


            Case "Cancel"
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load Report").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Cancel").Visible = (ScreenMode And EntryMode = "E")
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"POTCONF1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        ' Absx1.txtFor("CUST_CODE").Text = ""
        ' Absx1.txtFor("SREP_CODE").Text = ""

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "E" Then
        Else

            ASCMAIN1.sql = "TRUNCATE TABLE " & POTCONF1
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "INSERT INTO " & POTCONF1 & " SELECT X.* FROM (" & sqlPOTCONF1 & ") X "
            ' ASCMAIN1.sql = "INSERT INTO " & POTCONF1 & " SELECT X.*,'','','','','' FROM (" & sqlPOTCONF1 & ") X "
            ASCDATA1.ExecuteSQL()

            'DANAC= INSERT X.*,'','','','','' FROM (   X 


            Fill_Records("POTCONF1")
        End If


        Sort_grdColumns(grdPOTCONF1, "PO_ORDER_NO")


        POTPPRM1_CODE = "Z"
        'rowPOTPPRM1 = Fill_Record("POTPPRM1", POTPPRM1_CODE)
        rowPOTPPRM1 = LookUp("POTPPRM1", POTPPRM1_CODE)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub



    Sub SAVE_Record()

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTCONF1, "SBS", "Show Filter", "PO Inquiry", "Show GroupBox")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing


        Select Case e.SourceControl.Name


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdPOTCONF1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.OwningMenu IsNot Nothing Then grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        ' PO Inq view


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = ""

                PO_ORDER_NO =
                    grd.ActiveRow.Cells("PO_ORDER_NO").Value & ""

                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI")

        End Select
    End Sub

#End Region




#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "CUST_CODE"

        End Select

    End Sub

    Overrides Sub dte_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select

    End Sub


#End Region

    Private Sub UltraGroupBox3_Click(sender As System.Object, e As System.EventArgs) Handles UltraGroupBox3.Click

    End Sub
End Class