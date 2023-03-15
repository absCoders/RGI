Imports Infragistics.Win.UltraWinGrid

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
            ASCMAIN1.sql = "select T2.PO_ORDER_NO,   T1.VEND_CODE,   t1.WHSE_CODE, T1.PORT_CODE_ORIG, T1.PO_DATE_ORDERED, T2.PO_DATE_SHIP_BY, T1.PO_DATE_CANCEL " & vbCrLf _
            & " , Case When NVL(T2.PO_CONF_DATE,'') IS NULL THEN T2.PO_CONF_DATE ELSE T2.PO_DATE_SHIP_BY END AS VEND_ETD_DATE " & vbCrLf _
            & " , CASE WHEN NVL(T6.PO_SHIP_ETA,'') IS NULL THEN T2.PO_DATE_ETA ELSE T6.PO_SHIP_ETA END AS PO_DATE_ETA, T2.PO_BOOK_BY_DATE " & vbCrLf _
            & " , T2.PO_ON_BOARD_DATE,T6.PO_DATE_SHIPPED ACT_SHIP_DATE, T1.CUST_CODE, T3.CUST_NAME, T1.PO_CARTON_MARKS, T1.ORDR_NO" & vbCrLf _
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
            & ",T6.PO_DATE_SHIPPED, T1.PO_DATE_SHIP_BY, T1.CUST_CODE, T3.CUST_NAME, T1.PO_CARTON_MARKS, T1.ORDR_NO,CASE WHEN NVL(T1.PO_DATE_CANCELLED,'') IS NULL THEN T1.PO_STATUS ELSE 'X' END, T2.VEND_CARGO_READY_DATE , t2.PO_ORIG_DATE_SHIP_BY" & vbCrLf _
            & ", case WHEN NVL(T6.PO_DATE_SHIPPED,'') IS NULL THEN '0'  ELSE to_char(T2.PO_DATE_SHIP_BY - T6.PO_DATE_SHIPPED) END,  case WHEN NVL(T6.PO_DATE_SHIPPED,'') IS NULL THEN '0'  ELSE to_char(T2.PO_ORIG_DATE_SHIP_BY - T6.PO_DATE_SHIPPED) END " & vbCrLf _
            & "ORDER BY T2.PO_ORDER_NO " & vbCrLf
            ' & " And t1.po_status = 'O'" & vbCrLf _
            sqlPOTCONF1 = ASCMAIN1.sql

            POTCONF1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from " & POTCONF1
            Create_TDA(.Tables.Add("POTCONF1"), POTCONF1, "**", 0, True)
            With .Tables("POTCONF1")
                'ETA_PORT, ORDR_CUST_PO, ORDR_NO, ITEMS, ORDR_ARRIVAL_DATE, ORDR_LAST_ARRIVAL_DATE, DAYS_ARRIVAL_VS_ETA (r7-i7)
                .Columns.Add("ETA_PORT", GetType(System.DateTime))
                .Columns.Add("ORDR_CUST_PO", GetType(System.String))
                .Columns.Add("ITEMS", GetType(System.String))
                .Columns.Add("CTNS_OPEN", GetType(System.Decimal))
                .Columns.Add("ORDR_ARRIVAL_DATE", GetType(System.DateTime))
                .Columns.Add("ORDR_LAST_ARRIVAL_DATE", GetType(System.DateTime))
                .Columns.Add("DAYS_ARRIVAL_VS_ETA", GetType(System.Decimal))
            End With
            '.Tables("POTCONF1").Columns.Add("SHIP_DAYS_ACT", GetType(System.Decimal))
            '.Tables("POTCONF1").Columns.Add("SHIP_DAYS_ORIG", GetType(System.Decimal))

            With .Tables.Add("PIVOT_TABLE")
                .Columns.Add("PO_ORDER_NO")
                .Columns.Add("VEND_CODE")
                .Columns.Add("WHSE_CODE")
                .Columns.Add("PORT_CODE_ORIG")
                .Columns.Add("PO_STATUS")
                .Columns.Add("PO_DATE_ORDERED")
                .Columns.Add("PO_DATE_SHIP_BY")
                .Columns.Add("VEND_ETD_DATE")
                .Columns.Add("ETA_PORT")
                '.Columns.Add("PO_DATE_ETA")
                .Columns.Add("CUST_NAME")
                .Columns.Add("PO_CARTON_MARKS", GetType(System.String))
                .Columns.Add("PO_CUBE_OPN")
                .Columns.Add("OPEN_QTY")
                .Columns.Add("ORDR_CUST_PO", GetType(System.String))
                .Columns.Add("ORDR_NO", GetType(System.String))
                .Columns.Add("ITEMS", GetType(System.String))
                .Columns.Add("CTNS_OPEN", GetType(System.Decimal))
                .Columns.Add("ORDR_ARRIVAL_DATE", GetType(System.DateTime))
                .Columns.Add("ORDR_LAST_ARRIVAL_DATE", GetType(System.DateTime))
                .Columns.Add("DAYS_ARRIVAL_VS_ETA", GetType(System.Decimal))
            End With

            With .Tables("POTCONF1")
                For r As Integer = 0 To .Columns.Count - 1
                    For n As Integer = 0 To dst.Tables("PIVOT_TABLE").Columns.Count - 1
                        If dst.Tables("POTCONF1").Columns(r).ColumnName = dst.Tables("PIVOT_TABLE").Columns(n).ColumnName Then
                            dst.Tables("PIVOT_TABLE").Columns(n).DataType = dst.Tables("POTCONF1").Columns(r).DataType
                        End If
                    Next
                Next
            End With

            ASCMAIN1.sql = "Select DISTINCT T2.PO_ORDER_NO, T2.PO_ORDER_LNO, o1.ORDR_CUST_PO, T1.VEND_CODE, t1.WHSE_CODE, T1.PORT_CODE_ORIG, T1.PO_DATE_ORDERED, T2.PO_DATE_SHIP_BY, T1.PO_DATE_CANCEL
                            , CASE WHEN NVL(T1.PO_DATE_CANCELLED,'') IS NULL THEN T1.PO_STATUS ELSE 'X' END AS PO_STATUS, T1.CUST_CODE, T3.CUST_NAME, T1.PO_CARTON_MARKS, O2.STYLE_CODE, O2.COLOR_CODE, O2.CUST_SKU,  O2.CUST_COLOR_CODE
                            , (TRUNC(T2.PO_QTY_ORD  * NVL(T4.CASE_CUBE,0) / DECODE(NVL(T4.CARTON_PACK_QTY,0),0,1,NVL(T4.CARTON_PACK_QTY,0)) * 100) / 100) PO_CUBE_ORD
                            ,(T2.PO_QTY_ORD) ORDER_QTY, round((T2.PO_QTY_ORD / T2.CARTON_PACK_QTY),2) CARTONS
                            FROM POTORDR2 T2, POTORDR1 T1, ARTCUST1 T3, ICTSTYL1 T4, POTSHIP3 T5, POTSHIP1 T6, SOTORDR2 O2, SOTORDR1 O1
                            where(T2.PO_ORDER_NO = T1.PO_ORDER_NO)
                            and T1.CUST_CODE = T3.CUST_CODE (+)
                            and T2.STYLE_CODE = T4.STYLE_CODE 
                            and T2.PO_ORDER_NO = T5.PO_ORDER_NO (+)
                            and T2.PO_ORDER_LNO = T5.PO_ORDER_LNO (+)
                            and T5.PO_SHIPMENT_NO = T6.PO_SHIPMENT_NO(+)
                            and t2.ORDR_NO = o1.ORDR_NO
                            and t2.ORDR_NO = o2.ORDR_NO
                            and t2.ORDR_LNO = o2.ORDR_LNO
                            and t1.PO_STATUS = 'O'
                            and T1.PO_DATE_ORDERED > '" & horizon & "'"
            Create_TDA(.Tables.Add, "POTCONFD", "**", 0, False, "", 2)
            With .Tables("POTCONFD")
                .Columns.Add("SELECTED")
                .Columns("SELECTED").DefaultValue = "0"
                .Columns.Add("CTN_WIDTH", GetType(System.Decimal))
                .Columns.Add("CTN_HEIGHT", GetType(System.Decimal))
                .Columns.Add("CTN_LENGTH", GetType(System.Decimal))
                .Columns.Add("CTN_WEIGHT", GetType(System.Decimal))
            End With

            Dim DetailTable As DataTable = .Tables("POTCONFD").Clone()
            DetailTable.TableName = "DetailTable"
            .Tables.Add(DetailTable)
            With .Tables("DetailTable")
                .PrimaryKey = Nothing
                .Columns.Remove("SELECTED")
                .Columns.Remove("PO_ORDER_LNO")
                .Columns.Remove("WHSE_CODE")
                .Columns.Remove("PORT_CODE_ORIG")
                .Columns.Remove("PO_STATUS")
                .Columns.Remove("CUST_CODE")
                .Columns.Remove("CUST_NAME")
                ' .Columns.Remove("STYLE_CODE")
                .Columns.Remove("COLOR_CODE")
            End With

            ASCMAIN1.sql = "Select * FROM POTPPRM1 WHERE POTPPRM1_CODE = 'Z' " & vbCrLf
                sqlPOTPPRM1 = ASCMAIN1.sql

                POTPPRM1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

                ASCMAIN1.sql = "Select * from " & POTPPRM1
                Create_TDA(.Tables.Add("POTPPRM1"), POTPPRM1, "**", 0, True)


            End With


        grdPOTCONF1.DataSource = dst.Tables("POTCONF1")
        grdPOTCONFD.DataSource = dst.Tables("POTCONFD")


        Create_Summary(grdPOTCONF1, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTCONFD, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTCONFD, "SELECTED", "Count")

        Show_Filter(grdPOTCONFD, True)

        Dim CustColumn As UltraGridColumn = grdPOTCONFD.DisplayLayout.Bands(0).Columns("CUST_CODE")
        Dim CustColumnFilter As ColumnFilter = CustColumn.Band.ColumnFilters(CustColumn)
        CustColumnFilter.ClearFilterConditions()
        CustColumnFilter.FilterConditions.Add(FilterComparisionOperator.Equals, "171659")


        grdPOTCONFD.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        With grdPOTCONFD.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                If gcol.Key = "SELECTED" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                End If
            Next
        End With

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
        splGrids.Panel2Collapsed = True

        'MyBase.Absx1.dteFor("DTE1").DateTime = DateTime.Now
        Dim rowPOTPPRM1 As DataRow = LookUp("POTPPRM1", "Z")
        MyBase.Absx1.txtFor("BOOK_RPT_UPDATE_DATE").Text = rowPOTPPRM1.Item("BOOK_RPT_UPDATE_DATE") & ""
        MyBase.Absx1.txtFor("BOOK_RPT_UPDATE_OPER").Text = rowPOTPPRM1.Item("BOOK_RPT_UPDATE_OPER") & ""

        MyBase.Absx1.txtFor("CENT_IMP_UPDATE_DATE").Text = rowPOTPPRM1.Item("CENT_IMP_UPDATE_DATE") & ""
        MyBase.Absx1.txtFor("CENT_IMP_UPDATE_OPER").Text = rowPOTPPRM1.Item("CENT_IMP_UPDATE_OPER") & ""

        MyBase.Absx1.txtFor("CENT_IMP_EXECUTE_DATE").Text = rowPOTPPRM1.Item("CENT_IMP_EXECUTE_DATE") & ""
        MyBase.Absx1.txtFor("CENT_IMP_EXECUTE_OPER").Text = rowPOTPPRM1.Item("CENT_IMP_EXECUTE_OPER") & ""

        MyBase.Absx1.numFor("PORT_ADD_DAYS").Value = 60

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

            Case "Load QVC Report"
                EntryMode = "Q"
                Load_Record()
                Mode_Settings(True)

            Case "QVC Pivot Table"
                Create_Pivot()

            Case "Detail Style Extract"
                Create_Detail()

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
                    .Items("Load QVC Report").Settings.Enabled = IIf(not_iScreenMode = 1 Or EntryMode = "Q", 1, 2)
                    .Items("QVC Pivot Table").Settings.Enabled = IIf(EntryMode = "Q", 1, 2)
                    .Items("Detail Style Extract").Settings.Enabled = IIf(EntryMode = "Q", 1, 2)
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Cancel").Visible = (ScreenMode And EntryMode = "E")
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then

            With grdPOTCONF1.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If New String() {"PO_ORIG_DATE_SHIP_BY", "PO_DATE_CANCEL", "ACT_SHIP_DATE", "VEND_CARGO_READY_DATE",
                                    "PO_BOOK_BY_DATE", "PO_ON_BOARD_DATE", "CUST_CODE", "ORIG_COST", "OPEN_COST", "PO_CUBE_ORD",
                                    "ORDER_QTY", "SHIP_DAYS_ORIG", "SHIP_DAYS_ACT"}.Contains(gcol.Key) Then
                        gcol.Hidden = (EntryMode = "Q")
                    End If
                    If New String() {"ETA_PORT", "ORDR_CUST_PO", "ORDR_NO", "ITEMS", "CTNS_OPEN", "ORDR_ARRIVAL_DATE",
                                        "ORDR_LAST_ARRIVAL_DATE", "DAYS_ARRIVAL_VS_ETA"}.Contains(gcol.Key) Then
                        gcol.Hidden = Not (EntryMode = "Q")
                        If (EntryMode = "Q") Then
                            gcol.Header.Appearance.BackColor = Drawing.Color.White
                            gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                            gcol.Header.Appearance.BackColor2 = Drawing.Color.DodgerBlue
                        End If
                    End If
                    If New String() {"OPEN_QTY", "CTNS_OPEN", "DAYS_ARRIVAL_VS_ETA"}.Contains(gcol.Key) Then
                        gcol.Format = "#####0"
                    End If
                    If New String() {"PO_DATE_ORDERED", "VEND_ETD_DATE", "ETA_PORT", "PO_DATE_SHIP_BY"}.Contains(gcol.Key) Then
                        gcol.Format = "MM/dd/yy"
                    End If
                Next
            End With
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

        UltraExplorerBar1.Groups("Supplier").Visible = False

        ' Absx1.txtFor("CUST_CODE").Text = ""
        ' Absx1.txtFor("SREP_CODE").Text = ""
        splGrids.Panel1Collapsed = False
        splGrids.Panel2Collapsed = True

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        Save_Header_Fields(UltraGroupBox1)
        UltraExplorerBar1.Groups("Supplier").Visible = False
        splGrids.Panel1Collapsed = False
        splGrids.Panel2Collapsed = True

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

        If EntryMode = "Q" Then
            For Each row As DataRow In dst.Tables("POTCONF1").Select("WHSE_CODE = 'NC'")
                If row("VEND_ETD_DATE") & "" <> "" Then
                    row("ETA_PORT") = DateAdd(DateInterval.Day, MyBase.Absx1.numFor("PORT_ADD_DAYS").Value, row("VEND_ETD_DATE"))
                End If
            Next
            Dim dvw As DataView = DirectCast(grdPOTCONF1.DataSource, DataTable).DefaultView
            dvw.RowFilter = "WHSE_CODE = 'NC'"
            grdPOTCONF1.Text = "QVC Domestic Collect"

            ASCMAIN1.sql = "select POTCONF1.PO_ORDER_NO, SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
            & ", SOTORDR1.ORDR_ARRIVAL_DATE, SOTORDR1.ORDR_LAST_ARRIVAL_DATE" & vbCrLf _
            & ", ROUND(SUM(CTNS_OPEN),0) CTNS_OPEN, listagg (STYLE_CODE, ',') WITHIN GROUP (ORDER BY STYLE_CODE) ITEMS " & vbCrLf _
            & $" from SOTORDR1, {POTCONF1} POTCONF1, " & vbCrLf _
            & " (Select ORDR_NO, STYLE_CODE, SUM(ORDR_QTY_OPEN/CARTON_PACK_QTY) CTNS_OPEN" & vbCrLf _
            & " from SOTORDR2 GROUP BY ORDR_NO, STYLE_CODE ) SOTORDR2" & vbCrLf _
            & " Where SOTORDR1.ordr_no = POTCONF1.ORDR_NO" & vbCrLf _
            & " and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & " and POTCONF1.WHSE_CODE = 'NC'" & vbCrLf _
            & " group by POTCONF1.PO_ORDER_NO,SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
            & ", SOTORDR1.ORDR_ARRIVAL_DATE, SOTORDR1.ORDR_LAST_ARRIVAL_DATE"
            For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
                Dim rowPOTCONF1 As DataRow = dst.Tables("POTCONF1").Select("PO_ORDER_NO = '" & row("PO_ORDER_NO") & "'").FirstOrDefault
                rowPOTCONF1("ORDR_CUST_PO") = row("ORDR_CUST_PO")
                rowPOTCONF1("ORDR_ARRIVAL_DATE") = row("ORDR_ARRIVAL_DATE")
                rowPOTCONF1("ORDR_LAST_ARRIVAL_DATE") = row("ORDR_LAST_ARRIVAL_DATE")
                rowPOTCONF1("ITEMS") = row("ITEMS")
                rowPOTCONF1("CTNS_OPEN") = row("CTNS_OPEN")
                If IsDBNull(rowPOTCONF1("ETA_PORT")) Or IsDBNull(rowPOTCONF1("ORDR_LAST_ARRIVAL_DATE")) Then
                    rowPOTCONF1("DAYS_ARRIVAL_VS_ETA") = 0
                Else
                    rowPOTCONF1("DAYS_ARRIVAL_VS_ETA") = DateDiff(DateInterval.Day, rowPOTCONF1("ETA_PORT"), rowPOTCONF1("ORDR_LAST_ARRIVAL_DATE"))
                End If
            Next
        End If
        Sort_grdColumns(grdPOTCONF1, "PO_ORDER_NO")

        dst.Tables("POTCONF1").AcceptChanges()
        POTPPRM1_CODE = "Z"
        'rowPOTPPRM1 = Fill_Record("POTPPRM1", POTPPRM1_CODE)
        rowPOTPPRM1 = LookUp("POTPPRM1", POTPPRM1_CODE)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Create_Pivot()

        ASCMAIN1.Progress("Now Creating Workbook")

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsm"
        Dim DataTable As DataTable
        Dim r As Integer = 0

        If ASCMAIN1.Running_in_VS Then FILENAME = ASCMAIN1.Folders.Item("Work") & "Templates\" & Me.Name & ".xlsm"

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(FILENAME)
        ws = wb.Worksheets("Data")

        With dst.Tables("POTCONF1")
            For Each row As DataRow In .Select("WHSE_CODE = 'NC'")
                Dim rowPIVOT As DataRow = dst.Tables("PIVOT_TABLE").NewRow
                For n As Integer = 0 To dst.Tables("PIVOT_TABLE").Columns.Count - 1
                    Dim col_name As String = dst.Tables("PIVOT_TABLE").Columns(n).ColumnName
                    rowPIVOT(col_name) = row(col_name)
                Next
                dst.Tables("PIVOT_TABLE").Rows.Add(rowPIVOT)
            Next
        End With

        DataTable = dst.Tables("PIVOT_TABLE")

        r = 0
        For Each row As DataRow In DataTable.Select("")
            r += 1
            ws.Range("A" & CStr(7 + r) & ":U" & CStr(7 + r)).Value2 = row.ItemArray
        Next
        wb.Names.Add("DataPivotBase", "=DATA!$A$7:$U$" & CStr(7 + DataTable.Rows.Count))

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "FRANKENSTEIN_REPORT"
                XLS_FILENAME &= "-" & Format(Today, "yyyyMMdd") & ".xlsm"

                Dim objOpt As Object = Nothing ' Missing.Value
                wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
                wb.Close(False, objOpt, objOpt)

                success = True

            Catch ex As Exception
                ' Stop
            End Try
        Loop

        excel.Quit()
        ws = Nothing
        wb = Nothing
        excel = Nothing
        xlSourceRange = Nothing
        xlDestRange = Nothing

        ReleaseCOMObject(xlDestRange)
        ReleaseCOMObject(xlSourceRange)
        ReleaseCOMObject(ws)
        ReleaseCOMObject(wb)
        ReleaseCOMObject(excel)

        'Add_Document_to_ASTSPRF1(ASCMAIN1.Folders("Work") & XLS_FILENAME)
        Show_Document(ASCMAIN1.Folders("Work") & XLS_FILENAME)

        ASCMAIN1.Progress("")
    End Sub

    Sub Create_Detail()
        UltraExplorerBar1.Groups("Supplier").Visible = True
        UltraExplorerBar1.Groups("Screen Control").Items("QVC Pivot Table").Settings.Enabled = DefaultableBoolean.False
        splGrids.Panel1Collapsed = True
        splGrids.Panel2Collapsed = False
        Absx1.txtFor("VEND_CODE").Value = ""
        Dim dvw As DataView = DirectCast(grdPOTCONFD.DataSource, DataTable).DefaultView
        dvw.RowFilter = ""
        grdPOTCONFD.Text = "PO Details for All Supliers"

        Fill_Records("POTCONFD")

    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click

        If Absx1.txtFor("VEND_CODE").Value = "" Then
            MessageBox.Show("Supplier Code can not be empty!", "Generate Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If dst.Tables("POTCONFD").Select("VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Value & "'").Length = 0 Then
            MessageBox.Show("Supplier Code has no records!", "Generate Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim FILENAME As String = GenerateExcelExport()

        If FILENAME = "" Then
            MessageBox.Show("No records were generated for extract!", "Generate Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim ATTACHMENTs As New Dictionary(Of String, String)
        Dim VEND_CODE As String = Absx1.txtFor("VEND_CODE").Value
        Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)
        Dim PO_ORDER_NOs As String = ""

        ATTACHMENTs.Add(FILENAME, ASCMAIN1.Folders("Temp") & FILENAME)

        Dim last_po As String = ""
        For Each row As DataRow In dst.Tables("DetailTable").Select("", "PO_ORDER_NO")
            If last_po <> row("PO_ORDER_NO") Then
                last_po = row("PO_ORDER_NO")
                PO_ORDER_NOs &= "," & last_po
            End If
        Next

        Dim SUBJECT As String = ""
        Dim PFX As String = ""
        If ASCMAIN1.CLIENT = "RGI" Then PFX = "Regency  "
        If Split(PO_ORDER_NOs, ",").Count = 1 Then
            SUBJECT = PFX & "PO " & PO_ORDER_NOs
        Else
            SUBJECT = PFX & "POs " & PO_ORDER_NOs.Substring(1)
        End If

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        EMAIL_ADDRESSs.Add(rowAPTVEND1.Item("VEND_PURCH_EMAIL") & "", rowAPTVEND1.Item("VEND_PURCH_CONTACT") & "")

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                SUBJECT, "PO_CTN", False, True, VEND_CODE, rowAPTVEND1.Item("VEND_NAME"), "Supplier")

    End Sub

    Function GenerateExcelExport() As String

        ASCMAIN1.Progress("Now Creating Workbook")

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\POFCONFD.xlsx"
        Dim DataTable As DataTable
        Dim r As Integer = 0

        If ASCMAIN1.Running_in_VS Then FILENAME = ASCMAIN1.Folders.Item("Work") & "Templates\POFCONFD.xlsx"

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(FILENAME)
        ws = wb.Worksheets("Data")

        dst.Tables("DetailTable").Rows.Clear()

        With dst.Tables("POTCONFD")
            For Each grow As UltraWinGrid.UltraGridRow In grdPOTCONFD.Rows.GetFilteredInNonGroupByRows
                If grow.Cells("SELECTED").Value = "1" Then
                    For Each row As DataRow In .Select($"PO_ORDER_NO = '{grow.Cells("PO_ORDER_NO").Value}' and PO_ORDER_LNO = '{grow.Cells("PO_ORDER_LNO").Value}'")
                        Dim rowDetails As DataRow = dst.Tables("DetailTable").NewRow
                        Dim addrow As Boolean = False
                        For n As Integer = 0 To dst.Tables("DetailTable").Columns.Count - 1
                            Dim col_name As String = dst.Tables("DetailTable").Columns(n).ColumnName
                            rowDetails(col_name) = row(col_name)
                            If ("CTN_WIDTH,CTN_HEIGHT,CTN_LENGTH,CTN_WEIGHT".Contains(col_name) And Val(row(col_name) & "") = 0) Then
                                addrow = True
                            End If
                        Next
                        If addrow = True Then
                            dst.Tables("DetailTable").Rows.Add(rowDetails)
                        End If
                    Next
                End If
            Next
        End With
        If dst.Tables("DetailTable").Rows.Count = 0 Then
            ASCMAIN1.Progress("")
            Return ""
        End If

        DataTable = dst.Tables("DetailTable")
        ws.Range("A1").Value2 = Today
        r = 0
        For Each row As DataRow In DataTable.Select("")
            r += 1
            ws.Range("A" & CStr(6 + r) & ":Q" & CStr(6 + r)).Value2 = row.ItemArray
        Next

        ws.Protect(UserInterfaceOnly:=True)

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = $"DataRequest_{Absx1.txtFor("VEND_CODE").Value}"
                XLS_FILENAME &= "-" & Format(Now, "yyyyMMddHHmm") & ".xlsx"

                Dim objOpt As Object = Nothing ' Missing.Value
                wb.SaveAs(ASCMAIN1.Folders("Temp") & XLS_FILENAME, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
                wb.Close(False, objOpt, objOpt)

                success = True

            Catch ex As Exception
                ' Stop
            End Try
        Loop

        excel.Quit()
        ws = Nothing
        wb = Nothing
        excel = Nothing
        xlSourceRange = Nothing
        xlDestRange = Nothing

        ReleaseCOMObject(xlDestRange)
        ReleaseCOMObject(xlSourceRange)
        ReleaseCOMObject(ws)
        ReleaseCOMObject(wb)
        ReleaseCOMObject(excel)

        'Add_Document_to_ASTSPRF1(ASCMAIN1.Folders("Work") & XLS_FILENAME)
        'Show_Document(ASCMAIN1.Folders("Work") & XLS_FILENAME)

        ASCMAIN1.Progress("")
        Return XLS_FILENAME

    End Function
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
        Load_Popup_Menu(grdPOTCONFD, "SBBBBB", "Show Filter", "De-Select All", "De-Select Selected", "Select All for Supplier", "Select Selected", "Select All")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.SourceControl.Name


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdPOTCONFD"
                    tlb_btn = DirectCast(tlb_pop.Tools("Select All for Supplier"), UltraWinToolbars.ButtonTool)
                    Dim VEND_CODE As String = grd.ActiveRow.Cells("VEND_CODE").Value
                    tlb_btn.SharedProps.Caption = $"Select All for {VEND_CODE}"

                    'tlb_btn.SharedProps.Visible = enableBtn

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

            Case "De-Select All"
                Dim tname As String = ""
                Select Case grd.Name
                    Case "grdPOTCONFD"
                        tname = "POTCONFD"
                        'Case "grdWHTINST1"
                        '    tname = "WHTINST1"
                        'Case "grdWHTWAVEP"
                        '    tname = "WHTWAVEP"
                End Select
                If tname <> "" Then
                    If tname = "POTCONFD" Then
                        For Each row As DataRow In dst.Tables(tname).Select("SELECTED = '1'")
                            row.Item("SELECTED") = "0"
                        Next
                    Else
                        For Each row As DataRow In dst.Tables(tname).Select("SEL = '1'")
                            row.Item("SEL") = "0"
                        Next
                    End If

                End If


            Case "Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grd.Name = "grdPOTCONFD" Then
                        If Not grow.IsFilteredOut And grow.IsDataRow Then
                            grow.Cells("SELECTED").Value = "1"
                        End If
                    Else
                        If Not grow.IsFilteredOut And grow.IsDataRow Then
                            grow.Cells("SEL").Value = "1"
                        End If
                    End If
                Next
                grd.UpdateData()

            Case "Select Selected", "DE-Select Selected"
                Dim sval As String = "1"
                If e.Tool.Key = "DE-Select Selected" Then sval = "0"

                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    If grd.Name = "grdPOTCONFD" Then
                        If Not grow.IsFilteredOut And grow.IsDataRow Then
                            grow.Cells("SELECTED").Value = sval
                        End If
                    Else
                        If Not grow.IsFilteredOut And grow.IsDataRow Then
                            grow.Cells("SEL").Value = sval
                        End If
                    End If
                Next
                grd.UpdateData()

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Select All for Supplier"
                Dim VEND_CODE As String = grd.ActiveRow.Cells("VEND_CODE").Value
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If Not grow.IsFilteredOut And grow.Cells("VEND_CODE").Value = VEND_CODE Then
                        grow.Cells("SELECTED").Value = "1"
                    End If
                Next
                grd.UpdateData()
                Absx1.txtFor("VEND_CODE").Value = VEND_CODE
        End Select
    End Sub

#End Region




#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "CUST_CODE"

            Case "VEND_CODE"


        End Select

    End Sub

    Overrides Sub txt_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "VEND_CODE"
                If sender.Value = "" Then
                    Dim dvw As DataView = DirectCast(grdPOTCONFD.DataSource, DataTable).DefaultView
                    dvw.RowFilter = ""
                    Sort_grdColumns(grdPOTCONFD, "PO_ORDER_NO")
                    grdPOTCONFD.Text = "PO Details for All Supliers"
                Else
                    If dst.Tables("POTCONFD").Select($"VEND_CODE = '{sender.Value}'").Length > 0 Then
                        Dim dvw As DataView = DirectCast(grdPOTCONFD.DataSource, DataTable).DefaultView
                        dvw.RowFilter = $"VEND_CODE = '{sender.Value}'"
                        Sort_grdColumns(grdPOTCONFD, "PO_ORDER_NO")
                        grdPOTCONFD.Text = "PO Details for Suplier - " & sender.Value
                    End If
                End If


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

    Private Sub grdPOTCONF1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdPOTCONF1.InitializeRow
        If EntryMode = "Q" And String.IsNullOrEmpty(e.Row.Cells("VEND_ETD_DATE").Text) Then
            e.Row.Cells("VEND_ETD_DATE").Value = e.Row.Cells("PO_DATE_SHIP_BY").Value
            e.Row.Cells("VEND_ETD_DATE").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("ETA_PORT").Value = DateAdd(DateInterval.Day, MyBase.Absx1.numFor("PORT_ADD_DAYS").Value, e.Row.Cells("PO_DATE_SHIP_BY").Value)
            e.Row.Update()
        End If
    End Sub

    Private Sub grdPOTCONFD_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdPOTCONFD.AfterRowUpdate
        'if a single line is selected pick all PO lines
        If Not (String.IsNullOrEmpty(e.Row.Cells("PO_ORDER_NO").Value)) Then
            Dim PO_NO As String = e.Row.Cells("PO_ORDER_NO").Value
            Dim PO_LNO As Integer = e.Row.Cells("PO_ORDER_LNO").Value
            For Each row As DataRow In dst.Tables("POTCONFD").Select($"PO_ORDER_NO = '{PO_NO}'")
                If row("PO_ORDER_LNO") <> PO_LNO Then
                    row("SELECTED") = e.Row.Cells("SELECTED").Value
                End If
            Next
        End If
    End Sub
End Class