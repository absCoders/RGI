Public Class POFFRTA1
    ' GET AVGS TO SHOW ON TOTAL LINES
    ' FLEX CODES FOR TOP AND BOTTON GRID
 
    Dim POTFRTA1 As String
    Dim POTFRTA2 As String
    Dim STYLE_CODES As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -36, 0, -11)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -36, 0, 0)

        With dst

            Create_POTFRTA0("", "")
            ASCMAIN1.sql = "Select * from " & POTFRTA1
            Create_TDA(.Tables.Add, "POTFRTA1", "**", 0, False, "", 0)
            With .Tables("POTFRTA1").Columns
                .Add("AVGLDD", GetType(System.Decimal), "IIF(UNITS=0,0,LANDED/UNITS)")
                .Add("AVGFRT", GetType(System.Decimal), "IIF(UNITS=0,0,FRTIN/UNITS)")
            End With

            ASCMAIN1.sql = "Select X.SUB_BODY_CODE G_CODE, ICTBODY2.SUB_BODY_DESC G_DESC, SUM (UNITS) UNITS" & vbCrLf _
                & ", SUM (LANDED) LANDED, SUM (FRTIN) FRTIN" & vbCrLf _
                & " from  " & POTFRTA1 & " X, ICTBODY2 WHERE ICTBODY2.SUB_BODY_CODE (+) = X.SUB_BODY_CODE" & vbCrLf _
                & " group by X.SUB_BODY_CODE, ICTBODY2.SUB_BODY_DESC"
            Create_TDA(.Tables.Add, "POTFRTA0", "**", 0, False, "", 1)
            With .Tables("POTFRTA0").Columns
                .Add("AVGLDD", GetType(System.Decimal), "IIF(UNITS=0,0,LANDED/UNITS)")
                .Add("AVGFRT", GetType(System.Decimal), "IIF(UNITS=0,0,FRTIN/UNITS)")
            End With

            ASCMAIN1.sql = "Select * from " & POTFRTA2
            Create_TDA(.Tables.Add, "POTFRTA2", "**", 0, False, "", 0)

        End With

        grdPOTFRTA2.DataSource = dst.Tables("POTFRTA2")
        Create_Summary(grdPOTFRTA2, "STYLE_CODE", "Count")
        Create_Summary(grdPOTFRTA2, "PO_QTY_REC")

        grdPOTFRTA1.DataSource = dst.Tables("POTFRTA1")
        Create_Summary(grdPOTFRTA1, "FACTORY_CODE", "Count")

        grdPOTFRTA0.DataSource = dst.Tables("POTFRTA0")
        Create_Summary(grdPOTFRTA0, "G_CODE", "Count")

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTFRTA0, grdPOTFRTA1}
            With grd.DisplayLayout.Bands(0)
                For Each C As String In New String() {"UNITS", "LANDED", "FRTIN"}
                    With .Columns(C)
                        .Width = 100
                        .Format = "#,##0"
                        Create_Summary(grd, New String() {C})
                    End With
                Next
                For Each C As String In New String() {"AVGLDD", "AVGFRT"}
                    With .Columns(C)
                        .Width = 70
                        .Format = "#,##0.0000"
                    End With
                Next
            End With
        Next

        spl.Panel1Collapsed = True
        optVendorSubBody.Value = "S"

        MakeTransparent(chkShowAll)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Print"
                Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Period Range").Enabled = Not ScreenMode
                .Groups("Special Functions").Enabled = Not ScreenMode
            End With

        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpSpecialFunctions, ScreenMode)

        splSales.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"POTFRTA0", "POTFRTA1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        lblVendor.Text = ""
        lblSubBody.Text = ""
        txtSTYLE_CODE.Text = ""
        STYLE_CODES = ""


    End Sub

    Sub Load_Record()
        Dim STYLE_CHK As String = txtSTYLE_CODE.Text
        If STYLE_CHK <> "" Then
            STYLE_CHK = " - Style Code " & STYLE_CHK & ""
        End If


        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Reading from Sales History Data")

        Save_Header_Fields(UltraGroupBox1)

        Dim RYP0 As String = Absx1.cmbFor("RYP0").Value
        Dim RYP1 As String = Absx1.cmbFor("RYP1").Value

        Create_POTFRTA0(RYP0, RYP1)

        ASCMAIN1.Progress("Now Loading Data")

        EnforceConstraints(False)

        If optVendorSubBody.Value = "V" Then
            ASCMAIN1.sql = "Select NVL(X.FACTORY_CODE,'NoVend') G_CODE, SUBSTR(APTVEND1.VEND_NAME,1,35) G_DESC, SUM (UNITS) UNITS" & vbCrLf _
                & ", SUM (LANDED) LANDED, SUM (FRTIN) FRTIN" & vbCrLf _
                & " from  " & POTFRTA1 & " X, APTVEND1 WHERE APTVEND1.VEND_CODE (+) = X.FACTORY_CODE" & vbCrLf _
                & " group by X.FACTORY_CODE, APTVEND1.VEND_NAME"
            Fill_Records("POTFRTA0", String.Empty, True, ASCMAIN1.sql)
        ElseIf optVendorSubBody.Value = "C" Then
            ASCMAIN1.sql = "Select NVL(X.FACTORY_CODE,'NoCust') G_CODE, SUBSTR(ARTCUST1.CUST_NAME,1,35) G_DESC, SUM (UNITS) UNITS" & vbCrLf _
                & ", SUM (LANDED) LANDED, SUM (FRTIN) FRTIN" & vbCrLf _
                & " from  " & POTFRTA1 & " X, ARTCUST1 WHERE ARTCUST1.CUST_CODE (+) = X.FACTORY_CODE" & vbCrLf _
                & " group by X.FACTORY_CODE, ARTCUST1.CUST_NAME"
            Fill_Records("POTFRTA0", String.Empty, True, ASCMAIN1.sql)
        Else
            Fill_Records("POTFRTA0")
        End If


        Fill_Records("POTFRTA1")
        Fill_Records("POTFRTA2")

        EnforceConstraints(True)

        Sort_grdColumns(grdPOTFRTA0, "G_CODE")

        If optVendorSubBody.Value = "V" Then
            grdPOTFRTA0.Text = "Freight Analysis by Factory from " & Absx1.cmbFor("RYP0").Text & " thru " & Absx1.cmbFor("RYP1").Text & STYLE_CHK
        ElseIf optVendorSubBody.Value = "C" Then
            grdPOTFRTA0.Text = "Freight Analysis by Customer from " & Absx1.cmbFor("RYP0").Text & " thru " & Absx1.cmbFor("RYP1").Text & STYLE_CHK
        Else
            grdPOTFRTA0.Text = "Freight Analysis by Sub-Body from " & Absx1.cmbFor("RYP0").Text & " thru " & Absx1.cmbFor("RYP1").Text & STYLE_CHK
        End If

        Setup_grdPOTFRTA1()

        ASCMAIN1.Progress("Now Setting Up Screen")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTFRTA0, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdPOTFRTA1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))

        Select Case e.SourceControl.Name

            Case Else
        End Select
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Customer Inquiry"
            '    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
            '    Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
        End Select


    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            'Case "STATE_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
            '        Call Click_Command("Load", e)
            '    End If
            Case "STYLE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Dim row As DataRow = LookUp("ICTSTYL1", txtSTYLE_CODE.Text)

                    If row Is Nothing Then
                        lblVendor.Text = ""
                        lblSubBody.Text = ""

                        MsgBox("No record of Style " & txtSTYLE_CODE.Text)
                        Exit Sub
                    Else
                        '    STYLE_CODES = STYLE_CODES & txtSTYLE_CODE.Text & ","
                        ' lblVendor.Text = "Styles Sel " & STYLE_CODES
                        lblVendor.Text = "Vendor: " & row.Item("VEND_CODE") & ""
                        lblSubBody.Text = "Sub-Body: " & row.Item("SUB_BODY_CODE") & ""
                        'If row.Item("FACTORY_CODE") & "" = "" Then
                        '    MsgBox("No Factory Code for Style " & txtSTYLE_CODE.Text)
                        'End If
                        'If row.Item("SUB_BODY_CODE") & "" = "" Then
                        '    MsgBox("No Sub-Body Code for Style " & txtSTYLE_CODE.Text)
                        'End If
                        ' IF MULTIPLE txtSTYLE_CODE.Text = ""
                    End If
                    End If
        End Select

    End Sub

#End Region

    Sub Create_POTFRTA0(ByVal FYP As String, ByVal RYP As String)
        '& ", POTSHIP3.PO_COST_LANDED, POTSHIP3.PO_COST_VCOST, POTSHIP3.PO_COST_FREIGHT_IN" & vbCrLf _

        Dim STYLE_CHK As String = txtSTYLE_CODE.Text
        If STYLE_CHK <> "" Then
            STYLE_CHK = " AND POTORDR2.STYLE_CODE = '" & STYLE_CHK & "'"
        End If

        If optVendorSubBody.Value = "C" Then
            ASCMAIN1.sql = "Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR1.FACTORY_CODE , ICTSTYL1.CUST_CODE , ICTSTYL1.SUB_BODY_CODE," & vbCrLf _
            & "POTSHIP2.OPS_YYYYPP, POTSHIP3.*" & vbCrLf _
            & " from POTORDR2,POTSHIP3,POTORDR1,POTSHIP2,ICTSTYL1" & vbCrLf _
            & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            & "   and POTSHIP2.OPS_YYYYPP >= '" & FYP & "'" & vbCrLf _
            & "   and POTSHIP2.OPS_YYYYPP <= '" & RYP & "'" & vbCrLf _
            & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & STYLE_CHK

        Else
            ASCMAIN1.sql = "Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR1.FACTORY_CODE, ICTSTYL1.CUST_CODE, ICTSTYL1.SUB_BODY_CODE," & vbCrLf _
            & "POTSHIP2.OPS_YYYYPP, POTSHIP3.*" & vbCrLf _
            & " from POTORDR2,POTSHIP3,POTORDR1,POTSHIP2,ICTSTYL1" & vbCrLf _
            & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            & "   and POTSHIP2.OPS_YYYYPP >= '" & FYP & "'" & vbCrLf _
            & "   and POTSHIP2.OPS_YYYYPP <= '" & RYP & "'" & vbCrLf _
            & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & STYLE_CHK

        End If


        If POTFRTA2 = "" Then
            POTFRTA2 = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & POTFRTA2)
            ASCDATA1.ExecuteSQL("Insert into " & POTFRTA2 & " " & ASCMAIN1.sql)
        End If

        STYLE_CHK = txtSTYLE_CODE.Text
        If STYLE_CHK <> "" Then
            STYLE_CHK = " WHERE STYLE_CODE = '" & STYLE_CHK & "'"
        End If
        If optVendorSubBody.Value = "" Then
            ASCMAIN1.sql = "Select SUB_BODY_CODE || '    ' SUB_BODY_CODE, FACTORY_CODE || '    ' FACTORY_CODE" & vbCrLf _
            & ", SUM (PO_QTY_REC) UNITS, SUM (PO_QTY_REC * PO_COST_LANDED) LANDED" & vbCrLf _
            & ", SUM (PO_QTY_REC * PO_COST_FREIGHT_IN) FRTIN" & vbCrLf _
            & "  from " & POTFRTA2 & STYLE_CHK & vbCrLf _
            & "  group by SUB_BODY_CODE || '    '  ,  FACTORY_CODE || '    '"
        ElseIf optVendorSubBody.Value = "C" Then
            ASCMAIN1.sql = "Select SUB_BODY_CODE, CUST_CODE" & vbCrLf _
            & ", SUM (PO_QTY_REC) UNITS, SUM (PO_QTY_REC * PO_COST_LANDED) LANDED" & vbCrLf _
            & ", SUM (PO_QTY_REC * PO_COST_FREIGHT_IN) FRTIN" & vbCrLf _
            & "  from " & POTFRTA2 & STYLE_CHK & vbCrLf _
            & "  group by SUB_BODY_CODE, CUST_CODE"
        Else
            ASCMAIN1.sql = "Select SUB_BODY_CODE, FACTORY_CODE" & vbCrLf _
            & ", SUM (PO_QTY_REC) UNITS, SUM (PO_QTY_REC * PO_COST_LANDED) LANDED" & vbCrLf _
            & ", SUM (PO_QTY_REC * PO_COST_FREIGHT_IN) FRTIN" & vbCrLf _
            & "  from " & POTFRTA2 & STYLE_CHK & vbCrLf _
            & "  group by SUB_BODY_CODE,  FACTORY_CODE"
        End If

        If POTFRTA1 = "" Then
            POTFRTA1 = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & POTFRTA1)
            ASCDATA1.ExecuteSQL("Insert into " & POTFRTA1 & " " & ASCMAIN1.sql)
        End If

    End Sub

    Sub Print_Report()
        Dim SUBT As String = ""
        Print_Report_Begin()
        Print_Report_End()
    End Sub
     
    Private Sub grdPOTFRTA1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTFRTA1.InitializeRow
        'If e.Row.IsDataRow And Not e.Row.IsFilterRow Then
        '    If e.Row.Cells("STATE_CODE").Value & "" <> e.Row.Cells("CUST_STATE").Value & "" Then
        '        e.Row.Cells("CUST_STATE").Appearance.ForeColor = Drawing.Color.Red
        '    End If
        'End If
    End Sub

    Private Sub grdPOTFRTA0_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTFRTA0.AfterRowActivate
        Setup_grdPOTFRTA1()
    End Sub

    Sub Setup_grdPOTFRTA1()
        Dim STYLE_CHK As String = txtSTYLE_CODE.Text
        If STYLE_CHK <> "" Then
            STYLE_CHK = " - Style Code " & STYLE_CHK & ""
        End If

        If grdPOTFRTA0.ActiveRow Is Nothing Then
            grdPOTFRTA1.Visible = False
        Else
            If optVendorSubBody.Value = "V" Then
                Dim FACTORY_CODE As String = grdPOTFRTA0.ActiveRow.Cells("G_CODE").Value
                Dim dvw As DataView = DirectCast(grdPOTFRTA1.DataSource, DataTable).DefaultView
                If chkShowAll.Checked Then
                    grdPOTFRTA1.DisplayLayout.Bands(0).Columns("FACTORY_CODE").Hidden = False
                    grdPOTFRTA1.DisplayLayout.Bands(0).Columns("SUB_BODY_CODE").Hidden = False
                    dvw.RowFilter = ""
                    grdPOTFRTA1.Text = "Freight Analysis by Sub-Body from " & Absx1.cmbFor("RYP0").Text & " thru " & Absx1.cmbFor("RYP1").Text & STYLE_CHK
                Else
                    grdPOTFRTA1.DisplayLayout.Bands(0).Columns("FACTORY_CODE").Hidden = True
                    grdPOTFRTA1.DisplayLayout.Bands(0).Columns("SUB_BODY_CODE").Hidden = False
                    dvw.RowFilter = "FACTORY_CODE = '" & FACTORY_CODE & "'"
                    grdPOTFRTA1.Text = "Freight Analysis by Sub-Body within Factory " & FACTORY_CODE & STYLE_CHK
                End If

                Sort_grdColumns(grdPOTFRTA1, "SUB_BODY_CODE")
                grdPOTFRTA1.Visible = True

            ElseIf optVendorSubBody.Value = "C" Then
                Dim FACTORY_CODE As String = grdPOTFRTA0.ActiveRow.Cells("G_CODE").Value
                Dim dvw As DataView = DirectCast(grdPOTFRTA1.DataSource, DataTable).DefaultView
                If chkShowAll.Checked Then
                    grdPOTFRTA1.DisplayLayout.Bands(0).Columns("FACTORY_CODE").Hidden = False
                    grdPOTFRTA1.DisplayLayout.Bands(0).Columns("SUB_BODY_CODE").Hidden = False
                    dvw.RowFilter = ""
                    grdPOTFRTA1.Text = "Freight Analysis by Sub-Body from " & Absx1.cmbFor("RYP0").Text & " thru " & Absx1.cmbFor("RYP1").Text & STYLE_CHK
                Else
                    grdPOTFRTA1.DisplayLayout.Bands(0).Columns("FACTORY_CODE").Hidden = True
                    grdPOTFRTA1.DisplayLayout.Bands(0).Columns("SUB_BODY_CODE").Hidden = False
                    dvw.RowFilter = "FACTORY_CODE = '" & FACTORY_CODE & "'"
                    grdPOTFRTA1.Text = "Freight Analysis by Sub-Body within Customer " & FACTORY_CODE & STYLE_CHK
                End If

                Sort_grdColumns(grdPOTFRTA1, "SUB_BODY_CODE")
                grdPOTFRTA1.Visible = True

            Else

                Dim SUB_BODY_CODE As String = grdPOTFRTA0.ActiveRow.Cells("G_CODE").Value
                Dim dvw As DataView = DirectCast(grdPOTFRTA1.DataSource, DataTable).DefaultView
                If chkShowAll.Checked Then
                    grdPOTFRTA1.DisplayLayout.Bands(0).Columns("SUB_BODY_CODE").Hidden = False
                    grdPOTFRTA1.DisplayLayout.Bands(0).Columns("FACTORY_CODE").Hidden = False
                    dvw.RowFilter = ""
                    grdPOTFRTA1.Text = "Freight Analysis by Factory from " & Absx1.cmbFor("RYP0").Text & " thru " & Absx1.cmbFor("RYP1").Text & STYLE_CHK
                Else
                    grdPOTFRTA1.DisplayLayout.Bands(0).Columns("SUB_BODY_CODE").Hidden = True
                    grdPOTFRTA1.DisplayLayout.Bands(0).Columns("FACTORY_CODE").Hidden = False
                    dvw.RowFilter = "SUB_BODY_CODE = '" & SUB_BODY_CODE & "'"
                    grdPOTFRTA1.Text = "Freight Analysis by Factory within Sub-Body " & SUB_BODY_CODE & STYLE_CHK
                End If

                Sort_grdColumns(grdPOTFRTA1, "FACTORY_CODE")
                grdPOTFRTA1.Visible = True
            End If


        End If
    End Sub

    Private Sub grdPOTFRTA1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTFRTA1.AfterRowActivate
        Setup_grdPOTFRTA2()
    End Sub

    Sub Setup_grdPOTFRTA2()
        Dim STYLE_CHK As String = txtSTYLE_CODE.Text
        If STYLE_CHK <> "" Then
            STYLE_CHK = " - Style Code " & STYLE_CHK & ""
        End If

        If grdPOTFRTA1.ActiveRow Is Nothing Then
            grdPOTFRTA2.Visible = False
        Else
            'If optVendorSubBody.Value = "V" Then
            'ElseIf optVendorSubBody.Value = "C" Then
            'Else

            'End If
            Dim SUB_BODY_CODE As String = grdPOTFRTA1.ActiveRow.Cells("SUB_BODY_CODE").Value & ""
            Dim FACTORY_CODE As String = grdPOTFRTA1.ActiveRow.Cells("FACTORY_CODE").Value & ""
            '  Dim CUST_CODE As String = grdPOTFRTA1.ActiveRow.Cells("CUST_CODE").Value & ""
            Dim dvw As DataView = DirectCast(grdPOTFRTA2.DataSource, DataTable).DefaultView
            If chkShowAll.Checked Then
                dvw.RowFilter = ""
                grdPOTFRTA2.Text = "Freight Analysis by Shipment from " & Absx1.cmbFor("RYP0").Text & " thru " & Absx1.cmbFor("RYP1").Text & STYLE_CHK
            Else
                If optVendorSubBody.Value = "C" Then
                    dvw.RowFilter = $"ISNULL(SUB_BODY_CODE,'') = '{SUB_BODY_CODE}' and ISNULL(CUST_CODE,'') = '{FACTORY_CODE}'"
                Else
                    dvw.RowFilter = $"ISNULL(SUB_BODY_CODE,'') = '{SUB_BODY_CODE}' and ISNULL(FACTORY_CODE,'') = '{FACTORY_CODE}'"
                End If
                If optVendorSubBody.Value = "V" Then
                    grdPOTFRTA2.Text = "Shipment Details within Factory " & FACTORY_CODE & " for Sub-Body " & SUB_BODY_CODE & STYLE_CHK
                ElseIf optVendorSubBody.Value = "C" Then
                    grdPOTFRTA2.Text = "Shipment Details within Customer " & FACTORY_CODE & " for Sub-Body " & SUB_BODY_CODE & STYLE_CHK
                Else
                    grdPOTFRTA2.Text = "Shipment Details within Sub-Body " & SUB_BODY_CODE & " for Factory " & FACTORY_CODE & STYLE_CHK
                End If
            End If

            Sort_grdColumns(grdPOTFRTA2, "STYLE_CODE,COLOR_CODE")
            grdPOTFRTA2.Visible = True

        End If
    End Sub

    Private Sub chkShowAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowAll.CheckedChanged
        Setup_grdPOTFRTA1()
        Setup_grdPOTFRTA2()
    End Sub

    Private Sub txtSTYLE_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtSTYLE_CODE.ValueChanged

    End Sub
End Class