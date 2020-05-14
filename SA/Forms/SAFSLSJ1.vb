Public Class SAFSLSJ1

    Dim SATSLSJ0 As String = ""
    Dim SATSLSJ1 As String = ""

    Dim RYP0 As String
    Dim RYP1 As String

    Dim MOS As Integer

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -36, 0, -11)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -36, 0, 0)

        With dst

            Create_SATSLSJ0("", "")
            ASCMAIN1.sql = "Select * from " & SATSLSJ1
            Create_TDA(.Tables.Add, "SATSLSJ1", "**", 0, False, "", 0)
            With .Tables("SATSLSJ1")
                For Cx As Integer = 0 To 12
                    Dim C As String = IIf(Cx = 0, "TOTAL", "SHP" & Format(Cx, "00"))
                    .Columns(C).DataType = GetType(System.Int64)
                Next
            End With

            ASCMAIN1.sql = "Select * from " & SATSLSJ0 & " SATSLSJ0"
            Create_TDA(.Tables.Add, "SATSLSJ0", "**", 0, False, "", 1)
            With .Tables("SATSLSJ0")
                For Each C As String In New String() {"ONHD", "ONPO", "INXT", "OPEN", "PICK"}
                    .Columns(C).DataType = GetType(System.Int64)
                Next
                For Cx As Integer = 0 To 12
                    Dim C As String = IIf(Cx = 0, "TOTAL", "SHP" & Format(Cx, "00"))
                    .Columns(C).DataType = GetType(System.Int64)
                Next
                .Columns.Add("AVA2SHIP", GetType(System.Int64), "ISNULL(ONHD,0)-ISNULL(PICK,0)")
                .Columns.Add("AVA2SELL", GetType(System.Int64), "ISNULL(ONHD,0)+ISNULL(ONPO,0)+ISNULL(INXT,0)-ISNULL(PICK,0)-ISNULL(OPEN,0)")
                .Columns.Add("OVER_SHORT", GetType(System.Int64), "IIF(ISNULL(SAFETY_STOCK,0)>0,ISNULL(SAFETY_STOCK,0)-ISNULL(AVA2SELL,0),0)")
                .Columns.Add("AVGMOSLS", GetType(System.Decimal), "(ISNULL(TOTAL,0)+ISNULL(OPEN,0)+ISNULL(PICK,0))/1")
                .Columns.Add("MOSUP", GetType(System.Decimal), "IIF((ISNULL(TOTAL,0)+ISNULL(OPEN,0)+ISNULL(PICK,0)) = 0, 0, ISNULL(AVA2SELL,0)/AVGMOSLS)")
            End With
        End With

        grdSATSLSJ1.DataSource = dst.Tables("SATSLSJ1")
        Create_Summary(grdSATSLSJ1, "CUST_CODE", "Count")
        For Cx As Integer = 0 To 12
            Dim C As String = IIf(Cx = 0, "TOTAL", "SHP" & Format(Cx, "00"))
            Create_Summary(grdSATSLSJ1, C)
        Next

        grdSATSLSJ0.DataSource = dst.Tables("SATSLSJ0")
        Create_Summary(grdSATSLSJ0, "STYLE_CODE", "Count")
        'Create_Summary(grdSATSLSJ0, New String() {"SHP3", "SHP2", "SHP1", "SHP0", "TOTAL"})
        For Cx As Integer = 0 To 12
            Dim C As String = IIf(Cx = 0, "TOTAL", "SHP" & Format(Cx, "00"))
            Create_Summary(grdSATSLSJ0, C)
        Next

        With grdSATSLSJ0.DisplayLayout.Bands(0)
            For Each C As String In New String() {"STYLE_CODE", "STYLE_DESC"}
                .Columns(C).Header.Fixed = True
            Next
            For Each C As String In New String() {"SAFETY_STOCK", "OVER_SHORT"}
                .Columns(C).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(C).Header.Appearance.BackColor2 = Drawing.Color.Gold
                .Columns(C).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next

            For Each C As String In New String() {"AVGMOSLS"}
                .Columns(C).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(C).Header.Appearance.BackColor2 = Drawing.Color.Orange
                .Columns(C).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next


            For Each C As String In New String() {"STYLE_CODE", "STYLE_DESC", "STYLE_CLASS_CODE", "STYLE_GROUP_CODE", "SALES_DIVISION_CODE"}
                .Columns(C).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(C).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                .Columns(C).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next

            For Each C As String In New String() {"ONHD", "ONPO", "INXT", "OPEN", "PICK", "AVA2SHIP", "AVA2SELL", "MOSUP"}
                .Columns(C).Width = 70
                .Columns(C).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(C).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                .Columns(C).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next

            For Cx As Integer = 0 To 12
                Dim C As String = IIf(Cx = 0, "TOTAL", "SHP" & Format(Cx, "00"))
                .Columns(C).Width = 70
                .Columns(C).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(C).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .Columns(C).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next

        End With


        With grdSATSLSJ1.DisplayLayout.Bands(0)
            For Each C As String In New String() {"CUST_CODE", "CUST_NAME", "SREP_CODE"}
                .Columns(C).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(C).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                .Columns(C).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            For Cx As Integer = 0 To 12
                Dim C As String = IIf(Cx = 0, "TOTAL", "SHP" & Format(Cx, "00"))
                .Columns(C).Width = 70
                .Columns(C).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(C).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .Columns(C).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
        End With

        spl.Panel1Collapsed = True

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                RYP0 = Absx1.cmbFor("RYP0").Value
                RYP1 = Absx1.cmbFor("RYP1").Value
  
                Dim N As Integer = ASCMAIN1.Period_Diff(RYP0, RYP1)

                If N > 11 Then
                    EMsg &= vbCr & "Cannot select a period range with more than 12 months"
                End If
                If N < 0 Then
                    EMsg &= vbCr & "Cannot select a period, Invalid Range"
                End If

                ' MOS = Val(numMonths.Value & "")
                MOS = N + 1
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
                .Groups("Period Range").Visible = True
                '  .Groups("Paramaters").Visible = False
                .Groups("Show Styles").Visible = ScreenMode
                .Groups("Options").Visible = ScreenMode
            End With

        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpPERIOD_RANGE, ScreenMode)
        Set_Read_Only(grpParameters, ScreenMode)
        Show_Filter(grdSATSLSJ0, True)

        splSales.Visible = ScreenMode
        optShowStyles.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SATSLSJ0", "SATSLSJ1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Sales && Inventory Data")

        dst.Tables("SATSLSJ0").Columns("AVGMOSLS").Expression = _
            "(ISNULL(TOTAL,0)+ISNULL(OPEN,0)+ISNULL(PICK,0))/" & CStr(MOS)

        Save_Header_Fields(UltraGroupBox1)

        Dim RYP0 As String = Absx1.cmbFor("RYP0").Value
        Dim RYP1 As String = Absx1.cmbFor("RYP1").Value
        

        ' dgj
        Create_SATSLSJ0(RYP0, RYP1)

        '  Create_SATSLSJ0(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * (MOS - 1)), ASCMAIN1.CYP)

        ASCMAIN1.Progress("Now Loading Data")

        EnforceConstraints(False)

        Fill_Records("SATSLSJ0")
        Set_Styles_Filter()
        Fill_Records("SATSLSJ1")

        EnforceConstraints(True)

        Sort_grdColumns(grdSATSLSJ0, "STYLE_CODE")
        grdSATSLSJ0.Text = "Sales & Inventory by Style  from " & Absx1.cmbFor("RYP0").Text & " thru " & Absx1.cmbFor("RYP1").Text
        Setup_grdSATSLSJ1()

        ASCMAIN1.Progress("Now Setting Up Screen")

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATSLSJ0, grdSATSLSJ1}
            With grd.DisplayLayout.Bands(0)
                'dgjfR
                For I As Integer = 1 To MOS
                    Dim YP As String = ASCMAIN1.Period_Calc(RYP1, -1 * (MOS - I))
                    '  Dim YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * (MOS - I))

                    Dim LEGEND As String = ASCMAIN1.Get_Legend(YP)
                    .Columns("SHP" & Format(I, "00")).Header.Caption = Mid(LEGEND, 10, 6)
                Next
            End With
        Next

        If MOS <> 12 Then


            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATSLSJ0, grdSATSLSJ1}
                With grd.DisplayLayout.Bands(0)
                    'dgjfR
                    For I As Integer = (MOS + 1) To 12
                        .Columns("SHP" & Format(I, "00")).Hidden = True
                        '     this.customersDataGridView.Columns[0].Visible = false
                    Next
                End With
            Next
        End If


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
        Load_Popup_Menu(grdSATSLSJ0, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Inquiry")
        Load_Popup_Menu(grdSATSLSJ1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Customer Inquiry")
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
            Case "Style Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")

            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
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
        End Select

    End Sub

#End Region

    Sub Create_SATSLSJ0(ByVal YP_from As String, ByVal YP_to As String)

        Dim YPi As String = YP_from
        If YPi = "" Then YPi = ASCMAIN1.CYP
        Dim sqlP As String = ""
        For i As Integer = 1 To 12
            Dim YP As String = ASCMAIN1.Period_Calc(YPi, i - 1)
            sqlP &= ", SUM (DECODE(ORDR_YYYYPP_UPDATED,'" & YP & "',ORDR_QTY_SHIP,0)) SHP" & Format(i, "00") & vbCrLf
        Next

        ASCMAIN1.sql = "Select B.STYLE_CODE, B.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", ARTCUST1.SREP_CODE" & vbCrLf _
            & ", B.TOTAL, B.SHP01, B.SHP02, B.SHP03, B.SHP04, B.SHP05, B.SHP06, B.SHP07, B.SHP08, B.SHP09, B.SHP10, B.SHP11, B.SHP12" & vbCrLf _
            & " from ARTCUST1," & vbCrLf _
            & "(Select STYLE_CODE, CUST_CODE" & vbCrLf _
            & ", SUM (ORDR_QTY_SHIP) TOTAL" & vbCrLf _
            & sqlP _
            & " from SOTINVH2 where ORDR_YYYYPP_UPDATED BETWEEN '" & YP_from & "' and '" & YP_to & "'" & vbCrLf _
            & " group by STYLE_CODE,CUST_CODE) B" _
            & " where ARTCUST1.CUST_CODE = B.CUST_CODE"

        If SATSLSJ1 = "" Then
            SATSLSJ1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SATSLSJ1 & " add Primary Key (STYLE_CODE,CUST_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SATSLSJ1)
            ASCDATA1.ExecuteSQL("Insert into " & SATSLSJ1 & " " & ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
            & ", ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.STYLE_GROUP_CODE, ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
            & ", ICTSTYL1.SAFETY_STOCK" & vbCrLf _
            & ", A.ONHD, A.ONPO, A.INXT, A.OPEN, A.PICK, B.TOTAL, B.SHP01, B.SHP02, B.SHP03, B.SHP04, B.SHP05, B.SHP06, B.SHP07, B.SHP08, B.SHP09, B.SHP10, B.SHP11, B.SHP12" & vbCrLf _
            & " from ICTSTYL1," & vbCrLf _
            & "(Select STYLE_CODE" & vbCrLf _
            & ", SUM (WHSE_QTY_ON_HAND) ONHD" & vbCrLf _
            & ", SUM (WHSE_QTY_ON_ORDER) ONPO" & vbCrLf _
            & ", SUM (WHSE_QTY_TRAN) INXT" & vbCrLf _
            & ", SUM (WHSE_QTY_OPEN) OPEN" & vbCrLf _
            & ", SUM (WHSE_QTY_PICK) PICK" _
            & " from ICTSTAT2" & vbCrLf _
            & " group by STYLE_CODE) A," & vbCrLf _
            & "(Select STYLE_CODE" & vbCrLf _
            & ", SUM (ORDR_QTY_SHIP) TOTAL" & vbCrLf _
            & sqlP _
            & " from SOTINVH2 WHERE ORDR_YYYYPP_UPDATED BETWEEN '" & YP_from & "' and '" & YP_to & "'" & vbCrLf _
            & " group by STYLE_CODE) B" & vbCrLf _
            & " where A.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
            & "   and B.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
            & "   and (A.STYLE_CODE IS NOT NULL OR B.STYLE_CODE IS NOT NULL)" & vbCrLf _
            & "   and (NVL(B.TOTAL,0) <> 0 OR " & vbCrLf _
            & "NVL(A.ONHD,0) <> 0 OR NVL(A.ONPO,0) <> 0 OR NVL(A.INXT,0) <> 0 OR NVL(A.OPEN,0) <> 0 OR NVL(A.PICK,0) <> 0)"

        If SATSLSJ0 = "" Then
            SATSLSJ0 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SATSLSJ0 & " add Primary Key (STYLE_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SATSLSJ0)
            ASCDATA1.ExecuteSQL("Insert into " & SATSLSJ0 & " " & ASCMAIN1.sql)
        End If
    End Sub

    Sub Print_Report()
        Dim SUBT As String = ""
        Print_Report_Begin()
        Print_Report_End()
    End Sub
     
    Private Sub grdSATSLSJ1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs)
        If e.Row.IsDataRow And Not e.Row.IsFilterRow Then
            'If e.Row.Cells("STATE_CODE").Value & "" <> e.Row.Cells("CUST_STATE").Value & "" Then
            '    e.Row.Cells("CUST_STATE").Appearance.ForeColor = Drawing.Color.Red
            'End If
        End If
    End Sub

    Private Sub grdSATSLSJ0_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSATSLSJ0.AfterRowActivate
        Setup_grdSATSLSJ1()
    End Sub

    Sub Setup_grdSATSLSJ1()
        If grdSATSLSJ0.ActiveRow Is Nothing OrElse Not grdSATSLSJ0.ActiveRow.IsDataRow Then
            grdSATSLSJ1.Visible = False
        Else
            Dim STYLE_CODE As String = grdSATSLSJ0.ActiveRow.Cells("STYLE_CODE").Value
            Dim dvw As DataView = DirectCast(grdSATSLSJ1.DataSource, DataTable).DefaultView
            dvw.RowFilter = "STYLE_CODE = '" & STYLE_CODE & "'"
            Sort_grdColumns(grdSATSLSJ1, "CUST_CODE")
            grdSATSLSJ1.Visible = True
            grdSATSLSJ1.Text = "Sales by Customer within Style " & STYLE_CODE
        End If
    End Sub

    Private Sub grdSATSLSJ0_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSATSLSJ0.InitializeLayout

    End Sub

    Private Sub optShowStyles_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optShowStyles.ValueChanged
        Set_Styles_Filter()
    End Sub

    Sub Set_Styles_Filter()
        If SELECTION_NO = 0 Then Exit Sub

        grdSATSLSJ0.Text = "Sales & Inventory Analysis - " & optShowStyles.Text
        Dim sql As String = ""
        Select Case optShowStyles.Value
            Case "A"
                sql = ""
            Case "S"
                sql = "ISNULL(TOTAL,0) <> 0"
            Case "N"
                sql = "ISNULL(TOTAL,0) = 0"
        End Select

        Dim dvw As DataView = DirectCast(grdSATSLSJ0.DataSource, DataTable).DefaultView
        dvw.RowFilter = sql
    End Sub

    Private Sub chkSSSOnly_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSSSOnly.CheckedChanged
        Dim dvw As DataView = DirectCast(dst.Tables("SATSLSJ0"), DataTable).DefaultView
        If chkSSSOnly.Checked Then
            dvw.RowFilter = "ISNULL(SAFETY_STOCK,0) > 0"
        Else
            dvw.RowFilter = ""
        End If


    End Sub

    Private Sub UltraGroupBox1_Click(sender As System.Object, e As System.EventArgs) Handles UltraGroupBox1.Click

    End Sub
End Class