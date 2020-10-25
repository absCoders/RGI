Imports System
Imports System.Text

Public Class WBFDISC1
    Dim SATCSLS1 As String
    Dim SATCSLSH As String
    Dim SATCSLSX As String
    Dim sqlSATCSLSH As String

    Dim RYP0 As String
    Dim RYP1 As String
    Dim RYW0 As String
    Dim RYW1 As String
    Dim Periods As Integer

    Dim CUST_CODE As String
    Dim SREP_CODE As String
    Dim CHECK_BOX As String
    Dim IMAGE_FOLDER_NAME As String
    Dim TTM As New UltraWinToolTip.UltraToolTipManager

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -36, 0, -11)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -36, 0, 0)
        Dim SQL As New Text.StringBuilder With {.Length = 0}
        With dst
            ASCMAIN1.sql = MakesqlSATCSLSX()
            SATCSLSX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SATCSLSX & " Add Primary Key (STYLE_CODE, COLOR_CODE, CUST_CODE)")

            ASCMAIN1.sql = MakesqlSATCSLS1()
            SATCSLS1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SATCSLS1 & " Add Primary Key (STYLE_CODE, COLOR_CODE)")

            ASCMAIN1.sql = "Select * from " & SATCSLS1
            Create_TDA(.Tables.Add, "SATCSLS1", "**", 0, False)
            ' Create_TDA(.Tables.Add, "SATCSLS1", "**", 0, False, "VVVV")
            With .Tables("SATCSLS1")
                .Columns.Add("SLS_AMT_WHSEX", GetType(System.Decimal), "ISNULL(SLS_AMT,0) - ISNULL(SLS_AMT_WHSE1,0) - ISNULL(SLS_AMT_WHSE2,0)")
                .Columns.Add("SLS_QTY_WHSEX", GetType(System.Int32), "ISNULL(SLS_QTY,0) - ISNULL(SLS_QTY_WHSE1,0) - ISNULL(SLS_QTY_WHSE2,0)")
                .Columns.Add("FUT_QTY_WHSEX", GetType(System.Int32), "ISNULL(FUT_QTY,0) - ISNULL(FUT_QTY_WHSE1,0) - ISNULL(FUT_QTY_WHSE2,0)")
                .Columns.Add("THEME_DESC")
                .Columns.Add("SEASON_CODE")
            End With

            '--Begin New Table
            With SQL
                .Length = 0
                .AppendLine("SELECT")
                .AppendLine("STYLE_CODE,")
                .AppendLine("STYLE_DESC,")
                .AppendLine("STYLE_STATUS,")
                .AppendLine("VEND_CODE,")
                .AppendLine("FACTORY_CODE,")
                .AppendLine("STYLE_UOM,")
                .AppendLine("STYLE_CLASS_CODE,")
                .AppendLine("CARTON_PACK_QTY,")
                .AppendLine("PO_COST,")
                .AppendLine("SLS_AMT,")
                .AppendLine("SLS_AMT_WHSE1,")
                .AppendLine("SLS_AMT_WHSE2,")
                .AppendLine("SLS_QTY,")
                .AppendLine("SLS_QTY_WHSE1,")
                .AppendLine("SLS_QTY_WHSE2,")
                .AppendLine("SLS_CUSTS,")
                .AppendLine("SLS_CUSTS_WHSE1,")
                .AppendLine("SLS_CUSTS_WHSE2,")
                .AppendLine("SLS_CUSTS_WHSEX,")
                .AppendLine("FUT_QTY,")
                .AppendLine("FUT_QTY_WHSE1,")
                .AppendLine("FUT_QTY_WHSE2,")
                .AppendLine("CUST_CODE,")
                .AppendLine("CUST_NAME")
                .AppendLine("from " & SATCSLS1)
            End With
            sqlSATCSLSH = SQL.ToString
            ASCMAIN1.sql = sqlSATCSLSH
            SATCSLSH = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SATCSLSH & " Add Primary Key (STYLE_CODE)")

            ASCMAIN1.sql = "Select * from " & SATCSLSH
            Create_TDA(.Tables.Add, "SATCSLSH", "**", 0, False)
            With .Tables("SATCSLSH")
                .Columns.Add("SLS_AMT_WHSEX", GetType(System.Decimal), "ISNULL(SLS_AMT,0) - ISNULL(SLS_AMT_WHSE1,0) - ISNULL(SLS_AMT_WHSE2,0)")
                .Columns.Add("SLS_QTY_WHSEX", GetType(System.Int32), "ISNULL(SLS_QTY,0) - ISNULL(SLS_QTY_WHSE1,0) - ISNULL(SLS_QTY_WHSE2,0)")
                .Columns.Add("FUT_QTY_WHSEX", GetType(System.Int32), "ISNULL(FUT_QTY,0) - ISNULL(FUT_QTY_WHSE1,0) - ISNULL(FUT_QTY_WHSE2,0)")
                .Columns.Add("ATTR_CODE1")
                .Columns.Add("ATTR_CODE2")
                .Columns.Add("ATTR_CODE3")
                .Columns.Add("ATTR_CODE4")
                .Columns.Add("ATTR_CODE5")
            End With
            '--End New Table

            ASCMAIN1.sql = MakesqlSOTINVHX("", "", RYP0, RYP1)
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "", 0)

            SQL.Length = 0
            With SQL
                .AppendLine("SELECT")
                .AppendLine("ST.STYLE_CODE,")
                .AppendLine("ST.COLOR_CODE,")
                .AppendLine("CL.COLOR_DESC,")
                .AppendLine("SC.UPC_CODE,")
                .AppendLine("TH.THEME_DESC,")
                .AppendLine("TH.SEASON_CODE,")
                .AppendLine("SUM(NVL(WHSE_QTY_ON_HAND,0)) AS ON_HAND,")
                .AppendLine("SUM(NVL(WHSE_QTY_PICK,0)) AS PICK,")
                .AppendLine("(SUM(NVL(WHSE_QTY_ON_HAND,0)) - SUM(NVL(WHSE_QTY_PICK,0))) AS OTS,")
                .AppendLine("SUM(NVL(WHSE_QTY_TRAN,0)) AS TRAN,")
                .AppendLine("SUM(NVL(WHSE_QTY_ON_ORDER,0)) AS ON_ORDER,")
                .AppendLine("SUM(NVL(WHSE_QTY_OPEN,0)) AS OPEN,")
                .AppendLine("((SUM(NVL(WHSE_QTY_ON_HAND,0)) - SUM(NVL(WHSE_QTY_PICK,0))) + SUM(NVL(WHSE_QTY_TRAN,0)) + SUM(NVL(WHSE_QTY_ON_ORDER,0)) - SUM(NVL(WHSE_QTY_OPEN,0))) AS OTS_WIP")
                .AppendLine("FROM ICTSTAT2 ST, ICTSTYC1 SC, ICTCOLR1 CL, ICTTHEME TH")
                .AppendLine("WHERE ST.STYLE_CODE = SC.STYLE_CODE")
                .AppendLine("AND ST.COLOR_CODE = SC.COLOR_CODE")
                .AppendLine("AND ST.COLOR_CODE = CL.COLOR_CODE")
                .AppendLine("AND SC.THEME_CODE = TH.THEME_CODE")
                .AppendLine("GROUP BY ST.STYLE_CODE,")
                .AppendLine("ST.COLOR_CODE,")
                .AppendLine("CL.COLOR_DESC,")
                .AppendLine("SC.UPC_CODE,")
                .AppendLine("TH.THEME_DESC,")
                .AppendLine("TH.SEASON_CODE")
                .AppendLine("ORDER BY ST.STYLE_CODE,")
                .AppendLine("ST.COLOR_CODE")
            End With
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "ICTSTATA", "**", 0, False, "", 2)

            SQL.Length = 0
            With SQL
                .AppendLine("SELECT *")
                .AppendLine(" FROM ASTGRID1")
                .AppendLine(" WHERE USER_ID = :PARM1")
                .AppendLine(" AND FORM_NAME = :PARM2")
                .AppendLine(" AND GRID_NAME = :PARM3")
            End With
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "ASTGRID1", "**", 0, True, "VVV")

            Dim s As New Text.StringBuilder With {.Length = 0}
            s.AppendLine("SELECT S3.STYLE_CODE,")
            s.AppendLine("S3.ATTR_CODE,")
            s.AppendLine("NVL(A1.ATT_RANK,9) ATT_RANK")
            s.AppendLine("FROM ICTSTYL3 S3, ICTATTR1 A1")
            s.AppendLine("WHERE S3.ATTR_CODE = A1.ATTR_CODE")
            ASCMAIN1.sql = s.ToString
            Create_TDA(.Tables.Add, "ICTSTYL3", "**", 0, False)

            ASCMAIN1.sql = "Select * from ICTTHEME"
            Create_TDA(.Tables.Add, "ICTTHEME", "**", 0, False, "", 1)

        End With

        Fill_Records("ICTSTYL3")
        Fill_Records("ICTTHEME")

        grdSATCSLS1.DataSource = dst.Tables("SATCSLS1")
        grdSATCSLSH.DataSource = dst.Tables("SATCSLSH")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")
        grdICTSTATA.DataSource = dst.Tables("ICTSTATA")

        Create_Summary(grdSOTINVHX, "INV_NO", "Count")
        Create_Summary(grdSOTINVHX, "ORDR_QTY_SHIP")
        Create_Summary(grdSOTINVHX, "ORDR_AMT_SHIP")
        Create_Summary(grdSOTINVHX, "ORDR_QTY_OPEN")
        Create_Summary(grdSOTINVHX, "ORDR_QTY_PICK")
        Create_Summary(grdSOTINVHX, "ORDR_QTY_CANC")

        Create_Summary(grdSATCSLS1, "STYLE_CODE", "Count")
        Create_Summary(grdSATCSLS1, New String() {"SLS_AMT", "SLS_AMT_WHSE1", "SLS_AMT_WHSE2", "SLS_AMT_WHSEX", "SLS_QTY", "SLS_QTY_WHSE1", "SLS_QTY_WHSE2", "SLS_QTY_WHSEX", "FUT_QTY", "FUT_QTY_WHSE1", "FUT_QTY_WHSE2", "FUT_QTY_WHSEX"}, , , "#,##0")
        Create_Summary(grdSATCSLSH, "STYLE_CODE", "Count")
        Create_Summary(grdSATCSLSH, New String() {"SLS_AMT", "SLS_AMT_WHSE1", "SLS_AMT_WHSE2", "SLS_AMT_WHSEX", "SLS_QTY", "SLS_QTY_WHSE1", "SLS_QTY_WHSE2", "SLS_QTY_WHSEX", "FUT_QTY", "FUT_QTY_WHSE1", "FUT_QTY_WHSE2", "FUT_QTY_WHSEX"}, , , "#,##0")

        With grdSATCSLS1.DisplayLayout.Bands("SATCSLS1")
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("STYLE_DESC").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
            .Columns("COLOR_DESC").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"SLS_AMT", "SLS_AMT_WHSE1", "SLS_AMT_WHSE2", "SLS_AMT_WHSEX"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.Width = 80
                    gcol.Format = "#,##0"
                End If
                If New String() {"SLS_QTY", "SLS_QTY_WHSE1", "SLS_QTY_WHSE2", "SLS_QTY_WHSEX"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Width = 80
                    gcol.Format = "#,##0"
                End If
                If New String() {"FUT_QTY", "FUT_QTY_WHSE1", "FUT_QTY_WHSE2", "FUT_QTY_WHSEX"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.Width = 80
                    gcol.Format = "#,##0"
                End If
                If New String() {"SLS_CUSTS", "SLS_CUSTS_WHSE1", "SLS_CUSTS_WHSE2", "SLS_CUSTS_WHSEX"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Olive
                    gcol.Width = 60
                    gcol.Format = "#,##0"
                End If
            Next
        End With

        With grdSATCSLSH.DisplayLayout.Bands("SATCSLSH")
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("STYLE_DESC").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"SLS_AMT", "SLS_AMT_WHSE1", "SLS_AMT_WHSE2", "SLS_AMT_WHSEX"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.Width = 80
                    gcol.Format = "#,##0"
                End If
                If New String() {"SLS_QTY", "SLS_QTY_WHSE1", "SLS_QTY_WHSE2", "SLS_QTY_WHSEX"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Width = 80
                    gcol.Format = "#,##0"
                End If
                If New String() {"FUT_QTY", "FUT_QTY_WHSE1", "FUT_QTY_WHSE2", "FUT_QTY_WHSEX"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.Width = 80
                    gcol.Format = "#,##0"
                End If
                If New String() {"SLS_CUSTS", "SLS_CUSTS_WHSE1", "SLS_CUSTS_WHSE2", "SLS_CUSTS_WHSEX"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Olive
                    gcol.Width = 60
                    gcol.Format = "#,##0"
                End If
            Next
        End With

        grdICTSTATA.DisplayLayout.Override.HeaderPlacement = UltraWinGrid.HeaderPlacement.FixedOnTop
        With grdICTSTATA.DisplayLayout.Bands(0)
            .Columns("OTS").CellAppearance.BackColor = System.Drawing.Color.Yellow
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ON_HAND", "PICK", "OTS", "TRAN", "ON_ORDER", "OPEN", "OTS_WIP"}
                .Columns(COLUMN_NAME).Width = 65
                .Columns(COLUMN_NAME).Format = "#,##0"
            Next
        End With

        With grdSOTINVHX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_CANC"}
                .Columns(COLUMN_NAME).Format = "#,##0"
            Next
        End With

        ASCMAIN1.Add_Value_List(grdSATCSLS1, "STYLE_STATUS")
        ASCMAIN1.Add_Value_List(grdSATCSLS1, "STYLE_COLOR_STATUS")

        ASCMAIN1.Add_Value_List(grdSATCSLSH, "STYLE_STATUS")

        Dim rowICTPARM1 As DataRow = LookUp("ICTPARM1", "Z")
        IMAGE_FOLDER_NAME = rowICTPARM1.Item("IC_PARM_STYLE_IMG_DIR") & ""

    End Sub

    Sub Print_Report()
        Call Print_Report_Begin()

        Dim SUBT As String = ""
        Dim RecordSelectionFormula As String = ""
        Generate_Report("SARCSLS1", "", SUBT, RecordSelectionFormula)

        Call Print_Report_End()
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If Absx1.txtFor("CUST_CODE").Text <> "" Then Validate_Code("CUST_CODE")
                If Absx1.txtFor("SREP_CODE").Text <> "" Then Validate_Code("SREP_CODE")


                If EMsg = "" Then
                    If Absx1.cmbFor("RYP0").Value & "" = "" Then
                        EMsg &= vbCr & "You must Specify a Starting Period"
                    End If
                    If Absx1.cmbFor("RYP1").Value & "" = "" Then
                        EMsg &= vbCr & "You must Specify an Ending Period"
                    End If

                    If EMsg = "" Then
                        RYP0 = Absx1.cmbFor("RYP0").Value
                        RYP1 = Absx1.cmbFor("RYP1").Value
                        Periods = ASCMAIN1.Period_Diff(RYP0, RYP1) + 1
                    End If

                    If Periods > 24 Or Periods < 1 Then
                        EMsg &= vbCr & "Periods must be in chronological order and not more than 24 months apart"
                    End If
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Print Report"
                Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Data Options").Visible = tf
                .Groups("Options").Visible = Not tf
                .Groups("Period Range").Visible = Not tf
            End With
        End If
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            UltraExplorerBar1.Groups("Style Image").Visible = True
        Else
            UltraExplorerBar1.Groups("Style Image").Visible = False
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = tf

        spl.Panel1Collapsed = ScreenMode And Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("SREP_CODE").Text = ""

        With grdSOTINVHX.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Hidden = (Absx1.txtFor("CUST_CODE").Text <> "")
            .Columns("CUST_NAME").Hidden = (Absx1.txtFor("CUST_CODE").Text <> "")
            .Columns("SREP_CODE").Hidden = (Absx1.txtFor("CUST_CODE").Text <> "")
        End With

        ASCMAIN1.TACMAIN1.loadGridLayout(Me, grdSATCSLSH)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SATCSLS1", "SATCSLSH", "SOTINVHX", "ICTSTATA", "ASTAUDT1"} ', "SATCSLS1_DTL", "SATCSLS2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        'chkShow0Sales.Checked = True
        chkOpenOrders.Checked = True

        ' Absx1.txtFor("CUST_CODE").Text = ""
        ' Absx1.txtFor("SREP_CODE").Text = ""

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Customer Sales Data")

        Save_Header_Fields(UltraGroupBox1)

        CUST_CODE = HFs("CUST_CODE")
        Create_SATCSLS1()
        Set_DataSource()
        ShowActive()

        Sort_grdColumns(grdSATCSLS1, "STYLE_CODE,COLOR_CODE")
        Setup_grdSOTINVHX()

        With grdSATCSLS1.DisplayLayout.Bands(0)
            For Each COL As String In New String() {"SLS_AMT", "SLS_QTY", "FUT_QTY", "SLS_CUSTS"}
                Dim PFX As String = IIf(COL = "SLS_AMT", "$", IIf(COL = "SLS_QTY", "#", IIf(COL = "FUT_QTY", "Fut", "#C")))
                .Columns(COL).Header.Caption = PFX & " Total"
                .Columns(COL & "_WHSE1").Header.Caption = PFX & " " & Absx1.txtFor("WHSE_CODE1").Text
                .Columns(COL & "_WHSE2").Header.Caption = PFX & " " & Absx1.txtFor("WHSE_CODE2").Text
                .Columns(COL & "_WHSEX").Header.Caption = PFX & " " & "Other"
                .Columns(COL & "_WHSE1").Hidden = (Absx1.txtFor("WHSE_CODE1").Text = "")
                .Columns(COL & "_WHSE2").Hidden = (Absx1.txtFor("WHSE_CODE2").Text = "")
                .Columns(COL & "_WHSEX").Hidden = (Absx1.txtFor("WHSE_CODE1").Text = "") And (Absx1.txtFor("WHSE_CODE2").Text = "")
            Next
        End With

        With grdSATCSLSH.DisplayLayout.Bands(0)
            For Each COL As String In New String() {"SLS_AMT", "SLS_QTY", "FUT_QTY", "SLS_CUSTS"}
                Dim PFX As String = IIf(COL = "SLS_AMT", "$", IIf(COL = "SLS_QTY", "#", IIf(COL = "FUT_QTY", "Fut", "#C")))
                .Columns(COL).Header.Caption = PFX & " Total"
                .Columns(COL & "_WHSE1").Header.Caption = PFX & " " & Absx1.txtFor("WHSE_CODE1").Text
                .Columns(COL & "_WHSE2").Header.Caption = PFX & " " & Absx1.txtFor("WHSE_CODE2").Text
                .Columns(COL & "_WHSEX").Header.Caption = PFX & " " & "Other"
                .Columns(COL & "_WHSE1").Hidden = (Absx1.txtFor("WHSE_CODE1").Text = "")
                .Columns(COL & "_WHSE2").Hidden = (Absx1.txtFor("WHSE_CODE2").Text = "")
                .Columns(COL & "_WHSEX").Hidden = (Absx1.txtFor("WHSE_CODE1").Text = "") And (Absx1.txtFor("WHSE_CODE2").Text = "")
            Next
        End With

        grdSATCSLS1.Text = "Sales & Inventory by Style Color, Showing Sales from " & Trim(ASCMAIN1.Get_Legend(RYP0)) & " thru " & Trim(ASCMAIN1.Get_Legend(RYP1)) _
            & IIf(Absx1.txtFor("CUST_CODE").Text <> "", ", Customer " & Absx1.txtFor("CUST_CODE").Text & ":" & Absx1.txtFor("CUST_NAME").Text, "") _
            & IIf(Absx1.txtFor("SREP_CODE").Text <> "", ", Sales Rep " & Absx1.txtFor("SREP_CODE").Text & ":" & Absx1.txtFor("SREP_NAME").Text, "")

        grdSATCSLSH.Text = "Sales & Inventory by Style Color, Showing Sales from " & Trim(ASCMAIN1.Get_Legend(RYP0)) & " thru " & Trim(ASCMAIN1.Get_Legend(RYP1)) _
            & IIf(Absx1.txtFor("CUST_CODE").Text <> "", ", Customer " & Absx1.txtFor("CUST_CODE").Text & ":" & Absx1.txtFor("CUST_NAME").Text, "") _
            & IIf(Absx1.txtFor("SREP_CODE").Text <> "", ", Sales Rep " & Absx1.txtFor("SREP_CODE").Text & ":" & Absx1.txtFor("SREP_NAME").Text, "")

        Setup_SATCSLS1()
        ShowActive()

        ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSATCSLS1, "SSB", "Show Filter", "Show GroupBox", "Discontinue Color", "DNR Color")
        Load_Popup_Menu(grdSATCSLSH, "SSBBBBBS", "Show Filter", "Show GroupBox", "Style Status Inquiry", "Style Master File", "Discontinue Item", "DNR Item", "Save Grid Layout", "Show All Attributes")
        Load_Popup_Menu(grdSOTINVHX, "SSBB", "Show Filter", "Show GroupBox", "Sales Order Inquiry", "Show Invoice")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSATCSLSH"
                    If grdSATCSLSH.Selected.Rows.Count = 1 Then
                        Dim STYLE_STATUS As String = grdSATCSLSH.Selected.Rows(0).Cells.Item("STYLE_STATUS").Value
                        If STYLE_STATUS <> "D" Then
                            e.Tool.ToolbarsManager.Tools("Discontinue Item").SharedProps.Visible = True
                        Else
                            e.Tool.ToolbarsManager.Tools("Discontinue Item").SharedProps.Visible = False
                        End If
                        If STYLE_STATUS <> "N" Then
                            e.Tool.ToolbarsManager.Tools("DNR Item").SharedProps.Visible = True
                        Else
                            e.Tool.ToolbarsManager.Tools("DNR Item").SharedProps.Visible = False
                        End If
                    Else
                        e.Tool.ToolbarsManager.Tools("Discontinue Item").SharedProps.Visible = False
                        e.Tool.ToolbarsManager.Tools("DNR Item").SharedProps.Visible = False
                    End If
                Case "grdSATCSLS1"
                    If grdSATCSLS1.Selected.Rows.Count = 1 Then
                        Dim STYLE_COLOR_STATUS As String = grdSATCSLS1.Selected.Rows(0).Cells.Item("STYLE_COLOR_STATUS").Text
                        If STYLE_COLOR_STATUS <> "D" Then
                            e.Tool.ToolbarsManager.Tools("Discontinue Color").SharedProps.Visible = True
                        Else
                            e.Tool.ToolbarsManager.Tools("Discontinue Color").SharedProps.Visible = False
                        End If
                        If STYLE_COLOR_STATUS <> "N" Then
                            e.Tool.ToolbarsManager.Tools("DNR Color").SharedProps.Visible = True
                        Else
                            e.Tool.ToolbarsManager.Tools("DNR Color").SharedProps.Visible = False
                        End If
                    Else
                        e.Tool.ToolbarsManager.Tools("Discontinue Color").SharedProps.Visible = False
                        e.Tool.ToolbarsManager.Tools("DNR Color").SharedProps.Visible = False
                    End If
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If


            Case "Style Master File"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    'Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICTSTYL1")
                    Dim keys As New Dictionary(Of String, Object)
                    keys.Add("STYLE_CODE", STYLE_CODE)
                    ' If ASCMAIN1.Running_in_VS Then Stop ' NOT WORKING
                    Context_Launch("View", keys, e.Tool.Key, "ICTSTYL1")
                End If
            Case "Show Invoice"
                Dim FILENAME As String = ""
                If grd.ActiveRow IsNot Nothing Then
                    If Not grd.ActiveRow.Selected Then
                        grd.Selected.Rows.Clear()
                        grd.ActiveRow.Selected = True
                    End If
                End If

                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.Selected Then
                    Exit Sub
                End If

                Dim INV_NO As String = grd.ActiveRow.Cells("INV_NO").Value & ""

                FILENAME = TAC.SOCMAIN1.Create_Invoice(Me, INV_NO)

                Show_Document(FILENAME)

            Case "Discontinue Item", "DNR Item"
                Dim NEW_STATUS As String = "D"
                Dim NEW_STATUS_DESC As String = "Discontinue"
                If e.Tool.Key = "DNR Item" Then
                    NEW_STATUS = "N"
                    NEW_STATUS_DESC = "DNR"
                End If
                Dim STYLE_CODE As String = grdSATCSLSH.Selected.Rows(0).Cells.Item("STYLE_CODE").Text
                Dim iResult As MsgBoxResult
                Dim iMSG As New System.Text.StringBuilder
                iMSG.AppendLine(String.Format("This Will {0} The Following Item", NEW_STATUS_DESC))
                iMSG.AppendLine(String.Format("And All Of It's Non-{0}ed Colors:", NEW_STATUS_DESC))
                iMSG.AppendLine(STYLE_CODE)
                iMSG.AppendLine("")
                iMSG.AppendLine("Is This Really What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, NEW_STATUS_DESC & " Item")
                If iResult = MsgBoxResult.Yes Then
                    If ASCMAIN1.Logical_Lock("ICTSTYL1", STYLE_CODE) Then
                        BeginTrans()
                        Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
                        SQLS.AppendLine(String.Format("UPDATE ICTSTYL1 SET STYLE_STATUS = '{0}' WHERE STYLE_CODE = '{1}'", NEW_STATUS, STYLE_CODE))
                        ASCMAIN1.sql = SQLS.ToString
                        ASCDATA1.ExecuteSQL()
                        Dim ORIG_STATUS As String = grdSATCSLSH.Selected.Rows(0).Cells.Item("STYLE_STATUS").Value
                        grdSATCSLSH.Selected.Rows(0).Cells.Item("STYLE_STATUS").Value = NEW_STATUS
                        Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
                        With rowASTAUDT1
                            .Item("TABLE_NAME") = "ICTSTYL1"
                            .Item("KEY_VALUE") = STYLE_CODE
                            .Item("COLUMN_NAME") = "STYLE_STATUS"
                            .Item("FM_MODE") = "E"
                            .Item("USER_ID") = ASCMAIN1.USER_ID
                            .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                            .Item("OLD_VALUE") = ORIG_STATUS
                            .Item("NEW_VALUE") = NEW_STATUS
                            .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                            .Item("SELECTION_NO") = Me.SELECTION_NO
                            .Item("XNO") = Me.XNO
                        End With
                        dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
                        Update_Record_TDA("ASTAUDT1")
                        For Each rowSATCSLS1 As DataRow In dst.Tables("SATCSLS1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                            If rowSATCSLS1.Item("STYLE_COLOR_STATUS") & "" <> NEW_STATUS Then
                                Dim ORIG_COLOR_STATUS As String = rowSATCSLS1.Item("STYLE_COLOR_STATUS") & ""
                                If NEW_STATUS = "N" And rowSATCSLS1.Item("STYLE_COLOR_STATUS") & "" = "D" Then
                                    'Don't DNR colors that are Discontinued
                                Else
                                    Dim COLOR_CODE As String = rowSATCSLS1.Item("COLOR_CODE").ToString
                                    SQLS.Length = 0
                                    SQLS.AppendLine(String.Format("UPDATE ICTSTYC1 SET STYLE_COLOR_STATUS = '{0}' WHERE STYLE_CODE = '{1}' AND COLOR_CODE = '{2}'", NEW_STATUS, STYLE_CODE, COLOR_CODE))
                                    ASCMAIN1.sql = SQLS.ToString
                                    ASCDATA1.ExecuteSQL()
                                    rowSATCSLS1.Item("STYLE_COLOR_STATUS") = NEW_STATUS

                                    Dim rowASTAUDTC As DataRow = dst.Tables("ASTAUDT1").NewRow
                                    With rowASTAUDTC
                                        .Item("TABLE_NAME") = "ICTSTYC1"
                                        .Item("KEY_VALUE") = STYLE_CODE & ":" & COLOR_CODE
                                        .Item("COLUMN_NAME") = "STYLE_COLOR_STATUS"
                                        .Item("FM_MODE") = "E"
                                        .Item("USER_ID") = ASCMAIN1.USER_ID
                                        .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                                        .Item("OLD_VALUE") = ORIG_COLOR_STATUS
                                        .Item("NEW_VALUE") = NEW_STATUS
                                        .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                                        .Item("SELECTION_NO") = Me.SELECTION_NO
                                        .Item("XNO") = Me.XNO
                                    End With
                                    dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDTC)
                                    Update_Record_TDA("ASTAUDT1")
                                End If
                            End If
                        Next
                        CommitTrans()
                        ASCMAIN1.MultiTask_Release(, , )
                    End If
                Else
                    MsgBox("Nothing Was Done.", vbOKOnly, NEW_STATUS_DESC & " Item")
                End If
            Case "Discontinue Color", "DNR Color"
                Dim NEW_STATUS As String = "D"
                Dim NEW_STATUS_DESC As String = "Discontinue"
                If e.Tool.Key = "DNR Color" Then
                    NEW_STATUS = "N"
                    NEW_STATUS_DESC = "DNR"
                End If
                Dim STYLE_CODE As String = grdSATCSLS1.Selected.Rows(0).Cells.Item("STYLE_CODE").Text
                Dim COLOR_CODE As String = grdSATCSLS1.Selected.Rows(0).Cells.Item("COLOR_CODE").Text

                Dim S As New Text.StringBuilder With {.Length = 0}
                S.AppendLine("SELECT COUNT(*)")
                S.AppendLine("FROM ICTSTYC1")
                S.AppendLine("WHERE STYLE_CODE = '" & STYLE_CODE & "'")
                S.AppendLine("AND COLOR_CODE <> '" & COLOR_CODE & "'")
                S.AppendLine("AND STYLE_COLOR_STATUS <> '" & NEW_STATUS & "'")
                ASCMAIN1.sql = S.ToString()
                Dim dCount As Int16 = Val(ASCDATA1.GetDataValue)
                Dim aMsg As New StringBuilder With {.Length = 0}
                If dCount = 0 Then
                    aMsg.AppendLine("It Will Also Set The Status of The Style")
                    aMsg.AppendLine("To: " & NEW_STATUS_DESC)
                    aMsg.AppendLine("")
                End If
                Dim iResult As MsgBoxResult
                Dim iMSG As New System.Text.StringBuilder
                iMSG.AppendLine(String.Format("This Will {0} The Following Color For Style {1}", NEW_STATUS_DESC, STYLE_CODE))
                iMSG.AppendLine("")
                iMSG.AppendLine(COLOR_CODE)
                iMSG.AppendLine("")
                iMSG.AppendLine(aMsg.ToString())
                iMSG.AppendLine("Is This Really What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, NEW_STATUS_DESC & " Item")
                If iResult = MsgBoxResult.Yes Then

                    If ASCMAIN1.Logical_Lock("ICTSTYL1", STYLE_CODE) Then
                        BeginTrans()
                        Dim SQLS As New StringBuilder() With {.Length = 0}
                        SQLS.AppendLine(String.Format("UPDATE ICTSTYC1 SET STYLE_COLOR_STATUS = '{0}' WHERE STYLE_CODE = '{1}' AND COLOR_CODE = '{2}'", NEW_STATUS, STYLE_CODE, COLOR_CODE))
                        ASCMAIN1.sql = SQLS.ToString
                        ASCDATA1.ExecuteSQL()
                        Dim ORIG_COLOR_STATUS As String = grdSATCSLS1.Selected.Rows(0).Cells.Item("STYLE_COLOR_STATUS").Text
                        grdSATCSLS1.Selected.Rows(0).Cells.Item("STYLE_COLOR_STATUS").Value = NEW_STATUS
                        Dim rowASTAUDTC As DataRow = dst.Tables("ASTAUDT1").NewRow
                        With rowASTAUDTC
                            .Item("TABLE_NAME") = "ICTSTYC1"
                            .Item("KEY_VALUE") = STYLE_CODE & ":" & COLOR_CODE
                            .Item("COLUMN_NAME") = "STYLE_COLOR_STATUS"
                            .Item("FM_MODE") = "E"
                            .Item("USER_ID") = ASCMAIN1.USER_ID
                            .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                            .Item("OLD_VALUE") = ORIG_COLOR_STATUS
                            .Item("NEW_VALUE") = NEW_STATUS
                            .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                            .Item("SELECTION_NO") = Me.SELECTION_NO
                            .Item("XNO") = Me.XNO
                        End With
                        dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDTC)
                        Update_Record_TDA("ASTAUDT1")

                        If dCount = 0 Then
                            SQLS.Length = 0
                            SQLS.AppendLine(String.Format("UPDATE ICTSTYL1 SET STYLE_STATUS = '{0}' WHERE STYLE_CODE = '{1}'", NEW_STATUS, STYLE_CODE))
                            ASCMAIN1.sql = SQLS.ToString
                            ASCDATA1.ExecuteSQL()
                            Dim ORIG_STATUS As String = grdSATCSLSH.Selected.Rows(0).Cells.Item("STYLE_STATUS").Value
                            grdSATCSLSH.ActiveRow.Cells.Item("STYLE_STATUS").Value = NEW_STATUS
                            Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
                            With rowASTAUDT1
                                .Item("TABLE_NAME") = "ICTSTYL1"
                                .Item("KEY_VALUE") = STYLE_CODE
                                .Item("COLUMN_NAME") = "STYLE_STATUS"
                                .Item("FM_MODE") = "E"
                                .Item("USER_ID") = ASCMAIN1.USER_ID
                                .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                                .Item("OLD_VALUE") = ORIG_STATUS
                                .Item("NEW_VALUE") = NEW_STATUS
                                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                                .Item("SELECTION_NO") = Me.SELECTION_NO
                                .Item("XNO") = Me.XNO
                            End With
                            dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
                        End If
                        CommitTrans()
                        ASCMAIN1.MultiTask_Release(, , )
                    End If
                End If
            Case "Save Grid Layout"
                ASCMAIN1.TACMAIN1.SaveGridLayout(Me, grdSATCSLSH)
            Case "Show All Attributes"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
                tlb_sbt = DirectCast(tlb.Tools("Show All Attributes"), UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    Me.Cursor = Cursors.WaitCursor
                    For Each rowSATCSLSH As DataRow In dst.Tables("SATCSLSH").Select("", "STYLE_CODE")
                        Dim STYLE_CODE As String = rowSATCSLSH.Item("STYLE_CODE").ToString.Replace("'", "")
                        'Dim THEME_CODE As String = rowSATCSLSH.Item("THEME_CODE").ToString & String.Empty
                        'If THEME_CODE.Length > 0 Then
                        '    rowSATCSLSH.Item("THEME_DESC") = GET_THEME_INFO(THEME_CODE, "THEME_DESC")
                        '    rowSATCSLSH.Item("SEASON_CODE") = GET_THEME_INFO(THEME_CODE, "SEASON_CODE")
                        'End If

                        Dim rowICTSTYL31 As DataRow = dst.Tables("ICTSTYL3").Select(String.Format("STYLE_CODE = '{0}' AND ATT_RANK = '1'", STYLE_CODE)).FirstOrDefault
                        If Not IsNothing(rowICTSTYL31) Then
                            rowSATCSLSH.Item("ATTR_CODE1") = rowICTSTYL31.Item("ATTR_CODE")
                        End If
                        Dim nextI As Integer = 2
                        For Each rowICTSTYL3 As DataRow In dst.Tables("ICTSTYL3").Select(String.Format("STYLE_CODE = '{0}' AND ATT_RANK <> '1'", STYLE_CODE), "ATTR_CODE")
                            If nextI > 5 Then Exit For
                            rowSATCSLSH.Item(String.Format("ATTR_CODE{0}", nextI)) = rowICTSTYL3.Item("ATTR_CODE")
                            nextI += 1
                        Next
                    Next
                    grdSATCSLSH.DisplayLayout.Bands(0).Columns("ATTR_CODE1").Hidden = False
                    grdSATCSLSH.DisplayLayout.Bands(0).Columns("ATTR_CODE2").Hidden = False
                    grdSATCSLSH.DisplayLayout.Bands(0).Columns("ATTR_CODE3").Hidden = False
                    grdSATCSLSH.DisplayLayout.Bands(0).Columns("ATTR_CODE4").Hidden = False
                    grdSATCSLSH.DisplayLayout.Bands(0).Columns("ATTR_CODE5").Hidden = False
                    grdSATCSLSH.UpdateData()
                Else
                    grdSATCSLSH.DisplayLayout.Bands(0).Columns("ATTR_CODE1").Hidden = True
                    grdSATCSLSH.DisplayLayout.Bands(0).Columns("ATTR_CODE2").Hidden = True
                    grdSATCSLSH.DisplayLayout.Bands(0).Columns("ATTR_CODE3").Hidden = True
                    grdSATCSLSH.DisplayLayout.Bands(0).Columns("ATTR_CODE4").Hidden = True
                    grdSATCSLSH.DisplayLayout.Bands(0).Columns("ATTR_CODE5").Hidden = True
                End If
                Me.Cursor = Cursors.Default
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Call Click_Command("Load", e)
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

#Region "Custom Methods"
    Sub Create_SATCSLS1()

        Dim SQL As New Text.StringBuilder With {.Length = 0}

        ASCMAIN1.sql = "Truncate Table " & SATCSLSX
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & SATCSLSX & " " _
            & Replace(Replace(Replace(Replace(Replace(Replace(Replace(MakesqlSATCSLSX(), _
                                    "'WHSE1'", "'" & Absx1.txtFor("WHSE_CODE1").Text & "'"), _
                                    "'WHSE2'", "'" & Absx1.txtFor("WHSE_CODE2").Text & "'"), _
                                    "BETWEEN 'Z' AND 'Z'", IIf(chkOpenOrders.Checked, "BETWEEN 'O' AND 'P'", "BETWEEN 'Z' AND 'Z'")), _
                                    "'YP1'", "'" & RYP0 & "'"), "'YP2'", "'" & RYP1 & "'"), _
                             "   and SOTINVH2.CUST_CODE = SOTINVH2.CUST_CODE", "" _
                 & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTINVH1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                 & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTINVH1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf), _
                             "   and SOTORDR1.CUST_CODE = SOTORDR1.CUST_CODE", "" _
                 & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTORDR1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                 & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTORDR1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf)
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "Truncate Table " & SATCSLS1
        ASCDATA1.ExecuteSQL()

        Dim DANA As String = chkOpenOrders.CheckedValue

        ASCMAIN1.sql = "Insert into " & SATCSLS1 & " " _
            & Replace(Replace(Replace(Replace(Replace(Replace(Replace(MakesqlSATCSLS1, _
                                    "'WHSE1'", "'" & Absx1.txtFor("WHSE_CODE1").Text & "'"), _
                                    "'WHSE2'", "'" & Absx1.txtFor("WHSE_CODE2").Text & "'"), _
                                    "BETWEEN 'Z' AND 'Z'", IIf(chkOpenOrders.Checked, "BETWEEN 'O' AND 'P'", "BETWEEN 'Z' AND 'Z'")), _
                                    "'YP1'", "'" & RYP0 & "'"), "'YP2'", "'" & RYP1 & "'"), _
                             "   and SOTINVH2.CUST_CODE = SOTINVH2.CUST_CODE", "" _
                 & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTINVH1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                 & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTINVH1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf), _
                             "   and SOTORDR1.CUST_CODE = SOTORDR1.CUST_CODE", "" _
                 & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTORDR1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                 & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTORDR1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf)

        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {Absx1.txtFor("WHSE_CODE1").Text, Absx1.txtFor("WHSE_CODE2").Text, RYP0, RYP1})
        ASCDATA1.ExecuteSQL()

        If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("SREP_CODE").Text = "" Then

            'ASCMAIN1.sql = "Insert into " & SATCSLS1 & " (STYLE_CODE,COLOR_CODE,STYLE_DESC,STYLE_STATUS,VEND_CODE,FACTORY_CODE,STYLE_UOM,STYLE_CLASS_CODE,CARTON_PACK_QTY,COLOR_DESC,PO_COST, STYLE_COLOR_STATUS) " _
            '    & " Select X.STYLE_CODE, X.COLOR_CODE" & vbCrLf _
            '    & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_STATUS, ICTSTYL1.VEND_CODE, ICTSTYL1.FACTORY_CODE, ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.CARTON_PACK_QTY" & vbCrLf _
            '    & ", ICTCOLR1.COLOR_DESC, CASE WHEN NVL(NEW_PO_COST_DATE,TRUNC(SYSDATE+1)) <= TRUNC(SYSDATE) THEN NEW_PO_COST ELSE PO_COST END PO_COST, X.STYLE_COLOR_STATUS" & vbCrLf _
            '    & " from ICTSTYL1,ICTCOLR1,ICTSTYV1" & vbCrLf _
            '    & ", (Select ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTSTYC1.STYLE_COLOR_STATUS from ICTSTYC1,ICTSTYL1" & vbCrLf _
            '    & "     where ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE and ICTSTYL1.STYLE_STATUS = 'A'" & vbCrLf _
            '    & "   minus Select STYLE_CODE, COLOR_CODE, STYLE_COLOR_STATUS from " & SATCSLS1 & ") X" & vbCrLf _
            '    & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
            '    & "   and ICTCOLR1.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
            '    & "   and ICTSTYV1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
            '    & "   and ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE"

            'Fix You Are looking for issues with Styles that have no sales in the sales grid this is the SQL you are looking for.
            Dim S As New StringBuilder With {.Length = 0}
            S.AppendLine(String.Format("Insert into {0} (STYLE_CODE,COLOR_CODE,STYLE_DESC,STYLE_STATUS,VEND_CODE,FACTORY_CODE,STYLE_UOM,STYLE_CLASS_CODE,CARTON_PACK_QTY,COLOR_DESC,PO_COST, STYLE_COLOR_STATUS, FUT_QTY_WHSE1, FUT_QTY_WHSE2, FUT_QTY)", SATCSLS1))
            S.AppendLine("Select X.STYLE_CODE, X.COLOR_CODE")
            S.AppendLine(", ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_STATUS, ICTSTYL1.VEND_CODE, ICTSTYL1.FACTORY_CODE, ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.CARTON_PACK_QTY")
            S.AppendLine(", ICTCOLR1.COLOR_DESC, CASE WHEN NVL(NEW_PO_COST_DATE,TRUNC(SYSDATE+1)) <= TRUNC(SYSDATE) THEN NEW_PO_COST ELSE PO_COST END PO_COST, X.STYLE_COLOR_STATUS")
            S.AppendLine(", Z.FUT_QTY_WHSE1 ,Z.FUT_QTY_WHSE2 ,FUT_QTY_TOT")
            S.AppendLine(" from ICTSTYL1,ICTCOLR1,ICTSTYV1")
            S.AppendLine(", (Select ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTSTYC1.STYLE_COLOR_STATUS from ICTSTYC1,ICTSTYL1")
            S.AppendLine("     where ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE and ICTSTYL1.STYLE_STATUS = 'A'")
            S.AppendLine(String.Format("   minus Select STYLE_CODE, COLOR_CODE, STYLE_COLOR_STATUS from {0}) X", SATCSLS1))
            S.AppendLine(", (SELECT")
            S.AppendLine("STYLE_CODE, COLOR_CODE,")
            S.AppendLine(String.Format("SUM(DECODE(WHSE_CODE,'{0}',(NVL(WHSE_QTY_ON_HAND,0) + (NVL(WHSE_QTY_ON_ORDER,0) - NVL(WHSE_QTY_TRAN,0)) - (NVL(WHSE_QTY_OPEN,0) + NVL(WHSE_QTY_PICK,0))),0)) AS FUT_QTY_WHSE1,", Absx1.txtFor("WHSE_CODE1").Text))
            S.AppendLine(String.Format("SUM(DECODE(WHSE_CODE,'{0}',(NVL(WHSE_QTY_ON_HAND,0) + (NVL(WHSE_QTY_ON_ORDER,0) - NVL(WHSE_QTY_TRAN,0)) - (NVL(WHSE_QTY_OPEN,0) + NVL(WHSE_QTY_PICK,0))),0)) AS FUT_QTY_WHSE2,", Absx1.txtFor("WHSE_CODE2").Text))
            S.AppendLine("SUM(NVL(WHSE_QTY_ON_HAND,0) + (NVL(WHSE_QTY_ON_ORDER,0) - NVL(WHSE_QTY_TRAN,0)) - (NVL(WHSE_QTY_OPEN,0) + NVL(WHSE_QTY_PICK,0))) as FUT_QTY_TOT")
            S.AppendLine("FROM ICTSTAT2")
            S.AppendLine("GROUP BY STYLE_CODE,")
            S.AppendLine("COLOR_CODE) Z")
            S.AppendLine(" where ICTSTYL1.STYLE_CODE = X.STYLE_CODE")
            S.AppendLine("   and ICTCOLR1.COLOR_CODE = X.COLOR_CODE")
            S.AppendLine("AND ICTSTYL1.STYLE_CODE = Z.STYLE_CODE AND ICTCOLR1.COLOR_CODE = Z.COLOR_CODE")
            S.AppendLine("   and ICTSTYV1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE")
            S.AppendLine("   and ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE")
            ASCMAIN1.sql = S.ToString()
            ASCDATA1.ExecuteSQL()

            'chkShow0Sales.Visible = True
        Else
            'chkShow0Sales.Visible = False
        End If
        Fill_Records("SATCSLS1")
        For Each rowSATCSLS1 As DataRow In dst.Tables("SATCSLS1").Select()
            Dim THEME_CODE As String = rowSATCSLS1.Item("THEME_CODE").ToString & String.Empty
            If THEME_CODE.Length > 0 Then
                rowSATCSLS1.Item("THEME_DESC") = GET_THEME_INFO(THEME_CODE, "THEME_DESC")
                rowSATCSLS1.Item("SEASON_CODE") = GET_THEME_INFO(THEME_CODE, "SEASON_CODE")
            End If
        Next

        ASCMAIN1.sql = "Truncate Table " & SATCSLSH
        ASCDATA1.ExecuteSQL()
        With SQL
            .Length = 0
            .AppendLine(String.Format("Insert into {0} ", SATCSLSH))
            .AppendLine("SELECT")
            .AppendLine("STYLE_CODE,")
            .AppendLine("STYLE_DESC,")
            .AppendLine("STYLE_STATUS,")
            .AppendLine("VEND_CODE,")
            .AppendLine("FACTORY_CODE,")
            .AppendLine("STYLE_UOM,")
            .AppendLine("STYLE_CLASS_CODE,")
            .AppendLine("MIN(CARTON_PACK_QTY),")
            .AppendLine("MAX(PO_COST),")
            .AppendLine("SUM(SLS_AMT),")
            .AppendLine("SUM(SLS_AMT_WHSE1),")
            .AppendLine("SUM(SLS_AMT_WHSE2),")
            .AppendLine("SUM(SLS_QTY),")
            .AppendLine("SUM(SLS_QTY_WHSE1),")
            .AppendLine("SUM(SLS_QTY_WHSE2),")
            .AppendLine("SUM(SLS_CUSTS),")
            .AppendLine("SUM(SLS_CUSTS_WHSE1),")
            .AppendLine("SUM(SLS_CUSTS_WHSE2),")
            .AppendLine("SUM(SLS_CUSTS_WHSEX),")
            .AppendLine("SUM(FUT_QTY),")
            .AppendLine("SUM(FUT_QTY_WHSE1),")
            .AppendLine("SUM(FUT_QTY_WHSE2),")
            .AppendLine("MAX(CUST_CODE),")
            .AppendLine("MAX(CUST_NAME)")
            .AppendLine("FROM " & SATCSLS1)
            .AppendLine("GROUP BY")
            .AppendLine("STYLE_CODE,")
            .AppendLine("STYLE_DESC,")
            .AppendLine("STYLE_STATUS,")
            .AppendLine("VEND_CODE,")
            .AppendLine("FACTORY_CODE,")
            .AppendLine("STYLE_UOM,")
            .AppendLine("STYLE_CLASS_CODE")
        End With
        ASCMAIN1.sql = SQL.ToString
        ASCDATA1.ExecuteSQL()

        If chkShow0Sales.Checked Then
            With SQL
                .Length = 0
                .AppendLine(String.Format("Insert into {0} ", SATCSLSH))
                .AppendLine("SELECT")
                .AppendLine("S1.STYLE_CODE,")
                .AppendLine("S1.STYLE_DESC,")
                .AppendLine("S1.STYLE_STATUS,")
                .AppendLine("S1.VEND_CODE,")
                .AppendLine("S1.FACTORY_CODE,")
                .AppendLine("S1.STYLE_UOM,")
                .AppendLine("S1.STYLE_CLASS_CODE,")
                .AppendLine("MIN(S1.CARTON_PACK_QTY),")
                .AppendLine("MAX(CASE WHEN NVL(NEW_PO_COST_DATE,TRUNC(SYSDATE+1)) <= TRUNC(SYSDATE) THEN NEW_PO_COST ELSE PO_COST END) AS PO_COST,")
                .AppendLine("0 AS SLS_AMT,")
                .AppendLine("0 AS SLS_AMT_WHSE1,")
                .AppendLine("0 AS SLS_AMT_WHSE2,")
                .AppendLine("0 AS SLS_QTY,")
                .AppendLine("0 AS SLS_QTY_WHSE1,")
                .AppendLine("0 AS SLS_QTY_WHSE2,")
                .AppendLine("0 AS SLS_CUSTS,")
                .AppendLine("0 AS SLS_CUSTS_WHSE1,")
                .AppendLine("0 AS SLS_CUSTS_WHSE2,")
                .AppendLine("0 AS SLS_CUSTS_WHSEX,")
                .AppendLine("SUM (NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0)) FUT_QTY,")
                .AppendLine("SUM (DECODE(S2.WHSE_CODE,'" & Absx1.txtFor("WHSE_CODE1").Text & "',NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0),0)) AS FUT_QTY_WHSE1,")
                .AppendLine("SUM (DECODE(S2.WHSE_CODE,'" & Absx1.txtFor("WHSE_CODE2").Text & "',NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0),0)) AS FUT_QTY_WHSE2,")
                .AppendLine("S1.CUST_CODE,")
                .AppendLine("C1.CUST_NAME")
                .AppendLine("FROM ICTSTYL1 S1, ICTSTAT2 S2, ICTSTYV1 V1, ARTCUST1 C1")
                .AppendLine("WHERE S1.STYLE_CODE = S2.STYLE_CODE (+)")
                .AppendLine("AND S1.STYLE_CODE = V1.STYLE_CODE (+)")
                .AppendLine("AND S1.CUST_CODE = C1.CUST_CODE (+)")
                .AppendLine("AND S1.STYLE_CODE NOT IN (")
                .AppendLine("SELECT STYLE_CODE")
                .AppendLine(String.Format("FROM {0})", SATCSLSH))
                .AppendLine("GROUP BY")
                .AppendLine("S1.STYLE_CODE,")
                .AppendLine("S1.STYLE_DESC,")
                .AppendLine("S1.STYLE_STATUS,")
                .AppendLine("S1.VEND_CODE,")
                .AppendLine("S1.FACTORY_CODE,")
                .AppendLine("S1.STYLE_UOM,")
                .AppendLine("S1.STYLE_CLASS_CODE,")
                .AppendLine("S1.CUST_CODE,")
                .AppendLine("C1.CUST_NAME")
            End With
            ASCMAIN1.sql = SQL.ToString
            ASCDATA1.ExecuteSQL()
        End If

        Fill_Records("SATCSLSH")

        SQL.Length = 0
        With SQL
            .AppendLine("SELECT")
            .AppendLine("ST.STYLE_CODE,")
            .AppendLine("ST.COLOR_CODE,")
            .AppendLine("CL.COLOR_DESC,")
            .AppendLine("SC.UPC_CODE,")
            .AppendLine("TH.THEME_DESC,")
            .AppendLine("TH.SEASON_CODE,")
            .AppendLine("SUM(NVL(WHSE_QTY_ON_HAND,0)) AS ON_HAND,")
            .AppendLine("SUM(NVL(WHSE_QTY_PICK,0)) AS PICK,")
            .AppendLine("(SUM(NVL(WHSE_QTY_ON_HAND,0)) - SUM(NVL(WHSE_QTY_PICK,0))) AS OTS,")
            .AppendLine("SUM(NVL(WHSE_QTY_TRAN,0)) AS TRAN,")
            .AppendLine("SUM(NVL(WHSE_QTY_ON_ORDER,0)) AS ON_ORDER,")
            .AppendLine("SUM(NVL(WHSE_QTY_OPEN,0)) AS OPEN,")
            .AppendLine("((SUM(NVL(WHSE_QTY_ON_HAND,0)) - SUM(NVL(WHSE_QTY_PICK,0))) + SUM(NVL(WHSE_QTY_TRAN,0)) + SUM(NVL(WHSE_QTY_ON_ORDER,0)) - SUM(NVL(WHSE_QTY_OPEN,0))) AS OTS_WIP")
            .AppendLine("FROM ICTSTAT2 ST, ICTSTYC1 SC, ICTCOLR1 CL, ICTTHEME TH")
            .AppendLine("WHERE ST.STYLE_CODE = SC.STYLE_CODE")
            .AppendLine("AND ST.COLOR_CODE = SC.COLOR_CODE")
            .AppendLine("AND ST.COLOR_CODE = CL.COLOR_CODE")
            .AppendLine("AND SC.THEME_CODE = TH.THEME_CODE (+)")
            If chkInvMSOnly.Checked Then
                .AppendLine("AND ST.WHSE_CODE = 'MS'")
            End If
            .AppendLine("GROUP BY ST.STYLE_CODE,")
            .AppendLine("ST.COLOR_CODE,")
            .AppendLine("CL.COLOR_DESC,")
            .AppendLine("SC.UPC_CODE,")
            .AppendLine("TH.THEME_DESC,")
            .AppendLine("TH.SEASON_CODE")
            .AppendLine("ORDER BY ST.STYLE_CODE,")
            .AppendLine("ST.COLOR_CODE")
        End With
        Fill_Records("ICTSTATA", , True, SQL.ToString)

    End Sub

    Public Shared Function Get_Image( _
    ByVal IMAGE_FOLDER As String, _
    ByVal IMAGE_FILE As String) As String

        Dim RetVal As String

        Dim image_file_found As Boolean = True

        If IMAGE_FILE = "\.jpg" Then
            image_file_found = False
            RetVal = ""
            Return RetVal
        End If

        If Not IMAGE_FOLDER.EndsWith("\") Then IMAGE_FOLDER &= "\"
        Dim IMAGE_FILENAME As String = IMAGE_FOLDER & IMAGE_FILE
        Try
            If My.Computer.FileSystem.FileExists(IMAGE_FILENAME) Then
                RetVal = IMAGE_FILENAME
            ElseIf My.Computer.FileSystem.FileExists(IMAGE_FILENAME & ".PNG") Then
                RetVal = IMAGE_FILENAME & ".PNG"
            ElseIf My.Computer.FileSystem.FileExists(IMAGE_FILENAME & ".JPG") Then
                RetVal = IMAGE_FILENAME & ".JPG"
            Else
                image_file_found = False
                RetVal = ""
            End If
        Catch ex As Exception
            image_file_found = False
            RetVal = ""
        End Try

        Return RetVal
    End Function

    Private Sub picMaster_DoubleClick(sender As Object, e As System.EventArgs) Handles picMaster.DoubleClick
        Dim frmSOFIMGV1 As New WBFIMGV1(Me, picMaster.ImageLocation)
        frmSOFIMGV1.Show()
    End Sub

    Sub Set_DataSource()
        If Not IsNothing(grdSATCSLS1.DataSource) Then
            Dim dvw As DataView = DirectCast(grdSATCSLS1.DataSource, DataTable).DefaultView
            If chkShow0Sales.Checked Then
                dvw.RowFilter = ""
            Else
                dvw.RowFilter = "SLS_QTY <> 0"
            End If
        End If
    End Sub

    Sub Setup_grdSOTINVHX()
        If grdSATCSLS1.ActiveRow Is Nothing OrElse Not grdSATCSLS1.ActiveRow.IsDataRow Then
            grdSOTINVHX.Visible = False
            grdICTSTATA.Visible = False
            picMaster.Image = Nothing
            picMaster.Visible = False
        Else

            ASCMAIN1.Progress("Now Retrieving Sales Documents")
            Me.Cursor = Cursors.WaitCursor

            grdSOTINVHX.Visible = True
            grdICTSTATA.Visible = True

            Dim STYLE_CODE As String = grdSATCSLS1.ActiveRow.Cells("STYLE_CODE").Value & ""
            Dim COLOR_CODE As String = grdSATCSLS1.ActiveRow.Cells("COLOR_CODE").Value & ""

            ASCMAIN1.sql = MakesqlSOTINVHX(STYLE_CODE, COLOR_CODE, RYP0, RYP1)
            Fill_Records("SOTINVHX", "", , ASCMAIN1.sql)
            Fill_Extra_XRecords()

            If chkOpenOrders.Checked Then
                ASCMAIN1.sql = "Select SOTORDR2.ORDR_STATUS INV_TYPE, SOTORDR2.ORDR_NO INV_NO" & vbCrLf _
                    & ", SOTORDR1.ORDR_SHIP_DATE INV_DATE, SOTORDR1.ORDR_TYPE_CODE, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                    & ", '000000' OPS_YYYYPP, SOTORDR1.ORDR_NO" & vbCrLf _
                    & ", SOTORDR1.SREP_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE, ARTCUST1.CUST_NAME, SOTORDR1.CUST_STORE_NO" & vbCrLf _
                    & ", NVL(ARTCUST2.CUST_ADDR_NAME,ARTCUST2.CUST_NAME) CUST_STORE_LOCATION" & vbCrLf _
                    & ", SOTORDR2.STYLE_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
                    & ", NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) ORDR_QTY_SHIP, SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
                    & ", (NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" & vbCrLf _
                    & ", ORDR_QTY_OPEN" & vbCrLf _
                    & ", ORDR_QTY_PICK" & vbCrLf _
                    & ", ORDR_QTY_CANC" & vbCrLf _
                    & ", SOTORDR1.ORDR_DATE_RECD" & vbCrLf _
                    & " from SOTORDR2,ICTSTYL1,ARTCUST2,SOTORDR1,ARTCUST1 " & vbCrLf _
                    & " where ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE " & vbCrLf _
                    & " and ARTCUST2.CUST_CODE (+) = SOTORDR1.CUST_CODE " & vbCrLf _
                    & " and ARTCUST2.CUST_ADDR_TYPE (+) = 'MK' " & vbCrLf _
                    & " and ARTCUST2.CUST_ADDR_CODE (+) = SOTORDR1.CUST_STORE_NO " & vbCrLf _
                    & " and SOTORDR2.ORDR_STATUS BETWEEN 'O' AND 'P'" & vbCrLf _
                    & " and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                    & " and SOTORDR1.ORDR_YYYYPP_BOOKED between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf _
                    & " and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                    & " and SOTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                    & " and SOTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                    & " and SOTORDR1.ORDR_YYYYPP_BOOKED between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf _
                    & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTORDR1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                    & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTORDR1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf

                Fill_Records("SOTINVHX", "", False, ASCMAIN1.sql)
                Fill_Extra_XRecords()
            End If


            grdSOTINVHX.Text = "Sales Documents for " & STYLE_CODE & "-" & COLOR_CODE


            Dim IMAGE_NAME As String = String.Format("{0}-{1}", STYLE_CODE, COLOR_CODE)

            Dim imgba() As Byte = Nothing
            If IMAGE_NAME <> "" Then
                picMaster.ImageLocation = Get_Image(IMAGE_FOLDER_NAME, IMAGE_NAME)
                If picMaster.ImageLocation.ToString.Length > 0 Then
                    picMaster.Visible = True
                Else
                    picMaster.Visible = False
                End If
            Else
                picMaster.Image = Nothing
                picMaster.Visible = False
            End If

            ASCMAIN1.Progress("")
            Me.Cursor = Cursors.Default
        End If

    End Sub

    Private Sub Fill_Extra_XRecords()
        For Each rowSOTINVHX As DataRow In dst.Tables("SOTINVHX").Select()
            Dim ORDR_NO As String = rowSOTINVHX.Item("ORDR_NO") & ""
            Dim STYLE_CODE As String = rowSOTINVHX.Item("STYLE_CODE")
            Dim sql As New Text.StringBuilder With {.Length = 0}
            sql.AppendLine("SELECT")
            sql.AppendLine("ORDR_DATE_RECD,")
            sql.AppendLine("SUM(ORDR_QTY_OPEN) ORDR_QTY_OPEN,")
            sql.AppendLine("SUM(ORDR_QTY_PICK) ORDR_QTY_PICK,")
            sql.AppendLine("SUM(ORDR_QTY_CANC) ORDR_QTY_CANC")
            sql.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2")
            sql.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
            sql.AppendLine(String.Format("AND O1.ORDR_NO = '{0}'", ORDR_NO))
            sql.AppendLine(String.Format("AND O2.STYLE_CODE = '{0}'", STYLE_CODE))
            sql.AppendLine("GROUP BY ORDR_DATE_RECD")
            Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
            If tbl.Rows.Count = 1 Then
                rowSOTINVHX.Item("ORDR_DATE_RECD") = tbl.Rows(0).Item("ORDR_DATE_RECD")
                rowSOTINVHX.Item("ORDR_QTY_OPEN") = tbl.Rows(0).Item("ORDR_QTY_OPEN")
                rowSOTINVHX.Item("ORDR_QTY_PICK") = tbl.Rows(0).Item("ORDR_QTY_PICK")
                rowSOTINVHX.Item("ORDR_QTY_CANC") = tbl.Rows(0).Item("ORDR_QTY_CANC")
            Else
                rowSOTINVHX.Item("ORDR_DATE_RECD") = Null
                rowSOTINVHX.Item("ORDR_QTY_OPEN") = 0
                rowSOTINVHX.Item("ORDR_QTY_PICK") = 0
                rowSOTINVHX.Item("ORDR_QTY_CANC") = 0
            End If
        Next
        grdSOTINVHX.UpdateData()
    End Sub

    Sub Setup_SATCSLS1()
        Dim STYLE_CODE As String = ""
        Dim dvw As DataView = DirectCast(grdSATCSLS1.DataSource, DataTable).DefaultView
        If grdSATCSLSH.ActiveRow Is Nothing OrElse (Not grdSATCSLSH.ActiveRow.IsDataRow Or grdSATCSLSH.ActiveRow.IsAddRow) Then
            STYLE_CODE = "NONE"
        Else
            STYLE_CODE = grdSATCSLSH.ActiveRow.Cells("STYLE_CODE").Value
        End If
        dvw.RowFilter = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        grdSATCSLS1.Text = "Sales & Inventory For Style " & STYLE_CODE

        Dim dvws As DataView = DirectCast(grdICTSTATA.DataSource, DataTable).DefaultView
        dvws.RowFilter = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        grdICTSTATA.Text = "Status For Style " & STYLE_CODE
        picMaster.Image = Nothing
        picMaster.Visible = False
    End Sub

    Sub Setup_tabDetails()
        If SELECTION_NO = 0 Then Exit Sub

    End Sub

    Private Sub ShowActive()
        If Not IsNothing(grdSATCSLSH.DataSource) Then
            Dim dvw As DataView = DirectCast(grdSATCSLSH.DataSource, DataTable).DefaultView
            Dim Filter As String = ""
            If chkShowActive.Checked Then
                Filter = "STYLE_STATUS = 'A'"
            Else
                Filter = ""
            End If
            dvw.RowFilter = Filter
        End If
    End Sub
#End Region

#Region "Form Controls"

    Private Sub chkShow0Sales_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShow0Sales.CheckedChanged
        If EntryMode = "E" Then
            ASCMAIN1.Progress("Now Retrieving Sales Documents")
            Me.Cursor = Cursors.WaitCursor
            Create_SATCSLS1()
            Set_DataSource()
            ShowActive()
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub chkShowActive_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowActive.CheckedChanged
        ShowActive()
    End Sub

    Private Sub chkShowDetails_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowDetails.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        SplitContainer1.Panel2Collapsed = Not chkShowDetails.Checked
    End Sub

    Private Sub grdSATCSLS1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSATCSLS1.AfterRowActivate
        Setup_grdSOTINVHX()
        If Not grdSATCSLS1.ActiveRow Is Nothing And grdSATCSLS1.ActiveRow.IsDataRow Then
            Dim STYLE_CODE_IMG As String = grdSATCSLS1.ActiveRow.Cells.Item("STYLE_CODE").Text & String.Empty
            Dim COLOR_CODE_IMG As String = grdSATCSLS1.ActiveRow.Cells.Item("COLOR_CODE").Text & String.Empty
            If STYLE_CODE_IMG.Length > 0 And COLOR_CODE_IMG.Length > 0 Then
                FetchImage(STYLE_CODE_IMG, COLOR_CODE_IMG)
            End If
        End If
    End Sub

    Private Sub grdSATCSLS1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATCSLS1.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("STYLE_STATUS").Value & "" <> e.Row.Cells("STYLE_COLOR_STATUS").Value & "" Then
                e.Row.Cells("STYLE_COLOR_STATUS").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("STYLE_COLOR_STATUS").ToolTipText = "Color Status is not in agreement with Style Status"
            End If
        End If
    End Sub

    Private Sub grdSATCSLSH_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSATCSLSH.AfterRowActivate
        Setup_SATCSLS1()
        EcomIndicator()
    End Sub

    Private Sub tabDetails_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs)
        Setup_tabDetails()
    End Sub
#End Region

    Private Sub chkOpenOrders_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkOpenOrders.CheckedChanged
        If chkOpenOrders.Checked Then
            chkCreditsOnly.Checked = False
        End If
    End Sub

    Private Sub chkCreditsOnly_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkCreditsOnly.CheckedChanged
        If chkCreditsOnly.Checked Then
            chkOpenOrders.Checked = False
            chkShow0Sales.Checked = False
            chkShow0Sales.Visible = False
        Else
            chkShow0Sales.Visible = True
        End If
    End Sub

    Private Function MakesqlSATCSLSX() As String
        Dim Retval As System.Text.StringBuilder = New System.Text.StringBuilder() With {.Length = 0}
        If chkCreditsOnly.Checked Then
            Retval.AppendLine("Select STYLE_CODE, COLOR_CODE, CUST_CODE")
            Retval.AppendLine(", Sum (SLS_AMT) SLS_AMT, Sum (SLS_AMT_WHSE1) SLS_AMT_WHSE1, Sum (SLS_AMT_WHSE2) SLS_AMT_WHSE2, Sum (SLS_AMT_WHSEX) SLS_AMT_WHSEX")
            Retval.AppendLine(", Sum (SLS_QTY) SLS_QTY, Sum (SLS_QTY_WHSE1) SLS_QTY_WHSE1, Sum (SLS_QTY_WHSE2) SLS_QTY_WHSE2, Sum (SLS_QTY_WHSEX) SLS_QTY_WHSEX")
            Retval.AppendLine(" from (")
            Retval.AppendLine(" Select SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH1.CUST_CODE")
            Retval.AppendLine(", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)) SLS_AMT")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE1")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE2")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',0,'WHSE2',0,NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0))) SLS_AMT_WHSEX")
            Retval.AppendLine(", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) SLS_QTY")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE1")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE2")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',0,'WHSE2',0,NVL(SOTINVH2.ORDR_QTY_SHIP,0))) SLS_QTY_WHSEX")
            Retval.AppendLine(" from SOTINVH2,SOTINVH1")
            Retval.AppendLine(" where SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN 'YP1' AND 'YP2'")
            Retval.AppendLine("   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
            Retval.AppendLine("   and SOTINVH2.ORDR_QTY_SHIP < 0")
            Retval.AppendLine("   and SOTINVH2.INV_TYPE = 'C'")
            Retval.AppendLine("   and SOTINVH2.CUST_CODE = SOTINVH2.CUST_CODE")
            Retval.AppendLine(" group by SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH1.CUST_CODE")
            Retval.AppendLine(") group by STYLE_CODE, COLOR_CODE, CUST_CODE")
        Else
            Retval.AppendLine("Select STYLE_CODE, COLOR_CODE, CUST_CODE")
            Retval.AppendLine(", Sum (SLS_AMT) SLS_AMT, Sum (SLS_AMT_WHSE1) SLS_AMT_WHSE1, Sum (SLS_AMT_WHSE2) SLS_AMT_WHSE2, Sum (SLS_AMT_WHSEX) SLS_AMT_WHSEX")
            Retval.AppendLine(", Sum (SLS_QTY) SLS_QTY, Sum (SLS_QTY_WHSE1) SLS_QTY_WHSE1, Sum (SLS_QTY_WHSE2) SLS_QTY_WHSE2, Sum (SLS_QTY_WHSEX) SLS_QTY_WHSEX")
            Retval.AppendLine(" from (")
            Retval.AppendLine("Select SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH1.CUST_CODE")
            Retval.AppendLine(", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)) SLS_AMT")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE1")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE2")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',0,'WHSE2',0,NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0))) SLS_AMT_WHSEX")
            Retval.AppendLine(", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) SLS_QTY")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE1")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE2")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',0,'WHSE2',0,NVL(SOTINVH2.ORDR_QTY_SHIP,0))) SLS_QTY_WHSEX")
            Retval.AppendLine(" from SOTINVH2,SOTINVH1")
            Retval.AppendLine(" where SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN 'YP1' AND 'YP2'")
            Retval.AppendLine("   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
            Retval.AppendLine("   and SOTINVH2.ORDR_QTY_SHIP <> 0")
            Retval.AppendLine("   and SOTINVH2.INV_TYPE = 'I'")
            Retval.AppendLine("   and SOTINVH2.CUST_CODE = SOTINVH2.CUST_CODE")
            Retval.AppendLine(" group by SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH1.CUST_CODE")
            If chkIncludeCredits.Checked Then
                Retval.AppendLine(" union ")
                Retval.AppendLine(" Select SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH1.CUST_CODE")
                Retval.AppendLine(", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)) SLS_AMT")
                Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE1")
                Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE2")
                Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',0,'WHSE2',0,NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0))) SLS_AMT_WHSEX")
                Retval.AppendLine(", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) SLS_QTY")
                Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE1")
                Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE2")
                Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',0,'WHSE2',0,NVL(SOTINVH2.ORDR_QTY_SHIP,0))) SLS_QTY_WHSEX")
                Retval.AppendLine(" from SOTINVH2,SOTINVH1")
                Retval.AppendLine(" where SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN 'YP1' AND 'YP2'")
                Retval.AppendLine("   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
                Retval.AppendLine("   and SOTINVH2.ORDR_QTY_SHIP < 0")
                Retval.AppendLine("   and SOTINVH2.INV_TYPE = 'C'")
                Retval.AppendLine("   and SOTINVH2.CUST_CODE = SOTINVH2.CUST_CODE")
                Retval.AppendLine(" group by SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH1.CUST_CODE")
            End If
            Retval.AppendLine(" union ")
            Retval.AppendLine("Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR1.CUST_CODE")
            Retval.AppendLine(", SUM ((NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) SLS_AMT")
            Retval.AppendLine(", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE1',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE1")
            Retval.AppendLine(", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE2',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE2")
            Retval.AppendLine(", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE1',0,'WHSE2',0,(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0))) SLS_AMT_WHSEX")
            Retval.AppendLine(", SUM ((NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0))) SLS_QTY")
            Retval.AppendLine(", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE1',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)),0)) SLS_QTY_WHSE1")
            Retval.AppendLine(", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE2',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)),0)) SLS_QTY_WHSE2")
            Retval.AppendLine(", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE1',0,'WHSE2',0,(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)))) SLS_QTY_WHSEX")
            Retval.AppendLine(" from SOTORDR2,SOTORDR1")
            Retval.AppendLine(" where SOTORDR2.ORDR_STATUS BETWEEN 'Z' AND 'Z'")
            Retval.AppendLine(" and SOTORDR1.ORDR_YYYYPP_BOOKED BETWEEN 'YP1' AND 'YP2'")
            Retval.AppendLine("   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO")
            Retval.AppendLine("   and (NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) <> 0")
            Retval.AppendLine("   and SOTORDR1.CUST_CODE = SOTORDR1.CUST_CODE")
            Retval.AppendLine(" group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR1.CUST_CODE")
            Retval.AppendLine(") group by STYLE_CODE, COLOR_CODE, CUST_CODE")
        End If
        Return Retval.ToString
    End Function

    Private Function MakesqlSATCSLS1() As String
        Dim Retval As System.Text.StringBuilder = New System.Text.StringBuilder() With {.Length = 0}
        If chkCreditsOnly.Checked Then
            Retval.AppendLine("Select X.STYLE_CODE, X.COLOR_CODE")
            Retval.AppendLine(", ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_STATUS, ICTSTYC1.STYLE_COLOR_STATUS, ICTSTYL1.VEND_CODE, ICTSTYL1.FACTORY_CODE, ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.CARTON_PACK_QTY")
            Retval.AppendLine(", ICTCOLR1.COLOR_DESC, CASE WHEN NVL(NEW_PO_COST_DATE,TRUNC(SYSDATE+1)) <= TRUNC(SYSDATE) THEN NEW_PO_COST ELSE PO_COST END PO_COST")
            Retval.AppendLine(", X.SLS_AMT, X.SLS_AMT_WHSE1, X.SLS_AMT_WHSE2")
            Retval.AppendLine(", X.SLS_QTY, X.SLS_QTY_WHSE1, X.SLS_QTY_WHSE2")
            Retval.AppendLine(", Z.SLS_CUSTS, Z.SLS_CUSTS_WHSE1, Z.SLS_CUSTS_WHSE2, Z.SLS_CUSTS_WHSEX")
            Retval.AppendLine(", Y.FUT_QTY, Y.FUT_QTY_WHSE1, Y.FUT_QTY_WHSE2, ICTSTYL1.CUST_CODE, ARTCUST1.CUST_NAME, ICTSTYC1.THEME_CODE")
            Retval.AppendLine(" from ICTSTYL1, ICTCOLR1, ICTSTYV1, ICTSTYC1, ARTCUST1")
            Retval.AppendLine(",(Select STYLE_CODE, COLOR_CODE")
            Retval.AppendLine(", Sum (SLS_AMT) SLS_AMT, Sum (SLS_AMT_WHSE1) SLS_AMT_WHSE1, Sum (SLS_AMT_WHSE2) SLS_AMT_WHSE2")
            Retval.AppendLine(", Sum (SLS_QTY) SLS_QTY, Sum (SLS_QTY_WHSE1) SLS_QTY_WHSE1, Sum (SLS_QTY_WHSE2) SLS_QTY_WHSE2")
            Retval.AppendLine(" from (")
            Retval.AppendLine("Select SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE")
            Retval.AppendLine(", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)) SLS_AMT")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE1")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE2")
            Retval.AppendLine(", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) SLS_QTY")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE1")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE2")
            Retval.AppendLine(" from SOTINVH2,SOTINVH1")
            Retval.AppendLine(" where SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN 'YP1' AND 'YP2'")
            Retval.AppendLine("   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
            Retval.AppendLine("   and SOTINVH2.ORDR_QTY_SHIP < 0")
            Retval.AppendLine("   and SOTINVH2.INV_TYPE = 'C'")
            Retval.AppendLine("   and SOTINVH2.CUST_CODE = SOTINVH2.CUST_CODE")
            Retval.AppendLine(" group by SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE")
            Retval.AppendLine(") group by STYLE_CODE, COLOR_CODE")
            Retval.AppendLine(") X")
            Retval.AppendLine(", (Select STYLE_CODE, COLOR_CODE")
            Retval.AppendLine(", SUM (NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0)) FUT_QTY")
            Retval.AppendLine(", SUM (DECODE(WHSE_CODE,'WHSE1',NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0),0)) FUT_QTY_WHSE1")
            Retval.AppendLine(", SUM (DECODE(WHSE_CODE,'WHSE2',NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0),0)) FUT_QTY_WHSE2")
            Retval.AppendLine(" from ICTSTAT2 GROUP BY STYLE_CODE, COLOR_CODE) Y")
            Retval.AppendLine(", (Select STYLE_CODE, COLOR_CODE")
            Retval.AppendLine(", SUM(CASE WHEN SLS_QTY > 0 THEN 1 ELSE 0 END) SLS_CUSTS")
            Retval.AppendLine(", SUM(CASE WHEN SLS_QTY_WHSE1 > 0 THEN 1 ELSE 0 END) SLS_CUSTS_WHSE1")
            Retval.AppendLine(", SUM(CASE WHEN SLS_QTY_WHSE2 > 0 THEN 1 ELSE 0 END) SLS_CUSTS_WHSE2")
            Retval.AppendLine(", SUM(CASE WHEN SLS_QTY_WHSEX > 0 THEN 1 ELSE 0 END) SLS_CUSTS_WHSEX")
            Retval.AppendLine(String.Format(" from {0} group by STYLE_CODE, COLOR_CODE) Z", SATCSLSX))
            Retval.AppendLine(" where Y.STYLE_CODE (+) = X.STYLE_CODE")
            Retval.AppendLine("   and Y.COLOR_CODE (+) = X.COLOR_CODE")
            Retval.AppendLine("   and Z.STYLE_CODE (+) = X.STYLE_CODE")
            Retval.AppendLine("   and Z.COLOR_CODE (+) = X.COLOR_CODE")
            Retval.AppendLine("   and ICTSTYC1.STYLE_CODE (+) = X.STYLE_CODE")
            Retval.AppendLine("   and ICTSTYC1.COLOR_CODE (+) = X.COLOR_CODE")
            Retval.AppendLine("   and ICTSTYL1.STYLE_CODE = X.STYLE_CODE")
            Retval.AppendLine("   and ICTCOLR1.COLOR_CODE = X.COLOR_CODE")
            Retval.AppendLine("   and ICTSTYV1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE")
            Retval.AppendLine("   and ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE")
            Retval.AppendLine("   and ICTSTYL1.CUST_CODE = ARTCUST1.CUST_CODE (+)")
        Else
            Retval.AppendLine("Select X.STYLE_CODE, X.COLOR_CODE")
            Retval.AppendLine(", ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_STATUS, ICTSTYC1.STYLE_COLOR_STATUS, ICTSTYL1.VEND_CODE, ICTSTYL1.FACTORY_CODE, ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.CARTON_PACK_QTY")
            Retval.AppendLine(", ICTCOLR1.COLOR_DESC, CASE WHEN NVL(NEW_PO_COST_DATE,TRUNC(SYSDATE+1)) <= TRUNC(SYSDATE) THEN NEW_PO_COST ELSE PO_COST END PO_COST")
            Retval.AppendLine(", X.SLS_AMT, X.SLS_AMT_WHSE1, X.SLS_AMT_WHSE2")
            Retval.AppendLine(", X.SLS_QTY, X.SLS_QTY_WHSE1, X.SLS_QTY_WHSE2")
            Retval.AppendLine(", Z.SLS_CUSTS, Z.SLS_CUSTS_WHSE1, Z.SLS_CUSTS_WHSE2, Z.SLS_CUSTS_WHSEX")
            Retval.AppendLine(", Y.FUT_QTY, Y.FUT_QTY_WHSE1, Y.FUT_QTY_WHSE2, ICTSTYL1.CUST_CODE, ARTCUST1.CUST_NAME, ICTSTYC1.THEME_CODE")
            Retval.AppendLine(" from ICTSTYL1, ICTCOLR1, ICTSTYV1, ICTSTYC1, ARTCUST1")
            Retval.AppendLine(",(Select STYLE_CODE, COLOR_CODE")
            Retval.AppendLine(", Sum (SLS_AMT) SLS_AMT, Sum (SLS_AMT_WHSE1) SLS_AMT_WHSE1, Sum (SLS_AMT_WHSE2) SLS_AMT_WHSE2")
            Retval.AppendLine(", Sum (SLS_QTY) SLS_QTY, Sum (SLS_QTY_WHSE1) SLS_QTY_WHSE1, Sum (SLS_QTY_WHSE2) SLS_QTY_WHSE2")
            Retval.AppendLine(" from (")
            Retval.AppendLine("Select SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE")
            Retval.AppendLine(", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)) SLS_AMT")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE1")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE2")
            Retval.AppendLine(", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) SLS_QTY")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE1")
            Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE2")
            Retval.AppendLine(" from SOTINVH2,SOTINVH1")
            Retval.AppendLine(" where SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN 'YP1' AND 'YP2'")
            Retval.AppendLine("   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
            Retval.AppendLine("   and SOTINVH2.ORDR_QTY_SHIP <> 0")
            Retval.AppendLine("   and SOTINVH2.INV_TYPE = 'I'")
            Retval.AppendLine("   and SOTINVH2.CUST_CODE = SOTINVH2.CUST_CODE")
            Retval.AppendLine(" group by SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE")
            If chkIncludeCredits.Checked Then
                Retval.AppendLine(" union ")
                Retval.AppendLine("Select SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE")
                Retval.AppendLine(", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)) SLS_AMT")
                Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE1")
                Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE2")
                Retval.AppendLine(", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) SLS_QTY")
                Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE1")
                Retval.AppendLine(", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE2")
                Retval.AppendLine(" from SOTINVH2,SOTINVH1")
                Retval.AppendLine(" where SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN 'YP1' AND 'YP2'")
                Retval.AppendLine("   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
                Retval.AppendLine("   and SOTINVH2.ORDR_QTY_SHIP < 0")
                Retval.AppendLine("   and SOTINVH2.INV_TYPE = 'C'")
                Retval.AppendLine("   and SOTINVH2.CUST_CODE = SOTINVH2.CUST_CODE")
                Retval.AppendLine(" group by SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE")
            End If
            Retval.AppendLine(" union ")
            Retval.AppendLine("Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE")
            Retval.AppendLine(", SUM ((NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) SLS_AMT")
            Retval.AppendLine(", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE1',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE1")
            Retval.AppendLine(", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE2',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE2")
            Retval.AppendLine(", SUM ((NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0))) SLS_QTY")
            Retval.AppendLine(", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE1',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)),0)) SLS_QTY_WHSE1")
            Retval.AppendLine(", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE2',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)),0)) SLS_QTY_WHSE2")
            Retval.AppendLine(" from SOTORDR2,SOTORDR1")
            Retval.AppendLine(" where SOTORDR2.ORDR_STATUS BETWEEN 'Z' AND 'Z'")
            Retval.AppendLine(" and SOTORDR1.ORDR_YYYYPP_BOOKED BETWEEN 'YP1' AND 'YP2'")
            Retval.AppendLine("   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO")
            Retval.AppendLine("   and (NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) <> 0")
            Retval.AppendLine("   and SOTORDR1.CUST_CODE = SOTORDR1.CUST_CODE")
            Retval.AppendLine(" group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE")
            Retval.AppendLine(") group by STYLE_CODE, COLOR_CODE")
            Retval.AppendLine(") X")
            Retval.AppendLine(", (Select STYLE_CODE, COLOR_CODE")
            Retval.AppendLine(", SUM (NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0)) FUT_QTY")
            Retval.AppendLine(", SUM (DECODE(WHSE_CODE,'WHSE1',NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0),0)) FUT_QTY_WHSE1")
            Retval.AppendLine(", SUM (DECODE(WHSE_CODE,'WHSE2',NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0),0)) FUT_QTY_WHSE2")
            Retval.AppendLine(" from ICTSTAT2 GROUP BY STYLE_CODE, COLOR_CODE) Y")
            Retval.AppendLine(", (Select STYLE_CODE, COLOR_CODE")
            Retval.AppendLine(", SUM(CASE WHEN SLS_QTY > 0 THEN 1 ELSE 0 END) SLS_CUSTS")
            Retval.AppendLine(", SUM(CASE WHEN SLS_QTY_WHSE1 > 0 THEN 1 ELSE 0 END) SLS_CUSTS_WHSE1")
            Retval.AppendLine(", SUM(CASE WHEN SLS_QTY_WHSE2 > 0 THEN 1 ELSE 0 END) SLS_CUSTS_WHSE2")
            Retval.AppendLine(", SUM(CASE WHEN SLS_QTY_WHSEX > 0 THEN 1 ELSE 0 END) SLS_CUSTS_WHSEX")
            Retval.AppendLine(String.Format(" from {0} group by STYLE_CODE, COLOR_CODE) Z", SATCSLSX))
            Retval.AppendLine(" where Y.STYLE_CODE (+) = X.STYLE_CODE")
            Retval.AppendLine("   and Y.COLOR_CODE (+) = X.COLOR_CODE")
            Retval.AppendLine("   and Z.STYLE_CODE (+) = X.STYLE_CODE")
            Retval.AppendLine("   and Z.COLOR_CODE (+) = X.COLOR_CODE")
            Retval.AppendLine("   and ICTSTYC1.STYLE_CODE (+) = X.STYLE_CODE")
            Retval.AppendLine("   and ICTSTYC1.COLOR_CODE (+) = X.COLOR_CODE")
            Retval.AppendLine("   and ICTSTYL1.STYLE_CODE = X.STYLE_CODE")
            Retval.AppendLine("   and ICTCOLR1.COLOR_CODE = X.COLOR_CODE")
            Retval.AppendLine("   and ICTSTYV1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE")
            Retval.AppendLine("   and ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE")
            Retval.AppendLine("   and ICTSTYL1.CUST_CODE = ARTCUST1.CUST_CODE (+)")
        End If
        Return Retval.ToString
    End Function

    Private Function MakesqlSOTINVHX(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal RYP0 As String, ByVal RYP1 As String) As String
        Dim Retval As System.Text.StringBuilder = New System.Text.StringBuilder() With {.Length = 0}
        If chkCreditsOnly.Checked Then
            Retval.AppendLine("Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO")
            Retval.AppendLine(", SOTINVH1.INV_DATE, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.ORDR_CUST_PO")
            Retval.AppendLine(", SOTINVH2.ORDR_YYYYPP_UPDATED OPS_YYYYPP, SOTINVH1.ORDR_NO")
            Retval.AppendLine(", SOTINVH1.SREP_CODE, SOTINVH1.WHSE_CODE, SOTINVH2.CUST_CODE, ARTCUST1.CUST_NAME, SOTINVH1.CUST_STORE_NO")
            Retval.AppendLine(", NVL(ARTCUST2.CUST_ADDR_NAME,ARTCUST2.CUST_NAME) CUST_STORE_LOCATION")
            Retval.AppendLine(", SOTINVH2.STYLE_CODE, ICTSTYL1.STYLE_DESC")
            Retval.AppendLine(", SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE")
            Retval.AppendLine(", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE ORDR_AMT_SHIP,")
            Retval.AppendLine(", 0 ORDR_QTY_OPEN,")
            Retval.AppendLine(", 0 ORDR_QTY_PICK,")
            Retval.AppendLine(", 0 ORDR_QTY_CANC,")
            Retval.AppendLine(", SYSDATE ORDR_DATE_RECD")
            Retval.AppendLine(" from SOTINVH2,ICTSTYL1,ARTCUST2,SOTINVH1,ARTCUST1 ")
            Retval.AppendLine(" where ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE ")
            Retval.AppendLine(" and ARTCUST2.CUST_CODE (+) = SOTINVH1.CUST_CODE ")
            Retval.AppendLine(" and ARTCUST2.CUST_ADDR_TYPE (+) = 'MK' ")
            Retval.AppendLine(" and ARTCUST2.CUST_ADDR_CODE (+) = SOTINVH1.CUST_STORE_NO ")
            Retval.AppendLine(" and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE ")
            Retval.AppendLine(" and SOTINVH1.INV_NO = SOTINVH2.INV_NO")
            Retval.AppendLine(" and SOTINVH2.ORDR_QTY_SHIP < 0")
            Retval.AppendLine(" and SOTINVH2.INV_TYPE = 'C'")
            Retval.AppendLine(" and ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE")
            Retval.AppendLine(" AND SOTINVH2.STYLE_CODE = '" & STYLE_CODE & "'")
            Retval.AppendLine(" and SOTINVH2.COLOR_CODE = '" & COLOR_CODE & "'")
            Retval.AppendLine(" and SOTINVH1.ORDR_YYYYPP_UPDATED between '" & RYP0 & "' and '" & RYP1 & "'")
        Else
            Retval.AppendLine("Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO")
            Retval.AppendLine(", SOTINVH1.INV_DATE, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.ORDR_CUST_PO")
            Retval.AppendLine(", SOTINVH2.ORDR_YYYYPP_UPDATED OPS_YYYYPP, SOTINVH1.ORDR_NO")
            Retval.AppendLine(", SOTINVH1.SREP_CODE, SOTINVH1.WHSE_CODE, SOTINVH2.CUST_CODE, ARTCUST1.CUST_NAME, SOTINVH1.CUST_STORE_NO")
            Retval.AppendLine(", NVL(ARTCUST2.CUST_ADDR_NAME,ARTCUST2.CUST_NAME) CUST_STORE_LOCATION")
            Retval.AppendLine(", SOTINVH2.STYLE_CODE, ICTSTYL1.STYLE_DESC")
            Retval.AppendLine(", SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE")
            Retval.AppendLine(", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE ORDR_AMT_SHIP")
            Retval.AppendLine(", 0 ORDR_QTY_OPEN")
            Retval.AppendLine(", 0 ORDR_QTY_PICK")
            Retval.AppendLine(", 0 ORDR_QTY_CANC")
            Retval.AppendLine(", SYSDATE ORDR_DATE_RECD")
            Retval.AppendLine(" from SOTINVH2,ICTSTYL1,ARTCUST2,SOTINVH1,ARTCUST1 ")
            Retval.AppendLine(" where ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE ")
            Retval.AppendLine(" and ARTCUST2.CUST_CODE (+) = SOTINVH1.CUST_CODE ")
            Retval.AppendLine(" and ARTCUST2.CUST_ADDR_TYPE (+) = 'MK' ")
            Retval.AppendLine(" and ARTCUST2.CUST_ADDR_CODE (+) = SOTINVH1.CUST_STORE_NO ")
            Retval.AppendLine(" and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE ")
            Retval.AppendLine(" and SOTINVH1.INV_NO = SOTINVH2.INV_NO")
            Retval.AppendLine(" and SOTINVH2.ORDR_QTY_SHIP <> 0")
            Retval.AppendLine(" and SOTINVH2.INV_TYPE = 'I'")
            Retval.AppendLine(" and ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE")
            Retval.AppendLine(String.Format(" AND SOTINVH2.STYLE_CODE = '{0}'", STYLE_CODE))
            Retval.AppendLine(" and SOTINVH2.COLOR_CODE = '" & COLOR_CODE & "'")
            Retval.AppendLine(" and SOTINVH1.ORDR_YYYYPP_UPDATED between '" & RYP0 & "' and '" & RYP1 & "'")
            If chkIncludeCredits.Checked Then
                Retval.AppendLine(" UNION ")
                Retval.AppendLine("Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO")
                Retval.AppendLine(", SOTINVH1.INV_DATE, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.ORDR_CUST_PO")
                Retval.AppendLine(", SOTINVH2.ORDR_YYYYPP_UPDATED OPS_YYYYPP, SOTINVH1.ORDR_NO")
                Retval.AppendLine(", SOTINVH1.SREP_CODE, SOTINVH1.WHSE_CODE, SOTINVH2.CUST_CODE, ARTCUST1.CUST_NAME, SOTINVH1.CUST_STORE_NO")
                Retval.AppendLine(", NVL(ARTCUST2.CUST_ADDR_NAME,ARTCUST2.CUST_NAME) CUST_STORE_LOCATION")
                Retval.AppendLine(", SOTINVH2.STYLE_CODE, ICTSTYL1.STYLE_DESC")
                Retval.AppendLine(", SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE")
                Retval.AppendLine(", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE ORDR_AMT_SHIP")
                Retval.AppendLine(", 0 ORDR_QTY_OPEN")
                Retval.AppendLine(", 0 ORDR_QTY_PICK")
                Retval.AppendLine(", 0 ORDR_QTY_CANC")
                Retval.AppendLine(", SYSDATE ORDR_DATE_RECD")
                Retval.AppendLine(" from SOTINVH2,ICTSTYL1,ARTCUST2,SOTINVH1,ARTCUST1 ")
                Retval.AppendLine(" where ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE ")
                Retval.AppendLine(" and ARTCUST2.CUST_CODE (+) = SOTINVH1.CUST_CODE ")
                Retval.AppendLine(" and ARTCUST2.CUST_ADDR_TYPE (+) = 'MK' ")
                Retval.AppendLine(" and ARTCUST2.CUST_ADDR_CODE (+) = SOTINVH1.CUST_STORE_NO ")
                Retval.AppendLine(" and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE ")
                Retval.AppendLine(" and SOTINVH1.INV_NO = SOTINVH2.INV_NO")
                Retval.AppendLine(" and SOTINVH2.ORDR_QTY_SHIP < 0")
                Retval.AppendLine(" and SOTINVH2.INV_TYPE = 'C'")
                Retval.AppendLine(" and ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE")
                Retval.AppendLine(" AND SOTINVH2.STYLE_CODE = '" & STYLE_CODE & "'")
                Retval.AppendLine(" and SOTINVH2.COLOR_CODE = '" & COLOR_CODE & "'")
                Retval.AppendLine(" and SOTINVH1.ORDR_YYYYPP_UPDATED between '" & RYP0 & "' and '" & RYP1 & "'")
            End If
        End If
        Return Retval.ToString
    End Function

    Function FetchImage(ByVal STYLE_CODE_IMG As String, ByVal COLOR_CODE_IMG As String) As Byte()
        Dim IMAGE_NAME As String = STYLE_CODE_IMG & "-" & COLOR_CODE_IMG

        Dim imgba() As Byte = Nothing
        If IMAGE_NAME <> "" Then
            imgSTYLE.Image = Get_Style_Image(IMAGE_NAME, imgba)
            UltraExplorerBar1.Groups("Style Image").Text = "Style " & STYLE_CODE_IMG & "-" & COLOR_CODE_IMG
        Else
            imgSTYLE.Image = Nothing
            UltraExplorerBar1.Groups("Style Image").Text = "Style Image"
        End If

        Return imgba
    End Function

    Function Get_Style_Image(
        ByVal IMAGE_NAME As String,
        Optional ByRef imgba() As Byte = Nothing) As System.Drawing.Bitmap
        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
        Return ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba)
    End Function

    Private Sub imgSTYLE_DoubleClick(sender As Object, e As System.EventArgs) Handles imgSTYLE.DoubleClick
        If Not IsNothing(grdSATCSLS1.ActiveRow) And grdSATCSLS1.ActiveRow.IsDataRow Then
            Dim STYLE_CODE_IMG As String = grdSATCSLS1.ActiveRow.Cells.Item("STYLE_CODE").Text & String.Empty
            Dim STYLE_DESC_IMG As String = grdSATCSLS1.ActiveRow.Cells.Item("STYLE_DESC").Text & String.Empty
            Dim COLOR_CODE_IMG As String = grdSATCSLS1.ActiveRow.Cells.Item("COLOR_CODE").Text & String.Empty
            Using F As New ASFMSGBF
                F.Show_img(imgSTYLE.Image, Me, "Style " & STYLE_CODE_IMG & ":" & STYLE_DESC_IMG)
            End Using
        End If
    End Sub

    Private Function GET_THEME_INFO(ByVal THEME_CODE As String, ByVal COL_NAME As String) As String
        Dim RetVal As String = ""
        Dim filter As String = String.Format("THEME_CODE = '{0}'", THEME_CODE)
        Dim rowICTTHEME As DataRow = dst.Tables("ICTTHEME").Select(filter).FirstOrDefault
        If Not IsNothing(rowICTTHEME) Then
            RetVal = rowICTTHEME.Item(COL_NAME).ToString & String.Empty
        End If
        Return RetVal
    End Function

    Private Sub EcomIndicator()
        Try
            If Not (grdSATCSLSH.ActiveRow Is Nothing OrElse Not grdSATCSLSH.ActiveRow.IsDataRow) Then
                Dim STYLE_CODE As String = grdSATCSLSH.ActiveRow.Cells("STYLE_CODE").Value
                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    Dim ECOM_MSG As String = TAC.TACMAIN1.getEcomInfo(Me, STYLE_CODE)
                    If ECOM_MSG.Length > 0 Then
                        lblEcomStyle.Visible = True
                        Dim TTI As New UltraWinToolTip.UltraToolTipInfo
                        If Not IsNothing(TTM.GetUltraToolTip(lblEcomStyle)) Then
                            TTI.ToolTipTitle = "E-Commerce Information:"
                            TTM.AutoPopDelay = 20000
                            TTI.ToolTipTextFormatted = ECOM_MSG
                            TTM.SetUltraToolTip(lblEcomStyle, TTI)
                        Else
                            TTI.ToolTipTextFormatted = ECOM_MSG
                        End If
                    Else
                        lblEcomStyle.Visible = False
                    End If

                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

End Class