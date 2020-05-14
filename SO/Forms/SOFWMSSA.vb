Imports System.Drawing
Imports System.IO
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Http.Formatting
Imports Newtonsoft.Json
Imports Newtonsoft
Imports System.Reflection
Imports Infragistics.Win.UltraWinGrid

Public Class SOFWMSSA
    Private SQL As New System.Text.StringBuilder With {.Length = 0}
    Private RYP0 As String = ""
    Private RYP1 As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        txtCUST_CODE.Text = "WALMART"
        txtCUST_CODE.ReadOnly = True
        txtCUST_CODE.Enabled = False
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, -6)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -24, 0, 0)

        Dim RYPLEGEND0 = Absx1.cmbFor("RYP0", True).Value
        Dim RYPLEGEND1 = Absx1.cmbFor("RYP1", True).Value
        RYP0 = Mid(RYPLEGEND0, 1, 4) & Mid(RYPLEGEND0, 6, 2)
        RYP1 = Mid(RYPLEGEND1, 1, 4) & Mid(RYPLEGEND1, 6, 2)

        With dst
            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("'0' AS SEL,")
            SQL.AppendLine("I1.CUST_CODE,")
            SQL.AppendLine("C2.CUST_ADDR_CODE,")
            SQL.AppendLine("C2.CUST_NAME,")
            SQL.AppendLine("C2.CUST_STATE,")
            SQL.AppendLine("SUM(I2.ORDR_QTY_SHIP) AS UNITS,")
            SQL.AppendLine("SUM(I2.ORDR_QTY_SHIP * I2.ORDR_UNIT_PRICE) AS TOTAL")
            SQL.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2, ARTCUST2 C2")
            SQL.AppendLine("WHERE I1.INV_TYPE = I2.INV_TYPE")
            SQL.AppendLine("AND I1.INV_NO = I2.INV_NO")
            SQL.AppendLine("AND I1.CUST_CODE = C2.CUST_CODE")
            SQL.AppendLine("AND I1.CUST_STORE_NO = C2.CUST_ADDR_CODE")
            SQL.AppendLine("AND C2.CUST_ADDR_TYPE = 'MK'")
            SQL.AppendLine(String.Format("AND I1.ORDR_YYYYPP_UPDATED >= '{0}'", RYP0))
            SQL.AppendLine(String.Format("AND I1.ORDR_YYYYPP_UPDATED <= '{0}'", RYP1))
            SQL.AppendLine("AND I1.CUST_CODE = 'WALMART'")
            SQL.AppendLine("GROUP BY")
            SQL.AppendLine("I1.CUST_CODE,")
            SQL.AppendLine("C2.CUST_ADDR_CODE,")
            SQL.AppendLine("C2.CUST_NAME,")
            SQL.AppendLine("C2.CUST_STATE")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "SOTWMSS1", "**", 0, False)

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("I1.WHSE_CODE,")
            SQL.AppendLine("I1.CUST_CODE,")
            SQL.AppendLine("I1.CUST_STORE_NO,")
            SQL.AppendLine("I1.INV_DATE,")
            SQL.AppendLine("I1.INV_NO,")
            SQL.AppendLine("I1.INV_NO_CONS,")
            SQL.AppendLine("I1.PICK_NO,")
            SQL.AppendLine("I1.ORDR_CUST_PO,")
            SQL.AppendLine("SUM(I2.ORDR_QTY_SHIP) AS UNITS,")
            SQL.AppendLine("SUM(I2.ORDR_QTY_SHIP * I2.ORDR_UNIT_PRICE) AS TOTAL")
            SQL.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2")
            SQL.AppendLine("WHERE I1.INV_TYPE = I2.INV_TYPE")
            SQL.AppendLine("AND I1.INV_NO = I2.INV_NO")
            SQL.AppendLine(String.Format("AND I1.ORDR_YYYYPP_UPDATED >= '{0}'", RYP0))
            SQL.AppendLine(String.Format("AND I1.ORDR_YYYYPP_UPDATED <= '{0}'", RYP1))
            SQL.AppendLine("AND I1.CUST_CODE = 'WALMART'")
            SQL.AppendLine("GROUP BY I1.CUST_CODE,")
            SQL.AppendLine("I1.CUST_STORE_NO,")
            SQL.AppendLine("I1.INV_NO,")
            SQL.AppendLine("I1.INV_NO_CONS,")
            SQL.AppendLine("I1.PICK_NO,")
            SQL.AppendLine("I1.ORDR_CUST_PO,")
            SQL.AppendLine("I1.WHSE_CODE,")
            SQL.AppendLine("I1.INV_DATE")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "SOTWMSS2", "**", 0, False)

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("S1.CUST_STORE_NO,")
            SQL.AppendLine("S1.INV_NO,")
            SQL.AppendLine("S1.INV_DATE,")
            SQL.AppendLine("S1.ORDR_CUST_PO,")
            SQL.AppendLine("S1.ORDR_NO,")
            SQL.AppendLine("S2.STYLE_CODE,")
            SQL.AppendLine("S2.COLOR_CODE,")
            SQL.AppendLine("S2.ORDR_QTY_SHIP,")
            SQL.AppendLine("S2.ORDR_UNIT_PRICE,")
            SQL.AppendLine("(S2.ORDR_QTY_SHIP * S2.ORDR_UNIT_PRICE) AS TOTAL,")
            SQL.AppendLine("MIN(C1.CUST_STYLE_CODE) AS CUST_STYLE_CODE")
            SQL.AppendLine("FROM SOTINVH1 S1, SOTINVH2 S2, SOTCSTY1 C1")
            SQL.AppendLine("WHERE S1.INV_TYPE = S2.INV_TYPE")
            SQL.AppendLine("AND S1.INV_NO = S2.INV_NO")
            SQL.AppendLine("AND S1.INV_NO = '0000000000'")
            SQL.AppendLine("AND S2.STYLE_CODE = C1.STYLE_CODE (+)")
            SQL.AppendLine("AND S2.COLOR_CODE = C1.COLOR_CODE (+)")
            SQL.AppendLine("AND C1.CUST_CODE (+) = 'XXXXXX'")
            SQL.AppendLine("GROUP BY")
            SQL.AppendLine("S1.CUST_STORE_NO,")
            SQL.AppendLine("S1.INV_NO,")
            SQL.AppendLine("S1.INV_DATE,")
            SQL.AppendLine("S1.ORDR_CUST_PO,")
            SQL.AppendLine("S1.ORDR_NO,")
            SQL.AppendLine("S2.STYLE_CODE,")
            SQL.AppendLine("S2.COLOR_CODE,")
            SQL.AppendLine("S2.ORDR_QTY_SHIP,")
            SQL.AppendLine("S2.ORDR_UNIT_PRICE,")
            SQL.AppendLine("(s2.ORDR_QTY_SHIP * s2.ORDR_UNIT_PRICE)")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "SOTWMSS3", "**", 0, False)
        End With

        grdSOTWMSS1.DataSource = dst.Tables("SOTWMSS1")
        grdSOTWMSS2.DataSource = dst.Tables("SOTWMSS2")
        grdSOTWMSS3.DataSource = dst.Tables("SOTWMSS3")

        With grdSOTWMSS1.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "SEL" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                Select Case GCOL.Key
                    Case "UNITS"
                        GCOL.Format = "#,##0"
                    Case "TOTAL"
                        GCOL.Format = "#,##0.00"
                End Select
            Next
        End With

        With grdSOTWMSS2.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                Select Case GCOL.Key
                    Case "UNITS"
                        GCOL.Format = "#,##0"
                    Case "TOTAL"
                        GCOL.Format = "#,##0.00"
                    Case "INV_DATE"
                        GCOL.Format = "MM/dd/yy"
                End Select
            Next
        End With

        With grdSOTWMSS3.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                Select Case GCOL.Key
                    Case "TOTAL"
                        GCOL.Format = "#,##0.00"
                    Case "ORDR_QTY_SHIP"
                        GCOL.Format = "#,##0"
                    Case "INV_DATE"
                        GCOL.Format = "MM/dd/yy"
                End Select
            Next
        End With

        Create_Summary(grdSOTWMSS1, New String() {"UNITS", "TOTAL"})
        Create_Summary(grdSOTWMSS2, New String() {"UNITS", "TOTAL"})
        Create_Summary(grdSOTWMSS3, New String() {"ORDR_QTY_SHIP", "TOTAL"})

        Sort_grdColumns(grdSOTWMSS1, "CUST_CODE, CUST_ADDR_CODE", False)
        Sort_grdColumns(grdSOTWMSS2, "CUST_CODE, CUST_STORE_NO, INV_DATE, INV_NO, INV_NO_CONS", False)
        Sort_grdColumns(grdSOTWMSS3, "CUST_STORE_NO, INV_NO, INV_DATE, ORDR_CUST_PO", False)

    End Sub

    Sub Check_Inquiry_Mode()
        If InquiryMode Then
        Else
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Update"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Done"
                Call Mode_Settings(False)
                Me.Close()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            '.Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
            '.Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        'dst.Tables("SOFCGTIN").Rows.Clear()

        dst.EnforceConstraints = False


        dst.EnforceConstraints = True

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record()
        'BeginTrans()
        'INIT_LAST("PMTVIST1", True, "", True)
        'Update_Record_TDA("PMTVIST1")
        'CommitTrans("Update Complete")
    End Sub

    Sub Print_Report()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'CR_params.Add("GROUP_BY", optGroupBy.Value)
        'Generate_Report("PMRVIST1", "Open Site Visit Report")
        'Print_Report_End()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        'Select Case COLUMN_NAME
        '    Case "JOB_NO"
        '        sql_where = "JOB_STATUS = 'O' and SITE_VISITS > 0"
        'End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTWMSS1, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTWMSS2, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTWMSS3, "SS", "Show Filter", "Show GroupBox")

    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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
        'Dim tlb_sbt As UltraWinToolbars.StateButtonTool
        'Dim tlb_btn As UltraWinToolbars.ButtonTool

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case ""

                    If grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.True Then
                        e.Cancel = True
                    End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub

#End Region

#Region "ABSColumn Controls"
    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

    End Sub

    Private Sub btnFETCH_Click(sender As Object, e As EventArgs) Handles btnFETCH.Click
        ASCMAIN1.Progress("Refreshing Data", "")
        Me.Cursor = Cursors.WaitCursor

        Dim RYPLEGEND0 = Absx1.cmbFor("RYP0", True).Value
        Dim RYPLEGEND1 = Absx1.cmbFor("RYP1", True).Value
        RYP0 = Mid(RYPLEGEND0, 1, 4) & Mid(RYPLEGEND0, 6, 2)
        RYP1 = Mid(RYPLEGEND1, 1, 4) & Mid(RYPLEGEND1, 6, 2)

        SQL.Length = 0
        SQL.AppendLine("SELECT")
        SQL.AppendLine("'0' AS SEL,")
        SQL.AppendLine("I1.CUST_CODE,")
        SQL.AppendLine("C2.CUST_ADDR_CODE,")
        SQL.AppendLine("C2.CUST_NAME,")
        SQL.AppendLine("C2.CUST_STATE,")
        SQL.AppendLine("SUM(I2.ORDR_QTY_SHIP) AS UNITS,")
        SQL.AppendLine("SUM(I2.ORDR_QTY_SHIP * I2.ORDR_UNIT_PRICE) AS TOTAL")
        SQL.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2, ARTCUST2 C2")
        SQL.AppendLine("WHERE I1.INV_TYPE = I2.INV_TYPE")
        SQL.AppendLine("AND I1.INV_NO = I2.INV_NO")
        SQL.AppendLine("AND I1.CUST_CODE = C2.CUST_CODE")
        SQL.AppendLine("AND I1.CUST_STORE_NO = C2.CUST_ADDR_CODE")
        SQL.AppendLine("AND C2.CUST_ADDR_TYPE = 'MK'")
        SQL.AppendLine(String.Format("AND I1.ORDR_YYYYPP_UPDATED >= '{0}'", RYP0))
        SQL.AppendLine(String.Format("AND I1.ORDR_YYYYPP_UPDATED <= '{0}'", RYP1))
        SQL.AppendLine("AND I1.CUST_CODE = 'WALMART'")
        SQL.AppendLine("GROUP BY")
        SQL.AppendLine("I1.CUST_CODE,")
        SQL.AppendLine("C2.CUST_ADDR_CODE,")
        SQL.AppendLine("C2.CUST_NAME,")
        SQL.AppendLine("C2.CUST_STATE")
        ASCMAIN1.sql = SQL.ToString
        Fill_Records("SOTWMSS1", "", True, SQL.ToString())

        SQL.Length = 0
        SQL.AppendLine("SELECT")
        SQL.AppendLine("I1.WHSE_CODE,")
        SQL.AppendLine("I1.CUST_CODE,")
        SQL.AppendLine("I1.CUST_STORE_NO,")
        SQL.AppendLine("I1.INV_DATE,")
        SQL.AppendLine("I1.INV_NO,")
        SQL.AppendLine("I1.INV_NO_CONS,")
        SQL.AppendLine("I1.PICK_NO,")
        SQL.AppendLine("I1.ORDR_CUST_PO,")
        SQL.AppendLine("SUM(I2.ORDR_QTY_SHIP) AS UNITS,")
        SQL.AppendLine("SUM(I2.ORDR_QTY_SHIP * I2.ORDR_UNIT_PRICE) AS TOTAL")
        SQL.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2")
        SQL.AppendLine("WHERE I1.INV_TYPE = I2.INV_TYPE")
        SQL.AppendLine("AND I1.INV_NO = I2.INV_NO")
        SQL.AppendLine(String.Format("AND I1.ORDR_YYYYPP_UPDATED >= '{0}'", RYP0))
        SQL.AppendLine(String.Format("AND I1.ORDR_YYYYPP_UPDATED <= '{0}'", RYP1))
        SQL.AppendLine(String.Format("AND I1.CUST_CODE = '{0}'", Absx1.txtFor("CUST_CODE").Text))
        SQL.AppendLine("GROUP BY I1.CUST_CODE,")
        SQL.AppendLine("I1.CUST_STORE_NO,")
        SQL.AppendLine("I1.INV_NO,")
        SQL.AppendLine("I1.INV_NO_CONS,")
        SQL.AppendLine("I1.PICK_NO,")
        SQL.AppendLine("I1.ORDR_CUST_PO,")
        SQL.AppendLine("I1.WHSE_CODE,")
        SQL.AppendLine("I1.INV_DATE")
        ASCMAIN1.sql = SQL.ToString
        Fill_Records("SOTWMSS2", "", True, SQL.ToString())

        RefreshSOTWMSS2()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

#End Region

#Region "Custom Methods"
    Private Sub RefreshSOTWMSS2()
        'ASCMAIN1.Progress("Refreshing Data", "")
        'Me.Cursor = Cursors.WaitCursor
        'grdSOTWMSS1.UpdateData()
        'grdSOTWMSS1.Update()

        Dim STORE_LIST As String = ""
        Dim FILTER As String = "CUST_STORE_NO = 'X9999X'"
        For Each rowSOTWMSS1 As DataRow In dst.Tables("SOTWMSS1").Select()
            If rowSOTWMSS1.Item("SEL").ToString & String.Empty = "1" Then
                STORE_LIST = rowSOTWMSS1.Item("CUST_ADDR_CODE").ToString & "','" & STORE_LIST
            End If
        Next
        If STORE_LIST.Length > 0 Then
            FILTER = "CUST_STORE_NO IN ('" & STORE_LIST.Substring(0, STORE_LIST.Length - 2) & ")"
        End If
        Dim dvw As DataView = DirectCast(grdSOTWMSS2.DataSource, DataTable).DefaultView
        dvw.RowFilter = String.Format(FILTER, "CUST_STORE_NO")

        'Me.Cursor = Cursors.Default
        'ASCMAIN1.Progress("", "")
    End Sub

    Sub RefreshSOTWMSS3()
        If grdSOTWMSS2.ActiveRow Is Nothing OrElse (Not grdSOTWMSS2.ActiveRow.IsDataRow Or grdSOTWMSS2.ActiveRow.IsAddRow) Then
            dst.Tables("SOTWMSS3").Clear()
        Else
            Dim INV_FILTER As String = "INV_NO"
            If chkSHOWPODTL.Checked = True Then
                INV_FILTER = "INV_NO_CONS"
            End If
            Dim INV_VAL As String = grdSOTWMSS2.ActiveRow.Cells.Item(INV_FILTER).Text.ToString & String.Empty

            If INV_VAL.Length > 0 Then
                SQL.Length = 0
                SQL.AppendLine("SELECT")
                SQL.AppendLine("S1.CUST_STORE_NO,")
                SQL.AppendLine("S1.INV_NO,")
                SQL.AppendLine("S1.INV_DATE,")
                SQL.AppendLine("S1.ORDR_CUST_PO,")
                SQL.AppendLine("S1.ORDR_NO,")
                SQL.AppendLine("S2.STYLE_CODE,")
                SQL.AppendLine("S2.COLOR_CODE,")
                SQL.AppendLine("S2.ORDR_QTY_SHIP,")
                SQL.AppendLine("S2.ORDR_UNIT_PRICE,")
                SQL.AppendLine("(S2.ORDR_QTY_SHIP * S2.ORDR_UNIT_PRICE) AS TOTAL,")
                SQL.AppendLine("MIN(C1.CUST_STYLE_CODE) AS CUST_STYLE_CODE")
                SQL.AppendLine("FROM SOTINVH1 S1, SOTINVH2 S2, SOTCSTY1 C1")
                SQL.AppendLine("WHERE S1.INV_TYPE = S2.INV_TYPE")
                SQL.AppendLine("AND S1.INV_NO = S2.INV_NO")
                SQL.AppendLine(String.Format("And S1.{0} = '{1}'", INV_FILTER, INV_VAL))
                SQL.AppendLine("AND S2.STYLE_CODE = C1.STYLE_CODE (+)")
                SQL.AppendLine("AND S2.COLOR_CODE = C1.COLOR_CODE (+)")
                SQL.AppendLine(String.Format("AND C1.CUST_CODE (+) = '{0}'", Absx1.txtFor("CUST_CODE").Text))
                SQL.AppendLine("GROUP BY")
                SQL.AppendLine("S1.CUST_STORE_NO,")
                SQL.AppendLine("S1.INV_NO,")
                SQL.AppendLine("S1.INV_DATE,")
                SQL.AppendLine("S1.ORDR_CUST_PO,")
                SQL.AppendLine("S1.ORDR_NO,")
                SQL.AppendLine("S2.STYLE_CODE,")
                SQL.AppendLine("S2.COLOR_CODE,")
                SQL.AppendLine("S2.ORDR_QTY_SHIP,")
                SQL.AppendLine("S2.ORDR_UNIT_PRICE,")
                SQL.AppendLine("(s2.ORDR_QTY_SHIP * s2.ORDR_UNIT_PRICE)")
                ASCMAIN1.sql = SQL.ToString
                Fill_Records("SOTWMSS3", "", True, SQL.ToString())
            Else
                dst.Tables("SOTWMSS3").Clear()
            End If

        End If
    End Sub

    Private Sub grdSOTWMSS1_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdSOTWMSS1.AfterRowUpdate
        Me.Cursor = Cursors.WaitCursor
        RefreshSOTWMSS2()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub grdSOTWMSS2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTWMSS2.AfterRowActivate
        RefreshSOTWMSS3()
    End Sub

    Private Sub grdSOTWMSS3_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTWMSS3.InitializeRow
        If Not grdSOTWMSS2.ActiveRow Is Nothing Then
            Dim CUST_STORE_NO As String = grdSOTWMSS2.ActiveRow.Cells.Item("CUST_STORE_NO").Text & String.Empty
            If e.Row.Cells("CUST_STORE_NO").Text & String.Empty = CUST_STORE_NO Then
                e.Row.Appearance.BackColor = Color.Yellow
            Else
                e.Row.Appearance.BackColor = Color.Empty
            End If
        End If
    End Sub

    Private Sub chkSHOWPODTL_CheckedChanged(sender As Object, e As EventArgs) Handles chkSHOWPODTL.CheckedChanged
        RefreshSOTWMSS3()
    End Sub
#End Region

#Region "Form Controls"

#End Region

End Class