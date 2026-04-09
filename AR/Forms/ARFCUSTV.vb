Imports System.Net.Mail
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Xml
Imports Infragistics.Documents.Excel
Imports Infragistics.Win.UltraWinGrid
Imports Infragistics.Win.UltraWinSchedule

Public Class ARFCUSTV
    Dim Remote As New REMOTE(Me)
    Dim S As New System.Text.StringBuilder With {.Length = 0}
    Dim Loading As Boolean = True
    Dim TEMP_SALES As String = ""
    Dim TEMP_MAX As String = ""
    Dim YR1 As String = ""
    Dim YR2 As String = ""
    Dim YR3 As String = ""
    Dim YR4 As String = ""
    Dim YRFR As String = ""
    Dim YRTO As String = ""
    Dim EditingContacts As Boolean = False


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Check_Form_Options()

        Dim BaseYear As Int64 = Now().Year
        YR1 = (BaseYear).ToString
        YR2 = (BaseYear - 1).ToString
        YR3 = (BaseYear - 2).ToString
        YR4 = (BaseYear - 3).ToString

        YRFR = Now.AddYears(-2).Year.ToString().Substring(2, 2).ToString
        YRTO = Now.Year.ToString().Substring(2, 2).ToString

        'Fill In The Gaps

        RefreshTempTables()

        With dst
            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("'0' AS SEL,")
            S.AppendLine("ARTCUST1.CUST_CODE,")
            S.AppendLine("ARTCUST1.CUST_NAME,")
            S.AppendLine("ARTCUST1.CUST_ADDR1,")
            S.AppendLine("ARTCUST1.CUST_ADDR2,")
            S.AppendLine("ARTCUST1.CUST_ADDR3,")
            S.AppendLine("ARTCUST1.CUST_CITY,")
            S.AppendLine("ARTCUST1.CUST_STATE,")
            S.AppendLine("ARTCUST1.CUST_ZIP_CODE,")
            S.AppendLine("ARTCUST1.CUST_COUNTRY,")
            S.AppendLine("ARTCUST1.CUST_CONTACT,")
            S.AppendLine("ARTCUST1.CUST_EMAIL,")
            S.AppendLine("ARTCUST1.SREP_CODE,")
            S.AppendLine("ARTCUST1.INIT_DATE,")
            S.AppendLine("ARTCUST1.CUST_SALES_HOLD,")
            S.AppendLine("ARTCUST1.CUST_CREDIT_HOLD,")
            S.AppendLine("NVL(ARTCUST1.CUST_CREDIT_LIMIT,0) AS CUST_CREDIT_LIMIT,")
            S.AppendLine("ARTCUST1.TERM_CODE,")
            S.AppendLine("TATTERM1.TERM_DESC,")
            S.AppendLine("ARTCUST1.CUST_TERMS_NOTE,")
            S.AppendLine("ARTCUST1.CUST_CREDIT_LIMIT_NOTES,")
            S.AppendLine("V1.ACTIVITY_TYPE,")
            S.AppendLine("V1.ACTIVITY_DATE,")
            S.AppendLine("V1.ACTIVITY_NOTE,")
            S.AppendLine("SUM(SALES.YR1) AS YR1,")
            S.AppendLine("SUM(SALES.YR2) AS YR2,")
            S.AppendLine("SUM(SALES.YR3) AS YR3,")
            S.AppendLine("SUM(SALES.YR4) AS YR4,")
            S.AppendLine("'0' AS HAS_SALES")
            S.AppendLine($"FROM ARTCUST1, {TEMP_SALES} SALES, TATTERM1, {TEMP_MAX} V1")
            S.AppendLine("WHERE ARTCUST1.CUST_CODE = SALES.CUST_CODE (+)")
            S.AppendLine("AND ARTCUST1.TERM_CODE = TATTERM1.TERM_CODE (+)")
            S.AppendLine("AND ARTCUST1.CUST_CODE = V1.CUST_CODE (+)")
            S.AppendLine("GROUP BY")
            S.AppendLine("ARTCUST1.CUST_CODE,")
            S.AppendLine("ARTCUST1.CUST_NAME,")
            S.AppendLine("ARTCUST1.CUST_ADDR1,")
            S.AppendLine("ARTCUST1.CUST_ADDR2,")
            S.AppendLine("ARTCUST1.CUST_ADDR3,")
            S.AppendLine("ARTCUST1.CUST_CITY,")
            S.AppendLine("ARTCUST1.CUST_STATE,")
            S.AppendLine("ARTCUST1.CUST_ZIP_CODE,")
            S.AppendLine("ARTCUST1.CUST_COUNTRY,")
            S.AppendLine("ARTCUST1.CUST_CONTACT,")
            S.AppendLine("ARTCUST1.CUST_EMAIL,")
            S.AppendLine("ARTCUST1.SREP_CODE,")
            S.AppendLine("ARTCUST1.INIT_DATE,")
            S.AppendLine("ARTCUST1.CUST_SALES_HOLD,")
            S.AppendLine("ARTCUST1.CUST_CREDIT_HOLD,")
            S.AppendLine("NVL(ARTCUST1.CUST_CREDIT_LIMIT,0),")
            S.AppendLine("ARTCUST1.TERM_CODE,")
            S.AppendLine("TATTERM1.TERM_DESC,")
            S.AppendLine("ARTCUST1.CUST_TERMS_NOTE,")
            S.AppendLine("ARTCUST1.CUST_CREDIT_LIMIT_NOTES,")
            S.AppendLine("V1.ACTIVITY_TYPE,")
            S.AppendLine("V1.ACTIVITY_DATE,")
            S.AppendLine("V1.ACTIVITY_NOTE")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ARTCUSTX", "**", 0, False)
            'With .Tables("ARFCUSTX").Columns
            '    .Add("XXX", GetType(String))
            'End With

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ARTCUSTD")
            S.AppendLine("WHERE CUST_CODE = :PARM1")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ARTCUSTD", "**", 0, True, "V")
            Create_TDA(.Tables.Add("ARTCUSTO"), "ARTCUSTD", "**", 0, False, "V")

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ARTCUSV1")
            S.AppendLine("WHERE CUST_CODE = :PARM1")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ARTCUSV1", "**", 0, True, "V")
        End With

        grdARTCUSTX.DataSource = dst.Tables("ARTCUSTX")
        grdARTCUSTD.DataSource = dst.Tables("ARTCUSTD")
        grdARTCUSV1.DataSource = dst.Tables("ARTCUSV1")

        ASCMAIN1.Add_Value_List(grdARTCUSTD, "CONTACT_TYPE")

        Create_Summary(grdARTCUSTX, "CUST_CODE", "Count")
        Create_Summary(grdARTCUSTX, "YR1", "Sum")
        Create_Summary(grdARTCUSTX, "YR2", "Sum")
        Create_Summary(grdARTCUSTX, "YR3", "Sum")
        Create_Summary(grdARTCUSTX, "YR4", "Sum")

        'Create_Summary(grdSOTORDR2, New String() {"ORDR_QTY", "ORDR_AMT", "TCUFT", "ORDR_QTY_ALLO", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_AMT_OPEN", "ORDR_AMT_ALLO", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"})

        Sort_grdColumns(grdARTCUSTX, "activity_date, CUST_NAME", False)
        Sort_grdColumns(grdARTCUSV1, "ACTIVITY_NO", False)

        With grdARTCUSTX.DisplayLayout.Bands(0)
            For Each COL_NAME As String In New String() {"CUST_CODE", "CUST_NAME"}
                .Columns(COL_NAME).Header.Fixed = True
            Next
            .Columns("YR1").Header.Caption = YR1
            .Columns("YR1").Format = "###,##0.00"
            .Columns("YR2").Header.Caption = YR2
            .Columns("YR2").Format = "###,##0.00"
            .Columns("YR3").Header.Caption = YR3
            .Columns("YR3").Format = "###,##0.00"
            .Columns("YR4").Header.Caption = YR4
            .Columns("YR4").Format = "###,##0.00"
            For i As Integer = 0 To .Columns.Count - 1
                .Columns(i).CellActivation = Activation.NoEdit
            Next i
            For Each COLNAME As String In New String() {"SEL"}
                .Columns(COLNAME).CellActivation = Activation.AllowEdit
            Next
            For Each COLNAME As String In New String() {"SEL"}
                .Columns(COLNAME).CellClickAction = CellClickAction.EditAndSelectText
            Next
        End With

        TABLE_NAME = "ARFCUSTL"

        EntryMode = "E"

        'Call Load_Record()
        Setup_Summary()
        Call Mode_Settings(True)
        Loading = False
    End Sub

    Private Sub RefreshTempTables()
        S.Length = 0
        S.AppendLine("SELECT CUST_CODE,")
        S.AppendLine("SUM(YR1) AS YR1,")
        S.AppendLine("SUM(YR2) AS YR2,")
        S.AppendLine("SUM(YR3) AS YR3,")
        S.AppendLine("SUM(YR4) AS YR4")
        S.AppendLine("FROM")
        S.AppendLine("(")
        S.AppendLine("  SELECT")
        S.AppendLine("  CUST_CODE,")
        S.AppendLine("  SUM(INV_SALES) AS YR1,")
        S.AppendLine("  0 AS YR2,")
        S.AppendLine("  0 AS YR3,")
        S.AppendLine("  0 AS YR4")
        S.AppendLine("  FROM SOTINVH1")
        S.AppendLine(String.Format("  WHERE EXTRACT(year FROM inv_date) = '{0}'", YR1))
        S.AppendLine("  GROUP BY CUST_CODE")
        S.AppendLine("  UNION")
        S.AppendLine("")
        S.AppendLine("SELECT")
        S.AppendLine("O1.CUST_CODE,")
        S.AppendLine("SUM((NVL(O2.ORDR_QTY_PICK,0) + NVL(O2.ORDR_QTY_OPEN,0)) * O2.ORDR_UNIT_PRICE) AS YR1,")
        S.AppendLine("0 AS YR2,")
        S.AppendLine("0 AS YR3,")
        S.AppendLine("0 AS YR4")
        S.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2")
        S.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
        S.AppendLine(String.Format("  AND EXTRACT(year FROM ORDR_DATE_RECD) = '{0}'", YR1))
        S.AppendLine("GROUP BY CUST_CODE")
        S.AppendLine("  UNION")
        S.AppendLine("  SELECT")
        S.AppendLine("  CUST_CODE,")
        S.AppendLine("  0 AS YR1,")
        S.AppendLine("  SUM(INV_SALES) AS YR2,")
        S.AppendLine("  0 AS YR3,")
        S.AppendLine("  0 AS YR4")
        S.AppendLine("  FROM SOTINVH1")
        S.AppendLine(String.Format("  WHERE EXTRACT(year FROM inv_date) = '{0}'", YR2))
        S.AppendLine("  GROUP BY CUST_CODE")
        S.AppendLine("  UNION")
        S.AppendLine("SELECT")
        S.AppendLine("O1.CUST_CODE,")
        S.AppendLine("0 AS YR1,")
        S.AppendLine("SUM((NVL(O2.ORDR_QTY_PICK,0) + NVL(O2.ORDR_QTY_OPEN,0)) * O2.ORDR_UNIT_PRICE) AS YR2,")
        S.AppendLine("0 AS YR3,")
        S.AppendLine("0 AS YR4")
        S.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2")
        S.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
        S.AppendLine(String.Format("  AND EXTRACT(year FROM ORDR_DATE_RECD) = '{0}'", YR2))
        S.AppendLine("GROUP BY CUST_CODE")
        S.AppendLine("  UNION")
        S.AppendLine("  SELECT")
        S.AppendLine("  CUST_CODE,")
        S.AppendLine("  0 AS YR1,")
        S.AppendLine("  0 AS YR2,")
        S.AppendLine("  SUM(INV_SALES) AS YR3,")
        S.AppendLine("  0 AS YR4")
        S.AppendLine("  FROM SOTINVH1")
        S.AppendLine(String.Format("  WHERE EXTRACT(year FROM inv_date) = '{0}'", YR3))
        S.AppendLine("  GROUP BY CUST_CODE")
        S.AppendLine("  UNION")
        S.AppendLine("  SELECT")
        S.AppendLine("  CUST_CODE,")
        S.AppendLine("  0 AS YR1,")
        S.AppendLine("  0 AS YR2,")
        S.AppendLine("  0 AS YR3,")
        S.AppendLine("  SUM(INV_SALES) AS YR4")
        S.AppendLine("  FROM SOTINVH1")
        S.AppendLine(String.Format("  WHERE EXTRACT(year FROM inv_date) = '{0}'", YR4))
        S.AppendLine("  GROUP BY CUST_CODE")
        S.AppendLine(") RSLT")
        S.AppendLine("GROUP BY CUST_CODE")
        ASCMAIN1.sql = S.ToString
        TEMP_SALES = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Create Index I_" & TEMP_SALES & "_IND on " & TEMP_SALES & " (CUST_CODE)")

        S.Length = 0
        S.AppendLine("SELECT V1.*")
        S.AppendLine("FROM ARTCUSV1 V1,")
        S.AppendLine("(")
        S.AppendLine("    SELECT")
        S.AppendLine("    CUST_CODE,")
        S.AppendLine("    MAX(ACTIVITY_NO) AS MAX_ACTIVITY_NO")
        S.AppendLine("    FROM ARTCUSV1")
        S.AppendLine("    GROUP BY CUST_CODE")
        S.AppendLine(") MX")
        S.AppendLine("WHERE V1.CUST_CODE = MX.CUST_CODE")
        S.AppendLine("AND V1.ACTIVITY_NO = MX.MAX_ACTIVITY_NO")
        ASCMAIN1.sql = S.ToString
        TEMP_MAX = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Create Index I_" & TEMP_MAX & "_IND on " & TEMP_MAX & " (CUST_CODE)")
    End Sub

    Sub Check_Inquiry_Mode()
        If InquiryMode Then
        Else
        End If
    End Sub

    Sub Check_Form_Options()
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Refresh"
            Case "Cancel"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Cancel?"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("Are You Sure You Want To Cancel?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult <> MsgBoxResult.Yes Then
                    EMsg += "Cancel Aboorted."
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
            Case "Done"
                Call Update_Record("Your Data Is Saved", False)
                Call Mode_Settings(False)
                Loading = True
                Me.Close()
                Loading = False
            Case "Cancel"
                Call Mode_Settings(False)
                Me.Close()
            Case "Refresh"
                Setup_Summary()
                UltraExplorerBar1.Groups("Screen Control").Items("Refresh").Settings.Enabled = DefaultableBoolean.False
            Case "Save"
                Update_Record("Your Data Is Saved", False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            '.Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Refresh").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Cancel").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Save").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
        End With

        UltraExplorerBar1.Groups("Customer Filters").Visible = True
        'UltraExplorerBar1.Groups("List Options").Visible = False

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        'dst.EnforceConstraints = False
        'dst.Tables("PMTVIST1").Rows.Clear()
        'dst.EnforceConstraints = True
        'Setup_Summary()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Setup_Summary()

        'Setup_SOTCSTMX()

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record(ByVal MsgToShow As String, Optional ByVal AutoSaving As Boolean = False)
        If MsgToShow.Length > 0 Then
            AutoSaving = False
        End If
        If AutoSaving Then
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Auto-Saving Your Data", "")
            Application.DoEvents()
        End If
        BeginTrans()
        CommitTrans(MsgToShow)
        If AutoSaving Then
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
            Application.DoEvents()
        End If
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
        Call Load_Popup_Menu(grdARTCUSTX, "SSB", "Show Filter", "Show GroupBox", "Customer Master File", "Customer Inquiry", "Send Email")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool
        'Dim tlb_btn As UltraWinToolbars.ButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If

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
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
            Case "Customer Master File"
                If Not IsNothing(grd.ActiveRow) Then
                    Dim CUST_CODE As String = grd.ActiveRow.Cells.Item("CUST_CODE").Text
                    If CUST_CODE.Length > 0 Then
                        Context_Launch("View", Column_Values("CUST_CODE", CUST_CODE), e.Tool.Key, "ARTCUST1")
                    End If
                End If
            Case "Customer Inquiry"
                If Not IsNothing(grd.ActiveRow) Then
                    Dim CUST_CODE As String = grd.ActiveRow.Cells.Item("CUST_CODE").Text
                    If CUST_CODE.Length > 0 Then
                        'Context_Launch("Select Customer", Column_Values("CUST_CODE", CUST_CODE), e.Tool.Key, "ARFCINQ1")
                        Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
                    End If
                End If
            Case "Remove Contact"
                If grd.Selected.Rows.Count = 0 Then
                    MsgBox("You Must Select At Least One Row To Remove", vbOKOnly, "Row Removal")
                Else
                    For Each rw As UltraGridRow In grd.Selected.Rows
                        Dim CUST_CODE As String = rw.Cells.Item("CUST_CODE").Text
                        Dim CONTACT_NO As Int64 = Val(rw.Cells.Item("CONTACT_NO").Text)
                    Next

                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Project Center"
                Dim JOB_NO As String = grd.ActiveRow.Cells("JOB_NO").Text
                Context_Launch("Edit", Column_Values("JOB_NO", JOB_NO), e.Tool.Key, "PMFJOBM1")
            Case "Show Report"
                Dim FILENAME As String = "C:\Documents and Settings\wjz\Desktop\randfromdrc\RandInvoices\310 West 52nd Street - 30760.pdf"
                Show_Document(FILENAME)

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        'Select Case Absx1.GetABSColumnName(sender)
        'Case "EMPLOYEE_CODE"
        '    If e.KeyCode = Windows.Forms.Keys.Enter Then
        '        Setup_Summary()
        '    End If
        'End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        'Select Case COLUMN_NAME
        '    Case "EMPLOYEE_CODE"
        '        Setup_Summary()
        'End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        'Select Case Absx1.GetABSColumnName(txtctl)
        '    Case "EMPLOYEE_CODE"
        '        Setup_Summary()
        'End Select
    End Sub

#End Region

    Sub Setup_Summary()
        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        Update_Record("", True)

        dst.Tables("ARTCUSTX").Rows.Clear()
        'dst.Tables("ARTCUSTD").Rows.Clear()

        dst.EnforceConstraints = False
        Fill_Records("ARTCUSTX")
        'Fill_Records("ARTCUSTD")

        FILL_ARTCUSTX_EXTRA()

        grdARTCUSTX.Update()
        grdARTCUSTX.Refresh()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub FILL_ARTCUSTX_EXTRA()
        For Each rowARTCUSTX As DataRow In dst.Tables("ARTCUSTX").Select()
            Dim YR1 As Int64 = Val(rowARTCUSTX.Item("YR1").ToString & String.Empty)
            Dim YR2 As Int64 = Val(rowARTCUSTX.Item("YR2").ToString & String.Empty)
            Dim YR3 As Int64 = Val(rowARTCUSTX.Item("YR3").ToString & String.Empty)
            Dim YR4 As Int64 = Val(rowARTCUSTX.Item("YR4").ToString & String.Empty)
            Dim SALES As String = "0"
            If (YR1 + YR2 + YR3 + YR4) > 0 Then
                SALES = "1"
            End If
            rowARTCUSTX.Item("HAS_SALES") = SALES
        Next
    End Sub

    Sub Print_Report()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'CR_params.Add("GROUP_BY", optGroupBy.Value)
        'Generate_Report("PMRVIST1", "Open Site Visit Report")
        'Print_Report_End()
    End Sub

    Private Sub Fill_Extra_Fields()
        Dim RecTotal As Int64 = dst.Tables("SOTCSTMX").Rows.Count
        Dim OnRow As Int64 = 0
        Dim PCT As String = ""
    End Sub

    Private Sub chkHideZeros_CheckedChanged(sender As Object, e As EventArgs)
        filterNonActive()
    End Sub

    Private Sub filterNonActive()
        'Dim Filter As String = ""
        'If chkHideZeros.Checked Then
        '    Filter = "ORDRED_TY <> 0 OR  SHIPPED_TY <> 0 OR CANCELLED_TY <> 0 OR ORDRED_LY <> 0 OR  SHIPPED_LY <> 0 OR CANCELLED_LY <> 0"
        'End If
        'Dim dvw As DataView = DirectCast(grdARTCUSTX.DataSource, DataTable).DefaultView
        'dvw.RowFilter = String.Format(Filter)
    End Sub

    Private Sub grdARTCUSTX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdARTCUSTX.AfterRowActivate
        Setup_ARTCUSTD()
    End Sub

    Sub Setup_ARTCUSTD()
        dst.Tables("ARTCUSTD").Rows.Clear()
        dst.Tables("ARTCUSTO").Rows.Clear()
        If Not (grdARTCUSTX.ActiveRow Is Nothing OrElse (Not grdARTCUSTX.ActiveRow.IsDataRow)) Then
            Dim CUST_CODE As String = grdARTCUSTX.ActiveRow.Cells("CUST_CODE").Value & String.Empty
            Fill_Records("ARTCUSTD", CUST_CODE)
            'Fill_Records("ARTCUSTO", CUST_CODE)
            Fill_Records("ARTCUSV1", CUST_CODE)
        End If
    End Sub

    Private Sub grdARTCUSTD_AfterRowActivate(sender As Object, e As EventArgs)

    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        If Not Loading Then
            Select Case UltraTabControl1.SelectedTab.Key
                Case "Contact Maint"
                    UltraExplorerBar1.Groups("Customer Filters").Visible = True
                Case Else
                    UltraExplorerBar1.Groups("Customer Filters").Visible = True
            End Select
        End If
    End Sub

    Private Function IsMasterDuplicate(ByVal CUST_CODE As String, ByVal CONTACT_NAME As String, ByVal CONTACT_EMAIL As String) As Boolean
        Dim RetVal As Boolean = False
        CONTACT_NAME = CONTACT_NAME.ToUpper
        CONTACT_EMAIL = CONTACT_EMAIL.ToUpper
        Dim FILTER As String = String.Format("CUST_CODE = '{0}' AND CONTACT_TYPE = 'B'", CUST_CODE)
        For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(FILTER)
            If rowARTCUSTD.Item("CONTACT_TYPE").ToString & String.Empty <> "X" Then
                If (rowARTCUSTD.Item("CONTACT_NAME").ToString & String.Empty).ToUpper = CONTACT_NAME Then
                    RetVal = True
                End If
                If (rowARTCUSTD.Item("CONTACT_EMAIL").ToString & String.Empty).ToUpper = CONTACT_EMAIL Then
                    RetVal = True
                End If
            End If
        Next
        Return RetVal
    End Function

    Private Sub cboCLIST_CODE_SelectedIndexChanged(sender As Object, e As EventArgs)
        If Not Loading Then
            RefreshList()
        End If
    End Sub

    Private Sub RefreshList()
        ASCMAIN1.Progress("Now Loading List")
        Me.Cursor = Cursors.WaitCursor
        Update_Record("", True)

        dst.EnforceConstraints = False

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub btnSendSelected_Click(sender As Object, e As EventArgs)
        Stop
        Dim emsg As New Text.StringBuilder With {.Length = 0}
        'If txtOBSendName.Text.Length = 0 Then
        '    emsg.AppendLine("Missing Sender Name.")
        'End If
        'If txtOBSendEmail.Text.Length = 0 Then
        '    emsg.AppendLine("Missing Sender Email.")
        'End If
        If emsg.Length > 0 Then
            Dim iTitle As String = "Errors"
            MsgBox(emsg.ToString(), MsgBoxStyle.OkOnly, iTitle)
        Else
            Dim content As String = MakeHTMLBody()
            Dim fileName As String = ASCMAIN1.Folders("Temp") & "ContactVerfy.html"
            If System.IO.File.Exists(fileName) Then
                System.IO.File.Delete(fileName)
            End If
            System.IO.File.WriteAllText(fileName, content)
            If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop
            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            Dim ATTACHMENTs As New Dictionary(Of String, String)
            'EMAIL_ADDRESSs.Add(txtOBSendEmail.Text, txtOBSendName.Text)
            EMAIL_ADDRESSs.Add("whr@waynerichmond.net", "Wayne Richmond")

            Dim TEMPLATE_NAME As String = "CREDIT"
            Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                 "Contact Verification Request", TEMPLATE_NAME, True, False, TEMPLATE_NAME, TEMPLATE_NAME, "Contact Verification Request", content)
            'If chkBCCAR.Checked Then
            '    Dim EMAIL_ADDRESSs_AR As New Dictionary(Of String, String)
            '    EMAIL_ADDRESSs_AR.Add("ar@regency-rib.com", "Accounts Receivable")
            '    SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
            '        (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs_AR, ATTACHMENTs,
            '        "Credit Reference Request (Copy)", TEMPLATE_NAME, True, False, TEMPLATE_NAME, TEMPLATE_NAME, "Credit Reference Request (Copy)", content)
            'End If
            MsgBox("Mail Sent", vbOKOnly, "Done")
        End If
    End Sub

    Private Function MakeHTMLBody() As String
        Dim RetVal As String
        Dim TEMPLATE As String = $"{If(ASCMAIN1.useUNCPath, ASCMAIN1.Folders("SharedRoot"), "S:")}\Archive\templates\ContactVerification.html"
        If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
            TEMPLATE = "C:\Users\Wayne\Dropbox\Regency International\Shopsite Integration\Customers\ContactVerification.html"
        End If
        Dim WEB_FIELDS As New Dictionary(Of String, String)
        'WEB_FIELDS.Add("{AccountNumber}", Absx1.txtFor("CUST_CODE").Text.ToString & String.Empty)
        'WEB_FIELDS.Add("{AccountName}", Absx1.txtFor("CUST_NAME").Text.ToString & String.Empty)
        WEB_FIELDS.Add("{SendName}", "Wayne Richmond")
        WEB_FIELDS.Add("{SendEmail}", "whr@waynerichmond.net")

        Dim tableHtml As New System.Text.StringBuilder()

        tableHtml.Append("<table style='border-collapse:collapse; font-family:Arial, sans-serif; font-size:13px; margin:0; padding:0;'>")
        ' Header row
        tableHtml.Append("<tr style='background-color:#f2f2f2;'>")
        tableHtml.Append("<th style='border:1px solid #ccc; padding:6px;'>No</th>")
        tableHtml.Append("<th style='border:1px solid #ccc; padding:6px;'>Contact Type</th>")
        tableHtml.Append("<th style='border:1px solid #ccc; padding:6px;'>Name</th>")
        tableHtml.Append("<th style='border:1px solid #ccc; padding:6px;'>Title</th>")
        tableHtml.Append("<th style='border:1px solid #ccc; padding:6px;'>Email</th>")
        tableHtml.Append("<th style='border:1px solid #ccc; padding:6px;'>Phone</th>")
        tableHtml.Append("<th style='border:1px solid #ccc; padding:6px;'>Cell</th>")
        tableHtml.Append("</tr>")
        ' Row 1
        tableHtml.Append("<tr>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>1</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>Buyer</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>CHRIS LIGHT</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>owner</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>info@nygardenworld.com</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'></td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'></td>")
        tableHtml.Append("</tr>")
        ' Row 2
        tableHtml.Append("<tr>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>2</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>Buyer</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>teresa</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'></td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'></td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'></td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'></td>")
        tableHtml.Append("</tr>")
        ' Row 3
        tableHtml.Append("<tr>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>3</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>Other</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>CLAIRE</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>Main Contact</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>russ@nygardenworld.com</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>7182246789</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'></td>")
        tableHtml.Append("</tr>")
        ' Row 4
        tableHtml.Append("<tr>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>4</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>Buyer</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>James ODriscoll</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'></td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>JAMESOD@NYGARDENWORLD.COM</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'>3477682142</td>")
        tableHtml.Append("<td style='border:1px solid #ccc; padding:6px;'></td>")
        tableHtml.Append("</tr>")
        tableHtml.Append("</table>")

        Dim html As String = tableHtml.ToString()
        ' Final HTML string:
        Dim datatable As String = tableHtml.ToString()
        'datatable
        WEB_FIELDS.Add("{datatable}", datatable)

        Dim BodyContent As String = System.IO.File.ReadAllText(TEMPLATE)
        BodyContent = BodyContent.Replace(vbCrLf, "")
        For Each WEB_FIELD As KeyValuePair(Of String, String) In WEB_FIELDS
            BodyContent = BodyContent.Replace(WEB_FIELD.Key, WEB_FIELD.Value)
        Next
        RetVal = BodyContent

        Return RetVal
    End Function

    Private Sub grdARTCUSTX_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdARTCUSTX.ClickCell
        If e.Cell.Column.Key = "SEL" Then
            If e.Cell.Row.Cells("SEL").Value = "1" Then
                e.Cell.Row.Cells("SEL").Value = "0"
            Else
                e.Cell.Row.Cells("SEL").Value = "1"
            End If
        End If
    End Sub

    Private Sub btnEditContacts_Click(sender As Object, e As EventArgs) Handles btnEditContacts.Click
        If EditingContacts Then
            EditingContacts = False
            btnEditContacts.Text = "Edit Contacts"
        Else
            grdARTCUSTX.ActiveRow.Selected = True
            EditingContacts = True
            btnEditContacts.Text = "Save Changes"
        End If
        grdARTCUSTX.Enabled = Not EditingContacts
        With grdARTCUSTD.DisplayLayout
            For i As Integer = 0 To .Bands(0).Columns.Count - 1
                .Bands(0).Columns(i).CellActivation = Activation.NoEdit
            Next i
            If EditingContacts Then
                .Override.AllowAddNew = AllowAddNew.FixedAddRowOnTop
                .Override.AllowDelete = DefaultableBoolean.True
                .Override.AllowUpdate = DefaultableBoolean.True
                For Each COLNAME As String In New String() {"CONTACT_NAME", "CONTACT_TITLE", "CONTACT_PHONE", "CONTACT_EXT", "CONTACT_TYPE", "CONTACT_PRIMARY", "CONTACT_CELL"}
                    .Bands(0).Columns(COLNAME).CellActivation = Activation.AllowEdit
                Next
                For Each COLNAME As String In New String() {"CONTACT_NAME", "CONTACT_TITLE", "CONTACT_PHONE", "CONTACT_EXT", "CONTACT_TYPE", "CONTACT_PRIMARY", "CONTACT_CELL"}
                    .Bands(0).Columns(COLNAME).CellClickAction = CellClickAction.EditAndSelectText
                Next
            Else
                .Override.AllowAddNew = AllowAddNew.No
                .Override.AllowDelete = DefaultableBoolean.False
                .Override.AllowUpdate = DefaultableBoolean.True
            End If
        End With
        grdARTCUSTD.Update()
    End Sub

#Region "Space Code"
    'Private Sub btnMakeMasterContacts_Click(sender As Object, e As EventArgs)
    '    Dim iResult As MsgBoxResult
    '    Dim iTitle As String = "Make Masterfile Contacts?"
    '    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
    '    iMSG.AppendLine("Are You Sure You Want To Do This?")
    '    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
    '    If iResult <> MsgBoxResult.Yes Then
    '        Exit Sub
    '    End If

    '    Setup_Summary()
    '    Me.Cursor = Cursors.WaitCursor
    '    For Each rowARTCUSTX As DataRow In dst.Tables("ARTCUSTX").Select()
    '        Dim CUST_CODE As String = rowARTCUSTX.Item("CUST_CODE").ToString & String.Empty
    '        Dim FILTER As String = String.Format("CUST_CODE = '{0}'", CUST_CODE)
    '        If dst.Tables.Item("ARTCUSTD").Select(FILTER).Count = 0 Then
    '            AddMasterfileContact(rowARTCUSTX, False)
    '        Else
    '            Dim MFNAME As String = rowARTCUSTX.Item("CUST_CONTACT").ToString & String.Empty
    '            If MFNAME.Length > 0 Then
    '                MFNAME = MFNAME.ToUpper
    '                Dim MFFOUND As Boolean = False
    '                For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(FILTER)
    '                    Dim CTNAME As String = rowARTCUSTD.Item("CONTACT_NAME").ToString & String.Empty
    '                    If CTNAME.Length > 0 Then
    '                        CTNAME = CTNAME.ToUpper
    '                    End If
    '                    If MFNAME = CTNAME Then
    '                        MFFOUND = True
    '                    End If
    '                Next
    '                If Not MFFOUND Then
    '                    AddMasterfileContact(rowARTCUSTX, False)
    '                End If
    '            End If
    '        End If
    '    Next
    '    Me.Cursor = Cursors.Default
    'End Sub

    'Private Sub AddMasterfileContact(ByRef rowARTCUSTX As DataRow, ByVal isPrimary As Boolean)
    '    Dim CUST_CODE As String = rowARTCUSTX.Item("CUST_CODE").ToString & String.Empty
    '    Dim FILTER As String = String.Format("CUST_CODE = '{0}'", CUST_CODE)
    '    Dim CONTACT_NO As Long = 1
    '    For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(FILTER, "CONTACT_NO")
    '        If Val(rowARTCUSTD.Item("CONTACT_NO").ToString & String.Empty) >= CONTACT_NO Then
    '            CONTACT_NO = Val(rowARTCUSTD.Item("CONTACT_NO").ToString & String.Empty) + 1
    '        End If
    '    Next

    '    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)

    '    Dim newARTCUSTD As DataRow = dst.Tables("ARTCUSTD").NewRow
    '    newARTCUSTD.Item("CUST_CODE") = CUST_CODE
    '    newARTCUSTD.Item("CONTACT_NO") = CONTACT_NO
    '    newARTCUSTD.Item("CONTACT_NAME") = rowARTCUST1.Item("CUST_CONTACT").ToString & String.Empty
    '    newARTCUSTD.Item("CONTACT_TITLE") = "Master Contact"
    '    newARTCUSTD.Item("CONTACT_EMAIL") = rowARTCUST1.Item("CUST_EMAIL").ToString & String.Empty
    '    newARTCUSTD.Item("CONTACT_PHONE") = rowARTCUST1.Item("CUST_PHONE").ToString & String.Empty
    '    newARTCUSTD.Item("CONTACT_EXT") = rowARTCUST1.Item("CUST_EXT").ToString & String.Empty
    '    newARTCUSTD.Item("CONTACT_FAX") = rowARTCUST1.Item("CUST_FAX").ToString & String.Empty
    '    newARTCUSTD.Item("CONTACT_TYPE") = "X"
    '    If isPrimary Then
    '        newARTCUSTD.Item("CONTACT_PRIMARY") = "1"
    '    Else
    '        newARTCUSTD.Item("CONTACT_PRIMARY") = "0"
    '    End If
    '    newARTCUSTD.Item("CONTACT_NOTE") = "Added By Cont Maint"
    '    newARTCUSTD.Item("INIT_OPER") = ASCMAIN1.USER_ID
    '    newARTCUSTD.Item("LAST_DATE") = DATETIME_STAMP
    '    newARTCUSTD.Item("LAST_OPER") = ASCMAIN1.USER_ID
    '    newARTCUSTD.Item("INIT_DATE") = DATETIME_STAMP

    '    newARTCUSTD.Item("CONTACT_CELL") = Null
    '    dst.Tables("ARTCUSTD").Rows.Add(newARTCUSTD)
    'End Sub

    'Private Sub btnBuyerGroups_Click(sender As Object, e As EventArgs) Handles btnBuyerGroups.Click
    '    'Stop
    '    'Dim CUST_CODE As String = "200138"
    '    'Dim sql As New Text.StringBuilder With {.Length = 0}
    '    'Dim FILTER As String = String.Format("CUST_CODE = '{0}'", CUST_CODE)
    '    'Dim CONTACT_NO As Long = 1
    '    'For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(FILTER, "CONTACT_NO")
    '    '    If Val(rowARTCUSTD.Item("CONTACT_NO").ToString & String.Empty) >= CONTACT_NO Then
    '    '        CONTACT_NO = Val(rowARTCUSTD.Item("CONTACT_NO").ToString & String.Empty) + 1
    '    '    End If
    '    'Next

    '    'sql.AppendLine("SELECT *")
    '    'sql.AppendLine("FROM ARTCUST2")
    '    'sql.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
    '    'Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString())
    '    'For Each rowARTCUST2 As DataRow In tbl.Rows
    '    '    Dim newARTCUSTD As DataRow = dst.Tables("ARTCUSTD").NewRow
    '    '    newARTCUSTD.Item("CUST_CODE") = CUST_CODE
    '    '    newARTCUSTD.Item("CONTACT_NO") = CONTACT_NO
    '    '    newARTCUSTD.Item("CONTACT_NAME") = rowARTCUST1.Item("CUST_CONTACT").ToString & String.Empty
    '    '    newARTCUSTD.Item("CONTACT_TITLE") = "Ship To Contact"
    '    '    newARTCUSTD.Item("CONTACT_EMAIL") = rowARTCUST1.Item("CUST_EMAIL").ToString & String.Empty
    '    '    newARTCUSTD.Item("CONTACT_PHONE") = rowARTCUST1.Item("CUST_PHONE").ToString & String.Empty
    '    '    newARTCUSTD.Item("CONTACT_EXT") = rowARTCUST1.Item("CUST_EXT").ToString & String.Empty
    '    '    newARTCUSTD.Item("CONTACT_FAX") = rowARTCUST1.Item("CUST_FAX").ToString & String.Empty
    '    '    newARTCUSTD.Item("CONTACT_TYPE") = "X"
    '    '    newARTCUSTD.Item("CONTACT_PRIMARY") = "0"
    '    '    newARTCUSTD.Item("CONTACT_NOTE") = "Added By Cont Maint"
    '    '    newARTCUSTD.Item("INIT_OPER") = ASCMAIN1.USER_ID
    '    '    newARTCUSTD.Item("LAST_DATE") = DATETIME_STAMP
    '    '    newARTCUSTD.Item("LAST_OPER") = ASCMAIN1.USER_ID
    '    '    newARTCUSTD.Item("INIT_DATE") = DATETIME_STAMP
    '    '    newARTCUSTD.Item("CONTACT_CELL") = Null
    '    '    dst.Tables("ARTCUSTD").Rows.Add(newARTCUSTD)
    '    'Next
    'End Sub
#End Region
End Class