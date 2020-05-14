Public Class SAFSLST1

    Dim SATSLST0 As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -36, 0, -11)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -36, 0, 0)

        With dst

            Create_SATSLST0("", "")
            ASCMAIN1.sql = "Select * from " & SATSLST0
            Create_TDA(.Tables.Add, "SATSLST1", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select SATSLST0.STATE_CODE, TATSTATE.STATE_NAME" _
                & ", SUM (SATSLST0.GRS) GRS, SUM (SATSLST0.CRD) CRD, SUM (SATSLST0.INV_SALES) INV_SALES" _
                & " from TATSTATE, " & SATSLST0 & " SATSLST0 where TATSTATE.STATE_CODE (+) = SATSLST0.STATE_CODE" _
                & " group by SATSLST0.STATE_CODE, TATSTATE.STATE_NAME"
            Create_TDA(.Tables.Add, "SATSLST0", "**", 0, False, "", 1)
        End With

        grdSATSLST1.DataSource = dst.Tables("SATSLST1")
        Create_Summary(grdSATSLST1, "STATE_CODE", "Count")
        Create_Summary(grdSATSLST1, New String() {"GRS", "CRD", "INV_SALES"})

        grdSATSLST0.DataSource = dst.Tables("SATSLST0")
        Create_Summary(grdSATSLST0, "STATE_CODE", "Count")
        Create_Summary(grdSATSLST0, New String() {"GRS", "CRD", "INV_SALES"})

        spl.Panel1Collapsed = True

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
            End With

        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        splSales.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SATSLST0", "SATSLST1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Reading from Sales History Data")

        Save_Header_Fields(UltraGroupBox1)

        Dim RYP0 As String = Absx1.cmbFor("RYP0").Value
        Dim RYP1 As String = Absx1.cmbFor("RYP1").Value

        Create_SATSLST0(RYP0, RYP1)

        ASCMAIN1.Progress("Now Loading Data")

        EnforceConstraints(False)

        Fill_Records("SATSLST0")
        Fill_Records("SATSLST1")

        EnforceConstraints(True)

        Sort_grdColumns(grdSATSLST0, "STATE_CODE")
        grdSATSLST0.Text = "Sales by State from " & Absx1.cmbFor("RYP0").Text & " thru " & Absx1.cmbFor("RYP0").Text
        Setup_grdSATSLST1()

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
        Load_Popup_Menu(grdSATSLST0, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSATSLST1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Customer Inquiry")
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

    Sub Create_SATSLST0(ByVal FYP As String, ByVal RYP As String)

        ASCMAIN1.sql = "SELECT SOTINVH1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.CUST_STATE" & vbCrLf _
            & ", NVL(NVL (SOTORDR5.CUST_STATE,ARTCUST1.CUST_STATE),'?') STATE_CODE" & vbCrLf _
            & ", SUM (DECODE(INV_TYPE,'I',INV_SALES,0)) GRS" & vbCrLf _
            & ", SUM (DECODE(INV_TYPE,'C',INV_SALES,0)) CRD" & vbCrLf _
            & ", SUM (INV_SALES) INV_SALES" & vbCrLf _
            & " from SOTINVH1,SOTORDR5,ARTCUST1" & vbCrLf _
            & " where SOTINVH1. ORDR_YYYYPP_UPDATED between '" & FYP & "' and '" & RYP & "'" & vbCrLf _
            & "   and SOTORDR5.ORDR_NO (+) = SOTINVH1.ORDR_NO" & vbCrLf _
            & "   and SOTORDR5.CUST_ADDR_CODE (+) = 'ST'" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE" & vbCrLf _
            & " group by SOTINVH1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.CUST_STATE" & vbCrLf _
            & ", NVL(NVL (SOTORDR5.CUST_STATE,ARTCUST1.CUST_STATE),'?')"

        If SATSLST0 = "" Then
            SATSLST0 = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SATSLST0)
            ASCDATA1.ExecuteSQL("Insert into " & SATSLST0 & " " & ASCMAIN1.sql)
        End If
    End Sub

    Sub Print_Report()
        Dim SUBT As String = ""
        Print_Report_Begin()
        Print_Report_End()
    End Sub

    Private Sub grdSATSLST1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSATSLST1.InitializeLayout

    End Sub

    Private Sub grdSATSLST1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATSLST1.InitializeRow
        If e.Row.IsDataRow And Not e.Row.IsFilterRow Then
            If e.Row.Cells("STATE_CODE").Value & "" <> e.Row.Cells("CUST_STATE").Value & "" Then
                e.Row.Cells("CUST_STATE").Appearance.ForeColor = Drawing.Color.Red
            End If
        End If
    End Sub

    Private Sub grdSATSLST0_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSATSLST0.AfterRowActivate
        Setup_grdSATSLST1()
    End Sub

    Sub Setup_grdSATSLST1()
        If grdSATSLST0.ActiveRow Is Nothing Then
            grdSATSLST1.Visible = False
        Else
            Dim STATE_CODE As String = grdSATSLST0.ActiveRow.Cells("STATE_CODE").Value
            Dim dvw As DataView = DirectCast(grdSATSLST1.DataSource, DataTable).DefaultView
            dvw.RowFilter = "STATE_CODE = '" & STATE_CODE & "'"
            Sort_grdColumns(grdSATSLST1, "CUST_CODE")
            grdSATSLST1.Visible = True
            grdSATSLST1.Text = "Sales by Customer within State " & STATE_CODE
        End If
    End Sub
End Class