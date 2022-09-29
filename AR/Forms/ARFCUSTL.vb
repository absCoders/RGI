Imports System.Net.Mail
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Xml
Imports Infragistics.Documents.Excel
Imports Infragistics.Win.UltraWinGrid
Imports Infragistics.Win.UltraWinSchedule

Public Class ARFCUSTL
    Dim Remote As New REMOTE(Me)
    Dim S As New System.Text.StringBuilder With {.Length = 0}
    Dim Loading As Boolean = True
    Dim TEMP_SALES As String = ""
    Dim YR1 As String = ""
    Dim YR2 As String = ""
    Dim YR3 As String = ""
    Dim YR4 As String = ""


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Check_Form_Options()

        Dim BaseYear As Int64 = Now().Year
        YR1 = (BaseYear).ToString
        YR2 = (BaseYear - 1).ToString
        YR3 = (BaseYear - 2).ToString
        YR4 = (BaseYear - 3).ToString

        'Fill In The Gaps
        S.AppendLine("INSERT INTO ARTCUSTL")
        S.AppendLine("SELECT")
        S.AppendLine("ARTCUSTD.CUST_CODE,")
        S.AppendLine("ARTCUSTD.CONTACT_NO,")
        S.AppendLine("ARTCLST1.CLIST_CODE,")
        S.AppendLine("'0' AS CLIST_ACTIVE,")
        S.AppendLine("'wayne' AS INIT_OPER,")
        S.AppendLine("SYSDATE AS INIT_DATE,")
        S.AppendLine("'wayne' AS LAST_OPER,")
        S.AppendLine("SYSDATE AS LAST_DATE")
        S.AppendLine("FROM ARTCUSTD, ARTCLST1")
        S.AppendLine("WHERE (ARTCUSTD.CUST_CODE, ARTCUSTD.CONTACT_NO, ARTCLST1.CLIST_CODE)")
        S.AppendLine("NOT IN")
        S.AppendLine("(")
        S.AppendLine("  SELECT")
        S.AppendLine("  CUST_CODE,")
        S.AppendLine("  CONTACT_NO,")
        S.AppendLine("  CLIST_CODE")
        S.AppendLine("  FROM ARTCUSTL")
        S.AppendLine(")")
        ASCMAIN1.sql = S.ToString
        ASCDATA1.ExecuteSQL()

        RefreshSalesTempTable()

        With dst
            S.Length = 0
            S.AppendLine("SELECT")
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
            S.AppendLine("SUM(SALES.YR1) AS YR1,")
            S.AppendLine("SUM(SALES.YR2) AS YR2,")
            S.AppendLine("SUM(SALES.YR3) AS YR3,")
            S.AppendLine("SUM(SALES.YR4) AS YR4")
            S.AppendLine($"FROM ARTCUST1, {TEMP_SALES} SALES")
            S.AppendLine("WHERE ARTCUST1.CUST_CODE = SALES.CUST_CODE (+)")
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
            S.AppendLine("ARTCUST1.INIT_DATE")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ARTCUSTX", "**", 0, False)
            'With .Tables("ARFCUSTX").Columns
            '    .Add("XXX", GetType(String))
            'End With

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ARTCUSTD")
            'S.AppendLine("WHERE CONTACT_TYPE = 'B'")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ARTCUSTD", "*", 0, True)

            'Fill_Records("ARTCUSTD")

            S.Length = 0
            S.AppendLine("SELECT ARTCUSTL.*, ARTCLST1.CLIST_DESC")
            S.AppendLine("FROM ARTCUSTL, ARTCLST1")
            S.AppendLine("WHERE ARTCUSTL.CLIST_CODE = ARTCLST1.CLIST_CODE")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ARTCUSTL", "**", 0, True)
            With .Tables("ARTCUSTL").Columns
                .Add("CNT", GetType(Int64))
                .Add("CLIST_ACTIVE_TMP", GetType(String))
            End With

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ARTCLST1")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ARTCLST1", "**", 0, False)
            Create_TDA(.Tables.Add, "ARTCLSTF", "**", 0, False)
            Create_TDA(.Tables.Add, "ARTCLSTT", "**", 0, False)
            Fill_Records("ARTCLST1")
            Fill_Records("ARTCLSTF")
            Fill_Records("ARTCLSTT")

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("C1.CUST_CODE,")
            S.AppendLine("C1.CUST_NAME,")
            S.AppendLine("C1.CUST_ADDR1,")
            S.AppendLine("C1.CUST_ADDR2,")
            S.AppendLine("C1.CUST_ADDR3,")
            S.AppendLine("C1.CUST_CITY,")
            S.AppendLine("C1.CUST_STATE,")
            S.AppendLine("C1.CUST_ZIP_CODE,")
            S.AppendLine("C1.CUST_COUNTRY,")
            S.AppendLine("C1.INIT_DATE,")
            S.AppendLine("C1.SREP_CODE,")
            S.AppendLine("CD.CONTACT_NAME,")
            S.AppendLine("CD.CONTACT_TITLE,")
            S.AppendLine("CD.CONTACT_EMAIL,")
            S.AppendLine("CD.CONTACT_TYPE,")
            S.AppendLine("CD.CONTACT_PRIMARY,")
            S.AppendLine("CD.CONTACT_NO,")
            S.AppendLine("CL.CLIST_ACTIVE,")
            S.AppendLine("C1.CUST_STATUS,")
            S.AppendLine("SUM(SALES.YR1) AS YR1,")
            S.AppendLine("SUM(SALES.YR2) AS YR2,")
            S.AppendLine("SUM(SALES.YR3) AS YR3,")
            S.AppendLine("SUM(SALES.YR4) AS YR4")
            S.AppendLine($"FROM ARTCUST1 C1, ARTCUSTD CD, ARTCUSTL CL, {TEMP_SALES} SALES")
            S.AppendLine("WHERE C1.CUST_CODE = SALES.CUST_CODE (+)")
            S.AppendLine("AND C1.CUST_CODE = CD.CUST_CODE")
            S.AppendLine("AND CD.CUST_CODE = CL.CUST_CODE")
            S.AppendLine("AND CD.CONTACT_NO = CL.CONTACT_NO")
            S.AppendLine("AND CL.CLIST_CODE = :PARM1")
            S.AppendLine("GROUP BY")
            S.AppendLine("C1.CUST_CODE,")
            S.AppendLine("C1.CUST_NAME,")
            S.AppendLine("C1.CUST_ADDR1,")
            S.AppendLine("C1.CUST_ADDR2,")
            S.AppendLine("C1.CUST_ADDR3,")
            S.AppendLine("C1.CUST_CITY,")
            S.AppendLine("C1.CUST_STATE,")
            S.AppendLine("C1.CUST_ZIP_CODE,")
            S.AppendLine("C1.CUST_COUNTRY,")
            S.AppendLine("C1.INIT_DATE,")
            S.AppendLine("C1.SREP_CODE,")
            S.AppendLine("CD.CONTACT_NAME,")
            S.AppendLine("CD.CONTACT_TITLE,")
            S.AppendLine("CD.CONTACT_EMAIL,")
            S.AppendLine("CD.CONTACT_TYPE,")
            S.AppendLine("CD.CONTACT_PRIMARY,")
            S.AppendLine("CD.CONTACT_NO,")
            S.AppendLine("CL.CLIST_ACTIVE,")
            S.AppendLine("C1.CUST_STATUS")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ARTLIST", "**", 0, False, "V")
            With .Tables("ARTLIST").Columns
                '.Add("YR1", GetType(Double))
                '.Add("YR2", GetType(Double))
                '.Add("YR3", GetType(Double))
                '.Add("YR4", GetType(Double))
                .Add("CLIST_ACTIVE_TMP", GetType(String))
                .Add("YRT", GetType(Double), "YR1 + YR2 + YR3 + YR4")
            End With
        End With

        grdARTCUSTX.DataSource = dst.Tables("ARTCUSTX")
        grdARTCUSTL.DataSource = dst.Tables("ARTCUSTL")
        grdARTCUSTD.DataSource = dst.Tables("ARTCUSTD")
        grdARTLIST.DataSource = dst.Tables("ARTLIST")

        ASCMAIN1.Add_Value_List(grdARTCUSTD, "CONTACT_TYPE")
        ASCMAIN1.Add_Value_List(grdARTLIST, "CONTACT_TYPE")

        ASCMAIN1.Add_Value_List(grdARTLIST, "CUST_STATUS", , New String() {":", "A:Active", "I:Inactive", "C:Credit"})

        Create_Summary(grdARTCUSTX, "CUST_CODE", "Count")
        Create_Summary(grdARTCUSTX, "YR1", "Sum")
        Create_Summary(grdARTCUSTX, "YR2", "Sum")
        Create_Summary(grdARTCUSTX, "YR3", "Sum")
        Create_Summary(grdARTCUSTX, "YR4", "Sum")

        Create_Summary(grdARTLIST, "CUST_CODE", "Count")

        'Create_Summary(grdSOTORDR2, New String() {"ORDR_QTY", "ORDR_AMT", "TCUFT", "ORDR_QTY_ALLO", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_AMT_OPEN", "ORDR_AMT_ALLO", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"})

        Sort_grdColumns(grdARTCUSTL, "CLIST_CODE", False)

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
        End With

        With grdARTCUSTL.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.True
            For i As Integer = 0 To .Bands(0).Columns.Count - 1
                .Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
            Next i
            'For Each COLNAME As String In New String() {"CLIST_ACTIVE"}
            '    .Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
            '    .Bands(0).Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
            'Next
        End With

        With grdARTLIST.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.True
            For i As Integer = 0 To .Bands(0).Columns.Count - 1
                .Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
            Next i
            'For Each COLNAME As String In New String() {"CLIST_ACTIVE"}
            '    .Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
            '    .Bands(0).Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
            'Next
        End With

        With grdARTLIST.DisplayLayout.Bands(0)
            .Columns("YR1").Header.Caption = YR1
            .Columns("YR1").Format = "###,##0.00"
            .Columns("YR2").Header.Caption = YR2
            .Columns("YR2").Format = "###,##0.00"
            .Columns("YR3").Header.Caption = YR3
            .Columns("YR3").Format = "###,##0.00"
            .Columns("YR4").Header.Caption = YR4
            .Columns("YR4").Format = "###,##0.00"
        End With

        TABLE_NAME = "ARFCUSTL"

        EntryMode = "E"

        Dim lstCLIST_CODE As New Dictionary(Of String, String)
        For Each rowARTCLST1 As DataRow In dst.Tables.Item("ARTCLST1").Select("", "CLIST_CODE")
            lstCLIST_CODE.Add(rowARTCLST1.Item("CLIST_CODE").ToString & String.Empty, rowARTCLST1.Item("CLIST_DESC").ToString & String.Empty)
        Next

        cboCLIST_CODE.DataSource = dst.Tables("ARTCLST1")
        cboCLIST_CODE.ValueMember = "CLIST_CODE"
        cboCLIST_CODE.DisplayMember = "CLIST_DESC"

        cboCopyToList.DataSource = dst.Tables("ARTCLSTT")
        cboCopyToList.ValueMember = "CLIST_CODE"
        cboCopyToList.DisplayMember = "CLIST_DESC"

        cboCopyFromList.DataSource = dst.Tables("ARTCLSTF")
        cboCopyFromList.ValueMember = "CLIST_CODE"
        cboCopyFromList.DisplayMember = "CLIST_DESC"

        'cboCopyListTo.DataSource = dst.Tables("ARTCLST1")
        'cboCopyListTo.ValueMember = "CLIST_CODE"
        'cboCopyListTo.DisplayMember = "CLIST_DESC"

        numSalesGreater.Value = 0

        dteInitDate.DateTime = Now()

        'Call Load_Record()
        Call Mode_Settings(True)
        Loading = False
    End Sub

    Private Sub RefreshSalesTempTable()
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
            Case "Load"
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
                Call Update_Record()
                Call Mode_Settings(False)
                UltraTabControl1.Tabs.Item("Data Maint").Visible = False
                UltraTabControl1.Tabs.Item("List Maint").Visible = False
                Loading = True
                Me.Close()
                Loading = False
            Case "Cancel"
                Call Mode_Settings(False)
                UltraTabControl1.Tabs.Item("Data Maint").Visible = False
                UltraTabControl1.Tabs.Item("List Maint").Visible = False
                Me.Close()
            Case "Load"
                Setup_Summary()
                UltraTabControl1.Tabs.Item("Data Maint").Visible = True
                UltraTabControl1.Tabs.Item("List Maint").Visible = True
                UltraExplorerBar1.Groups("Screen Control").Items("Load").Settings.Enabled = DefaultableBoolean.False
            Case "Save"
                Update_Record()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            '.Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Cancel").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Save").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
        End With

        UltraExplorerBar1.Groups("Customer Options").Visible = True
        UltraExplorerBar1.Groups("List Options").Visible = False

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

    Sub Update_Record(Optional ByVal showMsg As Boolean = True)
        Dim msg As String = "Records Updated"
        If Not showMsg Then
            msg = ""
        End If
        BeginTrans()
        'INIT_LAST("PMTVIST1", True, "", True)
        Update_Record_TDA("ARTCUSTL")
        'Update_Record_TDA("ARTCUSTD")
        CommitTrans(msg)
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
        Call Load_Popup_Menu(grdARTCUSTX, "SSB", "Show Filter", "Show GroupBox", "Customer Master File", "Customer Inquiry")
        Call Load_Popup_Menu(grdARTCUSTL, "SSB", "Show Filter", "Show GroupBox")
        Call Load_Popup_Menu(grdARTLIST, "SSB", "Show Filter", "Show GroupBox", "Remove Contact", "Customer Master File", "Customer Inquiry")
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
                        Dim CLIST_CODE As String = cboCLIST_CODE.SelectedValue
                        Dim FILTER As String = String.Format("CUST_CODE = '{0}' AND CONTACT_NO = {1} AND CLIST_CODE = '{2}'", CUST_CODE, CONTACT_NO, CLIST_CODE)
                        Dim rowARTCUSTL As DataRow = dst.Tables("ARTCUSTL").Select(FILTER).FirstOrDefault
                        If Not IsNothing(rowARTCUSTL) Then
                            rowARTCUSTL.Item("CLIST_ACTIVE") = "0"
                            rw.Cells.Item("CLIST_ACTIVE").Value = "0"
                            rw.Cells.Item("CLIST_ACTIVE_TMP").Value = "1"
                        End If
                    Next
                    grdARTLIST.UpdateData()
                    grdARTLIST.Refresh()
                    ListActiveOnly()

                    'If Not IsNothing(grdARTLIST.ActiveRow) Then
                    '    Dim CUST_CODE As String = grdARTLIST.ActiveRow.Cells.Item("CUST_CODE").Text
                    '    Dim CONTACT_NO As Int64 = Val(grdARTLIST.ActiveRow.Cells.Item("CONTACT_NO").Text)
                    '    Dim CLIST_CODE As String = cboCLIST_CODE.SelectedValue
                    '    Dim FILTER As String = String.Format("CUST_CODE = '{0}' AND CONTACT_NO = {1} AND CLIST_CODE = '{2}'", CUST_CODE, CONTACT_NO, CLIST_CODE)
                    '    Dim rowARTCUSTL As DataRow = dst.Tables("ARTCUSTL").Select(FILTER).FirstOrDefault
                    '    If Not IsNothing(rowARTCUSTL) Then
                    '        rowARTCUSTL.Item("CLIST_ACTIVE") = "0"
                    '        grdARTLIST.ActiveRow.Cells.Item("CLIST_ACTIVE").Value = "0"
                    '        ListActiveOnly()
                    '        grdARTLIST.Refresh()
                    '    End If
                    'End If
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
        Update_Record(False)

        dst.Tables("ARTCUSTX").Rows.Clear()
        dst.Tables("ARTCUSTD").Rows.Clear()
        dst.Tables("ARTCUSTL").Rows.Clear()

        dst.EnforceConstraints = False
        Fill_Records("ARTCUSTX")
        Fill_Records("ARTCUSTL")
        Fill_Records("ARTCUSTD")

        grdARTCUSTX.Update()
        grdARTCUSTX.Refresh()

        If UltraTabControl1.SelectedTab.Key = "List Maint" Then
            Fill_Records("ARTLIST", cboCLIST_CODE.SelectedValue)
            grdARTLIST.Text = cboCLIST_CODE.Text
            'AddSalesToList()
            ListActiveOnly()
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub AddSalesToList()
        For Each rowARTLIST As DataRow In dst.Tables("ARTLIST").Select()
            Dim CUST_CODE As String = rowARTLIST.Item("CUST_CODE")
            Dim XFILTER As String = String.Format("CUST_CODE = '{0}'", CUST_CODE)
            Dim rowARTCUSTX As DataRow = dst.Tables.Item("ARTCUSTX").Select(XFILTER).FirstOrDefault
            If Not IsNothing(rowARTCUSTX) Then
                For i As Integer = 1 To 4
                    rowARTLIST.Item("YR" & i) = rowARTCUSTX.Item("YR" & i)
                Next
            End If
        Next
        grdARTLIST.UpdateData()
        grdARTLIST.Refresh()
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

    Sub Setup_ARTCUSTD()
        Dim filter As String = ""
        Dim dvw As DataView = DirectCast(grdARTCUSTD.DataSource, DataTable).DefaultView
        If grdARTCUSTX.ActiveRow Is Nothing OrElse (Not grdARTCUSTX.ActiveRow.IsDataRow) Then
            dvw.RowFilter = "CUST_CODE = 'X'"
        Else
            Dim CUST_CODE As String = grdARTCUSTX.ActiveRow.Cells("CUST_CODE").Value & String.Empty
            If chkOnlyBuyers.Checked Then
                dvw.RowFilter = String.Format("CUST_CODE = '{0}' and CONTACT_TYPE = 'B'", CUST_CODE)
            Else
                dvw.RowFilter = String.Format("CUST_CODE = '{0}'", CUST_CODE)
            End If
            If dvw.Count = 0 Then
                Setup_ARTCUSTL()
            End If
            ' grdSOTORDR3.Text = "Customer Style / Color Details for Order Line " & CStr(ORDR_LNO)
        End If
    End Sub

    Sub Setup_ARTCUSTL()
        Dim dvw As DataView = DirectCast(grdARTCUSTL.DataSource, DataTable).DefaultView
        If grdARTCUSTD.ActiveRow Is Nothing OrElse (Not grdARTCUSTD.ActiveRow.IsDataRow) Then
            dvw.RowFilter = "CUST_CODE = 'X'"
        Else
            Dim CUST_CODE As String = grdARTCUSTD.ActiveRow.Cells("CUST_CODE").Value & String.Empty
            Dim CONTACT_NO As Integer = Val(grdARTCUSTD.ActiveRow.Cells("CONTACT_NO").Value)
            Dim FLT As String = String.Format("CUST_CODE = '{0}' AND CONTACT_NO = {1}", CUST_CODE, CONTACT_NO)
            dvw.RowFilter = FLT
            SetListCounts()
        End If
    End Sub

    Private Sub SetListCounts()
        If Not IsNothing(grdARTCUSTD.ActiveRow) Then
            Dim CUST_CODE As String = grdARTCUSTD.ActiveRow.Cells("CUST_CODE").Value & String.Empty
            Dim CONTACT_NO As Integer = Val(grdARTCUSTD.ActiveRow.Cells("CONTACT_NO").Value)
            Dim FLT As String = String.Format("CUST_CODE = '{0}' AND CONTACT_NO = {1}", CUST_CODE, CONTACT_NO)
            For Each rowARTCUSTL As DataRow In dst.Tables("ARTCUSTL").Select(FLT)
                'rowARTCUSTL.Item("CNT") = 3
                Dim CLIST_CODE As String = rowARTCUSTL.Item("CLIST_CODE").ToString & String.Empty
                Dim FTR As String = $"CLIST_CODE = '{CLIST_CODE}' AND CLIST_ACTIVE = '1'"
                Dim CNT As Int64 = dst.Tables.Item("ARTCUSTL").Select(FTR).Count
                rowARTCUSTL.Item("CNT") = CNT
            Next
        End If
    End Sub

    Private Sub grdARTCUSTX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdARTCUSTX.AfterRowActivate
        Setup_ARTCUSTD()
    End Sub

    Private Sub grdARTCUSTD_AfterRowActivate(sender As Object, e As EventArgs) Handles grdARTCUSTD.AfterRowActivate
        Setup_ARTCUSTL()
    End Sub

    Private Sub grdARTCUSTL_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdARTCUSTL.BeforeRowUpdate
        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        grd.DisplayLayout.Rows(e.Row.Index).Cells("LAST_DATE").Value = Now()
        grd.DisplayLayout.Rows(e.Row.Index).Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        If Not Loading Then
            Select Case UltraTabControl1.SelectedTab.Key
                Case "Contact Maint"
                    UltraExplorerBar1.Groups("Customer Options").Visible = True
                    UltraExplorerBar1.Groups("List Options").Visible = False
                Case "List Maint"
                    UltraExplorerBar1.Groups("Customer Options").Visible = False
                    UltraExplorerBar1.Groups("List Options").Visible = True
                Case "Data Maint"
                    UltraExplorerBar1.Groups("Customer Options").Visible = False
                    UltraExplorerBar1.Groups("List Options").Visible = False
                Case Else
                    UltraExplorerBar1.Groups("Customer Options").Visible = False
                    UltraExplorerBar1.Groups("List Options").Visible = False
            End Select
        End If
    End Sub

    Private Sub ListActiveOnly()
        Dim filter As String = ""
        Dim dvw As DataView = DirectCast(grdARTLIST.DataSource, DataTable).DefaultView
        If chkListActiveOnly.Checked Then
            filter = "CLIST_ACTIVE = '1' OR CLIST_ACTIVE_TMP = '1'"
        End If
        dvw.RowFilter = String.Format(filter)
    End Sub

    Private Sub chkListActiveOnly_CheckedChanged(sender As Object, e As EventArgs) Handles chkListActiveOnly.CheckedChanged
        If Not Loading Then
            ListActiveOnly()
        End If
    End Sub

    Private Sub chkOnlyBuyers_CheckedChanged(sender As Object, e As EventArgs) Handles chkOnlyBuyers.CheckedChanged
        If Not Loading Then
            Setup_ARTCUSTD()
        End If
    End Sub

    Private Sub btnMakeMasterContacts_Click(sender As Object, e As EventArgs) Handles btnMakeMasterContacts.Click
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Make Masterfile Contacts?"
        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
        iMSG.AppendLine("Are You Sure You Want To Do This?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult <> MsgBoxResult.Yes Then
            Exit Sub
        End If

        Setup_Summary()
        Me.Cursor = Cursors.WaitCursor
        For Each rowARTCUSTX As DataRow In dst.Tables("ARTCUSTX").Select()
            Dim CUST_CODE As String = rowARTCUSTX.Item("CUST_CODE").ToString & String.Empty
            Dim FILTER As String = String.Format("CUST_CODE = '{0}'", CUST_CODE)
            If dst.Tables.Item("ARTCUSTD").Select(FILTER).Count = 0 Then
                AddMasterfileContact(rowARTCUSTX, False)
            Else
                Dim MFNAME As String = rowARTCUSTX.Item("CUST_CONTACT").ToString & String.Empty
                If MFNAME.Length > 0 Then
                    MFNAME = MFNAME.ToUpper
                    Dim MFFOUND As Boolean = False
                    For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(FILTER)
                        Dim CTNAME As String = rowARTCUSTD.Item("CONTACT_NAME").ToString & String.Empty
                        If CTNAME.Length > 0 Then
                            CTNAME = CTNAME.ToUpper
                        End If
                        If MFNAME = CTNAME Then
                            MFFOUND = True
                        End If
                    Next
                    If Not MFFOUND Then
                        AddMasterfileContact(rowARTCUSTX, False)
                    End If
                End If
            End If
        Next
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub AddMasterfileContact(ByRef rowARTCUSTX As DataRow, ByVal isPrimary As Boolean)
        Dim CUST_CODE As String = rowARTCUSTX.Item("CUST_CODE").ToString & String.Empty
        Dim FILTER As String = String.Format("CUST_CODE = '{0}'", CUST_CODE)
        Dim CONTACT_NO As Long = 1
        For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(FILTER, "CONTACT_NO")
            If Val(rowARTCUSTD.Item("CONTACT_NO").ToString & String.Empty) >= CONTACT_NO Then
                CONTACT_NO = Val(rowARTCUSTD.Item("CONTACT_NO").ToString & String.Empty) + 1
            End If
        Next

        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)

        Dim newARTCUSTD As DataRow = dst.Tables("ARTCUSTD").NewRow
        newARTCUSTD.Item("CUST_CODE") = CUST_CODE
        newARTCUSTD.Item("CONTACT_NO") = CONTACT_NO
        newARTCUSTD.Item("CONTACT_NAME") = rowARTCUST1.Item("CUST_CONTACT").ToString & String.Empty
        newARTCUSTD.Item("CONTACT_TITLE") = "Master Contact"
        newARTCUSTD.Item("CONTACT_EMAIL") = rowARTCUST1.Item("CUST_EMAIL").ToString & String.Empty
        newARTCUSTD.Item("CONTACT_PHONE") = rowARTCUST1.Item("CUST_PHONE").ToString & String.Empty
        newARTCUSTD.Item("CONTACT_EXT") = rowARTCUST1.Item("CUST_EXT").ToString & String.Empty
        newARTCUSTD.Item("CONTACT_FAX") = rowARTCUST1.Item("CUST_FAX").ToString & String.Empty
        newARTCUSTD.Item("CONTACT_TYPE") = "X"
        If isPrimary Then
            newARTCUSTD.Item("CONTACT_PRIMARY") = "1"
        Else
            newARTCUSTD.Item("CONTACT_PRIMARY") = "0"
        End If
        newARTCUSTD.Item("CONTACT_NOTE") = "Added By Cont Maint"
        newARTCUSTD.Item("INIT_OPER") = ASCMAIN1.USER_ID
        newARTCUSTD.Item("LAST_DATE") = DATETIME_STAMP
        newARTCUSTD.Item("LAST_OPER") = ASCMAIN1.USER_ID
        newARTCUSTD.Item("INIT_DATE") = DATETIME_STAMP

        newARTCUSTD.Item("CONTACT_CELL") = Null
        dst.Tables("ARTCUSTD").Rows.Add(newARTCUSTD)
    End Sub

    Private Sub btnBuyerGroups_Click(sender As Object, e As EventArgs) Handles btnBuyerGroups.Click
        'Stop
        'Dim CUST_CODE As String = "200138"
        'Dim sql As New Text.StringBuilder With {.Length = 0}
        'Dim FILTER As String = String.Format("CUST_CODE = '{0}'", CUST_CODE)
        'Dim CONTACT_NO As Long = 1
        'For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(FILTER, "CONTACT_NO")
        '    If Val(rowARTCUSTD.Item("CONTACT_NO").ToString & String.Empty) >= CONTACT_NO Then
        '        CONTACT_NO = Val(rowARTCUSTD.Item("CONTACT_NO").ToString & String.Empty) + 1
        '    End If
        'Next

        'sql.AppendLine("SELECT *")
        'sql.AppendLine("FROM ARTCUST2")
        'sql.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
        'Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString())
        'For Each rowARTCUST2 As DataRow In tbl.Rows
        '    Dim newARTCUSTD As DataRow = dst.Tables("ARTCUSTD").NewRow
        '    newARTCUSTD.Item("CUST_CODE") = CUST_CODE
        '    newARTCUSTD.Item("CONTACT_NO") = CONTACT_NO
        '    newARTCUSTD.Item("CONTACT_NAME") = rowARTCUST1.Item("CUST_CONTACT").ToString & String.Empty
        '    newARTCUSTD.Item("CONTACT_TITLE") = "Ship To Contact"
        '    newARTCUSTD.Item("CONTACT_EMAIL") = rowARTCUST1.Item("CUST_EMAIL").ToString & String.Empty
        '    newARTCUSTD.Item("CONTACT_PHONE") = rowARTCUST1.Item("CUST_PHONE").ToString & String.Empty
        '    newARTCUSTD.Item("CONTACT_EXT") = rowARTCUST1.Item("CUST_EXT").ToString & String.Empty
        '    newARTCUSTD.Item("CONTACT_FAX") = rowARTCUST1.Item("CUST_FAX").ToString & String.Empty
        '    newARTCUSTD.Item("CONTACT_TYPE") = "X"
        '    newARTCUSTD.Item("CONTACT_PRIMARY") = "0"
        '    newARTCUSTD.Item("CONTACT_NOTE") = "Added By Cont Maint"
        '    newARTCUSTD.Item("INIT_OPER") = ASCMAIN1.USER_ID
        '    newARTCUSTD.Item("LAST_DATE") = DATETIME_STAMP
        '    newARTCUSTD.Item("LAST_OPER") = ASCMAIN1.USER_ID
        '    newARTCUSTD.Item("INIT_DATE") = DATETIME_STAMP
        '    newARTCUSTD.Item("CONTACT_CELL") = Null
        '    dst.Tables("ARTCUSTD").Rows.Add(newARTCUSTD)
        'Next
    End Sub

    Private Sub btnManualUpdate_Click(sender As Object, e As EventArgs) Handles btnManualUpdate.Click
        Dim ToList As String = cboCopyToList.SelectedValue.ToString()
        Dim S As New Text.StringBuilder With {.Length = 0}
        S.AppendLine("SELECT CUST_CODE, CUST_NAME FROM ARTCUST1")
        With ASCMAIN1.CodeSelector
            .SQL = S.ToString
            .MultipleSelections = True
            .PreviouslySelectedCodes0 = ""
            .Caption = "Select Customers to Add To " & cboCopyToList.SelectedText.ToString & String.Empty
            .TABLE_NAME = ""
            .VIEW_NAME = ""
            .VIEW_DESC = ""
            .COLUMN_NAME = ""
            .COLUMN_PREKEYs = New Dictionary(Of String, String)
            .Custom_sql_where = ""
            .tblASTVIEW1 = New DataTable
        End With
        Dim F As New ASFCODE1
        F.ShowDialog()
        If ASCMAIN1.CodeSelector.Selections <> 0 Then
            For Each row As DataRow In ASCMAIN1.CodeSelector.SelectedRows
                Dim CUST_CODE As String = row.Item("CUST_CODE") & ""
                Dim filterD As String = String.Format("CUST_CODE = '{0}'", CUST_CODE)
                filterD = filterD & GetContactList()
                For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(filterD, "CUST_CODE, CONTACT_NO")
                    Dim SkipMaster As Boolean = False
                    If rowARTCUSTD.Item("CONTACT_TYPE").ToString & String.Empty = "X" Then
                        Dim CONTACT_NAME As String = rowARTCUSTD.Item("CONTACT_NAME").ToString & String.Empty
                        Dim CONTACT_EMAIL As String = rowARTCUSTD.Item("CONTACT_EMAIL").ToString & String.Empty
                        SkipMaster = IsMasterDuplicate(CUST_CODE, CONTACT_NAME, CONTACT_EMAIL)
                    End If
                    If Not SkipMaster Then
                        Dim CONTACT_NO As Int64 = Val(rowARTCUSTD.Item("CONTACT_NO").ToString & String.Empty)
                        Dim filterTo As String = String.Format("CUST_CODE = '{0}' AND CONTACT_NO = {1} AND CLIST_CODE = '{2}'", CUST_CODE, CONTACT_NO, ToList)
                        Dim rowARTCUSTL As DataRow = dst.Tables.Item("ARTCUSTL").Select(filterTo).FirstOrDefault
                        If Not IsNothing(rowARTCUSTL) Then
                            rowARTCUSTL.Item("CLIST_ACTIVE") = "1"
                        Else
                            Dim newARTCUSTL As DataRow = dst.Tables.Item("ARTCUSTL").NewRow
                            newARTCUSTL.Item("CUST_CODE") = CUST_CODE
                            newARTCUSTL.Item("CONTACT_NO") = CONTACT_NO
                            newARTCUSTL.Item("CLIST_CODE") = ToList
                            newARTCUSTL.Item("CLIST_ACTIVE") = "1"
                            newARTCUSTL.Item("INIT_OPER") = ASCMAIN1.USER_ID
                            newARTCUSTL.Item("INIT_DATE") = DATETIME_STAMP
                            newARTCUSTL.Item("LAST_OPER") = ASCMAIN1.USER_ID
                            newARTCUSTL.Item("LAST_DATE") = DATETIME_STAMP
                            dst.Tables.Item("ARTCUSTL").Rows.Add(newARTCUSTL)
                        End If
                    End If
                Next
            Next
            MsgBox("Done", vbOKOnly, "Done")
        End If
    End Sub

    Private Function GetContactList() As String
        Dim RetVal As String = ""
        If chkContactsX.Checked Then
            RetVal = " CONTACT_TYPE = 'X'"
        End If
        If chkContactsB.Checked Then
            If RetVal.Length = 0 Then
                RetVal = " CONTACT_TYPE = 'B'"
            Else
                RetVal = RetVal & " OR CONTACT_TYPE = 'B'"
            End If
        End If
        If chkContactsP.Checked Then
            If RetVal.Length = 0 Then
                RetVal = " CONTACT_TYPE = 'P'"
            Else
                RetVal = RetVal & " OR CONTACT_TYPE = 'P'"
            End If
        End If
        If chkContactsW.Checked Then
            If RetVal.Length = 0 Then
                RetVal = " CONTACT_TYPE = 'W'"
            Else
                RetVal = RetVal & " OR CONTACT_TYPE = 'W'"
            End If
        End If
        If chkContactsM.Checked Then
            If RetVal.Length = 0 Then
                RetVal = " CONTACT_TYPE = 'M'"
            Else
                RetVal = RetVal & " OR CONTACT_TYPE = 'M'"
            End If
        End If
        'Nothing was selected Make Nothing Get Seleted
        If RetVal = "" Then
            RetVal = RetVal & " AND CONTACT_TYPE = 'Z'"
        Else
            RetVal = " AND (" & RetVal & ")"
        End If

        Return RetVal
    End Function

    Private Sub btnSalesGreater_Click(sender As Object, e As EventArgs) Handles btnSalesGreater.Click
        If chkSettings() Then
            If Not IsNumeric(numSalesGreater.Value.ToString & String.Empty) Then
                MsgBox("Sales Must Be a Number > 0", vbOKOnly, "Problem")
            Else
                Dim SalesGreater As Int64 = Val(numSalesGreater.Value.ToString & String.Empty)
                If SalesGreater <= 0 Then
                    MsgBox("Sales Must Be a Number > 0", vbOKOnly, "Problem")
                Else
                    Dim ToList As String = cboCopyToList.SelectedValue.ToString()
                    For Each rowARTCUSTX As DataRow In dst.Tables("ARTCUSTX").Select("", "CUST_CODE")
                        Dim Y1 As Double = Val(rowARTCUSTX.Item("YR1").ToString & String.Empty)
                        Dim Y2 As Double = Val(rowARTCUSTX.Item("YR2").ToString & String.Empty)
                        Dim Y3 As Double = Val(rowARTCUSTX.Item("YR3").ToString & String.Empty)
                        Dim Y4 As Double = Val(rowARTCUSTX.Item("YR4").ToString & String.Empty)
                        If Y1 >= SalesGreater Or Y2 >= SalesGreater Or Y3 >= SalesGreater Or Y4 >= SalesGreater Then
                            Dim CUST_CODE As String = rowARTCUSTX.Item("CUST_CODE").ToString & String.Empty
                            'Dim CONTACT_NO As Int64 = Val(rowARTCUSTX.Item("CONTACT_NO").ToString & String.Empty)
                            Dim filterD As String = String.Format("CUST_CODE = '{0}'", CUST_CODE)
                            filterD = filterD & GetContactList()
                            For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(filterD, "CUST_CODE, CONTACT_NO")
                                Dim SkipMaster As Boolean = False
                                If rowARTCUSTD.Item("CONTACT_TYPE").ToString & String.Empty = "X" Then
                                    Dim CONTACT_NAME As String = rowARTCUSTD.Item("CONTACT_NAME").ToString & String.Empty
                                    Dim CONTACT_EMAIL As String = rowARTCUSTD.Item("CONTACT_EMAIL").ToString & String.Empty
                                    SkipMaster = IsMasterDuplicate(CUST_CODE, CONTACT_NAME, CONTACT_EMAIL)
                                End If
                                If Not SkipMaster Then
                                    Dim CONTACT_NO As Int64 = Val(rowARTCUSTD.Item("CONTACT_NO").ToString & String.Empty)
                                    Dim filterTo As String = String.Format("CUST_CODE = '{0}' AND CONTACT_NO = {1} AND CLIST_CODE = '{2}'", CUST_CODE, CONTACT_NO, ToList)
                                    Dim rowARTCUSTL As DataRow = dst.Tables.Item("ARTCUSTL").Select(filterTo).FirstOrDefault
                                    If Not IsNothing(rowARTCUSTL) Then
                                        rowARTCUSTL.Item("CLIST_ACTIVE") = "1"
                                    Else
                                        Dim newARTCUSTL As DataRow = dst.Tables.Item("ARTCUSTL").NewRow
                                        newARTCUSTL.Item("CUST_CODE") = CUST_CODE
                                        newARTCUSTL.Item("CONTACT_NO") = CONTACT_NO
                                        newARTCUSTL.Item("CLIST_CODE") = ToList
                                        newARTCUSTL.Item("CLIST_ACTIVE") = "1"
                                        newARTCUSTL.Item("INIT_OPER") = ASCMAIN1.USER_ID
                                        newARTCUSTL.Item("INIT_DATE") = DATETIME_STAMP
                                        newARTCUSTL.Item("LAST_OPER") = ASCMAIN1.USER_ID
                                        newARTCUSTL.Item("LAST_DATE") = DATETIME_STAMP
                                        dst.Tables.Item("ARTCUSTL").Rows.Add(newARTCUSTL)
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
                MsgBox("Done", vbOKOnly, "Done")
            End If
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

    Private Sub btnInitDate_Click(sender As Object, e As EventArgs) Handles btnInitDate.Click
        If chkSettings() Then
            If Not IsDate(dteInitDate.Value.ToString & String.Empty) Then
                MsgBox("Invalid Date", vbOKOnly, "Problem")
            Else
                Dim ToList As String = cboCopyToList.SelectedValue.ToString()
                For Each rowARTCUSTX As DataRow In dst.Tables("ARTCUSTX").Select("", "CUST_CODE")
                    If IsDate(rowARTCUSTX.Item("INIT_DATE").ToString & String.Empty) Then
                        Dim INIT_DATE As Date = CDate(rowARTCUSTX.Item("INIT_DATE").ToString & String.Empty)
                        If INIT_DATE >= CDate(dteInitDate.Value.ToString & String.Empty) Then
                            Dim CUST_CODE As String = rowARTCUSTX.Item("CUST_CODE").ToString & String.Empty
                            Dim filterD As String = String.Format("CUST_CODE = '{0}'", CUST_CODE)
                            filterD = filterD & GetContactList()
                            For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(filterD, "CUST_CODE, CONTACT_NO")
                                Dim SkipMaster As Boolean = False
                                If rowARTCUSTD.Item("CONTACT_TYPE").ToString & String.Empty = "X" Then
                                    Dim CONTACT_NAME As String = rowARTCUSTD.Item("CONTACT_NAME").ToString & String.Empty
                                    Dim CONTACT_EMAIL As String = rowARTCUSTD.Item("CONTACT_EMAIL").ToString & String.Empty
                                    SkipMaster = IsMasterDuplicate(CUST_CODE, CONTACT_NAME, CONTACT_EMAIL)
                                End If
                                If Not SkipMaster Then
                                    Dim CONTACT_NO As Int64 = Val(rowARTCUSTD.Item("CONTACT_NO").ToString & String.Empty)
                                    Dim filterTo As String = String.Format("CUST_CODE = '{0}' AND CONTACT_NO = {1} AND CLIST_CODE = '{2}'", CUST_CODE, CONTACT_NO, ToList)
                                    Dim rowARTCUSTL As DataRow = dst.Tables.Item("ARTCUSTL").Select(filterTo).FirstOrDefault
                                    If Not IsNothing(rowARTCUSTL) Then
                                        rowARTCUSTL.Item("CLIST_ACTIVE") = "1"
                                    Else
                                        Dim newARTCUSTL As DataRow = dst.Tables.Item("ARTCUSTL").NewRow
                                        newARTCUSTL.Item("CUST_CODE") = CUST_CODE
                                        newARTCUSTL.Item("CONTACT_NO") = CONTACT_NO
                                        newARTCUSTL.Item("CLIST_CODE") = ToList
                                        newARTCUSTL.Item("CLIST_ACTIVE") = "1"
                                        newARTCUSTL.Item("INIT_OPER") = ASCMAIN1.USER_ID
                                        newARTCUSTL.Item("INIT_DATE") = DATETIME_STAMP
                                        newARTCUSTL.Item("LAST_OPER") = ASCMAIN1.USER_ID
                                        newARTCUSTL.Item("LAST_DATE") = DATETIME_STAMP
                                        dst.Tables.Item("ARTCUSTL").Rows.Add(newARTCUSTL)
                                    End If
                                End If
                            Next
                        End If

                    End If

                Next
                MsgBox("Done", vbOKOnly, "Done")
            End If
        End If
    End Sub

    Private Sub btnCopyLists_Click(sender As Object, e As EventArgs) Handles btnCopyLists.Click
        If chkSettings() Then
            Dim FrList As String = cboCopyFromList.SelectedValue.ToString()
            Dim ToList As String = cboCopyToList.SelectedValue.ToString()
            Dim filterFr As String = String.Format("CLIST_CODE = '{0}' AND CLIST_ACTIVE = '1'", FrList)
            For Each rowTABLE_FROM As DataRow In dst.Tables("ARTCUSTL").Select(filterFr, "CUST_CODE, CONTACT_NO")
                Dim CUST_CODE As String = rowTABLE_FROM.Item("CUST_CODE").ToString & String.Empty
                Dim CONTACT_NO As Int64 = Val(rowTABLE_FROM.Item("CONTACT_NO").ToString & String.Empty)
                Dim CLIST_CODE As String = rowTABLE_FROM.Item("CLIST_CODE").ToString & String.Empty
                Dim filterTo As String = String.Format("CUST_CODE = '{0}' AND CONTACT_NO = {1} AND CLIST_CODE = '{2}'", CUST_CODE, CONTACT_NO, ToList)
                Dim rowARTCUSTL As DataRow = dst.Tables.Item("ARTCUSTL").Select(filterTo).FirstOrDefault
                If Not IsNothing(rowARTCUSTL) Then
                    rowARTCUSTL.Item("CLIST_ACTIVE") = "1"
                Else
                    Dim newARTCUSTL As DataRow = dst.Tables.Item("ARTCUSTL").NewRow
                    newARTCUSTL.Item("CUST_CODE") = CUST_CODE
                    newARTCUSTL.Item("CONTACT_NO") = CONTACT_NO
                    newARTCUSTL.Item("CLIST_CODE") = ToList
                    newARTCUSTL.Item("CLIST_ACTIVE") = "1"
                    newARTCUSTL.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    newARTCUSTL.Item("INIT_DATE") = DATETIME_STAMP
                    newARTCUSTL.Item("LAST_OPER") = ASCMAIN1.USER_ID
                    newARTCUSTL.Item("LAST_DATE") = DATETIME_STAMP
                    dst.Tables.Item("ARTCUSTL").Rows.Add(newARTCUSTL)
                End If
            Next
            MsgBox("Done", vbOKOnly, "Done")
        End If
    End Sub

    Private Function chkSettings() As Boolean
        Dim retval As Boolean = True
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Are You Ready?"
        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
        Dim ToList As String = cboCopyToList.Text.ToString()
        iMSG.AppendLine("You Are About To Update The Following List")
        iMSG.AppendLine("Based On Your Selection:")
        iMSG.AppendLine("")
        iMSG.AppendLine(ToList)
        iMSG.AppendLine("")
        iMSG.AppendLine("Are You Sure?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult <> MsgBoxResult.Yes Then
            retval = False
        End If
        Return retval
    End Function

    Private Sub grdARTCUSTL_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdARTCUSTL.ClickCell
        If e.Cell.Column.Key = "CLIST_ACTIVE" Then
            Dim CLIST_CODE As String = cboCLIST_CODE.SelectedValue
            Dim CUST_CODE As String = e.Cell.Row.Cells.Item("CUST_CODE").Value
            Dim CONTACT_NO As String = e.Cell.Row.Cells.Item("CONTACT_NO").Value
            Dim FLT As String = $"CUST_CODE = '{CUST_CODE}' AND CONTACT_NO = '{CONTACT_NO}'"
            Dim rowARTLIST As DataRow = dst.Tables("ARTLIST").Select(FLT).FirstOrDefault
            If e.Cell.Value = "1" Then
                e.Cell.Value = "0"
                e.Cell.Row.Cells.Item("CLIST_ACTIVE_TMP").Value = "1"
                If Not IsNothing(rowARTLIST) Then
                    rowARTLIST.Item("CLIST_ACTIVE") = "0"
                    rowARTLIST.Item("CLIST_ACTIVE_TMP") = "1"
                End If
            Else
                e.Cell.Value = "1"
                e.Cell.Row.Cells.Item("CLIST_ACTIVE_TMP").Value = "0"
                If Not IsNothing(rowARTLIST) Then
                    rowARTLIST.Item("CLIST_ACTIVE") = "1"
                    rowARTLIST.Item("CLIST_ACTIVE_TMP") = "0"
                End If
            End If
        End If

    End Sub

    Private Sub cboCLIST_CODE_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCLIST_CODE.SelectedIndexChanged
        If Not Loading Then
            RefreshList()
        End If
    End Sub

    Private Sub grdARTCUSTL_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdARTCUSTL.AfterRowUpdate
        SetListCounts()
        grdARTCUSTL.Update()
        grdARTCUSTL.Refresh()
    End Sub

    Private Sub grdARTLIST_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdARTLIST.ClickCell
        If e.Cell.Column.Key = "CLIST_ACTIVE" Then
            Dim CLIST_CODE As String = cboCLIST_CODE.SelectedValue
            Dim CUST_CODE As String = e.Cell.Row.Cells.Item("CUST_CODE").Value
            Dim CONTACT_NO As String = e.Cell.Row.Cells.Item("CONTACT_NO").Value
            Dim FLT As String = $"CLIST_CODE = '{CLIST_CODE}' AND CUST_CODE = '{CUST_CODE}' AND CONTACT_NO = '{CONTACT_NO}'"
            Dim rowARTCUSTL As DataRow = dst.Tables("ARTCUSTL").Select(FLT).FirstOrDefault

            If e.Cell.Value = "1" Then
                e.Cell.Value = "0"
                e.Cell.Row.Cells.Item("CLIST_ACTIVE_TMP").Value = "1"
                If Not IsNothing(rowARTCUSTL) Then
                    rowARTCUSTL.Item("CLIST_ACTIVE") = "0"
                End If
            Else
                e.Cell.Value = "1"
                e.Cell.Row.Cells.Item("CLIST_ACTIVE_TMP").Value = "0"
                If Not IsNothing(rowARTCUSTL) Then
                    rowARTCUSTL.Item("CLIST_ACTIVE") = "1"
                End If
            End If
            grdARTLIST.UpdateData()
            grdARTLIST.Refresh()
            ListActiveOnly()
        End If
    End Sub

    Private Sub grdARTLIST_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdARTLIST.AfterRowUpdate

        SetListCounts()
        grdARTLIST.Update()
        grdARTLIST.Refresh()
    End Sub

    Private Sub btnRefreshList_Click(sender As Object, e As EventArgs) Handles btnRefreshList.Click
        RefreshList()
    End Sub

    Private Sub RefreshList()
        ASCMAIN1.Progress("Now Loading List")
        Me.Cursor = Cursors.WaitCursor
        Update_Record(False)

        dst.EnforceConstraints = False

        Fill_Records("ARTLIST", cboCLIST_CODE.SelectedValue)
        grdARTLIST.Text = cboCLIST_CODE.Text
        'AddSalesToList()
        ListActiveOnly()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
End Class