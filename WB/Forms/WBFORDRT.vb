
Public Class WBFORDRT
    Dim InquiryOnly As Boolean = False
    Dim TimerOn As Boolean = False
    Dim TimerCount As Integer = 0
    Dim Processing As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        With dst

            SQLs.Length = 0
            SQLs.AppendLine("INSERT INTO SOTORDRT_L")
            SQLs.AppendLine("SELECT")
            SQLs.AppendLine("ORDR_NO,")
            SQLs.AppendLine("'0' AS LOCK_STATUS,")
            SQLs.AppendLine("NULL AS LOCK_BY,")
            SQLs.AppendLine("0 AS TO_PRINT,")
            SQLs.AppendLine("0 AS TO_EMAIL,")
            SQLs.AppendLine("NULL AS OTHER_EMAIL,")
            SQLs.AppendLine("'sys' AS WRITTEN_BY,")
            SQLs.AppendLine("NULL AS INIT_OPER,")
            SQLs.AppendLine("SYSDATE AS INIT_DATE,")
            SQLs.AppendLine("'sys' AS LAST_OPER,")
            SQLs.AppendLine("SYSDATE AS LAST_DATE")
            SQLs.AppendLine("FROM SOTORDR1_L")
            SQLs.AppendLine("WHERE ORDR_SOURCE = 'T'")
            SQLs.AppendLine("AND ORDR_STATUS = 'W'")
            SQLs.AppendLine("AND ORDR_NO NOT IN (SELECT ORDR_NO FROM SOTORDRT_L)")
            ASCMAIN1.sql = SQLs.ToString
            ASCDATA1.ExecuteSQL()

            SQLs.Length = 0
            SQLs.AppendLine("SELECT")
            SQLs.AppendLine("S1.*,")
            SQLs.AppendLine("NVL(ST.LOCK_STATUS,'0') LOCK_STATUS,")
            SQLs.AppendLine("NVL(LOCK_BY,'') LOCK_BY,")
            SQLs.AppendLine("NVL(ST.TO_PRINT,'0') TO_PRINT,")
            SQLs.AppendLine("NVL(ST.TO_EMAIL,'0') TO_EMAIL,")
            SQLs.AppendLine("NVL(OTHER_EMAIL,'') OTHER_EMAIL,")
            SQLs.AppendLine("NVL(WRITTEN_BY,'') WRITTEN_BY")
            SQLs.AppendLine("FROM SOTORDR1_L S1, SOTORDRT_L ST")
            SQLs.AppendLine("WHERE S1.ORDR_NO = ST.ORDR_NO (+)")
            SQLs.AppendLine("AND S1.ORDR_SOURCE = 'T'")
            SQLs.AppendLine("AND S1.ORDR_STATUS = 'W'")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "SOTORDR1_L", "**", 0, False, "", 1)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM SOTORDR2_L")
            SQLs.AppendLine("WHERE ORDR_NO IN")
            SQLs.AppendLine("(")
            SQLs.AppendLine("SELECT ORDR_NO FROM SOTORDR1_L WHERE ORDR_SOURCE = 'T' AND ORDR_STATUS = 'W'")
            SQLs.AppendLine(")")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "SOTORDR2_L", "**", 0, False, "", 2)
            .Tables("SOTORDR2_L").Columns.Add("LINE_TOTAL", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")

            'SQLs.Length = 0
            'SQLs.AppendLine("SELECT * FROM SOTORDRT_L")
            'ASCMAIN1.sql = SQLs.ToString()
            'Create_TDA(.Tables.Add, "SOTORDRT_L", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1)

            ASCMAIN1.sql = "SELECT * FROM SOTORDR1_L where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDP1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "SELECT * FROM SOTORDR2_L where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDP2", "**", 0, False, "V", 2)
            .Tables("SOTORDP2").Columns.Add("COLOR_DESC", GetType(System.String))
            .Tables("SOTORDP2").Columns.Add("UPC_CODE", GetType(System.String))
            .Tables("SOTORDP2").Columns.Add("FACTORY_CODE", GetType(System.String))

            ASCMAIN1.sql = "SELECT * FROM SOTORDR5_L where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDP5", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "SELECT * FROM ICTSTYL1"
            Create_TDA(.Tables.Add, "ICTSTYLD", "**", 1, False)
            Fill_Records("ICTSTYLD", "", , ASCMAIN1.sql)

            Create_TDA(.Tables.Add, "ICTPRICE", "*", 3)
        End With

        grdSOTORDR1_L.DataSource = dst.Tables("SOTORDR1_L")
        grdSOTORDR2_L.DataSource = dst.Tables("SOTORDR2_L")
        Setup_SOTORDR2_L()

        Create_Summary(grdSOTORDR2_L, "LINE_TOTAL", "Sum", "", "###,##0")

        With grdSOTORDR1_L.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With
        For i As Integer = 0 To grdSOTORDR1_L.DisplayLayout.Bands(0).Columns.Count - 1
            grdSOTORDR1_L.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        With grdSOTORDR2_L.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With
        For i As Integer = 0 To grdSOTORDR2_L.DisplayLayout.Bands(0).Columns.Count - 1
            grdSOTORDR2_L.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        Load_Record()

        tab.Visible = False
    End Sub

    Private Function CheckGridSelection() As String
        Dim EMsg As String = ""
        If grdSOTORDR1_L.Selected.Rows.Count = 0 Then
            EMsg &= vbCr & "You Must Select At Least One Order From The Grid"
        End If
        If grdSOTORDR1_L.Selected.Rows.Count > 1 Then
            EMsg &= vbCr & "You May Only Select One Order At A Time"
        End If
        Return EMsg
    End Function

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Finalize Orders"
                EMsg &= CheckGridSelection()
                If EMsg.Length = 0 Then
                    Dim ORDR_GROUP_NO As String = grdSOTORDR1_L.Selected.Rows(0).Cells.Item("ORDR_GROUP_NO").Text
                    Dim SQLS As New System.Text.StringBuilder
                    SQLS.Length = 0
                    SQLS.AppendLine("SELECT *")
                    SQLS.AppendLine("FROM SOTORDRT_L")
                    SQLS.AppendLine("WHERE ORDR_NO IN (")
                    SQLS.AppendLine("SELECT ORDR_NO FROM SOTORDR1_L")
                    SQLS.AppendLine("WHERE ORDR_SOURCE = 'T'")
                    SQLS.AppendLine(String.Format("AND ORDR_GROUP_NO = '{0}')", ORDR_GROUP_NO))
                    Dim tblSOTORDRT_L As DataTable = ASCDATA1.GetDataTable(SQLS.ToString(), String.Empty, "V", "")
                    For Each rowSOTORDRT_L As DataRow In tblSOTORDRT_L.Rows
                        If rowSOTORDRT_L.Item("LOCK_STATUS") & "" = "1" Then
                            EMsg &= vbCr & String.Format("Order {0} Is Locked By {1}", rowSOTORDRT_L.Item("ORDR_NO"), rowSOTORDRT_L.Item("LOCK_BY") & "")
                        End If
                    Next
                End If
                If EMsg.Length = 0 Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Finalize Tablet Order"
                    Dim iMSG As New System.Text.StringBuilder
                    iMSG.AppendLine("This Will Finalize The Selected Order")
                    iMSG.AppendLine("And Release It To ABSolution Lite.")
                    iMSG.AppendLine("Is This What You Want?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult <> MsgBoxResult.Yes Then
                        EMsg &= vbCr & "Finalize Canceled"
                    End If
                End If
            Case "Print Order"
                EMsg &= CheckGridSelection()
            Case "E-Mail Order"
                MsgBox("This Feature Is Still Under Construction", MsgBoxStyle.Information, "Not Ready Yet")
                EMsg &= vbCr & "Feature Under Construction"
            Case "Refresh"

            Case "Item Prices"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Item Pricing"
                Dim iMSG As New System.Text.StringBuilder
                iMSG.AppendLine("This Will Create New Pricing For")
                iMSG.AppendLine("All Items in The System For The")
                iMSG.AppendLine("iPads.  It May Take A While.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Is This What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult <> MsgBoxResult.Yes Then
                    EMsg &= vbCr & "Item Pricing Canceled"
                End If
            Case "Clear Lock"
                EMsg &= CheckGridSelection()
                If EMsg = "" Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "SLOW DOWN!!!!"
                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    iMSG.AppendLine("You Are In Danger And Better")
                    iMSG.AppendLine("Know What You Are Doing!!")
                    iMSG.AppendLine("")
                    iMSG.AppendLine("If This Order Is Still On A")
                    iMSG.AppendLine("Tablet Somewhere And You Proceed")
                    iMSG.AppendLine("There WILL Be People Angry With You!!")
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Still Wanna Do this?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult <> MsgBoxResult.Yes Then
                        EMsg &= "Lock Clear Request Cancelled."
                    End If
                End If
            Case "Exit"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Finalize Orders"
                FinalizeOrder()
                Load_Record()
            Case "Print Order"
                Dim ORDR_NO As String = grdSOTORDR1_L.Selected.Rows(0).Cells.Item("ORDR_NO").Text
                Dim CUST_CODE As String = grdSOTORDR1_L.Selected.Rows(0).Cells.Item("CUST_CODE").Text
                Print_Record(True, ORDR_NO, CUST_CODE)
            Case "E-Mail Order"

            Case "Refresh"
                Load_Record()
            Case "Exit"
                Call Mode_Settings(False)
            Case "Clear Lock"
                ClearLock()
                Load_Record()
            Case "Item Prices"
                ItemPrices()
                Load_Record()
            Case "Who Am I"
                Dim host As Net.IPHostEntry = Net.Dns.GetHostEntry(Net.Dns.GetHostName())
                Dim myIP As String = ""
                Dim ipFound As Boolean = False
                For Each ip As Net.IPAddress In host.AddressList
                    If ip.ToString.StartsWith("192.168") Then
                        ipFound = True
                        myIP = String.Format("https://{0}:4200/api/", ip.ToString)
                    End If
                Next
                If ipFound Then
                    MsgBox(myIP, vbOKOnly, "Your Address")
                Else
                    MsgBox("Address Not Found", vbOKOnly, "Your Address")
                End If
        End Select
    End Sub

    Private Sub ItemPrices()
        dst.Tables.Item("ICTPRICE").Clear()
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT *")
        sql.AppendLine("FROM ICTSTYL1")
        sql.AppendLine("WHERE STYLE_CODE IN")
        sql.AppendLine("(")
        sql.AppendLine("  SELECT S1.STYLE_CODE")
        sql.AppendLine("  FROM ICTSTYL1 S1, ICTSTAT2 S2")
        sql.AppendLine("  WHERE S1.STYLE_CODE = S2.STYLE_CODE (+)")
        sql.AppendLine("  AND S2.WHSE_CODE (+) = 'MS'")
        sql.AppendLine("  HAVING (SUM((NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))) > 0 OR MIN(S1.STYLE_STATUS) = 'A')")
        sql.AppendLine("  GROUP BY S1.STYLE_CODE")
        sql.AppendLine(")")
        sql.AppendLine("AND NVL(CUST_CODE,'NULL') = 'NULL'")
        Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
        Dim StyleNo As Int64 = 0
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", "180000")
        For Each rowICTSTYL1 As DataRow In tbl.Select("", "STYLE_CODE")
            StyleNo += 1
            Dim STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE").ToString & String.Empty
            ASCMAIN1.Progress("Style: " & STYLE_CODE, "Record: " & StyleNo)

            If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                If STYLE_CODE = "MTX71769" Then Stop
            End If

            For PriceLevel As Int64 = 0 To 5
                Select Case PriceLevel
                    Case 0
                        rowARTCUST1.Item("CUST_PRICE_TIER") = "PC"
                        rowARTCUST1.Item("CUST_DISC_PCT_EXTRA") = "0"
                        rowARTCUST1.Item("CUST_DISC_PCT") = 0
                    Case 1
                        rowARTCUST1.Item("CUST_PRICE_TIER") = "PC"
                        rowARTCUST1.Item("CUST_DISC_PCT_EXTRA") = "1"
                        rowARTCUST1.Item("CUST_DISC_PCT") = 0
                    Case 2
                        rowARTCUST1.Item("CUST_PRICE_TIER") = "PC"
                        rowARTCUST1.Item("CUST_DISC_PCT_EXTRA") = "2"
                        rowARTCUST1.Item("CUST_DISC_PCT") = 0
                    Case 3
                        rowARTCUST1.Item("CUST_PRICE_TIER") = "HC"
                        rowARTCUST1.Item("CUST_DISC_PCT_EXTRA") = "0"
                        rowARTCUST1.Item("CUST_DISC_PCT") = 0
                    Case 4
                        rowARTCUST1.Item("CUST_PRICE_TIER") = "FC"
                        rowARTCUST1.Item("CUST_DISC_PCT_EXTRA") = "0"
                        rowARTCUST1.Item("CUST_DISC_PCT") = 0
                    Case 5
                        rowARTCUST1.Item("CUST_PRICE_TIER") = "SP"
                        rowARTCUST1.Item("CUST_DISC_PCT_EXTRA") = "1"
                        rowARTCUST1.Item("CUST_DISC_PCT") = 54
                    Case Else
                        Stop 'This should never happen.
                End Select
                Dim DISCOUNTS As List(Of DISCOUNTS) = SOCMAIN2.Price_Discounts(Me, "", rowARTCUST1, STYLE_CODE, True,, True)
                For i As Int64 = 0 To DISCOUNTS.Count - 1
                    Dim newICTPRICE As DataRow = dst.Tables.Item("ICTPRICE").NewRow
                    newICTPRICE.Item("STYLE_CODE") = STYLE_CODE
                    newICTPRICE.Item("PRICE_LEVEL") = PriceLevel
                    'For i As Int64 = 0 To DISCOUNTS.Count - 1
                    newICTPRICE.Item("DISCOUNT_LEVEL") = i
                    newICTPRICE.Item("DISCOUNT_DESC") = DISCOUNTS(i).DISCOUNT_DESC
                    newICTPRICE.Item("DISCOUNT_QTY") = DISCOUNTS(i).DISCOUNT_QTY
                    newICTPRICE.Item("DISCOUNT_PRICE") = DISCOUNTS(i).DISCOUNT_PRICE
                    newICTPRICE.Item("DISCOUNT_PCT") = DISCOUNTS(i).DISCOUNT_PCT
                    newICTPRICE.Item("LAST_DATE") = Now()
                    dst.Tables.Item("ICTPRICE").Rows.Add(newICTPRICE)
                Next
            Next
        Next
        'Next
        ASCMAIN1.Progress("Updating Database", "")
        BeginTrans()
        Call Update_Record_TDA("ICTPRICE", "DELETE FROM ICTPRICE")
        CommitTrans()
        ASCMAIN1.Progress("", "")
        MsgBox("Prices Are Updated", vbOKOnly, "Done")
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Finalize Orders").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Print Order").Visible = Not ScreenMode
                .Groups("Screen Control").Items("E-Mail Order").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Exit").Visible = Not ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
    End Sub

    Sub Clear_Record()
        dst.Tables("ICTSTYL1").Rows.Clear()
    End Sub

    Sub Load_Record()
        'Call Save_Header_Fields(UltraGroupBox1)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Refreshing Data", "")

        EnforceConstraints(False)

        Fill_Records("SOTORDR1_L")
        Fill_Records("SOTORDR2_L")
        'Fill_Records("SOTORDRT_L")

        EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
        'Call BeginTrans()
        'For Each TABLE_NAME As String In {"SOTORDR1", "SOTORDR2", "SOTORDR5"}
        '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where ORDR_NO = '" & ORDR_NO & "'"
        '    ASCDATA1.ExecuteSQL()
        '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & "_L where ORDR_NO = '" & ORDR_NO & "'"
        '    ASCDATA1.ExecuteSQL()
        'Next
        'Call CommitTrans("Order / Quote Deleted")
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        'TODO: Remove this stop before going live.
        Stop
        Call CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
        Print_Report_Begin()
        'frm.CR_params.Add("SUBT", "")
        'Fill SOTORDRP records
        Fill_Records("ARTCUST1", CUST_CODE, True)
        Fill_Records("SOTORDP1", ORDR_NO, True)
        Fill_Records("SOTORDP2", ORDR_NO, True)
        Fill_Records("SOTORDP5", ORDR_NO, True)
        If dst.Tables.Item("SOTORDP1").Rows.Count > 0 Then
            Dim TERM_CODE As String = dst.Tables.Item("SOTORDP1").Rows(0).Item("TERM_CODE")
            Dim rowTATTERM1 As DataRow = LookUp("TATTERM1", TERM_CODE)
            CR_params.Add("TERM_DESC", rowTATTERM1.Item("TERM_DESC").ToString)
        Else
            CR_params.Add("TERM_DESC", "")
        End If
        For Each rowSOTORDP2 As DataRow In dst.Tables("SOTORDP2").Select()
            Dim STYLE_CODE As String = rowSOTORDP2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDP2.Item("COLOR_CODE")
            Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
            rowSOTORDP2.Item("COLOR_DESC") = rowICTCOLR1.Item("COLOR_DESC")
            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            If IsNothing(rowICTSTYC1) Then
                rowSOTORDP2.Item("UPC_CODE") = ""
            Else
                rowSOTORDP2.Item("UPC_CODE") = rowICTSTYC1.Item("UPC_CODE") & ""
            End If

            Dim FACTORY_CODE As String = GetVendorData(rowICTSTYL1.Item("VEND_CODE").ToString, "VEND_SUPPLIER_ID")
            rowSOTORDP2.Item("FACTORY_CODE") = FACTORY_CODE
        Next
        'Generate_Report("SORORDRT")
        Generate_Report("SORORDRO")
        If chkDefaultPrinter.Checked Then
            Dim PS As New System.Drawing.Printing.PrinterSettings
            Print_Report_End(True, False, PS.PrinterName, No_Of_Copies)
        Else
            Print_Report_End()
        End If
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdSOTORDRX, "SSB", "Show Filter", "Show GroupBox")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.SourceControl.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        End If

        Select Case e.SourceControl.Name
            'Case "grdSOTORDR1"
            '    If Not InquiryOnly Then
            '        e.Tool.ToolbarsManager.Tools("Edit Ship To").SharedProps.Visible = True
            '    End If
        End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Edit Ship To"
            '    If Not InquiryOnly Then
            '        MsgBox("Edit Ship To Feature Coming Soon", MsgBoxStyle.Exclamation, "Waiting For Feature")
            '    End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        'Select Case Absx1.GetABSColumnName(sender)
        '    Case "BANK_CODE"
        '        If e.KeyCode = Windows.Forms.Keys.Enter Then
        '            Call Click_Command("New", e)
        '        End If
        '    Case "PYMT_BATCH_NO"
        '        If e.KeyCode = Windows.Forms.Keys.Enter Then
        '            Call Click_Command("Edit", e)
        '        End If
        'End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        'Select Case COLUMN_NAME
        '    Case "STYLE_CODE"
        '        'FillStyle()
        'End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        'Select Case Absx1.GetABSColumnName(txtctl)
        '    Case "PYMT_BATCH_NO"
        '        Call Click_Command("Edit")
        'End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        If IsLoading Then
            Exit Sub
        Else
            MyBase.txt_ValueChanged(sender, e)
        End If
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub ClearLock()
        Dim ORDR_GROUP_NO As String = grdSOTORDR1_L.Selected.Rows(0).Cells.Item("ORDR_GROUP_NO").Text
        Dim SQLS As New System.Text.StringBuilder
        SQLS.Length = 0
        SQLS.AppendLine("UPDATE SOTORDRT_L")
        SQLS.AppendLine("SET LOCK_STATUS = '0',")
        SQLS.AppendLine("LOCK_BY = null,")
        SQLS.AppendLine(String.Format("LAST_OPER = '{0}',", ASCMAIN1.USER_ID))
        SQLS.AppendLine("LAST_DATE = SYSDATE")
        SQLS.AppendLine("WHERE ORDR_NO IN (")
        SQLS.AppendLine("SELECT ORDR_NO FROM SOTORDR1_L")
        SQLS.AppendLine("WHERE ORDR_SOURCE = 'T'")
        SQLS.AppendLine(String.Format("AND ORDR_GROUP_NO = '{0}')", ORDR_GROUP_NO))
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()
        MsgBox("Lock Cleared, Grid Will Be Refreshed.", MsgBoxStyle.OkOnly, "Complete")
    End Sub

    Private Sub FinalizeOrder()
        Call BeginTrans()
        Dim SQLS As New System.Text.StringBuilder

        Dim ORDR_GROUP_NO As String = grdSOTORDR1_L.Selected.Rows(0).Cells.Item("ORDR_GROUP_NO").Text
        Dim ORDR_NOs As String()

        SQLS.Length = 0
        SQLS.AppendLine("SELECT *")
        SQLS.AppendLine("FROM SOTORDR1_L")
        SQLS.AppendLine("WHERE ORDR_GROUP_NO = :PARM1")
        Dim tblSOTORDR1_L As DataTable = ASCDATA1.GetDataTable(SQLS.ToString(), String.Empty, "V", ORDR_GROUP_NO)
        Dim Index As Integer = 0
        Dim MaxIndex As Integer = tblSOTORDR1_L.Rows.Count - 1
        ReDim ORDR_NOs(MaxIndex)
        For Each rowSOTORDR1_L As DataRow In tblSOTORDR1_L.Rows
            ORDR_NOs(Index) = rowSOTORDR1_L.Item("ORDR_NO").ToString
            Index += 1
        Next

        For Each ORDR_NO As String In ORDR_NOs
            Dim ORDR_BATCH_NO As String = ASCMAIN1.Next_Control_No("ORDR_BATCH_NO")
            'SQLS.Length = 0
            SQLS.Length = 0
            SQLS.AppendLine("UPDATE SOTORDRT_L")
            SQLS.AppendLine("SET LOCK_STATUS = '0',")
            SQLS.AppendLine("LOCK_BY = null,")
            SQLS.AppendLine(String.Format("LAST_OPER = '{0}',", ASCMAIN1.USER_ID))
            SQLS.AppendLine("LAST_DATE = SYSDATE")
            SQLS.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO))
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()

            SQLS.Length = 0
            SQLS.AppendLine("UPDATE SOTORDR1_L")
            SQLS.AppendLine("SET ORDR_STATUS = 'L',")
            SQLS.AppendLine("ORDR_SOURCE = 'L',")
            SQLS.AppendLine("ORDR_GROUP_NO = NULL,")
            SQLS.AppendLine(String.Format("ORDR_BATCH_NO = '{0}'", ORDR_BATCH_NO))
            SQLS.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO))
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()

            'Add Style Attributes not being added by iPads.
            SQLS.Length = 0
            SQLS.AppendLine("BEGIN DECLARE CURSOR C1 IS")
            SQLS.AppendLine("  SELECT L2.ORDR_NO, L2.ORDR_LNO, L2.STYLE_CODE, L2.COLOR_CODE, S1.INNER_PACK_QTY, S1.STYLE_UOM, S1.CARTON_PACK_QTY")
            SQLS.AppendLine("  FROM SOTORDR2_L L2, ICTSTYL1 S1")
            SQLS.AppendLine("  WHERE L2.STYLE_CODE = S1.STYLE_CODE")
            SQLS.AppendLine("  AND NVL(L2.STYLE_UOM,'NULL') = 'NULL'")
            SQLS.AppendLine(String.Format("  AND L2.ORDR_NO = '{0}';", ORDR_NO))
            SQLS.AppendLine("BEGIN FOR R1 IN C1 LOOP")
            SQLS.AppendLine("  UPDATE SOTORDR2_L")
            SQLS.AppendLine("  SET INNER_PACK_QTY = R1.INNER_PACK_QTY,")
            SQLS.AppendLine("  STYLE_UOM = R1.STYLE_UOM,")
            SQLS.AppendLine("  CARTON_PACK_QTY = R1.CARTON_PACK_QTY")
            SQLS.AppendLine("  WHERE ORDR_NO = R1.ORDR_NO")
            SQLS.AppendLine("  AND ORDR_LNO = R1.ORDR_LNO;")
            SQLS.AppendLine("END LOOP; END; END;")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()

            For Each TableName As String In New String() {"SOTORDR1", "SOTORDR2", "SOTORDR5"}
                SQLS.Length = 0
                SQLS.AppendLine("INSERT INTO " & TableName)
                SQLS.AppendLine("SELECT *")
                SQLS.AppendLine(String.Format("FROM {0}_L", TableName))
                SQLS.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO))
                ASCMAIN1.sql = SQLS.ToString
                ASCDATA1.ExecuteSQL()
            Next

        Next

        Call CommitTrans("Order(s) Finalized.")
    End Sub

    Private Function GetVendorData(ByVal VEND_CODE As String, ByVal COLUMN As String) As String
        Dim RetVal As String = ""
        If VEND_CODE.Length > 0 And COLUMN.Length > 0 Then
            ASCMAIN1.sql = String.Format("SELECT {0} FROM APTVEND1 WHERE VEND_CODE = '{1}'", COLUMN, VEND_CODE)
            RetVal = ASCDATA1.GetDataValue
        End If
        Return RetVal
    End Function

    Sub Setup_SOTORDR2_L()
        If grdSOTORDR1_L.ActiveRow Is Nothing OrElse (Not grdSOTORDR1_L.ActiveRow.IsDataRow Or grdSOTORDR1_L.ActiveRow.IsAddRow) Then
            grdSOTORDR2_L.Visible = False
        Else
            Dim dvw As DataView = DirectCast(grdSOTORDR2_L.DataSource, DataTable).DefaultView
            Dim ORDR_NO As String = grdSOTORDR1_L.ActiveRow.Cells("ORDR_NO").Value
            dvw.RowFilter = String.Format("ORDR_NO = '{0}'", ORDR_NO)
            grdSOTORDR2_L.Text = "Order Details for Order " & CStr(ORDR_NO)
            grdSOTORDR2_L.Visible = True
        End If
    End Sub

    Sub RefreshLocks()
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine("NVL(ST.LOCK_STATUS,'0') LOCK_STATUS,")
        sql.AppendLine("NVL(LOCK_BY,'') LOCK_BY")
        sql.AppendLine("FROM SOTORDRT_L ST")
        sql.AppendLine("WHERE ST.ORDR_NO IN (")
        sql.AppendLine(" SELECT ORDR_NO")
        sql.AppendLine(" FROM SOTORDR1_L")
        sql.AppendLine(" WHERE ORDR_SOURCE = 'T'")
        sql.AppendLine(" AND ORDR_STATUS = 'W'")
        sql.AppendLine(")")
        Dim tblSOTORDRT_L As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
        For Each rowSOTORDR1_L As DataRow In dst.Tables("SOTORDR1_L").Select()
            Dim FORDR_NO As String = rowSOTORDR1_L.Item("ORDR_NO").ToString & String.Empty
            Dim FILTER As String = String.Format("ORDR_NO = '{0}'", FORDR_NO)
            Dim rowSOTORDRT_L As DataRow = tblSOTORDRT_L.Select(FILTER).FirstOrDefault
            If IsNothing(rowSOTORDRT_L) Then
                rowSOTORDR1_L.Item("LOCK_STATUS") = "0"
                rowSOTORDR1_L.Item("LOCK_BY") = ""
            Else
                rowSOTORDR1_L.Item("LOCK_STATUS") = rowSOTORDRT_L.Item("LOCK_STATUS").ToString & String.Empty
                rowSOTORDR1_L.Item("LOCK_BY") = rowSOTORDRT_L.Item("LOCK_BY").ToString & String.Empty
            End If
        Next
    End Sub
#End Region

#Region "Form Controls"
#Region "Buttons"
    Private Sub btnTimerRefresh_Click(sender As System.Object, e As System.EventArgs) Handles btnTimerRefresh.Click
        MsgBox("This Feature No Longer Applicable", MsgBoxStyle.Information, "Not For ABSolution Lite")
        'If TimerOn Then
        '    Timer1.Stop()
        '    btnTimerRefresh.Text = "Start Timer"
        '    TimerOn = False
        '    TimerCount = 0
        '    ProgressBar1.Value = 0
        '    ProgressBar1.Visible = False
        'Else
        '    btnTimerRefresh.Text = "Pause Timer"
        '    TimerOn = True
        '    TimerCount = 0
        '    ProgressBar1.Value = 0
        '    ProgressBar1.Visible = True
        '    Timer1.Start()
        'End If
    End Sub
#End Region

#Region "Grids"
    Private Sub grdSOTORDR1_L_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDR1_L.AfterRowActivate
        Setup_SOTORDR2_L()
    End Sub
#End Region

#Region "Timers"
    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        'TimerCount += 1

        'If Not IsNumeric(txtTimerRefresh.Text) Then
        '    txtTimerRefresh.Text = "15"
        'Else
        '    If ((TimerCount / Val(txtTimerRefresh.Text)) * 100) > 100 Then
        '        ProgressBar1.Value = 100
        '    Else
        '        ProgressBar1.Value = (TimerCount / Val(txtTimerRefresh.Text)) * 100
        '    End If

        'End If
        'If TimerCount >= Val(txtTimerRefresh.Text) And Not Processing Then
        '    Processing = True
        '    Load_Record()
        '    For Each rowSOTORDR1_L As DataRow In dst.Tables("SOTORDR1_L").Select()
        '        If rowSOTORDR1_L.Item("TO_PRINT") = "1" Then
        '            Dim ORDR_NO As String = rowSOTORDR1_L.Item("ORDR_NO").ToString()
        '            Dim CUST_CODE As String = rowSOTORDR1_L.Item("CUST_CODE").ToString()
        '            Print_Record(True, ORDR_NO, CUST_CODE)
        '            Dim SQLS As New System.Text.StringBuilder
        '            SQLS.Length = 0
        '            SQLS.AppendLine("UPDATE SOTORDRT_L")
        '            SQLS.AppendLine("SET TO_PRINT = '0'")
        '            SQLS.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO))
        '            ASCMAIN1.sql = SQLS.ToString
        '            ASCDATA1.ExecuteSQL()
        '        End If
        '    Next
        '    TimerCount = 0
        '    Processing = False
        'End If
    End Sub
#End Region
#End Region

End Class