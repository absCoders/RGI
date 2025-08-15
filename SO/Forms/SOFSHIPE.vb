Imports DPayments.DShippingSDK
Imports System.Net
Imports System.IO
Imports Newtonsoft.Json
Imports Infragistics.Win.UltraWinGrid

Public Class SOFSHIPE

    Private PICK_BATCH_NO As String = String.Empty
    Private WHSE_CODE_TRUCK As String = String.Empty
    Private isCustomTruck As Boolean = False
    Private dictCustomTruck As New Dictionary(Of String, String)
    Private drSOTTRCK1 As DataRow = Nothing

    Private allItemsOnBackOrder As Boolean = False
    Private AutoCancel As Boolean = False
    Private dictAppearances As New Dictionary(Of String, Infragistics.Win.Appearance)
    Private Appearance_Incomplete As New Infragistics.Win.Appearance
    Private sqlSOTPICK2 As String = String.Empty

    Private clsShip As New TAC.WHCSHIP1
    Private tblTATSTATE As DataTable
    Private defaultPACKAGING_TYPE As String = "31"
    Private defaultPKG_CODE As String = "OTHER"
    Private clsTACZPLT1 As New TAC.TACZPLT1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            Get_PARM("ASTPARM1")
            tblTATSTATE = ASCDATA1.GetDataTable("SELECT * FROM TATSTATE", "TATSTATE")

            ASCMAIN1.sql = $"SELECT SOTTRCK1.TRUCK_NO, SOTTRCK1.TRUCK_TYPE, SOTTRCK1.PICK_BATCH_NO, SOTTRCK1.WHSE_CODE, 
                                SOTPICK0.PICK_BATCH_STATUS, SOTPICK0.INIT_DATE, SOTPICK0.INIT_OPER, SOTTOTE1X.NUM_TOTES
                                FROM SOTTRCK1, SOTPICK0,
                                        (
                                        SELECT SOTTOTE1.TRUCK_NO, COUNT(*) NUM_TOTES
                                        FROM SOTTOTE1, SOTPICK1
                                        WHERE SOTTOTE1.PICK_NO = SOTPICK1.PICK_NO
                                        AND SOTPICK1.PICK_STATUS = 'P'
                                        GROUP BY SOTTOTE1.TRUCK_NO
                                        ) SOTTOTE1X
                                WHERE SOTTRCK1.PICK_BATCH_NO = SOTPICK0.PICK_BATCH_NO
                                AND SOTTRCK1.WHSE_CODE = SOTPICK0.WHSE_CODE
                                AND SOTPICK0.PICK_BATCH_STATUS IN ('K', 'N')
                                AND SOTTRCK1.TRUCK_NO = SOTTOTE1X.TRUCK_NO (+)"
            Create_TDA(.Tables.Add, "SOTTRCK1X", ASCMAIN1.sql, 0, False, String.Empty)

            ASCMAIN1.sql = $"SELECT SOTPICK1.*, SOTORDR1.CUST_NAME, SOTORDR1.CUST_SHIP_TO_NAME, SOTORDR1.CUST_SHIP_TO_DC,
                                SOTORDR1.ORDR_CUST_PO, SOTSVIA1.SHIP_VIA_DESC, SOTTOTE1.TRUCK_NO, '0' SELECTED,
                                SOTPICK2.PICK_QTY, SOTPICK2.PICK_QTY_CONF, SOTPICK2.PICK_QTY_CANC, SOTPICK2.PICK_QTY_BACK
                                FROM SOTPICK1, SOTORDR1, SOTSVIA1, SOTTOTE1, SOTTRCK1,
                                    (SELECT PICK_NO, 
                                        SUM(DECODE(NVL(PICK_QTY, 0), 0, NVL(PICK_QTY_BACK, 0) + NVL(PICK_QTY_CANC, 0), NVL(PICK_QTY, 0))) PICK_QTY,  
                                        SUM(NVL(PICK_QTY_CONF, 0)) PICK_QTY_CONF, 
                                        SUM(NVL(PICK_QTY_CANC, 0)) PICK_QTY_CANC, 
                                        SUM(NVL(PICK_QTY_BACK, 0)) PICK_QTY_BACK 
                                        FROM SOTPICK2
                                        GROUP BY PICK_NO
                                    ) SOTPICK2
                                WHERE SOTPICK1.PICK_NO = SOTPICK2.PICK_NO
                                AND SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO
                                AND SOTPICK1.PICK_STATUS = 'P'
                                AND SOTPICK1.TOTE_NO = SOTTOTE1.TOTE_NO
                                AND SOTPICK1.PICK_NO = SOTTOTE1.PICK_NO 
                                AND SOTTOTE1.TRUCK_NO = SOTTRCK1.TRUCK_NO
                                AND SOTTRCK1.TRUCK_NO = :PARM1
                                AND SOTPICK1.PICK_BATCH_NO = :PARM2
                                AND SOTTRCK1.PICK_BATCH_NO = SOTPICK1.PICK_BATCH_NO
                                AND SOTTRCK1.PICK_BATCH_NO IS NOT NULL 
                                AND SOTORDR1.SHIP_VIA_CODE = SOTSVIA1.SHIP_VIA_CODE (+)"
            Create_TDA(.Tables.Add, "SOTPICK1X", ASCMAIN1.sql, 0, False, "VV", 1)
            .Tables("SOTPICK1X").Columns.Add("TOTAL_SCANNED", GetType(System.Int16), "ISNULL(PICK_QTY_CONF, 0) + ISNULL(PICK_QTY_CANC, 0) + ISNULL(PICK_QTY_BACK, 0)")
            .Tables("SOTPICK1X").Columns.Add("INCOMPLETE", GetType(System.Int16), "IIF(ISNULL(PICK_QTY, 0) <> ISNULL(TOTAL_SCANNED, 0), '1', '0')")
            .Tables("SOTPICK1X").Columns.Add("BAY_COLOR", GetType(System.String))
            .Tables("SOTPICK1X").Columns.Add("ALL_ITEMS_BACK", GetType(System.String))

            Create_TDA(.Tables.Add, "SOTPICK0", "*")
            Create_TDA(.Tables.Add, "SOTSHIP1", "*")
            Create_TDA(.Tables.Add, "ARTOPEN1", "*")
            Create_TDA(.Tables.Add, "ARTCUST1", "*")
            Create_TDA(.Tables.Add, "ARTCUST2", "*")
            Create_TDA(.Tables.Add, "ARTCUSTS", "*")

            Create_TDA(.Tables.Add, "SOTPICK1", "*")
            Create_TDA(.Tables.Add, "SOTPICK2", "*", 1)
            .Tables("SOTPICK2").Columns.Add("STYLE_CODE", GetType(System.Int32))
            .Tables("SOTPICK2").Columns.Add("STYLE_DESC", GetType(System.Int32))
            .Tables("SOTPICK2").Columns.Add("ORDR_UNIT_PRICE", GetType(System.Int32))
            .Tables("SOTPICK2").Columns.Add("PICK_QTY_SCAN", GetType(System.Int32), "ISNULL(PICK_QTY_CONF, 0) + ISNULL(PICK_QTY_CANC, 0) + ISNULL(PICK_QTY_BACK, 0)")

            sqlSOTPICK2 = "SELECT SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.STYLE_DESC, SOTORDR2.ORDR_UNIT_PRICE 
                            FROM SOTPICK2, SOTORDR2, ICTSTYL1
                            WHERE SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO
                            AND SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO
                            AND SOTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE (+)
                            AND SOTPICK2.PICK_NO = :PARM1"

            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            Create_TDA(.Tables.Add, "SOTORDR2", "*", 1)
            Create_TDA(.Tables.Add, "SOTORDR5", "*")

            Create_TDA(.Tables.Add, "TATEVNT1", "*")

            Create_TDA(.Tables.Add, "SOTINVH1", "*")
            Create_TDA(.Tables.Add, "SOTINVH2", "*", 1)
            .Tables("SOTINVH2").Columns.Add("EXT_PRICE", GetType(System.Decimal), "ISNULL(ORDR_UNIT_PRICE, 0) * ISNULL(ORDR_QTY_SHIP, 0)")
            Create_TDA(.Tables.Add, "SOTINVH9", "*")
            Create_TDA(.Tables.Add, "SOTINVHM", "*")
            Create_TDA(.Tables.Add, "SOTRNGA1", "*")

            Create_TDA(.Tables.Add, "ICTWHSE1", "*", -1, False)
            Fill_Records("ICTWHSE1", String.Empty, True, "SELECT * FROM ICTWHSE1")

            Create_TDA(.Tables.Add, "ASTUSER1", "*", -1, False)
            Fill_Records("ASTUSER1", String.Empty, True, "SELECT * FROM ASTUSER1")

            Create_TDA(.Tables.Add, "WHTPKGM1", "*", -1, False)
            Fill_Records("WHTPKGM1", String.Empty, True, "SELECT * FROM WHTPKGM1")

            Create_TDA(.Tables.Add, "SOTSVIA1", "*", -1, False)
            Fill_Records("SOTSVIA1", String.Empty, True, "'SELECT * FROM SOTSVIA1 WHERE CARRIER_PROD_CODE IS NOT NULL")

            Create_TDA(.Tables.Add, "SOTCARR3", "*")
            .Tables("SOTCARR3").Columns.Add("CARRIER_REMOTE_HOST_IP", GetType(System.String))
            Fill_Records("SOTCARR3", "", True, "SELECT SOTCARR3.*, SOTCARR1.CARRIER_REMOTE_HOST_IP FROM SOTCARR3, SOTCARR1 WHERE SOTCARR3.CARRIER_CODE = SOTCARR1.CARRIER_CODE (+)")

            Create_TDA(.Tables.Add, "SOTCART1", "*")
            Create_TDA(.Tables.Add, "SOTCART2", "*")
            Create_Relation("SOTCART1", "SOTCART2", "CART_NO")

            Create_TDA(.Tables.Add, "WHTSHPC4", "*", 1)
            .Tables("WHTSHPC4").Columns.Add("SHIP_VIA_CODE", GetType(System.String))

            Create_TDA(.Tables.Add, "WHTSHPC1", "*")
            Create_TDA(.Tables.Add, "WHTSHPC2", "*")
            Create_TDA(.Tables.Add, "WHTSHPC5", "*")
            Create_TDA(.Tables.Add, "WHTSHPCG", "*")
            Create_TDA(.Tables.Add, "WHTSHPCS", "*")
            Create_TDA(.Tables.Add, "WHTSHPCC", "*")
            Create_TDA(.Tables.Add, "WHTSHPCP", "*")
        End With

        grdSOTTRCK1X.DataSource = dst.Tables("SOTTRCK1X")
        Create_Summary(grdSOTTRCK1X, "TRUCK_NO", "Count")

        grdSOTPICK1X.DataSource = dst.Tables("SOTPICK1X")
        Create_Summary(grdSOTPICK1X, "PICK_NO", "Count")

        grdSOTPICK2.DataSource = dst.Tables("SOTPICK2")
        Create_Summary(grdSOTPICK2, "PICK_LNO", "Count")

        grdWHTSHPC4.DataSource = dst.Tables("WHTSHPC4")

        grdSOTTRCK1X.Parent = tab.Parent
        splTotes.Parent = tab.Parent

        SetUpPortsAndPrinters()
        SetupScanner()
        CreateAppearances()
        txtUSER_ID.Text = ASCMAIN1.USER_ID

        Timer1.Start()
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = String.Empty

        Select Case eItemKey

            Case "Cancel"
                If Not AutoCancel Then
                    If MessageBox.Show($"Do you want to Cancel processing Truck {HFs("TRUCK_NO")}", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Refresh"

            Case "Request Rates"

            Case "Ship Package"
                If txtSHIP_VIA_CODE.TextLength = 0 Then
                    EMsg &= vbCr & "You must select one shipping method to ship the package."
                    Exit Sub
                End If

                Dim ErrorMessage As String = String.Empty
                If Not RequestShippingLabel("", ErrorMessage, False, "") Then
                    EMsg &= vbCr & ErrorMessage
                End If

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Select"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Me.Cursor = Cursors.WaitCursor
                Mode_Settings(False)
                Me.Cursor = Cursors.Default

            Case "Refresh"
                Me.Cursor = Cursors.WaitCursor
                Mode_Settings(False)
                Me.Cursor = Cursors.Default

            Case "Request Rates"
                Dim PICK_NO As String = dst.Tables("SOTPICK1").Rows(0).Item("PICK_NO")
                RequestRates(PICK_NO)

            Case "Ship Package"
                Dim ErrorMessage As String = String.Empty

                ' Create Invoice
                Dim PICK_NO As String = dst.Tables("SOTPICK1").Rows(0).Item("PICK_NO")
                If CreateSalesOrderInvoice(PICK_NO, txtSHIP_VIA_CODE.Text) Then
                    Dim INV_NO As String = dst.Tables("SOTINVH1").Rows(0).Item("INV_NO")
                    If Not RequestShippingLabel(INV_NO, ErrorMessage, False, PICK_NO) Then
                        MessageBox.Show(ErrorMessage, "Generate Shipping Label", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End If
                Mode_Settings(False)
        End Select

        ASCMAIN1.Progress("", "")

        Timer1.Start()

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("New Box").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Ship Package").Settings.Enabled = not_iScreenMode
            End With
        End If

        If ScreenMode Then
            splTotes.Visible = True
            txtTOTE_NO.ReadOnly = False
            txtTRUCK_NO.ReadOnly = True
        Else
            Clear_Record()
            splTotes.Visible = False
        End If

        grdSOTPICK1X.DisplayLayout.Bands(0).Columns("TRUCK_NO").Hidden = True

        Set_Read_Only(grpHeader, ScreenMode)

        If ScreenMode Then
            txtTOTE_NO.ReadOnly = False
        Else
            txtTOTE_NO.ReadOnly = True
        End If

    End Sub

    Private Sub Clear_Record()

        EnforceConstraints(False)
        For Each tableName As String In New String() {"SOTPICK0", "SOTPICK1", "SOTPICK2",
            "SOTORDR1", "SOTORDR2", "SOTORDR5", "TATEVNT1",
            "SOTINVH1", "SOTINVH2", "SOTINVH9", "SOTINVHM", "ARTOPEN1", "SOTRNGA1"}
            If dst.Tables.Contains(tableName) Then
                dst.Tables(tableName).Rows.Clear()
            End If
        Next
        EnforceConstraints(True)

        Clear_All_Filters(grdSOTPICK1X)
        Clear_All_Filters(grdSOTPICK2)
        Sort_grdColumns(grdSOTPICK1X, "TOTE_NO")

        txtTRUCK_NO.Clear()
        txtTOTE_NO.Clear()

        PICK_BATCH_NO = String.Empty
        WHSE_CODE_TRUCK = String.Empty

        grdSOTPICK2.Text = String.Empty
        FillPickTicketsInPick(String.Empty)

        txtTRUCK_NO.ReadOnly = False
        txtTOTE_NO.ReadOnly = True

        txtUSER_ID.Text = ASCMAIN1.USER_ID
        txtWHSE_CODE.Clear()
        txtSHIP_VIA_CODE.Clear()

        isCustomTruck = False
        dictCustomTruck.Clear()
        lblProcessingMode.Text = String.Empty
        drSOTTRCK1 = Nothing

        ASCMAIN1.MultiTask_Release()

    End Sub

    Private Sub Load_Record()
        Save_Header_Fields(grpHeader)
    End Sub

    Private Sub Update_Record()

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTTRCK1X, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTPICK1X, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTPICK2, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        If Not GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then
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

            Case "grdSOTORDR0"

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
        Else

            Select Case e.SourceControl.Name
                Case "grdSOTALLOX"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing

        If grd Is Nothing OrElse (grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow) Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case ""

        End Select

        Me.Cursor = Cursors.Default

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"

        End Select

    End Sub

#End Region

#Region "Form Procedures"

    Private Delegate Sub ScannerDelegate(ByVal ScannedString As String)
    Private scannedDelegate As ScannerDelegate = Nothing

    Private Sub BackOrderedAndCancelledItems(ByVal PICK_NO As String)

        ' All items in Back or Canc need to come out of Pick
        Dim drSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
        Dim WHSE_CODE As String = drSOTPICK1.Item("WHSE_CODE") & String.Empty
        Dim ORDR_NO As String = drSOTPICK1.Item("ORDR_NO") & String.Empty
        Dim drSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
        Dim PARTNER_CODE As String = drSOTORDR1.Item("PARTNER_CODE") & String.Empty

        For Each drSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}'")
            Dim ITEM_CODE As String = drSOTPICK2.Item("ITEM_CODE") & String.Empty
            Dim TakeOutOfPick As Int32 = Val(drSOTPICK2.Item("PICK_QTY_BACK") & String.Empty) + Val(drSOTPICK2.Item("PICK_QTY_CANC") & String.Empty)

            If TakeOutOfPick > 0 Then
                ASCMAIN1.sql = $"Update ICTSTAT2 SET WHSE_QTY_PICK = NVL(WHSE_QTY_PICK, 0) - :PARM1, 
                                    WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN, 0) + :PARM2
                                    WHERE ITEM_CODE = :PARM3 AND WHSE_CODE = :PARM4"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NNVV", New Object() {TakeOutOfPick, Val(drSOTPICK2.Item("PICK_QTY_BACK") & String.Empty), ITEM_CODE, WHSE_CODE})

                ASCMAIN1.sql = $"Update ICTSTAG2 SET WHSE_QTY_PICK = NVL(WHSE_QTY_PICK, 0) - :PARM1,
                                    WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN, 0) + :PARM2
                                    WHERE PARTNER_CODE = :PARM3 AND ITEM_CODE = :PARM4 AND WHSE_CODE = :PARM5"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NNVVV", New Object() {TakeOutOfPick, Val(drSOTPICK2.Item("PICK_QTY_BACK") & String.Empty), PARTNER_CODE, ITEM_CODE, WHSE_CODE})
            End If
        Next

    End Sub

    Private Function ProcessTote(ByVal TOTE_NO As String) As Boolean

        Dim PKG_WT As Decimal = 0

        Try
            Dim drSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Select($"TOTE_NO = '{TOTE_NO}'")(0)
            Dim PICK_NO As String = drSOTPICK1.Item("PICK_NO") & String.Empty
            Dim ORDR_NO As String = drSOTPICK1.Item("ORDR_NO") & String.Empty
            Dim WHSE_CODE As String = drSOTPICK1.Item("WHSE_CODE") & String.Empty

            Dim drSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
            drSOTORDR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            drSOTORDR1.Item("LAST_DATE") = DateTime.Now + ASCMAIN1.NowTSD
            drSOTPICK1.Item("PICK_STATUS") = "F"
            drSOTPICK1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            drSOTPICK1.Item("LAST_DATE") = DateTime.Now + ASCMAIN1.NowTSD

            For Each drSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}'", "ITEM_CODE")

                Dim PICK_LNO As Int16 = drSOTPICK2.Item("PICK_LNO")
                Dim PICK_QTY_CONF As Int16 = Val(drSOTPICK2.Item("PICK_QTY_CONF") & String.Empty)
                Dim PICK_QTY As Int16 = Val(drSOTPICK2.Item("PICK_QTY") & String.Empty)
                Dim PICK_QTY_SCAN As Int16 = Val(drSOTPICK2.Item("PICK_QTY_SCAN") & String.Empty)
                Dim ITEM_CODE As String = drSOTPICK2.Item("ITEM_CODE") & String.Empty

                ' Cancel any unpicked quantities
                If PICK_QTY_SCAN < PICK_QTY Then
                    drSOTPICK2.Item("PICK_QTY_CANC") = Val(drSOTPICK2.Item("PICK_QTY_CANC") & String.Empty) + (PICK_QTY - PICK_QTY_SCAN)
                End If

                Dim ORDR_LNO As Int32 = drSOTPICK2.Item("ORDR_LNO")
                Dim drSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})

                ' Change on 5/16/2022 - Back Ordered Pick Tickets are coming to me with no Pick Qty and SOTORDR2.ORDR_QTY_BACK is already updated.
                If Val(drSOTPICK2.Item("PICK_QTY") & String.Empty) > 0 Then
                    drSOTORDR2.Item("ORDR_QTY_BACK") = Val(drSOTORDR2.Item("ORDR_QTY_BACK") & String.Empty) + Val(drSOTPICK2.Item("PICK_QTY_BACK") & String.Empty)
                    drSOTORDR2.Item("ORDR_QTY_CANC") = Val(drSOTORDR2.Item("ORDR_QTY_CANC") & String.Empty) + Val(drSOTPICK2.Item("PICK_QTY_CANC") & String.Empty)
                End If

                drSOTORDR2.Item("ORDR_QTY_SHIP") = Val(drSOTORDR2.Item("ORDR_QTY_SHIP") & String.Empty) + Val(drSOTPICK2.Item("PICK_QTY_CONF") & String.Empty)
                drSOTORDR2.Item("ORDR_QTY_PICK") = 0

                If Val(drSOTORDR2.Item("ORDR_QTY_BACK") & String.Empty) < 0 Then drSOTORDR2.Item("ORDR_QTY_BACK") = 0
                If Val(drSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty) < 0 Then drSOTORDR2.Item("ORDR_QTY_PICK") = 0

                If Val(drSOTORDR2.Item("ORDR_QTY_BACK") & String.Empty) > 0 Then
                    drSOTORDR2.Item("ORDR_LINE_STATUS") = "B"
                ElseIf Val(drSOTORDR2.Item("ORDR_QTY_OPEN") & String.Empty) = 0 Then
                    drSOTORDR2.Item("ORDR_LINE_STATUS") = "F"
                End If
            Next

            PKG_WT = 0

            If dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' and ISNULL(ORDR_QTY_BACK, 0) > 0").Length = 0 _
                            AndAlso dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' and ISNULL(ORDR_QTY_OPEN, 0) > 0").Length = 0 Then
                drSOTORDR1.Item("ORDR_STATUS") = "F"
            End If

            ProcessTote = True

        Catch ex As Exception
            Return False
        End Try

    End Function

    Private Enum ValidateTruckToteTypes
        Truck
        Tote
    End Enum

    Private Function ValidateTruckTote(ByVal ValidationType As ValidateTruckToteTypes, ByVal InputValue As String) As Boolean

        Try
            Dim drSOTPICK1 As DataRow = Nothing
            Me.Cursor = Cursors.WaitCursor

            allItemsOnBackOrder = False

            Select Case ValidationType
                Case ValidateTruckToteTypes.Truck
                    WHSE_CODE_TRUCK = String.Empty
                    isCustomTruck = False
                    drSOTTRCK1 = Nothing

                    ASCMAIN1.Progress("Loading Truck", InputValue)

                    If InputValue.Length = 0 Then
                        MessageBox.Show("The supplied Truck is invalid.", "Validate Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    drSOTTRCK1 = Fill_Record("SOTTRCK1", InputValue)
                    If dst.Tables("SOTTRCK1").Rows.Count = 0 OrElse drSOTTRCK1 Is Nothing Then
                        MessageBox.Show("The supplied Truck is invalid or does not have any Totes to process", "Validate Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    If drSOTTRCK1.Item("PICK_BATCH_NO") & String.Empty = String.Empty Then
                        MessageBox.Show("The supplied Truck is NOT assigned to a Pick Batch", "Validate Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    PICK_BATCH_NO = drSOTTRCK1.Item("PICK_BATCH_NO") & String.Empty

                    Fill_Records("SOTPICK0", PICK_BATCH_NO)
                    If dst.Tables("SOTPICK0").Rows.Count = 0 Then
                        MessageBox.Show("The supplied Truck is NOT assigned to a Pick Batch that can be found.", "Validate Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    Dim drSOTPICK0 As DataRow = dst.Tables("SOTPICK0").Rows(0)

                    Dim PICK_BATCH_STATUS As String = drSOTPICK0.Item("PICK_BATCH_STATUS") & String.Empty
                    If Not "NK".Contains(PICK_BATCH_STATUS) Then
                        MessageBox.Show($"The supplied Truck's Pick Batch Status must be Picked or In Pack.", "Validate Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    WHSE_CODE_TRUCK = drSOTPICK0.Item("WHSE_CODE") & String.Empty
                    txtWHSE_CODE.Text = WHSE_CODE_TRUCK

                    ASCMAIN1.Progress("Fill Pick Tickets In Pick", "")
                    FillPickTicketsInPick(InputValue)

                    If dst.Tables("SOTPICK1X").Rows.Count = 0 Then
                        MessageBox.Show("The supplied Truck does not have any Totes to process", "Validate Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    Dim numIncomplete As Int32 = dst.Tables("SOTPICK1X").Select("INCOMPLETE = '1'").Length
                    If numIncomplete > 0 Then
                        Dim zMsg As String = "The following Totes are not Complete:" & Environment.NewLine
                        For Each drSOTTRCK1X As DataRow In dst.Tables("SOTPICK1X").Select("INCOMPLETE = '1'")
                            zMsg &= drSOTTRCK1X.Item("TOTE_NO") & Environment.NewLine
                        Next

                        MessageBox.Show(zMsg, "Process Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        dst.Tables("SOTPICK1X").Rows.Clear()
                        Clear_Record()
                        Return False
                    End If

                    Dim PICK_REQ_RES As Int32 = dst.Tables("SOTPICK1X").Select("PICK_REQ_RES = '1'").Length
                    If PICK_REQ_RES > 0 Then
                        Dim zMsg As String = "The following Totes Require Resolution:" & Environment.NewLine
                        For Each drSOTTRCK1X As DataRow In dst.Tables("SOTTRCK1X").Select("PICK_REQ_RES = '1'")
                            zMsg &= drSOTTRCK1X.Item("TOTE_NO") & Environment.NewLine
                        Next

                        MessageBox.Show(zMsg, "Process Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        dst.Tables("SOTTRCK1X").Rows.Clear()
                        Clear_Record()
                        Return False
                    End If

                    If Not ASCMAIN1.Logical_Lock("SOTTRCK1", InputValue) Then
                        dst.Tables("SOTTRCK1X").Rows.Clear()
                        Clear_Record()
                        Return False
                    End If

                    If Not ASCMAIN1.Logical_Lock("SOTPICK0", PICK_BATCH_NO) Then
                        dst.Tables("SOTTRCK1X").Rows.Clear()
                        Clear_Record()
                        Return False
                    End If

                    drSOTPICK1 = dst.Tables("SOTTRCK1X").Rows(0)

                    ASCMAIN1.Progress("Load Pick Tickets and Sales Orders", "")
                    Dim lstPICK_NOs As New List(Of String)
                    Dim lstORDR_NOs As New List(Of String)

                    For Each drSOTTRCK1X As DataRow In dst.Tables("SOTTRCK1X").Select("")

                        Dim PICK_NO As String = drSOTTRCK1X.Item("PICK_NO") & String.Empty
                        Dim ORDR_NO As String = drSOTTRCK1X.Item("ORDR_NO") & String.Empty
                        Dim TOTE_NO As String = drSOTTRCK1X.Item("TOTE_NO") & String.Empty

                        ASCMAIN1.Progress("Load Pick Tickets and Sales Orders", TOTE_NO)

                        If Not ASCMAIN1.Logical_Lock("SOTTOTE1", TOTE_NO) Then
                            Clear_Record()
                            Return False
                        End If

                        If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO) Then
                            Clear_Record()
                            Return False
                        End If

                        If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then
                            Clear_Record()
                            Return False
                        End If

                        lstPICK_NOs.Add(PICK_NO)
                        lstORDR_NOs.Add(ORDR_NO)
                    Next

                    ASCMAIN1.Progress("Load Pick Tickets and Sales Orders", "")
                    Fill_Records("SOTPICK1", String.Join(",", lstPICK_NOs.ToArray))
                    Fill_Records("SOTPICK2", String.Join(",", lstPICK_NOs.ToArray))
                    Fill_Records("SOTORDR1", String.Join(",", lstORDR_NOs.ToArray))
                    Fill_Records("SOTORDR2", String.Join(",", lstORDR_NOs.ToArray))
                    Fill_Records("SOTORDR5", String.Join(",", lstORDR_NOs.ToArray))
                    Fill_Records("ARTCUST2", String.Join(",", lstORDR_NOs.ToArray))


                    If dst.Tables("SOTTRCK1X").Rows.Count > 0 Then
                        grdSOTPICK1X.ActiveRow = grdSOTPICK1X.Rows(0)
                        grdSOTPICK1X_AfterRowActivate(Nothing, Nothing)
                    End If

                    Dim numOrdertypes As Int16 = ASCDATA1.SelectDistinct(dst.Tables("SOTORDR1"), New String() {"ORDR_TYPE_CODE"}).Rows.Count
                    If numOrdertypes > 1 Then
                        MessageBox.Show("The supplied Truck's sales orders have more than 1 Order Type Code", "Validate Truck", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Clear_Record()
                        Return False
                    End If

                    Dim CUST_CODE As String = drSOTPICK1.Item("CUST_CODE") & String.Empty
                    Dim CUST_SHIP_TO_NO As String = drSOTPICK1.Item("CUST_SHIP_TO_NO") & String.Empty
                    Dim SHIP_VIA_CODE As String = drSOTPICK1.Item("SHIP_VIA_CODE") & String.Empty

                    Dim CUST_SHIP_TO_DC As String = drSOTPICK1.Item("CUST_SHIP_TO_DC") & String.Empty
                    ' See if shipping to a DC
                    If CUST_SHIP_TO_DC.Length > 0 Then
                        CUST_SHIP_TO_NO = CUST_SHIP_TO_DC
                    End If

                    isCustomTruck = drSOTTRCK1.Item("TRUCK_TYPE") & String.Empty = "X"

                    ' Message if there are pick tickets where all items are back ordered.
                    Dim lstTotes As New List(Of String)
                    For Each drSOTPICK1 In dst.Tables("SOTTRCK1X").Select
                        Dim PICK_NO As String = drSOTPICK1.Item("PICK_NO") & String.Empty
                        Dim TOTE_NO As String = drSOTPICK1.Item("TOTE_NO") & String.Empty

                        If dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}' AND ISNULL(PICK_QTY_CONF, 0) > 0").Length = 0 Then
                            lstTotes.Add(TOTE_NO)
                            drSOTPICK1.Item("ALL_ITEMS_BACK") = "1"
                        End If
                    Next

                    If lstTotes.Count > 0 Then
                        MessageBox.Show($"The following totes have all items back ordered. You must scan these totes to close them out.{Environment.NewLine}{String.Join(Environment.NewLine, lstTotes.ToArray)}", "Back Orders", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                    Click_Command("Select")
                    Return True

                Case ValidateTruckToteTypes.Tote

                    ASCMAIN1.Progress("Validate Tote", InputValue)

                    drSOTPICK1 = Nothing
                    If dst.Tables("SOTTRCK1X").Select($"TOTE_NO = '{InputValue}'").Length = 1 Then
                        drSOTPICK1 = dst.Tables("SOTTRCK1X").Select($"TOTE_NO = '{InputValue}'")(0)
                    Else
                        MessageBox.Show($"Cannot locate Tote: {InputValue}", "Validate Tote", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    If drSOTPICK1.Item("SELECTED") & String.Empty = "1" Then
                        MessageBox.Show($"Tote: {InputValue} was already processed.", "Validate Tote", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    If drSOTPICK1.Item("PICK_REQ_RES") & String.Empty = "1" Then
                        MessageBox.Show($"Tote: {InputValue} Requires Resolution.", "Validate Tote", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    If drSOTPICK1.Item("INCOMPLETE") & String.Empty = "1" Then
                        MessageBox.Show($"Tote: {InputValue} is incomplete.", "Validate Tote", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If

                    ' Do not want this to stop procesisng if IT causes an error
                    Try
                        For Each grdRow As UltraWinGrid.UltraGridRow In grdSOTPICK1X.Rows
                            If grdRow.Band.Key = grdSOTPICK1X.DisplayLayout.Bands(0).Key Then
                                If grdRow.Cells("TOTE_NO").Value & String.Empty = InputValue Then
                                    grdSOTPICK1X.ActiveRow = grdRow
                                    grdSOTPICK1X.Selected.Rows.Clear()
                                    grdSOTPICK1X.Selected.Rows.Add(grdRow)
                                    grdSOTPICK1X.DisplayLayout.RowScrollRegions(0).FirstRow = grdRow
                                    Exit For
                                End If
                            End If
                        Next
                    Catch ex As Exception

                    End Try

                    drSOTPICK1.Item("SELECTED") = "1"

                    If Not ProcessTote(InputValue) Then
                        drSOTPICK1.Item("SELECTED") = "0"
                        drSOTPICK1.Item("BOX_NO") = String.Empty
                        Dim PICK_NO As String = drSOTPICK1.Item("PICK_NO") = String.Empty
                        drSOTPICK1.Item("PICK_STATUS") = "P"
                        ASCMAIN1.sql = "UPDATE SOTPICK1 SET PICK_STATUS = 'P' WHERE PICK_NO = :PARM1 AND PICK_STATUS = 'F'"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PICK_NO})
                        Return False
                    End If

                    Return True

            End Select

        Catch ex As Exception
            Me.Cursor = Cursors.Default

            MessageBox.Show($"Validate Error: {ex.Message}", "Validate Tote", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If ValidationType = ValidateTruckToteTypes.Truck Then
                Clear_Record()
            End If
            Return False

        Finally
            ASCMAIN1.Progress("", "")
            dst.Tables("SOTINVH1").Rows.Clear()
            dst.Tables("SOTINVH2").Rows.Clear()
            Me.Cursor = Cursors.Default
            Timer1.Start()
        End Try
    End Function

    Private Sub FillPickTicketsInPick(ByVal TRUCK_NO As String)
        grdSOTPICK2.Text = String.Empty
        Fill_Records("SOTPICK1X", New Object() {TRUCK_NO, PICK_BATCH_NO})
        grdSOTPICK1X.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
        Clear_All_Filters(grdSOTPICK1X)
        Sort_grdColumns(grdSOTPICK1X, "TOTE_NO")
        If dst.Tables("SOTTRCK1X").Rows.Count > 0 Then
            grdSOTPICK1X.ActiveRow = grdSOTPICK1X.Rows(0)
        End If
    End Sub

    Private Sub Record_Event(ByVal ORDR_NO As String, ByVal EVENT_CODE As String, ByVal EVENT_DESC As String)

        Dim drSOTORDRE As DataRow = dst.Tables("SOTORDRE").NewRow
        drSOTORDRE.Item("ORDR_NO") = ORDR_NO
        drSOTORDRE.Item("INIT_DATE") = DateTime.Now + ASCMAIN1.NowTSD
        drSOTORDRE.Item("INIT_OPER") = ASCMAIN1.USER_ID
        drSOTORDRE.Item("EVENT_CODE") = EVENT_CODE
        If EVENT_DESC.Trim.Length > dst.Tables("SOTORDRE").Columns("EVENT_DESC").MaxLength Then
            EVENT_DESC = EVENT_DESC.Trim.Substring(0, dst.Tables("SOTORDRE").Columns("EVENT_DESC").MaxLength).Trim
        End If
        drSOTORDRE.Item("EVENT_DESC") = EVENT_DESC
        dst.Tables("SOTORDRE").Rows.Add(drSOTORDRE)
    End Sub

    Private Function CreateSalesOrderInvoice(ByVal PICK_NO As String, ByVal SelectedShipViaCode As String) As Boolean
        Try
            dst.Tables("SOTINVH1").Rows.Clear()
            dst.Tables("SOTINVH2").Rows.Clear()
            dst.Tables("SOTINVH9").Rows.Clear()
            dst.Tables("SOTINVHM").Rows.Clear()

            Dim RFIXMSG As Boolean = False
            Dim drSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
            Dim SHIP_BOL_NO As String = drSOTPICK1.Item("SHIP_BOL_NO") & String.Empty
            Fill_Records("SOTSHIP1", SHIP_BOL_NO)
            Dim drSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows(0)

            'rowSOTSHIP1.ITEM("SHIP_BOL_NO") = ""
            drSOTSHIP1.Item("SHIP_DATE_SHIPPED") = DateTime.Now.ToShortDateString
            drSOTSHIP1.Item("SHIP_VIA_CODE") = SelectedShipViaCode
            'rowSOTSHIP1.ITEM("SHIP_REF") = ""
            drSOTSHIP1.Item("SHIP_TOTAL_WGT") = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "") & String.Empty)
            drSOTSHIP1.Item("SHIP_CNT_CARTONS") = dst.Tables("SOTCART1").Rows.Count
            'rowSOTSHIP1.ITEM("SHIP_ADDR_TYPE") = ""
            'rowSOTSHIP1.ITEM("SHIP_ADDR_CODE") = ""
            'rowSOTSHIP1.ITEM("ORDR_GROUP_NO") = ""
            'rowSOTSHIP1.ITEM("SHIP_PICK_PRINTED") = ""
            'rowSOTSHIP1.ITEM("PICK_BATCH_NO") = ""
            drSOTSHIP1.Item("SHIP_STATUS") = "F"
            'rowSOTSHIP1.ITEM("SHIP_PULL_BY_STYLE") = ""
            'rowSOTSHIP1.ITEM("SHIP_856_BATCH_NO") = ""
            'rowSOTSHIP1.ITEM("FRT_TERMS") = ""
            'rowSOTSHIP1.ITEM("WHSE_CODE") = ""
            drSOTSHIP1.Item("INV_DATE") = DateTime.Now.ToShortDateString
            'rowSOTSHIP1.ITEM("SHIP_MANIFEST_NO") = ""
            'rowSOTSHIP1.ITEM("SHIP_810_BATCH_NO") = ""
            'rowSOTSHIP1.ITEM("INIT_DATE") = ""
            'rowSOTSHIP1.ITEM("INIT_OPER") = ""
            drSOTSHIP1.Item("LAST_DATE") = DateTime.Now
            drSOTSHIP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            'rowSOTSHIP1.ITEM("BILL_OF_LADING_NO") = ""
            'rowSOTSHIP1.ITEM("REGISTER_XNO") = ""
            'rowSOTSHIP1.ITEM("REASON_CODE") = ""
            'rowSOTSHIP1.ITEM("SHIP_BOL_NO_REV") = ""
            'rowSOTSHIP1.ITEM("TERM_CODE") = ""
            'rowSOTSHIP1.ITEM("SREP_CODE") = ""
            'rowSOTSHIP1.ITEM("ORDR_DEPT") = ""
            drSOTSHIP1.Item("SHIP_DATE_RECEIVED") = DateTime.Now.ToShortDateString
            'rowSOTSHIP1.ITEM("SHIP_NOTES") = ""
            drSOTSHIP1.Item("SHIPPED_ACTUAL") = DateTime.Now.ToShortDateString
            'rowSOTSHIP1.ITEM("CUST_FACTOR_TRANS_IND") = ""
            'rowSOTSHIP1.ITEM("SHIP_SEAL_NO") = ""
            'rowSOTSHIP1.ITEM("SHIP_BOL_NO_ORIG") = ""
            'rowSOTSHIP1.ITEM("SHIP_BOL_NO_SPLIT") = ""
            'rowSOTSHIP1.ITEM("SREP2_CODE") = ""
            'rowSOTSHIP1.ITEM("BOL_PRINTED") = ""
            'rowSOTSHIP1.ITEM("SHIP_SPEC_INST") = ""
            'rowSOTSHIP1.ITEM("FACTOR_TRANS_BATCH_LAST") = ""
            'rowSOTSHIP1.ITEM("FACTOR_TRANS_LAST_OPER") = ""
            'rowSOTSHIP1.ITEM("FACTOR_TRANS_LAST_DATE") = ""
            'rowSOTSHIP1.ITEM("MASTER_SHIP_BOL_NO") = ""
            'rowSOTSHIP1.ITEM("SHIP_940_BATCH_NO") = ""
            'rowSOTSHIP1.ITEM("SHIP_753_IND") = ""
            'rowSOTSHIP1.ITEM("HANDLING_TYPE") = ""
            'rowSOTSHIP1.ITEM("HANDLING_UNITS") = ""
            'rowSOTSHIP1.ITEM("GEN_IND") = ""
            'rowSOTSHIP1.ITEM("GEN_XNO") = ""
            'rowSOTSHIP1.ITEM("GEN_DATE") = ""
            'rowSOTSHIP1.ITEM("DOCUMENTKEY") = ""
            'rowSOTSHIP1.ITEM("THIRD_PARTY") = ""
            'rowSOTSHIP1.ITEM("OPT_LINE1") = ""
            'rowSOTSHIP1.ITEM("OPT_LINE2") = ""
            'rowSOTSHIP1.ITEM("SHIP_DATE_PACKED") = ""
            'rowSOTSHIP1.ITEM("LP_STATUS") = ""
            'rowSOTSHIP1.ITEM("LP_XNO") = ""
            'rowSOTSHIP1.ITEM("MASTER_BILL_OF_LADING_NO") = ""
            'rowSOTSHIP1.ITEM("SHIP_TRAILER_NO") = ""
            'rowSOTSHIP1.ITEM("SHIP_LOAD_NO") = ""
            'rowSOTSHIP1.ITEM("SHIP_APPT_NO") = ""
            'rowSOTSHIP1.ITEM("ORDR_PICK_TYPE") = ""
            'rowSOTSHIP1.ITEM("SHIP_856_IND") = ""
            'rowSOTSHIP1.ITEM("SHIP_810_IND") = ""
            'rowSOTSHIP1.ITEM("LP_XMIT_DATE") = ""
            'rowSOTSHIP1.ITEM("LP_CODE") = ""
            'rowSOTSHIP1.ITEM("OPS_YYYYPP") = ""
            'rowSOTSHIP1.ITEM("SHIP_CART_REQD") = ""
            'rowSOTSHIP1.ITEM("EDI_856_CREATED") = ""
            'rowSOTSHIP1.ITEM("EDI_810_CREATED") = ""
            'rowSOTSHIP1.ITEM("SHIP_753_BATCH_NO") = ""
            'rowSOTSHIP1.ITEM("SHIP_DATE_PLANNED") = ""
            'rowSOTSHIP1.ITEM("SHIP_DATE_ROUTED") = ""
            'rowSOTSHIP1.ITEM("SHIP_NOTES_3PL") = ""
            'rowSOTSHIP1.ITEM("INSURED_VALUE") = ""
            'rowSOTSHIP1.ITEM("INSURED_SHIPMENT") = ""
            'rowSOTSHIP1.ITEM("EDI_LOAD_ID") = ""
            'rowSOTSHIP1.ITEM("SHIP_WAVE_STATUS") = ""
            'rowSOTSHIP1.ITEM("WAVE_NO") = ""
            'rowSOTSHIP1.ITEM("BTB_BOL_NO") = ""

            Dim SOCINVH1 As New TAC.SOCINVH1(dst)

            SOCINVH1.ProcessPickTicketsAndUpdateSalesDetails(DateTime.Now.ToShortDateString)

            ' Record event where the Ship via was changed
            Dim ORDR_NO As String = drSOTPICK1.Item("ORDR_NO")
            Dim drSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
            If drSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty <> SelectedShipViaCode Then
                Dim drTATEVNT1 As DataRow = dst.Tables("TATEVNT1").Rows.Add
                drTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
                drTATEVNT1.Item("TABLE_KEY") = ORDR_NO
                drTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
                drTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                drTATEVNT1.Item("EVENT_TYPE") = "SHPMTC"
                drTATEVNT1.Item("EVENT_DESC") = $"Ship Via was changed from {drSOTORDR1.Item("SHIP_VIA_CODE")} to {SelectedShipViaCode}"
                drTATEVNT1.Item("EVENT_KEY") = ""
                drTATEVNT1.Item("FORM_NAME") = "SOFSHIPE"
            End If

            If dst.Tables("SOTPICK2").Select("PICK_QTY > 0 AND PICK_QTY_CONF < PICK_QTY", "").Length > 0 Then
                Dim drTATEVNT1 As DataRow = dst.Tables("TATEVNT1").Rows.Add
                drTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
                drTATEVNT1.Item("TABLE_KEY") = ORDR_NO
                drTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
                drTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                drTATEVNT1.Item("EVENT_TYPE") = "SHSHP"
                drTATEVNT1.Item("EVENT_DESC") = "User chose to short ship Ecommerce order."
                drTATEVNT1.Item("EVENT_KEY") = ""
                drTATEVNT1.Item("FORM_NAME") = "SOFSHIPE"
            End If

            Dim CUST_FACTOR_TRANS_IND As String = "0"

            ' Log factoring change
            If Val(drSOTSHIP1.Item("CUST_FACTOR_TRANS_IND") & String.Empty) <> Val(CUST_FACTOR_TRANS_IND) Then
                Dim drTATEVNT1 As DataRow = dst.Tables("TATEVNT1").Rows.Add
                drTATEVNT1.Item("TABLE_NAME") = "SOTSHIP1"
                drTATEVNT1.Item("TABLE_KEY") = SHIP_BOL_NO
                drTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
                drTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                drTATEVNT1.Item("EVENT_TYPE") = "SHPFAC"
                drTATEVNT1.Item("EVENT_DESC") = "Factor Setting was changed from " _
                    & IIf(Val(drSOTSHIP1.Item("CUST_FACTOR_TRANS_IND") & String.Empty) = 1, "True", "False") & " to " & IIf(Val(CUST_FACTOR_TRANS_IND) = 1, "True", "False")
                drTATEVNT1.Item("EVENT_KEY") = ""
                drTATEVNT1.Item("FORM_NAME") = "SOFSHIP0"
            End If

            ' Needed when Back Orders exist and nothing is getting shipped.
            ' Also cancel a sales order when all items are cancelled.
            ASCMAIN1.sql = $"Begin Declare Cursor C1 Is
                                 Select ORDR_NO
                                    , Sum (NVL(ORDR_QTY_OPEN, 0)) ORDR_QTY_OPEN
                                    , Sum (NVL(ORDR_QTY_BACK, 0)) ORDR_QTY_BACK
                                    , Sum (NVL(ORDR_QTY_PICK, 0)) ORDR_QTY_PICK
                                    , Sum (NVL(ORDR_QTY_SHIP, 0)) ORDR_QTY_SHIP
                                  From SOTORDR2 Where ORDR_NO = :PARM1 group by ORDR_NO;
                                 Begin For R1 In C1 Loop
                                    Update SOTORDR1 Set
                                      ORDR_STATUS = 
                                        CASE WHEN R1.ORDR_QTY_OPEN > 0 OR R1.ORDR_QTY_BACK > 0 THEN 'O'
                                             ELSE CASE WHEN R1.ORDR_QTY_PICK > 0 THEN 'P'
                                                       ELSE CASE WHEN R1.ORDR_QTY_SHIP > 0 THEN 'F'
                                                                 ELSE 'C' END END END
                                    where ORDR_NO = R1.ORDR_NO;
                                 End Loop; End;
                                End;"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {ORDR_NO})

            ' If nothing to Invoice then get out of here. All items are back ordered.
            allItemsOnBackOrder = dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}' AND ISNULL(PICK_QTY_CONF, 0) > 0 ").Length = 0

            If allItemsOnBackOrder Then
                drSOTPICK1.Item("PICK_STATUS") = "C"
            Else
                SOCINVH1.CreateInvoices(SHIP_BOL_NO, RFIXMSG)
            End If

            Try
                BeginTrans()
                Update_Record_TDA("SOTORDR1")
                Update_Record_TDA("SOTORDR2")
                Update_Record_TDA("SOTORDR5")

                Update_Record_TDA("SOTPICK1")
                Update_Record_TDA("SOTPICK2")
                Update_Record_TDA("SOTSHIP1")

                Update_Record_TDA("SOTINVH1")
                Update_Record_TDA("SOTINVH2")
                Update_Record_TDA("SOTINVH9")
                Update_Record_TDA("SOTINVHM")

                Update_Record_TDA("TATEVNT1")

                Dim PICK_BATCH_NO As String = drSOTTRCK1.Item("PICK_BATCH_NO") & String.Empty

                'CHANGE SOTPICK0.PICK_STATUS FROM P -> K (IN PACK)
                If dst.Tables("SOTPICK1X").Select("SELECTED = '1'").Length = "1" Then
                    ASCMAIN1.sql = $"UPDATE SOTPICK0 SET PICK_BATCH_STATUS = 'K' WHERE PICK_BATCH_STATUS IN ('N', 'K') AND TRUCK_NO = :PARM1 AND PICK_BATCH_NO = :PARM2 AND DC_CODE = :PARM3"
                    Dim numrows As Int16 = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New Object() {HFs("TRUCK_NO"), PICK_BATCH_NO, HFs("DC_CODE")})
                    If numrows <> 1 Then
                        Throw New Exception("Unable to properly Update SOTPICK0.PICK_BATCH_STATUS")
                    End If
                End If

                'WHEN PACKING STATION IS "COMPLETE" WITH TRANSFERRING ORDERS IN PICK INTO PACK
                If dst.Tables("SOTTRCK1X").Select("SELECTED = '1'").Length = dst.Tables("SOTTRCK1X").Rows.Count Then
                    ASCMAIN1.Progress("Finalizing Truck", "SOTPICK0")
                    'CHANGE SOTPICK0.PICK_STATUS FROM K -> F (FINISHED)
                    ASCMAIN1.sql = "UPDATE SOTPICK0 SET PICK_BATCH_STATUS = 'F' WHERE PICK_BATCH_STATUS = 'K' AND PICK_BATCH_NO = :PARM1"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {PICK_BATCH_NO})

                    'CLEAR PICK BATCH FROM TRUCK
                    ASCMAIN1.Progress("Finalizing Truck", "SOTPICK1")
                    ASCMAIN1.sql = "UPDATE SOTTRCK1 SET PICK_BATCH_NO = NULL WHERE PICK_BATCH_NO = :PARM1"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {PICK_BATCH_NO})

                    'CLEAR PICK NO FROM TOTE (ALL TRUCK TYPES)
                    ASCMAIN1.Progress("Finalizing Truck", "SOTTOTE1")
                    'ASCMAIN1.sql = "UPDATE SOTTOTE1 SET PICK_NO = NULL WHERE TRUCK_NO = :PARM1"
                    'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {HFs("TRUCK_NO")})
                    If drSOTTRCK1.Item("TRUCK_TYPE") = "R" Then
                        ASCMAIN1.sql = $"UPDATE SOTTOTE1 SET PICK_NO = NULL, SLOT_NO = NULL, TRUCK_NO = NULL, INIT_OPER = '{ASCMAIN1.USER_ID}' WHERE PICK_NO IN (SELECT PICK_NO FROM SOTPICK1 WHERE PICK_BATCH_NO = :PARM1)"
                    Else
                        ASCMAIN1.sql = "UPDATE SOTTOTE1 SET PICK_NO = NULL WHERE PICK_NO IN (SELECT PICK_NO FROM SOTPICK1 WHERE PICK_BATCH_NO = :PARM1)"
                    End If
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {PICK_BATCH_NO})

                    'IF TRUCK IS A CUSTOM TRUCK, CHANGE TRUCK TYPE FROM X -> R
                    ASCMAIN1.Progress("Finalizing Truck", "SOTTRCK1")
                    If drSOTTRCK1.Item("TRUCK_TYPE") = "X" Then
                        ASCMAIN1.sql = "UPDATE SOTTRCK1 SET TRUCK_TYPE = 'R' WHERE TRUCK_TYPE = 'X' AND TRUCK_NO = :PARM1 AND DC_CODE = :PARM2"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {HFs("TRUCK_NO"), HFs("DC_CODE")})
                    End If

                    'DELETE CUSTOM TOTES THAT BELONG TO THE CUSTOM TRUCK
                    ASCMAIN1.Progress("Finalizing Truck", "SOTTOTE1")
                    ASCMAIN1.sql = "DELETE FROM SOTTOTE1 WHERE TRUCK_NO = :PARM1 AND DC_CODE = :PARM2 AND TOTE_TYPE = 'X'"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {HFs("TRUCK_NO"), HFs("DC_CODE")})
                End If

                CommitTrans()
            Catch ex As Exception
                Rollback(ex.Message)
                Return False
            End Try

            txtPickNo.Text = PICK_NO
            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Sends the Scanned Bar Code to the Appropriate Control based on the Current Processing State
    ''' </summary>
    ''' <param name="scannedData"></param>
    ''' <remarks></remarks>
    Private Sub ProcessScannedData(ByVal scannedData As String)

        Static dataReceived As String

        dataReceived += scannedData
        If InStr(dataReceived, Chr(13), CompareMethod.Text) = 0 Then
            Exit Sub
        End If

        Dim sender As Object = Nothing
        Dim e As New System.Windows.Forms.KeyEventArgs(Keys.Enter)

        ' Trim Off line feeds
        dataReceived = Replace(dataReceived, Chr(10), String.Empty)
        dataReceived = Replace(dataReceived, Chr(13), String.Empty)

        If ScreenMode Then
            sender = txtTRUCK_NO
            txtTRUCK_NO.Clear()
            txtTRUCK_NO.Focus()
            txtTRUCK_NO.Text = dataReceived
        Else
            sender = txtTOTE_NO
            txtTOTE_NO.Clear()
            txtTOTE_NO.Focus()
            txtTOTE_NO.Text = dataReceived
        End If

        txt_KeyDown(sender, e)
        dataReceived = String.Empty

    End Sub

#End Region

#Region "Carrier Procedures"

    Private Sub RequestRates(ByVal PICK_NO As String)

        Try
            Me.Cursor = Cursors.WaitCursor
            Dim drWHTSHPC4 As DataRow = Nothing
            Dim drWHTSHPCA As DataRow = Nothing

            Dim CARRIER_SURCHARGE_PERC As Int16 = 0
            Dim FRT_PER_SALES_HOLD As Int16 = 0
            Dim CARRIER_PPA_TYPE As String = "L"
            Dim CARRIER_SURCHARGE_BASE As String = "L"
            Dim drSOTCARR1 As DataRow = Nothing

            For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select($"PICK_NO = '{PICK_NO}'", "", DataViewRowState.CurrentRows)
                If Val(drSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") & String.Empty) <= 0 Then
                    MessageBox.Show("All Cartons must have a weight.", "Request Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            Next

            ASCMAIN1.Progress("Request Carrier Rates")

            dst.Tables("WHTSHPC4").Rows.Clear()

            Dim rUPSList(1) As WHCSHIP1.RateList
            Dim rFEDEXList(1) As WHCSHIP1.RateList
            Dim rUPSFreightList(1) As WHCSHIP1.RateList
            Dim rUSPSList(1) As WHCSHIP1.RateList

            'Me.Cursor = Cursors.WaitCursor
            'ASCMAIN1.Progress("-", "USPS")
            'rUSPSList = GetUSPSRates()

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("-", "UPS")
            rUPSList = GetUpsRates()

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("-", "FedEx")
            rFEDEXList = GetFedExRates()

            Me.Cursor = Cursors.WaitCursor

            If rUSPSList Is Nothing Then
                ReDim rUSPSList(1)
            End If

            If rUPSList Is Nothing Then
                ReDim rUPSList(1)
            End If

            If rFEDEXList Is Nothing Then
                ReDim rFEDEXList(1)
            End If

            If rUPSFreightList Is Nothing Then
                ReDim rUPSFreightList(1)
            End If

            Dim selected As Boolean = False
            Dim CARRIER_CODE As String = String.Empty

            For iCtr As Int16 = 1 To 4
                Dim freightShipment As String = " and ISNULL(FREIGHT_SHIPMENT, '0') = '0'"
                Dim rList(1) As WHCSHIP1.RateList

                Select Case iCtr
                    Case 1
                        rList = rUPSList
                        CARRIER_CODE = "UPS"
                        ASCMAIN1.Progress("-", "UPS")

                    Case 2
                        rList = rFEDEXList
                        CARRIER_CODE = "FEDEX"
                        ASCMAIN1.Progress("-", "FedEx")

                    Case 3
                        rList = rUPSFreightList
                        CARRIER_CODE = "UPS"
                        ASCMAIN1.Progress("-", "UPS Freight")
                        freightShipment = " and ISNULL(FREIGHT_SHIPMENT, '0') = '1'"

                    Case 4
                        rList = rUSPSList
                        CARRIER_CODE = "USPS"
                        ASCMAIN1.Progress("-", "USPS")

                End Select

                If rList IsNot Nothing Then
                    For iLoop As Integer = 0 To rList.Count - 1
                        With rList(iLoop)
                            If .ServiceType Is Nothing OrElse (.ServiceType = 0 AndAlso .ServiceTypeDescription.Length = 0) Then
                                Continue For
                            End If

                            ' Display only those services that are mapped to ship vias
                            If dst.Tables("SOTSVIA1").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND CARRIER_PROD_CODE = '" & .ServiceType & "' AND SHIP_VIA_STATUS = 'A'" & freightShipment).Length = 0 Then
                                Continue For
                            End If

                            'CARRIER_SURCHARGE_PERC
                            drSOTCARR1 = dst.Tables("SOTCARR1").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)

                            If drSOTCARR1.Table.Columns.Contains("CARRIER_SURCHARGE_PERC") Then
                                CARRIER_SURCHARGE_PERC = Val(drSOTCARR1.Item("CARRIER_SURCHARGE_PERC") & String.Empty)
                            End If

                            If drSOTCARR1.Table.Columns.Contains("FRT_PER_SALES_HOLD") Then
                                FRT_PER_SALES_HOLD = Val(drSOTCARR1.Item("FRT_PER_SALES_HOLD") & String.Empty)
                            End If

                            If drSOTCARR1.Table.Columns.Contains("CARRIER_PPA_TYPE") Then
                                CARRIER_PPA_TYPE = drSOTCARR1.Item("CARRIER_PPA_TYPE") & String.Empty
                                ' If not set then set to List
                                If CARRIER_PPA_TYPE.Length = 0 Then
                                    CARRIER_PPA_TYPE = "L"
                                End If
                            End If

                            If drSOTCARR1.Table.Columns.Contains("CARRIER_SURCHARGE_BASE") Then
                                CARRIER_SURCHARGE_BASE = drSOTCARR1.Item("CARRIER_SURCHARGE_BASE") & String.Empty
                                ' If not set then set to List
                                If CARRIER_SURCHARGE_BASE.Length = 0 Then
                                    CARRIER_SURCHARGE_BASE = "L"
                                End If
                            End If

                            drWHTSHPC4 = dst.Tables("WHTSHPC4").NewRow
                            drWHTSHPC4.Item("SHIP_CNTL_NO") = "*"

                            Select Case CARRIER_CODE
                                Case "UPS"
                                    drWHTSHPC4.Item("SERVICE_INDEX") = iLoop + (100 * iCtr)
                                Case "FEDEX"
                                    drWHTSHPC4.Item("SERVICE_INDEX") = iLoop + 200
                                Case "USPS"
                                    drWHTSHPC4.Item("SERVICE_INDEX") = iLoop + 300
                            End Select

                            drWHTSHPC4.Item("SERVICE_TYPE_DESC") = .ServiceTypeDescription
                            drWHTSHPC4.Item("DISCLAIMER") = .Disclaimer

                            drWHTSHPC4.Item("CARRIER_CODE") = CARRIER_CODE '
                            drWHTSHPC4.Item("SHIP_VIA_CODE") = dst.Tables("SOTSVIA1").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND CARRIER_PROD_CODE = '" & .ServiceType & "' AND SHIP_VIA_STATUS = 'A'" & freightShipment)(0).Item("SHIP_VIA_CODE")

                            If (.AccountNetCharge & String.Empty <> "") Then
                                drWHTSHPC4.Item("ACCT_NET_CHARGE") = Convert.ToDecimal(.AccountNetCharge)
                            Else
                                drWHTSHPC4.Item("ACCT_NET_CHARGE") = Convert.ToDecimal(.ListNetCharge)
                            End If

                            drWHTSHPC4.Item("SERVICE_TYPE") = .ServiceType
                            drWHTSHPC4.Item("SURCHARGE") = 0
                            drWHTSHPC4.Item("DELIVERY_TIME") = .DeliveryTime
                            drWHTSHPC4.Item("LIST_NET_CHARGE") = .ListNetCharge
                            If .TransitTime <> "" Then
                                drWHTSHPC4.Item("TRANSIT_TIME") = .TransitTime
                            End If

                            drWHTSHPC4.Item("CARRIER_CODE") = CARRIER_CODE

                            Select Case CARRIER_PPA_TYPE
                                Case "F" ' None
                                    drWHTSHPC4.Item("CUSTOMER_BASE_CHARGE") = 0
                                Case "N" ' Negotiated
                                    drWHTSHPC4.Item("CUSTOMER_BASE_CHARGE") = drWHTSHPC4.Item("ACCT_NET_CHARGE")
                                Case "L" ' List
                                    drWHTSHPC4.Item("CUSTOMER_BASE_CHARGE") = drWHTSHPC4.Item("LIST_NET_CHARGE")
                                Case Else
                                    ' If not set then use  List
                                    drWHTSHPC4.Item("CUSTOMER_BASE_CHARGE") = drWHTSHPC4.Item("LIST_NET_CHARGE")
                            End Select

                            ' Additional Surcharge based off List
                            If CARRIER_SURCHARGE_PERC > 0 Then
                                Select Case CARRIER_SURCHARGE_BASE
                                    Case "N" ' Negotiated
                                        drWHTSHPC4.Item("SURCHARGE") = Val(drWHTSHPC4.Item("ACCT_NET_CHARGE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                                    Case "L" ' List
                                        drWHTSHPC4.Item("SURCHARGE") = Val(drWHTSHPC4.Item("LIST_NET_CHARGE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                                    Case Else
                                        ' If not set then use  List
                                        drWHTSHPC4.Item("SURCHARGE") = Val(drWHTSHPC4.Item("LIST_NET_CHARGE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                                End Select
                            End If

                            dst.Tables("WHTSHPC4").Rows.Add(drWHTSHPC4)

                        End With
                    Next
                End If
            Next

            grdWHTSHPC4.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
            grdWHTSHPC4.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

            Try
                Sort_grdColumns(grdWHTSHPC4, "TOTAL_CHARGE")
            Catch ex As Exception
            End Try

            With grdWHTSHPC4.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With

        Catch ex As Exception
            MessageBox.Show("Get Rates Error: " & ex.Message, "Get Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

    End Sub

    Private Function GetUpsRates() As WHCSHIP1.RateList()
        Try

            Dim rList(1) As WHCSHIP1.RateList

            If dst.Tables("SOTCARR1").Select("CARRIER_CODE = 'UPS'").Length = 0 Then
                Return Nothing
            End If

            If dst.Tables("SOTCARR3").Select("CARRIER_CODE = 'UPS'").Length = 0 Then
                Return Nothing
            End If

            Dim CARRIER_CODE As String = "UPS"
            Dim drSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Select("CARRIER_CODE = 'UPS'")(0)
            Dim drSOTCARR3 As DataRow = Nothing
            Dim CUST_CODE As String = dst.Tables("SOTORDR1").Rows(0).Item("CUST_CODE") & String.Empty
            Dim ORDR_NO As String = dst.Tables("SOTORDR1").Rows(0).Item("ORDR_NO") & String.Empty

            If dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "'").Length > 0 Then
                drSOTCARR3 = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "'")(0)
            ElseIf dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = DIVISION_CODE AND DIVISION_CODE = '" & ASCMAIN1.CLIENT & "'").Length > 0 Then
                drSOTCARR3 = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = DIVISION_CODE AND DIVISION_CODE = '" & ASCMAIN1.CLIENT & "'")(0)
            Else
                Return Nothing
            End If

            Dim drICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(txtWHSE_CODE.Text)
            If drICTWHSE1 Is Nothing Then
                Return Nothing
            End If

            clsShip.Reset()
            clsShip.Service = WHCSHIP1.ServiceProviders.UPS

            ' Credentials
            With clsShip
                .Server = drSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
                .UserId = drSOTCARR3.Item("SHIPPER_ID") & String.Empty
                .Password = drSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
                .AccountNumber = drSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
                .UPSAccessKey = drSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                .FedexMeterNumber = drSOTCARR3.Item("METER_NUMBER") & String.Empty
                .FedexDeveloperKey = drSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                .LabelStockType = (drSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim
            End With

            clsShip.RequestedServiceType = ServiceTypes.stUnspecified
            clsShip.UPSPickupType = UpsratesPickupTypes.ptDailyPickup
            clsShip.CustomerType = UpsratesCustomerTypes.ccRetail
            clsShip.ShipDate = DateTime.Now.ToShortDateString

            Dim listSeqNo As New List(Of Int16)

            For Each drSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SELECTED = '1'")
                Dim PICK_NO As String = drSOTPICK1.Item("PICK_NO")
                For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'", "CART_SEQ, CART_TOTAL_WGT_ACTUAL DESC", DataViewRowState.CurrentRows)

                    If listSeqNo.Contains(drSOTCART1.Item("CART_SEQ")) Then
                        Continue For
                    End If

                    listSeqNo.Add(drSOTCART1.Item("CART_SEQ"))

                    Dim pkgDetail As New PackageDetail

                    pkgDetail.Id = StrReverse(StrReverse(drSOTCART1.Item("CART_NO") & String.Empty).Substring(0, 8))
                    pkgDetail.Weight = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "CART_SEQ = " & drSOTCART1.Item("CART_SEQ")) & String.Empty)

                    ' Convert Pounds to Ounces
                    pkgDetail.Weight *= 16

                    If pkgDetail.Weight = "0" Then
                        pkgDetail.Weight = "16.0"
                    End If

                    pkgDetail.PackagingType = CType(Val(drSOTCART1.Item("PACKAGING_TYPE") & String.Empty), UpsratesPickupTypes)
                    pkgDetail.Length = Val(drSOTCART1.Item("PKG_L") & String.Empty)
                    pkgDetail.Width = Val(drSOTCART1.Item("PKG_W") & String.Empty)
                    pkgDetail.Height = Val(drSOTCART1.Item("PKG_H") & String.Empty)

                    pkgDetail.InsuredValue = 0 ' numInsureValue.Value * -1 / dst.Tables("SOTCART1").Rows.Count
                    clsShip.PackageDetailList.Add(pkgDetail)
                Next
            Next

            With clsShip.Sender
                .Company = (drICTWHSE1.Item("WHSE_DESC") & String.Empty).ToString.Trim
                .FirstName = (drICTWHSE1.Item("WHSE_CONTACT") & String.Empty).ToString.Trim
                .MiddleInitial = String.Empty
                .LastName = String.Empty
                .Address1 = (drICTWHSE1.Item("WHSE_ADDR1") & String.Empty).ToString.Trim
                .Address2 = (drICTWHSE1.Item("WHSE_ADDR2") & String.Empty).ToString.Trim
                .City = (drICTWHSE1.Item("WHSE_CITY") & String.Empty).ToString.Trim
                .State = (drICTWHSE1.Item("WHSE_STATE") & String.Empty).ToString.Trim
                .ZipCode = (drICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                .CountryCode = (drICTWHSE1.Item("WHSE_COUNTRY") & String.Empty).ToString.Trim.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                If .CountryCode = "CAN" Then .CountryCode = "CA"
                .Phone = (drICTWHSE1.Item("WHSE_PHONE") & String.Empty).ToString.Trim
            End With

            Dim drSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Rows.Find({ORDR_NO, "ST"})

            With clsShip.Recipient
                If drSOTORDR5.Item("CUST_CONTACT") & String.Empty <> String.Empty Then
                    .FirstName = drSOTORDR5.Item("CUST_CONTACT") & String.Empty
                Else
                    .FirstName = drSOTORDR5.Item("CUST_NAME") & String.Empty
                End If

                .MiddleInitial = String.Empty
                .LastName = String.Empty

                .Address1 = drSOTORDR5.Item("CUST_ADDR1") & String.Empty
                .Address2 = drSOTORDR5.Item("CUST_ADDR2") & String.Empty
                .City = drSOTORDR5.Item("CUST_CITY") & String.Empty
                .State = drSOTORDR5.Item("CUST_STATE") & String.Empty
                .ZipCode = drSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty
                .CountryCode = (drSOTORDR5.Item("CUST_COUNTRY") & String.Empty).ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                If .CountryCode = "CAN" Then .CountryCode = "CA"

                .Company = .FirstName
                .Phone = drSOTORDR5.Item("CUST_PHONE") & String.Empty

                If .Phone.Trim.Length = 0 Then
                    .Phone = clsShip.Sender.Phone
                End If

                If .Phone.Trim.Length = 0 Then
                    .Phone = "1234567890"
                End If

                .IsResidental = True
                .IsPOBox = False
            End With

            clsShip.RatesTotalValue = 0
            clsShip.ShipmentSpecialServices = 0
            clsShip.SignatureRequired = False
            rList = clsShip.GetUPSRatesList()

            If clsShip.LastError.Length > 0 Then
                MessageBox.Show("UPS Error: " & clsShip.LastError, "UPS Rates Error")
            End If

            If rList Is Nothing Then
                ReDim rList(1)
            End If

            Return rList

        Catch ex As Exception
            MessageBox.Show("The following error occurred getting UPS Rates: " & ex.Message, "Get UPS Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try

    End Function

    Private Function GetFedExRates() As WHCSHIP1.RateList()
        Try

            Dim rList(1) As WHCSHIP1.RateList

            If dst.Tables("SOTCARR1").Select("CARRIER_CODE = 'FEDEX'").Length = 0 Then
                Return Nothing
            End If

            If dst.Tables("SOTCARR3").Select("CARRIER_CODE = 'FEDEX'").Length = 0 Then
                Return Nothing
            End If

            Dim CARRIER_CODE As String = "FEDEX"
            Dim drSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Select("CARRIER_CODE = 'FEDEX'")(0)
            Dim drSOTCARR3 As DataRow = Nothing
            Dim CUST_CODE As String = dst.Tables("SOTORDR1").Rows(0).Item("CUST_CODE") & String.Empty
            Dim ORDR_NO As String = dst.Tables("SOTORDR1").Rows(0).Item("ORDR_NO") & String.Empty

            If dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "' AND CARRIER_PROD_CODE IS NULL").Length > 0 Then
                drSOTCARR3 = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "' AND CARRIER_PROD_CODE IS NULL")(0)
            ElseIf dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = DIVISION_CODE AND DIVISION_CODE = '" & ASCMAIN1.CLIENT & "'").Length > 0 Then
                drSOTCARR3 = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = DIVISION_CODE AND DIVISION_CODE = '" & ASCMAIN1.CLIENT & "'")(0)
            Else
                Return Nothing
            End If

            Dim drICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(txtWHSE_CODE.Text)
            If drICTWHSE1 Is Nothing Then
                Return Nothing
            End If

            clsShip.Reset()
            clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpress

            ' Credentials
            With clsShip
                .Server = drSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
                .UserId = drSOTCARR3.Item("SHIPPER_ID") & String.Empty
                .Password = drSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
                .AccountNumber = drSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
                .UPSAccessKey = drSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                .FedexMeterNumber = drSOTCARR3.Item("METER_NUMBER") & String.Empty
                .FedexDeveloperKey = drSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                .LabelStockType = (drSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim
            End With

            clsShip.RequestedServiceType = ServiceTypes.stUnspecified
            clsShip.UPSPickupType = UpsratesPickupTypes.ptDailyPickup
            clsShip.CustomerType = UpsratesCustomerTypes.ccRetail
            clsShip.ShipDate = DateTime.Now.ToShortDateString

            Dim listSeqNo As New List(Of Int16)

            For Each drSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SELECTED = '1'")
                Dim PICK_NO As String = drSOTPICK1.Item("PICK_NO")
                For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'", "CART_SEQ, CART_TOTAL_WGT_ACTUAL DESC", DataViewRowState.CurrentRows)

                    If listSeqNo.Contains(drSOTCART1.Item("CART_SEQ")) Then
                        Continue For
                    End If

                    listSeqNo.Add(drSOTCART1.Item("CART_SEQ"))

                    Dim pkgDetail As New PackageDetail

                    pkgDetail.Id = drSOTCART1.Item("CART_NO").ToString.Substring(2)
                    pkgDetail.Weight = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "CART_SEQ = " & drSOTCART1.Item("CART_SEQ")) & String.Empty)
                    ' Convert to Ounces
                    pkgDetail.Weight *= 16

                    If pkgDetail.Weight = "0" Then
                        pkgDetail.Weight = "16.0"
                    End If

                    pkgDetail.PackagingType = CType(Val(drSOTCART1.Item("PACKAGING_TYPE") & String.Empty), UpsratesPickupTypes)
                    pkgDetail.Length = Val(drSOTCART1.Item("PKG_L") & String.Empty)
                    pkgDetail.Width = Val(drSOTCART1.Item("PKG_W") & String.Empty)
                    pkgDetail.Height = Val(drSOTCART1.Item("PKG_H") & String.Empty)

                    ' Can have either Insured or Declared not Both
                    pkgDetail.InsuredValue = 0
                    clsShip.PackageDetailList.Add(pkgDetail)
                Next
            Next

            With clsShip.Sender
                .Company = (drICTWHSE1.Item("WHSE_DESC") & String.Empty).ToString.Trim
                .FirstName = (drICTWHSE1.Item("WHSE_CONTACT") & String.Empty).ToString.Trim
                .MiddleInitial = String.Empty
                .LastName = String.Empty
                .Address1 = (drICTWHSE1.Item("WHSE_ADDR1") & String.Empty).ToString.Trim
                .Address2 = (drICTWHSE1.Item("WHSE_ADDR2") & String.Empty).ToString.Trim
                .City = (drICTWHSE1.Item("WHSE_CITY") & String.Empty).ToString.Trim
                .State = (drICTWHSE1.Item("WHSE_STATE") & String.Empty).ToString.Trim
                .ZipCode = (drICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                .CountryCode = (drICTWHSE1.Item("WHSE_COUNTRY") & String.Empty).ToString.Trim.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                .Phone = (drICTWHSE1.Item("WHSE_PHONE") & String.Empty).ToString.Trim
            End With

            Dim drSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Rows.Find({ORDR_NO, "ST"})

            With clsShip.Recipient
                If drSOTORDR5.Item("CUST_CONTACT") & String.Empty <> String.Empty Then
                    .FirstName = drSOTORDR5.Item("CUST_CONTACT") & String.Empty
                Else
                    .FirstName = drSOTORDR5.Item("CUST_NAME") & String.Empty
                End If

                .MiddleInitial = String.Empty
                .LastName = String.Empty

                .Address1 = drSOTORDR5.Item("CUST_ADDR1") & String.Empty
                .Address2 = drSOTORDR5.Item("CUST_ADDR2") & String.Empty
                .City = drSOTORDR5.Item("CUST_CITY") & String.Empty
                .State = drSOTORDR5.Item("CUST_STATE") & String.Empty
                .ZipCode = drSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty
                .CountryCode = (drSOTORDR5.Item("CUST_COUNTRY") & String.Empty).ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                If .CountryCode = "CAN" Then .CountryCode = "CA"

                .Company = .FirstName
                .Phone = drSOTORDR5.Item("CUST_PHONE") & String.Empty

                If .Phone.Trim.Length = 0 Then
                    .Phone = clsShip.Sender.Phone
                End If

                If .Phone.Trim.Length = 0 Then
                    .Phone = "1234567890"
                End If

                .IsResidental = True
                .IsPOBox = False
            End With

            clsShip.ShipmentSpecialServices = 0
            clsShip.SignatureRequired = False
            rList = clsShip.GetFedExRatesList()

            If clsShip.LastError.Length > 0 Then
                MessageBox.Show("FedEx Error: " & clsShip.LastError, "FedEx Rates Error")
            End If

            If rList Is Nothing Then
                ReDim rList(1)
            End If

            Return rList

        Catch ex As Exception
            MessageBox.Show("The following error occurred getting FedEx Rates: " & ex.Message, "Get FedEx Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try

    End Function

    Private Function RequestShippingLabel(ByVal INV_NO As String,
                                          ByRef ErrorMessage As String,
                                          ByVal PreScreenForErrorsOnly As Boolean,
                                          ByVal SHIP_BOL_NO As String) As Boolean

        ErrorMessage = String.Empty
        Dim ShippingLabels As New List(Of String)
        Dim drSOTSHIP1 As DataRow = Nothing
        Dim drSOTPICK1 As DataRow = Nothing
        Dim drSOTORDR1 As DataRow = Nothing
        Dim drSOTSVIA1 As DataRow = Nothing
        Dim drSOTCARR1 As DataRow = Nothing
        Dim drSOTORDR5 As DataRow = Nothing

        Dim SHIP_VIA_CODE As String = txtSHIP_VIA_CODE.Text
        Dim CARRIER_CODE As String = String.Empty
        Dim CARRIER_PROD_CODE As String = String.Empty

        Dim ORDR_NO As String = String.Empty
        Dim CUST_CODE As String = String.Empty
        Dim PICK_NO As String = String.Empty
        Dim ORDR_NO_WEB As String = String.Empty
        Dim ORDR_CUST_PO As String = String.Empty

        Dim SHIP_PACKAGE_NO As Int64 = 0
        Dim pkgId As Int64 = 0
        Dim isPitneyBowes As Boolean = False

        Dim CARRIER_SURCHARGE_PERC As Int16 = 0
        Dim CARRIER_SURCHARGE_BASE As String = "L"

        Dim FRT_PER_SALES_HOLD As Int16 = 0
        Dim CARRIER_PPA_TYPE As String = "L"

        RequestShippingLabel = True

        Try
            drSOTSHIP1 = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
            drSOTSVIA1 = dst.Tables("SOTSVIA1").Rows.Find(SHIP_VIA_CODE)
            CARRIER_CODE = drSOTSVIA1.Item("CARRIER_CODE") & String.Empty
            CARRIER_PROD_CODE = drSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty

            drSOTCARR1 = dst.Tables("SOTCARR1").Rows.Find(CARRIER_CODE)
            drSOTPICK1 = dst.Tables("SOTPICK1").Select($"SHIP_BOL_NO = '{SHIP_BOL_NO}'")(0)
            ORDR_NO = drSOTPICK1.Item("ORDR_NO") & String.Empty
            PICK_NO = drSOTPICK1.Item("PICK_NO") & String.Empty

            drSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
            CUST_CODE = drSOTORDR1.Item("CUST_CODE") & String.Empty
            ORDR_NO_WEB = drSOTORDR1.Item("ORDR_NO_WEB") & String.Empty
            ORDR_CUST_PO = drSOTORDR1.Item("ORDR_CUST_PO") & String.Empty
            Fill_Records("ARTCUSTS", CUST_CODE)
            Fill_Records("ARTCUST1", CUST_CODE)
            Dim drARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)

            drSOTORDR5 = dst.Tables("SOTORDR5").Rows.Find({ORDR_NO, "ST"})

            ' Logic added 3/18/2017 for Regency
            If drSOTCARR1.Table.Columns.Contains("CARRIER_SURCHARGE_PERC") Then
                CARRIER_SURCHARGE_PERC = Val(drSOTCARR1.Item("CARRIER_SURCHARGE_PERC") & String.Empty)
            End If

            If drSOTCARR1.Table.Columns.Contains("CARRIER_SURCHARGE_BASE") Then
                CARRIER_SURCHARGE_BASE = drSOTCARR1.Item("CARRIER_SURCHARGE_BASE") & String.Empty
                ' If not set then set to List
                If CARRIER_SURCHARGE_BASE.Length = 0 Then
                    CARRIER_SURCHARGE_BASE = "L"
                End If
            End If

            If drSOTCARR1.Table.Columns.Contains("FRT_PER_SALES_HOLD") Then
                FRT_PER_SALES_HOLD = Val(drSOTCARR1.Item("FRT_PER_SALES_HOLD") & String.Empty)
            End If

            If drSOTCARR1.Table.Columns.Contains("CARRIER_PPA_TYPE") Then
                CARRIER_PPA_TYPE = drSOTCARR1.Item("CARRIER_PPA_TYPE") & String.Empty
                ' If not set then set tp list
                If CARRIER_PPA_TYPE.Length = 0 Then
                    CARRIER_PPA_TYPE = "L"
                End If
            End If

            ' Load and Validate Carrier/Ship Method
            Dim drSOTCARR2 As DataRow = LookUp("SOTCARR2", New String() {CARRIER_CODE, CARRIER_PROD_CODE})
            If drSOTCARR2 Is Nothing Then
                ErrorMessage = "Invalid or missing Carrier / Ship Method combination for shipping label request"
                Return False
            End If

            ' Credentials
            Dim drSOTCARR3 As DataRow = Nothing

            ' SHIPPER_DIVISION_CODE holds a customer code,  SHIPPER_ID
            If dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "' AND CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "'").Length > 0 Then
                drSOTCARR3 = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "' AND CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "'")(0)
            ElseIf dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "' AND CARRIER_PROD_CODE IS NULL").Length > 0 Then
                drSOTCARR3 = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = '" & CUST_CODE & "' AND CARRIER_PROD_CODE IS NULL")(0)
            ElseIf dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = DIVISION_CODE AND DIVISION_CODE = '" & ASCMAIN1.CLIENT & "'").Length > 0 Then
                drSOTCARR3 = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = DIVISION_CODE AND DIVISION_CODE = '" & ASCMAIN1.CLIENT & "'")(0)
            End If

            If drSOTCARR3 Is Nothing Then
                ErrorMessage = "Cannot determine the Carrier Account to use for the shipping label request"
                Return False
            End If

            Dim DIVISION_CODE As String = drSOTCARR3.Item("DIVISION_CODE") & String.Empty
            Dim CARRIER_ACCOUNT_NO As String = drSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty

            ' See if we need to use a different From Address.
            Dim drSOTCARR5 As DataRow = Nothing
            ASCMAIN1.sql = "CARRIER_CODE = '" & CARRIER_CODE & "' and DIVISION_CODE = '" & DIVISION_CODE & "' and CARRIER_ACCOUNT_NO = '" & CARRIER_ACCOUNT_NO & "' and CUST_CODE = '" & CUST_CODE & "'"
            If dst.Tables("SOTCARR5").Select(ASCMAIN1.sql).Length > 0 Then
                drSOTCARR5 = dst.Tables("SOTCARR5").Select(ASCMAIN1.sql)(0)
            Else
                ASCMAIN1.sql = "CARRIER_CODE = '" & CARRIER_CODE & "' and DIVISION_CODE = '" & DIVISION_CODE & "' and CARRIER_ACCOUNT_NO = '" & CARRIER_ACCOUNT_NO & "' and CUST_CODE = '" & "*" & "'"
                If dst.Tables("SOTCARR5").Select(ASCMAIN1.sql).Length > 0 Then
                    drSOTCARR5 = dst.Tables("SOTCARR5").Select(ASCMAIN1.sql)(0)
                End If
            End If

            Dim ShippingLabelDirectory As String = (drSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim
            Dim PROVIDER_TYPE As String = (drSOTCARR1.Item("PROVIDER_TYPE") & String.Empty).ToString.Trim

            If drSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty = String.Empty Then
                ErrorMessage = "Invalid or missing Carrier Account Number for shipping label request"
                Return False
            End If

            Try
                If ShippingLabelDirectory.Length > 0 Then
                    If Not My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                        My.Computer.FileSystem.CreateDirectory(ShippingLabelDirectory)
                    End If
                End If
            Catch ex As Exception
                ShippingLabelDirectory = String.Empty
            End Try

            If ShippingLabelDirectory.Length > 0 AndAlso Not ShippingLabelDirectory.EndsWith("\") Then
                ShippingLabelDirectory = ShippingLabelDirectory & "\"
            End If

            Dim drICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(Absx1.txtFor("WHSE_CODE").Text)
            If drICTWHSE1 Is Nothing Then
                ErrorMessage = "Invalid or missing Warehouse"
                Return False
            End If

            Dim CUST_NAME As String = (drSOTORDR5.Item("CUST_NAME") & String.Empty).ToString.Trim
            Dim CUST_CONTACT As String = (drSOTORDR5.Item("CUST_CONTACT") & String.Empty).ToString.Trim
            Dim CUST_ADDR1 As String = (drSOTORDR5.Item("CUST_ADDR1") & String.Empty).ToString.Trim
            Dim CUST_ADDR2 As String = (drSOTORDR5.Item("CUST_ADDR2") & String.Empty).ToString.Trim
            Dim CUST_CITY As String = (drSOTORDR5.Item("CUST_CITY") & String.Empty).ToString.Trim
            Dim CUST_STATE As String = (drSOTORDR5.Item("CUST_STATE") & String.Empty).ToString.Trim
            Dim CUST_COUNTRY As String = (drSOTORDR5.Item("CUST_COUNTRY") & String.Empty).ToString.Trim
            Dim CUST_ZIP_CODE As String = (drSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty).ToString.Trim
            Dim CUST_PHONE As String = (drSOTORDR5.Item("CUST_PHONE") & String.Empty).ToString.Trim

            If CUST_ADDR1.Length = 0 AndAlso CUST_ADDR2.Length = 0 Then
                ErrorMessage = "Invalid or missing Ship To Street Address"
                Return False
            ElseIf Not CUST_COUNTRY.StartsWith("US") AndAlso (CUST_CITY.Length = 0 OrElse CUST_ZIP_CODE.Length = 0) Then
                ErrorMessage = "Invalid or missing International Ship To City and/or Zip Code"
                Return False
            ElseIf CUST_CITY.Length = 0 OrElse CUST_STATE.Length = 0 OrElse CUST_ZIP_CODE.Length = 0 Then
                ErrorMessage = "Invalid or missing Ship To City, State or Zip Code"
                Return False
            ElseIf CUST_COUNTRY.Length = 0 Then
                Dim drTATSTATE As DataRow = tblTATSTATE.Rows.Find(CUST_STATE)
                If drTATSTATE IsNot Nothing Then
                    CUST_COUNTRY = "US"
                Else
                    ErrorMessage = "Invalid or missing Country Code"
                    Return False
                End If
            End If

            ' 02/25/2020 - Evaluate Cartons to make sure carton dimensions are sent to UPS/ FedEx
            For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select($"PICK_NO = '{PICK_NO}'")
                Dim CART_NO As String = drSOTCART1.Item("CART_NO") & String.Empty
                Dim PACKAGING_TYPE As String = drSOTCART1.Item("PACKAGING_TYPE") & String.Empty
                Dim PKG_CODE As String = drSOTCART1.Item("PKG_CODE") & String.Empty

                ' Make sure FedEx and UPS shipments have box dimensions.
                Dim LENGTH As Decimal = Val(drSOTCART1.Item("PKG_L") & String.Empty)
                Dim WIDTH As Decimal = Val(drSOTCART1.Item("PKG_W") & String.Empty)
                Dim HEIGHT As Decimal = Val(drSOTCART1.Item("PKG_H") & String.Empty)

                If LENGTH <= 0 OrElse WIDTH <= 0 OrElse HEIGHT <= 0 Then
                    ErrorMessage &= vbCr & "Carton " & CART_NO & " has invalid dimensions."
                End If

                If dst.Tables("SOTCART2").Select("CART_NO = '" & CART_NO & "'").Length = 0 Then
                    ErrorMessage &= vbCr & "Carton " & CART_NO & " does not have any assigned products."
                End If

                If PACKAGING_TYPE = defaultPACKAGING_TYPE Then
                    If dst.Tables("WHTPKGM1").Rows.Find(PKG_CODE) Is Nothing Then
                        ErrorMessage &= vbCr & "Carton " & CART_NO & " has an invalid package code."
                    End If
                End If
            Next

            If ErrorMessage.Length > 0 Then
                Return True
            End If

            If PreScreenForErrorsOnly Then Return True

            '*******************************************************************************

            Dim isInternationalShipment As Boolean = False
            Dim fedexSmartPost As Int16 = 26

            Dim FRT_TERMS As String = Absx1.txtFor("FRT_TERMS").Text
            Dim PPA_FREIGHT As Decimal = 0
            Dim OUR_FREIGHT As Decimal = 0

            dst.Tables("WHTSHPC1").Rows.Clear()
            dst.Tables("WHTSHPC2").Rows.Clear()
            dst.Tables("WHTSHPC5").Rows.Clear()
            dst.Tables("WHTSHPCG").Rows.Clear()
            dst.Tables("WHTSHPCS").Rows.Clear()
            dst.Tables("WHTSHPCC").Rows.Clear()
            dst.Tables("WHTSHPCP").Rows.Clear()

            Dim SHIP_CNTL_NO As String = String.Empty
            clsShip.Reset()

            ' Credentials
            clsShip.Server = drSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
            clsShip.UserId = drSOTCARR3.Item("SHIPPER_ID") & String.Empty
            clsShip.Password = drSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
            clsShip.AccountNumber = drSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
            clsShip.UPSAccessKey = drSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.FedexMeterNumber = drSOTCARR3.Item("METER_NUMBER") & String.Empty
            clsShip.FedexDeveloperKey = drSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.LabelStockType = (drSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

            Dim drWHTSHPC1 As DataRow = Nothing
            Dim drWHTSHPC2 As DataRow = Nothing
            Dim drWHTSHPC5 As DataRow = Nothing
            Dim drWHTSHPCG As DataRow = Nothing

            drWHTSHPC1 = dst.Tables("WHTSHPC1").NewRow
            SHIP_CNTL_NO = ASCMAIN1.Next_Control_No("WHTSHPC1.SHIP_CNTL_NO")
            drWHTSHPC1.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
            drWHTSHPC1.Item("CARRIER_CODE") = CARRIER_CODE
            drWHTSHPC1.Item("CARRIER_PROD_CODE") = CARRIER_PROD_CODE
            drWHTSHPC1.Item("CARRIER_ACCOUNT_NO") = drSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
            dst.Tables("WHTSHPC1").Rows.Add(drWHTSHPC1)

            drWHTSHPC1.Item("STATUS") = "I"
            drWHTSHPC1.Item("ERROR_MSG") = String.Empty
            drWHTSHPC1.Item("SHIP_DATE") = DateTime.Now.ToString("MM/dd/yyyy")
            drWHTSHPC1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            drWHTSHPC1.Item("OPS_YYYYWW") = ASCMAIN1.CYW
            drWHTSHPC1.Item("CUST_CODE") = CUST_CODE
            drWHTSHPC1.Item("INIT_DATE") = DateTime.Now
            drWHTSHPC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            drWHTSHPC1.Item("LAST_DATE") = DateTime.Now
            drWHTSHPC1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            drWHTSHPC1.Item("MASTER_TRACKING_NO") = String.Empty
            drWHTSHPC1.Item("CUSTOMS_VALUE") = 0
            drWHTSHPC1.Item("SHIP_BOL_NO") = SHIP_BOL_NO
            drWHTSHPC1.Item("SHIP_VIA_CODE") = SHIP_VIA_CODE

            drWHTSHPC1.Item("INSURED_VALUE") = 0
            drWHTSHPC1.Item("INSURED_SHIPMENT") = IIf(Absx1.chkFor("INSURED_SHIPMENT").Checked, "1", "0")

            ' Update the Key in these tables
            For Each tableName As String In New String() {"WHTSHPC4", "WHTSHPCA"}
                For Each dr As DataRow In dst.Tables(tableName).Select("")
                    dr.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                Next
            Next

            ' Sender Information
            With clsShip.Sender
                .Company = (drICTWHSE1.Item("WHSE_DESC") & String.Empty).ToString.Trim
                .Phone = (drICTWHSE1.Item("WHSE_PHONE") & String.Empty).ToString.Trim
                .FirstName = (drICTWHSE1.Item("WHSE_CONTACT") & String.Empty).ToString.Trim
                .MiddleInitial = String.Empty
                .LastName = String.Empty
                .Address1 = (drICTWHSE1.Item("WHSE_ADDR1") & String.Empty).ToString.Trim
                .Address2 = (drICTWHSE1.Item("WHSE_ADDR2") & String.Empty).ToString.Trim
                .City = (drICTWHSE1.Item("WHSE_CITY") & String.Empty).ToString.Trim
                .State = (drICTWHSE1.Item("WHSE_STATE") & String.Empty).ToString.Trim
                .ZipCode = (drICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                .CountryCode = (drICTWHSE1.Item("WHSE_COUNTRY") & String.Empty).ToString.Trim.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                If .CountryCode = "CAN" Then .CountryCode = "CA"

                drWHTSHPC5 = dst.Tables("WHTSHPC5").NewRow
                drWHTSHPC5.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                drWHTSHPC5.Item("SHIP_ADDR_TYPE") = "SF"
                drWHTSHPC5.Item("SHIP_FIRST_NAME") = .FirstName
                drWHTSHPC5.Item("SHIP_LAST_NAME") = .LastName
                drWHTSHPC5.Item("SHIP_MIDDLE_INIT") = .MiddleInitial
                drWHTSHPC5.Item("SHIP_PHONE") = .Phone
                drWHTSHPC5.Item("SHIP_FAX") = .Fax
                drWHTSHPC5.Item("SHIP_EMAIL") = .eMail
                drWHTSHPC5.Item("SHIP_COMPANY") = .Company
                drWHTSHPC5.Item("SHIP_ADDR1") = .Address1
                drWHTSHPC5.Item("SHIP_ADDR2") = .Address2
                drWHTSHPC5.Item("SHIP_CITY") = .City
                drWHTSHPC5.Item("SHIP_STATE") = .State
                drWHTSHPC5.Item("SHIP_ZIP_CODE") = .ZipCode
                drWHTSHPC5.Item("SHIP_COUNTRY_CODE") = .CountryCode
                drWHTSHPC5.Item("SHIP_RESIDENTIAL") = IIf(.IsResidental, "1", "0")
                drWHTSHPC5.Item("SHIP_PO_BOX") = IIf(.IsPOBox, "1", "0")
                dst.Tables("WHTSHPC5").Rows.Add(drWHTSHPC5)
            End With

            ' This is an override address that will print as the Ship From address in the upper left hand corner of the shipping label.
            If drSOTCARR5 IsNot Nothing Then
                With clsShip.Account
                    .Company = drSOTCARR5.Item("ACCOUNT_NAME") & String.Empty
                    .Phone = drSOTCARR5.Item("ACCOUNT_PHONE") & String.Empty

                    .FirstName = drSOTCARR5.Item("ACCOUNT_CONTACT") & String.Empty
                    .MiddleInitial = String.Empty
                    .LastName = String.Empty
                    .Address1 = drSOTCARR5.Item("ACCOUNT_ADDR1") & String.Empty
                    .Address2 = drSOTCARR5.Item("ACCOUNT_ADDR2") & String.Empty
                    .Address3 = drSOTCARR5.Item("ACCOUNT_ADDR3") & String.Empty
                    .City = drSOTCARR5.Item("ACCOUNT_CITY") & String.Empty
                    .State = drSOTCARR5.Item("ACCOUNT_STATE") & String.Empty
                    .ZipCode = drSOTCARR5.Item("ACCOUNT_ZIP_CODE") & String.Empty
                    .CountryCode = drSOTCARR5.Item("ACCOUNT_COUNTRY") & String.Empty
                    If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                    If .CountryCode = "USA" Then .CountryCode = "US"
                    If .CountryCode = "CAN" Then .CountryCode = "CA"

                    drWHTSHPC5 = dst.Tables("WHTSHPC5").NewRow
                    drWHTSHPC5.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    drWHTSHPC5.Item("SHIP_ADDR_TYPE") = "AC"
                    drWHTSHPC5.Item("SHIP_FIRST_NAME") = .FirstName
                    drWHTSHPC5.Item("SHIP_LAST_NAME") = .LastName
                    drWHTSHPC5.Item("SHIP_MIDDLE_INIT") = .MiddleInitial
                    drWHTSHPC5.Item("SHIP_PHONE") = .Phone
                    drWHTSHPC5.Item("SHIP_FAX") = .Fax
                    drWHTSHPC5.Item("SHIP_EMAIL") = .eMail
                    drWHTSHPC5.Item("SHIP_COMPANY") = .Company
                    drWHTSHPC5.Item("SHIP_ADDR1") = .Address1
                    drWHTSHPC5.Item("SHIP_ADDR2") = .Address2
                    drWHTSHPC5.Item("SHIP_CITY") = .City
                    drWHTSHPC5.Item("SHIP_STATE") = .State
                    drWHTSHPC5.Item("SHIP_ZIP_CODE") = .ZipCode
                    drWHTSHPC5.Item("SHIP_COUNTRY_CODE") = .CountryCode
                    drWHTSHPC5.Item("SHIP_RESIDENTIAL") = IIf(.IsResidental, "1", "0")
                    drWHTSHPC5.Item("SHIP_PO_BOX") = IIf(.IsPOBox, "1", "0")
                    dst.Tables("WHTSHPC5").Rows.Add(drWHTSHPC5)
                End With
            Else
                With clsShip.Account
                    .Company = String.Empty
                    .Phone = String.Empty

                    .FirstName = String.Empty
                    .MiddleInitial = String.Empty
                    .LastName = String.Empty
                    .Address1 = String.Empty
                    .Address2 = String.Empty
                    .Address3 = String.Empty
                    .City = String.Empty
                    .State = String.Empty
                    .ZipCode = String.Empty
                    .CountryCode = String.Empty
                End With
            End If

            ' Recipient
            With clsShip.Recipient
                .FirstName = IIf(CUST_CONTACT.Length > 0, CUST_CONTACT, CUST_NAME)
                .MiddleInitial = ""
                .LastName = ""

                .Address1 = CUST_ADDR1
                .Address2 = CUST_ADDR2
                .City = CUST_CITY
                .State = CUST_STATE
                .ZipCode = CUST_ZIP_CODE
                .CountryCode = CUST_COUNTRY.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
                If .CountryCode = "CAN" Then .CountryCode = "CA"

                .Company = CUST_NAME

                If .Company = .FirstName Then
                    .FirstName = String.Empty
                End If
                'End If

                .Phone = CUST_PHONE

                If .Phone.Trim.Length = 0 Then
                    .Phone = clsShip.Sender.Phone
                End If

                If .Phone.Trim.Length = 0 Then
                    .Phone = "1234567890"
                End If

                ' Force FedEx Ground Home Delivery to residental
                If drSOTSVIA1 IsNot Nothing Then
                    'CARRIER_CODE = 'FEDEX' AND CARRIER_PROD_CODE = '16'
                    If drSOTSVIA1.Item("CARRIER_CODE") & String.Empty = "FEDEX" Then
                        If drSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty = "16" Then
                        End If
                    End If
                End If

                .IsResidental = True
                .IsPOBox = False

                drWHTSHPC5 = dst.Tables("WHTSHPC5").NewRow
                drWHTSHPC5.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                drWHTSHPC5.Item("SHIP_ADDR_TYPE") = "ST"
                drWHTSHPC5.Item("SHIP_FIRST_NAME") = .FirstName
                drWHTSHPC5.Item("SHIP_LAST_NAME") = .LastName
                drWHTSHPC5.Item("SHIP_MIDDLE_INIT") = .MiddleInitial
                drWHTSHPC5.Item("SHIP_PHONE") = .Phone
                drWHTSHPC5.Item("SHIP_FAX") = .Fax
                drWHTSHPC5.Item("SHIP_EMAIL") = .eMail
                drWHTSHPC5.Item("SHIP_COMPANY") = .Company
                drWHTSHPC5.Item("SHIP_ADDR1") = .Address1
                drWHTSHPC5.Item("SHIP_ADDR2") = .Address2
                drWHTSHPC5.Item("SHIP_CITY") = .City
                drWHTSHPC5.Item("SHIP_STATE") = .State
                drWHTSHPC5.Item("SHIP_ZIP_CODE") = .ZipCode
                drWHTSHPC5.Item("SHIP_COUNTRY_CODE") = .CountryCode
                drWHTSHPC5.Item("SHIP_RESIDENTIAL") = IIf(.IsResidental, "1", "0")
                drWHTSHPC5.Item("SHIP_PO_BOX") = IIf(.IsPOBox, "1", "0")
                dst.Tables("WHTSHPC5").Rows.Add(drWHTSHPC5)
            End With

            If dst.Tables("ARTCUSTS").Rows.Count = 1 Then
                Dim drARTCUSTS As DataRow = dst.Tables("ARTCUSTS").Rows(0)
                With clsShip.ReturnAddress
                    Select Case drSOTCARR1.Item("PROVIDER_TYPE") & String.Empty
                        Case "F" ' Federal Express
                            If drARTCUSTS.Item("FDX_RTN_SHIP_COMPANY") & String.Empty <> String.Empty Then
                                .Address1 = drARTCUSTS.Item("FDX_RTN_SHIP_ADDR1") & String.Empty
                                .Address2 = drARTCUSTS.Item("FDX_RTN_SHIP_ADDR2") & String.Empty
                                .Address3 = String.Empty
                                .City = drARTCUSTS.Item("FDX_RTN_SHIP_CITY") & String.Empty
                                .Company = drARTCUSTS.Item("FDX_RTN_SHIP_COMPANY") & String.Empty
                                .CountryCode = drARTCUSTS.Item("FDX_RTN_SHIP_COUNTRY_CODE") & String.Empty
                                .eMail = String.Empty
                                .Fax = String.Empty
                                .FirstName = String.Empty
                                .IsPOBox = False
                                .IsResidental = False
                                .LastName = String.Empty
                                .MiddleInitial = String.Empty
                                .Phone = drARTCUSTS.Item("FDX_RTN_SHIP_PHONE") & String.Empty
                                ' This is required
                                If .Phone.Length = 0 Then
                                    .Phone = clsShip.Sender.Phone
                                End If
                                .State = drARTCUSTS.Item("FDX_RTN_SHIP_STATE") & String.Empty
                                .ZipCode = drARTCUSTS.Item("FDX_RTN_SHIP_ZIP_CODE") & String.Empty
                            End If

                        Case "U" ' UPS
                            If drARTCUSTS.Item("UPS_RTN_SHIP_COMPANY") & String.Empty <> String.Empty Then
                                .Address1 = drARTCUSTS.Item("UPS_RTN_SHIP_ADDR1") & String.Empty
                                .Address2 = drARTCUSTS.Item("UPS_RTN_SHIP_ADDR2") & String.Empty
                                .Address3 = String.Empty
                                .City = drARTCUSTS.Item("UPS_RTN_SHIP_CITY") & String.Empty
                                .Company = drARTCUSTS.Item("UPS_RTN_SHIP_COMPANY") & String.Empty
                                .CountryCode = drARTCUSTS.Item("UPS_RTN_SHIP_COUNTRY_CODE") & String.Empty
                                .eMail = String.Empty
                                .Fax = String.Empty
                                .FirstName = String.Empty
                                .IsPOBox = False
                                .IsResidental = False
                                .LastName = String.Empty
                                .MiddleInitial = String.Empty
                                .Phone = drARTCUSTS.Item("UPS_RTN_SHIP_PHONE") & String.Empty
                                ' This is required
                                If .Phone.Length = 0 Then
                                    .Phone = clsShip.Sender.Phone
                                End If
                                .State = drARTCUSTS.Item("UPS_RTN_SHIP_STATE") & String.Empty
                                .ZipCode = drARTCUSTS.Item("UPS_RTN_SHIP_ZIP_CODE") & String.Empty
                            End If
                    End Select
                End With
            End If

            ' US Puerto Rico is considered International
            isInternationalShipment = (clsShip.Recipient.CountryCode <> clsShip.Sender.CountryCode) OrElse (clsShip.Recipient.CountryCode = "US" AndAlso clsShip.Recipient.State = "PR")

            Select Case PROVIDER_TYPE
                Case WHCSHIP1.ProviderTypeFedex
                    If Not isInternationalShipment Then
                        clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpress
                    Else
                        clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpressInternational
                    End If

                Case WHCSHIP1.ProviderTypeUPS
                    If Not isInternationalShipment Then
                        clsShip.Service = WHCSHIP1.ServiceProviders.UPS
                    Else
                        clsShip.Service = WHCSHIP1.ServiceProviders.UPSInternational
                    End If

                Case WHCSHIP1.ProviderTypeUSPS
                    clsShip.Service = WHCSHIP1.ServiceProviders.USPS
                    Select Case drSOTCARR1.Item("USPS_PARTNER") & String.Empty
                        Case "1"
                            clsShip.USPSPostageProvider = WHCSHIP1.USPSPostageProviders.Endicia
                        Case "2"
                            clsShip.USPSPostageProvider = WHCSHIP1.USPSPostageProviders.StampsCom
                        Case "3"
                            clsShip.USPSPostageProvider = WHCSHIP1.USPSPostageProviders.PitneyBowes
                            clsShip.PitneyBowesUniqueTransactionID = "USPS_PB_" & ASCMAIN1.Next_Control_No("SOTCARR1.USPS")
                            clsShip.PitneyBowesInductionPostalCode = (drICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                            isPitneyBowes = True
                        Case Else
                            Return False
                    End Select

                Case WHCSHIP1.ProviderTypeCanada
                    clsShip.Service = WHCSHIP1.ServiceProviders.CanadaPost

                Case Else
                    Return False
            End Select

            ' Build a package for each Carton for the current Pick Ticket
            ' Change as of 1/21/2013
            ' Some shipments are multi Pick Tickets and some Pick Tickets are combined into 1 carton.
            ' The carton sequence will be used to group pick tickets into one carton and also
            ' be used to identify the sequence the Shipping label will get printed
            ' The user is not permitted to deselect a pick ticket; therefore, no londfer need to use dst.Tables("SOTPICK1").Select("SELECTED = '1'", "PICK_NO")
            clsShip.PackageDetailList.Clear()

            Dim cartSequenceNos As List(Of Int16) = New List(Of Int16)

            ' Commodities for international shipments
            clsShip.TotalCustomsValue = 0
            clsShip.CommodityDetailList.Clear()
            Dim COMMODITY_LNO As Int16 = 1
            Dim itemList As List(Of String) = New List(Of String)

            Refresh_Refs(CARRIER_CODE)

            For Each drSOTPICK1 In dst.Tables("SOTPICK1").Select("SELECTED = '1'", "PICK_NO")

                PICK_NO = drSOTPICK1.Item("PICK_NO") & String.Empty
                ORDR_NO = drSOTPICK1.Item("ORDR_NO") & String.Empty
                SHIP_BOL_NO = drSOTPICK1.Item("SHIP_BOL_NO") & String.Empty
                ORDR_CUST_PO = drSOTPICK1.Item("ORDR_CUST_PO") & String.Empty
                PPA_FREIGHT = 0
                OUR_FREIGHT = 0
                itemList.Clear()

                '' Get the Invoice Number now so we can put it on the label
                'Dim INV_NO As String = String.Empty

                'If ASCMAIN1.CLIENT = "VAN" Then
                '    INV_NO = ASCMAIN1.Next_Control_No("INV_NO_01")
                'Else
                '    INV_NO = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
                'End If

                drSOTPICK1.Item("INV_NO") = INV_NO
                drSOTPICK1.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO

                For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'", "CART_SEQ, CART_TOTAL_WGT_ACTUAL DESC")
                    ' This is done to place multi pick tickets into one carton
                    Dim CART_SEQ As Int32 = drSOTCART1.Item("CART_SEQ")
                    If cartSequenceNos.Contains(CART_SEQ) Then
                        Continue For
                    End If
                    cartSequenceNos.Add(CART_SEQ)

                    Dim PACKAGING_TYPE As String = drSOTCART1.Item("PACKAGING_TYPE") & String.Empty
                    Dim PKG_CODE As String = drSOTCART1.Item("PKG_CODE") & String.Empty
                    Dim drWHTPKGM1 As DataRow = dst.Tables("WHTPKGM1").Rows.Find(PKG_CODE) ' LookUp("WHTPKGM1", PKG_CODE)
                    pkgId = CART_SEQ ' (Val(StrReverse(StrReverse(rowSOTCART1.Item("CART_NO").ToString).Substring(0, 8))))

                    Dim shipPackageDetail As New PackageDetail
                    With shipPackageDetail
                        .PackagingType = Val(PACKAGING_TYPE)

                        ' This is done to place multi pick tickets into one carton. Need combined weight 
                        If ASCMAIN1.CLIENT = "RGI" Then
                            .Weight = Val(drSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") & String.Empty)
                        Else
                            .Weight = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "CART_SEQ = " & CART_SEQ) & String.Empty)
                        End If
                        If .Weight = 0 Then
                            .Weight = 1
                        End If

                        '*************************************
                        '        Convert to Ounces
                        '*************************************
                        .Weight = Convert.ToInt16(.Weight * 16)
                        ' Take what is in the grid
                        .Length = Val(drSOTCART1.Item("PKG_L") & String.Empty)
                        .Width = Val(drSOTCART1.Item("PKG_W") & String.Empty)
                        .Height = Val(drSOTCART1.Item("PKG_H") & String.Empty)

                        Dim reference As String = String.Empty
                        Dim refCount As Int16 = 0

                        Select Case PROVIDER_TYPE
                            Case WHCSHIP1.ProviderTypeFedex
                                ' Fedex allows up to 3 References

                                If ASCMAIN1.CLIENT = "VAN" Then
                                    If (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 0 Then
                                        reference &= "; " & (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    End If

                                    If (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 0 Then
                                        reference &= "; " & (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    End If

                                    If (drSOTCART1.Item("REFERENCE3") & String.Empty).ToString.Trim.Length > 0 Then
                                        reference &= "; " & (drSOTCART1.Item("REFERENCE3") & String.Empty).ToString
                                    End If

                                    refCount = 5
                                End If

                                If refCount < 3 _
                                    AndAlso (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 2 _
                                    AndAlso (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Substring(2, 1) = ":" Then
                                    reference &= "; " & (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If refCount < 3 _
                                    AndAlso (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 2 _
                                    AndAlso (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Substring(2, 1) = ":" Then
                                    reference &= "; " & (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If refCount < 3 _
                                    AndAlso (drSOTCART1.Item("REFERENCE3") & String.Empty).ToString.Trim.Length > 2 _
                                    AndAlso (drSOTCART1.Item("REFERENCE3") & String.Empty).ToString.Substring(2, 1) = ":" Then
                                    reference &= "; " & (drSOTCART1.Item("REFERENCE3") & String.Empty).ToString
                                    refCount += 1
                                End If

                                ' This is done because some customers want specific information on the label
                                If reference.Length > 0 Then
                                    refCount = 5
                                End If


                                ' Fedex allows up to 3 References
                                If (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If (drSOTCART1.Item("REFERENCE3") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & (drSOTCART1.Item("REFERENCE3") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If ORDR_CUST_PO.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & ORDR_CUST_PO
                                    refCount += 1
                                End If

                                If INV_NO.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; IN:" & INV_NO
                                    refCount += 1
                                End If

                                If (drSOTSHIP1.Item("ORDR_DEPT") & String.Empty).ToString.Trim <> String.Empty AndAlso refCount < 3 Then
                                    reference &= "; DN:" & drSOTSHIP1.Item("ORDR_DEPT") & String.Empty
                                    refCount += 1
                                End If

                                If reference.Length > 0 Then
                                    reference = reference.Substring(1).Trim
                                End If

                            Case WHCSHIP1.ProviderTypeUPS
                                ' Ups allows up to 2 References

                                If ASCMAIN1.CLIENT = "VAN" Then
                                    reference &= (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    reference &= (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    refCount = 5
                                End If

                                If refCount < 2 _
                                    AndAlso (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 2 _
                                    AndAlso (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Substring(2, 1) = ":" Then
                                    reference &= "; " & (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If refCount < 2 _
                                    AndAlso (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 2 _
                                    AndAlso (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Substring(2, 1) = ":" Then
                                    reference &= "; " & (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    refCount += 1
                                End If

                                ' This is done because some customers want specific information on the label
                                If reference.Length > 0 Then
                                    refCount = 5
                                End If

                                If (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 2 Then
                                    reference &= "; CR:" & (drSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    refCount += 1
                                Else
                                    If (drSOTPICK1.Item("CUST_STORE_NO") & String.Empty).ToString.Trim <> String.Empty AndAlso refCount < 2 Then
                                        reference &= "; ST:" & drSOTPICK1.Item("CUST_STORE_NO") & String.Empty
                                        refCount += 1
                                    End If
                                End If

                                If (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 2 Then
                                    reference &= "; CR:" & (drSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    refCount += 1
                                Else
                                    If (drSOTPICK1.Item("ORDR_CUST_PO") & String.Empty).ToString.Trim <> String.Empty AndAlso refCount < 2 Then
                                        reference &= "; PO:" & drSOTPICK1.Item("ORDR_CUST_PO") & String.Empty
                                        refCount += 1
                                    End If
                                End If

                                If reference.Length > 0 Then
                                    reference = reference.Substring(1).Trim
                                End If

                        End Select

                        If reference.Length > 0 Then
                            If reference.StartsWith(";") Then
                                reference = reference.Substring(1).Trim
                            End If

                            If Not reference.EndsWith(";") Then
                                reference &= ";"
                            End If
                        End If

                        .Reference = reference
                        .Id = pkgId.ToString("D8")
                    End With

                    clsShip.PackageDetailList.Add(shipPackageDetail)

                    drWHTSHPC2 = dst.Tables("WHTSHPC2").NewRow
                    drWHTSHPC2.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    drWHTSHPC2.Item("SHIP_PACKAGE_NO") = pkgId
                    drWHTSHPC2.Item("HEIGHT") = shipPackageDetail.Height
                    drWHTSHPC2.Item("INSURED_VALUE") = 0
                    drWHTSHPC2.Item("LENGTH") = shipPackageDetail.Length
                    drWHTSHPC2.Item("NET_CHARGE") = 0
                    drWHTSHPC2.Item("PACKAGING_TYPE") = Val(shipPackageDetail.PackagingType)
                    drWHTSHPC2.Item("TOTAL_DISCOUNT") = 0
                    drWHTSHPC2.Item("TOTAL_SURCHARGES") = 0
                    drWHTSHPC2.Item("TRACKING_NUMBER") = String.Empty
                    drWHTSHPC2.Item("WEIGHT") = Convert.ToInt16(shipPackageDetail.Weight)
                    drWHTSHPC2.Item("WIDTH") = shipPackageDetail.Width
                    drWHTSHPC2.Item("TRACKING_NO") = String.Empty

                    drWHTSHPC2.Item("CUST_REF") = ORDR_CUST_PO
                    drWHTSHPC2.Item("INV_BOL_NO") = SHIP_BOL_NO
                    drWHTSHPC2.Item("CART_NO") = drSOTCART1.Item("CART_NO") & String.Empty
                    drWHTSHPC2.Item("INV_NO") = INV_NO
                    drWHTSHPC2.Item("PO_ORDER_NO") = String.Empty
                    drWHTSHPC2.Item("DEPT_NO") = (drSOTPICK1.Item("ORDR_DEPT") & String.Empty).ToString.Trim

                    dst.Tables("WHTSHPC2").Rows.Add(drWHTSHPC2)
                Next

                If isInternationalShipment Then
                    ' Set the Customs value
                    clsShip.TotalCustomsValue = Val(drSOTPICK1.Item("PICK_AMT_CONF") & String.Empty)

                    For Each drSOTCART2 As DataRow In dst.Tables("SOTCART2").Select("PICK_NO = '" & PICK_NO & "'")
                        Dim STYLE_CODE As String = drSOTCART2.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = drSOTCART2.Item("COLOR_CODE")

                        If itemList.Contains(STYLE_CODE) Then Continue For

                        itemList.Add(STYLE_CODE)

                        Dim drICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        ' Just in case a non item is permitted in the shipment
                        If drICTSTYL1 Is Nothing Then Continue For

                        Dim CommodityDetail As New CommodityDetail
                        CommodityDetail.Description = drICTSTYL1.Item("STYLE_DESC") & String.Empty

                        Dim NumberOfPieces As Int32 = Val(dst.Tables("SOTCART2").Compute("SUM(QTY_PACKED)", "STYLE_CODE = '" & STYLE_CODE & "' and PICK_NO = '" & PICK_NO & "'") & String.Empty)

                        CommodityDetail.NumberOfPieces = NumberOfPieces
                        CommodityDetail.Quantity = NumberOfPieces
                        CommodityDetail.QuantityUnit = "EA"

                        Dim pickUnitPrice As Decimal = Val(dst.Tables("SOTPICK2").Compute("MAX(PICK_UNIT_PRICE)", "PICK_NO = '" & PICK_NO & "' AND STYLE_CODE = '" & STYLE_CODE & "'") & String.Empty)
                        CommodityDetail.UnitPrice = pickUnitPrice

                        CommodityDetail.Weight = Val(drICTSTYL1.Item("STYLE_WEIGHT") & String.Empty) ' Leave as pounds
                        CommodityDetail.Manufacturer = (drICTSTYL1.Item("COUNTRY_CODE") & String.Empty).ToString.ToUpper.Trim ' "US" '
                        If CommodityDetail.Manufacturer.Length = 0 Then
                            CommodityDetail.Manufacturer = "US"
                        End If

                        clsShip.CommodityDetailList.Add(CommodityDetail)

                        Dim drWHTSHPCC As DataRow = dst.Tables("WHTSHPCC").NewRow
                        drWHTSHPCC.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                        drWHTSHPCC.Item("COMMODITY_LNO") = COMMODITY_LNO
                        COMMODITY_LNO += 1
                        drWHTSHPCC.Item("COMMODITY_DESC") = CommodityDetail.Description
                        drWHTSHPCC.Item("NUM_PIECES") = CommodityDetail.NumberOfPieces
                        drWHTSHPCC.Item("MANUFACTURER") = CommodityDetail.Manufacturer
                        drWHTSHPCC.Item("HARMONIZED_CODE") = String.Empty
                        drWHTSHPCC.Item("WEIGHT") = CommodityDetail.Weight
                        drWHTSHPCC.Item("QUANTITY") = CommodityDetail.Quantity
                        drWHTSHPCC.Item("QUANTITY_UOM") = CommodityDetail.QuantityUnit
                        drWHTSHPCC.Item("UNIT_PRICE") = CommodityDetail.UnitPrice
                        dst.Tables("WHTSHPCC").Rows.Add(drWHTSHPCC)
                    Next
                End If
            Next  ' This is where the For Sotpick1, for sotcart1, for sotcart2 should end 

            clsShip.TotalCustomsValue = 0
            clsShip.SignatureRequired = False

            ' Shipping Method
            If isInternationalShipment Then
                If drSOTSVIA1.Item("CARRIER_PROD_CODE_INTL") & String.Empty <> String.Empty Then
                    clsShip.RequestedServiceType = Val(drSOTSVIA1.Item("CARRIER_PROD_CODE_INTL") & String.Empty)
                Else
                    clsShip.RequestedServiceType = Val(drSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty)
                End If
            Else
                clsShip.RequestedServiceType = Val(drSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty)
            End If

            If clsShip.RequestedServiceType = fedexSmartPost Then
                clsShip.FedexSmartPost.HubId = drSOTCARR3.Item("FEDEX_HUB_ID") & String.Empty
            End If

            clsShip.USPSEndorsement = WHCSHIP1.USPSEndorsements.NoServiceSelected

            ' The COLLECT payment type is only supported in FedEx Ground services. The CONSIGNEE type is only supported in UPS service.

            ' For FedEx, when this field is set to a value other than 0 (ptSender), the AccountNumber and 
            ' CountryCode are required to be provided in the request as well. Otherwise, those will default to AccountNumber and CountryCode.

            ' For UPS, when set to ptSender, the AccountNumber is automatically set to AccountNumber. 
            ' When ptRecipient is specified, AccountNumber and ZipCode are required to be provided in the request. 
            ' For return international shipments, this option is invalid for transportation charges. 
            ' And, when ptThirdParty has been specified, the AccountNumber, ZipCode and CountryCode are 
            ' required to be provided in the request. When ptConsignee is specified, it indicates that UPS Consignee Billing 
            ' option is selected, no other fields need to be set. ptConsignee only applies to US/PR and PR/US shipment origins and destination. 

            ' Payor of the Shipmenet
            clsShip.Payor = TPayorTypes.ptSender

            Dim drWHTSHPCP As DataRow
            drWHTSHPCP = dst.Tables("WHTSHPCP").NewRow
            drWHTSHPCP("SHIP_CNTL_NO") = SHIP_CNTL_NO
            drWHTSHPCP("PAYOR_TYPE") = "S"
            drWHTSHPCP("PAYOR_ACCT_NO") = clsShip.PayorContact.AccountNumber & String.Empty
            drWHTSHPCP("PAYOR_COUNTRY") = clsShip.PayorContact.CountryCode & String.Empty
            dst.Tables("WHTSHPCP").Rows.Add(drWHTSHPCP)

            ' Payor of the Duties
            clsShip.DutiesPayor = TPayorTypes.ptSender
            If isInternationalShipment Then
                clsShip.DutiesPayor = clsShip.Payor
                clsShip.DutiesPayorContact.AccountNumber = clsShip.PayorContact.AccountNumber
                clsShip.DutiesPayorContact.CountryCode = clsShip.PayorContact.CountryCode
                clsShip.DutiesPayorContact.ZipCode = clsShip.PayorContact.ZipCode
            End If

            drWHTSHPCP = dst.Tables("WHTSHPCP").NewRow
            drWHTSHPCP("SHIP_CNTL_NO") = SHIP_CNTL_NO
            drWHTSHPCP("PAYOR_TYPE") = "D"
            drWHTSHPCP("PAYOR_ACCT_NO") = clsShip.DutiesPayorContact.AccountNumber & String.Empty
            drWHTSHPCP("PAYOR_COUNTRY") = clsShip.DutiesPayorContact.CountryCode & String.Empty
            dst.Tables("WHTSHPCP").Rows.Add(drWHTSHPCP)

            With clsShip
                .EzshipLabelImage = EzshipLabelImageTypes.itZPL
                .ShippingLabelDirectory = ShippingLabelDirectory
                .ShippingLabelPrefix = SHIP_CNTL_NO
                .ShipDate = DateTime.Now.ToString("yyyy-MM-dd")
            End With

            Try
                BeginTrans()
                Update_Record_TDA("WHTSHPC1")
                Update_Record_TDA("WHTSHPC2")
                Update_Record_TDA("WHTSHPC3")
                Update_Record_TDA("WHTSHPC4")
                Update_Record_TDA("WHTSHPC5")
                Update_Record_TDA("WHTSHPCG")
                Update_Record_TDA("WHTSHPCA")
                Update_Record_TDA("WHTSHPCS")
                Update_Record_TDA("WHTSHPCP")
                Update_Record_TDA("WHTSHPCC")
                CommitTrans()
            Catch ex As Exception
                ErrorMessage &= " " & ex.Message
                Rollback()
            End Try

            ' Notifications
            Dim CUST_EMAIL As String = drARTCUST1.Item("CUST_EMAIL") & String.Empty
            CUST_EMAIL = CUST_EMAIL.Trim

            clsShip.ShipmentNotifications.Clear()
            If CUST_EMAIL.Length > 0 AndAlso drSOTCARR1.Item("CARRIER_SEND_NOTIFY") & String.Empty = "1" Then

                Dim notify As New WHCSHIP1.Notifications
                With notify
                    .email = CUST_EMAIL
                    .NotificationFlags = WHCSHIP1.NotifictaionTypes.On_Shipment
                    .Message = "Your Shipment from " & ROWs("ASTPARM1").Item("AS_PARM_INST_NAME") & " was picked up for shipment."
                End With
                clsShip.ShipmentNotifications.Add(notify)

                notify = New WHCSHIP1.Notifications
                With notify
                    .email = CUST_EMAIL
                    .NotificationFlags = WHCSHIP1.NotifictaionTypes.On_Deleivery
                    .Message = "Your Shipment from " & ROWs("ASTPARM1").Item("AS_PARM_INST_NAME") & " was delivered."
                End With
                clsShip.ShipmentNotifications.Add(notify)

                notify = New WHCSHIP1.Notifications
                With notify
                    .email = CUST_EMAIL
                    .NotificationFlags = WHCSHIP1.NotifictaionTypes.On_Exception
                    .Message = "Your Shipment from " & ROWs("ASTPARM1").Item("AS_PARM_INST_NAME") & " has a delivery problem."
                End With
                clsShip.ShipmentNotifications.Add(notify)
            End If

            If Not isInternationalShipment Then
                clsShip.CommodityDetailList.Clear()
            End If

            Select Case ASCMAIN1.CLIENT
                Case "RGI"
                    clsShip.ShipmentDescription = "Artificial Flowers / Home Decorations"
                Case "VAN"
                    clsShip.ShipmentDescription = "Undergarments"
                Case Else
                    clsShip.ShipmentDescription = "Garments"
            End Select

            clsShip.RequestedUPSInternationalForms.ShippersExportDeclarationInfo = New WHCSHIP1.ShippersExportDeclaration
            clsShip.RequestedUPSInternationalForms.ShippersExportDeclaration = False
            clsShip.RequestedUPSInternationalForms.CommercialInvoice = False

            If isInternationalShipment Then
                clsShip.RequestedUPSInternationalForms.ShippersExportDeclaration = True
                With clsShip.RequestedUPSInternationalForms.ShippersExportDeclarationInfo
                    .ImportEntryNumber = String.Empty
                    .InBond = TInBondCodes.ibcNotInBond
                    .LicenseDate = String.Empty
                    .LicenseExceptionCode = TExceptionCodes.ecNLR
                    .LicenseNumber = String.Empty
                    .PointOfOrigin = "US"
                    .ShippersTaxID = String.Empty
                    .TransPortType = String.Empty
                    .ExportingCarrier = CARRIER_CODE
                    .ExportingDate = System.DateTime.Now.ToString("yyyyMMdd")
                End With

                clsShip.RequestedUPSInternationalForms.CommercialInvoice = True
                With clsShip.RequestedUPSInternationalForms.CommercialInvoiceInfo
                    .Comments = String.Empty
                    .CustomersInvoiceNumber = CUST_CODE
                    .FreightCharge = 0
                    .InvoiceDate = System.DateTime.Now
                    .Purpose = CommercialInvoicePurposes.cipSold
                    .ShipperInsurance = 0
                    .Terms = CommercialInvoiceTerms.citCpt

                End With
            End If

            If clsShip.RequestLabel() Then

                drWHTSHPC1.Item("ERROR_MSG") = clsShip.LastError & String.Empty
                drWHTSHPC1.Item("STATUS") = "P"
                If drWHTSHPC1 IsNot Nothing AndAlso (drWHTSHPC1.Item("ERROR_MSG") & String.Empty).ToString.Length > 200 Then
                    drWHTSHPC1.Item("ERROR_MSG") = drWHTSHPC1("ERROR_MSG").ToString.Substring(0, 200).Trim
                End If

                If isPitneyBowes Then
                    drWHTSHPC1.Item("MASTER_TRACKING_NO") = clsShip.MasterTrackingNumber & String.Empty
                Else
                    drWHTSHPC1.Item("MASTER_TRACKING_NO") = clsShip.MasterTrackingNumber & String.Empty
                End If

                ' Update Pro Number if it is blank
                For Each dr As DataRow In dst.Tables("SOTSHIP1").Select("")
                    If dr.Item("SHIP_REF") & String.Empty = String.Empty Then
                        dr.Item("SHIP_REF") = clsShip.MasterTrackingNumber & String.Empty
                    End If
                Next

                For Each shipPackageDetail As PackageDetail In clsShip.PackageDetailList
                    SHIP_PACKAGE_NO = Val(shipPackageDetail.Id)
                    If dst.Tables("WHTSHPC2").Select("SHIP_PACKAGE_NO = " & SHIP_PACKAGE_NO, "").Length > 0 Then
                        drWHTSHPC2 = dst.Tables("WHTSHPC2").Select("SHIP_PACKAGE_NO = " & SHIP_PACKAGE_NO)(0)

                        Dim pitneyBowesshipdata As New TAC.WHCSHIP1.PitneyBowesPackageInformation

                        If isPitneyBowes Then
                            pitneyBowesshipdata = JsonConvert.DeserializeObject(Of TAC.WHCSHIP1.PitneyBowesPackageInformation)(shipPackageDetail.Reference)
                            drWHTSHPC2.Item("TRACKING_NO") = pitneyBowesshipdata.TrackingNumber & String.Empty
                            drWHTSHPC2.Item("TRACKING_NUMBER") = pitneyBowesshipdata.ShipmentID & String.Empty
                        Else
                            drWHTSHPC2.Item("TRACKING_NO") = shipPackageDetail.TrackingNumber & String.Empty
                        End If

                        drWHTSHPC2.Item("BASE_CHARGE") = Val(clsShip.ShipmentBaseCharge(SHIP_PACKAGE_NO) & String.Empty)
                        drWHTSHPC2.Item("NET_CHARGE") = Val(clsShip.ShipmentNetCharge(SHIP_PACKAGE_NO) & String.Empty)
                        drWHTSHPC2.Item("TOTAL_DISCOUNT") = Val(clsShip.ShipmentDiscountCharge(SHIP_PACKAGE_NO) & String.Empty)
                        drWHTSHPC2.Item("TOTAL_SURCHARGES") = Val(clsShip.ShipmentSurCharge(SHIP_PACKAGE_NO) & String.Empty)

                        drWHTSHPC2.Item("LENGTH") = Val(shipPackageDetail.Length & String.Empty)
                        drWHTSHPC2.Item("WIDTH") = Val(shipPackageDetail.Width & String.Empty)
                        drWHTSHPC2.Item("HEIGHT") = Val(shipPackageDetail.Height & String.Empty)

                        If clsShip.ShipmentListCharge.ContainsKey(SHIP_PACKAGE_NO) Then
                            drWHTSHPC2.Item("LIST_PRICE") = Val(clsShip.ShipmentListCharge(SHIP_PACKAGE_NO) & String.Empty)
                        Else
                            drWHTSHPC2.Item("LIST_PRICE") = drWHTSHPC2.Item("NET_CHARGE")
                        End If

                        OUR_FREIGHT = Val(drWHTSHPC2.Item("NET_CHARGE") & String.Empty)

                        ' Logic added 3/18/2017 for Regency
                        Select Case CARRIER_PPA_TYPE
                            Case "F" ' None
                                PPA_FREIGHT = 0

                            Case "L" ' List Rates
                                PPA_FREIGHT = Val(drWHTSHPC2.Item("LIST_PRICE") & String.Empty)

                            Case "N" ' Negioated Rates
                                PPA_FREIGHT = Val(drWHTSHPC2.Item("NET_CHARGE") & String.Empty)

                            Case Else
                                ' If not set then use List Price
                                PPA_FREIGHT = Val(drWHTSHPC2.Item("LIST_PRICE") & String.Empty)
                        End Select

                        If CARRIER_SURCHARGE_PERC > 0 Then
                            Select Case CARRIER_SURCHARGE_BASE
                                Case "N" ' Negotiated
                                    PPA_FREIGHT += Val(drWHTSHPC2.Item("NET_CHARGE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                                Case "L" ' List
                                    PPA_FREIGHT += Val(drWHTSHPC2.Item("LIST_PRICE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                                Case Else
                                    ' If not set then use List
                                    PPA_FREIGHT += Val(drWHTSHPC2.Item("LIST_PRICE") & String.Empty) * (CARRIER_SURCHARGE_PERC / 100)
                            End Select
                        End If

                        PICK_NO = String.Empty
                        drSOTPICK1 = Nothing

                        ' We may have multi pick tickets in a single carton. This stamps them with the same tracking number
                        ' Spread the Customer Freight Cost and Our freight cost across the Pick Tickets
                        Dim numPickTickets As Int16 = dst.Tables("SOTCART1").Select("CART_SEQ = " & SHIP_PACKAGE_NO).Length
                        For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("CART_SEQ = " & SHIP_PACKAGE_NO)

                            If isPitneyBowes Then
                                drSOTCART1.Item("CART_TRACKING_NO") = pitneyBowesshipdata.TrackingNumber & String.Empty
                            Else
                                drSOTCART1.Item("CART_TRACKING_NO") = shipPackageDetail.TrackingNumber & String.Empty
                            End If

                            PICK_NO = drSOTCART1.Item("PICK_NO") & String.Empty
                            drSOTPICK1 = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
                            If Absx1.txtFor("FRT_TERMS").Text = "PPA" Then
                                ' RGI charges freight for all Orders.
                                If ASCMAIN1.CLIENT = "RGI" Then
                                    drSOTPICK1.Item("PICK_FREIGHT") = Val(drSOTPICK1.Item("PICK_FREIGHT") & String.Empty) + Math.Round(PPA_FREIGHT / numPickTickets, 2)
                                ElseIf drSOTPICK1("ORDR_SOURCE") & String.Empty <> "W" Then
                                    drSOTPICK1.Item("PICK_FREIGHT") = Val(drSOTPICK1.Item("PICK_FREIGHT") & String.Empty) + Math.Round(PPA_FREIGHT / numPickTickets, 2)
                                End If
                            End If
                            drSOTPICK1.Item("OUR_FREIGHT") = Val(drSOTPICK1.Item("OUR_FREIGHT") & String.Empty) + Math.Round(OUR_FREIGHT / numPickTickets, 2)
                        Next
                        pitneyBowesshipdata = Nothing
                    End If

                    If isPitneyBowes Then
                        Dim file As String = shipPackageDetail.ShippingLabelFile
                        Using sr As New StreamReader(file)
                            ShippingLabels.Add(sr.ReadToEnd)
                            sr.Close()
                            sr.Dispose()
                        End Using
                    Else
                        ShippingLabels.Add(shipPackageDetail.ShippingLabel)
                        ShippingLabels.Add(shipPackageDetail.CODLabel)
                        ShippingLabels.Add(shipPackageDetail.ReturnReceipt)
                    End If
                Next

                Dim totalLabelCharge As Decimal = Math.Round(Val(dst.Tables("SOTPICK1").Compute("SUM(PICK_FREIGHT)", "") & String.Empty), 2)
                Dim rateCharge As Decimal = Math.Round(Val(dst.Tables("WHTSHPC4").Select($"CARRIER_CODE = '{CARRIER_CODE}' AND SERVICE_TYPE = '{CARRIER_PROD_CODE}'", "")(0).Item("TOTAL_CHARGE") & String.Empty))

                If CInt(totalLabelCharge) > CInt(rateCharge) Then
                    Dim diff As Decimal = Math.Round(totalLabelCharge - rateCharge, 2)
                    Dim userMessage As String = $"The Customer Freight Rate is {rateCharge.ToString("#,##0.00")} and the Label Charge is {totalLabelCharge.ToString("#,##0.00")}. This is a difference of {diff.ToString("#,##0.00")}."
                    userMessage &= Environment.NewLine & Environment.NewLine & "Do you want to continue?"
                    If MessageBox.Show(userMessage, "Freight Difference", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        ErrorMessage = "Discrepency in Customer Freight Rate and Label Charge. User cancelled Finalization."
                        For Each drSOTPICK1x As DataRow In dst.Tables("SOTPICK1").Select("")
                            drSOTPICK1x.Item("PICK_FREIGHT") = drSOTPICK1x.Item("PICK_FREIGHT_ORIG")
                            drSOTPICK1x.Item("OUR_FREIGHT") = 0
                        Next
                        Return False
                    Else
                        For Each drSOTPICK1x As DataRow In dst.Tables("SOTPICK1").Select("")
                            Dim drTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
                            drTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
                            drTATEVNT1.Item("TABLE_KEY") = drSOTPICK1x.Item("ORDR_NO")
                            drTATEVNT1.Item("INIT_DATE") = DateTime.Now
                            drTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                            drTATEVNT1.Item("EVENT_TYPE") = "LBFRT"
                            drTATEVNT1.Item("EVENT_DESC") = $"User {ASCMAIN1.USER_ID} choose to ship when the Customer Freight Rate was {rateCharge.ToString("#,##0.00")} and the Label Charge was {totalLabelCharge.ToString("#,##0.00")}"
                            drTATEVNT1.Item("EVENT_KEY") = ""
                            drTATEVNT1.Item("FORM_NAME") = "SOFSHIPB"
                            dst.Tables("TATEVNT1").Rows.Add(drTATEVNT1)
                        Next
                    End If
                End If

                Try
                    BeginTrans()
                    Update_Record_TDA("WHTSHPC1")
                    Update_Record_TDA("WHTSHPC2")
                    CommitTrans()
                Catch ex As Exception
                    ErrorMessage &= " " & ex.Message
                    Rollback()
                End Try

                If clsShip.InternationalFormsFile & String.Empty <> String.Empty Then
                    If My.Computer.FileSystem.FileExists(clsShip.InternationalFormsFile) Then
                        ShippingLabels.Add(clsShip.InternationalFormsFile)
                    End If
                End If

            Else
                ErrorMessage &= " " & clsShip.LastError
                RequestShippingLabel = False
            End If

        Catch ex As Exception
            ErrorMessage &= " " & ex.Message
            RequestShippingLabel = False
        End Try

        For Each shippingLabel As String In ShippingLabels
            If shippingLabel.Trim.Length > 0 Then PrintShippingLabels(shippingLabel)
        Next

        ErrorMessage = ErrorMessage.Trim

    End Function

    Sub Refresh_Refs(ByVal CarrierCode As String, Optional CART_NO As String = "")

        Dim REFERENCE1 As String = String.Empty
        Dim REF_CODE_1 As String = String.Empty
        Dim REF1_PREF As String = String.Empty
        Dim REF1_SUFF As String = String.Empty

        Dim REFERENCE2 As String = String.Empty
        Dim REF_CODE_2 As String = String.Empty
        Dim REF2_PREF As String = String.Empty
        Dim REF2_SUFF As String = String.Empty

        Dim REFERENCE3 As String = String.Empty
        Dim REF_CODE_3 As String = String.Empty
        Dim REF3_PREF As String = String.Empty
        Dim REF3_SUFF As String = String.Empty

        Dim drSOTCARRR As DataRow = Nothing

        If Not (CarrierCode = "FEDEX" OrElse CarrierCode = "UPS") Then
            Exit Sub
        End If

        If dst.Tables("ARTCUSTS").Rows.Count = 0 Then
            Exit Sub
        End If

        Dim drARTCUSTS As DataRow = dst.Tables("ARTCUSTS").Rows(0)

        Select Case CarrierCode

            Case "UPS"
                If drARTCUSTS.Item("UPS_REF1") & String.Empty <> String.Empty Then
                    REF_CODE_1 = drARTCUSTS.Item("UPS_REF1") & String.Empty
                    REF1_PREF = drARTCUSTS.Item("UPS_REF1_PREF") & String.Empty
                    REF1_SUFF = drARTCUSTS.Item("UPS_REF1_SUFF") & String.Empty

                    drSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {CarrierCode, REF_CODE_1})
                    If drSOTCARRR IsNot Nothing Then
                        REFERENCE1 = REF_CODE_1.Substring(0, 2) & ":" & drSOTCARRR.Item("TABLE_NAME") & "." & drSOTCARRR.Item("COLUMN_NAME")
                    End If
                End If

                If drARTCUSTS.Item("UPS_REF2") & String.Empty <> String.Empty Then
                    REF_CODE_2 = drARTCUSTS.Item("UPS_REF2") & String.Empty
                    REF2_PREF = drARTCUSTS.Item("UPS_REF2_PREF") & String.Empty
                    REF2_SUFF = drARTCUSTS.Item("UPS_REF2_SUFF") & String.Empty

                    drSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {CarrierCode, REF_CODE_2})
                    If drSOTCARRR IsNot Nothing Then
                        REFERENCE2 = REF_CODE_2.Substring(0, 2) & ":" & drSOTCARRR.Item("TABLE_NAME") & "." & drSOTCARRR.Item("COLUMN_NAME")
                    End If
                End If

                If drARTCUSTS.Item("UPS_REF3") & String.Empty <> String.Empty Then
                    REF_CODE_3 = drARTCUSTS.Item("UPS_REF3") & String.Empty
                    REF3_PREF = drARTCUSTS.Item("UPS_REF3_PREF") & String.Empty
                    REF3_SUFF = drARTCUSTS.Item("UPS_REF3_SUFF") & String.Empty

                    drSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {CarrierCode, REF_CODE_3})
                    If drSOTCARRR IsNot Nothing Then
                        REFERENCE3 = REF_CODE_3.Substring(0, 2) & ":" & drSOTCARRR.Item("TABLE_NAME") & "." & drSOTCARRR.Item("COLUMN_NAME")
                    End If
                End If

            Case "FEDEX"

                If drARTCUSTS.Item("FDX_REF1") & String.Empty <> String.Empty Then
                    REF_CODE_1 = drARTCUSTS.Item("FDX_REF1") & String.Empty
                    REF1_PREF = drARTCUSTS.Item("FDX_REF1_PREF") & String.Empty
                    REF1_SUFF = drARTCUSTS.Item("FDX_REF1_SUFF") & String.Empty

                    drSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {CarrierCode, REF_CODE_1})
                    If drSOTCARRR IsNot Nothing Then
                        REFERENCE1 = REF_CODE_1.Substring(0, 2) & ":" & drSOTCARRR.Item("TABLE_NAME") & "." & drSOTCARRR.Item("COLUMN_NAME")
                    End If
                End If

                If drARTCUSTS.Item("FDX_REF2") & String.Empty <> String.Empty Then
                    REF_CODE_2 = drARTCUSTS.Item("FDX_REF2") & String.Empty
                    REF2_PREF = drARTCUSTS.Item("FDX_REF2_PREF") & String.Empty
                    REF2_SUFF = drARTCUSTS.Item("FDX_REF2_SUFF") & String.Empty

                    drSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {CarrierCode, REF_CODE_2})
                    If drSOTCARRR IsNot Nothing Then
                        REFERENCE2 = REF_CODE_2.Substring(0, 2) & ":" & drSOTCARRR.Item("TABLE_NAME") & "." & drSOTCARRR.Item("COLUMN_NAME")
                    End If
                End If

                If drARTCUSTS.Item("FDX_REF3") & String.Empty <> String.Empty Then
                    REF_CODE_3 = drARTCUSTS.Item("FDX_REF3") & String.Empty
                    REF3_PREF = drARTCUSTS.Item("FDX_REF3_PREF") & String.Empty
                    REF3_SUFF = drARTCUSTS.Item("FDX_REF3_SUFF") & String.Empty

                    drSOTCARRR = dst.Tables("SOTCARRR").Rows.Find(New Object() {CarrierCode, REF_CODE_3})
                    If drSOTCARRR IsNot Nothing Then
                        REFERENCE3 = REF_CODE_3.Substring(0, 2) & ":" & drSOTCARRR.Item("TABLE_NAME") & "." & drSOTCARRR.Item("COLUMN_NAME")
                    End If
                End If

        End Select

        Dim temp1 As String = String.Empty
        Dim temp2 As String = String.Empty
        Dim temp3 As String = String.Empty

        If CART_NO.Length > 0 Then
            CART_NO = "CART_NO = '" & CART_NO & "'"
        End If

        For Each drSOTCART1 As DataRow In dst.Tables("SOTCART1").Select(CART_NO, "SHIP_BOL_NO,CART_NO")

            temp1 = String.Empty
            temp2 = String.Empty
            temp3 = String.Empty

            If REFERENCE1.Length > 0 Then
                temp1 = GetReferenceValue(REF1_PREF, REF1_SUFF, REFERENCE1.Split(":")(1), drSOTCART1.Item("CART_NO"))
                If temp1.Length > 0 Then
                    temp1 = REFERENCE1.Split(":")(0).Substring(0, 2) & ":" & temp1
                End If
            End If

            If REFERENCE2.Length > 0 Then
                temp2 = GetReferenceValue(REF2_PREF, REF2_SUFF, REFERENCE2.Split(":")(1), drSOTCART1.Item("CART_NO"))
                If temp2.Length > 0 Then
                    temp2 = REFERENCE2.Split(":")(0).Substring(0, 2) & ":" & temp2
                End If
            End If

            If REFERENCE3.Length > 0 Then
                temp3 = GetReferenceValue(REF3_PREF, REF3_SUFF, REFERENCE3.Split(":")(1), drSOTCART1.Item("CART_NO"))
                If temp3.Length > 0 Then
                    temp3 = REFERENCE3.Split(":")(0).Substring(0, 2) & ":" & temp3
                End If
            End If

            If CarrierCode = "FEDEX" AndAlso drSOTCARRR IsNot Nothing Then
                If temp1 <> "" Then
                    If REF_CODE_1.Length = 0 Then
                        REF_CODE_1 = temp1.Substring(0, 2)
                    End If
                    temp1 = Replace(temp1, temp1.Substring(0, 3), Mid(REF_CODE_1, 1, 2) & ":")
                End If

                If temp2 <> "" Then
                    If REF_CODE_2.Length = 0 Then
                        REF_CODE_2 = temp2.Substring(0, 2)
                    End If
                    temp2 = Replace(temp2, temp2.Substring(0, 3), Mid(REF_CODE_2, 1, 2) & ":")
                End If

                If temp3 <> "" Then
                    If REF_CODE_3.Length = 0 Then
                        REF_CODE_3 = temp3.Substring(0, 2)
                    End If
                    temp3 = Replace(temp3, temp3.Substring(0, 3), Mid(REF_CODE_3, 1, 2) & ":")
                End If

                If temp1.StartsWith("ST:") Then
                    temp1 = temp1.Replace("ST:", "CR:")
                End If

                If temp2.StartsWith("ST:") Then
                    temp2 = temp2.Replace("ST:", "CR:")
                End If

                If temp3.StartsWith("ST:") Then
                    temp3 = temp3.Replace("ST:", "CR:")
                End If
            End If

            If temp1.Length > 0 Then drSOTCART1.Item("REFERENCE1") = temp1
            If temp2.Length > 0 Then drSOTCART1.Item("REFERENCE2") = temp2
            If temp3.Length > 0 Then drSOTCART1.Item("REFERENCE3") = temp3
        Next

    End Sub

    Private Function GetReferenceValue(ByVal Prefix As String, ByVal Suffix As String, ByVal field As String, ByVal CART_NO As String) As String

        Dim TABLE_NAME As String = String.Empty
        Dim COLUMN_NAME As String = String.Empty
        Dim dataRow As DataRow = Nothing
        Dim referenceValue As String = String.Empty

        Try
            TABLE_NAME = field.Split(".")(0)
            COLUMN_NAME = field.Split(".")(1)

            Select Case TABLE_NAME
                Case "SOTCART1"
                    dataRow = dst.Tables("SOTCART1").Rows.Find(CART_NO)

                Case "SOTPICK1"
                    dataRow = dst.Tables("SOTCART1").Rows.Find(CART_NO)
                    If dataRow Is Nothing Then
                        Exit Select
                    End If
                    dataRow = dst.Tables("SOTPICK1").Rows.Find(dataRow.Item("PICK_NO") & String.Empty)

                Case "SOTORDR1"
                    dataRow = dst.Tables("SOTCART1").Rows.Find(CART_NO)
                    If dataRow Is Nothing Then
                        Exit Select
                    End If
                    dataRow = dst.Tables("SOTPICK1").Rows.Find(dataRow.Item("PICK_NO") & String.Empty)
                    If dataRow Is Nothing Then
                        Exit Select
                    End If
                    dataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR1 WHERE ORDR_NO = '" & dataRow.Item("ORDR_NO") & "'")

                Case "SOTSHIP1"
                    dataRow = dst.Tables("SOTCART1").Rows.Find(CART_NO)
                    If dataRow Is Nothing Then
                        Exit Select
                    End If
                    dataRow = dst.Tables("SOTPICK1").Rows.Find(dataRow.Item("PICK_NO") & String.Empty)
                    If dataRow Is Nothing Then
                        Exit Select
                    End If
                    dataRow = dst.Tables("SOTSHIP1").Rows.Find(dataRow.Item("SHIP_BOL_NO") & String.Empty)

                Case "EDT850T1"
                    dataRow = dst.Tables("SOTCART1").Rows.Find(CART_NO)
                    If dataRow Is Nothing Then Exit Select

                    dataRow = dst.Tables("SOTPICK1").Rows.Find(dataRow.Item("PICK_NO") & String.Empty)
                    If dataRow Is Nothing Then Exit Select

                    dataRow = dst.Tables("SOTORDR1").Rows.Find(dataRow.Item("ORDR_NO") & String.Empty)
                    If dataRow Is Nothing Then Exit Select

                    dataRow = dst.Tables("EDT850T1").Rows.Find(dataRow.Item("EDI_DOC_SEQ_NO") & String.Empty)

                Case String.Empty
                    If COLUMN_NAME = String.Empty Then
                        referenceValue = Prefix & Suffix
                        Return referenceValue.Trim
                    End If

            End Select

            If dataRow Is Nothing Then
                Return String.Empty
            End If

            If dataRow.Item(COLUMN_NAME) & String.Empty = String.Empty Then
                Return String.Empty
            End If

            referenceValue = Prefix & dataRow.Item(COLUMN_NAME) & String.Empty & Suffix
            referenceValue = referenceValue.Trim
            Return referenceValue


        Catch ex As Exception
            Return String.Empty
        End Try

    End Function

    Public Function PrintShippingLabels(ByVal LabelData As String) As Boolean

        Try
            If LabelData.ToUpper.EndsWith(".PDF") Then
                If My.Computer.FileSystem.FileExists(LabelData) Then
                    Dim waitAmount As Int16 = 0
                    Dim requestedIpAddress As String = txtLaserPrinter.Text
                    If LabelData.ToUpper.EndsWith(".ZPL") Then
                        requestedIpAddress = txtLabelPrinter.Text
                    End If
                    Dim requestedStreamPort As String = "9100"

                    If requestedIpAddress.Contains(":") Then
                        requestedStreamPort = requestedIpAddress.Split(":")(1)
                        requestedIpAddress = requestedIpAddress.Split(":")(0)
                    End If

                    Using ipp As New nsoftware.IPWorks.Ipport
                        ipp.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareipportkey")
                        'ipp.Config("SSLEnabledProtocols=" & TAC.TACMAIN1.SSLEnabledProtocols)
                        ipp.Connect(requestedIpAddress, Val(requestedStreamPort))
                        Using binaryReader As New System.IO.BinaryReader(System.IO.File.Open(LabelData, System.IO.FileMode.Open))
                            waitAmount = 0
                            Do
                                'ErrorMessage = "If Not ipp.Connected Then"
                                If Not ipp.Connected Then
                                    System.Threading.Thread.Sleep(1000)
                                    waitAmount += 1
                                End If
                            Loop While Not ipp.Connected And waitAmount < 15

                            'ErrorMessage = "ipp.Send(binaryReader.ReadBytes(binaryReader.BaseStream.Length))"
                            ipp.Send(binaryReader.ReadBytes(binaryReader.BaseStream.Length))
                            System.Threading.Thread.Sleep(1000)
                            binaryReader.Close()
                            binaryReader.Dispose()
                        End Using

                        'ErrorMessage = "ipp.Disconnect()"
                        ipp.Disconnect()
                        ipp.Dispose()
                    End Using
                End If
                Return True
            End If

            If IsIPAddress(txtLabelPrinter.Text) Then
                clsTACZPLT1.SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, LabelData)
            Else
                ASCMAIN1.LabelPrinterSerialPort.WriteLine(LabelData)
            End If


        Catch ex As Exception
            MessageBox.Show("Print Shipping Label Error: " & ex.Message)
        End Try

    End Function

    Function IsIPAddress(value As String) As Boolean
        Dim ip As IPAddress = Nothing
        Return IPAddress.TryParse(value, ip)
    End Function


#End Region

#Region "grdWHTSHPC4"

    Private Sub grdWHTSHPC4_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTSHPC4.DoubleClickRow
        Dim CARRIER_CODE As String = e.Row.Cells("CARRIER_CODE").Value
        Dim CARRIER_PROD_CODE As String = e.Row.Cells("SERVICE_TYPE").Value & String.Empty
        Dim SHIP_VIA_CODE As String = e.Row.Cells("SHIP_VIA_CODE").Value & String.Empty

        ASCMAIN1.sql = $"CARRIER_CODE = '{CARRIER_CODE}' AND CARRIER_PROD_CODE = '{CARRIER_PROD_CODE}' AND SHIP_VIA_CODE = '{SHIP_VIA_CODE}' AND SHIP_VIA_STATUS = 'A'"
        If dst.Tables("SOTSVIA1").Select(ASCMAIN1.sql).Length > 0 Then
            txtSHIP_VIA_CODE.Text = SHIP_VIA_CODE
        End If

    End Sub

    Private Sub grdWHTSHPC4_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdWHTSHPC4.InitializeRow

        If e.Row.Cells("DISCLAIMER").Value & String.Empty <> String.Empty Then
            e.Row.Cells("CARRIER_CODE").Appearance.BackColor = Drawing.Color.DarkMagenta
            e.Row.Cells("CARRIER_CODE").Appearance.ForeColor = Drawing.Color.White
        End If

        If e.Row.Cells("TOTAL_CHARGE").Value > e.Row.Cells("LIST_NET_CHARGE").Value + e.Row.Cells("SURCHARGE").Value Then
            e.Row.Cells("TOTAL_CHARGE").Appearance.FontData.Bold = DefaultableBoolean.True
            e.Row.Cells("TOTAL_CHARGE").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("TOTAL_CHARGE").Appearance.FontData.Bold = DefaultableBoolean.False
            e.Row.Cells("TOTAL_CHARGE").Appearance.ForeColor = Drawing.Color.Black
        End If

        Dim CARRIER_CODE As String = e.Row.Cells("CARRIER_CODE").Value & String.Empty
        Dim CARRIER_PROD_CODE As String = e.Row.Cells("SERVICE_TYPE").Value & String.Empty
        Dim SHIP_VIA_CODE As String = e.Row.Cells("SHIP_VIA_CODE").Value & String.Empty

        If CARRIER_CODE.Length > 0 AndAlso CARRIER_PROD_CODE.Length > 0 Then
            ASCMAIN1.sql = "SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "' and CARRIER_CODE = '" & CARRIER_CODE & "' AND CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "' and SHIP_VIA_STATUS = 'A'"
            If dst.Tables("SOTSVIA1").Select(ASCMAIN1.sql).Length = 0 Then
                If dst.Tables("SOTSVIA1").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "' and SHIP_VIA_STATUS = 'A'")(0).Item("SHIP_VIA_CODE").length > 0 Then
                    e.Row.Cells("SHIP_VIA_CODE").Value = dst.Tables("SOTSVIA1").Select("CARRIER_CODE = '" & CARRIER_CODE & "' AND CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "' and SHIP_VIA_STATUS = 'A'")(0).Item("SHIP_VIA_CODE")
                    grdWHTSHPC4.UpdateData()
                End If
            End If
        End If
    End Sub

    Private Sub grdWHTSHPC4_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdWHTSHPC4.BeforeRowUpdate

        If e.Row.Band.Key <> grdWHTSHPC4.DisplayLayout.Bands(0).Key Then
            Exit Sub
        End If

        If e.Row.Cells("SELECTED").Value & String.Empty = "1" Then
            For Each dRow As DataRow In dst.Tables("WHTSHPC4").Select("SELECTED = '1'", "", DataViewRowState.CurrentRows)
                dRow.Item("SELECTED") = "0"
            Next
        End If

    End Sub

#End Region

#Region "Overrides"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

            Case "TRUCK_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    txtTRUCK_NO.Text = txtTRUCK_NO.Text.Trim.ToUpper
                    If txtTRUCK_NO.TextLength > 0 Then
                        If Not ValidateTruckTote(ValidateTruckToteTypes.Truck, txtTRUCK_NO.Text) Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "TOTE_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    txtTOTE_NO.Text = txtTOTE_NO.Text.ToUpper.Trim
                    If txtTOTE_NO.TextLength > 0 Then
                        ValidateTruckTote(ValidateTruckToteTypes.Tote, txtTOTE_NO.Text)
                    End If

                    If dst.Tables("SOTPICK1X").Select("ISNULL(SELECTED, '0') = '0'").Length = 0 Then
                        AutoCancel = True
                        ASCMAIN1.Progress("Resetting screen for next Truck", "")
                        Click_Command("Cancel")
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "Devices"

    ''' <summary>
    ''' Sets up and Initializes the Scanner Control
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetupScanner()

        scannedDelegate = AddressOf ProcessScannedData
        Try
            txtComPort.Appearance.BackColor = Drawing.Color.Red
            txtComPort.Clear()

            'If ASCMAIN1.ScannerSerialPort IsNot Nothing Then
            '    txtComPort.Appearance.BackColor = Drawing.Color.Green
            '    txtComPort.Text = ASCMAIN1.ScannerSerialPort.PortName
            'End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub SetUpPortsAndPrinters()

        Dim prtdoc As New System.Drawing.Printing.PrintDocument
        txtLaserPrinter.Text = prtdoc.PrinterSettings.PrinterName
        'ASCMAIN1.InvoicePrinterIpAddress = txtInvoicePrinter.Text

        txtLabelPrinter.Appearance.BackColor = Drawing.Color.LightGreen
        If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing Then
            txtLabelPrinter.Text = "Serial " & ASCMAIN1.LabelPrinterSerialPort.PortName
            Try
                If Not ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                    ASCMAIN1.LabelPrinterSerialPort.Open()
                End If
            Catch ex As Exception
                txtLabelPrinter.Appearance.BackColor = Drawing.Color.Red
                txtLabelPrinter.Appearance.ForeColor = Drawing.Color.White
            End Try
        ElseIf ASCMAIN1.LabelPrinterIPAddress.Length > 0 Then
            ' AEG IP Port label printing
            txtLabelPrinter.Text = ASCMAIN1.LabelPrinterIPAddress
        ElseIf ASCMAIN1.LabelPrinterName.Length Then
            txtLabelPrinter.Text = ASCMAIN1.LabelPrinterName
        Else
            txtLabelPrinter.Text = "No Port"
            txtLabelPrinter.Appearance.BackColor = Drawing.Color.Red
            txtLabelPrinter.Appearance.ForeColor = Drawing.Color.White
        End If

    End Sub

    Private Sub CreateAppearances()

        dictAppearances.Add("BLUE", New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.LightBlue})
        dictAppearances.Add("GREEN", New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.LightGreen})
        dictAppearances.Add("RED", New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.LightPink})
        dictAppearances.Add("TAN", New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.Tan})
        dictAppearances.Add("YELLOW", New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.LightYellow})
        dictAppearances.Add("GRAY", New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.LightGray})
        dictAppearances.Add("ORANGE", New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.Orange})
        dictAppearances.Add("PINK", New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.Pink})

        dictAppearances.Add("BLACK", New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.DarkGray, .ForeColor = Drawing.Color.White})
        dictAppearances.Add("ALL_ITEMS_BACK", New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.DarkSlateGray, .ForeColor = Drawing.Color.White})

        Appearance_Incomplete.BackColor = System.Drawing.Color.Red
        Appearance_Incomplete.ForeColor = System.Drawing.Color.White

    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdSOTTRCK1X_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTTRCK1X.DoubleClickRow

        If grdSOTTRCK1X.ActiveRow Is Nothing Then
            Exit Sub
        End If

        If grdSOTTRCK1X.ActiveRow.IsFilterRow OrElse grdSOTTRCK1X.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        txtTRUCK_NO.Text = grdSOTTRCK1X.ActiveRow.Cells("TRUCK_NO").Value & String.Empty
        ValidateTruckTote(ValidateTruckToteTypes.Truck, txtTRUCK_NO.Text)
    End Sub

    Private Sub grdSOTPICK1X_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTPICK1X.AfterRowActivate

        grdSOTPICK2.Text = String.Empty
        Dim PICK_NO As String = grdSOTPICK1X.ActiveRow.Cells("PICK_NO").Value & String.Empty
        Dim TOTE_NO As String = grdSOTPICK1X.ActiveRow.Cells("TOTE_NO").Value & String.Empty

        Dim dvw As DataView = DirectCast(grdSOTPICK2.DataSource, DataTable).DefaultView
        grdSOTPICK2.Text = $"Details for Tote: {TOTE_NO}, Pick Ticket: {PICK_NO}"
        dvw.RowFilter = $"PICK_NO = '{PICK_NO}'"

    End Sub

    Private Sub grdSOTPICK1X_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTPICK1X.InitializeRow
        If e.Row.Cells("INCOMPLETE").Value & String.Empty = "1" Then
            e.Row.Appearance = Appearance_Incomplete
        ElseIf e.Row.Cells("ALL_ITEMS_BACK").Value & String.Empty = "1" Then
            e.Row.Cells("TOTE_NO").Appearance = dictAppearances("ALL_ITEMS_BACK")
        Else
            e.Row.Appearance = Nothing
            If dictAppearances.ContainsKey(e.Row.Cells("BAY_COLOR").Value & String.Empty) Then
                e.Row.Appearance = dictAppearances(e.Row.Cells("BAY_COLOR").Value & String.Empty)
            End If
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Stop()

        If ScreenMode Then
            txtTOTE_NO.Clear()
            txtTOTE_NO.Focus()
        Else
            txtTRUCK_NO.Clear()
            txtTRUCK_NO.Focus()
        End If

    End Sub

    Private Sub btnLabelPrinter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLabelPrinter.Click
        Try
            If txtLabelPrinter.Text.Trim.Length = 0 Then
                MessageBox.Show("There is no assigned Label Printer.", "Test Print", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            'clsTACZPLT1.PrintSampleShippingLabel()
            MessageBox.Show("Test Label sent to Printer " & txtLabelPrinter.Text, "Test Print", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Label Printer Test Print Error {ex.Message}", "Test Print", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnLaserPrinter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLaserPrinter.Click
        Try
            If txtLaserPrinter.Text.Trim.Length = 0 Then
                MessageBox.Show("There is no assigned Invoice Printer.", "Test Print", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            'PrintInvoice("", True)

        Catch ex As Exception
            MessageBox.Show($"Test Invoice Printer Error: {ex.Message}", "Test Print", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyDown(e)

        If e.KeyCode = System.Windows.Forms.Keys.F8 Then
            'Try
            '    With UltraExplorerBar1.Groups("Screen Control").Items("Update")
            '        If .Visible AndAlso .Settings.Enabled = DefaultableBoolean.True Then
            '            Me.Validate()
            '            UltraExplorerBar1.Focus()
            '            Click_Command("Update")
            '            e.Handled = True
            '            Exit Sub
            '        End If
            '    End With

            'Catch ex As Exception

            'End Try
        End If
    End Sub

#End Region

End Class