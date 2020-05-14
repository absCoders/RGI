Public Class WHFDASH1

    Dim connection_is_down As Boolean = False
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Get_PARM("POTPARM1")

        With dst
            With .Tables.Add("WHTDASH1")
                .Columns.Add("LINE_NO", GetType(System.Int32))
                .Columns.Add("LINE_DESC")
                .Columns.Add("LINE_CATGY")
                .Columns.Add("LINE_COUNT", GetType(System.Int32))
                .Columns.Add("LINE_SQL")
                .Columns.Add("LINE_WHOS_BALL")
                .Columns.Add("LINE_ACTION")
                .PrimaryKey = New DataColumn() {.Columns("LINE_NO")}
            End With

            With .Tables.Add("WHTCOLN1")
                .Columns.Add("COLUMN_NAME")
                .Columns.Add("COLUMN_CAPTION")
                .Columns.Add("COLUMN_WIDTH", GetType(System.Int32))
                .Columns.Add("COLUMN_FORMAT")
                .PrimaryKey = New DataColumn() {.Columns("COLUMN_NAME")}
            End With

            'ASCMAIN1.sql = "Select POTSHIP1.*" & vbCrLf _
            '& ", POTSHIP2.CONTAINER_NO, POTSHIP2.PO_SHIP_STATUS" & vbCrLf _
            '& ", POTSHIP2.PO_SHIPMENT_LNO, POTSHIP2.BOL_NO, POTSHIP2.PO_SHIP_CTNS" & vbCrLf _
            '& " from POTSHIP1,POTSHIP2 " & vbCrLf _
            '& " where POTSHIP2.PO_SHIPMENT_NO = POTSHIP1.PO_SHIPMENT_NO" & vbCrLf _
            '& "   and POTSHIP2.PO_SHIP_STATUS = 'O' "
            'Create_TDA(.Tables.Add, "POTSHIPX", "**", 0, False, "", 0)
        End With
        Initialize_WHTCOLN1()
        Initialize_WHTDASH1()

        grdWHTDASH1.DataSource = dst.Tables("WHTDASH1")
        'grdWHTDASH2.DataSource = dst.Tables("POTSHIP3")

        Create_Summary(grdWHTDASH1, "LINE_NO", "Count")

        'Create_Summary(grdWHTDASH2, "PO_SHIPMENT_LNO", "Count")
        'Create_Summary(grdWHTDASH2, New String() {"PO_QTY_SHP", "UNITS"})

        Set_Read_Only_for_ctl(grpLINE_WHOS_BALL, True)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Refresh"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Refresh"
                Load_WHTDASH1()

            Case "Print"
                Print_Record()

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    '.Items("View").Settings.Enabled = not_iScreenMode
                    '.Items("Done").Settings.Enabled = iScreenMode
                    '.Items("Print").Settings.Enabled = iScreenMode

                    '.Items("View").Visible = (EntryMode = "L" Or Not ScreenMode)
                    '.Items("Done").Visible = (EntryMode = "L" And ScreenMode)
                    '.Items("Print").Visible = ScreenMode
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        splWHTDASHx.Visible = Not tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"WHTDASH1"} ' , "POTSHIP3"}
            ' dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Load_WHTDASH1()
    End Sub

    Sub Load_Record()
        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        CommitTrans("Delete")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Stop
        CommitTrans("Update Complete")
    End Sub

    Sub Print_Record()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'Generate_Report("PORWREC2")
        'Print_Report_End()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "PO_SHIPMENT_NO"
                'sql_where = "STATUS = '0'"
        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        '  Load_Popup_Menu(grdWHTDASH1, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "PO Inquiry", "PO Shipment Inquiry")
        Load_Popup_Menu(grdWHTDASH2, "BBB", "PO Inquiry", "PO Shipment Inquiry", "Customer Order Inquiry")
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

        Select Case e.SourceControl.Name
            'Case "grdPOTORDRR"
            '    If EntryMode = "V" Then e.Cancel = True

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdWHTDASH2"
                Dim TBL As DataTable = DirectCast(grdWHTDASH2.DataSource, DataTable)
                tlb_pop.Tools("Customer Order Inquiry").SharedProps.Visible = TBL.Columns.Contains("CUST_CODE") ' (grdWHTDASH2.DisplayLayout.Bands(0).Columns.Contains("CUST_CODE"))
                tlb_pop.Tools("PO Shipment Inquiry").SharedProps.Visible = TBL.Columns.Contains("PO_SHIPMENT_NO")
                tlb_pop.Tools("PO Inquiry").SharedProps.Visible = TBL.Columns.Contains("PO_ORDER_NO")
                '       tlb_btn = DirectCast(tlb_pop.Tools("Update Style Master (+ POs) with Changes made to Case Packs"), UltraWinToolbars.ButtonTool)
                '       tlb_btn.SharedProps.Visible = grdPOTORDRS.DisplayLayout.Bands(0).Columns("STYLE_ACTION").Hidden And Not (MENU_ITEM_OBJECT = "POFORDRI")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdPOTORDR3"
                '    tlb_sbt = DirectCast(tlb.Tools("Show Cartons"), UltraWinToolbars.StateButtonTool)
                '    e.Tool.SharedProps.Visible = tlb_sbt.Checked

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            'Case "Style Multi-Color"
            '    Using F As New TAC.ICFSTYCX
            '        F.STYLE_CODE = ""
            '        F.Price_Caption = "Cost" & IIf(ssdDZGRD.Value = 1, "", "/Dz")
            '        F.ShowDialog()
            '        If F.STYLE_CODE <> "" Then
            '            Add_Colors(F.STYLE_CODE, F.dst.Tables("ICTCOLRM"), F.PRICE)
            '        End If
            '    End Using

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "POE")

            Case "PO Shipment Inquiry"
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Text
                Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPI", "F", "POE")

            Case "Customer Order Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Dim FIND_BY As String = CUST_CODE
                Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Text
                FIND_BY &= ":" & ORDR_GROUP_NO
                Context_Launch("Select", FIND_BY, e.Tool.Key, "SOFCORD1")
        End Select
    End Sub


#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "PO_SHIPMENT_NO"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Call Click_Command("Load", e)
            '    End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "PO_SHIPMENT_NO"
            '    Call Click_Command("View")
        End Select
    End Sub

#End Region

    Private Sub grdPOTSHIPX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTDASH1.AfterRowActivate
        Setup_WHTDASH2()
    End Sub

    Private Sub grdPOTSHIPX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTDASH1.DoubleClickRow
        'If grdPOTSHIPX.ActiveRow IsNot Nothing Then
        '    Absx1.txtFor("PO_SHIPMENT_NO").Text = grdPOTSHIPX.ActiveRow.Cells("PO_SHIPMENT_NO").Text
        '    Click_Command("View")
        'End If

    End Sub

    Sub Setup_WHTDASH2()
        If grdWHTDASH1.ActiveRow Is Nothing OrElse Not grdWHTDASH1.ActiveRow.IsDataRow Then
            grdWHTDASH2.Visible = False
        Else
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Scanning Database")
            grdWHTDASH2.DataSource = Nothing
            grdWHTDASH2.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
            grdWHTDASH2.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
            Dim sql As String = grdWHTDASH1.ActiveRow.Cells("LINE_SQL").Value & ""
            If sql <> "" Then
                Try
                    ASCMAIN1.sql = sql
                    Dim TBL As DataTable = ASCDATA1.GetDataTable
                    grdWHTDASH2.DataSource = TBL
                    grdWHTDASH1.ActiveRow.Cells("LINE_COUNT").Value = TBL.Rows.Count
                    grdWHTDASH1.ActiveRow.Update()

                    For Each gcol As UltraWinGrid.UltraGridColumn In grdWHTDASH2.DisplayLayout.Bands(0).Columns
                        Dim rowWHTCOLN1 As DataRow = dst.Tables("WHTCOLN1").Rows.Find(gcol.Key)
                        If rowWHTCOLN1 IsNot Nothing Then
                            gcol.Header.Caption = rowWHTCOLN1.Item("COLUMN_CAPTION")
                            gcol.Width = Val(rowWHTCOLN1.Item("COLUMN_WIDTH") & "")
                            gcol.Format = rowWHTCOLN1.Item("COLUMN_FORMAT")
                        End If
                    Next
                    Dim DTS As String = " as of " & Format(Now, "MM/dd/yyyy HH:mm")
                    grdWHTDASH2.Text = grdWHTDASH1.ActiveRow.Cells("LINE_DESC").Value & DTS
                    grdWHTDASH2.Visible = True
                Catch ex As Exception
                    grdWHTDASH2.Visible = False
                End Try
           
            Else
                grdWHTDASH2.Visible = False
            End If
           
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If
    End Sub

    Sub Load_WHTDASH1()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Status")

        Refresh_WHTDASH1()

        'grdWHTDASH1.DisplayLayout.GroupByBox.Hidden = False
        ' grdWHTDASH1.DisplayLayout.Override.AllowGroupBy = DefaultableBoolean.True
        'Sort_grdColumns(grdWHTDASH1, "LINE_NO")
        With grdWHTDASH1.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .Columns("LINE_CATGY").HiddenWhenGroupBy = DefaultableBoolean.True
            .SortedColumns.Add("LINE_CATGY", False, True)
        End With
        grdWHTDASH1.Rows.ExpandAll(True)

        'grdWHTDASH1.DisplayLayout.Override.AllowGroupBy =DefaultableBoolean.FALSE

        Setup_WHTDASH2()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Initialize_WHTDASH1()
        Dim LINE_CATGY As String = ""
        Dim AT As String = "@ADSIIS"
        If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then
            AT = ""
        End If

        LINE_CATGY = "Sales Orders"
        Add_WHTDASH1("Released", "Transmitted", "WHFSPCK1", "Uu", LINE_CATGY, "U", "Select SOTSHIP1.WHSE_CODE, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE, SOTSHIP1.ORDR_GROUP_NO, SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.PICK_BATCH_NO, SOTSHIP1.SHIP_NOTES, SOTSHIP1.SHIP_SPEC_INST from SOTSHIP1,SOTORDR0 where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO and SOTSHIP1.LP_STATUS = '0' and SOTSHIP1.SHIP_STATUS = 'P' and LP_XMIT_DATE is Null")
        Add_WHTDASH1("Transmitted", "Acknowledged", "", "uT", LINE_CATGY, "T", "Select  SHIP_BOL_NO, WHSE_CODE, CUST_NAME, CUST_CITY, CUST_STATE, ORDR_SHIP_DATE, ORDR_CANCEL_DATE, SHIP_ADDR_TYPE, SHIP_ADDR_CODE, ORDR_GROUP_NO, PICK_BATCH_NO, SHIP_NOTES, SHIP_SPEC_INST, LP_STATUS, LP_XNO from ADS.SOTSHIP1_3PL" & AT & " where LP_STATUS = '0'")
        Add_WHTDASH1("Acknowledged", "Shipped", "", "TT", LINE_CATGY, "P", "Select SOTSHIP1.WHSE_CODE, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE, SOTSHIP1.ORDR_GROUP_NO, SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.PICK_BATCH_NO, SOTSHIP1.SHIP_NOTES, SOTSHIP1.SHIP_SPEC_INST from SOTSHIP1,SOTORDR0 where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO and SOTSHIP1.SHIP_BOL_NO in (Select SHIP_BOL_NO from ADS.SOTSHIP1_3PL" & AT & " where LP_STATUS = '1')")
        Add_WHTDASH1("Modified", "Confirmed", "", "UT", LINE_CATGY, "U", "Select SOTSHIP1.WHSE_CODE, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE, SOTSHIP1.ORDR_GROUP_NO, SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.PICK_BATCH_NO, SOTSHIP1.SHIP_NOTES, SOTSHIP1.SHIP_SPEC_INST from SOTSHIP1,SOTORDR0 where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO and SOTSHIP1.SHIP_BOL_NO in (Select SHIP_BOL_NO from ADS.SOTSHIP3_3PL" & AT & " where LP_STATUS = '0')")
        Add_WHTDASH1("Shipped", "Confirmed", "SOFSHIP0", "TU", LINE_CATGY, "U", "Select SOTSHIP1.WHSE_CODE, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE, SOTSHIP1.ORDR_GROUP_NO, SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.PICK_BATCH_NO, SOTSHIP1.SHIP_NOTES, SOTSHIP1.SHIP_SPEC_INST from SOTSHIP1,SOTORDR0 where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO and SOTSHIP1.SHIP_BOL_NO in (Select SHIP_BOL_NO from ADS.SOTSHIP1_3PL" & AT & " where LP_STATUS in ('V','2'))")
        Add_WHTDASH1("Shipments Deleted", "De-Transmitted", "SORDREL1", "TU", LINE_CATGY, "U", "Select SOTSHIP1.WHSE_CODE, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE, SOTSHIP1.ORDR_GROUP_NO, SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.PICK_BATCH_NO, SOTSHIP1.SHIP_NOTES, SOTSHIP1.SHIP_SPEC_INST from SOTSHIP1,SOTORDR0 where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO and SOTSHIP1.SHIP_BOL_NO in (Select SHIP_BOL_NO from ADS.SOTSHIP1_3PL" & AT & " where LP_STATUS = 'D')")
        Add_WHTDASH1("De-Transmitted", "Re-Transmitted", "WHFSPCK1", "uU", LINE_CATGY, "U", "Select SOTSHIP1.WHSE_CODE, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE, SOTSHIP1.ORDR_GROUP_NO, SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.PICK_BATCH_NO, SOTSHIP1.SHIP_NOTES, SOTSHIP1.SHIP_SPEC_INST from SOTSHIP1,SOTORDR0 where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO and SOTSHIP1.LP_STATUS = '0' and SOTSHIP1.SHIP_STATUS = 'P' and LP_XMIT_DATE is Not Null")

        LINE_CATGY = "PO Shipments"
        Add_WHTDASH1("Inbound Shipments Built", "Transmitted", "POFSHIP1", "Uu", LINE_CATGY, "U", "Select POTSHIP1.PO_SHIPMENT_NO, POTSHIP1.PO_SHIP_VESSEL, POTSHIP1.PO_SHIP_ETA, POTSHIP1.PO_SHIP_ADV_DATE, POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.WHSE_CODE, POTSHIP1.PO_NOTES from POTSHIP1 where POTSHIP1.LP_STATUS = '0'")
        Add_WHTDASH1("Transmitted", "Acknowledged", "", "uT", LINE_CATGY, "T", "Select POTSHIP1.PO_SHIPMENT_NO, POTSHIP1.PO_SHIP_VESSEL, POTSHIP1.PO_SHIP_ETA, POTSHIP1.PO_SHIP_ADV_DATE, POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.WHSE_CODE, POTSHIP1.PO_NOTES from POTSHIP1 where POTSHIP1.PO_SHIPMENT_NO in (Select PO_SHIPMENT_NO from ADS.WHTPORD1" & AT & " where LP_STATUS = '0' and NVL(AIR_SHIP,'0') <> 'T')")
        Add_WHTDASH1("Shipments Acknowledged", "Received", "", "TT", LINE_CATGY, "P", "Select POTSHIP1.PO_SHIPMENT_NO, POTSHIP1.PO_SHIP_VESSEL, POTSHIP1.PO_SHIP_ETA, POTSHIP1.PO_SHIP_ADV_DATE, POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.WHSE_CODE, POTSHIP1.PO_NOTES from POTSHIP1 where POTSHIP1.PO_SHIPMENT_NO in (Select PO_SHIPMENT_NO from ADS.WHTPORD1" & AT & " where LP_STATUS = '1' and NVL(AIR_SHIP,'0') <> 'T' minus Select Distinct PO_SHIPMENT_NO from ADS.RCPTHDR@ADSIIS)")
        Add_WHTDASH1("Container Receipts Closed", "Updated as Received", "POFSHIPR", "TU", LINE_CATGY, "U", "Select POTSHIP1.PO_SHIP_VESSEL, POTSHIP1.PO_SHIP_ADV_DATE, POTSHIP1.PO_SHIP_ETA, POTSHIP1.PO_DATE_SHIPPED, RCPTHDR.* from POTSHIP1, ADS.RCPTHDR" & AT & " where RCPTHDR.STATUS in ('0','V') and RCPTHDR.INVTYP = 'P' and POTSHIP1.PO_SHIPMENT_NO (+) = RCPTHDR.PO_SHIPMENT_NO")

        LINE_CATGY = "Customer Returns"
        Add_WHTDASH1("Customer Returns Received", "Credited", "", "TU", LINE_CATGY, "U", "Select * from ADS.RCPTHDR" & AT & " where RCPTHDR.STATUS in ('0','V') and RCPTHDR.INVTYP = 'R'")

        LINE_CATGY = "Work Orders"
        Add_WHTDASH1("Created", "Transmitted", "", "Uu", LINE_CATGY, "U", "SELECT WKORDER_NO, WKORDER_DESC, CUST_CODE, ORDR_GROUP_NO, ORDR_GROUP_PO, SHIP_DATE, DATE_SENT, CANCEL_DATE, INIT_OPER FROM WOTORDR1 WHERE WKORDER_STATUS = 'P' ORDER BY WKORDER_NO")
        Add_WHTDASH1("Transmitted", "Completed", "", "uT", LINE_CATGY, "P", "SELECT WKORDER_NO, WKORDER_DESC, CUST_CODE, ORDR_GROUP_NO, ORDR_GROUP_PO, SHIP_DATE, DATE_SENT, CANCEL_DATE, INIT_OPER FROM ADS.WOTORDR1@ADSIIS WHERE WKORDER_STATUS in ('0','1','2') ORDER BY WKORDER_NO")
        Add_WHTDASH1("Completed", "Confirmed", "", "TU", LINE_CATGY, "U", "SELECT WKORDER_NO, WKORDER_DESC, CUST_CODE, ORDR_GROUP_NO, ORDR_GROUP_PO, SHIP_DATE, DATE_SENT, CANCEL_DATE, INIT_OPER FROM ADS.WOTORDR1@ADSIIS WHERE WKORDER_STATUS = '3' ORDER BY WKORDER_NO")
        Add_WHTDASH1("Recalled", "Confirmed", "", "UT", LINE_CATGY, "T", "SELECT WKORDER_NO, WKORDER_DESC, CUST_CODE, ORDR_GROUP_NO, ORDR_GROUP_PO, SHIP_DATE, DATE_SENT, CANCEL_DATE, INIT_OPER FROM ADS.WOTORDR1@ADSIIS WHERE WKORDER_STATUS = '4' ORDER BY WKORDER_NO")
        Add_WHTDASH1("Recall Accepted", "Confirmed", "", "TU", LINE_CATGY, "U", "SELECT WKORDER_NO, WKORDER_DESC, CUST_CODE, ORDR_GROUP_NO, ORDR_GROUP_PO, SHIP_DATE, DATE_SENT, CANCEL_DATE, INIT_OPER FROM ADS.WOTORDR1@ADSIIS WHERE WKORDER_STATUS = '5' ORDER BY WKORDER_NO")
        Add_WHTDASH1("Recalled", "Accepted", "", "UT", LINE_CATGY, "P", "SELECT WKORDER_NO, WKORDER_DESC, CUST_CODE, ORDR_GROUP_NO, ORDR_GROUP_PO, SHIP_DATE, DATE_SENT, CANCEL_DATE, INIT_OPER FROM ADS.WOTORDR1@ADSIIS WHERE WKORDER_STATUS = '6' ORDER BY WKORDER_NO")

        LINE_CATGY = "Inventory Adjustments & Warehouse Transfers"
        Add_WHTDASH1("Invty Adjs Submitted", "Loaded", "", "TU", LINE_CATGY, "U", "Select * from ADS.INVADJ" & AT & " where STATUS = '0'")
        'Add_WHTDASH1("Substitution Adjs Submitted", "Processed", "", "TU", LINE_CATGY, "U", "Select * from ADS.INVADJ" & AT & " where STATUS = '0' and NVL(REACOD,'?') not in ('WOK','SHP','RTN')")
        'Add_WHTDASH1("Work Order Adjs Submitted", "Processed", "", "TU", LINE_CATGY, "U", "Select * from ADS.INVADJ" & AT & " where STATUS = '0' and REACOD = 'WRK'")
        'Add_WHTDASH1("Shipping Adjs Submitted", "Processed", "", "TU", LINE_CATGY, "U", "Select * from ADS.INVADJ" & AT & " where STATUS = '0' and REACOD = 'SHP'")
        'Add_WHTDASH1("Returns Adjs Submitted", "Processed", "", "TU", LINE_CATGY, "U", "Select * from ADS.INVADJ" & AT & " where STATUS = '0' and REACOD = 'RTN'")
        Add_WHTDASH1("Invty Adjs Loaded", "Acknowledged", "WHFIADJ1", "UU", LINE_CATGY, "U", "Select * from WHTIADJ1 where ABS_STATUS = 'N'")
        Add_WHTDASH1("Invty Adjs Acknowledged", "Approved", "WHFIADJ1", "UU", LINE_CATGY, "U", "Select * from WHTIADJ1 where ABS_STATUS = 'S'")
        Add_WHTDASH1("Whse Xfrs Submitted", "Acknowledged", "", "UT", LINE_CATGY, "T", "Select SOTSHIP1.WHSE_CODE, SOTINVH1.CUST_CODE, SOTSHIP1.SHIP_DATE_SHIPPED, SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.SHIP_NOTES, SOTSHIP1.SHIP_SPEC_INST from SOTSHIP1,SOTINVH1,ICTIXFR1 where SOTINVH1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO and SOTINVH1.INV_TYPE = 'I' AND SOTINVH1.INV_NO = ICTIXFR1.CTL_NO and ICTIXFR1.XFR_NO in (Select PO_SHIPMENT_NO from ADS.WHTPORD1" & AT & " where LP_STATUS = '0' and NVL(AIR_SHIP,'0') = 'T')")
        Add_WHTDASH1("Whse Xfrs Acknowledged", "Received", "", "TT", LINE_CATGY, "T", "Select SOTSHIP1.WHSE_CODE, SOTINVH1.CUST_CODE, SOTSHIP1.SHIP_DATE_SHIPPED, SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.SHIP_NOTES, SOTSHIP1.SHIP_SPEC_INST from SOTSHIP1,SOTINVH1,ICTIXFR1 where SOTINVH1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO and SOTINVH1.INV_TYPE = 'I' AND SOTINVH1.INV_NO = ICTIXFR1.CTL_NO and ICTIXFR1.XFR_NO in (Select PO_SHIPMENT_NO from ADS.WHTPORD1" & AT & " where LP_STATUS = '1' and NVL(AIR_SHIP,'0') = 'T')")

        txtLINE_WHOS_BALL_U.Text = Replace(txtLINE_WHOS_BALL_U.Text, "XXX", ASCMAIN1.DBS_COMPANY)
        txtLINE_WHOS_BALL_U.Appearance.BackColor = Drawing.Color.LightBlue
        txtLINE_WHOS_BALL_P.Appearance.BackColor = Drawing.Color.Yellow
        txtLINE_WHOS_BALL_T.Appearance.BackColor = Drawing.Color.HotPink
    End Sub

    Sub Add_WHTDASH1( _
        LINE_DESC1 As String, _
        LINE_DESC2 As String, _
        LINE_ACTION As String, _
        UsThem As String, _
        LINE_CATGY As String, _
        LINE_WHOS_BALL As String, _
        LINE_SQL As String)

        Dim Us As String = ASCMAIN1.DBS_COMPANY
        Dim Them As String = "3PL"

        Dim BY_or_TO As New Dictionary(Of String, String)
        BY_or_TO.Add("U", " by " & Us)
        BY_or_TO.Add("u", " to " & Them)
        BY_or_TO.Add("T", " by " & Them)
        BY_or_TO.Add("t", " to " & Us)
        BY_or_TO.Add("", "")

        Dim LINE_DESC As String = LINE_DESC1 & BY_or_TO(Mid(UsThem, 1, 1))
        If LINE_DESC2 <> "" Then
            LINE_DESC &= " Not " & LINE_DESC2 & BY_or_TO(Mid(UsThem, 2, 1))
        End If
        Dim rowWHTDASH1 As DataRow = dst.Tables("WHTDASH1").NewRow
        Dim LINE_NO As Integer = Val(dst.Tables("WHTDASH1").Compute("MAX(LINE_NO)", "") & "") + 1
        rowWHTDASH1.Item("LINE_NO") = LINE_NO
        rowWHTDASH1.Item("LINE_DESC") = LINE_DESC
        rowWHTDASH1.Item("LINE_CATGY") = LINE_CATGY
        Dim LINE_COUNT As Int64 = 0
        rowWHTDASH1.Item("LINE_COUNT") = LINE_COUNT
        rowWHTDASH1.Item("LINE_SQL") = LINE_SQL
        rowWHTDASH1.Item("LINE_WHOS_BALL") = LINE_WHOS_BALL
        rowWHTDASH1.Item("LINE_ACTION") = LINE_ACTION
        dst.Tables("WHTDASH1").Rows.Add(rowWHTDASH1)
    End Sub

    Sub Refresh_WHTDASH1()
        For Each rowWHTDASH1 As DataRow In dst.Tables("WHTDASH1").Select("")
            Dim LINE_SQL As String = rowWHTDASH1.Item("LINE_SQL")
            If LINE_SQL <> "" Then
                Try
                    ASCMAIN1.sql = "Select Count (*) from (" & LINE_SQL & ")"
                    Dim LINE_COUNT As Int64 = Val(ASCDATA1.GetDataValue)
                    rowWHTDASH1.Item("LINE_COUNT") = LINE_COUNT
                Catch ex As Exception
                    connection_is_down = True
                    Exit For
                End Try

            End If
        Next
        Dim DTS As String = " as of " & Format(Now, "MM/dd/yyyy HH:mm")
        If connection_is_down Then
            DTS &= " - Connection Is Down"
        End If
        grdWHTDASH1.Text = "Status Counts" & DTS
    End Sub

    Sub Initialize_WHTCOLN1()
        With dst.Tables("WHTCOLN1").Rows
            .Add("ABS_COMMENT", "ABS Comment", 100, "")
            .Add("ABS_STATUS", "Sta", 30, "")
            .Add("ADJ_REF1", "Ref1", 60, "")
            .Add("ADJ_REF2", "Ref2", 60, "")
            .Add("ADJQTY", "Qty", 60, "")
            .Add("ARRDTE", "Arrived", 85, "MM/dd/yy")
            .Add("CANCEL_DATE", "Cancel Date", 85, "MM/dd/yy")
            .Add("CLOSE_DTE", "Dt Closed", 90, "")
            .Add("CONTAINER_NO", "Container", 130, "")
            .Add("CUST_CITY", "City", 80, "")
            .Add("CUST_CODE", "Customer", 100, "")
            .Add("CUST_NAME", "Customer Name", 110, "")
            .Add("CUST_STATE", "Customer", 60, "")
            .Add("DATE_SENT", "Sent", 85, "MM/dd/yy")
            .Add("INIT_OPER", "By", 60, "")
            .Add("INVTYP", "Typ", 40, "")
            .Add("ITEM_CODE", "Item", 60, "")
            .Add("LP_CODE", "LP", 60, "")
            .Add("LP_STATUS_TS_3PL", "TS ADS", 100, "MM/dd HH:mm")
            .Add("LP_STATUS_TS_ERP", "TS VAN", 100, "MM/dd HH:mm")
            .Add("LP_XNO", "XNo", 60, "")
            .Add("ORDR_CANCEL_DATE", "Cancel", 85, "MM/dd/yy")
            .Add("ORDR_CUST_PO", "Customer PO", 120, "")
            .Add("ORDR_GROUP_NO", "Group", 120, "")
            .Add("ORDR_GROUP_PO", "Group PO", 120, "")
            .Add("ORDR_SHIP_DATE", "Ship", 85, "MM/dd/yy")
            .Add("PICK_BATCH_NO", "Batch", 70, "")
            .Add("PO_DATE_SHIPPED", "Shipped", 85, "MM/dd/yy")
            .Add("PO_NOTES", "PO Notes", 70, "")
            .Add("PO_SHIP_ADV_DATE", "Advised", 85, "MM/dd/yy")
            .Add("PO_SHIP_ETA", "ETA", 85, "MM/dd/yy")
            .Add("PO_SHIPMENT_NO", "Ship No", 70, "")
            .Add("PO_SHIP_VESSEL", "Vessel", 120, "")
            .Add("REACOD", "Reason", 60, "")
            .Add("REF1", "Ref1", 60, "")
            .Add("REF2", "Ref2", 60, "")
            .Add("STATUS", "LP Sta", 60, "")
            .Add("SHIP_ADDR_TYPE", "ST", 40, "")
            .Add("SHIP_ADDR_CODE", "ShipTo", 70, "")
            .Add("SHIP_BOL_NO", "Shipment", 110, "")
            .Add("SHIP_DATE", "Ship Date", 85, "MM/dd/yy")
            .Add("SHIP_DATE_SHIPPED", "Shipped", 85, "MM/dd/yy")
            .Add("SHIP_NOTES", "Shipment Notes", 80, "")
            .Add("SHIP_SPEC_INST", "Special Instructions", 80, "")
            .Add("TRANS_SEQ", "Trans No", 60, "")
            .Add("TRNDTE", "Date", 60, "")
            .Add("WHSE_CODE", "Whs", 40, "")
            .Add("WKORDER_NO", "Wk Order #", 120, "")
            .Add("WKORDER_DESC", "Description", 200, "")
        End With
    End Sub

    Private Sub grdWHTDASH1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTDASH1.InitializeRow

        Dim LINE_WHOS_BALL As String = e.Row.Cells("LINE_WHOS_BALL").Value
        Dim LINE_COUNT As Int64 = Val(e.Row.Cells("LINE_COUNT").Value & "")
        If LINE_COUNT <> 0 Then
            Select Case LINE_WHOS_BALL
                Case "U"
                    e.Row.Cells("LINE_COUNT").Appearance.BackColor = Drawing.Color.LightBlue
                Case "T"
                    e.Row.Cells("LINE_COUNT").Appearance.BackColor = Drawing.Color.HotPink
                Case "P"
                    e.Row.Cells("LINE_COUNT").Appearance.BackColor = Drawing.Color.Yellow
                Case Else
                    e.Row.Cells("LINE_COUNT").Appearance.BackColor = Drawing.Color.Empty
            End Select
        Else
            e.Row.Cells("LINE_COUNT").Appearance.BackColor = Drawing.Color.Empty
        End If

        Dim LINE_ACTION As String = e.Row.Cells("LINE_ACTION").Value & ""
        Select Case LINE_ACTION
            Case "WHFSPCK1"
                e.Row.Cells("LINE_ACTION_BTN").Value = "Xmit 3PL"
            Case "SOFSHIP0"
                e.Row.Cells("LINE_ACTION_BTN").Value = "Billing"
            Case "SORDREL1"
                e.Row.Cells("LINE_ACTION_BTN").Value = "De-Release"
            Case "POFSHIP1"
                e.Row.Cells("LINE_ACTION_BTN").Value = "PO Shipments"
            Case "POFSHIPR"
                e.Row.Cells("LINE_ACTION_BTN").Value = "PO Receipts"
            Case "WHFIADJ1"
                e.Row.Cells("LINE_ACTION_BTN").Value = "Post Adj"
        End Select

    End Sub

    Private Sub grdWHTDASH1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdWHTDASH1.InitializeLayout

    End Sub

    Private Sub grdWHTDASH1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTDASH1.ClickCellButton
        Dim LINE_ACTION As String = e.Cell.Row.Cells("LINE_ACTION").Value & ""
        Select Case LINE_ACTION
            Case "POFSHIP1"
                Context_Launch("", "", "", "POFSHIP1", "F", "POE")
            Case "POFSHIPR"
                Context_Launch("", "", "", "POFSHIPR", "F", "POE")
            Case "WHFSPCK1"
                Context_Launch("", "", "", "WHFSPCK1")
            Case "SOFSHIP0"
                Context_Launch("", "", "", "SOFSHIP0")
            Case "SORDREL1"
                Context_Launch("", "", "", "SORDREL1")
            Case "WHFIADJ1"
                Context_Launch("", "", "", "WHFIADJ1")
        End Select
    End Sub
End Class