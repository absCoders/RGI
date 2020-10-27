
Imports System.Text
Imports Microsoft.Office.Interop.Word

Public Class ECFSZIO1
    Dim S As New System.Text.StringBuilder() With {.Length = 0}
    Dim isFormLoading As Boolean = True
    Dim COL_MAP As New Dictionary(Of String, Type)
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MAKE_COL_MAP()

        With dst

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("WD.STYLE_CODE,")
            S.AppendLine("WD.COLOR_CODE")
            'S.AppendLine("(WD.STYLE_CODE || '-' || WD.COLOR_CODE) AS ITEMID")
            S.AppendLine("FROM WBTSTYLD WD, ICTSTYL1 S1, ICTSTAT2 S2")
            S.AppendLine("WHERE WD.STYLE_CODE = S1.STYLE_CODE")
            S.AppendLine("AND WD.STYLE_CODE = S2.STYLE_CODE (+)")
            S.AppendLine("AND WD.COLOR_CODE = S2.COLOR_CODE (+)")
            S.AppendLine("AND")
            S.AppendLine("(")
            S.AppendLine("  S1.STYLE_STATUS = 'A'")
            S.AppendLine("  OR")
            S.AppendLine("  (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_TRAN,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) - NVL(S2.WHSE_QTY_OPEN,0)) > 0")
            S.AppendLine(")")
            S.AppendLine("AND S2.WHSE_CODE (+) = 'MS'")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ECTSZIO1", "**", 0, False)
            Create_TDA(.Tables.Add, "ECTSZIO2", "**", 0, False)
            For Each KVP As KeyValuePair(Of String, Type) In COL_MAP
                .Tables("ECTSZIO1").Columns.Add(KVP.Key.ToUpper, KVP.Value)
            Next
            .Tables("ECTSZIO2").Columns.Add("ItemID", GetType(System.String))
            .Tables("ECTSZIO2").Columns.Add("Qty", GetType(System.Int64))
            .Tables("ECTSZIO2").Columns.Add("PriceLevel0", GetType(System.Double))
            .Tables("ECTSZIO2").Columns.Add("IsDeleted", GetType(System.String))

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM WBTSTYLH")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "WBTSTYLH", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ICTSTYL1")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ICTSTYC1")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ICTSTYC1", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT S2.STYLE_CODE,")
            S.AppendLine("S2.COLOR_CODE,")
            S.AppendLine("(NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0) + NVL(WHSE_QTY_TRAN,0) + NVL(WHSE_QTY_ON_ORDER,0) - NVL(WHSE_QTY_OPEN,0)) AVAIL")
            S.AppendLine("FROM ICTSTAT2 S2")
            S.AppendLine("WHERE S2.WHSE_CODE = 'MS'")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ICTSTAT2", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ICTSTYLD")
            S.AppendLine(" WHERE PACK_CODE = 'ITM'")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ICTSTYLD", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ICTSIZE1")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ICTSIZE1", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT WD.STYLE_CODE, WH.PAGE_NAME")
            S.AppendLine("FROM WBTPAGEH WH, WBTPAGED WD")
            S.AppendLine("WHERE WH.PAGE_CODE = WD.PAGE_CODE")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "WBTPAGEX", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM TATCNTRY")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "TATCNTRY", "**", 0, False)
        End With

        'Fill_Records("ECTSZIO1")
        Fill_Records("ICTSTYL1")
        Fill_Records("ICTSTYC1")
        Fill_Records("ICTSTAT2")
        Fill_Records("ICTSTYLD")
        Fill_Records("ICTSIZE1")
        Fill_Records("TATCNTRY")
        Fill_Records("WBTPAGEX")

        MAKE_GRID_COLS()

        'Fill_COLS()

        grdECTSZIO1.DataSource = dst.Tables("ECTSZIO1")
        grdECTSZIO2.DataSource = dst.Tables("ECTSZIO2")

        Create_Summary(grdECTSZIO1, "ItemID", "Count", "", "###,##0")
        Create_Summary(grdECTSZIO2, "ItemID", "Count", "", "###,##0")

        'ASCMAIN1.Add_Value_List(grdSOTQRDR1, "CALC_STATUS", , New String() {":", "I:Imported From Web", "L:Pulled To Laptop", "O:Finalized As Order", "X:Deleted", "M:Marked Complete", "T:Testing"})
        Sort_grdColumns(grdECTSZIO1, "ItemID", False)
        Sort_grdColumns(grdECTSZIO2, "ItemID, Qty", False)

        With grdECTSZIO1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        For i As Integer = 0 To grdECTSZIO1.DisplayLayout.Bands(0).Columns.Count - 1
            grdECTSZIO1.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        With grdECTSZIO2.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        For i As Integer = 0 To grdECTSZIO2.DisplayLayout.Bands(0).Columns.Count - 1
            grdECTSZIO2.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        Load_Record()

        tab.Visible = False
        isFormLoading = False

    End Sub

    Private Sub MAKE_GRID_COLS()
        grdECTSZIO1.DisplayLayout.Bands(0).Columns.Item("STYLE_CODE").Hidden = True
        grdECTSZIO1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE").Hidden = True
        For Each KVP As KeyValuePair(Of String, Type) In COL_MAP
            With grdECTSZIO1.DisplayLayout.Bands(0)
                .Columns.Add(KVP.Key.ToUpper, KVP.Key)
            End With
        Next
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        EMsg = ""
        Select Case eItemKey
            Case "Refresh"

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
            Case "Refresh"
                'Load_Record()
                RefreshData()
                BuildPricing()
            Case "Exit"
                Call Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Refresh").Visible = True
                .Groups("Screen Control").Items("Exit").Visible = True
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
    End Sub

    Sub Clear_Record()
        'dst.Tables("SOTQRDR1").Rows.Clear()
    End Sub

    Sub Load_Record()
        'Call Save_Header_Fields(UltraGroupBox1)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Refreshing Data", "")

        EnforceConstraints(False)

        'Fill_Records("SOTQRDR1")

        EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        'Update_Record_TDA("SOTQRDR1")
        Call CommitTrans("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    'Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
    '    Print_Report_Begin()
    '    'frm.CR_params.Add("SUBT", "")
    '    'Fill SOTORDRP records
    '    Fill_Records("SOTQRDR5", ORDR_NO, True)
    '    For Each rowSOTQRDR1 As DataRow In dst.Tables("SOTQRDR1").Select()
    '        If rowSOTQRDR1.Item("ORDR_NO") = ORDR_NO Then
    '            rowSOTQRDR1.Item("ERRORS") = "NEW"
    '        Else
    '            rowSOTQRDR1.Item("ERRORS") = ""
    '        End If
    '    Next
    '    'Generate_Report("SORQRDRO")
    '    Generate_Report("WBRWEBQT", "Quotes Imported From Web", "Re-printed From Quote Maint.")
    '    '    Print_Report_End()
    'End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdECTSZIO1, "SS", "Show Filter", "Show GroupBox", "Release Quote For Re-import", "Mark Quote As Complete", "Mark Quote As Testing", "Re-Assign Quote To New Order", "Delete Quote")
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

        If grd.Selected.Rows.Count = 0 Then
            MsgBox("You Must Select One And Only One Row First", vbOKOnly, "Select A Row")
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Something"
                'grd.ActiveRow.Cells.Item("ORDR_NO_WEB").Value = ""
        End Select

        Update_Record()
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
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

    Private Sub BuildPricing()
        ASCMAIN1.Progress("Building Prices", "")
        dst.Tables.Item("ECTSZIO2").Clear()
        Dim Discounts As List(Of DISCOUNTS) = Nothing
        Dim STYLE_CODE_LAST As String = ""
        Dim rowARTCUST1 As DataRow = Nothing
        For Each rowECTSZIO1 As DataRow In dst.Tables("ECTSZIO1").Select("", "STYLE_CODE, COLOR_CODE")
            Dim STYLE_CODE As String = rowECTSZIO1.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowECTSZIO1.Item("COLOR_CODE").ToString & String.Empty
            Dim ITEM_CODE As String = String.Format("{0}-{1}", STYLE_CODE, COLOR_CODE)
            ASCMAIN1.Progress("-", ITEM_CODE)
            If STYLE_CODE <> STYLE_CODE_LAST Then
                STYLE_CODE_LAST = STYLE_CODE
                Discounts = SOCMAIN2.Price_Discounts(Me, "", rowARTCUST1, STYLE_CODE, False)
            End If
            For i As Int64 = 3 To 0 Step -1
                If Discounts(i).DISCOUNT_QTY > 0 Then
                    Dim rowECTSZIO2 As DataRow = dst.Tables.Item("ECTSZIO2").NewRow
                    rowECTSZIO2.Item("STYLE_CODE") = STYLE_CODE
                    rowECTSZIO2.Item("COLOR_CODE") = COLOR_CODE
                    rowECTSZIO2.Item("ItemID") = ITEM_CODE
                    rowECTSZIO2.Item("Qty") = Discounts(i).DISCOUNT_QTY
                    rowECTSZIO2.Item("PriceLevel0") = Discounts(i).DISCOUNT_PRICE
                    rowECTSZIO2.Item("IsDeleted") = "False"
                    dst.Tables.Item("ECTSZIO2").Rows.Add(rowECTSZIO2)
                End If
            Next
        Next
        ASCMAIN1.Progress("", "")
    End Sub
    Private Sub RefreshData()
        ASCMAIN1.Progress("Refreshing Styles", "")

        Fill_Records("ECTSZIO1")

        For Each rowECTSZIO1 As DataRow In dst.Tables("ECTSZIO1").Select("", "STYLE_CODE, COLOR_CODE")
            Dim STYLE_CODE As String = rowECTSZIO1.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowECTSZIO1.Item("COLOR_CODE").ToString & String.Empty
            Dim ITEM_CODE As String = String.Format("{0}-{1}", STYLE_CODE, COLOR_CODE)
            ASCMAIN1.Progress("-", ITEM_CODE)

            Dim SFilter As String = String.Format("STYLE_CODE='{0}'", STYLE_CODE)
            Dim SCFilter As String = String.Format("STYLE_CODE='{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)

            Dim rowWBTSTYLH As DataRow = dst.Tables.Item("WBTSTYLH").Select(SFilter).FirstOrDefault
            Dim rowICTSTYL1 As DataRow = dst.Tables.Item("ICTSTYL1").Select(SFilter).FirstOrDefault
            Dim rowICTSTYC1 As DataRow = dst.Tables.Item("ICTSTYC1").Select(SCFilter).FirstOrDefault
            Dim rowICTSTAT2 As DataRow = dst.Tables.Item("ICTSTAT2").Select(SCFilter).FirstOrDefault
            Dim rowICTSTYLD As DataRow = dst.Tables.Item("ICTSTYLD").Select(SFilter).FirstOrDefault

            Dim BEST_DESC As String = rowICTSTYL1.Item("STYLE_DESC").ToString & String.Empty
            If Not IsNothing(rowWBTSTYLH) Then
                If (rowWBTSTYLH.Item("STYLE_DESC_SHORT").ToString & String.Empty).Length > 0 Then
                    BEST_DESC = rowWBTSTYLH.Item("STYLE_DESC_SHORT").ToString & String.Empty
                End If
                If (rowWBTSTYLH.Item("STYLE_DESC_LONG").ToString & String.Empty).Length > 0 Then
                    BEST_DESC = rowWBTSTYLH.Item("STYLE_DESC_LONG").ToString & String.Empty
                End If
            End If

            For Each KVP As KeyValuePair(Of String, Type) In COL_MAP
                Select Case KVP.Key
                    Case "ItemID"
                        rowECTSZIO1.Item(KVP.Key.ToUpper) = ITEM_CODE
                    Case "ItemName"
                        rowECTSZIO1.Item(KVP.Key.ToUpper) = rowICTSTYL1.Item("STYLE_DESC").ToString & String.Empty
                    Case "Description"
                        rowECTSZIO1.Item(KVP.Key.ToUpper) = BEST_DESC
                    Case "Dimensions"
                        If Not IsNothing(rowICTSTYLD) Then
                            Dim LENGTH As Double = Val(rowICTSTYLD.Item("LENGTH").ToString & String.Empty)
                            Dim WIDTH As Double = Val(rowICTSTYLD.Item("WIDTH").ToString & String.Empty)
                            Dim HEIGHT As Double = Val(rowICTSTYLD.Item("HEIGHT").ToString & String.Empty)
                            If Not (HEIGHT = 0 Or WIDTH = 0 Or HEIGHT = 0) Then
                                rowECTSZIO1.Item(KVP.Key.ToUpper) = String.Format("{0} X {1} X {2}", LENGTH, WIDTH, HEIGHT)
                            Else
                                Dim SIZE_CODE As String = rowICTSTYL1.Item("SIZE_CODE").ToString & String.Empty
                                Dim rowICTSIZE1 As DataRow = dst.Tables.Item("ICTSIZE1").Select(String.Format("SIZE_CODE = '{0}'", SIZE_CODE)).FirstOrDefault
                                If Not IsNothing(rowICTSIZE1) Then
                                    Dim SIZE_DESC As String = rowICTSIZE1.Item("SIZE_DESC").ToString & String.Empty
                                    If SIZE_DESC.Length > 0 Then
                                        rowECTSZIO1.Item(KVP.Key.ToUpper) = SIZE_DESC
                                    End If
                                End If

                            End If
                        End If
                    Case "DimensionsMetric"
                        Dim SIZE_CODE As String = rowICTSTYL1.Item("SIZE_CODE").ToString & String.Empty
                        If SIZE_CODE.Length > 2 Then
                            rowECTSZIO1.Item(KVP.Key.ToUpper) = SIZE_CODE.Substring(SIZE_CODE.Length - 2, 2)
                        End If
                    Case "Category"
                        Dim CATG As String = ""
                        For Each rowWBTPAGEX As DataRow In dst.Tables("WBTPAGEX").Select(SFilter)
                            CATG = CATG & " | " & rowWBTPAGEX.Item("PAGE_NAME").ToString & String.Empty
                        Next
                        If CATG.Length > 3 Then
                            CATG = CATG.Substring(3, CATG.Length - 3)
                        End If
                        rowECTSZIO1.Item(KVP.Key.ToUpper) = CATG
                    Case "OrderMinimumQuantity"
                        If Val(rowICTSTYL1.Item("STYLE_SO_QTY_MIN").ToString & String.Empty) > 0 Then
                            rowECTSZIO1.Item(KVP.Key.ToUpper) = Val(rowICTSTYL1.Item("STYLE_SO_QTY_MIN").ToString & String.Empty)
                        Else
                            rowECTSZIO1.Item(KVP.Key.ToUpper) = 1
                        End If
                    Case "OrderMultipleQuantity"
                        If Val(rowICTSTYL1.Item("STYLE_ASST_QTY").ToString & String.Empty) > 0 Then
                            rowECTSZIO1.Item(KVP.Key.ToUpper) = Val(rowICTSTYL1.Item("STYLE_ASST_QTY").ToString & String.Empty)
                        Else
                            rowECTSZIO1.Item(KVP.Key.ToUpper) = 1
                        End If
                    Case "OnHandQuantity"
                        If IsNothing(rowICTSTAT2) Then
                            rowECTSZIO1.Item(KVP.Key.ToUpper) = 0
                        Else
                            If Val(rowICTSTAT2.Item("AVAIL").ToString & String.Empty) > 0 Then
                                rowECTSZIO1.Item(KVP.Key.ToUpper) = Val(rowICTSTAT2.Item("AVAIL").ToString & String.Empty)
                            Else
                                rowECTSZIO1.Item(KVP.Key.ToUpper) = 0
                            End If
                        End If
                    Case "UPC"
                        rowECTSZIO1.Item(KVP.Key.ToUpper) = rowICTSTYC1.Item("UPC_CODE").ToString & String.Empty
                    Case "BasePrice"
                        If Val(rowICTSTYL1.Item("STYLE_PRICE").ToString & String.Empty) > 0 Then
                            rowECTSZIO1.Item(KVP.Key.ToUpper) = Val(rowICTSTYL1.Item("STYLE_PRICE").ToString & String.Empty)
                        Else
                            rowECTSZIO1.Item(KVP.Key.ToUpper) = 0
                        End If
                    Case "Cubes"
                        If Val(rowICTSTYL1.Item("CASE_CUBE").ToString & String.Empty) > 0 Then
                            rowECTSZIO1.Item(KVP.Key.ToUpper) = Val(rowICTSTYL1.Item("CASE_CUBE").ToString & String.Empty)
                        Else
                            rowECTSZIO1.Item(KVP.Key.ToUpper) = 0
                        End If
                    Case "Source"
                        Dim COUNTRY_CODE As String = rowICTSTYL1.Item("COUNTRY_CODE").ToString & String.Empty
                        If COUNTRY_CODE.Length = 0 Then
                            COUNTRY_CODE = "CHN"
                        End If
                        Dim ctFilter As String = String.Format("COUNTRY_CODE = '{0}'", COUNTRY_CODE)
                        Dim rowTATCNTRY As DataRow = dst.Tables.Item("TATCNTRY").Select(ctFilter).FirstOrDefault
                        If Not IsNothing(rowTATCNTRY) Then
                            rowECTSZIO1.Item(KVP.Key.ToUpper) = rowTATCNTRY.Item("COUNTRY_NAME").ToString & String.Empty
                        Else
                            rowECTSZIO1.Item(KVP.Key.ToUpper) = "CHINA"
                        End If
                    Case "UnitOfMeasure"
                        rowECTSZIO1.Item(KVP.Key.ToUpper) = rowICTSTYL1.Item("STYLE_UOM").ToString & String.Empty
                    Case "Weight"
                        If Val(rowICTSTYL1.Item("STYLE_WEIGHT").ToString & String.Empty) <> 0 Then
                            rowECTSZIO1.Item(KVP.Key.ToUpper) = Val(rowICTSTYL1.Item("STYLE_WEIGHT").ToString & String.Empty)
                        End If
                    Case "AdditionalImageCount"
                        rowECTSZIO1.Item(KVP.Key.ToUpper) = 0
                    Case "PhotoName"
                        rowECTSZIO1.Item(KVP.Key.ToUpper) = ITEM_CODE
                    Case "IsDeleted"
                        rowECTSZIO1.Item(KVP.Key.ToUpper) = "False"
                    Case Else 'All Fields Not Specified Above Get Left Blank.
                End Select
            Next
            'FILL PRICELEVELS HERE
        Next
        ASCMAIN1.Progress("", "")
        'grdECTSZIO1.DisplayLayout.AutoFitStyle = UltraWinGrid.AutoFitStyle.ResizeAllColumns
    End Sub
    Private Sub MAKE_COL_MAP()
        COL_MAP.Add("ItemID", GetType(System.String))
        COL_MAP.Add("ItemName", GetType(System.String))
        COL_MAP.Add("Description", GetType(System.String))
        COL_MAP.Add("Notes", GetType(System.String))
        COL_MAP.Add("Dimensions", GetType(System.String))
        COL_MAP.Add("DimensionsMetric", GetType(System.String))
        COL_MAP.Add("Category", GetType(System.String))
        COL_MAP.Add("OrderMinimumQuantity", GetType(System.Int64))
        COL_MAP.Add("OrderMultipleQuantity", GetType(System.Int64))
        COL_MAP.Add("OnHandQuantity", GetType(System.Int64))
        COL_MAP.Add("InventoryStatus", GetType(System.String))
        COL_MAP.Add("ReportCategory", GetType(System.String))
        COL_MAP.Add("IntroDate", GetType(System.String))
        COL_MAP.Add("UPC", GetType(System.String))
        COL_MAP.Add("BasePrice", GetType(System.Double))
        For I As Integer = 1 To 20
            COL_MAP.Add(String.Format("PriceLevel{0}", I), GetType(System.Double))
        Next
        COL_MAP.Add("SpecialPrice", GetType(System.Double))
        COL_MAP.Add("PieceBox", GetType(System.Int64))
        COL_MAP.Add("Cubes", GetType(System.Double))
        COL_MAP.Add("Source", GetType(System.String))
        COL_MAP.Add("ContainerMinQty", GetType(System.Int64))
        COL_MAP.Add("UnitOfMeasure", GetType(System.String))
        COL_MAP.Add("Weight", GetType(System.Double))
        For I As Integer = 1 To 20
            COL_MAP.Add(String.Format("UDF{0}",I), GetType(System.String))
        Next
        COL_MAP.Add("AdditionalImageCount", GetType(System.Int64))
        COL_MAP.Add("AdditionalPhotos", GetType(System.String))
        COL_MAP.Add("PhotoName", GetType(System.String))
        COL_MAP.Add("CatalogCode", GetType(System.String))
        COL_MAP.Add("CatalogName", GetType(System.String))
        COL_MAP.Add("showChildFor", GetType(System.String))
        COL_MAP.Add("showRelatedFor", GetType(System.String))
        COL_MAP.Add("IsDeleted", GetType(System.String))
    End Sub
#End Region

#Region "Form Controls"

#Region "Grids"

#End Region
#End Region

End Class