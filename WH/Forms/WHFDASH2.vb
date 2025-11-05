
Imports System.Text
Imports System.IO
Imports System.Drawing
Imports Infragistics.Win.UltraWinGrid

Public Class WHFDASH2
    Dim S As New System.Text.StringBuilder() With {.Length = 0}
    Dim isFormLoading As Boolean = True
    Dim TTM As New UltraWinToolTip.UltraToolTipManager
    Dim IMAGES_FOLDER_HIGH As String = ""
    Dim IMAGES_FOLDER_LOW As String = ""
    Dim IMAGE_DEFAULT As String = ""


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        With dst

            ASCMAIN1.sql = "SELECT COUNT(DISTINCT CART_NO)CARTONS, COUNT(1) SCANS, 
                            nvl(SUM(CASE WHEN DISPOSITION = 'Y' THEN 1 ELSE 0 END),0) CARTONSGOOD,
                            nvl(SUM(CASE WHEN (VARIANCE > 0 AND DISPOSITION = 'N') THEN 1 ELSE 0 END),0) CARTONSOVER, 
                            nvl(SUM(CASE WHEN (VARIANCE < 0 AND DISPOSITION = 'N') THEN 1 ELSE 0 END),0) CARTONSSHORT,
                            nvl(SUM(CASE WHEN (VARIANCE = 0 AND DISPOSITION = 'N') THEN 1 ELSE 0 END),0) WRONGUPC,
                            nvl(SUM(SCANNED),0) SCANQTY, nvl(SUM(REQUESTED),0) REQUESTQTY
                            FROM (
                            SELECT WHTRFID2.SCAN_NO, WHTRFID1.CART_NO, WHTRFID1.DISPOSITION, SUM(WHTRFID2.SCAN_QTY) SCANNED, 
                            SUM(WHTRFID2.PICK_QTY) REQUESTED, SUM(WHTRFID2.SCAN_QTY - WHTRFID2.PICK_QTY) VARIANCE, COUNT(1) UPCS
                            FROM WHTRFID1, WHTRFID2
                            WHERE WHTRFID2.SCAN_NO =  WHTRFID1.SCAN_NO
                            AND WHTRFID1.SCAN_DATE > trunc(sysdate)
                            GROUP BY WHTRFID2.SCAN_NO, WHTRFID1.CART_NO,  WHTRFID1.DISPOSITION)"
            Create_TDA(.Tables.Add, "WHTRFIDS", "**", 0, False)
            Fill_Records("WHTRFIDS")

            ASCMAIN1.sql = "SELECT WHTRFID1.SCAN_NO, WHTRFID1.CART_NO, WHTRFID1.DISPOSITION, WHTRFID1.DISPOSITION_REASON,  
                            SUM(WHTRFID2.SCAN_QTY) SCANNED, SUM(WHTRFID2.PICK_QTY) REQUESTED, 
                            SUM(WHTRFID2.SCAN_QTY - WHTRFID2.PICK_QTY) VARIANCE, COUNT(1) UPCS
                            FROM WHTRFID1, WHTRFID2
                            WHERE WHTRFID2.SCAN_NO =  WHTRFID1.SCAN_NO
                            AND WHTRFID1.SCAN_DATE > TRUNC(sysdate)
                            GROUP BY WHTRFID1.SCAN_NO, WHTRFID1.CART_NO, WHTRFID1.DISPOSITION, WHTRFID1.DISPOSITION_REASON"
            Create_TDA(.Tables.Add, "WHTRFID1", "**", 0, False)
            Fill_Records("WHTRFID1")

            ASCMAIN1.sql = "select WHTRFID2.UPC_CODE, MIN(WHTSCSEQ.STYLE_SEQ) STYLE_SEQ, MIN(WHTLOCM1.LOCATION_ZONE) LOCATION_ZONE,
                            MIN(ICVLUPC1.STYLE_CODE) STYLE_CODE, min(ICVLUPC1.COLOR_CODE) COLOR_CODE,
                            sum(WHTRFID2.SCAN_QTY) Scanned, sum(WHTRFID2.PICK_QTY) Requested, 
                            sum(WHTRFID2.SCAN_QTY - WHTRFID2.PICK_QTY) variance, count(1) Cartons, 
                            sum( case when (WHTRFID2.SCAN_QTY - WHTRFID2.PICK_QTY) <> 0 then 1 else 0 end) Errors
                            from WHTRFID1, WHTRFID2, ICVLUPC1, WHTSCSEQ, WHTLOCM1
                            where WHTRFID2.SCAN_NO =  WHTRFID1.SCAN_NO
                            and WHTRFID2.UPC_CODE = ICVLUPC1.UPC_CODE
                            and WHTSCSEQ.CUST_CODE = 'WALMART'
                            and WHTSCSEQ.STYLE_CODE = ICVLUPC1.STYLE_CODE
                            and WHTSCSEQ.COLOR_CODE = ICVLUPC1.COLOR_CODE
                            and WHTSCSEQ.STYLE_SEQ = WHTLOCM1.LOCATION_ROUTE_SEQ
                            and WHTLOCM1.WHSE_CODE = 'NJC'
                            and WHTLOCM1.LOCATION_CODE like 'F1%'
                            and WHTRFID1.scan_date > TRUNC(sysdate)
                            group by WHTRFID2.UPC_CODE
                            having sum(WHTRFID2.SCAN_QTY - WHTRFID2.PICK_QTY) <> 0"
            Create_TDA(.Tables.Add, "WHTRFID2", "**", 0, False)
            Fill_Records("WHTRFID2")

            ASCMAIN1.sql = "select WHTRFID3.RFID, WHTRFID3.UPC_CODE, MIN(WHTSCSEQ.STYLE_SEQ) STYLE_SEQ, MIN(WHTLOCM1.LOCATION_ZONE) LOCATION_ZONE,
                            MIN(ICVLUPC1.STYLE_CODE) STYLE_CODE, min(ICVLUPC1.COLOR_CODE) COLOR_CODE, count(distinct WHTRFID1.CART_NO) Cartons
                            from WHTRFID1, WHTRFID3, ICVLUPC1, WHTSCSEQ, WHTLOCM1
                            where WHTRFID3.SCAN_NO =  WHTRFID1.SCAN_NO
                            and WHTRFID3.UPC_CODE = ICVLUPC1.UPC_CODE
                            and WHTSCSEQ.CUST_CODE = 'WALMART'
                            and WHTSCSEQ.STYLE_CODE = ICVLUPC1.STYLE_CODE
                            and WHTSCSEQ.COLOR_CODE = ICVLUPC1.COLOR_CODE
                            and WHTSCSEQ.STYLE_SEQ = WHTLOCM1.LOCATION_ROUTE_SEQ
                            and WHTLOCM1.WHSE_CODE = 'NJC'
                            and WHTLOCM1.LOCATION_CODE like 'F1%'
                            and WHTRFID1.scan_date > TRUNC(sysdate)
                            group by WHTRFID3.RFID, WHTRFID3.UPC_CODE
                            having count(distinct WHTRFID1.CART_NO) > 1"
            Create_TDA(.Tables.Add, "WHTRFID3", "**", 0, False)
            Fill_Records("WHTRFID3")

        End With

        grdWHTRFID1.DataSource = dst.Tables("WHTRFID1")
        grdWHTRFID2.DataSource = dst.Tables("WHTRFID2")
        grdWHTRFID3.DataSource = dst.Tables("WHTRFID3")

        Create_Summary(grdWHTRFID3, "STYLE_CODE", "Count", "", "###,##0")
        Create_Summary(grdWHTRFID2, "STYLE_CODE", "Count", "", "###,##0")
        Create_Summary(grdWHTRFID1, "SCAN_NO", "Count", "", "###,##0")


        Sort_grdColumns(grdWHTRFID1, "scan_no, CART_NO", False)
        Sort_grdColumns(grdWHTRFID2, "errors", False)
        Sort_grdColumns(grdWHTRFID3, "cartons", False)

        'grdWHTRFID3.DisplayLayout.Bands(0).Columns("AVAIL").Format = "###,##0"

        With grdWHTRFID1.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Color.White
                GCOL.Header.Appearance.BackColor2 = Color.Gray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                GCOL.CellActivation = Activation.NoEdit

                If New String() {"SCANNED", "REQUESTED", "VARIANCE", "UPCS"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.Orange
                    GCOL.Format = "####,##0"
                Else
                    GCOL.Header.Appearance.BackColor2 = Color.LightGreen
                End If
            Next
        End With

        With grdWHTRFID2.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Color.White
                GCOL.Header.Appearance.BackColor2 = Color.Gray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                GCOL.CellActivation = Activation.NoEdit

                If New String() {"SCANNED", "REQUESTED", "VARIANCE", "CARTONS", "ERRORS", "STYLE_SEQ"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.Orange
                    If GCOL.Key = "STYLE_SEQ" Then
                        GCOL.Header.Appearance.BackColor2 = Color.LightGreen
                    End If
                    GCOL.Format = "####,##0"
                Else
                    GCOL.Header.Appearance.BackColor2 = Color.LightGreen
                End If
            Next
        End With

        With grdWHTRFID3.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Color.White
                GCOL.Header.Appearance.BackColor2 = Color.Gray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                GCOL.CellActivation = Activation.NoEdit

                If New String() {"CARTONS", "STYLE_SEQ"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.Orange
                    If GCOL.Key = "STYLE_SEQ" Then
                        GCOL.Header.Appearance.BackColor2 = Color.LightGreen
                    End If
                    GCOL.Format = "####,##0"
                Else
                    GCOL.Header.Appearance.BackColor2 = Color.LightGreen
                End If
            Next
        End With

        With grdWHTRFID3.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        With grdWHTRFID2.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        Load_Record()

        tab.Visible = False
        isFormLoading = False

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
        RefreshData()

        EnforceConstraints(True)

        'If EntryMode = "N" Then
        'Else
        '    dst.AcceptChanges()
        'End If

        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        'Update_Record_TDA("XXXXXXX")
        Call CommitTrans("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
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
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTRFID1, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdWHTRFID2, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdWHTRFID3, "SS", "Show Filter", "Show GroupBox")
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

        'If grd.Selected.Rows.Count = 0 Then
        '    MsgBox("You Must Select One And Only One Row First", vbOKOnly, "Select A Row")
        '    Exit Sub
        'End If

        Select Case e.Tool.Key
            Case "View Image"
                'Dim STYLE_CODE As String = grd.ActiveRow.Cells.Item("STYLE_CODE").Value
                'Dim COLOR_CODE As String = grd.ActiveRow.Cells.Item("COLOR_CODE").Value
                'Dim frmIMAGE As New TAC.TAFIMGV1(Me, STYLE_CODE, COLOR_CODE, "M")
                'With frmIMAGE
                '    .ShowDialog(Me)
                'End With
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
    Private Sub RefreshData()
        ASCMAIN1.Progress("Refreshing Statistics", "")

        Fill_Records("WHTRFIDS")
        Fill_Records("WHTRFID1")
        Fill_Records("WHTRFID2")
        Fill_Records("WHTRFID3")

        lstWHTRFIDS.View = View.Details
        If lstWHTRFIDS.Columns.Count = 0 Then
            lstWHTRFIDS.Columns.Add("Statistic", 120)
            lstWHTRFIDS.Columns.Add("Value", 80)
        End If

        Dim rowSTATS As DataRow = dst.Tables("WHTRFIDS").Rows(0)
        lstWHTRFIDS.Items.Clear()

        For Each col As DataColumn In dst.Tables("WHTRFIDS").Columns
            Dim item As New ListViewItem(FormatColumnText(col.ColumnName))
            item.SubItems.Add(rowSTATS(col.ColumnName))
            lstWHTRFIDS.Items.Add(item)
        Next

        lblUpdated.Text = "Analyzed at " & DateTime.Now.ToString("g")
        grdWHTRFID1.Text = "Cartons Scanned "
        grdWHTRFID2.Text = "UPCs with Errors "
        grdWHTRFID3.Text = "RFIDs with multiple scans "

        'MsgBox("Analysis Complete", vbOKOnly, "Done")
        ASCMAIN1.Progress("", "")
    End Sub
    Private Function FormatColumnText(input As String) As String
        ' Replace underscores with spaces
        Dim spaced As String = input.Replace("_", " ")

        ' Split into words
        Dim words() As String = spaced.Split(" "c)

        ' Convert each word to Title Case
        For i As Integer = 0 To words.Length - 1
            If words(i).Length > 0 Then
                words(i) = Char.ToUpper(words(i)(0)) & words(i).Substring(1).ToLower()
            End If
        Next

        ' Join the words back into a single string
        Return String.Join(" ", words)
    End Function

    Private Sub grdWBTIMGLT_MouseHover(sender As Object, e As EventArgs) Handles grdWHTRFID3.MouseHover
        If 1 = 1 Then

        End If
    End Sub

    Private Sub grdWHTRFID3_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdWHTRFID3.InitializeLayout

    End Sub
#End Region

#Region "Form Controls"

#Region "Grids"

#End Region
#End Region

End Class