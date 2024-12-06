Imports System.Text
Imports Infragistics.Win.UltraWinGrid

Public Class ICFPROM1
    Private SQL As New StringBuilder With {.Length = 0}
    Private rowICTPROMX As DataRow
    Private PROMO_CTL_NO As String = ""
    Private FormLoading As Boolean = True

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        FormLoading = True

        If MENU_ITEM_OBJECT = "ICTPROM1" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")

        With dst

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("P1.PROMO_CTL_NO,")
            SQL.AppendLine("P1.PROMO_DESC,")
            SQL.AppendLine("P1.PROMO_START_DATE,")
            SQL.AppendLine("P1.PROMO_END_DATE,")
            SQL.AppendLine("COUNT(P2.STYLE_CODE) AS STYLE_CNT")
            SQL.AppendLine("FROM ICTPROM1 P1, ICTPROM2 P2")
            SQL.AppendLine("WHERE P1.PROMO_CTL_NO = P2.PROMO_CTL_NO (+)")
            SQL.AppendLine("GROUP BY")
            SQL.AppendLine("P1.PROMO_CTL_NO,")
            SQL.AppendLine("P1.PROMO_DESC,")
            SQL.AppendLine("P1.PROMO_START_DATE,")
            SQL.AppendLine("P1.PROMO_END_DATE")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ICTPROMX", "**", 0, False)
            'With .Tables("ICTPROMX").Columns
            '    .Add("SEL", GetType(System.String))
            'End With

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("P2.PROMO_CTL_NO,")
            SQL.AppendLine("P2.STYLE_CODE,")
            SQL.AppendLine("P2.PROMO_STYLE_NOTES,")
            SQL.AppendLine("S1.STYLE_DESC,")
            SQL.AppendLine("P2.PROMO_UNIT_PRICE")
            SQL.AppendLine("FROM ICTPROM2 P2, ICTSTYL1 S1")
            SQL.AppendLine("WHERE P2.STYLE_CODE = S1.STYLE_CODE")
            SQL.AppendLine("AND P2.PROMO_CTL_NO = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ICTPROM2", "**", 0, True, "V", 2)

            SQL.Length = 0
            SQL.AppendLine("SELECT ECOM_CODE, ECOM_UNIT_PRICE")
            SQL.AppendLine("FROM ECTESTY1")
            SQL.AppendLine("WHERE (NVL(SHIP_ECOM,'0') = '1' OR NVL(SHIP_DROP,'0') = '1')")
            SQL.AppendLine("AND STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTESTY1", "**", 0, False, "V")
        End With

        grdICTPROMX.DataSource = dst.Tables("ICTPROMX")
        grdICTPROM2.DataSource = dst.Tables("ICTPROM2")
        grdECTECOM1.DataSource = dst.Tables("ECTESTY1")

        Create_Summary(grdICTPROMX, "PROMO_CTL_NO", "Count")

        Sort_grdColumns(grdICTPROMX, "PROMO_CTL_NO", False)
        Sort_grdColumns(grdICTPROM2, "STYLE_CODE", False)
        Sort_grdColumns(grdECTECOM1, "ECOM_CODE", False)

        Fill_Records("ICTPROMX")

        'Bind_Controls(pnlICTSTYL1, "ICTSTYL1")

        With grdICTPROMX.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("STYLE_CNT").Format = "###,##0"
            '.Columns("SEL").CellActivation = UltraWinGrid.Activation.AllowEdit
        End With

        With grdICTPROM2.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.True
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("PROMO_UNIT_PRICE").Format = "###,##0.00"
            .Columns("PROMO_UNIT_PRICE").CellActivation = UltraWinGrid.Activation.AllowEdit
            .Columns("PROMO_STYLE_NOTES").CellActivation = UltraWinGrid.Activation.AllowEdit
        End With

        With grdECTECOM1.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("ECOM_UNIT_PRICE").Format = "###,##0.00"
        End With

        'ASCMAIN1.Add_Value_List(grdX, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")

        'Set_Read_Only(grpTotals, True)

        'grpHeader.Visible = False
        txtPROMO_START_DATE.Value = ""
        txtPROMO_END_DATE.Value = ""
        dteUpdateEndDate.Value = ""
        FormLoading = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                'Validate_Code("STYLE_CODE")
                If Absx1.txtFor("PROMO_CTL_NO").Text & "" <> "" Then
                    EMsg &= vbCr & "Promo # Must Be Blank"
                End If

                If Absx1.txtFor("PROMO_DESC").Value & "" = "" Then
                    EMsg &= vbCr & "Missing Promo Description"
                End If

                If Not IsDate(txtPROMO_START_DATE.Value) Or Not IsDate(txtPROMO_END_DATE.Value) Then
                    EMsg &= vbCr & "Missing Start/End Date"
                Else
                    If txtPROMO_END_DATE.Value < txtPROMO_START_DATE.Value Then
                        EMsg &= vbCr & "End Date Can Not Be Before Start Date"
                    End If
                End If

            Case "View"
                If Not ASCMAIN1.Logical_Lock("ICFPROM1", Absx1.txtFor("PROMO_CTL_NO").Text) Then Exit Sub
            Case "Update"

            Case "Delete"
                If MsgBox("Are You Sure You Want To Delete This Promo?", MsgBoxStyle.YesNo,
                          "Make Sure!!") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If

            Case "Cancel"
                If MsgBox("OK To Lose Changes?", MsgBoxStyle.YesNo,
                          "You May Have Pending Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Upload Spreadsheet"
                Dim msg As New System.Text.StringBuilder With {.Length = 0}
                msg.AppendLine("This Will Prompt You For An Excel Spreadsheet.")
                msg.AppendLine("")
                msg.AppendLine("It Should Be In The Same Format As If You")
                msg.AppendLine("Downloaded It From This Grid!!")
                msg.AppendLine("")
                msg.AppendLine("Once Selected,It Will Verify Then Replace All")
                msg.AppendLine("Of The Styles On This Promo With The Styles,")
                msg.AppendLine("Colors and Pricing Listed On Your Spreadsheet.")
                msg.AppendLine("")
                msg.AppendLine("Are You Ready For This?")
                If MsgBox(msg.ToString, MsgBoxStyle.YesNo,
                          "Upload Spreadsheet") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)
                RefreshSummary()
            Case "Delete"
                Delete_Promo()
                Mode_Settings(False)
                RefreshSummary()
            Case "Cancel"
                Mode_Settings(False)
                RefreshSummary()
            Case "Add Style"
                AddStyle()

            Case "Upload Spreadsheet"
                UploadSpreadsheet()
            Case "Change Description"
                Dim frmASFMSGBF As New ASFMSGBF
                'UpDatesReqNotes = True
                Dim UpDatesNotes As String = frmASFMSGBF.Get_txtblock_from_User("New Description", "Change Description", txtPROMO_DESC.Text, False, 75)
                If UpDatesNotes = "" Or UpDatesNotes = txtPROMO_DESC.Text Then
                    MsgBox("Blank Or No Change Detected", vbCritical, "No Change")
                Else
                    txtPROMO_DESC.Text = UpDatesNotes
                End If
        End Select

    End Sub

    Private Sub Delete_Promo()
        BeginTrans()
        Dim PROMO_CTL_NO As String = Absx1.txtFor("PROMO_CTL_NO").Text.ToString & String.Empty
        If PROMO_CTL_NO.Length > 0 Then

            Dim BATCH_NO As String = ASCMAIN1.Next_Control_No("ICTPROM2_U.BATCH_NO")
            SQL.Length = 0
            SQL.AppendLine("INSERT INTO ICTPROM2_U")
            SQL.AppendLine(String.Format("SELECT '{0}' AS BATCH_NO, ICTPROM2.* FROM ICTPROM2", BATCH_NO))
            SQL.AppendLine(String.Format("WHERE PROMO_CTL_NO = {0}", PROMO_CTL_NO))
            ASCMAIN1.sql = SQL.ToString
            ASCDATA1.ExecuteSQL()

            For Each DELTABLE As String In New String() {"ICTPROM1", "ICTPROM2"}
                SQL.Length = 0
                SQL.AppendLine(String.Format("DELETE FROM {0} WHERE PROMO_CTL_NO = '{1}'", DELTABLE, PROMO_CTL_NO))
                ASCMAIN1.sql = SQL.ToString
                ASCDATA1.ExecuteSQL()
            Next
        Else
            MsgBox("Invalid Promo Code", vbOKOnly, "Promo Deletion Error")
            Rollback()
            Exit Sub
        End If
        CommitTrans()
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = ScreenMode
                    .Items("View").Settings.Enabled = ScreenMode
                    .Items("Update").Settings.Enabled = Not ScreenMode
                    .Items("Delete").Settings.Enabled = Not ScreenMode
                    .Items("Cancel").Settings.Enabled = Not ScreenMode
                    .Items("Add Style").Settings.Enabled = Not ScreenMode
                    .Items("Upload Spreadsheet").Settings.Enabled = Not ScreenMode
                    .Items("Change Description").Settings.Enabled = Not ScreenMode
                End With
                .Groups("Pricing").Visible = ScreenMode
                .Groups("Options").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        'SplitContainer1.Visible = ScreenMode

        Panel1.Visible = Not ScreenMode
        'UltraTabControl1.Visible = Not ScreenMode

        txtPROMO_START_DATE.Enabled = True
        txtPROMO_END_DATE.Enabled = True

        If ScreenMode Then

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        Dim clearTables As New List(Of String)
        clearTables.Add("ICTPROM2")
        clearTables.Add("ECTESTY1")

        For Each clearTable As String In clearTables
            dst.Tables(clearTable).Rows.Clear()
        Next

        EnforceConstraints(True)

        txtPROMO_START_DATE.Value = ""
        txtPROMO_END_DATE.Value = ""
        Absx1.txtFor("PROMO_CTL_NO").Text = ""
        Absx1.txtFor("PROMO_DESC").Text = ""
        PROMO_CTL_NO = ""
        numSTYLE_RETAIL.Value = 0
        numSTYLE_PRICE.Value = 0
        numSTYLE_PROMO_PRICE.Value = 0

        imgSTYL1.Visible = False
        For i As Integer = 1 To 4
            Absx1.CtlFor(String.Format("lblDISC{0}", i)).Visible = False
            Absx1.CtlFor(String.Format("lblDISC{0}QP", i)).Visible = False
            Absx1.txtFor(String.Format("qtyDISC{0}", i)).Visible = False
            Absx1.txtFor(String.Format("priceDISC{0}", i)).Visible = False
            Absx1.CtlFor(String.Format("lblDISC{0}", i)).Text = ""
            Absx1.CtlFor(String.Format("lblDISC{0}", i)).Tag = ""
            Absx1.txtFor(String.Format("qtyDISC{0}", i)).Text = ""
            Absx1.txtFor(String.Format("priceDISC{0}", i)).Text = ""
        Next

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        PROMO_CTL_NO = Absx1.txtFor("PROMO_CTL_NO").Text & ""

        Dim loadTables As New List(Of String)
        loadTables.Add("ICTPROM2")

        If EntryMode = "N" Then
            Dim PROMO_CTL_NO As String = ASCMAIN1.Next_Control_No("ICTPROM1.PROMO_CTL_NO")

            Absx1.txtFor("PROMO_CTL_NO").Text = PROMO_CTL_NO
            Fill_Records("ICTPROM2", PROMO_CTL_NO)
        Else
            For Each loadTable As String In loadTables
                Fill_Records(loadTable, PROMO_CTL_NO)
            Next
        End If

        'init_Misc()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        'This is stupid but I spent too much time fighting the standards.
        Dim PROMO_CTL_NO As String = Absx1.txtFor("PROMO_CTL_NO").Text
        Dim PROMO_DESC As String = Absx1.txtFor("PROMO_DESC").Text
        SQL.Length = 0
        SQL.AppendLine("SELECT COUNT(*) AS RECCNT")
        SQL.AppendLine("FROM ICTPROM1")
        SQL.AppendLine(String.Format("WHERE PROMO_CTL_NO = '{0}'", PROMO_CTL_NO))
        ASCMAIN1.sql = SQL.ToString()
        Dim RECCNT As Integer = Val(ASCDATA1.GetDataValue)
        If RECCNT = 0 Then
            SQL.Length = 0
            SQL.AppendLine("INSERT INTO ICTPROM1")
            SQL.AppendLine("VALUES(")
            SQL.AppendLine(String.Format("'{0}',", PROMO_CTL_NO))
            SQL.AppendLine(String.Format("'{0}',", PROMO_DESC))
            SQL.AppendLine(String.Format("'{0}',", Format(txtPROMO_START_DATE.DateTime, "dd-MMM-yy")))
            SQL.AppendLine(String.Format("'{0}'", Format(txtPROMO_END_DATE.DateTime, "dd-MMM-yy")))
            SQL.AppendLine(")")
            ASCMAIN1.sql = SQL.ToString
            ASCDATA1.ExecuteSQL()
        Else
            SQL.Length = 0
            SQL.AppendLine("UPDATE ICTPROM1")
            SQL.AppendLine(String.Format("SET PROMO_END_DATE = '{0}'", Format(txtPROMO_END_DATE.DateTime, "dd-MMM-yy")))
            SQL.AppendLine(String.Format(", PROMO_DESC = '{0}'", txtPROMO_DESC.Text))
            SQL.AppendLine(String.Format("WHERE PROMO_CTL_NO = '{0}'", PROMO_CTL_NO))
            ASCMAIN1.sql = SQL.ToString
            ASCDATA1.ExecuteSQL()
        End If

        'Dim updateTables As New List(Of String)
        'updateTables.Add("ECTESTY3")

        BeginTrans()
        Update_Record_TDA("ICTPROM2", String.Format("PROMO_CTL_NO = '{0}'", Absx1.txtFor("PROMO_CTL_NO").Text.ToString & String.Empty))
        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        Dim deleteTables As New List(Of String)
        'deleteTables.Add("")

        BeginTrans()
        For Each deleteTable As String In deleteTables
            'Delete_Records(deleteTable)
        Next

        CommitTrans("Delete Complete")
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTPROMX, "SS", "Show Filter", "Show GroupBox")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'Select Case e.SourceControl.Name
            '    Case = "grdECTECOM1_PARTNER"
            'End Select

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

        Select Case e.Tool.Key
            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"
    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "STYLE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode Then
                        Click_Command("View", e)
                    End If
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "STYLE_CODE"
                If Not InquiryMode Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "SOME_COLUMN"
        End Select
    End Sub
#End Region

#Region "Form Controls"
    Private Sub btnUpdateEndDate_Click(sender As Object, e As EventArgs) Handles btnUpdateEndDate.Click
        If Not IsDate(dteUpdateEndDate.Value) Then
            MsgBox("You Must Pick A Valid Date", vbOKOnly, "Invalid Date")
        Else
            If txtPROMO_END_DATE.Value < txtPROMO_START_DATE.Value Then
                EMsg &= vbCr & "End Date Can Not Be Before Start Date"
            Else
                txtPROMO_END_DATE.Value = dteUpdateEndDate.Value
                'For Each rowICTPROM1 As DataRow In dst.Tables("ICTPROM1").Select()
                '    rowICTPROM1.Item("PROMO_END_DATE") = dteUpdateEndDate.Value
                'Next
            End If
        End If
    End Sub

    Private Sub btnMarkDownPct_Click(sender As Object, e As EventArgs) Handles btnMarkDownPct.Click
        If numMarkDownPct.Value <= 0 Then
            MsgBox("Markdown Pct Must Be Greater Than 0", vbOKOnly, "Check Pct")
        Else
            Dim MDP As Double = ((100 - numMarkDownPct.Value) / 100)
            For Each rowICTPROM2 As DataRow In dst.Tables("ICTPROM2").Select()
                If chkOffListPrice.Checked Then
                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowICTPROM2.Item("STYLE_CODE").ToString)
                    If Not IsNothing(rowICTSTYL1) Then
                        Dim STYLE_PRICE As Double = Val(rowICTSTYL1.Item("STYLE_PRICE").ToString & String.Empty)
                        If STYLE_PRICE > 0 Then
                            rowICTPROM2.Item("PROMO_UNIT_PRICE") = Val(STYLE_PRICE) * MDP
                        End If
                    End If
                Else
                    rowICTPROM2.Item("PROMO_UNIT_PRICE") = Val(rowICTPROM2.Item("PROMO_UNIT_PRICE")) * MDP
                End If
            Next
        End If
    End Sub

#End Region

#Region "Custom Methods"
    Private Sub AddStyle()
        'ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STYLE_CODE")

        SQL.Length = 0
        SQL.AppendLine("SELECT")
        SQL.AppendLine("S1.STYLE_CODE,")
        SQL.AppendLine("S1.STYLE_DESC,")
        SQL.AppendLine("S1.STYLE_STATUS,")
        SQL.AppendLine("S1.STYLE_PRICE LIST_PRICE")
        SQL.AppendLine("FROM ICTSTYL1 S1")
        With ASCMAIN1.CodeSelector
            .SQL = SQL.ToString
            .MultipleSelections = True
            .PreviouslySelectedCodes0 = ""
            .Caption = "Select Style To Add"
            .TABLE_NAME = ""
            .VIEW_NAME = ""
            .VIEW_DESC = ""
            .COLUMN_NAME = ""
            .COLUMN_PREKEYs = New Dictionary(Of String, String)
            .Custom_sql_where = ""
            .tblASTVIEW1 = New DataTable
        End With
        'Dim F As New ASFCODE1
        'F.ShowDialog()

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            Dim STYLE_CODEs As String = ""

            '  ASCMAIN1.CodeSelector.UseDataFromTable = tbl
            Dim FM As New ASFCODE1
            FM.ShowDialog()
            FM.Dispose()
            If ASCMAIN1.CodeSelector.Selections = 0 Then
                MsgBox("Nothing Selected", vbOKOnly, "Please Select Something")
            Else
                Dim ERR_MSG1 As New StringBuilder With {.Length = 0}
                Dim ERR_MSG2 As New StringBuilder With {.Length = 0}
                For i As Integer = 0 To ASCMAIN1.CodeSelector.Selections - 1
                    Dim STYLE_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(i).Item("STYLE_CODE") & String.Empty
                    Dim Filter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
                    If dst.Tables.Item("ICTPROM2").Select(Filter).Count <> 0 Then
                        ERR_MSG1.AppendLine(String.Format("{0}", STYLE_CODE))
                    Else
                        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                        SQLS.AppendLine("Select P2.*,")
                        SQLS.AppendLine("P1.PROMO_START_DATE,")
                        SQLS.AppendLine("P1.PROMO_END_DATE")
                        SQLS.AppendLine("From ICTPROM1 P1, ICTPROM2 P2")
                        SQLS.AppendLine("Where P1.PROMO_CTL_NO = P2.PROMO_CTL_NO")
                        SQLS.AppendLine(String.Format("AND P2.STYLE_CODE = '{0}'", STYLE_CODE))
                        SQLS.AppendLine(String.Format("AND (P1.PROMO_START_DATE BETWEEN '{0}' AND '{1}'", Format(txtPROMO_START_DATE.DateTime, "dd-MMM-yy"), Format(txtPROMO_END_DATE.DateTime, "dd-MMM-yy")))
                        SQLS.AppendLine(String.Format("Or P1.PROMO_END_DATE BETWEEN '{0}' AND '{1}'", Format(txtPROMO_START_DATE.DateTime, "dd-MMM-yy"), Format(txtPROMO_END_DATE.DateTime, "dd-MMM-yy")))
                        SQLS.AppendLine(String.Format("Or (P1.PROMO_START_DATE < '{0}' AND '{1}' < '{1}' ))", Format(txtPROMO_START_DATE.DateTime, "dd-MMM-yy"), Format(txtPROMO_END_DATE.DateTime, "dd-MMM-yy")))

                        Dim tblICTPROM2 As DataTable = ASCDATA1.GetDataTable(SQLS.ToString(), String.Empty)
                        If tblICTPROM2.Rows.Count > 0 Then
                            For Each rowICTPROM2 As DataRow In tblICTPROM2.Rows
                                Dim PROMO_CTL_NO As String = rowICTPROM2.Item("PROMO_CTL_NO").ToString & String.Empty
                                Dim SD As String = "No Date"
                                If IsDate(rowICTPROM2.Item("PROMO_START_DATE").ToString & String.Empty) Then
                                    SD = CDate(rowICTPROM2.Item("PROMO_START_DATE").ToString & String.Empty).ToShortDateString
                                End If
                                Dim ED As String = "No Date"
                                If IsDate(rowICTPROM2.Item("PROMO_END_DATE").ToString & String.Empty) Then
                                    ED = CDate(rowICTPROM2.Item("PROMO_END_DATE").ToString & String.Empty).ToShortDateString
                                End If
                                ERR_MSG2.AppendLine(String.Format("{0} | {1} | {2} | {3}", PROMO_CTL_NO, SD, ED, STYLE_CODE))
                            Next
                        Else
                            Dim rowICTPROM2 As DataRow = dst.Tables.Item("ICTPROM2").NewRow
                            rowICTPROM2.Item("PROMO_CTL_NO") = Absx1.txtFor("PROMO_CTL_NO").Text & String.Empty
                            rowICTPROM2.Item("STYLE_CODE") = STYLE_CODE
                            rowICTPROM2.Item("STYLE_DESC") = ASCMAIN1.CodeSelector.SelectedRows(i).Item("STYLE_DESC") & String.Empty
                            rowICTPROM2.Item("PROMO_UNIT_PRICE") = 0
                            dst.Tables.Item("ICTPROM2").Rows.Add(rowICTPROM2)
                        End If
                    End If
                Next
                grdICTPROM2.Refresh()
                If ERR_MSG1.Length > 0 Then
                    MsgBox(ERR_MSG1.ToString, vbOKOnly, "Styles Already On This Promo")
                End If
                If ERR_MSG2.Length > 0 Then
                    Dim emsg As String = "Promo | Start | End | Style" & vbCrLf & ERR_MSG2.ToString
                    MsgBox(emsg, vbOKOnly, "Styles Conflicting With Other Promos")
                End If
            End If
        End If
    End Sub

    Private Function GetImageLocation(ByVal STYLE_CODE As String) As String
        Dim RetVal As String = ""
        Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
        Dim RO_PARM_STYLE_IMG_DIR As String = ""
        Dim FileMatch As String
        'Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
        'Dim COLOR_CODE_LONG As String = ""
        'If Not IsNothing(rowICTCOLR1) Then
        '    COLOR_CODE_LONG = rowICTCOLR1.Item("COLOR_CODE_LONG").ToString()
        'End If
        Dim WebVal As String = ""
        If Not IsNothing(rowSOTPARM3) Then
            RO_PARM_STYLE_IMG_DIR = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString
            FileMatch = Dir(String.Format("{0}\{1}*", RO_PARM_STYLE_IMG_DIR, STYLE_CODE))
            If FileMatch.Length > 0 Then
                RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
            End If
            'We just get the first color now.
            'If RO_PARM_STYLE_IMG_DIR.Length > 0 Then
            '    FileMatch = Dir(String.Format("{0}\{1}-{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE))
            '    If FileMatch.Length > 0 Then
            '        RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
            '    Else
            '        FileMatch = Dir(String.Format("{0}\{1}{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE))
            '        If FileMatch.Length > 0 Then
            '            RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
            '        Else
            '            FileMatch = Dir(String.Format("{0}\{1}{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE_LONG))
            '            If FileMatch.Length > 0 Then
            '                RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
            '            Else
            '                FileMatch = Dir(String.Format("{0}\{1}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE))
            '                If FileMatch.Length > 0 Then
            '                    RetVal = String.Format("{0}\{1}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE)
            '                Else
            '                    FileMatch = Dir(String.Format("{0}\{1}*", RO_PARM_STYLE_IMG_DIR, STYLE_CODE))
            '                    If FileMatch.Length > 0 Then
            '                        RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
            '                    End If
            '                End If
            '            End If
            '        End If
            '    End If
            'End If
        End If
        'Try
        '    If WebVal.Length > 0 Then
        '        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(WebVal)
        '        Dim response As System.Net.WebResponse = req.GetResponse()
        '        Dim stream As IO.Stream = response.GetResponseStream()
        '        Dim img As System.Drawing.Image = System.Drawing.Image.FromStream(stream)
        '        stream.Close()
        '        If System.IO.File.Exists(RetVal) Then
        '            System.IO.File.Delete(RetVal)
        '            img.Save(RetVal)
        '        Else
        '            RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, String.Format("{0}-{1}.jpg", STYLE_CODE, COLOR_CODE))
        '            img.Save(RetVal)
        '        End If
        '    End If
        'Catch ex As Exception
        'End Try
        Return RetVal
    End Function

    Private Sub getStyleInfo()
        If Not IsNothing(grdICTPROM2.ActiveRow) Then
            Dim STYLE_CODE As String = grdICTPROM2.ActiveRow.Cells.Item("STYLE_CODE").Text & String.Empty
            'Dim COLOR_CODE As String = grdICTPROM2.ActiveRow.Cells.Item("STYLE_CODE").Text & String.Empty
            imgSTYL1.ImageLocation = GetImageLocation(STYLE_CODE)
            If imgSTYL1.ImageLocation.Length > 0 Then
                imgSTYL1.Visible = True
            Else
                imgSTYL1.Visible = False
            End If

            Dim rowARTCUST1 As DataRow = Nothing
            Dim Discounts As List(Of DISCOUNTS)
            Discounts = SOCMAIN2.Price_Discounts(Me, "", rowARTCUST1, STYLE_CODE, False)
            For i As Integer = 1 To 4
                If Discounts(i - 1).DISCOUNT_QTY = 0 Then
                    Absx1.CtlFor(String.Format("lblDISC{0}", i)).Visible = False
                    Absx1.CtlFor(String.Format("lblDISC{0}QP", i)).Visible = False
                    Absx1.txtFor(String.Format("qtyDISC{0}", i)).Visible = False
                    Absx1.txtFor(String.Format("priceDISC{0}", i)).Visible = False
                Else
                    Absx1.CtlFor(String.Format("lblDISC{0}", i)).Visible = True
                    Absx1.CtlFor(String.Format("lblDISC{0}QP", i)).Visible = True
                    Absx1.txtFor(String.Format("qtyDISC{0}", i)).Visible = True
                    Absx1.txtFor(String.Format("priceDISC{0}", i)).Visible = True
                    Absx1.CtlFor(String.Format("lblDISC{0}", i)).Text = Discounts(i - 1).DISCOUNT_DESC
                    Absx1.CtlFor(String.Format("lblDISC{0}", i)).Tag = Discounts(i - 1).DISCOUNT_PCT 'Use for hover over.
                    Absx1.txtFor(String.Format("qtyDISC{0}", i)).Text = Discounts(i - 1).DISCOUNT_QTY
                    Absx1.txtFor(String.Format("priceDISC{0}", i)).Text = Format$(Discounts(i - 1).DISCOUNT_PRICE, "###,##0.00")
                End If
            Next

            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            If Not IsNothing(rowICTSTYL1) Then
                numSTYLE_RETAIL.Value = Val(rowICTSTYL1.Item("STYLE_RETAIL").ToString & String.Empty)
                numSTYLE_PRICE.Value = Val(rowICTSTYL1.Item("STYLE_PRICE").ToString & String.Empty)
                numSTYLE_PROMO_PRICE.Value = Val(rowICTSTYL1.Item("STYLE_PROMO_PRICE").ToString & String.Empty)
            End If

            Fill_Record("ECTESTY1", STYLE_CODE)

        End If
    End Sub

    Private Sub RefreshSummary()
        Fill_Records("ICTPROMX")
    End Sub

    Private Sub UploadSpreadsheet()
        Dim fileToImport As String = String.Empty
        Dim UpsertData As New DataTable
        Dim FirstRow As Int64
        Dim LastRow As Int64
        Dim thisRow As Int64
        Dim iResult As MsgBoxResult
        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
        Dim PROBLEM_STYLES As New System.Text.StringBuilder With {.Length = 0}

        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Open File To Upsert"
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx"
            openFileDialog1.FilterIndex = 1
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                fileToImport = openFileDialog1.FileName
            End If
            openFileDialog1.Dispose()
        End Using

        If fileToImport.Length = 0 Then
            Exit Sub
        End If

        ASCMAIN1.Progress("Reading File")
        Me.Cursor = Cursors.WaitCursor

        Using cn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileToImport & ";Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""")
            Using cmd As New System.Data.OleDb.OleDbDataAdapter("select * from [Styles On Promotion$]", cn)
                ' Select the data from Sheet1 of the workbook.
                cn.Open()
                cmd.Fill(UpsertData)
                cn.Close()
                cmd.Dispose()
            End Using
            cn.Dispose()
        End Using

        For Each rowUPSERT As DataRow In UpsertData.Select()
            If thisRow > 2000 Then
                Exit For
            End If
            If rowUPSERT.Item(0).ToString & String.Empty = "Style" Then
                FirstRow = thisRow
                If rowUPSERT.Item(3).ToString & String.Empty <> "Promo Price" Then
                    iMSG.Length = 0
                    iMSG.AppendLine("Spreadsheet Is Not In The Correct Format")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, "Problems With File.")
                    Exit Sub
                End If
            End If
            If FirstRow > 0 And rowUPSERT.Item(0).ToString & String.Empty = "" Then
                LastRow = thisRow - 1
                Exit For
            End If
            LastRow = thisRow
            thisRow += 1
        Next
        If FirstRow > 0 Then
            BeginTrans()
            Dim BATCH_NO As String = ASCMAIN1.Next_Control_No("ICTPROM2_U.BATCH_NO")
            Dim PROMO_CTL_NO As String = Absx1.txtFor("PROMO_CTL_NO").Text.ToString & String.Empty
            SQL.Length = 0
            SQL.AppendLine("INSERT INTO ICTPROM2_U")
            SQL.AppendLine(String.Format("SELECT '{0}' AS BATCH_NO, ICTPROM2.* FROM ICTPROM2", BATCH_NO))
            SQL.AppendLine(String.Format("WHERE PROMO_CTL_NO = {0}", PROMO_CTL_NO))
            ASCMAIN1.sql = SQL.ToString
            ASCDATA1.ExecuteSQL()
            dst.Tables.Item("ICTPROM2").Clear()

            For i As Int64 = FirstRow + 1 To LastRow
                Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                Dim HAS_PROBLEM As Boolean = False
                Dim STYLE_CODE As String = UpsertData.Rows(i).Item(0).ToString.ToUpper
                Dim STYLE_DESC As String = ""
                Dim PROMO_STYLE_NOTES As String = UpsertData.Rows(i).Item(2).ToString.ToUpper
                Dim PROMO_UNIT_PRICE As String = UpsertData.Rows(i).Item(3).ToString.ToUpper

                'SQLS.Length = 0
                'SQLS.AppendLine("SELECT STYLE_DESC")
                'SQLS.AppendLine("FROM ICTSTYL1")
                'SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                'ASCMAIN1.sql = SQLS.ToString()
                'STYLE_DESC = ASCDATA1.GetDataValue
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If IsNothing(rowICTSTYL1) Then
                    PROBLEM_STYLES.AppendLine(String.Format("Style {0} Not Found In Masterfile", STYLE_CODE))
                    HAS_PROBLEM = True
                Else
                    STYLE_DESC = rowICTSTYL1.Item("STYLE_DESC").ToString & String.Empty
                End If
                If Not IsNumeric(PROMO_UNIT_PRICE) Then
                    PROBLEM_STYLES.AppendLine(String.Format("{0} Has Invalid Promo Price", STYLE_CODE))
                    HAS_PROBLEM = True
                End If
                If HAS_PROBLEM = False Then
                    Dim newICTPROM2 As DataRow = dst.Tables("ICTPROM2").NewRow
                    newICTPROM2.Item("PROMO_CTL_NO") = Absx1.txtFor("PROMO_CTL_NO").Text.ToString & String.Empty
                    newICTPROM2.Item("STYLE_CODE") = STYLE_CODE
                    newICTPROM2.Item("PROMO_STYLE_NOTES") = PROMO_STYLE_NOTES
                    newICTPROM2.Item("PROMO_UNIT_PRICE") = Val(PROMO_UNIT_PRICE)
                    newICTPROM2.Item("STYLE_DESC") = STYLE_DESC
                    dst.Tables("ICTPROM2").Rows.Add(newICTPROM2)
                End If
            Next
            'Update_Record_TDA("ECTESTY3", String.Format("ECOM_PROMO_CTL_NO = '{0}'", Absx1.txtFor("ECOM_PROMO_CTL_NO").Text.ToString & String.Empty))
            iMSG.Length = 0
            iMSG.AppendLine("Promo Backed Up / Cleared")
            iMSG.AppendLine("And Replaced With Data From Spreadsheet")
            If PROBLEM_STYLES.Length > 0 Then
                iMSG.AppendLine("")
                iMSG.AppendLine("However The Following Problems Were Found:")
                iMSG.AppendLine(PROBLEM_STYLES.ToString)
            End If
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, "Please review.")
            CommitTrans()
        Else
            iMSG.Length = 0
            iMSG.AppendLine("No Matching Records Founds To Upload.")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, "Problems With File.")
        End If
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
#End Region

#Region "grdICTPROMX"
    Private Sub grdICTPROMX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTPROMX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("PROMO_CTL_NO").Text = e.Row.Cells("PROMO_CTL_NO").Text
            Absx1.txtFor("PROMO_DESC").Text = e.Row.Cells("PROMO_DESC").Text
            If IsDate(e.Row.Cells("PROMO_START_DATE").Text) Then
                Absx1.dteFor("PROMO_START_DATE").DateTime = CDate(e.Row.Cells("PROMO_START_DATE").Text)
            End If
            If IsDate(e.Row.Cells("PROMO_END_DATE").Text) Then
                Absx1.dteFor("PROMO_END_DATE").DateTime = CDate(e.Row.Cells("PROMO_END_DATE").Text)
            End If

            Click_Command("View")
        End If
    End Sub

#End Region

#Region "grdICTPROM2"
    Private Sub grdICTPROM2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdICTPROM2.InitializeRow

    End Sub

    Private Sub grdICTPROM2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTPROM2.AfterRowActivate
        getStyleInfo()
    End Sub

    Private Sub chkOffListPrice_CheckedChanged(sender As Object, e As EventArgs) Handles chkOffListPrice.CheckedChanged
        If chkOffListPrice.Checked Then
            lblOffPrice.Text = "Off List Price"
        Else
            lblOffPrice.Text = "Off Current Price"
        End If
    End Sub
#End Region
End Class