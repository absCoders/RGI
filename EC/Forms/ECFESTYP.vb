Imports System.Text
Imports Infragistics.Win.UltraWinGrid

Public Class ECFESTYP
    Private SQL As New StringBuilder With {.Length = 0}
    Private rowECTESTYX As DataRow
    Private ECOM_PROMO_CTL_NO As String = ""
    Private FormLoading As Boolean = True

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        FormLoading = True

        If MENU_ITEM_OBJECT = "ECTESTYP" Then
            InquiryMode = True
        End If

        Get_PARM("ECTPARM1")

        initDST()

        setGridDataSources()

        setGridSummaries()

        setGridSorts()

        initFillDST()

        setBinding()

        setGridDefaults()

        setGridValueLists()

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
                If Absx1.txtFor("ECOM_PROMO_CTL_NO").Text & "" <> "" Then
                    EMsg &= vbCr & "Promo # Must Be Blank"
                End If

                If Absx1.txtFor("ECOM_CODE").Value & "" = "" Then
                    EMsg &= vbCr & "Missing Partner"
                End If

                If Not IsDate(txtPROMO_START_DATE.Value) Or Not IsDate(txtPROMO_END_DATE.Value) Then
                    EMsg &= vbCr & "Missing Start/End Date"
                Else
                    If txtPROMO_END_DATE.Value < txtPROMO_START_DATE.Value Then
                        EMsg &= vbCr & "End Date Can Not Be Before Start Date"
                    End If
                End If

            Case "View"

            Case "Update"

            Case "Delete"
                If MsgBox("Are You Sure You Want To Delete This Promo?", MsgBoxStyle.YesNo,
                          "Make Sure!!") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Upload Spreadsheet"
                Dim msg As New System.Text.StringBuilder With {.Length = 0}
                msg.AppendLine("This Will Prompt You For An Excel Spreadsheet.")
                msg.AppendLine("")
                msg.AppendLine("It Should Be In The Same Format As If You")
                msg.AppendLine("Downloaded It From This Grid!!")
                msg.AppendLine("")
                msg.AppendLine("Once Selected,It Will Verfy Then Preplace All")
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
                initFillDST()
            Case "Delete"
                Delete_Promo()
                Mode_Settings(False)
                initFillDST()
            Case "Cancel"
                Mode_Settings(False)
                initFillDST()

            Case "Add Style/Color"
                AddStyleColor()

            Case "Upload Spreadsheet"
                UploadSpreadsheet()

        End Select

    End Sub

    Private Sub Delete_Promo()
        BeginTrans()
        Dim ECOM_PROMO_CTL_NO As String = Absx1.txtFor("ECOM_PROMO_CTL_NO").Text.ToString & String.Empty
        If ECOM_PROMO_CTL_NO.Length > 0 Then
            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}

            Dim BATCH_NO As String = ASCMAIN1.Next_Control_No("ECTESTY3_U.BATCH_NO")
            SQLS.Length = 0
            SQLS.AppendLine("INSERT INTO ECTESTY3_U")
            SQLS.AppendLine(String.Format("SELECT '{0}' AS BATCH_NO, ECTESTY3.* FROM ECTESTY3", BATCH_NO))
            SQLS.AppendLine(String.Format("WHERE ECOM_PROMO_CTL_NO = {0}", ECOM_PROMO_CTL_NO))
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()

            For Each DELTABLE As String In New String() {"ECTESTYP", "ECTESTY3"}
                SQLS.Length = 0
                SQLS.AppendLine(String.Format("DELETE FROM {0} WHERE ECOM_PROMO_CTL_NO = '{1}'", DELTABLE, ECOM_PROMO_CTL_NO))
                ASCMAIN1.sql = SQLS.ToString
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
                    .Items("Add Style/Color").Settings.Enabled = Not ScreenMode
                    .Items("Upload Spreadsheet").Settings.Enabled = Not ScreenMode
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
        clearTables.Add("ECTESTY3")

        For Each clearTable As String In clearTables
            dst.Tables(clearTable).Rows.Clear()
        Next

        EnforceConstraints(True)

        txtPROMO_START_DATE.Value = ""
        txtPROMO_END_DATE.Value = ""
        Absx1.txtFor("ECOM_PROMO_CTL_NO").Text = ""
        Absx1.txtFor("ECOM_CODE").Text = ""
        ECOM_PROMO_CTL_NO = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        ECOM_PROMO_CTL_NO = Absx1.txtFor("ECOM_PROMO_CTL_NO").Text & ""

        Dim loadTables As New List(Of String)
        loadTables.Add("ECTESTY3")

        If EntryMode = "N" Then
            Dim ECOM_PROMO_CTL_NO As String = ASCMAIN1.Next_Control_No("ECTESTYP.ECOM_PROMO_CTL_NO")
            Dim ECOM_CODE As String = Absx1.txtFor("ECOM_CODE").Text & String.Empty

            Absx1.txtFor("ECOM_PROMO_CTL_NO").Text = ECOM_PROMO_CTL_NO
            Fill_Records("ECTESTY3", ECOM_PROMO_CTL_NO)
        Else
            For Each loadTable As String In loadTables
                Fill_Records(loadTable, ECOM_PROMO_CTL_NO)
            Next
        End If

        init_Misc()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        'This is stupid but I spent too much time fighting the standards.
        Dim ECOM_PROMO_CTL_NO As String = Absx1.txtFor("ECOM_PROMO_CTL_NO").Text
        Dim ECOM_CODE As String = Absx1.txtFor("ECOM_CODE").Text
        SQL.Length = 0
        SQL.AppendLine("SELECT COUNT(*) AS RECCNT")
        SQL.AppendLine("FROM ECTESTYP")
        SQL.AppendLine(String.Format("WHERE ECOM_PROMO_CTL_NO = '{0}'", ECOM_PROMO_CTL_NO))
        ASCMAIN1.sql = SQL.ToString()
        Dim RECCNT As Integer = Val(ASCDATA1.GetDataValue)
        If RECCNT = 0 Then
            SQL.Length = 0
            SQL.AppendLine("INSERT INTO ECTESTYP")
            SQL.AppendLine("VALUES(")
            SQL.AppendLine(String.Format("'{0}',", ECOM_PROMO_CTL_NO))
            SQL.AppendLine(String.Format("'{0}',", ECOM_CODE))
            SQL.AppendLine(String.Format("'{0}',", Format(txtPROMO_START_DATE.DateTime, "dd-MMM-yy")))
            SQL.AppendLine(String.Format("'{0}'", Format(txtPROMO_END_DATE.DateTime, "dd-MMM-yy")))
            SQL.AppendLine(")")
            ASCMAIN1.sql = SQL.ToString
            ASCDATA1.ExecuteSQL()
        Else
            SQL.Length = 0
            SQL.AppendLine("UPDATE ECTESTYP")
            SQL.AppendLine(String.Format("SET PROMO_END_DATE = '{0}'", Format(txtPROMO_END_DATE.DateTime, "dd-MMM-yy")))
            SQL.AppendLine(String.Format("WHERE ECOM_PROMO_CTL_NO = '{0}'", ECOM_PROMO_CTL_NO))
            ASCMAIN1.sql = SQL.ToString
            ASCDATA1.ExecuteSQL()
        End If

        'Dim updateTables As New List(Of String)
        'updateTables.Add("ECTESTY3")

        BeginTrans()
        Update_Record_TDA("ECTESTY3", String.Format("ECOM_PROMO_CTL_NO = '{0}'", Absx1.txtFor("ECOM_PROMO_CTL_NO").Text.ToString & String.Empty))
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
        Load_Popup_Menu(grdECFESTYX, "SS", "Show Filter", "Show GroupBox")
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
                For Each rowECTESTY3 As DataRow In dst.Tables("ECTESTY3").Select()
                    rowECTESTY3.Item("PROMO_END_DATE") = dteUpdateEndDate.Value
                Next
            End If
        End If
    End Sub

    Private Sub btnMarkDownPct_Click(sender As Object, e As EventArgs) Handles btnMarkDownPct.Click
        If numMarkDownPct.Value <= 0 Then
            MsgBox("Markdown Pct Must Be Greater Than 0", vbOKOnly, "Check Pct")
        Else
            For Each rowECTESTY3 As DataRow In dst.Tables("ECTESTY3").Select()
                rowECTESTY3.Item("PROMO_UNIT_PRICE") = Val(rowECTESTY3.Item("ECOM_UNIT_PRICE")) - (Val(rowECTESTY3.Item("ECOM_UNIT_PRICE")) * (numMarkDownPct.Value / 100))
            Next
        End If
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub AddStyleColor()
        Dim S As New Text.StringBuilder With {.Length = 0}
        S.AppendLine("Select")
        S.AppendLine("E2.STYLE_CODE, ")
        S.AppendLine("E2.COLOR_CODE, ")
        S.AppendLine("E2.ECOM_CODE, ")
        S.AppendLine("S1.STYLE_DESC, ")
        S.AppendLine("E1.ECOM_UNIT_PRICE")
        S.AppendLine("FROM ECTESTY1 E1, ECTESTY2 E2, ICTSTYL1 S1")
        S.AppendLine("WHERE E1.STYLE_CODE = E2.STYLE_CODE")
        S.AppendLine("And E2.STYLE_CODE = S1.STYLE_CODE")
        S.AppendLine(String.Format("And E1.ECOM_CODE = '{0}'", Absx1.txtFor("ECOM_CODE").Text))
        S.AppendLine(String.Format("And E2.ECOM_CODE = '{0}'", Absx1.txtFor("ECOM_CODE").Text))
        S.AppendLine("AND E2.ECOM_STYLE_COLOR_STATUS = 'A'")
        With ASCMAIN1.CodeSelector
            .SQL = S.ToString
            .MultipleSelections = True
            .PreviouslySelectedCodes0 = ""
            .Caption = "Select Style/Colors To Add"
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
        If ASCMAIN1.CodeSelector.Selections = 0 Then
            MsgBox("Nothing Selected", vbOKOnly, "Please Select Something")
        Else
            Dim ERR_MSG1 As New StringBuilder With {.Length = 0}
            Dim ERR_MSG2 As New StringBuilder With {.Length = 0}
            For i As Integer = 0 To ASCMAIN1.CodeSelector.Selections - 1
                Dim STYLE_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(i).Item("STYLE_CODE") & String.Empty
                Dim COLOR_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(i).Item("COLOR_CODE") & String.Empty
                Dim ECOM_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(i).Item("ECOM_CODE") & String.Empty
                Dim Filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
                If dst.Tables.Item("ECTESTY3").Select(Filter).Count <> 0 Then
                    ERR_MSG1.AppendLine(String.Format("{0}-{1}", STYLE_CODE, COLOR_CODE))
                Else
                    Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                    SQLS.AppendLine("SELECT *")
                    SQLS.AppendLine("FROM ECTESTY3")
                    SQLS.AppendLine(String.Format("WHERE ECOM_CODE = '{0}'", ECOM_CODE))
                    SQLS.AppendLine(String.Format("AND STYLE_CODE = '{0}'", STYLE_CODE))
                    SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
                    SQLS.AppendLine(String.Format("AND (PROMO_START_DATE BETWEEN '{0}' AND '{1}'", Format(txtPROMO_START_DATE.DateTime, "dd-MMM-yy"), Format(txtPROMO_END_DATE.DateTime, "dd-MMM-yy")))
                    SQLS.AppendLine(String.Format("Or PROMO_END_DATE BETWEEN '{0}' AND '{1}'", Format(txtPROMO_START_DATE.DateTime, "dd-MMM-yy"), Format(txtPROMO_END_DATE.DateTime, "dd-MMM-yy")))
                    SQLS.AppendLine(String.Format("Or (PROMO_START_DATE < '{0}' AND '{1}' < PROMO_END_DATE ))", Format(txtPROMO_START_DATE.DateTime, "dd-MMM-yy"), Format(txtPROMO_END_DATE.DateTime, "dd-MMM-yy")))

                    Dim tblECTESTY3 As DataTable = ASCDATA1.GetDataTable(SQLS.ToString(), String.Empty)
                    If tblECTESTY3.Rows.Count > 0 Then
                        For Each rowECTESTY3 As DataRow In tblECTESTY3.Rows
                            Dim ECOM_PROMO_CTL_NO As String = rowECTESTY3.Item("ECOM_PROMO_CTL_NO").ToString & String.Empty
                            Dim SD As String = "No Date"
                            If IsDate(rowECTESTY3.Item("PROMO_START_DATE").ToString & String.Empty) Then
                                SD = CDate(rowECTESTY3.Item("PROMO_START_DATE").ToString & String.Empty).ToShortDateString
                            End If
                            Dim ED As String = "No Date"
                            If IsDate(rowECTESTY3.Item("PROMO_END_DATE").ToString & String.Empty) Then
                                ED = CDate(rowECTESTY3.Item("PROMO_END_DATE").ToString & String.Empty).ToShortDateString
                            End If
                            ERR_MSG2.AppendLine(String.Format("{0} | {1} | {2} | {3}-{4}", ECOM_PROMO_CTL_NO, SD, ED, STYLE_CODE, COLOR_CODE))
                        Next
                    Else
                        Dim rowECTESTY3 As DataRow = dst.Tables.Item("ECTESTY3").NewRow
                        rowECTESTY3.Item("ECOM_PROMO_CTL_NO") = Absx1.txtFor("ECOM_PROMO_CTL_NO").Text & String.Empty
                        rowECTESTY3.Item("STYLE_CODE") = STYLE_CODE
                        rowECTESTY3.Item("COLOR_CODE") = COLOR_CODE
                        rowECTESTY3.Item("ECOM_CODE") = ECOM_CODE
                        rowECTESTY3.Item("STYLE_DESC") = ASCMAIN1.CodeSelector.SelectedRows(i).Item("STYLE_DESC") & String.Empty
                        rowECTESTY3.Item("PROMO_START_DATE") = txtPROMO_START_DATE.DateTime
                        rowECTESTY3.Item("PROMO_END_DATE") = txtPROMO_END_DATE.DateTime
                        rowECTESTY3.Item("ECOM_UNIT_PRICE") = Val(ASCMAIN1.CodeSelector.SelectedRows(i).Item("ECOM_UNIT_PRICE") & String.Empty)
                        rowECTESTY3.Item("PROMO_UNIT_PRICE") = Val(ASCMAIN1.CodeSelector.SelectedRows(i).Item("ECOM_UNIT_PRICE") & String.Empty)
                        dst.Tables.Item("ECTESTY3").Rows.Add(rowECTESTY3)
                    End If
                End If
            Next
            grdECTESTY3.Refresh()
            If ERR_MSG1.Length > 0 Then
                MsgBox(ERR_MSG1.ToString, vbOKOnly, "Styles/Colors Already On This Promo")
            End If
            If ERR_MSG2.Length > 0 Then
                Dim emsg As String = "Promo | Start | End | Style-Color" & vbCrLf & ERR_MSG2.ToString
                MsgBox(emsg, vbOKOnly, "Styles/Colors Conflicting With Other Promos")
            End If
        End If
    End Sub

    Private Sub initDST()
        With dst
            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine("FROM ECTPARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTPARM1", "**", 0, False)

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("YP.ECOM_PROMO_CTL_NO,")
            SQL.AppendLine("YP.ECOM_CODE,")
            SQL.AppendLine("C1.ECOM_NAME,")
            SQL.AppendLine("YP.PROMO_START_DATE,")
            SQL.AppendLine("YP.PROMO_END_DATE,")
            SQL.AppendLine("COUNT(Y3.STYLE_CODE) AS STYLE_CNT")
            SQL.AppendLine("FROM ECTESTYP YP, ECTECOM1 C1, ECTESTY3 Y3")
            SQL.AppendLine("WHERE YP.ECOM_CODE= C1.ECOM_CODE")
            SQL.AppendLine("AND YP.ECOM_PROMO_CTL_NO = Y3.ECOM_PROMO_CTL_NO")
            SQL.AppendLine("GROUP BY")
            SQL.AppendLine("YP.ECOM_PROMO_CTL_NO,")
            SQL.AppendLine("YP.ECOM_CODE,")
            SQL.AppendLine("C1.ECOM_NAME,")
            SQL.AppendLine("YP.PROMO_START_DATE,")
            SQL.AppendLine("YP.PROMO_END_DATE")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTESTYX", "**", 0, False)
            'With .Tables("ECTECOM1_PARTNER").Columns
            '    .Add("SEL", GetType(System.String))
            'End With

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("E3.ECOM_PROMO_CTL_NO,")
            SQL.AppendLine("E3.STYLE_CODE,")
            SQL.AppendLine("E3.COLOR_CODE,")
            SQL.AppendLine("E3.ECOM_CODE,")
            SQL.AppendLine("S1.STYLE_DESC,")
            SQL.AppendLine("E3.PROMO_START_DATE,")
            SQL.AppendLine("E3.PROMO_END_DATE,")
            SQL.AppendLine("E1.ECOM_UNIT_PRICE,")
            SQL.AppendLine("E3.PROMO_UNIT_PRICE")
            SQL.AppendLine("FROM ECTESTY3 E3, ECTESTY1 E1, ICTSTYL1 S1")
            SQL.AppendLine("WHERE E3.STYLE_CODE = E1.STYLE_CODE")
            SQL.AppendLine("AND E3.ECOM_CODE = E1.ECOM_CODE")
            SQL.AppendLine("AND E1.STYLE_CODE = S1.STYLE_CODE")
            SQL.AppendLine("AND E3.ECOM_PROMO_CTL_NO = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTESTY3", "**", 0, True, "V", 4)

            'SQL.Length = 0
            'SQL.AppendLine("SELECT *")
            'SQL.AppendLine("FROM ECTESTYH")
            'SQL.AppendLine("WHERE ECOM_PROMO_CTL_NO = :PARM1")
            'ASCMAIN1.sql = SQL.ToString()
            'Create_TDA(.Tables.Add, "ECTESTYH", "**", 0, True, "V")
        End With
    End Sub

    Private Sub initFillDST()
        Fill_Records("ECTPARM1")
        Fill_Records("ECTESTYX")
        'For Each rowECTECOM1_FILTER As DataRow In dst.Tables("ECTECOM1_FILTER").Select()
        '    rowECTECOM1_FILTER.Item("SEL") = "1"
        'Next
    End Sub

    Private Sub init_Misc()

    End Sub

    Private Sub setBinding()
        'Bind_Controls(pnlICTSTYL1, "ICTSTYL1")
    End Sub

    Private Sub setGridDataSources()
        grdECFESTYX.DataSource = dst.Tables("ECTESTYX")
        grdECTESTY3.DataSource = dst.Tables("ECTESTY3")
    End Sub

    Private Sub setGridDefaults()
        With grdECFESTYX.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("STYLE_CNT").Format = "###,##0"
            '.Columns("SEL").CellActivation = UltraWinGrid.Activation.AllowEdit
        End With

        With grdECTESTY3.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.True
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("PROMO_UNIT_PRICE").Format = "###,##0.00"
            .Columns("ECOM_UNIT_PRICE").Format = "###,##0.00"

            .Columns("PROMO_UNIT_PRICE").CellActivation = UltraWinGrid.Activation.AllowEdit
        End With
    End Sub

    Private Sub setGridSorts()
        Sort_grdColumns(grdECFESTYX, "ECOM_PROMO_CTL_NO", False)
        Sort_grdColumns(grdECTESTY3, "STYLE_CODE", False)
    End Sub

    Private Sub setGridSummaries()
        Create_Summary(grdECFESTYX, "ECOM_PROMO_CTL_NO", "Count")
        'Create_Summary(grdECTESTY3, "STYLE_CODE", "Count")
    End Sub

    Private Sub setGridValueLists()
        'ASCMAIN1.Add_Value_List(grdECTSTYLX, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")
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
            openFileDialog1.Filter = "Excel files (*.xls)|*.xls"
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

        Dim strConnection As String = "Provider=Microsoft.Jet.OleDb.4.0;" &
                "data source=" & fileToImport & ";" &
                "Extended Properties=Excel 8.0;"

        Using cn As New System.Data.OleDb.OleDbConnection(strConnection)
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
                If rowUPSERT.Item(1).ToString & String.Empty <> "Color" Or rowUPSERT.Item(4).ToString & String.Empty <> "Promo Price" Then
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
            Dim BATCH_NO As String = ASCMAIN1.Next_Control_No("ECTESTY3_U.BATCH_NO")
            Dim ECOM_PROMO_CTL_NO As String = Absx1.txtFor("ECOM_PROMO_CTL_NO").Text.ToString & String.Empty
            SQL.Length = 0
            SQL.AppendLine("INSERT INTO ECTESTY3_U")
            SQL.AppendLine(String.Format("SELECT '{0}' AS BATCH_NO, ECTESTY3.* FROM ECTESTY3", BATCH_NO))
            SQL.AppendLine(String.Format("WHERE ECOM_PROMO_CTL_NO = {0}", ECOM_PROMO_CTL_NO))
            ASCMAIN1.sql = SQL.ToString
            ASCDATA1.ExecuteSQL()
            dst.Tables.Item("ECTESTY3").Clear()

            For i As Int64 = FirstRow + 1 To LastRow
                Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                Dim HAS_PROBLEM As Boolean = False
                Dim STYLE_CODE As String = UpsertData.Rows(i).Item(0).ToString.ToUpper
                Dim STYLE_DESC As String = ""
                Dim COLOR_CODE As String = UpsertData.Rows(i).Item(1).ToString.ToUpper
                Dim PROMO_UNIT_PRICE As String = UpsertData.Rows(i).Item(4).ToString.ToUpper
                Dim ECOM_UNIT_PRICE As Double = 0

                SQLS.Length = 0
                SQLS.AppendLine("SELECT COUNT(*) AS E_CNT")
                SQLS.AppendLine("FROM ECTESTY2")
                SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
                ASCMAIN1.sql = SQLS.ToString()
                Dim E_CNT As Int16 = Val(ASCDATA1.GetDataValue)

                If E_CNT = 0 Then
                    PROBLEM_STYLES.AppendLine(String.Format("{0}-{1} Is Not In The E-Commerce System.", STYLE_CODE, COLOR_CODE))
                    HAS_PROBLEM = True
                Else
                    SQLS.Length = 0
                    SQLS.AppendLine("SELECT STYLE_DESC")
                    SQLS.AppendLine("FROM ICTSTYL1")
                    SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                    ASCMAIN1.sql = SQLS.ToString()
                    STYLE_DESC = ASCDATA1.GetDataValue

                    SQLS.Length = 0
                    SQLS.AppendLine("Select")
                    SQLS.AppendLine("MAX(E1.ECOM_UNIT_PRICE) AS ECOM_UNIT_PRICE")
                    SQLS.AppendLine("FROM ECTESTY1 E1, ECTESTY2 E2, ICTSTYL1 S1")
                    SQLS.AppendLine("WHERE E1.STYLE_CODE = E2.STYLE_CODE")
                    SQLS.AppendLine("And E2.STYLE_CODE = S1.STYLE_CODE")
                    SQLS.AppendLine(String.Format("And E1.ECOM_CODE = '{0}'", Absx1.txtFor("ECOM_CODE").Text))
                    SQLS.AppendLine(String.Format("And E2.ECOM_CODE = '{0}'", Absx1.txtFor("ECOM_CODE").Text))
                    SQLS.AppendLine(String.Format("And E2.STYLE_CODE = '{0}'", STYLE_CODE))
                    SQLS.AppendLine(String.Format("And E2.COLOR_CODE = '{0}'", COLOR_CODE))
                    SQLS.AppendLine("AND E2.ECOM_STYLE_COLOR_STATUS = 'A'")
                    ASCMAIN1.sql = SQLS.ToString()
                    ECOM_UNIT_PRICE = Val(ASCDATA1.GetDataValue)

                End If

                If Not IsNumeric(PROMO_UNIT_PRICE) Then
                    PROBLEM_STYLES.AppendLine(String.Format("{0}-{1} Has Invalid Promo Price", STYLE_CODE, COLOR_CODE))
                    HAS_PROBLEM = True
                End If
                If HAS_PROBLEM = False Then
                    Dim newECTESTY3 As DataRow = dst.Tables("ECTESTY3").NewRow
                    newECTESTY3.Item("ECOM_PROMO_CTL_NO") = Absx1.txtFor("ECOM_PROMO_CTL_NO").Text.ToString & String.Empty
                    newECTESTY3.Item("STYLE_CODE") = STYLE_CODE
                    newECTESTY3.Item("COLOR_CODE") = COLOR_CODE
                    newECTESTY3.Item("ECOM_CODE") = Absx1.txtFor("ECOM_CODE").Text.ToString & String.Empty
                    newECTESTY3.Item("PROMO_START_DATE") = CDate(Absx1.dteFor("PROMO_START_DATE").Value & String.Empty)
                    newECTESTY3.Item("PROMO_END_DATE") = CDate(Absx1.dteFor("PROMO_END_DATE").Value & String.Empty)
                    newECTESTY3.Item("PROMO_UNIT_PRICE") = Val(PROMO_UNIT_PRICE)
                    newECTESTY3.Item("STYLE_DESC") = STYLE_DESC
                    newECTESTY3.Item("ECOM_UNIT_PRICE") = ECOM_UNIT_PRICE
                    dst.Tables("ECTESTY3").Rows.Add(newECTESTY3)
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

#Region "grdECTSTYLX"
    Private Sub grdECFESTYX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdECFESTYX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("ECOM_PROMO_CTL_NO").Text = e.Row.Cells("ECOM_PROMO_CTL_NO").Text
            Absx1.txtFor("ECOM_CODE").Text = e.Row.Cells("ECOM_CODE").Text
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
End Class