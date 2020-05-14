Public Class SOFCSTY1

#Region "Declarations"
    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim rowICTSTYL1 As DataRow
    Dim STYLE_CODE_last As String
    Dim COLOR_CODE_last As String
    Dim COLOR_CODEs As New List(Of String)    ' table of COLOR_CODEs associated with a STYLE_CODE
#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst
            ASCMAIN1.sql = "Select X.CUST_CODE, ARTCUST1.CUST_NAME" _
                & " from (Select Distinct CUST_CODE from SOTCSTY1) X,ARTCUST1" _
                & " where ARTCUST1.CUST_CODE = X.CUST_CODE"
            Create_TDA(.Tables.Add, "SOTCSTYX", "**", 0, False)

            ASCMAIN1.sql = "Select SOTCSTY1.*,ICTSTYL1.STYLE_DESC,ICTCOLR1.COLOR_DESC" _
                & " from SOTCSTY1,ICTSTYL1,ICTCOLR1" _
                & " where ICTSTYL1.STYLE_CODE = SOTCSTY1.STYLE_CODE" _
                & "   and ICTCOLR1.COLOR_CODE = SOTCSTY1.COLOR_CODE" _
                & "   and SOTCSTY1.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SOTCSTY1", "**", 0, True, "V")
            .Tables("SOTCSTY1").Columns.Add("STATE")
            .Tables("SOTCSTY1").Columns.Add("SELECTED")
            .Tables("SOTCSTY1").Columns.Add("CUST_STYLE_CODE_ORIG")

            ASCMAIN1.sql = "Select ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC" _
               & " from ICTSTYC1,ICTCOLR1 where ICTSTYC1.STYLE_CODE = :PARM1" _
               & "  and ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTCOLRS", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTCSTY1.*,ICTSTYL1.STYLE_DESC,ICTCOLR1.COLOR_DESC" _
                & " from SOTCSTY1,ICTSTYL1,ICTCOLR1" _
                & " where ICTSTYL1.STYLE_CODE = SOTCSTY1.STYLE_CODE" _
                & "   and ICTCOLR1.COLOR_CODE = SOTCSTY1.COLOR_CODE"
            Create_TDA(.Tables.Add, "SOTCSTYO", "**", 0, False, "", 0)
            .Tables("SOTCSTYO").Columns.Add("STATE")
            .Tables("SOTCSTYO").Columns.Add("SELECTED")

            ASCMAIN1.sql = "Select * from ICTSTYC1 where STYLE_CODE = :PARM1 and STYLE_COLOR_STATUS = 'A'"
            Create_TDA(.Tables.Add, "ICTSTYCX", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select X.UPC_CODE, X.STYLE_CODE, X.COLOR_CODE, X.STYLE_DESC, X.CUST_CODE" & vbCrLf _
                & ", SOTCSTY1.CUST_PRICE, SOTCSTY1.SIZE_DESC, SOTCSTY1.CUST_STYLE_CODE" & vbCrLf _
                & " from SOTCSTY1, (" & vbCrLf _
                & "Select ICTUPCH1.*, ICTSTYL1.STYLE_DESC, ICTSTYL1.CUST_CODE" & vbCrLf _
                & " from ICTUPCH1,ICTSTYL1" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE (+) = ICTUPCH1.STYLE_CODE" & vbCrLf _
                & "   and (:PARM1 is Null or ICTSTYL1.CUST_CODE = :PARM1)" & vbCrLf _
                & ") X" & vbCrLf _
                & " where SOTCSTY1.CUST_CODE (+) = X.CUST_CODE" & vbCrLf _
                & "   and SOTCSTY1.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                & "   and SOTCSTY1.COLOR_CODE (+) = X.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTUPCHX", "**", 0, False, "V", 0)

        End With

        grdSOTCSTYX.DataSource = dst.Tables("SOTCSTYX")
        grdSOTCSTY1.DataSource = dst.Tables("SOTCSTY1")
        grdSOTCSTYO.DataSource = dst.Tables("SOTCSTYO")

        grdICTUPCHX.DataSource = dst.Tables("ICTUPCHX")

        grdSOTCSTY1.DisplayLayout.UseFixedHeaders = True
        With grdSOTCSTY1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_STYLE_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        grdSOTCSTYX.DisplayLayout.UseFixedHeaders = True
        With grdSOTCSTYX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSOTCSTY1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "STYLE_DESC" Or gcol.Key = "COLOR_DESC" Or gcol.Key = "STATE" Then
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                ElseIf New String() {"INIT_OPER", "INIT_DATE", "LAST_OPER", "LAST_DATE"}.Contains(gcol.Key) Then
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            Next
        End With

        Create_Summary(grdSOTCSTYX, "CUST_CODE", "Count")
        Create_Summary(grdSOTCSTY1, "CUST_STYLE_CODE", "Count")

        Create_Summary(grdICTUPCHX, "CUST_STYLE_CODE", "Count")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New", "Edit", "View"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 IsNot Nothing Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                    Else
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SOTCSTY1", CUST_CODE) Then Exit Sub
                End If

            Case "Update"

                If Not splSOTCSTY1.Panel2Collapsed Then
                    If MsgBox("OK to continue with the Update leaving the EDI Styles Screen Up without Updating?", _
                              MsgBoxStyle.YesNo, _
                              "Maybe you forgot to click the big Update button in the Lower Left?") = MsgBoxResult.No Then Exit Sub
                End If

                For Each rowSOTCSTY1 As DataRow In dst.Tables("SOTCSTY1").Select("")
                    rowSOTCSTY1.AcceptChanges()
                    If rowSOTCSTY1.Item("STATE") & "" = "Added" Then
                        rowSOTCSTY1.SetAdded()
                    ElseIf rowSOTCSTY1.Item("STATE") & "" = "Edited" Then
                        rowSOTCSTY1.SetModified()
                    End If
                    If rowSOTCSTY1.Item("STATE") & "" = "Added" Or rowSOTCSTY1.Item("STATE") & "" = "Edited" Then
                        Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {rowSOTCSTY1.Item("STYLE_CODE"), rowSOTCSTY1.Item("COLOR_CODE")})
                        If rowICTSTYC1 Is Nothing Then
                            EMsg &= vbCr & "Bad Style/Color for Customer Style " & rowSOTCSTY1.Item("CUST_STYLE_CODE")
                        End If
                    End If
                Next
                '  Stop

            Case "Cancel"
                If EMsg = "" Then
                    If MsgBox("Do you really want to Cancel all changes made to this record?", _
                              MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Verification") <> MsgBoxResult.Yes Then
                        Exit Sub
                    End If
                End If

            Case "Delete"

                If EntryMode = "" Then
                    Exit Sub
                End If

                If EMsg = "" Then
                    If MsgBox("Do you really want to Delete all Pricing Data for Customer " & CUST_CODE & "?", _
                              MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "WARNING! - Answering 'Yes' will PERMANENTLY DELETE these records") <> MsgBoxResult.Yes Then
                        Exit Sub
                    End If
                End If

            Case "Show UPCs"

                'If Absx1.txtFor("CUST_CODE").Text = "" Then
                '    If MsgBox("", MsgBoxStyle.YesNo, "You have not Selected to view the UPCs of a Single Customer") = MsgBoxResult.No Then
                '        Exit Sub
                '    End If

                'Else
                '    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                '    If rowARTCUST1 IsNot Nothing Then
                '        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                '    Else
                '        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                '    End If
                'End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Select"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Show UPCs"

                Show_UPCs()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Find Styles By").Visible = ScreenMode And (EntryMode = "E")
            With .Groups("Screen Control")

                .Items("New").Settings.Enabled = not_iScreenMode
                If (EntryMode = "V" And ScreenMode) Then
                    .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                Else
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                End If
                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Delete").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("View").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
                .Items("Delete").Visible = False ' (EntryMode = "E")
                .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode)

                .Items("Show UPCs").Visible = Not ScreenMode And (ASCMAIN1.CLIENT = "VAN")
            End With
        End With

        ' grdSOTCSTYX.Visible = Not ScreenMode
        splMain.Visible = Not ScreenMode

        splSOTCSTY1.Visible = ScreenMode
        splSOTCSTY1.Panel2Collapsed = True

        If ScreenMode Then
            grdICTUPCHX.Parent = tabSOTCSTY1.Tabs("UPCs").TabPage
            grdICTUPCHX.Visible = True
        Else
            grdICTUPCHX.Parent = splMain.Panel2
            grdICTUPCHX.Visible = False
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTCSTY1}
                If EntryMode = "V" Then
                    grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                Else
                    grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                End If
            Next

            tabSOTCSTY1.Tabs("UPCs").Visible = (EntryMode = "V")
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        Absx1.txtFor("CUST_CODE").Text = ""
        CUST_CODE = ""
        txtFindBy.Text = ""

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTCSTY1", "SOTCSTYO"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        dst.Tables("ASTSQLX1").Rows.Clear() ' TAKES WAY TOO LONG TO UPDATE TO ORACLE

        grdSOTCSTY1.Rows.ColumnFilters.ClearAllFilters()

        STYLE_CODE_last = ""
        COLOR_CODE_last = ""

        Load_SOTCSTYX("")
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Fill_Records("SOTCSTY1", CUST_CODE)

        For Each rowSOTCSTY1 As DataRow In dst.Tables("SOTCSTY1").Select("")
            rowSOTCSTY1.Item("CUST_STYLE_CODE_ORIG") = rowSOTCSTY1.Item("CUST_STYLE_CODE")
        Next
        dst.Tables("SOTCSTY1").AcceptChanges()

        Sort_grdColumns(grdSOTCSTY1, "CUST_STYLE_CODE")

        EnforceConstraints(True)

        Show_UPCs()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Delete_Record()
        If EntryMode <> "E" Then Exit Sub
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        Stop
        'If EntryMode = "N" Then Exit Sub
        'For Each TABLE_NAME As String In New String() _
        '    {"SOTRSRV1", "SOTRSRV2"}
        '    Delete_Records_1(TABLE_NAME)
        'Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        'ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where CUST_CODE = '" & CUST_CODE & "'"
        'ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        BeginTrans()

        For Each row As DataRow In dst.Tables("SOTCSTY1").Select("CUST_STYLE_CODE <> CUST_STYLE_CODE_ORIG", "", DataViewRowState.ModifiedCurrent)
            Dim CUST_STYLE_CODE As String = row.Item("CUST_STYLE_CODE") & ""
            Dim CUST_STYLE_CODE_ORIG As String = row.Item("CUST_STYLE_CODE_ORIG") & ""
            If CUST_STYLE_CODE <> CUST_STYLE_CODE_ORIG And CUST_STYLE_CODE_ORIG <> "" Then
                ASCMAIN1.sql = "Delete from SOTCSTY1 where CUST_CODE = '" & CUST_CODE & "' and CUST_STYLE_CODE = '" & CUST_STYLE_CODE_ORIG & "'"
                ASCDATA1.ExecuteSQL()
                row.AcceptChanges()
                row.SetAdded()
            End If
        Next

        Update_Record_TDA("SOTCSTY1") ', sqldelete)
        ' Rollback()
        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "EDI_DOC_SEQ_NO"
                Select Case optFindBy.Value
                    Case "P"

                    Case "E"

                    Case "O"

                End Select


                Cancel = True
        End Select
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "Select"

                Dim CUST_CODE As String = Split(key, ":")(0)
                Dim ORDR_GROUP_NO As String = Split(key & ":", ":")(1)
                Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                Click_Command("View")
        End Select

        Return return_key
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTCSTYX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdICTUPCHX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTCSTY1, "SSSBSBB", "Show Filter", "Show GroupBox", "Show Pins", _
                        "Style Status Inquiry", "Show Selected Only", "Load from Spreadsheet", "Load from Retail Link")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '    e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdSOTCSTY1"
                    tlb_btn = DirectCast(tlb_pop.Tools("Load from Spreadsheet"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

                    tlb_btn = DirectCast(tlb_pop.Tools("Load from Retail Link"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And ASCMAIN1.CLIENT = "VAN"
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Show Selected Only"
                Dim dvw As DataView = DirectCast(grdSOTCSTY1.DataSource, DataTable).DefaultView
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                dvw.RowFilter = IIf(tlb_sbt.Checked, "SELECTED = '1'", "")
                Exit Sub ' because e.tool.OwningMenu is set to nothing when checking this tbl_sbt from code


            Case "Load from Spreadsheet"
                Try
                    Dim FILENAME As String = ""
                    Using openFileDialog1 As New OpenFileDialog
                        openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                        openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
                        openFileDialog1.RestoreDirectory = True

                        '  Excel_Import = -1

                        If openFileDialog1.ShowDialog() = DialogResult.OK Then
                            FILENAME = openFileDialog1.FileName
                        End If
                    End Using

                    If FILENAME <> "" Then

                        Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
                        Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
                        Dim range As SpreadsheetGear.IRange = Nothing
                        range = oSheet.UsedRange
                        Dim rmax As Integer = range.RowCount
                        ASCMAIN1.Progress("Now Loading XLS")

                        Dim started As Boolean = False
                        grdSOTCSTY1.Visible = False

                        Dim INVALID_STYLES As String = ""
                        Dim r As Integer = 1
                        Do While r < rmax ' oSheet.Cells(r, 0).Value & "" <> ""

                            Dim STYLE_CODE As String = ""
                            Dim COLOR_CODE As String = ""
                            Dim SIZE_DESC As String = ""
                            Dim CUST_STYLE_CODE As String = ""
                            Dim CUST_UPC As String = ""
                            Dim CUST_PRICE As Decimal = 0
                            Dim VENDOR_STOCK_NO As String = ""

                            CUST_STYLE_CODE = oSheet.Cells(r, 0).Value & ""
                            SIZE_DESC = oSheet.Cells(r, 2).Value & ""
                            STYLE_CODE = oSheet.Cells(r, 3).Value & ""
                            COLOR_CODE = oSheet.Cells(r, 5).Value & ""
                            CUST_UPC = oSheet.Cells(r, 10).Value & ""
                            CUST_PRICE = Val(oSheet.Cells(r, 7).Value & "")
                            VENDOR_STOCK_NO = oSheet.Cells(r, 16).Value & ""

                            ASCMAIN1.Progress("-", CStr(r) & ":" & CUST_STYLE_CODE)

                            If LookUp("ICTSTYL1", STYLE_CODE) IsNot Nothing AndAlso LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE}) IsNot Nothing Then
                                If grdSOTCSTY1.ActiveRow IsNot Nothing AndAlso grdSOTCSTY1.ActiveRow.IsAddRow Then
                                    grdSOTCSTY1.ActiveRow.CancelUpdate()
                                End If

                                Dim rowSOTCSTY1 As DataRow = dst.Tables("SOTCSTY1").Rows.Find(New Object() {CUST_CODE, CUST_STYLE_CODE})
                                If rowSOTCSTY1 IsNot Nothing Then
                                    'INVALID_STYLES &= ", " & STYLE_CODE & "-" & COLOR_CODE & " exists already"
                                    If VENDOR_STOCK_NO.Length > 0 Then
                                        rowSOTCSTY1.Item("VENDOR_STOCK_NO") = VENDOR_STOCK_NO

                                        rowSOTCSTY1.Item("SIZE_DESC") = SIZE_DESC
                                        rowSOTCSTY1.Item("STYLE_CODE") = STYLE_CODE
                                        rowSOTCSTY1.Item("COLOR_CODE") = COLOR_CODE
                                        rowSOTCSTY1.Item("CUST_UPC") = CUST_UPC
                                        rowSOTCSTY1.Item("CUST_PRICE") = CUST_PRICE

                                        If rowSOTCSTY1.Item("STATE") & "" = "Added" Then
                                            INVALID_STYLES &= ", " & CUST_STYLE_CODE & "- Duplicate Customer Style in XLS"
                                        Else
                                            rowSOTCSTY1.Item("STATE") = "Edited"
                                        End If
                                    End If
                                Else
                                    grdSOTCSTY1.DisplayLayout.Bands(0).AddNew.Activate()
                                    With grdSOTCSTY1.ActiveRow
                                        .Cells("CUST_STYLE_CODE").Value = CUST_STYLE_CODE
                                        .Cells("SIZE_DESC").Value = SIZE_DESC
                                        .Cells("STYLE_CODE").Value = STYLE_CODE
                                        .Cells("COLOR_CODE").Value = COLOR_CODE
                                        .Cells("CUST_UPC").Value = CUST_UPC
                                        .Cells("CUST_PRICE").Value = CUST_PRICE
                                        .Cells("VENDOR_STOCK_NO").Value = VENDOR_STOCK_NO

                                        .Update()
                                        started = True
                                        grdSOTCSTY1.ActiveRow.CancelUpdate()
                                    End With
                                End If


                            Else
                                If STYLE_CODE <> "" And COLOR_CODE <> "" Then INVALID_STYLES &= ", " & STYLE_CODE & "-" & COLOR_CODE & " bad Style or Color"
                            End If
                            r += 1
                        Loop

                        grdSOTCSTY1.Visible = True
                        ASCMAIN1.Progress("")

                        ' grdICTRSTYM.ResumeLayout()

                        Sort_grdColumns(grdSOTCSTY1, "CUST_STYLE_CODE")

                        If INVALID_STYLES <> "" Then
                            MsgBox(Mid(INVALID_STYLES, 3), MsgBoxStyle.OkOnly, "The following Style-Color codes were invalid")
                        End If

                        MsgBox("Spreadsheet has been Loaded", MsgBoxStyle.OkOnly, "Success")
                    End If
                Catch ex As Exception
                    MsgBox("Error " & ex.Message, MsgBoxStyle.OkOnly, "Cannot Load this Spreadsheet")
                End Try
                grdSOTCSTY1.UpdateData()


            Case "Load from Retail Link"
                Try
                    Dim FILENAME As String = ""
                    Using openFileDialog1 As New OpenFileDialog
                        openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                        openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
                        openFileDialog1.RestoreDirectory = True

                        '  Excel_Import = -1

                        If openFileDialog1.ShowDialog() = DialogResult.OK Then
                            FILENAME = openFileDialog1.FileName
                        End If
                    End Using

                    If FILENAME <> "" Then

                        Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
                        Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
                        Dim range As SpreadsheetGear.IRange = Nothing
                        range = oSheet.UsedRange
                        Dim rmax As Integer = range.RowCount
                        ASCMAIN1.Progress("Now Loading XLS")

                        Dim started As Boolean = False
                        grdSOTCSTY1.Visible = False

                        Dim INVALID_STYLES As String = ""
                        Dim r As Integer = 1
                        Do While r < rmax ' oSheet.Cells(r, 0).Value & "" <> ""

                            Dim STYLE_CODE As String = ""
                            Dim COLOR_CODE As String = ""
                            Dim SIZE_DESC As String = ""
                            Dim CUST_STYLE_CODE As String = ""
                            Dim CUST_UPC As String = ""
                            Dim CUST_PRICE As Decimal = 0
                            Dim VENDOR_STOCK_NO As String = ""

                            CUST_STYLE_CODE = oSheet.Cells(r, 0).Value & ""
                            SIZE_DESC = oSheet.Cells(r, 1).Value & ""
                            STYLE_CODE = oSheet.Cells(r, 2).Value & ""
                            COLOR_CODE = oSheet.Cells(r, 4).Value & ""
                            CUST_UPC = oSheet.Cells(r, 5).Value & ""
                            CUST_PRICE = Val(oSheet.Cells(r, 7).Value & "")
                            VENDOR_STOCK_NO = oSheet.Cells(r, 6).Value & ""

                            If CUST_UPC.Length = 13 And CUST_UPC.StartsWith("00") Then
                                CUST_UPC = CUST_UPC.Substring(2)
                                CUST_UPC = TAC.SOCMAIN1.UPC(Me, CUST_UPC.Substring(6), CUST_UPC.Substring(0, 6))
                            End If

                            ASCMAIN1.Progress("-", CStr(r) & ":" & CUST_STYLE_CODE)

                            Dim colors As Integer = Fill_Records("ICTSTYCX", STYLE_CODE)
                            If colors = 1 Then
                                COLOR_CODE = dst.Tables("ICTSTYCX").Rows(0).Item("COLOR_CODE")
                            End If

                            If LookUp("ICTSTYL1", STYLE_CODE) IsNot Nothing AndAlso LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE}) IsNot Nothing Then
                                If grdSOTCSTY1.ActiveRow IsNot Nothing AndAlso grdSOTCSTY1.ActiveRow.IsAddRow Then
                                    grdSOTCSTY1.ActiveRow.CancelUpdate()
                                End If

                                Dim rowSOTCSTY1 As DataRow = dst.Tables("SOTCSTY1").Rows.Find(New Object() {CUST_CODE, CUST_STYLE_CODE})
                                If rowSOTCSTY1 IsNot Nothing Then
                                    'INVALID_STYLES &= ", " & STYLE_CODE & "-" & COLOR_CODE & " exists already"
                                    If VENDOR_STOCK_NO.Length > 0 Then
                                        rowSOTCSTY1.Item("VENDOR_STOCK_NO") = VENDOR_STOCK_NO

                                        rowSOTCSTY1.Item("SIZE_DESC") = SIZE_DESC
                                        rowSOTCSTY1.Item("STYLE_CODE") = STYLE_CODE
                                        rowSOTCSTY1.Item("COLOR_CODE") = COLOR_CODE
                                        rowSOTCSTY1.Item("CUST_UPC") = CUST_UPC
                                        rowSOTCSTY1.Item("CUST_PRICE") = CUST_PRICE

                                        If rowSOTCSTY1.Item("STATE") & "" = "Added" Then
                                            INVALID_STYLES &= ", " & CUST_STYLE_CODE & "- Duplicate Customer Style in XLS"
                                        Else
                                            rowSOTCSTY1.Item("STATE") = "Edited"
                                        End If

                                    End If
                                Else
                                    grdSOTCSTY1.DisplayLayout.Bands(0).AddNew.Activate()
                                    With grdSOTCSTY1.ActiveRow
                                        .Cells("CUST_STYLE_CODE").Value = CUST_STYLE_CODE
                                        .Cells("SIZE_DESC").Value = SIZE_DESC
                                        .Cells("STYLE_CODE").Value = STYLE_CODE
                                        .Cells("COLOR_CODE").Value = COLOR_CODE
                                        .Cells("CUST_UPC").Value = CUST_UPC
                                        .Cells("CUST_PRICE").Value = CUST_PRICE
                                        .Cells("VENDOR_STOCK_NO").Value = VENDOR_STOCK_NO

                                        .Update()
                                        started = True
                                        grdSOTCSTY1.ActiveRow.CancelUpdate()
                                    End With
                                End If


                            Else
                                If STYLE_CODE <> "" And COLOR_CODE <> "" Then INVALID_STYLES &= ", " & STYLE_CODE & "-" & COLOR_CODE & " bad Style or Color"
                            End If
                            r += 1
                        Loop

                        grdSOTCSTY1.Visible = True
                        ASCMAIN1.Progress("")

                        ' grdICTRSTYM.ResumeLayout()

                        Sort_grdColumns(grdSOTCSTY1, "CUST_STYLE_CODE")

                        If INVALID_STYLES <> "" Then
                            MsgBox(Mid(INVALID_STYLES, 3), MsgBoxStyle.OkOnly, "The following Style-Color codes were invalid")
                        End If

                        MsgBox("Spreadsheet has been Loaded", MsgBoxStyle.OkOnly, "Success")
                    End If
                Catch ex As Exception
                    MsgBox("Error " & ex.Message, MsgBoxStyle.OkOnly, "Cannot Load this Spreadsheet")
                End Try
                grdSOTCSTY1.UpdateData()



        End Select

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View")
                End If

            Case "EDI_DOC_SEQ_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Dim DT As DataTable = Nothing
                    Select Case optFindBy.Value
                        Case "P"
                            ASCMAIN1.sql = "Select * from (" _
                                & "Select DISTINCT '" & CUST_CODE & "' CUST_CODE" _
                                & ",NVL(CUST_STYLE_CODE,CUST_SKU) CUST_STYLE_CODE" _
                                & ",CUST_SIZE_CODE SIZE_DESC" _
                                & ",ORDR_UNIT_PRICE CUST_PRICE" _
                                & ",STYLE_CODE" _
                                & ",COLOR_CODE" _
                                & ",DECODE(CUST_UPC,NULL,'0','1') CUST_STYLE_UPC_FLAG" _
                                & ",CUST_UPC CUST_UPC" _
                                & " from SOTORDR2" _
                                & " where ORDR_NO IN (Select ORDR_NO from SOTORDR1" _
                                & " where CUST_CODE = '" & CUST_CODE & "' and ORDR_CUST_PO = :PARM1)" _
                                & ") where CUST_STYLE_CODE IS NOT NULL"
                            ' TROUBLE USING SOTORDR2 BECAUSE WHEN WE SUBSTITUTE STYLES FOR CUSTOMER STYLES WE WINDUP WITH DUPS
                            ASCMAIN1.sql = "Select * from (" & vbCrLf _
                                & "Select DISTINCT '" & CUST_CODE & "' CUST_CODE" & vbCrLf _
                                & ",EDI_SLN_SKU CUST_STYLE_CODE" & vbCrLf _
                                & ",EDI_SLN_SIZE_DESC SIZE_DESC" & vbCrLf _
                                & ",EDI_SLN_PRICE CUST_PRICE" & vbCrLf _
                                & ",NULL STYLE_CODE" & vbCrLf _
                                & ",NULL COLOR_CODE" & vbCrLf _
                                & ",DECODE(EDI_SLN_UPC,NULL,'0','1') CUST_STYLE_UPC_FLAG" & vbCrLf _
                                & ",EDI_SLN_UPC CUST_UPC" & vbCrLf _
                                & " from EDT850T6" & vbCrLf _
                                & " where EDI_DOC_SEQ_NO IN (Select EDI_DOC_SEQ_NO from EDT850T1" & vbCrLf _
                                & " where (EDI_TP_QUAL, EDI_TP_ID) IN " & vbCrLf _
                                & "(Select Distinct EDI_TP_QUAL, EDI_TP_ID from EDTTRPM1 " & vbCrLf _
                                & " where EDI_DOC_NO = '850' and CUST_CODE = '" & CUST_CODE & "')" & vbCrLf _
                                & " and EDI_PO_NO = :PARM1)" & vbCrLf _
                                & ") where CUST_STYLE_CODE IS NOT NULL"
                            DT = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New String() {txtFindBy.Text})

                        Case "E"
                            ASCMAIN1.sql = "Select * from (" _
                                & "Select DISTINCT '" & CUST_CODE & "' CUST_CODE" _
                                & ",EDI_SLN_SKU CUST_STYLE_CODE" _
                                & ",EDI_SLN_SIZE_DESC SIZE_DESC" _
                                & ",EDI_SLN_PRICE CUST_PRICE" _
                                & ",NULL STYLE_CODE" _
                                & ",NULL COLOR_CODE" _
                                & ",DECODE(EDI_SLN_UPC,NULL,'0','1') CUST_STYLE_UPC_FLAG" _
                                & ",EDI_SLN_UPC CUST_UPC" _
                                & " from EDT850T6" _
                                & " where EDI_DOC_SEQ_NO IN (Select EDI_DOC_SEQ_NO from EDT850T1" _
                                & " where (EDI_TP_QUAL, EDI_TP_ID) IN " _
                                & "(Select Distinct EDI_TP_QUAL, EDI_TP_ID from EDTTRPM1 " _
                                & " where EDI_DOC_NO = '850' and CUST_CODE = '" & CUST_CODE & "')" _
                                & " and EDI_DOC_SEQ_NO = :PARM1)" _
                                & ") where CUST_STYLE_CODE IS NOT NULL"
                            DT = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New String() {txtFindBy.Text})
                            
                        Case "O"
                            ASCMAIN1.sql = "Select * from (" _
                               & "Select DISTINCT '" & CUST_CODE & "' CUST_CODE" _
                               & ",NVL(CUST_STYLE_CODE,CUST_SKU) CUST_STYLE_CODE" _
                               & ",CUST_SIZE_CODE SIZE_DESC" _
                               & ",ORDR_UNIT_PRICE CUST_PRICE" _
                               & ",STYLE_CODE" _
                               & ",COLOR_CODE" _
                               & ",DECODE(CUST_UPC,NULL,'0','1') CUST_STYLE_UPC_FLAG" _
                               & ",CUST_UPC CUST_UPC" _
                               & " from SOTORDR2" _
                               & " where ORDR_NO = :PARM1" _
                               & ") WHERE CUST_STYLE_CODE IS NOT NULL"
                            DT = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New String() {txtFindBy.Text})
                    End Select

                    dst.Tables("SOTCSTYO").Rows.Clear()
                    For Each row As DataRow In DT.Rows
                        Dim rowO As DataRow = dst.Tables("SOTCSTYO").Rows.Add(row.ItemArray)
                        Dim rowSOTCSTY1 As DataRow = dst.Tables("SOTCSTY1").Rows.Find(New String() {CUST_CODE, row.Item("CUST_STYLE_CODE")})
                        If rowSOTCSTY1 Is Nothing Then
                            rowO.Item("STATE") = "Added"
                        Else
                            If Val(rowO.Item("CUST_PRICE") & "") = Val(rowSOTCSTY1.Item("CUST_PRICE") & "") Then
                                rowO.Item("STATE") = "Existing"
                                rowO.Item("STYLE_CODE") = rowSOTCSTY1.Item("STYLE_CODE")
                                rowO.Item("COLOR_CODE") = rowSOTCSTY1.Item("COLOR_CODE")
                            Else
                                rowO.Item("STATE") = "New Price"
                            End If
                        End If
                    Next

                    splSOTCSTY1.Panel2Collapsed = False
                End If

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "CUST_CODE"
            '    If Not ScreenMode Then
            '        Load_SOTORDRX()
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Click_Command("View")

        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub
#End Region

    Sub Load_SOTCSTYX(Optional PARM1 As String = "", Optional CUST_CODE As String = "")
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Customers", "")

        Fill_Records("SOTCSTYX")
        Sort_grdColumns(grdSOTCSTYX, "CUST_CODE")
        grdSOTCSTYX.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub grdSOTCSTYX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTCSTYX.DoubleClickRow
        If Not ScreenMode Then
            Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value
            Click_Command("View")
        End If
    End Sub

#Region "grdSOTCSTY1"

    Private Sub grdSOTCSTY1_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCSTY1.AfterCellUpdate
        With grdSOTCSTY1.ActiveRow
            Select Case e.Cell.Column.Key
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Value & "")
                    If STYLE_CODE <> "" Then
                        .Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                        If COLOR_CODEs.Count = 1 Then
                            .Cells("COLOR_CODE").Value = COLOR_CODEs(0)
                        End If
                    End If

                Case "COLOR_CODE"
                    Dim COLOR_CODE As String = e.Cell.Value & ""
                    If COLOR_CODE <> "" Then
                        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                        If rowICTCOLR1 IsNot Nothing Then
                            .Cells("COLOR_DESC").Value = rowICTCOLR1.Item("COLOR_DESC")
                        End If
                    End If

            End Select
        End With
    End Sub

    Private Sub grdSOTCSTY1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCSTY1.AfterRowActivate

    End Sub

    Private Sub grdSOTCSTY1_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTCSTY1.AfterRowsDeleted

    End Sub

    Private Sub grdSOTCSTY1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCSTY1.AfterRowUpdate

    End Sub

    Private Sub grdSOTCSTY1_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTCSTY1.BeforeCellUpdate

    End Sub

    Private Sub grdSOTCSTY1_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTCSTY1.BeforeExitEditMode
        If grdSOTCSTY1.ActiveCell IsNot Nothing Then
            With grdSOTCSTY1.ActiveCell
                Select Case .Column.Key
                    Case "STYLE_CODE", "COLOR_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTCSTY1_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTCSTY1.BeforeRowsDeleted

    End Sub

    Private Sub grdSOTCSTY1_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCSTY1.BeforeRowUpdate

        Validate_Columns(grdSOTCSTY1, "STYLE_CODE", e.Cancel)
        If e.Cancel Then
            MsgBox("Issue with Style Code")
            Exit Sub
        End If
        If Not e.Cancel Then
            Validate_Columns(grdSOTCSTY1, "COLOR_CODE", e.Cancel)
            If e.Cancel Then
                MsgBox("Issue with Color Code, or Style-Color Setup")
                Exit Sub
            End If
        End If

        If e.Cancel = True Then
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("CUST_CODE").Value = CUST_CODE
            e.Row.Cells("STATE").Value = "Added"
            e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
            e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
        Else
            If e.Row.Cells("STATE").Value & "" <> "Added" Then
                e.Row.Cells("STATE").Value = "Edited"
                e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
                e.Row.Cells("LAST_DATE").Value = DATETIME_STAMP
            End If
        End If
    End Sub

    Private Sub grdSOTCSTY1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCSTY1.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "STYLE_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdSOTCSTY1, sql_where)

                Case "COLOR_CODE"
                    Dim sql_where As String = "COLOR_CODE IN (Select COLOR_CODE from ICTSTYC1 where STYLE_CODE ='" & grdSOTCSTY1.ActiveRow.Cells("STYLE_CODE").Value & "')"

                    grdClickCellButton(grdSOTCSTY1, sql_where)
            End Select
        End With

    End Sub
#End Region

    Sub Validate_Columns(grd As UltraWinGrid.UltraGrid, COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grd.ActiveRow
            Select Case COLUMN_NAME
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = ""
                    If Trim(.Cells("STYLE_CODE").Value & "") <> "" Then
                        STYLE_CODE = Validate_Style(.Cells("STYLE_CODE").Value & "")
                    End If
                    Cancel = (STYLE_CODE = "")

                Case "COLOR_CODE"
                    If .Cells("COLOR_CODE").Value & "" <> "" Then
                        If Not COLOR_CODEs.Contains(.Cells("COLOR_CODE").Value & "") Then
                            Cancel = True
                        End If
                    Else
                        Cancel = True
                    End If
            End Select
        End With
    End Sub

    Function Validate_Style(STYLE_CODE_z As String) As String
        Dim EMsg As String = ""
        If STYLE_CODE_z = "" Then Return ""

        Dim STYLE_CODE As String = ""
        rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE_z)

        If rowICTSTYL1 Is Nothing Then
            EMsg = "Style is Not on File" & vbCrLf
        Else
            If rowICTSTYL1.Item("STYLE_STATUS") & "" <> "A" Then
                EMsg = "Item Status is not Active" & vbCrLf
            End If
            If rowICTSTYL1.Item("STYLE_UOM") & "" = "" Then
                EMsg = "Item does not have a valid Unit of Measure" & vbCrLf
            End If
            If rowICTSTYL1.Item("SALES_DIVISION_CODE") & "" = "" Then
                EMsg = "Item does not have a valid Division Code" & vbCrLf
            End If
        End If

        If EMsg = "" Then
            COLOR_CODEs.Clear()
            Fill_Records("ICTCOLRS", STYLE_CODE_z)
            For Each row As DataRow In dst.Tables("ICTCOLRS").Select("")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                COLOR_CODEs.Add(COLOR_CODE)
            Next
        End If

        If EMsg <> "" And grdSOTCSTY1.ActiveRow.IsAddRow Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Style Code Entered is Invalid because ...")
        Else
            If EMsg = "" Then
                STYLE_CODE = rowICTSTYL1.Item(0)
            End If
        End If
        Return STYLE_CODE
    End Function

    Private Sub txtFind_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtFindBy.KeyDown
        'If e.KeyCode = Keys.Enter Then

        '    grdSOTCSTYO.Visible = True
        'End If
    End Sub

    Private Sub cmdUpdate_Click(sender As System.Object, e As System.EventArgs) Handles cmdUpdate.Click
        If Not chkAddNewStyles.Checked And Not chkUpdatePrices.Checked Then
            MsgBox("You must select one of 'Add New ...' or 'Update Existing ...'", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating")

        Dim sqlw As String = ""
        If chkAddNewStyles.Checked Then sqlw &= " or  STATE = 'Added'"
        If chkUpdatePrices.Checked Then sqlw &= " or  STATE = 'New Price' or  STATE = 'Chg Style/Color'"

        If chkSelectedOnly.Checked Then sqlw = " and SELECTED = '1' and (" & Mid(sqlw, 5) & ")"

        For Each row As DataRow In dst.Tables("SOTCSTY1").Select("SELECTED = '1'")
            row.Item("SELECTED") = "0"
        Next

        Dim c As Integer = 0
        For Each row As DataRow In dst.Tables("SOTCSTYO").Select(Mid(sqlw, 5))
            If row.Item("STATE") = "Added" Then
                Dim rowSOTCSTY1 As DataRow = dst.Tables("SOTCSTY1").Rows.Add(row.ItemArray)
                rowSOTCSTY1.Item("SELECTED") = "1"
                rowSOTCSTY1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowSOTCSTY1.Item("INIT_DATE") = DATETIME_STAMP
                rowSOTCSTY1.Item("STATE") = "Added"
                Dim SIZE_DESC As String = rowSOTCSTY1.Item("SIZE_DESC") & ""
                If SIZE_DESC.Length > 0 Then
                    SIZE_DESC = SIZE_DESC.ToUpper
                    For i As Integer = 1 To SIZE_DESC.Length
                        If Mid(SIZE_DESC, i, 1) >= "A" And Mid(SIZE_DESC, i, 1) <= "Z" _
                        Or Mid(SIZE_DESC, i, 1) >= "0" And Mid(SIZE_DESC, i, 1) <= "9" Then
                        Else
                            Mid(SIZE_DESC, i, 1) = "#"
                        End If
                    Next
                    SIZE_DESC = Replace(SIZE_DESC, "#", "")
                    rowSOTCSTY1.Item("SIZE_DESC") = SIZE_DESC
                End If

            ElseIf row.Item("STATE") = "New Price" Or row.Item("STATE") = "Chg Style/Color" Then
                Dim rowSOTCSTY1 As DataRow = dst.Tables("SOTCSTY1").Rows.Find(New String() {CUST_CODE, row.Item("CUST_STYLE_CODE")})
                rowSOTCSTY1.Item("CUST_PRICE") = row.Item("CUST_PRICE")
                rowSOTCSTY1.Item("SELECTED") = "1"
                rowSOTCSTY1.Item("STYLE_CODE") = row.Item("STYLE_CODE")
                rowSOTCSTY1.Item("COLOR_CODE") = row.Item("COLOR_CODE")
                rowSOTCSTY1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowSOTCSTY1.Item("LAST_DATE") = DATETIME_STAMP
                rowSOTCSTY1.Item("STATE") = "Edited"
            End If
            c += 1
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        MsgBox(CStr(c) & " Customer Styles Updated/Added", MsgBoxStyle.OkOnly, "Verification")

        If c <> 0 Then
            Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Selected Only"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = True

            splSOTCSTY1.Panel2Collapsed = True
        End If
    End Sub

#Region "grdSOTCSTYO"

    Private Sub grdSOTCSTYO_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCSTYO.AfterCellUpdate
        With grdSOTCSTYO.ActiveRow
            Select Case e.Cell.Column.Key
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Value)
                    If STYLE_CODE <> "" Then
                        .Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                        If COLOR_CODEs.Count = 1 Then
                            .Cells("COLOR_CODE").Value = COLOR_CODEs(0)
                        End If
                    End If

                Case "COLOR_CODE"
                    Dim COLOR_CODE As String = e.Cell.Value & ""
                    If COLOR_CODE <> "" Then
                        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                        If rowICTCOLR1 IsNot Nothing Then
                            .Cells("COLOR_DESC").Value = rowICTCOLR1.Item("COLOR_DESC")
                        End If
                    End If

            End Select
        End With
    End Sub

    Private Sub grdSOTCSTYO_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCSTYO.AfterRowActivate

    End Sub

    Private Sub grdSOTCSTYO_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTCSTYO.AfterRowsDeleted

    End Sub

    Private Sub grdSOTCSTYO_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCSTYO.AfterRowUpdate
        STYLE_CODE_last = e.Row.Cells("STYLE_CODE").Value
        COLOR_CODE_last = e.Row.Cells("COLOR_CODE").Value
    End Sub

    Private Sub grdSOTCSTYO_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTCSTYO.BeforeCellUpdate

    End Sub

    Private Sub grdSOTCSTYO_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTCSTYO.BeforeExitEditMode
        If grdSOTCSTYO.ActiveCell IsNot Nothing Then
            With grdSOTCSTYO.ActiveCell
                Select Case .Column.Key
                    Case "STYLE_CODE", "COLOR_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTCSTYO_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTCSTYO.BeforeRowsDeleted

    End Sub

    Private Sub grdSOTCSTYO_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCSTYO.BeforeRowUpdate

        Validate_Columns(grdSOTCSTYO, "STYLE_CODE", e.Cancel)
        If Not e.Cancel Then
            Validate_Columns(grdSOTCSTYO, "COLOR_CODE", e.Cancel)
        End If

        If e.Cancel = True Then
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            'e.Row.Cells("CUST_CODE").Value = CUST_CODE
            'e.Row.Cells("STATE").Value = "Added"
            'e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
            'e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
        Else
            If e.Row.Cells("STATE").Value = "Added" Then
            Else
                e.Row.Cells("STATE").Value = "Chg Style/Color"
            End If
        End If
    End Sub

    Private Sub grdSOTCSTYO_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCSTYO.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "STYLE_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdSOTCSTYO, sql_where)

                Case "COLOR_CODE"
                    Dim sql_where As String = "COLOR_CODE IN (Select COLOR_CODE from ICTSTYC1 where STYLE_CODE ='" & grdSOTCSTYO.ActiveRow.Cells("STYLE_CODE").Value & "')"

                    grdClickCellButton(grdSOTCSTYO, sql_where)
            End Select
        End With

    End Sub

    Private Sub grdSOTCSTYO_DoubleClickCell(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdSOTCSTYO.DoubleClickCell
        If e.Cell.Column.Key = "STYLE_CODE" Then
            If e.Cell.Value & "" = "" Then
                e.Cell.Value = STYLE_CODE_last
                e.Cell.Row.Cells("COLOR_CODE").Value = COLOR_CODE_last
            End If
        End If
    End Sub

    Private Sub grdSOTCSTYO_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTCSTYO.InitializeRow
        If e.Row.Cells("STATE").Value = "Added" Then
            e.Row.Cells("STATE").Appearance.BackColor = Drawing.Color.LightGreen
        ElseIf e.Row.Cells("STATE").Value = "Existing" Then
            e.Row.Cells("STATE").Appearance.BackColor = Drawing.Color.Empty
        ElseIf e.Row.Cells("STATE").Value = "New Price" Then
            e.Row.Cells("CUST_PRICE").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("STATE").Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub
#End Region

    Private Sub txtFindBy_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtFindBy.ValueChanged

    End Sub

    Public Overrides Sub Excel_Import_Pre_Process _
    (ByVal grd As UltraWinGrid.UltraGrid, _
     Optional ByRef load_by_table As Boolean = False, _
     Optional ByRef load_handled As Boolean = False, _
     Optional ByRef F As ASFEXCL1 = Nothing)

    End Sub

    'Public Overrides Sub Excel_Import_Post_Process _
    '(ByVal grd As UltraWinGrid.UltraGrid)

    'End Sub

    Sub Show_UPCs()
        If ScreenMode Then
            ASCMAIN1.Progress("Now Creating UPC List fo Styles for " & CUST_CODE, "")
            Fill_Records("ICTUPCHX", CUST_CODE)
            Sort_grdColumns(grdICTUPCHX, "UPC_CODE")
            grdICTUPCHX.Visible = True
            grdICTUPCHX.Text = "All UPCs" & IIf(CUST_CODE = "", "", " for " & CUST_CODE)
            ASCMAIN1.Progress("", "")
        Else
            ASCMAIN1.Progress("Now Creating UPC List", "")

            ASCMAIN1.sql = "Select X.UPC_CODE, X.STYLE_CODE, X.COLOR_CODE, X.STYLE_DESC, X.CUST_CODE" & vbCrLf _
                & ", 0 CUST_PRICE, X.SIZE_CODE SIZE_DESC, NULL CUST_STYLE_CODE" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select ICVLUPC1.*, ICTSTYL1.STYLE_DESC, ICTSTYL1.CUST_CODE" & vbCrLf _
                & " from ICVLUPC1,ICTSTYL1" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE (+) = ICVLUPC1.STYLE_CODE" & vbCrLf _
                & ") X"

            Fill_Records("ICTUPCHX", "", True, ASCMAIN1.sql)

            Sort_grdColumns(grdICTUPCHX, "UPC_CODE")
            grdICTUPCHX.Visible = True
            grdICTUPCHX.Text = "All UPCs"
            ASCMAIN1.Progress("", "")
        End If

        With grdICTUPCHX.DisplayLayout.Bands(0)
            .Columns("CUST_PRICE").Hidden = Not ScreenMode
            .Columns("CUST_STYLE_CODE").Hidden = Not ScreenMode
            If ScreenMode Then
                .Columns("STYLE_DESC").Width = 130
            Else
                .Columns("STYLE_DESC").Width = 260
            End If
        End With


    End Sub
End Class