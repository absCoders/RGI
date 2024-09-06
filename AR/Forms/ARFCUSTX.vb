Public Class ARFCUSTX

    ' formatting of store no

    Dim rowARTCUST1 As DataRow
    Dim ADDRESS() As String = {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3", _
             "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY", _
             "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_EMAIL", "GLOBAL_LOCATION_NUMBER"}
    Dim ADDRESS_EXT() As String = {"CUST_ADDR3",
          "CUST_COUNTRY",
         "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_EMAIL"}

    Dim LAST_CHANGE_COLUMN_NAME As String
    Dim LAST_CHANGE_CELL_VALUE As String
    Dim COPY_VALUE_clipboard As String
    Dim COLUMN_NAME_clipboard As String

    Dim CUST_CODE As String
    Dim CUST_DC_NOs_to_delete As New List(Of String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "ARTCUST1", "*")
            Create_TDA(.Tables.Add, "ARTCUST2", "*", 1)
            .Tables("ARTCUST2").Columns.Add("CUST_DC_IND")
            .Tables("ARTCUST2").Columns("CUST_DC_IND").DefaultValue = "0"
            Create_TDA(.Tables.Add, "ARTCUST3", "*", 1)

            If ASCMAIN1.CLIENT = "VAN" Then
                ASCMAIN1.sql = "Select ARTCUST2.* FROM ARTCUST2 WHERE ARTCUST2.CUST_CODE = :PARM1"
                '    Create_TDA(.Tables.Add, "ARTCUSTM", "**", 0, True, "V")

                Create_TDA(.Tables.Add, "ARTCUSTM", "**", 0, False, "V", 3)

                With .Tables.Add("ERROR_TBL")
                    .Columns.Add("ERROR_CODE", GetType(System.String))
                    .Columns.Add("ERROR_DETAIL", GetType(System.String))
                End With
            End If


        End With
        grdARTCUSTM.DataSource = dst.Tables("ARTCUSTM")
        Create_Summary(grdARTCUSTM, "CUST_ADDR_CODE", "Count")

        grdARTCUST2.DataSource = dst.Tables("ARTCUST2")

        Create_Summary(grdARTCUST2, "CUST_ADDR_CODE", "Count")

        With grdARTCUST2.DisplayLayout.Bands("ARTCUST2")
            .Columns("CUST_ADDR_CODE").Header.Fixed = True
            'For Each COLUMN_NAME As String In New String() {"CUST_ADDR_TYPE", "CUST_RANK", "CUST_ADDR_ROUTING_INST", "CUST_ROUTING_INST", "CUST_ADDR_CODE_DOMESTIC", "CONSIGNEE_BILLED", "MIN_ORDR_QTY", "MIN_ORDR_AMT", "MIN_STYLE_QTY", "MIN_STYLE_AMT", "FDX_ACCT_NO", "CUST_ADDR_GROUP"}
            For Each COLUMN_NAME As String In New String() {"CUST_ADDR_TYPE", "CUST_ADDR_NAME"}
                .Columns(COLUMN_NAME).Hidden = True
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
        End With

        For Each COLUMN_NAME As String In ADDRESS
            If ASCMAIN1.CLIENT = "VANXXX" And COLUMN_NAME = "CUST_ADDR3" Then
            Else

                With grdARTCUST2.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Appearance
                    .BackColor2 = Drawing.Color.Yellow
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
                If ADDRESS_EXT.Contains(COLUMN_NAME) Then grdARTCUST2.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True
            End If

        Next
        grdARTCUST2.DisplayLayout.Bands(0).Columns("CUST_ADDR_CODE").CellAppearance.BackColor = Drawing.Color.Beige
        grdARTCUST2.DisplayLayout.Bands(0).Columns("STAX_CODE").Hidden = Not (ASCMAIN1.CLIENT = "NYA")

        ASCMAIN1.Add_Value_List(grdARTCUST2, "CUST_ADDR_STATUS")

        Dim udd As New UltraWinGrid.UltraDropDown
        Dim DVW As DataView = New DataView(dst.Tables("ARTCUST2"), "CUST_DC_IND = '1'", "CUST_ADDR_CODE", DataViewRowState.CurrentRows)
        udd.DataSource = DVW
        For Each GC As UltraWinGrid.UltraGridColumn In udd.DisplayLayout.Bands(0).Columns
            If GC.Key <> "CUST_ADDR_CODE" Then
                GC.Hidden = True
            Else
                GC.Header.Caption = "DC No"
            End If
        Next
        udd.ValueMember = "CUST_ADDR_CODE"
        With grdARTCUST2.DisplayLayout.Bands(0).Columns("CUST_DC_NO")
            .ValueList = udd
            .ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
        End With

        grpStores.Visible = False

        AUDIT.Add("ARTCUST2", "NED")

        ASCMAIN1.Add_Value_List(grdARTCUST2, "CUST_ADDR_STATUS", Nothing, New String() {":", "A:Active", "I:Inactive", "C:Closed"})


        If ASCMAIN1.CLIENT = "VAN" Then
            With UltraExplorerBar1.Groups("Special Functions")
                .Visible = False
                With .Items.Add("Update Addresses/Excel")
                    .Text = .Key
                End With
            End With
        End If

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("CUST_CODE")

                CUST_CODE = Absx1.txtFor("CUST_CODE").Text

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ARTCUST1", CUST_CODE) Then
                        Exit Sub
                    End If
                End If

            Case "Update"
                For Each row As DataRow In dst.Tables("ARTCUST2").Select("CUST_DC_NO is Not Null")
                    Dim CUST_ADDR_CODE As String = row.Item("CUST_ADDR_CODE")
                    Dim CUST_DC_NO As String = row.Item("CUST_DC_NO")
                    Dim rowARTCUST2_DC As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "DC", CUST_DC_NO})
                    If rowARTCUST2_DC Is Nothing Then
                        EMsg &= "Invalid DC Specified for Store " & CUST_ADDR_CODE
                    Else
                        Dim rowARTCUST2_MK As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "MK", CUST_DC_NO})
                        If rowARTCUST2_MK Is Nothing OrElse rowARTCUST2_MK.Item("CUST_DC_IND") & "" <> "1" Then
                            EMsg &= "Invalid DC Specified for Store " & CUST_ADDR_CODE
                        End If
                    End If
                Next

                For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select("CUST_ADDR_TYPE = 'MK'", "CUST_ADDR_CODE")
                    Dim CUST_ADDR_CODE As String = rowARTCUST2.Item("CUST_ADDR_CODE")

                    Dim CUST_NAME As String = rowARTCUST2.Item("CUST_NAME") & ""
                    Dim CUST_ADDR1 As String = rowARTCUST2.Item("CUST_ADDR1") & ""
                    Dim CUST_CITY As String = rowARTCUST2.Item("CUST_CITY") & ""
                    Dim CUST_STATE As String = rowARTCUST2.Item("CUST_STATE") & ""
                    Dim CUST_ZIP_CODE As String = rowARTCUST2.Item("CUST_ZIP_CODE") & ""
                    Dim CUST_COUNTRY As String = rowARTCUST2.Item("CUST_COUNTRY") & ""

                    Dim STAX_CODE As String = rowARTCUST2.Item("STAX_CODE") & ""

                    If STAX_CODE <> "" AndAlso LookUp("ARTSTAX1", STAX_CODE) Is Nothing Then
                        EMsg &= vbCr & "Invalid Sales Tax Code (see Store " & CUST_ADDR_CODE & ")"
                    End If

                    If CUST_NAME = "" Or CUST_ADDR1 = "" Then
                        EMsg &= vbCr & "Name and Address Line 1 is mandatory (see Store " & CUST_ADDR_CODE & ")"
                    End If

                    If CUST_COUNTRY <> "" And CUST_COUNTRY <> "USA" Then
                        Dim rowTATCNTRY As DataRow = LookUp("TATCNTRY", CUST_COUNTRY)
                        If rowTATCNTRY Is Nothing Then
                            EMsg &= vbCr & "Invalid Country for Store " & CUST_ADDR_CODE
                        Else
                            If CUST_CITY = "" Then
                                EMsg &= vbCr & "City is mandatory for International Addresses (see Store " & CUST_ADDR_CODE & ")"
                            End If
                        End If
                    Else

                        If CUST_CITY = "" Or CUST_STATE = "" Or CUST_ZIP_CODE = "" Then
                            EMsg &= vbCr & "City, State and Zip Code is mandatory (see Store " & CUST_ADDR_CODE & ")"
                        End If
                    End If

                    If Split(EMsg, vbCr).Length > 10 Then
                        EMsg &= vbCr & "..."
                        Exit For
                    End If
                Next

                Get_DCs_to_Delete()

                If EMsg = "" Then
                    If CUST_DC_NOs_to_delete.Count > 0 Then
                        If MsgBox("Please Note - the following Address Codes will no longer be defined as DCs:" _
                                  & vbCrLf & " " & Join(CUST_DC_NOs_to_delete.ToArray, ","),
                                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub
                    End If
                End If

            Case "Update Excel"
                If dst.Tables("ERROR_TBL").Rows.Count <> 0 Then
                    EMsg &= vbCr & "There are errors to address in Excel Import. Cannot Update"

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

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Cancel Excel", "Done"
                Call SET_EXCEL_CONTROLS(False)
                Mode_Settings(False)

            Case "Update Excel", "Done"
                Update_Excel()
                Call SET_EXCEL_CONTROLS(False)
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update Excel").Settings.Enabled = True
                .Groups("Screen Control").Items("Cancel Excel").Settings.Enabled = True
                .Groups("Screen Control").Items("Update Excel").Visible = False
                .Groups("Screen Control").Items("Cancel Excel").Visible = False
                .Groups("Display Options").Visible = ScreenMode
                .Groups("Customer Info").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grpStores.Visible = tf
        grdARTCUST2.Visible = True
        grdARTCUSTM.Visible = False



        If ScreenMode Then
            If ASCMAIN1.CLIENT = "VAN" And (CUST_CODE = "WALMART" Or CUST_CODE = "SAMSCLUB") Then
                With UltraExplorerBar1.Groups("Special Functions")
                    .Visible = True
                End With
                ' OPTION IN EXCEL BASED ON CUST
                With optEXCEL.ValueList
                    If CUST_CODE = "WALMART" Then
                        .ValueListItems(0).DisplayText = "Walmart"
                        .ValueListItems(1).DisplayText = "Walmart DC"
                    ElseIf CUST_CODE = "SAMSCLUB" Then
                        .ValueListItems(0).DisplayText = "SamsClub"
                        .ValueListItems(1).DisplayText = "SamsClub DC"
                    End If
                End With

            End If
        Else
            Clear_Record()
            With UltraExplorerBar1.Groups("Special Functions")
                .Visible = False
            End With
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ARTCUST1", "ARTCUST2", "ARTCUST3"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        If ASCMAIN1.CLIENT = "VAN" Then
            dst.Tables("ARTCUSTM").Rows.Clear()
            dst.Tables("ERROR_TBL").Rows.Clear()
        End If

        Absx1.txtFor("CUST_CODE").Text = ""

        LAST_CHANGE_CELL_VALUE = ""
        LAST_CHANGE_COLUMN_NAME = ""
        COPY_VALUE_clipboard = ""
        COLUMN_NAME_clipboard = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)
        rowARTCUST1 = Fill_Record("ARTCUST1", HFs("CUST_CODE"))
        Fill_Records("ARTCUST2", HFs("CUST_CODE"))
        Fill_Records("ARTCUST3", HFs("CUST_CODE"))
        If ASCMAIN1.CLIENT = "VAN" Then
            '  Fill_Records("ARTCUSTM", HFs("CUST_CODE"))
        End If

        EnforceConstraints(True)

        Dim CUST_DC_NOs_nogood As New List(Of String)
        ' Each MK that references a DC may be referencing an address that was set up as an MK, but not as a DC
        For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select("CUST_ADDR_TYPE = 'MK' and CUST_DC_NO is Not Null")
            Dim CUST_ADDR_CODE As String = rowARTCUST2.Item("CUST_DC_NO")
            Dim rowARTCUST2_DC As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "DC", CUST_ADDR_CODE})
            If rowARTCUST2_DC Is Nothing Then
                Dim rowARTCUST2_MK As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "MK", CUST_ADDR_CODE})
                If rowARTCUST2_MK IsNot Nothing Then
                    rowARTCUST2_DC = dst.Tables("ARTCUST2").NewRow
                    rowARTCUST2_DC.ItemArray = rowARTCUST2_MK.ItemArray
                    rowARTCUST2_DC.Item("CUST_ADDR_TYPE") = "DC"
                    dst.Tables("ARTCUST2").Rows.Add(rowARTCUST2_DC)
                Else
                    CUST_DC_NOs_nogood.Add(CUST_ADDR_CODE)
                End If
            End If
        Next

        If CUST_DC_NOs_nogood.Count > 0 Then
            MsgBox("The following Address Codes are referenced as DCs but no Address Record exists" _
                   & vbCrLf & Join(CUST_DC_NOs_nogood.ToArray, ","),
                   MsgBoxStyle.OkOnly, "Data Integrity Issue")
        End If


        ' Each DC record must also have an MK record
        For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select("CUST_ADDR_TYPE = 'DC'")
            Dim CUST_ADDR_CODE As String = rowARTCUST2.Item("CUST_ADDR_CODE")
            Dim rowARTCUST2_MK As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "MK", CUST_ADDR_CODE})
            If rowARTCUST2_MK Is Nothing Then
                rowARTCUST2_MK = dst.Tables("ARTCUST2").NewRow
                rowARTCUST2_MK.ItemArray = rowARTCUST2.ItemArray
                rowARTCUST2_MK.Item("CUST_ADDR_TYPE") = "MK"
                dst.Tables("ARTCUST2").Rows.Add(rowARTCUST2_MK)
            End If
            rowARTCUST2_MK.Item("CUST_DC_IND") = "1"
        Next

        Sort_grdColumns(grdARTCUST2, "CUST_ADDR_CODE")
        optShow.Value = "A"
        Setup_ARTCUST2()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        rowARTCUST1.Item("LAST_DATE") = DATETIME_STAMP
        rowARTCUST1.Item("LAST_OPER") = ASCMAIN1.USER_ID

        For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.CurrentRows)
            If rowARTCUST2.RowState = DataRowState.Added Then
                rowARTCUST2.Item("LAST_DATE") = DATETIME_STAMP
                rowARTCUST2.Item("LAST_OPER") = ASCMAIN1.USER_ID
            End If

            rowARTCUST2.Item("LAST_DATE") = DATETIME_STAMP
            rowARTCUST2.Item("LAST_OPER") = ASCMAIN1.USER_ID
        Next

        For Each CUST_ADDR_CODE As String In CUST_DC_NOs_to_delete
            Dim row As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "DC", CUST_ADDR_CODE})
            row.Delete()
        Next

        ' Fix all DC addresses (which were never visible in the form) so that they reflect changes made to the corresponding MK record
        Dim SQLW As String = "CUST_CODE = '" & CUST_CODE & "' and CUST_ADDR_TYPE = 'MK' and CUST_DC_IND = '1'"
        For Each rowARTCUST2_MK As DataRow In dst.Tables("ARTCUST2").Select(SQLW)
            Dim CUST_ADDR_CODE As String = rowARTCUST2_MK.Item("CUST_ADDR_CODE")
            Dim rowARTCUST2_DC As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "DC", CUST_ADDR_CODE})
            If rowARTCUST2_DC Is Nothing Then
                rowARTCUST2_DC = dst.Tables("ARTCUST2").NewRow
                rowARTCUST2_DC.ItemArray = rowARTCUST2_MK.ItemArray
                rowARTCUST2_DC.Item("CUST_ADDR_TYPE") = "DC"
                dst.Tables("ARTCUST2").Rows.Add(rowARTCUST2_DC)
            Else
                For Each DC As DataColumn In dst.Tables("ARTCUST2").Columns
                    If DC.ColumnName <> "CUST_ADDR_TYPE" Then
                        rowARTCUST2_DC.Item(DC.ColumnName) = rowARTCUST2_MK.Item(DC.ColumnName)
                    End If
                Next
            End If
        Next

        ' not being used - using the sql below
        'dst.Tables("ARTCUST3").Rows.Clear()
        'For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select("CUST_ADDR_TYPE = 'MK' and CUST_DC_NO is Not Null")
        '    Dim CUST_ADDR_CODE = rowARTCUST2.Item("CUST_ADDR_CODE")
        '    Dim CUST_DC_NO = rowARTCUST2.Item("CUST_DC_NO")
        '    Dim rowARTCUST3 As DataRow = dst.Tables("ARTCUST3").NewRow
        '    rowARTCUST3.Item("CUST_CODE") = CUST_CODE
        '    rowARTCUST3.Item("CUST_ADDR_TYPE") = "MK"
        '    rowARTCUST3.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
        '    rowARTCUST3.Item("CUST_CODE") = CUST_CODE
        '    rowARTCUST3.Item("CUST_ADDR_TYPE2") = "DC"
        '    rowARTCUST3.Item("CUST_ADDR_CODE2") = CUST_DC_NO
        '    dst.Tables("ARTCUST3").Rows.Add(rowARTCUST3)
        'Next


        BeginTrans()

        Update_Record_TDA("ARTCUST1")
        Update_Record_TDA("ARTCUST2")
        'Update_Record_TDA("ARTCUST3", "CUST_CODE = '" & CUST_CODE & "'")

        ASCMAIN1.sql = "Delete from ARTCUST3" & vbCrLf _
            & " where CUST_CODE = '" & CUST_CODE & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into ARTCUST3" & vbCrLf _
            & " Select CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE, 'DC', CUST_DC_NO" & vbCrLf _
            & " from ARTCUST2" & vbCrLf _
            & " where CUST_ADDR_TYPE = 'MK'" & vbCrLf _
            & "   and CUST_DC_NO is Not Null" & vbCrLf _
            & "   and CUST_CODE = '" & CUST_CODE & "'"
        ASCDATA1.ExecuteSQL()

        CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        '  Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub
#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdARTCUST2, "SSSSSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Show Vertically", "Add New Stores", "Show Full Address", "Clear Column", "Copy Value and Paste to All Stores", "Copy Value to Clipboard", "Paste Value to Selected Stores")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
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

        Select Case grd.Name
            Case "grdARTCUST2"

                tlb_btn = DirectCast(tlb_pop.Tools("Copy Value and Paste to All Stores"), UltraWinToolbars.ButtonTool)
                If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    tlb_btn.SharedProps.Caption = "Copy Value and Paste to All Stores"
                    tlb_btn.SharedProps.Visible = True
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Copy Value to Clipboard"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (grd.Selected.Rows.Count = 0)
                If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Or grd.ActiveCell Is Nothing Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    tlb_btn.SharedProps.Caption = "Copy '" & grd.ActiveCell.Value & "' to Clipboard"
                    tlb_btn.SharedProps.Visible = True
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Paste Value to Selected Stores"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (COLUMN_NAME_clipboard <> "") And grd.Selected.Rows.Count > 0
                tlb_btn.SharedProps.Caption = "Paste '" & COPY_VALUE_clipboard & "' to Selected Stores"
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else

            Select Case e.SourceControl.Name
                Case "grdARTCUST2"
                    'If grdARTCUST2.Tag = "" Then
                    '    e.Cancel = True
                    'End If
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Show Vertically"
                tlb_sbt = DirectCast(tlb.Tools("Show Vertically"), UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    Dim tlb_sbt2 As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Add New Stores"), UltraWinToolbars.StateButtonTool)
                    tlb_sbt2.Checked = False
                End If
                grdARTCUST2.DisplayLayout.Bands(0).CardView = tlb_sbt.Checked

            Case "Add New Stores"
                tlb_sbt = DirectCast(tlb.Tools("Add New Stores"), UltraWinToolbars.StateButtonTool)

                If tlb_sbt.Checked Then
                    Dim tlb_sbt2 As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Vertically"), UltraWinToolbars.StateButtonTool)
                    tlb_sbt2.Checked = False
                End If

                If tlb_sbt.Checked Then
                    grdARTCUST2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                Else
                    grdARTCUST2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                End If

            Case "Show Full Address"
                tlb_sbt = DirectCast(tlb.Tools("Show Full Address"), UltraWinToolbars.StateButtonTool)
                With grdARTCUST2.DisplayLayout.Bands(0)
                    For Each COLUMN_NAME As String In ADDRESS_EXT
                        .Columns(COLUMN_NAME).Hidden = Not tlb_sbt.Checked
                    Next
                End With

            Case "Paste Value to Selected Stores"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells(COLUMN_NAME_clipboard).Value = COPY_VALUE_clipboard
                    grow.Update()
                Next
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Clear Column"
                If grd.ActiveCell IsNot Nothing Then
                    Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                    If COLUMN_NAME = "" Then Exit Sub
                    If COLUMN_NAME = "CUST_ADDR_CODE" Then Exit Sub
                    For Each row As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.CurrentRows)
                        row.Item(COLUMN_NAME) = DBNull.Value
                    Next
                End If

            Case "Copy Value and Paste to All Stores"
                Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                If COLUMN_NAME = "" Then Exit Sub
                If grdARTCUST2.ActiveRow Is Nothing OrElse grdARTCUST2.ActiveRow.IsAddRow OrElse Not grdARTCUST2.ActiveRow.IsDataRow Then Exit Sub
                Dim COPY_VALUE As String = grdARTCUST2.ActiveRow.Cells(COLUMN_NAME).Value & ""
                For Each row As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.CurrentRows)
                    row.Item(COLUMN_NAME) = COPY_VALUE
                Next

            Case "Copy Value to Clipboard"
                Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                If COLUMN_NAME = "" Then Exit Sub
                If COLUMN_NAME = "CUST_ADDR_CODE" Then
                    MsgBox("Cannot Copy and Paste Store Numbers")
                    Exit Sub
                End If
                If grdARTCUST2.ActiveRow Is Nothing OrElse grdARTCUST2.ActiveRow.IsAddRow OrElse Not grdARTCUST2.ActiveRow.IsDataRow Then Exit Sub
                COPY_VALUE_clipboard = grdARTCUST2.ActiveRow.Cells(COLUMN_NAME).Value
                COLUMN_NAME_clipboard = COLUMN_NAME

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Call Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "grdARTCUST2"

    Private Sub grdARTCUST2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUST2.AfterCellUpdate
        If e.Cell.Column.Key = "CUST_DC_IND" Then
            grdARTCUST2.UpdateData()
        End If

        If e.Cell.Value & "" = "" Then
            LAST_CHANGE_CELL_VALUE = e.Cell.Value & ""
            LAST_CHANGE_COLUMN_NAME = e.Cell.Column.Key
        End If
    End Sub

    Private Sub grdARTCUST2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUST2.BeforeRowUpdate
        If e.Row.IsAddRow Then
            If Trim(e.Row.Cells("CUST_ADDR_CODE").Value & "") = "" Then
                e.Cancel = True
                MsgBox("Cannot use a blank Store No", vbOKOnly, "Update Denied")
            Else

                e.Row.Cells("CUST_CODE").Value = HFs("CUST_CODE")
                e.Row.Cells("CUST_ADDR_TYPE").Value = "MK"
            End If
        End If
    End Sub

    Private Sub grdARTCUST2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUST2.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdARTCUST2.ActiveCell.Column.Key
            Case "CUST_STATE"
                grdClickCellButton(grdARTCUST2, sql_where, False, "STATE_CODE")

            Case "CUST_COUNTRY"
                grdClickCellButton(grdARTCUST2, sql_where, False, "COUNTRY_CODE")

            Case "STAX_CODE"
                grdClickCellButton(grdARTCUST2, sql_where, False, "STAX_CODE")
        End Select


    End Sub

    Private Sub grdARTCUST2_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdARTCUST2.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In grdARTCUST2.Selected.Rows
            Dim CUST_ADDR_CODE As String = grow.Cells("CUST_ADDR_CODE").Value
            If dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "MK", CUST_ADDR_CODE}).RowState = DataRowState.Added Then
            Else
                MsgBox("Cannot Delete Existing Store Records", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                e.Cancel = True
                Exit For
            End If
        Next
    End Sub

    Private Sub grdARTCUST2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCUST2.AfterRowActivate
        If grdARTCUST2.ActiveRow.IsAddRow Then
            grdARTCUST2.DisplayLayout.Bands(0).Columns("CUST_ADDR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdARTCUST2.DisplayLayout.Bands(0).Columns("CUST_ADDR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Private Sub grdARTCUST2_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdARTCUST2.BeforeCellUpdate

    End Sub

    Private Sub grdARTCUST2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCUST2.AfterExitEditMode
        With grdARTCUST2
            Select Case .ActiveCell.Column.Key
                Case "CUST_ADDR_CODE"
                    If .ActiveCell.Text <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(.ActiveCell.Text, .ActiveCell.Column.Key)
                        If IsNumeric(.ActiveCell.Value & "") Then
                            .ActiveCell.Value = Format(Val(.ActiveCell.Value & ""), "000000")
                        End If
                    End If
            End Select
        End With
    End Sub

    Private Sub grdARTCUST2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdARTCUST2.AfterRowUpdate
        If e.Row.IsAddRow Then

        End If
    End Sub

    Private Sub grdARTCUST2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles grdARTCUST2.KeyPress

    End Sub

    Private Sub grdARTCUST2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles grdARTCUST2.KeyDown
        If e.KeyData = Keys.Delete Then
            If grdARTCUST2.ActiveCell IsNot Nothing Then
                If grdARTCUST2.ActiveCell.Column.Key = "CUST_DC_NO" Then
                    grdARTCUST2.ActiveCell.Value = DBNull.Value
                End If
            End If
        End If
    End Sub

#End Region

    Private Sub optShow_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optShow.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_ARTCUST2()
    End Sub

    Sub Setup_ARTCUST2()
        Dim DVW As DataView = DirectCast(grdARTCUST2.DataSource, DataTable).DefaultView
        Select Case optShow.Value
            Case "A"
                DVW.RowFilter = "CUST_ADDR_TYPE = 'MK'"
            Case "S"
                DVW.RowFilter = "CUST_ADDR_TYPE = 'MK' AND ISNULL(CUST_DC_IND,'0') <> '1'"
            Case "D"
                DVW.RowFilter = "CUST_ADDR_TYPE = 'MK' AND CUST_DC_IND = '1'"
        End Select
    End Sub

    Sub Get_DCs_to_Delete()
        CUST_DC_NOs_to_delete.Clear()
        For Each rowARTCUST2_DC As DataRow In dst.Tables("ARTCUST2").Select("CUST_ADDR_TYPE = 'DC'")
            Dim CUST_ADDR_CODE As String = rowARTCUST2_DC.Item("CUST_ADDR_CODE")
            Dim rowARTCUST2_MK As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "MK", CUST_ADDR_CODE})
            If rowARTCUST2_MK.Item("CUST_DC_IND") & "" <> "1" Then
                CUST_DC_NOs_to_delete.Add(CUST_ADDR_CODE)
            End If
        Next
    End Sub

    Private Sub cmdImportExcel_Click(sender As Object, e As EventArgs) Handles cmdImportExcel.Click
        If ASCMAIN1.CLIENT = "VAN" Then
            dst.Tables("ARTCUSTM").Rows.Clear()
            dst.Tables("ERROR_TBL").Rows.Clear()
            Import_Excel()
            Call SET_EXCEL_CONTROLS(True)

        End If


    End Sub
    Sub SET_EXCEL_CONTROLS(EXCELMODE As Boolean)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Visible = Not EXCELMODE
                .Groups("Screen Control").Items("Update").Visible = Not EXCELMODE
                .Groups("Screen Control").Items("Cancel").Visible = Not EXCELMODE
                .Groups("Screen Control").Items("Update Excel").Settings.Enabled = Not EXCELMODE
                .Groups("Screen Control").Items("Cancel Excel").Settings.Enabled = Not EXCELMODE
                .Groups("Screen Control").Items("Update Excel").Visible = EXCELMODE
                .Groups("Screen Control").Items("Cancel Excel").Visible = EXCELMODE
                .Groups("Display Options").Visible = Not EXCELMODE
                .Groups("Customer Info").Visible = Not EXCELMODE

            End With
        End If

        If EXCELMODE Then
            grdARTCUST2.Visible = False
            grdARTCUSTM.Visible = True
        Else
            grdARTCUST2.Visible = True
            grdARTCUSTM.Visible = False

        End If


    End Sub
    Sub Update_Excel()


        BeginTrans()

        '     If Not ASCMAIN1.Logical_Lock("ARTCUST2", CUST_CODE, , , , 1) Then Exit Sub
        Fill_Records("ARTCUST2", HFs("CUST_CODE"))


        Try

            Dim CUST_ADDR_TYPE As String = ""
            Dim CUST_ADDR_CODE As String = ""

            For Each rowARTCUSTM As DataRow In dst.Tables("ARTCUSTM").Select("", "")
                CUST_ADDR_TYPE = rowARTCUSTM.Item("CUST_ADDR_TYPE")
                CUST_ADDR_CODE = rowARTCUSTM.Item("CUST_ADDR_CODE")

                Dim rowARTCUST2 As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE})
                If rowARTCUST2 Is Nothing Then

                    rowARTCUST2 = dst.Tables("ARTCUST2").NewRow
                    With rowARTCUST2
                        .Item("CUST_CODE") = CUST_CODE
                        .Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
                        .Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    End With
                    dst.Tables("ARTCUST2").Rows.Add(rowARTCUST2)

                    ' NEW
                End If
                'OLD
                With rowARTCUST2
                    .Item("CUST_NAME") = Replace(rowARTCUSTM.Item("CUST_NAME") & "", "`", "") & ""
                    .Item("CUST_ADDR1") = Replace(rowARTCUSTM.Item("CUST_ADDR1") & "", "`", "") & ""
                    .Item("CUST_ADDR2") = Replace(rowARTCUSTM.Item("CUST_ADDR2") & "", "`", "") & ""
                    .Item("CUST_CITY") = Replace(rowARTCUSTM.Item("CUST_CITY") & "", "`", "") & ""
                    .Item("CUST_STATE") = rowARTCUSTM.Item("CUST_STATE") & ""
                    .Item("CUST_ZIP_CODE") = rowARTCUSTM.Item("CUST_ZIP_CODE")
                    .Item("CUST_COUNTRY") = rowARTCUSTM.Item("CUST_COUNTRY")
                    .Item("CUST_CONTACT") = rowARTCUSTM.Item("CUST_CONTACT")
                    .Item("CUST_PHONE") = rowARTCUSTM.Item("CUST_PHONE")
                    '   .Item("INIT_DATE") = DATETIME_STAMP
                    '   .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_DATE") = DATETIME_STAMP
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    .Item("CUST_ADDR_NAME") = rowARTCUSTM.Item("CUST_ADDR_NAME")
                    .Item("CUST_ADDR_STATUS") = rowARTCUSTM.Item("CUST_ADDR_STATUS")
                    .Item("CUST_EMAIL") = rowARTCUSTM.Item("CUST_EMAIL")
                    .Item("CUST_RANK") = rowARTCUSTM.Item("CUST_RANK")
                    .Item("GLOBAL_LOCATION_NUMBER") = rowARTCUSTM.Item("GLOBAL_LOCATION_NUMBER")

                End With
            Next



            Update_Record_TDA("ARTCUST2")

            Dim FN_TO As String = ""
            Dim SESSION_NO As String = ASCMAIN1.Next_Control_No(String.Format("{0}.SESSION_NO", "STYLE_UPLOAD"))
            Dim S As String = Format(DATETIME_STAMP, "yyMMdd") & "_" & Format(DATETIME_STAMP, "HHmmss")

            CommitTrans()
            MsgBox("This Excel File has been successfully Updated to the Customer Store Table",
                          MsgBoxStyle.OkOnly, "Verification")

            ASCMAIN1.MultiTask_Release(, , 2)

            dst.Tables("ARTCUSTM").Rows.Clear()
            dst.Tables("ERROR_TBL").Rows.Clear()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Errors In Updating Item Spreadsheet, Send to ABS")
            Rollback()
        End Try

        ''    If CUST_CODE <> "" Then
        ''        iResponse = MultiTask("ARTCUST2", CUST_CODE, "x", -1, "", sessionid)
        ''    End If

        ''    Dim LAST_DATE As Date
        ''    LAST_DATE = Now + NowTSD
        ''    Screen.MousePointer = 11
        ''    OraS.BeginTrans

        ''    Dim dynWK As Recordset
        ''    Dim tblARWCUST2 As Recordset
        ''    Dim CUST_ADDR_TYPE As String
        ''    Dim CUST_ADDR_CODE As String

        ''    Call Prompt("Loading Master File to Update:", CUST_CODE)
        ''    Sql = " Select * from ARTCUST2 where CUST_CODE = '" & CUST_CODE & "'"
        ''    Call Ora_to_Acc(Nothing, "ARWCUST2", 3, "", Sql)

        ''Set tblARWCUST2 = AccD.OpenRecordset("ARWCUST2", dbOpenTable)
        ''tblARWCUST2.Index = "PrimaryKey"

        ''Set dynWK = AccD.OpenRecordset("ARWCUSTI", dbOpenDynaset)

        ''Do While Not dynWK.EOF
        ''        CUST_ADDR_TYPE = dynWK.Fields("CUST_ADDR_TYPE").Value & ""
        ''        CUST_ADDR_CODE = dynWK.Fields("CUST_ADDR_CODE").Value & ""
        ''        tblARWCUST2.Seek "=", CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE
        ''    If tblARWCUST2.NoMatch Then
        ''            Call Prompt("Adding " & CUST_CODE & " Store: ", CUST_ADDR_CODE)
        ''            tblARWCUST2.AddNew
        ''            tblARWCUST2.Fields("CUST_CODE").Value = CUST_CODE
        ''            tblARWCUST2.Fields("CUST_ADDR_TYPE").Value = CUST_ADDR_TYPE
        ''            tblARWCUST2.Fields("CUST_ADDR_CODE").Value = CUST_ADDR_CODE
        ''            tblARWCUST2.Fields("CUST_RANK").Value = Val(dynWK.Fields("CUST_RANK").Value & "")
        ''            tblARWCUST2.Fields("INIT_DATE").Value = LAST_DATE
        ''            tblARWCUST2.Fields("INIT_OPER").Value = UserID
        ''        Else
        ''            Call Prompt("Modifying " & CUST_CODE & " Store: ", CUST_ADDR_CODE)
        ''            tblARWCUST2.Edit
        ''            '  If tblARWCUST2.Fields("GLOBAL_LOCATION_NUMBER").Value <> dynWK.Fields("GLOBAL_LOCATION_NUMBER").Value & "" Then
        ''            '  Stop
        ''            ' End If
        ''        End If
        ''        tblARWCUST2.Fields("CUST_NAME").Value = Replace(dynWK.Fields("CUST_NAME").Value, "`", "") & ""
        ''        tblARWCUST2.Fields("CUST_ADDR1").Value = Replace(dynWK.Fields("CUST_ADDR1").Value, "`", "") & ""
        ''        tblARWCUST2.Fields("CUST_ADDR2").Value = Replace(dynWK.Fields("CUST_ADDR2").Value, "`", "") & ""
        ''        tblARWCUST2.Fields("CUST_CITY").Value = Replace(dynWK.Fields("CUST_CITY").Value, "`", "") & ""
        ''        tblARWCUST2.Fields("CUST_STATE").Value = dynWK.Fields("CUST_STATE").Value & ""
        ''        tblARWCUST2.Fields("CUST_ZIP_CODE").Value = dynWK.Fields("CUST_ZIP_CODE").Value & ""
        ''        tblARWCUST2.Fields("CUST_COUNTRY").Value = dynWK.Fields("CUST_COUNTRY").Value & ""
        ''        tblARWCUST2.Fields("CUST_CONTACT").Value = dynWK.Fields("CUST_CONTACT").Value & ""
        ''        tblARWCUST2.Fields("CUST_PHONE").Value = dynWK.Fields("CUST_PHONE").Value & ""
        ''        tblARWCUST2.Fields("CUST_EXT").Value = dynWK.Fields("CUST_EXT").Value & ""
        ''        tblARWCUST2.Fields("CUST_FAX").Value = dynWK.Fields("CUST_FAX").Value & ""
        ''        tblARWCUST2.Fields("LAST_DATE").Value = LAST_DATE
        ''        tblARWCUST2.Fields("LAST_OPER").Value = UserID
        ''        tblARWCUST2.Fields("CUST_ADDR_NAME").Value = ""
        ''        tblARWCUST2.Fields("CUST_ADDR_STATUS").Value = dynWK.Fields("CUST_ADDR_STATUS").Value & ""
        ''        tblARWCUST2.Fields("CUST_EMAIL").Value = dynWK.Fields("CUST_EMAIL").Value & ""
        ''        tblARWCUST2.Fields("GLOBAL_LOCATION_NUMBER").Value = dynWK.Fields("GLOBAL_LOCATION_NUMBER").Value & ""
        ''        tblARWCUST2.Update
        ''        dynWK.MoveNext
        ''    Loop

        ''    Call Delete_Records()
        ''    Call Prompt("Updating Master File...", "")
        ''    Call Acc_to_Ora("ARWCUST2", "")
        ''    Call Prompt("Finished!", "")
        ''    OraS.CommitTrans
        ''    Call cmdExecute(4)




    End Sub
    Sub Import_Excel()

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            Dim filter As String = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
            openFileDialog1.Filter = filter
            openFileDialog1.RestoreDirectory = True
            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using
        'Try
        Dim Vs As New Dictionary(Of String, Integer)
        Dim SHEETNAME As String = ""
        If FILENAME <> "" Then
            If Absx1.txtFor("CUST_CODE").Text & "" = "WALMART" Then
                If optEXCEL.Value = "H" Then
                    SHEETNAME = "Wal-Mart Stores"
                ElseIf optEXCEL.Value = "D" Then
                    SHEETNAME = "Wal-Mart DC Receiving"
                End If
            ElseIf Absx1.txtFor("CUST_CODE").Text & "" = "SAMSCLUB" Then
                If optEXCEL.Value = "H" Then
                    SHEETNAME = "Sams Clubs"
                ElseIf optEXCEL.Value = "D" Then
                    SHEETNAME = "Sams DC Receiving Addesses"
                End If
            End If

            Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
            Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(SHEETNAME)
            Dim range As SpreadsheetGear.IRange = Nothing
            Dim r As Integer = 0
            Dim ERROR_CODEs As List(Of String) = New List(Of String)
            Dim BLANKCODES As Integer = 0

            r = 4

            Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text & ""

            Do While oSheet.Cells(r, 0).Value & "" <> "END"
                Try
                    If BLANKCODES > 20 Then
                        Exit Do
                    End If
                    Dim rowARTCUSTM As DataRow
                    Dim PHONEVAR As String = ""
                    Dim CUST_ADDR_TYPE As String = "MK"
                    Dim DC_TYPE(8) As String
                    Dim I As Integer

                    Dim CUST_ADDR_CODE As String = ""
                    If Val(Trim(oSheet.Cells(r, 0).Value & "")) <> 0 Then
                        CUST_ADDR_CODE = Format(Val(Trim(oSheet.Cells(r, 0).Value & "")), "000000") & ""
                    End If

                    If CUST_ADDR_CODE & "" <> "" Then
                        BLANKCODES = 0
                    Else
                        BLANKCODES = BLANKCODES + 1
                    End If

                    If optEXCEL.Value = "H" Then
                        ' WALMART SAMSCLUB SHEET
                        Select Case CUST_CODE
                            Case "WALMART"
                                ' CHECK TO SEE IF RECORD EXISTS (ERROR) ELSE ADD NEW
                                Dim row As DataRow = dst.Tables("ARTCUSTM").Rows.Find(New Object() {CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE})
                                If row Is Nothing And CUST_ADDR_CODE <> "" Then

                                    rowARTCUSTM = dst.Tables("ARTCUSTM").NewRow
                                    With rowARTCUSTM
                                        .Item("CUST_CODE") = CUST_CODE
                                        .Item("CUST_ADDR_TYPE") = "MK"
                                        .Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                                        .Item("CUST_NAME") = Trim(oSheet.Cells(r, 3).Value & "") & "" & " " & Format(Val(Trim(oSheet.Cells(r, 0).Value & "")), "0000") & ""
                                        .Item("CUST_ADDR1") = Replace((Trim(oSheet.Cells(r, 4).Value & "")), "`", "") & ""
                                        .Item("CUST_ADDR2") = ""
                                        .Item("CUST_CITY") = Replace((Trim(oSheet.Cells(r, 5).Value & "")), "`", "") & ""
                                        .Item("CUST_STATE") = Trim(oSheet.Cells(r, 6).Value & "")
                                        .Item("CUST_ZIP_CODE") = Trim(oSheet.Cells(r, 7).Value & "")
                                        .Item("CUST_COUNTRY") = "USA"
                                        .Item("CUST_CONTACT") = "Receiving"
                                        PHONEVAR = Replace((Trim(oSheet.Cells(r, 12).Value & "")), "-", "")
                                        PHONEVAR = Replace((PHONEVAR & ""), " ", "")
                                        .Item("CUST_PHONE") = PHONEVAR & ""
                                        .Item("INIT_DATE") = DATETIME_STAMP
                                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                        .Item("LAST_DATE") = DATETIME_STAMP
                                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                                        .Item("CUST_ADDR_NAME") = ""
                                        .Item("CUST_ADDR_STATUS") = "A"
                                        .Item("CUST_EMAIL") = ""
                                        .Item("CUST_RANK") = Null
                                        .Item("GLOBAL_LOCATION_NUMBER") = Trim(oSheet.Cells(r, 1).Value & "")
                                    End With
                                    dst.Tables("ARTCUSTM").Rows.Add(rowARTCUSTM)

                                    rowARTCUSTM = dst.Tables("ARTCUSTM").NewRow
                                    With rowARTCUSTM
                                        .Item("CUST_CODE") = CUST_CODE
                                        .Item("CUST_ADDR_TYPE") = "DC"
                                        .Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                                        .Item("CUST_NAME") = Trim(oSheet.Cells(r, 3).Value & "") & "" & " " & Format(Val(Trim(oSheet.Cells(r, 0).Value & "")), "0000") & ""
                                        .Item("CUST_ADDR1") = Replace((Trim(oSheet.Cells(r, 4).Value & "")), "`", "") & ""
                                        .Item("CUST_ADDR2") = ""
                                        .Item("CUST_CITY") = Replace((Trim(oSheet.Cells(r, 5).Value & "")), "`", "") & ""
                                        .Item("CUST_STATE") = Trim(oSheet.Cells(r, 6).Value & "")
                                        .Item("CUST_ZIP_CODE") = Trim(oSheet.Cells(r, 7).Value & "")
                                        .Item("CUST_COUNTRY") = "USA"
                                        .Item("CUST_CONTACT") = "Receiving"
                                        PHONEVAR = Replace((Trim(oSheet.Cells(r, 12).Value & "")), "-", "")
                                        PHONEVAR = Replace((PHONEVAR & ""), " ", "")
                                        .Item("CUST_PHONE") = PHONEVAR & ""
                                        .Item("INIT_DATE") = DATETIME_STAMP
                                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                        .Item("LAST_DATE") = DATETIME_STAMP
                                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                                        .Item("CUST_ADDR_NAME") = ""
                                        .Item("CUST_ADDR_STATUS") = "A"
                                        .Item("CUST_EMAIL") = ""
                                        .Item("CUST_RANK") = Null
                                        .Item("GLOBAL_LOCATION_NUMBER") = Trim(oSheet.Cells(r, 1).Value & "")
                                    End With
                                    dst.Tables("ARTCUSTM").Rows.Add(rowARTCUSTM)
                                Else
                                    If CUST_ADDR_CODE <> "" Then
                                        ERROR_CODEs.Add("Customer Store already present on Line No " & r)
                                        Dim rowERROR_TBL As DataRow = Nothing
                                        rowERROR_TBL = dst.Tables("ERROR_TBL").NewRow
                                        With rowERROR_TBL
                                            .Item("ERROR_CODE") = "Duplicate Customer/Store in Excel " & CUST_ADDR_CODE
                                            .Item("ERROR_DETAIL") = "Ln# " & r
                                        End With
                                        dst.Tables("ERROR_TBL").Rows.Add(rowERROR_TBL)
                                    End If
                                End If
                            Case "SAMSCLUB"
                                ' CHECK TO SEE IF RECORD EXISTS (ERROR) ELSE ADD NEW
                                Dim row As DataRow = dst.Tables("ARTCUSTM").Rows.Find(New Object() {CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE})
                                If row Is Nothing And CUST_ADDR_CODE <> "" Then

                                    rowARTCUSTM = dst.Tables("ARTCUSTM").NewRow
                                    With rowARTCUSTM
                                        .Item("CUST_CODE") = CUST_CODE
                                        .Item("CUST_ADDR_TYPE") = "MK"
                                        .Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                                        .Item("CUST_NAME") = Trim(oSheet.Cells(r, 3).Value & "") & "" & " " & Format(Val(Trim(oSheet.Cells(r, 0).Value & "")), "0000") & ""
                                        .Item("CUST_ADDR1") = Replace((Trim(oSheet.Cells(r, 3).Value & "")), "`", "") & ""
                                        .Item("CUST_ADDR2") = ""
                                        .Item("CUST_CITY") = Replace((Trim(oSheet.Cells(r, 4).Value & "")), "`", "") & ""
                                        .Item("CUST_STATE") = Trim(oSheet.Cells(r, 5).Value & "")
                                        .Item("CUST_ZIP_CODE") = Trim(oSheet.Cells(r, 6).Value & "")
                                        .Item("CUST_COUNTRY") = "USA"
                                        .Item("CUST_CONTACT") = "Receiving"
                                        PHONEVAR = Replace((Trim(oSheet.Cells(r, 7).Value & "")), "-", "")
                                        PHONEVAR = Replace((PHONEVAR & ""), " ", "")
                                        PHONEVAR = Replace((PHONEVAR & ""), "(", "")
                                        PHONEVAR = Replace((PHONEVAR & ""), ")", "")

                                        .Item("CUST_PHONE") = PHONEVAR & ""
                                        .Item("INIT_DATE") = DATETIME_STAMP
                                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                        .Item("LAST_DATE") = DATETIME_STAMP
                                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                                        .Item("CUST_ADDR_NAME") = ""
                                        .Item("CUST_ADDR_STATUS") = "A"
                                        .Item("CUST_EMAIL") = ""
                                        .Item("CUST_RANK") = Null
                                        .Item("GLOBAL_LOCATION_NUMBER") = Trim(oSheet.Cells(r, 1).Value & "")
                                    End With
                                    dst.Tables("ARTCUSTM").Rows.Add(rowARTCUSTM)

                                    rowARTCUSTM = dst.Tables("ARTCUSTM").NewRow
                                    With rowARTCUSTM
                                        .Item("CUST_CODE") = CUST_CODE
                                        .Item("CUST_ADDR_TYPE") = "DC"
                                        .Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                                        .Item("CUST_NAME") = Trim(oSheet.Cells(r, 3).Value & "") & "" & " " & Format(Val(Trim(oSheet.Cells(r, 0).Value & "")), "0000") & ""
                                        .Item("CUST_ADDR1") = Replace((Trim(oSheet.Cells(r, 3).Value & "")), "`", "") & ""
                                        .Item("CUST_ADDR2") = ""
                                        .Item("CUST_CITY") = Replace((Trim(oSheet.Cells(r, 4).Value & "")), "`", "") & ""
                                        .Item("CUST_STATE") = Trim(oSheet.Cells(r, 5).Value & "")
                                        .Item("CUST_ZIP_CODE") = Trim(oSheet.Cells(r, 6).Value & "")
                                        .Item("CUST_COUNTRY") = "USA"
                                        .Item("CUST_CONTACT") = "Receiving"
                                        PHONEVAR = Replace((Trim(oSheet.Cells(r, 7).Value & "")), "-", "")
                                        PHONEVAR = Replace((PHONEVAR & ""), " ", "")
                                        PHONEVAR = Replace((PHONEVAR & ""), "(", "")
                                        PHONEVAR = Replace((PHONEVAR & ""), ")", "")

                                        .Item("CUST_PHONE") = PHONEVAR & ""
                                        .Item("INIT_DATE") = DATETIME_STAMP
                                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                        .Item("LAST_DATE") = DATETIME_STAMP
                                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                                        .Item("CUST_ADDR_NAME") = ""
                                        .Item("CUST_ADDR_STATUS") = "A"
                                        .Item("CUST_EMAIL") = ""
                                        .Item("CUST_RANK") = Null
                                        .Item("GLOBAL_LOCATION_NUMBER") = Trim(oSheet.Cells(r, 1).Value & "")
                                    End With
                                    dst.Tables("ARTCUSTM").Rows.Add(rowARTCUSTM)
                                Else
                                    If CUST_ADDR_CODE <> "" Then
                                        ERROR_CODEs.Add("Customer Store already present on Line No " & r)
                                        Dim rowERROR_TBL As DataRow = Nothing
                                        rowERROR_TBL = dst.Tables("ERROR_TBL").NewRow
                                        With rowERROR_TBL
                                            .Item("ERROR_CODE") = "Duplicate Customer/Store in Excel " & CUST_ADDR_CODE
                                            .Item("ERROR_DETAIL") = "Ln# " & r
                                        End With
                                        dst.Tables("ERROR_TBL").Rows.Add(rowERROR_TBL)
                                    End If
                                End If


                        End Select
                    ElseIf optEXCEL.Value = "D" Then
                        ' DC RECEIVING SHEET
                        Select Case CUST_CODE
                            Case "WALMART"
                                ReDim DC_TYPE(8)

                                DC_TYPE(0) = "R"
                                DC_TYPE(1) = "G"
                                DC_TYPE(2) = "T"
                                DC_TYPE(3) = "D"
                                DC_TYPE(4) = "J"
                                DC_TYPE(5) = "P"
                                DC_TYPE(6) = "W"
                                DC_TYPE(7) = "I"
                                DC_TYPE(8) = "A"
                                I = Val(Mid(Trim(oSheet.Cells(r, 4).Value & ""), 1, 1))
                                If Trim(oSheet.Cells(r, 1).Value & "") & "" <> "" And Mid(Trim(oSheet.Cells(r, 4).Value & ""), 1, 1) <> "0" Then
                                    CUST_ADDR_CODE = Format(Val(Trim(oSheet.Cells(r, 1).Value & "")), "0000") & "" & DC_TYPE(I - 1)
                                    CUST_ADDR_TYPE = "DC"
                                    ' CHECK TO SEE IF RECORD EXISTS (ERROR) ELSE ADD NEW

                                    Dim row As DataRow = dst.Tables("ARTCUSTM").Rows.Find(New Object() {CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE})
                                    If row Is Nothing And CUST_ADDR_CODE <> "" Then

                                        rowARTCUSTM = dst.Tables("ARTCUSTM").NewRow
                                        With rowARTCUSTM
                                            .Item("CUST_CODE") = CUST_CODE
                                            .Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
                                            .Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                                            '        .Item("CUST_NAME") = Trim(oSheet.Cells(r, 7).Value & "") & ""
                                            .Item("CUST_ADDR1") = Replace((Trim(oSheet.Cells(r, 8).Value & "")), "`", "") & ""
                                            .Item("CUST_ADDR2") = ""
                                            .Item("CUST_CITY") = Replace((Trim(oSheet.Cells(r, 11).Value & "")), "`", "") & ""
                                            .Item("CUST_STATE") = Trim(oSheet.Cells(r, 12).Value & "")
                                            .Item("CUST_ZIP_CODE") = Trim(oSheet.Cells(r, 13).Value & "")
                                            .Item("CUST_COUNTRY") = "USA"
                                            .Item("CUST_CONTACT") = "Receiving"
                                            PHONEVAR = Replace((Trim(oSheet.Cells(r, 15).Value & "")), "-", "")
                                            PHONEVAR = Replace((PHONEVAR & ""), " ", "")
                                            .Item("CUST_PHONE") = PHONEVAR & ""
                                            .Item("INIT_DATE") = DATETIME_STAMP
                                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                            .Item("LAST_DATE") = DATETIME_STAMP
                                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                                            .Item("CUST_ADDR_NAME") = ""
                                            .Item("CUST_ADDR_STATUS") = "A"
                                            .Item("CUST_EMAIL") = ""
                                            .Item("CUST_RANK") = Null
                                            .Item("GLOBAL_LOCATION_NUMBER") = Trim(oSheet.Cells(r, 2).Value & "")
                                        End With
                                        dst.Tables("ARTCUSTM").Rows.Add(rowARTCUSTM)
                                    Else
                                        ERROR_CODEs.Add("Customer Store already present on Line No " & r)
                                        Dim rowERROR_TBL As DataRow = Nothing
                                        rowERROR_TBL = dst.Tables("ERROR_TBL").NewRow
                                        With rowERROR_TBL
                                            .Item("ERROR_CODE") = "Duplicate Customer/Store in Excel " & CUST_ADDR_CODE
                                            .Item("ERROR_DETAIL") = "Ln# " & r
                                        End With
                                        dst.Tables("ERROR_TBL").Rows.Add(rowERROR_TBL)
                                    End If
                                End If

                            Case "SAMSCLUB"
                                ReDim DC_TYPE(8)

                                DC_TYPE(0) = "R"
                                DC_TYPE(1) = "G"
                                DC_TYPE(2) = "T"
                                DC_TYPE(3) = "D"
                                DC_TYPE(4) = "J"
                                DC_TYPE(5) = "P"
                                DC_TYPE(6) = "W"
                                DC_TYPE(7) = "I"
                                DC_TYPE(8) = "A"
                                I = Val(Mid(Trim(oSheet.Cells(r, 4).Value & ""), 1, 1))
                                If Trim(oSheet.Cells(r, 1).Value & "") & "" <> "" Then
                                    '  CUST_ADDR_CODE = Format(Val(Trim(oSheet.Cells(r, 1).Value & "")), "0000") & "" & DC_TYPE(I - 1)
                                    ' CUST_ADDR_CODE = Format(Val(Trim(oSheet.Cells(r, 1).Value & "")), "0000") & "S"
                                    CUST_ADDR_CODE = Format(Val(Trim(oSheet.Cells(r, 1).Value & "")), "000000")

                                    CUST_ADDR_TYPE = "DC"
                                    ' CHECK TO SEE IF RECORD EXISTS (ERROR) ELSE ADD NEW

                                    Dim row As DataRow = dst.Tables("ARTCUSTM").Rows.Find(New Object() {CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE})
                                    If row Is Nothing And CUST_ADDR_CODE <> "" Then

                                        rowARTCUSTM = dst.Tables("ARTCUSTM").NewRow
                                        With rowARTCUSTM
                                            .Item("CUST_CODE") = CUST_CODE
                                            .Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
                                            .Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                                            .Item("CUST_NAME") = Trim(oSheet.Cells(r, 7).Value & "") & ""
                                            .Item("CUST_ADDR1") = Replace((Trim(oSheet.Cells(r, 8).Value & "")), "`", "") & ""
                                            .Item("CUST_ADDR2") = ""
                                            .Item("CUST_CITY") = Replace((Trim(oSheet.Cells(r, 11).Value & "")), "`", "") & ""
                                            .Item("CUST_STATE") = Trim(oSheet.Cells(r, 12).Value & "")
                                            .Item("CUST_ZIP_CODE") = Trim(oSheet.Cells(r, 13).Value & "")
                                            .Item("CUST_COUNTRY") = "USA"
                                            .Item("CUST_CONTACT") = "Receiving"
                                            PHONEVAR = Replace((Trim(oSheet.Cells(r, 15).Value & "")), "-", "")
                                            PHONEVAR = Replace((PHONEVAR & ""), " ", "")
                                            .Item("CUST_PHONE") = PHONEVAR & ""
                                            .Item("INIT_DATE") = DATETIME_STAMP
                                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                            .Item("LAST_DATE") = DATETIME_STAMP
                                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                                            .Item("CUST_ADDR_NAME") = ""
                                            .Item("CUST_ADDR_STATUS") = "A"
                                            .Item("CUST_EMAIL") = ""
                                            .Item("CUST_RANK") = Null
                                            .Item("GLOBAL_LOCATION_NUMBER") = Trim(oSheet.Cells(r, 2).Value & "")
                                        End With
                                        dst.Tables("ARTCUSTM").Rows.Add(rowARTCUSTM)
                                    Else
                                        ERROR_CODEs.Add("Customer Store already present on Line No " & r)
                                        Dim rowERROR_TBL As DataRow = Nothing
                                        rowERROR_TBL = dst.Tables("ERROR_TBL").NewRow
                                        With rowERROR_TBL
                                            .Item("ERROR_CODE") = "Duplicate Customer/Store in Excel " & CUST_ADDR_CODE
                                            .Item("ERROR_DETAIL") = "Ln# " & r
                                        End With
                                        dst.Tables("ERROR_TBL").Rows.Add(rowERROR_TBL)
                                    End If
                                End If
                        End Select
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Errors In Updating Item Spreadsheet, Send to ABS")
                    ' MsgBox("Errors In Updating Item Spreadsheet, Send to ABS")
                    Exit Sub
                End Try
                r = r + 1
            Loop
            If ERROR_CODEs.Count <> 0 Then

                If dst.Tables("ERROR_TBL").Rows.Count <> 0 Then
                    Using F As New ASFMSGBF
                        F.Show_grd(dst.Tables("ERROR_TBL"), Me, "The following Import Errors have been identified", "DGJ")
                    End Using
                End If

            Else
                '  Stop ' GOOD TO UPDATE
                MsgBox("This Excel File has been successfully Imported with no Errors. Click Excel Update to Update Database",
                          MsgBoxStyle.OkOnly, "Verification")

            End If

        End If
    End Sub

End Class