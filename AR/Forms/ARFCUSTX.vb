Public Class ARFCUSTX

    ' formatting of store no

    Dim rowARTCUST1 As DataRow
    Dim ADDRESS() As String = {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3", _
             "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY", _
             "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_EMAIL", "GLOBAL_LOCATION_NUMBER"}
    Dim ADDRESS_EXT() As String = {"CUST_ADDR2", "CUST_ADDR3", _
          "CUST_COUNTRY", _
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
        End With

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
                                  & vbCrLf & " " & Join(CUST_DC_NOs_to_delete.ToArray, ","), _
                                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub
                    End If
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

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Display Options").Visible = ScreenMode
                .Groups("Customer Info").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grpStores.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ARTCUST1", "ARTCUST2", "ARTCUST3"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

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
                   & vbCrLf & Join(CUST_DC_NOs_nogood.ToArray, ","), _
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
End Class