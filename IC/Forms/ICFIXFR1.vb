Imports System.Drawing
Imports System.Math

Public Class ICFIXFR1
    ' SHOULD PROBABLY ADD A LOCATION_CODE TO ICTIXFR1 AND PROMPT FOR IT IF THE XFR-TO WHSE IS A LOCATOR - THEN CHG SP TO NOT USE DEFAULT LOC FOR XIN
    Dim rowICTIXFR1 As DataRow
    Dim location_support As Boolean = False
    Dim rowICTWHSE1 As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ICFIXFRI" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")
        Get_PARM("WHTPARM1")

        With dst
            ASCMAIN1.sql = "Select ICTIXFR1.*" _
            & " from ICTIXFR1 where ICTIXFR1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "ICTIXFRX", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "ICTIXFR1", "*")

            ASCMAIN1.sql = "Select ICTIXFR2.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" _
            & " from ICTIXFR2,ICTSTYL1,ICTCOLR1" _
            & " where ICTSTYL1.STYLE_CODE = ICTIXFR2.STYLE_CODE" _
            & " and ICTCOLR1.COLOR_CODE = ICTIXFR2.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTIXFR2", "**", 1)
            .Tables("ICTIXFR2").Columns.Add("LINE_COSTS", GetType(System.Decimal), "ISNULL(XFR_QTY,0) * ISNULL(STYLE_COST,0)")

            ASCMAIN1.sql = "Select ICTIXFR3.*, GLTACCT1.ACCT_DESC" _
            & " from ICTIXFR3,GLTACCT1 where GLTACCT1.ACCT_CODE = ICTIXFR3.ACCT_CODE"
            Create_TDA(.Tables.Add, "ICTIXFR3", "**", 1)

            ASCMAIN1.sql = "Select ICTSTAT2.*" _
            & " from ICTSTAT2 where STYLE_CODE = :PARM1 and WHSE_CODE = :PARM2"
            Create_TDA(.Tables.Add, "ICTSTAT2", "**", 0, False, "VV")

            .Tables.Add("ICTIXFR0")
            .Tables("ICTIXFR0").Columns.Add("KEY")
            .Tables("ICTIXFR0").Columns.Add("DESCRIPTION")
        End With

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)

        grdICTIXFR0.DataSource = dst.Tables("ICTIXFR0")
        grdICTIXFR2.DataSource = dst.Tables("ICTIXFR2")
        grdICTIXFR3.DataSource = dst.Tables("ICTIXFR3")
        grdICTIXFRX.DataSource = dst.Tables("ICTIXFRX")

        Create_Summary(grdICTIXFRX, "XFR_NO", "Count")
        Create_Summary(grdICTIXFRX, "TOTAL_COSTS")

        Create_Summary(grdICTIXFR2, "XFR_LNO", "Count")
        Create_Summary(grdICTIXFR2, "XFR_QTY")
        Create_Summary(grdICTIXFR2, "LINE_COSTS")

        Create_Summary(grdICTIXFR3, "XFR_GNO", "Count")
        Create_Summary(grdICTIXFR3, "DIST_AMT")


        With grdICTIXFRX.DisplayLayout.Bands("ICTIXFRX")
            .Columns("XFR_NO").Header.Fixed = True
        End With

        'ASCMAIN1.Add_Value_List(grdICTIXFRX, "WHSE_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 where REASON_TYPE = 'A' order by REASON_DESC")

        grdICTIXFR0.DisplayLayout.Bands(0).ColHeadersVisible = False
        Set_SEGS(grdICTIXFR3, "ICTIXFR3")

        Set_Read_Only(grpTotals, True)
        If InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") = 0 Then
            grpTotals.Visible = False
            With grdICTIXFR2.DisplayLayout.Bands(0)
                .Columns("STYLE_COST").Hidden = True
                .Columns("LINE_COSTS").Hidden = True
                .Columns("STYLE_CLASS_CODE").Hidden = True
                .Columns("SALES_DIVISION_CODE").Hidden = True
            End With
        End If

        grpHeader.Visible = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("WHSE_CODE")

                If ASCMAIN1.CLIENT = "NYA" Then
                    If Absx1.txtFor("WHSE_CODE").Text = "21" Then
                        EMsg &= vbCr & "Warehouse Transfers not allowed for NYAG Candada - need to set up Intercompany Sale"
                    End If
                End If

                If Absx1.dteFor("XFR_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Invalid Date Specified for Entry"
                End If

                If Absx1.txtFor("WHSE_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If IsNothing(rowICTWHSE1) Then
                        EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                    Else
                        If rowICTWHSE1.Item("WHSE_STATUS").ToString <> "A" Then
                            EMsg &= vbCr & "Warehouse Entered Is Not Active"
                        Else
                            If rowICTWHSE1.Item("LP_CODE") & "" <> "" Then
                                If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                                Else
                                    EMsg &= vbCr & "Warehouse Entered Is A 3PL.  No Transfer Allowed"
                                End If
                            End If
                        End If
                    End If
                End If

            Case "View"
                If Absx1.txtFor("XFR_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowICTIXFR1 = LookUp("ICTIXFR1", Absx1.txtFor("XFR_NO").Text)
                    If rowICTIXFR1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & Absx1.txtFor("XFR_NO").Text & " on File"
                    End If
                End If

            Case "Update"
                If Absx1.txtFor("WHSE_CODE_TO").Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Transfer-To Warehouse"
                Else
                    'Validate_Code("WHSE_CODE_TO") - did not work - gets object reference error
                    Dim row As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE_TO").Text)
                    If row Is Nothing Then
                        EMsg &= vbCr & "Invalid Value specified for Transfer-To Warehouse"
                    Else
                        If row.Item("LP_CODE") & "" <> "" Then
                            EMsg &= vbCr & "Warehouse Entered Is A 3PL.  No Transfer Allowed"
                        End If
                        If Absx1.txtFor("WHSE_CODE").Text = Absx1.txtFor("WHSE_CODE_TO").Text Then
                            EMsg &= "Transfer-From and Transfer-To Warehouses must be different." & vbCr
                        End If
                    End If
                End If

                If grdICTIXFR2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Details Entered"
                Else
                    For Each rowICTIXFR2 As DataRow In dst.Tables("ICTIXFR2").Select("", "", DataViewRowState.CurrentRows)
                        If rowICTIXFR2.Item("STYLE_CLASS_CODE") & "" = "" Then
                            EMsg &= "Unable to determine Item Class Code for " & rowICTIXFR2.Item("STYLE_CODE") & ""
                        End If
                        If rowICTIXFR2.Item("SALES_DIVISION_CODE") & "" = "" Then
                            EMsg &= "Unable to determine Sales Division Code for " & rowICTIXFR2.Item("STYLE_CODE") & ""
                        End If

                        If Val(rowICTIXFR2.Item("XFR_QTY") & "") <= 0 Then
                            If ASCMAIN1.CLIENT = "RGI" And ASCMAIN1.Running_in_VS Then
                                Stop
                            Else
                                EMsg &= "Positive Values Only (see " & rowICTIXFR2.Item("STYLE_CODE") & ")"
                            End If
                        End If
                    Next
                End If

                If ASCMAIN1.CLIENT = "NYA" Then
                    If Absx1.txtFor("WHSE_CODE").Text = "21" Or Absx1.txtFor("WHSE_CODE_TO").Text = "21" Then
                        EMsg &= vbCr & "Warehouse Transfers not allowed for NYAG Candada - need to set up Intercompany Sale"
                    End If
                End If


                If EMsg = "" Then
                    Dim msg As String = Check_Qty("ICTIXFR2", Absx1.txtFor("WHSE_CODE").Text, "XFR_QTY", -1)
                    If msg <> "" Then
                        If MsgBox(msg & vbCr & vbCr & "OK to Continue Anyway?", MsgBoxStyle.YesNo, "The following Items do not have Sufficent Qty for this Transaction") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Reverse"
                If MessageBox.Show("Are you sure you want to reverse this Entry?", "Confirm Reversal", _
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
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

            Case "Reverse"
                Set_Up_Reversal()
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
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode

                    If ScreenMode And EntryMode <> "N" Then
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                    End If

                    .Items("Reverse").Visible = (ScreenMode AndAlso EntryMode = "V" AndAlso Not InquiryMode) _
                        AndAlso rowICTIXFR1 IsNot Nothing _
                        AndAlso rowICTIXFR1.Item("REVERSED_BY_XFR_NO") Is DBNull.Value _
                        AndAlso rowICTIXFR1.Item("REVERSES_XFR_NO") Is DBNull.Value
                End With

                .Groups("GL Distribution").Visible = ScreenMode And (EntryMode = "V") And InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") <> 0
                .Groups("Show if Entered in").Visible = Not ScreenMode And InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") <> 0
                .Groups("Totals").Visible = ScreenMode
                .Groups("Events").Visible = ScreenMode And (EntryMode <> "N")
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode
        grdICTIXFRX.Visible = Not ScreenMode

        If ScreenMode Then
            grdICTIXFR0.Visible = (EntryMode = "V")
            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                End With
            End If

            SplitContainer2.Panel2Collapsed = (EntryMode <> "V") Or InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") = 0
            Set_Read_Only(grpHeader, (EntryMode = "V"))
            Set_Read_Only(SplitContainer2, (EntryMode = "V"))
            If EntryMode = "N" Then
                With grdICTIXFR2.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = DefaultableBoolean.True
                    .AllowUpdate = DefaultableBoolean.True
                End With
                With grdICTIXFR2.DisplayLayout.Bands(0)
                    .Columns("STYLE_CODE").CellAppearance.BackColor = Color.LightYellow
                    .Columns("COLOR_CODE").CellAppearance.BackColor = Color.LightYellow

                    .Columns("XFR_QTY").CellAppearance.BackColor = Color.LightYellow
                End With
                With grdICTIXFR3.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = DefaultableBoolean.True
                    .AllowUpdate = DefaultableBoolean.True
                End With
            Else
                With grdICTIXFR2.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowDelete = DefaultableBoolean.False
                    .AllowUpdate = DefaultableBoolean.False
                End With
                With grdICTIXFR2.DisplayLayout.Bands(0)
                    .Columns("STYLE_CODE").CellAppearance.BackColor = Color.Empty
                    .Columns("COLOR_CODE").CellAppearance.BackColor = Color.Empty
                    .Columns("XFR_QTY").CellAppearance.BackColor = Color.Empty
                End With
                With grdICTIXFR3.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowDelete = DefaultableBoolean.False
                    .AllowUpdate = DefaultableBoolean.False
                End With
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTIXFR0", "ICTIXFR1", "ICTIXFR2", "ICTIXFR3"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()

        Absx1.txtFor("WHSE_CODE").Text = ""
        Absx1.dteFor("XFR_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.txtFor("XFR_NO").Text = ""

        optGL.Tag = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        'If EntryMode = "N" Then
        '    Absx1.txtFor("XFR_NO").Text = ASCMAIN1.Next_Control_No("SOTINVH1.XFR_NO")
        'End If

        If EntryMode = "N" Then
            rowICTIXFR1 = dst.Tables("ICTIXFR1").NewRow
            rowICTIXFR1.Item("XFR_NO") = ASCMAIN1.Next_Control_No("ICTIXFR1.XFR_NO")
            rowICTIXFR1.Item("WHSE_CODE") = HFs("WHSE_CODE")
            rowICTIXFR1.Item("XFR_DATE") = HFs("XFR_DATE")
            rowICTIXFR1.Item("XFR_SOURCE") = "E"
            rowICTIXFR1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowICTIXFR1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTIXFR1.Item("INIT_DATE") = DATETIME_STAMP
            rowICTIXFR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowICTIXFR1.Item("LAST_DATE") = DATETIME_STAMP
            rowICTIXFR1.Item("REGISTER_IND") = "0"
            rowICTIXFR1.Item("JOURNAL_IND") = "0"
            dst.Tables("ICTIXFR1").Rows.Add(rowICTIXFR1)
        Else
            Fill_Record("ICTIXFR1", Absx1.txtFor("XFR_NO").Text)
            dst.AcceptChanges()

            With dst.Tables("ICTIXFR0")
                .Rows.Add(New String() {"Entered By", rowICTIXFR1.Item("INIT_OPER")})
                .Rows.Add(New String() {"Entered On", Format(rowICTIXFR1.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
                .Rows.Add(New String() {"Source", rowICTIXFR1.Item("XFR_SOURCE")})

                If rowICTIXFR1.Item("XFR_SOURCE") & "" = "S" Then
                    .Rows.Add(New String() {"Xfr Inv No", rowICTIXFR1.Item("CTL_NO")})
                End If

                If rowICTIXFR1.Item("REVERSED_BY_XFR_NO") & "" <> "" Then
                    Dim row As DataRow = LookUp("ICTIXFR1", rowICTIXFR1.Item("REVERSED_BY_XFR_NO"))
                    .Rows.Add(New String() {"Reversed", Format(row.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
                    .Rows.Add(New String() {"By", row.Item("INIT_OPER")})
                    .Rows.Add(New String() {"using", rowICTIXFR1.Item("REVERSED_BY_XFR_NO")})
                ElseIf rowICTIXFR1.Item("REVERSES_XFR_NO") & "" <> "" Then
                    .Rows.Add(New String() {"Reverses", rowICTIXFR1.Item("REVERSES_XFR_NO")})
                End If
            End With
        End If

        rowICTWHSE1 = LookUp("ICTWHSE1", rowICTIXFR1.Item("WHSE_CODE"))
        location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")
        With grdICTIXFR2.DisplayLayout.Bands(0)
            .Columns("BAR_CODE").Hidden = True ' Not location_support
            .Columns("LOCATION_CODE").Hidden = Not location_support
        End With

        Fill_Records("ICTIXFR2", Absx1.txtFor("XFR_NO").Text)
        Fill_Records("ICTIXFR3", Absx1.txtFor("XFR_NO").Text)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        ICCMAIN1.Update_Transfer(Me)

        If location_support Then

            'Update_WHTLOCBX("T")
            TAC.ICCMAIN1.Update_WHTLOCBX("T", rowICTIXFR1.Item("XFR_NO"))

            'ASCMAIN1.sql = "Select ICTIXFR2.XFR_NO WHSE_TRAN_NO, ICTIXFR2.XFR_LNO WHSE_TRAN_LNO" _
            '               & ", 'T' WHSE_TRAN_TYPE, ICTIXFR1.WHSE_CODE" _
            '               & ", ICTIXFR2.LOCATION_CODE, ICTIXFR2.STYLE_CODE, ICTIXFR2.COLOR_CODE" _
            '               & ", -1 * ICTIXFR2.XFR_QTY WHSE_TRAN_QTY" _
            '               & " from ICTIXFR1,ICTIXFR2" _
            '               & " where ICTIXFR1.XFR_NO = ICTIXFR2.XFR_NO" _
            '               & "   and ICTIXFR2.XFR_NO = '" & rowICTIXFR1.Item("XFR_NO") & "'"
            'WHCMAIN1.Update_WHTLOCBX(Me)
        End If

        Dim rowICTWHSE1_WHSE_CODE_TO As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE_TO").Text)
        If rowICTWHSE1_WHSE_CODE_TO.Item("WHSE_LOCATOR") & "" = "1" Then

            'Update_WHTLOCBX("X")
            TAC.ICCMAIN1.Update_WHTLOCBX("X", rowICTIXFR1.Item("XFR_NO"))

            'ASCMAIN1.sql = "Select ICTIXFR2.XFR_NO WHSE_TRAN_NO, ICTIXFR2.XFR_LNO WHSE_TRAN_LNO" _
            '   & ", 'T' WHSE_TRAN_TYPE, ICTIXFR1.WHSE_CODE_TO WHSE_CODE" _
            '   & ", ICTWHSE1.WHSE_LOC_REC LOCATION_CODE, ICTIXFR2.STYLE_CODE, ICTIXFR2.COLOR_CODE" _
            '   & ", ICTIXFR2.XFR_QTY WHSE_TRAN_QTY" _
            '   & " from ICTIXFR1,ICTIXFR2,ICTWHSE1" _
            '   & " where ICTIXFR1.XFR_NO = ICTIXFR2.XFR_NO" _
            '   & "   and ICTWHSE1.WHSE_CODE = ICTIXFR1.WHSE_CODE_TO" _
            '   & "   and ICTIXFR2.XFR_NO = '" & rowICTIXFR1.Item("XFR_NO") & "'"
            'WHCMAIN1.Update_WHTLOCBX(Me)
        End If

        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub


    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("XFR_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTIXFRX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTIXFR2, "B", "Style Status Inquiry")

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
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

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

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)

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
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode Then
                        Click_Command("New", e)
                    End If
                End If
            Case "XFR_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                If Not InquiryMode Then
                    Click_Command("New")
                End If
            Case "XFR_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "INV_FREIGHT"
        End Select
    End Sub

#End Region

#Region "grdICTIXFR2"

    Private Sub grdICTIXFR2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIXFR2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"

                grdCodeDesc(grdICTIXFR2, "ICTSTYL1", "STYLE_CODE", "STYLE_DESC")
                ' FOR SOME REASON THE ABOVE CALL IS NOT LOADING THE STYLE_DESC
                If cdr IsNot Nothing Then
                    Dim STYLE_CODE As String = e.Cell.Value
                    e.Cell.Row.Cells("SALES_DIVISION_CODE").Value = cdr.Item("SALES_DIVISION_CODE")
                    e.Cell.Row.Cells("STYLE_DESC").Value = cdr.Item("STYLE_DESC")
                    Dim STYLE_CLASS_CODE As String = cdr.Item("STYLE_CLASS_CODE") & ""
                    Dim SALES_DIVISION_CODE As String = cdr.Item("SALES_DIVISION_CODE") & ""
                    Dim STYLE_COST As Decimal = Val(cdr.Item("STYLE_COST") & "")
                    e.Cell.Row.Cells("STYLE_CLASS_CODE").Value = STYLE_CLASS_CODE
                    e.Cell.Row.Cells("SALES_DIVISION_CODE").Value = SALES_DIVISION_CODE
                    e.Cell.Row.Cells("STYLE_COST").Value = STYLE_COST
                    If location_support Then
                        e.Cell.Row.Cells("LOCATION_CODE").Value = rowICTWHSE1.Item("WHSE_LOC_SHP")
                        ' USE ITEM_BIN AS A DEFAULT FOR AHA
                    End If

                    ASCMAIN1.sql = "Select * from ICTSTYC1 where STYLE_CODE = '" & STYLE_CODE & "'"
                    Dim rowICTSTYC1s() As DataRow = ASCDATA1.GetDataTable.Select("")
                    If rowICTSTYC1s.Length = 1 Then
                        e.Cell.Row.Cells("COLOR_CODE").Value = rowICTSTYC1s(0).Item("COLOR_CODE")
                    End If
                Else
                    grdICTIXFR2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "COLOR_CODE"
                grdCodeDesc(grdICTIXFR2, "ICTCOLR1", "COLOR_CODE", "COLOR_DESC")
                ' FOR SOME REASON THE ABOVE CALL IS NOT LOADING THE COLOR_DESC
                If cdr IsNot Nothing Then
                    e.Cell.Row.Cells("COLOR_DESC").Value = cdr.Item("COLOR_DESC")
                End If

            Case "XFR_QTY"
        End Select
    End Sub

    Private Sub grdICTIXFR2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIXFR2.AfterExitEditMode
        Select Case grdICTIXFR2.ActiveCell.Column.Key
            'Case "ACCT_CODE"
            '    Dim ACCT_CODE As String = grdICTIXFR2.ActiveCell.Text
            '    If ACCT_CODE <> "" Then
            '        grdICTIXFR2.ActiveCell.Value = ASCMAIN1.Format_Field(ACCT_CODE, grdGLTJRNL2.ActiveCell.Column.Key)
            '    End If
        End Select
    End Sub

    Private Sub grdICTIXFR2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIXFR2.AfterRowActivate
        With grdICTIXFR2.DisplayLayout.Bands(0)
            If grdICTIXFR2.ActiveRow.IsAddRow Then
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit

                grdICTIXFR2.ActiveCell = grdICTIXFR2.ActiveRow.Cells("STYLE_CODE")
                grdICTIXFR2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        If EntryMode = "V" Then
            Show_GL()
        End If
    End Sub

    Private Sub grdICTIXFR2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIXFR2.AfterRowsDeleted
        DisplayTotals()
    End Sub

    Private Sub grdICTIXFR2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTIXFR2.AfterRowUpdate
        DisplayTotals()
    End Sub


    'End Sub

    Private Sub grdICTIXFR2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTIXFR2.BeforeExitEditMode
        If grdICTIXFR2.ActiveCell Is Nothing Then Exit Sub
        With grdICTIXFR2.ActiveCell
            Select Case .Column.Key
                Case "STYLE_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTSTYL1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Style Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If

                Case "COLOR_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTCOLR1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Color Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                        If Not e.Cancel Then
                            cdr = LookUp("ICTSTYC1", New String() {.Row.Cells("STYLE_CODE").Value, .Text})
                            If cdr Is Nothing Then
                                ASCMAIN1.Progress("Color Code (" & .Text & ") not set up with Style (" & .Row.Cells("STYLE_CODE").Value & ")")
                                If .Value IsNot Nothing Then
                                    .Value = ""
                                End If
                                e.Cancel = True
                            End If
                        End If
                    End If

                    'Case "BAR_CODE"
                    '    If location_support Then
                    '        If .Text <> "" Then
                    '            If .Value IsNot Nothing Then
                    '                .Value = .Text.ToUpper
                    '            End If

                    '        End If
                    '        If .Text <> "" Then
                    '            cdr = LookUp("WHTBARC1", .Text)
                    '            If cdr Is Nothing Then
                    '                ASCMAIN1.Progress("Invalid Bar Code (" & .Text & ")")
                    '                If .Value IsNot Nothing Then
                    '                    .Value = ""
                    '                End If
                    '                e.Cancel = True
                    '            End If
                    '        End If
                    '    End If

                Case "LOCATION_CODE"
                    If location_support Then
                        If .Text <> "" Then
                            If .Value IsNot Nothing Then
                                .Value = .Text.ToUpper
                            End If

                        End If
                        If .Text <> "" Then
                            cdr = LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, .Text})
                            If cdr Is Nothing Then
                                ASCMAIN1.Progress("Invalid Location Code (" & .Text & ")")
                                If .Value IsNot Nothing Then
                                    .Value = ""
                                End If
                                e.Cancel = True
                            End If
                        End If
                    End If

            End Select
        End With
    End Sub

    Private Sub grdICTIXFR2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTIXFR2.BeforeRowUpdate
        With grdICTIXFR2
            If e.Row.Cells("STYLE_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTSTYL1", e.Row.Cells("STYLE_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Style Code (" & e.Row.Cells("STYLE_CODE").Text & ")", _
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If e.Row.Cells("COLOR_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTCOLR1", e.Row.Cells("COLOR_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Color Code (" & e.Row.Cells("COLOR_CODE").Text & ")", _
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
                If Not e.Cancel Then
                    LookUp("ICTSTYC1", New String() {e.Row.Cells("STYLE_CODE").Text, e.Row.Cells("COLOR_CODE").Text})
                    If cdr Is Nothing Then
                        MsgBox("Color Code (" & e.Row.Cells("COLOR_CODE").Text & ") not set up for Style (" & e.Row.Cells("STYLE_CODE").Text & ")", _
                               MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If
            End If

            If location_support Then
                'If e.Row.Cells("BAR_CODE").Text = "" Then
                '    e.Cancel = True
                'Else
                '    LookUp("WHTBARC1", e.Row.Cells("BAR_CODE").Text)
                '    If cdr Is Nothing Then
                '        MsgBox("Invalid Value entered for Bar Code (" & e.Row.Cells("BAR_CODE").Text & ")", _
                '               MsgBoxStyle.OkOnly, "Cannot Update Row")
                '        e.Cancel = True
                '    End If
                'End If

                If e.Row.Cells("LOCATION_CODE").Text = "" Then
                e.Cancel = True
            Else
                LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, e.Row.Cells("LOCATION_CODE").Text})
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Location Code (" & e.Row.Cells("LOCATION_CODE").Text & ")", _
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            End If

            If Val(e.Row.Cells("XFR_QTY").Text) = 0 Then
                MsgBox("Invalid Value entered for Qty (" & e.Row.Cells("XFR_QTY").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("XFR_NO").Text = "" Then
                    .ActiveRow.Cells("XFR_NO").Value = Absx1.CtlFor("XFR_NO").Text
                    .ActiveRow.Cells("XFR_LNO").Value = Val(dst.Tables("ICTIXFR2").Compute("Max(XFR_LNO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdICTIXFR2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIXFR2.ClickCellButton

        If grdICTIXFR2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
            Case "COLOR_CODE"
                sql_where = "COLOR_CODE in (Select COLOR_CODE from ICTSTYC1 where STYLE_CODE = '" & e.Cell.Row.Cells("STYLE_CODE").Value & "')"
            Case "LOCATION_CODE"
                sql_where = "WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
        End Select
        grdClickCellButton(grdICTIXFR2, sql_where, False) 'sql_where <> ""

    End Sub

    Private Sub grdICTIXFR2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdICTIXFR2.Error
        grdICTIXFR2.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Sub DisplayTotals()
        Dim TOTAL_COSTS As Decimal = Val(dst.Tables("ICTIXFR2").Compute("SUM(LINE_COSTS)", "") & "")
        Absx1.numFor("TOTAL_COSTS").Value = TOTAL_COSTS
    End Sub

    Private Sub grdICTIXFRX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTIXFRX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("XFR_NO").Text = e.Row.Cells("XFR_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub optGL_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optGL.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_GL()
    End Sub

    Sub Show_GL()

        If optGL.Tag <> optGL.Value Or optGL.Value = "L" Then
            optGL.Tag = optGL.Value
            If optGL.Value = "A" Then
                grdICTIXFR3.DataSource = dst.Tables("ICTIXFR3")
                Dim dvw As DataView = dst.Tables("ICTIXFR3").DefaultView
                dvw.RowFilter = ""
            ElseIf optGL.Value = "L" Then
                grdICTIXFR3.DataSource = dst.Tables("ICTIXFR3")
                Dim dvw As DataView = dst.Tables("ICTIXFR3").DefaultView
                Dim XFR_LNO As Integer = 0
                If grdICTIXFR2.ActiveRow IsNot Nothing Then
                    XFR_LNO = Val(grdICTIXFR2.ActiveRow.Cells("XFR_LNO").Text)
                End If
                dvw.RowFilter = "XFR_LNO = " & CStr(XFR_LNO)
            ElseIf optGL.Value = "S" Then
                Dim tbl As DataTable = dst.Tables("ICTIXFR3").Clone
                Dim XFR_GNO As Integer = 0
                For Each rowA234 As DataRow In ASCDATA1.SelectDistinct _
                ("ICTIXFR3", New String() {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ACCT_DESC"}).Rows
                    Dim DIST_AMT As Decimal = dst.Tables("ICTIXFR3").Compute _
                    ("SUM(DIST_AMT)", _
                     "ACCT_CODE = '" & rowA234.Item("ACCT_CODE") & "' and SEG2_CODE = '" & rowA234.Item("SEG2_CODE") & "' and SEG3_CODE = '" & rowA234.Item("SEG3_CODE") & "' and SEG4_CODE = '" & rowA234.Item("SEG4_CODE") & "'")
                    Dim row As DataRow = tbl.NewRow
                    row.Item("XFR_NO") = Absx1.txtFor("XFR_NO").Text
                    row.Item("XFR_LNO") = 0
                    XFR_GNO += 1
                    row.Item("XFR_GNO") = XFR_GNO
                    row.Item("ACCT_CODE") = rowA234.Item("ACCT_CODE")
                    row.Item("SEG2_CODE") = rowA234.Item("SEG2_CODE")
                    row.Item("SEG3_CODE") = rowA234.Item("SEG3_CODE")
                    row.Item("SEG4_CODE") = rowA234.Item("SEG4_CODE")
                    row.Item("ACCT_DESC") = rowA234.Item("ACCT_DESC")
                    row.Item("DIST_AMT") = DIST_AMT
                    tbl.Rows.Add(row)
                Next

                grdICTIXFR3.DataSource = tbl
            End If
        End If
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        Dim YP As String = cbeYP.Value
        Fill_Records("ICTIXFRX", YP)
        Sort_grdColumns(grdICTIXFRX, "XFR_NO".ToLower)
        grdICTIXFRX.Text = "Entered in " & cbeYP.Text
    End Sub

    Function Check_Qty(ByVal TABLE_NAME As String, _
                       ByVal WHSE_CODE As String, _
                       ByVal QTY_FIELD As String, _
                       ByVal S As Integer) As String

        Dim msg As String = ""

        For Each row As DataRow In dst.Tables(TABLE_NAME).Rows
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
            Dim QTY As Integer = row.Item(QTY_FIELD)
            ASCMAIN1.sql = "Select * from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and WHSE_CODE = '" & WHSE_CODE & "'"
            Dim rowICTSTAT2 As DataRow = ASCDATA1.GetDataRow
            Dim WHSE_QTY_ON_HAND As Integer = 0
            If rowICTSTAT2 IsNot Nothing Then
                WHSE_QTY_ON_HAND = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & "")
            End If
            If WHSE_QTY_ON_HAND + S * QTY < 0 Then
                msg &= vbCr & Format("Style/Color " & STYLE_CODE & "/" & COLOR_CODE & " has only " & CStr(WHSE_QTY_ON_HAND) & " On Hand")
            End If
        Next

        Return msg
    End Function

    Sub Set_Up_Reversal()
 
        Dim REVERSED_BY_XFR_NO = ASCMAIN1.Next_Control_No("ICTIXFR1.XFR_NO")
        Dim rowICTIXFR1_orig As DataRow = dst.Tables("ICTIXFR1").NewRow
        rowICTIXFR1_orig.ItemArray = rowICTIXFR1.ItemArray

        rowICTIXFR1 = dst.Tables("ICTIXFR1").Rows(0)
        rowICTIXFR1.Item("REVERSED_BY_XFR_NO") = REVERSED_BY_XFR_NO
        rowICTIXFR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowICTIXFR1.Item("LAST_DATE") = DATETIME_STAMP
        Update_Record_TDA("ICTIXFR1")

        rowICTIXFR1.ItemArray = rowICTIXFR1_orig.ItemArray
        rowICTIXFR1.AcceptChanges()
        rowICTIXFR1.SetAdded()

        With rowICTIXFR1
            .Item("REVERSES_XFR_NO") = .Item("XFR_NO")
            .Item("XFR_NO") = REVERSED_BY_XFR_NO
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("XFR_DATE") = DATETIME_STAMP.Date
            .Item("TOTAL_COSTS") *= -1

            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("REGISTER_IND") = "0"
            .Item("REGISTER_XNO") = DBNull.Value
        End With

        For Each row As DataRow In dst.Tables("ICTIXFR2").Rows
            row.Item("XFR_NO") = REVERSED_BY_XFR_NO
            If row.Item("XFR_QTY") IsNot DBNull.Value Then
                row.Item("XFR_QTY") *= -1
            End If
            If row.Item("OPS_YYYYPP") IsNot DBNull.Value Then
                row.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            End If

            row.AcceptChanges()
            row.SetAdded()
        Next
    End Sub
End Class