Imports Microsoft.Office.Interop
Imports System.Drawing

Public Class POFWPDM1

    Dim rowPOTWPDM1 As DataRow
    Dim sqlPOTWPDMX As String = ""
    Dim STYLE_GROUP_NO As String
    Dim images_folder As String = "C:\dmp\Images"
    Dim images As New Dictionary(Of String, List(Of System.Drawing.Bitmap))

    Dim Calendar_Tasks As New Dictionary(Of String, UltraWinSchedule.Appointment)

    '   Dim dvwSPTSCHD1 As DataView
    '  Dim deleted_rows() As String
    '  Dim apptEdit As Infragistics.Win.UltraWinSchedule.Appointment = Nothing
    'Dim SALES_DIVISION_CODE As String
    'Dim DEPT_CODE As String

    ' Dim sqlSPTCOOPX As String
    ' Dim SPTCODE1 As String = ""

    Dim APPR_STATUS_CODE_BackColors As New Dictionary(Of String, System.Drawing.Color)
    Dim APPR_STATUS_CODE_ForeColors As New Dictionary(Of String, System.Drawing.Color)

    '  Dim dte1 As Date = CDate(Format(Now.Date, "MM/01/yyyy"))
    '  Dim dte2 As Date = CDate(Format(dte1.AddMonths(1), "MM/01/yyyy")).AddDays(-1)

    '  Dim calTimeLine As New UltraWinSchedule.UltraCalendarInfo

    Dim colors() As System.Drawing.Color = {Color.Green, Color.Purple, Color.Blue, Color.Red}

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("POTPARM1")

        With dst
            sqlPOTWPDMX = "Select POTWPDM1.*, X.STYLE_CODE_1, X.STYLES" _
            & " from POTWPDM1, (Select STYLE_GROUP_NO, MIN (STYLE_CODE_PLM) STYLE_CODE_1, Count (*) STYLES from POTWPDM2 group by STYLE_GROUP_NO) X" _
            & " where X.STYLE_GROUP_NO = POTWPDM1.STYLE_GROUP_NO"
            ASCMAIN1.sql = sqlPOTWPDMX
            Create_TDA(.Tables.Add, "POTWPDMX", "**", 0, False, "")

            Create_TDA(.Tables.Add, "POTWPDM1", "*")
            With .Tables("POTWPDM1")
                .Columns.Add("LOGO", GetType(System.Byte()))
                '.PrimaryKey = New DataColumn() {.Columns("STYLE_GROUP_NO")}
            End With

            ASCMAIN1.sql = "Select POTWPDM2.*, ICTPLIN2.STYLE_DESC" _
                & " from POTWPDM2, ICTPLIN2 where ICTPLIN2.STYLE_CODE_PLM = POTWPDM2.STYLE_CODE_PLM" _
                & " and POTWPDM2.STYLE_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTWPDM2", "**", 0, True, "V", 2)
            'With .Tables("POTWPDM2")
            '    .Columns.Add("IMAGE", GetType(System.Byte()))
            '    ' .Columns("SEQ").DataType = GetType(System.Int32)
            'End With

            ASCMAIN1.sql = "Select POTWPDM3.*, ICTCOLR1.COLOR_DESC" _
                & " from POTWPDM3, ICTCOLR1 where ICTCOLR1.COLOR_CODE = POTWPDM3.COLOR_CODE" _
                & " and POTWPDM3.STYLE_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTWPDM3", "**", 0, True, "V", 2)

            ASCMAIN1.sql = "Select POTWPDM4.*" _
                & " from POTWPDM4" _
                & " where POTWPDM4.STYLE_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTWPDM4", "**", 0, True, "V", 3)

            Create_TDA(.Tables.Add, "POTWPDM5", "*", 1)
            Create_TDA(.Tables.Add, "POTWPDM6", "*", 1)

            .Tables("POTWPDM6").Columns("TASK_ASSIGNED").DefaultValue = Now.Date
            .Tables("POTWPDM6").Columns("TASK_DUE").DefaultValue = Now.Date

            Create_Relation("POTWPDM5", "POTWPDM6", "STYLE_GROUP_NO,STEP_LNO")

            With .Tables("POTWPDM6").Columns
                .Add("STEP_DESC", GetType(System.String), "PARENT.STEP_DESC")
                .Add("STEP_STAGE", GetType(System.String), "PARENT.STEP_STAGE")
                .Add("STEP_ACTION_DATE_NAME", GetType(System.String), "PARENT.STEP_ACTION_DATE_NAME")
            End With
            With .Tables("POTWPDM5").Columns
                .Add("TASK_COUNT", GetType(System.Int32), "COUNT(CHILD(POTWPDM5_POTWPDM6).TASK_LNO)")
            End With

            ASCMAIN1.sql = "Select * from POTWPDM7 where STYLE_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTWPDM7", "**", 0, True, "V", 0)

            Create_TDA(.Tables.Add, "ASTATTA2", "*", 3)

            Create_TDA(.Tables.Add, "POTWPDW1", "*", 1, False)
            Create_TDA(.Tables.Add, "POTWPDW2", "*", 1, False)

            ASCMAIN1.sql = "" _
                & "Select USER_ID from ASTUSER2 where SECURITY_CODE = 'PM'" & vbCrLf _
                & " union " & vbCrLf _
                & "Select USER_ID from ASTUSER1 where USER_ID = '" & ASCMAIN1.USER_ID & "'" & vbCrLf _
                & " union " & vbCrLf _
                & "Select Distinct TASK_ASSIGNED_TO from POTWPDM6 where STYLE_GROUP_NO = :PARM1 and TASK_ASSIGNED_TO is Not Null"
            Create_TDA(.Tables.Add, "POTWPDMO", "**", 0, False, "V")

        End With

        grdPOTWPDM2.DataSource = dst.Tables("POTWPDM2")
        grdPOTWPDM3.DataSource = dst.Tables("POTWPDM3")
        grdPOTWPDM4.DataSource = dst.Tables("POTWPDM4")
        grdPOTWPDM5.DataSource = dst.Tables("POTWPDM5")
        grdPOTWPDM6.DataSource = dst.Tables("POTWPDM6")
        grdPOTWPDMO.DataSource = dst.Tables("POTWPDMO")
        ' grdPOTWPDM7.DataSource = dst.Tables("POTWPDM7")
        grdPOTWPDMX.DataSource = dst.Tables("POTWPDMX")
        grdASTATTA2.DataSource = dst.Tables("ASTATTA2")

        Create_Summary(grdPOTWPDMX, "STYLE_GROUP_NO", "Count")


        For Each gcol As UltraWinGrid.UltraGridColumn In grdPOTWPDM6.DisplayLayout.Bands(0).Columns
            If gcol.Key = "" Or gcol.Key = "" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        For Each gcol As UltraWinGrid.UltraGridColumn In grdPOTWPDM5.DisplayLayout.Bands(0).Columns
            gcol.CellActivation = UltraWinGrid.Activation.NoEdit
        Next
        grdPOTWPDM5.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        grdPOTWPDM5.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdPOTWPDM5.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No

        'grdPOTWPDM5.DisplayLayout.Bands(0).Override.AllowUpdate = DefaultableBoolean.False
        'grdPOTWPDM5.DisplayLayout.Bands(0).Override.AllowDelete = DefaultableBoolean.False
        'grdPOTWPDM5.DisplayLayout.Bands(0).Override.AllowAddNew = UltraWinGrid.AllowAddNew.No

        'grdPOTWPDM5.DisplayLayout.Bands(1).Override.AllowUpdate = DefaultableBoolean.False
        'grdPOTWPDM5.DisplayLayout.Bands(1).Override.AllowDelete = DefaultableBoolean.False
        'grdPOTWPDM5.DisplayLayout.Bands(1).Override.AllowAddNew = UltraWinGrid.AllowAddNew.No

        For Each gcol As UltraWinGrid.UltraGridColumn In grdPOTWPDM5.DisplayLayout.Bands(1).Columns
            Dim gcol2 As UltraWinGrid.UltraGridColumn = grdPOTWPDM6.DisplayLayout.Bands(0).Columns(gcol.Key)
            gcol.Width = gcol2.Width
            gcol.Header.Caption = gcol2.Header.Caption
            If New String() {"STEP_LNO", "STEP_DESC", "STEP_STAGE"}.Contains(gcol.Key) Then
                gcol.Hidden = True
            Else
                gcol.Hidden = gcol2.Hidden
            End If
        Next
        '   Create_Summary(grdPOTWPDM2, "SEQ", "Count")

        eventMonthView.CalendarInfo = UltraCalendarInfo1
        UltraGanttView1.CalendarInfo = UltraCalendarInfo1
        'UltraTimelineView1.CalendarInfo = calTimeLine
        UltraTimelineView1.CalendarInfo = UltraCalendarInfo1
        Dim ultraCalendarLook1 As New UltraWinSchedule.UltraCalendarLook
        ultraCalendarLook1.ViewStyle = Infragistics.Win.UltraWinSchedule.ViewStyle.VisualStudio2005
        Me.UltraTimelineView1.CalendarLook = ultraCalendarLook1

        ASCMAIN1.Add_Value_List(grdPOTWPDM5, "STEP_STAGE")
        ASCMAIN1.Add_Value_List(grdPOTWPDM5, "STEP_STATUS")

        ASCMAIN1.Add_Value_List(grdPOTWPDM6, "STEP_STAGE")
        ASCMAIN1.Add_Value_List(grdPOTWPDM6, "TASK_DIR")
        ASCMAIN1.Add_Value_List(grdPOTWPDM6, "TASK_STATUS")

        ASCMAIN1.Add_Value_List(grdPOTWPDM5, "TASK_DIR", , , 1)
        ASCMAIN1.Add_Value_List(grdPOTWPDM5, "TASK_STATUS", , , 1)

        grdPOTWPDM5.DisplayLayout.Bands(0).Override.HeaderPlacement = UltraWinGrid.HeaderPlacement.FixedOnTop
        grdPOTWPDM5.DisplayLayout.Bands(1).Override.HeaderPlacement = UltraWinGrid.HeaderPlacement.OncePerRowIsland
        grdPOTWPDM5.DisplayLayout.Bands(1).Override.RowSelectors = DefaultableBoolean.False

        SplitContainer1.Panel1Collapsed = True

        'tabWIP.Tabs("Timeline").Visible = False
        'tabWIP.Tabs("Gantt").Visible = False

        txtUSER_ID.Text = ASCMAIN1.USER_ID

        '   splCalendar.Panel2Collapsed = True
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("STEP_TEMPLATE")

            Case "View", "Edit"
                If Absx1.txtFor("STYLE_GROUP_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a valid Quote No"
                Else
                    STYLE_GROUP_NO = Absx1.txtFor("STYLE_GROUP_NO").Text
                    rowPOTWPDM1 = LookUp("POTWPDM1", STYLE_GROUP_NO)
                    If rowPOTWPDM1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Style Group No " & STYLE_GROUP_NO
                    End If
                End If

            Case "Update"
                If dst.Tables("POTWPDM2").Select("").Length = 0 Then
                    EMsg &= vbCr & "No Details Entered"
                Else
                    For Each rowPOTWPDM2 As DataRow In dst.Tables("POTWPDM2").Select("", "", DataViewRowState.CurrentRows)
                    Next
                End If

                If EMsg = "" Then

                End If

            Case "Delete"

                If ASCMAIN1.USER_ID <> rowPOTWPDM1.Item("INIT_OPER") & "" Then
                    EMsg &= vbCr & "Only " & rowPOTWPDM1.Item("INIT_OPER") & " may Delete this Quote"
                End If

                If EMsg = "" Then
                    If MsgBox("Do you really want to Delete this Quote", _
                              MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Print", "email"
                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Customer Code Specified"
                    End If
                End If
                If dst.Tables("POTWPDM2").Select("").Length = 0 Then
                    EMsg &= vbCr & "No Styles on the Quote Sheet"
                End If

                'Case "Save Quote Sheet"
                '    If txtQUOTE_DESC.Text = "" Then
                '        EMsg &= vbCr & "Please enter a Description for the Quote Sheet"
                '    End If

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

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Print", "email"

                '  Update_Record_TDA("POTWPDM2", "1=1")
                Synch_TABLE_NAME("POTWPDM1")

                Dim FILENAME As String = ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG"
                If My.Computer.FileSystem.FileExists(FILENAME) Then
                    rowPOTWPDM1.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
                End If


                Print_Report_Begin()


                Dim RPT As String = "ICRQUOT1"


                If eItemKey = "email" Then
                    Dim tempFileName As String = rowPOTWPDM1.Item("STYLE_GROUP_NO")
                    Dim REPORT_NO As String = Generate_Report(RPT, "Quote Sheet", "", "", "PDF", tempFileName, False)
                    ' Dim FILENAME As String = REPORT_FILENAMES(REPORT_NO)
                    Print_Report_End(, True)
                    email_Quote(tempFileName)
                Else
                    Generate_Report(RPT, "Quote Sheet")
                    Print_Report_End()
                End If

                'Case "Clear Quote Sheet"
                '    dst.Tables("POTWPDM2").Rows.Clear()
                '    Setup_Style_Quoted()
                '    txtQUOTE_DESC.Text = ""
                '    Absx1.txtFor("CUST_CODE").Text = ""

                'Case "Save Quote Sheet"
                '    Update_Record_TDA("POTWPDM1")
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
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
                    .Items("Print").Settings.Enabled = iScreenMode

                    .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                    .Items("Print").Visible = ScreenMode
                    .Items("email").Visible = ScreenMode
                    .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
                    .Items("Delete").Visible = (EntryMode = "E")
                    .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
                End With

                Setup_tabPOTWPDMX()
                .Groups("Show Style Groups").Visible = Not ScreenMode

                .Groups("WIP").Visible = Not ScreenMode
                .Groups("Task Owners").Visible = ScreenMode
                .Groups("Template").Visible = Not ScreenMode
                .Groups("Group Attributes").Visible = ScreenMode

            End With
        End If


        Set_Read_Only(UltraGroupBox1, ScreenMode)
        grpAttributes.Visible = ScreenMode

        tabPOTWPDMX.Visible = Not ScreenMode

        ' lblSTEP_TEMPLATE.Visible = Not ScreenMode
        txtSTEP_TEMPLATE.Visible = Not ScreenMode

        Setup_Tasks_grid()

        If ScreenMode Then
            Set_Read_Only(grpHeader, (EntryMode = "V"))
            Set_Read_Only(grpAttributes, (EntryMode = "V"))
            Set_Read_Only_for_ctl(Absx1.txtFor("STYLE_GROUP_NAME"), (EntryMode = "V"))

            If EntryMode = "V" Then
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTWPDM2, grdPOTWPDM3, grdPOTWPDM4, grdPOTWPDM6, grdPOTWPDMO}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                Next
            Else
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTWPDM2, grdPOTWPDM3, grdPOTWPDM4, grdPOTWPDM6, grdPOTWPDMO}
                    With grd.DisplayLayout.Override
                        If grd.Name <> "grdPOTWPDM6" Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        End If
                        .AllowDelete = DefaultableBoolean.True
                        .AllowUpdate = DefaultableBoolean.True
                    End With
                Next
            End If

            splCalendar.Parent = tabWIP.Tabs("Calendar").TabPage
        Else
            Clear_Record()

            splCalendar.Parent = tabPOTWPDMX.Tabs("Calendar").TabPage
        End If

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"POTWPDMX", "POTWPDM1", "POTWPDM2", "POTWPDM3", "POTWPDM4", "POTWPDM5", "POTWPDM6", "POTWPDM7"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()

        grdPOTWPDM2.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

        Absx1.txtFor("STYLE_GROUP_NAME").Text = ""
        'Absx1.dteFor("SHIP_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.txtFor("STYLE_GROUP_NO").Text = ""

        STYLE_GROUP_NO = ""
        images.Clear()

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If EntryMode = "N" Then
            rowPOTWPDM1 = dst.Tables("POTWPDM1").NewRow
            STYLE_GROUP_NO = ASCMAIN1.Next_Control_No("POTWPDM1.STYLE_GROUP_NO")
            rowPOTWPDM1.Item("STYLE_GROUP_NO") = STYLE_GROUP_NO
            rowPOTWPDM1.Item("STYLE_GROUP_NAME") = HFs("STYLE_GROUP_NAME")
            'rowPOTWPDM1.Item("SEASON_CODE") = HFs("SEASON_CODE")
            'rowPOTWPDM1.Item("SHIP_DATE") = HFs("SHIP_DATE")
            ' rowPOTWPDM1.Item("TOTAL_QTY") = Val(HFs("TOTAL_QTY"))
            ' rowPOTWPDM1.Item("CUST_CODE") = HFs("CUST_CODE")
            ' rowPOTWPDM1.Item("VEND_CODE") = HFs("VEND_CODE")
            rowPOTWPDM1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowPOTWPDM1.Item("INIT_DATE") = DATETIME_STAMP
            rowPOTWPDM1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowPOTWPDM1.Item("LAST_DATE") = DATETIME_STAMP
            dst.Tables("POTWPDM1").Rows.Add(rowPOTWPDM1)
        Else
            rowPOTWPDM1 = Fill_Record("POTWPDM1", STYLE_GROUP_NO)
            dst.AcceptChanges()
        End If

        images.Clear()
        Calendar_Tasks.Clear()

        Fill_Records("POTWPDM4", STYLE_GROUP_NO)
        Sort_grdColumns(grdPOTWPDM4, "SEQ")

        Fill_Records("POTWPDM6", STYLE_GROUP_NO)
        Fill_Records("POTWPDM5", STYLE_GROUP_NO)
        Sort_grdColumns(grdPOTWPDM5, "SEQ")
        Setup_grdPOTWPDM5()

        Fill_Records("POTWPDM2", STYLE_GROUP_NO)
        Sort_grdColumns(grdPOTWPDM2, "STYLE_CODE_PLM")

        Fill_Records("POTWPDM3", STYLE_GROUP_NO)
        Sort_grdColumns(grdPOTWPDM3, "COLOR_CODE")

        Fill_Records("POTWPDMO", STYLE_GROUP_NO)
        Sort_grdColumns(grdPOTWPDMO, "USER_ID")

        Load_Cover_Graphic()

        EnforceConstraints(True)

        APPR_STATUS_CODE_BackColors.Clear()
        APPR_STATUS_CODE_BackColors.Add("", Color.Empty)
        APPR_STATUS_CODE_ForeColors.Clear()
        APPR_STATUS_CODE_ForeColors.Add("", Color.Orange)

        ' TEMPORARY DATA LOADING
        If EntryMode = "N" Then

            Fill_Records("POTWPDW1", Absx1.txtFor("STEP_TEMPLATE").Text)
            For Each rowPOTWPDW1 As DataRow In dst.Tables("POTWPDW1").Select("", "STEP_LNO")
                With dst.Tables("POTWPDM5").Rows
                    .Add(STYLE_GROUP_NO, _
                         rowPOTWPDW1.Item("STEP_LNO"), _
                         rowPOTWPDW1.Item("STEP_DESC"), _
                         rowPOTWPDW1.Item("SEQ"), _
                         rowPOTWPDW1.Item("STEP_STAGE"), _
                         rowPOTWPDW1.Item("STEP_BY_STYLE"), _
                         rowPOTWPDW1.Item("STEP_BY_COLOR"), _
                         rowPOTWPDW1.Item("STEP_ACTION_DATE_NAME"), DBNull.Value, "U")
                End With
            Next

            Fill_Records("POTWPDW2", Absx1.txtFor("STEP_TEMPLATE").Text)
            For Each rowPOTWPDW2 As DataRow In dst.Tables("POTWPDW2").Select("", "STEP_LNO, TASK_LNO")
                Dim TASK_ID As String = ASCMAIN1.Next_Control_No("POTWPDM6.TASK_ID")
                With dst.Tables("POTWPDM6").Rows
                    .Add(STYLE_GROUP_NO, _
                         rowPOTWPDW2.Item("STEP_LNO"), _
                         rowPOTWPDW2.Item("TASK_LNO"), _
                         rowPOTWPDW2.Item("TASK_DESC"), _
                         rowPOTWPDW2.Item("TASK_DIR"), _
                         rowPOTWPDW2.Item("TASK_NOTE"), _
                         "U", DBNull.Value, DBNull.Value, DBNull.Value, _
                         DBNull.Value, DBNull.Value, DATETIME_STAMP, ASCMAIN1.USER_ID, DBNull.Value, DBNull.Value, TASK_ID)
                End With
            Next

            Sort_grdColumns(grdPOTWPDM5, "SEQ")
            Setup_grdPOTWPDM5()

        End If

        ' calTimeLine.Owners.Clear()
        UltraCalendarInfo1.Owners.Clear()

        For Each rowPOTWPDMO As DataRow In dst.Tables("POTWPDMO").Select("")
            Dim USER_ID As String = rowPOTWPDMO.Item("USER_ID") & ""
            If USER_ID <> "" Then
                Add_Owner(USER_ID)
            End If
        Next

        UltraCalendarInfo1.Projects.Clear()
        UltraCalendarInfo1.Projects.Add("Style Group " & STYLE_GROUP_NO, Now.AddDays(10))
        UltraCalendarInfo1.Tasks.Clear()

        Create_Gantt()
        Create_Calendar()

        Setup_grdPOTWPDM2()

        grdPOTWPDMO.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Update_Record_TDA("ASTATTA2")
        Update_Record_TDA("POTWPDM1")
        Update_Record_TDA("POTWPDM2")
        Update_Record_TDA("POTWPDM3")
        Update_Record_TDA("POTWPDM4")
        Update_Record_TDA("POTWPDM5")
        Update_Record_TDA("POTWPDM6")
        Update_Record_TDA("POTWPDM7")

        For Each row As DataRow In dst.Tables("POTWPDM6").Select("")
            Send_email_alert(row)
        Next

        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Delete_Records("POTWPDM1")
        Delete_Records("POTWPDM2")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where STYLE_GROUP_NO = '" & Absx1.txtFor("STYLE_GROUP_NO").Text & "'")
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("STYLE_GROUP_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "POTWPDM1"
            E.COLUMN_NAME = "STYLE_GROUP_NO"
            E.CODE_VALUE = Absx1.txtFor("STYLE_GROUP_NO").Text
            E.DESC_VALUE = "Style Group No"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTORDR1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME


            Case "VEND_CODE"
                sql_where = "VEND_TYPE = 'S'"


        End Select

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTWPDMX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdPOTWPDM2, "SBBBB", "Show Filter", "Get Styles", "Product Line Maintenance", "Product Line Inquiry", "Sequence as Shown")
        Load_Popup_Menu(grdPOTWPDM3, "B", "Get Colors")
        Load_Popup_Menu(grdPOTWPDM4, "B", "Get Specifications")
        Load_Popup_Menu(grdPOTWPDM5, "BB", "Assign Tasks", "Expand All")
        Load_Popup_Menu(grdPOTWPDM6, "B", "Assign Tasks")
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
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdPOTWPDM2"
                    tlb_btn = DirectCast(tlb_pop.Tools("Sequence as Shown"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")
                    tlb_btn = DirectCast(tlb_pop.Tools("Get Styles"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

                Case "grdPOTWPDM3"
                    tlb_btn = DirectCast(tlb_pop.Tools("Get Colors"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

                Case "grdPOTWPDM4"
                    tlb_btn = DirectCast(tlb_pop.Tools("Get Specifications"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

                Case "grdPOTWPDM5"
                    tlb_btn = DirectCast(tlb_pop.Tools("Assign Tasks"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And grdPOTWPDMO.ActiveRow IsNot Nothing AndAlso grdPOTWPDM5.ActiveRow IsNot Nothing AndAlso grdPOTWPDM5.ActiveRow.Band.Key = "POTWPDM5" AndAlso grdPOTWPDM5.ActiveRow.Cells("STEP_STATUS").Value <> "C"
                    If grdPOTWPDMO.ActiveRow IsNot Nothing Then
                        tlb_btn.SharedProps.Caption = "Assign Tasks to " & grdPOTWPDMO.ActiveRow.Cells("USER_ID").Value
                    End If

                Case "grdPOTWPDM6"
                    tlb_btn = DirectCast(tlb_pop.Tools("Assign Tasks"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And grdPOTWPDMO.ActiveRow IsNot Nothing AndAlso grdPOTWPDM6.ActiveRow IsNot Nothing AndAlso grdPOTWPDM6.ActiveRow.Cells("TASK_STATUS").Value <> "C"
                    If grdPOTWPDMO.ActiveRow IsNot Nothing Then
                        tlb_btn.SharedProps.Caption = "Assign Tasks to " & grdPOTWPDMO.ActiveRow.Cells("USER_ID").Value
                    End If

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Get Styles"
                Get_Styles()

            Case "Get Colors"
                Get_Colors()

            Case "Get Specifications"
                Get_Specifications()

            Case "Expand All"
                grdPOTWPDM5.Rows.ExpandAll(True)
                grdPOTWPDM5.DisplayLayout.Bands(1).Override.RowAppearance.BackColor = Color.Beige
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)


            Case "Sequence as Shown"
                Dim SEQ As Integer = 0
                For Each grow As UltraWinGrid.UltraGridRow In grdPOTWPDM2.Rows
                    SEQ += 10
                    grow.Cells("SEQ").Value = SEQ
                    grow.Update()
                Next

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE_PLM").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Product Line Maintenance"
                Dim STYLE_CODE_PLM As String = grd.ActiveRow.Cells("STYLE_CODE_PLM").Text
                Dim rowICTPLIN2 As DataRow = LookUp("ICTPLIN2", STYLE_CODE_PLM)
                If rowICTPLIN2 IsNot Nothing Then
                    Context_Launch("View", STYLE_CODE_PLM, e.Tool.Key, "ICFPLIN1")
                End If

            Case "Product Line Inquiry"
                Dim STYLE_CODE_PLM As String = grd.ActiveRow.Cells("STYLE_CODE_PLM").Text
                Dim rowICTPLIN2 As DataRow = LookUp("ICTPLIN2", STYLE_CODE_PLM)
                If rowICTPLIN2 IsNot Nothing Then
                    Context_Launch("View", STYLE_CODE_PLM, e.Tool.Key, "ICFPLINI")
                End If

            Case "Assign Tasks"


                If grdPOTWPDMO.ActiveRow IsNot Nothing Then
                    If grdPOTWPDM6.ActiveRow IsNot Nothing AndAlso grdPOTWPDM6.ActiveRow.DataChanged Then
                        grdPOTWPDM6.ActiveRow.Update()
                    End If

                    Dim USER_ID As String = grdPOTWPDMO.ActiveRow.Cells("USER_ID").Value & ""
                    If USER_ID = "" Then
                        Exit Sub
                    End If

                    If grd.Name = "grdPOTWPDM5" Then
                        If grdPOTWPDM5.Selected.Rows.Count = 0 Then
                            If grdPOTWPDM5.ActiveRow IsNot Nothing Then
                                grdPOTWPDM5.ActiveRow.Selected = True
                            End If
                        End If
                        For Each grow As UltraWinGrid.UltraGridRow In grdPOTWPDM5.Selected.Rows
                            If grow.Band.Key = "POTWPDM5" Then
                                For Each grow2 As UltraWinGrid.UltraGridRow In grow.ChildBands("POTWPDM5_POTWPDM6").Rows
                                    If grow2.Cells("TASK_STATUS").Value <> "C" Then
                                        Dim TASK_ID As String = grow2.Cells("TASK_ID").Value
                                        Assign_Task(TASK_ID, USER_ID)
                                    End If
                                Next
                                Dim STEP_LNO As Int32 = Val(grow.Cells("STEP_LNO").Value & "")
                                grow.Cells("STEP_STATUS").Value = Set_STEP_STATUS(STEP_LNO)
                                grow.Update()
                            End If
                        Next
                        grdPOTWPDM5.Selected.Rows.Clear()
                    Else
                        If grdPOTWPDM6.Selected.Rows.Count = 0 Then
                            If grdPOTWPDM6.ActiveRow IsNot Nothing Then
                                grdPOTWPDM6.ActiveRow.Selected = True
                            End If
                        End If

                        Dim STEP_LNOs As New List(Of Int32)
                        Dim STEP_LNO As Int32 = 0
                        For Each grow As UltraWinGrid.UltraGridRow In grdPOTWPDM6.Selected.Rows
                            STEP_LNO = Val(grow.Cells("STEP_LNO").Value & "")
                            If Not STEP_LNOs.Contains(STEP_LNO) Then STEP_LNOs.Add(STEP_LNO)
                            If grow.Cells("TASK_STATUS").Value <> "C" Then
                                Dim TASK_ID As String = grow.Cells("TASK_ID").Value
                                Assign_Task(TASK_ID, USER_ID)
                            End If
                        Next
                        For Each STEP_LNO In STEP_LNOs
                            Dim rowPOTWPDM5 As DataRow = dst.Tables("POTWPDM5").Rows.Find(New Object() {STYLE_GROUP_NO, STEP_LNO})
                            rowPOTWPDM5.Item("STEP_STATUS") = Set_STEP_STATUS(STEP_LNO)
                        Next

                        grdPOTWPDM6.Selected.Rows.Clear()
                    End If

                    grdPOTWPDM6.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)
                    grdPOTWPDM6.ActiveRow.Activate()
                    With grdPOTWPDM6.DisplayLayout.Bands(0)
                        .Columns("TASK_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                        .Columns("TASK_NOTE").CellActivation = UltraWinGrid.Activation.NoEdit
                    End With
                End If
        End Select
    End Sub
#End Region

    Sub Assign_Task(TASK_ID As String, USER_ID As String)
        Dim rowPOTWPDM6 As DataRow = dst.Tables("POTWPDM6").Select("TASK_ID = '" & TASK_ID & "'")(0)
        rowPOTWPDM6.Item("TASK_STATUS") = "O"
        rowPOTWPDM6.Item("TASK_ASSIGNED_TO") = USER_ID
        rowPOTWPDM6.Item("TASK_ASSIGNED") = DATETIME_STAMP
        rowPOTWPDM6.Item("LAST_DATE") = DATETIME_STAMP
        rowPOTWPDM6.Item("LAST_OPER") = ASCMAIN1.USER_ID
        If rowPOTWPDM6.Item("TASK_DUE") & "" = "" Then
            rowPOTWPDM6.Item("TASK_DUE") = DATETIME_STAMP.Date
        End If
        Dim rowPOTWPDM7 As DataRow = dst.Tables("POTWPDM7").NewRow
        For Each dcol As DataColumn In dst.Tables("POTWPDM7").Columns
            If dst.Tables("POTWPDM6").Columns.Contains(dcol.ColumnName) Then
                rowPOTWPDM7.Item(dcol.ColumnName) = rowPOTWPDM6.Item(dcol.ColumnName)
            End If
        Next
        rowPOTWPDM7.Item("WORK_ID") = ASCMAIN1.Next_Control_No("POTWPDM7.WORK_ID")
        rowPOTWPDM7.Item("WORK_PERFORMED") = "Task Assigned"
        dst.Tables("POTWPDM7").Rows.Add(rowPOTWPDM7)

        Remove_Calendar_Task(rowPOTWPDM6)
        Add_Calendar_Task(rowPOTWPDM6)

    End Sub

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    If Not InquiryMode Then
                        Click_Command("New", e)
                    End If
                End If
            Case "STYLE_GROUP_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("View", e)
                End If

        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                If Not InquiryMode And Not ScreenMode Then
                    Click_Command("New")
                End If

        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case ""
        End Select
    End Sub

#End Region

#Region "grdPOTWPDM2"

    Private Sub grdPOTWPDM2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM2.AfterCellUpdate
        If Not e.Cell.Row.IsDataRow Then Exit Sub
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE_PLM"

                grdCodeDesc(grdPOTWPDM2, "ICTPLIN2", "STYLE_CODE_PLM", "STYLE_DESC")
                ' FOR SOME REASON THE ABOVE CALL IS NOT LOADING THE STYLE_DESC
                If cdr IsNot Nothing Then
                    Dim STYLE_CODE_PLM As String = e.Cell.Value
                    e.Cell.Row.Cells("STYLE_DESC").Value = cdr.Item("STYLE_DESC")
                    'e.Cell.Row.Cells("SALES_DIVISION_CODE").Value = cdr.Item("SALES_DIVISION_CODE")
                    'e.Cell.Row.Cells("STYLE_CLASS_CODE").Value = cdr.Item("STYLE_CLASS_CODE") & ""
                    'e.Cell.Row.Cells("STYLE_PRICE").Value = cdr.Item("STYLE_PRICE")

                Else
                    grdPOTWPDM2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If
        End Select
    End Sub

    Private Sub grdPOTWPDM2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTWPDM2.AfterRowActivate

        If Not grdPOTWPDM2.ActiveRow.IsDataRow Then Exit Sub

        With grdPOTWPDM2.DisplayLayout.Bands(0)
            If grdPOTWPDM2.ActiveRow.IsAddRow Then
                .Columns("STYLE_CODE_PLM").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdPOTWPDM2.ActiveCell = grdPOTWPDM2.ActiveRow.Cells("STYLE_CODE_PLM")
                grdPOTWPDM2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("STYLE_CODE_PLM").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        Setup_grdPOTWPDM2()
    End Sub

    Sub Setup_grdPOTWPDM2()

        tplStyle.Tiles.Clear()

        If grdPOTWPDM2.ActiveRow Is Nothing OrElse (grdPOTWPDM2.ActiveRow.IsAddRow Or Not grdPOTWPDM2.ActiveRow.IsDataRow) Then
            'grdPOTWPDM4.Visible = False
            tabStyleDetails.Visible = False
        Else
            'grdPOTWPDM4.Visible = True
            tabStyleDetails.Visible = True
            Dim STYLE_CODE_PLM As String = grdPOTWPDM2.ActiveRow.Cells("STYLE_CODE_PLM").Value & ""
            Dim dvw As DataView = DirectCast(grdPOTWPDM4.DataSource, DataTable).DefaultView
            dvw.RowFilter = "STYLE_GROUP_NO = '" & STYLE_GROUP_NO & "' and STYLE_CODE_PLM = '" & STYLE_CODE_PLM & "'"
            Sort_grdColumns(grdPOTWPDM4, "SEQ")

            If images.ContainsKey(STYLE_CODE_PLM) Then
                For Each I As System.Drawing.Bitmap In images(STYLE_CODE_PLM)
                    Dim t As New Infragistics.Win.Misc.UltraTile
                    Dim P As New UltraWinEditors.UltraPictureBox
                    P.Image = I
                    '   pic.Image = I
                    t.Control = P
                    t.Text = "MY TILE"
                    tplStyle.Tiles.Add(t)
                    P.Visible = True
                    t.Visible = True
                Next
                tplStyle.Visible = True
            End If
        End If
    End Sub

    Private Sub grdPOTWPDM2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdPOTWPDM2.AfterRowsDeleted

        If images.Count > 0 Then
            Dim STYLE_CODE_PLMs As New List(Of String)
            For Each rowPOTWPDM2 As DataRow In dst.Tables("POTWPDM2").Select("")
                Dim STYLE_CODE_PLM As String = rowPOTWPDM2.Item("STYLE_CODE_PLM")
                If Not STYLE_CODE_PLMs.Contains(STYLE_CODE_PLM) Then
                    STYLE_CODE_PLMs.Add(STYLE_CODE_PLM)
                End If
            Next
            Dim STYLE_CODE_PLMs_to_delete As New List(Of String)
            For Each k As String In images.Keys
                If Not STYLE_CODE_PLMs.Contains(k) Then
                    STYLE_CODE_PLMs_to_delete.Add(k)
                End If
            Next
            For Each k As String In STYLE_CODE_PLMs_to_delete
                images.Remove(k)
            Next
        End If



        Setup_grdPOTWPDM2()
    End Sub
    Private Sub grdPOTWPDM2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTWPDM2.AfterRowUpdate

        Dim STYLE_GROUP_NO As String = e.Row.Cells("STYLE_GROUP_NO").Value
        Dim STYLE_CODE_PLM As String = e.Row.Cells("STYLE_CODE_PLM").Value
        Dim rowPOTWPDM2 As DataRow = dst.Tables("POTWPDM2").Rows.Find(New Object() {STYLE_GROUP_NO, STYLE_CODE_PLM})
        If Not images.ContainsKey(STYLE_CODE_PLM) Then
            Get_Images(STYLE_CODE_PLM)
        End If
    End Sub

    Private Sub grdPOTWPDM2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTWPDM2.BeforeExitEditMode
        If grdPOTWPDM2.ActiveCell Is Nothing Then Exit Sub
        If Not grdPOTWPDM2.ActiveRow.IsDataRow Then Exit Sub
        With grdPOTWPDM2.ActiveCell
            Select Case .Column.Key
                Case "STYLE_CODE_PLM"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTPLIN2", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Style Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If

            End Select
        End With
    End Sub

    Private Sub grdPOTWPDM2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTWPDM2.BeforeRowUpdate
        With grdPOTWPDM2
            If e.Row.Cells("STYLE_CODE_PLM").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTPLIN2", e.Row.Cells("STYLE_CODE_PLM").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Style Code (" & e.Row.Cells("STYLE_CODE_PLM").Text & ")", _
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("STYLE_GROUP_NO").Text = "" Then
                    .ActiveRow.Cells("STYLE_GROUP_NO").Value = Absx1.CtlFor("STYLE_GROUP_NO").Text
                    .ActiveRow.Cells("SEQ").Value = Val(dst.Tables("POTWPDM2").Compute("Max(SEQ)", "") & "") + 10
                End If
            End If
        End With


    End Sub

    Private Sub grdPOTWPDM2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM2.ClickCellButton

        If grdPOTWPDM2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE_PLM"
        End Select
        grdClickCellButton(grdPOTWPDM2, sql_where, True)

    End Sub

#End Region

#Region "grdPOTWPDM3"
    Private Sub grdPOTWPDM3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM3.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "COLOR_CODE"
                Dim COLOR_CODE As String = e.Cell.Value & ""
                grdCodeDesc(grdPOTWPDM3, "ICTCOLR1", "COLOR_CODE", "COLOR_DESC")
        End Select
    End Sub

    Private Sub grdPOTWPDM3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTWPDM3.BeforeRowUpdate
        With grdPOTWPDM3
            If Not e.Cancel Then
                If e.Row.Cells("STYLE_GROUP_NO").Text = "" Then
                    .ActiveRow.Cells("STYLE_GROUP_NO").Value = Absx1.CtlFor("STYLE_GROUP_NO").Text
                End If
            End If
        End With
    End Sub

    Private Sub grdPOTWPDM3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM3.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdPOTWPDM3, sql_where, sql_where <> "")
    End Sub
#End Region

#Region "grdPOTWPDM4"
    Private Sub grdPOTWPDM4_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM4.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "SPEC_CODE"
                Dim SPEC_CODE As String = e.Cell.Value & ""
                grdCodeDesc(grdPOTWPDM4, "POTWPDMS", "SPEC_CODE", "SPEC_DESC")
        End Select
    End Sub

    Private Sub grdPOTWPDM4_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTWPDM4.BeforeRowUpdate
        With grdPOTWPDM4
            If Not e.Cancel Then
                If e.Row.Cells("STYLE_GROUP_NO").Text = "" Then
                    .ActiveRow.Cells("STYLE_GROUP_NO").Value = Absx1.CtlFor("STYLE_GROUP_NO").Text
                    .ActiveRow.Cells("STYLE_CODE_PLM").Value = grdPOTWPDM2.ActiveRow.Cells("STYLE_CODE_PLM").Value
                    Dim sqlx As String = "STYLE_GROUP_NO = '" & .ActiveRow.Cells("STYLE_GROUP_NO").Value & "' and STYLE_CODE_PLM = '" & .ActiveRow.Cells("STYLE_CODE_PLM").Value & "'"
                    .ActiveRow.Cells("SPEC_LNO").Value = Val(dst.Tables("POTWPDM4").Compute("MAX(SPEC_LNO)", sqlx) & "") + 1
                    .ActiveRow.Cells("SEQ").Value = Val(dst.Tables("POTWPDM4").Compute("Max(SEQ)", "") & "") + 10
                End If
            End If
        End With
    End Sub

    Private Sub grdPOTWPDM4_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM4.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdPOTWPDM4, sql_where, sql_where <> "")
    End Sub
#End Region

#Region "grdPOTWPDM5"
    Private Sub grdPOTWPDM5_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM5.AfterCellUpdate
        Select Case e.Cell.Column.Key

        End Select
    End Sub

    Private Sub grdPOTWPDM5_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTWPDM5.AfterRowActivate
        Setup_grdPOTWPDM5()
    End Sub

    Private Sub grdPOTWPDM5_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTWPDM5.BeforeRowUpdate
        With grdPOTWPDM5
            If Not e.Cancel Then
                If e.Row.Cells("STYLE_GROUP_NO").Text = "" Then
                    .ActiveRow.Cells("STYLE_GROUP_NO").Value = Absx1.CtlFor("STYLE_GROUP_NO").Text
                    Dim sqlx As String = "STYLE_GROUP_NO = '" & .ActiveRow.Cells("STYLE_GROUP_NO").Value & "'"
                    .ActiveRow.Cells("STEP_LNO").Value = Val(dst.Tables("POTWPDM5").Compute("MAX(STEP_LNO)", sqlx) & "") + 1
                    .ActiveRow.Cells("SEQ").Value = Val(dst.Tables("POTWPDM5").Compute("Max(SEQ)", "") & "") + 10
                End If
            End If
        End With
    End Sub

    Private Sub grdPOTWPDM5_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM5.ClickCellButton
        'Dim sql_where As String = ""
        'grdClickCellButton(grdPOTWPDM5, sql_where, sql_where <> "")

        If e.Cell.Row.Band.Key = "POTWPDM5_POTWPDM6" Then
            If e.Cell.Row.Cells("TASK_STATUS").Value & "" = "U" Then
                MsgBox("Task has not yet been assigned")
            Else
                Dim EVENT_KEY As String = e.Cell.Row.Cells("STYLE_GROUP_NO").Value & ":" & e.Cell.Row.Cells("STEP_LNO").Value & ":" & e.Cell.Row.Cells("TASK_LNO").Value
                Edit_Task(EVENT_KEY)
            End If
        End If

    End Sub

    Sub Setup_grdPOTWPDM5()

        If grdPOTWPDM5.ActiveRow Is Nothing Then
            grdPOTWPDM6.Visible = False
        Else
            grdPOTWPDM6.Visible = True
            Dim STEP_LNO As Int32 = Val(grdPOTWPDM5.ActiveRow.Cells("STEP_LNO").Value & "")
            Dim STEP_DESC As String = grdPOTWPDM5.ActiveRow.Cells("STEP_DESC").Value & ""
            Dim dvw As DataView = DirectCast(grdPOTWPDM6.DataSource, DataTable).DefaultView
            dvw.RowFilter = "STEP_LNO = " & CStr(STEP_LNO)
            grdPOTWPDM6.Text = "Tasks defined to Step " & CStr(STEP_LNO) & ":" & STEP_DESC
        End If
    End Sub
#End Region

#Region "grdPOTWPDM6"
    Private Sub grdPOTWPDM6_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM6.AfterCellUpdate
        Select Case e.Cell.Column.Key

        End Select
    End Sub

    Private Sub grdPOTWPDM6_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTWPDM6.AfterRowActivate

        If Not grdPOTWPDM6.ActiveRow.IsDataRow Then Exit Sub

        With grdPOTWPDM6.DisplayLayout.Bands(0)
            If grdPOTWPDM6.ActiveRow.IsAddRow OrElse grdPOTWPDM6.ActiveRow.Cells("TASK_STATUS").Value & "" = "U" Then
                .Columns("TASK_DESC").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("TASK_NOTE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("TASK_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("TASK_NOTE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdPOTWPDM6_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTWPDM6.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            ASCMAIN1.sql = "Select Count (*) from POTWPDM7" _
                & " where STYLE_GROUP_NO = '" & grow.Cells("STYLE_GROUP_NO").Value & "'" _
                & "   and STEP_LNO = " & grow.Cells("STEP_LNO").Value _
                & "   and TASK_LNO = " & grow.Cells("TASK_LNO").Value _
                & "   and WORK_PERFORMED <> 'Task Assigned'"
            Dim C As Integer = Val(ASCDATA1.GetDataValue)
            If C > 0 Then
                MsgBox("Cannot Delete Task " & grow.Cells("TASK_LNO").Value _
                       & " within Step " & grow.Cells("STEP_LNO").Value _
                       & " of Style Group " & grow.Cells("STYLE_GROUP_NO").Value _
                       & vbCrLf & vbCrLf & "Work has already been recorded under this task", _
                       MsgBoxStyle.OkOnly, "Verification")
                e.Cancel = True
            End If
        Next
    End Sub

    Private Sub grdPOTWPDM6_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTWPDM6.BeforeRowUpdate
        With grdPOTWPDM6
            If Not e.Cancel Then
                If e.Row.Cells("STYLE_GROUP_NO").Text = "" Then
                    .ActiveRow.Cells("STYLE_GROUP_NO").Value = Absx1.CtlFor("STYLE_GROUP_NO").Text
                    Dim STEP_LNO As Int32 = Val(grdPOTWPDM5.ActiveRow.Cells("STEP_LNO").Value & "")
                    .ActiveRow.Cells("STEP_LNO").Value = STEP_LNO
                    Dim sqlx As String = "STYLE_GROUP_NO = '" & .ActiveRow.Cells("STYLE_GROUP_NO").Value & "' and STEP_LNO = " & CStr(STEP_LNO)
                    .ActiveRow.Cells("TASK_LNO").Value = Val(dst.Tables("POTWPDM6").Compute("MAX(TASK_LNO)", sqlx) & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdPOTWPDM6_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM6.ClickCellButton
        'Dim sql_where As String = ""
        'grdClickCellButton(grdPOTWPDM6, sql_where, sql_where <> "")
        If e.Cell.Row.Cells("TASK_STATUS").Value & "" = "U" Then
            MsgBox("Task has not yet been assigned")
        Else
            Dim EVENT_KEY As String = e.Cell.Row.Cells("STYLE_GROUP_NO").Value & ":" & e.Cell.Row.Cells("STEP_LNO").Value & ":" & e.Cell.Row.Cells("TASK_LNO").Value
            Edit_Task(EVENT_KEY)
        End If
    End Sub
#End Region


#Region "grdPOTWPDMO"

    Private Sub grdPOTWPDMO_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDMO.AfterCellUpdate
        If Not e.Cell.Row.IsDataRow Then Exit Sub
        Select Case e.Cell.Column.Key
            Case "USER_ID"
                'grdCodeDesc(grdPOTWPDMO, "ASTUSER1", "USER_ID", "USER_NAME")
                ' FOR SOME REASON THE ABOVE CALL IS NOT LOADING THE STYLE_DESC
                LookUp("ASTUSER1", e.Cell.Row.Cells("USER_ID").Value)
                If cdr IsNot Nothing Then
                    'Dim STYLE_CODE_PLM As String = e.Cell.Value
                    ' e.Cell.Row.Cells("STYLE_DESC").Value = cdr.Item("STYLE_DESC")
                Else
                    grdPOTWPDMO.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If
        End Select
    End Sub

    Private Sub grdPOTWPDMO_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTWPDMO.AfterRowActivate

        If Not grdPOTWPDMO.ActiveRow.IsDataRow Then Exit Sub

        With grdPOTWPDMO.DisplayLayout.Bands(0)
            If grdPOTWPDMO.ActiveRow.IsAddRow Then
                .Columns("USER_ID").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdPOTWPDMO.ActiveCell = grdPOTWPDMO.ActiveRow.Cells("USER_ID")
                grdPOTWPDMO.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("USER_ID").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdPOTWPDMO_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTWPDMO.AfterRowUpdate

    End Sub

    Private Sub grdPOTWPDMO_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTWPDMO.BeforeExitEditMode
        If grdPOTWPDMO.ActiveCell Is Nothing Then Exit Sub
        If Not grdPOTWPDMO.ActiveRow.IsDataRow Then Exit Sub
        With grdPOTWPDMO.ActiveCell
            Select Case .Column.Key
                Case "USER_ID"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToLower
                        End If
                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ASTUSER1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid User ID (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If

            End Select
        End With
    End Sub

    Private Sub grdPOTWPDMO_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTWPDMO.BeforeRowsDeleted

    End Sub

    Private Sub grdPOTWPDMO_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTWPDMO.BeforeRowUpdate
        With grdPOTWPDMO
            If e.Row.Cells("USER_ID").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ASTUSER1", e.Row.Cells("USER_ID").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for User ID (" & e.Row.Cells("USER_ID").Text & ")", _
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            ElseIf e.Row.IsAddRow Then
                Add_Owner(e.Row.Cells("USER_ID").Value)
            End If
        End With
    End Sub

    Private Sub grdPOTWPDMO_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDMO.ClickCellButton

        If grdPOTWPDMO.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "USER_ID"
        End Select
        grdClickCellButton(grdPOTWPDMO, sql_where, True)

    End Sub

#End Region

    Private Sub grdPOTWPDMX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTWPDMX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("STYLE_GROUP_NO").Text = e.Row.Cells("STYLE_GROUP_NO").Text
            Click_Command("View")
        End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        ASCMAIN1.sql = sqlPOTWPDMX
        Dim STYLE_GROUP_NO As String = Absx1.txtFor("STYLE_GROUP_NO").Text
        If optShow.Value = "A" And STYLE_GROUP_NO = "" Then
            grdPOTWPDMX.Text = "All Quotes"
        ElseIf optShow.Value = "M" Then
            ASCMAIN1.sql &= " and (INIT_OPER = '" & ASCMAIN1.USER_ID & "' or LAST_OPER = '" & ASCMAIN1.USER_ID & "')"
            grdPOTWPDMX.Text = "Quotes entered or modified by Me"
        ElseIf optShow.Value = "C" Or STYLE_GROUP_NO <> "" Then
            ASCMAIN1.sql &= " and STYLE_GROUP_NO = '" & STYLE_GROUP_NO & "'"
            grdPOTWPDMX.Text = "Quotes associated with " & STYLE_GROUP_NO
        End If
        Fill_Records("POTWPDMX")
        Sort_grdColumns(grdPOTWPDMX, "STYLE_GROUP_NO".ToLower)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub email_Quote(tempFileName As String)
        Dim ATTACHMENTs As New Dictionary(Of String, String)
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim CUST_NAME As String = Absx1.txtFor("CUST_NAME").Text
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
        ATTACHMENTs.Add(tempFileName & ".pdf", ASCMAIN1.Folders("Temp") & tempFileName & ".pdf")

        Dim SUBJECT As String = "Quote Sheet"
        Dim PFX As String = ""

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        If CUST_CODE <> "" Then
            EMAIL_ADDRESSs.Add(rowARTCUST1.Item("CUST_EMAIL") & "", rowARTCUST1.Item("CUST_CONTACT") & "")
        End If

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                SUBJECT, "POTWPDM1", False, True, CUST_CODE, CUST_NAME, "Customer")
        If SEND_NO <> "" Then
            TAC.TACMAIN1.Record_Event("ARTCUST1", CUST_CODE, Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, "QUOEML", "Quote Sheet emailed", SEND_NO)
        End If
    End Sub

    Private Sub optShow_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optShow.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Refresh_Documents()
    End Sub

    Sub Get_Styles()
        Add_Codes(grdPOTWPDM2, "ICTPLIN2", "STYLE_CODE_PLM", "Styles")
    End Sub

    Sub Get_Colors()
        Add_Codes(grdPOTWPDM3, "ICTCOLR1", "COLOR_CODE", "Colors")
    End Sub

    Sub Get_Specifications()
        Add_Codes(grdPOTWPDM4, "POTWPDMS", "SPEC_CODE", "Specifications")
    End Sub

    Sub Get_Images(STYLE_CODE_PLM As String)
        Dim I As New List(Of System.Drawing.Bitmap)

        Dim IMAGE_FOLDER As String = images_folder & "\COLUMN_NAME\STYLE_GROUP_NO\STYLE_CODE_PLM\" & STYLE_CODE_PLM
        If My.Computer.FileSystem.DirectoryExists(IMAGE_FOLDER) Then
            For Each file As String In My.Computer.FileSystem.GetFiles(IMAGE_FOLDER)
                Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(file)
                'Dim myBitmapMetadata As New System.Windows.Forms .me
                Dim imgba() As Byte = Nothing
                Dim b As System.Drawing.Bitmap = ASCMAIN1.Get_Image(IMAGE_FOLDER, fi.Name, True, , , imgba)
                I.Add(b)
                ' we will need imgba when we go to print
            Next
        End If

        If images.ContainsKey(STYLE_CODE_PLM) Then
            images(STYLE_CODE_PLM) = I
        Else
            images.Add(STYLE_CODE_PLM, I)
        End If
    End Sub

    Private Sub cmdLocateCoverGraphic_Click(sender As System.Object, e As System.EventArgs) Handles cmdLocateCoverGraphic.Click
        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select a Graphic to use as the Cover Sheet for this Style Group"
            ' openFileDialog1.Filter = "png files (*.png)|*.png"
            openFileDialog1.Filter = "jpg files (*.jpg)|*.jpg"
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then
            Dim FOLDER_NAME As String = images_folder & "\COLUMN_NAME\STYLE_GROUP_NO\"

            My.Computer.FileSystem.CopyFile(FILENAME, FOLDER_NAME & STYLE_GROUP_NO & ".png", True)
            Load_Cover_Graphic()
        End If
    End Sub

    Sub Load_Cover_Graphic()
        Dim b As System.Drawing.Bitmap = ASCMAIN1.Get_Image(images_folder & "\COLUMN_NAME\STYLE_GROUP_NO\", STYLE_GROUP_NO & ".png", False, , , Nothing)
        picCover.Image = b
    End Sub

    Overloads Sub Process_DragDrop()
        ENTITY = Dropped_On_Context()

        If ENTITY.READ_ONLY Then
            Exit Sub
        End If

        lblNowProcessing.Visible = True
        Application.DoEvents()

        Dim files() As String = eDND.Data.GetData(DataFormats.FileDrop)

        If files IsNot Nothing Then
            For Each FILENAME As String In files
                Dim Msg As String = Attach_File(FILENAME, , , , False)
                If Msg <> "" Then
                    MsgBox(Msg, MsgBoxStyle.OkOnly, "Error Attempting to Attach File ")
                End If
            Next
        Else
            Try
                Dim outlook As Outlook.Application = CType(Microsoft.VisualBasic.Interaction.GetObject("", "Outlook.Application"), Outlook.Application)
                Dim explorer As Outlook.Explorer = outlook.ActiveExplorer

                For i As Int32 = 0 To explorer.Selection.Count - 1
                    Dim mail As Outlook.MailItem = CType(explorer.Selection.Item(i + 1), Outlook.MailItem)
                    mail.SaveAs(ASCMAIN1.Folders("Temp") & "mailitem.msg")

                    Dim FILENAME As String = ASCMAIN1.Folders("Temp") & "mailitem.msg"
                    Dim Msg As String = Attach_File(FILENAME, mail.Subject, mail.SenderName, mail.SentOn, False)
                    If Msg <> "" Then
                        MsgBox(Msg, MsgBoxStyle.OkOnly, "Error Attempting to Attach File ")
                    End If
                    mail = Nothing
                Next

                outlook = Nothing
                explorer = Nothing

            Catch ex As System.Exception

                MsgBox(ex, "Error - Outlook request not found")

            End Try

        End If

        lblNowProcessing.Visible = False
        Application.DoEvents()

        Me.Activate()
        'ASCMAIN1.ActiveForm.Activate()

    End Sub

#Region "grdASTATTA2"
    Private Sub grdASTATTA2_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles grdASTATTA2.DragEnter
        If grdASTATTA2.AllowDrop Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub grdASTATTA2_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles grdASTATTA2.DragDrop
        lblNowProcessing.Visible = True
        eDND = e
        Process_DragDrop()
        lblNowProcessing.Visible = False
    End Sub

#End Region

    Private Sub tabWIP_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabWIP.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_Tasks_grid()
    End Sub

    Sub Setup_Tasks_grid()
        If Not ScreenMode Then
            grdPOTWPDM6.Parent = tabPOTWPDMX.Tabs("Tasks").TabPage
       
            With grdPOTWPDM6.DisplayLayout.Bands(0)
                .Columns("STEP_LNO").Hidden = False
                .Columns("STEP_DESC").Hidden = False
                .Columns("STEP_STAGE").Hidden = False
            End With

            Get_Tasks()


        Else

            If tabWIP.SelectedTab.Text = "Tasks" Then
                grdPOTWPDM6.Parent = tabWIP.SelectedTab.TabPage
                Dim dvw As DataView = DirectCast(grdPOTWPDM6.DataSource, DataTable).DefaultView
                dvw.RowFilter = ""
                With grdPOTWPDM6.DisplayLayout.Bands(0)
                    .Columns("STEP_LNO").Hidden = False
                    .Columns("STEP_DESC").Hidden = False
                    .Columns("STEP_STAGE").Hidden = False
                End With
                grdPOTWPDM6.Text = "All Tasks defined to Style Group"
            ElseIf tabWIP.SelectedTab.Text = "Steps" Then
                grdPOTWPDM6.Parent = splWorkflow.Panel2
                With grdPOTWPDM6.DisplayLayout.Bands(0)
                    .Columns("STEP_LNO").Hidden = True
                    .Columns("STEP_DESC").Hidden = True
                    .Columns("STEP_STAGE").Hidden = True
                End With
                Setup_grdPOTWPDM5()
            ElseIf tabWIP.SelectedTab.Text = "Calendar" Then
                grdPOTWPDM6.Parent = splCalendar.Panel2
                With grdPOTWPDM6.DisplayLayout.Bands(0)
                    .Columns("STEP_LNO").Hidden = False
                    .Columns("STEP_DESC").Hidden = False
                    .Columns("STEP_STAGE").Hidden = False
                End With
                Setup_Tasks_Due()
                ' Setup_grdPOTWPDM5()
            End If
        End If
    End Sub

    Sub Get_Tasks()

        Dim dvw As DataView = DirectCast(grdPOTWPDM6.DataSource, DataTable).DefaultView
        dvw.RowFilter = ""
        Dim sql As String = ""
        Select Case optTasks.Value
            Case "A"
                grdPOTWPDM6.Text = "All Open Tasks"
            Case "M"
                grdPOTWPDM6.Text = "All Open Tasks assigned to Me (" & ASCMAIN1.USER_ID & ")"
                sql = " and TASK_ASSIGNED_TO = '" & ASCMAIN1.USER_ID & "'"
            Case "U"
                grdPOTWPDM6.Text = "All Open Tasks assigned to " & txtUSER_ID.Text
                sql = " and TASK_ASSIGNED_TO = '" & txtUSER_ID.Text & "'"
        End Select

        EnforceConstraints(False)
        ASCMAIN1.sql = "Select * from POTWPDM6 where TASK_STATUS = 'O'" & sql
        Fill_Records("POTWPDM6", "", True, ASCMAIN1.sql)
        ASCMAIN1.sql = "Select * from POTWPDM5 where (STYLE_GROUP_NO,STEP_LNO) in (Select STYLE_GROUP_NO,STEP_LNO from POTWPDM6 where TASK_STATUS = 'O'" & sql & ")"
        Fill_Records("POTWPDM5", "", True, ASCMAIN1.sql)
        EnforceConstraints(False)

    End Sub
#Region "eventMonthView"

    Private Sub eventMonthView_ActivitiesDragComplete(sender As Object, e As Infragistics.Win.UltraWinSchedule.ActivitiesDragCompleteEventArgs) Handles eventMonthView.ActivitiesDragComplete

    End Sub

    Private Sub eventMonthView_AfterActiveOwnerChanged(sender As Object, e As Infragistics.Win.UltraWinSchedule.AfterActiveOwnerChangedEventArgs) Handles eventMonthView.AfterActiveOwnerChanged

    End Sub

    Private Sub eventMonthView_AfterAppointmentEdit(sender As Object, e As Infragistics.Win.UltraWinSchedule.AfterAppointmentEditEventArgs) Handles eventMonthView.AfterAppointmentEdit

    End Sub

    Private Sub eventMonthView_AppointmentsDragDrop(sender As Object, e As Infragistics.Win.UltraWinSchedule.AppointmentsDragDropEventArgs) Handles eventMonthView.AppointmentsDragDrop

    End Sub

    Private Sub eventMonthView_BeforeAppointmentEdit(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinSchedule.BeforeAppointmentEditEventArgs) Handles eventMonthView.BeforeAppointmentEdit
        e.Cancel = True
        'Edit_Appointment(e.Appointment.Tag)
    End Sub

    Private Sub eventMonthView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles eventMonthView.Click
        Setup_Tasks_Due()
    End Sub

    Private Sub eventMonthView_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles eventMonthView.MouseDoubleClick

        If InquiryMode Then Exit Sub
        If Not e.Button = MouseButtons.Left Then Exit Sub

        Dim point As System.Drawing.Point = New System.Drawing.Point(e.X, e.Y)
        Dim objAppointment As Infragistics.Win.UltraWinSchedule.Appointment
        Dim objDay As Infragistics.Win.UltraWinSchedule.Day

        ' See if we clicked an Appointment
        objAppointment = Me.eventMonthView.GetAppointmentFromPoint(e.X, e.Y)
        objDay = Me.eventMonthView.GetDayFromPoint(e.X, e.Y)
        If objAppointment Is Nothing AndAlso objDay Is Nothing Then
            Exit Sub
        End If

        'Setup_SPTSCHD1()

        If objAppointment IsNot Nothing Then
            Dim APPT_TAG As String = objAppointment.Tag
            Edit_Task(APPT_TAG)
        End If
    End Sub

#End Region

#Region "Form Procedures"

    Sub Setup_Tasks_Due()
        Dim SCHED_DATE As Date = eventMonthView.CalendarInfo.ActiveDay.Date
        Dim dvwPOTWPDM6 As DataView = DirectCast(grdPOTWPDM6.DataSource, DataTable).DefaultView
        dvwPOTWPDM6.RowFilter = "TASK_DUE <= #" & Format(SCHED_DATE, "MM/dd/yyyy") & "# AND TASK_DUE >= #" & Format(SCHED_DATE, "MM/dd/yyyy") & "#"
        grdPOTWPDM6.Text = "Tasks Due on " & Format(SCHED_DATE, "MM/dd/yyyy")
    End Sub

    Sub Remove_Calendar_Task(rowPOTWPDM6 As DataRow) ' ByVal SCHED_NO As String)

        Dim TASK_key As String = rowPOTWPDM6.Item("STYLE_GROUP_NO") & ":" & rowPOTWPDM6.Item("STEP_LNO") & ":" & rowPOTWPDM6.Item("TASK_LNO")

        If Calendar_Tasks.ContainsKey(TASK_key) Then
            eventMonthView.CalendarInfo.Appointments.Remove(Calendar_Tasks(TASK_key))
            Calendar_Tasks.Remove(TASK_key)
        End If
    End Sub

    Sub Edit_Task(ByVal EVENT_key As String)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Editing Task")

        If EVENT_key.Length = 0 Then Exit Sub
        Dim STYLE_GROUP_NO As String = Split(EVENT_key, ":")(0)
        Dim STEP_LNO As Integer = Val(Split(EVENT_key, ":")(1))
        Dim TASK_LNO As Integer = Val(Split(EVENT_key, ":")(2))
        Dim rowPOTWPDM6 As DataRow = dst.Tables("POTWPDM6").Rows.Find(New Object() {STYLE_GROUP_NO, STEP_LNO, TASK_LNO})
        Dim TASK_ID As String = rowPOTWPDM6.Item("TASK_ID")

        Using F As New POFWPDM2

            F.frmASFBASE0 = Me
            F.rowPOTWPDM6 = rowPOTWPDM6
            F.CUST_CODE = Absx1.txtFor("CUST_CODE").Text
            F.VEND_CODE = Absx1.txtFor("VEND_CODE").Text
            'F.TASK_ASSIGNED_TO = F.rowPOTWPDM6.Item("TASK_ASSIGNED_TO")
            F.STYLE_GROUP_NO = Absx1.txtFor("STYLE_GROUP_NO").Text

            F.CUST_NAME = Absx1.txtFor("CUST_NAME").Text
            F.VEND_NAME = Absx1.txtFor("VEND_NAME").Text
            F.STYLE_GROUP_NAME = Absx1.txtFor("STYLE_GROUP_NAME").Text

            F.ShowDialog()

            If F.UPDATED Then

                ' Remove_Calendar_Task(EVENT_key)
                Remove_Calendar_Task(rowPOTWPDM6)

                Dim rowPOTWPDM5 As DataRow = dst.Tables("POTWPDM5").Rows.Find(New Object() {STYLE_GROUP_NO, STEP_LNO})
                rowPOTWPDM5.Item("STEP_STATUS") = Set_STEP_STATUS(STEP_LNO)

                If rowPOTWPDM6.Item("TASK_DUE") & "" <> "" Then
                    Add_Calendar_Task(rowPOTWPDM6)
                End If

                '  Add_Appointment(F.rowPOTWPDM6)
                'row.ItemArray = F.rowPOTWPDM6.ItemArray
                ' dst.Tables("POTWPDM6").AcceptChanges()
            End If
        End Using

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Function Set_STEP_STATUS(STEP_LNO As Int32) As String
        Dim STEP_STATUS As String = ""
        For Each rowPOTWPDM6_status As DataRow In dst.Tables("POTWPDM6").Select("STEP_LNO = " & CStr(STEP_LNO))
            If rowPOTWPDM6_status.Item("TASK_STATUS") & "" = "U" Then
                STEP_STATUS = "U"
                Exit For
            ElseIf rowPOTWPDM6_status.Item("TASK_STATUS") & "" = "O" Then
                STEP_STATUS = "O"
            ElseIf rowPOTWPDM6_status.Item("TASK_STATUS") & "" = "C" And STEP_STATUS = "" Then
                STEP_STATUS = "C"
            End If
        Next

        Return STEP_STATUS

    End Function
#End Region

    Private Sub tabPOTWPDMX_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabPOTWPDMX.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabPOTWPDMX()
    End Sub

    Sub Setup_tabPOTWPDMX()
        With UltraExplorerBar1
            .Groups("Screen Control").Visible = (tabPOTWPDMX.SelectedTab.Key = "Groups")
            .Groups("Show Style Groups").Visible = (tabPOTWPDMX.SelectedTab.Key = "Groups")
            .Groups("Filters").Visible = (tabPOTWPDMX.SelectedTab.Key = "Calendar")
        End With

        spl.Panel1Collapsed = (tabPOTWPDMX.SelectedTab.Key = "Calendar")
    End Sub

    Private Sub UltraTimelineView1_ActiveOwnerChanged(sender As Object, e As Infragistics.Win.UltraWinSchedule.ActiveOwnerChangedEventArgs) Handles UltraTimelineView1.ActiveOwnerChanged

    End Sub

    Private Sub UltraTimelineView1_AppointmentEditModeEntered(sender As Object, e As Infragistics.Win.UltraWinSchedule.AppointmentEditModeEnteredEventArgs) Handles UltraTimelineView1.AppointmentEditModeEntered

    End Sub

    Private Sub UltraTimelineView1_AppointmentEditModeExited(sender As Object, e As Infragistics.Win.UltraWinSchedule.AppointmentEditModeExitedEventArgs) Handles UltraTimelineView1.AppointmentEditModeExited

    End Sub

    Private Sub UltraTimelineView1_AppointmentEnteringEditMode(sender As Object, e As Infragistics.Win.UltraWinSchedule.AppointmentEnteringEditModeEventArgs) Handles UltraTimelineView1.AppointmentEnteringEditMode

        ' If InquiryMode Then Exit Sub

        Dim objAppointment As Infragistics.Win.UltraWinSchedule.Appointment = e.Appointment


        If objAppointment IsNot Nothing Then
            Edit_Task(objAppointment.Tag)
        Else ' New Appointment

            'Using f As New POFWPDM2

            '    f.frmASFBASE0 = Me
            '    Dim STEP_LNO As Int32 = 0
            '    Dim TASK_LNO As Int32 = 0
            '    f.rowPOTWPDM6 = dst.Tables("POTWPDM6").Rows.Find(New Object() {STYLE_GROUP_NO, STEP_LNO, TASK_LNO})
            '    f.CUST_CODE = Absx1.txtFor("CUST_CODE").Text
            '    f.VEND_CODE = Absx1.txtFor("VEND_CODE").Text
            '    f.STYLE_GROUP_NO = Absx1.txtFor("STYLE_GROUP_NO").Text

            '    f.ShowDialog()

            '    If f.UPDATED Then
            '        Add_Appointment(f.rowPOTWPDM6)
            '        Dim SCHED_NO As String = f.rowPOTWPDM6.Item("SCHED_NO")
            '        Dim row As DataRow = dst.Tables("SPTSCHD1").Rows.Find(SCHED_NO)
            '        If row IsNot Nothing Then
            '            row.ItemArray = f.rowPOTWPDM6.ItemArray
            '        Else
            '            row = dst.Tables("SPTSCHD1").NewRow
            '            row.ItemArray = f.rowPOTWPDM6.ItemArray
            '            dst.Tables("SPTSCHD1").Rows.Add(row)
            '        End If
            '        dst.Tables("SPTSCHD1").AcceptChanges()
            '    End If
            'End Using
        End If

        e.Cancel = True
    End Sub

    Private Sub UltraTimelineView1_AppointmentResized(sender As Object, e As Infragistics.Win.UltraWinSchedule.AppointmentResizedEventArgs) Handles UltraTimelineView1.AppointmentResized

    End Sub

    Private Sub UltraTimelineView1_AppointmentResizing(sender As Object, e As Infragistics.Win.UltraWinSchedule.AppointmentResizingEventArgs) Handles UltraTimelineView1.AppointmentResizing
        e.Cancel = True
    End Sub

    Private Sub UltraTimelineView1_AppointmentsDragDrop(sender As Object, e As Infragistics.Win.UltraWinSchedule.AppointmentsDragDropEventArgs) Handles UltraTimelineView1.AppointmentsDragDrop

    End Sub

    Private Sub UltraTimelineView1_AppointmentsDragging(sender As Object, e As Infragistics.Win.UltraWinSchedule.AppointmentsDraggingEventArgs) Handles UltraTimelineView1.AppointmentsDragging

    End Sub


    Private Sub UltraTimelineView1_MouseDoubleClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles UltraTimelineView1.MouseDoubleClick


    End Sub

    Private Sub grdPOTWPDM5_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTWPDM5.InitializeRow

        If e.Row.Band.Key = "POTWPDM5" Then

            With e.Row.Cells("STEP_STATUS")
                If .Value & "" = "U" Then
                    .Appearance.BackColor = Color.LightGray
                ElseIf .Value & "" = "O" Then
                    .Appearance.BackColor = Color.Yellow
                ElseIf .Value & "" = "C" Then
                    .Appearance.BackColor = Color.LightGreen
                Else
                    .Appearance.BackColor = Color.Red
                End If
            End With
        Else
            With e.Row.Cells("TASK_STATUS")
                If .Value & "" = "U" Then
                    .Appearance.BackColor = Color.LightGray
                ElseIf .Value & "" = "O" Then
                    .Appearance.BackColor = Color.Yellow
                ElseIf .Value & "" = "C" Then
                    .Appearance.BackColor = Color.LightGreen
                Else
                    .Appearance.BackColor = Color.Red
                End If
            End With

            Dim TASK_ASSIGNED_TO As String = e.Row.Cells("TASK_ASSIGNED_TO").Value & ""
            If TASK_ASSIGNED_TO <> "" And APPR_STATUS_CODE_ForeColors.ContainsKey(TASK_ASSIGNED_TO) Then
                e.Row.Cells("TASK_ASSIGNED_TO").Appearance.ForeColor = APPR_STATUS_CODE_ForeColors(TASK_ASSIGNED_TO)

            End If
 
        End If
    End Sub

    Private Sub grdPOTWPDM6_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTWPDM6.InitializeRow
        With e.Row.Cells("TASK_STATUS")
            If .Value & "" = "U" Then
                .Appearance.BackColor = Color.LightGray
            ElseIf .Value & "" = "O" Then
                .Appearance.BackColor = Color.Yellow
            ElseIf .Value & "" = "C" Then
                .Appearance.BackColor = Color.LightGreen
            Else
                .Appearance.BackColor = Color.Red
            End If
        End With

    End Sub

    Sub Add_Owner(USER_ID As String)

        Dim i As Integer = UltraCalendarInfo1.Owners.Count Mod colors.Length
 
        If Not APPR_STATUS_CODE_BackColors.ContainsKey(USER_ID) Then
            APPR_STATUS_CODE_BackColors.Add(USER_ID, Color.Empty)
        End If

        If Not APPR_STATUS_CODE_ForeColors.ContainsKey(USER_ID) Then
            APPR_STATUS_CODE_ForeColors.Add(USER_ID, colors(i))
        End If

        If Not UltraCalendarInfo1.Owners.Contains(USER_ID) Then
            UltraCalendarInfo1.Owners.Add(USER_ID)
        End If
    End Sub

    Sub Create_Gantt()

        'eventMonthView.CalendarInfo.MinDate = dst.Tables("SPTCOOPX").Compute("MIN(DROP_DATE)", "")
        'eventMonthView.CalendarInfo.MaxDate = dst.Tables("SPTCOOPX").Compute("MAX(DROP_DATE)", "")


        For Each rowPOTWPDM5 As DataRow In dst.Tables("POTWPDM5").Select("", "SEQ")
            Dim STEP_DESC As String = rowPOTWPDM5.Item("STEP_DESC") & ""
            Dim DT_INIT As Date = DATETIME_STAMP.Date
            If EntryMode <> "N" Then
                DT_INIT = rowPOTWPDM1.Item("INIT_DATE")
            End If

            Dim T As UltraWinSchedule.Task = UltraCalendarInfo1.Tasks.Add(DT_INIT, TimeSpan.FromDays(3), STEP_DESC)
            Dim STEP_LNO As Int32 = Val(rowPOTWPDM5.Item("STEP_LNO") & "")
            Dim D1 As Integer = 0 ' Now.Ticks Mod 20
            Dim D2 As Integer = 1 ' Now.Ticks Mod 7
            Dim DT As Date = Now.Date.AddDays(D1)
            Dim DT_DUE As Date = DT.AddDays(D2)

            For Each rowPOTWPDM6 As DataRow In dst.Tables("POTWPDM6").Select("STEP_LNO = " & CStr(STEP_LNO), "TASK_LNO")

                Dim TASK_DESC As String = rowPOTWPDM6.Item("TASK_DESC") & ""
                If TASK_DESC = "" Then TASK_DESC = STEP_DESC

                'rowPOTWPDM6.Item("TASK_ASSIGNED") = DT
                'rowPOTWPDM6.Item("TASK_DUE") = DT_DUE
                Dim owner As String = rowPOTWPDM6.Item("TASK_ASSIGNED_TO") & ""

                Dim T2 As UltraWinSchedule.Task = T.Tasks.Add(DT, DT_DUE.Subtract(DT), TASK_DESC)
                T2.Tag = rowPOTWPDM6.Item("TASK_ID")

                ' Determine where in the control the right button was pressed
                '  Dim objAppointment As New Infragistics.Win.UltraWinSchedule.Appointment(DT, DT_DUE)

                Dim appt As Infragistics.Win.UltraWinSchedule.Appointment = eventMonthView.CalendarInfo.Appointments.Add(DT, DT_DUE, TASK_DESC)
                appt.Tag = rowPOTWPDM6.Item("STYLE_GROUP_NO") & ":" & rowPOTWPDM6.Item("STEP_LNO") & ":" & rowPOTWPDM6.Item("TASK_LNO")

                Dim APPR_STATUS_CODE As String = owner
                appt.Appearance.BackColor = APPR_STATUS_CODE_BackColors(APPR_STATUS_CODE)
                appt.Appearance.ForeColor = APPR_STATUS_CODE_ForeColors(APPR_STATUS_CODE)

                appt.Description = TASK_DESC
                If owner <> "" Then appt.Owner = appt.CalendarInfo.Owners(owner)
                appt.AllDayEvent = True
            Next
        Next
    End Sub

    Sub Create_Calendar()
        For Each rowPOTWPDM6 As DataRow In dst.Tables("POTWPDM6").Select("", "STEP_LNO,TASK_LNO")
            If rowPOTWPDM6.Item("TASK_DUE") & "" <> "" Then
                Add_Calendar_Task(rowPOTWPDM6)
            End If
        Next
    End Sub

    Sub Add_Calendar_Task(rowPOTWPDM6 As DataRow)

        Dim STEP_LNO As Int32 = Val(rowPOTWPDM6.Item("STEP_LNO") & "")
        Dim TASK_LNO As Int32 = Val(rowPOTWPDM6.Item("TASK_LNO") & "")
        Dim TASK_DUE As Date = rowPOTWPDM6.Item("TASK_DUE")

        'If TASK_DUE.DayOfWeek = DayOfWeek.Saturday Then
        '    TASK_DUE = TASK_DUE.AddDays(-1)
        'ElseIf TASK_DUE.DayOfWeek = DayOfWeek.Sunday Then
        '    TASK_DUE = TASK_DUE.AddDays(-2)
        'End If

        Dim rowPOTWPDM5 As DataRow = dst.Tables("POTWPDM5").Rows.Find(New Object() {STYLE_GROUP_NO, STEP_LNO})
        Dim STEP_DESC As String = rowPOTWPDM5.Item("STEP_DESC") & ""

        Dim TASK_DESC As String = rowPOTWPDM6.Item("TASK_DESC") & ""
        If TASK_DESC = "" Then TASK_DESC = STEP_DESC

        Dim TASK_ASSIGNED_TO As String = rowPOTWPDM6.Item("TASK_ASSIGNED_TO") & ""
        Dim appt As Infragistics.Win.UltraWinSchedule.Appointment = eventMonthView.CalendarInfo.Appointments.Add(TASK_DUE, TASK_DUE, TASK_DESC)
        appt.Tag = rowPOTWPDM6.Item("STYLE_GROUP_NO") & ":" & rowPOTWPDM6.Item("STEP_LNO") & ":" & rowPOTWPDM6.Item("TASK_LNO")

        'appt.Appearance.BackColor = APPR_STATUS_CODE_BackColors(TASK_ASSIGNED_TO)
        If Format(TASK_DUE, "yyyyMMdd") < Format(Now, "yyyyMMdd") Then
            appt.Appearance.BackColor = Color.Red
        ElseIf Format(TASK_DUE, "yyyyMMdd") < Format(Now.AddDays(7), "yyyyMMdd") Then
            appt.Appearance.BackColor = Color.Yellow
        End If

        appt.Appearance.ForeColor = APPR_STATUS_CODE_ForeColors(TASK_ASSIGNED_TO)

        appt.Description = TASK_DESC
        If TASK_ASSIGNED_TO <> "" Then appt.Owner = appt.CalendarInfo.Owners(TASK_ASSIGNED_TO)
        appt.AllDayEvent = True

        If Calendar_Tasks.ContainsKey(appt.Tag) Then
            Calendar_Tasks(appt.Tag) = appt
        Else
            Calendar_Tasks.Add(appt.Tag, appt)
        End If
    End Sub

    Private Sub UltraGanttView1_ClientSizeChanged(sender As Object, e As System.EventArgs) Handles UltraGanttView1.ClientSizeChanged

    End Sub

    Private Sub UltraGanttView1_MouseDoubleClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles UltraGanttView1.MouseDoubleClick

        'If InquiryMode Then Exit Sub
        'If Not e.Button = MouseButtons.Left Then Exit Sub

        Dim point As System.Drawing.Point = New System.Drawing.Point(e.X, e.Y)
        'Dim t As UltraWinSchedule.Task = UltraGanttView1.GetChildAtPoint(e.Location)

        'Dim objAppointment As Infragistics.Win.UltraWinSchedule.Appointment
        'Dim objDay As Infragistics.Win.UltraWinSchedule.Day

        '' See if we clicked an Appointment
        'objAppointment = Me.eventMonthView.GetAppointmentFromPoint(e.X, e.Y)
        'objDay = Me.eventMonthView.GetDayFromPoint(e.X, e.Y)
        'If objAppointment Is Nothing AndAlso objDay Is Nothing Then
        '    Exit Sub
        'End If

        ''Setup_SPTSCHD1()

        'If objAppointment IsNot Nothing Then
        '    Dim APPT_TAG As String = objAppointment.Tag
        '    Edit_Task(APPT_TAG)
        'End If

        '    UltraGanttView1.AutoDisplayTaskDialog = UltraWinGanttView.AutoDisplayTaskDialog.GridRowSelector

        If UltraGanttView1.ActiveTask IsNot Nothing Then
            Dim T2 As UltraWinSchedule.Task = UltraGanttView1.ActiveTask
            If T2.Tasks.Count = 0 Then
                Dim row As DataRow = dst.Tables("POTWPDM6").Select("TASK_ID = '" & T2.Tag & "'")(0)
                Edit_Task(row.Item("STYLE_GROUP_NO") & ":" & row.Item("STEP_LNO") & ":" & row.Item("TASK_LNO"))
            End If

            'If objAppointment IsNot Nothing Then
            '    Dim APPT_TAG As String = objAppointment.Tag
            '    Edit_Task(APPT_TAG)
            'End If
            '    e.Cancel = True
        End If
    End Sub

    Private Sub UltraGanttView1_TaskDialogDisplaying(sender As Object, e As Infragistics.Win.UltraWinGanttView.TaskDialogDisplayingEventArgs) Handles UltraGanttView1.TaskDialogDisplaying
        If 1 = 1 Then Exit Sub

        Dim T2 As UltraWinSchedule.Task = e.Task
        If T2.Tasks.Count = 0 Then
            Dim row As DataRow = dst.Tables("POTWPDM6").Select("TASK_ID = '" & T2.Tag & "'")(0)
            Dim EVENT_key As String = row.Item("STYLE_GROUP_NO") & ":" & row.Item("STEP_LNO") & ":" & row.Item("TASK_LNO")
            Edit_Task(EVENT_key)
        End If

        'If objAppointment IsNot Nothing Then
        '    Dim APPT_TAG As String = objAppointment.Tag
        '    Edit_Task(APPT_TAG)
        'End If
        e.Cancel = True
    End Sub

    Private Sub UltraGanttView1_Click(sender As System.Object, e As System.EventArgs) Handles UltraGanttView1.Click
        With UltraGanttView1.GridSettings
            .AllowColumnMoving = False
            .ColumnSettings(UltraWinSchedule.TaskField.Dependencies).Visible = False
        End With
       
        Me.UltraGanttView1.GridSettings.ColumnSettings(UltraWinSchedule.TaskField.Resources).Visible = False
        Me.UltraGanttView1.AutoDisplayDefaultContextMenu = UltraWinGanttView.AutoDisplayDefaultContextMenu.No

        Me.UltraGanttView1.GridSettings.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Me.UltraGanttView1.GridSettings.AllowColumnMoving = False


    End Sub

    Private Sub UltraGanttView1_TaskElementDragComplete(sender As Object, e As Infragistics.Win.UltraWinGanttView.TaskElementDragCompleteEventArgs) Handles UltraGanttView1.TaskElementDragComplete
     

    End Sub

    Private Sub grdPOTWPDM5_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTWPDM5.InitializeLayout

    End Sub

    Private Sub grdPOTWPDMO_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTWPDMO.InitializeRow

        Dim USER_ID As String = e.Row.Cells("USER_ID").Value & ""
        If APPR_STATUS_CODE_ForeColors.ContainsKey(USER_ID) Then
            e.Row.Cells("USER_ID").Appearance.ForeColor = APPR_STATUS_CODE_ForeColors(USER_ID)
        End If
    End Sub

    Private Sub optTasks_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optTasks.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Get_Tasks()
    End Sub

    Private Sub txtUSER_ID_ValueChanged(sender As Object, e As System.EventArgs) Handles txtUSER_ID.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Get_Tasks()
    End Sub

    Sub Send_email_alert(rowPOTWPDM6 As DataRow)
        If 1 <> 1 Then
            Dim ATTACHMENTs As Dictionary(Of String, String) = Nothing
            Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", rowPOTWPDM6.Item("TASK_ASSIGNED_TO") & "")

            Dim SUBJECT As String = "Task"
            'SUBJECT = "Regency PO " & PO_ORDER_NOs(0)

            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            'EMAIL_ADDRESSs.Add("wjz@absolution.com", "Walter J. Zielenski")
            EMAIL_ADDRESSs.Add(rowASTUSER1.Item("USER_EMAIL") & "", rowASTUSER1.Item("USER_NAME") & "")

            Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                   (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                    SUBJECT, "PDTASK", False, True, "", "", "")

            If SEND_NO <> "" Then
                ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
                    & " Select 'POTWPDM6', STYLE_GROUP_NO || ':' || STEP_LNO || ':' ||  TASK_LNO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'TASKEM','Task Emailed', '" & SEND_NO & "'" _
                    & " from POTWPDM6 " & vbCrLf _
                    & " where STYLE_GROUP_NO = '" & rowPOTWPDM6.Item("STYLE_GROUP_NO") & "' and STEP_LNO = '" & rowPOTWPDM6.Item("STEP_LNO") & "' and TASK_LNO ='" & rowPOTWPDM6.Item("STEP_LNO") & "'"
                ASCDATA1.ExecuteSQL()
            End If
        End If
    End Sub
End Class